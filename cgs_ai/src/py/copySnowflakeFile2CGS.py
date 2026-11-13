"""
=====================================================================
  Program Name  : copySnowflakeFile2CGS.py
  Author        : Manuel Figallo
  Purpose       : Copy a file from a Snowflake workspace or stage to a
                  local or UNC target path, replacing the manual
                  download-and-drag step.
  Version       : 1.0beta
  Created       : 2026-09-23
  Last Modified : 2026-09-23

  Dependencies:
    STANDARD LIBRARY ONLY to import. snowflake.snowpark and
    snowflake.connector are imported LAZILY, and only when the source
    actually has to be fetched from Snowflake. A plain file-to-file copy
    needs neither.

  Description:
    WHERE THIS RUNS DECIDES WHETHER IT CAN WORK.

    In a Snowflake notebook the kernel executes on SNOWFLAKE infrastructure,
    not on your laptop and not in your Citrix session. From there:

        \\\\Client\\C$\\Temp\\...                    is YOUR machine    -> unreachable
        \\\\anf-citrix-cd96.cms.local\\Ctx-XA-...    is a CGS file server -> unreachable
        /tmp/..., ./data/...                     is the kernel's disk -> fine

    No library can bridge that: there is no network route from Snowflake
    compute to a Citrix client drive, and "\\\\Client\\C$" is not even a real
    host -- it is a redirection that only exists inside your own session.
    So this function does NOT pretend. It works out where it is running,
    and when the target cannot be reached from there it says so precisely,
    names the reason, and prints the two ways to actually get the file
    across (see UnreachableTargetError).

    Run it FROM THE MACHINE THAT OWNS THE TARGET -- a CGS Windows desktop or
    your Citrix session, with snowflake-connector-python installed -- and it
    does the whole job unattended: pull from Snowflake, write to the UNC
    path, verify the byte count.

  Input Parameters (required first):
    SOURCE_LOCATION (REQUIRED, str) - one of:
        snow://workspace/USER$FHER.PUBLIC.DEFAULT$/versions/head/data/x.csv
        @MY_DB.MY_SCHEMA.MY_STAGE/data/x.csv
        MY_DB.MY_SCHEMA.MY_STAGE/data/x.csv
        /any/ordinary/path/x.csv        (copied straight through)
    TARGET_LOCATION (REQUIRED, str) - destination file path, local or UNC,
        e.g. \\\\Client\\C$\\Temp\\synthetic_medicare_claims.csv
    Session       (optional) - an existing Snowpark Session. Omitted, the
        active session is used, else a connector session is opened.
    Overwrite     (optional, bool, default True) - replace an existing target.
    DryRun        (optional, bool, default False) - resolve, check access and
        report, writing nothing. Use it to test a target path safely.
    TempDir       (optional, str) - where a staged file lands before the copy.
    TimeoutSeconds(optional, int, default 300) - cap on the Snowflake fetch.

  Returns:
    dict with SourceLocation, TargetLocation, BytesCopied, Strategy,
    Runtime, Steps (every attempt, in order) and Copied (bool).

  Exit codes (CLI):
    0 = success, 2 = config error, 3 = I/O or access failure.
=====================================================================
"""

from __future__ import annotations

import argparse
import os
import platform
import re
import shutil
import sys
import tempfile
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

_HERE = Path(__file__).resolve()
if str(_HERE.parent.parent.parent) not in sys.path:
    sys.path.insert(0, str(_HERE.parent.parent.parent))

from src.utils.helpers import ensureParent              # noqa: E402
from src.utils.logger import logError, logInfo, logWarn  # noqa: E402

__version__ = "1.0beta"

EXIT_OK, EXIT_CONFIG_ERROR, EXIT_IO_ERROR = 0, 2, 3

#: snow://workspace/<workspace>/versions/<version>/<path inside the workspace>
WORKSPACE_URI = re.compile(
    r"^snow://workspace/(?P<workspace>[^/]+)/versions/(?P<version>[^/]+)/"
    r"(?P<path>.+)$", re.IGNORECASE)

#: @DB.SCHEMA.STAGE/path, with the leading @ optional.
STAGE_REFERENCE = re.compile(
    r"^@?(?P<stage>[A-Za-z0-9_$]+\.[A-Za-z0-9_$]+\.[A-Za-z0-9_$]+)"
    r"(?:/(?P<path>.*))?$")

#: A Windows drive letter or a UNC share. Either is unreachable from Linux.
WINDOWS_TARGET = re.compile(r"^(?:[A-Za-z]:[\\/]|\\\\)")

#: Where Snowflake notebooks have been seen to mount workspace files. Tried
#: in order; the first that exists wins, and every miss is logged.
WORKSPACE_MOUNT_CANDIDATES = (
    "/snowflake/workspace",
    "/home/udf/workspace",
    "/tmp/workspace",
    ".",
)


class SnowflakeCopyError(RuntimeError):
    """Base class for every failure this module raises."""


class UnreachableTargetError(SnowflakeCopyError):
    """The target cannot be written from where this code is running.

    Raised instead of a confusing permission error when the code is on
    Snowflake compute and the target is a Windows or UNC path. The message
    carries the two ways round it, because this is the failure a first-time
    user will hit and a bare OSError tells them nothing.
    """


class SourceNotFoundError(SnowflakeCopyError):
    """The source could not be located or fetched."""


class TargetWriteError(SnowflakeCopyError):
    """The target path exists in principle but could not be written."""


def describeRuntime() -> Dict[str, Any]:
    """Work out where this code is executing.

    Parameters: none.
    Returns:
        dict with OnSnowflake (bool), Platform, Reasons (list[str]) - the
        evidence behind the verdict, which is logged so a wrong guess can
        be seen and argued with rather than silently acted on.
    """
    reasons: List[str] = []
    # Snowflake's notebook and UDF containers set these; none of them exist
    # on a CGS desktop or in a Citrix session.
    for variable in ("SNOWFLAKE_ACCOUNT", "SNOWFLAKE_HOST",
                     "SNOWFLAKE_WAREHOUSE", "SNOWFLAKE_QUERY_ID"):
        if os.environ.get(variable):
            reasons.append(f"environment variable {variable} is set")
    for marker in ("/snowflake", "/home/udf"):
        if Path(marker).is_dir():
            reasons.append(f"{marker} exists")
    return {
        "OnSnowflake": bool(reasons),
        "Platform": platform.system(),
        "Reasons": reasons or ["no Snowflake markers found"],
    }


def classifySource(SOURCE_LOCATION: str) -> Dict[str, str]:
    """Decide what kind of source location this is.

    Parameters: SOURCE_LOCATION (str).
    Returns:
        dict with Kind ("workspace" | "stage" | "file"), plus the parts that
        kind needs: Workspace/Version/Path, Stage/Path, or Path.
    Raises: ValueError - the location is empty.
    """
    text = str(SOURCE_LOCATION or "").strip().strip('"').strip("'")
    # Smart quotes survive a copy-paste out of Word and Teams, and would
    # otherwise show up as a baffling "file not found".
    text = text.replace("“", "").replace("”", "")
    if not text:
        raise ValueError("required parameter 'SOURCE_LOCATION' is missing or empty")

    workspace = WORKSPACE_URI.match(text)
    if workspace:
        return {"Kind": "workspace", "Raw": text,
                "Workspace": workspace.group("workspace"),
                "Version": workspace.group("version"),
                "Path": workspace.group("path")}
    if text.lower().startswith("snow://"):
        raise ValueError(
            f"{text} starts with snow:// but is not a workspace URI. "
            f"Expected snow://workspace/<workspace>/versions/<version>/<path>")

    if not WINDOWS_TARGET.match(text) and "/" not in text.split(".")[0]:
        stage = STAGE_REFERENCE.match(text)
        if stage and not Path(text).exists():
            return {"Kind": "stage", "Raw": text,
                    "Stage": "@" + stage.group("stage"),
                    "Path": stage.group("path") or ""}
    return {"Kind": "file", "Raw": text, "Path": text}


def checkTargetReachable(TARGET_LOCATION: str, runtime: Dict[str, Any]) -> None:
    """Refuse early when the target cannot be reached from here.

    Parameters:
        TARGET_LOCATION (str) - the destination.
        runtime (dict)        - from describeRuntime().
    Raises:
        UnreachableTargetError - running on Snowflake compute (or any
        non-Windows host) with a Windows or UNC target.

    This is the whole point of the function's error handling. A Snowflake
    kernel writing to \\\\Client\\C$ does not fail with "access denied"; it
    fails with something far more confusing, because \\\\Client is not a host
    that exists anywhere except inside a Citrix session.
    """
    if not WINDOWS_TARGET.match(str(TARGET_LOCATION)):
        return
    if runtime["Platform"] == "Windows":
        return
    where = ("Snowflake compute" if runtime["OnSnowflake"]
             else f"a {runtime['Platform']} host")
    raise UnreachableTargetError(
        f"cannot write {TARGET_LOCATION} from {where}.\n"
        f"  WHY: the kernel is executing on {where}, not on your machine. "
        f"A path beginning \\\\Client\\ or a mapped drive letter only exists "
        f"inside your own Windows or Citrix session, and there is no network "
        f"route to it from here. Evidence: "
        f"{'; '.join(runtime['Reasons'])}.\n"
        f"  FIX 1 (recommended): run this function ON the machine that owns "
        f"the target -- a CGS Windows desktop or your Citrix session -- with "
        f"snowflake-connector-python installed. It will then pull from "
        f"Snowflake and write the UNC path unattended.\n"
        f"  FIX 2: in the notebook, copy to the kernel's own disk first "
        f"(TARGET_LOCATION='./synthetic_medicare_claims.csv'), then bring it "
        f"down with the notebook's download control or st.download_button.")


def resolveWorkspaceLocally(source: Dict[str, str],
                            steps: List[str]) -> Optional[Path]:
    """Look for a workspace file already mounted on this filesystem.

    Parameters:
        source (dict) - from classifySource(), Kind "workspace".
        steps (list)  - appended to with every candidate tried.
    Returns: the Path when found, else None.

    Inside a Snowflake notebook the workspace is usually on the kernel's
    disk already, which makes the copy a plain file copy and avoids a stage
    round-trip. The mount point has moved between Snowflake releases, so
    every candidate is tried and each miss is recorded rather than assumed.
    """
    relative = source["Path"]
    for root in WORKSPACE_MOUNT_CANDIDATES:
        candidate = Path(root) / relative
        if candidate.is_file():
            steps.append(f"workspace mount hit: {candidate}")
            return candidate
        steps.append(f"workspace mount miss: {candidate}")
    # The bare filename, in case only the leaf was mounted.
    leaf = Path(relative).name
    for root in WORKSPACE_MOUNT_CANDIDATES:
        candidate = Path(root) / leaf
        if candidate.is_file():
            steps.append(f"workspace mount hit by filename: {candidate}")
            return candidate
    return None


def workspaceStageReference(source: Dict[str, str]) -> str:
    """Translate a workspace URI into the stage reference Snowflake accepts.

    Parameters: source (dict) - from classifySource(), Kind "workspace".
    Returns: str e.g. "snow://workspace/USER$FHER.PUBLIC.DEFAULT$/versions/head/data/x.csv"

    Snowflake's GET understands the snow:// form directly, so the URI is
    passed through unchanged. Kept as a named function because that is the
    line most likely to need changing when the workspace API moves.
    """
    return source["Raw"]


def fetchFromSnowflake(source: Dict[str, str], tempDir: str,
                       Session: Any, steps: List[str],
                       TimeoutSeconds: int) -> Path:
    """Download a staged or workspace file to `tempDir`.

    Parameters:
        source (dict)  - from classifySource().
        tempDir (str)  - an existing directory to download into.
        Session        - a Snowpark Session, or None to find/open one.
        steps (list)   - appended to with every attempt and its outcome.
        TimeoutSeconds (int).
    Returns: Path of the downloaded file.
    Raises: SourceNotFoundError - every strategy failed; the message lists
            what was tried and why each failed.

    Two strategies, in order: Snowpark's session.file.get, then a raw SQL
    GET through snowflake-connector-python. Both are lazy imports, so a
    machine with neither installed still imports this module.
    """
    reference = (workspaceStageReference(source) if source["Kind"] == "workspace"
                 else f'{source["Stage"]}/{source["Path"]}'.rstrip("/"))
    failures: List[str] = []
    started = time.time()

    # --- strategy 1: Snowpark -------------------------------------------- #
    try:
        session = Session
        if session is None:
            from snowflake.snowpark.context import get_active_session
            session = get_active_session()
            steps.append("snowpark: using the active session")
        else:
            steps.append("snowpark: using the session passed in")
        session.file.get(reference, tempDir)
        steps.append(f"snowpark: GET {reference} -> {tempDir}")
    except Exception as exc:                       # noqa: BLE001 - reported
        failures.append(f"snowpark: {type(exc).__name__}: {exc}")
        steps.append(f"snowpark FAILED: {type(exc).__name__}: {exc}")
    else:
        found = _firstFile(tempDir, Path(source["Path"]).name)
        if found:
            return found
        failures.append("snowpark: the GET reported success but wrote nothing")

    if time.time() - started > TimeoutSeconds:
        raise SourceNotFoundError(
            f"gave up after {TimeoutSeconds}s fetching {reference!r}. "
            f"Tried: {'; '.join(failures)}")

    # --- strategy 2: the connector, with a raw GET ------------------------ #
    try:
        import snowflake.connector                  # noqa: F401
        from snowflake.connector import connect
        connection = connect(
            account=os.environ.get("SNOWFLAKE_ACCOUNT", ""),
            user=os.environ.get("SNOWFLAKE_USER", ""),
            authenticator=os.environ.get("SNOWFLAKE_AUTHENTICATOR",
                                         "externalbrowser"),
            warehouse=os.environ.get("SNOWFLAKE_WAREHOUSE", ""),
            database=os.environ.get("SNOWFLAKE_DATABASE", ""),
            schema=os.environ.get("SNOWFLAKE_SCHEMA", ""))
        steps.append("connector: opened a connection from the environment")
        try:
            # file:// needs forward slashes on every platform, including
            # Windows, or the driver reads the backslashes as escapes.
            destination = Path(tempDir).as_posix()
            cursor = connection.cursor()
            cursor.execute(f"GET '{reference}' 'file://{destination}'")
            steps.append(f"connector: GET {reference} -> file://{destination}")
        finally:
            connection.close()
    except Exception as exc:                        # noqa: BLE001 - reported
        failures.append(f"connector: {type(exc).__name__}: {exc}")
        steps.append(f"connector FAILED: {type(exc).__name__}: {exc}")
    else:
        found = _firstFile(tempDir, Path(source["Path"]).name)
        if found:
            return found
        failures.append("connector: the GET reported success but wrote nothing")

    raise SourceNotFoundError(
        f"could not fetch {reference!r}. Every strategy failed:\n  - "
        + "\n  - ".join(failures)
        + "\n  CHECK: the path is spelled exactly as Snowflake shows it; "
          "your role has READ on the stage or workspace; and a warehouse is "
          "running. Set SNOWFLAKE_ACCOUNT / SNOWFLAKE_USER / "
          "SNOWFLAKE_WAREHOUSE when no active session exists.")


def _firstFile(directory: str, preferredName: str = "") -> Optional[Path]:
    """Return a downloaded file from `directory`, preferring `preferredName`.

    Parameters: directory (str); preferredName (str) - the expected filename.
    Returns: Path, or None when the directory holds no file.

    Snowflake gzips on the way out often enough that the name on disk is not
    always the name asked for, so the preferred name is a preference and not
    a requirement.
    """
    root = Path(directory)
    if not root.is_dir():
        return None
    files = [item for item in sorted(root.iterdir()) if item.is_file()]
    if not files:
        return None
    for item in files:
        if item.name == preferredName:
            return item
    return files[0]


def writeTarget(sourcePath: Path, TARGET_LOCATION: str, Overwrite: bool,
                steps: List[str]) -> int:
    """Copy `sourcePath` onto the target, classifying any failure.

    Parameters:
        sourcePath (Path)     - the file to copy.
        TARGET_LOCATION (str) - the destination.
        Overwrite (bool)      - replace an existing target.
        steps (list)          - appended to.
    Returns: int bytes written.
    Raises:
        TargetWriteError - with the reason spelled out: access, missing
        share, read-only file, or a full disk.
    """
    target = Path(TARGET_LOCATION)
    if target.exists() and not Overwrite:
        raise TargetWriteError(
            f"{target} already exists and Overwrite=False. Pass "
            f"Overwrite=True to replace it, or choose another name.")
    try:
        ensureParent(str(target))
        steps.append(f"target folder ready: {target.parent}")
    except OSError as exc:
        raise TargetWriteError(
            f"cannot create the folder {target.parent}: {exc.strerror or exc}.\n"
            f"  CHECK: the share exists and is reachable from this machine "
            f"(try opening {target.parent} in File Explorer), and that your "
            f"account may create folders there.") from exc

    try:
        shutil.copyfile(sourcePath, target)
    except PermissionError as exc:
        raise TargetWriteError(
            f"access denied writing {target}: {exc.strerror or exc}.\n"
            f"  CHECK: the file is not open in Excel, it is not marked "
            f"read-only, and your account has WRITE on {target.parent}. "
            f"Insufficient access to a CGS share is the usual cause."
        ) from exc
    except OSError as exc:
        hint = ("the destination disk or share is full"
                if getattr(exc, "errno", None) == 28
                else "the share may be unreachable or the path too long")
        raise TargetWriteError(
            f"write failed for {target}: {exc.strerror or exc}.\n"
            f"  LIKELY: {hint}.") from exc

    written = target.stat().st_size
    expected = sourcePath.stat().st_size
    if written != expected:
        raise TargetWriteError(
            f"wrote {written} bytes to {target} but the source is "
            f"{expected} bytes. The copy is incomplete -- treat the target "
            f"as bad and run again.")
    steps.append(f"copied {written} byte(s) -> {target}")
    return written


def copySnowflakeFile2CGS(SOURCE_LOCATION: str, TARGET_LOCATION: str,
                          Session: Any = None, Overwrite: bool = True,
                          DryRun: bool = False, TempDir: Optional[str] = None,
                          TimeoutSeconds: int = 300) -> Dict[str, Any]:
    """Copy a file from a Snowflake workspace or stage to a local/UNC path.

    Parameters:
        SOURCE_LOCATION (str)  - REQUIRED. A snow://workspace/... URI, a
                                 @DB.SCHEMA.STAGE/path reference, or an
                                 ordinary file path.
        TARGET_LOCATION (str)  - REQUIRED destination file path.
        Session                - an existing Snowpark Session; optional.
        Overwrite (bool)       - replace an existing target; default True.
        DryRun (bool)          - resolve and check access, write nothing.
        TempDir (str)          - staging directory for the download.
        TimeoutSeconds (int)   - cap on the Snowflake fetch; default 300.
    Returns:
        dict with SourceLocation, TargetLocation, BytesCopied, Strategy,
        Runtime, Steps and Copied.
    Raises:
        ValueError              - a required parameter is missing.
        UnreachableTargetError  - the target cannot be reached from here.
        SourceNotFoundError     - the source could not be fetched.
        TargetWriteError        - the write failed; the message says why.

    Use in claims processing:
        Land a claims extract produced in Snowflake onto the CGS share the
        downstream SAS and Excel jobs read from, on a schedule, instead of
        somebody downloading it and dragging it across every morning.
    """
    if not str(TARGET_LOCATION or "").strip():
        raise ValueError("required parameter 'TARGET_LOCATION' is missing or empty")

    steps: List[str] = []
    runtime = describeRuntime()
    logInfo(f"copySnowflakeFile2CGS {__version__} starting on "
            f"{runtime['Platform']}"
            f"{' (Snowflake compute)' if runtime['OnSnowflake'] else ''}")
    steps.append(f"runtime: {runtime['Platform']}; "
                 f"OnSnowflake={runtime['OnSnowflake']}; "
                 f"{'; '.join(runtime['Reasons'])}")

    source = classifySource(SOURCE_LOCATION)
    steps.append(f"source classified as {source['Kind']}: {source['Raw']}")
    logInfo(f"source is a {source['Kind']} location")

    # Fail on an unreachable target BEFORE downloading anything: there is no
    # point pulling a large extract out of Snowflake to then discard it.
    checkTargetReachable(TARGET_LOCATION, runtime)
    steps.append("target reachability: OK from this host")

    temporary: Optional[tempfile.TemporaryDirectory] = None
    try:
        if source["Kind"] == "file":
            resolved = Path(source["Path"])
            strategy = "file"
            if not resolved.is_file():
                raise SourceNotFoundError(
                    f"source file not found: {resolved}.\n"
                    f"  CHECK: the path is spelled correctly and is visible "
                    f"from this machine.")
            steps.append(f"source found on this filesystem: {resolved}")
        else:
            resolved = None
            strategy = ""
            if source["Kind"] == "workspace":
                resolved = resolveWorkspaceLocally(source, steps)
                if resolved is not None:
                    strategy = "workspace-mount"
            if resolved is None:
                if TempDir:
                    stagingDir = TempDir
                    Path(stagingDir).mkdir(parents=True, exist_ok=True)
                else:
                    temporary = tempfile.TemporaryDirectory(prefix="cgs_snow_")
                    stagingDir = temporary.name
                steps.append(f"staging directory: {stagingDir}")
                resolved = fetchFromSnowflake(source, stagingDir, Session,
                                              steps, TimeoutSeconds)
                strategy = "snowflake-get"
                logInfo(f"fetched {resolved.name} "
                        f"({resolved.stat().st_size} bytes)")

        if DryRun:
            probe = _probeTarget(TARGET_LOCATION, steps)
            logInfo(f"DRY RUN: would copy {resolved} -> {TARGET_LOCATION}")
            return {"SourceLocation": source["Raw"],
                    "TargetLocation": str(TARGET_LOCATION),
                    "BytesCopied": 0, "Strategy": strategy + "+dryrun",
                    "Runtime": runtime, "Steps": steps, "Copied": False,
                    "TargetWritable": probe}

        written = writeTarget(resolved, TARGET_LOCATION, Overwrite, steps)
    finally:
        if temporary is not None:
            temporary.cleanup()
            steps.append("staging directory removed")

    logInfo(f"copied {written} byte(s) to {TARGET_LOCATION}")
    return {"SourceLocation": source["Raw"],
            "TargetLocation": str(TARGET_LOCATION),
            "BytesCopied": written, "Strategy": strategy,
            "Runtime": runtime, "Steps": steps, "Copied": True}


def _probeTarget(TARGET_LOCATION: str, steps: List[str]) -> bool:
    """Check the target could be written, without writing the real file.

    Parameters: TARGET_LOCATION (str); steps (list) - appended to.
    Returns: bool - True when a test file could be created and removed.
    """
    target = Path(TARGET_LOCATION)
    try:
        ensureParent(str(target))
        probe = target.with_name(target.name + ".cgs_probe")
        probe.write_bytes(b"")
        probe.unlink()
    except OSError as exc:
        steps.append(f"dry run: target NOT writable: {exc.strerror or exc}")
        logWarn(f"target is not writable: {exc.strerror or exc}")
        return False
    steps.append("dry run: target is writable")
    return True


def buildArgParser() -> argparse.ArgumentParser:
    """Build the command-line parser. Parameters: none."""
    parser = argparse.ArgumentParser(
        prog="copySnowflakeFile2CGS.py",
        description="Copy a file from a Snowflake workspace or stage to a "
                    "local or UNC path.")
    parser.add_argument("--source-location", required=True)
    parser.add_argument("--target-location", required=True)
    parser.add_argument("--no-overwrite", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--temp-dir", default=None)
    parser.add_argument("--timeout-seconds", type=int, default=300)
    parser.add_argument("--version", action="version",
                        version=f"copySnowflakeFile2CGS {__version__}")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    """CLI entry point. Parameters: argv (sequence[str]). Returns: exit code."""
    args = buildArgParser().parse_args(argv)
    try:
        result = copySnowflakeFile2CGS(
            SOURCE_LOCATION=args.source_location,
            TARGET_LOCATION=args.target_location,
            Overwrite=not args.no_overwrite,
            DryRun=args.dry_run,
            TempDir=args.temp_dir,
            TimeoutSeconds=args.timeout_seconds)
    except ValueError as exc:
        logError(str(exc))
        return EXIT_CONFIG_ERROR
    except SnowflakeCopyError as exc:
        # Every line of the explanation, not just the first.
        for line in str(exc).splitlines():
            logError(line)
        return EXIT_IO_ERROR
    for step in result["Steps"]:
        logInfo(f"  step: {step}")
    logInfo(f"done; {result['BytesCopied']} byte(s) via {result['Strategy']}")
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
