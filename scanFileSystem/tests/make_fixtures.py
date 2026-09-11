"""Generate the synthetic fixture tree used by the self-tests.

The default UNC roots in scanFileSystem.py are not reachable in dev, so the
tests run against this tree instead. Regenerating is idempotent.

Hand-summed expectations (asserted by test_scanFileSystem.py):

  jobA.log   3 steps  real 0.05 + 1.20 + 2.00 = 3.25   cpu 0.03 + 0.90 + 1.50 = 2.43
  jobB.log   2 steps  real 63.05 (1:03.05) + 0.10 = 63.15
                      cpu  60.00 (1:00.00) + 0.05 = 60.05
  nested/jobC.log  1 step  real 0.40  cpu 0.20
"""

from __future__ import annotations

import os
import time
from pathlib import Path

FIXTURES = Path(__file__).resolve().parent / "fixtures" / "logs"


def _sas_step(label: str, real: str, cpu: str) -> str:
    return (
        f"NOTE: {label} used (Total process time):\n"
        f"      real time           {real}\n"
        f"      cpu time            {cpu}\n"
    )


def write(path: Path, text: str, *, mtime: str | None = None) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    if mtime:
        stamp = time.mktime(time.strptime(mtime, "%Y-%m-%d"))
        os.utime(path, (stamp, stamp))
    return path


def build() -> Path:
    root = FIXTURES
    root.mkdir(parents=True, exist_ok=True)

    # --- jobA.log: three steps, plain seconds, 1 ERROR + 2 WARNINGs ---
    write(root / "jobA.log",
          "1    data work.a; set sashelp.class; run;\n"
          + _sas_step("DATA statement", "0.05 seconds", "0.03 seconds")
          + "WARNING: Variable height is uninitialized.\n"
          + "2    proc means data=work.a; run;\n"
          + _sas_step("PROCEDURE MEANS", "1.20 seconds", "0.90 seconds")
          + "ERROR: File WORK.MISSING.DATA does not exist.\n"
          + "WARNING: The data set WORK.B may be incomplete.\n"
          + "3    proc sort data=work.a; by name; run;\n"
          + _sas_step("PROCEDURE SORT", "2.00 seconds", "1.50 seconds")
          + "NOTE: The SAS System used the .accdb bridge for lookup.\n",
          mtime="2026-03-15")

    # --- jobB.log: clock-format durations (mm:ss and h:mm:ss) ---
    write(root / "jobB.log",
          "1    proc sql; create table big as select * from huge; quit;\n"
          + _sas_step("PROCEDURE SQL", "1:03.05", "1:00.00")
          + "2    data _null_; run;\n"
          + _sas_step("DATA statement", "0.10 seconds", "0.05 seconds"),
          mtime="2026-05-20")

    # --- nested subfolder ---
    write(root / "nested" / "jobC.log",
          "1    proc print data=sashelp.class; run;\n"
          + _sas_step("PROCEDURE PRINT", "0.40 seconds", "0.20 seconds"),
          mtime="2026-02-01")

    # --- plain .txt with Access database references (keyword sweep) ---
    write(root / "notes.txt",
          "Migration notes\n"
          "The legacy tracker lives in S:/shared/claims.accdb today.\n"
          "A second extract still reads Archive.mdb every Monday.\n"
          "Both should move to Snowflake in Q3.\n"
          "Contact the DME team before touching claims.accdb again.\n",
          mtime="2026-04-10")

    # --- excluded folders (only excluded when folder_exclusion_list is set) ---
    write(root / "Old" / "legacy.log",
          "1    data old; run;\n" + _sas_step("DATA statement", "9.99 seconds", "9.99 seconds"),
          mtime="2026-01-05")
    write(root / "Test" / "scratch.log",
          "1    data t; run;\n" + _sas_step("DATA statement", "8.88 seconds", "8.88 seconds"),
          mtime="2026-01-06")
    # near-miss sibling: excluding "Old" must NOT drop this
    write(root / "Older" / "keepme.log",
          "1    data keep; run;\n" + _sas_step("DATA statement", "0.11 seconds", "0.11 seconds"),
          mtime="2026-01-07")

    # --- date-range fixtures ---
    write(root / "dated_2025.log", "NOTE: nothing to see here.\n", mtime="2025-06-15")
    write(root / "dated_2026H1.log", "NOTE: first half.\n", mtime="2026-04-01")
    write(root / "dated_2026H2.log", "NOTE: second half.\n", mtime="2026-11-30")

    # --- a .sas program (exercises the default extension list) ---
    write(root / "PGM_report.sas",
          "%let dsn=claims;\nproc print data=&dsn; run;\n", mtime="2026-03-01")

    # --- malformed: invalid UTF-8 bytes (must fall back to latin-1, stay OK) ---
    bad = root / "malformed_binary.log"
    bad.parent.mkdir(parents=True, exist_ok=True)
    bad.write_bytes(b"NOTE: header\n\xff\xfe\x00\x81 broken bytes \xc3\x28\nERROR: bad\n")
    os.utime(bad, (time.mktime(time.strptime("2026-03-20", "%Y-%m-%d")),) * 2)

    # --- malformed: broken symlink (must yield non-OK parse_status, not abort) ---
    broken = root / "broken_link.log"
    if broken.is_symlink() or broken.exists():
        broken.unlink()
    try:
        broken.symlink_to(root / "does_not_exist_target.log")
    except (OSError, NotImplementedError):
        pass    # symlinks unavailable (e.g. Windows without privileges) -- skip

    # --- ignored extension (must never appear in output) ---
    write(root / "ignore_me.dat", "binary-ish payload\n", mtime="2026-03-01")

    return root


if __name__ == "__main__":
    print(build())
