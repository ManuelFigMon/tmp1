"""Static checks on the SAS wrappers.

There is no SAS interpreter available in CI, so these catch the macro-syntax
mistakes that are easy to make and expensive to debug in a live SAS session.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

SAS_DIR = Path(__file__).resolve().parent.parent / "sas"
SAS_FILES = sorted(SAS_DIR.glob("*.sas"))

# %str(...) / %nrstr(...) / %bquote(...) mask semicolons; strip them first.
MASKED = re.compile(r"%n?r?b?str\([^)]*\)", re.IGNORECASE)


def put_statements(path: Path):
    for number, line in enumerate(path.read_text().splitlines(), 1):
        stripped = line.strip()
        if stripped.lower().startswith("%put"):
            yield number, stripped


@pytest.mark.parametrize("path", SAS_FILES, ids=lambda p: p.name)
def test_put_message_text_has_no_unmasked_semicolon(path):
    """A ';' inside %put text ends the statement early.

    Everything after it becomes stray open code, which SAS rejects with
    ERROR 180-322 -- and because a bare `%else %put ...; trailing text;` puts
    that trailing text OUTSIDE the branch, it fires on every invocation.
    Wrap such messages in %str(...).
    """
    offenders = []
    for number, statement in put_statements(path):
        body = MASKED.sub("", statement[4:])
        semicolons = body.count(";")
        if semicolons > 1 or (semicolons == 1 and not body.rstrip().endswith(";")):
            offenders.append(f"{path.name}:{number}: {statement}")
    assert not offenders, (
        "unmasked ';' in %put text (wrap the message in %str(...)):\n  "
        + "\n  ".join(offenders))


@pytest.mark.parametrize("path", SAS_FILES, ids=lambda p: p.name)
def test_macro_definitions_are_balanced(path):
    text = path.read_text()
    opens = len(re.findall(r"(?m)^\s*%macro\b", text, re.IGNORECASE))
    closes = len(re.findall(r"(?m)^\s*%mend\b", text, re.IGNORECASE))
    assert opens == closes, f"{path.name}: {opens} %macro vs {closes} %mend"

    do_count = len(re.findall(r"%do\b", text, re.IGNORECASE))
    end_count = len(re.findall(r"%end\b", text, re.IGNORECASE))
    assert do_count == end_count, f"{path.name}: {do_count} %do vs {end_count} %end"


def test_both_wrappers_expose_the_same_parameters():
    """%scanFileSystem() and %scanFileSystemPS() must stay interchangeable."""
    def params(path: Path, macro: str):
        text = path.read_text()
        block = re.search(rf"%macro\s+{macro}\s*\((.*?)\);", text,
                          re.IGNORECASE | re.DOTALL)
        assert block, f"{macro} not found in {path.name}"
        body = re.sub(r"/\*.*?\*/", "", block.group(1), flags=re.DOTALL)
        return {m.strip().lower() for m in re.findall(r"(\w+)\s*=", body)}

    python_params = params(SAS_DIR / "Run_scanFileSystem_v1.sas", "scanFileSystem")
    ps_params = params(SAS_DIR / "Run_scanFileSystem_PS_v1.sas", "scanFileSystemPS")
    assert python_params == ps_params, (
        f"parameter lists diverged\n  only in Python wrapper: {python_params - ps_params}"
        f"\n  only in PS wrapper: {ps_params - python_params}")


def test_examples_file_has_no_live_macro_calls():
    """Every example must stay commented out so submitting the file is inert.

    Examples are wrapped in `%*` ... `;` macro comments, which span lines and
    end at the first semicolon -- so track that state rather than looking at
    each line in isolation.
    """
    path = SAS_DIR / "Examples_scanFileSystem_v1.sas"
    live, in_macro_comment = [], False
    for number, line in enumerate(path.read_text().splitlines(), 1):
        stripped = line.strip()
        if not in_macro_comment and stripped.startswith("%*"):
            in_macro_comment = ";" not in stripped[2:]
            continue
        if in_macro_comment:
            if ";" in stripped:
                in_macro_comment = False
            continue
        if re.match(r"^\s*%scanFileSystem(PS)?\s*\(", line):
            live.append(f"{number}: {stripped}")
    assert not live, "uncommented example call(s):\n  " + "\n  ".join(live)


def test_wrappers_do_not_pass_macro_quoted_text_to_systask():
    """v1.3.4: the command must NOT be assembled with %str(%") into a macro
    variable handed to SYSTASK -- macro-quoting characters could reach the
    SAS shell module and crash it (sasxshel access violation). The command
    is built in a DATA step via symget() and run from a .bat instead."""
    for name in ("Run_scanFileSystem_v1.sas", "Run_scanFileSystem_PS_v1.sas"):
        raw = (SAS_DIR / name).read_text()
        # Strip /* */ comments -- the header legitimately *describes* the old
        # approach; only live code matters here.
        code = re.sub(r"/\*.*?\*/", "", raw, flags=re.DOTALL)
        assert "%str(%\")" not in code, f"{name}: macro-quoted quote reintroduced"
        assert 'systask command """&_bat"""' in code, \
            f"{name}: must launch the generated .bat"
        assert "symget(" in code, f"{name}: command must be built via symget()"


def test_wrappers_refuse_to_run_under_noxcmd():
    """SYSTASK needs XCMD; without it the wrapper must abort cleanly rather
    than letting the shell module fail."""
    for name in ("Run_scanFileSystem_v1.sas", "Run_scanFileSystem_PS_v1.sas"):
        text = (SAS_DIR / name).read_text()
        assert "getoption(xcmd)" in text, f"{name}: missing XCMD guard"


def test_wrappers_expose_a_debug_switch():
    """debug=1 must build the command without executing it, so a crashing
    site can still inspect exactly what would have run."""
    for name in ("Run_scanFileSystem_v1.sas", "Run_scanFileSystem_PS_v1.sas"):
        text = (SAS_DIR / name).read_text()
        assert "debug" in text.lower() and "%return" in text, f"{name}: no debug path"
