"""
=====================================================================
  Program Name  : build_vscode_tutorial_docx.py
  Author        : Manuel Figallo
  Purpose       : Generate "VS Studio for Python in the Azure Cloud in
                  3 Steps" -- the beginner tutorial that comes before the
                  cgs_ai training.
  Version       : 1.0beta
  Created       : 2026-09-11
  Last Modified : 2026-09-11

  Dependencies:
    STANDARD LIBRARY ONLY. OOXML helpers come from build_readme_docx.py.
    The ten screenshots live in docs/media and were taken from the source
    outline; nothing here regenerates them.

  Description:
    Three steps, one screenshot per action, every action on its own line.
    The source outline left five questions marked "???"; each is answered
    in full here, which is most of what this file adds.

    Written for somebody who has never opened VS Code. That drives the
    format: a Goal box at the top of each step, a Key Takeaway box at the
    bottom, numbered actions in between, and a caption under every picture
    saying what the reader should be looking at.

  Usage:
    python src/utils/build_vscode_tutorial_docx.py [output_path]
=====================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_readme_docx import (BLUE, NAVY, bullet, esc, image,  # noqa: E402
                               pageBreak, para, writeDocx)

__version__ = "1.0beta"

ROOT = Path(__file__).resolve().parent.parent.parent
MEDIA = ROOT / "docs" / "media"

BODY = 13          # points -- large, for a printed tutorial
CAPTION = 10
GREEN = "2E7D4F"
PANEL = "EEF3FB"
GOLD = "8A6D1F"

#: The folder every step refers to. <your-id> replaces one person's login.
FOLDER = r"C:\Users\<your-id>\OneDrive - bcbsscgov\code\python\tests"


def box(title: str, lines, fill: str = PANEL, accent: str = NAVY,
        monoFirst: bool = False, trailing: int = 160) -> str:
    """Build a single-cell callout table -- a Goal or Key Takeaway panel.

    Parameters:
        title (str)       - the label, e.g. "GOAL".
        lines (seq|str)   - body line(s).
        fill (str)        - background hex.
        accent (str)      - title colour, also the left edge.
        monoFirst (bool)  - render the first line as code, for a panel whose
                            opening line is something the reader will type.
        trailing (int)    - spacing after, in twips. Pass 0 when a page break
                            follows: the spacer would otherwise flow onto the
                            next page and the break would then leave it blank.
    Returns: the table XML plus trailing spacing.
    """
    if isinstance(lines, str):
        lines = [lines]
    runs = (f'<w:p><w:pPr><w:spacing w:after="40"/></w:pPr>'
            f'<w:r><w:rPr><w:b/><w:sz w:val="20"/><w:color w:val="{accent}"/>'
            f'</w:rPr><w:t xml:space="preserve">{esc(title)}</w:t></w:r></w:p>')
    for index, line in enumerate(lines):
        style = (f'<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>'
                 f'<w:b/><w:sz w:val="24"/><w:color w:val="{accent}"/>'
                 if monoFirst and index == 0 else f'<w:sz w:val="{BODY * 2}"/>')
        runs += (f'<w:p><w:pPr><w:spacing w:after="40"/></w:pPr>'
                 f'<w:r><w:rPr>{style}</w:rPr>'
                 f'<w:t xml:space="preserve">{esc(line)}</w:t></w:r></w:p>')
    return (
        '<w:tbl><w:tblPr><w:tblW w:w="10080" w:type="dxa"/>'
        '<w:tblBorders>'
        f'<w:left w:val="single" w:sz="18" w:space="0" w:color="{accent}"/>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="D6DEEB"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D6DEEB"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="D6DEEB"/>'
        '</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="10080"/></w:tblGrid>'
        '<w:tr><w:trPr><w:cantSplit/></w:trPr>'
        f'<w:tc><w:tcPr><w:tcW w:w="10080" w:type="dxa"/>'
        f'<w:shd w:val="clear" w:fill="{fill}"/>'
        '<w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="180" w:type="dxa"/>'
        '<w:bottom w:val="120" w:type="dxa"/><w:right w:w="180" w:type="dxa"/>'
        '</w:tcMar></w:tcPr>'
        f'{runs}</w:tc></w:tr></w:tbl>'
        + (para("", spaceAfter=trailing) if trailing else ""))


def stepBanner(number: int, title: str) -> str:
    """Build the full-width navy banner that opens each step."""
    return (
        '<w:tbl><w:tblPr><w:tblW w:w="10080" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="10080"/></w:tblGrid>'
        f'<w:tr><w:tc><w:tcPr><w:tcW w:w="10080" w:type="dxa"/>'
        f'<w:shd w:val="clear" w:fill="{NAVY}"/>'
        '<w:tcMar><w:top w:w="140" w:type="dxa"/><w:left w:w="200" w:type="dxa"/>'
        '<w:bottom w:w="140" w:type="dxa"/></w:tcMar></w:tcPr>'
        '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
        '<w:r><w:rPr><w:b/><w:sz w:val="34"/><w:color w:val="FFFFFF"/></w:rPr>'
        f'<w:t xml:space="preserve">STEP {number}   {esc(title)}</w:t>'
        '</w:r></w:p></w:tc></w:tr></w:tbl>' + para("", spaceAfter=140))


def action(number: str, text: str) -> str:
    """Build one numbered action line -- the thing the reader actually does.

    keepNext ties it to whatever follows -- the screenshot, or the command to
    type -- so an instruction is never stranded at the foot of a page with
    its picture overleaf.
    """
    return (f'<w:p><w:pPr><w:keepNext/><w:spacing w:before="120" w:after="60"/>'
            f'<w:ind w:left="400" w:hanging="400"/></w:pPr>'
            f'<w:r><w:rPr><w:b/><w:sz w:val="{BODY * 2}"/>'
            f'<w:color w:val="{BLUE}"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(number)}   </w:t></w:r>'
            f'<w:r><w:rPr><w:sz w:val="{BODY * 2}"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')


def leadPara(lead: str, rest: str, spaceAfter: int = 110) -> str:
    """One paragraph: a bold navy lead sentence, then the explanation.

    Parameters: lead (str), rest (str), spaceAfter (int, twips).
    Returns: the paragraph XML.

    Keeping both in ONE paragraph rather than two halves the vertical space
    a list of points costs, which is what lets the introduction and the
    prerequisites panel share a page.
    """
    return (f'<w:p><w:pPr><w:spacing w:after="{spaceAfter}"/></w:pPr>'
            f'<w:r><w:rPr><w:b/><w:sz w:val="{BODY * 2}"/>'
            f'<w:color w:val="{NAVY}"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(lead)}  </w:t></w:r>'
            f'<w:r><w:rPr><w:sz w:val="{BODY * 2}"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(rest)}</w:t></w:r></w:p>')


def note(text: str) -> str:
    """Build the small grey explanation that follows a command or a click.

    keepNext, because a note sits between an action and its screenshot:
    without it the chain breaks and the picture drifts to the next page.
    """
    return para(text, size=CAPTION + 1, color="6B7280", spaceAfter=120,
                keepNext=True)


def caption(text: str) -> str:
    """Build the grey line under a screenshot saying what to look at."""
    return para(text, size=CAPTION, color="6B7280", align="center",
                spaceAfter=200)


def code(text: str) -> str:
    """Build a monospace block for something the reader types verbatim."""
    return para(text, style="Code", size=11, mono=True, spaceAfter=140,
                keepNext=True)


#: (relationship id, file, caption) for every screenshot, in reading order.
SHOTS = [
    ("rId10", "image1.png",
     "The Windows search box. Type the name, then press Enter."),
    ("rId11", "image2.png",
     "The VS Code Welcome screen. 'New File...' is top-left, under Start."),
    ("rId12", "image3.png",
     "The Create File box. Note the folder path along the top: "
     "code > python > tests."),
    ("rId13", "image4.png",
     "hello_world in your OneDrive tests folder. The cloud tick means it is "
     "backed up."),
    ("rId14", "image5.png",
     "Two lines of Python. Line 1 is a comment; line 2 is the program."),
    ("rId15", "image6.png",
     "Save Workspace As... Note 'Save as type: Code Workspace'."),
    ("rId16", "image7.png",
     "Toggle Panel, top-right. The keyboard shortcut is Ctrl+J."),
    ("rId17", "image8.png",
     "The terminal, open at the bottom. Yours starts in your home folder, "
     "not in tests."),
    ("rId18", "image9.png",
     "The cd command typed in. The double quotes matter -- the path has "
     "spaces in it."),
    ("rId19", "image10.png",
     "Hello, World! Your first Python program has run."),
]
IMAGES = {relationshipId: str(MEDIA / name)
          for relationshipId, name, _ in SHOTS}


def shot(index: int, maxWidth: float = 5.1) -> str:
    """Emit screenshot `index` (0-based) followed by its caption.

    5.1 inches, not the full 6.0 the page allows: a full VS Code window is
    4:3, so at 6 inches it stands 4 inches tall and only one action fits on
    a page. At 5.1 two fit, and the screenshots are still legible in print.
    """
    relationshipId, name, text = SHOTS[index]
    return (image(relationshipId, str(MEDIA / name), maxWidthInches=maxWidth,
                  index=index + 1, keepNext=True)
            + caption(text))


def buildBody() -> str:
    """Assemble the tutorial body XML. Returns: <w:body> content."""
    parts = []

    # ===================== COVER / INTRODUCTION ===================== #
    parts.append(para("VS Studio for Python", style="Title", size=30,
                      spaceAfter=0))
    parts.append(para("in the Azure Cloud — in 3 Steps", style="Title",
                      size=30, spaceAfter=60))
    parts.append(para("A beginner's tutorial  ·  CGS  ·  no experience needed",
                      style="Subtitle", size=14, spaceAfter=200))

    parts.append(para("Introduction", style="Heading1", spaceAfter=120))
    parts.append(para(
        "By the end of this tutorial you will have written and run a Python "
        "program, saved in the Azure cloud, on your own machine. It takes "
        "about fifteen minutes, and assumes you have never programmed.",
        size=BODY, spaceAfter=160))

    parts.append(para("Why is Python so important?", style="Heading2",
                      spaceAfter=100))
    for lead, rest in [
        ("It is the language CMS uses.",
         "The Centers for Medicare and Medicaid Services use Python "
         "throughout, so sharing work with them means reading Python."),
        ("It is the common language of data science.",
         "Whatever the specialism, the examples and the answers online are "
         "written in Python."),
        ("It is already inside the tools you use.",
         "Excel and Power BI both run Python now, so it extends software "
         "you already have open."),
        ("It is free, and it is everywhere.",
         "No licence to request, and the skill moves with you between teams "
         "and jobs."),
    ]:
        parts.append(leadPara(lead, rest))

    parts.append(para("Why Visual Studio Code?", style="Heading2",
                      spaceAfter=100))
    parts.append(para(
        "Python is the language. Visual Studio Code — \"VS Code\" — is where "
        "you write it. You could use Notepad, but VS Code gives you four "
        "things Notepad cannot:", size=BODY, spaceAfter=100))
    for text in [
        "Write and run in one window — editor on top, terminal underneath.",
        "It understands Python: keywords are coloured and typos are "
        "underlined before you run anything.",
        "It opens your cloud folder directly, leaving the files in OneDrive.",
        "It is what the rest of the team uses, so instructions match your "
        "screen.",
    ]:
        parts.append(bullet(text))
    parts.append(para("", spaceAfter=80))

    parts.append(box(
        "WHAT YOU NEED BEFORE YOU START",
        ["A CGS Windows machine with Visual Studio Code installed.",
         "Your OneDrive account signed in — the cloud icon in the taskbar.",
         "That is all. Python itself is already on the machine."],
        fill="F2F7F2", accent=GREEN, trailing=0))

    parts.append(pageBreak())

    # ========================== STEP 1 ========================== #
    parts.append(stepBanner(1, "Prepare your folder"))
    parts.append(box("GOAL",
                     "Understand where VS Code keeps your work, and make the "
                     "one folder this tutorial uses."))

    parts.append(para("A note about storage", style="Heading2",
                      spaceAfter=100))
    parts.append(para(
        "Your code has to live somewhere. Save it to your laptop's hard disk "
        "and it exists in exactly one place. We are going to use OneDrive "
        "instead — the Azure cloud — for four reasons:",
        size=BODY, spaceAfter=120))
    for lead, rest in [
        ("It future-proofs the code base.",
         "Work saved in the cloud outlives the machine it was written on."),
        ("You get 100 GB, reachable anywhere.",
         "From any CGS machine you sign in to, without copying anything."),
        ("It is backed up, and it is reliable.",
         "Previous versions can be restored. A local folder has neither."),
        ("It is shareable.",
         "A colleague can be given the folder. No zipping, no emailing."),
    ]:
        parts.append(leadPara(lead, rest, spaceAfter=120))

    parts.append(action("1.1", "Open File Explorer and go to your OneDrive "
                               "folder."))
    parts.append(action("1.2", "Create this folder path inside it, one folder "
                               "at a time: code, then python, then tests."))
    parts.append(code(FOLDER))
    parts.append(para(
        "Replace <your-id> with your own Windows user name. The screenshots "
        "in this tutorial show 731o, which is the author's — yours will be "
        "different, and everything else will look the same.",
        size=CAPTION + 1, color="6B7280", spaceAfter=160))

    parts.append(box(
        "KEY TAKEAWAY",
        ["Your Python work lives in the Azure cloud, not on your laptop.",
         "Every file in this tutorial goes in the tests folder you just made."],
        fill="FBF7EC", accent=GOLD, trailing=0))

    parts.append(pageBreak())

    # ========================== STEP 2 ========================== #
    parts.append(stepBanner(2, "Learn two definitions and three rules"))
    parts.append(box("GOAL",
                     "Know what a workspace is, and know which parts of VS "
                     "Code you must not use at CGS."))

    parts.append(para("What is a workspace?", style="Heading2",
                      spaceAfter=100))
    parts.append(para(
        "A FILE is one program — hello_world.py. A WORKSPACE is a saved note "
        "of which folder you were working in, which files you had open, and "
        "how the window was arranged.", size=BODY, spaceAfter=120))
    parts.append(para(
        "Think of it as a bookmark for a whole project rather than one page. "
        "Open it tomorrow and VS Code puts everything back as you left it. A "
        "workspace is a small .code-workspace file; it does not contain your "
        "code, it only points at it.", size=BODY, spaceAfter=180))

    parts.append(para("Three things you must not use", style="Heading2",
                      spaceAfter=100))
    parts.append(para(
        "VS Code ships with features aimed at the public internet. CGS "
        "restricts them, because your code and the data it touches must not "
        "leave the organisation. You will see all three on screen — leave "
        "them alone:", size=BODY, spaceAfter=140))
    for name, rest in [
        ("Copilot",
         "The AI suggestion feature, and the \"Type copilot to use Copilot "
         "CLI\" message you will see in the terminal. Ignore it. It sends "
         "what you are writing to an outside service."),
        ("GitHub",
         "\"Clone Git Repository...\" on the Welcome screen, and the source "
         "control icon down the left edge. Do not publish or clone code. "
         "Internal code belongs on the network share, not on a public host."),
        ("\"Sign In\"",
         "The account icon at the bottom-left, and any prompt to sign in or "
         "turn on Settings Sync. Stay signed out. Signing in links this "
         "machine to an external account."),
    ]:
        parts.append(leadPara(name, rest, spaceAfter=120))
    parts.append(para(
        "If VS Code prompts you for any of the three, dismiss the prompt and "
        "carry on — nothing in this tutorial needs them. If you are unsure "
        "whether something is permitted, ask before you click: your team owns "
        "the current policy.", size=BODY, spaceAfter=160))

    parts.append(box(
        "KEY TAKEAWAY",
        ["A workspace remembers your project; a file holds your code.",
         "Copilot, GitHub and Sign In are off limits. You need none of them."],
        fill="FBF7EC", accent=GOLD, trailing=0))

    parts.append(pageBreak())

    # ========================== STEP 3 ========================== #
    parts.append(stepBanner(3, "Write and run your first program"))
    parts.append(box("GOAL",
                     "Launch VS Code, learn the screen, and run a Python "
                     "program that prints Hello, World!"))

    parts.append(action("3.1", "Find VS Code. In the Windows taskbar search "
                               "box, type: Visual Studio Code — then press "
                               "Enter."))
    parts.append(shot(0, maxWidth=3.6))

    parts.append(action("3.2", "The Welcome screen opens. Under Start, click "
                               "\"New File...\"."))
    parts.append(shot(1))

    parts.append(action("3.3", "Name the file hello_world.py and save it into "
                               "your tests folder. The .py on the end is what "
                               "makes it a Python program."))
    parts.append(shot(2))

    parts.append(action("3.4", "Check it arrived. Open the tests folder in "
                               "File Explorer — hello_world should be there."))
    parts.append(shot(3))

    parts.append(action("3.5", "Type these two lines into the file, then press "
                               "Ctrl+S to save."))
    parts.append(code('#This is my first python comment\n'
                      'print("Hello, World!")'))
    parts.append(note(
        "Line 1 starts with # , which makes it a COMMENT — a note for humans "
        "that Python ignores. Line 2 is the program: print puts whatever is "
        "inside the brackets on the screen."))
    parts.append(shot(4))

    parts.append(action("3.6", "Save a workspace: File → Save Workspace As... "
                               "Name it hello_world_workspace and save it in "
                               "the same tests folder."))
    parts.append(note(
        "Why bother? Because tomorrow you will want to carry on. Opening the "
        "workspace reopens this folder and this file, arranged as you left "
        "them — instead of navigating back through OneDrive each time. It "
        "costs one click now and saves several every day after."))
    parts.append(shot(5))

    parts.append(action("3.7", "Open the terminal. Click Toggle Panel at the "
                               "top-right, or press Ctrl+J."))
    parts.append(shot(6))

    parts.append(action("3.8", "The terminal appears along the bottom. This is "
                               "where you type commands and where your "
                               "program's output will show up."))
    parts.append(shot(7))

    parts.append(action("3.9", "Move the terminal into your folder with the cd "
                               "command. Type it exactly, including the double "
                               "quotes:"))
    parts.append(code(f'cd "{FOLDER}"'))
    parts.append(note(
        "cd means \"change directory\". The double quotes are REQUIRED here, "
        "because \"OneDrive - bcbsscgov\" has spaces in it; without them the "
        "command stops at the first space and fails."))
    parts.append(shot(8))

    parts.append(action("3.10", "Run it. Type this and press Enter:"))
    parts.append(code("python hello_world.py"))
    parts.append(shot(9))

    parts.append(box(
        "KEY TAKEAWAY",
        ["You created a Python file in the Azure cloud, wrote a comment and "
         "a program, saved a workspace, and ran it from the terminal.",
         "That is the whole loop: write, save, run. Everything else builds "
         "on it."],
        fill="FBF7EC", accent=GOLD))

    parts.append(para("If something goes wrong", style="Heading2",
                      spaceAfter=100))
    for message, fix in [
        ("python: command not found, or 'python' is not recognized",
         "Python is not on the path for this terminal. Try py hello_world.py "
         "instead, and tell your team if neither works."),
        ("can't open file ... No such file or directory",
         "The terminal is not in your tests folder. Repeat action 3.9, and "
         "check the quotes are there."),
        ("The terminal shows a message about Copilot",
         "Expected, and harmless. Ignore it and type your command — see "
         "Step 2."),
    ]:
        parts.append(para(message, size=CAPTION + 1, mono=True, color=NAVY,
                          spaceAfter=30))
        parts.append(para(fix, size=CAPTION + 1, spaceAfter=120))

    # ========================== CLOSING ========================== #
    parts.append(para("", spaceAfter=120))
    parts.append(para("What comes next", style="Heading1", spaceAfter=120))
    parts.append(para(
        "This tutorial looked small on purpose. What you actually built is "
        "the foundation every later session depends on: a cloud folder, a "
        "working editor, a terminal you can run code from, and a file that "
        "runs.", size=BODY, spaceAfter=140))
    parts.append(para(
        "That foundation is what the cgs_ai Python modules will be delivered "
        "onto. cgs_ai is our own Python package — it holds the work we all "
        "repeat, such as turning an extract into a formatted report and "
        "sending it on, as functions you call by name. Those modules are "
        "forthcoming, and when they arrive the instructions will begin exactly "
        "where this tutorial ends: open your workspace, open the terminal, "
        "run the file.", size=BODY, spaceAfter=140))
    parts.append(para(
        "Before then, practise the loop. Change the words inside the quotes "
        "and run it again. Add another print line. Getting comfortable with "
        "write, save, run is the only preparation needed.",
        size=BODY, spaceAfter=200))

    parts.append(box(
        "COMING SOON — cgs_ai",
        ["import cgs_ai",
         "The same three steps, with our own package doing the work. "
         "Keep this tutorial; the next one starts from it."],
        fill="F2F7F2", accent=GREEN, monoFirst=True))

    parts.append(para(
        "Questions, or a correction to this tutorial? Please contact "
        "Manuel Figallo. cgs_ai is in beta — feedback is welcome now.",
        size=BODY, spaceAfter=0))

    return "".join(parts)


if __name__ == "__main__":
    output = sys.argv[1] if len(sys.argv) > 1 else str(
        ROOT / "docs" / "VS_Studio_for_Python_in_the_Azure_Cloud_in_3_Steps.docx")
    print("wrote", writeDocx(buildBody(), output, margin=1080, images=IMAGES))
