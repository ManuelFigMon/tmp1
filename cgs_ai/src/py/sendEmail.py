"""
=====================================================================
  Program Name  : sendEmail.py
  Author        : Manuel Figallo
  Purpose       : Send email alerts over SMTP, used to notify operators
                  when a pipeline or scan completes or fails.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-09-06

  Dependencies:
    Standard library only (smtplib, email).

  Description:
    Mirrors the PowerShell twin. Multiple recipients are accepted in `To`
    as a list or a ';'-delimited string. SmtpServer and Port default to
    smtp.example.com:25 (override in .env with SMTP_SERVER / SMTP_PORT);
    every other parameter is required.

  HOW TO FORMAT AN EMAIL
    1. ADDRESSES. Give every address a real name, so the recipient sees
       "Manuel Figallo" in their inbox instead of a raw mailbox:
           From="Manuel Figallo <manuel.figallo@cgsadmin.com>"
           To="Al Cordoba <al.cordoba@cgsadmin.com>"
       The form is  Display Name <address>. Separate several recipients
       with a SEMICOLON, never a comma -- a comma is legal INSIDE a
       display name ("Cordoba, Al"), so splitting on it would tear the
       address in half.

    2. GREETING. corporateBody() reads the display name out of `To` and
       opens with "Dear Al". With no display name it falls back to the
       part of the address before the '@', so an address-only call still
       produces a sentence rather than a blank.

    3. PLAIN TEXT OR HTML. Set Html=True to send HTML. This matters more
       than it looks: alignment, bold and spacing DO NOT EXIST in a plain
       text mail. Anything centered, indented or coloured needs Html=True.

    4. LAYOUT. Keep one idea per paragraph and put the deliverable's
       location on its own line -- an operator reading on a phone should
       be able to find the path without scrolling sideways.

    5. URGENCY. Urgent=True marks the message high priority AND sets the
       Outlook message flag, so it arrives with a red flag. Use it for
       something that needs action tonight, not for every run: a mailbox
       where everything is urgent has nothing urgent in it.

    6. SUBJECT. Say the outcome, not the mechanism -- "Issue Log of DB
       Tables - scan complete" beats "SAS job finished".

  Input Parameters (required first):
    To          (REQUIRED) - recipient(s); list or ';'-delimited string,
                             each optionally "Display Name <address>".
    From        (REQUIRED) - sender address, optionally with a display name.
    Subject     (REQUIRED) - subject line.
    Body        (REQUIRED) - message body; HTML when Html=True.
    SmtpServer  (optional, default smtp.example.com or SMTP_SERVER)
    Port        (optional, default 25 or SMTP_PORT)
    Html        (optional, default False) - send the body as HTML.
    Urgent      (optional, default False) - flag the message as urgent.
    Attachments (optional) - file path(s) to attach.
=====================================================================
"""

from __future__ import annotations

import smtplib
import sys
from email.message import EmailMessage
from email.utils import formataddr, getaddresses, parseaddr
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

_HERE = Path(__file__).resolve()
if str(_HERE.parent.parent.parent) not in sys.path:
    sys.path.insert(0, str(_HERE.parent.parent.parent))

from src.utils.config import getConfig          # noqa: E402
from src.utils.helpers import asList            # noqa: E402
from src.utils.logger import logError, logInfo  # noqa: E402

__version__ = "1.0beta"

DEFAULT_SMTP_SERVER = "smtp.example.com"
DEFAULT_SMTP_PORT = 25

#: Outlook shows this text as a red flag banner above the message.
DEFAULT_URGENT_FLAG = "Follow up"

#: Signature block for corporateBody(). Pass your own to override.
DEFAULT_SIGNATURE = (
    "Manuel A. Figallo | Statistical Programmer IV and Analyst | CGS",
    "26 Century Blvd., Ste NT600",
    "Nashville, TN 37214-3685",
    "email: manuel.figallo@cgsadmin.com",
)

#: The three lines that are centered at the foot of a corporate message.
DEFAULT_CONFIDENTIALITY = (
    "Confidential, unpublished property of CGS Administrators, LLC. "
    "Do not duplicate or distribute.",
    "Use and distribution limited solely to authorized personnel.",
    "© 2026 Copyright, CGS Administrators, LLC.",
)

FONT_STACK = "Calibri, 'Segoe UI', Arial, sans-serif"


def escapeHtml(value: Any) -> str:
    """Escape a value for HTML text content.

    Parameters: value (any) - None becomes an empty string.
    Returns: str with & < > escaped, so a path or a name cannot break the
             markup or be read as a tag.
    """
    if value is None:
        return ""
    return (str(value).replace("&", "&amp;")
            .replace("<", "&lt;").replace(">", "&gt;"))


def recipientFirstName(To: Any) -> str:
    """Work out what to call the first recipient.

    Parameters: To (list|str) - recipient(s), "Display Name <address>" or bare.
    Returns: str the given name, e.g. "Al" from "Al Cordoba <al.cordoba@...>".
             Falls back to the mailbox before the '@', title-cased, so an
             address-only call still greets somebody by name.
    """
    entries = asList(To)
    if not entries:
        return "colleague"
    name, address = parseaddr(str(entries[0]))
    if name.strip():
        # "Cordoba, Al" is written surname-first; the given name is last.
        if "," in name:
            return name.split(",")[-1].strip().split()[0]
        return name.strip().split()[0]
    mailbox = address.split("@")[0]
    return (mailbox.split(".")[0] or "colleague").title()


def corporateBody(To: Any, Message: Any, ReportPath: str = "",
                  Closing: str = "Regards,",
                  Signature: Sequence[str] = DEFAULT_SIGNATURE,
                  Confidentiality: Sequence[str] = DEFAULT_CONFIDENTIALITY
                  ) -> str:
    """Build an HTML body in the corporate house style.

    Parameters:
        To (list|str)          - recipients; the first one's display name
                                 becomes the greeting.
        Message (str|sequence) - one paragraph, or a sequence of paragraphs.
        ReportPath (str)       - where the output landed; omitted when blank.
        Closing (str)          - sign-off line above the signature.
        Signature (sequence)   - signature lines.
        Confidentiality (seq)  - the notice, CENTERED at the foot.
    Returns:
        str of HTML. Send it with Html=True -- as plain text the markup
        would be shown literally and nothing would be centered.

    Use in claims processing:
        Give an overnight scan's completion notice the same look as the rest
        of the department's correspondence, without retyping the footer.
    """
    paragraphs = [Message] if isinstance(Message, str) else list(Message or [])
    body = [f'<div style="font-family:{FONT_STACK};font-size:11pt;color:#1F2937;">',
            f"<p>Dear {escapeHtml(recipientFirstName(To))},</p>"]
    for paragraph in paragraphs:
        body.append(f"<p>{escapeHtml(paragraph)}</p>")
    if ReportPath:
        body.append("<p>You can review the detailed report here:</p>")
        body.append(f'<p style="font-family:Consolas,monospace;font-size:10pt;">'
                    f"{escapeHtml(ReportPath)}</p>")
    body.append(f"<p>{escapeHtml(Closing)}</p>")
    body.append('<p style="margin-bottom:18px;">'
                + "<br>".join(escapeHtml(line) for line in Signature) + "</p>")
    # The notice is centered, which is the one thing plain text cannot do.
    body.append('<hr style="border:none;border-top:1px solid #D1D5DB;">')
    body.append('<div style="text-align:center;font-size:8pt;color:#6B7280;">'
                + "<br>".join(escapeHtml(line) for line in Confidentiality)
                + "</div>")
    body.append("</div>")
    return "\n".join(body)


def normalizeAddresses(entries: Sequence[str]) -> List[str]:
    """Tidy each "Display Name <address>" entry without losing the name.

    Parameters: entries (sequence) - raw address strings.
    Returns: list[str] re-encoded with formataddr, so a display name holding
             a comma is quoted and cannot be mistaken for two recipients.
    """
    tidied = []
    for entry in entries:
        name, address = parseaddr(str(entry))
        tidied.append(formataddr((name, address)) if name else address)
    return tidied


def sendEmail(To: Any, From: str, Subject: str, Body: str,
              SmtpServer: Optional[str] = None, Port: Optional[int] = None,
              Html: bool = False, Urgent: bool = False,
              Attachments: Any = (),
              UrgentFlag: str = DEFAULT_URGENT_FLAG) -> Dict[str, Any]:
    """Send an email alert over SMTP.

    Parameters:
        To (list|str)     - REQUIRED recipient(s); ';'-delimited string allowed,
                            each optionally "Display Name <address>".
        From (str)        - REQUIRED sender; "Manuel Figallo <manuel...>" shows
                            the name in the recipient's inbox.
        Subject (str)     - REQUIRED subject line.
        Body (str)        - REQUIRED message body.
        SmtpServer (str)  - SMTP host; defaults to SMTP_SERVER in .env,
                            else smtp.example.com.
        Port (int)        - SMTP port; defaults to SMTP_PORT in .env, else 25.
        Html (bool)       - send the body as HTML instead of plain text.
        Urgent (bool)     - mark the message urgent; default False.
        Attachments       - optional file path(s) to attach.
        UrgentFlag (str)  - the flag text Outlook shows; only used when Urgent.
    Returns:
        dict with To (list), Subject, SmtpServer, Port, Urgent and Sent.
    Raises:
        ValueError            - a required parameter is missing.
        smtplib.SMTPException - the server rejected the message.
        OSError               - the server is unreachable.

    Urgency uses two different mechanisms, because Outlook draws them
    differently: Importance/X-Priority produce the red exclamation mark and
    sort the message to the top, while X-Message-Flag is what actually puts
    the red FLAG on it. Both are set, so the message stands out whichever
    view the recipient uses. Clients other than Outlook honour Importance
    and ignore the flag.

    Use in claims processing:
        Notify the claims-operations mailbox when an overnight log scan or
        bulk attachment download finishes, including the row count and the
        output location, so nobody has to watch the job.
    """
    recipients: List[str] = normalizeAddresses(asList(To))
    if not recipients:
        raise ValueError("required parameter 'To' is missing or empty")
    for name, value in (("From", From), ("Subject", Subject), ("Body", Body)):
        if value is None or str(value).strip() == "":
            raise ValueError(f"required parameter '{name}' is missing or empty")

    server = SmtpServer or getConfig("SMTP_SERVER", DEFAULT_SMTP_SERVER)
    port = int(Port or getConfig("SMTP_PORT", str(DEFAULT_SMTP_PORT)))

    message = EmailMessage()
    message["To"] = ", ".join(recipients)
    message["From"] = normalizeAddresses([From])[0]
    message["Subject"] = Subject
    message.set_content(Body, subtype="html" if Html else "plain")

    if Urgent:
        message["X-Priority"] = "1 (Highest)"
        message["X-MSMail-Priority"] = "High"
        message["Importance"] = "High"
        message["X-Message-Flag"] = UrgentFlag

    for attachment in asList(Attachments):
        path = Path(attachment)
        if not path.is_file():
            logError(f"attachment not found, skipping: {path}")
            continue
        message.add_attachment(path.read_bytes(), maintype="application",
                               subtype="octet-stream", filename=path.name)

    logInfo(f"sending {'URGENT ' if Urgent else ''}mail to "
            f"{len(recipients)} recipient(s) via {server}:{port}")
    with smtplib.SMTP(server, port, timeout=30) as smtp:
        smtp.send_message(message)
    logInfo(f"sent: {Subject}")
    # Report the bare addresses the server was actually handed.
    delivered = [address for _, address in getaddresses(recipients)]
    return {"To": delivered, "Subject": Subject, "SmtpServer": server,
            "Port": port, "Urgent": bool(Urgent), "Sent": True}
