"""
=====================================================================
  Program Name  : sendEmail.py
  Author        : Manuel Figallo
  Purpose       : Send email alerts over SMTP, used to notify operators
                  when a pipeline or scan completes or fails.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies:
    Standard library only (smtplib, email).

  Description:
    Mirrors the PowerShell Send-EmailAlert pattern. Multiple recipients are
    accepted in `To` as a list or a ';'-delimited string. SmtpServer and
    Port default to smtp.example.com:25 (override in .env with SMTP_SERVER /
    SMTP_PORT); every other parameter is required.

  Input Parameters (required first):
    To          (REQUIRED) - recipient(s); list or ';'-delimited string.
    From        (REQUIRED) - sender address.
    Subject     (REQUIRED) - subject line.
    Body        (REQUIRED) - message body.
    SmtpServer  (optional, default smtp.example.com or SMTP_SERVER)
    Port        (optional, default 25 or SMTP_PORT)
=====================================================================
"""

from __future__ import annotations

import smtplib
import sys
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Dict, List, Optional

_HERE = Path(__file__).resolve()
if str(_HERE.parent.parent.parent) not in sys.path:
    sys.path.insert(0, str(_HERE.parent.parent.parent))

from src.utils.config import getConfig          # noqa: E402
from src.utils.helpers import asList            # noqa: E402
from src.utils.logger import logError, logInfo  # noqa: E402

__version__ = "1.0beta"

DEFAULT_SMTP_SERVER = "smtp.example.com"
DEFAULT_SMTP_PORT = 25


def sendEmail(To: Any, From: str, Subject: str, Body: str,
              SmtpServer: Optional[str] = None, Port: Optional[int] = None,
              Html: bool = False, Attachments: Any = ()) -> Dict[str, Any]:
    """Send an email alert over SMTP.

    Parameters:
        To (list|str)     - REQUIRED recipient(s); ';'-delimited string allowed.
        From (str)        - REQUIRED sender address.
        Subject (str)     - REQUIRED subject line.
        Body (str)        - REQUIRED message body.
        SmtpServer (str)  - SMTP host; defaults to SMTP_SERVER in .env,
                            else smtp.example.com.
        Port (int)        - SMTP port; defaults to SMTP_PORT in .env, else 25.
        Html (bool)       - send the body as HTML instead of plain text.
        Attachments       - optional file path(s) to attach.
    Returns:
        dict with To (list), Subject, SmtpServer, Port and Sent (bool).
    Raises:
        ValueError            - a required parameter is missing.
        smtplib.SMTPException - the server rejected the message.
        OSError               - the server is unreachable.

    Use in claims processing:
        Notify the claims-operations mailbox when an overnight log scan or
        bulk attachment download finishes, including the row count and the
        output location, so nobody has to watch the job.
    """
    recipients: List[str] = asList(To)
    if not recipients:
        raise ValueError("required parameter 'To' is missing or empty")
    for name, value in (("From", From), ("Subject", Subject), ("Body", Body)):
        if value is None or str(value).strip() == "":
            raise ValueError(f"required parameter '{name}' is missing or empty")

    server = SmtpServer or getConfig("SMTP_SERVER", DEFAULT_SMTP_SERVER)
    port = int(Port or getConfig("SMTP_PORT", str(DEFAULT_SMTP_PORT)))

    message = EmailMessage()
    message["To"] = ", ".join(recipients)
    message["From"] = From
    message["Subject"] = Subject
    message.set_content(Body, subtype="html" if Html else "plain")

    for attachment in asList(Attachments):
        path = Path(attachment)
        if not path.is_file():
            logError(f"attachment not found, skipping: {path}")
            continue
        message.add_attachment(path.read_bytes(), maintype="application",
                               subtype="octet-stream", filename=path.name)

    logInfo(f"sending mail to {len(recipients)} recipient(s) via {server}:{port}")
    with smtplib.SMTP(server, port, timeout=30) as smtp:
        smtp.send_message(message)
    logInfo(f"sent: {Subject}")
    return {"To": recipients, "Subject": Subject, "SmtpServer": server,
            "Port": port, "Sent": True}
