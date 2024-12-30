# sendemail.py
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
import logging
from typing import Optional
import ssl

def setup_logging() -> logging.Logger:
    """Configure logging for email operations."""
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    
    if not logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    
    return logger

def send_email(
    subject: str,
    body: str,
    to: str,
    from_email: str,
    password: str,
    smtp_server: str,
    smtp_port: int,
    attachment: Optional[str] = None
) -> None:
    logger = setup_logging()
    
    try:
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        if attachment:
            attachment_path = Path(attachment)
            if not attachment_path.exists():
                raise FileNotFoundError(f"Attachment not found: {attachment}")

            with open(attachment_path, "rb") as file:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename="{attachment_path.name}"'
                )
                msg.attach(part)

        # Gunakan SMTP_SSL untuk port 465
        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.login(from_email, password)
            server.send_message(msg)
            
        logger.info(f"Email successfully sent to {to}")
        
    except Exception as e:
        logger.error(f"Failed to send email to {to}: {e}")
        raise