import os
import smtplib
import ssl
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

try:
    import certifi
except ImportError:
    certifi = None

ALWAYS_INCLUDE_RECIPIENT = "it@cordialsausa.com"


def _get_ssl_context():
    if certifi:
        return ssl.create_default_context(cafile=certifi.where())
    return ssl.create_default_context()


def send_notification_email(subject, body, log_file_path=None, logger=None):
    smtp_server = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    sender_email = os.getenv("SENDER_EMAIL", "")
    sender_password = os.getenv("SENDER_PASSWORD", "")
    recipients_raw = os.getenv("EMAIL_RECIPIENTS", "")

    if not sender_email or not sender_password:
        if logger:
            logger.warning("No se envio correo: SENDER_EMAIL o SENDER_PASSWORD no configurados en .env")
        return False

    recipients = [r.strip() for r in recipients_raw.split(",") if r.strip()]
    if ALWAYS_INCLUDE_RECIPIENT not in recipients:
        recipients.append(ALWAYS_INCLUDE_RECIPIENT)

    if not recipients:
        if logger:
            logger.warning("No se envio correo: no hay destinatarios configurados")
        return False

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    html_body = f"""
    <html>
    <body style="font-family: Arial, sans-serif; font-size: 14px;">
        <h2 style="color: #333;">SAP Registration Info - Reporte de Ejecucion</h2>
        <p><strong>Fecha/Hora:</strong> {timestamp}</p>
        <hr>
        <pre style="background: #f4f4f4; padding: 15px; border-radius: 5px; white-space: pre-wrap; word-wrap: break-word;">{body}</pre>
    </body>
    </html>
    """

    msg.attach(MIMEText(html_body, "html", "utf-8"))

    if log_file_path and os.path.isfile(log_file_path):
        try:
            with open(log_file_path, "r", encoding="utf-8") as f:
                log_attachment = MIMEBase("text", "plain")
                log_attachment.set_payload(f.read().encode("utf-8"))
                encoders.encode_base64(log_attachment)
                log_filename = os.path.basename(log_file_path)
                log_attachment.add_header(
                    "Content-Disposition",
                    f"attachment; filename={log_filename}"
                )
                msg.attach(log_attachment)
        except Exception as e:
            if logger:
                logger.warning(f"No se pudo adjuntar el log: {e}")

    try:
        context = _get_ssl_context()
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls(context=context)
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipients, msg.as_string())
        if logger:
            logger.info(f"Correo enviado exitosamente a: {', '.join(recipients)}")
        return True
    except Exception as e:
        if logger:
            logger.error(f"Error al enviar correo: {e}")
        return False
