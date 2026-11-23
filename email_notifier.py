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
    smtp_server = os.getenv("SMTP_SERVER", "")
    smtp_port = int(os.getenv("SMTP_PORT", "25"))
    sender_email = os.getenv("SENDER_EMAIL", "")
    # SENDER_PASSWORD es OPCIONAL: solo para servidores que exigen login
    # (ej. Gmail con password de aplicacion). En el relay corporativo
    # (autorizado por IP) se deja vacio y se envia sin autenticacion.
    sender_password = os.getenv("SENDER_PASSWORD", "")
    # Escape para relays que exigen STARTTLS incluso sin credenciales
    use_tls = os.getenv("SMTP_USE_TLS", "").strip() in ("1", "true", "True")
    recipients_raw = os.getenv("EMAIL_RECIPIENTS", "")

    if not smtp_server or not sender_email:
        if logger:
            logger.warning("No se envio correo: SMTP_SERVER o SENDER_EMAIL no configurados en .env")
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
        if smtp_port == 465:
            # Puerto SSL implicito (estilo Gmail SSL)
            with smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30, context=context) as server:
                if sender_password:
                    server.login(sender_email, sender_password)
                server.sendmail(sender_email, recipients, msg.as_string())
        else:
            with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as server:
                # Relay corporativo sin credenciales: envio plano en puerto 25.
                # TLS solo si hay password (Gmail) o se pide con SMTP_USE_TLS=1.
                if sender_password or use_tls:
                    try:
                        server.starttls(context=context)
                    except smtplib.SMTPNotSupportedError as tls_err:
                        if use_tls and not sender_password and logger:
                            logger.warning(f"El servidor no soporto STARTTLS, se envia plano: {tls_err}")
                if sender_password:
                    server.login(sender_email, sender_password)
                server.sendmail(sender_email, recipients, msg.as_string())
        if logger:
            logger.info(f"Correo enviado exitosamente a: {', '.join(recipients)}")
        return True
    except Exception as e:
        if logger:
            logger.error(f"Error al enviar correo: {e}")
        return False
