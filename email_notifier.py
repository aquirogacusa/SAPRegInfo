import os
import base64
import smtplib
import ssl
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formataddr

import requests

try:
    import certifi
except ImportError:
    certifi = None

ALWAYS_INCLUDE_RECIPIENT = "it@cordialsausa.com"
RESEND_API_URL = "https://api.resend.com/emails"
BREVO_API_URL = "https://api.brevo.com/v3/smtp/email"


def _get_ssl_context():
    if certifi:
        return ssl.create_default_context(cafile=certifi.where())
    return ssl.create_default_context()


def _build_html_body(body):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return f"""
    <html>
    <body style="font-family: Arial, sans-serif; font-size: 14px;">
        <h2 style="color: #333;">SAP Registration Info - Reporte de Ejecucion</h2>
        <p><strong>Fecha/Hora:</strong> {timestamp}</p>
        <hr>
        <pre style="background: #f4f4f4; padding: 15px; border-radius: 5px; white-space: pre-wrap; word-wrap: break-word;">{body}</pre>
    </body>
    </html>
    """


def _enviar_via_resend(api_key, from_addr, subject, body, recipients, log_file_path, logger):
    """Envio por API de Resend (HTTPS): sin SMTP, sin puertos y sin
    permisos por IP publica. Solo requiere la API key."""
    payload = {
        "from": from_addr,
        "to": recipients,
        "subject": subject,
        "html": _build_html_body(body),
    }

    # Nombre de visualizacion si se definio y el from no lo trae ya
    from_name = os.getenv("MAIL_FROM_NAME", "").strip()
    if from_name and "<" not in from_addr:
        payload["from"] = f"{from_name} <{from_addr}>"

    # Adjuntar el log (Resend acepta adjuntos en base64)
    if log_file_path and os.path.isfile(log_file_path):
        try:
            with open(log_file_path, "r", encoding="utf-8") as f:
                contenido = f.read().encode("utf-8")
            payload["attachments"] = [{
                "filename": os.path.basename(log_file_path),
                "content": base64.b64encode(contenido).decode("ascii"),
            }]
        except Exception as e:
            if logger:
                logger.warning(f"No se pudo adjuntar el log: {e}")

    try:
        resp = requests.post(
            RESEND_API_URL,
            json=payload,
            headers={"Authorization": f"Bearer {api_key}"},
            timeout=30,
        )
        if resp.status_code in (200, 201):
            if logger:
                logger.info(f"Correo enviado via Resend a: {', '.join(recipients)}")
            return True
        detalle = (resp.text or "")[:300]
        if logger:
            logger.error(f"Error al enviar via Resend (HTTP {resp.status_code}): {detalle}")
            if resp.status_code == 403 and "resend.dev" in from_addr:
                logger.warning(
                    "Recordatorio: sin dominio verificado, Resend solo entrega a la "
                    "propia cuenta con la que te registraste. Revisa EMAIL_RECIPIENTS."
                )
        return False
    except Exception as e:
        if logger:
            logger.error(f"Error al enviar via Resend: {e}")
        return False


def _enviar_via_brevo(api_key, from_addr, subject, body, recipients, log_file_path, logger):
    """Envio por API de Brevo (ex-Sendinblue): plan gratuito con 300
    correos/dia, admite adjuntos y multiples destinatarios con solo
    verificar la direccion remitente (sin DNS del dominio)."""
    payload = {
        "sender": {"email": from_addr},
        "to": [{"email": r} for r in recipients],
        "subject": subject,
        "htmlContent": _build_html_body(body),
    }

    # Nombre de visualizacion (Brevo lo soporta en el objeto sender)
    from_name = os.getenv("MAIL_FROM_NAME", "").strip()
    if from_name:
        payload["sender"]["name"] = from_name

    if log_file_path and os.path.isfile(log_file_path):
        try:
            with open(log_file_path, "r", encoding="utf-8", errors="replace") as f:
                contenido = f.read().encode("utf-8")
            payload["attachment"] = [{
                "name": os.path.basename(log_file_path),
                "content": base64.b64encode(contenido).decode("ascii"),
            }]
        except Exception as e:
            if logger:
                logger.warning(f"No se pudo adjuntar el log: {e}")

    try:
        resp = requests.post(
            BREVO_API_URL,
            json=payload,
            headers={"api-key": api_key, "content-type": "application/json"},
            timeout=30,
        )
        if resp.status_code in (200, 201):
            if logger:
                logger.info(f"Correo enviado via Brevo a: {', '.join(recipients)}")
            return True
        detalle = (resp.text or "")[:300]
        if logger:
            logger.error(f"Error al enviar via Brevo (HTTP {resp.status_code}): {detalle}")
            if resp.status_code == 401:
                logger.warning("Brevo devolvio 401: revisa que BREVO_API_KEY sea una key valida (v3).")
            elif resp.status_code == 400 and "sender" in detalle.lower():
                logger.warning("Brevo rechazo el remitente: verifica la direccion en Brevo -> Senders, Domains & Dedicated IPs.")
        return False
    except Exception as e:
        if logger:
            logger.error(f"Error al enviar via Brevo: {e}")
        return False


def _enviar_via_smtp(smtp_server, smtp_port, sender_email, sender_password, use_tls,
                     subject, body, recipients, log_file_path, logger):
    """Envio por SMTP clasico: modo relay sin credenciales (puerto 25) o
    modo con login (ej. Gmail con password de aplicacion)."""
    msg = MIMEMultipart()
    # Nombre de visualizacion (disfraz legitimo): la direccion real no cambia,
    # solo como la muestran los clientes de correo. formataddr cita y codifica
    # el nombre correctamente (acentos, comas, etc.).
    from_name = os.getenv("MAIL_FROM_NAME", "").strip()
    msg["From"] = formataddr((from_name, sender_email)) if from_name else sender_email
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject

    msg.attach(MIMEText(_build_html_body(body), "html", "utf-8"))

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
        if int(smtp_port) == 465:
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


def _enviar_via_teams(webhook_url, subject, body, log_file_path, logger):
    """Notifica a un canal de Teams via webhook. Funciona con el Incoming
    Webhook clasico y con el Workflow de Power Automate ("Post to a channel
    when a webhook request is received"): ambos aceptan POST con {"text": ...}.
    Sin correos, sin dominios, sin permisos por IP.
    Teams no permite adjuntar archivos via webhook, por lo que el detalle del
    log se incluye como cola del propio mensaje (el final es donde quedan el
    resumen y las estadisticas)."""
    texto = f"{subject}\n\n{body}"

    if log_file_path and os.path.isfile(log_file_path):
        try:
            with open(log_file_path, "r", encoding="utf-8", errors="replace") as f:
                lineas = f.readlines()
            cola = "".join(lineas[-120:])
            texto += (
                "\n\n--------------------------------------------------\n"
                f"DETALLE: ultimas lineas de {os.path.basename(log_file_path)}:\n"
                f"{cola}"
            )
        except Exception as e:
            if logger:
                logger.warning(f"No se pudo leer el log para el mensaje de Teams: {e}")

    if len(texto) > 20000:
        texto = texto[:20000] + "\n...(contenido truncado; el log completo queda en logs\\ del servidor)"

    try:
        resp = requests.post(webhook_url, json={"text": texto}, timeout=30)
        if resp.status_code in (200, 202):
            if logger:
                logger.info("Notificacion enviada al canal de Teams via webhook")
            return True
        if logger:
            logger.error(f"Error al enviar a Teams (HTTP {resp.status_code}): {(resp.text or '')[:300]}")
        return False
    except Exception as e:
        if logger:
            logger.error(f"Error al enviar a Teams via webhook: {e}")
        return False


def send_notification_email(subject, body, log_file_path=None, logger=None):
    """Notifica por TODOS los canales configurados (cada uno es backup de los
    demas): Teams (webhook), Resend (API), Brevo (API) y SMTP (relay/login).
    Retorna True si al menos un canal logro enviar."""
    teams_url = os.getenv("TEAMS_WEBHOOK_URL", "").strip()
    resend_key = os.getenv("RESEND_API_KEY", "").strip()
    brevo_key = os.getenv("BREVO_API_KEY", "").strip()
    smtp_server = os.getenv("SMTP_SERVER", "").strip()
    smtp_port = int(os.getenv("SMTP_PORT", "25"))
    sender_email = os.getenv("SENDER_EMAIL", "").strip()
    # SENDER_PASSWORD es OPCIONAL: solo para servidores SMTP que exigen login
    # (ej. Gmail con password de aplicacion). En relay corporativo se deja vacio.
    sender_password = os.getenv("SENDER_PASSWORD", "")
    use_tls = os.getenv("SMTP_USE_TLS", "").strip() in ("1", "true", "True")
    recipients_raw = os.getenv("EMAIL_RECIPIENTS", "")

    destinatarios_base = [r.strip() for r in recipients_raw.split(",") if r.strip()]

    def lista_con_it():
        """Copia de destinatarios con el destinatario corporativo forzado."""
        lista = list(destinatarios_base)
        if ALWAYS_INCLUDE_RECIPIENT not in lista:
            lista.append(ALWAYS_INCLUDE_RECIPIENT)
        return lista

    resultados = []
    canales_ok = []

    # Canal 1: Teams (instantaneo, multi-persona, sin adjuntos)
    if teams_url:
        ok = _enviar_via_teams(teams_url, subject, body, log_file_path, logger)
        resultados.append(ok)
        if ok:
            canales_ok.append("Teams")

    # Canal 2: Resend (API). Mientras el remitente sea el dominio de pruebas
    # (resend.dev), Resend solo entrega a la propia cuenta registrada: no
    # forzamos el destinatario corporativo porque rechazaria el envio completo.
    if resend_key:
        from_addr = os.getenv("RESEND_FROM", "onboarding@resend.dev").strip()
        dest_resend = list(destinatarios_base)
        if "resend.dev" not in from_addr:
            dest_resend = lista_con_it()
        if not dest_resend:
            if logger:
                logger.warning("Resend omitido: no hay destinatarios configurados (EMAIL_RECIPIENTS)")
            resultados.append(False)
        else:
            ok = _enviar_via_resend(resend_key, from_addr, subject, body, dest_resend, log_file_path, logger)
            resultados.append(ok)
            if ok:
                canales_ok.append("Resend")

    # Canal 3: Brevo (plan gratuito, 300 correos/dia, con adjunto)
    if brevo_key:
        from_addr = os.getenv("BREVO_FROM", "").strip() or sender_email
        if not from_addr:
            if logger:
                logger.warning("Brevo omitido: define BREVO_FROM o SENDER_EMAIL (el remitente verificado en Brevo)")
            resultados.append(False)
        else:
            dest_brevo = lista_con_it()
            if not dest_brevo:
                if logger:
                    logger.warning("Brevo omitido: no hay destinatarios configurados (EMAIL_RECIPIENTS)")
                resultados.append(False)
            else:
                ok = _enviar_via_brevo(brevo_key, from_addr, subject, body, dest_brevo, log_file_path, logger)
                resultados.append(ok)
                if ok:
                    canales_ok.append("Brevo")

    # Canal 4: SMTP clasico (relay corporativo o con login)
    if smtp_server and sender_email:
        dest_smtp = lista_con_it()
        if not dest_smtp:
            if logger:
                logger.warning("SMTP omitido: no hay destinatarios configurados (EMAIL_RECIPIENTS)")
            resultados.append(False)
        else:
            ok = _enviar_via_smtp(smtp_server, smtp_port, sender_email, sender_password,
                                  use_tls, subject, body, dest_smtp, log_file_path, logger)
            resultados.append(ok)
            if ok:
                canales_ok.append("SMTP")

    if not resultados:
        if logger:
            logger.warning("No se envio notificacion: no hay canales configurados (TEAMS_WEBHOOK_URL, RESEND_API_KEY, BREVO_API_KEY o SMTP_SERVER/SENDER_EMAIL)")
        return False

    if logger:
        if canales_ok:
            logger.info(f"Canales de notificacion exitosos: {', '.join(canales_ok)}")
        else:
            logger.warning("Ningun canal de notificacion pudo enviar (todos fallaron)")
    return any(resultados)
