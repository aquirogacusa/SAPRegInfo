import subprocess
import sys
import os
import csv
import json
import traceback
from datetime import datetime

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from logger import setup_logger, get_latest_log_file
from email_notifier import send_notification_email


def run_script(script_name, logger):
    logger.info(f"{'=' * 60}")
    logger.info(f"Ejecutando: {script_name}")
    logger.info(f"{'=' * 60}")

    try:
        result = subprocess.run(
            [sys.executable, script_name],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=3600
        )

        if result.stdout.strip():
            for line in result.stdout.strip().splitlines():
                logger.info(line)

        if result.stderr.strip():
            for line in result.stderr.strip().splitlines():
                logger.warning(line)

        if result.returncode != 0:
            logger.error(f"{script_name} finalizo con codigo de salida {result.returncode}")
            return False, result.stdout + result.stderr

        logger.info(f"{script_name} finalizado exitosamente.")
        return True, result.stdout + result.stderr

    except subprocess.TimeoutExpired:
        msg = f"{script_name} excedio el tiempo maximo de ejecucion (1 hora)"
        logger.error(msg)
        return False, msg
    except Exception as e:
        msg = f"Error ejecutando {script_name}: {e}\n{traceback.format_exc()}"
        logger.error(msg)
        return False, msg


def leer_asignaciones_plantas():
    ruta = "plantas_asignadas.json"
    if not os.path.isfile(ruta):
        return {}
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def leer_clientes_para_resumen():
    clientes = []
    try:
        with open("clientes.csv", mode="r", encoding="utf-8-sig") as f:
            lector = csv.DictReader(f)
            for fila in lector:
                clientes.append(fila)
    except Exception:
        pass
    return clientes


def main():
    logger = setup_logger("sap_reginfo")
    start_time = datetime.now()

    logger.info(f"{'#' * 60}")
    logger.info(f"# INICIO DE EJECUCION: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"{'#' * 60}")

    results = []
    overall_success = True

    success_1, output_1 = run_script("asignar_plantas.py", logger)
    results.append(("asignar_plantas.py", success_1, output_1))
    if not success_1:
        overall_success = False

    logger.info("")

    success_2, output_2 = run_script("captura_sap.py", logger)
    results.append(("captura_sap.py", success_2, output_2))
    if not success_2:
        overall_success = False

    end_time = datetime.now()
    duration = end_time - start_time

    logger.info(f"{'#' * 60}")
    logger.info(f"# FIN DE EJECUCION: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"# Duracion total: {duration}")
    logger.info(f"# Resultado global: {'EXITOSO' if overall_success else 'CON ERRORES'}")
    logger.info(f"{'#' * 60}")

    email_body_lines = []
    email_body_lines.append(f"Fecha/Hora Inicio: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    email_body_lines.append(f"Fecha/Hora Fin:    {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    email_body_lines.append(f"Duracion Total:    {duration}")
    email_body_lines.append(f"Resultado Global:  {'EXITOSO' if overall_success else 'CON ERRORES'}")
    email_body_lines.append("")
    email_body_lines.append("Asignacion de plantas:")
    email_body_lines.append("-" * 40)
    plantas = leer_asignaciones_plantas()
    clientes_resumen = leer_clientes_para_resumen()
    if clientes_resumen:
        for c in clientes_resumen:
            kunnr = c.get("Cliente", "").strip()
            if not kunnr:
                continue
            manual = (c.get("Planta") or c.get("Plant") or "").strip()
            determinada = plantas.get(kunnr, "")
            if manual:
                email_body_lines.append(f"  {kunnr} -> Planta {manual} (manual)")
            elif determinada:
                email_body_lines.append(f"  {kunnr} -> Planta {determinada} (asignada)")
            else:
                email_body_lines.append(f"  {kunnr} -> Sin planta asignada")
    else:
        email_body_lines.append("  (no se encontraron clientes en clientes.csv)")
    email_body_lines.append("")
    email_body_lines.append("Detalle por script:")
    email_body_lines.append("-" * 40)
    for script_name, success, output in results:
        status = "EXITOSO" if success else "ERROR"
        email_body_lines.append(f"  [{status}] {script_name}")
        if not success and output:
            error_lines = output.strip().splitlines()
            last_errors = error_lines[-10:] if len(error_lines) > 10 else error_lines
            email_body_lines.append("  Ultimas lineas:")
            for line in last_errors:
                email_body_lines.append(f"    {line}")
        email_body_lines.append("")

    email_body = "\n".join(email_body_lines)

    status_text = "EXITOSO" if overall_success else "CON ERRORES"
    subject = f"[SAP RegInfo] Ejecucion {status_text} - {start_time.strftime('%Y-%m-%d %H:%M')}"

    log_file = get_latest_log_file("sap_reginfo")
    send_notification_email(subject, email_body, log_file_path=log_file, logger=logger)


if __name__ == "__main__":
    main()
