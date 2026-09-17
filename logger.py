import logging
import os
from datetime import datetime

LOG_DIR = "logs"


def setup_logger(name="sap_reginfo"):
    os.makedirs(LOG_DIR, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(LOG_DIR, f"{timestamp}_{name}.log")

    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    if logger.handlers:
        return logger

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    file_handler.setFormatter(file_formatter)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter("%(message)s")
    console_handler.setFormatter(console_formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    logger.info(f"Log file: {log_file}")
    return logger


def get_latest_log_file(name="sap_reginfo"):
    if not os.path.isdir(LOG_DIR):
        return None
    files = [f for f in os.listdir(LOG_DIR) if f.endswith(f"_{name}.log")]
    if not files:
        return None
    files.sort(reverse=True)
    return os.path.join(LOG_DIR, files[0])
