"""
logger.py — Rotating daily file logger + coloured console output.
"""

import logging
import os
import sys
from logging.handlers import TimedRotatingFileHandler

RESET  = "\033[0m"
GREEN  = "\033[32m"
YELLOW = "\033[33m"
RED    = "\033[31m"
CYAN   = "\033[36m"
GREY   = "\033[90m"

LEVEL_COLOURS = {
    "DEBUG":    GREY,
    "INFO":     GREEN,
    "WARNING":  YELLOW,
    "ERROR":    RED,
    "CRITICAL": RED,
}


class ColouredFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        colour = LEVEL_COLOURS.get(record.levelname, RESET)
        msg = super().format(record)
        return f"{colour}{msg}{RESET}"


def build_logger(log_dir: str, name: str = "thermopac_agent") -> logging.Logger:
    os.makedirs(log_dir, exist_ok=True)

    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    if logger.handlers:
        return logger

    fmt_plain  = logging.Formatter("%(asctime)s [%(levelname)-8s] %(message)s",
                                   datefmt="%Y-%m-%d %H:%M:%S")
    fmt_colour = ColouredFormatter("%(asctime)s [%(levelname)-8s] %(message)s",
                                   datefmt="%Y-%m-%d %H:%M:%S")

    # Rotating file handler — new file each day, keep 30 days
    log_path = os.path.join(log_dir, "agent.log")
    fh = TimedRotatingFileHandler(log_path, when="midnight", backupCount=30,
                                   encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt_plain)
    logger.addHandler(fh)

    # Console handler
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt_colour)
    logger.addHandler(ch)

    return logger
