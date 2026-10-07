"""日志配置"""

from __future__ import annotations
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional


def init_logging(debug: bool = False, log_file: Optional[str] = None) -> None:
    """配置 psalter_generator 日志；debug 模式下默认写入 logs/ 目录"""
    logger = logging.getLogger("psalter_generator")
    if logger.handlers:
        return

    level = logging.DEBUG if debug else logging.INFO
    logger.setLevel(level)
    formatter = logging.Formatter(
        '%(asctime)s [%(levelname)s] %(name)s: %(message)s', '%Y-%m-%d %H:%M:%S'
    )

    if log_file is None and debug:
        log_file = f"logs/psalter_{datetime.now():%Y%m%d_%H%M%S}.log"

    handlers: list[logging.Handler] = [logging.StreamHandler(sys.stdout)]
    if log_file:
        Path(log_file).parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(log_file, encoding='utf-8'))

    for handler in handlers:
        handler.setLevel(level)
        handler.setFormatter(formatter)
        logger.addHandler(handler)
