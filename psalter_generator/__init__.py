"""Psalter LaTeX Generator：生成拉丁文-中文对照日课经 LaTeX 文档的图形界面工具"""

from .config import AppConfig
from .logger import init_logging
from .app import PsalterApp

__version__ = "3.1.0"

__all__ = ["AppConfig", "init_logging", "PsalterApp", "__version__"]
