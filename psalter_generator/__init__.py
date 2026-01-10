#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Psalter LaTeX Generator v3.0
一个用于生成礼仪书LaTeX文档的工具

v3.0 改进内容:
- 支持 LuaLaTeX 编译引擎
- 支持 Gregorio 乐谱
- 时辰模块系统（封装礼仪时辰结构）
- 新增 LaTeX 宏（红黑双色、签名块等）
- 完整的类型注解
- 统一的日志系统

模块结构:
- theme.py: UI主题配置
- models.py: 数据模型
- config.py: LaTeX命令映射配置
- modules.py: 时辰模块定义
- module_dialogs.py: 模块配置对话框
- app.py: 主应用
"""

from .theme import Theme, DEFAULT_THEME, LIGHT_THEME
from .models import ContentItem, MultiLineContentItem, TitlePageData, ModuleItem, MODULE_TYPE_NAMES
from .config import TexMappingConfig, CONTENT_CATEGORIES, AppConfig
from .exceptions import (
    PsalterError, ConfigError, FileError, 
    ContentError, ValidationError, CompileError, TemplateError
)
from .logger import setup_logger, get_logger, init_logging, LoggerMixin
from .content_manager import ContentManager, FileContentLoader
from .latex_generator import TexGenerator, BodyTexGenerator, MainTexGenerator
from .compiler_service import CompilerService, CompileResult, CompileOutput, LaTeXCompiler
from .modules import (
    HourType, HourModule, ModuleConfig, create_hour_module,
    MatinsModule1962, LaudsModule1962, VespersModule1962, 
    ComplineModule1962, SmallHourModule1962,
    OfficeOfReadingsModule, LaudsModuleLOTH, VespersModuleLOTH,
    ComplineModuleLOTH, DaytimePrayerModuleLOTH
)
from .app import PsalterApp

__version__ = "3.1.0"
__author__ = "Psalter Generator Team"

__all__ = [
    # Version
    "__version__",
    # Theme
    "Theme",
    "DEFAULT_THEME",
    "LIGHT_THEME",
    "DARK_THEME",
    # Models
    "ContentItem",
    "MultiLineContentItem", 
    "TitlePageData",
    "ModuleItem",
    "MODULE_TYPE_NAMES",
    # Config
    "TexMappingConfig",
    "CONTENT_CATEGORIES",
    "AppConfig",
    # Exceptions
    "PsalterError",
    "ConfigError",
    "FileError",
    "ContentError",
    "ValidationError",
    "CompileError",
    "TemplateError",
    # Logger
    "setup_logger",
    "get_logger",
    "init_logging",
    "LoggerMixin",
    # Content
    "ContentManager",
    "FileContentLoader",
    # LaTeX
    "TexGenerator",
    "BodyTexGenerator",
    "MainTexGenerator",
    # Compiler
    "CompilerService",
    "CompileResult",
    "CompileOutput",
    "LaTeXCompiler",
    # Modules - 1962
    "HourType",
    "HourModule",
    "ModuleConfig",
    "create_hour_module",
    "MatinsModule1962",
    "LaudsModule1962",
    "VespersModule1962",
    "ComplineModule1962",
    "SmallHourModule1962",
    # Modules - LOTH
    "OfficeOfReadingsModule",
    "LaudsModuleLOTH",
    "VespersModuleLOTH",
    "ComplineModuleLOTH",
    "DaytimePrayerModuleLOTH",
    # App
    "PsalterApp",
]
