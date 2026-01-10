#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
logger.py - 统一日志管理
提供应用程序级别的日志配置
"""

from __future__ import annotations
import logging
import sys
from pathlib import Path
from typing import Optional
from datetime import datetime


def setup_logger(
    name: str = "psalter_generator",
    level: int = logging.INFO,
    log_file: Optional[str] = None,
    console_output: bool = True
) -> logging.Logger:
    """
    配置并返回日志记录器
    
    Args:
        name: 日志记录器名称
        level: 日志级别
        log_file: 日志文件路径（可选）
        console_output: 是否输出到控制台
    
    Returns:
        配置好的Logger实例
    """
    logger = logging.getLogger(name)
    logger.setLevel(level)
    
    # 避免重复添加handler
    if logger.handlers:
        return logger
    
    # 格式化器
    formatter = logging.Formatter(
        fmt='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # 控制台输出
    if console_output:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(level)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
    
    # 文件输出
    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(level)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    return logger


def get_logger(name: str = "psalter_generator") -> logging.Logger:
    """获取已配置的日志记录器"""
    return logging.getLogger(name)


class LoggerMixin:
    """日志混入类，为类提供日志功能"""
    
    @property
    def logger(self) -> logging.Logger:
        """获取类专属的日志记录器"""
        if not hasattr(self, '_logger'):
            self._logger = logging.getLogger(
                f"psalter_generator.{self.__class__.__name__}"
            )
        return self._logger


# 模块级日志记录器
_module_logger: Optional[logging.Logger] = None


def init_logging(debug: bool = False, log_file: Optional[str] = None) -> None:
    """初始化应用程序日志"""
    global _module_logger
    
    level = logging.DEBUG if debug else logging.INFO
    
    # 如果未指定日志文件，默认创建带时间戳的日志
    if log_file is None and debug:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = f"logs/psalter_{timestamp}.log"
    
    _module_logger = setup_logger(
        name="psalter_generator",
        level=level,
        log_file=log_file,
        console_output=True
    )
    
    _module_logger.info("日志系统初始化完成")
