#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
exceptions.py - 自定义异常类
提供统一的错误处理机制
"""

from __future__ import annotations
from typing import Optional, List


class PsalterError(Exception):
    """Psalter应用程序基础异常"""
    
    def __init__(self, message: str, details: Optional[str] = None):
        super().__init__(message)
        self.message = message
        self.details = details
    
    def __str__(self) -> str:
        if self.details:
            return f"{self.message}\n详情: {self.details}"
        return self.message


class ConfigError(PsalterError):
    """配置相关错误"""
    pass


class FileError(PsalterError):
    """文件操作相关错误"""
    
    def __init__(self, message: str, filepath: Optional[str] = None, details: Optional[str] = None):
        super().__init__(message, details)
        self.filepath = filepath


class ContentError(PsalterError):
    """内容处理相关错误"""
    
    def __init__(self, message: str, item_type: Optional[str] = None, details: Optional[str] = None):
        super().__init__(message, details)
        self.item_type = item_type


class ValidationError(PsalterError):
    """数据验证错误"""
    
    def __init__(self, message: str, field: Optional[str] = None, details: Optional[str] = None):
        super().__init__(message, details)
        self.field = field


class CompileError(PsalterError):
    """LaTeX编译错误"""
    
    def __init__(
        self, 
        message: str, 
        log_content: Optional[str] = None,
        missing_files: Optional[List[str]] = None,
        details: Optional[str] = None
    ):
        super().__init__(message, details)
        self.log_content = log_content
        self.missing_files = missing_files or []
    
    def get_error_summary(self) -> str:
        """获取错误摘要"""
        lines = [self.message]
        
        if self.missing_files:
            lines.append(f"缺失文件: {', '.join(self.missing_files)}")
        
        if self.log_content:
            # 提取关键错误行
            error_lines = [
                line for line in self.log_content.split('\n')
                if 'error' in line.lower() or 'fatal' in line.lower()
            ]
            if error_lines:
                lines.append("关键错误:")
                lines.extend(error_lines[:5])  # 最多显示5行
        
        return '\n'.join(lines)


class TemplateError(PsalterError):
    """模板处理错误"""
    
    def __init__(self, message: str, template_name: Optional[str] = None, details: Optional[str] = None):
        super().__init__(message, details)
        self.template_name = template_name
