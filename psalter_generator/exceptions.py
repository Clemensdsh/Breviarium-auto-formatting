"""自定义异常"""

from __future__ import annotations
from typing import Optional


class PsalterError(Exception):
    """应用基础异常"""

    def __init__(self, message: str, details: Optional[str] = None):
        super().__init__(message)
        self.message = message
        self.details = details

    def __str__(self) -> str:
        if self.details:
            return f"{self.message}\n详情: {self.details}"
        return self.message


class FileError(PsalterError):
    """文件读写错误"""

    def __init__(self, message: str, filepath: Optional[str] = None, details: Optional[str] = None):
        super().__init__(message, details)
        self.filepath = filepath


class ContentError(PsalterError):
    """内容处理错误"""


class TemplateError(PsalterError):
    """模板处理错误"""

    def __init__(self, message: str, template_name: Optional[str] = None, details: Optional[str] = None):
        super().__init__(message, details)
        self.template_name = template_name
