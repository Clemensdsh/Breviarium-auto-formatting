#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
compiler_service.py - LaTeX编译服务
负责调用XeLaTeX编译并管理构建目录
支持超时控制和更详细的错误报告
"""

from __future__ import annotations
import os
import shutil
import subprocess
import platform
import logging
from typing import Optional, Tuple, Callable, List
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

from .models import TitlePageData, ContentItem
from .latex_generator import MainTexGenerator, BodyTexGenerator
from .config import AppConfig
from .exceptions import CompileError, FileError

logger = logging.getLogger(__name__)


class CompileResult(Enum):
    """编译结果枚举"""
    SUCCESS = "success"
    MISSING_FILES = "missing_files"
    COMPILER_NOT_FOUND = "compiler_not_found"
    COMPILE_ERROR = "compile_error"
    COMPILE_TIMEOUT = "compile_timeout"
    PDF_NOT_GENERATED = "pdf_not_generated"
    DIRECTORY_ERROR = "directory_error"


@dataclass
class CompileOutput:
    """编译输出结果"""
    result: CompileResult
    pdf_path: Optional[str] = None
    error_log: Optional[str] = None
    message: str = ""
    warnings: List[str] = field(default_factory=list)
    
    @property
    def is_success(self) -> bool:
        return self.result == CompileResult.SUCCESS
    
    def get_summary(self) -> str:
        """获取结果摘要"""
        lines = [f"结果: {self.result.value}", self.message]
        if self.warnings:
            lines.append(f"警告: {len(self.warnings)} 条")
        return "\n".join(lines)


class BuildDirectoryManager:
    """构建目录管理器"""
    
    def __init__(self, base_dir: str, build_dir_name: str = "build"):
        self.base_dir = Path(base_dir)
        self.build_dir = self.base_dir / build_dir_name
    
    def prepare(self) -> bool:
        """准备构建目录（清空或创建）"""
        try:
            if self.build_dir.exists():
                self._clean_directory()
            else:
                self.build_dir.mkdir(parents=True)
            logger.debug(f"构建目录已准备: {self.build_dir}")
            return True
        except Exception as e:
            logger.error(f"准备构建目录失败: {e}")
            return False
    
    def _clean_directory(self) -> None:
        """清空构建目录"""
        for item in self.build_dir.iterdir():
            try:
                if item.is_file() or item.is_symlink():
                    item.unlink()
                elif item.is_dir():
                    shutil.rmtree(item)
            except Exception as e:
                logger.warning(f"删除失败 {item}: {e}")
    
    def copy_resources(self, files: List[str], directories: Optional[List[str]] = None) -> bool:
        """复制资源文件到构建目录"""
        try:
            # 复制文件
            for filename in files:
                src = self.base_dir / filename
                if src.exists():
                    shutil.copy2(src, self.build_dir)
                    logger.debug(f"已复制: {filename}")
                else:
                    logger.warning(f"资源文件不存在: {src}")
            
            # 复制目录
            if directories:
                for dirname in directories:
                    src = self.base_dir / dirname
                    dst = self.build_dir / dirname
                    if src.exists():
                        shutil.copytree(src, dst, dirs_exist_ok=True)
                        logger.debug(f"已复制目录: {dirname}")
                    else:
                        dst.mkdir(parents=True, exist_ok=True)
            
            return True
        except Exception as e:
            logger.error(f"复制资源失败: {e}")
            return False
    
    def get_output_path(self, filename: str) -> Path:
        """获取构建目录中的文件路径"""
        return self.build_dir / filename
    
    def write_file(self, filename: str, content: str) -> bool:
        """写入文件到构建目录"""
        try:
            filepath = self.build_dir / filename
            filepath.write_text(content, encoding='utf-8')
            return True
        except Exception as e:
            logger.error(f"写入文件失败 {filename}: {e}")
            return False


class LaTeXCompiler:
    """LaTeX编译器 - 支持 LuaLaTeX (默认) 和 XeLaTeX"""
    
    def __init__(self, timeout: int = 120, engine: str = "lualatex"):
        self.timeout = timeout
        self.engine = engine  # "lualatex" 或 "xelatex"
        self.compiler_cmd = engine
    
    def is_available(self) -> bool:
        """检查编译器是否可用"""
        return shutil.which(self.compiler_cmd) is not None
    
    def set_engine(self, engine: str) -> None:
        """设置编译引擎"""
        if engine in ("lualatex", "xelatex"):
            self.engine = engine
            self.compiler_cmd = engine
        else:
            raise ValueError(f"不支持的编译引擎: {engine}")
    
    def compile(
        self, 
        working_dir: Path, 
        tex_file: str = "main.tex", 
        runs: int = 2
    ) -> Tuple[bool, str, List[str]]:
        """
        编译LaTeX文件
        
        Args:
            working_dir: 工作目录
            tex_file: 主tex文件名
            runs: 编译次数
        
        Returns:
            (成功与否, 日志内容, 警告列表)
        """
        cmd = [self.compiler_cmd, '-interaction=nonstopmode', '-shell-escape', tex_file]
        full_log = ""
        warnings: List[str] = []
        
        try:
            for i in range(runs):
                logger.info(f"使用 {self.engine} 编译第 {i+1}/{runs} 次...")
                
                result = subprocess.run(
                    cmd,
                    cwd=str(working_dir),
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    errors='replace',
                    timeout=self.timeout
                )
                
                full_log += f"\n{'='*20} Run {i+1} ({self.engine}) {'='*20}\n{result.stdout}"
                
                # 提取警告
                for line in result.stdout.split('\n'):
                    if 'warning' in line.lower():
                        warnings.append(line.strip())
                
                # 第一次编译失败就直接返回
                if i == 0 and result.returncode != 0:
                    logger.error("第一次编译失败")
                    return False, full_log, warnings
            
            logger.info("编译完成")
            return True, full_log, warnings
            
        except subprocess.TimeoutExpired:
            msg = f"编译超时（{self.timeout}秒）"
            logger.error(msg)
            return False, msg, warnings
        except Exception as e:
            logger.error(f"编译异常: {e}")
            return False, str(e), warnings


# 保留旧名称以兼容
XeLatexCompiler = LaTeXCompiler


class PDFViewer:
    """PDF查看器 - 跨平台打开PDF"""
    
    @staticmethod
    def open(pdf_path: str) -> bool:
        """打开PDF文件"""
        if not os.path.exists(pdf_path):
            logger.error(f"PDF文件不存在: {pdf_path}")
            return False
        
        try:
            system = platform.system()
            logger.info(f"打开PDF: {pdf_path} (系统: {system})")
            
            if system == 'Windows':
                os.startfile(pdf_path)
            elif system == 'Darwin':
                subprocess.call(['open', pdf_path])
            else:
                subprocess.call(['xdg-open', pdf_path])
            return True
        except Exception as e:
            logger.error(f"打开PDF失败: {e}")
            return False


class CompilerService:
    """编译服务 - 整合所有编译相关功能"""
    
    def __init__(self, base_dir: str, config: Optional[AppConfig] = None, engine: str = "lualatex"):
        self.base_dir = Path(base_dir)
        self.config = config or AppConfig()
        self.engine = engine
        
        self.build_manager = BuildDirectoryManager(
            str(base_dir), 
            self.config.build_dir_name
        )
        self.compiler = LaTeXCompiler(self.config.compiler_timeout, engine=engine)
        self.main_generator = MainTexGenerator(str(self.base_dir / "main.tex"))
        self.body_generator = BodyTexGenerator()
        
        # 添加gabc到资源目录列表
        self._resource_directories = list(self.config.resource_directories) + ["gabc"]
        
        logger.info(f"CompilerService 初始化: {base_dir} (引擎: {engine})")
    
    def set_engine(self, engine: str) -> None:
        """设置编译引擎"""
        self.engine = engine
        self.compiler.set_engine(engine)
        logger.info(f"编译引擎已切换为: {engine}")
    
    def check_prerequisites(self) -> Tuple[bool, List[str]]:
        """检查编译前提条件"""
        missing = []
        
        for filename in self.config.required_files:
            if not (self.base_dir / filename).exists():
                missing.append(filename)
        
        return len(missing) == 0, missing
    
    def compile(
        self, 
        items: List[ContentItem], 
        title_data: TitlePageData,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> CompileOutput:
        """
        执行完整的编译流程
        
        Args:
            items: 内容项列表
            title_data: 标题页数据
            progress_callback: 进度回调函数
        
        Returns:
            CompileOutput对象
        """
        def report(msg: str) -> None:
            if progress_callback:
                progress_callback(msg)
            logger.info(msg)
        
        # 1. 检查前提条件
        report("检查必要文件...")
        ok, missing = self.check_prerequisites()
        if not ok:
            return CompileOutput(
                result=CompileResult.MISSING_FILES,
                message=f"缺失文件: {', '.join(missing)}"
            )
        
        # 2. 检查编译器
        if not self.compiler.is_available():
            return CompileOutput(
                result=CompileResult.COMPILER_NOT_FOUND,
                message=f"未找到 {self.engine} 命令，请确保已安装 TeX Live 或 MiKTeX"
            )
        
        # 3. 准备构建目录
        report("准备构建目录...")
        if not self.build_manager.prepare():
            return CompileOutput(
                result=CompileResult.DIRECTORY_ERROR,
                message="无法创建构建目录"
            )
        
        # 4. 复制资源
        report("复制资源文件...")
        self.build_manager.copy_resources(
            files=["psalter.sty"],
            directories=self._resource_directories
        )
        
        # 5. 生成main.tex
        report("生成 main.tex...")
        try:
            main_content = self.main_generator.generate(title_data)
            self.build_manager.write_file("main.tex", main_content)
        except Exception as e:
            return CompileOutput(
                result=CompileResult.COMPILE_ERROR,
                error_log=str(e),
                message="生成 main.tex 失败"
            )
        
        # 6. 生成body.tex
        report("生成 body.tex...")
        try:
            body_content = self.body_generator.generate(items)
            self.build_manager.write_file("body.tex", body_content)
        except Exception as e:
            return CompileOutput(
                result=CompileResult.COMPILE_ERROR,
                error_log=str(e),
                message="生成 body.tex 失败"
            )
        
        # 7. 编译
        report(f"调用 {self.engine.upper()} 编译...")
        success, log, warnings = self.compiler.compile(
            self.build_manager.build_dir,
            runs=self.config.compile_runs
        )
        
        if not success:
            # 检查是否超时
            if "超时" in log:
                return CompileOutput(
                    result=CompileResult.COMPILE_TIMEOUT,
                    error_log=log,
                    message="编译超时"
                )
            return CompileOutput(
                result=CompileResult.COMPILE_ERROR,
                error_log=log,
                message="编译失败",
                warnings=warnings
            )
        
        # 8. 检查PDF生成
        pdf_path = self.build_manager.get_output_path("main.pdf")
        if not pdf_path.exists():
            return CompileOutput(
                result=CompileResult.PDF_NOT_GENERATED,
                error_log=log,
                message="编译完成但未生成PDF",
                warnings=warnings
            )
        
        report("编译成功！")
        return CompileOutput(
            result=CompileResult.SUCCESS,
            pdf_path=str(pdf_path),
            message="编译成功",
            warnings=warnings
        )
    
    def compile_and_preview(
        self, 
        items: List[ContentItem], 
        title_data: TitlePageData,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> CompileOutput:
        """编译并自动打开预览"""
        output = self.compile(items, title_data, progress_callback)
        
        if output.is_success and output.pdf_path:
            PDFViewer.open(output.pdf_path)
        
        return output
