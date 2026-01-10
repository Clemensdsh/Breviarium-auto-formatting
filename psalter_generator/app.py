#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app.py - 主应用程序
Psalter LaTeX Generator - Magnificat风格界面
"""

from __future__ import annotations
import os
import sys
import shutil
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from typing import Optional

from .theme import Theme, DEFAULT_THEME
from .models import ContentItem, TitlePageData, MultiLineContentItem
from .config import CONTENT_CATEGORIES, AppConfig
from .content_manager import ContentManager
from .latex_generator import BodyTexGenerator
from .compiler_service import CompilerService, CompileResult
from .ui_components import (
    PanedWindow, ButtonFactory, ScrollableListbox, 
    ScrollableTreeview, TelegramScrollbar
)
from .dialogs import (
    CustomContentDialog, TitlePageDialog, 
    CompileErrorDialog, LoadingDialog, ScoreDialog,
    RuleTypeDialog, ImageDialog
)
from .module_dialogs import (
    ModuleSelectionDialog, ModuleConfigDialog, 
    ModuleEditorDialog, ImprimaturDialog
)
from .modules import HourType, HourModule, ModuleConfig, create_hour_module

logger = logging.getLogger(__name__)


def get_application_path() -> str:
    """获取应用程序运行目录"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class PsalterApp:
    """Psalter LaTeX Generator 主应用 - Magnificat风格"""
    
    def __init__(
        self, 
        root: tk.Tk, 
        theme: Theme = DEFAULT_THEME,
        config: Optional[AppConfig] = None
    ):
        self.root = root
        self.theme = theme
        self.config = config or AppConfig()
        self.button_factory = ButtonFactory(theme)
        
        # 路径
        self.base_dir = get_application_path()
        self.content_dir = os.path.join(self.base_dir, "content")
        self.images_dir = os.path.join(self.base_dir, "images")
        self.gabc_dir = os.path.join(self.base_dir, "gabc")
        
        # 业务组件
        self.content_manager = ContentManager(self.content_dir)
        self.body_generator = BodyTexGenerator()
        self.compiler_service = CompilerService(self.base_dir, self.config)
        
        # 数据
        self.title_data = TitlePageData()
        
        # 初始化
        os.makedirs(self.images_dir, exist_ok=True)
        os.makedirs(self.gabc_dir, exist_ok=True)
        self._configure_window()
        self._setup_ui()
        
        logger.info("PsalterApp 初始化完成")
    
    def _configure_window(self) -> None:
        """配置窗口 - 确保完整显示"""
        self.root.title("Psalter LaTeX Generator")
        
        # 获取屏幕尺寸
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # 窗口尺寸（略小于屏幕，留出任务栏空间）
        window_width = min(1280, screen_width - 100)
        window_height = min(850, screen_height - 100)  # 留出任务栏空间
        
        # 居中位置
        x = (screen_width - window_width) // 2
        y = max(20, (screen_height - window_height) // 2 - 40)  # 略微靠上，确保底部可见
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.minsize(900, 600)
        self.root.configure(bg=self.theme.BG_DARK)
        
        # 延迟设置窗口图标，确保窗口完全初始化
        self.root.after(100, self._set_window_icon)
    
    def _set_window_icon(self) -> None:
        """设置主窗口图标"""
        try:
            from .icon import set_window_icon
            set_window_icon(self.root)
        except Exception as e:
            logger.warning(f"设置窗口图标失败: {e}")
    
    def _setup_ui(self) -> None:
        """设置UI"""
        main = tk.Frame(self.root, bg=self.theme.BG_DARK)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        main.columnconfigure(0, weight=0, minsize=220)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)
        
        self._setup_left_panel(main)
        self._setup_center_right_panel(main)
    
    def _setup_left_panel(self, parent: tk.Frame) -> None:
        """左侧面板 - 文件浏览"""
        left = tk.Frame(parent, bg=self.theme.BG_DARK, width=230)
        left.grid(row=0, column=0, sticky='nsew', padx=(0, 8))
        left.grid_propagate(False)
        
        # 标题
        tk.Label(
            left, text="内容来源", bg=self.theme.BG_DARK, fg=self.theme.RUBRIC_RED,
            font=self.theme.get_font("title", bold=True)
        ).pack(anchor='w', pady=(0, 8))
        
        # 文件树
        tree_frame = tk.Frame(left, bg=self.theme.BG_LIGHT, highlightthickness=1, 
                             highlightbackground=self.theme.BORDER)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        
        self.file_tree = ScrollableTreeview(tree_frame, self.theme)
        self.file_tree.pack(fill=tk.BOTH, expand=True)
        self._load_file_tree()
        
        # 按钮区域
        btn_frame = tk.Frame(left, bg=self.theme.BG_DARK)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=5)
        
        # 主要操作按钮
        buttons = [
            ("添加选中文件", self._add_selected_file, "accent"),
            ("添加自定义内容", self._add_custom_content, "default"),
            ("添加时辰模块", self._add_hour_module, "accent"),
            ("添加图片 *", self._add_image, "default"),
            ("添加乐谱", self._add_score, "accent"),
            ("添加分隔线", self._add_rule, "default"),
            ("添加分页", self._add_pagebreak, "default"),
            ("添加目录起始", self._add_tocstart, "default"),
            ("切换单栏/双栏", self._add_singlecol, "default"),  # 与分页按钮相同颜色
            ("添加签名块", self._add_imprimatur, "default"),
        ]
        
        for text, cmd, style in buttons:
            btn = self.button_factory.create(btn_frame, text, cmd, style)
            btn.pack(fill=tk.X, pady=2)
        
        # 图片说明
        tk.Label(
            btn_frame, text="* 图片须放在 images 文件夹", 
            bg=self.theme.BG_DARK, fg=self.theme.TEXT_SEC,
            font=self.theme.get_font("small")
        ).pack(anchor='w', pady=(2, 5))
        
        # 分隔线
        tk.Frame(btn_frame, bg=self.theme.BORDER, height=1).pack(fill=tk.X, pady=5)
        
        # 封面设置
        self.button_factory.create(
            btn_frame, "设置封面标题", self._edit_title_page, "accent"
        ).pack(fill=tk.X, pady=2)
    
    def _setup_center_right_panel(self, parent: tk.Frame) -> None:
        """中右侧面板"""
        paned = PanedWindow(parent, self.theme, initial_ratio=0.35)
        paned.grid(row=0, column=1, sticky='nsew')
        
        self._setup_center_panel(paned.left_frame)
        self._setup_right_panel(paned.right_frame)
    
    def _setup_center_panel(self, parent: tk.Frame) -> None:
        """中间面板 - 内容列表"""
        # 标题
        tk.Label(
            parent, text="当前内容", bg=self.theme.BG_LIGHT, fg=self.theme.RUBRIC_RED,
            font=self.theme.get_font("title", bold=True)
        ).pack(anchor='w', pady=(8, 8), padx=8)
        
        # 列表框
        list_frame = tk.Frame(parent, bg=self.theme.BG_LIGHT)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8), padx=8)
        
        self.content_listbox = ScrollableListbox(list_frame, self.theme)
        self.content_listbox.pack(fill=tk.BOTH, expand=True)
        self.content_listbox.bind('<Double-1>', lambda e: self._edit_item())
        
        # 操作按钮
        ops = tk.Frame(parent, bg=self.theme.BG_LIGHT)
        ops.pack(fill=tk.X, padx=8, pady=(0, 8))
        
        row1 = tk.Frame(ops, bg=self.theme.BG_LIGHT)
        row1.pack(fill=tk.X, pady=2)
        
        for text, cmd, style in [
            ("上移", self._move_up, "default"),
            ("下移", self._move_down, "default"),
            ("编辑", self._edit_item, "accent"),
            ("删除", self._delete_item, "danger"),
        ]:
            self.button_factory.create(row1, text, cmd, style, width=6).pack(side=tk.LEFT, padx=1)
        
        row2 = tk.Frame(ops, bg=self.theme.BG_LIGHT)
        row2.pack(fill=tk.X, pady=2)
        self.button_factory.create(row2, "清空全部", self._clear_all, "danger", width=13).pack(side=tk.LEFT, padx=1)
    
    def _setup_right_panel(self, parent: tk.Frame) -> None:
        """右侧面板 - 预览和导出"""
        # 标题
        tk.Label(
            parent, text="预览和导出", bg=self.theme.BG_LIGHT, fg=self.theme.RUBRIC_RED,
            font=self.theme.get_font("title", bold=True)
        ).pack(anchor='w', pady=(8, 8), padx=8)
        
        # 预览区
        preview_frame = tk.Frame(parent, bg=self.theme.BG_LIGHT, highlightthickness=1,
                                highlightbackground=self.theme.BORDER)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8), padx=8)
        
        self.preview_text = tk.Text(
            preview_frame, bg="#FFFFFF", fg=self.theme.TEXT,
            insertbackground=self.theme.TEXT, font=self.theme.get_mono_font(),
            borderwidth=0, highlightthickness=0, wrap=tk.NONE, padx=8, pady=8
        )
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scroll = TelegramScrollbar(preview_frame, command=self.preview_text.yview, theme=self.theme)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_text.config(yscrollcommand=scroll.set)
        
        # 导出按钮区
        export_frame = tk.Frame(parent, bg=self.theme.BG_LIGHT)
        export_frame.pack(fill=tk.X, pady=(0, 8), padx=8)
        
        for text, cmd, style, w in [
            ("刷新预览", self._refresh_preview, "default", 8),
            ("保存CSV", self._export_csv, "default", 8),
            ("导出 Body.tex", self._export_tex, "danger", 12),  # 改为红色
        ]:
            self.button_factory.create(export_frame, text, cmd, style, width=w).pack(side=tk.LEFT, padx=2)
        
        # 编译按钮 - 礼仪红醒目
        compile_btn = tk.Label(
            export_frame, text="编译并预览 PDF",
            bg=self.theme.RUBRIC_RED, fg="#FFFFFF",
            font=self.theme.get_font("normal", bold=True),
            cursor='hand2', padx=15, pady=6
        )
        compile_btn.pack(side=tk.LEFT, padx=5)
        compile_btn.bind('<Enter>', lambda e: compile_btn.config(bg=self.theme.RUBRIC_DARK))
        compile_btn.bind('<Leave>', lambda e: compile_btn.config(bg=self.theme.RUBRIC_RED))
        compile_btn.bind('<Button-1>', lambda e: self._compile_preview())
    
    # ========== 文件树 ==========
    
    def _load_file_tree(self) -> None:
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        files = self.content_manager.get_available_files()
        for category, name in CONTENT_CATEGORIES.items():
            cat_id = self.file_tree.insert("", tk.END, text=name, open=False)
            for _, filename in files.get(category, []):
                self.file_tree.insert(cat_id, tk.END, text=filename, values=(category, filename))
    
    # ========== 内容操作 ==========
    
    def _add_selected_file(self) -> None:
        selection = self.file_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要添加的文件")
            return
        
        for sid in selection:
            item = self.file_tree.item(sid)
            vals = item.get('values', [])
            if len(vals) >= 2:
                self.content_manager.add_from_file(vals[0], vals[1])
        
        self._refresh_listbox()
        self._refresh_preview()
    
    def _add_custom_content(self) -> None:
        dialog = CustomContentDialog(self.root, theme=self.theme)
        self.root.wait_window(dialog.top)
        
        if dialog.result:
            self.content_manager.add(dialog.result)
            self._refresh_listbox()
            self._refresh_preview()
    
    def _add_image(self) -> None:
        """添加图片 - 自动定位到images目录"""
        dialog = ImageDialog(self.root, self.images_dir, self.theme)
        self.root.wait_window(dialog.top)
        
        if dialog.result:
            self.content_manager.add(dialog.result)
            self._refresh_listbox()
            self._refresh_preview()
    
    def _add_score(self) -> None:
        """添加乐谱"""
        dialog = ScoreDialog(self.root, gabc_dir=self.gabc_dir, theme=self.theme)
        self.root.wait_window(dialog.top)
        
        if dialog.result:
            if dialog.result.arg != "inline":
                src_path = dialog.result.latin
                if os.path.exists(src_path):
                    filename = os.path.basename(src_path)
                    dest = os.path.join(self.gabc_dir, filename)
                    if not os.path.exists(dest):
                        shutil.copy2(src_path, dest)
                    dialog.result.latin = f"gabc/{filename}"
            
            self.content_manager.add(dialog.result)
            self._refresh_listbox()
            self._refresh_preview()
    
    def _add_rule(self) -> None:
        """添加分隔线 - 使用正式选择对话框"""
        dialog = RuleTypeDialog(self.root, self.theme)
        result = dialog.show()
        
        if result == "thin":
            self.content_manager.add(ContentItem("rule", "", "", ""))
        elif result == "thick":
            self.content_manager.add(ContentItem("thickrule", "", "", ""))
        
        if result:
            self._refresh_listbox()
            self._refresh_preview()
    
    def _add_pagebreak(self) -> None:
        self.content_manager.add(ContentItem("pagebreak", "", "", ""))
        self._refresh_listbox()
        self._refresh_preview()
    
    def _add_tocstart(self) -> None:
        self.content_manager.add(ContentItem("tocstart", "", "", ""))
        self._refresh_listbox()
        self._refresh_preview()
    
    def _add_singlecol(self) -> None:
        self.content_manager.add(ContentItem("singlecol", "", "", ""))
        self._refresh_listbox()
        self._refresh_preview()
    
    def _move_up(self) -> None:
        sel = self.content_listbox.curselection()
        if not sel or sel[0] == 0:
            return
        
        idx = sel[0]
        if self.content_manager.move_up(idx):
            self._refresh_listbox()
            self.content_listbox.selection_set(idx - 1)
            self._refresh_preview()
    
    def _move_down(self) -> None:
        sel = self.content_listbox.curselection()
        if not sel or sel[0] >= self.content_manager.count - 1:
            return
        
        idx = sel[0]
        if self.content_manager.move_down(idx):
            self._refresh_listbox()
            self.content_listbox.selection_set(idx + 1)
            self._refresh_preview()
    
    def _edit_item(self) -> None:
        sel = self.content_listbox.curselection()
        if not sel:
            return
        
        item = self.content_manager.get(sel[0])
        if isinstance(item, MultiLineContentItem):
            messagebox.showinfo("提示", "多行文件内容无法直接编辑。")
            return
        
        dialog = CustomContentDialog(self.root, item, self.theme)
        self.root.wait_window(dialog.top)
        
        if dialog.result:
            self.content_manager.update(sel[0], dialog.result)
            self._refresh_listbox()
            self._refresh_preview()
    
    def _delete_item(self) -> None:
        sel = self.content_listbox.curselection()
        if not sel:
            return
        
        if messagebox.askyesno("确认", "确定要删除选中的项目吗？"):
            self.content_manager.remove(sel[0])
            self._refresh_listbox()
            self._refresh_preview()
    
    def _clear_all(self) -> None:
        if messagebox.askyesno("确认", "确定要清空所有内容吗？"):
            self.content_manager.clear()
            self._refresh_listbox()
            self._refresh_preview()
    
    # ========== UI刷新 ==========
    
    def _refresh_listbox(self) -> None:
        self.content_listbox.delete(0, tk.END)
        for item in self.content_manager.items:
            self.content_listbox.insert(tk.END, item.get_display_text())
    
    def _refresh_preview(self) -> None:
        self.preview_text.delete(1.0, tk.END)
        try:
            text = self.content_manager.to_preview_text()
            self.preview_text.insert(tk.END, text)
        except Exception as e:
            logger.error(f"预览出错: {e}")
            self.preview_text.insert(tk.END, f"预览出错: {str(e)}")
    
    # ========== 标题页 ==========
    
    def _edit_title_page(self) -> None:
        dialog = TitlePageDialog(self.root, self.title_data, self.theme)
        self.root.wait_window(dialog.top)
        
        if dialog.result:
            self.title_data = dialog.result
    
    # ========== 导出 ==========
    
    def _export_csv(self) -> None:
        if self.content_manager.is_empty():
            messagebox.showwarning("提示", "没有内容可导出")
            return
        
        filepath = filedialog.asksaveasfilename(
            title="保存CSV工程文件",
            defaultextension=".csv",
            filetypes=[("CSV文件", "*.csv")],
            initialfile="psalter_project.csv"
        )
        
        if filepath:
            try:
                self.content_manager.save_to_csv(filepath)
                messagebox.showinfo("成功", f"工程文件已保存到:\n{filepath}")
            except Exception as e:
                logger.error(f"保存CSV失败: {e}")
                messagebox.showerror("错误", f"保存失败: {str(e)}")
    
    def _export_tex(self) -> None:
        if self.content_manager.is_empty():
            messagebox.showwarning("提示", "没有内容可导出")
            return
        
        filepath = filedialog.asksaveasfilename(
            title="生成 TeX 文件",
            defaultextension=".tex",
            filetypes=[("TeX 文件", "*.tex")],
            initialfile="body.tex"
        )
        
        if not filepath:
            return
        
        try:
            self.body_generator.save(self.content_manager.items, filepath)
            messagebox.showinfo("成功", f"文件已生成:\n{filepath}")
        except Exception as e:
            logger.error(f"导出TeX失败: {e}")
            messagebox.showerror("错误", f"生成失败: {str(e)}")
    
    # ========== 编译 ==========
    
    def _compile_preview(self) -> None:
        if self.content_manager.is_empty():
            messagebox.showwarning("提示", "内容为空，无法编译")
            return
        
        loading = LoadingDialog(self.root, "正在编译...", self.theme)
        
        def progress(msg: str) -> None:
            loading.update_message(msg)
        
        try:
            output = self.compiler_service.compile_and_preview(
                self.content_manager.items,
                self.title_data,
                progress
            )
            
            loading.close()
            
            if output.result == CompileResult.SUCCESS:
                if output.warnings:
                    logger.warning(f"编译警告: {len(output.warnings)} 条")
            elif output.result == CompileResult.MISSING_FILES:
                messagebox.showerror("错误", f"缺失核心文件:\n{output.message}")
            elif output.result == CompileResult.COMPILER_NOT_FOUND:
                messagebox.showerror("错误", output.message)
            elif output.result == CompileResult.COMPILE_ERROR:
                CompileErrorDialog(self.root, output.error_log or "", self.theme)
            elif output.result == CompileResult.COMPILE_TIMEOUT:
                messagebox.showerror("错误", "编译超时，请检查文档是否过于复杂")
            elif output.result == CompileResult.PDF_NOT_GENERATED:
                messagebox.showerror("失败", "编译完成但未生成PDF")
            else:
                messagebox.showerror("错误", output.message)
                
        except Exception as e:
            loading.close()
            logger.exception("编译异常")
            messagebox.showerror("系统错误", str(e))
    
    # ========== 模块操作 ==========
    
    def _add_hour_module(self) -> None:
        """添加时辰模块"""
        selection_dialog = ModuleSelectionDialog(self.root, self.theme)
        hour_type = selection_dialog.show()
        
        if not hour_type:
            return
        
        config_dialog = ModuleConfigDialog(self.root, hour_type, self.theme)
        config = config_dialog.show()
        
        if not config:
            return
        
        try:
            module = create_hour_module(hour_type, config)
        except Exception as e:
            logger.error(f"创建模块失败: {e}")
            messagebox.showerror("错误", f"创建模块失败: {str(e)}")
            return
        
        editor = ModuleEditorDialog(self.root, module, self.theme)
        result = editor.show()
        
        if not result:
            return
        
        try:
            items = result.expand()
            for item in items:
                self.content_manager.add(item)
            
            self._refresh_listbox()
            self._refresh_preview()
            
            messagebox.showinfo("成功", f"已添加 {result.name_zh}，共 {len(items)} 个内容项")
        except Exception as e:
            logger.error(f"展开模块失败: {e}")
            messagebox.showerror("错误", f"展开模块失败: {str(e)}")
    
    def _add_imprimatur(self) -> None:
        """添加Imprimatur/Nihil Obstat签名块"""
        dialog = ImprimaturDialog(self.root, self.theme)
        result = dialog.show()
        
        if result:
            for item in result:
                self.content_manager.add(item)
            
            self._refresh_listbox()
            self._refresh_preview()
            messagebox.showinfo("成功", "签名块已添加")
