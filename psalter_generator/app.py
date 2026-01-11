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
from typing import Optional, List
import copy

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
    CompileErrorDialog, CompileProgressDialog, ScoreDialog,
    RuleTypeDialog, ImageDialog
)
from .module_dialogs import (
    ModuleSelectionDialog, ModuleConfigDialog, 
    ModuleEditorDialog, ImprimaturDialog
)
from .modules import HourType, HourModule, ModuleConfig, create_hour_module

logger = logging.getLogger(__name__)


def get_resource_path() -> str:
    """获取资源文件目录（兼容PyInstaller打包）"""
    if hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    else:
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_working_path() -> str:
    """获取工作目录（exe所在目录，用于输出文件）"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def ensure_external_content(resource_dir: str, working_dir: str) -> str:
    """确保外部content文件夹存在"""
    internal_content = os.path.join(resource_dir, "content")
    external_content = os.path.join(working_dir, "content")
    
    if not os.path.exists(external_content):
        if os.path.exists(internal_content):
            try:
                shutil.copytree(internal_content, external_content)
                logger.info(f"已将内置content复制到: {external_content}")
            except Exception as e:
                logger.error(f"复制content失败: {e}")
                return internal_content
    
    return external_content


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
        self.resource_dir = get_resource_path()
        self.working_dir = get_working_path()
        
        # 使用外部content文件夹
        self.content_dir = ensure_external_content(self.resource_dir, self.working_dir)
        self.images_dir = os.path.join(self.working_dir, "images")
        self.gabc_dir = os.path.join(self.working_dir, "gabc")
        
        # 业务组件
        self.content_manager = ContentManager(self.content_dir)
        self.body_generator = BodyTexGenerator()
        self.compiler_service = CompilerService(
            resource_dir=self.resource_dir,
            working_dir=self.working_dir,
            config=self.config
        )
        
        # 数据
        self.title_data = TitlePageData()
        
        # 撤销/重做栈
        self._undo_stack: List = []
        self._redo_stack: List = []
        self._clipboard: List = []
        
        # 初始化
        os.makedirs(self.images_dir, exist_ok=True)
        os.makedirs(self.gabc_dir, exist_ok=True)
        self._configure_window()
        self._setup_ui()
        
        logger.info("PsalterApp 初始化完成")
    
    def _configure_window(self) -> None:
        """配置窗口"""
        self.root.title("Psalter LaTeX Generator")
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        window_width = min(1280, screen_width - 100)
        window_height = min(950, screen_height - 80)
        
        x = (screen_width - window_width) // 2
        y = max(10, (screen_height - window_height) // 2 - 30)
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.minsize(900, 700)
        self.root.configure(bg=self.theme.BG_DARK)
        
        self.root.after(100, self._set_window_icon)
    
    def _set_window_icon(self) -> None:
        try:
            from .icon import set_window_icon
            set_window_icon(self.root)
        except Exception as e:
            logger.warning(f"设置窗口图标失败: {e}")
    
    def _setup_ui(self) -> None:
        main = tk.Frame(self.root, bg=self.theme.BG_DARK)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        main.columnconfigure(0, weight=0, minsize=220)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)
        
        self._setup_left_panel(main)
        self._setup_center_right_panel(main)
    
    def _setup_left_panel(self, parent: tk.Frame) -> None:
        left = tk.Frame(parent, bg=self.theme.BG_DARK, width=230)
        left.grid(row=0, column=0, sticky='nsew', padx=(0, 8))
        left.grid_propagate(False)
        
        tk.Label(
            left, text="内容来源", bg=self.theme.BG_DARK, fg=self.theme.RUBRIC_RED,
            font=self.theme.get_font("title", bold=True)
        ).pack(anchor='w', pady=(0, 8))
        
        tree_frame = tk.Frame(left, bg=self.theme.BG_LIGHT, highlightthickness=1, 
                             highlightbackground=self.theme.BORDER)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        
        self.file_tree = ScrollableTreeview(tree_frame, self.theme)
        self.file_tree.pack(fill=tk.BOTH, expand=True)
        self._load_file_tree()
        
        btn_container = tk.Frame(left, bg=self.theme.BG_DARK)
        btn_container.pack(side=tk.BOTTOM, fill=tk.X, pady=5)
        
        buttons = [
            ("添加选中文件", self._add_selected_file, "accent"),
            ("添加自定义内容", self._add_custom_content, "default"),
            ("添加时辰模块", self._add_hour_module, "accent"),
            ("添加图片 *", self._add_image, "default"),
            ("添加乐谱", self._add_score, "accent"),
            ("添加分隔线", self._add_rule, "default"),
            ("添加分页", self._add_pagebreak, "default"),
            ("添加目录起始", self._add_tocstart, "default"),
            ("切换单栏/双栏", self._add_singlecol, "default"),
            ("添加签名块", self._add_imprimatur, "default"),
        ]
        
        for text, cmd, style in buttons:
            btn = self.button_factory.create(btn_container, text, cmd, style)
            btn.pack(fill=tk.X, pady=2)
        
        tk.Label(
            btn_container, text="* 图片须放在 images 文件夹", 
            bg=self.theme.BG_DARK, fg=self.theme.TEXT_SEC,
            font=self.theme.get_font("small")
        ).pack(anchor='w', pady=(2, 5))
        
        tk.Frame(btn_container, bg=self.theme.BORDER, height=1).pack(fill=tk.X, pady=5)
        
        self.button_factory.create(
            btn_container, "设置封面标题", self._edit_title_page, "accent"
        ).pack(fill=tk.X, pady=2)
        
        self.button_factory.create(
            btn_container, "刷新文件列表", self._reload_file_tree, "default"
        ).pack(fill=tk.X, pady=2)
    
    def _setup_center_right_panel(self, parent: tk.Frame) -> None:
        paned = PanedWindow(parent, self.theme, initial_ratio=0.35)
        paned.grid(row=0, column=1, sticky='nsew')
        
        self._setup_center_panel(paned.left_frame)
        self._setup_right_panel(paned.right_frame)
    
    def _setup_center_panel(self, parent: tk.Frame) -> None:
        """中间面板 - 内容列表"""
        tk.Label(
            parent, text="当前内容", bg=self.theme.BG_LIGHT, fg=self.theme.RUBRIC_RED,
            font=self.theme.get_font("title", bold=True)
        ).pack(anchor='w', pady=(8, 8), padx=8)
        
        list_frame = tk.Frame(parent, bg=self.theme.BG_LIGHT)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8), padx=8)
        
        self.content_listbox = ScrollableListbox(list_frame, self.theme, selectmode=tk.EXTENDED)
        self.content_listbox.pack(fill=tk.BOTH, expand=True)
        self.content_listbox.bind('<Double-1>', lambda e: self._edit_item())
        
        # 绑定快捷键到listbox
        self.content_listbox.listbox.bind('<Delete>', lambda e: self._delete_item())
        self.content_listbox.listbox.bind('<Control-a>', lambda e: self._select_all())
        self.content_listbox.listbox.bind('<Control-A>', lambda e: self._select_all())
        self.content_listbox.listbox.bind('<Control-c>', lambda e: self._copy_selected())
        self.content_listbox.listbox.bind('<Control-C>', lambda e: self._copy_selected())
        self.content_listbox.listbox.bind('<Control-x>', lambda e: self._cut_selected())
        self.content_listbox.listbox.bind('<Control-X>', lambda e: self._cut_selected())
        self.content_listbox.listbox.bind('<Control-v>', lambda e: self._paste())
        self.content_listbox.listbox.bind('<Control-V>', lambda e: self._paste())
        self.content_listbox.listbox.bind('<Control-z>', lambda e: self._undo())
        self.content_listbox.listbox.bind('<Control-Z>', lambda e: self._undo())
        self.content_listbox.listbox.bind('<Control-y>', lambda e: self._redo())
        self.content_listbox.listbox.bind('<Control-Y>', lambda e: self._redo())
        
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
        # 标题行
        title_frame = tk.Frame(parent, bg=self.theme.BG_LIGHT)
        title_frame.pack(anchor='w', pady=(8, 8), padx=8, fill=tk.X)
        
        tk.Label(
            title_frame, text="预览和导出", bg=self.theme.BG_LIGHT, fg=self.theme.RUBRIC_RED,
            font=self.theme.get_font("title", bold=True)
        ).pack(side=tk.LEFT)
        
        tk.Label(
            title_frame, text="（只读，编辑请双击左侧列表项）", 
            bg=self.theme.BG_LIGHT, fg=self.theme.TEXT_SEC,
            font=self.theme.get_font("small")
        ).pack(side=tk.LEFT, padx=(10, 0))
        
        # 预览区
        preview_frame = tk.Frame(parent, bg=self.theme.BG_LIGHT, highlightthickness=1,
                                highlightbackground=self.theme.BORDER)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8), padx=8)
        
        self.preview_text = tk.Text(
            preview_frame, bg="#FFFFFF", fg=self.theme.TEXT,
            insertbackground=self.theme.TEXT, font=self.theme.get_mono_font(),
            borderwidth=0, highlightthickness=0, wrap=tk.NONE, padx=8, pady=8,
            state=tk.DISABLED
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
            ("导出 Body.tex", self._export_tex, "danger", 12),
        ]:
            self.button_factory.create(export_frame, text, cmd, style, width=w).pack(side=tk.LEFT, padx=2)
        
        # 编译按钮
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
    
    # ========== 撤销/重做/剪贴板 ==========
    
    def _save_state(self) -> None:
        """保存当前状态用于撤销"""
        # 获取items的真实副本
        state = [copy.deepcopy(item) for item in self.content_manager]
        self._undo_stack.append(state)
        self._redo_stack.clear()
        if len(self._undo_stack) > 50:
            self._undo_stack.pop(0)
    
    def _restore_state(self, state: List) -> None:
        """恢复状态"""
        self.content_manager.clear()
        for item in state:
            self.content_manager.add(copy.deepcopy(item))
    
    def _undo(self) -> None:
        """撤销"""
        if not self._undo_stack:
            return
        # 保存当前状态到redo栈
        current = [copy.deepcopy(item) for item in self.content_manager]
        self._redo_stack.append(current)
        # 恢复上一个状态
        self._restore_state(self._undo_stack.pop())
        self._refresh_listbox()
        self._refresh_preview()
    
    def _redo(self) -> None:
        """重做"""
        if not self._redo_stack:
            return
        # 保存当前状态到undo栈
        current = [copy.deepcopy(item) for item in self.content_manager]
        self._undo_stack.append(current)
        # 恢复redo状态
        self._restore_state(self._redo_stack.pop())
        self._refresh_listbox()
        self._refresh_preview()
    
    def _select_all(self) -> None:
        """全选"""
        self.content_listbox.listbox.selection_set(0, tk.END)
        return "break"
    
    def _copy_selected(self) -> None:
        """复制选中项"""
        sel = list(self.content_listbox.curselection())
        if sel:
            self._clipboard = [copy.deepcopy(self.content_manager.get(i)) for i in sel]
        return "break"
    
    def _cut_selected(self) -> None:
        """剪切选中项"""
        self._copy_selected()
        if self._clipboard:
            self._delete_item()
        return "break"
    
    def _paste(self) -> None:
        """粘贴"""
        if not self._clipboard:
            return "break"
        self._save_state()
        sel = list(self.content_listbox.curselection())
        insert_pos = sel[-1] + 1 if sel else self.content_manager.count
        for i, item in enumerate(self._clipboard):
            self.content_manager.add(copy.deepcopy(item))
            # 移动到正确位置
            current_pos = self.content_manager.count - 1
            target_pos = insert_pos + i
            while current_pos > target_pos:
                self.content_manager.move_up(current_pos)
                current_pos -= 1
        self._refresh_listbox()
        self._refresh_preview()
        return "break"
    
    # ========== 文件树 ==========
    
    def _load_file_tree(self) -> None:
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        files = self.content_manager.get_available_files()
        for category, name in CONTENT_CATEGORIES.items():
            cat_id = self.file_tree.insert("", tk.END, text=name, open=False)
            for _, filename in files.get(category, []):
                self.file_tree.insert(cat_id, tk.END, text=filename, values=(category, filename))
    
    def _reload_file_tree(self) -> None:
        """刷新文件列表（重新加载content目录）"""
        self.content_manager = ContentManager(self.content_dir)
        self._load_file_tree()
    
    # ========== 内容操作 ==========
    
    def _add_selected_file(self) -> None:
        selection = self.file_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要添加的文件")
            return
        
        self._save_state()
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
            self._save_state()
            self.content_manager.add(dialog.result)
            self._refresh_listbox()
            self._refresh_preview()
    
    def _add_image(self) -> None:
        dialog = ImageDialog(self.root, self.images_dir, self.theme)
        self.root.wait_window(dialog.top)
        
        if dialog.result:
            self._save_state()
            self.content_manager.add(dialog.result)
            self._refresh_listbox()
            self._refresh_preview()
    
    def _add_score(self) -> None:
        dialog = ScoreDialog(self.root, gabc_dir=self.gabc_dir, theme=self.theme)
        self.root.wait_window(dialog.top)
        
        if dialog.result:
            self._save_state()
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
        dialog = RuleTypeDialog(self.root, self.theme)
        self.root.wait_window(dialog.top)
        
        if dialog.result:
            self._save_state()
            self.content_manager.add(dialog.result)
            self._refresh_listbox()
            self._refresh_preview()
    
    def _add_pagebreak(self) -> None:
        self._save_state()
        self.content_manager.add(ContentItem("pagebreak", "", ""))
        self._refresh_listbox()
        self._refresh_preview()
    
    def _add_tocstart(self) -> None:
        self._save_state()
        self.content_manager.add(ContentItem("tocstart", "", ""))
        self._refresh_listbox()
        self._refresh_preview()
    
    def _add_singlecol(self) -> None:
        self._save_state()
        result = messagebox.askquestion("切换栏模式", "选择栏模式：\n\n是 = 进入单栏模式\n否 = 退出单栏模式")
        if result == 'yes':
            self.content_manager.add(ContentItem("singlecol_enter", "", ""))
        else:
            self.content_manager.add(ContentItem("singlecol_exit", "", ""))
        self._refresh_listbox()
        self._refresh_preview()
    
    # ========== 列表操作（支持多选）==========
    
    def _move_up(self) -> None:
        """上移选中项（支持多选）"""
        sel = list(self.content_listbox.curselection())
        if not sel or sel[0] == 0:
            return
        
        self._save_state()
        # 从上到下依次移动
        for idx in sel:
            self.content_manager.move_up(idx)
        
        self._refresh_listbox()
        # 重新选中（位置-1）
        for idx in sel:
            self.content_listbox.selection_set(idx - 1)
        self._refresh_preview()
    
    def _move_down(self) -> None:
        """下移选中项（支持多选）"""
        sel = list(self.content_listbox.curselection())
        if not sel or sel[-1] >= self.content_manager.count - 1:
            return
        
        self._save_state()
        # 从下到上依次移动
        for idx in reversed(sel):
            self.content_manager.move_down(idx)
        
        self._refresh_listbox()
        # 重新选中（位置+1）
        for idx in sel:
            self.content_listbox.selection_set(idx + 1)
        self._refresh_preview()
    
    def _edit_item(self) -> None:
        """编辑选中项（只编辑第一个）"""
        sel = self.content_listbox.curselection()
        if not sel:
            return
        
        idx = sel[0]
        item = self.content_manager.get(idx)
        
        if isinstance(item, MultiLineContentItem):
            messagebox.showinfo("提示", "多行文件内容无法直接编辑。")
            return
        
        dialog = CustomContentDialog(self.root, item, self.theme)
        self.root.wait_window(dialog.top)
        
        if dialog.result:
            self._save_state()
            self.content_manager.update(idx, dialog.result)
            self._refresh_listbox()
            self._refresh_preview()
    
    def _delete_item(self) -> None:
        """删除选中项（支持多选）"""
        sel = list(self.content_listbox.curselection())
        if not sel:
            return
        
        self._save_state()
        # 从后往前删除，避免索引错位
        for idx in reversed(sel):
            self.content_manager.remove(idx)
        
        self._refresh_listbox()
        self._refresh_preview()
    
    def _clear_all(self) -> None:
        if self.content_manager.is_empty():
            return
        
        if messagebox.askyesno("确认", "确定要清空所有内容吗？"):
            self._save_state()
            self.content_manager.clear()
            self._refresh_listbox()
            self._refresh_preview()
    
    # ========== UI刷新 ==========
    
    def _refresh_listbox(self) -> None:
        self.content_listbox.delete(0, tk.END)
        for item in self.content_manager.items:
            self.content_listbox.insert(tk.END, item.get_display_text())
    
    def _refresh_preview(self) -> None:
        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete(1.0, tk.END)
        try:
            text = self.content_manager.to_preview_text()
            self.preview_text.insert(tk.END, text)
        except Exception as e:
            logger.error(f"预览出错: {e}")
            self.preview_text.insert(tk.END, f"预览出错: {str(e)}")
        self.preview_text.config(state=tk.DISABLED)
    
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
        
        progress_dialog = CompileProgressDialog(self.root, self.theme)
        
        def progress(msg: str) -> None:
            progress_dialog.update_message(msg)
        
        def do_compile():
            try:
                output = self.compiler_service.compile_and_preview(
                    self.content_manager.items,
                    self.title_data,
                    progress
                )
                
                progress_dialog.close()
                
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
                progress_dialog.close()
                logger.exception("编译异常")
                messagebox.showerror("系统错误", str(e))
        
        self.root.after(100, do_compile)
    
    # ========== 模块操作 ==========
    
    def _add_hour_module(self) -> None:
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
            self._save_state()
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
        dialog = ImprimaturDialog(self.root, self.theme)
        result = dialog.show()
        
        if result:
            self._save_state()
            for item in result:
                self.content_manager.add(item)
            
            self._refresh_listbox()
            self._refresh_preview()
            messagebox.showinfo("成功", "签名块已添加")
