#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dialogs.py - 对话框组件
包含内容编辑对话框、标题页设置对话框等
Magnificat礼仪风格
"""

from __future__ import annotations
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional, Dict, List, Tuple, Any

from .theme import Theme, DEFAULT_THEME
from .models import ContentItem, TitlePageData, get_format_types_list
from .ui_components import LargeRadioButton


class BaseDialog:
    """对话框基类 - 统一的Magnificat风格"""
    
    def __init__(
        self, 
        parent: tk.Widget, 
        title: str,
        size: str = "600x400",
        theme: Theme = DEFAULT_THEME
    ):
        self.result: Optional[Any] = None
        self.theme = theme
        
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry(size)
        self.top.configure(bg=theme.BG_LIGHT)
        self.top.transient(parent)
        self.top.grab_set()
        
        # 设置窗口图标
        self._set_window_icon()
        
        # 居中显示
        self.top.update_idletasks()
        width = self.top.winfo_width()
        height = self.top.winfo_height()
        x = (self.top.winfo_screenwidth() // 2) - (width // 2)
        y = (self.top.winfo_screenheight() // 2) - (height // 2)
        self.top.geometry(f'+{x}+{y}')
    
    def _set_window_icon(self) -> None:
        """设置窗口图标"""
        try:
            from .icon import set_toplevel_icon
            set_toplevel_icon(self.top)
        except Exception:
            pass
    
    def _create_button_frame(self) -> tk.Frame:
        frame = tk.Frame(self.top, bg=self.theme.BG_LIGHT)
        frame.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=15)
        return frame
    
    def _add_button(
        self, 
        parent: tk.Frame, 
        text: str, 
        command: callable,
        bg: str,
        fg: str = None,
        side: str = tk.RIGHT
    ) -> tk.Label:
        if fg is None:
            c = bg.lstrip('#')
            if len(c) == 6:
                r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
                luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
                fg = "#FFFFFF" if luminance < 0.5 else self.theme.TEXT
            else:
                fg = self.theme.TEXT
        
        btn = tk.Label(
            parent, text=text, bg=bg, fg=fg,
            font=self.theme.get_font(), cursor='hand2', padx=20, pady=8
        )
        btn.pack(side=side, padx=5)
        btn.bind('<Button-1>', lambda e: command())
        btn.bind('<Enter>', lambda e: btn.config(bg=self.theme.darken(bg, 15)))
        btn.bind('<Leave>', lambda e: btn.config(bg=bg))
        return btn


class CustomContentDialog(BaseDialog):
    """自定义内容编辑对话框"""
    
    def __init__(
        self, 
        parent: tk.Widget, 
        item: Optional[ContentItem] = None,
        theme: Theme = DEFAULT_THEME
    ):
        super().__init__(
            parent, 
            "编辑内容" if item else "添加自定义内容",
            "700x550",
            theme
        )
        self.result: Optional[ContentItem] = None
        self._setup_ui(item)
    
    def _setup_ui(self, item: Optional[ContentItem]) -> None:
        theme = self.theme
        
        btn_frame = self._create_button_frame()
        self._add_button(btn_frame, "取消", self._cancel, theme.BG_HOVER)
        self._add_button(btn_frame, "确定", self._ok, theme.RUBRIC_RED)
        
        content = tk.Frame(self.top, bg=theme.BG_LIGHT)
        content.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        # 类型选择
        type_frame = tk.Frame(content, bg=theme.BG_LIGHT)
        type_frame.pack(fill=tk.X, padx=15, pady=(15, 10))
        
        tk.Label(
            type_frame, text="格式类型:", bg=theme.BG_LIGHT, fg=theme.RUBRIC_RED,
            font=theme.get_font("normal", bold=True)
        ).pack(side=tk.LEFT)
        
        self.type_var = tk.StringVar()
        format_types = get_format_types_list()
        
        self.type_combo = ttk.Combobox(
            type_frame, textvariable=self.type_var,
            values=[f"{t[0]} - {t[1]}" for t in format_types],
            width=50, font=theme.get_font(), state='readonly'
        )
        self.type_combo.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        self.type_combo.bind('<Button-1>', self._show_dropdown)
        
        # 拉丁文
        self._add_text_field(content, "拉丁文/路径:", "latin_text", 5)
        
        # 中文
        self._add_text_field(content, "中文:", "chinese_text", 5)
        
        # 附加参数
        arg_frame = tk.Frame(content, bg=theme.BG_LIGHT)
        arg_frame.pack(fill=tk.X, padx=15, pady=10)
        
        tk.Label(
            arg_frame, text="附加参数:", bg=theme.BG_LIGHT, fg=theme.TEXT,
            font=theme.get_font()
        ).pack(side=tk.LEFT)
        
        self.arg_entry = tk.Entry(
            arg_frame, width=35, bg=theme.BG_PANEL, fg=theme.TEXT,
            insertbackground=theme.TEXT, font=theme.get_font(), relief='flat'
        )
        self.arg_entry.pack(side=tk.LEFT, padx=10)
        
        tk.Label(
            arg_frame, text="(如对经编号等)", bg=theme.BG_LIGHT, fg=theme.TEXT_SEC,
            font=theme.get_font("small")
        ).pack(side=tk.LEFT)
        
        if item:
            self._populate(item, format_types)
    
    def _show_dropdown(self, event: tk.Event) -> None:
        self.type_combo.event_generate('<Down>')
    
    def _add_text_field(self, parent: tk.Frame, label: str, attr_name: str, height: int) -> None:
        theme = self.theme
        
        tk.Label(
            parent, text=label, bg=theme.BG_LIGHT, fg=theme.RUBRIC_RED,
            font=theme.get_font("normal", bold=True)
        ).pack(anchor='w', padx=15, pady=(10, 5))
        
        frame = tk.Frame(parent, bg=theme.BG_PANEL, highlightthickness=1, highlightbackground=theme.BORDER)
        frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))
        
        text = tk.Text(
            frame, height=height, wrap=tk.WORD, bg=theme.BG_PANEL, fg=theme.TEXT,
            insertbackground=theme.TEXT, font=theme.get_font(),
            borderwidth=0, padx=8, pady=8
        )
        text.pack(fill=tk.BOTH, expand=True)
        setattr(self, attr_name, text)
    
    def _populate(self, item: ContentItem, format_types: List[Tuple[str, str]]) -> None:
        for i, (t, _) in enumerate(format_types):
            if t == item.item_type:
                self.type_combo.current(i)
                break
        
        self.latin_text.insert(tk.END, item.latin)
        self.chinese_text.insert(tk.END, item.chinese)
        self.arg_entry.insert(0, item.arg)
    
    def _ok(self) -> None:
        type_str = self.type_var.get()
        if not type_str:
            messagebox.showwarning("提示", "请选择格式类型")
            return
        
        item_type = type_str.split(" - ")[0]
        self.result = ContentItem(
            item_type=item_type,
            latin=self.latin_text.get(1.0, tk.END).strip(),
            chinese=self.chinese_text.get(1.0, tk.END).strip(),
            arg=self.arg_entry.get().strip()
        )
        self.top.destroy()
    
    def _cancel(self) -> None:
        self.top.destroy()


class TitlePageDialog(BaseDialog):
    """标题页设置对话框"""
    
    FIELDS = [
        ("title_zh", "中文主标题:", "羅馬大日課\\\\[0.5em]耶穌聖誕瞻禮"),
        ("title_lat", "拉丁文标题:", "Breviarium Romanum\\\\[0.5em]In Nativitate Domini"),
        ("edition", "版本/编者:", "中拉對照\\\\[0.5em]Editio Sinico-Latina"),
        ("footer", "底部文字:", "Pro Manuscripto"),
    ]
    
    def __init__(
        self, 
        parent: tk.Widget, 
        initial_data: TitlePageData,
        theme: Theme = DEFAULT_THEME
    ):
        super().__init__(parent, "设置封面标题", "600x650", theme)
        self.result: Optional[TitlePageData] = None
        self.entries: Dict[str, tk.Text] = {}
        self._setup_ui(initial_data)
    
    def _setup_ui(self, initial_data: TitlePageData) -> None:
        theme = self.theme
        data_dict = initial_data.to_dict()
        
        tk.Label(
            self.top, text="设置 PDF 封面文本", bg=theme.BG_DARK, fg=theme.RUBRIC_RED,
            font=theme.get_font("title", bold=True)
        ).pack(pady=15)
        
        form = tk.Frame(self.top, bg=theme.BG_DARK)
        form.pack(fill=tk.BOTH, expand=True, padx=20)
        
        for key, label, default in self.FIELDS:
            self._add_field(form, key, label, data_dict.get(key, default))
        
        tk.Label(
            form, text="提示：使用 \\\\ 表示换行，\\\\[0.5em] 表示带间距换行",
            bg=theme.BG_DARK, fg=theme.TEXT_SEC, font=theme.get_font("small")
        ).pack(pady=10)
        
        btn_frame = self._create_button_frame()
        self._add_button(btn_frame, "取消", lambda: self.top.destroy(), theme.BG_HOVER)
        self._add_button(btn_frame, "保存设置", self._save, theme.RUBRIC_RED)
    
    def _add_field(self, parent: tk.Frame, key: str, label: str, value: str) -> None:
        theme = self.theme
        
        tk.Label(
            parent, text=label, bg=theme.BG_DARK, fg=theme.TEXT,
            font=theme.get_font()
        ).pack(anchor='w', pady=(8, 2))
        
        frame = tk.Frame(parent, bg=theme.BG_LIGHT, highlightthickness=1, highlightbackground=theme.BORDER)
        frame.pack(fill=tk.X, pady=(0, 5))
        
        text = tk.Text(
            frame, height=2, wrap=tk.WORD, bg=theme.BG_LIGHT, fg=theme.TEXT,
            insertbackground=theme.TEXT, font=theme.get_font(),
            borderwidth=0, padx=8, pady=5
        )
        text.pack(fill=tk.X)
        text.insert(tk.END, value)
        self.entries[key] = text
    
    def _save(self) -> None:
        self.result = TitlePageData(
            title_zh=self.entries["title_zh"].get(1.0, tk.END).strip(),
            title_lat=self.entries["title_lat"].get(1.0, tk.END).strip(),
            edition=self.entries["edition"].get(1.0, tk.END).strip(),
            footer=self.entries["footer"].get(1.0, tk.END).strip()
        )
        self.top.destroy()


class CompileErrorDialog(BaseDialog):
    """编译错误对话框"""
    
    def __init__(self, parent: tk.Widget, error_log: str, theme: Theme = DEFAULT_THEME):
        super().__init__(parent, "编译错误", "800x500", theme)
        self._setup_ui(error_log)
    
    def _setup_ui(self, error_log: str) -> None:
        theme = self.theme
        
        tk.Label(
            self.top, text="编译过程中出现错误", bg=theme.BG_DARK, fg=theme.DANGER,
            font=theme.get_font("title", bold=True)
        ).pack(pady=10)
        
        frame = tk.Frame(self.top, bg=theme.BG_LIGHT, highlightthickness=1, highlightbackground=theme.BORDER)
        frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        text = tk.Text(
            frame, wrap=tk.WORD, bg=theme.BG_LIGHT, fg=theme.TEXT,
            font=theme.get_mono_font(), borderwidth=0, padx=8, pady=8
        )
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scroll = tk.Scrollbar(frame, command=text.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        text.config(yscrollcommand=scroll.set)
        
        text.insert(tk.END, error_log)
        text.config(state=tk.DISABLED)
        
        btn_frame = self._create_button_frame()
        self._add_button(btn_frame, "关闭", lambda: self.top.destroy(), theme.RUBRIC_RED)




class CompileProgressDialog:
    """编译进度对话框 - 替代命令行输出"""
    
    def __init__(self, parent: tk.Widget, theme: Theme = DEFAULT_THEME):
        self.theme = theme
        self.cancelled = False
        
        self.top = tk.Toplevel(parent)
        self.top.title("☩ 正在编译...")
        self.top.geometry("500x350")
        self.top.configure(bg=theme.BG_LIGHT)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)
        
        # 设置窗口图标
        self._set_window_icon()
        
        # 居中
        self.top.update_idletasks()
        x = (self.top.winfo_screenwidth() // 2) - 250
        y = (self.top.winfo_screenheight() // 2) - 175
        self.top.geometry(f'+{x}+{y}')
        
        self._create_widgets()
    
    def _set_window_icon(self) -> None:
        try:
            from .icon import set_toplevel_icon
            set_toplevel_icon(self.top)
        except Exception:
            pass
    
    def _create_widgets(self) -> None:
        theme = self.theme
        
        main = tk.Frame(self.top, bg=theme.BG_LIGHT, padx=20, pady=20)
        main.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        tk.Label(
            main, text="☩ 正在编译文档", bg=theme.BG_LIGHT, fg=theme.RUBRIC_RED,
            font=theme.get_font("title", bold=True)
        ).pack(anchor='w', pady=(0, 15))
        
        # 当前状态
        self.status_label = tk.Label(
            main, text="准备中...", bg=theme.BG_LIGHT, fg=theme.TEXT,
            font=theme.get_font(), anchor='w'
        )
        self.status_label.pack(fill=tk.X, pady=(0, 10))
        
        # 进度条
        self.progress = ttk.Progressbar(main, mode='indeterminate', length=400)
        self.progress.pack(fill=tk.X, pady=(0, 15))
        self.progress.start(10)
        
        # 日志区域
        tk.Label(
            main, text="编译日志:", bg=theme.BG_LIGHT, fg=theme.TEXT_SEC,
            font=theme.get_font("small")
        ).pack(anchor='w')
        
        log_frame = tk.Frame(main, bg=theme.BORDER)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 15))
        
        self.log_text = tk.Text(
            log_frame, bg=theme.BG_PANEL, fg=theme.TEXT,
            font=theme.get_mono_font(), height=8,
            borderwidth=0, highlightthickness=0, padx=8, pady=8
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set, state=tk.DISABLED)
        
        # 取消按钮
        cancel_btn = tk.Label(
            main, text="取消", bg=theme.BG_HOVER, fg=theme.TEXT,
            font=theme.get_font(), cursor='hand2', padx=20, pady=8
        )
        cancel_btn.pack(side=tk.RIGHT)
        cancel_btn.bind('<Button-1>', lambda e: self._on_cancel())
        cancel_btn.bind('<Enter>', lambda e: cancel_btn.config(bg=theme.darken(theme.BG_HOVER, 10)))
        cancel_btn.bind('<Leave>', lambda e: cancel_btn.config(bg=theme.BG_HOVER))
    
    def update_message(self, message: str) -> None:
        """更新状态消息"""
        self.status_label.config(text=message)
        
        # 添加到日志
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"> {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        self.top.update()
    
    def _on_cancel(self) -> None:
        self.cancelled = True
        self.close()
    
    def close(self) -> None:
        try:
            self.progress.stop()
            self.top.destroy()
        except:
            pass


class LoadingDialog:
    """简单加载对话框"""
    
    def __init__(self, parent: tk.Widget, message: str = "处理中...", theme: Theme = DEFAULT_THEME):
        self.top = tk.Toplevel(parent)
        self.top.title("请稍候")
        self.top.geometry("300x100")
        self.top.configure(bg=theme.BG_LIGHT)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.overrideredirect(True)
        
        self.top.update_idletasks()
        x = (self.top.winfo_screenwidth() // 2) - 150
        y = (self.top.winfo_screenheight() // 2) - 50
        self.top.geometry(f'+{x}+{y}')
        
        self.label = tk.Label(
            self.top, text=message, font=theme.get_font(),
            bg=theme.BG_LIGHT, fg=theme.TEXT
        )
        self.label.pack(expand=True)
    
    def update_message(self, message: str) -> None:
        self.label.config(text=message)
        self.top.update()
    
    def close(self) -> None:
        self.top.destroy()



class ScoreDialog(BaseDialog):
    """乐谱编辑对话框 - 改进版大型选择按钮"""
    
    EXAMPLE_GABC = """name:Antiphona;
%%
(c4)Al(f)le(gf)lu(gh)ia.(g.) (::)"""
    
    def __init__(
        self, 
        parent: tk.Widget, 
        item: Optional[ContentItem] = None,
        gabc_dir: str = "",
        theme: Theme = DEFAULT_THEME
    ):
        super().__init__(parent, "添加乐谱 (Gregorio)", "750x600", theme)
        self.result: Optional[ContentItem] = None
        self.gabc_dir = gabc_dir
        self._setup_ui(item)
    
    def _setup_ui(self, item: Optional[ContentItem]) -> None:
        theme = self.theme
        
        btn_frame = self._create_button_frame()
        self._add_button(btn_frame, "取消", self._cancel, theme.BG_HOVER)
        self._add_button(btn_frame, "确定", self._ok, theme.RUBRIC_RED)
        
        content = tk.Frame(self.top, bg=theme.BG_DARK)
        content.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        # 模式选择 - 大型选择按钮
        tk.Label(
            content, text="选择乐谱来源", bg=theme.BG_DARK, fg=theme.RUBRIC_RED,
            font=theme.get_font("title", bold=True)
        ).pack(anchor='w', pady=(0, 10))
        
        mode_frame = tk.Frame(content, bg=theme.BG_DARK)
        mode_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.mode_var = tk.StringVar(value="file")
        
        # GABC文件按钮
        file_btn = LargeRadioButton(
            mode_frame, "GABC 文件", self.mode_var, "file", theme,
            "选择 .gabc 乐谱文件，编译时需要 GregorioTeX 支持"
        )
        file_btn.pack(side=tk.LEFT, padx=(0, 10), fill=tk.BOTH, expand=True)
        
        # 内联代码按钮
        inline_btn = LargeRadioButton(
            mode_frame, "内联 GABC 代码", self.mode_var, "inline", theme,
            "直接输入 GABC 代码，适合简短的乐谱片段"
        )
        inline_btn.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.mode_var.trace_add('write', lambda *args: self._on_mode_change())
        
        # 文件选择区 - 添加背景框架以与外围白色区分
        self.file_frame = tk.Frame(content, bg=theme.BG_DARK)
        
        # 添加一个带边框的容器来区分
        file_container = tk.Frame(self.file_frame, bg=theme.BG_PANEL, 
                                  highlightthickness=1, highlightbackground=theme.BORDER)
        file_container.pack(fill=tk.X, pady=(0, 10))
        
        file_inner = tk.Frame(file_container, bg=theme.BG_PANEL, padx=10, pady=10)
        file_inner.pack(fill=tk.X)
        
        tk.Label(
            file_inner, text="GABC文件:", bg=theme.BG_PANEL, fg=theme.TEXT,
            font=theme.get_font()
        ).pack(anchor='w', pady=(0, 5))
        
        file_row = tk.Frame(file_inner, bg=theme.BG_PANEL)
        file_row.pack(fill=tk.X)
        
        self.file_entry = tk.Entry(
            file_row, width=50, bg=theme.BG_LIGHT, fg=theme.TEXT,
            insertbackground=theme.TEXT, font=theme.get_font(), relief='flat',
            highlightthickness=1, highlightbackground=theme.BORDER
        )
        self.file_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        
        browse_btn = tk.Label(
            file_row, text="浏览...", bg=theme.RUBRIC_RED, fg="#FFFFFF",
            font=theme.get_font(), cursor='hand2', padx=15, pady=5
        )
        browse_btn.pack(side=tk.LEFT)
        browse_btn.bind('<Button-1>', lambda e: self._browse_file())
        
        # 内联代码区
        self.inline_frame = tk.Frame(content, bg=theme.BG_DARK)
        
        tk.Label(
            self.inline_frame, text="GABC代码:", bg=theme.BG_DARK, fg=theme.TEXT,
            font=theme.get_font()
        ).pack(anchor='w', pady=(0, 5))
        
        code_frame = tk.Frame(self.inline_frame, bg=theme.BG_LIGHT, 
                             highlightthickness=1, highlightbackground=theme.BORDER)
        code_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.code_text = tk.Text(
            code_frame, height=8, wrap=tk.NONE, bg=theme.BG_LIGHT, fg=theme.TEXT,
            insertbackground=theme.TEXT, font=theme.get_mono_font(),
            borderwidth=0, padx=8, pady=8
        )
        self.code_text.pack(fill=tk.BOTH, expand=True)
        
        example_btn = tk.Label(
            self.inline_frame, text="插入示例代码", bg=theme.BG_HOVER, fg=theme.TEXT,
            font=theme.get_font("small"), cursor='hand2', padx=10, pady=3
        )
        example_btn.pack(anchor='w')
        example_btn.bind('<Button-1>', lambda e: self._insert_example())
        
        # 标题
        title_frame = tk.Frame(content, bg=theme.BG_DARK)
        title_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(
            title_frame, text="乐谱标题(可选):", bg=theme.BG_DARK, fg=theme.TEXT,
            font=theme.get_font()
        ).pack(side=tk.LEFT)
        
        self.title_entry = tk.Entry(
            title_frame, width=40, bg=theme.BG_LIGHT, fg=theme.TEXT,
            insertbackground=theme.TEXT, font=theme.get_font(), relief='flat'
        )
        self.title_entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        # 提示
        tip_text = "提示：乐谱将自动切换到单栏显示，完成后恢复原来的分栏模式"
        tk.Label(
            content, text=tip_text, bg=theme.BG_DARK, fg=theme.TEXT_SEC,
            font=theme.get_font("small")
        ).pack(anchor='w', pady=5)
        
        self._on_mode_change()
        
        if item:
            self._populate(item)
    
    def _on_mode_change(self) -> None:
        if self.mode_var.get() == "file":
            self.inline_frame.pack_forget()
            self.file_frame.pack(fill=tk.X, pady=5)
        else:
            self.file_frame.pack_forget()
            self.inline_frame.pack(fill=tk.BOTH, expand=True, pady=5)
    
    def _browse_file(self) -> None:
        filepath = filedialog.askopenfilename(
            title="选择GABC乐谱文件",
            initialdir=self.gabc_dir,
            filetypes=[("GABC文件", "*.gabc"), ("所有文件", "*.*")]
        )
        if filepath:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, filepath)
    
    def _insert_example(self) -> None:
        self.code_text.delete(1.0, tk.END)
        self.code_text.insert(tk.END, self.EXAMPLE_GABC)
    
    def _populate(self, item: ContentItem) -> None:
        if item.arg == "inline":
            self.mode_var.set("inline")
            self._on_mode_change()
            self.code_text.insert(tk.END, item.latin)
        else:
            self.mode_var.set("file")
            self._on_mode_change()
            self.file_entry.insert(0, item.latin)
        
        if item.chinese:
            self.title_entry.insert(0, item.chinese)
    
    def _ok(self) -> None:
        mode = self.mode_var.get()
        
        if mode == "file":
            filepath = self.file_entry.get().strip()
            if not filepath:
                messagebox.showwarning("提示", "请选择GABC文件")
                return
            self.result = ContentItem(
                item_type="score",
                latin=filepath,
                chinese=self.title_entry.get().strip(),
                arg=""
            )
        else:
            code = self.code_text.get(1.0, tk.END).strip()
            if not code:
                messagebox.showwarning("提示", "请输入GABC代码")
                return
            self.result = ContentItem(
                item_type="score",
                latin=code,
                chinese=self.title_entry.get().strip(),
                arg="inline"
            )
        
        self.top.destroy()
    
    def _cancel(self) -> None:
        self.top.destroy()


class RuleTypeDialog:
    """分隔线类型选择对话框 - 统一风格"""
    
    def __init__(self, parent: tk.Widget, theme: Theme = DEFAULT_THEME):
        self.result: Optional[str] = None
        self.theme = theme
        
        self.top = tk.Toplevel(parent)
        self.top.title("选择分隔线类型")
        self.top.geometry("320x220")
        self.top.configure(bg=theme.BG_LIGHT)
        self.top.transient(parent)
        self.top.grab_set()
        
        # 居中
        self.top.update_idletasks()
        x = (self.top.winfo_screenwidth() // 2) - 160
        y = (self.top.winfo_screenheight() // 2) - 110
        self.top.geometry(f'+{x}+{y}')
        
        self._setup_ui()
    
    def _setup_ui(self) -> None:
        theme = self.theme
        
        tk.Label(
            self.top, text="请选择分隔线类型", bg=theme.BG_LIGHT, fg=theme.RUBRIC_RED,
            font=theme.get_font("title", bold=True)
        ).pack(pady=(20, 15))
        
        btn_frame = tk.Frame(self.top, bg=theme.BG_LIGHT)
        btn_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # 普通分隔线
        thin_btn = tk.Label(
            btn_frame, text="普通分隔线", bg=theme.BG_HOVER, fg=theme.TEXT,
            font=theme.get_font(), cursor='hand2', padx=15, pady=10
        )
        thin_btn.pack(fill=tk.X, pady=3)
        thin_btn.bind('<Button-1>', lambda e: self._select("thin"))
        thin_btn.bind('<Enter>', lambda e: thin_btn.config(bg=theme.darken(theme.BG_HOVER, 10)))
        thin_btn.bind('<Leave>', lambda e: thin_btn.config(bg=theme.BG_HOVER))
        
        # 粗分隔线
        thick_btn = tk.Label(
            btn_frame, text="粗分隔线", bg=theme.RUBRIC_RED, fg="#FFFFFF",
            font=theme.get_font(), cursor='hand2', padx=15, pady=10
        )
        thick_btn.pack(fill=tk.X, pady=3)
        thick_btn.bind('<Button-1>', lambda e: self._select("thick"))
        thick_btn.bind('<Enter>', lambda e: thick_btn.config(bg=theme.RUBRIC_DARK))
        thick_btn.bind('<Leave>', lambda e: thick_btn.config(bg=theme.RUBRIC_RED))
        
        # 取消
        cancel_btn = tk.Label(
            btn_frame, text="取消", bg=theme.BORDER, fg=theme.TEXT,
            font=theme.get_font(), cursor='hand2', padx=15, pady=10
        )
        cancel_btn.pack(fill=tk.X, pady=3)
        cancel_btn.bind('<Button-1>', lambda e: self._cancel())
        cancel_btn.bind('<Enter>', lambda e: cancel_btn.config(bg=theme.darken(theme.BORDER, 10)))
        cancel_btn.bind('<Leave>', lambda e: cancel_btn.config(bg=theme.BORDER))
    
    def _select(self, rule_type: str) -> None:
        self.result = rule_type
        self.top.destroy()
    
    def _cancel(self) -> None:
        self.result = None
        self.top.destroy()
    
    def show(self) -> Optional[str]:
        self.top.wait_window()
        return self.result


class ImageDialog(BaseDialog):
    """图片添加对话框"""
    
    def __init__(
        self, 
        parent: tk.Widget, 
        images_dir: str,
        theme: Theme = DEFAULT_THEME
    ):
        super().__init__(parent, "添加图片", "550x300", theme)
        self.result: Optional[ContentItem] = None
        self.images_dir = images_dir
        self._setup_ui()
    
    def _setup_ui(self) -> None:
        theme = self.theme
        
        # 先创建按钮区域，pack到BOTTOM（与其他对话框保持一致）
        btn_frame = self._create_button_frame()
        self._add_button(btn_frame, "取消", self._cancel, theme.BG_HOVER)
        self._add_button(btn_frame, "确定", self._ok, theme.RUBRIC_RED)
        
        # 然后创建内容区域
        content = tk.Frame(self.top, bg=theme.BG_LIGHT)
        content.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # 标题区域 - 使用深色背景
        title_frame = tk.Frame(content, bg=theme.BG_DARK, padx=10, pady=10)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(
            title_frame, text="选择图片文件", bg=theme.BG_DARK, fg=theme.RUBRIC_RED,
            font=theme.get_font("title", bold=True)
        ).pack(anchor='w')
        
        tk.Label(
            title_frame, text="图片必须放在项目的 images 文件夹下", 
            bg=theme.BG_DARK, fg=theme.TEXT_SEC, font=theme.get_font("small")
        ).pack(anchor='w', pady=(5, 0))
        
        # 文件选择区域
        file_frame = tk.Frame(content, bg=theme.BG_LIGHT)
        file_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(
            file_frame, text="文件名:", bg=theme.BG_LIGHT, fg=theme.TEXT,
            font=theme.get_font()
        ).pack(side=tk.LEFT)
        
        self.file_entry = tk.Entry(
            file_frame, width=35, bg=theme.BG_PANEL, fg=theme.TEXT,
            insertbackground=theme.TEXT, font=theme.get_font(), relief='flat',
            highlightthickness=1, highlightbackground=theme.BORDER
        )
        self.file_entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        browse_btn = tk.Label(
            file_frame, text="浏览", bg=theme.RUBRIC_RED, fg="#FFFFFF",
            font=theme.get_font(), cursor='hand2', padx=15, pady=5
        )
        browse_btn.pack(side=tk.LEFT)
        browse_btn.bind('<Button-1>', lambda e: self._browse())
        browse_btn.bind('<Enter>', lambda e: browse_btn.config(bg=theme.RUBRIC_DARK))
        browse_btn.bind('<Leave>', lambda e: browse_btn.config(bg=theme.RUBRIC_RED))
        
        # 示例提示
        tk.Label(
            content, text="示例: images/cover.png", bg=theme.BG_LIGHT, fg=theme.TEXT_SEC,
            font=theme.get_font("small")
        ).pack(anchor='w')
        
    def _browse(self) -> None:
        filepath = filedialog.askopenfilename(
            title="选择图片",
            initialdir=self.images_dir,
            filetypes=[
                ("图片文件", "*.png *.jpg *.jpeg *.pdf"),
                ("所有文件", "*.*")
            ]
        )
        if filepath:
            import os
            # 尝试获取相对路径
            try:
                rel_path = os.path.relpath(filepath, os.path.dirname(self.images_dir)).replace('\\', '/')
                self.file_entry.delete(0, tk.END)
                self.file_entry.insert(0, rel_path)
            except ValueError:
                self.file_entry.delete(0, tk.END)
                self.file_entry.insert(0, filepath.replace('\\', '/'))
    
    def _ok(self) -> None:
        filepath = self.file_entry.get().strip().replace('\\', '/')
        if not filepath:
            messagebox.showwarning("提示", "请输入或选择图片文件")
            return
        
        self.result = ContentItem(
            item_type="image",
            latin=filepath,
            chinese="",
            arg=""
        )
        self.top.destroy()
    
    def _cancel(self) -> None:
        self.top.destroy()
