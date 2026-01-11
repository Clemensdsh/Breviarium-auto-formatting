#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
module_dialogs.py - 时辰模块配置对话框
提供模块选择、配置和编辑的UI界面
"""

from __future__ import annotations
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from typing import Optional, Dict, List, Callable, Any, TYPE_CHECKING

from .modules import (
    HourType, HourModule, ModuleConfig,
    create_hour_module, HOUR_TYPE_NAMES, COMPONENT_TYPE_NAMES,
    BaseComponent, ModuleSlot, PsalmWithAntiphonComponent, 
    LessonWithResponsoryComponent
)
from .models import ContentItem, ModuleItem
from .theme import Theme

if TYPE_CHECKING:
    from .app import PsalterApp


class ModuleSelectionDialog:
    """模块选择对话框 - 选择要添加的时辰类型"""
    
    def __init__(self, parent: tk.Tk, theme: Theme):
        self.parent = parent
        self.theme = theme
        self.result: Optional[HourType] = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("选择时辰模块")
        self.dialog.geometry("400x550")  # 增加高度
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.configure(bg=theme.BG_LIGHT)
        
        # 设置窗口图标
        self._set_window_icon()
        
        self._create_widgets()
        self._center_dialog()
    
    def _set_window_icon(self) -> None:
        """设置窗口图标"""
        try:
            import os, sys
            if hasattr(sys, '_MEIPASS'):
                # PyInstaller 打包后，资源在临时目录 _MEIPASS 中
                base_dir = sys._MEIPASS
            else:
                base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            
            icon_paths = [
                os.path.join(base_dir, 'cross_of_saint_stephen.png'),
                os.path.join(base_dir, 'Cross_of_saint_stephen.png'),
                os.path.join(base_dir, 'maltese_cross.ico'),
                os.path.join(base_dir, 'maltese_cross.png'),
            ]
            
            for icon_path in icon_paths:
                if os.path.exists(icon_path):
                    if icon_path.endswith('.ico'):
                        self.dialog.iconbitmap(icon_path)
                    else:
                        photo = tk.PhotoImage(file=icon_path)
                        self.dialog.iconphoto(False, photo)
                    break
        except Exception:
            pass
    
    def _create_widgets(self) -> None:
        theme = self.theme
        
        # 主框架
        main = tk.Frame(self.dialog, bg=theme.BG_LIGHT, padx=15, pady=15)
        main.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        tk.Label(
            main, text="选择要添加的时辰", bg=theme.BG_LIGHT, fg=theme.RUBRIC_RED,
            font=theme.get_font("title", bold=True)
        ).pack(anchor='w')
        
        tk.Label(
            main, text="请选择一个时辰类型，将创建对应的模块结构",
            bg=theme.BG_LIGHT, fg=theme.TEXT_SEC, font=theme.get_font("small")
        ).pack(anchor='w', pady=(0, 10))
        
        # 时辰列表框架（带滚动条）
        list_frame = tk.Frame(main, bg=theme.BG_PANEL, highlightthickness=1, 
                             highlightbackground=theme.BORDER)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.hour_listbox = tk.Listbox(
            list_frame, font=theme.get_font(), selectmode=tk.SINGLE,
            bg=theme.BG_PANEL, fg=theme.TEXT, selectbackground=theme.RUBRIC_RED,
            selectforeground="#FFFFFF", borderwidth=0, highlightthickness=0,
            yscrollcommand=scrollbar.set
        )
        self.hour_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar.config(command=self.hour_listbox.yview)
        
        # 填充列表
        for hour_type in HourType:
            display_name = HOUR_TYPE_NAMES.get(hour_type.value, hour_type.value)
            self.hour_listbox.insert(tk.END, display_name)
        
        self.hour_listbox.bind('<Double-1>', lambda e: self._on_select())
        
        # 描述区域
        desc_frame = tk.Frame(main, bg=theme.BG_DARK, padx=10, pady=10)
        desc_frame.pack(fill=tk.X, pady=(10, 0))
        
        tk.Label(
            desc_frame, text="说明", bg=theme.BG_DARK, fg=theme.RUBRIC_RED,
            font=theme.get_font("normal", bold=True)
        ).pack(anchor='w')
        
        self.desc_label = tk.Label(
            desc_frame, text="选择一个时辰以查看说明",
            bg=theme.BG_DARK, fg=theme.TEXT, font=theme.get_font(),
            wraplength=350, justify=tk.LEFT
        )
        self.desc_label.pack(anchor='w', pady=(5, 0))
        
        self.hour_listbox.bind('<<ListboxSelect>>', self._on_listbox_select)
        
        # 按钮区域
        btn_frame = tk.Frame(self.dialog, bg=theme.BG_LIGHT, pady=10)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        # 取消按钮
        cancel_btn = tk.Label(
            btn_frame, text="取消", bg=theme.BG_HOVER, fg=theme.TEXT,
            font=theme.get_font(), cursor='hand2', padx=20, pady=8
        )
        cancel_btn.pack(side=tk.RIGHT, padx=10)
        cancel_btn.bind('<Button-1>', lambda e: self._on_cancel())
        cancel_btn.bind('<Enter>', lambda e: cancel_btn.config(bg=theme.darken(theme.BG_HOVER, 15)))
        cancel_btn.bind('<Leave>', lambda e: cancel_btn.config(bg=theme.BG_HOVER))
        
        # 确定按钮
        ok_btn = tk.Label(
            btn_frame, text="确定", bg=theme.RUBRIC_RED, fg="#FFFFFF",
            font=theme.get_font(), cursor='hand2', padx=20, pady=8
        )
        ok_btn.pack(side=tk.RIGHT, padx=5)
        ok_btn.bind('<Button-1>', lambda e: self._on_select())
        ok_btn.bind('<Enter>', lambda e: ok_btn.config(bg=theme.RUBRIC_DARK))
        ok_btn.bind('<Leave>', lambda e: ok_btn.config(bg=theme.RUBRIC_RED))
    
    def _on_listbox_select(self, event) -> None:
        selection = self.hour_listbox.curselection()
        if selection:
            hour_type = list(HourType)[selection[0]]
            descriptions = {
                # 1962年旧礼
                HourType.MATINS_1962: "【旧礼】夜课经是日课中最长的部分，包含三个夜课，每个夜课有圣咏和读经。",
                HourType.LAUDS_1962: "【旧礼】赞美经是黎明时的祈祷，包含5首圣咏和赞主曲。",
                HourType.PRIME_1962: "【旧礼】一时经是早晨六点的祈祷，是小时课之一。",
                HourType.TERCE_1962: "【旧礼】三时经是上午九点的祈祷，纪念圣神降临。",
                HourType.SEXT_1962: "【旧礼】六时经是正午的祈祷，纪念耶稳被钉十字架。",
                HourType.NONE_1962: "【旧礼】九时经是下午三点的祈祷，纪念耶稳在十字架上断气。",
                HourType.VESPERS_1962: "【旧礼】晚祷是日落时的祈祷，包含5首圣咏和圣母赞主曲。",
                HourType.COMPLINE_1962: "【旧礼】夜祷是睡前的祈祷，以西默盎赞主曲和圣母对经结束。",
                # 新礼 LOTH
                HourType.OFFICE_OF_READINGS: "【新礼】诵读是可在任何时间诵念的祈祷，包含3首圣咏和2篇读经。",
                HourType.LAUDS_LOTH: "【新礼】晨祷是清晨的祈祷，包含圣咏、旧约圣歌、赞美圣咏和赞主曲。",
                HourType.TERCE_LOTH: "【新礼】午前祈祷约在上午9时诵念，是日间祈祷之一。",
                HourType.SEXT_LOTH: "【新礼】午时祈祷约在正午诵念，是日间祈祷之一。",
                HourType.NONE_LOTH: "【新礼】午后祈祷约在下午3时诵念，是日间祈祷之一。",
                HourType.VESPERS_LOTH: "【新礼】晚祷是傍晚的祈祷，包含2首圣咏、新约圣歌和圣母赞主曲。",
                HourType.COMPLINE_LOTH: "【新礼】夜祷是睡前的祈祷，以西默盎赞主曲和圣母对经结束。",
            }
            self.desc_label.config(text=descriptions.get(hour_type, ""))
    
    def _on_select(self) -> None:
        selection = self.hour_listbox.curselection()
        if selection:
            self.result = list(HourType)[selection[0]]
            self.dialog.destroy()
        else:
            messagebox.showwarning("提示", "请先选择一个时辰类型")
    
    def _on_cancel(self) -> None:
        self.result = None
        self.dialog.destroy()
    
    def _center_dialog(self) -> None:
        self.dialog.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() - self.dialog.winfo_width()) // 2
        y = self.parent.winfo_y() + (self.parent.winfo_height() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")
    
    def show(self) -> Optional[HourType]:
        self.dialog.wait_window()
        return self.result


class ModuleConfigDialog:
    """模块配置对话框 - 配置时辰参数（统一风格）"""
    
    def __init__(self, parent: tk.Tk, hour_type: HourType, 
                 theme: Theme, existing_config: Optional[ModuleConfig] = None):
        self.parent = parent
        self.hour_type = hour_type
        self.theme = theme
        self.config = existing_config or ModuleConfig()
        self.result: Optional[ModuleConfig] = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"配置 - {HOUR_TYPE_NAMES.get(hour_type.value, hour_type.value)}")
        self.dialog.geometry("500x520")  # 增加高度以容纳所有配置项
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.configure(bg=theme.BG_LIGHT)
        
        # 设置窗口图标
        self._set_window_icon()
        
        self._create_widgets()
        self._center_dialog()
    
    def _set_window_icon(self) -> None:
        """设置窗口图标"""
        try:
            from .icon import set_toplevel_icon
            set_toplevel_icon(self.dialog)
        except Exception:
            pass
    
    def _create_widgets(self) -> None:
        theme = self.theme
        
        # 主框架 - 使用统一背景色
        main_frame = tk.Frame(self.dialog, bg=theme.BG_LIGHT, padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题信息区
        info_frame = tk.Frame(main_frame, bg=theme.BG_DARK, padx=10, pady=10)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(
            info_frame, text="基本信息", bg=theme.BG_DARK, fg=theme.RUBRIC_RED,
            font=theme.get_font("title", bold=True)
        ).pack(anchor='w', pady=(0, 8))
        
        # 拉丁文标题
        row1 = tk.Frame(info_frame, bg=theme.BG_DARK)
        row1.pack(fill=tk.X, pady=2)
        tk.Label(row1, text="拉丁文标题:", bg=theme.BG_DARK, fg=theme.TEXT,
                font=theme.get_font()).pack(side=tk.LEFT)
        self.title_lat_var = tk.StringVar(value=self.config.title_lat)
        tk.Entry(row1, textvariable=self.title_lat_var, width=35, 
                bg=theme.BG_PANEL, fg=theme.TEXT, insertbackground=theme.TEXT,
                font=theme.get_font(), relief='flat').pack(side=tk.LEFT, padx=10)
        
        # 中文标题
        row2 = tk.Frame(info_frame, bg=theme.BG_DARK)
        row2.pack(fill=tk.X, pady=2)
        tk.Label(row2, text="中文标题:", bg=theme.BG_DARK, fg=theme.TEXT,
                font=theme.get_font()).pack(side=tk.LEFT)
        self.title_zh_var = tk.StringVar(value=self.config.title_zh)
        tk.Entry(row2, textvariable=self.title_zh_var, width=35, 
                bg=theme.BG_PANEL, fg=theme.TEXT, insertbackground=theme.TEXT,
                font=theme.get_font(), relief='flat').pack(side=tk.LEFT, padx=10)
        
        # 根据时辰类型显示不同的配置选项
        # 夜课类型
        if self.hour_type == HourType.MATINS_1962:
            self._create_matins_config(main_frame)
        # 小时课类型（旧礼和新礼）
        elif self.hour_type in [
            HourType.PRIME_1962, HourType.TERCE_1962, HourType.SEXT_1962, HourType.NONE_1962,
            HourType.TERCE_LOTH, HourType.SEXT_LOTH, HourType.NONE_LOTH
        ]:
            self._create_small_hour_config(main_frame)
        else:
            self._create_general_config(main_frame)
        
        # 显示选项区
        options_frame = tk.Frame(main_frame, bg=theme.BG_DARK, padx=10, pady=10)
        options_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            options_frame, text="显示选项", bg=theme.BG_DARK, fg=theme.RUBRIC_RED,
            font=theme.get_font("title", bold=True)
        ).pack(anchor='w', pady=(0, 8))
        
        self.show_gloria_var = tk.BooleanVar(value=self.config.show_gloria)
        gloria_cb = tk.Checkbutton(
            options_frame, text="显示圣三光荣颂",
            variable=self.show_gloria_var,
            bg=theme.BG_DARK, fg=theme.TEXT, selectcolor=theme.BG_PANEL,
            activebackground=theme.BG_DARK, activeforeground=theme.TEXT,
            font=theme.get_font()
        )
        gloria_cb.pack(anchor=tk.W)
        
        self.show_antiphon_var = tk.BooleanVar(value=self.config.show_antiphon_repeat)
        antiphon_cb = tk.Checkbutton(
            options_frame, text="圣咏后重复对经",
            variable=self.show_antiphon_var,
            bg=theme.BG_DARK, fg=theme.TEXT, selectcolor=theme.BG_PANEL,
            activebackground=theme.BG_DARK, activeforeground=theme.TEXT,
            font=theme.get_font()
        )
        antiphon_cb.pack(anchor=tk.W)
        
        # 按钮区域
        btn_frame = tk.Frame(self.dialog, bg=theme.BG_LIGHT, pady=10)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        # 取消按钮
        cancel_btn = tk.Label(
            btn_frame, text="取消", bg=theme.BG_HOVER, fg=theme.TEXT,
            font=theme.get_font(), cursor='hand2', padx=20, pady=8
        )
        cancel_btn.pack(side=tk.RIGHT, padx=10)
        cancel_btn.bind('<Button-1>', lambda e: self._on_cancel())
        cancel_btn.bind('<Enter>', lambda e: cancel_btn.config(bg=theme.darken(theme.BG_HOVER, 15)))
        cancel_btn.bind('<Leave>', lambda e: cancel_btn.config(bg=theme.BG_HOVER))
        
        # 确定按钮
        ok_btn = tk.Label(
            btn_frame, text="确定", bg=theme.RUBRIC_RED, fg="#FFFFFF",
            font=theme.get_font(), cursor='hand2', padx=20, pady=8
        )
        ok_btn.pack(side=tk.RIGHT, padx=5)
        ok_btn.bind('<Button-1>', lambda e: self._on_ok())
        ok_btn.bind('<Enter>', lambda e: ok_btn.config(bg=theme.RUBRIC_DARK))
        ok_btn.bind('<Leave>', lambda e: ok_btn.config(bg=theme.RUBRIC_RED))
    
    def _create_matins_config(self, parent: tk.Frame) -> None:
        """创建夜课特定配置（统一风格）"""
        theme = self.theme
        config_frame = tk.Frame(parent, bg=theme.BG_DARK, padx=10, pady=10)
        config_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            config_frame, text="夜课配置", bg=theme.BG_DARK, fg=theme.RUBRIC_RED,
            font=theme.get_font("title", bold=True)
        ).pack(anchor='w', pady=(0, 8))
        
        # 夜课数量
        row1 = tk.Frame(config_frame, bg=theme.BG_DARK)
        row1.pack(fill=tk.X, pady=2)
        tk.Label(row1, text="夜课数量:", bg=theme.BG_DARK, fg=theme.TEXT,
                font=theme.get_font()).pack(side=tk.LEFT)
        self.nocturn_count_var = tk.IntVar(value=self.config.nocturn_count)
        nocturn_combo = ttk.Combobox(
            row1, textvariable=self.nocturn_count_var,
            values=[1, 3], state="readonly", width=10
        )
        nocturn_combo.pack(side=tk.LEFT, padx=10)
        
        # 每夜课圣咏数量
        row2 = tk.Frame(config_frame, bg=theme.BG_DARK)
        row2.pack(fill=tk.X, pady=2)
        tk.Label(row2, text="每夜课圣咏数:", bg=theme.BG_DARK, fg=theme.TEXT,
                font=theme.get_font()).pack(side=tk.LEFT)
        self.psalms_per_var = tk.IntVar(value=self.config.psalms_per_nocturn)
        ttk.Spinbox(row2, from_=1, to=6, textvariable=self.psalms_per_var, width=10).pack(side=tk.LEFT, padx=10)
        
        # 每夜课读经数量
        row3 = tk.Frame(config_frame, bg=theme.BG_DARK)
        row3.pack(fill=tk.X, pady=2)
        tk.Label(row3, text="每夜课读经数:", bg=theme.BG_DARK, fg=theme.TEXT,
                font=theme.get_font()).pack(side=tk.LEFT)
        self.lessons_per_var = tk.IntVar(value=self.config.lessons_per_nocturn)
        ttk.Spinbox(row3, from_=1, to=6, textvariable=self.lessons_per_var, width=10).pack(side=tk.LEFT, padx=10)
    
    def _create_small_hour_config(self, parent: tk.Frame) -> None:
        """创建小时课配置（统一风格）"""
        theme = self.theme
        config_frame = tk.Frame(parent, bg=theme.BG_DARK, padx=10, pady=10)
        config_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            config_frame, text="小时课配置", bg=theme.BG_DARK, fg=theme.RUBRIC_RED,
            font=theme.get_font("title", bold=True)
        ).pack(anchor='w', pady=(0, 8))
        
        row1 = tk.Frame(config_frame, bg=theme.BG_DARK)
        row1.pack(fill=tk.X, pady=2)
        tk.Label(row1, text="圣咏数量:", bg=theme.BG_DARK, fg=theme.TEXT,
                font=theme.get_font()).pack(side=tk.LEFT)
        self.psalm_count_var = tk.IntVar(value=self.config.psalm_count)
        ttk.Spinbox(row1, from_=1, to=6, textvariable=self.psalm_count_var, width=10).pack(side=tk.LEFT, padx=10)
    
    def _create_general_config(self, parent: tk.Frame) -> None:
        """创建通用配置（赞美经、晚祷、夜祷）（统一风格）"""
        theme = self.theme
        config_frame = tk.Frame(parent, bg=theme.BG_DARK, padx=10, pady=10)
        config_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            config_frame, text="结构配置", bg=theme.BG_DARK, fg=theme.RUBRIC_RED,
            font=theme.get_font("title", bold=True)
        ).pack(anchor='w', pady=(0, 8))
        
        tk.Label(
            config_frame, text="此时辰使用固定结构，包含5首圣咏和相关祷文。",
            bg=theme.BG_DARK, fg=theme.TEXT, font=theme.get_font()
        ).pack(anchor='w')
    
    def _on_ok(self) -> None:
        # 收集配置
        self.config.title_lat = self.title_lat_var.get()
        self.config.title_zh = self.title_zh_var.get()
        self.config.show_gloria = self.show_gloria_var.get()
        self.config.show_antiphon_repeat = self.show_antiphon_var.get()
        
        if self.hour_type == HourType.MATINS_1962:
            self.config.nocturn_count = self.nocturn_count_var.get()
            self.config.psalms_per_nocturn = self.psalms_per_var.get()
            self.config.lessons_per_nocturn = self.lessons_per_var.get()
        elif hasattr(self, 'psalm_count_var'):
            self.config.psalm_count = self.psalm_count_var.get()
        
        self.result = self.config
        self.dialog.destroy()
    
    def _on_cancel(self) -> None:
        self.result = None
        self.dialog.destroy()
    
    def _center_dialog(self) -> None:
        self.dialog.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() - self.dialog.winfo_width()) // 2
        y = self.parent.winfo_y() + (self.parent.winfo_height() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")
    
    def show(self) -> Optional[ModuleConfig]:
        self.dialog.wait_window()
        return self.result


class SlotEditDialog:
    """插槽编辑对话框 - 编辑单个插槽的内容（统一风格，支持类型选择）"""
    
    def __init__(self, parent: tk.Tk, slot: ModuleSlot, theme: Theme, allow_type_change: bool = True):
        self.parent = parent
        self.slot = slot
        self.theme = theme
        self.allow_type_change = allow_type_change
        self.result: Optional[ContentItem] = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"编辑 - {slot.label_zh}")
        self.dialog.geometry("650x600")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.configure(bg=theme.BG_LIGHT)
        
        # 设置窗口图标
        self._set_window_icon()
        
        self._create_widgets()
        self._center_dialog()
    
    def _set_window_icon(self) -> None:
        """设置窗口图标"""
        try:
            from .icon import set_toplevel_icon
            set_toplevel_icon(self.dialog)
        except Exception:
            pass
    
    def _create_widgets(self) -> None:
        theme = self.theme
        
        main_frame = tk.Frame(self.dialog, bg=theme.BG_LIGHT, padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 类型选择（可选）
        type_frame = tk.Frame(main_frame, bg=theme.BG_DARK, padx=10, pady=8)
        type_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(
            type_frame, text="类型:",
            bg=theme.BG_DARK, fg=theme.RUBRIC_RED,
            font=theme.get_font("normal", bold=True)
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        # 获取格式类型列表
        from .models import FORMAT_TYPE_NAMES
        self.format_types = [(k, f"{k} - {v}") for k, v in FORMAT_TYPE_NAMES.items()]
        
        self.type_var = tk.StringVar()
        
        if self.allow_type_change:
            # 下拉框选择类型
            self.type_combo = ttk.Combobox(
                type_frame, textvariable=self.type_var,
                values=[display for _, display in self.format_types],
                state='readonly', width=35
            )
            self.type_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            # 设置默认值
            default_type = self.slot.slot_type
            for i, (type_key, _) in enumerate(self.format_types):
                if type_key == default_type:
                    self.type_combo.current(i)
                    break
            else:
                # 如果没找到匹配的，选择text
                for i, (type_key, _) in enumerate(self.format_types):
                    if type_key == "text":
                        self.type_combo.current(i)
                        break
        else:
            # 只读显示
            tk.Label(
                type_frame,
                text=f"{self.slot.slot_type} | {self.slot.label_lat}",
                bg=theme.BG_DARK, fg=theme.TEXT_SEC,
                font=theme.get_font()
            ).pack(side=tk.LEFT)
        
        # 拉丁文
        tk.Label(
            main_frame, text="拉丁文", 
            bg=theme.BG_LIGHT, fg=theme.RUBRIC_RED,
            font=theme.get_font("normal", bold=True)
        ).pack(anchor=tk.W, pady=(5, 3))
        
        lat_container = tk.Frame(main_frame, bg=theme.BORDER)
        lat_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.latin_text = scrolledtext.ScrolledText(
            lat_container, height=5, wrap=tk.WORD,
            bg=theme.BG_PANEL, fg=theme.TEXT,
            font=theme.get_font(), relief='flat',
            padx=8, pady=8
        )
        self.latin_text.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        # 中文
        tk.Label(
            main_frame, text="中文", 
            bg=theme.BG_LIGHT, fg=theme.RUBRIC_RED,
            font=theme.get_font("normal", bold=True)
        ).pack(anchor=tk.W, pady=(5, 3))
        
        zh_container = tk.Frame(main_frame, bg=theme.BORDER)
        zh_container.pack(fill=tk.BOTH, expand=True)
        
        self.chinese_text = scrolledtext.ScrolledText(
            zh_container, height=5, wrap=tk.WORD,
            bg=theme.BG_PANEL, fg=theme.TEXT,
            font=theme.get_font(), relief='flat',
            padx=8, pady=8
        )
        self.chinese_text.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        # 填充现有内容
        content = self.slot.get_content()
        if content:
            self.latin_text.insert(tk.END, content.latin)
            self.chinese_text.insert(tk.END, content.chinese)
            # 如果有内容，使用内容的类型
            if self.allow_type_change and hasattr(content, 'item_type'):
                for i, (type_key, _) in enumerate(self.format_types):
                    if type_key == content.item_type:
                        self.type_combo.current(i)
                        break
        
        # 按钮
        btn_frame = tk.Frame(self.dialog, bg=theme.BG_LIGHT)
        btn_frame.pack(fill=tk.X, padx=15, pady=(0, 15))
        
        # 清空按钮
        clear_btn = tk.Label(
            btn_frame, text="清空", 
            bg=theme.BG_HOVER, fg=theme.TEXT,
            font=theme.get_font(), cursor='hand2', padx=15, pady=8
        )
        clear_btn.pack(side=tk.LEFT)
        clear_btn.bind('<Button-1>', lambda e: self._on_clear())
        clear_btn.bind('<Enter>', lambda e: clear_btn.config(bg=theme.darken(theme.BG_HOVER, 10)))
        clear_btn.bind('<Leave>', lambda e: clear_btn.config(bg=theme.BG_HOVER))
        
        # 确定按钮
        ok_btn = tk.Label(
            btn_frame, text="确定", 
            bg=theme.RUBRIC_RED, fg="#FFFFFF",
            font=theme.get_font(), cursor='hand2', padx=20, pady=8
        )
        ok_btn.pack(side=tk.RIGHT, padx=(5, 0))
        ok_btn.bind('<Button-1>', lambda e: self._on_ok())
        ok_btn.bind('<Enter>', lambda e: ok_btn.config(bg=theme.RUBRIC_DARK))
        ok_btn.bind('<Leave>', lambda e: ok_btn.config(bg=theme.RUBRIC_RED))
        
        # 取消按钮
        cancel_btn = tk.Label(
            btn_frame, text="取消", 
            bg=theme.BG_HOVER, fg=theme.TEXT,
            font=theme.get_font(), cursor='hand2', padx=20, pady=8
        )
        cancel_btn.pack(side=tk.RIGHT)
        cancel_btn.bind('<Button-1>', lambda e: self._on_cancel())
        cancel_btn.bind('<Enter>', lambda e: cancel_btn.config(bg=theme.darken(theme.BG_HOVER, 10)))
        cancel_btn.bind('<Leave>', lambda e: cancel_btn.config(bg=theme.BG_HOVER))
    
    def _on_ok(self) -> None:
        latin = self.latin_text.get("1.0", tk.END).strip()
        chinese = self.chinese_text.get("1.0", tk.END).strip()
        
        if latin or chinese:
            # 获取选择的类型
            if self.allow_type_change and hasattr(self, 'type_combo'):
                type_str = self.type_var.get()
                if type_str and " - " in type_str:
                    item_type = type_str.split(" - ")[0]
                else:
                    item_type = self.slot.slot_type
            else:
                # 根据slot类型确定item_type
                type_mapping = {
                    "antiphon": "antiphon",
                    "psalm": "verse",
                    "psalm_title": "psalmtitle",
                    "lesson_title": "lesson",
                    "text": "text",
                    "versicle": "V",
                    "responsory": "text",
                    "rubric": "rubric",
                    "gloria": "gloria",
                    "preces": "text",
                    "verse": "verse",
                }
                item_type = type_mapping.get(self.slot.slot_type, "text")
            
            self.result = ContentItem(
                item_type=item_type,
                latin=latin,
                chinese=chinese
            )
        else:
            self.result = None
        
        self.dialog.destroy()
    
    def _on_cancel(self) -> None:
        self.result = None
        self.dialog.destroy()
    
    def _on_clear(self) -> None:
        self.latin_text.delete("1.0", tk.END)
        self.chinese_text.delete("1.0", tk.END)
    
    def _center_dialog(self) -> None:
        self.dialog.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() - self.dialog.winfo_width()) // 2
        y = self.parent.winfo_y() + (self.parent.winfo_height() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")
    
    def show(self) -> Optional[ContentItem]:
        self.dialog.wait_window()
        return self.result


class ModuleEditorDialog:
    """模块编辑器对话框 - 编辑整个模块的内容（统一风格）"""
    
    def __init__(self, parent: tk.Tk, module: HourModule, theme: Theme):
        self.parent = parent
        self.module = module
        self.theme = theme
        self.result: Optional[HourModule] = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"编辑模块 - {module.name_zh}")
        self.dialog.geometry("950x750")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.configure(bg=theme.BG_LIGHT)
        
        # 设置窗口图标
        self._set_window_icon()
        
        self._create_widgets()
        self._populate_tree()
        self._center_dialog()
    
    def _set_window_icon(self) -> None:
        """设置窗口图标"""
        try:
            from .icon import set_toplevel_icon
            set_toplevel_icon(self.dialog)
        except Exception:
            pass
    
    def _create_widgets(self) -> None:
        theme = self.theme
        
        # 底部按钮 - 先pack以确保显示
        btn_frame = tk.Frame(self.dialog, bg=theme.BG_LIGHT)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        
        # 预览全部按钮
        preview_btn = tk.Label(
            btn_frame, text="预览全部", 
            bg=theme.BG_HOVER, fg=theme.TEXT,
            font=theme.get_font(), cursor='hand2', padx=15, pady=8
        )
        preview_btn.pack(side=tk.LEFT)
        preview_btn.bind('<Button-1>', lambda e: self._preview_all())
        preview_btn.bind('<Enter>', lambda e: preview_btn.config(bg=theme.darken(theme.BG_HOVER, 10)))
        preview_btn.bind('<Leave>', lambda e: preview_btn.config(bg=theme.BG_HOVER))
        
        # 完成按钮
        ok_btn = tk.Label(
            btn_frame, text="完成", 
            bg=theme.RUBRIC_RED, fg="#FFFFFF",
            font=theme.get_font(), cursor='hand2', padx=20, pady=8
        )
        ok_btn.pack(side=tk.RIGHT, padx=(5, 0))
        ok_btn.bind('<Button-1>', lambda e: self._on_ok())
        ok_btn.bind('<Enter>', lambda e: ok_btn.config(bg=theme.RUBRIC_DARK))
        ok_btn.bind('<Leave>', lambda e: ok_btn.config(bg=theme.RUBRIC_RED))
        
        # 取消按钮
        cancel_btn = tk.Label(
            btn_frame, text="取消", 
            bg=theme.BG_HOVER, fg=theme.TEXT,
            font=theme.get_font(), cursor='hand2', padx=20, pady=8
        )
        cancel_btn.pack(side=tk.RIGHT)
        cancel_btn.bind('<Button-1>', lambda e: self._on_cancel())
        cancel_btn.bind('<Enter>', lambda e: cancel_btn.config(bg=theme.darken(theme.BG_HOVER, 10)))
        cancel_btn.bind('<Leave>', lambda e: cancel_btn.config(bg=theme.BG_HOVER))
        
        # 主框架 - 使用tk.PanedWindow以支持主题颜色
        paned = tk.PanedWindow(
            self.dialog, 
            orient=tk.HORIZONTAL, 
            bg=theme.BG_LIGHT,
            sashwidth=6,
            sashrelief=tk.FLAT,
            bd=0
        )
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))
        
        # 左侧 - 组件树
        left_frame = tk.Frame(paned, bg=theme.BG_LIGHT)
        paned.add(left_frame, minsize=200, width=280)  # 增加左侧默认宽度
        
        # 左侧标题
        tk.Label(
            left_frame, text="模块结构", 
            bg=theme.BG_LIGHT, fg=theme.RUBRIC_RED,
            font=theme.get_font("title", bold=True)
        ).pack(anchor=tk.W, pady=(0, 8))
        
        # 树框架
        tree_container = tk.Frame(left_frame, bg=theme.BORDER)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        tree_inner = tk.Frame(tree_container, bg=theme.BG_LIGHT)
        tree_inner.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        self.tree = ttk.Treeview(tree_inner, selectmode="browse")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(tree_inner, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.bind('<Double-1>', self._on_tree_double_click)
        self.tree.bind('<<TreeviewSelect>>', self._on_tree_select)
        
        # 左侧按钮 - 第一行
        left_btn_frame = tk.Frame(left_frame, bg=theme.BG_LIGHT)
        left_btn_frame.pack(fill=tk.X, pady=(8, 0))
        
        edit_btn = tk.Label(
            left_btn_frame, text="编辑", 
            bg=theme.RUBRIC_RED, fg="#FFFFFF",
            font=theme.get_font(), cursor='hand2', padx=12, pady=5
        )
        edit_btn.pack(side=tk.LEFT, padx=(0, 3))
        edit_btn.bind('<Button-1>', lambda e: self._edit_selected())
        edit_btn.bind('<Enter>', lambda e: edit_btn.config(bg=theme.RUBRIC_DARK))
        edit_btn.bind('<Leave>', lambda e: edit_btn.config(bg=theme.RUBRIC_RED))
        
        add_btn = tk.Label(
            left_btn_frame, text="+同类", 
            bg=theme.BG_HOVER, fg=theme.TEXT,
            font=theme.get_font(), cursor='hand2', padx=12, pady=5
        )
        add_btn.pack(side=tk.LEFT, padx=(0, 3))
        add_btn.bind('<Button-1>', lambda e: self._add_similar())
        add_btn.bind('<Enter>', lambda e: add_btn.config(bg=theme.darken(theme.BG_HOVER, 10)))
        add_btn.bind('<Leave>', lambda e: add_btn.config(bg=theme.BG_HOVER))
        
        custom_btn = tk.Label(
            left_btn_frame, text="+自定义", 
            bg=theme.BG_HOVER, fg=theme.TEXT,
            font=theme.get_font(), cursor='hand2', padx=12, pady=5
        )
        custom_btn.pack(side=tk.LEFT, padx=(0, 3))
        custom_btn.bind('<Button-1>', lambda e: self._add_custom())
        custom_btn.bind('<Enter>', lambda e: custom_btn.config(bg=theme.darken(theme.BG_HOVER, 10)))
        custom_btn.bind('<Leave>', lambda e: custom_btn.config(bg=theme.BG_HOVER))
        
        del_btn = tk.Label(
            left_btn_frame, text="删除", 
            bg="#D32F2F", fg="#FFFFFF",
            font=theme.get_font(), cursor='hand2', padx=12, pady=5
        )
        del_btn.pack(side=tk.LEFT)
        del_btn.bind('<Button-1>', lambda e: self._delete_selected())
        del_btn.bind('<Enter>', lambda e: del_btn.config(bg="#B71C1C"))
        del_btn.bind('<Leave>', lambda e: del_btn.config(bg="#D32F2F"))
        
        # 左侧按钮 - 第二行（上移/下移）
        left_btn_frame2 = tk.Frame(left_frame, bg=theme.BG_LIGHT)
        left_btn_frame2.pack(fill=tk.X, pady=(5, 0))
        
        up_btn = tk.Label(
            left_btn_frame2, text="↑ 上移", 
            bg=theme.BG_HOVER, fg=theme.TEXT,
            font=theme.get_font(), cursor='hand2', padx=12, pady=5
        )
        up_btn.pack(side=tk.LEFT, padx=(0, 3))
        up_btn.bind('<Button-1>', lambda e: self._move_up())
        up_btn.bind('<Enter>', lambda e: up_btn.config(bg=theme.darken(theme.BG_HOVER, 10)))
        up_btn.bind('<Leave>', lambda e: up_btn.config(bg=theme.BG_HOVER))
        
        down_btn = tk.Label(
            left_btn_frame2, text="↓ 下移", 
            bg=theme.BG_HOVER, fg=theme.TEXT,
            font=theme.get_font(), cursor='hand2', padx=12, pady=5
        )
        down_btn.pack(side=tk.LEFT)
        down_btn.bind('<Button-1>', lambda e: self._move_down())
        down_btn.bind('<Enter>', lambda e: down_btn.config(bg=theme.darken(theme.BG_HOVER, 10)))
        down_btn.bind('<Leave>', lambda e: down_btn.config(bg=theme.BG_HOVER))
        
        # 右侧 - 预览
        right_frame = tk.Frame(paned, bg=theme.BG_LIGHT)
        paned.add(right_frame, minsize=200)
        
        # 右侧标题
        tk.Label(
            right_frame, text="内容预览", 
            bg=theme.BG_LIGHT, fg=theme.RUBRIC_RED,
            font=theme.get_font("title", bold=True)
        ).pack(anchor=tk.W, pady=(0, 8))
        
        # 预览区域
        preview_container = tk.Frame(right_frame, bg=theme.BORDER)
        preview_container.pack(fill=tk.BOTH, expand=True)
        
        preview_inner = tk.Frame(preview_container, bg=theme.BG_PANEL)
        preview_inner.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        self.preview_text = scrolledtext.ScrolledText(
            preview_inner, 
            wrap=tk.WORD, 
            state=tk.DISABLED,
            bg=theme.BG_PANEL,
            fg=theme.TEXT,
            font=theme.get_font(),
            relief='flat',
            padx=10,
            pady=10
        )
        self.preview_text.pack(fill=tk.BOTH, expand=True)
        
        # 绑定快捷键
        self._bind_shortcuts()
    
    def _bind_shortcuts(self) -> None:
        """绑定快捷键到对话框和树"""
        # 绑定到对话框
        for key in ['z', 'Z']:
            self.dialog.bind(f'<Control-{key}>', lambda e: None)  # 模块编辑暂不支持撤销
        for key in ['c', 'C']:
            self.dialog.bind(f'<Control-{key}>', lambda e: self._copy_slot())
        for key in ['v', 'V']:
            self.dialog.bind(f'<Control-{key}>', lambda e: self._paste_slot())
        self.dialog.bind('<Delete>', lambda e: self._delete_selected())
        
        # 绑定到树
        for key in ['c', 'C']:
            self.tree.bind(f'<Control-{key}>', lambda e: self._copy_slot())
        for key in ['v', 'V']:
            self.tree.bind(f'<Control-{key}>', lambda e: self._paste_slot())
        self.tree.bind('<Delete>', lambda e: self._delete_selected())
    
    def _copy_slot(self) -> None:
        """复制选中的插槽内容"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        tags = self.tree.item(item, "tags")
        
        if len(tags) >= 2:
            comp_id, slot_id = tags[0], tags[1]
            comp = self.module.get_component(comp_id)
            if comp:
                slot = comp.get_slot(slot_id)
                if slot and slot.is_filled():
                    import copy
                    self._clipboard_slot = copy.deepcopy(slot)
    
    def _paste_slot(self) -> None:
        """粘贴插槽内容到选中位置"""
        if not hasattr(self, '_clipboard_slot') or not self._clipboard_slot:
            return
        
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        tags = self.tree.item(item, "tags")
        
        if len(tags) >= 2:
            comp_id, slot_id = tags[0], tags[1]
            comp = self.module.get_component(comp_id)
            if comp:
                slot = comp.get_slot(slot_id)
                if slot:
                    slot.content = self._clipboard_slot.get_content()
                    self._populate_tree()
                    self._select_slot_in_tree(comp_id, slot_id)
    
    def _populate_tree(self) -> None:
        """填充组件树"""
        self.tree.delete(*self.tree.get_children())
        
        # 根节点
        root_id = self.tree.insert("", tk.END, text=f"📖 {self.module.name_zh}", open=True)
        
        # 组件
        for comp in self.module.components:
            comp_id = self.tree.insert(root_id, tk.END, text=f"📑 {comp.name_zh}", open=True)
            
            # 插槽
            for slot in comp.slots:
                status = "✓" if slot.is_filled() else "○"
                required = "*" if slot.required else ""
                slot_text = f"{status} {slot.label_zh}{required}"
                self.tree.insert(comp_id, tk.END, text=slot_text, tags=(comp.component_id, slot.slot_id))
    
    def _on_tree_select(self, event) -> None:
        """树选择变化时更新预览"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        tags = self.tree.item(item, "tags")
        
        if len(tags) >= 2:
            comp_id, slot_id = tags[0], tags[1]
            comp = self.module.get_component(comp_id)
            if comp:
                slot = comp.get_slot(slot_id)
                if slot:
                    self._show_slot_preview(slot)
    
    def _on_tree_double_click(self, event) -> None:
        """双击编辑"""
        self._edit_selected()
    
    def _edit_selected(self) -> None:
        """编辑选中的插槽"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("提示", "请先选择要编辑的项目")
            return
        
        item = selection[0]
        tags = self.tree.item(item, "tags")
        
        if len(tags) >= 2:
            comp_id, slot_id = tags[0], tags[1]
            comp = self.module.get_component(comp_id)
            if comp:
                slot = comp.get_slot(slot_id)
                if slot:
                    dialog = SlotEditDialog(self.dialog, slot, self.theme)
                    result = dialog.show()
                    if result is not None:
                        slot.content = result
                        self._populate_tree()
                        self._show_slot_preview(slot)
    
    def _add_custom(self) -> None:
        """添加自定义内容到选中的组件"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("提示", "请先选择一个组件")
            return
        
        item = selection[0]
        tags = self.tree.item(item, "tags")
        
        if len(tags) >= 1:
            comp_id = tags[0]
            comp = self.module.get_component(comp_id)
            if comp:
                # 创建简单的自定义内容对话框
                temp_slot = ModuleSlot(
                    slot_id="custom",
                    slot_type="text",
                    label_lat="Custom",
                    label_zh="自定义内容",
                    required=False
                )
                dialog = SlotEditDialog(self.dialog, temp_slot, self.theme)
                result = dialog.show()
                if result:
                    comp.add_custom_item(result)
                    messagebox.showinfo("成功", "自定义内容已添加")
    
    def _add_similar(self) -> None:
        """添加与选中项相同类型的新项（插入到当前项之后）"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("提示", "请先选择一个项目")
            return
        
        item = selection[0]
        tags = self.tree.item(item, "tags")
        
        if len(tags) >= 2:
            comp_id, slot_id = tags[0], tags[1]
            comp = self.module.get_component(comp_id)
            if comp:
                slot = comp.get_slot(slot_id)
                if slot:
                    # 找到当前插槽的位置
                    current_idx = -1
                    for i, s in enumerate(comp.slots):
                        if s.slot_id == slot_id:
                            current_idx = i
                            break
                    
                    # 创建相同类型的新插槽
                    new_slot = ModuleSlot(
                        slot_id=f"{slot.slot_type}_custom_{len(comp.slots)}",
                        slot_type=slot.slot_type,
                        label_lat=f"{slot.label_lat} (新)",
                        label_zh=f"{slot.label_zh} (新)",
                        required=False
                    )
                    
                    # 插入到当前项之后
                    if current_idx >= 0:
                        comp.slots.insert(current_idx + 1, new_slot)
                    else:
                        comp.slots.append(new_slot)
                    
                    self._populate_tree()
                    messagebox.showinfo("成功", f"已添加新的 {slot.label_zh}")
    
    def _move_up(self) -> None:
        """上移选中的插槽"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        tags = self.tree.item(item, "tags")
        
        if len(tags) >= 2:
            comp_id, slot_id = tags[0], tags[1]
            comp = self.module.get_component(comp_id)
            if comp:
                for i, slot in enumerate(comp.slots):
                    if slot.slot_id == slot_id and i > 0:
                        comp.slots[i], comp.slots[i-1] = comp.slots[i-1], comp.slots[i]
                        self._populate_tree()
                        # 重新选中移动后的项
                        self._select_slot_in_tree(comp_id, slot_id)
                        return
    
    def _move_down(self) -> None:
        """下移选中的插槽"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        tags = self.tree.item(item, "tags")
        
        if len(tags) >= 2:
            comp_id, slot_id = tags[0], tags[1]
            comp = self.module.get_component(comp_id)
            if comp:
                for i, slot in enumerate(comp.slots):
                    if slot.slot_id == slot_id and i < len(comp.slots) - 1:
                        comp.slots[i], comp.slots[i+1] = comp.slots[i+1], comp.slots[i]
                        self._populate_tree()
                        # 重新选中移动后的项
                        self._select_slot_in_tree(comp_id, slot_id)
                        return
    
    def _select_slot_in_tree(self, comp_id: str, slot_id: str) -> None:
        """在树中选中指定的插槽"""
        def find_and_select(parent=""):
            for item in self.tree.get_children(parent):
                tags = self.tree.item(item, "tags")
                if len(tags) >= 2 and tags[0] == comp_id and tags[1] == slot_id:
                    self.tree.selection_set(item)
                    self.tree.see(item)
                    return True
                if find_and_select(item):
                    return True
            return False
        find_and_select()
    
    def _delete_selected(self) -> None:
        """删除选中的项目"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("提示", "请先选择要删除的项目")
            return
        
        item = selection[0]
        tags = self.tree.item(item, "tags")
        
        if len(tags) >= 2:
            comp_id, slot_id = tags[0], tags[1]
            comp = self.module.get_component(comp_id)
            if comp:
                # 找到并删除插槽
                for i, slot in enumerate(comp.slots):
                    if slot.slot_id == slot_id:
                        if slot.required:
                            if not messagebox.askyesno("确认", f"'{slot.label_zh}' 是必需项，确定要删除吗？"):
                                return
                        del comp.slots[i]
                        self._populate_tree()
                        messagebox.showinfo("成功", "已删除")
                        return
    
    def _show_slot_preview(self, slot: ModuleSlot) -> None:
        """显示插槽预览"""
        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete("1.0", tk.END)
        
        content = slot.get_content()
        if content:
            self.preview_text.insert(tk.END, f"类型: {slot.slot_type}\n")
            self.preview_text.insert(tk.END, f"标签: {slot.label_lat} / {slot.label_zh}\n")
            self.preview_text.insert(tk.END, "-" * 40 + "\n")
            self.preview_text.insert(tk.END, f"拉丁文:\n{content.latin}\n\n")
            self.preview_text.insert(tk.END, f"中文:\n{content.chinese}\n")
        else:
            self.preview_text.insert(tk.END, "（未填充）")
        
        self.preview_text.config(state=tk.DISABLED)
    
    def _preview_all(self) -> None:
        """预览展开后的全部内容"""
        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete("1.0", tk.END)
        
        items = self.module.expand()
        self.preview_text.insert(tk.END, f"=== {self.module.name_zh} ===\n")
        self.preview_text.insert(tk.END, f"共 {len(items)} 个内容项\n")
        self.preview_text.insert(tk.END, "=" * 40 + "\n\n")
        
        for i, item in enumerate(items, 1):
            self.preview_text.insert(tk.END, f"[{i}] {item.item_type}\n")
            if item.latin:
                preview = item.latin[:50] + "..." if len(item.latin) > 50 else item.latin
                self.preview_text.insert(tk.END, f"  L: {preview}\n")
            if item.chinese:
                preview = item.chinese[:30] + "..." if len(item.chinese) > 30 else item.chinese
                self.preview_text.insert(tk.END, f"  C: {preview}\n")
            self.preview_text.insert(tk.END, "\n")
        
        self.preview_text.config(state=tk.DISABLED)
    
    def _on_ok(self) -> None:
        self.result = self.module
        self.dialog.destroy()
    
    def _on_cancel(self) -> None:
        self.result = None
        self.dialog.destroy()
    
    def _center_dialog(self) -> None:
        self.dialog.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() - self.dialog.winfo_width()) // 2
        y = self.parent.winfo_y() + (self.parent.winfo_height() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")
    
    def show(self) -> Optional[HourModule]:
        self.dialog.wait_window()
        return self.result


class ImprimaturDialog:
    """Imprimatur/Nihil Obstat 签名块对话框（统一风格）"""
    
    def __init__(self, parent: tk.Tk, theme: Theme):
        self.parent = parent
        self.theme = theme
        self.result: Optional[List[ContentItem]] = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("添加签名块")
        self.dialog.geometry("480x380")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.configure(bg=theme.BG_LIGHT)
        
        # 设置窗口图标
        self._set_window_icon()
        
        self._create_widgets()
        self._center_dialog()
    
    def _set_window_icon(self) -> None:
        """设置窗口图标"""
        try:
            from .icon import set_toplevel_icon
            set_toplevel_icon(self.dialog)
        except Exception:
            pass
    
    def _create_widgets(self) -> None:
        theme = self.theme
        
        main_frame = tk.Frame(self.dialog, bg=theme.BG_LIGHT, padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Nihil Obstat 区域
        no_frame = tk.Frame(main_frame, bg=theme.BG_DARK, padx=10, pady=10)
        no_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            no_frame, text="Nihil Obstat（可选）", bg=theme.BG_DARK, fg=theme.RUBRIC_RED,
            font=theme.get_font("title", bold=True)
        ).pack(anchor='w', pady=(0, 8))
        
        row1 = tk.Frame(no_frame, bg=theme.BG_DARK)
        row1.pack(fill=tk.X, pady=2)
        tk.Label(row1, text="审查员:", bg=theme.BG_DARK, fg=theme.TEXT, 
                font=theme.get_font(), width=8, anchor='w').pack(side=tk.LEFT)
        self.censor_var = tk.StringVar()
        tk.Entry(row1, textvariable=self.censor_var, width=35, bg=theme.BG_PANEL,
                fg=theme.TEXT, insertbackground=theme.TEXT, font=theme.get_font(),
                relief='flat').pack(side=tk.LEFT, padx=5)
        
        row2 = tk.Frame(no_frame, bg=theme.BG_DARK)
        row2.pack(fill=tk.X, pady=2)
        tk.Label(row2, text="日期:", bg=theme.BG_DARK, fg=theme.TEXT,
                font=theme.get_font(), width=8, anchor='w').pack(side=tk.LEFT)
        self.no_date_var = tk.StringVar()
        tk.Entry(row2, textvariable=self.no_date_var, width=35, bg=theme.BG_PANEL,
                fg=theme.TEXT, insertbackground=theme.TEXT, font=theme.get_font(),
                relief='flat').pack(side=tk.LEFT, padx=5)
        
        # Imprimatur 区域
        imp_frame = tk.Frame(main_frame, bg=theme.BG_DARK, padx=10, pady=10)
        imp_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            imp_frame, text="Imprimatur", bg=theme.BG_DARK, fg=theme.RUBRIC_RED,
            font=theme.get_font("title", bold=True)
        ).pack(anchor='w', pady=(0, 8))
        
        row3 = tk.Frame(imp_frame, bg=theme.BG_DARK)
        row3.pack(fill=tk.X, pady=2)
        tk.Label(row3, text="主教:", bg=theme.BG_DARK, fg=theme.TEXT,
                font=theme.get_font(), width=8, anchor='w').pack(side=tk.LEFT)
        self.bishop_var = tk.StringVar()
        tk.Entry(row3, textvariable=self.bishop_var, width=35, bg=theme.BG_PANEL,
                fg=theme.TEXT, insertbackground=theme.TEXT, font=theme.get_font(),
                relief='flat').pack(side=tk.LEFT, padx=5)
        
        row4 = tk.Frame(imp_frame, bg=theme.BG_DARK)
        row4.pack(fill=tk.X, pady=2)
        tk.Label(row4, text="教区:", bg=theme.BG_DARK, fg=theme.TEXT,
                font=theme.get_font(), width=8, anchor='w').pack(side=tk.LEFT)
        self.diocese_var = tk.StringVar()
        tk.Entry(row4, textvariable=self.diocese_var, width=35, bg=theme.BG_PANEL,
                fg=theme.TEXT, insertbackground=theme.TEXT, font=theme.get_font(),
                relief='flat').pack(side=tk.LEFT, padx=5)
        
        row5 = tk.Frame(imp_frame, bg=theme.BG_DARK)
        row5.pack(fill=tk.X, pady=2)
        tk.Label(row5, text="日期:", bg=theme.BG_DARK, fg=theme.TEXT,
                font=theme.get_font(), width=8, anchor='w').pack(side=tk.LEFT)
        self.imp_date_var = tk.StringVar()
        tk.Entry(row5, textvariable=self.imp_date_var, width=35, bg=theme.BG_PANEL,
                fg=theme.TEXT, insertbackground=theme.TEXT, font=theme.get_font(),
                relief='flat').pack(side=tk.LEFT, padx=5)
        
        # 按钮区域
        btn_frame = tk.Frame(self.dialog, bg=theme.BG_LIGHT, pady=10)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        cancel_btn = tk.Label(
            btn_frame, text="取消", bg=theme.BG_HOVER, fg=theme.TEXT,
            font=theme.get_font(), cursor='hand2', padx=20, pady=8
        )
        cancel_btn.pack(side=tk.RIGHT, padx=10)
        cancel_btn.bind('<Button-1>', lambda e: self._on_cancel())
        cancel_btn.bind('<Enter>', lambda e: cancel_btn.config(bg=theme.darken(theme.BG_HOVER, 15)))
        cancel_btn.bind('<Leave>', lambda e: cancel_btn.config(bg=theme.BG_HOVER))
        
        ok_btn = tk.Label(
            btn_frame, text="添加", bg=theme.RUBRIC_RED, fg="#FFFFFF",
            font=theme.get_font(), cursor='hand2', padx=20, pady=8
        )
        ok_btn.pack(side=tk.RIGHT, padx=5)
        ok_btn.bind('<Button-1>', lambda e: self._on_ok())
        ok_btn.bind('<Enter>', lambda e: ok_btn.config(bg=theme.RUBRIC_DARK))
        ok_btn.bind('<Leave>', lambda e: ok_btn.config(bg=theme.RUBRIC_RED))
    
    def _on_ok(self) -> None:
        items = []
        
        # Nihil Obstat
        if self.censor_var.get().strip():
            items.append(ContentItem(
                item_type="nihilobstat",
                latin=self.censor_var.get().strip(),
                chinese="",
                arg=self.no_date_var.get().strip()
            ))
        
        # Imprimatur
        if self.bishop_var.get().strip():
            items.append(ContentItem(
                item_type="imprimatur",
                latin=self.bishop_var.get().strip(),
                chinese=self.diocese_var.get().strip(),
                arg=self.imp_date_var.get().strip()
            ))
        
        if items:
            self.result = items
        
        self.dialog.destroy()
    
    def _on_cancel(self) -> None:
        self.result = None
        self.dialog.destroy()
    
    def _center_dialog(self) -> None:
        self.dialog.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() - self.dialog.winfo_width()) // 2
        y = self.parent.winfo_y() + (self.parent.winfo_height() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")
    
    def show(self) -> Optional[List[ContentItem]]:
        self.dialog.wait_window()
        return self.result
