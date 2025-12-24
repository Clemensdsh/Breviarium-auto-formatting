#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tex_generator.py - Psalter LaTeX 生成器 (完整全功能版 v6)
包含功能：
1. 完整的增删改查 UI。
2. 自动编译预览 (XeLaTeX)。
3. 封面标题自定义 (左下角紫色按钮)。
4. 智能 Paracol 分栏管理。
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, scrolledtext
import os, csv, shutil, re, subprocess, platform
import sys

# ==========================================
# 1. LaTeX 命令映射配置
# ==========================================
TEX_MAPPING = {
    'h1':           (r'\psHeaderOne{{{l}}}{{{c}}}',      r'\psSingleHeaderOne{{{c}}}'),
    'h1cap':        (r'\psHeaderOneCap{{{l}}}{{{c}}}',   r'\psSingleHeaderOneCap{{{l}}}{{{c}}}'),
    'h1lowercase':  (r'\psHeaderOneLowercase{{{l}}}{{{c}}}', r'\psSingleHeaderOneLowercase{{{l}}}{{{c}}}'),
    'h2':           (r'\psHeaderTwo{{{l}}}{{{c}}}',      r'\psSingleHeaderTwo{{{c}}}'),
    'h3':           (r'\psHeaderThree{{{l}}}{{{c}}}',    r'\psSingleHeaderThree{{{c}}}'),
    'psalmtitle':   (r'\psPsalmTitle{{{l}}}{{{c}}}',     r'\psSinglePsalmTitle{{{c}}}'),
    'canticletitle':(r'\psCanticleTitle{{{l}}}{{{c}}}',  r'\psSingleCanticleTitle{{{c}}}'),
    'hymntitle':    (r'\psHymnTitle{{{l}}}{{{c}}}',      r'\psSingleHymnTitle{{{c}}}'),
    'hymnheader':   (r'\psHymnHeader{{{l}}}{{{c}}}',     r'\psSingleHymnHeader{{{c}}}'),
    'antiphon':     (r'\psAntiphonRepeat{{{l}}}{{{c}}}', r'\psSingleAntiphon{{{c}}}'),
    'dropcap':      (r'\psVerseDropcap{{{l}}}{{{c}}}',   r'\psSingleVerseDropcap{{{c}}}'),
    'verse':        (r'\psVerse{{{l}}}{{{c}}}',          r'\psSingleVerse{{{c}}}'),
    'gloria':       (r'\psGloria{{{l}}}{{{c}}}',         r'\psSingleGloria{{{c}}}'),
    'rubric':       (r'\psRubric{{{l}}}{{{c}}}',         r'\psSingleRubric{{{c}}}'),
    'V':            (r'\psVR{{V}}{{{l}}}{{{c}}}',        r'\psSingleVR{{V}}{{{c}}}'),
    'R':            (r'\psVR{{R}}{{{l}}}{{{c}}}',        r'\psSingleVR{{R}}{{{c}}}'),
    'hymn':         (r'\psHymnStanza{{{l}}}{{{c}}}',     r'\psSingleHymnStanza{{{c}}}'),
    'capit':        (r'\psCapit{{{l}}}{{{c}}}',          r'\psSingleCapit{{{c}}}'),
    'capitheader':  (r'\psCapitHeader{{{l}}}{{{c}}}',    r'\psSingleCapitHeader{{{c}}}'),
    'scriptureref': (r'\psScriptureRef{{{l}}}{{{c}}}',   r'\psSingleScriptureRef{{{c}}}'),
    'collect':      (r'\psCollect{{{l}}}{{{c}}}',        r'\psSingleCollect{{{c}}}'),
    'lesson':       (r'\psLesson{{{l}}}{{{c}}}',         r'\psSingleLesson{{{c}}}'),
    'text':         (r'\psText{{{l}}}{{{c}}}',           r'\psSingleText{{{c}}}'),
    'rule':         (r'\psThinRule',                     r'\psSingleThinRule'),
    'thickrule':    (r'\psThickRule',                    r'\psSingleThickRule'),
}

# ==========================================
# 2. 核心样式与类定义
# ==========================================
class S:
    BG_DARK = "#17212b"
    BG_LIGHT = "#242f3d"
    BG_HOVER = "#2b5278"
    ACCENT = "#5288c1"
    ACCENT_LIGHT = "#6ab3f3"
    TEXT = "#f5f5f5"
    TEXT_SEC = "#8b9ba5"
    SUCCESS = "#50a550"
    WARNING = "#d4a535"
    DANGER = "#c45c5c"
    BORDER = "#3d4d5c"
    SCROLL_FG = "#4a5d6e"
    PURPLE = "#8e44ad"  # 封面设置按钮颜色

class ContentItem:
    def __init__(self, t, l="", c="", a="", src="", multi=False, cnt=1):
        self.item_type, self.latin, self.chinese, self.arg = t, l, c, a
        self.source_file, self.is_multiline, self.line_count = src, multi, cnt
    def to_csv_row(self): return [self.item_type, self.latin, self.chinese, self.arg]
    def get_display_text(self):
        t = self.item_type
        if self.is_multiline:
            return f"[{t}] {self.latin[:20]}... | {self.chinese[:10]}... (+{self.line_count-1}行)"
        if t == "image": return f"[图片] {os.path.basename(self.latin)}"
        if t == "rule": return "[分隔线]"
        if t == "thickrule": return "[粗分隔线]"
        if t == "pagebreak": return "[分页]"
        if t == "tocstart": return "[目录起始]"
        if t == "singlecol": return "[单栏/双栏切换]"
        pl = self.latin[:20] + "..." if len(self.latin) > 20 else self.latin
        pc = self.chinese[:10] + "..." if len(self.chinese) > 10 else self.chinese
        return f"[{t}] {pl} | {pc}"

class MultiLineContentItem(ContentItem):
    def __init__(self, src, items):
        self.items = items
        if items:
            f = items[0]
            super().__init__(f.item_type, f.latin, f.chinese, f.arg, src, True, len(items))
        else:
            super().__init__("", "", "", "", src, True, 0)
    def to_csv_rows(self): 
        return [i.to_csv_row() for i in self.items]
    def get_flat_items(self): 
        return self.items

class FileContentLoader:
    CATEGORIES = {
        "psalms": "圣咏 (Psalms)", "canticles": "圣歌 (Canticles)",
        "hymns": "赞美诗 (Hymns)", "antiphons": "对经 (Antiphons)",
        "lessons": "读经 (Lessons)", "responsories": "答唱咏 (Responsories)",
        "collects": "集祷经 (Collects)", "common": "通用文本 (Common)"
    }
    def __init__(self, d):
        self.content_dir = d
        for c in self.CATEGORIES: os.makedirs(os.path.join(d, c), exist_ok=True)
    def get_available_files(self):
        files = {c: [] for c in self.CATEGORIES}
        for cat in files:
            p = os.path.join(self.content_dir, cat)
            if os.path.exists(p):
                fl = []
                for f in os.listdir(p):
                    if f.endswith('.txt'):
                        nums = re.findall(r'\d+', f)
                        k = int(nums[0]) if nums else 999
                        if len(nums) > 1: k = k * 100 + int(nums[1])
                        fl.append((k, f))
                fl.sort(key=lambda x: x[0])
                files[cat] = fl
        return files
    def load_file_content(self, cat, fn):
        fp = os.path.join(self.content_dir, cat, fn)
        items = []
        if not os.path.exists(fp): return items
        with open(fp, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#'): continue
                parts = line.split('|')
                if len(parts) >= 3:
                    items.append(ContentItem(parts[0], parts[1], parts[2], parts[3] if len(parts) > 3 else ""))
        return items
    def load_file_as_multiline(self, cat, fn):
        items = self.load_file_content(cat, fn)
        return MultiLineContentItem(fn, items) if items else None

FORMAT_TYPES = [
    ("h1", "大标题"), ("h1cap", "目录大标题"), ("h1lowercase", "目录小标题"),
    ("h2", "副标题"), ("h3", "节次标题"), ("psalmtitle", "圣咏标题"),
    ("canticletitle", "圣歌标题"), ("hymntitle", "赞美诗标题"), ("hymnheader", "赞美诗加粗标题"),
    ("antiphon", "对经"), ("antiphonnum", "对经(带编号)"), ("dropcap", "首字下沉文本"),
    ("verse", "诗节"), ("gloria", "圣三光荣颂"), ("rubric", "礼仪指示"),
    ("V", "启(V)"), ("R", "应(R)"), ("hymn", "赞美诗节"),
    ("capit", "短读经"), ("capitheader", "短读经标题"), ("scriptureref", "圣经引用"),
    ("collect", "集祷经"), ("lesson", "读经标题"), ("text", "普通文本"),
    ("rule", "分隔线"), ("thickrule", "粗分隔线"), ("pagebreak", "分页"),
    ("tocstart", "目录起始"), ("singlecol", "单栏/双栏切换"), ("image", "图片"),
]

# ==========================================
# 3. 自定义控件类
# ==========================================
class TelegramScrollbar(tk.Canvas):
    def __init__(self, parent, command=None, **kw):
        super().__init__(parent, width=8, highlightthickness=0, bg=S.BG_LIGHT, **kw)
        self.command = command
        self.thumb_pos = 0; self.thumb_size = 0.3; self.dragging = False; self.drag_start = 0; self.hover = False
        self.bind('<Button-1>', self.on_click); self.bind('<B1-Motion>', self.on_drag)
        self.bind('<ButtonRelease-1>', lambda e: setattr(self, 'dragging', False))
        self.bind('<Configure>', self.draw)
        self.bind('<Enter>', lambda e: (setattr(self, 'hover', True), self.draw()))
        self.bind('<Leave>', lambda e: (setattr(self, 'hover', False), self.draw()))
    def set(self, first, last): self.thumb_pos = float(first); self.thumb_size = float(last) - float(first); self.draw()
    def draw(self, e=None):
        self.delete('all'); h, w = self.winfo_height(), self.winfo_width()
        if h <= 1: return
        self.create_rectangle(0, 0, w, h, fill=S.BG_LIGHT, outline='')
        th = max(30, h * self.thumb_size)
        tt = self.thumb_pos * (h - th) / (1 - self.thumb_size) if self.thumb_size < 1 else 0
        c = '#5a6d7e' if self.hover else S.SCROLL_FG
        self.create_rectangle(1, tt+2, w-1, tt+th-2, fill=c, outline='')
    def on_click(self, e):
        h = self.winfo_height(); th = max(30, h * self.thumb_size)
        tt = self.thumb_pos * (h - th) / (1 - self.thumb_size) if self.thumb_size < 1 else 0
        if tt <= e.y <= tt + th: self.dragging = True; self.drag_start = e.y - tt
        elif self.command: self.command('moveto', str(e.y / h))
    def on_drag(self, e):
        if self.dragging and self.command:
            h = self.winfo_height(); th = max(30, h * self.thumb_size)
            nt = e.y - self.drag_start; np = nt / (h - th) * (1 - self.thumb_size)
            np = max(0, min(1 - self.thumb_size, np)); self.command('moveto', str(np))

class PanedWindow(tk.Frame):
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=S.BG_DARK, **kw)
        self.left_frame = tk.Frame(self, bg=S.BG_DARK)
        self.sash = tk.Frame(self, bg=S.BORDER, width=5, cursor='sb_h_double_arrow')
        self.right_frame = tk.Frame(self, bg=S.BG_DARK)
        self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False)
        self.sash.pack(side=tk.LEFT, fill=tk.Y)
        self.right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.sash.bind('<Button-1>', self.start_drag); self.sash.bind('<B1-Motion>', self.do_drag)
        self.sash.bind('<Enter>', lambda e: self.sash.config(bg=S.ACCENT))
        self.sash.bind('<Leave>', lambda e: self.sash.config(bg=S.BORDER))
        self.drag_start = 0; self.left_width = 280
    def start_drag(self, e): self.drag_start = e.x_root; self.left_width = self.left_frame.winfo_width()
    def do_drag(self, e):
        delta = e.x_root - self.drag_start; nw = max(150, min(600, self.left_width + delta))
        self.left_frame.config(width=nw)

# ==========================================
# 4. 标题页编辑对话框
# ==========================================
class TitlePageDialog:
    def __init__(self, parent, initial_data):
        self.result = None
        self.top = tk.Toplevel(parent)
        self.top.title("设置封面标题")
        self.top.geometry("600x480")
        self.top.configure(bg=S.BG_DARK)
        self.top.transient(parent)
        self.top.grab_set()

        # 标题区域
        tk.Label(self.top, text="设置 PDF 封面文本", bg=S.BG_DARK, fg=S.ACCENT_LIGHT, 
                 font=('Segoe UI', 12, 'bold')).pack(pady=15)
        
        form_frame = tk.Frame(self.top, bg=S.BG_DARK)
        form_frame.pack(fill=tk.BOTH, expand=True, padx=20)

        # 辅助函数：创建输入行
        self.entries = {}
        def add_field(key, label_text, default_val, height=2):
            f = tk.Frame(form_frame, bg=S.BG_DARK)
            f.pack(fill=tk.X, pady=5)
            tk.Label(f, text=label_text, bg=S.BG_DARK, fg=S.TEXT, width=15, anchor='e').pack(side=tk.LEFT, padx=5)
            txt = tk.Text(f, height=height, bg=S.BG_LIGHT, fg=S.TEXT, insertbackground=S.TEXT, 
                          font=('Segoe UI', 10), borderwidth=0, padx=5, pady=5)
            txt.pack(side=tk.LEFT, fill=tk.X, expand=True)
            txt.insert(1.0, initial_data.get(key, default_val))
            self.entries[key] = txt

        add_field("title_zh", "中文主标题:", "羅馬大日課\\\\[0.5em]耶穌聖誕瞻禮")
        add_field("title_lat", "拉丁文标题:", "Breviárium Románum\\\\[0.5em]In Nativitáte Dómini")
        add_field("edition", "版本/编者:", "中拉對照\\\\[0.5em]Edítio Sínico-Latína")
        add_field("footer", "底部文字:", "Pro Manuscripto")

        tk.Label(form_frame, text="提示：使用 \\\\ 表示换行，\\\\[0.5em] 表示带间距换行", 
                 bg=S.BG_DARK, fg=S.TEXT_SEC, font=('Segoe UI', 9)).pack(pady=10)

        # 按钮区
        btn_f = tk.Frame(self.top, bg=S.BG_DARK)
        btn_f.pack(fill=tk.X, pady=15, padx=20)
        
        cancel_btn = tk.Label(btn_f, text="取消", bg=S.BG_HOVER, fg="white", padx=15, pady=6, cursor='hand2')
        cancel_btn.pack(side=tk.RIGHT, padx=5)
        cancel_btn.bind('<Button-1>', lambda e: self.top.destroy())

        save_btn = tk.Label(btn_f, text="保存设置", bg=S.SUCCESS, fg="white", padx=15, pady=6, cursor='hand2')
        save_btn.pack(side=tk.RIGHT, padx=5)
        save_btn.bind('<Button-1>', self.save)

    def save(self, e=None):
        data = {}
        for k, v in self.entries.items():
            data[k] = v.get(1.0, tk.END).strip()
        self.result = data
        self.top.destroy()

# ==========================================
# 5. 主程序类 CSVEditorApp
# ==========================================

def get_application_path():
    """获取应用程序运行目录（兼容打包后的 EXE 和源码运行）"""
    if getattr(sys, 'frozen', False):
        # 如果是打包后的 EXE，使用 EXE 所在的目录
        return os.path.dirname(sys.executable)
    else:
        # 如果是源码运行，使用脚本所在的目录
        return os.path.dirname(os.path.abspath(__file__))

class CSVEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Psalter LaTeX Generator")
        self.root.geometry("1200x900")
        self.root.minsize(900, 600)
        self.root.configure(bg=S.BG_DARK)
        self.content_items = []
        
        # 默认封面标题数据
        self.title_data = {
            "title_zh": "羅馬大日課\\\\[0.5em]耶穌聖誕瞻禮",
            "title_lat": "Breviárium Románum\\\\[0.5em]In Nativitáte Dómini",
            "edition": "中拉對照\\\\[0.5em]Edítio Sínico-Latína",
            "footer": "Pro Manuscripto"
        }

        self.base_dir = get_application_path()
        self.content_dir = os.path.join(self.base_dir, "content")
        self.images_dir = os.path.join(self.base_dir, "images")
        self.loader = FileContentLoader(self.content_dir)
        
        os.makedirs(self.images_dir, exist_ok=True)
        self.setup_ui()
    
    def setup_ui(self):
        main = tk.Frame(self.root, bg=S.BG_DARK)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        main.columnconfigure(0, weight=0, minsize=220)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)
        
        # === 左侧面板 ===
        left = tk.Frame(main, bg=S.BG_DARK, width=220)
        left.grid(row=0, column=0, sticky='nsew', padx=(0, 8))
        left.grid_propagate(False)
        
        tk.Label(left, text="内容来源", bg=S.BG_DARK, fg=S.ACCENT_LIGHT,
                font=('Segoe UI', 11, 'bold')).pack(anchor='w', pady=(0, 8))
        
        tree_f = tk.Frame(left, bg=S.BG_LIGHT)
        tree_f.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        
        self.file_tree = ttk.Treeview(tree_f, show='tree', selectmode='browse')
        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        style = ttk.Style()
        style.configure('Treeview', background=S.BG_LIGHT, foreground=S.TEXT,
                       fieldbackground=S.BG_LIGHT, font=('Segoe UI', 10), rowheight=26)
        style.map('Treeview', background=[('selected', S.BG_HOVER)], foreground=[('selected', S.TEXT)])
        
        ts = TelegramScrollbar(tree_f, command=self.file_tree.yview)
        ts.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_tree.config(yscrollcommand=ts.set)
        self.load_file_tree()
        
        # === 左下角按钮区 ===
        btn_f = tk.Frame(left, bg=S.BG_DARK)
        # 强制固定在底部
        btn_f.pack(side=tk.BOTTOM, fill=tk.X, pady=5)
        
        # 功能按钮循环
        for txt, cmd, bg in [
            ("添加选中文件", self.add_selected_file, S.ACCENT),
            ("添加自定义内容", self.add_custom_content, S.BG_HOVER),
            ("添加图片", self.add_image, S.BG_HOVER),
            ("添加分隔线", self.add_rule, S.BG_HOVER),
            ("添加分页", self.add_pagebreak, S.BG_HOVER),
            ("添加目录起始", self.add_tocstart, S.BG_HOVER),
            ("切换单栏/双栏", self.add_singlecol, S.SUCCESS),
        ]:
            self.make_btn(btn_f, txt, cmd, bg).pack(fill=tk.X, pady=2)
        
        # 【新增】紫色的封面设置按钮
        tk.Label(btn_f, text="----------------", bg=S.BG_DARK, fg=S.BORDER).pack(fill=tk.X, pady=2)
        self.make_btn(btn_f, "设置封面标题", self.edit_title_page, S.PURPLE).pack(fill=tk.X, pady=2)

        # === 中间面板 ===
        paned = PanedWindow(main)
        paned.grid(row=0, column=1, sticky='nsew')
        
        center = paned.left_frame
        center.config(width=280)
        
        tk.Label(center, text="当前内容", bg=S.BG_DARK, fg=S.ACCENT_LIGHT,
                font=('Segoe UI', 11, 'bold')).pack(anchor='w', pady=(0, 8))
        
        list_f = tk.Frame(center, bg=S.BG_LIGHT)
        list_f.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        
        self.content_listbox = tk.Listbox(list_f, bg=S.BG_LIGHT, fg=S.TEXT,
            selectbackground=S.BG_HOVER, selectforeground=S.TEXT,
            font=('Segoe UI', 10), borderwidth=0, highlightthickness=0, activestyle='none')
        self.content_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        ls = TelegramScrollbar(list_f, command=self.content_listbox.yview)
        ls.pack(side=tk.RIGHT, fill=tk.Y)
        self.content_listbox.config(yscrollcommand=ls.set)
        
        ops = tk.Frame(center, bg=S.BG_DARK)
        ops.pack(fill=tk.X)
        
        r1 = tk.Frame(ops, bg=S.BG_DARK)
        r1.pack(fill=tk.X, pady=2)
        for txt, cmd, bg in [("上移", self.move_up, S.BG_HOVER), ("下移", self.move_down, S.BG_HOVER),
                             ("编辑", self.edit_item, S.ACCENT), ("删除", self.delete_item, S.DANGER)]:
            self.make_btn(r1, txt, cmd, bg, 6).pack(side=tk.LEFT, padx=1)
        
        r2 = tk.Frame(ops, bg=S.BG_DARK)
        r2.pack(fill=tk.X, pady=2)
        self.make_btn(r2, "清空全部", self.clear_all, S.WARNING, 13).pack(side=tk.LEFT, padx=1)
        
        # === 右侧面板 ===
        right = paned.right_frame
        tk.Label(right, text="预览和导出", bg=S.BG_DARK, fg=S.ACCENT_LIGHT,
                font=('Segoe UI', 11, 'bold')).pack(anchor='w', pady=(0, 8), padx=(8, 0))
        
        pf = tk.Frame(right, bg=S.BG_LIGHT)
        pf.pack(fill=tk.BOTH, expand=True, pady=(0, 8), padx=(8, 0))
        
        self.preview_text = tk.Text(pf, bg=S.BG_LIGHT, fg=S.TEXT, insertbackground=S.TEXT,
            font=('Consolas', 10), borderwidth=0, highlightthickness=0, wrap=tk.NONE, padx=8, pady=8)
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        ps = TelegramScrollbar(pf, command=self.preview_text.yview)
        ps.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_text.config(yscrollcommand=ps.set)
        
        exp = tk.Frame(right, bg=S.BG_DARK)
        exp.pack(fill=tk.X, pady=(8, 0), padx=(8, 0))
        
        self.make_btn(exp, "刷新预览", self.refresh_preview, S.BG_HOVER, 8).pack(side=tk.LEFT, padx=2)
        self.make_btn(exp, "保存CSV", self.export_csv, S.ACCENT, 8).pack(side=tk.LEFT, padx=2)
        self.make_btn(exp, "导出 Body.tex", self.export_tex, S.SUCCESS, 12).pack(side=tk.LEFT, padx=2)
        
        # 编译按钮
        compile_btn = tk.Label(exp, text="编译并预览 PDF", bg="#c62828", fg="white", 
                             font=('Segoe UI', 10, 'bold'), cursor='hand2', padx=15, pady=6)
        compile_btn.pack(side=tk.LEFT, padx=5)
        compile_btn.bind('<Enter>', lambda e: compile_btn.config(bg="#d32f2f"))
        compile_btn.bind('<Leave>', lambda e: compile_btn.config(bg="#c62828"))
        compile_btn.bind('<Button-1>', lambda e: self.compile_preview())
        
        self.content_listbox.bind('<Double-1>', lambda e: self.edit_item())
    
    def make_btn(self, parent, text, cmd, bg, w=None):
        btn = tk.Label(parent, text=text, bg=bg, fg=S.TEXT, font=('Segoe UI', 10),
                      cursor='hand2', padx=12, pady=6)
        if w: btn.config(width=w)
        def enter(e): btn.config(bg=self.lighten(bg))
        def leave(e): btn.config(bg=bg)
        btn.bind('<Enter>', enter)
        btn.bind('<Leave>', leave)
        btn.bind('<Button-1>', lambda e: cmd())
        return btn
    
    def lighten(self, c):
        c = c.lstrip('#')
        r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
        return f'#{min(255,r+20):02x}{min(255,g+20):02x}{min(255,b+20):02x}'
    
    def load_file_tree(self):
        for i in self.file_tree.get_children(): self.file_tree.delete(i)
        files = self.loader.get_available_files()
        for k, n in FileContentLoader.CATEGORIES.items():
            cid = self.file_tree.insert("", tk.END, text=n, open=False)
            for _, fn in files.get(k, []):
                self.file_tree.insert(cid, tk.END, text=fn, values=(k, fn))
    
    def add_selected_file(self):
        sel = self.file_tree.selection()
        if not sel: messagebox.showwarning("提示", "请先选择要添加的文件"); return
        for sid in sel:
            item = self.file_tree.item(sid)
            vals = item.get('values', [])
            if len(vals) >= 2:
                mi = self.loader.load_file_as_multiline(vals[0], vals[1])
                if mi: self.content_items.append(mi)
        self.refresh_listbox()
        self.refresh_preview()
    
    def add_custom_content(self):
        d = CustomContentDialog(self.root)
        self.root.wait_window(d.top)
        if d.result: self.content_items.append(d.result); self.refresh_listbox(); self.refresh_preview()
    
    def add_image(self):
        fp = filedialog.askopenfilename(title="选择图片", filetypes=[("图片文件", "*.png *.jpg *.jpeg *.gif *.bmp")])
        if fp:
            fn = os.path.basename(fp)
            dp = os.path.join(self.images_dir, fn)
            if not os.path.exists(dp): shutil.copy2(fp, dp)
            h = simpledialog.askstring("图片高度", "请输入图片高度（留空使用默认值3.2cm）:", initialvalue="")
            self.content_items.append(ContentItem("image", f"images/{fn}", "", h or ""))
            self.refresh_listbox(); self.refresh_preview()
    
    def add_rule(self):
        c = messagebox.askyesnocancel("分隔线类型", "是 = 普通分隔线\n否 = 粗分隔线\n取消 = 不添加")
        if c is True: self.content_items.append(ContentItem("rule", "", "", ""))
        elif c is False: self.content_items.append(ContentItem("thickrule", "", "", ""))
        self.refresh_listbox(); self.refresh_preview()
    
    def add_pagebreak(self):
        self.content_items.append(ContentItem("pagebreak", "", "", ""))
        self.refresh_listbox(); self.refresh_preview()
    
    def add_tocstart(self):
        self.content_items.append(ContentItem("tocstart", "", "", ""))
        self.refresh_listbox(); self.refresh_preview()
    
    def add_singlecol(self):
        self.content_items.append(ContentItem("singlecol", "", "", ""))
        self.refresh_listbox(); self.refresh_preview()
    
    def move_up(self):
        sel = self.content_listbox.curselection()
        if not sel or sel[0] == 0: return
        i = sel[0]
        self.content_items[i], self.content_items[i-1] = self.content_items[i-1], self.content_items[i]
        self.refresh_listbox(); self.content_listbox.selection_set(i-1); self.refresh_preview()
    
    def move_down(self):
        sel = self.content_listbox.curselection()
        if not sel or sel[0] >= len(self.content_items) - 1: return
        i = sel[0]
        self.content_items[i], self.content_items[i+1] = self.content_items[i+1], self.content_items[i]
        self.refresh_listbox(); self.content_listbox.selection_set(i+1); self.refresh_preview()
    
    def edit_item(self):
        sel = self.content_listbox.curselection()
        if not sel: return
        item = self.content_items[sel[0]]
        if isinstance(item, MultiLineContentItem):
            messagebox.showinfo("提示", "多行文件内容无法直接编辑。"); return
        d = CustomContentDialog(self.root, item)
        self.root.wait_window(d.top)
        if d.result: self.content_items[sel[0]] = d.result; self.refresh_listbox(); self.refresh_preview()
    
    def delete_item(self):
        sel = self.content_listbox.curselection()
        if not sel: return
        if messagebox.askyesno("确认", "确定要删除选中的项目吗？"):
            del self.content_items[sel[0]]; self.refresh_listbox(); self.refresh_preview()
    
    def clear_all(self):
        if messagebox.askyesno("确认", "确定要清空所有内容吗？"):
            self.content_items.clear(); self.refresh_listbox(); self.refresh_preview()

    def refresh_listbox(self):
        self.content_listbox.delete(0, tk.END)
        for item in self.content_items: self.content_listbox.insert(tk.END, item.get_display_text())
    
    def refresh_preview(self):
        self.preview_text.delete(1.0, tk.END)
        try:
            lines = []
            for item in self.content_items:
                if isinstance(item, MultiLineContentItem):
                    for r in item.to_csv_rows(): lines.append(",".join(f'"{x}"' for x in r))
                else:
                    lines.append(",".join(f'"{x}"' for x in item.to_csv_row()))
            self.preview_text.insert(tk.END, "\n".join(lines))
        except Exception as e:
            self.preview_text.insert(tk.END, f"预览出错: {str(e)}")
    
    def export_csv(self):
        if not self.content_items: messagebox.showwarning("提示", "没有内容可导出"); return
        fp = filedialog.asksaveasfilename(title="保存CSV工程文件", defaultextension=".csv",
            filetypes=[("CSV文件", "*.csv")], initialfile="psalter_project.csv")
        if fp:
            try:
                with open(fp, 'w', encoding='utf-8', newline='') as f:
                    w = csv.writer(f)
                    for item in self.content_items:
                        if isinstance(item, MultiLineContentItem):
                            for r in item.to_csv_rows(): w.writerow(r)
                        else: w.writerow(item.to_csv_row())
                messagebox.showinfo("成功", f"工程文件已保存到:\n{fp}")
            except Exception as e: messagebox.showerror("错误", f"保存失败: {str(e)}")

    # ==========================================================
    # 打开封面设置对话框
    # ==========================================================
    def edit_title_page(self):
        d = TitlePageDialog(self.root, self.title_data)
        self.root.wait_window(d.top)
        if d.result:
            self.title_data = d.result

    # ==========================================================
    # 生成 LaTeX 内容 (Paracol管理)
    # ==========================================================
    def get_latex_content(self):
        latex_lines = [r"\begin{paracol}{2}"]
        is_single_col = False
        
        flat_items = []
        for item in self.content_items:
            if isinstance(item, MultiLineContentItem):
                flat_items.extend(item.get_flat_items())
            else:
                flat_items.append(item)
        
        for item in flat_items:
            t, l, c, a = item.item_type, item.latin, item.chinese, item.arg
            
            if t == 'tocstart':
                if not is_single_col:
                    latex_lines.append(r"\end{paracol}")
                
                latex_lines.append(r"\psPrintToc")
                latex_lines.append(r"\clearpage")
                latex_lines.append(r"\pagenumbering{arabic}")
                latex_lines.append(r"\pagestyle{fancy}")
                latex_lines.append(r"\begin{paracol}{2}")
                is_single_col = False
                continue
            
            if t == 'singlecol':
                if is_single_col:
                    latex_lines.append(r"\psExitSingleCol")
                    is_single_col = False
                else:
                    latex_lines.append(r"\psEnterSingleCol")
                    is_single_col = True
                continue

            if t == 'pagebreak':
                latex_lines.append(r"\psSinglePageBreak" if is_single_col else r"\psPageBreak")
                continue
            
            if t in TEX_MAPPING:
                double_cmd, single_cmd = TEX_MAPPING[t]
                cmd = single_cmd.format(l=l, c=c, a=a) if is_single_col else double_cmd.format(l=l, c=c, a=a)
                latex_lines.append(cmd)
            
            elif t == 'antiphonnum':
                if is_single_col:
                    latex_lines.append(rf"\psSingleAntiphonNum{{{a}}}{{{c}}}")
                else:
                    latex_lines.append(rf"\psAntiphonNum{{{a}}}{{{l}}}{{{c}}}")
            
            elif t == 'image':
                if is_single_col:
                    latex_lines.append(rf"\psSingleImage{{{l}}}")
                else:
                    latex_lines.append(rf"\psImageFullWidth{{{l}}}")
            
            else:
                latex_lines.append(f"% 未知类型: {t} | {l} | {c}")
        
        if not is_single_col:
            latex_lines.append(r"\end{paracol}")
        return "\n".join(latex_lines)

    def export_tex(self):
        if not self.content_items: messagebox.showwarning("提示", "没有内容可导出"); return
        fp = filedialog.asksaveasfilename(title="生成 TeX 文件", defaultextension=".tex",
            filetypes=[("TeX 文件", "*.tex")], initialfile="body.tex")
        if not fp: return
        try:
            content = self.get_latex_content()
            with open(fp, 'w', encoding='utf-8') as f:
                f.write("% Generated by Psalter Editor\n")
                f.write(content)
            messagebox.showinfo("成功", f"文件已生成:\n{fp}")
        except Exception as e:
            messagebox.showerror("错误", f"生成失败: {str(e)}")

    def compile_preview(self):
        if not self.content_items:
            messagebox.showwarning("提示", "内容为空，无法编译"); return
        
        req_files = ["main.tex", "psalter.sty"]
        missing = [f for f in req_files if not os.path.exists(os.path.join(self.base_dir, f))]
        if missing:
            messagebox.showerror("错误", f"缺失核心文件:\n{', '.join(missing)}")
            return

        build_dir = os.path.join(self.base_dir, "build")
        try:
            if os.path.exists(build_dir):
                for filename in os.listdir(build_dir):
                    file_path = os.path.join(build_dir, filename)
                    try:
                        if os.path.isfile(file_path) or os.path.islink(file_path): os.unlink(file_path)
                        elif os.path.isdir(file_path): shutil.rmtree(file_path)
                    except Exception as e: pass
            else:
                os.makedirs(build_dir)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建目录: {e}"); return

        try:
            shutil.copy2(os.path.join(self.base_dir, "psalter.sty"), build_dir)
            
            src_img = os.path.join(self.base_dir, "images")
            dst_img = os.path.join(build_dir, "images")
            if os.path.exists(src_img): shutil.copytree(src_img, dst_img, dirs_exist_ok=True)
            else: os.makedirs(dst_img)

            # 读取 main.tex 并注入标题
            with open(os.path.join(self.base_dir, "main.tex"), 'r', encoding='utf-8') as f:
                main_content = f.read()
            
            # 替换标题占位符
            main_content = main_content.replace("%TITLE_ZH%", self.title_data.get("title_zh", ""))
            main_content = main_content.replace("%TITLE_LAT%", self.title_data.get("title_lat", ""))
            main_content = main_content.replace("%EDITION_INFO%", self.title_data.get("edition", ""))
            main_content = main_content.replace("%FOOTER_TEXT%", self.title_data.get("footer", ""))

            with open(os.path.join(build_dir, "main.tex"), 'w', encoding='utf-8') as f:
                f.write(main_content)
            
            with open(os.path.join(build_dir, "body.tex"), 'w', encoding='utf-8') as f:
                f.write(self.get_latex_content())

        except Exception as e:
            messagebox.showerror("错误", f"准备文件失败: {e}"); return

        loading = tk.Toplevel(self.root)
        loading.title("编译中")
        loading.geometry("300x100")
        tk.Label(loading, text="正在调用 XeLaTeX 编译...", font=('Segoe UI', 10)).pack(expand=True)
        loading.update()
        
        try:
            if shutil.which("xelatex") is None:
                raise Exception("未找到 xelatex 命令。")

            cmd = ['xelatex', '-interaction=nonstopmode', 'main.tex']
            
            result = subprocess.run(cmd, cwd=build_dir, capture_output=True, text=True, encoding='utf-8', errors='replace')
            if result.returncode != 0:
                loading.destroy()
                self.show_error_log(result.stdout)
                return
            
            subprocess.run(cmd, cwd=build_dir, capture_output=True)
            loading.destroy()
            
            pdf_path = os.path.join(build_dir, "main.pdf")
            if os.path.exists(pdf_path):
                if platform.system() == 'Windows': os.startfile(pdf_path)
                elif platform.system() == 'Darwin': subprocess.call(('open', pdf_path))
                else: subprocess.call(('xdg-open', pdf_path))
            else:
                messagebox.showerror("失败", "编译似乎成功但没生成 PDF")
                
        except Exception as e:
            loading.destroy()
            messagebox.showerror("系统错误", str(e))

    def show_error_log(self, log_content):
        error_win = tk.Toplevel(self.root)
        error_win.title("编译失败 - 错误日志")
        error_win.geometry("800x600")
        
        lbl = tk.Label(error_win, text="LaTeX 编译出错。常见原因：\n1. 缺少字体 (Times New Roman, SimSun)\n2. 图片路径错误\n3. main.tex 里有冲突的 paracol\n以下是详细日志:", 
                       fg=S.DANGER, justify=tk.LEFT, padx=10, pady=10)
        lbl.pack(fill=tk.X)
        
        txt = scrolledtext.ScrolledText(error_win, font=('Consolas', 9))
        txt.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        txt.insert(tk.END, log_content)
        txt.see(tk.END)

class CustomContentDialog:
    def __init__(self, parent, item=None):
        self.result = None
        self.top = tk.Toplevel(parent)
        self.top.title("编辑内容" if item else "添加自定义内容")
        self.top.geometry("650x500")
        self.top.configure(bg=S.BG_DARK)
        self.top.transient(parent)
        self.top.grab_set()
        
        bf = tk.Frame(self.top, bg=S.BG_DARK)
        bf.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=15)
        
        cb = tk.Label(bf, text="取消", bg=S.BG_HOVER, fg=S.TEXT, font=('Segoe UI', 10), cursor='hand2', padx=20, pady=8)
        cb.pack(side=tk.RIGHT, padx=5)
        cb.bind('<Button-1>', lambda e: self.cancel())
        
        ob = tk.Label(bf, text="确定", bg=S.SUCCESS, fg=S.TEXT, font=('Segoe UI', 10), cursor='hand2', padx=20, pady=8)
        ob.pack(side=tk.RIGHT, padx=5)
        ob.bind('<Button-1>', lambda e: self.ok())

        content_frame = tk.Frame(self.top, bg=S.BG_DARK)
        content_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        tf = tk.Frame(content_frame, bg=S.BG_DARK)
        tf.pack(fill=tk.X, padx=15, pady=(15, 10))
        tk.Label(tf, text="格式类型:", bg=S.BG_DARK, fg=S.TEXT, font=('Segoe UI', 10)).pack(side=tk.LEFT)
        self.type_var = tk.StringVar()
        self.type_combo = ttk.Combobox(tf, textvariable=self.type_var,
            values=[f"{t[0]} - {t[1]}" for t in FORMAT_TYPES], width=45, font=('Segoe UI', 10))
        self.type_combo.pack(side=tk.LEFT, padx=10)
        
        tk.Label(content_frame, text="拉丁文/路径:", bg=S.BG_DARK, fg=S.ACCENT_LIGHT,
                font=('Segoe UI', 10, 'bold')).pack(anchor='w', padx=15, pady=(10, 5))
        lf = tk.Frame(content_frame, bg=S.BG_LIGHT)
        lf.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))
        self.latin_text = tk.Text(lf, height=4, wrap=tk.WORD, bg=S.BG_LIGHT, fg=S.TEXT,
            insertbackground=S.TEXT, font=('Segoe UI', 10), borderwidth=0, padx=8, pady=8)
        self.latin_text.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(content_frame, text="中文:", bg=S.BG_DARK, fg=S.ACCENT_LIGHT,
                font=('Segoe UI', 10, 'bold')).pack(anchor='w', padx=15, pady=(10, 5))
        cf = tk.Frame(content_frame, bg=S.BG_LIGHT)
        cf.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))
        self.chinese_text = tk.Text(cf, height=4, wrap=tk.WORD, bg=S.BG_LIGHT, fg=S.TEXT,
            insertbackground=S.TEXT, font=('Segoe UI', 10), borderwidth=0, padx=8, pady=8)
        self.chinese_text.pack(fill=tk.BOTH, expand=True)
        
        af = tk.Frame(content_frame, bg=S.BG_DARK)
        af.pack(fill=tk.X, padx=15, pady=10)
        tk.Label(af, text="附加参数:", bg=S.BG_DARK, fg=S.TEXT, font=('Segoe UI', 10)).pack(side=tk.LEFT)
        self.arg_entry = tk.Entry(af, width=35, bg=S.BG_LIGHT, fg=S.TEXT, insertbackground=S.TEXT,
            font=('Segoe UI', 10), relief='flat')
        self.arg_entry.pack(side=tk.LEFT, padx=10)
        tk.Label(af, text="(如对经编号等)", bg=S.BG_DARK, fg=S.TEXT_SEC, font=('Segoe UI', 9)).pack(side=tk.LEFT)
        
        if item:
            for i, (t, _) in enumerate(FORMAT_TYPES):
                if t == item.item_type: self.type_combo.current(i); break
            self.latin_text.insert(tk.END, item.latin)
            self.chinese_text.insert(tk.END, item.chinese)
            self.arg_entry.insert(0, item.arg)
    
    def ok(self):
        ts = self.type_var.get()
        if not ts: messagebox.showwarning("提示", "请选择格式类型"); return
        it = ts.split(" - ")[0]
        self.result = ContentItem(it, self.latin_text.get(1.0, tk.END).strip(),
            self.chinese_text.get(1.0, tk.END).strip(), self.arg_entry.get().strip())
        self.top.destroy()
    
    def cancel(self): self.top.destroy()

def main():
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except: pass
    root = tk.Tk()
    CSVEditorApp(root)
    root.mainloop()

if __name__ == '__main__': main()