#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
content_manager.py - 内容管理器
负责内容项的增删改查、文件加载与序列化
"""

from __future__ import annotations
import os
import re
import csv
import logging
from typing import List, Optional, Dict, Tuple, Iterator
from pathlib import Path

from .models import ContentItem, MultiLineContentItem
from .config import CONTENT_CATEGORIES
from .exceptions import FileError, ContentError

logger = logging.getLogger(__name__)


class FileContentLoader:
    """文件内容加载器"""
    
    def __init__(self, content_dir: str):
        self.content_dir = Path(content_dir)
        self._ensure_directories()
    
    def _ensure_directories(self) -> None:
        """确保所有分类目录存在"""
        for category in CONTENT_CATEGORIES:
            (self.content_dir / category).mkdir(parents=True, exist_ok=True)
        logger.debug(f"内容目录已初始化: {self.content_dir}")
    
    def get_available_files(self) -> Dict[str, List[Tuple[int, str]]]:
        """获取所有可用文件，按分类组织"""
        files: Dict[str, List[Tuple[int, str]]] = {cat: [] for cat in CONTENT_CATEGORIES}
        
        for category in files:
            cat_path = self.content_dir / category
            if not cat_path.exists():
                continue
            
            file_list: List[Tuple[int, str]] = []
            for filepath in cat_path.iterdir():
                if not filepath.suffix == '.txt':
                    continue
                
                # 提取数字用于排序
                nums = re.findall(r'\d+', filepath.name)
                sort_key = int(nums[0]) if nums else 999
                if len(nums) > 1:
                    sort_key = sort_key * 100 + int(nums[1])
                
                file_list.append((sort_key, filepath.name))
            
            file_list.sort(key=lambda x: x[0])
            files[category] = file_list
        
        return files
    
    def load_file_content(self, category: str, filename: str) -> List[ContentItem]:
        """加载单个文件的内容"""
        filepath = self.content_dir / category / filename
        items: List[ContentItem] = []
        
        if not filepath.exists():
            logger.warning(f"文件不存在: {filepath}")
            return items
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                for line_num, line in enumerate(f, 1):
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue
                    
                    parts = line.split('|')
                    if len(parts) >= 3:
                        items.append(ContentItem(
                            item_type=parts[0],
                            latin=parts[1],
                            chinese=parts[2],
                            arg=parts[3] if len(parts) > 3 else ""
                        ))
                    else:
                        logger.warning(f"跳过格式错误的行 {filepath}:{line_num}")
            
            logger.debug(f"已加载 {len(items)} 项内容: {filepath}")
            return items
            
        except UnicodeDecodeError as e:
            raise FileError(f"文件编码错误: {filename}", str(filepath), str(e))
        except IOError as e:
            raise FileError(f"读取文件失败: {filename}", str(filepath), str(e))
    
    def load_file_as_multiline(self, category: str, filename: str) -> Optional[MultiLineContentItem]:
        """将文件作为多行内容块加载"""
        items = self.load_file_content(category, filename)
        if not items:
            return None
        
        return MultiLineContentItem(
            item_type="",
            source_file=filename,
            items=items
        )


class ContentManager:
    """内容管理器 - 管理所有内容项"""
    
    def __init__(self, content_dir: str):
        self.loader = FileContentLoader(content_dir)
        self._items: List[ContentItem] = []
        logger.info(f"ContentManager 初始化完成: {content_dir}")
    
    @property
    def items(self) -> List[ContentItem]:
        """获取所有内容项（只读视图）"""
        return self._items.copy()
    
    @property
    def count(self) -> int:
        """获取内容项数量"""
        return len(self._items)
    
    def is_empty(self) -> bool:
        """检查是否为空"""
        return len(self._items) == 0
    
    def __iter__(self) -> Iterator[ContentItem]:
        """支持迭代"""
        return iter(self._items)
    
    def __len__(self) -> int:
        """支持len()"""
        return len(self._items)
    
    def __getitem__(self, index: int) -> ContentItem:
        """支持索引访问"""
        return self._items[index]
    
    # ========== CRUD 操作 ==========
    
    def add(self, item: ContentItem) -> None:
        """添加内容项"""
        self._items.append(item)
        logger.debug(f"添加内容项: {item.item_type}")
    
    def insert(self, index: int, item: ContentItem) -> None:
        """在指定位置插入内容项"""
        if 0 <= index <= len(self._items):
            self._items.insert(index, item)
            logger.debug(f"在位置 {index} 插入内容项: {item.item_type}")
        else:
            raise ContentError(f"无效的插入位置: {index}")
    
    def remove(self, index: int) -> Optional[ContentItem]:
        """移除指定位置的内容项"""
        if 0 <= index < len(self._items):
            item = self._items.pop(index)
            logger.debug(f"移除内容项: {item.item_type}")
            return item
        return None
    
    def get(self, index: int) -> Optional[ContentItem]:
        """获取指定位置的内容项"""
        if 0 <= index < len(self._items):
            return self._items[index]
        return None
    
    def update(self, index: int, item: ContentItem) -> bool:
        """更新指定位置的内容项"""
        if 0 <= index < len(self._items):
            old_type = self._items[index].item_type
            self._items[index] = item
            logger.debug(f"更新内容项: {old_type} -> {item.item_type}")
            return True
        return False
    
    def move_up(self, index: int) -> bool:
        """上移内容项"""
        if 0 < index < len(self._items):
            self._items[index], self._items[index - 1] = \
                self._items[index - 1], self._items[index]
            return True
        return False
    
    def move_down(self, index: int) -> bool:
        """下移内容项"""
        if 0 <= index < len(self._items) - 1:
            self._items[index], self._items[index + 1] = \
                self._items[index + 1], self._items[index]
            return True
        return False
    
    def clear(self) -> None:
        """清空所有内容项"""
        count = len(self._items)
        self._items.clear()
        logger.info(f"已清空 {count} 项内容")
    
    # ========== 批量操作 ==========
    
    def get_flat_items(self) -> List[ContentItem]:
        """获取扁平化的内容项列表（展开MultiLineContentItem）"""
        flat: List[ContentItem] = []
        for item in self._items:
            if isinstance(item, MultiLineContentItem):
                flat.extend(item.get_flat_items())
            else:
                flat.append(item)
        return flat
    
    def add_from_file(self, category: str, filename: str) -> bool:
        """从文件加载内容并添加"""
        try:
            multi_item = self.loader.load_file_as_multiline(category, filename)
            if multi_item:
                self._items.append(multi_item)
                logger.info(f"已从文件加载内容: {category}/{filename}")
                return True
            return False
        except FileError as e:
            logger.error(f"加载文件失败: {e}")
            return False
    
    def get_available_files(self) -> Dict[str, List[Tuple[int, str]]]:
        """获取可用的内容文件列表"""
        return self.loader.get_available_files()
    
    # ========== 序列化方法 ==========
    
    def to_csv_rows(self) -> List[List[str]]:
        """转换为CSV行列表"""
        rows: List[List[str]] = []
        for item in self._items:
            if isinstance(item, MultiLineContentItem):
                rows.extend(item.to_csv_rows())
            else:
                rows.append(item.to_csv_row())
        return rows
    
    def save_to_csv(self, filepath: str) -> None:
        """保存到CSV文件"""
        try:
            with open(filepath, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                for row in self.to_csv_rows():
                    writer.writerow(row)
            logger.info(f"已保存CSV: {filepath}")
        except IOError as e:
            raise FileError(f"保存CSV失败", filepath, str(e))
    
    def load_from_csv(self, filepath: str) -> None:
        """从CSV文件加载"""
        try:
            self._items.clear()
            with open(filepath, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                for row in reader:
                    if len(row) >= 3:
                        self._items.append(ContentItem(
                            item_type=row[0],
                            latin=row[1],
                            chinese=row[2],
                            arg=row[3] if len(row) > 3 else ""
                        ))
            logger.info(f"已从CSV加载 {len(self._items)} 项: {filepath}")
        except IOError as e:
            raise FileError(f"加载CSV失败", filepath, str(e))
    
    def to_preview_text(self) -> str:
        """生成预览文本（CSV格式）"""
        lines: List[str] = []
        for row in self.to_csv_rows():
            lines.append(",".join(f'"{x}"' for x in row))
        return "\n".join(lines)
