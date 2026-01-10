#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
config.py - LaTeX命令映射与配置
支持从外部JSON文件加载，便于扩展
"""

from __future__ import annotations
import json
import os
import logging
from typing import Dict, Optional, Any
from dataclasses import dataclass

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class TexCommand:
    """LaTeX命令配置（不可变）"""
    double_col: str  # 双栏模式命令
    single_col: str  # 单栏模式命令
    
    def format(self, is_single: bool, **kwargs: str) -> str:
        """根据模式格式化命令"""
        template = self.single_col if is_single else self.double_col
        return template.format(**kwargs)


# 默认的LaTeX命令映射
_DEFAULT_MAPPING: Dict[str, Dict[str, str]] = {
    'h1': {
        'double_col': r'\psHeaderOne{{{l}}}{{{c}}}',
        'single_col': r'\psSingleHeaderOne{{{c}}}'
    },
    'h1cap': {
        'double_col': r'\psHeaderOneCap{{{l}}}{{{c}}}',
        'single_col': r'\psSingleHeaderOneCap{{{l}}}{{{c}}}'
    },
    'h1lowercase': {
        'double_col': r'\psHeaderOneLowercase{{{l}}}{{{c}}}',
        'single_col': r'\psSingleHeaderOneLowercase{{{l}}}{{{c}}}'
    },
    'h2': {
        'double_col': r'\psHeaderTwo{{{l}}}{{{c}}}',
        'single_col': r'\psSingleHeaderTwo{{{c}}}'
    },
    'h3': {
        'double_col': r'\psHeaderThree{{{l}}}{{{c}}}',
        'single_col': r'\psSingleHeaderThree{{{c}}}'
    },
    'psalmtitle': {
        'double_col': r'\psPsalmTitle{{{l}}}{{{c}}}',
        'single_col': r'\psSinglePsalmTitle{{{c}}}'
    },
    'canticletitle': {
        'double_col': r'\psCanticleTitle{{{l}}}{{{c}}}',
        'single_col': r'\psSingleCanticleTitle{{{c}}}'
    },
    'hymntitle': {
        'double_col': r'\psHymnTitle{{{l}}}{{{c}}}',
        'single_col': r'\psSingleHymnTitle{{{c}}}'
    },
    'hymnheader': {
        'double_col': r'\psHymnHeader{{{l}}}{{{c}}}',
        'single_col': r'\psSingleHymnHeader{{{c}}}'
    },
    'antiphon': {
        'double_col': r'\psAntiphonRepeat{{{l}}}{{{c}}}',
        'single_col': r'\psSingleAntiphon{{{c}}}'
    },
    'dropcap': {
        'double_col': r'\psVerseDropcap{{{l}}}{{{c}}}',
        'single_col': r'\psSingleVerseDropcap{{{c}}}'
    },
    'verse': {
        'double_col': r'\psVerse{{{l}}}{{{c}}}',
        'single_col': r'\psSingleVerse{{{c}}}'
    },
    'versedropcap': {
        'double_col': r'\psVerseDropcap{{{l}}}{{{c}}}',
        'single_col': r'\psSingleVerseDropcap{{{c}}}'
    },
    'gloria': {
        'double_col': r'\psGloria{{{l}}}{{{c}}}',
        'single_col': r'\psSingleGloria{{{c}}}'
    },
    'rubric': {
        'double_col': r'\psRubric{{{l}}}{{{c}}}',
        'single_col': r'\psSingleRubric{{{c}}}'
    },
    'V': {
        'double_col': r'\psVR{{V}}{{{l}}}{{{c}}}',
        'single_col': r'\psSingleVR{{V}}{{{c}}}'
    },
    'R': {
        'double_col': r'\psVR{{R}}{{{l}}}{{{c}}}',
        'single_col': r'\psSingleVR{{R}}{{{c}}}'
    },
    'hymn': {
        'double_col': r'\psHymnStanza{{{l}}}{{{c}}}',
        'single_col': r'\psSingleHymnStanza{{{c}}}'
    },
    'hymnstanza': {
        'double_col': r'\psHymnStanza{{{l}}}{{{c}}}',
        'single_col': r'\psSingleHymnStanza{{{c}}}'
    },
    'antiphonrepeat': {
        'double_col': r'\psAntiphonRepeat{{{l}}}{{{c}}}',
        'single_col': r'\psSingleAntiphon{{{c}}}'
    },
    'capit': {
        'double_col': r'\psCapit{{{l}}}{{{c}}}',
        'single_col': r'\psSingleCapit{{{c}}}'
    },
    'capitheader': {
        'double_col': r'\psCapitHeader{{{l}}}{{{c}}}',
        'single_col': r'\psSingleCapitHeader{{{c}}}'
    },
    'scriptureref': {
        'double_col': r'\psScriptureRef{{{l}}}{{{c}}}',
        'single_col': r'\psSingleScriptureRef{{{c}}}'
    },
    'collect': {
        'double_col': r'\psCollect{{{l}}}{{{c}}}',
        'single_col': r'\psSingleCollect{{{c}}}'
    },
    'lesson': {
        'double_col': r'\psLesson{{{l}}}{{{c}}}',
        'single_col': r'\psSingleLesson{{{c}}}'
    },
    'text': {
        'double_col': r'\psText{{{l}}}{{{c}}}',
        'single_col': r'\psSingleText{{{c}}}'
    },
    'rule': {
        'double_col': r'\psThinRule',
        'single_col': r'\psSingleThinRule'
    },
    'thickrule': {
        'double_col': r'\psThickRule',
        'single_col': r'\psSingleThickRule'
    },
    # 新增宏类型
    'imprimatur': {
        'double_col': r'\psImprimatur{{{l}}}{{{a}}}{{{c}}}',
        'single_col': r'\psImprimatur{{{l}}}{{{a}}}{{{c}}}'
    },
    'nihilobstat': {
        'double_col': r'\psNihilObstat{{{l}}}{{{a}}}',
        'single_col': r'\psNihilObstat{{{l}}}{{{a}}}'
    },
    'redinline': {
        'double_col': r'\psRed{{{l}}}',
        'single_col': r'\psRed{{{c}}}'
    },
    'scriptureinline': {
        'double_col': r'\psScriptureInline{{{l}}}',
        'single_col': r'\psScriptureInline{{{c}}}'
    },
}


def _build_default_mapping() -> Dict[str, TexCommand]:
    """构建默认映射"""
    return {
        k: TexCommand(v['double_col'], v['single_col'])
        for k, v in _DEFAULT_MAPPING.items()
    }


class TexMappingConfig:
    """LaTeX命令映射配置管理器"""
    
    def __init__(self, mapping: Optional[Dict[str, TexCommand]] = None):
        self._mapping = mapping or _build_default_mapping()
    
    def get(self, item_type: str) -> Optional[TexCommand]:
        """获取指定类型的命令配置"""
        return self._mapping.get(item_type)
    
    def has(self, item_type: str) -> bool:
        """检查是否存在指定类型"""
        return item_type in self._mapping
    
    def add(self, item_type: str, double_col: str, single_col: str) -> None:
        """添加新的命令映射"""
        self._mapping[item_type] = TexCommand(double_col, single_col)
    
    def remove(self, item_type: str) -> bool:
        """移除命令映射"""
        if item_type in self._mapping:
            del self._mapping[item_type]
            return True
        return False
    
    def all_types(self) -> list:
        """获取所有类型名称"""
        return list(self._mapping.keys())
    
    def to_dict(self) -> Dict[str, Dict[str, str]]:
        """转换为可序列化的字典"""
        return {
            k: {"double_col": v.double_col, "single_col": v.single_col}
            for k, v in self._mapping.items()
        }
    
    def save_to_json(self, filepath: str) -> None:
        """保存到JSON文件"""
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)
            logger.info(f"配置已保存到: {filepath}")
        except IOError as e:
            logger.error(f"保存配置失败: {e}")
            raise
    
    @classmethod
    def load_from_json(cls, filepath: str) -> TexMappingConfig:
        """从JSON文件加载"""
        if not os.path.exists(filepath):
            logger.warning(f"配置文件不存在，使用默认配置: {filepath}")
            return cls()
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            mapping = {
                k: TexCommand(v["double_col"], v["single_col"])
                for k, v in data.items()
            }
            logger.info(f"已加载配置: {filepath}")
            return cls(mapping)
        except (IOError, json.JSONDecodeError, KeyError) as e:
            logger.error(f"加载配置失败，使用默认配置: {e}")
            return cls()


# 内容文件分类配置
CONTENT_CATEGORIES: Dict[str, str] = {
    "psalms": "圣咏 (Psalms)",
    "canticles": "圣歌 (Canticles)",
    "hymns": "赞美诗 (Hymns)",
    "antiphons": "对经 (Antiphons)",
    "lessons": "读经 (Lessons)",
    "responsories": "答唱咏 (Responsories)",
    "collects": "集祷经 (Collects)",
    "common": "通用文本 (Common)",
    "test": "测试 (Test)",
}


# 应用配置
@dataclass
class AppConfig:
    """应用程序配置"""
    # 编译相关
    compiler_timeout: int = 120
    compile_runs: int = 3  # 增加到3次以确保目录正确生成
    
    # 构建目录
    build_dir_name: str = "build"
    
    # 必需文件
    required_files: tuple = ("main.tex", "psalter.sty")
    resource_directories: tuple = ("images",)
    
    @classmethod
    def load_from_json(cls, filepath: str) -> AppConfig:
        """从JSON加载配置"""
        if not os.path.exists(filepath):
            return cls()
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return cls(**{k: v for k, v in data.items() if hasattr(cls, k)})
        except Exception as e:
            logger.error(f"加载应用配置失败: {e}")
            return cls()
