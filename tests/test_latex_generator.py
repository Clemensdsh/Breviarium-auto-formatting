#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tests/test_latex_generator.py - LaTeX生成器单元测试
"""

import unittest
import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from psalter_generator.models import ContentItem
from psalter_generator.latex_generator import TexGenerator, GeneratorState, ColumnMode
from psalter_generator.config import TexMappingConfig


class TestGeneratorState(unittest.TestCase):
    """测试生成器状态"""
    
    def test_initial_state(self):
        state = GeneratorState()
        self.assertEqual(state.mode, ColumnMode.DOUBLE)
        self.assertTrue(state.in_paracol)
        self.assertFalse(state.is_single_col)
    
    def test_toggle_to_single(self):
        state = GeneratorState()
        new_mode = state.toggle_mode()
        self.assertEqual(new_mode, ColumnMode.SINGLE)
        self.assertTrue(state.is_single_col)
        self.assertFalse(state.in_paracol)
    
    def test_toggle_back_to_double(self):
        state = GeneratorState()
        state.toggle_mode()
        state.toggle_mode()
        self.assertEqual(state.mode, ColumnMode.DOUBLE)
        self.assertFalse(state.is_single_col)
    
    def test_reset(self):
        state = GeneratorState()
        state.set_single()
        state.reset()
        self.assertEqual(state.mode, ColumnMode.DOUBLE)
        self.assertTrue(state.in_paracol)


class TestTexGenerator(unittest.TestCase):
    """测试LaTeX生成器"""
    
    def setUp(self):
        self.generator = TexGenerator()
    
    def test_empty_items(self):
        result = self.generator.generate([])
        self.assertIn(r"\begin{paracol}{2}", result)
        self.assertIn(r"\end{paracol}", result)
    
    def test_single_verse(self):
        items = [ContentItem("verse", "Dominus", "上主", "")]
        result = self.generator.generate(items)
        self.assertIn(r"\psVerse{Dominus}{上主}", result)
    
    def test_rule(self):
        items = [ContentItem("rule", "", "", "")]
        result = self.generator.generate(items)
        self.assertIn(r"\psThinRule", result)
    
    def test_pagebreak_double_col(self):
        items = [ContentItem("pagebreak", "", "", "")]
        result = self.generator.generate(items)
        self.assertIn(r"\psPageBreak", result)
    
    def test_singlecol_toggle(self):
        items = [
            ContentItem("singlecol", "", "", ""),
            ContentItem("verse", "", "单栏内容", ""),
            ContentItem("singlecol", "", "", ""),
        ]
        result = self.generator.generate(items)
        self.assertIn(r"\psEnterSingleCol", result)
        self.assertIn(r"\psSingleVerse{单栏内容}", result)
        self.assertIn(r"\psExitSingleCol", result)
    
    def test_h1cap(self):
        items = [ContentItem("h1cap", "Latin Title", "中文标题", "")]
        result = self.generator.generate(items)
        self.assertIn(r"\psHeaderOneCap{Latin Title}{中文标题}", result)
    
    def test_antiphon(self):
        items = [ContentItem("antiphon", "Ant. lat", "对经中文", "")]
        result = self.generator.generate(items)
        self.assertIn(r"\psAntiphonRepeat{Ant. lat}{对经中文}", result)
    
    def test_unknown_type(self):
        items = [ContentItem("unknown_type", "a", "b", "")]
        result = self.generator.generate(items)
        self.assertIn("% 未知类型", result)
    
    def test_tocstart(self):
        items = [ContentItem("tocstart", "", "", "")]
        result = self.generator.generate(items)
        self.assertIn(r"\psPrintToc", result)
        self.assertIn(r"\pagenumbering{arabic}", result)
    
    def test_score_file_in_double_col(self):
        """测试在双栏模式下插入乐谱文件"""
        items = [ContentItem("score", "gabc/antiphon.gabc", "对经", "")]
        result = self.generator.generate(items)
        # 应该先切换到单栏
        self.assertIn(r"\psEnterSingleCol", result)
        # 插入乐谱
        self.assertIn(r"\gregorioscore{gabc/antiphon}", result)
        # 然后切换回双栏
        self.assertIn(r"\psExitSingleCol", result)
    
    def test_score_inline_in_double_col(self):
        """测试在双栏模式下插入内联乐谱"""
        gabc_code = "(c4)Al(f)le(gf)lú(gh)ia.(g.)"
        items = [ContentItem("score", gabc_code, "", "inline")]
        result = self.generator.generate(items)
        self.assertIn(r"\psEnterSingleCol", result)
        self.assertIn(r"\gabcsnippet{" + gabc_code + "}", result)
        self.assertIn(r"\psExitSingleCol", result)
    
    def test_score_in_single_col(self):
        """测试在单栏模式下插入乐谱（不需要切换）"""
        items = [
            ContentItem("singlecol", "", "", ""),  # 先切换到单栏
            ContentItem("score", "gabc/test", "", ""),
            ContentItem("singlecol", "", "", ""),  # 切换回双栏
        ]
        result = self.generator.generate(items)
        # 乐谱部分不应该有额外的切换
        lines = result.split('\n')
        enter_count = sum(1 for l in lines if 'psEnterSingleCol' in l)
        exit_count = sum(1 for l in lines if 'psExitSingleCol' in l)
        # 应该只有一对切换（由singlecol产生）
        self.assertEqual(enter_count, 1)
        self.assertEqual(exit_count, 1)
    
    def test_score_with_title(self):
        """测试乐谱带标题"""
        items = [ContentItem("score", "gabc/kyrie", "Kyrie XVII", "")]
        result = self.generator.generate(items)
        self.assertIn(r"\psSingleRubric{Kyrie XVII}", result)


class TestTexMappingConfig(unittest.TestCase):
    """测试配置管理"""
    
    def test_default_mapping(self):
        config = TexMappingConfig()
        self.assertTrue(config.has("verse"))
        self.assertTrue(config.has("h1"))
        self.assertFalse(config.has("nonexistent"))
    
    def test_get_command(self):
        config = TexMappingConfig()
        cmd = config.get("verse")
        self.assertIsNotNone(cmd)
        self.assertIn("psVerse", cmd.double_col)
        self.assertIn("psSingleVerse", cmd.single_col)
    
    def test_add_command(self):
        config = TexMappingConfig()
        config.add("custom", r"\customDouble{{{l}}}", r"\customSingle{{{c}}}")
        self.assertTrue(config.has("custom"))
        cmd = config.get("custom")
        self.assertIn("customDouble", cmd.double_col)
    
    def test_remove_command(self):
        config = TexMappingConfig()
        self.assertTrue(config.remove("verse"))
        self.assertFalse(config.has("verse"))
        self.assertFalse(config.remove("nonexistent"))
    
    def test_to_dict(self):
        config = TexMappingConfig()
        d = config.to_dict()
        self.assertIsInstance(d, dict)
        self.assertIn("verse", d)
        self.assertIn("double_col", d["verse"])


class TestContentItem(unittest.TestCase):
    """测试内容项模型"""
    
    def test_to_csv_row(self):
        item = ContentItem("verse", "latin", "chinese", "arg")
        row = item.to_csv_row()
        self.assertEqual(row, ["verse", "latin", "chinese", "arg"])
    
    def test_display_text_normal(self):
        item = ContentItem("verse", "short", "短", "")
        text = item.get_display_text()
        self.assertIn("[verse]", text)
        self.assertIn("short", text)
    
    def test_display_text_truncate(self):
        item = ContentItem("verse", "a" * 30, "中" * 20, "")
        text = item.get_display_text()
        self.assertIn("...", text)
    
    def test_display_text_special(self):
        item = ContentItem("rule", "", "", "")
        text = item.get_display_text()
        self.assertEqual(text, "[分隔线]")
    
    def test_validate_success(self):
        item = ContentItem("verse", "lat", "chi", "")
        ok, msg = item.validate()
        self.assertTrue(ok)
    
    def test_validate_empty_type(self):
        item = ContentItem("", "lat", "chi", "")
        ok, msg = item.validate()
        self.assertFalse(ok)
    
    def test_validate_special_type(self):
        item = ContentItem("rule", "", "", "")
        ok, msg = item.validate()
        self.assertTrue(ok)


if __name__ == '__main__':
    unittest.main(verbosity=2)
