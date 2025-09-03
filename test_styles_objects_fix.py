#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""测试脚本：验证styles_tools和objects_tools的修复结果

这个脚本验证以下修复：
1. styles_tools.py中的set_paragraph_style操作修复，解决'Selection' object is not iterable错误
2. objects_tools.py中的bookmark_operations修复，解决'SelectorEngine' object has no attribute 'select_ranges'错误

脚本会模拟必要的对象和方法，验证修复逻辑是否正确。
"""

import unittest
from unittest.mock import Mock, patch
import json
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 模拟COM对象
class MockDocument:
    def __init__(self):
        self.Styles = Mock()
        self.Bookmarks = Mock()
        self.Hyperlinks = Mock()
        
        # 模拟样式
        mock_style = Mock()
        mock_style.Name = "Heading 1"
        mock_style.NameLocal = "标题 1"
        self.Styles.return_value = mock_style
        
        # 模拟样式迭代
        def mock_styles_iter():
            yield mock_style
        self.Styles.__iter__ = mock_styles_iter
        
        # 模拟书签创建
        mock_bookmark = Mock()
        mock_bookmark.Name = "TestBookmark"
        self.Bookmarks.Add.return_value = mock_bookmark
        
        # 模拟超链接创建
        mock_hyperlink = Mock()
        mock_hyperlink.Index = 1
        self.Hyperlinks.Add.return_value = mock_hyperlink

class MockRange:
    def __init__(self):
        self.Start = 0
        self.End = 0
        self.Text = ""
    
    def InsertAfter(self, text):
        pass
    
    def Collapse(self, direction):
        pass

class MockSelection:
    def __init__(self):
        self._com_ranges = [MockRange()]

class MockParagraph:
    def __init__(self):
        self.Style = None
        self.Range = MockRange()

# 导入需要测试的模块
with patch('word_docx_tools.tools.styles_tools.SelectorEngine'):
    with patch('word_docx_tools.tools.objects_tools.SelectorEngine'):
        from word_docx_tools.tools.styles_tools import styles_tools
        from word_docx_tools.tools.objects_tools import objects_tools

class TestStyleAndObjectToolsFix(unittest.TestCase):
    
    def setUp(self):
        # 创建模拟上下文
        self.mock_ctx = Mock()
        self.mock_ctx.request_context.lifespan_context.get_active_document.return_value = MockDocument()
        self.mock_ctx.request_context.lifespan_context = Mock()
        self.mock_ctx.request_context.lifespan_context.get_active_document.return_value = MockDocument()
        
    def test_styles_tools_set_paragraph_style_fix(self):
        """测试styles_tools中的set_paragraph_style操作修复"""
        
        # 模拟SelectorEngine返回Selection对象
        def mock_select(document, locator):
            return MockSelection()
            
        with patch('word_docx_tools.tools.styles_tools.SelectorEngine.select', side_effect=mock_select):
            # 执行set_paragraph_style操作
            result = styles_tools(
                ctx=self.mock_ctx,
                operation_type="set_paragraph_style",
                style_name="Heading 1",
                locator={"type": "paragraph", "value": "1"}
            )
            
            # 验证结果
            self.assertIn('success', json.loads(result))
            self.assertTrue(json.loads(result)['success'])
        
        # 测试模拟返回单个段落对象的情况
        def mock_select_single(document, locator):
            return MockParagraph()
            
        with patch('word_docx_tools.tools.styles_tools.SelectorEngine.select', side_effect=mock_select_single):
            result = styles_tools(
                ctx=self.mock_ctx,
                operation_type="set_paragraph_style",
                style_name="Heading 1",
                locator={"type": "paragraph", "value": "1"}
            )
            
            self.assertIn('success', json.loads(result))
            self.assertTrue(json.loads(result)['success'])
        
        # 测试模拟返回段落列表的情况
        def mock_select_list(document, locator):
            return [MockParagraph(), MockParagraph()]
            
        with patch('word_docx_tools.tools.styles_tools.SelectorEngine.select', side_effect=mock_select_list):
            result = styles_tools(
                ctx=self.mock_ctx,
                operation_type="set_paragraph_style",
                style_name="Heading 1",
                locator={"type": "paragraph", "value": "1"}
            )
            
            self.assertIn('success', json.loads(result))
            self.assertTrue(json.loads(result)['success'])
    
    def test_objects_tools_bookmark_operations_fix(self):
        """测试objects_tools中的bookmark_operations操作修复"""
        
        # 模拟SelectorEngine返回Selection对象
        def mock_select(document, locator):
            return MockSelection()
            
        with patch('word_docx_tools.tools.objects_tools.SelectorEngine.select', side_effect=mock_select):
            # 模拟get_active_document
            with patch('word_docx_tools.tools.objects_tools.get_active_document', return_value=MockDocument()):
                # 模拟require_active_document_validation
                with patch('word_docx_tools.tools.objects_tools.require_active_document_validation'):
                    # 执行bookmark_operations操作
                    result = objects_tools(
                        ctx=self.mock_ctx,
                        operation_type="bookmark_operations",
                        sub_operation="create",
                        bookmark_name="TestBookmark",
                        locator={"type": "paragraph", "value": "1"}
                    )
                    
                    # 验证结果
                    self.assertIn('success', result)
                    self.assertTrue(result['success'])
    
    def test_objects_tools_hyperlink_operations_fix(self):
        """测试objects_tools中的hyperlink_operations操作修复"""
        
        # 模拟SelectorEngine返回Selection对象
        def mock_select(document, locator):
            return MockSelection()
            
        with patch('word_docx_tools.tools.objects_tools.SelectorEngine.select', side_effect=mock_select):
            # 模拟get_active_document
            with patch('word_docx_tools.tools.objects_tools.get_active_document', return_value=MockDocument()):
                # 模拟require_active_document_validation
                with patch('word_docx_tools.tools.objects_tools.require_active_document_validation'):
                    # 执行hyperlink_operations操作
                    result = objects_tools(
                        ctx=self.mock_ctx,
                        operation_type="hyperlink_operations",
                        sub_operation="create",
                        url="https://example.com",
                        locator={"type": "paragraph", "value": "1"}
                    )
                    
                    # 验证结果
                    self.assertIn('success', result)
                    self.assertTrue(result['success'])

if __name__ == "__main__":
    print("======= 测试 styles_tools 和 objects_tools 修复结果 =======")
    print("1. 测试 set_paragraph_style 操作修复 - 解决'Selection' object is not iterable错误")
    print("2. 测试 bookmark_operations 操作修复 - 解决'SelectorEngine' object has no attribute 'select_ranges'错误")
    print("3. 测试 hyperlink_operations 操作修复 - 同样修复select_ranges调用问题")
    print("\n开始测试...\n")
    
    # 运行测试
    unittest.main(verbosity=2)
    
    print("\n======= 测试完成 =======")
    print("如果所有测试通过，说明修复成功解决了以下问题：")
    print("1. styles_tools.set_paragraph_style 现在可以正确处理Selection对象")
    print("2. objects_tools.bookmark_operations 不再调用不存在的select_ranges方法")
    print("3. objects_tools.hyperlink_operations 也不再调用不存在的select_ranges方法")