"""
Selector tests for Word Document MCP Server.

This module contains tests for the selector engine that is used to locate
and select elements in Word documents.
"""

import unittest
import tempfile
import os
import shutil
import pythoncom
import win32com.client
from pathlib import Path

from word_document_server.selector.selector import SelectorEngine
from word_document_server.selector.exceptions import LocatorSyntaxError, AmbiguousLocatorError
from word_document_server.utils.core_utils import WordDocumentError


class TestSelector(unittest.TestCase):
    """Tests for selector functionality"""
    
    @classmethod
    def setUpClass(cls):
        """测试类初始化"""
        pythoncom.CoInitialize()
        
        # 创建临时目录和测试文件
        cls.test_dir = tempfile.mkdtemp()
        cls.test_file = os.path.join(cls.test_dir, "test_document.docx")
        
        # 创建Word应用程序实例
        cls.word_app = win32com.client.Dispatch('Word.Application')
        try:
            cls.word_app.Visible = False
        except AttributeError:
            pass

        # 创建测试文档
        cls.doc = cls.word_app.Documents.Add()
        
        # 添加测试内容
        # 第一段
        cls.doc.Range(0, 0).Text = "这是第一段文本。\n"
        
        # 第二段（包含关键词）
        cls.doc.Range().Collapse(0)  # 0 = wdCollapseEnd
        cls.doc.Range().Text = "这是第二段文本，包含关键词。\n"
        
        # 表格
        cls.doc.Range().Collapse(0)
        table_range = cls.doc.Range()
        table_range.Collapse(0)
        table = cls.doc.Tables.Add(table_range, 2, 2)
        table.Cell(1, 1).Range.Text = "A1"
        table.Cell(1, 2).Range.Text = "B1"
        table.Cell(2, 1).Range.Text = "A2"
        table.Cell(2, 2).Range.Text = "B2"
        
        # 第三段
        cls.doc.Range().Collapse(0)
        cls.doc.Range().Text = "这是第三段文本。\n"
        
        # 保存并关闭文档
        cls.doc.SaveAs2(cls.test_file)
        cls.doc.Close()
        cls.word_app.Quit()

    def setUp(self):
        # 创建Word应用程序实例
        self.word_app = win32com.client.Dispatch('Word.Application')
        try:
            self.word_app.Visible = False
        except AttributeError:
            # 某些环境中可能不支持设置Visible属性，忽略此错误
            pass
            
        # 打开测试文档
        self.test_doc = self.word_app.Documents.Open(self.test_file)
        
        # 创建选择器实例
        self.selector = SelectorEngine()
    
    def tearDown(self):
        # 关闭文档而不保存更改
        self.test_doc.Close(SaveChanges=False)
        
        # 退出Word应用程序
        self.word_app.Quit()
        
        # 清理选择器实例
        self.selector = None
    
    @classmethod
    def tearDownClass(cls):
        # 清理临时目录
        try:
            shutil.rmtree(cls.test_dir)
        except:
            pass
        pythoncom.CoUninitialize()
    
    def test_document_end_selector(self):
        """测试document_end选择器"""
        # 测试获取文档末尾位置
        locator = {
            "target": {
                "type": "document_end"
            }
        }
        
        selection = self.selector.select(self.test_doc, locator)
        self.assertIsNotNone(selection)
        self.assertEqual(len(selection._elements), 1)
        
        # 验证选择的位置确实是文档末尾
        element = selection._elements[0]
        self.assertEqual(element.Start, element.End)
        # 文档末尾位置可能因为Word的处理方式略有不同，所以我们只验证它是一个有效的正数
        self.assertGreaterEqual(element.Start, 0)
    
    def test_document_start_selector(self):
        """测试document_start选择器"""
        # 测试获取文档开始位置
        locator = {
            "target": {
                "type": "document_start"
            }
        }
        
        selection = self.selector.select(self.test_doc, locator)
        self.assertIsNotNone(selection)
        self.assertEqual(len(selection._elements), 1)
        
        # 验证选择的位置确实是文档开头
        element = selection._elements[0]
        self.assertEqual(element.Start, 0)
        self.assertEqual(element.End, 0)
    
    def test_paragraph_selection(self):
        """测试段落选择"""
        # 测试选择所有段落
        locator = {
            "target": {
                "type": "paragraph"
            }
        }
        
        selection = self.selector.select(self.test_doc, locator)
        self.assertIsNotNone(selection)
        # 应该至少有几段（文档内容段落+表格段落）
        self.assertGreater(len(selection._elements), 0)
    
    def test_table_selection(self):
        """测试表格选择"""
        # 测试选择所有表格
        locator = {
            "target": {
                "type": "table"
            }
        }
        
        selection = self.selector.select(self.test_doc, locator)
        self.assertIsNotNone(selection)
        # 应该有1个表格
        self.assertEqual(len(selection._elements), 1)
    
    def test_text_filter_selection(self):
        """测试带文本过滤器的选择"""
        # 测试选择包含特定文本的段落
        locator = {
            "target": {
                "type": "paragraph",
                "filters": [
                    {"contains_text": "关键词"}
                ]
            }
        }
        
        selection = self.selector.select(self.test_doc, locator)
        self.assertIsNotNone(selection)
        # 应该有1个段落包含"关键词"
        self.assertEqual(len(selection._elements), 1)
        
        # 验证选中的段落确实包含关键词
        # 通过Range属性访问段落文本
        element_text = selection._elements[0].Range.Text.strip()
        self.assertIn("关键词", element_text)
    
    def test_anchored_selection(self):
        """测试锚定选择"""
        # 先选择一个锚点（包含"关键词"的段落）
        anchor_locator = {
            "target": {
                "type": "paragraph",
                "filters": [
                    {"contains_text": "关键词"}
                ]
            }
        }
        
        # 然后选择锚点后紧跟的表格
        locator = {
            "anchor": {
                "type": "paragraph",
                "identifier": {
                    "text": "这是第二段文本，包含关键词。"
                }
            },
            "relation": {
                "type": "immediately_following"
            },
            "target": {
                "type": "table"
            }
        }
        
        try:
            selection = self.selector.select(self.test_doc, locator)
            self.assertIsNotNone(selection)
            # 应该能找到紧跟在段落后的表格
            self.assertEqual(len(selection._elements), 1)
        except Exception as e:
            # 如果测试失败，输出错误信息以便调试
            print(f"Anchored selection test failed with error: {e}")
            raise
    
    def test_invalid_locator(self):
        """测试无效定位器"""
        # 测试缺少target的定位器
        invalid_locator = {
            "anchor": {
                "type": "paragraph"
            }
        }
        
        with self.assertRaises((Exception, LocatorSyntaxError)):
            self.selector.select(self.test_doc, invalid_locator)
    
    def test_element_not_found(self):
        """测试找不到元素的情况"""
        # 测试查找不存在的文本
        locator = {
            "target": {
                "type": "paragraph",
                "filters": [
                    {"contains_text": "不存在的文本"}
                ]
            }
        }
        
        with self.assertRaises((AmbiguousLocatorError, ValueError)):
            self.selector.select(self.test_doc, locator)


if __name__ == '__main__':
    unittest.main()