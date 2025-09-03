#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""测试image_tools.py的修复效果

这个脚本用于验证image_tools.py中插入图片功能的修复是否成功解决了"AddPicture() got an unexpected keyword argument 'Range'"错误。
"""

import os
import sys
import json
import unittest
import asyncio
from unittest.mock import MagicMock, patch

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 导入需要测试的模块
from word_docx_tools.tools.image_tools import image_tools
from word_docx_tools.mcp_service.core_utils import WordDocumentError
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from word_docx_tools.utils.app_context import AppContext

# 创建异步测试的辅助类
class AsyncTestCase(unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)
        
    def tearDown(self):
        self.loop.close()
        super().tearDown()
        
    def async_test(self, coro):
        return self.loop.run_until_complete(coro)

class TestImageToolsFix(AsyncTestCase):
    
    def setUp(self):
        super().setUp()
        # 创建模拟对象
        self.mock_document = MagicMock()
        self.mock_inline_shape = MagicMock()
        self.mock_inline_shape.Index = 1
        self.mock_document.InlineShapes.AddPicture.return_value = self.mock_inline_shape
        
        # 创建模拟的范围对象
        self.mock_range = MagicMock()
        
        # 创建模拟的选择器引擎
        self.mock_selection = MagicMock()
        self.mock_selection._com_ranges = [self.mock_range]
        
        # 设置AppContext
        self.app_context = AppContext()
        self.app_context.get_active_document = MagicMock(return_value=self.mock_document)
        
        # 设置Context
        self.mock_request_context = MagicMock()
        self.mock_request_context.lifespan_context = self.app_context
        self.context = Context[ServerSession, AppContext](
            request_context=self.mock_request_context,
            session=MagicMock(spec=ServerSession)
        )
        
        # 创建一个临时测试图片文件
        self.test_image_path = os.path.join(os.path.dirname(__file__), "test_image.svg")
        with open(self.test_image_path, "w") as f:
            f.write('<?xml version="1.0" encoding="UTF-8"?><svg width="100" height="100" xmlns="http://www.w3.org/2000/svg"><circle cx="50" cy="50" r="40" fill="red" /></svg>')
    
    def tearDown(self):
        # 清理临时测试图片文件
        if os.path.exists(self.test_image_path):
            os.remove(self.test_image_path)
    
    @patch('word_docx_tools.tools.image_tools.SelectorEngine')
    @patch('os.path.exists')
    def test_insert_image_success(self, mock_exists, mock_selector_engine):
        """测试插入图片成功的情况"""
        # 配置mock
        mock_exists.return_value = True
        mock_selector_engine.return_value.select.return_value = self.mock_selection
        
        # 执行测试
        async def run_test():
            result = await image_tools(
                ctx=self.context,
                operation_type="insert",
                image_path=self.test_image_path,
                locator={"type": "document_end"},
                position="after"
            )
            
            # 验证结果
            result_data = json.loads(result)
            self.assertTrue(result_data["success"])
            self.assertEqual(result_data["result"]["shape_id"], 1)
            
            # 验证调用了正确的方法
            self.mock_document.InlineShapes.AddPicture.assert_called_once()
            call_args = self.mock_document.InlineShapes.AddPicture.call_args
            self.assertEqual(call_args[1]["FileName"], self.test_image_path)
            self.assertEqual(call_args[1]["LinkToFile"], False)
            self.assertEqual(call_args[1]["SaveWithDocument"], True)
        
        self.async_test(run_test())
        
    @patch('word_docx_tools.tools.image_tools.SelectorEngine')
    @patch('os.path.exists')
    def test_insert_image_fallback_method(self, mock_exists, mock_selector_engine):
        """测试备用插入图片方法的情况"""
        # 配置mock - 第一次调用失败，第二次调用成功
        mock_exists.return_value = True
        mock_selector_engine.return_value.select.return_value = self.mock_selection
        self.mock_document.InlineShapes.AddPicture.side_effect = [
            Exception("First method failed"),
            self.mock_inline_shape
        ]
        
        # 执行测试
        async def run_test():
            result = await image_tools(
                ctx=self.context,
                operation_type="insert",
                image_path=self.test_image_path,
                locator={"type": "document_end"},
                position="after"
            )
            
            # 验证结果
            result_data = json.loads(result)
            self.assertTrue(result_data["success"])
            
            # 验证调用了两次AddPicture，第二次没有Range参数
            self.assertEqual(self.mock_document.InlineShapes.AddPicture.call_count, 2)
            second_call_args = self.mock_document.InlineShapes.AddPicture.call_args_list[1]
            self.assertNotIn("Range", second_call_args[1])
            
            # 验证调用了Select方法
            self.mock_range.Select.assert_called_once()
        
        self.async_test(run_test())
    
    @patch('word_docx_tools.tools.image_tools.SelectorEngine')
    @patch('os.path.exists')
    def test_insert_image_failure(self, mock_exists, mock_selector_engine):
        """测试插入图片失败的情况"""
        # 配置mock - 两次调用都失败
        mock_exists.return_value = True
        mock_selector_engine.return_value.select.return_value = self.mock_selection
        self.mock_document.InlineShapes.AddPicture.side_effect = [
            Exception("First method failed"),
            Exception("Second method failed")
        ]
        
        # 执行测试并验证异常
        async def run_test():
            with self.assertRaises(WordDocumentError) as context:
                await image_tools(
                    ctx=self.context,
                    operation_type="insert",
                    image_path=self.test_image_path,
                    locator={"type": "document_end"},
                    position="after"
                )
            
            # 验证异常信息
            self.assertIn("Failed to insert image", str(context.exception))
        
        self.async_test(run_test())

if __name__ == "__main__":
    unittest.main()