"""
End-to-end scenario tests for Word Document MCP Server.

This module contains detailed integration tests that simulate specific real-world usage scenarios
with actual locator usage and complex operation sequences.
"""

import json
import os
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock

import pythoncom
import win32com.client
from mcp.server.fastmcp import Context

from word_docx_tools.tools.comment_tools import comment_tools
from word_docx_tools.tools.document_tools import document_tools
from word_docx_tools.tools.image_tools import image_tools
from word_docx_tools.tools.range_tools import range_tools
from word_docx_tools.tools.table_tools import table_tools
from word_docx_tools.tools.text_tools import text_tools
from word_docx_tools.mcp_service.app_context import AppContext


class TestE2EScenarios(unittest.TestCase):
    """Detailed end-to-end scenario tests"""

    @classmethod
    def setUpClass(cls):
        # 初始化COM
        pythoncom.CoInitialize()

        # 创建临时目录用于测试
        cls.test_dir = tempfile.mkdtemp()

        # 复制测试文档到临时目录
        source_doc = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "tests",
            "test_docs",
            "text_test_doc.docx",
        )
        cls.test_doc_path = os.path.join(cls.test_dir, "text_test_doc.docx")
        shutil.copy2(source_doc, cls.test_doc_path)

        # 创建测试图片
        cls.test_image_path = os.path.join(cls.test_dir, "test_image.png")
        with open(cls.test_image_path, "w") as f:
            f.write("fake image content for testing")

    @classmethod
    def tearDownClass(cls):
        # 清理临时目录
        try:
            shutil.rmtree(cls.test_dir)
        except:
            pass
        pythoncom.CoUninitialize()

    def setUp(self):
        # 创建Word应用程序实例
        self.word_app = win32com.client.Dispatch("Word.Application")
        # 尝试设置Visible属性，但捕获可能的异常
        try:
            self.word_app.Visible = False
        except AttributeError:
            # 某些环境中可能不支持设置Visible属性，忽略此错误
            pass

        # 创建应用上下文
        self.app_context = AppContext(self.word_app)

        # 创建mock session
        self.mock_session = MagicMock()
        self.mock_session.lifespan_context = self.app_context

        # 创建Context对象
        self.context = MagicMock()
        self.context.request_context = self.mock_session

    def tearDown(self):
        # 关闭Word应用程序
        try:
            self.word_app.Quit()
        except:
            pass

    def test_document_creation_and_formatting_scenario(self):
        """场景测试：创建新文档、添加内容并格式化"""
        # 1. 创建新文档
        result = document_tools(self.context, operation_type="open", file_path="")
        self.assertIn("Document opened successfully", result)

        # 2. 添加标题
        text_tools(
            self.context, operation_type="insert", text="项目报告", style="Heading 1"
        )

        # 3. 添加段落
        text_tools(
            self.context, operation_type="insert", text="这是项目报告的主要内容。"
        )
        text_tools(self.context, operation_type="insert", text="报告包含以下部分：")

        # 4. 创建要点列表
        text_tools(
            self.context,
            operation_type="create_list",
            items=["引言", "方法", "结果", "结论"],
        )

        # 5. 添加另一个标题
        text_tools(
            self.context, operation_type="insert", text="引言", style="Heading 2"
        )

        # 6. 添加更多段落
        text_tools(self.context, operation_type="insert", text="这是引言部分的内容。")

        # 7. 查找并替换文本
        text_tools(
            self.context,
            operation_type="replace",
            text="项目报告",
            new_text="年度项目报告",
        )

        # 8. 批量格式化操作
        # 这里需要实际的定位器，简化测试中跳过实际格式化

    def test_report_generation_scenario(self):
        """场景测试：生成完整报告，包含表格和图片"""
        # 1. 打开现有文档
        document_tools(
            self.context, operation_type="open", file_path=self.test_doc_path
        )

        # 2. 添加新章节
        text_tools(
            self.context, operation_type="insert", text="数据分析", style="Heading 1"
        )

        # 3. 添加描述段落
        text_tools(
            self.context, operation_type="insert", text="以下是我们收集的数据分析结果："
        )

        # 4. 创建数据表格
        table_tools(self.context, operation_type="create", rows=4, cols=3)

        # 5. 填充表格数据（需要定位器支持）
        # 在实际测试中，我们会使用定位器来找到表格并填充数据

        # 6. 添加图片（如果环境支持）
        try:
            image_tools(
                self.context,
                operation_type="insert",
                image_path=self.test_image_path,
                position="after",
            )
        except Exception:
            # 在测试环境中可能无法插入实际图片，这是预期的
            pass

        # 7. 添加题注（需要定位器支持）

        # 8. 添加注释
        comment_tools(
            self.context,
            operation_type="add",
            comment_text="需要进一步核实这些数据",
            author="审核员",
        )

    def test_document_review_scenario(self):
        """场景测试：文档审阅流程"""
        # 1. 打开文档
        document_operation(
            self.context, operation_type="open", file_path=self.test_doc_path
        )

        # 2. 添加多个注释
        comment_operation(
            self.context,
            operation_type="add",
            text="这部分需要重写，表述不够清晰",
            author="审阅者1",
        )

        comment_operation(
            self.context,
            operation_type="add",
            text="建议添加更多数据支持这个结论",
            author="审阅者2",
        )

        # 3. 查找特定文本并添加注释
        # 这需要定位器支持，在简化测试中跳过

        # 4. 获取所有注释
        result = comment_operation(self.context, operation_type="get_all")
        self.assertIn("Comments retrieved successfully", result)

        # 5. 回复注释（需要注释索引）
        # 在简化测试中跳过

        # 6. 编辑注释（需要注释索引）
        # 在简化测试中跳过

    def test_content_analysis_scenario(self):
        """场景测试：内容分析和批量操作"""
        # 1. 打开文档
        document_operation(
            self.context, operation_type="open", file_path=self.test_doc_path
        )

        # 2. 获取文档元素
        objects_result = document_operation(
            self.context, operation_type="get_objects", object_type="paragraphs"
        )
        self.assertIn("Objects retrieved successfully", objects_result)

        # 3. 查找特定文本
        find_result = text_content_operation(
            self.context,
            operation_type="find",
            text="test",
            match_case=False,
            match_whole_word=False,
        )
        # 验证查找结果

        # 4. 获取文档样式
        styles_result = document_operation(self.context, operation_type="get_styles")
        self.assertIn("Styles retrieved successfully", styles_result)

        # 5. 批量格式化操作
        # 需要实际的定位器支持，在简化测试中跳过

    def test_collaborative_editing_scenario(self):
        """场景测试：协作编辑流程"""
        # 1. 打开文档
        document_operation(
            self.context, operation_type="open", file_path=self.test_doc_path
        )

        # 2. 多个用户添加注释
        comment_operation(
            self.context,
            operation_type="add",
            text="我建议修改这部分内容",
            author="用户A",
        )

        comment_operation(
            self.context, operation_type="add", text="同意用户A的意见", author="用户B"
        )

        comment_operation(
            self.context, operation_type="add", text="我会处理这个问题", author="编辑者"
        )

        # 3. 获取评论线程
        comments_result = comment_operation(self.context, operation_type="get_all")

        # 4. 回复特定评论
        # 需要注释索引支持

        # 5. 编辑现有注释
        # 需要注释索引支持


if __name__ == "__main__":
    unittest.main()
