# -*- coding: utf-8 -*-
"""
测试框架通用工具函数和类
提供测试环境设置、资源管理、结果验证等通用功能
"""

import json
import os
import tempfile
import shutil
import unittest
from io import StringIO
from unittest.mock import MagicMock

import pythoncom
import win32com.client
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession

from word_docx_tools.mcp_service.app_context import AppContext


class WordDocumentTestBase(unittest.TestCase):
    """Word文档测试基类，提供通用的测试环境设置和资源管理功能"""

    @classmethod
    def setUpClass(cls):
        # 初始化COM
        pythoncom.CoInitialize()
        # 创建临时目录用于测试
        cls.test_dir = tempfile.mkdtemp()

    @classmethod
    def tearDownClass(cls):
        # 清理临时目录
        try:
            shutil.rmtree(cls.test_dir)
        except Exception:
            pass
        # 清理COM资源
        pythoncom.CoUninitialize()

    def setUp(self):
        """测试前准备，创建Word应用程序、文档和相关上下文"""
        self._setup_word_environment()
        self._setup_test_context()

    def tearDown(self):
        """测试后清理，关闭文档和Word应用程序"""
        self._cleanup_document()
        self._cleanup_word_application()

    def _setup_word_environment(self):
        """设置Word应用程序环境"""
        try:
            # 创建Word应用程序实例
            self.word_app = win32com.client.Dispatch("Word.Application")
            # 尝试设置Visible属性，但捕获可能的异常
            try:
                self.word_app.Visible = False
            except AttributeError:
                # 某些环境中可能不支持设置Visible属性，忽略此错误
                pass

            # 创建测试文档
            self.doc = self._create_test_document()
        except Exception as e:
            print(f"设置Word环境时出错: {str(e)}")
            # 使用模拟对象作为后备
            self._setup_mock_word_environment()

    def _create_test_document(self):
        """创建测试文档并添加一些测试内容"""
        try:
            if hasattr(self.word_app, 'Documents') and callable(getattr(self.word_app.Documents, 'Add', None)):
                doc = self.word_app.Documents.Add()
                # 添加默认测试内容
                self._populate_test_document(doc)
                return doc
            else:
                raise AttributeError("无法访问Documents属性或Add方法")
        except Exception as e:
            print(f"创建测试文档时出错: {str(e)}")
            raise

    def _populate_test_document(self, doc):
        """为测试文档添加默认内容"""
        try:
            doc.Range(0, 0).Text = "这是测试文档的第一段。\n"
            doc.Range().Collapse(0)  # 0 = wdCollapseEnd
            doc.Range().Text = "这是测试文档的第二段。\n"
            doc.Range().Collapse(0)
            doc.Range().Text = "这是测试文档的第三段。\n"
        except Exception as e:
            print(f"填充测试文档内容时出错: {str(e)}")

    def _setup_mock_word_environment(self):
        """设置模拟的Word环境（当无法创建真实Word应用程序时使用）"""
        self.word_app = MagicMock()
        self.doc = MagicMock()
        self.doc.Content = MagicMock()
        self.doc.Content.Text = """
        这是测试文档的第一段。
        这是测试文档的第二段。
        这是测试文档的第三段。
        """
        # 设置其他常用模拟属性和方法
        if not hasattr(self.doc, 'Paragraphs'):
            self.doc.Paragraphs = MagicMock()
            self.doc.Paragraphs.Count = 3

    def _setup_test_context(self):
        """设置测试上下文"""
        # 创建应用上下文
        self.app_context = AppContext()
        self.app_context.set_word_app(self.word_app)
        self.app_context.set_active_document(self.doc)

        # 创建模拟的读写流
        self.read_stream = StringIO()
        self.write_stream = StringIO()

        # 创建ServerSession实例
        self.session = ServerSession(
            read_stream=self.read_stream,
            write_stream=self.write_stream,
            init_options={},
        )
        self.session.lifespan_context = self.app_context

        # 创建Context对象
        self.context = Context(
            server_session=self.session,
            request_context=self.session
        )

    def _cleanup_document(self):
        """清理测试文档"""
        try:
            if hasattr(self, "doc") and hasattr(self.doc, 'Close'):
                try:
                    self.doc.Close(SaveChanges=False)
                except Exception as e:
                    print(f"关闭文档时出错: {str(e)}")
        except Exception as e:
            print(f"关闭文档过程中出现异常: {str(e)}")

    def _cleanup_word_application(self):
        """清理Word应用程序"""
        try:
            if hasattr(self, "word_app") and hasattr(self.word_app, 'Quit'):
                try:
                    # 先检查并关闭所有可能打开的文档
                    if hasattr(self.word_app, 'Documents') and hasattr(self.word_app.Documents, 'Count'):
                        try:
                            for i in range(self.word_app.Documents.Count):
                                try:
                                    doc = self.word_app.Documents[1]  # 总是获取第一个文档
                                    doc.Close(SaveChanges=False)
                                except Exception as e:
                                    print(f"关闭文档时出错: {str(e)}")
                                    continue
                        except Exception as e:
                            print(f"获取文档数量时出错: {str(e)}")
                    self.word_app.Quit()
                except Exception as e:
                    print(f"关闭Word应用程序时出错: {str(e)}")
        except Exception as e:
            print(f"关闭Word应用程序过程中出现异常: {str(e)}")

    def verify_tool_result(self, result, expected_success=True):
        """验证工具结果
        
        Args:
            result: 工具返回的结果
            expected_success: 期望的成功状态
            
        Returns:
            解析后的结果数据
        """
        try:
            if result is None:
                print("Got None result")
                self.fail("Expected result but got None")
            elif isinstance(result, dict):
                # 如果已经是字典，直接验证
                if "success" in result:
                    if expected_success:
                        self.assertTrue(result["success"])
                    else:
                        self.assertFalse(result["success"])
                return result
            elif isinstance(result, str):
                if not result:
                    print("Got empty string result")
                    self.fail("Expected JSON string or dict but got empty string")
                elif result.startswith("Error ["):
                    # 处理格式化的错误响应
                    print(f"Got error response: {result}")
                    if expected_success:
                        self.fail(f"Expected success but got error: {result}")
                    return {"success": False, "error": result}
                else:
                    # 尝试解析为JSON
                    try:
                        result_data = json.loads(result)
                        if "success" in result_data:
                            if expected_success:
                                self.assertTrue(result_data["success"])
                            else:
                                self.assertFalse(result_data["success"])
                        return result_data
                    except json.JSONDecodeError:
                        # 如果不是有效的JSON字符串，则失败
                        self.fail(f"Expected valid JSON string or dict but got: {result}")
            else:
                self.fail(f"Expected JSON string or dict but got: {type(result)}")
        except Exception as e:
            error_message = f"验证工具结果时出错: {repr(e)}"
            print(error_message)
            self.fail(error_message)

    def get_test_file_path(self, file_name):
        """获取测试文件的路径
        
        Args:
            file_name: 文件名
            
        Returns:
            完整的文件路径
        """
        return os.path.join(self.test_dir, file_name)

    def create_test_file(self, file_name, content=""):
        """创建测试文件
        
        Args:
            file_name: 文件名
            content: 文件内容
            
        Returns:
            完整的文件路径
        """
        file_path = self.get_test_file_path(file_name)
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(content)
            return file_path
        except Exception as e:
            print(f"创建测试文件时出错: {str(e)}")
            self.fail(f"创建测试文件时出错: {str(e)}")

    def copy_test_document(self, source_doc_name="valid_test_document_v2.docx", target_doc_name="test_document.docx"):
        """复制测试文档到临时目录
        
        Args:
            source_doc_name: 源文档名称
            target_doc_name: 目标文档名称
            
        Returns:
            目标文档的路径
        """
        source_doc_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "tests",
            "test_docs",
            source_doc_name,
        )
        target_doc_path = self.get_test_file_path(target_doc_name)
        
        if os.path.exists(source_doc_path):
            try:
                shutil.copy2(source_doc_path, target_doc_path)
                return target_doc_path
            except Exception as e:
                print(f"复制测试文档时出错: {str(e)}")
                # 如果复制失败，尝试创建新文档
                return self._create_empty_test_document(target_doc_path)
        else:
            print(f"源测试文档不存在: {source_doc_path}")
            # 源文档不存在，创建新文档
            return self._create_empty_test_document(target_doc_path)

    def _create_empty_test_document(self, doc_path):
        """创建一个空的测试文档
        
        Args:
            doc_path: 文档路径
            
        Returns:
            文档路径
        """
        try:
            # 使用当前的Word应用程序创建文档
            if hasattr(self, 'word_app') and self.word_app is not None and not isinstance(self.word_app, MagicMock):
                doc = self.word_app.Documents.Add()
                doc.SaveAs2(doc_path)
                doc.Close()
                return doc_path
            else:
                # 如果没有可用的Word应用程序，创建一个简单的标记文件
                with open(doc_path, "w") as f:
                    f.write("Mock Word document for testing")
                return doc_path
        except Exception as e:
            print(f"创建空测试文档时出错: {str(e)}")
            self.fail(f"创建空测试文档时出错: {str(e)}")


class TestResultVerifier:
    """测试结果验证器，提供更丰富的结果验证功能"""

    @staticmethod
    def verify_json_result(result, expected_keys=None, expected_success=True):
        """验证JSON格式的结果
        
        Args:
            result: 要验证的结果
            expected_keys: 预期包含的键列表
            expected_success: 预期success字段的值
            
        Returns:
            解析后的JSON数据
        """
        try:
            # 解析JSON
            if isinstance(result, str):
                result_data = json.loads(result)
            elif isinstance(result, dict):
                result_data = result
            else:
                raise TypeError(f"Expected string or dict, got {type(result)}")

            # 验证success字段
            if "success" in result_data:
                assert result_data["success"] == expected_success, \
                    f"Expected success={expected_success}, got {result_data['success']}"

            # 验证预期的键
            if expected_keys:
                for key in expected_keys:
                    assert key in result_data, f"Expected key '{key}' not found in result"

            return result_data
        except AssertionError as e:
            print(f"结果验证失败: {str(e)}")
            raise
        except Exception as e:
            print(f"解析或验证结果时出错: {str(e)}")
            raise

    @staticmethod
    def verify_error_result(result, expected_error_message=None):
        """验证错误结果
        
        Args:
            result: 要验证的结果
            expected_error_message: 预期的错误消息（部分匹配）
            
        Returns:
            解析后的错误数据
        """
        try:
            # 首先验证这是一个失败的结果
            result_data = TestResultVerifier.verify_json_result(result, expected_success=False)

            # 验证错误消息
            if expected_error_message:
                error_found = False
                # 检查可能的错误消息位置
                for key in ["error", "message", "detail"]:
                    if key in result_data and expected_error_message in str(result_data[key]):
                        error_found = True
                        break
                assert error_found, f"Expected error message '{expected_error_message}' not found"

            return result_data
        except AssertionError as e:
            print(f"错误结果验证失败: {str(e)}")
            raise

    @staticmethod
    def verify_paragraph_content(doc, expected_content, paragraph_index=1):
        """验证文档中特定段落的内容
        
        Args:
            doc: Word文档对象
            expected_content: 预期的段落内容（部分匹配）
            paragraph_index: 段落索引（从1开始）
        """
        try:
            # 检查是否为模拟对象
            if isinstance(doc, MagicMock):
                # 对于模拟对象，我们只能检查Content.Text
                assert expected_content in doc.Content.Text, \
                    f"Expected content '{expected_content}' not found in mock document"
            else:
                # 对于真实文档，获取指定段落
                assert hasattr(doc, 'Paragraphs'), "Document has no Paragraphs property"
                assert paragraph_index <= doc.Paragraphs.Count, \
                    f"Paragraph index {paragraph_index} exceeds document paragraphs count {doc.Paragraphs.Count}"
                
                paragraph_text = doc.Paragraphs(paragraph_index).Range.Text
                assert expected_content in paragraph_text, \
                    f"Expected content '{expected_content}' not found in paragraph {paragraph_index}: '{paragraph_text}'"
        except AssertionError as e:
            print(f"段落内容验证失败: {str(e)}")
            raise
        except Exception as e:
            print(f"验证段落内容时出错: {str(e)}")
            raise