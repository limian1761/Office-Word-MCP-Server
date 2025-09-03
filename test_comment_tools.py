import unittest
import asyncio
import logging
from unittest.mock import MagicMock, patch

# 设置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class TestCommentTools(unittest.TestCase):
    def setUp(self):
        # 创建模拟对象
        self.mock_document = MagicMock()
        self.mock_range = MagicMock()
        self.mock_comment = MagicMock()
        
        # 模拟文档内容
        self.mock_document.Content = self.mock_range
        self.mock_document.Comments = MagicMock()
        
        # 初始化评论计数为0
        self.mock_document.Comments.Count = 0
    
    async def test_add_and_get_comment(self):
        """测试添加评论后能否正确获取"""
        try:
            # 导入comment_tools
            from word_docx_tools.tools.comment_tools import comment_tools
            from word_docx_tools.mcp_service.app_context import AppContext
            from mcp.server.session import ServerSession
            
            # 创建模拟的上下文对象
            mock_context = MagicMock()
            mock_context.request_context.lifespan_context.get_active_document.return_value = self.mock_document
            
            # 模拟添加评论后评论计数变为1
            def side_effect_add(*args, **kwargs):
                self.mock_document.Comments.Count = 1
                return self.mock_comment
            
            # 模拟get_comments返回包含一条评论的列表
            def side_effect_get(*args, **kwargs):
                return [{"index": 0, "text": "Test comment", "author": "Test Author"}]
            
            # 应用模拟效果
            with patch('word_docx_tools.operations.comment_ops.add_comment', side_effect=side_effect_add):
                with patch('word_docx_tools.operations.comment_ops.get_comments', side_effect=side_effect_get):
                    # 测试添加评论
                    add_result = await comment_tools(
                        ctx=mock_context,
                        operation_type="add",
                        comment_text="Test comment",
                        author="Test Author"
                    )
                    logger.info(f"Add comment result: {add_result}")
                    self.assertTrue(add_result["success"])
                    
                    # 测试获取所有评论
                    get_result = await comment_tools(
                        ctx=mock_context,
                        operation_type="get_all"
                    )
                    logger.info(f"Get all comments result: {get_result}")
                    self.assertTrue(get_result["success"])
                    self.assertEqual(len(get_result["comments"]), 1)
                    self.assertEqual(get_result["comments"][0]["text"], "Test comment")
        except Exception as e:
            logger.error(f"Test failed with error: {str(e)}")
            raise
    
    async def test_real_com_integration(self):
        """尝试实际的COM集成测试"""
        try:
            import win32com.client
            
            # 导入comment_tools
            from word_docx_tools.tools.comment_tools import comment_tools
            from word_docx_tools.mcp_service.app_context import AppContext
            
            # 创建一个实际的Word实例
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = True  # 使Word可见以便观察
            
            # 创建一个新文档
            real_document = word_app.Documents.Add()
            
            # 创建模拟的上下文对象
            mock_context = MagicMock()
            mock_context.request_context.lifespan_context.get_active_document.return_value = real_document
            
            # 测试添加评论
            add_result = await comment_tools(
                ctx=mock_context,
                operation_type="add",
                comment_text="Integration test comment",
                author="Test Integration"
            )
            logger.info(f"Real add comment result: {add_result}")
            self.assertTrue(add_result["success"])
            
            # 保存文档以确保评论被保存
            doc_path = "c:/Users/lichao/Office-Word-MCP-Server/test_comment_doc.docx"
            real_document.SaveAs2(doc_path)
            logger.info(f"Document saved to: {doc_path}")
            
            # 测试获取所有评论
            get_result = await comment_tools(
                ctx=mock_context,
                operation_type="get_all"
            )
            logger.info(f"Real get all comments result: {get_result}")
            self.assertTrue(get_result["success"])
            logger.info(f"Number of comments retrieved: {len(get_result['comments'])}")
            
            # 显示获取到的评论内容
            if get_result["comments"]:
                for comment in get_result["comments"]:
                    logger.info(f"Comment: {comment}")
            
            # 关闭文档但不保存更改
            real_document.Close(SaveChanges=False)
            word_app.Quit()
            
        except Exception as e:
            logger.error(f"Real COM integration test failed with error: {str(e)}")
            # 确保清理资源
            try:
                real_document.Close(SaveChanges=False)
                word_app.Quit()
            except:
                pass
            raise

    def run_async_test(self, coro):
        """运行异步测试的辅助方法"""
        return asyncio.get_event_loop().run_until_complete(coro)
    
    def test_add_and_get_comment_sync(self):
        """同步运行添加和获取评论的测试"""
        self.run_async_test(self.test_add_and_get_comment())
    
    def test_real_com_integration_sync(self):
        """同步运行实际COM集成测试"""
        self.run_async_test(self.test_real_com_integration())

if __name__ == "__main__":
    unittest.main()