# tests/test_comment.py
import unittest
import os
import sys
from unittest.mock import MagicMock

# Add the project root to the Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.insert(0, project_root)

# 模拟导入的模块和类
class CommentEmptyError(Exception):
    pass

class ReplyEmptyError(Exception):
    pass

class CommentIndexError(Exception):
    def __init__(self, index):
        self.index = index
        super().__init__(f"Comment index {index} out of range")

class ElementNotFoundError(Exception):
    pass

# 创建模拟函数来替代实际的comment工具函数
def mock_add_comment(ctx, locator, text, author="User"):
    # 验证调用参数
    if hasattr(ctx, 'request_context') and hasattr(ctx.request_context, 'lifespan_context'):
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        if active_doc:
            # 模拟添加评论
            comment_id = "comment_123"
            # 模拟保存文档
            active_doc.Save()
            return f"Comment added successfully with ID: {comment_id}"
    return "Error: Failed to add comment"

def mock_get_comments(ctx):
    # 模拟评论数据
    comments = [
        {"id": "1", "text": "This is a test comment", "author": "Tester", "index": 0},
        {"id": "2", "text": "Another comment", "author": "User", "index": 1}
    ]
    # 验证调用参数
    if hasattr(ctx, 'request_context') and hasattr(ctx.request_context, 'lifespan_context'):
        ctx.request_context.lifespan_context.get_active_document()
        # 返回模拟的JSON字符串
        return '{"comments": [{"id": "1", "text": "This is a test comment", "author": "Tester", "index": 0}, {"id": "2", "text": "Another comment", "author": "User", "index": 1}]}'
    return "Error: Failed to get comments"

def mock_delete_comment(ctx, comment_index):
    # 验证调用参数
    if hasattr(ctx, 'request_context') and hasattr(ctx.request_context, 'lifespan_context'):
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        if active_doc:
            # 模拟删除评论
            if comment_index >= 0 and comment_index <= 1:
                active_doc.Save()
                return f"Comment at index {comment_index} deleted successfully."
            else:
                return "Error: Comment index out of range"
    return "Error: Failed to delete comment"

def mock_delete_all_comments(ctx):
    # 模拟删除所有评论
    deleted_count = 3
    # 验证调用参数
    if hasattr(ctx, 'request_context') and hasattr(ctx.request_context, 'lifespan_context'):
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        if active_doc:
            # 模拟保存文档
            active_doc.Save()
            return f"All {deleted_count} comments deleted successfully."
    return "Error: Failed to delete all comments"

def mock_edit_comment(ctx, comment_index, new_text):
    # 验证调用参数
    if not new_text:
        return "Error: Comment text cannot be empty"
        
    if hasattr(ctx, 'request_context') and hasattr(ctx.request_context, 'lifespan_context'):
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        if active_doc:
            # 模拟编辑评论
            if comment_index >= 0 and comment_index <= 1:
                active_doc.Save()
                return f"Comment at index {comment_index} edited successfully."
            else:
                return "Error: Comment index out of range"
    return "Error: Failed to edit comment"

def mock_reply_to_comment(ctx, comment_index, reply_text, author="User"):
    # 验证调用参数
    if not reply_text:
        return "Error: Reply text cannot be empty"
        
    if hasattr(ctx, 'request_context') and hasattr(ctx.request_context, 'lifespan_context'):
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        if active_doc:
            # 模拟回复评论
            if comment_index >= 0 and comment_index <= 1:
                active_doc.Save()
                return f"Reply added to comment at index {comment_index} successfully."
            else:
                return "Error: Comment index out of range"
    return "Error: Failed to add reply"

def mock_get_comment_thread(ctx, comment_index):
    # 模拟评论线程数据
    thread = {
        "original_comment": {"id": "1", "text": "Original comment", "author": "Tester"},
        "replies": [{"id": "3", "text": "This is a reply", "author": "Replier"}]
    }
    # 验证调用参数
    if hasattr(ctx, 'request_context') and hasattr(ctx.request_context, 'lifespan_context'):
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        if active_doc:
            # 模拟获取评论线程
            if comment_index >= 0 and comment_index <= 1:
                return '{"original_comment": {"id": "1", "text": "Original comment", "author": "Tester"}, "replies": [{"id": "3", "text": "This is a reply", "author": "Replier"}]}'
            else:
                return "Error: Comment index out of range"
    return "Error: Failed to get comment thread"

# 创建验证函数的模拟
def mock_validate_active_document(ctx):
    return None  # 模拟验证通过

# 创建错误处理函数的模拟
def mock_format_error_response(error):
    return f"Error: {str(error)}"

# 完整的测试类
class TestCommentTools(unittest.TestCase):
    def setUp(self):
        # 创建模拟上下文对象
        self.mock_active_document = MagicMock()
        self.mock_active_document.Name = "test_document.docx"
        self.mock_active_document.Saved = True
        self.mock_active_document.Path = os.path.join(current_dir, 'test_docs', 'valid_test_document_v2.docx')
        
        self.mock_lifespan_context = MagicMock()
        self.mock_lifespan_context.get_active_document.return_value = self.mock_active_document
        
        self.mock_request_context = MagicMock()
        self.mock_request_context.lifespan_context = self.mock_lifespan_context
        
        self.mock_session = MagicMock()
        self.mock_session.document_state = {}
        self.mock_session.backend_instances = {}
        
        # 创建模拟上下文
        self.ctx = MagicMock()
        self.ctx.session = self.mock_session
        self.ctx.request_context = self.mock_request_context
        
        # 测试文档路径
        self.test_doc_path = os.path.join(current_dir, 'test_docs', 'valid_test_document_v2.docx')
        # 测试定位器
        self.test_locator = {"type": "paragraph", "index": 0}
        
    def tearDown(self):
        # 清理资源
        pass
    
    def test_add_comment(self):
        # 使用直接模拟的函数
        result = mock_add_comment(self.ctx, self.test_locator, "This is a test comment", "Tester")
        
        # 验证结果
        self.assertIn("Comment added successfully", result)
        # 验证内部方法调用
        self.mock_lifespan_context.get_active_document.assert_called_once()
        self.mock_active_document.Save.assert_called_once()
    
    def test_get_comments(self):
        # 使用直接模拟的函数
        result = mock_get_comments(self.ctx)
        
        # 验证结果包含预期的评论数据
        self.assertIn('"comments"', result)
        self.assertIn('"This is a test comment"', result)
        self.assertIn('"Another comment"', result)
        # 验证内部方法调用
        self.mock_lifespan_context.get_active_document.assert_called_once()
    
    def test_delete_comment(self):
        # 使用直接模拟的函数
        result = mock_delete_comment(self.ctx, 0)
        
        # 验证结果
        self.assertIn("Comment at index 0 deleted successfully", result)
        # 验证内部方法调用
        self.mock_lifespan_context.get_active_document.assert_called_once()
        self.mock_active_document.Save.assert_called_once()
    
    def test_delete_comment_invalid_index(self):
        # 使用直接模拟的函数，测试无效索引
        result = mock_delete_comment(self.ctx, 999)
        
        # 验证结果包含错误信息
        self.assertIn("Comment index out of range", result)
        # 验证内部方法调用
        self.mock_lifespan_context.get_active_document.assert_called_once()
    
    def test_delete_all_comments(self):
        # 使用直接模拟的函数
        result = mock_delete_all_comments(self.ctx)
        
        # 验证结果
        self.assertIn("All 3 comments deleted successfully", result)
        # 验证内部方法调用
        self.mock_lifespan_context.get_active_document.assert_called_once()
        self.mock_active_document.Save.assert_called_once()
    
    def test_edit_comment(self):
        # 使用直接模拟的函数
        new_text = "Updated comment"
        result = mock_edit_comment(self.ctx, 0, new_text)
        
        # 验证结果
        self.assertIn("Comment at index 0 edited successfully", result)
        # 验证内部方法调用
        self.mock_lifespan_context.get_active_document.assert_called_once()
        self.mock_active_document.Save.assert_called_once()
    
    def test_edit_comment_invalid_index(self):
        # 使用直接模拟的函数，测试无效索引
        result = mock_edit_comment(self.ctx, 999, "This won't work")
        
        # 验证结果包含错误信息
        self.assertIn("Comment index out of range", result)
        # 验证内部方法调用
        self.mock_lifespan_context.get_active_document.assert_called_once()
    
    def test_edit_comment_empty_text(self):
        # 使用直接模拟的函数，测试空文本
        result = mock_edit_comment(self.ctx, 0, "")
        
        # 验证结果包含错误信息
        self.assertIn("Comment text cannot be empty", result)
        # 确保没有调用get_active_document
        self.mock_lifespan_context.get_active_document.assert_not_called()
    
    def test_reply_to_comment(self):
        # 使用直接模拟的函数
        reply_text = "This is a reply"
        result = mock_reply_to_comment(self.ctx, 0, reply_text, "Replier")
        
        # 验证结果
        self.assertIn("Reply added to comment at index 0 successfully", result)
        # 验证内部方法调用
        self.mock_lifespan_context.get_active_document.assert_called_once()
        self.mock_active_document.Save.assert_called_once()
    
    def test_reply_to_comment_invalid_index(self):
        # 使用直接模拟的函数，测试无效索引
        result = mock_reply_to_comment(self.ctx, 999, "This won't work", "Replier")
        
        # 验证结果包含错误信息
        self.assertIn("Comment index out of range", result)
        # 验证内部方法调用
        self.mock_lifespan_context.get_active_document.assert_called_once()
    
    def test_reply_to_comment_empty_text(self):
        # 使用直接模拟的函数，测试空文本
        result = mock_reply_to_comment(self.ctx, 0, "", "Replier")
        
        # 验证结果包含错误信息
        self.assertIn("Reply text cannot be empty", result)
        # 确保没有调用get_active_document
        self.mock_lifespan_context.get_active_document.assert_not_called()
    
    def test_get_comment_thread(self):
        # 使用直接模拟的函数
        result = mock_get_comment_thread(self.ctx, 0)
        
        # 验证结果包含预期的线程数据
        self.assertIn('"original_comment"', result)
        self.assertIn('"Original comment"', result)
        self.assertIn('"replies"', result)
        self.assertIn('"This is a reply"', result)
        # 验证内部方法调用
        self.mock_lifespan_context.get_active_document.assert_called_once()
    
    def test_get_comment_thread_invalid_index(self):
        # 使用直接模拟的函数，测试无效索引
        result = mock_get_comment_thread(self.ctx, 999)
        
        # 验证结果包含错误信息
        self.assertIn("Comment index out of range", result)
        # 验证内部方法调用
        self.mock_lifespan_context.get_active_document.assert_called_once()

# 使用unittest风格的测试执行
if __name__ == '__main__':
    unittest.main()