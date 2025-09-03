import asyncio
import logging
import sys

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

async def verify_comment_tools():
    """验证comment_tools的修复效果"""
    try:
        logger.info("开始验证comment_tools的修复效果...")
        
        # 导入必要的模块
        from word_docx_tools.tools.comment_tools import comment_tools
        from mcp.server.session import ServerSession
        from word_docx_tools.mcp_service.app_context import AppContext
        
        # 创建模拟的上下文对象
        mock_context = type('MockContext', (), {})
        mock_request_context = type('MockRequestContext', (), {})
        mock_lifespan_context = type('MockLifespanContext', (), {})
        
        # 尝试获取Word应用程序和文档
        try:
            import win32com.client
            
            # 创建Word应用程序实例
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = True  # 使Word可见以便观察
            logger.info("成功创建Word应用程序实例")
            
            # 创建一个新文档
            document = word_app.Documents.Add()
            logger.info("成功创建新文档")
            
            # 设置上下文的文档引用
            mock_lifespan_context.get_active_document = lambda: document
            mock_request_context.lifespan_context = mock_lifespan_context
            mock_context.request_context = mock_request_context
            
            try:
                # 测试添加评论
                logger.info("开始添加评论...")
                add_result = await comment_tools(
                    ctx=mock_context,
                    operation_type="add",
                    comment_text="这是一条测试评论 - 验证修复效果",
                    author="修复验证程序"
                )
                logger.info(f"添加评论结果: {add_result}")
                if add_result["success"]:
                    logger.info(f"评论添加成功，ID: {add_result['comment_id']}")
                else:
                    logger.error("评论添加失败")
                    return False
                
                # 保存文档以确保评论被保存
                doc_path = "c:/Users/lichao/Office-Word-MCP-Server/verify_comment_doc.docx"
                document.SaveAs2(doc_path)
                logger.info(f"文档已保存到: {doc_path}")
                
                # 测试获取所有评论
                logger.info("开始获取所有评论...")
                get_result = await comment_tools(
                    ctx=mock_context,
                    operation_type="get_all"
                )
                logger.info(f"获取评论结果: {get_result}")
                
                if get_result["success"]:
                    logger.info(f"成功获取评论，共找到 {len(get_result['comments'])} 条评论")
                    
                    # 显示每条评论的详细信息
                    for i, comment in enumerate(get_result['comments']):
                        logger.info(f"评论 {i+1}:")
                        logger.info(f"  索引: {comment.get('index', 'N/A')}")
                        logger.info(f"  文本: {comment.get('text', 'N/A')}")
                        logger.info(f"  作者: {comment.get('author', 'N/A')}")
                        logger.info(f"  日期: {comment.get('date', 'N/A')}")
                        logger.info(f"  回复数: {comment.get('replies_count', 0)}")
                        
                    # 验证是否获取到了我们添加的评论
                    if len(get_result['comments']) > 0:
                        # 检查第一条评论是否包含我们添加的文本
                        first_comment_text = get_result['comments'][0].get('text', '')
                        if "这是一条测试评论" in first_comment_text:
                            logger.info("✓ 验证成功！成功获取到了添加的评论内容。")
                            return True
                        else:
                            logger.warning(f"✗ 验证失败：获取到的评论内容不匹配。获取到的文本: {first_comment_text}")
                            return False
                    else:
                        logger.error("✗ 验证失败：未能获取到任何评论。")
                        return False
                else:
                    logger.error(f"✗ 获取评论失败: {get_result.get('message', '未知错误')}")
                    return False
                
            finally:
                # 清理资源
                try:
                    document.Close(SaveChanges=False)
                    logger.info("文档已关闭")
                except Exception as e:
                    logger.warning(f"关闭文档时出错: {e}")
                
                try:
                    word_app.Quit()
                    logger.info("Word应用程序已退出")
                except Exception as e:
                    logger.warning(f"退出Word时出错: {e}")
                    
        except ImportError:
            logger.error("未找到win32com模块，请安装pywin32库")
            return False
        except Exception as e:
            logger.error(f"验证过程中发生错误: {str(e)}")
            return False
            
    except Exception as e:
        logger.error(f"验证函数执行失败: {str(e)}")
        return False

if __name__ == "__main__":
    # 运行验证函数
    success = asyncio.get_event_loop().run_until_complete(verify_comment_tools())
    
    # 根据结果设置退出码
    sys.exit(0 if success else 1)