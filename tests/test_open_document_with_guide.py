import os
import sys
import json
import asyncio
from mcp.server.fastmcp.server import Context

# 创建一个模拟的上下文对象
class MockContext:
    def __init__(self):
        self.session = type('obj', (object,), {})
        self.session.document_state = {}
        self.session.backend_instances = {}
        self.session.document_cache = {}

async def call_tool(tool_name, params):
    # 模拟调用工具的函数
    ctx = MockContext()
    if tool_name == 'open_document':
        from word_document_server.tools.document import open_document
        return open_document(ctx, **params)
    elif tool_name == 'shutdown_word':
        from word_document_server.tools.document import shutdown_word
        return shutdown_word(ctx)
    else:
        return f"工具 {tool_name} 不支持"


async def test_open_document_with_guide():
    # 创建测试文档路径
    test_doc_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tests', 'test_docs', 'test_document.docx')
    print(f"测试文档路径: {test_doc_path}")

    try:
        # 调用open_document工具
        print("调用open_document工具...")
        response = await call_tool(
            "open_document",
            {
                "file_path": test_doc_path
            }
        )
        print("调用完成")

        # 检查响应内容
        print("\n--- 响应结果摘要 ---")
        if "Document opened successfully" in response and "Office-Word-MCP-Server Agent Guide" in response:
            print("✅ 测试成功: open_document工具同时返回了文档信息和agent_guide内容")
        else:
            print("❌ 测试失败: 响应中未同时包含文档信息和agent_guide内容")
            print(f"响应内容预览: {response[:500]}...")

        # 关闭Word应用
        print("\n调用shutdown_word工具关闭Word应用...")
        await call_tool("shutdown_word", {})
        print("Word应用已关闭")

    except Exception as e:
        print(f"❌ 测试出错: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    asyncio.run(test_open_document_with_guide())