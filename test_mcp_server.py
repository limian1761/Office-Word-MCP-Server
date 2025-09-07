import asyncio

import mcp


async def main():
    """测试MCP服务器连接"""
    try:
        # 尝试连接到本地MCP服务器
        print("尝试连接到Word MCP服务器...")
        client = mcp.Client("word-docx-tools")

        # 获取服务器信息
        info = await client.get_server_info()
        print(f"成功连接到服务器！服务器信息: {info}")

        # 列出可用的工具
        tools = await client.list_tools()
        print(f"可用工具数量: {len(tools)}")
        for i, tool in enumerate(tools[:5]):  # 只显示前5个工具
            print(f"{i+1}. {tool.name}")

        await client.close()
        return True
    except Exception as e:
        print(f"连接服务器失败: {e}")
        return False


if __name__ == "__main__":
    result = asyncio.run(main())
    print(f"测试结果: {'成功' if result else '失败'}")
