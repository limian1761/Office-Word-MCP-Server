import os
import sys

# 打印Python版本和路径
print(f"Python版本: {sys.version}")
print(f"Python路径: {sys.executable}")
print(f"当前工作目录: {os.getcwd()}")

# 检查是否在虚拟环境中
if hasattr(sys, "base_prefix") and sys.base_prefix != sys.prefix:
    print("✓ 已在虚拟环境中运行")
    print(f"虚拟环境路径: {sys.prefix}")
else:
    print("✗ 不在虚拟环境中运行")

# 检查已安装的包
try:
    import mcp

    print("✓ 成功导入mcp包")
    # 尝试不同的pywin32导入方式
    try:
        import win32com.client

        print("✓ 成功导入win32com.client")
    except ImportError:
        print("✗ 无法导入win32com.client")
except ImportError as e:
    print(f"✗ 导入包失败: {e}")

print("Python环境测试完成!")
