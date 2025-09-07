import pythoncom
import win32com.client
from word_docx_tools.selector.selector import SelectorEngine
import sys
import traceback

# 初始化COM
pythoncom.CoInitialize()

# 创建Word应用程序实例
word_app = win32com.client.Dispatch("Word.Application")
word_app.Visible = True

try:
    # 创建一个新文档
    doc = word_app.Documents.Add()
    
    # 添加测试内容
    print("创建测试文档并添加内容...")
    doc.Range(0, 0).Text = "这是第一段文本。\n"
    doc.Range().Collapse(0)  # 0 = wdCollapseEnd
    doc.Range().Text = "这是第二段文本，包含关键词。\n"
    doc.Range().Collapse(0)
    doc.Range().Text = "这是第三段文本。\n"
    
    # 显示实际创建的段落数量
    print(f"文档中的实际段落数量: {doc.Paragraphs.Count}")
    for i in range(1, doc.Paragraphs.Count + 1):
        print(f"段落{i}文本: {doc.Paragraphs(i).Range.Text.strip()}")
    
    # 初始化选择器引擎
    selector = SelectorEngine()
    
    # 测试1: 测试value作为索引
    print("\n测试1: 使用value作为索引")
    locator1 = {"type": "paragraph", "value": "2"}
    print(f"  使用定位器: {locator1}")
    try:
        selection1 = selector.select(doc, locator1)
        print(f"  选择的段落数量: {len(selection1._com_ranges)}")
        for i, range_obj in enumerate(selection1._com_ranges):
            print(f"  选择的段落{i+1}文本: {range_obj.Text.strip()}")
    except Exception as e:
        print(f"  发生错误: {e}")
        traceback.print_exc()
    
    # 测试2: 测试value作为文本内容
    print("\n测试2: 使用value作为文本内容")
    locator2 = {"type": "paragraph", "value": "关键词"}
    print(f"  使用定位器: {locator2}")
    try:
        selection2 = selector.select(doc, locator2)
        print(f"  选择的段落数量: {len(selection2._com_ranges)}")
        for i, range_obj in enumerate(selection2._com_ranges):
            print(f"  选择的段落{i+1}文本: {range_obj.Text.strip()}")
    except Exception as e:
        print(f"  发生错误: {e}")
        traceback.print_exc()
    
    # 测试3: 直接使用index过滤器
    print("\n测试3: 直接使用index过滤器")
    locator3 = {"type": "paragraph", "filters": [{"index": 1}]}
    print(f"  使用定位器: {locator3}")
    try:
        selection3 = selector.select(doc, locator3)
        print(f"  选择的段落数量: {len(selection3._com_ranges)}")
        for i, range_obj in enumerate(selection3._com_ranges):
            print(f"  选择的段落{i+1}文本: {range_obj.Text.strip()}")
    except Exception as e:
        print(f"  发生错误: {e}")
        traceback.print_exc()
    
    # 测试4: 直接使用contains_text过滤器
    print("\n测试4: 直接使用contains_text过滤器")
    locator4 = {"type": "paragraph", "filters": [{"contains_text": "关键词"}]}
    print(f"  使用定位器: {locator4}")
    try:
        selection4 = selector.select(doc, locator4)
        print(f"  选择的段落数量: {len(selection4._com_ranges)}")
        for i, range_obj in enumerate(selection4._com_ranges):
            print(f"  选择的段落{i+1}文本: {range_obj.Text.strip()}")
    except Exception as e:
        print(f"  发生错误: {e}")
        traceback.print_exc()
    
    print("\n测试完成。")
    
finally:
    # 清理
    # 取消下面的注释以自动关闭文档和Word应用程序
    # doc.Close(SaveChanges=0)  # 0 = wdDoNotSaveChanges
    # word_app.Quit()
    pythoncom.CoUninitialize()