#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试多种参数名（value、index、id等）处理功能
此脚本测试object_finder.py中select_core方法对不同参数名的处理能力。
"""
import os
import sys
import traceback
import pythoncom
import win32com.client
from word_docx_tools.selector.selector import SelectorEngine


if __name__ == "__main__":
    # 初始化COM
    pythoncom.CoInitialize()

    try:
        # 创建Word应用程序实例
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = True

        # 创建一个新文档
        doc = word_app.Documents.Add()
        
        # 添加测试内容
        print("创建测试文档并添加内容...")
        doc.Range(0, 0).Text = "这是第一段文本。\n"
        doc.Range().Collapse(0)  # 0 = wdCollapseEnd
        doc.Range().Text = "这是第二段文本，包含关键词。\n"
        doc.Range().Collapse(0)
        doc.Range().Text = "这是第三段文本。\n"
        doc.Range().Collapse(0)
        doc.Range().Text = "这是第四段文本，用于测试。\n"
        
        # 显示实际创建的段落数量
        print(f"文档中的实际段落数量: {doc.Paragraphs.Count}")
        for i in range(1, doc.Paragraphs.Count + 1):
            print(f"段落{i}文本: {doc.Paragraphs(i).Range.Text.strip()}")
        
        # 初始化选择器引擎
        selector = SelectorEngine()
        
        # 测试1: 使用value作为索引
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
        
        # 测试2: 使用index作为索引
        print("\n测试2: 使用index作为索引")
        locator2 = {"type": "paragraph", "index": "3"}
        print(f"  使用定位器: {locator2}")
        try:
            selection2 = selector.select(doc, locator2)
            print(f"  选择的段落数量: {len(selection2._com_ranges)}")
            for i, range_obj in enumerate(selection2._com_ranges):
                print(f"  选择的段落{i+1}文本: {range_obj.Text.strip()}")
        except Exception as e:
            print(f"  发生错误: {e}")
            traceback.print_exc()
        
        # 测试3: 使用id作为索引
        print("\n测试3: 使用id作为索引")
        locator3 = {"type": "paragraph", "id": "1"}
        print(f"  使用定位器: {locator3}")
        try:
            selection3 = selector.select(doc, locator3)
            print(f"  选择的段落数量: {len(selection3._com_ranges)}")
            for i, range_obj in enumerate(selection3._com_ranges):
                print(f"  选择的段落{i+1}文本: {range_obj.Text.strip()}")
        except Exception as e:
            print(f"  发生错误: {e}")
            traceback.print_exc()
        
        # 测试4: 使用value作为文本内容
        print("\n测试4: 使用value作为文本内容")
        locator4 = {"type": "paragraph", "value": "关键词"}
        print(f"  使用定位器: {locator4}")
        try:
            selection4 = selector.select(doc, locator4)
            print(f"  选择的段落数量: {len(selection4._com_ranges)}")
            for i, range_obj in enumerate(selection4._com_ranges):
                print(f"  选择的段落{i+1}文本: {range_obj.Text.strip()}")
        except Exception as e:
            print(f"  发生错误: {e}")
            traceback.print_exc()
        
        # 测试5: 使用index作为文本内容
        print("\n测试5: 使用index作为文本内容")
        locator5 = {"type": "paragraph", "index": "测试"}
        print(f"  使用定位器: {locator5}")
        try:
            selection5 = selector.select(doc, locator5)
            print(f"  选择的段落数量: {len(selection5._com_ranges)}")
            for i, range_obj in enumerate(selection5._com_ranges):
                print(f"  选择的段落{i+1}文本: {range_obj.Text.strip()}")
        except Exception as e:
            print(f"  发生错误: {e}")
            traceback.print_exc()
        
        # 测试6: 使用id作为文本内容
        print("\n测试6: 使用id作为文本内容")
        locator6 = {"type": "paragraph", "id": "第四段"}
        print(f"  使用定位器: {locator6}")
        try:
            selection6 = selector.select(doc, locator6)
            print(f"  选择的段落数量: {len(selection6._com_ranges)}")
            for i, range_obj in enumerate(selection6._com_ranges):
                print(f"  选择的段落{i+1}文本: {range_obj.Text.strip()}")
        except Exception as e:
            print(f"  发生错误: {e}")
            traceback.print_exc()
        
        # 测试7: 多个参数同时存在（优先级测试）
        print("\n测试7: 多个参数同时存在（优先级测试）")
        locator7 = {"type": "paragraph", "value": "5", "index": "关键词", "id": "不存在的文本"}
        print(f"  使用定位器: {locator7}")
        try:
            selection7 = selector.select(doc, locator7)
            print(f"  选择的段落数量: {len(selection7._com_ranges)}")
            for i, range_obj in enumerate(selection7._com_ranges):
                print(f"  选择的段落{i+1}文本: {range_obj.Text.strip()}")
        except Exception as e:
            print(f"  发生错误: {e}")
            traceback.print_exc()
        
        print("\n所有测试完成。")
        
    except Exception as e:
        print(f"测试过程中发生错误: {e}")
        traceback.print_exc()
    finally:
        # 清理
        input("按Enter键关闭Word文档...")
        # 取消下面的注释以自动关闭文档和Word应用程序
        # if 'doc' in locals():
        #     doc.Close(SaveChanges=0)  # 0 = wdDoNotSaveChanges
        # if 'word_app' in locals():
        #     word_app.Quit()
        pythoncom.CoUninitialize()