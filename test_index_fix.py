#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试索引超出范围修复功能
此脚本用于验证object_finder.py中修复的索引处理逻辑，确保当索引超出范围时不会删除所有段落。
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
        
        # 添加测试内容 - 只添加2个段落，用于测试索引超出范围的情况
        print("创建测试文档并添加内容...")
        rng = doc.Range(0, 0)
        rng.Text = "这是第一段文本。"
        rng.InsertParagraphAfter()
        
        rng = doc.Range(rng.End, rng.End)
        rng.Text = "这是第二段文本。"
        rng.InsertParagraphAfter()
        
        # 显示实际创建的段落数量
        print(f"文档中的实际段落数量: {doc.Paragraphs.Count}")
        for i in range(1, doc.Paragraphs.Count + 1):
            print(f"段落{i}文本: '{doc.Paragraphs(i).Range.Text.strip()}'")
        
        # 初始化选择器引擎
        selector = SelectorEngine()
        
        # 测试1: 使用超出范围的索引
        print("\n测试1: 使用超出范围的索引 (value: 10)")
        locator1 = {"type": "paragraph", "value": "10"}
        print(f"  使用定位器: {locator1}")
        try:
            selection1 = selector.select(doc, locator1)
            print(f"  选择的段落数量: {len(selection1._com_ranges)}")
            for i, range_obj in enumerate(selection1._com_ranges):
                print(f"  选择的段落{i+1}文本: '{range_obj.Text.strip()}'")
        except Exception as e:
            print(f"  发生错误: {e}")
            traceback.print_exc()
        
        # 测试2: 使用不存在的文本内容
        print("\n测试2: 使用不存在的文本内容 (value: '不存在的文本')")
        locator2 = {"type": "paragraph", "value": "不存在的文本"}
        print(f"  使用定位器: {locator2}")
        try:
            selection2 = selector.select(doc, locator2)
            print(f"  选择的段落数量: {len(selection2._com_ranges)}")
            for i, range_obj in enumerate(selection2._com_ranges):
                print(f"  选择的段落{i+1}文本: '{range_obj.Text.strip()}'")
        except Exception as e:
            print(f"  发生错误: {e}")
            traceback.print_exc()
        
        # 测试3: 使用有效的索引
        print("\n测试3: 使用有效的索引 (value: 1)")
        locator3 = {"type": "paragraph", "value": "1"}
        print(f"  使用定位器: {locator3}")
        try:
            selection3 = selector.select(doc, locator3)
            print(f"  选择的段落数量: {len(selection3._com_ranges)}")
            for i, range_obj in enumerate(selection3._com_ranges):
                print(f"  选择的段落{i+1}文本: '{range_obj.Text.strip()}'")
        except Exception as e:
            print(f"  发生错误: {e}")
            traceback.print_exc()
        
        # 显示修复后文档中的段落数量，确认没有意外删除
        print(f"\n修复后文档中的段落数量: {doc.Paragraphs.Count}")
        for i in range(1, doc.Paragraphs.Count + 1):
            print(f"段落{i}文本: '{doc.Paragraphs(i).Range.Text.strip()}'")
        
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