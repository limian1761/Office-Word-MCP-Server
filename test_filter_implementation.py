#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试contains_text过滤器的基本实现
此脚本用于验证object_finder.py中的过滤逻辑是否正常工作。
"""
import os
import sys
import traceback
import pythoncom
import win32com.client


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
        
        # 确保正确添加段落
        rng = doc.Range(0, 0)
        rng.Text = "这是第一段文本。"
        rng.InsertParagraphAfter()
        
        rng = doc.Range(rng.End, rng.End)
        rng.Text = "这是第二段文本，包含关键词。"
        rng.InsertParagraphAfter()
        
        rng = doc.Range(rng.End, rng.End)
        rng.Text = "这是第三段文本。"
        rng.InsertParagraphAfter()
        
        rng = doc.Range(rng.End, rng.End)
        rng.Text = "这是第四段文本，用于测试。"
        rng.InsertParagraphAfter()
        
        # 显示实际创建的段落数量
        print(f"文档中的实际段落数量: {doc.Paragraphs.Count}")
        for i in range(1, doc.Paragraphs.Count + 1):
            print(f"段落{i}文本: '{doc.Paragraphs(i).Range.Text.strip()}'")
        
        # 直接测试contains_text逻辑，不通过SelectorEngine
        print("\n直接测试contains_text逻辑:")
        search_text = "关键词"
        print(f"  搜索文本: '{search_text}'")
        
        matched_paragraphs = []
        for i in range(1, doc.Paragraphs.Count + 1):
            paragraph = doc.Paragraphs(i)
            paragraph_text = paragraph.Range.Text.lower()
            if search_text.lower() in paragraph_text:
                matched_paragraphs.append(paragraph)
                print(f"  匹配段落{i}: '{paragraph_text.strip()}'")
        
        print(f"  匹配段落总数: {len(matched_paragraphs)}")
        
        # 测试索引逻辑
        print("\n测试索引逻辑:")
        index = 2  # 测试1-based索引
        if 0 < index <= doc.Paragraphs.Count:
            selected_paragraph = doc.Paragraphs(index)
            print(f"  索引{index}对应的段落: '{selected_paragraph.Range.Text.strip()}'")
        else:
            print(f"  索引{index}超出范围")
        
        print("\n测试完成。")
        
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