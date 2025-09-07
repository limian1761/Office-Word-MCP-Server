#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试增强的value参数处理功能
此脚本测试object_finder.py中select_core方法对value参数的增强处理能力，支持索引、文本内容等多种类型。
"""
import os
import sys
import traceback

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from word_docx_tools.tools.document_tools import DocumentTools
from word_docx_tools.selector.selector import SelectorEngine
from win32com.client import Dispatch, constants


def create_test_document():
    """创建一个包含多种元素的测试文档"""
    try:
        # 创建Word应用程序
        word = Dispatch('Word.Application')
        word.Visible = False
        
        # 创建新文档
        doc = word.Documents.Add()
        
        # 添加测试段落
        doc.Content.InsertAfter('这是第一段文本\n')
        doc.Content.InsertAfter('这是第二段包含特定内容的文本\n')
        doc.Content.InsertAfter('这是第三段用于测试的段落\n')
        doc.Content.InsertAfter('这是一个特殊的段落，用于测试多种过滤器\n')
        doc.Content.InsertAfter('\n')  # 空段落
        doc.Content.InsertAfter('这是最后一段文本\n')
        
        # 保存文档
        test_doc_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_document.docx')
        doc.SaveAs(test_doc_path)
        
        # 关闭文档和Word
        doc.Close()
        word.Quit()
        
        return test_doc_path
    except Exception as e:
        print(f"创建测试文档时出错: {e}")
        traceback.print_exc()
        return None


def test_value_as_index():
    """测试value作为索引的情况"""
    print("\n=== 测试1: value作为索引 ===")
    
    # 创建选择器引擎
    selector = SelectorEngine()
    
    try:
        # 测试场景1: value作为索引，获取第3个段落
        locator = {"type": "paragraph", "value": "3"}
        result = selector.select(doc_path, locator)
        print(f"场景1 - value='3' 结果数量: {len(result)}")
        if result:
            print(f"场景1 - 内容: '{result[0].Range.Text.strip()}'")
            assert result[0].Range.Text.strip() == "这是第三段用于测试的段落", "索引3的段落内容不匹配"
        
        # 测试场景2: value作为索引，使用数字类型
        locator = {"type": "paragraph", "value": 4}
        result = selector.select(doc_path, locator)
        print(f"场景2 - value=4 结果数量: {len(result)}")
        if result:
            print(f"场景2 - 内容: '{result[0].Range.Text.strip()}'")
            assert result[0].Range.Text.strip() == "这是一个特殊的段落，用于测试多种过滤器", "索引4的段落内容不匹配"
            
    except Exception as e:
        print(f"测试1失败: {e}")
        traceback.print_exc()


def test_value_as_text_content():
    """测试value作为文本内容的情况"""
    print("\n=== 测试2: value作为文本内容 ===")
    
    # 创建选择器引擎
    selector = SelectorEngine()
    
    try:
        # 测试场景3: value作为文本内容，不区分大小写
        locator = {"type": "paragraph", "value": "特定内容"}
        result = selector.select(doc_path, locator)
        print(f"场景3 - value='特定内容' 结果数量: {len(result)}")
        for i, item in enumerate(result):
            print(f"场景3 - 结果{i+1}: '{item.Range.Text.strip()}'")
        assert len(result) >= 1, "没有找到包含'特定内容'的段落"
        
        # 测试场景4: value作为文本内容，与现有过滤器结合
        locator = {"type": "paragraph", "value": "段落", "filters": [{"contains_text": "测试"}]}
        result = selector.select(doc_path, locator)
        print(f"场景4 - value='段落' 带过滤器 结果数量: {len(result)}")
        for i, item in enumerate(result):
            print(f"场景4 - 结果{i+1}: '{item.Range.Text.strip()}'")
        assert len(result) >= 1, "没有找到包含'段落'和'测试'的段落"
        
    except Exception as e:
        print(f"测试2失败: {e}")
        traceback.print_exc()


def test_edge_cases():
    """测试一些边缘情况"""
    print("\n=== 测试3: 边缘情况测试 ===")
    
    # 创建选择器引擎
    selector = SelectorEngine()
    
    try:
        # 测试场景5: 无效的索引
        locator = {"type": "paragraph", "value": "100"}
        result = selector.select(doc_path, locator)
        print(f"场景5 - value='100'(无效索引) 结果数量: {len(result)}")
        assert len(result) == 0 or len(result) > 1, "无效索引应该返回所有匹配或空列表"
        
        # 测试场景6: 不存在的文本内容
        locator = {"type": "paragraph", "value": "不存在的文本内容12345"}
        result = selector.select(doc_path, locator)
        print(f"场景6 - 不存在的文本内容 结果数量: {len(result)}")
        assert len(result) == 0, "不存在的文本内容应该返回空列表"
        
        # 测试场景7: value为0的情况
        locator = {"type": "paragraph", "value": "0"}
        result = selector.select(doc_path, locator)
        print(f"场景7 - value='0' 结果数量: {len(result)}")
        
        # 测试场景8: value为空字符串
        locator = {"type": "paragraph", "value": ""}
        result = selector.select(doc_path, locator)
        print(f"场景8 - value为空字符串 结果数量: {len(result)}")
        
    except Exception as e:
        print(f"测试3失败: {e}")
        traceback.print_exc()


if __name__ == "__main__":
    # 创建测试文档
    doc_path = create_test_document()
    
    if not doc_path:
        print("无法创建测试文档，测试中止")
        sys.exit(1)
    
    print(f"成功创建测试文档: {doc_path}")
    
    try:
        # 运行各项测试
        test_value_as_index()
        test_value_as_text_content()
        test_edge_cases()
        
        print("\n所有测试完成！")
    finally:
        # 可选：删除测试文档
        # if os.path.exists(doc_path):
        #     os.remove(doc_path)
        #     print(f"已删除测试文档: {doc_path}")
        pass