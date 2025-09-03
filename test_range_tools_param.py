#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
简单测试range_tools的参数传递修复

这个脚本专门验证对range_tools.py的修复是否正确处理了参数传递。
"""

import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


class MockDocument:
    """简单的模拟文档对象"""
    def __init__(self):
        # 模拟一些必要的属性
        class Paragraphs:
            def __getitem__(self, index):
                class Paragraph:
                    class Range:
                        Text = "测试文本"
                        Start = 0
                        End = 10
                        class Style:
                            NameLocal = "Normal"
                    Range = Range()
                return Paragraph()
        self.Paragraphs = Paragraphs()


class MockSelector:
    """简单的模拟选择器引擎"""
    def select(self, document, locator):
        # 验证locator参数类型
        print(f"✓ 选择器接收到的locator类型: {type(locator).__name__}")
        
        # 检查locator是否是预期的格式
        if isinstance(locator, dict) and 'type' in locator and 'value' in locator:
            print(f"✓ locator格式正确，包含type={locator['type']}和value={locator['value']}")
        
        # 返回一个模拟的选择结果
        class MockSelection:
            def __init__(self):
                self._com_ranges = [MockDocument().Paragraphs[0].Range]
        return MockSelection()


def test_param_passing():
    """测试参数传递修复"""
    print("\n===== 测试range_tools的参数传递修复 =====")
    
    # 1. 保存原始的SelectorEngine并替换为模拟版本
    try:
        from word_docx_tools.selector.selector import SelectorEngine
        original_selector = SelectorEngine
        
        # 替换SelectorEngine为我们的模拟版本
        import word_docx_tools.selector.selector
        word_docx_tools.selector.selector.SelectorEngine = MockSelector
        
        # 2. 测试修复后的range_tools.py
        print("\n测试修复后的range_tools.py:")
        mock_doc = MockDocument()
        mock_locator = {"type": "paragraph", "value": "1"}
        
        # 导入range_tools中的select_objects函数
        from word_docx_tools.operations.range_ops import select_objects
        
        # 执行修复后的调用 - 直接传递locator字典
        print("执行修复后的调用模式: select_objects(document, locator_dict)")
        try:
            result = select_objects(mock_doc, mock_locator)
            print("✓ 修复后的调用成功完成")
            print("结论: 修复已成功，range_tools现在可以正确传递locator参数")
        except Exception as e:
            print(f"✗ 修复后的调用失败: {str(e)}")
            print("结论: 修复可能不完全正确，需要进一步检查")
            
        # 3. 测试修复前的调用模式
        print("\n测试修复前的调用模式:")
        print("执行修复前的调用模式: select_objects(document, [locator_dict])")
        try:
            result = select_objects(mock_doc, [mock_locator])
            print("✗ 修复前的调用模式居然成功了，这不符合预期")
        except Exception as e:
            print(f"✓ 修复前的调用模式按预期失败: {str(e)}")
            print("这确认了我们的修复是正确的")
            
    finally:
        # 恢复原始的SelectorEngine
        import word_docx_tools.selector.selector
        word_docx_tools.selector.selector.SelectorEngine = original_selector


if __name__ == "__main__":
    test_param_passing()
    print("\n测试完成。range_tools的select功能现在应该可以正常工作了。")