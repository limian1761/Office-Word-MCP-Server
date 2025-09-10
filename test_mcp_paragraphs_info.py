#!/usr/bin/env python3
"""
测试 MCP 工具的 get_paragraphs_info 功能
"""

import win32com.client
import json
import os

# 创建 Word 应用程序实例
word_app = win32com.client.Dispatch("Word.Application")
word_app.Visible = True

# 创建新文档
doc = word_app.Documents.Add()

# 添加测试内容
# 标题1
range_obj = doc.Range()
range_obj.InsertAfter("第一章 概述\n")
range_obj.SetRange(range_obj.End, range_obj.End)
range_obj.Style = doc.Styles("标题 1")

# 正文内容
range_obj.InsertAfter("这是第一章的概述内容。这是第一句话。这是第二句话。这是第三句话。\n")
range_obj.SetRange(range_obj.End, range_obj.End)

# 标题2
range_obj.InsertAfter("1.1 技术背景\n")
range_obj.SetRange(range_obj.End, range_obj.End)
range_obj.Style = doc.Styles("标题 2")

# 正文内容
range_obj.InsertAfter("这是技术背景的详细说明。包含多个句子。用于测试段落信息提取功能。\n")
range_obj.SetRange(range_obj.End, range_obj.End)

# 空段落（用于测试空段落属性）
range_obj.InsertAfter("\n")
range_obj.SetRange(range_obj.End, range_obj.End)

print("文档内容已创建")
print(f"总段落数: {doc.Paragraphs.Count}")

# 模拟 MCP 工具的 get_paragraphs_info 功能
def get_paragraphs_info(document):
    """模拟 MCP 工具的 get_paragraphs_info 功能"""
    try:
        total_paragraphs = document.Paragraphs.Count
        empty_paragraphs = 0
        total_characters = 0
        total_words = 0
        paragraphs_info = []

        # 获取所有段落的详细信息
        for i in range(1, total_paragraphs + 1):
            paragraph = document.Paragraphs(i)
            paragraph_range = paragraph.Range
            text = paragraph_range.Text.strip()
            
            # 计算字符数和字数
            paragraph_chars = len(text)
            paragraph_words = len(text.split()) if text else 0
            
            total_characters += paragraph_chars
            total_words += paragraph_words
            
            # 统计空段落
            if not text:
                empty_paragraphs += 1
            
            # 获取段落样式信息
            style_name = ""
            try:
                if hasattr(paragraph, "Style") and paragraph.Style:
                    style_name = paragraph.Style.NameLocal
            except:
                pass
            
            # 获取段落开头和结尾的句子
            opening_sentence = ""
            closing_sentence = ""
            if text:
                sentences = text.split('.')
                if sentences:
                    opening_sentence = sentences[0].strip() + ('.' if len(sentences) > 1 else '')
                    closing_sentence = sentences[-1].strip() + '.' if sentences and len(sentences) > 1 else ""
            
            # 构建段落信息
            paragraph_info = {
                "index": i,
                "text_length": paragraph_chars,
                "word_count": paragraph_words,
                "style_name": style_name,
                "opening_sentence": opening_sentence,
                "closing_sentence": closing_sentence,
                "is_empty": not bool(text)
            }
            
            # 如果段落包含文字，添加文字内容
            if text:
                paragraph_info["text_preview"] = text[:100] + "..." if len(text) > 100 else text
            else:
                # 对于空段落，添加其他属性信息
                paragraph_info["has_formatting"] = hasattr(paragraph, "Format") and paragraph.Format is not None
                try:
                    paragraph_info["outline_level"] = paragraph.OutlineLevel if hasattr(paragraph, "OutlineLevel") else 0
                except:
                    paragraph_info["outline_level"] = 0
                
            paragraphs_info.append(paragraph_info)
        
        # 计算平均值
        non_empty_paragraphs = total_paragraphs - empty_paragraphs
        avg_chars_per_paragraph = total_characters / non_empty_paragraphs if non_empty_paragraphs > 0 else 0
        avg_words_per_paragraph = total_words / non_empty_paragraphs if non_empty_paragraphs > 0 else 0
        
        # 构建结果
        stats = {
            "total_paragraphs": total_paragraphs,
            "empty_paragraphs": empty_paragraphs,
            "non_empty_paragraphs": non_empty_paragraphs,
            "total_characters": total_characters,
            "total_words": total_words,
            "avg_characters_per_paragraph": round(avg_chars_per_paragraph, 2),
            "avg_words_per_paragraph": round(avg_words_per_paragraph, 2)
        }
        
        return {
            "success": True, 
            "statistics": stats,
            "paragraphs": paragraphs_info
        }
        
    except Exception as e:
        return {"success": False, "error": str(e)}

# 测试函数
result = get_paragraphs_info(doc)

if result["success"]:
    print("\n✅ get_paragraphs_info 测试成功!")
    print(f"统计信息: {json.dumps(result['statistics'], ensure_ascii=False, indent=2)}")
    
    print("\n段落详细信息:")
    for paragraph in result['paragraphs']:
        print(f"段落 {paragraph['index']}:")
        print(f"  样式: {paragraph['style_name']}")
        print(f"  字数: {paragraph['word_count']}")
        print(f"  字符数: {paragraph['text_length']}")
        print(f"  开头句子: {paragraph['opening_sentence']}")
        print(f"  结尾句子: {paragraph['closing_sentence']}")
        print(f"  是否为空: {paragraph['is_empty']}")
        
        if paragraph['is_empty']:
            print(f"  是否有格式: {paragraph.get('has_formatting', 'N/A')}")
            print(f"  大纲级别: {paragraph.get('outline_level', 'N/A')}")
        else:
            print(f"  文本预览: {paragraph.get('text_preview', 'N/A')}")
        print()
else:
    print(f"❌ 测试失败: {result['error']}")

# 保存文档
doc_path = os.path.join(os.getcwd(), "test_mcp_paragraphs_info.docx")
doc.SaveAs(doc_path)
print(f"文档已保存到: {doc_path}")

print("\n测试完成!")