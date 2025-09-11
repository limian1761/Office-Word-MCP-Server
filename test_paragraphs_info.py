#!/usr/bin/env python3
"""
测试 get_paragraphs_info 函数
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

# 标题3
range_obj.InsertAfter("1.1.1 具体实现\n")
range_obj.SetRange(range_obj.End, range_obj.End)
range_obj.Style = doc.Styles("标题 3")

# 正文内容
range_obj.InsertAfter("这是具体实现的描述。包含技术细节和实现方法。\n")
range_obj.SetRange(range_obj.End, range_obj.End)

# 空段落（用于测试空段落属性）
range_obj.InsertAfter("\n")
range_obj.SetRange(range_obj.End, range_obj.End)

print("文档内容已创建，段落信息如下：")

# 获取所有段落信息
paragraphs_info = []
for i in range(1, doc.Paragraphs.Count + 1):
    paragraph = doc.Paragraphs(i)
    paragraph_range = paragraph.Range
    text = paragraph_range.Text.strip()
    
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
        "text": text,
        "text_length": len(text),
        "word_count": len(text.split()) if text else 0,
        "style_name": style_name,
        "opening_sentence": opening_sentence,
        "closing_sentence": closing_sentence,
        "is_empty": not bool(text)
    }
    
    # 如果段落包含文字，添加文字内容
    if text:
        paragraph_info["text_preview"] = text[:50] + "..." if len(text) > 50 else text
    else:
        # 对于空段落，添加其他属性信息
        paragraph_info["has_formatting"] = hasattr(paragraph, "Format") and paragraph.Format is not None
        try:
            paragraph_info["outline_level"] = paragraph.OutlineLevel if hasattr(paragraph, "OutlineLevel") else 0
        except:
            paragraph_info["outline_level"] = 0
    
    paragraphs_info.append(paragraph_info)
    
    print(f"段落 {i}: {json.dumps(paragraph_info, ensure_ascii=False, indent=2)}")

# 计算统计信息
total_paragraphs = doc.Paragraphs.Count
empty_paragraphs = sum(1 for p in paragraphs_info if p["is_empty"])
non_empty_paragraphs = total_paragraphs - empty_paragraphs
total_characters = sum(p["text_length"] for p in paragraphs_info)
total_words = sum(p["word_count"] for p in paragraphs_info)
avg_chars_per_paragraph = total_characters / non_empty_paragraphs if non_empty_paragraphs > 0 else 0
avg_words_per_paragraph = total_words / non_empty_paragraphs if non_empty_paragraphs > 0 else 0

stats = {
    "total_paragraphs": total_paragraphs,
    "empty_paragraphs": empty_paragraphs,
    "non_empty_paragraphs": non_empty_paragraphs,
    "total_characters": total_characters,
    "total_words": total_words,
    "avg_characters_per_paragraph": round(avg_chars_per_paragraph, 2),
    "avg_words_per_paragraph": round(avg_words_per_paragraph, 2)
}

print(f"\n统计信息: {json.dumps(stats, ensure_ascii=False, indent=2)}")

# 保存文档
doc_path = os.path.join(os.getcwd(), "test_paragraphs_info.docx")
doc.SaveAs(doc_path)
print(f"\n文档已保存到: {doc_path}")

# 不关闭文档，以便查看
print("\n测试完成，Word文档保持打开状态以供查看。")