import os
import traceback
from word_document_server.com_backend import WordBackend
from word_document_server.errors import WordDocumentError, ErrorCode


def create_test_document(doc_path):
    """创建一个包含中文标题的测试文档，用于验证样式应用"""
    # 确保目录存在
    os.makedirs(os.path.dirname(doc_path), exist_ok=True)
    
    # 如果文档已存在，先删除
    if os.path.exists(doc_path):
        try:
            os.remove(doc_path)
            print(f"已删除旧测试文档: {doc_path}")
        except Exception as e:
            print(f"删除旧测试文档失败: {e}")
            print(f"删除错误详细信息: {traceback.format_exc()}")

    try:
        # 创建新文档，不指定file_path
        print("尝试启动Word应用程序...")
        with WordBackend(visible=True) as backend:
            print("成功启动Word应用程序并创建文档")
            
            # 先获取并打印所有可用样式，以确认样式名称
            try:
                styles = backend.get_document_styles()
                print("可用样式列表:")
                for style in styles:
                    print(f"- {style['name']} (类型: {style['type']})")
            except Exception as e:
                print(f"获取样式失败: {e}")
                print(f"获取样式错误详细信息: {traceback.format_exc()}")

            # 获取文档的起始范围
            doc_range = backend.document.Range(0, 0)
            
            # 检查可用样式，确保标题样式存在
            style_info = {}
            available_styles = []
            try:
                print("可用段落样式列表:")
                for style in backend.document.Styles:
                    if style.Type == 1:  # 段落样式
                        style_name = style.NameLocal
                        available_styles.append(style_name)
                        # 不使用ID，直接存储样式名称
                        style_info[style_name] = style_name
                        print(f"- 名称: {style_name}, 类型: {style.Type}")
                print(f"可用段落样式列表: {', '.join(available_styles)}")
            except Exception as e:
                print(f"获取样式失败: {e}")
                print(f"样式获取错误详细信息: {traceback.format_exc()}")

            # 添加段落并应用样式
            try:
                # 确保文档为空
                backend.document.Range(0, backend.document.Content.End).Delete()
                print("已清空文档内容")

                # 获取文档范围
                doc_range = backend.document.Range(0, 0)

                # 添加标题 1 段落
                print("添加标题 1 段落...")
                doc_range.InsertAfter("第一章 介绍")
                doc_range.InsertParagraphAfter()  # 创建新段落
                # 移动范围到新段落
                doc_range.SetRange(backend.document.Content.End, backend.document.Content.End)
                # 获取刚添加的段落
                if backend.document.Paragraphs.Count >= 1:
                    para1 = backend.document.Paragraphs(1)
                    # 应用标题 1 样式
                    if "标题 1" in style_info:
                        try:
                            style_name = style_info["标题 1"]
                            para1.Style = style_name
                            actual_style = para1.Style.NameLocal
                            print(f"应用'标题 1'样式成功，实际样式: {actual_style}")
                            print(f"标题 1 段落内容: {para1.Range.Text.strip()}")
                        except Exception as style_e:
                            print(f"应用'标题 1'样式失败: {style_e}")
                    else:
                        print("警告: 未找到'标题 1'样式")

                # 添加标题 2 段落
                print("添加标题 2 段落...")
                doc_range.InsertAfter("1.1 研究背景")
                doc_range.InsertParagraphAfter()  # 创建新段落
                # 移动范围到新段落
                doc_range.SetRange(backend.document.Content.End, backend.document.Content.End)
                # 获取刚添加的段落
                if backend.document.Paragraphs.Count >= 2:
                    para2 = backend.document.Paragraphs(2)
                    # 应用标题 2 样式
                    if "标题 2" in style_info:
                        try:
                            style_name = style_info["标题 2"]
                            para2.Style = style_name
                            actual_style = para2.Style.NameLocal
                            print(f"应用'标题 2'样式成功，实际样式: {actual_style}")
                            print(f"标题 2 段落内容: {para2.Range.Text.strip()}")
                        except Exception as style_e:
                            print(f"应用'标题 2'样式失败: {style_e}")
                    else:
                        print("警告: 未找到'标题 2'样式")

                # 添加正文段落
                print("添加正文段落...")
                doc_range.InsertAfter("这是正文内容。")
                doc_range.InsertParagraphAfter()  # 创建新段落
                # 移动范围到文档末尾
                doc_range.SetRange(backend.document.Content.End, backend.document.Content.End)
                # 获取刚添加的段落
                if backend.document.Paragraphs.Count >= 3:
                    para3 = backend.document.Paragraphs(3)
                    # 应用正文样式
                    if "正文" in style_info:
                        try:
                            style_name = style_info["正文"]
                            para3.Style = style_name
                            actual_style = para3.Style.NameLocal
                            print(f"应用'正文'样式成功，实际样式: {actual_style}")
                            print(f"正文段落内容: {para3.Range.Text.strip()}")
                        except Exception as style_e:
                            print(f"应用'正文'样式失败: {style_e}")
                    else:
                        print("警告: 未找到'正文'样式")

                # 确认所有段落都已添加
                print(f"总段落数: {backend.document.Paragraphs.Count}")
                # 打印所有段落内容和样式
                for i in range(1, backend.document.Paragraphs.Count + 1):
                    para = backend.document.Paragraphs(i)
                    print(f"段落 {i} 内容: {para.Range.Text.strip()}")
                    print(f"段落 {i} 样式: {para.Style.NameLocal}")
            except Exception as e:
                print(f"添加段落失败: {e}")
                print(f"段落添加错误详细信息: {traceback.format_exc()}")

            # 添加标题 2 段落
            try:
                para2 = backend.document.Paragraphs.Add()
                if para2 is not None:
                    print("成功添加标题 2 段落")
                    para2.Range.Text = "1.1 研究背景"
                    # 应用样式
                    title2_id = style_info.get("标题 2")
                    if title2_id:
                        try:
                            para2.Style = backend.document.Styles(title2_id)
                            actual_style = para2.Style.NameLocal
                            print(f"使用ID应用'标题 2'样式，实际样式: {actual_style}")
                        except Exception as style_e:
                            print(f"使用ID应用'标题 2'样式失败: {style_e}")
                            try:
                                para2.Style = "标题 2"
                                actual_style = para2.Style.NameLocal
                                print(f"使用名称应用'标题 2'样式，实际样式: {actual_style}")
                            except Exception as name_e:
                                print(f"使用名称应用'标题 2'样式失败: {name_e}")
                    else:
                        print("警告: 未找到'标题 2'样式")
                else:
                    print("警告: Paragraphs.Add 返回 None")
            except Exception as e:
                print(f"添加标题 2 段落失败: {e}")
                print(f"段落添加错误详细信息: {traceback.format_exc()}")

            # 添加正文段落
            try:
                para3 = backend.document.Paragraphs.Add()
                if para3 is not None:
                    print("成功添加正文段落")
                    para3.Range.Text = "这是正文内容333。"
                    # 应用样式
                    body_id = style_info.get("正文")
                    if body_id:
                        try:
                            para3.Style = backend.document.Styles(body_id)
                            actual_style = para3.Style.NameLocal
                            print(f"使用ID应用'正文'样式，实际样式: {actual_style}")
                        except Exception as style_e:
                            print(f"使用ID应用'正文'样式失败: {style_e}")
                            try:
                                para3.Style = "正文"
                                actual_style = para3.Style.NameLocal
                                print(f"使用名称应用'正文'样式，实际样式: {actual_style}")
                            except Exception as name_e:
                                print(f"使用名称应用'正文'样式失败: {name_e}")
                    else:
                        print("警告: 未找到'正文'样式")
                else:
                    print("警告: Paragraphs.Add 返回 None")
            except Exception as e:
                print(f"添加正文段落失败: {e}")
                print(f"段落添加错误详细信息: {traceback.format_exc()}")

            # 保存前再次检查所有段落样式和内容
            print("\n===== 保存前最终样式和内容检查 ====")
            try:
                for i in range(1, len(backend.document.Paragraphs) + 1):
                    para = backend.document.Paragraphs(i)
                    print(f"段落 {i} 样式: {para.Style.NameLocal}")
                    print(f"段落内容: {para.Range.Text.strip()}")
                    # 确保段落有正确的结束标记
                    if not para.Range.Text.endswith('\r'):
                        print(f"段落 {i} 缺少结束标记，添加...")
                        para.Range.InsertParagraphAfter()
            except Exception as e:
                print(f"检查段落样式和内容失败: {e}")
                print(f"样式和内容检查错误详细信息: {traceback.format_exc()}")

            # 强制保存样式 - 使用更健壮的方式
            try:
                print("\n===== 强制重新应用样式 ====")
                for i in range(1, len(backend.document.Paragraphs) + 1):
                    para = backend.document.Paragraphs(i)
                    if i == 1 and "标题 1" in available_styles:
                        try:
                            para.Style = "标题 1"
                            print(f"段落 {i} 已重新应用'标题 1'样式")
                        except Exception as e:
                            print(f"段落 {i} 重新应用'标题 1'样式失败: {e}")
                    elif i == 2 and "标题 2" in available_styles:
                        try:
                            para.Style = "标题 2"
                            print(f"段落 {i} 已重新应用'标题 2'样式")
                        except Exception as e:
                            print(f"段落 {i} 重新应用'标题 2'样式失败: {e}")
                    elif i > 2 and "正文" in available_styles:
                        try:
                            para.Style = "正文"
                            print(f"段落 {i} 已重新应用'正文'样式")
                        except Exception as e:
                            print(f"段落 {i} 重新应用'正文'样式失败: {e}")
            except Exception as e:
                print(f"强制应用样式失败: {e}")
                print(f"强制应用样式错误详细信息: {traceback.format_exc()}")

            # 再次检查样式
            print("\n===== 强制应用样式后检查 ====")
            try:
                for i in range(1, len(backend.document.Paragraphs) + 1):
                    para = backend.document.Paragraphs(i)
                    print(f"段落 {i} 样式: {para.Style.NameLocal}")
            except Exception as e:
                print(f"强制应用后样式检查失败: {e}")
            
            # 在保存前检查所有段落的样式
            print("\n===== 文档保存前段落样式检查 ====")
            try:
                if hasattr(backend.document, 'Paragraphs') and backend.document.Paragraphs.Count > 0:
                    for i, para in enumerate(backend.document.Paragraphs):
                        try:
                            style_name = para.Style.NameLocal
                            content = para.Range.Text.strip()
                            print(f"段落 {i+1} 样式: {style_name}")
                            print(f"段落内容: {content}")
                        except Exception as para_e:
                            print(f"警告: 处理段落 {i+1} 时出错: {para_e}")
                else:
                    print("文档中没有段落")
            except Exception as e:
                print(f"警告: 检查段落样式时出错: {e}")
                print(f"样式检查错误详细信息: {traceback.format_exc()}")

            # 保存文档前最后检查样式
            print("\n===== 保存前最后样式检查 ====")
            try:
                if hasattr(backend.document, 'Paragraphs') and backend.document.Paragraphs.Count > 0:
                    for i, para in enumerate(backend.document.Paragraphs):
                        try:
                            style_name = para.Style.NameLocal
                            content = para.Range.Text.strip()
                            print(f"段落 {i+1} 样式: {style_name}")
                            print(f"段落内容: {content}")
                        except Exception as para_e:
                            print(f"警告: 处理段落 {i+1} 时出错: {para_e}")
                else:
                    print("文档中没有段落")
            except Exception as e:
                print(f"警告: 检查段落样式时出错: {e}")

            # 保存文档
            try:
                # 使用document.SaveAs方法保存文档，指定格式
                print("尝试保存文档...")
                backend.document.SaveAs(doc_path, FileFormat=16)  # 16 表示docx格式
                print(f"测试文档已保存: {doc_path}")
                # 强制保存并等待
                backend.document.Save()
                print("文档已强制保存")
                # 确认文档已保存
                if os.path.exists(doc_path):
                    print(f"确认文档已保存，文件大小: {os.path.getsize(doc_path)} 字节")
                else:
                    print("警告: 保存后文档不存在")
            except Exception as e:
                print(f"保存文档失败: {e}")
                print(f"保存错误详细信息: {traceback.format_exc()}")

            # 关闭前等待一段时间
            print("保存后等待2秒再关闭文档...")
            import time
            time.sleep(2)
    except Exception as e:
        print(f"创建测试文档失败: {e}")
        print(f"错误详细信息: {traceback.format_exc()}")


def test_document_structure():
    # 测试文档路径
    doc_path = os.path.join(os.path.dirname(__file__), 'tests', 'test_docs', 'chinese_heading_test.docx')

    # 首先创建测试文档
    create_test_document(doc_path)
    
    # 等待一段时间确保文档已完全保存
    import time
    print("等待2秒确保文档保存完成...")
    time.sleep(2)
    
    try:
        # 打开Word应用和文档
        with WordBackend(file_path=doc_path, visible=True) as backend:
            print(f"成功打开文档: {doc_path}")
            
            # 检查所有段落的样式
            print("\n===== 打开后段落样式检查 ====")
            try:
                if hasattr(backend.document, 'Paragraphs') and backend.document.Paragraphs.Count > 0:
                    for i, para in enumerate(backend.document.Paragraphs):
                        try:
                            style_name = para.Style.NameLocal
                            content = para.Range.Text.strip()
                            print(f"段落 {i+1} 样式: {style_name}")
                            print(f"段落内容: {content}")
                        except Exception as para_e:
                            print(f"警告: 处理段落 {i+1} 时出错: {para_e}")
                else:
                    print("文档中没有段落")
            except Exception as e:
                print(f"警告: 检查段落样式时出错: {e}")
            
            # 获取文档结构
            structure = backend.get_document_structure()
            
            # 打印文档结构
            if structure:
                print("文档结构:")
                for item in structure:
                    print(f"级别 {item['level']}: {item['text']}")
            else:
                print("文档结构为空。可能文档中没有使用'Heading'或'标题'样式的段落。")
                
            # 获取所有文本，以确认文档内容
            all_text = backend.get_all_text()
            print(f"文档文本长度: {len(all_text)} 字符")
            print("文档前100个字符:", all_text[:100] + ('...' if len(all_text) > 100 else ''))
            
    except WordDocumentError as e:
        print(f"发生错误: {e}")
    except Exception as e:
        print(f"发生未预期的错误: {e}")


if __name__ == "__main__":
    test_document_structure()