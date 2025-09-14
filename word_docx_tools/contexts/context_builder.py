"""
上下文构建工具模块

此模块提供文档上下文树的创建、构建和批量处理功能，支持高效地构建完整的文档结构上下文。
"""

import time
import logging
from typing import Optional, Dict, List, Any, Tuple, Set
from win32com.client import CDispatch
from ..utils.com_error_handler import handle_com_error
from ..utils.logger import get_logger
from .context_control import DocumentContext
from .metadata_processing import (
    create_document_metadata,
    create_section_metadata,
    create_paragraph_metadata,
    create_table_metadata,
    create_image_metadata
)

# 导入文档大纲相关的函数
from ..operations.document_ops import get_document_outline, build_hierarchical_outline_by_level

# 获取日志记录器
logger = get_logger(__name__)


@handle_com_error
def create_document_context_tree(document: CDispatch, word_app: Optional[CDispatch] = None) -> Optional[DocumentContext]:
    """
    创建文档上下文树，从文档对象构建完整的上下文层次结构，优先考虑文档大纲，同时结合section信息
    
    Args:
        document: Word文档对象
        word_app: Word应用程序实例，如果为None则尝试从文档中获取
        
    Returns:
        文档上下文树的根节点，如果创建失败则返回None
    """
    if not document:
        logger.error("No document provided to create context tree")
        return None
    
    try:
        start_time = time.time()
        
        # 尝试从文档获取Word应用实例
        if not word_app:
            try:
                word_app = document.Application
            except Exception:
                logger.warning("Failed to get Word application from document")
        
        # 创建文档根上下文
        document_metadata = create_document_metadata(document)
        root_context = DocumentContext.create_root_context(document, metadata=document_metadata)
        
        if not root_context:
            logger.error("Failed to create root document context")
            return None
        
        # 1. 首先获取并处理文档大纲（优先）
        outline_headings = []
        try:
            # 收集所有标题段落
            for i in range(1, document.Paragraphs.Count + 1):
                paragraph = document.Paragraphs(i)
                
                # 获取段落大纲级别
                outline_level = 0
                if hasattr(paragraph, 'OutlineLevel'):
                    outline_level = paragraph.OutlineLevel
                
                # 1~9为标题样式，10为正文
                if 1 <= outline_level <= 9:
                    # 获取段落样式信息
                    style_name = paragraph.Style.NameLocal if hasattr(paragraph.Style, 'NameLocal') else ""
                    
                    outline_headings.append({
                        "index": i,
                        "text": paragraph.Range.Text.strip(),
                        "outline_level": outline_level,
                        "style_name": style_name,
                        "paragraph_obj": paragraph,
                        "range_obj": paragraph.Range
                    })
            
            if outline_headings:
                logger.info(f"Found {len(outline_headings)} outline headings")
                
                # 构建层次化的大纲结构
                hierarchical_outline = build_hierarchical_outline_by_level(outline_headings)
                
                # 根据大纲结构构建上下文树
                _build_outline_context_tree(root_context, hierarchical_outline, word_app)
            else:
                logger.warning("No outline headings found in the document")
        except Exception as e:
            logger.error(f"Error processing document outline: {e}")
        
        # 2. 处理文档中的所有节（如果大纲不存在或不完整）
        sections_count = 0
        try:
            sections = document.Sections
            sections_count = len(sections)
            
            # 检查是否已经通过大纲处理了所有节内容
            # 如果已经有大纲上下文，我们仍然需要处理节但不重复创建内容上下文
            for section in sections:
                # 创建节上下文
                section_metadata = create_section_metadata(section)
                section_context = DocumentContext(
                    title=f"Section {section_metadata.get('section_index', 'unknown')}",
                    range_obj=section.Range,
                    metadata=section_metadata
                )
                
                # 检查这个节是否已经在大纲上下文中被处理
                # 我们可以通过比较范围来判断
                section_processed = False
                for child in root_context.child_contexts:
                    try:
                        if hasattr(child, 'range') and child.range and \
                           child.range.Start <= section.Range.Start and child.range.End >= section.Range.End:
                            section_processed = True
                            break
                    except Exception:
                        continue
                
                # 如果节没有被处理，则处理其内容
                if not section_processed:
                    # 构建节内容（优化版）
                    _build_section_content_optimized(section, section_context, word_app)
                
                # 将节上下文添加到根上下文
                root_context.add_child_context(section_context)
        except Exception as e:
            logger.error(f"Error processing sections: {e}")
        
        # 更新根上下文的元数据
        root_context._update_metadata({
            'sections_count': sections_count,
            'outline_headings_count': len(outline_headings),
            'build_time': time.time() - start_time,
            'build_completed': True
        })
        
        logger.info(f"Document context tree created successfully with {sections_count} sections and {len(outline_headings)} outline headings")
        return root_context
    except Exception as e:
        logger.error(f"Failed to create document context tree: {e}")
        return None


@handle_com_error
def _build_outline_context_tree(parent_context: DocumentContext, outline_items: list, word_app: Optional[CDispatch] = None) -> None:
    """
    根据层次化的大纲结构构建上下文树
    
    Args:
        parent_context: 父上下文对象
        outline_items: 层次化的大纲结构
        word_app: Word应用程序实例
    """
    if not outline_items:
        return
    
    try:
        for item in outline_items:
            # 创建大纲项目的上下文
            heading_metadata = {
                'type': 'outline_heading',
                'outline_level': item.get('outline_level', 0),
                'style_name': item.get('style_name', ''),
                'heading_text': item.get('text', ''),
                'heading_index': item.get('index', 0)
            }
            
            heading_context = DocumentContext(
                title=f"{item.get('text', '').strip()}",
                range_obj=item.get('range_obj'),
                metadata=heading_metadata
            )
            
            # 查找并添加该标题下的所有内容（直到下一个同级或更高级标题）
            if 'paragraph_obj' in item:
                _add_content_below_heading(heading_context, item['paragraph_obj'], item.get('outline_level', 0), word_app)
            
            # 添加子上下文
            parent_context.add_child_context(heading_context)
            
            # 递归处理子标题
            if 'children' in item and item['children']:
                _build_outline_context_tree(heading_context, item['children'], word_app)
    except Exception as e:
        logger.error(f"Error building outline context tree: {e}")


@handle_com_error
def _add_content_below_heading(heading_context: DocumentContext, heading_paragraph: CDispatch, heading_level: int, word_app: Optional[CDispatch] = None) -> None:
    """
    添加标题下的所有内容（直到下一个同级或更高级标题）
    
    Args:
        heading_context: 标题上下文对象
        heading_paragraph: 标题段落对象
        heading_level: 标题级别
        word_app: Word应用程序实例
    """
    try:
        document = heading_paragraph.Document
        current_paragraph = heading_paragraph
        
        # 获取当前标题的起始位置
        heading_start = heading_paragraph.Range.Start
        
        # 获取文档中的所有段落
        paragraphs = document.Paragraphs
        
        # 查找段落索引
        heading_index = 0
        for i in range(1, paragraphs.Count + 1):
            if paragraphs(i) == heading_paragraph:
                heading_index = i
                break
        
        # 处理标题后的内容
        if heading_index > 0:
            content_section_metadata = {
                'type': 'heading_content',
                'parent_heading_level': heading_level,
                'parent_heading_index': heading_index
            }
            
            # 创建内容部分的上下文
            content_context = DocumentContext(
                title=f"Content under heading {heading_index}",
                metadata=content_section_metadata
            )
            
            # 遍历标题后的段落
            for i in range(heading_index + 1, paragraphs.Count + 1):
                para = paragraphs(i)
                
                # 检查是否是下一个同级或更高级标题
                current_outline_level = 0
                if hasattr(para, 'OutlineLevel'):
                    current_outline_level = para.OutlineLevel
                
                # 如果遇到同级或更高级标题，则停止
                if 1 <= current_outline_level <= heading_level:
                    break
                
                # 处理普通段落内容
                _process_single_paragraph_content(para, content_context)
            
            # 将内容上下文添加到标题上下文
            if content_context.child_contexts:
                heading_context.add_child_context(content_context)
    except Exception as e:
        logger.error(f"Error adding content below heading: {e}")


@handle_com_error
def _process_single_paragraph_content(paragraph: CDispatch, parent_context: DocumentContext) -> None:
    """
    处理单个段落的内容，包括其中的表格和图片
    
    Args:
        paragraph: 段落对象
        parent_context: 父上下文对象
    """
    try:
        # 检查段落中是否包含表格
        if hasattr(paragraph.Range, 'Tables') and paragraph.Range.Tables.Count > 0:
            for table in paragraph.Range.Tables:
                try:
                    # 创建表格元数据
                    table_metadata = create_table_metadata(table)
                    
                    # 创建表格上下文
                    table_context = DocumentContext(
                        title=f"Table in paragraph",
                        range_obj=table.Range,
                        metadata=table_metadata
                    )
                    
                    # 添加表格对象信息
                    table_context.batch_add_objects([table_metadata])
                    
                    # 添加到父上下文
                    parent_context.add_child_context(table_context)
                except Exception as e:
                    logger.warning(f"Error processing table in paragraph: {e}")
        
        # 检查段落中是否包含图片
        if hasattr(paragraph.Range, 'InlineShapes'):
            for shape in paragraph.Range.InlineShapes:
                try:
                    # 只处理图片类型的形状
                    if hasattr(shape, 'Type') and shape.Type == 3:  # wdInlineShapePicture
                        # 创建图片元数据
                        image_metadata = create_image_metadata(shape)
                        
                        # 创建图片上下文
                        image_context = DocumentContext(
                            title=f"Image in paragraph",
                            range_obj=shape.Range,
                            metadata=image_metadata
                        )
                        
                        # 添加图片对象信息
                        image_context.batch_add_objects([image_metadata])
                        
                        # 添加到父上下文
                        parent_context.add_child_context(image_context)
                except Exception as e:
                    logger.warning(f"Error processing image in paragraph: {e}")
        
        # 创建段落元数据
        para_metadata = create_paragraph_metadata(paragraph)
        
        # 创建段落上下文
        para_context = DocumentContext(
            title=f"Content paragraph",
            range_obj=paragraph.Range,
            metadata=para_metadata
        )
        
        # 添加段落对象信息
        para_context.batch_add_objects([para_metadata])
        
        # 添加到父上下文
        parent_context.add_child_context(para_context)
    except Exception as e:
        logger.warning(f"Error processing paragraph content: {e}")


@handle_com_error
def _build_section_content_optimized(section: CDispatch, parent_context: DocumentContext, 
                                   word_app: Optional[CDispatch] = None) -> Optional[DocumentContext]:
    """
    构建单个节的内容上下文（优化版），批量处理节内的所有元素
    
    Args:
        section: Word节对象
        parent_context: 父上下文对象（通常是文档根上下文）
        word_app: Word应用程序实例
        
    Returns:
        包含节内容的上下文对象，如果构建失败则返回None
    """
    if not section:
        logger.warning("No section provided to build content")
        return None
    
    try:
        start_time = time.time()
        
        # 创建节上下文
        section_metadata = create_section_metadata(section)
        section_context = DocumentContext(
            title=f"Section {section_metadata.get('section_index', 'unknown')}",
            range_obj=section.Range,
            metadata=section_metadata
        )
        
        # 批量处理节内的表格
        tables_processed = _process_tables_in_section(section, section_context)
        
        # 批量处理节内的图片
        images_processed = _process_images_in_section(section, section_context)
        
        # 批量处理节内的段落（排除表格内的段落）
        paragraphs_processed = _process_paragraphs_in_section(section, section_context)
        
        # 更新节上下文的元数据
        section_context._update_metadata({
            'tables_count': tables_processed,
            'images_count': images_processed,
            'paragraphs_count': paragraphs_processed,
            'processing_time': time.time() - start_time,
            'is_processed': True
        })
        
        logger.debug(f"Processed section content: {paragraphs_processed} paragraphs, {tables_processed} tables, {images_processed} images")
        return section_context
    except Exception as e:
        logger.error(f"Failed to build section content: {e}")
        return None


@handle_com_error
def _process_tables_in_section(section: CDispatch, section_context: DocumentContext) -> int:
    """
    处理节内的所有表格，为每个表格创建上下文
    
    Args:
        section: Word节对象
        section_context: 节上下文对象
        
    Returns:
        成功处理的表格数量
    """
    processed_count = 0
    
    try:
        if not hasattr(section.Range, 'Tables'):
            return 0
        
        tables = section.Range.Tables
        
        # 使用批处理方式处理表格
        table_contexts = []
        processed_table_ids = set()  # 用于跟踪已处理的表格，避免重复处理
        
        for table in tables:
            try:
                # 获取表格的起始位置作为唯一标识
                table_start = table.Range.Start
                table_id = f"table_{table_start}"
                
                # 避免重复处理
                if table_id in processed_table_ids:
                    continue
                processed_table_ids.add(table_id)
                
                # 创建表格元数据
                table_metadata = create_table_metadata(table)
                
                # 创建表格上下文
                table_context = DocumentContext(
                    title=f"Table {table_metadata.get('table_index', 'unknown')}",
                    range_obj=table.Range,
                    metadata=table_metadata
                )
                
                # 添加表格对象信息
                table_context.batch_add_objects([table_metadata])
                
                table_contexts.append(table_context)
                processed_count += 1
            except Exception as e:
                logger.warning(f"Error processing table: {e}")
                continue
        
        # 批量添加表格上下文到节上下文
        for table_context in table_contexts:
            section_context.add_child_context(table_context)
    except Exception as e:
        logger.error(f"Error processing tables in section: {e}")
    
    return processed_count


@handle_com_error
def _process_images_in_section(section: CDispatch, section_context: DocumentContext) -> int:
    """
    处理节内的所有图片，为每个图片创建上下文
    
    Args:
        section: Word节对象
        section_context: 节上下文对象
        
    Returns:
        成功处理的图片数量
    """
    processed_count = 0
    
    try:
        if not hasattr(section.Range, 'InlineShapes'):
            return 0
        
        inline_shapes = section.Range.InlineShapes
        
        # 使用批处理方式处理图片
        image_contexts = []
        processed_image_ids = set()  # 用于跟踪已处理的图片，避免重复处理
        
        for shape in inline_shapes:
            try:
                # 只处理图片类型的形状
                if hasattr(shape, 'Type') and shape.Type != 3:  # wdInlineShapePicture
                    continue
                
                # 获取图片的起始位置作为唯一标识
                image_start = shape.Range.Start
                image_id = f"image_{image_start}"
                
                # 避免重复处理
                if image_id in processed_image_ids:
                    continue
                processed_image_ids.add(image_id)
                
                # 创建图片元数据
                image_metadata = create_image_metadata(shape)
                
                # 创建图片上下文
                image_context = DocumentContext(
                    title=f"Image {image_metadata.get('image_index', 'unknown')}",
                    range_obj=shape.Range,
                    metadata=image_metadata
                )
                
                # 添加图片对象信息
                image_context.batch_add_objects([image_metadata])
                
                image_contexts.append(image_context)
                processed_count += 1
            except Exception as e:
                logger.warning(f"Error processing image: {e}")
                continue
        
        # 批量添加图片上下文到节上下文
        for image_context in image_contexts:
            section_context.add_child_context(image_context)
    except Exception as e:
        logger.error(f"Error processing images in section: {e}")
    
    return processed_count


@handle_com_error
def _process_paragraphs_in_section(section: CDispatch, section_context: DocumentContext) -> int:
    """
    处理节内的所有段落，为每个段落创建上下文
    
    Args:
        section: Word节对象
        section_context: 节上下文对象
        
    Returns:
        成功处理的段落数量
    """
    processed_count = 0
    
    try:
        if not hasattr(section.Range, 'Paragraphs'):
            return 0
        
        paragraphs = section.Range.Paragraphs
        
        # 使用批处理方式处理段落
        paragraph_contexts = []
        processed_para_ids = set()  # 用于跟踪已处理的段落，避免重复处理
        
        # 获取所有表格的范围，避免处理表格内的段落
        table_ranges = []
        if hasattr(section.Range, 'Tables'):
            for table in section.Range.Tables:
                try:
                    table_ranges.append((table.Range.Start, table.Range.End))
                except Exception:
                    continue
        
        for paragraph in paragraphs:
            try:
                para_range = paragraph.Range
                para_start = para_range.Start
                para_id = f"paragraph_{para_start}"
                
                # 避免重复处理
                if para_id in processed_para_ids:
                    continue
                
                # 检查段落是否在表格内，如果是则跳过
                if _is_range_in_tables(para_range, table_ranges):
                    continue
                
                processed_para_ids.add(para_id)
                
                # 创建段落元数据
                para_metadata = create_paragraph_metadata(paragraph)
                
                # 创建段落上下文
                para_context = DocumentContext(
                    title=f"Paragraph {processed_count + 1}",
                    range_obj=para_range,
                    metadata=para_metadata
                )
                
                # 添加段落对象信息
                para_context.batch_add_objects([para_metadata])
                
                paragraph_contexts.append(para_context)
                processed_count += 1
            except Exception as e:
                logger.warning(f"Error processing paragraph: {e}")
                continue
        
        # 批量添加段落上下文到节上下文
        for para_context in paragraph_contexts:
            section_context.add_child_context(para_context)
    except Exception as e:
        logger.error(f"Error processing paragraphs in section: {e}")
    
    return processed_count


@handle_com_error
def _is_range_in_tables(range_obj: CDispatch, table_ranges: List[Tuple[int, int]]) -> bool:
    """
    检查指定的范围是否在表格内
    
    Args:
        range_obj: 要检查的Range对象
        table_ranges: 表格范围的列表，每个元素是(start, end)元组
        
    Returns:
        如果范围在表格内则返回True，否则返回False
    """
    try:
        range_start = range_obj.Start
        range_end = range_obj.End
        
        for table_start, table_end in table_ranges:
            if range_start >= table_start and range_end <= table_end:
                return True
        
        return False
    except Exception:
        # 如果发生错误，安全起见假设不在表格内
        return False


@handle_com_error
def find_section_for_range(document: CDispatch, range_obj: CDispatch) -> Optional[CDispatch]:
    """
    查找给定Range所在的节
    
    Args:
        document: Word文档对象
        range_obj: Range对象
        
    Returns:
        节对象，如果未找到则返回None
    """
    if not document or not range_obj:
        logger.warning("Invalid document or range object")
        return None
    
    try:
        sections = document.Sections
        
        # 获取范围的起始和结束位置
        try:
            range_start = range_obj.Start
            range_end = range_obj.End
        except Exception as e:
            logger.error(f"Failed to get range positions: {e}")
            return None
        
        # 遍历所有节查找包含该范围的节
        for section in sections:
            try:
                section_start = section.Range.Start
                section_end = section.Range.End
                
                # 检查范围是否完全包含在节内
                if range_start >= section_start and range_end <= section_end:
                    return section
            except Exception:
                continue
        
        logger.warning(f"No section found for range {range_start}-{range_end}")
        return None
    except Exception as e:
        logger.error(f"Failed to find section for range: {e}")
        return None


@handle_com_error
def batch_add_contexts_in_section(section: CDispatch, section_context: DocumentContext, 
                                 object_type: str = 'all') -> Dict[str, int]:
    """
    批量在节内添加上下文，支持指定对象类型
    
    Args:
        section: Word节对象
        section_context: 节上下文对象
        object_type: 要添加的对象类型，可选值: 'all', 'paragraph', 'table', 'image'
        
    Returns:
        包含添加结果统计的字典
    """
    result = {
        'paragraphs': 0,
        'tables': 0,
        'images': 0,
        'total': 0
    }
    
    try:
        if object_type == 'all' or object_type == 'table':
            result['tables'] = _process_tables_in_section(section, section_context)
        
        if object_type == 'all' or object_type == 'image':
            result['images'] = _process_images_in_section(section, section_context)
        
        if object_type == 'all' or object_type == 'paragraph':
            result['paragraphs'] = _process_paragraphs_in_section(section, section_context)
        
        # 计算总数
        result['total'] = result['paragraphs'] + result['tables'] + result['images']
        
        logger.info(f"Batch added {result['total']} contexts in section")
    except Exception as e:
        logger.error(f"Error in batch adding contexts: {e}")