import time
import traceback
from typing import Dict, List, Optional, Any, Callable, Set
from win32com.client import CDispatch
from ..common.logger import logger
from ..common.exceptions import DocumentContextError, ErrorCode
from ..com_backend.com_utils import handle_com_error
from .context_control import DocumentContext
from .context_manager import get_context_manager


class DocumentChangeHandler:
    """文档变更处理器，负责处理文档中的各种变更事件"""
    def __init__(self):
        # 更新处理器回调列表
        self._update_handlers: List[Callable[[Dict[str, Any]], None]] = []
        # 支持的变更类型
        self._supported_change_types = {
            'paragraph_inserted', 'paragraph_updated', 'paragraph_deleted',
            'table_inserted', 'table_updated', 'table_deleted',
            'image_inserted', 'image_updated', 'image_deleted',
            'document_structure_changed'
        }

    def register_update_handler(self, handler: Callable[[Dict[str, Any]], None]) -> None:
        """
        注册更新处理器回调函数
        
        Args:
            handler: 处理更新事件的回调函数
        """
        if handler not in self._update_handlers:
            self._update_handlers.append(handler)
            logger.info("Update handler registered")

    def unregister_update_handler(self, handler: Callable[[Dict[str, Any]], None]) -> None:
        """
        取消注册更新处理器回调函数
        
        Args:
            handler: 要取消注册的回调函数
        """
        if handler in self._update_handlers:
            self._update_handlers.remove(handler)
            logger.info("Update handler unregistered")

    def notify_update(self, update_info: Dict[str, Any]) -> None:
        """
        通知所有注册的更新处理器
        
        Args:
            update_info: 更新信息
        """
        for handler in self._update_handlers:
            try:
                handler(update_info)
            except Exception as e:
                logger.error(f"Error in update handler: {e}")

    def handle_document_change(self, change_type: str, changed_object: CDispatch) -> bool:
        """
        处理文档变更事件
        
        Args:
            change_type: 变更类型，如'paragraph_inserted', 'table_updated', 'image_deleted'等
            changed_object: 发生变更的对象（Range、Table、InlineShape等）
        
        Returns:
            处理是否成功
        """
        start_time = time.time()
        success = False

        try:
            # 验证变更类型
            if change_type not in self._supported_change_types:
                logger.warning(f"Unsupported change type: {change_type}")
                return False

            context_manager = get_context_manager()
            
            # 开始事务
            was_in_transaction = context_manager._in_transaction
            if not was_in_transaction:
                context_manager.begin_transaction()

            # 根据变更类型处理
            if change_type == 'paragraph_inserted' or change_type == 'paragraph_updated':
                success = self._update_paragraph_context(changed_object)
            elif change_type == 'table_inserted' or change_type == 'table_updated':
                success = self._update_table_context(changed_object)
            elif change_type == 'image_inserted' or change_type == 'image_updated':
                success = self._update_image_context(changed_object)
            elif change_type == 'paragraph_deleted' or change_type == 'table_deleted' or change_type == 'image_deleted':
                # 确定对象类型并删除
                object_type = 'paragraph' if change_type.startswith('paragraph') else \
                             'table' if change_type.startswith('table') else 'image'
                success = self._remove_object_context(object_type, changed_object)
            elif change_type == 'document_structure_changed':
                # 文档结构发生重大变化，刷新整个上下文树
                success = self._refresh_document_context_tree(changed_object)

            # 如果是新创建的事务，则提交
            if not was_in_transaction:
                context_manager.commit_transaction()

            # 记录性能指标
            context_manager._record_operation_time('handle_change', time.time() - start_time, change_type=change_type, success=success)

            if success:
                # 通知更新处理器
                self.notify_update({
                    'type': 'document_change',
                    'change_type': change_type,
                    'success': success
                })
                
                logger.info(f"Document change handled: {change_type}")
            else:
                logger.warning(f"Failed to handle document change: {change_type}")

            return success
        except Exception as e:
            logger.error(f"Error handling document change {change_type}: {e}")
            logger.error(f"Traceback: {traceback.format_exc()}")
            
            # 获取上下文管理器
            context_manager = get_context_manager()
            
            # 如果是新创建的事务，则回滚
            if not was_in_transaction:
                context_manager.rollback_transaction()

            # 记录性能指标
            context_manager._record_operation_time('handle_change', time.time() - start_time, success=False)

            # 在事务模式下抛出异常
            if context_manager._in_transaction:
                raise DocumentContextError(
                    error_code=ErrorCode.DOCUMENT_CHANGE_HANDLING_FAILED,
                    message=f"Failed to handle document change: {str(e)}"
                )

            return False

    def _update_paragraph_context(self, paragraph: CDispatch) -> bool:
        """
        更新段落上下文
        
        Args:
            paragraph: 段落对象
        
        Returns:
            更新是否成功
        """
        try:
            if not paragraph:
                return False

            context_manager = get_context_manager()
            paragraph_id = f"paragraph_{paragraph.Range.Start}"
            
            # 检查上下文是否已存在
            existing_context = context_manager.find_context_by_id(paragraph_id)
            
            if existing_context:
                # 更新现有上下文
                metadata = existing_context.metadata
                metadata['last_updated'] = time.time()
                metadata['text_preview'] = paragraph.Range.Text[:100] + ("..." if len(paragraph.Range.Text) > 100 else "")
                
                return context_manager.update_context(paragraph_id, {'metadata': metadata})
            else:
                # 创建新的上下文
                section = self._find_section_for_range(paragraph.Range)
                if not section:
                    return False
                
                section_id = f"section_{section.Range.Start}"
                section_context = context_manager.find_context_by_id(section_id)
                
                if not section_context:
                    return False
                
                # 创建段落上下文
                try:
                    text_preview = paragraph.Range.Text[:100] + ("..." if len(paragraph.Range.Text) > 100 else "")
                    style_name = paragraph.Style.NameLocal if hasattr(paragraph.Style, 'NameLocal') else "Normal"
                except Exception:
                    text_preview = ""
                    style_name = "Normal"
                
                paragraph_metadata = {
                    "type": "paragraph",
                    "id": str(paragraph.Range.Start),
                    "text_preview": text_preview,
                    "style": style_name,
                    "created_time": time.time(),
                    "last_updated": time.time()
                }
                
                new_paragraph_context = DocumentContext(
                    title=f"Paragraph {paragraph.Range.Start}",
                    range_obj=paragraph.Range,
                    metadata=paragraph_metadata
                )
                new_paragraph_context.batch_add_objects([paragraph_metadata])
                
                # 添加到父上下文和映射中
                return context_manager.add_context(new_paragraph_context, section_context)
        except Exception as e:
            logger.error(f"Error updating paragraph context: {e}")
            return False

    def _update_table_context(self, table: CDispatch) -> bool:
        """
        更新表格上下文
        
        Args:
            table: 表格对象
        
        Returns:
            更新是否成功
        """
        try:
            if not table:
                return False

            context_manager = get_context_manager()
            table_id = f"table_{table.Range.Start}"
            
            # 检查上下文是否已存在
            existing_context = context_manager.find_context_by_id(table_id)
            
            if existing_context:
                # 更新现有上下文
                metadata = existing_context.metadata
                try:
                    metadata['rows'] = table.Rows.Count
                    metadata['columns'] = table.Columns.Count
                    metadata['cell_count'] = metadata['rows'] * metadata['columns']
                except Exception:
                    pass
                metadata['last_updated'] = time.time()
                
                return context_manager.update_context(table_id, {'metadata': metadata})
            else:
                # 创建新的上下文
                section = self._find_section_for_range(table.Range)
                if not section:
                    return False
                
                section_id = f"section_{section.Range.Start}"
                section_context = context_manager.find_context_by_id(section_id)
                
                if not section_context:
                    return False
                
                # 创建表格上下文
                try:
                    rows = table.Rows.Count
                    cols = table.Columns.Count
                except Exception:
                    rows = 0
                    cols = 0
                
                table_metadata = {
                    "type": "table",
                    "id": str(table.Range.Start),
                    "rows": rows,
                    "columns": cols,
                    "cell_count": rows * cols,
                    "created_time": time.time(),
                    "last_updated": time.time()
                }
                
                new_table_context = DocumentContext(
                    title=f"Table {table.Range.Start}",
                    range_obj=table.Range,
                    metadata=table_metadata
                )
                new_table_context.batch_add_objects([table_metadata])
                
                # 添加到父上下文和映射中
                return context_manager.add_context(new_table_context, section_context)
        except Exception as e:
            logger.error(f"Error updating table context: {e}")
            return False

    def _update_image_context(self, image: CDispatch) -> bool:
        """
        更新图片上下文
        
        Args:
            image: 图片对象
        
        Returns:
            更新是否成功
        """
        try:
            if not image:
                return False

            context_manager = get_context_manager()
            image_id = f"image_{image.Range.Start}"
            
            # 检查上下文是否已存在
            existing_context = context_manager.find_context_by_id(image_id)
            
            if existing_context:
                # 更新现有上下文
                metadata = existing_context.metadata
                try:
                    metadata['width'] = image.Width
                    metadata['height'] = image.Height
                except Exception:
                    pass
                metadata['last_updated'] = time.time()
                
                return context_manager.update_context(image_id, {'metadata': metadata})
            else:
                # 创建新的上下文
                section = self._find_section_for_range(image.Range)
                if not section:
                    return False
                
                section_id = f"section_{section.Range.Start}"
                section_context = context_manager.find_context_by_id(section_id)
                
                if not section_context:
                    return False
                
                # 创建图片上下文
                try:
                    width = image.Width
                    height = image.Height
                except Exception:
                    width = 0
                    height = 0
                
                image_metadata = {
                    "type": "image",
                    "id": str(image.Range.Start),
                    "width": width,
                    "height": height,
                    "created_time": time.time(),
                    "last_updated": time.time()
                }
                
                new_image_context = DocumentContext(
                    title=f"Image {image.Range.Start}",
                    range_obj=image.Range,
                    metadata=image_metadata
                )
                new_image_context.batch_add_objects([image_metadata])
                
                # 添加到父上下文和映射中
                return context_manager.add_context(new_image_context, section_context)
        except Exception as e:
            logger.error(f"Error updating image context: {e}")
            return False

    def _remove_object_context(self, object_type: str, object_range: CDispatch) -> bool:
        """
        移除对象上下文
        
        Args:
            object_type: 对象类型
            object_range: 对象的Range
        
        Returns:
            移除是否成功
        """
        try:
            if not object_range:
                return False

            context_manager = get_context_manager()
            object_id = f"{object_type}_{object_range.Start}"
            
            # 移除上下文
            return context_manager.remove_context(object_id)
        except Exception as e:
            logger.error(f"Error removing {object_type} context: {e}")
            return False

    def _refresh_document_context_tree(self, document: CDispatch) -> bool:
        """
        刷新整个文档上下文树
        
        Args:
            document: 文档对象
        
        Returns:
            刷新是否成功
        """
        try:
            if not document:
                return False

            context_manager = get_context_manager()
            
            # 清除所有现有上下文
            context_manager.clear_all_contexts()
            
            # 创建根上下文
            root_context = DocumentContext.create_root_context(document)
            context_manager.add_context(root_context)
            
            # 遍历文档中的所有节
            for section in document.Sections:
                # 创建节上下文
                try:
                    section_metadata = {
                        "type": "section",
                        "id": str(section.Range.Start),
                        "page_setup": section.PageSetup.Orientation,  # 横版或竖版
                        "created_time": time.time()
                    }
                    
                    section_context = DocumentContext(
                        title=f"Section {section.Index}",
                        range_obj=section.Range,
                        metadata=section_metadata
                    )
                    section_context.batch_add_objects([section_metadata])
                    
                    # 添加到根上下文
                    context_manager.add_context(section_context, root_context)
                except Exception as e:
                    logger.error(f"Error creating section context: {e}")
                    continue
            
            # 批量处理段落、表格和图片
            self._batch_process_document_objects(document)
            
            logger.info("Document context tree refreshed successfully")
            return True
        except Exception as e:
            logger.error(f"Error refreshing document context tree: {e}")
            return False

    def _batch_process_document_objects(self, document: CDispatch) -> None:
        """
        批量处理文档中的所有对象
        
        Args:
            document: 文档对象
        """
        try:
            context_manager = get_context_manager()
            
            # 处理所有段落
            for paragraph in document.Paragraphs:
                self._update_paragraph_context(paragraph)
            
            # 处理所有表格
            for table in document.Tables:
                self._update_table_context(table)
            
            # 处理所有图片
            for image in document.InlineShapes:
                if image.Type == 1:  # wdInlineShapePicture
                    self._update_image_context(image)
        except Exception as e:
            logger.error(f"Error batch processing document objects: {e}")

    def _find_section_for_range(self, range_obj: CDispatch) -> Optional[CDispatch]:
        """
        查找Range所属的节
        
        Args:
            range_obj: Range对象
        
        Returns:
            节对象，如果未找到则返回None
        """
        try:
            if not range_obj or not hasattr(range_obj, 'Document'):
                return None
            
            document = range_obj.Document
            for section in document.Sections:
                if section.Range.Start <= range_obj.Start and section.Range.End >= range_obj.End:
                    return section
            
            # 如果没有完全包含的节，返回第一个节
            if document.Sections.Count > 0:
                return document.Sections(1)
            
            return None
        except Exception:
            return None


# 创建全局文档变更处理器实例
global_change_handler = DocumentChangeHandler()


def get_document_change_handler() -> DocumentChangeHandler:
    """
    获取全局文档变更处理器实例
    """
    return global_change_handler