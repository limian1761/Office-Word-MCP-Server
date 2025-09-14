"""
AppContext for managing the Word application instance and the active document state.
"""

import logging
import os
import shutil
import sys
import traceback
import time
from typing import Optional, Dict, List, Any, Callable, Set, Tuple

import pythoncom
from pythoncom import com_error
from win32com.client.dynamic import CDispatch
from win32com.client import constants as wd_constants

from .errors import ErrorCode, WordDocumentError, DocumentContextError
from ..contexts.context_control import DocumentContext

# Configure logger
logger = logging.getLogger(__name__)


class AppContext:
    """
    Application context that holds the Word application instance and the active document.
    This class acts as a state container. The lifecycle of the Word application
    itself is managed by the server's lifespan manager.
    优化版：增强上下文树管理、批量更新机制和错误处理
    """

    # Singleton instance
    _instance = None
    _lock = None

    def __new__(cls):
        """单例模式实现"""
        if cls._instance is None:
            if cls._lock is None:
                import threading
                cls._lock = threading.Lock()
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super(AppContext, cls).__new__(cls)
                    cls._instance._initialized = False
        return cls._instance

    @classmethod
    def get_instance(cls) -> "AppContext":
        """Get the singleton instance of AppContext, creating it if it doesn't exist"""
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        """
        Initialize the AppContext with a running Word application instance.
        """
        # Prevent duplicate initialization
        if self._initialized:
            return

        # Initialize attributes first
        self._temp_word_app: Optional[CDispatch] = None
        self._active_document: Optional[CDispatch] = None
        self._word_app: Optional[CDispatch] = None
        
        # Document context tree management
        self._document_context_tree: Optional[DocumentContext] = None  # Root of the context tree
        self._context_map: Dict[str, DocumentContext] = {}  # Map of context IDs to context objects
        self._active_context: Optional[DocumentContext] = None  # Currently active context
        self._update_handlers: List[Callable] = []  # List of update handlers for real-time mapping
        
        # 文档操作相关
        self._document_operations_count = 0
        self._last_document_operation_time = None
        
        # 缓存管理
        self._cache_enabled = True
        self._cache_size_limit = 100
        self._cache_hits = 0
        self._cache_misses = 0
        
        # 性能监控
        self._operation_times = {}
        
        # 事务状态
        self._in_transaction = False
        self._transaction_operations = []
        self._transaction_context_backups = {}

        self._initialized = True

    def set_word_app(self, word_app: Optional[CDispatch] = None) -> None:
        """
        Set the Word application instance for the context.

        Args:
            word_app: An active Word application dispatch object.
        """
        self._word_app = word_app

    def _clear_com_cache(self):
        """Clear win32com cache to resolve CLSIDToPackageMap errors"""
        try:
            # 获取gen_py目录路径，使用win32com的内置路径
            import win32com

            gen_path = os.path.join(win32com.__gen_path__, "win32com", "gen_py")
            logger.info(f"Checking win32com cache at: {gen_path}")

            # 检查目录是否存在
            if os.path.exists(gen_path):
                logger.info(f"Clearing win32com cache at: {gen_path}")
                # 尝试删除目录
                shutil.rmtree(gen_path)
                logger.info("Cleared win32com cache successfully")
                return True
            else:
                logger.warning(f"win32com cache directory not found: {gen_path}")
                # 尝试查找其他可能的缓存位置
                temp_gen_path = os.path.join(
                    os.environ.get("TEMP", "C:\\temp"), "gen_py"
                )
                if os.path.exists(temp_gen_path):
                    logger.info(
                        f"Clearing win32com cache at temp location: {temp_gen_path}"
                    )
                    shutil.rmtree(temp_gen_path)
                    logger.info(
                        "Cleared win32com cache from temp location successfully"
                    )
                    return True
                return False
        except PermissionError as e:
            logger.error(f"Permission denied when clearing win32com cache: {e}")
            return False
        except Exception as e:
            logger.error(f"Failed to clear win32com cache: {e}")
            logger.error(f"Traceback: {traceback.format_exc()}")
            return False

    def get_word_app(self, create_if_needed: bool = False) -> Optional[CDispatch]:
        """
        Get the Word application instance, optionally creating it if needed.

        Args:
            create_if_needed: Whether to create a new Word app instance if one doesn't exist.

        Returns:
            The Word application instance or None if not available and not created.
        """
        # Return existing Word app if available and validate it's still functional
        if self._word_app is not None:
            if self._validate_word_app(self._word_app):
                logger.debug("Returning existing valid Word application instance.")
                return self._word_app
            else:
                logger.warning("Existing Word application instance is invalid, will attempt to create a new one.")
                self._word_app = None

        # If we shouldn't create and don't have one, return None
        if not create_if_needed:
            return None

        # Try multiple connection methods with retries
        for attempt in range(3):  # Try up to 3 times
            try:
                logger.info(f"Attempt {attempt+1}/3 to create Word Application instance...")
                
                # 1. Try the standard Dispatch method first
                result = self._create_word_app_with_dispatch()
                if result:
                    return result
                
                # 2. If standard method fails, clear cache and try again
                if attempt == 0:
                    if self._clear_com_cache():
                        logger.info("Retrying after COM cache clear...")
                        result = self._create_word_app_with_dispatch(reload_module=True)
                        if result:
                            return result
                    else:
                        logger.warning("Failed to clear COM cache, moving to next method")
                
                # 3. Try with DispatchEx which creates a separate process
                logger.info("Trying with DispatchEx...")
                result = self._create_word_app_with_dispatchex()
                if result:
                    return result
                
                # 4. Try with early binding
                logger.info("Trying with early binding (gencache)...")
                result = self._create_word_app_with_early_binding()
                if result:
                    return result
                
                # If all methods failed for this attempt, wait before retrying
                if attempt < 2:  # Don't wait after the last attempt
                    import time
                    wait_time = 1.5 ** attempt  # Exponential backoff
                    logger.info(f"All methods failed, waiting {wait_time:.2f}s before retry...")
                    time.sleep(wait_time)
                
            except Exception as e:
                logger.error(f"Unexpected error during attempt {attempt+1}: {e}")
                logger.error(f"Error type: {type(e).__name__}")
                logger.error(f"Traceback: {traceback.format_exc()}")
                
        logger.error("Failed to create Word Application instance after multiple attempts.")
        return None
        
    def _create_word_app_with_dispatch(self, reload_module: bool = False) -> Optional[CDispatch]:
        """Create Word app using standard Dispatch method."""
        try:
            # 确保每次都使用新的win32com.client导入
            import win32com.client
            
            if reload_module:
                import importlib
                importlib.reload(win32com.client)
                logger.info("Reloaded win32com.client module.")
            
            self._word_app = win32com.client.Dispatch("Word.Application")
            logger.info("Successfully created Word application instance with Dispatch.")
            return self._word_app
        except Exception as e:
            logger.warning(f"Dispatch method failed: {e}")
            return None
            
    def _create_word_app_with_dispatchex(self) -> Optional[CDispatch]:
        """Create Word app using DispatchEx method (creates a separate process)."""
        try:
            import win32com.client
            
            self._word_app = win32com.client.DispatchEx("Word.Application")
            logger.info("Successfully created Word application instance with DispatchEx.")
            return self._word_app
        except Exception as e:
            logger.warning(f"DispatchEx method failed: {e}")
            return None
            
    def _create_word_app_with_early_binding(self) -> Optional[CDispatch]:
        """Create Word app using early binding (gencache)."""
        try:
            import win32com.client.gencache
            
            # Ensure we have the gen_py files for Word
            win32com.client.gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 7)
            self._word_app = win32com.client.gencache.Dispatch("Word.Application")
            logger.info("Successfully created Word application instance with early binding.")
            return self._word_app
        except Exception as e:
            logger.warning(f"Early binding method failed: {e}")
            return None
            
    def _validate_word_app(self, word_app: CDispatch) -> bool:
        """Validate that the Word application instance is still functional."""
        if not word_app:
            return False
            
        try:
            # Simple property access to verify the COM object is still alive
            app_version = word_app.Version
            logger.debug(f"Word application validation successful, version: {app_version}")
            return True
        except Exception as e:
            logger.warning(f"Word application validation failed: {e}")
            return False

    def get_active_document(self) -> Optional[CDispatch]:
        """Get the current active document."""
        return self._active_document

    def set_active_document(self, doc: Optional[CDispatch]) -> None:
        """
        Set the current active document.
        """
        self._active_document = doc
        
        # 当设置活动文档后，自动创建文档上下文树
        if doc is not None:
            self.on_document_opened()
        else:
            # 如果清除活动文档，也要清除上下文树
            self._document_context_tree = None
            self._context_map = {}  
            self._active_context = None
            self._update_handlers = []

    def clear_active_document(self) -> None:
        """Clear the current active document."""
        self._active_document = None

    def close_document(self):
        """关闭当前活动文档"""
        try:
            if self._active_document is not None:
                self._active_document.Close(SaveChanges=0)  # 不保存更改
                self._active_document = None
                # 清除上下文树相关信息
                self._document_context_tree = None
                self._context_map = {}
                self._active_context = None
                return True
            return False
        except Exception as e:
            logger.error(f"Error closing document: {e}")
            return False

    def quit_word_app(self):
        """退出Word应用程序"""
        try:
            if self._word_app is not None:
                self._word_app.Quit()
                self._word_app = None
                self._active_document = None
                # 清除上下文树相关信息
                self._document_context_tree = None
                self._context_map = {}
                self._active_context = None
                return True
            return False
        except Exception as e:
            logger.error(f"Error quitting Word application: {e}")
            return False

    def create_document_context_tree(self) -> Optional[DocumentContext]:
        """
        创建当前活动文档的上下文树
        优化版：使用增强的DocumentContext.create_root_context方法，提升创建效率和错误处理
        
        Returns:
            文档上下文树的根节点，如果没有活动文档则返回None
            
        Raises:
            DocumentContextError: 当创建上下文树失败时
        """
        start_time = time.time()
        
        if not self._active_document:
            self._logger.warning("No active document to create context tree")
            return None
        
        try:
            self._logger.info(f"Starting to create context tree for document: {self._active_document.Name}")
            
            # 使用增强版的DocumentContext.create_root_context方法
            document_name = self._active_document.Name
            document_path = getattr(self._active_document, 'FullName', 'Untitled')
            document_metadata = {
                'name': document_name,
                'path': document_path,
                'creation_time': getattr(self._active_document, 'BuiltInDocumentProperties')('Creation Date').Value if hasattr(self._active_document, 'BuiltInDocumentProperties') else None,
                'last_modified_time': getattr(self._active_document, 'BuiltInDocumentProperties')('Last Save Time').Value if hasattr(self._active_document, 'BuiltInDocumentProperties') else None,
                'word_count': getattr(self._active_document, 'BuiltInDocumentProperties')('Word Count').Value if hasattr(self._active_document, 'BuiltInDocumentProperties') else 0,
                'page_count': getattr(self._active_document, 'BuiltInDocumentProperties')('Number of Pages').Value if hasattr(self._active_document, 'BuiltInDocumentProperties') else 0
            }
            
            # 创建根上下文节点
            self._document_context_tree = DocumentContext.create_root_context(
                title=f"Document: {document_name}",
                range_obj=self._active_document.Content if hasattr(self._active_document, 'Content') else None,
                metadata=document_metadata
            )
            
            # 添加根上下文到映射中
            self._context_map[self._document_context_tree.context_id] = self._document_context_tree
            
            # 批量构建文档结构的上下文树
            self._build_document_structure_optimized(self._document_context_tree)
            
            self._logger.info(f"Successfully created document context tree with {len(self._context_map)} contexts for {document_name}")
            
            # 记录性能指标
            self._record_operation_time('create_document_context_tree', time.time() - start_time)
            return self._document_context_tree
        except Exception as e:
            self._logger.error(f"Failed to create document context tree: {e}")
            self._logger.error(f"Traceback: {traceback.format_exc()}")
            raise DocumentContextError(
                error_code=ErrorCode.DOCUMENT_CONTEXT_CREATION_FAILED,
                message=f"Failed to create document context tree: {str(e)}"
            )
    
    def _build_document_structure_optimized(self, root_context: DocumentContext) -> None:
        """
        构建文档结构的上下文树（优化版）
        使用批量添加和异步处理提高性能
        
        Args:
            root_context: 根上下文节点
        """
        if not self._active_document:
            return
        
        try:
            # 收集所有节信息用于批量处理
            sections = self._active_document.Sections
            section_contexts = []
            
            # 预处理所有节
            for i, section in enumerate(sections):
                section_metadata = {
                    "type": "section",
                    "id": str(section.Range.Start),
                    "index": i,
                    "range_start": section.Range.Start,
                    "range_end": section.Range.End,
                    "page_setup": {
                        "orientation": str(section.PageSetup.Orientation),
                        "paper_size": str(section.PageSetup.PaperSize),
                        "top_margin": section.PageSetup.TopMargin,
                        "bottom_margin": section.PageSetup.BottomMargin,
                        "left_margin": section.PageSetup.LeftMargin,
                        "right_margin": section.PageSetup.RightMargin
                    }
                }
                
                # 创建节上下文
                section_context = DocumentContext(
                    title=f"Section {i+1}",
                    range_obj=section.Range,
                    metadata=section_metadata
                )
                section_contexts.append((section_context, section))
                
                # 添加节上下文到映射中
                self._context_map[section_context.context_id] = section_context
            
            # 使用批量添加方法添加所有节上下文
            root_context.batch_add_child_contexts([ctx for ctx, _ in section_contexts])
            
            # 为每个节构建内容
            for section_context, section in section_contexts:
                self._build_section_content_optimized(section_context, section)
                
        except Exception as e:
            self._logger.error(f"Failed to build document structure: {e}")
    
    def _build_section_content_optimized(self, parent_context: DocumentContext, section: CDispatch) -> None:
        """
        构建节内容的上下文树（优化版）
        批量处理段落、表格和图片，减少重复操作
        
        Args:
            parent_context: 父上下文节点
            section: Word节对象
        """
        start_time = time.time()
        
        try:
            # 获取节内的所有对象范围
            range_start = section.Range.Start
            range_end = section.Range.End
            
            # 预收集所有需要处理的对象
            child_contexts = []
            objects_to_add = []
            
            # 1. 处理表格
            for table in self._active_document.Tables:
                if range_start <= table.Range.Start and table.Range.End <= range_end:
                    table_metadata = {
                        "type": "table",
                        "id": str(table.Range.Start),
                        "rows": table.Rows.Count,
                        "columns": table.Columns.Count,
                        "cell_count": table.Rows.Count * table.Columns.Count
                    }
                    
                    table_context = DocumentContext(
                        title=f"Table at {table.Range.Start}",
                        range_obj=table.Range,
                        metadata=table_metadata
                    )
                    
                    # 添加到批量处理列表
                    child_contexts.append(table_context)
                    self._context_map[table_context.context_id] = table_context
                    objects_to_add.append((table_context, table_metadata))
            
            # 2. 处理图片
            for i in range(1, self._active_document.InlineShapes.Count + 1):
                try:
                    shape = self._active_document.InlineShapes(i)
                    if range_start <= shape.Range.Start and shape.Range.End <= range_end:
                        # 检查是否已作为表格的一部分处理
                        is_processed = False
                        for ctx in child_contexts:
                            if ctx.range and shape.Range.Start >= ctx.range.Start and shape.Range.End <= ctx.range.End:
                                is_processed = True
                                break
                        
                        if not is_processed:
                            image_metadata = {
                                "type": "image",
                                "id": str(shape.Range.Start),
                                "width": shape.Width,
                                "height": shape.Height,
                                "type": str(getattr(shape, 'Type', 'Unknown'))
                            }
                            
                            image_context = DocumentContext(
                                title=f"Image at {shape.Range.Start}",
                                range_obj=shape.Range,
                                metadata=image_metadata
                            )
                            
                            # 添加到批量处理列表
                            child_contexts.append(image_context)
                            self._context_map[image_context.context_id] = image_context
                            objects_to_add.append((image_context, image_metadata))
                except Exception:
                    # 忽略无法访问的图片
                    continue
            
            # 3. 处理段落（排除已处理的表格和图片中的段落）
            processed_ranges = [(ctx.range.Start, ctx.range.End) for ctx in child_contexts if ctx.range]
            
            for i in range(1, self._active_document.Paragraphs.Count + 1):
                try:
                    paragraph = self._active_document.Paragraphs(i)
                    if range_start <= paragraph.Range.Start and paragraph.Range.End <= range_end:
                        # 检查是否已被处理
                        is_processed = False
                        for start, end in processed_ranges:
                            if paragraph.Range.Start >= start and paragraph.Range.End <= end:
                                is_processed = True
                                break
                        
                        if not is_processed:
                            # 只处理非空段落或包含重要内容的段落
                            if paragraph.Range.Text.strip():
                                text_preview = paragraph.Range.Text[:30] + ("..." if len(paragraph.Range.Text) > 30 else "")
                                para_metadata = {
                                    "type": "paragraph",
                                    "id": str(paragraph.Range.Start),
                                    "text_preview": text_preview,
                                    "style_name": getattr(paragraph, 'Style', '').Name if hasattr(getattr(paragraph, 'Style', ''), 'Name') else 'Normal',
                                    "is_heading": getattr(paragraph, 'Style', '').Name.startswith('Heading') if hasattr(getattr(paragraph, 'Style', ''), 'Name') else False
                                }
                                
                                para_context = DocumentContext(
                                    title=f"Paragraph at {paragraph.Range.Start}",
                                    range_obj=paragraph.Range,
                                    metadata=para_metadata
                                )
                                
                                # 添加到批量处理列表
                                child_contexts.append(para_context)
                                self._context_map[para_context.context_id] = para_context
                                objects_to_add.append((para_context, para_metadata))
                except Exception:
                    # 忽略无法访问的段落
                    continue
            
            # 批量添加子上下文
            parent_context.batch_add_child_contexts(child_contexts)
            
            # 批量添加对象
            for context, obj_metadata in objects_to_add:
                context.batch_add_objects([obj_metadata])
            
            # 记录性能指标
            self._record_operation_time('build_section_content', time.time() - start_time)
            
        except Exception as e:
            self._logger.error(f"Failed to build section content: {e}")
            self._logger.error(f"Traceback: {traceback.format_exc()}")
    
    def get_document_context_tree(self) -> Optional[DocumentContext]:
        """
        获取文档的上下文树
        
        返回:
            文档上下文树的根节点，如果没有则返回None
        """
        return self._document_context_tree
    
    def get_active_context(self) -> Optional[DocumentContext]:
        """
        获取当前活动上下文
        
        返回:
            当前活动上下文对象，如果没有则返回None
        """
        return self._active_context
    
    def set_active_context(self, context: Optional[DocumentContext]) -> None:
        """
        设置当前活动上下文
        
        参数:
            context: 要设置为活动的上下文对象
        """
        self._active_context = context
        
        # 如果设置了活动上下文且有Range对象，可以滚动到该位置
        if context and context.range:
            try:
                word_app = self.get_word_app()
                if word_app:
                    context.range.Select()
                    word_app.ActiveWindow.ScrollIntoView(context.range)
            except Exception as e:
                logger.error(f"Failed to select context range: {e}")
    
    def get_context_by_id(self, context_id: str) -> Optional[DocumentContext]:
        """
        通过ID获取上下文对象
        
        参数:
            context_id: 上下文ID
        
        返回:
            上下文对象，如果未找到则返回None
        """
        return self._context_map.get(context_id)
    
    def add_context_to_tree(self, context: DocumentContext, parent_context: Optional[DocumentContext] = None) -> bool:
        """
        向上下文树添加新的上下文（优化版）
        使用增强的DocumentContext功能和事务支持
        
        Args:
            context: 要添加的上下文对象
            parent_context: 父上下文对象，如果为None则添加到根节点
        
        Returns:
            添加是否成功
        
        Raises:
            DocumentContextError: 当添加上下文失败时（在事务模式下）
        """
        try:
            # 在事务模式下记录操作以便回滚
            if self._in_transaction:
                self._transaction_operations.append({
                    'type': 'add_context',
                    'context_id': context.context_id,
                    'parent_context_id': parent_context.context_id if parent_context else None
                })
                
                # 备份当前状态
                if context.context_id in self._context_map:
                    self._transaction_context_backups[context.context_id] = {
                        'context': self._context_map[context.context_id],
                        'parent': self._context_map[context.context_id].parent_context
                    }
            
            # 如果没有指定父上下文，则添加到根节点
            if parent_context is None:
                if self._document_context_tree:
                    self._document_context_tree.add_child_context(context)
                else:
                    # 如果没有根节点，则将此上下文设为根节点
                    self._document_context_tree = context
            else:
                parent_context.add_child_context(context)
            
            # 添加到映射中
            self._context_map[context.context_id] = context
            
            # 更新上下文的metadata
            context._update_metadata({
                'added_time': time.time(),
                'is_active': context == self._active_context
            })
            
            self._logger.info(f"Added context '{context.title}' (ID: {context.context_id}) to context tree")
            
            # 通知更新
            self.notify_update("context_added", context_id=context.context_id, parent_id=parent_context.context_id if parent_context else None)
            
            return True
        except Exception as e:
            self._logger.error(f"Failed to add context to tree: {e}")
            self._logger.error(f"Traceback: {traceback.format_exc()}")
            
            # 在事务模式下抛出异常以便回滚
            if self._in_transaction:
                raise DocumentContextError(
                    error_code=ErrorCode.CONTEXT_ADD_FAILED,
                    message=f"Failed to add context '{context.title}': {str(e)}"
                )
            
            return False
    
    def remove_context_from_tree(self, context_id: str) -> bool:
        """
        从上下文树中移除指定的上下文（增强版）
        支持事务处理、错误恢复和性能监控
        
        Args:
            context_id: 要移除的上下文ID
        
        Returns:
            移除是否成功
        
        Raises:
            DocumentContextError: 当移除上下文失败时（在事务模式下）
        """
        start_time = time.time()
        
        try:
            if context_id not in self._context_map:
                logger.warning(f"Context with ID {context_id} not found")
                return False
            
            context = self._context_map[context_id]
            parent_context = context.parent_context
            
            # 在事务模式下记录操作以便回滚
            if self._in_transaction:
                self._transaction_operations.append({
                    'type': 'remove_context',
                    'context_id': context_id,
                    'parent_context_id': parent_context.context_id if parent_context else None,
                    'context': context,
                    'children': context.child_contexts.copy()
                })
            
            # 递归移除所有子上下文
            for child in list(context.child_contexts):
                self.remove_context_from_tree(child.context_id)
            
            # 如果上下文有父上下文，则从父上下文中移除
            if parent_context:
                parent_context.remove_child_context(context)
            elif context == self._document_context_tree:
                # 如果是根节点，则清除上下文树
                self._document_context_tree = None
            
            # 从映射中移除
            del self._context_map[context_id]
            
            # 如果当前活动上下文是被移除的上下文，则清除活动上下文
            if self._active_context == context:
                self._active_context = None
            
            # 通知更新
            self.notify_update("context_removed", context_id=context_id)
            
            # 记录性能指标
            self._record_operation_time('remove_context', time.time() - start_time)
            
            logger.info(f"Removed context '{context.title}' (ID: {context_id}) from context tree")
            return True
        except Exception as e:
            logger.error(f"Failed to remove context from tree: {e}")
            logger.error(f"Traceback: {traceback.format_exc()}")
            
            # 记录性能指标
            self._record_operation_time('remove_context', time.time() - start_time, success=False)
            
            # 在事务模式下抛出异常以便回滚
            if self._in_transaction:
                raise DocumentContextError(
                    error_code=ErrorCode.CONTEXT_REMOVE_FAILED,
                    message=f"Failed to remove context '{context_id}': {str(e)}"
                )
            
            return False
    
    def on_document_opened(self) -> None:
        """
        在文档打开或新建后被调用，创建并初始化文档上下文树
        """
        try:
            logger.info("Initializing document context tree after document opened")
            
            # 清除旧的上下文树信息
            self._document_context_tree = None
            self._context_map = {}  
            self._active_context = None
            self._update_handlers = []
            
            # 创建新的上下文树
            self.create_document_context_tree()
            
            logger.info("Document context tree initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize document context tree: {e}")
            logger.error(f"Traceback: {traceback.format_exc()}")

    def get_context_tree_as_dict(self) -> Dict[str, Any]:
        """
        将上下文树转换为字典格式，便于序列化
        
        返回:
            上下文树的字典表示
        """
        if not self._document_context_tree:
            return {
                "success": False,
                "message": "No context tree available"
            }
        
        try:
            def context_to_dict(context):
                result = context.to_dict()
                result["children"] = [context_to_dict(child) for child in context.child_contexts]
                return result
            
            return {
                "success": True,
                "root_context": context_to_dict(self._document_context_tree),
                "context_count": len(self._context_map),
                "has_active_context": self._active_context is not None
            }
        except Exception as e:
            logger.error(f"Failed to convert context tree to dict: {e}")
            return {
                "success": False,
                "message": str(e)
            }
    
    def refresh_document_context_tree(self) -> Optional[DocumentContext]:
        """
        刷新文档上下文树，重新构建整个树结构
        
        返回:
            刷新后的文档上下文树的根节点，如果没有活动文档则返回None
        """
        return self.create_document_context_tree()
    
    def register_update_handler(self, handler: Callable) -> None:
        """
        注册上下文更新处理器，用于实时监听文档变化
        
        参数:
            handler: 处理更新的回调函数
        """
        if handler not in self._update_handlers:
            self._update_handlers.append(handler)
            logger.debug(f"Registered update handler: {handler.__name__}")
    
    def unregister_update_handler(self, handler: Callable) -> None:
        """
        注销上下文更新处理器
        
        参数:
            handler: 要注销的回调函数
        """
        if handler in self._update_handlers:
            self._update_handlers.remove(handler)
            logger.debug(f"Unregistered update handler: {handler.__name__}")
    
    def notify_update(self, update_type: str, **kwargs) -> None:
        """
        通知所有注册的处理器有更新发生
        
        参数:
            update_type: 更新类型
            **kwargs: 额外的更新信息
        """
        for handler in self._update_handlers:
            try:
                handler(update_type=update_type, **kwargs)
            except Exception as e:
                logger.error(f"Error in update handler {handler.__name__}: {e}")
    
    def update_paragraph_context(self, paragraph_range: CDispatch) -> bool:
        """
        更新段落上下文（增强版）
        支持事务处理、错误恢复和性能监控
        
        Args:
            paragraph_range: 段落的Range对象
        
        Returns:
            更新是否成功
        
        Raises:
            DocumentContextError: 当更新上下文失败时（在事务模式下）
        """
        start_time = time.time()
        
        if not self._active_document or not paragraph_range:
            return False
        
        try:
            # 查找与该段落相关的上下文
            para_id = f"paragraph_{paragraph_range.Start}"
            paragraph_context = self._context_map.get(para_id)
            
            # 提取段落样式和格式化信息
            text_preview = paragraph_range.Text[:30] + ("..." if len(paragraph_range.Text) > 30 else "")
            
            try:
                style_name = getattr(paragraph_range.ParagraphFormat.Style, 'Name', 'Normal')
                is_heading = style_name.startswith('Heading')
                formatting = {
                    'alignment': str(getattr(paragraph_range.ParagraphFormat, 'Alignment', 0)),
                    'space_before': getattr(paragraph_range.ParagraphFormat, 'SpaceBefore', 0),
                    'space_after': getattr(paragraph_range.ParagraphFormat, 'SpaceAfter', 0)
                }
            except Exception:
                style_name = 'Normal'
                is_heading = False
                formatting = {}
            
            para_metadata = {
                "type": "paragraph",
                "id": str(paragraph_range.Start),
                "text_preview": text_preview,
                "style_name": style_name,
                "is_heading": is_heading,
                "formatting": formatting,
                "last_updated": time.time()
            }
            
            if paragraph_context:
                # 在事务模式下记录操作以便回滚
                if self._in_transaction:
                    self._transaction_operations.append({
                        'type': 'update_paragraph',
                        'context_id': para_id,
                        'old_state': {
                            'title': paragraph_context.title,
                            'range': paragraph_context.range,
                            'metadata': paragraph_context.metadata.copy()
                        }
                    })
                
                # 更新上下文信息
                paragraph_context.title = f"Paragraph {paragraph_range.Start}"
                paragraph_context.range = paragraph_range
                
                # 更新对象信息和元数据
                paragraph_context.batch_add_objects([para_metadata])
                paragraph_context._update_metadata(para_metadata)
                
                logger.debug(f"Updated paragraph context: {para_id}")
            else:
                # 如果找不到对应的上下文，尝试为其创建新的上下文
                section = self._find_section_for_range(paragraph_range)
                if section:
                    section_id = f"section_{section.Range.Start}"
                    section_context = self._context_map.get(section_id)
                    
                    if section_context:
                        # 创建新的段落上下文
                        new_para_context = DocumentContext(
                            title=f"Paragraph {paragraph_range.Start}",
                            range_obj=paragraph_range,
                            metadata=para_metadata
                        )
                        new_para_context.batch_add_objects([para_metadata])
                        
                        # 添加到父上下文和映射中
                        self.add_context_to_tree(new_para_context, section_context)
                        
                        logger.debug(f"Created new paragraph context: {para_id}")
            
            # 通知更新
            self.notify_update("paragraph_updated", paragraph_id=para_id, metadata=para_metadata)
            
            # 记录性能指标
            self._record_operation_time('update_paragraph', time.time() - start_time)
            
            return True
        except Exception as e:
            logger.error(f"Failed to update paragraph context: {e}")
            logger.error(f"Traceback: {traceback.format_exc()}")
            
            # 记录性能指标
            self._record_operation_time('update_paragraph', time.time() - start_time, success=False)
            
            # 在事务模式下抛出异常以便回滚
            if self._in_transaction:
                raise DocumentContextError(
                    error_code=ErrorCode.PARAGRAPH_UPDATE_FAILED,
                    message=f"Failed to update paragraph context: {str(e)}"
                )
            
            return False
    
    def update_table_context(self, table: CDispatch) -> bool:
        """
        更新表格上下文（增强版）
        支持事务处理、错误恢复和性能监控
        
        Args:
            table: Word表格对象
        
        Returns:
            更新是否成功
        
        Raises:
            DocumentContextError: 当更新上下文失败时（在事务模式下）
        """
        start_time = time.time()
        
        if not self._active_document or not table:
            return False
        
        try:
            # 查找与该表格相关的上下文
            table_id = f"table_{table.Range.Start}"
            table_context = self._context_map.get(table_id)
            
            # 收集表格信息
            try:
                rows = table.Rows.Count
                cols = table.Columns.Count
                # 尝试获取表格样式信息
                try:
                    style_name = getattr(table, 'Style', None)
                    if style_name:
                        style_name = str(style_name)
                except Exception:
                    style_name = None
            except Exception:
                rows = 0
                cols = 0
                style_name = None
            
            table_metadata = {
                "type": "table",
                "id": str(table.Range.Start),
                "rows": rows,
                "columns": cols,
                "cell_count": rows * cols,
                "style_name": style_name,
                "last_updated": time.time()
            }
            
            if table_context:
                # 在事务模式下记录操作以便回滚
                if self._in_transaction:
                    self._transaction_operations.append({
                        'type': 'update_table',
                        'context_id': table_id,
                        'old_state': {
                            'title': table_context.title,
                            'range': table_context.range,
                            'metadata': table_context.metadata.copy()
                        }
                    })
                
                # 更新上下文信息
                table_context.range = table.Range
                
                # 更新对象信息和元数据
                table_context.batch_add_objects([table_metadata])
                table_context._update_metadata(table_metadata)
                
                logger.debug(f"Updated table context: {table_id}")
            else:
                # 如果找不到对应的上下文，尝试为其创建新的上下文
                section = self._find_section_for_range(table.Range)
                if section:
                    section_id = f"section_{section.Range.Start}"
                    section_context = self._context_map.get(section_id)
                    
                    if section_context:
                        # 创建新的表格上下文
                        new_table_context = DocumentContext(
                            title=f"Table {table.Range.Start}",
                            range_obj=table.Range,
                            metadata=table_metadata
                        )
                        new_table_context.batch_add_objects([table_metadata])
                        
                        # 添加到父上下文和映射中
                        self.add_context_to_tree(new_table_context, section_context)
                        
                        logger.debug(f"Created new table context: {table_id}")
            
            # 通知更新
            self.notify_update("table_updated", table_id=table_id, metadata=table_metadata)
            
            # 记录性能指标
            self._record_operation_time('update_table', time.time() - start_time)
            
            return True
        except Exception as e:
            logger.error(f"Failed to update table context: {e}")
            logger.error(f"Traceback: {traceback.format_exc()}")
            
            # 记录性能指标
            self._record_operation_time('update_table', time.time() - start_time, success=False)
            
            # 在事务模式下抛出异常以便回滚
            if self._in_transaction:
                raise DocumentContextError(
                    error_code=ErrorCode.TABLE_UPDATE_FAILED,
                    message=f"Failed to update table context: {str(e)}"
                )
            
            return False
    
    def update_image_context(self, image: CDispatch) -> bool:
        """
        更新图片上下文（增强版）
        支持事务处理、错误恢复和性能监控
        
        Args:
            image: Word图片对象
        
        Returns:
            更新是否成功
        
        Raises:
            DocumentContextError: 当更新上下文失败时（在事务模式下）
        """
        start_time = time.time()
        
        if not self._active_document or not image:
            return False
        
        try:
            # 查找与该图片相关的上下文
            image_id = f"image_{image.Range.Start}"
            image_context = self._context_map.get(image_id)
            
            # 收集图片信息
            try:
                width = image.Width
                height = image.Height
                # 尝试获取更多图片属性
                try:
                    shape_type = str(getattr(image, 'Type', 'Unknown'))
                    wrap_format = str(getattr(image, 'WrapFormat', 'Unknown'))
                except Exception:
                    shape_type = 'Unknown'
                    wrap_format = 'Unknown'
            except Exception:
                width = 0
                height = 0
                shape_type = 'Unknown'
                wrap_format = 'Unknown'
            
            image_metadata = {
                "type": "image",
                "id": str(image.Range.Start),
                "width": width,
                "height": height,
                "type": shape_type,
                "wrap_format": wrap_format,
                "last_updated": time.time()
            }
            
            if image_context:
                # 在事务模式下记录操作以便回滚
                if self._in_transaction:
                    self._transaction_operations.append({
                        'type': 'update_image',
                        'context_id': image_id,
                        'old_state': {
                            'title': image_context.title,
                            'range': image_context.range,
                            'metadata': image_context.metadata.copy()
                        }
                    })
                
                # 更新上下文信息
                image_context.range = image.Range
                
                # 更新对象信息和元数据
                image_context.batch_add_objects([image_metadata])
                image_context._update_metadata(image_metadata)
                
                logger.debug(f"Updated image context: {image_id}")
            else:
                # 如果找不到对应的上下文，尝试为其创建新的上下文
                section = self._find_section_for_range(image.Range)
                if section:
                    section_id = f"section_{section.Range.Start}"
                    section_context = self._context_map.get(section_id)
                    
                    if section_context:
                        # 创建新的图片上下文
                        new_image_context = DocumentContext(
                            title=f"Image {image.Range.Start}",
                            range_obj=image.Range,
                            metadata=image_metadata
                        )
                        new_image_context.batch_add_objects([image_metadata])
                        
                        # 添加到父上下文和映射中
                        self.add_context_to_tree(new_image_context, section_context)
                        
                        logger.debug(f"Created new image context: {image_id}")
            
            # 通知更新
            self.notify_update("image_updated", image_id=image_id, metadata=image_metadata)
            
            # 记录性能指标
            self._record_operation_time('update_image', time.time() - start_time)
            
            return True
        except Exception as e:
            logger.error(f"Failed to update image context: {e}")
            logger.error(f"Traceback: {traceback.format_exc()}")
            
            # 记录性能指标
            self._record_operation_time('update_image', time.time() - start_time, success=False)
            
            # 在事务模式下抛出异常以便回滚
            if self._in_transaction:
                raise DocumentContextError(
                    error_code=ErrorCode.IMAGE_UPDATE_FAILED,
                    message=f"Failed to update image context: {str(e)}"
                )
            
            return False
    
    def remove_object_context(self, object_type: str, object_range: CDispatch) -> bool:
        """
        移除对象的上下文
        
        参数:
            object_type: 对象类型 (paragraph, table, image等)
            object_range: 对象的Range对象
        
        返回:
            移除是否成功
        """
        if not self._active_document or not object_range:
            return False
        
        try:
            # 构建对象ID
            object_id = f"{object_type}_{object_range.Start}"
            
            # 从上下文树中移除
            result = self.remove_context_from_tree(object_id)
            
            if result:
                # 通知更新
                self.notify_update("object_removed", object_id=object_id, object_type=object_type)
                logger.debug(f"Removed {object_type} context: {object_id}")
            
            return result
        except Exception as e:
            logger.error(f"Failed to remove {object_type} context: {e}")
            return False
    
    def _find_section_for_range(self, range_obj: CDispatch) -> Optional[CDispatch]:
        """
        查找给定Range所在的节
        
        参数:
            range_obj: Range对象
        
        返回:
            节对象，如果未找到则返回None
        """
        if not self._active_document:
            return None
        
        try:
            for section in self._active_document.Sections:
                if section.Range.Start <= range_obj.Start and range_obj.End <= section.Range.End:
                    return section
            return None
        except Exception as e:
            logger.error(f"Failed to find section for range: {e}")
            return None
    
    def batch_update_contexts(self, update_operations: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        批量更新上下文（增强版）
        支持事务处理、错误恢复和性能监控
        
        Args:
            update_operations: 更新操作列表，每个操作包含type和必要的参数
        
        Returns:
            包含操作结果的详细字典
        
        Raises:
            DocumentContextError: 当批量更新在事务模式下失败时
        """
        start_time = time.time()
        
        results = {
            "success": True,
            "updated": 0,
            "failed": 0,
            "errors": [],
            "operation_details": [],
            "total_operations": len(update_operations),
            "transaction_id": None
        }
        
        # 如果不在事务模式中，创建一个新的事务
        was_in_transaction = self._in_transaction
        if not was_in_transaction:
            self.begin_transaction()
            results["transaction_id"] = self._current_transaction_id
        
        try:
            for op_index, op in enumerate(update_operations):
                op_type = op.get("type")
                op_details = {
                    "index": op_index,
                    "type": op_type,
                    "success": False,
                    "error": None
                }
                
                try:
                    if op_type == "update_paragraph":
                        success = self.update_paragraph_context(op.get("range"))
                    elif op_type == "update_table":
                        success = self.update_table_context(op.get("table"))
                    elif op_type == "update_image":
                        success = self.update_image_context(op.get("image"))
                    elif op_type == "remove_object":
                        success = self.remove_object_context(op.get("object_type"), op.get("range"))
                    elif op_type == "add_paragraph":
                        # 支持批量添加段落
                        success = self._add_paragraph_context_in_batch(op.get("range"), op.get("section"))
                    elif op_type == "add_table":
                        # 支持批量添加表格
                        success = self._add_table_context_in_batch(op.get("table"), op.get("section"))
                    elif op_type == "add_image":
                        # 支持批量添加图片
                        success = self._add_image_context_in_batch(op.get("image"), op.get("section"))
                    else:
                        raise ValueError(f"Unknown update operation type: {op_type}")
                    
                    if success:
                        results["updated"] += 1
                        op_details["success"] = True
                    else:
                        results["failed"] += 1
                        error_msg = f"Failed to perform {op_type}"
                        results["errors"].append(error_msg)
                        op_details["error"] = error_msg
                        results["success"] = False
                except Exception as e:
                    results["failed"] += 1
                    error_msg = str(e)
                    results["errors"].append(error_msg)
                    op_details["error"] = error_msg
                    results["success"] = False
                    logger.error(f"Error in batch update operation {op_type} at index {op_index}: {e}")
                    logger.error(f"Traceback: {traceback.format_exc()}")
                
                results["operation_details"].append(op_details)
            
            # 如果是新创建的事务，则提交
            if not was_in_transaction:
                self.commit_transaction()
            
            # 记录性能指标
            self._record_operation_time('batch_update', time.time() - start_time, operations_count=len(update_operations))
            
            logger.info(f"Batch update completed: {results['updated']} succeeded, {results['failed']} failed")
            
            return results
        except Exception as e:
            logger.error(f"Batch update failed: {e}")
            logger.error(f"Traceback: {traceback.format_exc()}")
            
            # 如果是新创建的事务，则回滚
            if not was_in_transaction:
                self.rollback_transaction()
            
            # 记录性能指标
            self._record_operation_time('batch_update', time.time() - start_time, success=False)
            
            # 在事务模式下抛出异常
            if self._in_transaction:
                raise DocumentContextError(
                    error_code=ErrorCode.BATCH_UPDATE_FAILED,
                    message=f"Failed to perform batch update: {str(e)}"
                )
            
            results["success"] = False
            results["errors"].append(f"Batch update failed: {str(e)}")
            
            return results
    
    def _add_paragraph_context_in_batch(self, paragraph_range: CDispatch, section: Optional[CDispatch] = None) -> bool:
        """
        在批量操作中添加段落上下文
        
        Args:
            paragraph_range: 段落Range对象
            section: 所属节对象
        
        Returns:
            添加是否成功
        """
        if not paragraph_range:
            return False
        
        try:
            # 如果没有提供节，则查找
            if not section:
                section = self._find_section_for_range(paragraph_range)
                if not section:
                    return False
            
            section_id = f"section_{section.Range.Start}"
            section_context = self._context_map.get(section_id)
            
            if section_context:
                para_id = f"paragraph_{paragraph_range.Start}"
                
                # 如果已存在则直接更新
                if para_id in self._context_map:
                    return self.update_paragraph_context(paragraph_range)
                
                # 创建新的段落上下文
                text_preview = paragraph_range.Text[:30] + ("..." if len(paragraph_range.Text) > 30 else "")
                
                try:
                    style_name = getattr(paragraph_range.ParagraphFormat.Style, 'Name', 'Normal')
                    is_heading = style_name.startswith('Heading')
                except Exception:
                    style_name = 'Normal'
                    is_heading = False
                
                para_metadata = {
                    "type": "paragraph",
                    "id": str(paragraph_range.Start),
                    "text_preview": text_preview,
                    "style_name": style_name,
                    "is_heading": is_heading,
                    "created_time": time.time()
                }
                
                new_para_context = DocumentContext(
                    title=f"Paragraph {paragraph_range.Start}",
                    range_obj=paragraph_range,
                    metadata=para_metadata
                )
                new_para_context.batch_add_objects([para_metadata])
                
                # 添加到父上下文和映射中
                self.add_context_to_tree(new_para_context, section_context)
                
                return True
            
            return False
        except Exception:
            return False

    def search_contexts_by_type(self, context_type: str, max_results: int = 100) -> List[Dict[str, Any]]:
        """
        按类型搜索文档中的上下文
        
        Args:
            context_type: 要搜索的上下文类型，如'paragraph', 'table', 'image', 'section'
            max_results: 最大返回结果数
        
        Returns:
            匹配的上下文列表，每个上下文以字典形式返回
        """
        start_time = time.time()
        results = []
        
        try:
            # 遍历上下文映射，查找匹配类型的上下文
            for context_id, context in self._context_map.items():
                if len(results) >= max_results:
                    break
                
                if context.metadata.get('type') == context_type:
                    # 转换为字典格式并添加到结果列表
                    results.append(self._context_to_dict(context))
            
            # 记录性能指标
            self._record_operation_time('search_contexts', time.time() - start_time, results_count=len(results))
            
            logger.info(f"Context search completed: found {len(results)} {context_type}(s)")
            return results
        except Exception as e:
            logger.error(f"Error searching contexts by type: {e}")
            logger.error(f"Traceback: {traceback.format_exc()}")
            
            # 记录性能指标
            self._record_operation_time('search_contexts', time.time() - start_time, success=False)
            
            return []
    
    def get_context_hierarchy(self, context_id: str) -> Optional[Dict[str, Any]]:
        """
        获取指定上下文的层次结构
        
        Args:
            context_id: 上下文ID
        
        Returns:
            包含层次结构信息的字典，或None（如果上下文不存在）
        """
        start_time = time.time()
        
        try:
            # 检查上下文是否存在
            context = self._context_map.get(context_id)
            if not context:
                logger.warning(f"Context not found: {context_id}")
                return None
            
            # 构建层次结构
            hierarchy = {
                'current': self._context_to_dict(context),
                'parent': None,
                'children': []
            }
            
            # 获取父上下文信息
            if context.parent_context:
                hierarchy['parent'] = {
                    'id': context.parent_context.context_id,
                    'title': context.parent_context.title,
                    'type': context.parent_context.metadata.get('type')
                }
            
            # 获取子上下文信息
            for child in context.child_contexts:
                hierarchy['children'].append({
                    'id': child.context_id,
                    'title': child.title,
                    'type': child.metadata.get('type')
                })
            
            # 记录性能指标
            self._record_operation_time('get_hierarchy', time.time() - start_time)
            
            return hierarchy
        except Exception as e:
            logger.error(f"Error getting context hierarchy: {e}")
            logger.error(f"Traceback: {traceback.format_exc()}")
            
            # 记录性能指标
            self._record_operation_time('get_hierarchy', time.time() - start_time, success=False)
            
            return None
    
    def _context_to_dict(self, context: DocumentContext) -> Dict[str, Any]:
        """
        将上下文对象转换为字典格式
        
        Args:
            context: DocumentContext对象
        
        Returns:
            上下文的字典表示
        """
        try:
            context_dict = {
                'id': context.context_id,
                'title': context.title,
                'type': context.metadata.get('type'),
                'metadata': context.metadata.copy(),
                'has_children': len(context.child_contexts) > 0
            }
            
            # 添加位置信息（如果可用）
            if hasattr(context, 'range_obj') and context.range_obj:
                try:
                    context_dict['start'] = context.range_obj.Start
                    context_dict['end'] = context.range_obj.End
                except Exception:
                    pass
            
            return context_dict
        except Exception as e:
            logger.error(f"Error converting context to dict: {e}")
            return {
                'id': context.context_id if hasattr(context, 'context_id') else 'unknown',
                'title': 'Error converting context',
                'type': 'error',
                'error': str(e)
            }
    
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
            # 开始事务
            was_in_transaction = self._in_transaction
            if not was_in_transaction:
                self.begin_transaction()
            
            # 根据变更类型处理
            if change_type == 'paragraph_inserted' or change_type == 'paragraph_updated':
                success = self.update_paragraph_context(changed_object)
            elif change_type == 'table_inserted' or change_type == 'table_updated':
                success = self.update_table_context(changed_object)
            elif change_type == 'image_inserted' or change_type == 'image_updated':
                success = self.update_image_context(changed_object)
            elif change_type == 'paragraph_deleted' or change_type == 'table_deleted' or change_type == 'image_deleted':
                # 确定对象类型并删除
                object_type = 'paragraph' if change_type.startswith('paragraph') else \
                             'table' if change_type.startswith('table') else 'image'
                success = self.remove_object_context(object_type, changed_object)
            elif change_type == 'document_structure_changed':
                # 文档结构发生重大变化，刷新整个上下文树
                success = self.refresh_document_context_tree()
            
            # 如果是新创建的事务，则提交
            if not was_in_transaction:
                self.commit_transaction()
            
            # 记录性能指标
            self._record_operation_time('handle_change', time.time() - start_time, change_type=change_type, success=success)
            
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
            
            # 如果是新创建的事务，则回滚
            if not was_in_transaction:
                self.rollback_transaction()
            
            # 记录性能指标
            self._record_operation_time('handle_change', time.time() - start_time, success=False)
            
            # 在事务模式下抛出异常
            if self._in_transaction:
                raise DocumentContextError(
                    error_code=ErrorCode.DOCUMENT_CHANGE_HANDLING_FAILED,
                    message=f"Failed to handle document change: {str(e)}"
                )
            
            return False
    
    def _record_operation_time(self, operation_type: str, duration: float, success: bool = True, **kwargs):
        """
        记录操作的性能指标
        
        Args:
            operation_type: 操作类型
            duration: 操作持续时间（秒）
            success: 操作是否成功
            **kwargs: 其他要记录的指标（如结果数量、操作计数等）
        """
        try:
            if operation_type not in self._operation_times:
                self._operation_times[operation_type] = {
                    'count': 0,
                    'total_time': 0,
                    'success_count': 0,
                    'fail_count': 0,
                    'metrics': {}
                }
            
            # 更新基本统计信息
            self._operation_times[operation_type]['count'] += 1
            self._operation_times[operation_type]['total_time'] += duration
            
            if success:
                self._operation_times[operation_type]['success_count'] += 1
            else:
                self._operation_times[operation_type]['fail_count'] += 1
            
            # 更新额外指标
            for key, value in kwargs.items():
                if key not in self._operation_times[operation_type]['metrics']:
                    self._operation_times[operation_type]['metrics'][key] = []
                
                # 对于数值类型，记录具体值；对于其他类型，记录计数或状态
                if isinstance(value, (int, float)):
                    self._operation_times[operation_type]['metrics'][key].append(value)
                else:
                    # 对于非数值类型，记录存在性
                    if key not in self._operation_times[operation_type]['metrics']:
                        self._operation_times[operation_type]['metrics'][key] = 0
                    self._operation_times[operation_type]['metrics'][key] += 1
            
            # 记录操作频率
            current_time = time.time()
            self._last_document_operation_time = current_time
            self._document_operations_count += 1
            
            # 性能监控：如果操作时间超过阈值，记录警告
            if duration > 1.0:  # 超过1秒的操作被视为慢操作
                logger.warning(f"Slow operation detected: {operation_type} took {duration:.2f} seconds")
                
        except Exception as e:
            # 记录性能指标本身的错误不应影响主流程
            logger.error(f"Error recording operation metrics: {e}")
    
    def _add_table_context_in_batch(self, table: CDispatch, section: Optional[CDispatch] = None) -> bool:
        """
        在批量操作中添加表格上下文
        
        Args:
            table: 表格对象
            section: 所属节对象
        
        Returns:
            添加是否成功
        """
        if not table:
            return False
        
        try:
            # 如果没有提供节，则查找
            if not section:
                section = self._find_section_for_range(table.Range)
                if not section:
                    return False
            
            section_id = f"section_{section.Range.Start}"
            section_context = self._context_map.get(section_id)
            
            if section_context:
                table_id = f"table_{table.Range.Start}"
                
                # 如果已存在则直接更新
                if table_id in self._context_map:
                    return self.update_table_context(table)
                
                # 创建新的表格上下文
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
                    "created_time": time.time()
                }
                
                new_table_context = DocumentContext(
                    title=f"Table {table.Range.Start}",
                    range_obj=table.Range,
                    metadata=table_metadata
                )
                new_table_context.batch_add_objects([table_metadata])
                
                # 添加到父上下文和映射中
                self.add_context_to_tree(new_table_context, section_context)
                
                return True
            
            return False
        except Exception:
            return False
    
    def _add_image_context_in_batch(self, image: CDispatch, section: Optional[CDispatch] = None) -> bool:
        """
        在批量操作中添加图片上下文
        
        Args:
            image: 图片对象
            section: 所属节对象
        
        Returns:
            添加是否成功
        """
        if not image:
            return False
        
        try:
            # 如果没有提供节，则查找
            if not section:
                section = self._find_section_for_range(image.Range)
                if not section:
                    return False
            
            section_id = f"section_{section.Range.Start}"
            section_context = self._context_map.get(section_id)
            
            if section_context:
                image_id = f"image_{image.Range.Start}"
                
                # 如果已存在则直接更新
                if image_id in self._context_map:
                    return self.update_image_context(image)
                
                # 创建新的图片上下文
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
                    "created_time": time.time()
                }
                
                new_image_context = DocumentContext(
                    title=f"Image {image.Range.Start}",
                    range_obj=image.Range,
                    metadata=image_metadata
                )
                new_image_context.batch_add_objects([image_metadata])
                
                # 添加到父上下文和映射中
                self.add_context_to_tree(new_image_context, section_context)
                
                return True
            
            return False
        except Exception:
            return False
