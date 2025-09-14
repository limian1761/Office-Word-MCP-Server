from typing import Dict, Any, Optional, List, Callable, Set
import time
from ..utils.exceptions import WordDocumentError, ErrorCode
from ..utils.logger import log_info, log_error, log_debug
from ..utils.decorators import handle_com_error, record_operation_time
from .context_control import DocumentContext
from .context_transaction import begin_transaction, commit_transaction, rollback_transaction

class DocumentChangeHandler:
    """文档变更处理器，负责处理文档变更事件"""
    
    def __init__(self):
        self._update_handlers: Dict[str, List[Callable]] = {}
        self._before_change_handlers: List[Callable] = []
        self._after_change_handlers: List[Callable] = []
        self._disabled_events: Set[str] = set()
    
    def register_update_handler(self, event_type: str, handler: Callable) -> None:
        """注册更新处理器
        
        Args:
            event_type: 事件类型
            handler: 处理函数
        """
        if event_type not in self._update_handlers:
            self._update_handlers[event_type] = []
        
        if handler not in self._update_handlers[event_type]:
            self._update_handlers[event_type].append(handler)
            log_info(f"Registered update handler for event type: {event_type}")
    
    def unregister_update_handler(self, event_type: str, handler: Callable) -> None:
        """注销更新处理器
        
        Args:
            event_type: 事件类型
            handler: 处理函数
        """
        if event_type in self._update_handlers and handler in self._update_handlers[event_type]:
            self._update_handlers[event_type].remove(handler)
            log_info(f"Unregistered update handler for event type: {event_type}")
            
            # 如果该事件类型没有处理器了，删除该事件类型
            if not self._update_handlers[event_type]:
                del self._update_handlers[event_type]
    
    def register_before_change_handler(self, handler: Callable) -> None:
        """注册变更前处理器
        
        Args:
            handler: 处理函数
        """
        if handler not in self._before_change_handlers:
            self._before_change_handlers.append(handler)
            log_info("Registered before change handler")
    
    def unregister_before_change_handler(self, handler: Callable) -> None:
        """注销变更前处理器
        
        Args:
            handler: 处理函数
        """
        if handler in self._before_change_handlers:
            self._before_change_handlers.remove(handler)
            log_info("Unregistered before change handler")
    
    def register_after_change_handler(self, handler: Callable) -> None:
        """注册变更后处理器
        
        Args:
            handler: 处理函数
        """
        if handler not in self._after_change_handlers:
            self._after_change_handlers.append(handler)
            log_info("Registered after change handler")
    
    def unregister_after_change_handler(self, handler: Callable) -> None:
        """注销变更后处理器
        
        Args:
            handler: 处理函数
        """
        if handler in self._after_change_handlers:
            self._after_change_handlers.remove(handler)
            log_info("Unregistered after change handler")
    
    def enable_event(self, event_type: str) -> None:
        """启用事件
        
        Args:
            event_type: 事件类型
        """
        if event_type in self._disabled_events:
            self._disabled_events.remove(event_type)
            log_info(f"Enabled event: {event_type}")
    
    def disable_event(self, event_type: str) -> None:
        """禁用事件
        
        Args:
            event_type: 事件类型
        """
        self._disabled_events.add(event_type)
        log_info(f"Disabled event: {event_type}")
    
    def notify_update_handlers(self, event_type: str, data: Dict[str, Any]) -> None:
        """通知更新处理器
        
        Args:
            event_type: 事件类型
            data: 事件数据
        """
        # 检查事件是否被禁用
        if event_type in self._disabled_events:
            log_debug(f"Event notification skipped (disabled): {event_type}")
            return
        
        # 通知特定事件类型的处理器
        if event_type in self._update_handlers:
            for handler in self._update_handlers[event_type]:
                try:
                    handler(data)
                except Exception as e:
                    log_error(f"Error in update handler for {event_type}: {str(e)}")
    
    def notify_before_change_handlers(self, change_data: Dict[str, Any]) -> None:
        """通知变更前处理器
        
        Args:
            change_data: 变更数据
        """
        for handler in self._before_change_handlers:
            try:
                handler(change_data)
            except Exception as e:
                log_error(f"Error in before change handler: {str(e)}")
    
    def notify_after_change_handlers(self, change_data: Dict[str, Any]) -> None:
        """通知变更后处理器
        
        Args:
            change_data: 变更数据
        """
        for handler in self._after_change_handlers:
            try:
                handler(change_data)
            except Exception as e:
                log_error(f"Error in after change handler: {str(e)}")

# 创建全局文档变更处理器实例
document_change_handler = DocumentChangeHandler()

@handle_com_error(ErrorCode.SERVER_ERROR, "handle document change")
@record_operation_time
def handle_document_change(
    document: object,
    change_type: str,
    change_data: Dict[str, Any],
    transaction_id: Optional[str] = None
) -> Dict[str, Any]:
    """处理文档变更

    Args:
        document: Word文档COM对象
        change_type: 变更类型
        change_data: 变更数据
        transaction_id: 事务ID（可选）

    Returns:
        包含处理结果的字典
    """
    log_info(f"Handling document change: {change_type}")
    
    start_time = time.time()
    
    # 如果没有提供事务ID，创建一个新事务
    transaction_provided = transaction_id is not None
    if not transaction_provided:
        transaction_result = begin_transaction()
        transaction_id = transaction_result["transaction_id"]
    
    try:
        # 通知变更前处理器
        document_change_handler.notify_before_change_handlers({
            "change_type": change_type,
            "change_data": change_data,
            "transaction_id": transaction_id
        })
        
        # 根据变更类型执行相应的处理逻辑
        result = process_change(document, change_type, change_data, transaction_id)
        
        # 通知更新处理器
        document_change_handler.notify_update_handlers(change_type, {
            "change_data": change_data,
            "result": result,
            "transaction_id": transaction_id
        })
        
        # 如果是自动创建的事务，提交它
        if not transaction_provided:
            commit_transaction(transaction_id)
        
        # 通知变更后处理器
        document_change_handler.notify_after_change_handlers({
            "change_type": change_type,
            "change_data": change_data,
            "result": result,
            "transaction_id": transaction_id
        })
        
        # 计算耗时
        elapsed_time = time.time() - start_time
        
        log_info(f"Document change handled successfully in {elapsed_time:.2f} seconds")
        
        return {
            "success": True,
            "message": f"Document change of type '{change_type}' handled successfully",
            "result": result,
            "transaction_id": transaction_id,
            "elapsed_time": elapsed_time
        }
    except Exception as e:
        # 如果是自动创建的事务，回滚它
        if not transaction_provided and transaction_id:
            try:
                rollback_transaction(transaction_id)
            except Exception as rollback_error:
                log_error(f"Failed to rollback transaction: {str(rollback_error)}")
        
        log_error(f"Failed to handle document change: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "process change")
def process_change(
    document: object,
    change_type: str,
    change_data: Dict[str, Any],
    transaction_id: str
) -> Dict[str, Any]:
    """处理具体的文档变更

    Args:
        document: Word文档COM对象
        change_type: 变更类型
        change_data: 变更数据
        transaction_id: 事务ID

    Returns:
        处理结果
    """
    # 这里实现具体的变更处理逻辑
    # 根据不同的变更类型执行不同的操作
    if change_type == "paragraph_update":
        # 处理段落更新
        context_id = change_data.get("context_id")
        new_content = change_data.get("content")
        formatting = change_data.get("formatting")
        
        # 这里应该调用段落更新函数
        # 简化实现
        return {
            "context_id": context_id,
            "change_type": change_type,
            "status": "processed"
        }
    elif change_type == "table_update":
        # 处理表格更新
        context_id = change_data.get("context_id")
        cell_updates = change_data.get("cell_updates")
        
        # 这里应该调用表格更新函数
        # 简化实现
        return {
            "context_id": context_id,
            "change_type": change_type,
            "status": "processed",
            "cells_updated": len(cell_updates) if cell_updates else 0
        }
    elif change_type == "image_update":
        # 处理图像更新
        context_id = change_data.get("context_id")
        image_properties = change_data.get("image_properties")
        
        # 这里应该调用图像更新函数
        # 简化实现
        return {
            "context_id": context_id,
            "change_type": change_type,
            "status": "processed"
        }
    elif change_type == "object_insert":
        # 处理对象插入
        object_type = change_data.get("object_type")
        insert_position = change_data.get("position")
        
        # 这里应该调用对象插入函数
        # 简化实现
        return {
            "object_type": object_type,
            "change_type": change_type,
            "status": "processed"
        }
    elif change_type == "object_delete":
        # 处理对象删除
        context_id = change_data.get("context_id")
        
        # 这里应该调用对象删除函数
        # 简化实现
        return {
            "context_id": context_id,
            "change_type": change_type,
            "status": "processed"
        }
    else:
        # 处理未知变更类型
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT,
            f"Unsupported change type: {change_type}"
        )

@handle_com_error(ErrorCode.SERVER_ERROR, "register change handler")
def register_change_handler(
    event_type: str,
    handler: Callable
) -> Dict[str, Any]:
    """注册变更处理器

    Args:
        event_type: 事件类型
        handler: 处理函数

    Returns:
        注册结果
    """
    document_change_handler.register_update_handler(event_type, handler)
    return {
        "success": True,
        "message": f"Handler registered for event type: {event_type}"
    }

@handle_com_error(ErrorCode.SERVER_ERROR, "unregister change handler")
def unregister_change_handler(
    event_type: str,
    handler: Callable
) -> Dict[str, Any]:
    """注销变更处理器

    Args:
        event_type: 事件类型
        handler: 处理函数

    Returns:
        注销结果
    """
    document_change_handler.unregister_update_handler(event_type, handler)
    return {
        "success": True,
        "message": f"Handler unregistered for event type: {event_type}"
    }

@handle_com_error(ErrorCode.SERVER_ERROR, "enable event")
def enable_event(event_type: str) -> Dict[str, Any]:
    """启用事件

    Args:
        event_type: 事件类型

    Returns:
        启用结果
    """
    document_change_handler.enable_event(event_type)
    return {
        "success": True,
        "message": f"Event enabled: {event_type}"
    }

@handle_com_error(ErrorCode.SERVER_ERROR, "disable event")
def disable_event(event_type: str) -> Dict[str, Any]:
    """禁用事件

    Args:
        event_type: 事件类型

    Returns:
        禁用结果
    """
    document_change_handler.disable_event(event_type)
    return {
        "success": True,
        "message": f"Event disabled: {event_type}"
    }