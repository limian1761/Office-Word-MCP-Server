from typing import Dict, Any, Optional, List, Union
import time
from ..utils.exceptions import WordDocumentError, ErrorCode
from ..utils.logger import log_info, log_error, log_debug
from ..utils.decorators import handle_com_error, record_operation_time
from .context_control import DocumentContext

@handle_com_error(ErrorCode.SERVER_ERROR, "update paragraph context")
@record_operation_time
def update_paragraph_context(
    document: object,
    context_id: str,
    new_content: Optional[str] = None,
    formatting: Optional[Dict[str, Any]] = None,
    transaction_id: Optional[str] = None
) -> Dict[str, Any]:
    """更新段落上下文

    Args:
        document: Word文档COM对象
        context_id: 上下文ID
        new_content: 新的段落内容（可选）
        formatting: 格式设置（可选）
        transaction_id: 事务ID（可选）

    Returns:
        包含操作结果的字典
    """
    log_info(f"Updating paragraph context: {context_id}")
    
    try:
        # 获取根上下文
        root_context = DocumentContext.create_root_context(document)
        
        # 查找目标上下文
        target_context = root_context.find_child_context_by_id(context_id)
        
        if not target_context:
            raise WordDocumentError(ErrorCode.CONTEXT_ERROR, f"Context with ID {context_id} not found")
        
        # 备份当前状态用于回滚
        original_state = {
            "content": target_context.metadata.get("content"),
            "formatting": target_context.metadata.get("formatting")
        }
        
        try:
            # 更新内容
            if new_content is not None:
                target_context.metadata["content"] = new_content
                
                # 如果Range存在，更新实际文档内容
                if hasattr(target_context, 'range') and target_context.range:
                    target_context.range.Text = new_content
            
            # 更新格式
            if formatting is not None:
                target_context.metadata["formatting"] = formatting
                
                # 应用格式（这里应该调用格式应用函数）
                if hasattr(target_context, 'range') and target_context.range:
                    # 简化的格式应用示例
                    if "font_size" in formatting:
                        target_context.range.Font.Size = formatting["font_size"]
                    if "font_name" in formatting:
                        target_context.range.Font.Name = formatting["font_name"]
                    if "bold" in formatting:
                        target_context.range.Font.Bold = formatting["bold"]
                    if "italic" in formatting:
                        target_context.range.Font.Italic = formatting["italic"]
                    
                    # 应用段落格式
                    if "alignment" in formatting:
                        align_map = {
                            "left": 0,  # wdAlignParagraphLeft
                            "center": 1,  # wdAlignParagraphCenter
                            "right": 2,  # wdAlignParagraphRight
                            "justify": 3  # wdAlignParagraphJustify
                        }
                        if formatting["alignment"] in align_map:
                            target_context.range.ParagraphFormat.Alignment = align_map[formatting["alignment"]]
        except Exception as inner_error:
            # 回滚操作
            target_context.metadata["content"] = original_state["content"]
            target_context.metadata["formatting"] = original_state["formatting"]
            raise WordDocumentError(
                ErrorCode.ROLLBACK_ERROR,
                f"Failed to update paragraph context, rolled back: {str(inner_error)}"
            )
        
        log_info(f"Successfully updated paragraph context: {context_id}")
        
        return {
            "success": True,
            "message": "Paragraph context updated successfully",
            "context_id": context_id,
            "transaction_id": transaction_id
        }
    except Exception as e:
        log_error(f"Failed to update paragraph context: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "update table context")
@record_operation_time
def update_table_context(
    document: object,
    context_id: str,
    cell_updates: Optional[List[Dict[str, Any]]] = None,
    table_properties: Optional[Dict[str, Any]] = None,
    transaction_id: Optional[str] = None
) -> Dict[str, Any]:
    """更新表格上下文

    Args:
        document: Word文档COM对象
        context_id: 上下文ID
        cell_updates: 单元格更新列表（可选）
        table_properties: 表格属性更新（可选）
        transaction_id: 事务ID（可选）

    Returns:
        包含操作结果的字典
    """
    log_info(f"Updating table context: {context_id}")
    
    try:
        # 获取根上下文
        root_context = DocumentContext.create_root_context(document)
        
        # 查找目标上下文
        target_context = root_context.find_child_context_by_id(context_id)
        
        if not target_context:
            raise WordDocumentError(ErrorCode.CONTEXT_ERROR, f"Context with ID {context_id} not found")
        
        # 备份当前状态用于回滚
        original_state = {
            "table_properties": target_context.metadata.get("table_properties"),
            "cells": target_context.metadata.get("cells", [])
        }
        
        try:
            # 更新表格属性
            if table_properties is not None:
                target_context.metadata["table_properties"] = table_properties
                
                # 如果表格对象存在，应用属性
                if hasattr(target_context, 'table') and target_context.table:
                    # 应用表格属性
                    if "style" in table_properties:
                        target_context.table.Style = table_properties["style"]
                    if "preferred_width" in table_properties:
                        target_context.table.PreferredWidthType = 3  # wdPreferredWidthPercent
                        target_context.table.PreferredWidth = table_properties["preferred_width"]
        except Exception as inner_error:
            # 回滚操作
            target_context.metadata["table_properties"] = original_state["table_properties"]
            target_context.metadata["cells"] = original_state["cells"]
            raise WordDocumentError(
                ErrorCode.ROLLBACK_ERROR,
                f"Failed to update table context, rolled back: {str(inner_error)}"
            )
        
        log_info(f"Successfully updated table context: {context_id}")
        
        return {
            "success": True,
            "message": "Table context updated successfully",
            "context_id": context_id,
            "transaction_id": transaction_id
        }
    except Exception as e:
        log_error(f"Failed to update table context: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "update image context")
@record_operation_time
def update_image_context(
    document: object,
    context_id: str,
    image_properties: Optional[Dict[str, Any]] = None,
    transaction_id: Optional[str] = None
) -> Dict[str, Any]:
    """更新图像上下文

    Args:
        document: Word文档COM对象
        context_id: 上下文ID
        image_properties: 图像属性更新（可选）
        transaction_id: 事务ID（可选）

    Returns:
        包含操作结果的字典
    """
    log_info(f"Updating image context: {context_id}")
    
    try:
        # 获取根上下文
        root_context = DocumentContext.create_root_context(document)
        
        # 查找目标上下文
        target_context = root_context.find_child_context_by_id(context_id)
        
        if not target_context:
            raise WordDocumentError(ErrorCode.CONTEXT_ERROR, f"Context with ID {context_id} not found")
        
        # 备份当前状态用于回滚
        original_state = {
            "image_properties": target_context.metadata.get("image_properties")
        }
        
        try:
            # 更新图像属性
            if image_properties is not None:
                target_context.metadata["image_properties"] = image_properties
                
                # 如果图像对象存在，应用属性
                if hasattr(target_context, 'image') and target_context.image:
                    # 应用图像属性
                    if "width" in image_properties:
                        target_context.image.Width = image_properties["width"]
                    if "height" in image_properties:
                        target_context.image.Height = image_properties["height"]
                    if "lock_aspect_ratio" in image_properties:
                        target_context.image.LockAspectRatio = image_properties["lock_aspect_ratio"]
        except Exception as inner_error:
            # 回滚操作
            target_context.metadata["image_properties"] = original_state["image_properties"]
            raise WordDocumentError(
                ErrorCode.ROLLBACK_ERROR,
                f"Failed to update image context, rolled back: {str(inner_error)}"
            )
        
        log_info(f"Successfully updated image context: {context_id}")
        
        return {
            "success": True,
            "message": "Image context updated successfully",
            "context_id": context_id,
            "transaction_id": transaction_id
        }
    except Exception as e:
        log_error(f"Failed to update image context: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "remove object context")
@record_operation_time
def remove_object_context(
    document: object,
    context_id: str,
    transaction_id: Optional[str] = None
) -> Dict[str, Any]:
    """移除对象上下文

    Args:
        document: Word文档COM对象
        context_id: 上下文ID
        transaction_id: 事务ID（可选）

    Returns:
        包含操作结果的字典
    """
    log_info(f"Removing object context: {context_id}")
    
    try:
        # 获取根上下文
        root_context = DocumentContext.create_root_context(document)
        
        # 查找目标上下文
        target_context = root_context.find_child_context_by_id(context_id)
        
        if not target_context:
            raise WordDocumentError(ErrorCode.CONTEXT_ERROR, f"Context with ID {context_id} not found")
        
        # 获取父上下文
        parent_context = target_context.parent_context
        
        if not parent_context:
            raise WordDocumentError(
                ErrorCode.CONTEXT_ERROR,
                f"Cannot remove root context or context with no parent: {context_id}"
            )
        
        # 备份当前状态用于回滚
        backup_info = {
            "context_id": target_context.context_id,
            "context_data": target_context.to_dict_full(),
            "parent_id": parent_context.context_id
        }
        
        try:
            # 从文档中删除实际对象
            if hasattr(target_context, 'range') and target_context.range:
                target_context.range.Delete()
            
            # 从上下文树中移除
            parent_context.remove_child_context(target_context.context_id)
        except Exception as inner_error:
            # 这里需要实现更复杂的回滚逻辑
            # 目前简化处理，记录错误
            log_error(f"Failed to remove context, potential rollback needed: {str(inner_error)}")
            raise WordDocumentError(
                ErrorCode.ROLLBACK_ERROR,
                f"Failed to remove object context, potential data inconsistency: {str(inner_error)}"
            )
        
        log_info(f"Successfully removed object context: {context_id}")
        
        return {
            "success": True,
            "message": "Object context removed successfully",
            "context_id": context_id,
            "transaction_id": transaction_id
        }
    except Exception as e:
        log_error(f"Failed to remove object context: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "batch update contexts")
@record_operation_time
def batch_update_contexts(
    document: object,
    updates: List[Dict[str, Any]],
    transaction_id: Optional[str] = None
) -> Dict[str, Any]:
    """批量更新上下文

    Args:
        document: Word文档COM对象
        updates: 更新操作列表
        transaction_id: 事务ID（可选）

    Returns:
        包含操作结果的字典
    """
    log_info(f"Starting batch update with {len(updates)} operations")
    
    start_time = time.time()
    
    # 操作结果记录
    results = {
        "success": [],
        "failed": []
    }
    
    try:
        for update in updates:
            context_id = update.get("context_id")
            update_type = update.get("type", "update")
            
            try:
                if update_type == "update":
                    object_type = update.get("object_type", "paragraph")
                    if object_type == "paragraph":
                        update_paragraph_context(
                            document,
                            context_id,
                            new_content=update.get("content"),
                            formatting=update.get("formatting"),
                            transaction_id=transaction_id
                        )
                    elif object_type == "table":
                        update_table_context(
                            document,
                            context_id,
                            cell_updates=update.get("cell_updates"),
                            table_properties=update.get("table_properties"),
                            transaction_id=transaction_id
                        )
                    elif object_type == "image":
                        update_image_context(
                            document,
                            context_id,
                            image_properties=update.get("image_properties"),
                            transaction_id=transaction_id
                        )
                elif update_type == "remove":
                    remove_object_context(
                        document,
                        context_id,
                        transaction_id=transaction_id
                    )
                
                # 记录成功
                results["success"].append({
                    "context_id": context_id,
                    "type": update_type
                })
            except Exception as e:
                # 记录失败
                results["failed"].append({
                    "context_id": context_id,
                    "type": update_type,
                    "error": str(e)
                })
        
        # 计算耗时
        elapsed_time = time.time() - start_time
        log_info(f"Batch update completed in {elapsed_time:.2f} seconds")
        
        return {
            "success": True,
            "message": f"Batch update completed with {len(results['success'])} successes and {len(results['failed'])} failures",
            "results": results,
            "transaction_id": transaction_id,
            "elapsed_time": elapsed_time
        }
    except Exception as e:
        log_error(f"Failed to complete batch update: {str(e)}")
        raise