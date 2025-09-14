import time
from typing import Dict, List, Optional, Any, Set
from win32com.client import CDispatch
from ..common.logger import logger
from ..common.exceptions import DocumentContextError, ErrorCode
from ..com_backend.com_utils import handle_com_error
from .context_control import DocumentContext


class ContextManager:
    """上下文管理器，负责管理文档上下文对象的创建、更新、删除和查询"""
    def __init__(self):
        # 上下文对象映射表，用于快速查找
        self._context_map: Dict[str, DocumentContext] = {}
        # 上下文层次结构，维护上下文树关系
        self._context_tree: Dict[str, str] = {}
        # 事务相关状态
        self._in_transaction = False
        self._transaction_operations: List[Dict[str, Any]] = []
        self._transaction_context_backups: Dict[str, Dict[str, Any]] = {}
        # 性能监控相关
        self._operation_times: Dict[str, Dict[str, Any]] = {}
        self._last_document_operation_time = 0
        self._document_operations_count = 0

    def add_context(self, context: DocumentContext, parent_context: Optional[DocumentContext] = None) -> bool:
        """
        添加新的上下文对象到管理系统
        
        Args:
            context: 要添加的上下文对象
            parent_context: 父上下文对象（可选）
        
        Returns:
            添加是否成功
        """
        start_time = time.time()
        success = False

        try:
            context_id = context.context_id
            
            # 检查上下文是否已存在
            if context_id in self._context_map:
                logger.warning(f"Context with ID {context_id} already exists")
                return False

            # 添加到映射表
            self._context_map[context_id] = context

            # 如果有父上下文，建立父子关系
            if parent_context:
                self._context_tree[context_id] = parent_context.context_id
                parent_context.add_child_context(context)

            # 记录事务操作
            if self._in_transaction:
                self._transaction_operations.append({
                    'type': 'add',
                    'context_id': context_id,
                    'parent_id': parent_context.context_id if parent_context else None
                })

            success = True
            logger.info(f"Context {context_id} added successfully")
        except Exception as e:
            logger.error(f"Error adding context: {e}")
        finally:
            # 记录性能指标
            self._record_operation_time('add_context', time.time() - start_time, success=success)
            
        return success

    def update_context(self, context_id: str, updates: Dict[str, Any]) -> bool:
        """
        更新上下文对象
        
        Args:
            context_id: 上下文ID
            updates: 要更新的字段和值
        
        Returns:
            更新是否成功
        """
        start_time = time.time()
        success = False

        try:
            # 检查上下文是否存在
            if context_id not in self._context_map:
                logger.warning(f"Context with ID {context_id} not found")
                return False

            context = self._context_map[context_id]
            
            # 记录事务备份
            if self._in_transaction and context_id not in self._transaction_context_backups:
                self._transaction_context_backups[context_id] = context.to_dict_full()

            # 应用更新
            for key, value in updates.items():
                if hasattr(context, key):
                    setattr(context, key, value)
                elif key == 'metadata':
                    context.update_multiple_metadata(value)
                else:
                    logger.warning(f"Cannot update unknown field {key} for context {context_id}")

            # 使缓存失效
            context._invalidate_cache()

            # 记录事务操作
            if self._in_transaction:
                self._transaction_operations.append({
                    'type': 'update',
                    'context_id': context_id,
                    'updates': updates
                })

            success = True
            logger.info(f"Context {context_id} updated successfully")
        except Exception as e:
            logger.error(f"Error updating context {context_id}: {e}")
        finally:
            # 记录性能指标
            self._record_operation_time('update_context', time.time() - start_time, success=success)
            
        return success

    def remove_context(self, context_id: str) -> bool:
        """
        移除上下文对象
        
        Args:
            context_id: 上下文ID
        
        Returns:
            移除是否成功
        """
        start_time = time.time()
        success = False

        try:
            # 检查上下文是否存在
            if context_id not in self._context_map:
                logger.warning(f"Context with ID {context_id} not found")
                return False

            context = self._context_map[context_id]
            
            # 记录事务备份和操作
            if self._in_transaction:
                self._transaction_context_backups[context_id] = context.to_dict_full()
                self._transaction_operations.append({
                    'type': 'remove',
                    'context_id': context_id
                })

            # 处理子上下文
            child_contexts_to_remove = self.find_child_contexts(context_id)
            for child_id in child_contexts_to_remove:
                self.remove_context(child_id)

            # 从父上下文中移除
            if context_id in self._context_tree:
                parent_id = self._context_tree[context_id]
                if parent_id in self._context_map:
                    self._context_map[parent_id].remove_child_context(context)
                del self._context_tree[context_id]

            # 从映射表中移除
            del self._context_map[context_id]

            success = True
            logger.info(f"Context {context_id} removed successfully")
        except Exception as e:
            logger.error(f"Error removing context {context_id}: {e}")
        finally:
            # 记录性能指标
            self._record_operation_time('remove_context', time.time() - start_time, success=success)
            
        return success

    def find_context_by_id(self, context_id: str) -> Optional[DocumentContext]:
        """
        通过ID查找上下文对象
        
        Args:
            context_id: 上下文ID
        
        Returns:
            找到的上下文对象，或None
        """
        return self._context_map.get(context_id)

    def find_child_contexts(self, parent_id: str) -> List[str]:
        """
        查找指定父上下文的所有子上下文
        
        Args:
            parent_id: 父上下文ID
        
        Returns:
            子上下文ID列表
        """
        return [context_id for context_id, p_id in self._context_tree.items() if p_id == parent_id]

    def get_all_contexts(self) -> List[DocumentContext]:
        """
        获取所有上下文对象
        
        Returns:
            上下文对象列表
        """
        return list(self._context_map.values())

    def begin_transaction(self) -> None:
        """\开始事务"""
        if not self._in_transaction:
            self._in_transaction = True
            self._transaction_operations = []
            self._transaction_context_backups = {}
            logger.info("Transaction began")

    def commit_transaction(self) -> None:
        """\提交事务"""
        if self._in_transaction:
            self._in_transaction = False
            self._transaction_operations = []
            self._transaction_context_backups = {}
            logger.info("Transaction committed")

    def rollback_transaction(self) -> None:
        """\回滚事务"""
        if self._in_transaction:
            try:
                # 处理回滚操作
                for op in reversed(self._transaction_operations):
                    if op['type'] == 'add' and op['context_id'] in self._context_map:
                        # 移除添加的上下文
                        self.remove_context(op['context_id'])
                    elif op['type'] == 'remove' and op['context_id'] in self._transaction_context_backups:
                        # 恢复移除的上下文
                        # 这部分逻辑需要根据DocumentContext的实现进行调整
                        pass
                    elif op['type'] == 'update' and op['context_id'] in self._transaction_context_backups:
                        # 恢复更新的上下文
                        # 这部分逻辑需要根据DocumentContext的实现进行调整
                        pass
            except Exception as e:
                logger.error(f"Error during transaction rollback: {e}")
            finally:
                self._in_transaction = False
                self._transaction_operations = []
                self._transaction_context_backups = {}
                logger.info("Transaction rolled back")

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

    def get_performance_metrics(self) -> Dict[str, Any]:
        """
        获取性能指标报告
        
        Returns:
            包含性能指标的字典
        """
        return self._operation_times.copy()

    def clear_all_contexts(self) -> bool:
        """
        清除所有上下文
        
        Returns:
            操作是否成功
        """
        start_time = time.time()
        success = False

        try:
            # 在事务中执行清除操作
            self.begin_transaction()
            
            # 复制所有上下文ID，避免在迭代中修改集合
            context_ids = list(self._context_map.keys())
            
            # 移除所有上下文
            for context_id in context_ids:
                self.remove_context(context_id)
            
            # 清空映射表和树结构
            self._context_map.clear()
            self._context_tree.clear()
            
            # 提交事务
            self.commit_transaction()
            
            success = True
            logger.info("All contexts cleared successfully")
        except Exception as e:
            logger.error(f"Error clearing all contexts: {e}")
            # 回滚事务
            self.rollback_transaction()
        finally:
            # 记录性能指标
            self._record_operation_time('clear_all_contexts', time.time() - start_time, success=success)
            
        return success


# 创建全局上下文管理器实例
global_context_manager = ContextManager()


def get_context_manager() -> ContextManager:
    """\获取全局上下文管理器实例"""
    return global_context_manager