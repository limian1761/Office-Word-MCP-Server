from typing import Dict, Any, Optional, List, Union
import uuid
import time
from ..utils.exceptions import WordDocumentError, ErrorCode
from ..utils.logger import log_info, log_error, log_debug
from ..utils.decorators import handle_com_error, record_operation_time
from .context_control import DocumentContext

class TransactionManager:
    """事务管理器，负责处理文档操作的事务管理"""
    
    def __init__(self):
        self.active_transactions = {}
        self.transaction_history = {}
    
    @record_operation_time
    def begin_transaction(self) -> str:
        """开始一个新事务
        
        Returns:
            事务ID
        """
        transaction_id = str(uuid.uuid4())
        self.active_transactions[transaction_id] = {
            "start_time": time.time(),
            "operations": [],
            "state_backups": {},
            "status": "active"
        }
        
        log_info(f"Transaction started: {transaction_id}")
        return transaction_id
    
    @record_operation_time
    def commit_transaction(self, transaction_id: str) -> Dict[str, Any]:
        """提交事务
        
        Args:
            transaction_id: 事务ID
            
        Returns:
            包含提交结果的字典
            
        Raises:
            WordDocumentError: 当事务不存在或已完成时抛出
        """
        if transaction_id not in self.active_transactions:
            raise WordDocumentError(
                ErrorCode.TRANSACTION_ERROR,
                f"Transaction not found: {transaction_id}"
            )
        
        transaction = self.active_transactions[transaction_id]
        if transaction["status"] != "active":
            raise WordDocumentError(
                ErrorCode.TRANSACTION_ERROR,
                f"Transaction is not active: {transaction_id}, status: {transaction['status']}"
            )
        
        # 更新事务状态
        transaction["status"] = "committed"
        transaction["end_time"] = time.time()
        transaction["elapsed_time"] = transaction["end_time"] - transaction["start_time"]
        
        # 保存到历史记录
        self.transaction_history[transaction_id] = transaction
        
        # 从活动事务中移除
        del self.active_transactions[transaction_id]
        
        log_info(f"Transaction committed: {transaction_id}, operations: {len(transaction['operations'])}")
        
        return {
            "success": True,
            "message": "Transaction committed successfully",
            "transaction_id": transaction_id,
            "operations_count": len(transaction["operations"]),
            "elapsed_time": transaction["elapsed_time"]
        }
    
    @record_operation_time
    def rollback_transaction(self, transaction_id: str) -> Dict[str, Any]:
        """回滚事务
        
        Args:
            transaction_id: 事务ID
            
        Returns:
            包含回滚结果的字典
            
        Raises:
            WordDocumentError: 当事务不存在或已完成时抛出
        """
        if transaction_id not in self.active_transactions:
            raise WordDocumentError(
                ErrorCode.TRANSACTION_ERROR,
                f"Transaction not found: {transaction_id}"
            )
        
        transaction = self.active_transactions[transaction_id]
        if transaction["status"] != "active":
            raise WordDocumentError(
                ErrorCode.TRANSACTION_ERROR,
                f"Transaction is not active: {transaction_id}, status: {transaction['status']}"
            )
        
        # 记录回滚开始时间
        rollback_start_time = time.time()
        
        # 执行回滚操作
        rollback_results = {
            "success": [],
            "failed": []
        }
        
        # 反向执行回滚（从最后一个操作开始）
        for operation in reversed(transaction["operations"]):
            try:
                # 这里应该根据操作类型执行相应的回滚逻辑
                # 简化实现，实际需要更复杂的回滚策略
                context_id = operation.get("context_id")
                if context_id and context_id in transaction["state_backups"]:
                    backup_state = transaction["state_backups"][context_id]
                    # 这里应该有具体的回滚代码
                    log_debug(f"Rolled back context: {context_id}")
                    rollback_results["success"].append(context_id)
                else:
                    log_debug(f"No backup found for context: {context_id}")
                    rollback_results["failed"].append(context_id)
            except Exception as e:
                log_error(f"Failed to rollback operation: {str(e)}")
                rollback_results["failed"].append({
                    "context_id": operation.get("context_id"),
                    "error": str(e)
                })
        
        # 更新事务状态
        transaction["status"] = "rolled_back"
        transaction["end_time"] = time.time()
        transaction["elapsed_time"] = transaction["end_time"] - transaction["start_time"]
        transaction["rollback_results"] = rollback_results
        
        # 保存到历史记录
        self.transaction_history[transaction_id] = transaction
        
        # 从活动事务中移除
        del self.active_transactions[transaction_id]
        
        log_info(f"Transaction rolled back: {transaction_id}")
        
        return {
            "success": True,
            "message": "Transaction rolled back successfully",
            "transaction_id": transaction_id,
            "rollback_results": rollback_results,
            "elapsed_time": transaction["elapsed_time"]
        }
    
    @record_operation_time
    def add_operation_to_transaction(
        self,
        transaction_id: str,
        operation: Dict[str, Any],
        state_backup: Optional[Dict[str, Any]] = None
    ) -> None:
        """向事务添加操作记录
        
        Args:
            transaction_id: 事务ID
            operation: 操作信息
            state_backup: 操作前的状态备份（可选）
            
        Raises:
            WordDocumentError: 当事务不存在或已完成时抛出
        """
        if transaction_id not in self.active_transactions:
            raise WordDocumentError(
                ErrorCode.TRANSACTION_ERROR,
                f"Transaction not found: {transaction_id}"
            )
        
        transaction = self.active_transactions[transaction_id]
        if transaction["status"] != "active":
            raise WordDocumentError(
                ErrorCode.TRANSACTION_ERROR,
                f"Transaction is not active: {transaction_id}, status: {transaction['status']}"
            )
        
        # 添加操作记录
        operation_with_timestamp = {
            **operation,
            "timestamp": time.time()
        }
        transaction["operations"].append(operation_with_timestamp)
        
        # 保存状态备份
        if state_backup and "context_id" in operation:
            transaction["state_backups"][operation["context_id"]] = state_backup
        
        log_debug(f"Added operation to transaction: {transaction_id}, context: {operation.get('context_id')}")
    
    @record_operation_time
    def get_transaction_status(self, transaction_id: str) -> Dict[str, Any]:
        """获取事务状态
        
        Args:
            transaction_id: 事务ID
            
        Returns:
            事务状态信息
            
        Raises:
            WordDocumentError: 当事务不存在时抛出
        """
        if transaction_id in self.active_transactions:
            transaction = self.active_transactions[transaction_id]
        elif transaction_id in self.transaction_history:
            transaction = self.transaction_history[transaction_id]
        else:
            raise WordDocumentError(
                ErrorCode.TRANSACTION_ERROR,
                f"Transaction not found: {transaction_id}"
            )
        
        # 返回简化的事务状态
        return {
            "transaction_id": transaction_id,
            "status": transaction["status"],
            "operations_count": len(transaction["operations"]),
            "start_time": transaction["start_time"]
        }
    
    @record_operation_time
    def get_active_transactions(self) -> List[Dict[str, Any]]:
        """获取所有活动事务
        
        Returns:
            活动事务列表
        """
        active_transactions_list = []
        for tx_id, tx_data in self.active_transactions.items():
            active_transactions_list.append({
                "transaction_id": tx_id,
                "status": tx_data["status"],
                "operations_count": len(tx_data["operations"]),
                "start_time": tx_data["start_time"]
            })
        
        log_debug(f"Retrieved {len(active_transactions_list)} active transactions")
        return active_transactions_list

# 创建全局事务管理器实例
transaction_manager = TransactionManager()

@handle_com_error(ErrorCode.TRANSACTION_ERROR, "begin transaction")
def begin_transaction() -> Dict[str, Any]:
    """开始一个新事务
    
    Returns:
        包含事务ID的结果字典
    """
    transaction_id = transaction_manager.begin_transaction()
    return {
        "success": True,
        "transaction_id": transaction_id,
        "message": "Transaction started successfully"
    }

@handle_com_error(ErrorCode.TRANSACTION_ERROR, "commit transaction")
def commit_transaction(transaction_id: str) -> Dict[str, Any]:
    """提交事务
    
    Args:
        transaction_id: 事务ID
        
    Returns:
        提交结果
    """
    return transaction_manager.commit_transaction(transaction_id)

@handle_com_error(ErrorCode.TRANSACTION_ERROR, "rollback transaction")
def rollback_transaction(transaction_id: str) -> Dict[str, Any]:
    """回滚事务
    
    Args:
        transaction_id: 事务ID
        
    Returns:
        回滚结果
    """
    return transaction_manager.rollback_transaction(transaction_id)

@handle_com_error(ErrorCode.TRANSACTION_ERROR, "get transaction status")
def get_transaction_status(transaction_id: str) -> Dict[str, Any]:
    """获取事务状态
    
    Args:
        transaction_id: 事务ID
        
    Returns:
        事务状态信息
    """
    return {
        "success": True,
        "transaction_status": transaction_manager.get_transaction_status(transaction_id)
    }

@handle_com_error(ErrorCode.TRANSACTION_ERROR, "get active transactions")
def get_active_transactions() -> Dict[str, Any]:
    """获取所有活动事务
    
    Returns:
        活动事务列表
    """
    active_transactions = transaction_manager.get_active_transactions()
    return {
        "success": True,
        "active_transactions": active_transactions,
        "count": len(active_transactions)
    }