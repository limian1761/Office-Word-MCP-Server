# Word MCP服务架构优化方案

## 1. 架构现状分析

通过对现有代码和架构文档的分析，我们已经了解到系统采用了以下分层架构：

1. **MCP服务层**：定义工具，管理会话，翻译AI意图
2. **AppContext上下文管理**：维护活动文档、活动上下文、活动对象
3. **选择器引擎**：根据上下文精确查找文档中的元素
4. **选择集抽象**：为找到的元素提供统一的原子操作接口
5. **操作层**：提供底层文档操作实现
6. **COM后端**：与Microsoft Word COM接口交互的底层实现

### 当前架构的主要特点

- **上下文感知**：通过AppContext自动维护操作上下文，简化工具调用
- **关注点分离**：各层职责清晰，相互独立
- **单例模式**：AppContext采用单例模式管理全局状态
- **观察者模式**：通过_update_handlers机制实现上下文更新通知

### 当前架构存在的问题

1. **DocumentContext功能相对简单**：只提供了基本的上下文树管理功能
2. **上下文更新机制不够灵活**：当前update_handlers机制较为简单
3. **错误处理粒度不够细**：缺乏针对不同操作类型的专门错误处理策略
4. **上下文树构建效率问题**：对于大型文档可能存在性能瓶颈
5. **接口设计不够统一**：部分方法命名和参数格式不一致

## 2. 架构优化建议

### 2.1 DocumentContext类增强

**设计目标**：增强DocumentContext的功能，使其成为更强大的上下文管理组件。

**实现方案**：

```python
class DocumentContext:
    """增强版文档上下文类，提供更丰富的上下文管理功能"""
    
    def __init__(self, title: str = "", range_obj: Optional[Any] = None):
        """初始化文档上下文对象"""
        self.title = title
        self.range = range_obj
        self.object_list: List[Dict[str, Any]] = []
        self.parent_context: Optional['DocumentContext'] = None
        self.child_contexts: List['DocumentContext'] = []
        
        # 新增属性
        self.context_id: str = f"context_{id(self)}"  # 唯一标识
        self.metadata: Dict[str, Any] = {}  # 存储附加信息
        self.created_at = datetime.now()  # 创建时间
        self.updated_at = datetime.now()  # 更新时间
        
    def add_object(self, object_info: Dict[str, Any]) -> None:
        """添加对象到上下文"""
        self.object_list.append(object_info)
        self._update_timestamp()
    
    def add_child_context(self, child_context: 'DocumentContext') -> None:
        """添加子上下文"""
        if child_context not in self.child_contexts:
            self.child_contexts.append(child_context)
            child_context.parent_context = self
            self._update_timestamp()
    
    def remove_child_context(self, child_context: 'DocumentContext') -> None:
        """移除子上下文"""
        if child_context in self.child_contexts:
            self.child_contexts.remove(child_context)
            child_context.parent_context = None
            self._update_timestamp()
    
    def find_node_by_range(self, start: int, end: int) -> Optional[Dict[str, Any]]:
        """根据范围查找节点"""
        # 实现基于范围的节点查找逻辑
        pass
    
    def find_nodes_by_type(self, object_type: str) -> List['DocumentContext']:
        """根据对象类型查找节点"""
        # 实现基于类型的节点查找逻辑
        pass
    
    def to_dict(self, include_children: bool = True) -> Dict[str, Any]:
        """将上下文对象转换为字典格式，支持选择性包含子节点"""
        result = {
            "context_id": self.context_id,
            "title": self.title,
            "has_range": self.range is not None,
            "object_count": len(self.object_list),
            "child_count": len(self.child_contexts),
            "has_parent": self.parent_context is not None,
            "metadata": self.metadata,
            "created_at": self.created_at.isoformat(),
            "updated_at": self.updated_at.isoformat()
        }
        
        if self.range:
            try:
                result["range_info"] = {
                    "start": self.range.Start,
                    "end": self.range.End,
                    "text_preview": self.range.Text[:50] + ("..." if len(self.range.Text) > 50 else "")
                }
            except Exception:
                result["range_info"] = {"error": "Failed to get range details"}
        
        result["objects_preview"] = [
            {"type": obj.get("type", "unknown"), "id": obj.get("id", "unknown")}
            for obj in self.object_list[:5]
        ]
        
        if include_children:
            result["children"] = [child.to_dict() for child in self.child_contexts]
        
        return result
    
    def update_metadata(self, key: str, value: Any) -> None:
        """更新元数据"""
        self.metadata[key] = value
        self._update_timestamp()
    
    def _update_timestamp(self) -> None:
        """更新时间戳"""
        self.updated_at = datetime.now()
    
    # 静态工具方法
    @staticmethod
    def merge_contexts(contexts: List['DocumentContext']) -> 'DocumentContext':
        """合并多个上下文"""
        # 实现上下文合并逻辑
        pass
```

### 2.2 AppContext重构

**设计目标**：重构AppContext，优化上下文树管理和事件通知机制。

**实现方案**：

```python
class AppContext:
    """重构后的应用上下文类"""
    
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(AppContext, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    @classmethod
    def get_instance(cls) -> "AppContext":
        """获取单例实例"""
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance
    
    def __init__(self):
        """初始化AppContext"""
        if self._initialized:
            return
        
        # 原有属性
        self._temp_word_app: Optional[CDispatch] = None
        self._active_document: Optional[CDispatch] = None
        self._word_app: Optional[CDispatch] = None
        self._document_context_tree: Optional[DocumentContext] = None
        self._context_map: Dict[str, DocumentContext] = {}
        self._active_context: Optional[DocumentContext] = None
        
        # 改进的事件系统
        self._event_handlers: Dict[str, List[Callable]] = {
            "document_opened": [],
            "document_closed": [],
            "context_updated": [],
            "object_created": [],
            "object_removed": [],
            "selection_changed": [],
            # 更细粒度的事件类型
            "paragraph_updated": [],
            "table_updated": [],
            "image_updated": [],
            "style_updated": []
        }
        
        # 操作历史记录
        self._operation_history: List[Dict[str, Any]] = []
        
        self._initialized = True
    
    # 事件系统方法
    def register_event_handler(self, event_type: str, handler: Callable) -> None:
        """注册事件处理器"""
        if event_type not in self._event_handlers:
            self._event_handlers[event_type] = []
        
        if handler not in self._event_handlers[event_type]:
            self._event_handlers[event_type].append(handler)
            logger.debug(f"Registered handler for event '{event_type}': {handler.__name__}")
    
    def unregister_event_handler(self, event_type: str, handler: Callable) -> None:
        """注销事件处理器"""
        if event_type in self._event_handlers and handler in self._event_handlers[event_type]:
            self._event_handlers[event_type].remove(handler)
            logger.debug(f"Unregistered handler for event '{event_type}': {handler.__name__}")
    
    def trigger_event(self, event_type: str, **kwargs) -> None:
        """触发事件"""
        handlers = self._event_handlers.get(event_type, [])
        for handler in handlers:
            try:
                handler(**kwargs)
            except Exception as e:
                logger.error(f"Error in event handler for '{event_type}': {e}")
        
        # 对于特定事件，同时触发通用事件
        if event_type in ["paragraph_updated", "table_updated", "image_updated", "style_updated"]:
            self.trigger_event("context_updated", event_type=event_type, **kwargs)
    
    # 上下文树管理优化
    def create_document_context_tree(self) -> Optional[DocumentContext]:
        """优化的上下文树创建方法"""
        if not self._active_document:
            logger.warning("No active document to create context tree")
            return None
        
        try:
            logger.info(f"Starting to create context tree for document: {self._active_document.Name}")
            
            # 清空旧的上下文树
            self._clear_context_tree()
            
            # 创建根上下文节点
            root_title = f"Document: {self._active_document.Name}"
            root_context = DocumentContext(title=root_title)
            self._document_context_tree = root_context
            
            # 添加根上下文到映射中
            self._context_map[root_context.context_id] = root_context
            
            # 优化的文档结构构建，使用批处理方式
            self._build_document_structure_optimized(root_context)
            
            logger.info(f"Successfully created document context tree with {len(self._context_map)} contexts")
            
            # 触发文档打开事件
            self.trigger_event("document_opened", document_name=self._active_document.Name)
            
            return root_context
        except Exception as e:
            logger.error(f"Failed to create document context tree: {e}")
            logger.error(f"Traceback: {traceback.format_exc()}")
            return None
    
    def _build_document_structure_optimized(self, root_context: DocumentContext) -> None:
        """优化的文档结构构建方法"""
        # 实现更高效的文档结构构建逻辑
        # 1. 首先获取所有节
        # 2. 然后批量处理每个节中的元素
        # 3. 使用并行处理或延迟加载大型文档
        pass
    
    def _clear_context_tree(self) -> None:
        """清空上下文树"""
        self._document_context_tree = None
        self._context_map.clear()
        self._active_context = None
    
    # 批量操作支持
    def batch_update_contexts(self, update_operations: List[Dict[str, Any]]) -> Dict[str, Any]:
        """增强的批量更新功能"""
        # 实现事务性的批量更新，支持回滚
        pass
    
    # 添加操作历史记录
    def add_operation_history(self, operation_info: Dict[str, Any]) -> None:
        """添加操作历史记录"""
        operation_info["timestamp"] = datetime.now().isoformat()
        self._operation_history.append(operation_info)
        
        # 限制历史记录长度
        if len(self._operation_history) > 1000:
            self._operation_history.pop(0)
    
    # 获取操作历史
    def get_operation_history(self, limit: int = 100) -> List[Dict[str, Any]]:
        """获取操作历史记录"""
        return self._operation_history[-limit:]
```

### 2.3 错误处理机制优化

**设计目标**：引入更完善的异常体系，提高系统的健壮性和可维护性。

**实现方案**：

```python
# 定义更细粒度的异常类
class WordMCPError(Exception):
    """Word MCP服务的基础异常类"""
    def __init__(self, error_code: int, message: str, details: Optional[Dict[str, Any]] = None):
        self.error_code = error_code
        self.details = details or {}
        super().__init__(message)

class DocumentError(WordMCPError):
    """文档相关异常"""
    pass

class ContextError(WordMCPError):
    """上下文相关异常"""
    pass

class SelectionError(WordMCPError):
    """选择相关异常"""
    pass

class OperationError(WordMCPError):
    """操作相关异常"""
    pass

class COMBackendError(WordMCPError):
    """COM后端相关异常"""
    pass

# 增强的错误处理装饰器
def handle_mcp_errors(error_type: type = WordMCPError, error_code: int = ErrorCode.SERVER_ERROR):
    """增强的错误处理装饰器"""
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except com_error as e:
                # 处理COM错误，提取更多信息
                com_error_info = {
                    "com_error_code": e.hresult,
                    "com_error_text": str(e),
                    "function_name": func.__name__,
                    "args": args,
                    "kwargs": kwargs
                }
                logger.error(f"COM error in {func.__name__}: {e}")
                raise COMBackendError(
                    error_code=ErrorCode.COM_ERROR,
                    message=f"COM error occurred in {func.__name__}",
                    details=com_error_info
                )
            except WordMCPError:
                # 已包装的异常直接重新抛出
                raise
            except Exception as e:
                # 未预期的异常包装为MCP错误
                logger.error(f"Unexpected error in {func.__name__}: {e}")
                logger.error(f"Traceback: {traceback.format_exc()}")
                raise error_type(
                    error_code=error_code,
                    message=f"Error in {func.__name__}: {str(e)}",
                    details={
                        "function_name": func.__name__,
                        "args": args,
                        "kwargs": kwargs,
                        "error_type": type(e).__name__
                    }
                )
        return wrapper
    return decorator
```

### 2.4 上下文更新机制优化

**设计目标**：优化DocumentContext的更新机制，提高系统响应速度和性能。

**实现方案**：

```python
def _update_document_context_for_style(range_obj, operation_type):
    """优化的样式操作上下文更新函数"""
    try:
        # 获取AppContext实例
        app_context = AppContext.get_instance()
        
        # 记录操作开始时间，用于性能监控
        start_time = time.time()
        
        # 获取范围的起始和结束位置
        start = range_obj.Start
        end = range_obj.End
        
        # 批量更新策略：对于大范围操作，使用批量更新而非单个更新
        range_size = end - start
        if range_size > 1000:  # 阈值可调整
            # 大范围操作使用批量更新
            app_context.batch_update_contexts([{
                "type": "style_update",
                "range_start": start,
                "range_end": end,
                "operation_type": operation_type
            }])
        else:
            # 小范围操作使用精确更新
            # 查找并更新相关节点
            updated_nodes = []
            
            # 遍历上下文映射，查找受影响的节点
            for context_id, context in app_context._context_map.items():
                if context.range and hasattr(context.range, 'Start') and hasattr(context.range, 'End'):
                    # 检查节点范围是否与更新范围有重叠
                    if (context.range.Start <= end and context.range.End >= start):
                        # 更新节点信息
                        context.update_metadata("style_modified", True)
                        context.update_metadata("last_style_operation", operation_type)
                        updated_nodes.append(context_id)
            
            # 添加操作记录
            app_context.add_operation_history({
                "type": "style",
                "operation": operation_type,
                "range": {"start": start, "end": end},
                "affected_nodes": updated_nodes
            })
            
            # 触发细粒度的事件通知
            app_context.trigger_event("style_updated", {
                "range": {"start": start, "end": end},
                "operation_type": operation_type,
                "affected_nodes": updated_nodes
            })
        
        # 记录性能信息
        elapsed_time = time.time() - start_time
        logger.info(f"Style operation '{operation_type}' completed in {elapsed_time:.3f}s")
        
    except Exception as e:
        logger.error(f"Failed to update DocumentContext for style operation: {str(e)}")
        # 异常记录但不中断主流程
```

### 2.5 性能优化建议

1. **延迟加载上下文**：对于大型文档，实现上下文树的延迟加载机制

```python
class DocumentContext:
    """支持延迟加载的文档上下文类"""
    def __init__(self, title: str = "", range_obj: Optional[Any] = None):
        # 原有初始化代码
        self._children_loaded = False  # 子节点是否已加载
    
    def ensure_children_loaded(self):
        """确保子节点已加载"""
        if not self._children_loaded and self.range:
            # 加载子节点逻辑
            self._children_loaded = True
```

2. **上下文缓存机制**：添加LRU缓存减少重复查找

```python
from functools import lru_cache

class AppContext:
    # ...其他代码...
    
    @lru_cache(maxsize=100)
    def get_context_by_range(self, start: int, end: int) -> List[DocumentContext]:
        """根据范围查找上下文，结果缓存"""
        # 实现范围查找逻辑
        pass
    
    def clear_cache(self):
        """清除缓存"""
        self.get_context_by_range.cache_clear()
```

3. **批量COM操作**：减少COM调用次数，提高性能

```python
def batch_process_com_objects(objects, batch_size=100):
    """批量处理COM对象"""
    results = []
    for i in range(0, len(objects), batch_size):
        batch = objects[i:i+batch_size]
        # 批量处理逻辑
        results.extend(processed_batch)
    return results
```

## 3. 实现路线图

### 第一阶段：核心组件重构
1. 实现增强版DocumentContext类
2. 重构AppContext，引入新的事件系统
3. 更新相关操作模块以使用新的上下文接口

### 第二阶段：性能优化
1. 实现上下文树的延迟加载机制
2. 添加上下文缓存系统
3. 优化COM操作批处理

### 第三阶段：错误处理和监控
1. 实现细粒度的异常体系
2. 添加性能监控和日志系统
3. 完善错误恢复机制

### 第四阶段：测试和验证
1. 编写单元测试和集成测试
2. 性能基准测试
3. 稳定性测试

## 4. 预期收益

1. **性能提升**：通过延迟加载、缓存机制和批量操作，显著提高大型文档的处理速度
2. **可维护性增强**：清晰的接口设计和异常处理，使代码更易于理解和维护
3. **扩展性提高**：灵活的事件系统和模块化设计，便于添加新功能
4. **稳定性增强**：更完善的错误处理和恢复机制，提高系统的稳定性
5. **用户体验优化**：更快的响应速度和更准确的操作结果

## 5. 风险评估

1. **兼容性风险**：接口变更可能影响现有代码，需要进行充分测试
2. **实现复杂度增加**：新功能可能引入额外的复杂度
3. **性能优化可能带来的副作用**：缓存机制可能导致数据不一致

## 6. 总结

通过以上优化方案，我们可以显著提升Word MCP服务的性能、可维护性和扩展性，为AI大语言模型提供更强大、更可靠的Word文档操作接口。优化后的架构将更好地支持复杂文档操作和大型文档处理，同时保持代码的清晰和可维护性。