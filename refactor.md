# Word Docx Tools 重构计划

## 问题总结

在代码审查和运行测试过程中，发现了以下几个关键问题需要修复：

### 1. 模块结构和缩进问题

- **context_control.py** 中的函数缩进不正确，导致函数被错误地嵌套在其他函数内部
- 需要确保所有函数都是模块级别的，不使用 `self` 引用

### 2. 导入语句错误

- **contexts/__init__.py** 导入了不存在的函数：
  - `get_context_by_position`
  - `get_all_contexts_of_type`
- **contexts/__init__.py** 导入了不存在的变量 `metadata_processor`（正确的名称是 `global_metadata_processor`）
- **document_change_handler.py** 从不存在的 `common.logger` 模块导入 `logger`
- **table_ops.py** 从错误的位置导入 `DocumentContext`（应该从 `models.context` 导入）
- **comment_tools.py** 从错误的位置导入 `_get_selection_range` 函数

### 3. Pydantic JSON Schema 生成问题

- 当尝试导入包含 MCP 服务相关类的模块时，Pydantic 无法为 `mcp.server.session.ServerSession` 生成 JSON Schema

## 已完成的修复

### 1. 修复 context_control.py 缩进问题

已确认 context_control.py 文件中的函数（如 `navigate_to_previous_object`、`get_context_information` 和 `set_zoom_level`）都已修复为正确的模块级结构。

### 2. 修复 contexts/__init__.py 导入错误

- 移除了对不存在的 `get_context_by_position` 和 `get_all_contexts_of_type` 函数的导入
- 添加了正确的 `search_contexts` 和 `find_contexts_by_range` 函数导入
- 将 `metadata_processor` 改为 `global_metadata_processor`
- 更新了 `__all__` 列表以保持一致性

### 3. 修复 document_change_handler.py 导入错误

- 将 logger 导入从 `from ..common.logger import logger` 修改为 `from ..mcp_service.core_utils import logger`

### 4. 修复 table_ops.py 导入错误

- 将 DocumentContext 的导入从 `mcp_service.core_utils` 改为 `models.context`

### 5. 修复 comment_tools.py 导入错误

- 将 `_get_selection_range` 的导入改为从 `com_backend.selector_utils` 导入 `get_selection_range`
- 更新了函数调用以匹配新的函数签名

## 剩余问题

### Pydantic JSON Schema 生成问题

需要进一步调查和解决 Pydantic 无法为 `mcp.server.session.ServerSession` 生成 JSON Schema 的问题。可能的解决方案包括：

1. 修改 Pydantic 模型定义
2. 使用自定义的 JSON Schema 生成逻辑
3. 在导入时采取延迟加载策略
4. 配置 Pydantic 以忽略特定类型的 JSON Schema 生成

## 测试策略

1. **导入测试**：使用简化版的 `test_imports.py` 测试基本导入功能
2. **单元测试**：运行项目中的单元测试套件验证核心功能
3. **集成测试**：验证不同模块之间的交互是否正常
4. **MCP 服务测试**：验证 MCP 服务能够正常启动和处理请求

## 重构建议

### 1. 代码组织改进

- 建立更清晰的模块层次结构
- 为每个模块定义明确的职责边界
- 避免循环导入

### 2. 导入语句优化

- 集中管理跨模块的导入依赖
- 使用相对导入路径确保模块可移植性
- 实现延迟导入机制以解决循环导入问题

### 3. 错误处理改进

- 统一错误处理模式
- 提供更详细的错误信息
- 为常见错误场景定义明确的错误代码

### 4. 文档改进

- 为每个模块添加清晰的文档字符串
- 记录所有公共函数和类的参数、返回值和使用示例
- 更新项目文档以反映最新的代码结构和API