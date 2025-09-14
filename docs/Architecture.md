### **架构设计文档：Word MCP服务 (COM接口版)**

**版本:** 1.2.0
**核心理念:** **奥卡姆剃刀 (Ockham's Razor)** - 如无必要，勿增实体。
**日期:** 2024年11月5日

**更新说明:** 完成了从传统定位器机制到AppContext上下文管理的全面迁移，重构了text_operations.py等核心操作文件，完全移除了对SelectorEngine的依赖。

---

### **1. 概述**

#### **1.1. 目标**
本服务旨在为AI大语言模型提供一个**简洁、强大且可靠**的接口，用于对本地Microsoft Word文档进行精确的自动化操作。服务遵循模型上下文协议 (MCP) 规范，通过一系列简单直接的工具 (Tools) 暴露Word的核心功能。

#### **1.2. 设计哲学**
我们严格遵循**奥卡姆剃刀**原则：
*   **简洁性优于完备性:** 优先实现最常用、最高价值的功能。避免为不常见的边缘情况增加复杂性。
*   **实用主义:** 架构和工具的设计以能用、好用为首要目标。
*   **关注点分离:** 将复杂的系统分解为简单、独立的组件，每个组件只做一件事。
*   **上下文感知:** 自动维护操作上下文，简化工具调用，提高用户体验。

---

### **2. 核心架构**

系统由五个高度解耦的组件构成，确保了逻辑的清晰和未来的可扩展性。

#### **2.1. 架构图**

```
+-------------------------------------------------------------------------+
|                      AI / MCP Client (e.g., LLM Agent)                  |
+---------------------------------+---------------------------------------+
                                  |
                                  | (MCP Tool Call)
                                  v
+-------------------------------------------------------------------------+
| **1. MCP服务层 (`main.py` 和 `mcp_service/`)**                          |
| (定义工具, 管理会话, 翻译AI意图)                                        |
| e.g., @mcp_server.tool()                                                |
| @handle_tool_errors                                                     |
| def document_tools(ctx: Context, operation_type, ...): ...             |
|                                                                         |
| @mcp_server.tool()                                                      |
| @handle_tool_errors                                                     |
| def view_control_tools(ctx: Context, operation_type, ...): ...         |
|                                                                         |
| @mcp_server.tool()                                                      |
| @handle_tool_errors                                                     |
| def navigate_tools(ctx: Context, operation_type, ...): ...            |
+---------------------------------+---------------------------------------+
                                  |
                                  | (Calls Selector & Selection)
                                  v
+-------------------------------------------------------------------------+
| **2. AppContext 上下文管理**                                            |
| (维护活动文档、活动上下文、活动对象，替代了传统的定位器机制)            |
| - `set_active_document()`: 设置当前活动文档                             |
| - `set_active_context()`: 设置当前活动上下文 (基于文档大纲)              |
| - `set_active_object()`: 设置当前活动对象 (基于上下文对象列表)           |
+---------------------------------+---------------------------------------+
                                  |
                                  | (Access Document Elements)
                                  v
+-------------------------------------------------------------------------+
| **3. 选择器引擎 (`selector.py`)**                                       |
| (优先根据 AppContext 上下文精确查找文档中的元素，替代了传统的Locator机制) |
+---------------------------------+---------------------------------------+
                                  |
                                  v
+-------------------------------------------------------------------------+
| **4. 选择集抽象 (`selection.py`)**                                      |
| (为找到的元素提供统一的原子操作接口, 如 .delete(), .apply_format())     |
+---------------------------------+---------------------------------------+
                                  |
                                  v
+-------------------------------------------------------------------------+
| **5. 操作层 (`operations/`)**                                           |
| (提供底层文档操作实现，如文本、表格、图片、注释等操作)                 |
| - `document_ops.py`: 文档基础操作                                      |
| - `text_operations.py`: 文本内容操作                                    |
| - `text_format_ops.py`: 文本格式操作                                    |
| - `paragraphs_ops.py`: 段落操作                                         |
| - `table_ops.py`: 表格操作                                              |
| - `image_ops.py`: 图片操作                                              |
| - `comment_ops.py`: 注释操作                                            |
| - `objects_ops.py`: 对象操作                                            |
| - `range_ops.py`: 范围操作                                              |
| - `styles_ops.py`: 样式操作                                             |
| - `others_ops.py`: 其他操作                                             |
| - `view_control_ops.py`: 视图控制操作                                   |
| - `navigate_tools.py`: 导航工具操作 (上下文管理)                        |
+---------------------------------+---------------------------------------+
                                  |
                                  v
+-------------------------------------------------------------------------+
| **6. COM后端 (`com_backend/`)**                                         |
| (与Microsoft Word COM接口交互的底层实现)                                |
+-------------------------------------------------------------------------+
```

#### **2.2. AppContext 上下文管理**

AppContext 是系统的核心组件之一，负责维护和管理文档操作的上下文环境，极大地简化了工具调用流程。它通过维护活动文档、活动上下文和活动对象，使其他工具可以在不需要指定任何定位参数的情况下进行操作，一切操作都围绕这些活动元素展开，完全替代了传统的定位器机制。

##### **2.2.1. 核心概念**

1. **活动文档 (Active Document)**
   - 当前正在操作的Word文档实例
   - 默认情况下，系统会自动选择最新打开的文档
   - 可以通过 `set_active_document()` 显式切换

2. **活动上下文 (Active Context)**
   - **根据活动文档大纲确定**的当前工作区域
   - 支持多种上下文类型：正文、标题、目录、脚注、尾注、批注等
   - 通过 `set_active_context()` 设置，系统会自动识别该上下文中的所有对象

3. **活动对象 (Active Object)**
   - **根据上下文的对象列表确定**的当前被选中或正在操作的特定对象
   - 可以是段落、表格、图片、注释等文档元素
   - 通过 `set_active_object()` 设置，后续操作将默认针对此对象

##### **2.2.2. 设计优势**

- **简化工具调用**: 工具不再需要显式指定定位参数，通过navigate_tools设置活动对象即可
- **操作连贯性**: 所有操作自动围绕当前上下文和对象进行
- **用户体验优化**: 提供了更加流畅和自然的操作流程
- **代码简化**: 完全替代了传统的定位器机制，提高了代码复用性

#### **2.3. 操作层 (Operations Layer)**

操作层是系统的核心组件之一，负责提供底层文档操作的具体实现。该层被设计为独立的模块集合，每个模块专注于特定类型的操作。

##### **2.3.1. 模块结构**

1. **document_ops.py** - 文档基础操作
   - 提供创建、打开、关闭、保存文档和获取文档大纲等基本操作
   - 处理文档级别的核心功能

2. **text_operations.py** - 文本内容操作
   - 提供文本的插入、替换、获取和字符计数等操作
   - 处理文档中文本内容的增删改查

3. **text_format_ops.py** - 文本格式操作
   - 提供文本格式化功能，如设置字体、字号、颜色、对齐方式等
   - 实现文本样式的精确控制

4. **paragraphs_ops.py** - 段落操作
   - 提供获取文档段落、段落信息和范围内段落的功能
   - 处理文档段落级别的查询操作

5. **table_ops.py** - 表格操作
   - 提供创建表格、获取表格信息、操作单元格内容和插入行列等功能
   - 处理文档中表格的所有操作

6. **image_ops.py** - 图片操作
   - 提供插入图片、获取图片信息、调整图片大小和设置图片颜色类型等功能
   - 处理文档中图片的所有操作

7. **comment_ops.py** - 注释操作
   - 提供添加、删除、编辑、回复注释和获取注释信息等功能
   - 处理文档中注释的所有操作

8. **objects_ops.py** - 对象操作
   - 提供创建和管理书签、引用和超链接等功能
   - 处理文档中的特殊对象元素

9. **range_ops.py** - 范围操作
   - 提供批量选择对象、应用格式和删除对象等功能
   - 处理文档中对象范围的操作

10. **styles_ops.py** - 样式操作
    - 提供应用格式化和设置字体等功能
    - 处理文档中的样式应用

11. **others_ops.py** - 其他操作
    - 提供比较文档、转换格式、导出PDF、获取文档统计信息、打印和保护文档等辅助功能
    - 处理文档的各种辅助操作

12. **view_control_ops.py** - 视图控制操作
    - 提供文档视图控制功能，如切换视图、设置缩放比例、显示/隐藏元素等
    - 处理文档的视图相关操作

13. **navigate_tools.py** - 导航工具操作
    - 提供上下文管理功能，如设置活动文档、活动上下文和活动对象
    - 支持的上下文类型：正文、标题、目录、脚注、尾注、批注
    - 支持的对象类型：段落、表格、图片、注释、超链接
    - 主要操作类型：set_active_context, set_active_object, get_active_context, get_active_object
    - 这是系统中唯一的定位机制，替代了传统的定位器(Locator)，通过设置活动对象，其他工具可以在不需要指定任何定位参数的情况下进行操作

##### **2.3.2. 设计原则**

1. **一致性** - 所有操作函数遵循相同的设计模式和编码规范
2. **可维护性** - 清晰的文档字符串和类型注解使代码易于理解和维护
3. **错误处理** - 统一的错误处理机制确保系统健壮性
4. **可测试性** - 独立的模块设计便于单元测试和集成测试

---

### **3. 数据流**

1. **工具调用** - AI客户端通过MCP协议调用特定工具
2. **上下文解析** - MCP服务层解析请求，优先使用AppContext
3. **上下文应用** - 系统自动使用AppContext中维护的活动文档、活动上下文和活动对象
4. **元素选择** - 选择器引擎根据AppContext查找文档元素
5. **操作执行** - 选择集抽象调用操作层执行具体操作
6. **COM交互** - 操作层通过COM后端与Word应用程序交互
7. **结果返回** - 操作结果逐层返回给AI客户端

#### **3.1. 上下文感知数据流优势**

- **默认操作路径**: 系统自动维护操作上下文，不再需要指定任何定位参数
- **智能定位**: 选择器引擎完全基于AppContext中的上下文信息进行定位
- **操作连续性**: 多次连续操作默认针对同一上下文和对象
- **上下文切换**: 通过`navigate_tools`可以轻松切换操作上下文，替代了传统的定位器机制

---

### **4. 错误处理**

系统采用分层错误处理机制：
- **COM层** - 捕获并转换COM接口错误
- **操作层** - 统一处理操作相关错误
- **工具层** - 提供用户友好的错误信息
- **MCP层** - 按照MCP协议格式化错误响应

---

### **5. 扩展性**

系统设计支持以下扩展方式：
- **新增操作** - 在操作层添加新功能模块
- **新增工具** - 在MCP服务层定义新工具接口
- **增强选择器** - 扩展选择器引擎支持更多元素类型和过滤器
- **上下文扩展** - 增强AppContext支持更多上下文类型和对象类型

# Word Document MCP Server Architecture

## Overview

The Word Document MCP Server is a Python-based implementation of the Model Context Protocol that enables AI assistants to manipulate Microsoft Word documents through COM automation. The architecture is designed to provide a robust, scalable, and maintainable system for document operations.

## System Components

### 1. Core MCP Service Layer

The core service layer is built on the FastMCP framework and handles:
- Protocol implementation and message routing
- Server lifecycle management
- Session and context management
- Error handling and logging

Key files:
- [core.py](../word_docx_tools/mcp_service/core.py) - Main server initialization and configuration
- [core_utils.py](../word_docx_tools/mcp_service/core_utils.py) - Utility functions for error handling and logging
- [app_context.py](../word_docx_tools/mcp_service/app_context.py) - Application context management

### 2. Document Context Management

The context management system provides:
- Word application instance management
- Active document tracking
- Document structure analysis and caching
- Context tree for efficient element targeting

Key files:
- [app_context.py](../word_docx_tools/mcp_service/app_context.py) - Main context management implementation
- [context_control.py](../word_docx_tools/contexts/context_control.py) - Document context representation

### 3. Operations Layer

The operations layer provides low-level document manipulation functions:
- Document operations (create, open, save, close)
- Text operations (insert, replace, format)
- Table operations (create, modify, format)
- Image operations (insert, modify)
- Paragraph operations (format, manipulate)
- Style operations (apply, modify)
- Comment operations (add, modify, delete)
- Object operations (select, manipulate)

Key files:
- [document_ops.py](../word_docx_tools/operations/document_ops.py) - Document-level operations
- [text_operations.py](../word_docx_tools/operations/text_operations.py) - Text manipulation operations
- [table_ops.py](../word_docx_tools/operations/table_ops.py) - Table operations
- [image_ops.py](../word_docx_tools/operations/image_ops.py) - Image operations
- [paragraphs_ops.py](../word_docx_tools/operations/paragraphs_ops.py) - Paragraph operations
- [styles_ops.py](../word_docx_tools/operations/styles_ops.py) - Style operations
- [comment_ops.py](../word_docx_tools/operations/comment_ops.py) - Comment operations
- [objects_ops.py](../word_docx_tools/operations/objects_ops.py) - Object operations

### 4. Tools Layer

The tools layer exposes operations through the MCP protocol:
- Tool definitions and registration
- Parameter validation and parsing
- Result formatting and error handling

Key files:
- [document_tools.py](../word_docx_tools/tools/document_tools.py) - Document management tools
- [text_tools.py](../word_docx_tools/tools/text_tools.py) - Text manipulation tools
- [table_tools.py](../word_docx_tools/tools/table_tools.py) - Table manipulation tools
- [image_tools.py](../word_docx_tools/tools/image_tools.py) - Image manipulation tools
- [paragraph_tools.py](../word_docx_tools/tools/paragraph_tools.py) - Paragraph manipulation tools
- [styles_tools.py](../word_docx_tools/tools/styles_tools.py) - Style manipulation tools
- [comment_tools.py](../word_docx_tools/tools/comment_tools.py) - Comment manipulation tools
- [objects_tools.py](../word_docx_tools/tools/objects_tools.py) - Object manipulation tools

### 5. COM Backend

The COM backend handles communication with Microsoft Word:
- Word application lifecycle management
- COM object creation and validation
- Error handling for COM operations

Key files:
- [word_backend.py](../word_docx_tools/com_backend/word_backend.py) - Word application management
- [com_utils.py](../word_docx_tools/com_backend/com_utils.py) - COM utility functions

## Data Flow

1. **Initialization**: Server starts and initializes the AppContext
2. **Connection**: MCP client connects and registers tools
3. **Operation Request**: Client sends tool call with parameters
4. **Parameter Processing**: Tools layer validates and processes parameters
5. **Context Management**: AppContext ensures Word application is available
6. **Operation Execution**: Operations layer performs the requested action
7. **Result Formatting**: Results are formatted and returned to the client
8. **Error Handling**: Errors are caught, logged, and returned to the client

## Key Design Principles

### 1. Separation of Concerns
Each layer has a specific responsibility:
- Tools: Interface with MCP protocol
- Operations: Document manipulation logic
- Context: State management
- Backend: COM communication

### 2. Error Resilience
- Comprehensive error handling at all levels
- Graceful degradation when Word is unavailable
- Detailed error reporting for debugging

### 3. Performance Optimization
- Context caching for repeated operations
- Efficient COM object management
- Lazy loading of document structures

### 4. Extensibility
- Modular design allows for easy addition of new operations
- Standardized tool interfaces
- Consistent error handling patterns

## State Management

The system maintains state through the AppContext which tracks:
- Active Word application instance
- Current document
- Document context tree
- Operation history

This state is managed per MCP session, allowing for multiple concurrent users.

## Error Handling

Error handling follows these principles:
- All errors are caught and logged
- User-friendly error messages are returned
- Specific error codes for different error types
- Automatic recovery attempts for transient errors

## Performance Considerations

1. **COM Object Management**: Efficient creation and reuse of COM objects
2. **Context Caching**: Caching of document structures to avoid repeated analysis
3. **Batch Operations**: Support for batch operations where possible
4. **Resource Cleanup**: Proper cleanup of Word instances and documents

## Security Considerations

1. **File System Access**: Operations are limited to Word document manipulation
2. **Input Validation**: All parameters are validated before use
3. **Error Information**: Sensitive information is not exposed in error messages
4. **Resource Limits**: Operations are designed to avoid excessive resource consumption
