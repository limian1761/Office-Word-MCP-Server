# Agent Guide for Office Word MCP Server

本指南提供了使用 Office Word MCP Server 的全面指导，包括工具使用方法、最佳实践和常见场景示例，帮助您高效地操作Word文档。

## 项目架构概述

Office Word MCP Server 采用五组件分层架构设计，包括：

1. **MCP服务层** - 处理工具调用请求，管理会话上下文
2. **选择器引擎** - 负责文档元素的精确定位和选择
3. **选择集抽象** - 提供统一的元素操作接口
4. **操作层** - 封装具体的文档操作逻辑
5. **COM后端** - 与Word应用程序进行交互的底层实现

这种分层架构确保了系统的模块化、可扩展性和可维护性，同时遵循了"关注点分离"原则。

## 使用流程

### 1. 打开文档

在进行任何操作之前，您需要先打开一个文档。使用 `document_tools` 工具的 `open` 操作来打开现有文档：

```python
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "document_tools",
  "args": {
    "operation_type": "open",
    "file_path": "C:\\path\\to\\your\\document.docx"
  }
}
```

### 2. 操作文档

成功打开文档后，您可以使用各种工具来操作文档内容，如添加文本、插入图片、创建表格等。以下是一些常用工具及其功能：

- **document_tools**: 文档级操作（创建、打开、保存、关闭等）
- **text_tools**: 文本内容操作（获取、插入、替换文本等）
- **image_tools**: 图片操作（插入、调整大小、添加标题等）
- **table_tools**: 表格操作（创建、插入行/列、设置单元格内容等）
- **comment_tools**: 注释操作（添加、删除、回复注释等）
- **styles_tools**: 样式操作（应用格式、设置字体、段落样式等）
- **object_tools**: 元素操作（选择、删除元素等）
- **objects_tools**: 对象操作（创建书签、超链接等）

### 3. 保存更改

完成操作后，使用 `document_tools` 工具的 `save` 或 `save_as` 操作来保存更改：

```python
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "document_tools",
  "args": {
    "operation_type": "save"
  }
}
```

或另存为新文件：

```python
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "document_tools",
  "args": {
    "operation_type": "save_as",
    "file_path": "C:\\path\\to\\new\\document.docx"
  }
}
```

### 4. 关闭文档

当您完成所有操作后，可以使用 `document_tools` 工具的 `close` 操作来关闭当前活动文档：

```python
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "document_tools",
  "args": {
    "operation_type": "close"
  }
}
```

### 5. Word应用程序生命周期管理

Word应用程序实例由MCP服务器的生命周期管理器统一管理。所有Word应用实例必须通过`get_word_app`函数获取，禁止直接使用`win32com.client.Dispatch("Word.Application")`或`win32com.client.GetActiveObject("Word.Application")`创建或获取实例。

当服务器启动时，会自动通过`get_word_app`函数创建Word应用程序实例；当服务器关闭时，会自动调用Word应用程序的退出方法。

目前，通过`document_tools`的`close`操作只能关闭文档，但不会退出Word应用程序。Word应用程序的退出是由服务器生命周期自动管理的。

## Locator使用指南

Locator是用于精确定位文档中元素的机制。所有需要操作特定元素的工具都接受locator参数。

### 基本语法

Locator支持两种主要格式：

1. **字符串格式**：`type:value[filter1][filter2]...`
   - `type`: 元素类型（如 paragraph、table、text 等）
   - `value`: 可选的元素值（如文本内容、标识符等）
   - `filter`: 可选的过滤器，用于进一步缩小选择范围

2. **JSON对象格式**：

```json
{
  "type": "paragraph",  // 元素类型
  "value": "1",         // 元素标识符（索引或唯一ID）
  "filters": [           // 可选的过滤条件
    {"contains_text": "example"}
  ]
}
```

### 带锚点的格式

对于相对定位，可以使用带锚点的格式：`type:value@anchor_id[relation]` 或JSON对象格式：

```json
{
  "anchor": {
    "type": "paragraph",
    "identifier": {
      "text": "锚点文本"
    }
  },
  "relation": {
    "type": "first_occurrence_after"
  },
  "target": {
    "type": "paragraph"
  }
}
```

### 元素类型

支持的元素类型包括：

- **paragraph**: 段落元素
  根据段落位置或内容定位。
  - 示例: `paragraph:3` (第3个段落)
  - 示例: `paragraph:"标题文本"` (包含"标题文本"的段落)
  
  **重要说明**：当段落定位器包含索引参数时（如`paragraph:5`），系统会强制返回单个对象，确保操作只针对特定段落执行，避免误操作影响整个文档。
- **table**: 表格元素
- **text**: 文本元素（会被转换为段落搜索）
- **inline_shape** 或 **image**: 图片元素
- **cell**: 表格单元格
- **run**: 文本运行（单词）
- **document_start**: 文档开始位置
- **document_end**: 文档结束位置
- **range**: 范围元素
- **bookmark**: 书签
- **comment**: 注释

### 过滤器

过滤器用于进一步缩小定位范围，常用的过滤器包括：

#### 文本相关过滤器
- **contains_text**: 元素包含指定的文本
  - 示例: `[contains_text=重要信息]`
- **text_matches_regex**: 元素文本与指定的正则表达式匹配
  - 示例: `[text_matches_regex=^第[0-9]+章]`
- **is_list_item**: 元素是列表项
  - 示例: `[is_list_item=true]`

#### 位置相关过滤器
- **index_in_parent**: 元素在父元素中的索引位置（0 开始）
  - 示例: `[index_in_parent=0]`（第一个元素）
- **row_index**: 单元格在表格中的行索引
  - 示例: `[row_index=2]`
- **column_index**: 单元格在表格中的列索引
  - 示例: `[column_index=3]`
- **table_index**: 元素所属表格的索引
  - 示例: `[table_index=0]`
- **range_start**: 范围元素的起始位置
  - 示例: `[range_start=100]`
- **range_end**: 范围元素的结束位置
  - 示例: `[range_end=200]`

#### 样式相关过滤器
- **style**: 元素具有指定的样式
  - 示例: `[style=标题1]`
- **is_bold**: 元素字体为粗体
  - 示例: `[is_bold=true]`

#### 形状相关过滤器
- **shape_type**: 形状类型
  - 示例: `[shape_type=Picture]`

### 关系类型

当使用带锚点的定位器时，可以指定以下关系类型：

1. **all_occurrences_within** - 查找锚点元素范围内的所有目标元素
2. **first_occurrence_after** - 查找锚点元素之后的第一个目标元素
3. **parent_of** - 查找锚点元素的父元素
4. **immediately_following** - 查找锚点元素之后紧接着的目标元素

### 定位器稳定性策略

为确保定位器在文档修改后仍然有效，建议采用以下策略：

1. **使用相对位置**：优先使用相对位置而非绝对索引或位置
2. **使用内容特征**：基于元素的文本内容、样式等不易受修改影响的特征进行定位
3. **使用书签或自定义标记**：在关键点位置插入书签或自定义标记作为定位锚点
4. **先查找后操作模式**：在执行修改操作前，先使用定位器查找元素并获取其实时位置信息
5. **批量执行策略**：将相关操作分组执行，以减少定位器失效的影响

### 示例

#### 定位特定段落

```json
{
  "type": "paragraph",
  "value": "3"
}
```

这将定位文档中的第3个段落。

#### 定位包含特定文本的段落

```json
{
  "type": "paragraph",
  "filters": [
    {"contains_text": "示例"}
  ]
}
```

这将定位包含"示例"文本的所有段落。

#### 定位文档中的第一个表格

```python
# 定位器字符串形式
locator_str = "table[index_in_parent=0]"
```

#### 相对定位示例

选择包含"关键词"的段落之后的第一个表格：

```python
# 字典形式的定位器
locator = {
    "anchor": {
        "type": "paragraph",
        "identifier": {
            "text": "关键词"
        }
    },
    "relation": {
        "type": "immediately_following"
    },
    "target": {
        "type": "table"
    }
}
```







## 最佳实践

### 1. 会话管理
- 在一个会话中完成所有相关操作，避免频繁创建和销毁Word实例
- 确保在完成操作后关闭文档以释放资源

### 2. 错误处理
- 所有操作都应该包含适当的错误处理逻辑
- 检查工具返回的错误代码和消息以获取详细信息
- 使用`@handle_com_error`装饰器统一处理COM错误
- 常见错误代码包括：
  - 通用错误（1001-1005）
  - 文档错误（2001-2005）
  - 元素错误（3001-3005）
  - 样式错误（4001-4005）
  - 格式化错误（5001-5005）
  - 图片错误（6001-6005）
  - 表格错误（7001-7005）
  - 注释错误（8001-8005）

### 3. 资源管理
- 避免长时间持有大型文档的打开状态
- 对于批量操作，考虑分批处理以减少内存占用
- 使用 `close` 操作及时关闭不再需要的文档
- 所有COM操作应使用上下文管理器处理，确保资源正确释放

### 4. Word应用实例管理
 - **必须通过get_word_app获取实例**: 所有Word应用实例必须通过`get_word_app`函数获取，禁止直接使用`win32com.client.Dispatch("Word.Application")`或`win32com.client.GetActiveObject("Word.Application")`
 - **统一生命周期管理**: Word应用实例由MCP服务器的生命周期管理器统一管理，确保正确的创建和销毁
 - **避免重复创建**: 在会话中重用已获取的Word应用实例，避免不必要的实例创建开销
 - **测试环境例外**: 测试代码中允许直接创建Word应用实例，但生产代码必须严格遵守get_word_app规范
 
 ### 5. 性能优化
- 对于大型文档，使用更具体的定位器来减少搜索范围
- 合并多个小操作成一个大操作以减少COM调用
- 避免在循环中执行复杂操作
- 优化选择器引擎，合理使用缓存提高元素查找效率

### 6. 安全性
- 避免在公共环境中传递敏感文档路径
- 确保有适当的文件系统权限来访问和修改文档
- 处理敏感信息时注意加密和保护
- 避免泄露文档内容

### 7. 兼容性
- 确保您的文档格式与Word版本兼容
- 对于特殊格式的文档，考虑先转换为标准格式
- 注意文件路径分隔符在不同操作系统上的差异

### 8. 测试与质量保证
- 在实际使用前，先在测试文档上验证您的操作序列
- 记录和分析测试结果，优化操作流程
- 使用单元测试、集成测试和端到端测试确保功能正常
- 利用代码质量工具如`mypy`、`black`和`isort`保持代码质量

## 常见问题与解决方案

### Q: 如何获取文档的基本信息？
A: 使用 `document_tools` 工具的 `get_info` 操作来获取文档的结构信息，包括段落、表格、图片等元素的数量和分布。

### Q: 如何在特定位置插入内容？
A: 使用定位器（locator）参数来指定插入位置。定位器可以基于元素类型、索引或文本内容进行定位。对于更复杂的定位需求，可以使用带锚点的定位器和相对关系。

### Q: 遇到COM错误怎么办？
A: COM错误通常与Word应用程序的状态有关。尝试关闭所有Word实例并重新启动MCP服务器。如果问题仍然存在，请检查文档格式和权限。使用`@handle_com_error`装饰器可以帮助统一处理COM错误。

### Q: 如何处理大型文档？
A: 对于大型文档，考虑使用更具体的定位器，减少每次操作的范围，并且在操作之间适当释放资源。同时，可以采用批量处理策略，将相关操作分组执行。

### Q: 如何保护文档？
A: 使用 `document_tools` 工具的 `protect` 操作来保护文档，支持只读、评论、表单和跟踪保护类型。

### Q: 定位器语法错误怎么办？
A: 检查定位器格式是否正确，确保遵循 `type:value[filter1][filter2]...` 格式，验证元素类型是否受支持，检查过滤器语法是否正确，特别是正则表达式。

### Q: 找不到匹配的元素怎么办？
A: 检查文档中是否存在符合定位器条件的元素，确认过滤器值是否正确（注意大小写），考虑使用更宽泛的过滤条件。

### Q: 文档修改后定位器失效怎么办？
A: 采用稳定性策略，如使用相对位置、内容特征、书签等方法，在文档修改后重新生成定位器，使用基于内容的定位而非基于位置的定位。