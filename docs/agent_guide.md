## Office-Word-MCP-Server 使用指南

如何使用Office-Word-MCP-Server进行Word文档操作，包括工具调用方法、参数说明和使用示例。

### 使用流程概述

1. **连接服务器**：确保MCP服务器已启动并与AI大模型客户端正确连接
2. **打开文档**：首先使用`open_document`工具打开目标Word文档
3. **执行操作**：调用相应工具进行文档编辑、内容提取等操作
4. **关闭文档**：完成操作后使用`shutdown_word`关闭Word应用实例

### 错误处理指南

在使用MCP服务器时，需要注意以下错误处理原则:

1. 确保在调用任何工具前，已经通过`open_document`打开了一个有效的文档
2. 对于可能返回错误的工具，检查返回值是否包含错误信息和标准化的错误代码
3. 当出现错误时，尝试理解错误原因并采取相应措施

以下是标准化的错误代码表:

| 错误代码 | 错误类型 | 描述 | 可能的解决方法 |
|---------|---------|------|--------------|
| [0] | 成功 | 操作成功完成 | - |
| [1001] | 参数错误 | 无效输入参数 | 检查参数格式和值是否正确 |
| [1002] | 资源错误 | 请求的资源未找到 | 确认资源存在且路径正确 |
| [1003] | 权限错误 | 权限不足 | 检查文件和系统权限设置 |
| [1004] | 服务器错误 | 内部服务器错误 | 查看服务器日志获取详细信息 |
| [2001] | 文档错误 | 没有活动文档 | 确保已通过`open_document`打开有效文档 |
| [2002] | 文档错误 | 无法打开文档 | 确认文件路径正确且文件未损坏 |
| [2003] | 文档错误 | 无法保存文档 | 检查文件权限和存储空间 |
| [2004] | 文档错误 | 无效的文档格式 | 确认文件是有效的.docx格式 |
| [3001] | 元素错误 | 元素未找到 | 检查Locator定位器是否正确 |
| [3002] | 元素错误 | 元素已锁定 | 检查文档保护设置 |
| [3003] | 元素错误 | 无效的元素类型 | 确保定位器选择了正确类型的元素 |
| [3004] | 元素错误 | 无法选择段落元素 | 检查文档结构和段落格式 |
| [4001] | 样式错误 | 样式未找到 | 确认样式名称正确 |
| [4002] | 样式错误 | 无法应用样式 | 检查样式定义和文档兼容性 |
| [5001] | 格式化错误 | 格式化错误 | 检查格式参数是否符合要求 |
| [6001] | 图像错误 | 图像未找到 | 确认图像文件路径正确 |
| [6002] | 图像错误 | 无效的图像格式 | 使用支持的图像格式（JPG、PNG等） |
| [6003] | 图像错误 | 无法加载图像 | 检查图像文件完整性 |
| [7001] | 表格错误 | 表格操作错误 | 检查表格结构是否完整 |
| [8001] | 注释错误 | 注释操作错误 | 检查文档是否包含注释组件 |

这些错误消息通常会以JSON格式返回，例如:
```json
{"error": "Element not found", "code": 2002}
```

### 核心工具列表及参数说明

#### 文档管理工具

```python
# 打开Word文档，这是使用其他工具的前提
open_document(file_path: str) -> str
# 参数：file_path - .docx文件的绝对路径
# 返回值：确认消息或错误信息

# 关闭文档并关闭Word应用实例
shutdown_word() -> str
# 返回值：关闭结果确认消息

# 获取文档结构（所有标题）
get_document_structure() -> List[Dict[str, Any]]
# 返回值：标题列表，每项包含text和level

# 获取文档中所有可用样式
get_document_styles() -> str
# 返回值：包含样式信息的JSON字符串，每项包含name和type

# 接受文档中的所有修订
accept_all_changes() -> str
# 返回值：操作结果确认消息

# 开启文档修订模式
enable_track_revisions() -> str
# 返回值：操作结果确认消息

# 关闭文档修订模式
disable_track_revisions() -> str
# 返回值：操作结果确认消息
```

#### 内容操作工具

```python
# 插入新段落
insert_paragraph(locator: Dict[str, Any], text: str, position: str = "after", style: str = None) -> str
# 参数：
#   locator - 定位锚点元素的定位器
#   text - 要插入的段落文本
#   position - 相对于锚点的位置（"before"或"after"）
#   style - 可选，指定段落样式名称（如"标题 1"、"正文"等）
# 返回值：操作结果确认消息

# 获取指定元素的文本或指定范围的文本
get_text(locator: Dict[str, Any] = None, start_pos: int = None, end_pos: int = None) -> str
# 功能：从所有由定位器找到的元素中检索文本内容，或从文档的特定位置范围获取文本内容，支持获取段落、标题、列表项等多种元素类型的文本
# 参数：
#   locator - 可选，定位目标元素的定位器，可使用多种过滤条件精确定位（详见Locator使用详解）
#   start_pos - 可选，文本范围的起始位置（整数）
#   end_pos - 可选，文本范围的结束位置（整数）
# 返回值：
#   - 成功时：返回格式为 "Success: '{提取的文本}'" 的字符串，其中包含提取的文本内容
#   - 失败时：返回错误信息字符串，如 "Error: No active document. Please use 'open_document' first."、"Error: Invalid range positions. start_pos must be >= 0 and end_pos must be > start_pos." 或 "An unexpected error occurred: {具体错误信息}"
# 注意事项：
#   - 使用前必须先调用open_document打开文档
#   - 如果定位器匹配多个元素，将返回所有匹配元素的文本内容
#   - 若要按范围获取文本，必须同时提供start_pos和end_pos参数
#   - start_pos必须大于等于0，end_pos必须大于start_pos
# 使用示例：
# ```python
# # 调用示例1 - 使用定位器获取第一个段落的文本
# response = await mcp_client.call_tool(
#     "get_text",
#     {
#         "locator": {
#             "type": "paragraph",
#             "index": 0
#         }
#     }
# )
# 
# # 调用示例2 - 使用范围参数获取文本（从位置100到200）
# response = await mcp_client.call_tool(
#     "get_text",
#     {
#         "start_pos": 100,
#         "end_pos": 200
#     }
# )
# 
# # 调用示例3 - 错误用法（未提供必需参数）
# response = await mcp_client.call_tool(
#     "get_text",
#     {
#         "locator": {
#             "target": {
#                 "type": "paragraph",
#                 "filters": [{"type": "index_in_parent", "value": 0}]
#             }
#         }
#     }
# )
# ```

# 在文档中查找文本
find_text(find_text: str, match_case: bool = False, match_whole_word: bool = False, match_wildcards: bool = False, match_synonyms: bool = False, ignore_punct: bool = False, ignore_space: bool = False) -> str
# 功能：在活动文档中查找指定文本的所有出现位置，并返回详细的匹配信息
# 参数：
#   find_text - 要搜索的文本内容
#   match_case - 是否区分大小写（默认：False）
#   match_whole_word - 是否仅匹配完整单词（默认：False）
#   match_wildcards - 是否允许使用通配符字符（默认：False）
#   match_synonyms - 是否匹配同义词（默认：False）【注意：此参数目前不受支持】
#   ignore_punct - 是否忽略标点符号差异（默认：False）
#   ignore_space - 是否忽略空格差异（默认：False）
# 返回值：
#   - 成功时：返回JSON格式字符串，包含找到的匹配项数量和每个匹配项的详细信息（位置、文本、段落索引、上下文预览）
#   - 失败时：返回错误信息字符串，如 "Error: No active document. Please use 'open_document' first." 或 "Error: Search text cannot be empty." 或 "An unexpected error occurred during text search: {具体错误信息}"
# 注意事项：
#   - 使用前必须先调用open_document打开文档
#   - 搜索文本不能为空
#   - 返回的JSON结构包含matches_found（匹配数量）和matches（匹配详情数组）
#   - match_synonyms参数目前不受支持，设置后不会产生任何效果
# 使用示例：
# ```python
# # 调用示例 - 查找文档中的特定文本，不区分大小写
# response = await mcp_client.call_tool(
#     "find_text",
#     {
#         "find_text": "example",
#         "match_case": False,
#         "match_whole_word": True
#     }
# )
# # 解析返回的JSON结果
# import json
# result = json.loads(response)
# print(f"找到 {result['matches_found']} 个匹配项")
# for match in result['matches']:
#     print(f"第{match['index']}个匹配: '{match['text']}'，位于段落{match['paragraph_index']}")
# ```

# 替换指定元素的文本
replace_text(locator: Dict[str, Any], new_text: str) -> str
# 参数：
#   locator - 定位目标元素的定位器
#   new_text - 替换的新文本
# 返回值：操作结果确认消息

# 删除定位元素
delete_element(locator: Dict[str, Any]) -> str
# 参数：locator - 定位目标元素的定位器
# 返回值：
#   - 成功时：返回格式为 "Element(s) deleted successfully" 的确认消息
#   - 失败时：返回标准化错误JSON，包含错误代码和描述，如 {"error": "Element not found", "code": 2002} 或 {"error": "Operation failed", "code": 4001}
# 可能的错误代码：[1001] 参数错误、[2001] 文档错误、[2002] 元素未找到、[3001] 定位器歧义、[4001] 操作失败、[4003] COM组件错误

# 添加评论到指定位置
add_comment(locator: Dict[str, Any], text: str) -> str
# 功能：在由定位器指定的文档位置添加评论
# 参数：
#   locator - 定位目标元素的定位器
#   text - 评论内容文本
# 返回值：
#   - 成功时：返回格式为 "Comment added successfully" 的确认消息
#   - 失败时：返回错误信息字符串，如 "Error: No active document" 或 "An unexpected error occurred" 

# 获取文档中的所有评论
get_comments() -> str
# 功能：获取文档中的所有评论
# 返回值：
#   - 成功时：包含所有评论信息的JSON字符串，每项包含评论的文本、作者、位置等信息
#   - 失败时：返回标准化错误JSON，包含错误代码和描述，如 {"error": "No active document available", "code": 2001} 或 {"error": "Comments access error", "code": 4004}
# 可能的错误代码：[1001] 参数错误、[2001] 文档错误、[4004] 特殊组件错误

# 删除指定评论
delete_comment(comment_id: int) -> str
# 功能：删除文档中指定ID的评论
# 参数：
#   comment_id - 评论的唯一标识符
# 返回值：
#   - 成功时：返回格式为 "Comment deleted successfully" 的确认消息
#   - 失败时：返回错误信息字符串，如 "Error: Comment not found" 或 "An unexpected error occurred" 

# 删除文档中的所有评论
delete_all_comments() -> str
# 功能：删除文档中的所有评论
# 返回值：
#   - 成功时：返回格式为 "All comments deleted successfully" 的确认消息
#   - 失败时：返回错误信息字符串

# 创建表格
create_table(locator: Dict[str, Any], rows: int, cols: int) -> str
# 参数：
#   locator - 定位表格插入位置的定位器
#   rows - 表格行数
#   cols - 表格列数
# 返回值：操作结果确认消息

# 创建项目符号列表
create_bulleted_list(locator: Dict[str, Any], items: List[str], position: str = "after") -> str
# 参数：
#   locator - 定位列表插入位置的定位器
#   items - 列表项内容列表
#   position - 相对于锚点的位置
# 返回值：操作结果确认消息

# 获取图片信息
get_image_info(locator: Dict[str, Any] = None) -> str
# 参数：
#   locator - 可选，定位特定图片的定位器，如果未提供，则返回所有图片
# 返回值：
#   - 成功时：包含图片信息的JSON字符串，每项包含图片的索引、类型、尺寸等信息
#   - 失败时：返回标准化错误JSON，包含错误代码和描述，如 {"error": "No active document available", "code": 2001} 或 {"error": "Element not found", "code": 2002}
# 可能的错误代码：[1001] 参数错误、[2001] 文档错误、[2002] 元素未找到、[3001] 定位器歧义、[4003] COM组件错误

# 插入嵌入式图片
insert_inline_picture(locator: Dict[str, Any], image_path: str, position: str = "after") -> str
# 参数：
#   locator - 定位图片插入位置的定位器
#   image_path - 图片文件的绝对路径
#   position - 相对于锚点的位置（"before"或"after"）
# 返回值：操作结果确认消息
```

#### 表格操作工具

```python
# 获取表格单元格文本
get_text_from_cell(locator: Dict[str, Any]) -> str
# 参数：locator - 定位表格单元格的定位器
# 返回值：
#   - 成功时：单元格文本内容
#   - 失败时：返回标准化错误信息，包含错误代码和描述
# 可能的错误代码：[1001] 参数错误、[2001] 文档错误、[2002] 元素未找到、[2003] 定位器歧义、[7001] 表格错误

# 设置表格单元格值
set_cell_value(locator: Dict[str, Any], text: str) -> str
# 参数：
#   locator - 定位表格单元格的定位器
#   text - 要设置的单元格文本
# 返回值：
#   - 成功时：操作结果确认消息
#   - 失败时：返回标准化错误信息，包含错误代码和描述
# 可能的错误代码：[1001] 参数错误、[2001] 文档错误、[2002] 元素未找到、[2003] 定位器歧义、[7001] 表格错误

# 创建表格
create_table(locator: Dict[str, Any], rows: int, cols: int) -> str
# 参数：
#   locator - 定位表格插入位置的定位器
#   rows - 表格行数（必须为正整数）
#   cols - 表格列数（必须为正整数）
# 返回值：
#   - 成功时：操作结果确认消息
#   - 失败时：返回标准化错误信息，包含错误代码和描述
# 可能的错误代码：[1001] 参数错误、[2001] 文档错误、[2002] 元素未找到、[2003] 定位器歧义、[7001] 表格错误
```

#### 页眉页脚工具

```python
# 设置文档页眉文本
set_header_text(text: str) -> str
# 参数：text - 页眉文本
# 返回值：操作结果确认消息

# 设置文档页脚文本
set_footer_text(text: str) -> str
# 参数：text - 页脚文本
# 返回值：操作结果确认消息
```

#### 格式设置工具

```python
# 应用格式到指定元素
apply_format(locator: Dict[str, Any], formatting: Dict[str, Any]) -> str
# 参数：
#   locator - 定位目标元素的定位器
#   formatting - 格式设置字典，如{"bold": True, "alignment": "center", "paragraph_style": "标题 1"}
# 返回值：
#   - 成功时：返回格式为 "Formatting applied successfully" 的确认消息
#   - 失败时：返回标准化错误JSON，包含错误代码和描述，如 {"error": "Invalid formatting specified", "code": 4002} 或 {"error": "Element not found", "code": 2002}
# 可能的错误代码：[1001] 参数错误、[2001] 文档错误、[2002] 元素未找到、[3001] 定位器歧义、[4001] 操作失败、[4002] 格式错误、[4003] COM组件错误

# 应用段落样式到指定元素（更稳定的样式应用方式）
apply_paragraph_style(locator: Dict[str, Any], style_name: str) -> str
# 参数：
#   locator - 定位目标元素的定位器
#   style_name - 要应用的段落样式名称
# 返回值：操作结果确认消息

# 批量应用格式到多个元素（高效的批量操作方式）
batch_apply_format(operations: List[Dict[str, Any]], save_document: bool = True) -> str
# 参数：
#   operations - 操作列表，每个操作包含locator和formatting键
#   save_document - 是否在应用所有格式后保存文档（默认：True）
# 返回值：批量操作结果摘要
```

### Locator使用详解

Locator是一个特殊的查询对象，用于精确定位Word文档中的元素。它支持多种定位方式和过滤条件，使AI大模型能够精确地找到需要操作的文档部分。

#### 基本结构

Locator是一个字典，必须包含`target`字段，`target`内部包含以下组件：
- `type`: 元素类型（如"paragraph", "table", "cell", "heading"等）
- `filters`: 过滤条件列表（可选）

对于相对定位，Locator还可以包含：
- `anchor`: 锚点元素的定位器（包含`type`和`filters`）
- `relation`: 相对位置描述

#### 支持的元素类型

- `paragraph`: 段落
- `table`: 表格
- `cell`: 表格单元格
- `heading`: 标题
- `image`: 图片
- `list`: 列表

#### 常用过滤器

- `contains_text`: 文本包含指定内容
- `text_matches_regex`: 文本匹配正则表达式
- `index_in_parent`: 元素在父容器中的索引位置
- `style`: 元素样式名称
- `is_bold`: 是否为粗体
- `row_index`: 表格行索引
- `column_index`: 表格列索引
- `is_list_item`: 是否为列表项

#### 相对定位

通过`relation`字段可以指定元素相对于其他元素的位置，如：
- `all_occurrences_within`
- `first_occurrence_after`
- `last_occurrence_before`

### 使用示例

下面是AI大模型客户端调用MCP工具的典型示例：

#### 1. 打开文档

```python
# 调用示例
response = await mcp_client.call_tool(
    "open_document",
    {
        "file_path": "C:/Users/username/Documents/report.docx"
    }
)
print(response)  # "Active document set to: C:/Users/username/Documents/report.docx"
```

#### 2. 读取文档结构

```python
# 获取文档标题结构
headings = await mcp_client.call_tool("get_document_structure", {})
print(headings)
# 输出示例: [{"text": "Introduction", "level": 1}, {"text": "Methods", "level": 1}, ...]
```

#### 3. 插入段落

```python
# 在文档开头插入段落（使用默认样式）
response = await mcp_client.call_tool(
    "insert_paragraph",
    {
        "locator": {"target": {"type": "paragraph", "filters": [{"index_in_parent": 0}]}},
        "text": "This is a new paragraph at the beginning of the document.",
        "position": "before"
    }
)
print(response)  # "Successfully inserted paragraph."

# 在文档开头插入具有指定样式的段落
response = await mcp_client.call_tool(
    "insert_paragraph",
    {
        "locator": {"target": {"type": "paragraph", "filters": [{"index_in_parent": 0}]}},
        "text": "这是一个标题样式的段落",
        "position": "before",
        "style": "标题 1"
    }
)
print(response)  # "Successfully inserted paragraph with style."
```

#### 4. 查找并替换文本

```python
# 查找包含特定文本的段落并替换
response = await mcp_client.call_tool(
    "replace_text",
    {
        "locator": {"target": {"type": "paragraph", "filters": [{"contains_text": "old information"}]}},
        "new_text": "Updated content with new information."
    }
)
print(response)  # "Successfully replaced text."
```

#### 5. 格式化文本

```python
# 将第一段设置为粗体并居中对齐
response = await mcp_client.call_tool(
    "apply_format",
    {
        "locator": {"target": {"type": "paragraph", "filters": [{"index_in_parent": 0}]}},
        "formatting": {"bold": True, "alignment": "center"}
    }
)
print(response)  # "Formatting applied successfully."

# 使用专用工具应用段落样式（更稳定的方法）
response = await mcp_client.call_tool(
    "apply_paragraph_style",
    {
        "locator": {"target": {"type": "paragraph", "filters": [{"index_in_parent": 1}]}},
        "style_name": "标题 2"
    }
)
print(response)  # "Successfully applied paragraph style: 标题 2"

# 批量应用不同的格式设置（提高效率）
response = await mcp_client.call_tool(
    "batch_apply_format",
    {
        "operations": [
            {
                "locator": {"target": {"type": "paragraph", "filters": [{"index_in_parent": 0}]}},
                "formatting": {"paragraph_style": "标题 1"}
            },
            {
                "locator": {"target": {"type": "paragraph", "filters": [{"index_in_parent": 1}]}},
                "formatting": {"paragraph_style": "标题 2", "italic": True}
            },
            {
                "locator": {"target": {"type": "paragraph", "filters": [{"index_in_parent": 2}]}},
                "formatting": {"alignment": "justify", "font_size": 12}
            }
        ],
        "save_document": True
    }
)
print(response)  # "Batch formatting completed: 3 successful, 0 failed out of 3 operations."
```

#### 6. 操作表格

```python
# 在文档末尾创建一个3x4的表格
response = await mcp_client.call_tool(
    "create_table",
    {
        "locator": {"target": {"type": "paragraph", "filters": [{"index_in_parent": -1}]}},  # 最后一个段落
        "rows": 3,
        "cols": 4
    }
)
print(response)  # "Successfully created table."

# 设置表格单元格值
response = await mcp_client.call_tool(
    "set_cell_value",
    {
        "locator": {
            "target": {
                "type": "cell", 
                "filters": [
                    {"table_index": 0},
                    {"row_index": 0},
                    {"column_index": 0}
                ]
            }
        },
        "text": "Header 1"
    }
)
print(response)  # "Successfully set cell value."

# 获取表格单元格文本
response = await mcp_client.call_tool(
    "get_text_from_cell",
    {
        "locator": {
            "target": {
                "type": "cell", 
                "filters": [
                    {"table_index": 0},
                    {"row_index": 0},
                    {"column_index": 0}
                ]
            }
        }
    }
)
print(response)  # "Header 1"

# 表格操作错误处理示例
response = await mcp_client.call_tool(
    "set_cell_value",
    {
        "locator": {
            "target": {
                "type": "cell", 
                "filters": [
                    {"table_index": 999},  # 不存在的表格
                    {"row_index": 0},
                    {"column_index": 0}
                ]
            }
        },
        "text": "Test"
    }
)
# 可能的错误响应: "Error [2002]: No table cell found matching the locator: Table index 999 not found"

# 表格创建参数验证错误示例
response = await mcp_client.call_tool(
    "create_table",
    {
        "locator": {"target": {"type": "paragraph", "filters": [{"index_in_parent": -1}]}},
        "rows": 0,  # 无效的行数
        "cols": 4
    }
)
# 可能的错误响应: "Error [1001]: Invalid row count. Must be a positive integer."

# COM错误处理示例（表格操作失败）
response = await mcp_client.call_tool(
    "set_cell_value",
    {
        "locator": {
            "target": {
                "type": "cell", 
                "filters": [
                    {"table_index": 0},
                    {"row_index": 0},
                    {"column_index": 0}
                ]
            }
        },
        "text": "Test"
    }
)
# 可能的错误响应: "Error [7001]: Failed to update table cell. This may occur if the document structure is corrupted, the table is protected, or Word is in an unstable state. Try closing and reopening the document."
```

#### 7. 文档修订模式控制

```python
# 开启文档修订模式
response = await mcp_client.call_tool("enable_track_revisions", {})
print(response)  # "Revision tracking enabled successfully."

# 在修订模式下进行编辑操作
response = await mcp_client.call_tool(
    "replace_text",
    {
        "locator": {"target": {"type": "paragraph", "filters": [{"contains_text": "draft content"}]}},
        "new_text": "Final content"
    }
) # 此更改会被标记为修订

# 接受所有修订
response = await mcp_client.call_tool("accept_all_changes", {})
print(response)  # "All changes accepted successfully."

# 关闭文档修订模式
response = await mcp_client.call_tool("disable_track_revisions", {})
print(response)  # "Revision tracking disabled successfully."

#### 8. 设置页眉页脚

```python
# 设置页眉文本
response = await mcp_client.call_tool(
    "set_header_text",
    {"text": "Confidential Report - 2023"}
)
print(response)  # "Header text set successfully."

# 设置页脚文本
response = await mcp_client.call_tool(
    "set_footer_text",
    {"text": "Page [Page] of [Pages]"}
)
print(response)  # "Footer text set successfully."
```

#### 9. 关闭文档

```python
# 关闭Word应用
response = await mcp_client.call_tool("shutdown_word", {})
print(response)  # "Word application shut down successfully."
```

### 最佳实践

1. **会话管理**：在一个会话中完成一组相关操作，但不要自动调用`shutdown_word`工具，应由人类用户自己接受修订和关闭文档
2. **修订模式**：文档编辑前打开修订模式，以便追踪和管理所有修改
3. **错误处理**：检查工具返回值是否包含错误代码，根据错误类型采取相应措施。特别注意处理[4003]、[4004]、[7001]、[8001]等COM组件错误，可能需要重启Word应用
4. **定位策略**：使用精确的locator避免误操作，特别是在批量操作时
5. **路径处理**：确保提供的文件路径是绝对路径
6. **文档检查**：使用`open_document`前确认文件存在且格式正确
7. **样式应用**：对于段落样式应用，优先使用`apply_paragraph_style`工具，提供更稳定的样式应用体验
8. **批量操作**：当需要对多个元素应用不同格式时，使用`batch_apply_format`工具提高效率
9. **样式验证**：在应用样式前，可以使用`get_document_styles`获取文档中可用的样式列表，避免使用不存在的样式
10. **错误重试**：对于临时性错误，实现适当的重试机制
11. **组件访问**：当访问特定Word组件（如Paragraphs、Comments、InlineShapes、Tables）时，确保文档包含相应组件并使用适当的访问方式
12. **表格操作**：
    - 在操作表格前，确保表格存在且结构完整
    - 使用精确的表格索引、行索引和列索引定位单元格
    - 处理表格操作时，注意文档保护状态和表格锁定状态
    - 对于[7001]表格错误，建议关闭并重新打开文档后重试
13. **参数验证**：
    - 在调用工具前验证参数的有效性（如行数、列数必须为正整数）
    - 使用适当的错误代码帮助用户快速定位问题
14. **异常处理**：
    - 实现嵌套异常处理结构，区分不同类型的错误
    - 对COM相关错误提供具体的解决建议
    - 保持错误消息的一致性和可操作性