# Office Word MCP Server 综合工具使用指南

## 1. 工具优化背景

为了响应"暴露给大模型的tool函数数量应充分精简，且充分说明，宁肯重复的注释和说明，也不要增加tool的数量。单工具应该足够强大"的要求，我们对原有的20多个独立工具函数进行了优化，合并为5个功能强大的综合工具函数。

**优化目标：**
- **减少工具数量**：从原来的20多个工具减少到5个综合工具
- **保持功能强大**：每个工具能够处理多种相关操作
- **使用更加简便**：用户只需学习少量工具的使用
- **文档更加完整**：每个工具都有详细的注释和说明
- **维护更加容易**：减少了代码重复，提高了可维护性

## 2. 综合工具概述

| 综合工具函数 | 功能描述 | 替换的原工具数量 |
|------------|---------|---------------|
| `document_operation` | 处理所有文档级操作 | 5个文档工具 |
| `text_operation` | 处理所有文本相关操作 | 8个文本工具 |
| `table_operation` | 处理所有表格相关操作 | 3个表格工具 |
| `image_operation` | 处理所有图片相关操作 | 3个图片工具 |
| `comment_operation` | 处理所有注释相关操作 | 6个注释工具 |

## 3. 综合工具详细说明

### 3.1 document_operation - 综合文档操作工具

处理所有文档级操作，如打开文档、关闭文档、获取文档信息等。

**支持的操作类型：**

| 操作类型 | 描述 | 必填参数 | 可选参数 | 返回值 |
|---------|------|---------|---------|-------|
| `open` | 打开Word文档 | `file_path` | - | 文档信息字典 |
| `close` | 关闭活动文档 | - | - | 文档关闭成功信息 |
| `shutdown` | 关闭Word应用程序 | - | - | Word应用程序关闭成功信息 |
| `get_styles` | 获取文档样式 | - | - | 包含样式名称和类型的JSON字符串 |
| `get_objects` | 获取特定类型元素 | `object_type` | - | 包含元素信息的JSON字符串 |

**使用示例：**

```python
# 打开文档
document_operation(operation_type="open", file_path="C:\path\to\document.docx")

# 获取文档样式
document_operation(operation_type="get_styles")

# 获取文档中的表格
document_operation(operation_type="get_objects", object_type="tables")

# 关闭文档
document_operation(operation_type="close")

# 关闭Word应用程序
document_operation(operation_type="shutdown")
```

### 3.2 text_operation - 综合文本操作工具

处理所有文本相关操作，如获取文本、插入文本、应用格式化、查找替换等。

**支持的操作类型：**

| 操作类型 | 描述 | 必填参数 | 可选参数 | 返回值 |
|---------|------|---------|---------|-------|
| `get` | 获取文档中的文本 | - | `locator` | 请求的文本 |
| `insert` | 插入文本或段落 | `text` | `locator`, `position`, `style` | 插入成功信息 |
| `apply_formatting` | 应用格式化（支持单格式和多格式） | `locator`, (`formatting` 或 `format_type` 和 `format_value`) | - | 格式化成功信息 |
| `batch_apply_format` | 批量应用格式化 | `operations` | `save_document` | 批量操作结果摘要 |
| `find` | 查找文本 | `text` | `match_case`, `match_whole_word`, `match_wildcards`, `ignore_punct`, `ignore_space` | 包含查找结果的JSON字符串 |
| `replace` | 替换文本 | `locator`, `text` | - | 替换成功信息和替换次数 |
| `create_list` | 创建项目符号列表 | `locator`, `items` | `position` | 创建列表成功信息 |

**使用示例：**

```python
# 获取所有文本
text_operation(operation_type="get")

# 插入文本
text_operation(operation_type="insert", 
              locator={"type": "paragraph", "value": 3}, 
              text="新插入的文本", 
              position="after")

# 应用格式化
text_operation(operation_type="apply_formatting", 
              locator={"type": "paragraph", "value": 3}, 
              formatting={"bold": True, "font_size": 12})

# 查找文本
text_operation(operation_type="find", 
              text="关键词", 
              match_case=True)

# 替换文本
text_operation(operation_type="replace", 
              locator={"type": "paragraph", "value": 3}, 
              text="替换后的文本")

# 批量应用格式化
text_operation(operation_type="batch_apply_format", 
              operations=[
                  {"locator": {"type": "paragraph", "value": 1}, "formatting": {"bold": True}},
                  {"locator": {"type": "paragraph", "value": 2}, "formatting": {"italic": True}}
              ])

# 创建项目符号列表
text_operation(operation_type="create_list", 
              locator={"type": "paragraph", "value": 5}, 
              items=["列表项1", "列表项2", "列表项3"])
```

### 3.3 table_operation - 综合表格操作工具

处理所有表格相关操作，如创建表格、获取/设置单元格内容等。

**支持的操作类型：**

| 操作类型 | 描述 | 必填参数 | 可选参数 | 返回值 |
|---------|------|---------|---------|-------|
| `create` | 创建新表格 | `locator`, `rows`, `cols` | - | 表格创建成功信息 |
| `get_cell_text` | 获取表格单元格文本 | `locator` | - | 单元格的文本内容 |
| `set_cell_text` | 设置表格单元格文本 | `locator`, `text` | - | 单元格值设置成功信息 |

**使用示例：**

```python
# 创建表格
table_operation(operation_type="create", 
               locator={"type": "paragraph", "value": 3}, 
               rows=3, 
               cols=4)

# 获取单元格文本
table_operation(operation_type="get_cell_text", 
               locator={"type": "cell", "value": "1,2"})

# 设置单元格文本
table_operation(operation_type="set_cell_text", 
               locator={"type": "cell", "value": "1,2"}, 
               text="新的单元格内容")
```

### 3.4 image_operation - 综合图片操作工具

处理所有图片相关操作，如获取图片信息、插入图片、添加题注等。

**支持的操作类型：**

| 操作类型 | 描述 | 必填参数 | 可选参数 | 返回值 |
|---------|------|---------|---------|-------|
| `get_info` | 获取所有图片的信息 | - | - | 包含所有图片信息的JSON字符串 |
| `insert` | 插入图片或对象 | `locator`, `object_path` | `object_type`, `position` | 图片插入成功信息 |
| `add_caption` | 为图片添加题注 | `locator`, `caption_text` | `label`, `position` | 题注添加成功信息 |

**使用示例：**

```python
# 获取所有图片信息
image_operation(operation_type="get_info")

# 插入图片
image_operation(operation_type="insert", 
               locator={"type": "paragraph", "value": 3}, 
               object_path="C:\path\to\image.jpg")

# 为图片添加题注
image_operation(operation_type="add_caption", 
               locator={"type": "image", "value": 1}, 
               caption_text="图片描述", 
               label="Figure")
```

### 3.5 comment_operation - 综合注释操作工具

处理所有注释相关操作，如添加注释、获取注释、删除注释、回复注释等。

**支持的操作类型：**

| 操作类型 | 描述 | 必填参数 | 可选参数 | 返回值 |
|---------|------|---------|---------|-------|
| `add` | 添加注释 | `locator`, `text` | `author` | 注释添加成功信息和注释ID |
| `get_all` | 获取所有注释 | - | - | 包含所有注释信息的JSON字符串 |
| `delete` | 删除注释 | - | `comment_index` (不提供则删除所有) | 注释删除成功信息 |
| `edit` | 编辑注释 | `comment_index`, `text` | - | 注释编辑成功信息 |
| `reply` | 回复注释 | `comment_index`, `text` | `author` | 回复添加成功信息 |
| `get_thread` | 获取注释线程 | `comment_index` | - | 包含注释线程信息的JSON字符串 |

**使用示例：**

```python
# 添加注释
comment_operation(operation_type="add", 
                 locator={"type": "paragraph", "value": 3}, 
                 text="这是一条注释")

# 获取所有注释
comment_operation(operation_type="get_all")

# 回复注释
comment_operation(operation_type="reply", 
                 comment_index=0, 
                 text="这是回复")

# 删除指定注释
comment_operation(operation_type="delete", 
                 comment_index=0)

# 删除所有注释
comment_operation(operation_type="delete")

# 获取注释线程
comment_operation(operation_type="get_thread", 
                 comment_index=0)
```

## 4. 定位器(Locator)使用指南

大多数操作都需要使用**定位器(Locator)**来指定操作的目标元素。定位器是一个字典对象，格式如下：

```python
locator = {"type": "object_type", "value": "object_identifier", "filters": {"filter_name": "filter_value"}}
```

### 支持的元素类型

| 元素类型 | 描述 | 值的格式 | 示例 |
|---------|------|---------|------|
| `document` | 整个文档 | 不需要值 | `{"type": "document"}` |
| `paragraph` | 段落 | 段落索引(从1开始) | `{"type": "paragraph", "value": 3}` |
| `table` | 表格 | 表格索引(从1开始) | `{"type": "table", "value": 2}` |
| `cell` | 表格单元格 | "行号,列号"(从1开始) | `{"type": "cell", "value": "1,2"}` |
| `image` | 图片 | 图片索引(从1开始) | `{"type": "image", "value": 1}` |
| `comment` | 注释 | 注释索引(从0开始) | `{"type": "comment", "value": 0}` |
| `text` | 文本范围 | 文本内容 | `{"type": "text", "value": "关键词"}` |

### 过滤器的使用

过滤器可以帮助更精确地定位元素，例如：

```python
# 查找包含特定文本的段落
locator = {"type": "paragraph", "filters": {"contains_text": "重要信息"}}

# 查找特定标题样式的段落
locator = {"type": "paragraph", "filters": {"style_name": "Heading 1"}}
```

## 5. 综合工具使用流程示例

下面是一个完整的文档操作流程示例，展示如何使用综合工具完成从打开文档到保存关闭的整个过程：

```python
import json

# 1. 打开文档
doc_info = document_operation(operation_type="open", file_path="C:\path\to\document.docx")
print(f"文档已打开: {json.dumps(doc_info, ensure_ascii=False)}")

# 2. 插入标题
text_operation(operation_type="insert", 
              locator={"type": "document"}, 
              text="综合工具使用示例", 
              position="after",
              style="Heading 1")

# 3. 插入正文段落
text_operation(operation_type="insert", 
              locator={"type": "paragraph", "value": 1}, 
              text="这是使用综合工具创建的第一个段落。", 
              position="after")

# 4. 应用格式化
text_operation(operation_type="apply_formatting", 
              locator={"type": "paragraph", "value": 2}, 
              formatting={"font_size": 12, "font_name": "Arial", "alignment": "justify"})

# 5. 创建表格
table_operation(operation_type="create", 
               locator={"type": "paragraph", "value": 2}, 
               rows=2, 
               cols=3, 
               position="after")

# 6. 填写表格
for i in range(1, 3):
    for j in range(1, 4):
        table_operation(operation_type="set_cell_text", 
                       locator={"type": "cell", "value": f"{i},{j}"}, 
                       text=f"单元格{i}-{j}")

# 7. 插入图片
image_operation(operation_type="insert", 
               locator={"type": "table", "value": 1}, 
               object_path="C:\path\to\image.jpg", 
               position="after")

# 8. 为图片添加题注
image_operation(operation_type="add_caption", 
               locator={"type": "image", "value": 1}, 
               caption_text="示例图片", 
               label="Figure")

# 9. 添加注释
comment_operation(operation_type="add", 
                 locator={"type": "paragraph", "value": 2}, 
                 text="这是一个示例注释")

# 10. 查找文本
results = text_operation(operation_type="find", text="示例")
print(f"找到 {json.loads(results)['matches_found']} 个匹配项")

# 11. 关闭文档
document_operation(operation_type="close")

# 12. 关闭Word应用程序
document_operation(operation_type="shutdown")
```

## 6. 错误处理和故障排除

### 常见错误及解决方案

| 错误类型 | 错误消息示例 | 解决方案 |
|---------|------------|---------|
| 文档未找到 | File not found: C:\path\to\document.docx | 检查文件路径是否正确 |
| 没有活动文档 | No active document found | 先使用 `document_operation("open", file_path="...")` 打开文档 |
| 元素未找到 | No object found matching the locator | 检查定位器参数是否正确，可能文档结构与预期不符 |
| 参数无效 | Invalid parameter: text is required | 确保提供了所有必需的参数 |
| COM错误 | Failed to open document: ... | 检查Word应用程序是否正常运行，可能需要重启Word |

### 调试技巧

1. 使用`log_info`和`log_error`函数记录操作过程
2. 检查`word_doc_server.log`文件获取详细的错误日志
3. 对于定位问题，可以先使用`document_operation("get_objects")`查看文档结构
4. 确保所有路径使用双反斜杠(`\\`)或正斜杠(`/`)，避免使用单反斜杠

## 7. 总结

综合工具函数是Office Word MCP Server的推荐API，它们提供了精简而强大的接口，使文档操作更加简单和一致。通过掌握这些工具的使用，您可以更高效地开发基于Word的应用程序和自动化工作流程。

如果您有任何问题或建议，请随时联系开发团队。