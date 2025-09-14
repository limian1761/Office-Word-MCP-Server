# Office Word MCP Server 综合工具使用指南

## 1. 工具优化背景

为了响应"暴露给大模型的tool函数数量应充分精简，且充分说明，宁肯重复的注释和说明，也不要增加tool的数量。单工具应该足够强大"的要求，我们对原有的20多个独立工具函数进行了优化，合并为6个功能强大的综合工具函数。

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
| `navigate_tools` | 综合导航工具 | 新增工具 |
| `view_control_tools` | 处理所有视图控制相关操作 | 7个视图控制工具 |

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

### 3.6 navigate_tools - 综合导航工具

管理文档上下文和对象的选择，用于设置活动文档、活动上下文和活动对象，这是当前推荐的定位方式，替代了传统的定位器(Locator)机制，使得其他工具可以在不需要指定定位参数的情况下进行操作。

**支持的操作类型：**

| 操作类型 | 描述 | 必填参数 | 可选参数 | 返回值 |
|---------|------|---------|---------|-------|
| `set_active_context` | 设置活动上下文 | `context_type`, `context_value` | - | 上下文设置成功信息 |
| `set_active_object` | 设置活动对象 | `object_type`, `object_value` | - | 对象设置成功信息 |
| `get_active_context` | 获取当前活动上下文信息 | - | - | 当前活动上下文信息JSON字符串 |
| `get_active_object` | 获取当前活动对象信息 | - | - | 当前活动对象信息JSON字符串 |

**使用示例：**

```python
# 设置活动上下文为特定章节
navigate_tools(operation_type="set_active_context", context_type="section", context_value=2)

# 设置活动对象为特定段落
navigate_tools(operation_type="set_active_object", object_type="paragraph", object_value=5)

# 获取当前活动上下文
current_context = navigate_tools(operation_type="get_active_context")

# 获取当前活动对象
current_object = navigate_tools(operation_type="get_active_object")
```

**支持的上下文类型：**
- `document`: 整个文档
- `section`: 文档章节
- `heading`: 标题
- `bookmark`: 书签
- `table`: 表格
- `selection`: 当前选择

**支持的对象类型：**
- `paragraph`: 段落
- `table`: 表格
- `cell`: 表格单元格
- `image`: 图片
- `comment`: 注释

### 3.7 view_control_tools - 综合视图控制工具

处理所有视图控制相关操作，如切换视图、调整缩放比例、显示/隐藏元素等。

**支持的操作类型：**

| 操作类型 | 描述 | 必填参数 | 可选参数 | 返回值 |
|---------|------|---------|---------|-------|
| `switch_view` | 切换文档视图 | `view_type` | - | 视图切换成功信息 |
| `set_zoom` | 设置文档缩放比例 | `zoom_level` | - | 缩放设置成功信息 |
| `show_element` | 显示特定元素 | `element_type` | - | 元素显示成功信息 |
| `hide_element` | 隐藏特定元素 | `element_type` | - | 元素隐藏成功信息 |
| `toggle_element` | 切换特定元素的显示状态 | `element_type` | - | 元素状态切换成功信息 |
| `get_view_info` | 获取当前视图信息 | - | - | 包含当前视图信息的JSON字符串 |
| `navigate` | 导航到文档特定位置 | `navigation_type`, `value` | - | 导航成功信息 |

**使用示例：**

```python
# 切换到阅读视图
view_control_tools(operation_type="switch_view", view_type="read")

# 设置缩放比例为150%
view_control_tools(operation_type="set_zoom", zoom_level=150)

# 显示网格线
view_control_tools(operation_type="show_element", element_type="gridlines")

# 隐藏标尺
view_control_tools(operation_type="hide_element", element_type="rulers")

# 切换导航窗格显示状态
view_control_tools(operation_type="toggle_element", element_type="navigation_pane")

# 获取当前视图信息
current_view = view_control_tools(operation_type="get_view_info")

# 导航到特定页码
view_control_tools(operation_type="navigate", navigation_type="page", value=5)
```

**支持的视图类型：**
- `print`: 打印视图
- `web`: Web视图
- `read`: 阅读视图
- `outline`: 大纲视图
- `draft`: 草稿视图

**支持的元素类型：**
- `rulers`: 标尺
- `gridlines`: 网格线
- `navigation_pane`: 导航窗格
- `status_bar`: 状态栏
- `task_pane`: 任务窗格
- `comments_pane`: 评论窗格
- `formatting_marks`: 格式标记

**支持的导航类型：**
- `page`: 按页码导航
- `heading`: 按标题导航
- `bookmark`: 按书签导航
- `section`: 按节导航
- `table`: 按表格导航
```

## 4. 综合工具使用流程示例

下面是一个完整的文档操作流程示例，展示如何使用综合工具完成从打开文档到保存关闭的整个过程：

```python
import json

# 1. 打开文档
doc_info = document_operation(operation_type="open", file_path="C:\path\to\document.docx")
print(f"文档已打开: {json.dumps(doc_info, ensure_ascii=False)}")

# 2. 插入标题（使用文档作为默认上下文）
text_operation(operation_type="insert", 
              text="综合工具使用示例", 
              position="after",
              style="Heading 1")

# 3. 设置活动对象为第1段落后，插入正文段落
navigate_tools(operation_type="set_active_object", object_type="paragraph", object_value=1)
text_operation(operation_type="insert", 
              text="这是使用综合工具创建的第一个段落。", 
              position="after")

# 4. 设置活动对象为第2段落后，应用格式化
navigate_tools(operation_type="set_active_object", object_type="paragraph", object_value=2)
text_operation(operation_type="apply_formatting", 
              formatting={"font_size": 12, "font_name": "Arial", "alignment": "justify"})

# 5. 设置活动对象为第2段落后，创建表格
navigate_tools(operation_type="set_active_object", object_type="paragraph", object_value=2)
table_operation(operation_type="create", 
               rows=2, 
               cols=3, 
               position="after")

# 6. 填写表格
for i in range(1, 3):
    for j in range(1, 4):
        # 设置活动对象为目标单元格，然后设置内容
        navigate_tools(operation_type="set_active_object", object_type="cell", object_value=f"{i},{j}")
        table_operation(operation_type="set_cell_text", text=f"单元格{i}-{j}")

# 7. 设置活动对象为第1个表格后，插入图片
navigate_tools(operation_type="set_active_object", object_type="table", object_value=1)
image_operation(operation_type="insert", 
               object_path="C:\path\to\image.jpg", 
               position="after")

# 8. 设置活动对象为第1个图片后，添加题注
navigate_tools(operation_type="set_active_object", object_type="image", object_value=1)
image_operation(operation_type="add_caption", 
               caption_text="示例图片", 
               label="Figure")

# 9. 设置活动对象为第2段落后，添加注释
navigate_tools(operation_type="set_active_object", object_type="paragraph", object_value=2)
comment_operation(operation_type="add", text="这是一个示例注释")

# 10. 查找文本
results = text_operation(operation_type="find", text="示例")
print(f"找到 {json.loads(results)['matches_found']} 个匹配项")

# 11. 关闭文档
document_operation(operation_type="close")

# 12. 关闭Word应用程序
document_operation(operation_type="shutdown")
```

## 5. 错误处理和故障排除

### 常见错误及解决方案

| 错误类型 | 错误消息示例 | 解决方案 |
|---------|------------|---------|
| 文档未找到 | File not found: C:\path\to\document.docx | 检查文件路径是否正确 |
| 没有活动文档 | No active document found | 先使用 `document_operation("open", file_path="...")` 打开文档 |
| 元素未找到 | No object found matching the specified object_type and object_value | 检查指定的对象类型和值是否正确，可能文档结构与预期不符 |
| 参数无效 | Invalid parameter: text is required | 确保提供了所有必需的参数 |
| COM错误 | Failed to open document: ... | 检查Word应用程序是否正常运行，可能需要重启Word |

### 调试技巧

1. 使用`log_info`和`log_error`函数记录操作过程
2. 检查`word_doc_server.log`文件获取详细的错误日志
3. 对于导航问题，可以先使用`document_operation("get_objects")`查看文档结构
4. 使用`navigate_tools("get_active_context")`和`navigate_tools("get_active_object")`检查当前活动上下文和对象
5. 确保所有路径使用双反斜杠(`\\`)或正斜杠(`/`)，避免使用单反斜杠

## 6. 总结

综合工具函数是Office Word MCP Server的推荐API，它们提供了精简而强大的接口，使文档操作更加简单和一致。特别是navigate_tools工具，作为当前推荐的定位方式，替代了传统的定位器(Locator)机制，通过设置活动文档、活动上下文和活动对象，使得其他工具可以在不需要指定定位参数的情况下进行操作。

通过掌握这些工具的使用，特别是利用navigate_tools来管理文档上下文和对象，您可以更高效地开发基于Word的应用程序和自动化工作流程。

如果您有任何问题或建议，请随时联系开发团队。