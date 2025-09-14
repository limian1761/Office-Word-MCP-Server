# Office Word MCP Server AI客户端使用指南

本指南提供AI客户端使用Office Word MCP Server的核心操作指导，帮助高效处理Word文档。

## 注意：架构更新说明

自版本1.2.0起，系统已完成从传统定位器(Locator)机制到AppContext上下文管理的全面迁移。核心变更包括：
1. 重构了`text_operations.py`等核心操作文件
2. 完全移除了对`SelectorEngine`的依赖
3. 统一使用`AppContext`和`_get_selection_range`进行文档元素定位
4. 操作函数不再需要直接处理`locator`参数

本指南已更新以反映这些变更。

## 1. 快速入门

### 1.1 基本调用格式

所有工具调用遵循统一JSON格式：

```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "工具名称",
  "args": {
    "参数1": "值1",
    "参数2": "值2"
  }
}
```

## 2. 核心工具操作

### 2.1 文档基础操作 (document_tools)

用于文档的基本管理。

#### 2.1.1 打开文档
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "document_tools",
  "args": {
    "operation_type": "open",
    "file_path": "C:\\path\\to\\your\\document.docx"
  }
}
```

#### 2.1.2 保存文档
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "document_tools",
  "args": {
    "operation_type": "save"
  }
}
```

#### 2.1.3 另存为文档
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "document_tools",
  "args": {
    "operation_type": "save_as",
    "file_path": "C:\\path\\to\\new\\document.docx"
  }
}
```

#### 2.1.4 关闭文档
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "document_tools",
  "args": {
    "operation_type": "close"
  }
}
```

#### 2.1.5 获取文档大纲
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "document_tools",
  "args": {
    "operation_type": "get_outline"
  }
}
```

### 2.2 文本内容操作 (text_tools)

用于处理文档中的文本。

#### 2.2.1 获取文本

**获取特定段落文本：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "get_text",
    "locator": "paragraph:1"
  }
}
```

**获取包含特定文本的段落：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "get_text",
    "locator": "paragraph[contains_text=重要信息]"
  }
}
```

**获取文档特定范围文本：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "get_text",
    "locator": "document_start[range_start=100][range_end=200]"
  }
}
```

**使用JSON对象格式的locator：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "get_text",
    "locator": {
      "type": "paragraph",
      "value": "3",
      "filters": [
        {"contains_text": "示例"}
      ]
    }
  }
}
```

#### 2.2.2 插入文本

**在文档末尾插入：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "insert_text",
    "locator": "document_end",
    "text": "要插入的文本内容"
  }
}
```

**在特定段落前插入：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "insert_text",
    "locator": "paragraph:2",
    "position": "before",
    "text": "插入到第2段前面的内容"
  }
}
```

**在包含特定文本的段落内插入：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "insert_text",
    "locator": "paragraph[contains_text=关键词]",
    "text": "插入的文本"
  }
}
```

#### 2.2.3 替换文本

**替换特定段落中的全部文本：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "replace_text",
    "locator": "paragraph:5",
    "text": "新的段落内容"
  }
}
```

**替换包含特定文本的段落：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "replace_text",
    "locator": "paragraph[contains_text=旧文本]",
    "text": "新文本内容"
  }
}
```

**使用正则表达式定位并替换：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "replace_text",
    "locator": "paragraph[text_matches_regex=^日期：.*]",
    "text": "日期：2023-12-31"
  }
}
```

#### 2.2.4 获取字符计数

**获取整个文档的字符数：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "get_char_count"
  }
}
```

**获取特定段落的字符数：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "get_char_count",
    "locator": "paragraph:3"
  }
}
```

#### 2.2.5 应用文本格式

**为特定段落应用格式：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "apply_formatting",
    "locator": "paragraph:1",
    "formatting": {
      "bold": true,
      "font_size": 16,
      "alignment": "center"
    }
  }
}
```

**为包含特定文本的段落应用格式：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "apply_formatting",
    "locator": "paragraph[contains_text=标题]",
    "formatting": {
      "font_name": "微软雅黑",
      "font_color": "#0000FF",
      "bold": true
    }
  }
}
```

### 2.3 表格管理 (table_tools)

用于创建和操作文档中的表格。

#### 2.3.1 创建表格

**在文档末尾创建表格：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "table_tools",
  "args": {
    "operation_type": "create",
    "rows": 3,
    "cols": 4,
    "locator": "document_end"
  }
}
```

**在特定段落之后创建表格：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "table_tools",
  "args": {
    "operation_type": "create",
    "rows": 2,
    "cols": 3,
    "locator": "paragraph:5",
    "position": "after"
  }
}
```

#### 2.3.2 获取单元格内容

**获取特定表格的单元格内容：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "table_tools",
  "args": {
    "operation_type": "get_cell",
    "table_index": 1,
    "row": 1,
    "col": 1
  }
}
```

#### 2.3.3 设置单元格内容

**设置特定表格的单元格内容：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "table_tools",
  "args": {
    "operation_type": "set_cell",
    "table_index": 1,
    "row": 1,
    "col": 1,
    "text": "单元格内容"
  }
}
```

**设置单元格内容并应用格式：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "table_tools",
  "args": {
    "operation_type": "set_cell",
    "table_index": 1,
    "row": 2,
    "col": 3,
    "text": "格式化内容",
    "formatting": {
      "bold": true,
      "font_color": "#FF0000"
    }
  }
}
```

#### 2.3.4 获取表格信息

**获取所有表格信息：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "table_tools",
  "args": {
    "operation_type": "get_info"
  }
}
```

**获取特定表格信息：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "table_tools",
  "args": {
    "operation_type": "get_info",
    "table_index": 2
  }
}
```

#### 2.3.5 插入行

**在表格末尾插入行：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "table_tools",
  "args": {
    "operation_type": "insert_row",
    "table_index": 1
  }
}
```

**在特定位置插入多行：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "table_tools",
  "args": {
    "operation_type": "insert_row",
    "table_index": 1,
    "position": 2,
    "count": 3
  }
}
```

#### 2.3.6 插入列

**在表格末尾插入列：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "table_tools",
  "args": {
    "operation_type": "insert_column",
    "table_index": 1
  }
}
```

**在特定位置插入多列：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "table_tools",
  "args": {
    "operation_type": "insert_column",
    "table_index": 1,
    "position": 1,
    "count": 2
  }
}
```

### 2.4 图片操作 (image_tools)

用于处理文档中的图片。

#### 2.4.1 获取图片信息

**获取所有图片信息：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "image_tools",
  "args": {
    "operation_type": "get_info"
  }
}
```

#### 2.4.2 插入图片

**在文档末尾插入图片：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "image_tools",
  "args": {
    "operation_type": "insert",
    "image_path": "C:\\path\\to\\image.jpg",
    "locator": "document_end"
  }
}
```

**在特定段落之后插入图片：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "image_tools",
  "args": {
    "operation_type": "insert",
    "image_path": "C:\\path\\to\\image.jpg",
    "locator": "paragraph:3",
    "position": "after"
  }
}
```

#### 2.4.3 为图片添加说明文字

**为特定位置的图片添加说明：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "image_tools",
  "args": {
    "operation_type": "add_caption",
    "caption_text": "这是图片说明",
    "locator": "paragraph:4"
  }
}
```

**为图片添加带标签的说明：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "image_tools",
  "args": {
    "operation_type": "add_caption",
    "caption_text": "产品截图",
    "locator": "paragraph:5",
    "label": "图1"
  }
}
```

#### 2.4.4 调整图片大小

**按宽度调整图片大小：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "image_tools",
  "args": {
    "operation_type": "resize",
    "width": 300,
    "locator": "paragraph:6"
  }
}
```

**同时按宽度和高度调整图片大小：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "image_tools",
  "args": {
    "operation_type": "resize",
    "width": 400,
    "height": 300,
    "locator": "paragraph[contains_text=图片]"
  }
}
```

#### 2.4.5 设置图片颜色类型

**将图片设置为灰度：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "image_tools",
  "args": {
    "operation_type": "set_color_type",
    "color_type": "grayscale",
    "locator": "paragraph:7"
  }
}
```

### 2.5 注释管理 (comment_tools)

用于添加和管理文档注释。

#### 2.5.1 添加注释

**为特定段落添加注释：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "comment_tools",
  "args": {
    "operation_type": "add",
    "comment_text": "这是一条注释",
    "locator": "paragraph:1"
  }
}
```

**添加带作者信息的注释：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "comment_tools",
  "args": {
    "operation_type": "add",
    "comment_text": "需要进一步修改",
    "locator": "paragraph[contains_text=关键内容]",
    "author": "审核人"
  }
}
```

#### 2.5.2 删除注释

```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "comment_tools",
  "args": {
    "operation_type": "delete",
    "comment_id": 1
  }
}
```

#### 2.5.3 获取所有注释

```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "comment_tools",
  "args": {
    "operation_type": "get_all"
  }
}
```

#### 2.5.4 回复注释

```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "comment_tools",
  "args": {
    "operation_type": "reply",
    "comment_text": "已处理",
    "comment_id": 1
  }
}
```

#### 2.5.5 获取特定评论线程

```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "comment_tools",
  "args": {
    "operation_type": "get_thread",
    "comment_id": 2
  }
}
```

#### 2.5.6 删除所有注释

```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "comment_tools",
  "args": {
    "operation_type": "delete_all"
  }
}
```

#### 2.5.7 编辑注释

```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "comment_tools",
  "args": {
    "operation_type": "edit",
    "comment_text": "更新后的注释内容",
    "comment_id": 3
  }
}
```

### 2.6 视图控制 (view_control_tools)

用于控制Word文档的视图显示方式。

#### 2.6.1 切换视图

**切换到阅读视图：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "view_control_tools",
  "args": {
    "operation_type": "switch_view",
    "view_type": "read"
  }
}
```

**切换到打印预览视图：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "view_control_tools",
  "args": {
    "operation_type": "switch_view",
    "view_type": "print"
  }
}
```

#### 2.6.2 设置缩放比例

**设置固定缩放比例：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "view_control_tools",
  "args": {
    "operation_type": "set_zoom",
    "percentage": 150
  }
}
```

**设置为页面宽度：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "view_control_tools",
  "args": {
    "operation_type": "set_zoom",
    "view_type": "page_width"
  }
}
```

#### 2.6.3 显示元素

**显示段落标记：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "view_control_tools",
  "args": {
    "operation_type": "show_element",
    "element_type": "paragraph_marks"
  }
}
```

#### 2.6.4 隐藏元素

**隐藏网格线：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "view_control_tools",
  "args": {
    "operation_type": "hide_element",
    "element_type": "gridlines"
  }
}
```

#### 2.6.5 切换元素显示状态

**切换标尺显示状态：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "view_control_tools",
  "args": {
    "operation_type": "toggle_element",
    "element_type": "rulers"
  }
}
```

#### 2.6.6 获取视图信息

```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "view_control_tools",
  "args": {
    "operation_type": "get_view_info"
  }
}
```

#### 2.6.7 导航到指定位置

**导航到下一页：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "view_control_tools",
  "args": {
    "operation_type": "navigate",
    "navigation_type": "page",
    "direction": "next"
  }
}
```

**导航到指定章节：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "view_control_tools",
  "args": {
    "operation_type": "navigate",
    "navigation_type": "heading",
    "heading_level": 1
  }
}
```

### 2.7 导航与上下文管理 (navigate_tools)

用于设置活动文档、活动上下文和活动对象，使其他工具可以在不需要指定locator参数的情况下进行操作。

#### 2.7.1 设置活动上下文

**设置活动上下文为特定章节：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "navigate_tools",
  "args": {
    "operation_type": "set_active_context",
    "context_type": "section",
    "context_value": 2
  }
}
```

**设置活动上下文为特定标题：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "navigate_tools",
  "args": {
    "operation_type": "set_active_context",
    "context_type": "heading",
    "context_value": "第一章 概述"
  }
}
```

**设置活动上下文为特定书签：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "navigate_tools",
  "args": {
    "operation_type": "set_active_context",
    "context_type": "bookmark",
    "context_value": "section_summary"
  }
}
```

#### 2.7.2 设置活动对象

**设置活动对象为特定段落：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "navigate_tools",
  "args": {
    "operation_type": "set_active_object",
    "object_type": "paragraph",
    "object_value": 5
  }
}
```

**设置活动对象为特定表格：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "navigate_tools",
  "args": {
    "operation_type": "set_active_object",
    "object_type": "table",
    "object_value": 2
  }
}
```

**设置活动对象为特定图片：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "navigate_tools",
  "args": {
    "operation_type": "set_active_object",
    "object_type": "image",
    "object_value": 3
  }
}
```

#### 2.7.3 获取当前上下文信息

**获取当前活动上下文：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "navigate_tools",
  "args": {
    "operation_type": "get_active_context"
  }
}
```

**获取当前活动对象：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "navigate_tools",
  "args": {
    "operation_type": "get_active_object"
  }
}
```

#### 2.7.4 上下文感知的工具调用示例

设置好活动上下文和活动对象后，其他工具可以在不指定locator参数的情况下工作：

**传统方式（需要指定locator）：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "insert_text",
    "locator": {
      "type": "paragraph",
      "value": 5
    },
    "text": "新的段落内容"
  }
}
```

**上下文感知方式（无需指定locator）：**
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "navigate_tools",
  "args": {
    "operation_type": "set_active_object",
    "object_type": "paragraph",
    "object_value": 5
  }
}

{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "insert_text",
    "text": "新的段落内容"
  }
}

## 3. Locator精确定位机制

Locator用于精确指定文档中需要操作的元素，是工具调用的核心参数。

### 3.1 基本语法

Locator支持两种格式：
1. **字符串格式**：`type:value[filter1][filter2]...`
2. **JSON对象格式**：
```json
{
  "type": "paragraph",
  "value": "1",
  "filters": [
    {"contains_text": "example"}
  ]
}
```

### 3.2 常用元素类型

- **paragraph**: 段落元素
  - 示例: `paragraph:3` (第3个段落) 或 `paragraph:"标题文本"` (含特定文本的段落)
- **table**: 表格元素
- **text**: 文本元素
- **image**: 图片元素
- **cell**: 表格单元格
- **document_start**: 文档开始位置
- **document_end**: 文档结束位置

### 3.3 实用过滤条件

#### 3.3.1 文本过滤
- **contains_text**: 元素包含指定文本
  - 示例: `[contains_text=重要信息]`
- **text_matches_regex**: 文本匹配正则表达式
  - 示例: `[text_matches_regex=^第[0-9]+章]`

#### 3.3.2 位置过滤
- **index_in_parent**: 元素在父元素中的索引
  - 示例: `[index_in_parent=0]` (第一个元素)
- **range_start/range_end**: 指定文本范围
  - 示例: `[range_start=100][range_end=200]`

## 4. 常见操作示例

### 4.1 读取文档全部内容
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "get_text"
  }
}
```

### 4.2 替换特定文本
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "replace_text",
    "locator": "paragraph[contains_text=公司名称]",
    "text": "新公司名称"
  }
}
```

### 4.3 提取表格数据
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "table_tools",
  "args": {
    "operation_type": "get_info"
  }
}
```

### 4.4 添加文档总结
```json
{
  "server_name": "mcp.config.usrlocalmcp.word-docx-tools",
  "tool_name": "text_tools",
  "args": {
    "operation_type": "insert_text",
    "locator": "document_end",
    "text": "\n\n### 文档总结\n这是根据文档内容生成的总结。"
  }
}
```

## 5. 使用技巧

### 5.1 元素精确定位
- 优先使用内容特征定位，比位置索引更稳定
- 组合多种过滤条件提高定位精度
- 处理长文档时，使用range_start和range_end分块读取

### 5.2 高效文档操作
- 在一个会话中完成相关操作，减少交互次数
- 操作完成后及时保存并关闭文档，释放资源
- 批量处理多个小操作以提高效率

### 5.3 常见问题处理
- **元素定位失败**：检查定位器条件，尝试使用更宽泛的过滤条件
- **语法错误**：确保遵循 `type:value[filter1][filter2]...` 标准格式
- **大文档性能问题**：使用更具体的定位器，减少每次操作范围