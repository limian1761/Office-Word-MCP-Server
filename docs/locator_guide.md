# Locator 使用指南与错误处理

## 概述

Locator 是 Word Document MCP Server 中用于定位文档元素的查询语言，通过 SelectorEngine 类实现。它支持直接查找和基于锚点的相对定位，并提供多种过滤器进行精确筛选。

## 基本结构

有效的 Locator 必须包含以下字段：

```python
{
    "target": {  # 必需，指定要查找的目标元素
        "type": "paragraph",  # 必需，元素类型
        "filters": []  # 可选，过滤器数组
    },
    "anchor": {  # 可选，锚点元素
        "type": "heading",
        "identifier": {}
    },
    "relation": {  # 可选，与锚点的关系
        "type": "first_occurrence_after"
    }
}
```

## 支持的元素类型

- `paragraph`: 段落
- `table`: 表格
- `heading`: 标题段落
- `cell`: 表格单元格
- `run`: 文本运行
- `inline_shape`/`image`: 内联形状/图片

## 过滤器用法

### 文本过滤器
- `contains_text`: 筛选包含指定文本的元素（不区分大小写）
  ```python
  {"contains_text": "示例文本"}
  ```

- `text_matches_regex`: 筛选文本匹配正则表达式的元素
  ```python
  {"text_matches_regex": "^\d+\.\s+"}
  ```

### 样式过滤器
- `style`: 筛选具有特定样式的元素
  ```python
  {"style": "Heading 1"}
  ```

- `is_bold`: 筛选字体为粗体的元素
  ```python
  {"is_bold": true}
  ```

### 位置过滤器
- `index_in_parent`: 筛选在父元素中特定索引位置的元素
  ```python
  {"index_in_parent": 0}  # 第一个元素
  ```

- `row_index`: 筛选表格中特定行的单元格
  ```python
  {"row_index": 2}
  ```

- `column_index`: 筛选表格中特定列的单元格
  ```python
  {"column_index": 3}
  ```

- `table_index`: 筛选特定表格中的单元格
  ```python
  {"table_index": 0}  # 第一个表格
  ```

### 其他过滤器
- `is_list_item`: 筛选是否为列表项
  ```python
  {"is_list_item": true}
  ```

- `shape_type`: 筛选特定类型的形状
  ```python
  {"shape_type": "Picture"}  # 图片类型
  ```

## 相对定位

通过 `anchor` 和 `relation` 字段实现相对定位：

```python
{
    "anchor": {  # 锚点元素
        "type": "paragraph",
        "identifier": {"text": "引言"}
    },
    "relation": {  # 与锚点的关系
        "type": "first_occurrence_after"
    },
    "target": {  # 目标元素
        "type": "table"
    }
}
```

支持的关系类型：
- `all_occurrences_within`: 锚点范围内的所有匹配元素
- `first_occurrence_after`: 锚点之后的第一个匹配元素
- `parent_of`: 锚点的父元素
- `immediately_following`: 紧跟锚点之后的元素

## 常见错误及解决方法

### 1. LocatorSyntaxError

**错误原因**：Locator 结构不完整或无效

**常见情况**：
- 缺少 `target` 字段
- `target` 缺少 `type` 字段
- 包含 `anchor` 但缺少 `relation` 字段
- 过滤器格式不正确

**解决方法**：
确保 Locator 包含所有必需字段，并遵循正确的格式规范。

### 2. ElementNotFoundError

**错误原因**：未找到匹配的元素

**常见情况**：
- 元素类型不正确
- 过滤器条件过于严格
- 锚点元素不存在

**解决方法**：
- 验证元素类型是否支持
- 检查过滤器条件是否合理
- 确保锚点元素存在于文档中

### 3. AmbiguousLocatorError

**错误原因**：期望单个元素但找到多个匹配

**解决方法**：
添加更具体的过滤器条件以缩小搜索范围，或使用 `index_in_parent` 过滤器选择特定索引的元素。

## 最佳实践

1. 尽量使用具体的元素类型和过滤器条件，避免过于宽泛的查询
2. 对于复杂查询，先测试基础部分再逐步添加条件
3. 当使用正则表达式时，确保模式正确且考虑文档中的特殊字符
4. 对于相对定位，确保锚点唯一且易于识别
5. 处理可能的异常，特别是 `ElementNotFoundError` 和 `AmbiguousLocatorError`

## 示例

### 示例 1: 查找包含特定文本的段落

```python
{
    "target": {
        "type": "paragraph",
        "filters": [{"contains_text": "研究方法"}]
    }
}
```

### 示例 2: 查找表格中的特定单元格

```python
{
    "target": {
        "type": "cell",
        "filters": [
            {"table_index": 0},
            {"row_index": 2},
            {"column_index": 3}
        ]
    }
}
```

### 示例 3: 基于标题的相对定位

```python
{
    "anchor": {
        "type": "heading",
        "identifier": {"text": "实验结果"}
    },
    "relation": {"type": "first_occurrence_after"},
    "target": {"type": "table"}
}
```