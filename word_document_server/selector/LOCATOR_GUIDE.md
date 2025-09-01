# Word Document MCP Server 定位器指南

## 文档索引

- [概述](#概述)
- [定位器语法](#定位器语法)
  - [基本格式](#基本格式)
  - [带锚点的格式](#带锚点的格式)
- [元素类型](#元素类型)
  - [基本元素类型](#基本元素类型)
- [过滤器](#过滤器)
  - [文本相关过滤器](#文本相关过滤器)
  - [位置相关过滤器](#位置相关过滤器)
  - [样式相关过滤器](#样式相关过滤器)
  - [形状相关过滤器](#形状相关过滤器)
- [关系类型](#关系类型)
- [定位器在文档修改操作后的稳定性](#定位器在文档修改操作后的稳定性)
  - [问题说明](#问题说明)
  - [解决策略](#解决策略)
  - [最佳实践](#最佳实践)
- [使用示例](#使用示例)
  - [基本选择](#基本选择)
  - [带过滤器的选择](#带过滤器的选择)
  - [带锚点的选择](#带锚点的选择)
  - [高级用法](#高级用法)
- [常见问题与解决方案](#常见问题与解决方案)

## 概述

定位器（Locator）是 Word Document MCP Server 中用于精确定位文档元素的查询语言。用户可以通过指定元素类型、值和过滤器来查找文档中的特定元素（如段落、表格、图片等）。定位器提供了一种灵活且强大的方式来识别和操作 Word 文档中的各种元素。

## 定位器语法

### 基本格式

基本的定位器格式为：`type:value[filter1][filter2]...`

- `type`: 元素类型（如 paragraph、table、text 等）
- `value`: 可选的元素值（如文本内容、标识符等）
- `filter`: 可选的过滤器，用于进一步缩小选择范围

### 带锚点的格式

带锚点的定位器格式为：`type:value@anchor_id[relation]`

- `type:value`: 目标元素的类型和值
- `anchor_id`: 锚点元素的标识符
- `relation`: 目标元素与锚点元素之间的关系类型

## 元素类型

### 基本元素类型

1. **paragraph** - 段落元素
2. **table** - 表格元素
3. **text** - 文本元素（会被转换为段落搜索）
4. **inline_shape** 或 **image** - 图片元素
5. **cell** - 表格单元格
6. **run** - 文本运行（单词）
7. **document_start** - 文档开始位置
8. **document_end** - 文档结束位置
9. **range** - 范围元素

## 过滤器

### 文本相关过滤器

- `contains_text`: 元素包含指定的文本（不区分大小写）
  - 示例: `[contains_text=重要信息]`
- `text_matches_regex`: 元素文本与指定的正则表达式匹配
  - 示例: `[text_matches_regex=^第[0-9]+章]`
- `is_list_item`: 元素是列表项
  - 示例: `[is_list_item=true]` 或 `[is_list_item=false]`

### 位置相关过滤器

- `index_in_parent`: 元素在父元素中的索引位置（0 开始）
  - 示例: `[index_in_parent=0]`（第一个元素）
- `row_index`: 单元格在表格中的行索引
  - 示例: `[row_index=2]`
- `column_index`: 单元格在表格中的列索引
  - 示例: `[column_index=3]`
- `table_index`: 元素所属表格的索引
  - 示例: `[table_index=0]`（第一个表格）
- `range_start`: 范围元素的起始位置
  - 示例: `[range_start=100]`
- `range_end`: 范围元素的结束位置
  - 示例: `[range_end=200]`

### 样式相关过滤器

- `style`: 元素具有指定的样式
  - 示例: `[style=标题1]` 或 `[style=Heading 1]`
- `is_bold`: 元素字体为粗体
  - 示例: `[is_bold=true]` 或 `[is_bold=false]`

### 形状相关过滤器

- `shape_type`: 形状类型（例如 "Picture", "Chart" 等）
  - 示例: `[shape_type=Picture]`

## 关系类型

当使用带锚点的定位器时，可以指定以下关系类型：

1. **all_occurrences_within** - 查找锚点元素范围内的所有目标元素
2. **first_occurrence_after** - 查找锚点元素之后的第一个目标元素
3. **parent_of** - 查找锚点元素的父元素
4. **immediately_following** - 查找锚点元素之后紧接着的目标元素

## 定位器在文档修改操作后的稳定性

### 问题说明

在使用 Word Document MCP Server 进行文档编辑时，定位器可能会因为文档结构的变化而失效。以下是一些常见的影响：

1. **段落编号变更**：插入或删除段落会改变后续段落的索引位置
2. **起始位置变更**：在文档前部插入内容会使后面所有内容的字符位置偏移
3. **行列变更**：表格操作可能会改变行/列索引
4. **样式变更**：格式化操作可能会影响基于样式的定位

### 解决策略

为确保定位器在文档修改后仍然有效，建议采用以下策略：

#### 1. 使用相对位置

优先使用相对位置而非绝对索引或位置：

```
{
  "anchor": {
    "type": "paragraph",
    "identifier": {
      "text": "锚点段落文本"
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

避免使用绝对索引：

```
{
  "target": {
    "type": "paragraph",
    "filters": [
      {"index_in_parent": 5}
    ]
  }
}
```

#### 2. 使用内容特征进行定位

基于元素的文本内容、样式等不易受修改影响的特征进行定位：

```
{
  "target": {
    "type": "paragraph",
    "filters": [
      {"contains_text": "特定标识文本"},
      {"style": "标题1"}
    ]
  }
}
```

#### 3. 使用书签或自定义标记

在关键点位置插入书签或自定义标记作为定位锚点：

```
{
  "anchor": {
    "type": "bookmark",
    "identifier": {
      "name": "关键点书签"
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

#### 4. 先查找后操作模式

在执行修改操作前，先使用定位器查找元素并获取其实时位置信息，然后基于该信息执行操作：

```python
# 1. 先查找元素
selected_objects = selector_engine.select(document, locator)

# 2. 基于查找的元素执行操作
for object in selected_objects:
    # 执行操作
    pass

# 3. 如果需要继续操作，更新定位器
```

#### 5. 批量执行策略

将相关操作分组执行，以减少定位器失效的影响：

```python
# 一次性获取所有需要操作的元素
objects_to_modify = selector_engine.select(document, locator)

# 对所有元素执行操作
for object in objects_to_modify:
    # 执行操作
    pass
```

### 最佳实践

1. **避免长时间引用定位器**：定位器应在使用后立即执行操作，避免在多个操作步骤中使用相同的定位器

2. **使用稳定的标识文本**：在文档中插入用于定位的标识文本，但不显示在实际内容中

3. **更新定位器**：在可能导致定位器失效的操作后，重新生成或更新定位器

4. **错误处理**：在使用定位器时添加适当的错误处理，以应对定位器失效的情况

5. **测试验证**：在复杂的操作序列中验证定位器的有效性

## 使用示例

### 基本选择

选择包含指定文本的所有段落：

```python
# 定位器字符串形式
locator_str = "paragraph[contains_text=重要信息]"

# 解析定位器
parsed_locator = selector_engine.parse_locator(locator_str)

# 选择元素
selection = selector_engine.select(document, parsed_locator)
```

### 带过滤器的选择

选择文档中的第一个表格：

```python
locator_str = "table[index_in_parent=0]"
parsed_locator = selector_engine.parse_locator(locator_str)
selection = selector_engine.select(document, parsed_locator)
```

选择特定样式的所有段落：

```python
locator_str = "paragraph[style=标题1]"
parsed_locator = selector_engine.parse_locator(locator_str)
selection = selector_engine.select(document, parsed_locator)
```

### 带锚点的选择

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

selection = selector_engine.select(document, locator)
```

### 高级用法

#### 过滤器组合

组合多个过滤器精确定位元素：

```python
locator_str = "paragraph[contains_text=报告][style=正文][is_bold=true]"
parsed_locator = selector_engine.parse_locator(locator_str)
selection = selector_engine.select(document, parsed_locator)
```

#### 范围操作

选择并操作特定范围的文本：

```python
locator_str = "range[range_start=100][range_end=200]"
parsed_locator = selector_engine.parse_locator(locator_str)
selection = selector_engine.select(document, parsed_locator)
```

#### 正则表达式过滤

使用正则表达式选择匹配特定模式的段落：

```python
locator_str = "paragraph[text_matches_regex=^第[0-9]+章]"
parsed_locator = selector_engine.parse_locator(locator_str)
selection = selector_engine.select(document, parsed_locator)
```

## 常见问题与解决方案

### 1. 定位器语法错误

**问题**：使用定位器时遇到 `LocatorSyntaxError` 异常。
**解决方案**：
- 检查定位器格式是否正确，确保遵循 `type:value[filter1][filter2]...` 格式
- 验证元素类型是否受支持
- 检查过滤器语法是否正确，特别是正则表达式

### 2. 找不到匹配的元素

**问题**：使用定位器时遇到 `ObjectNotFoundError` 异常。
**解决方案**：
- 检查文档中是否存在符合定位器条件的元素
- 确认过滤器值是否正确（注意大小写）
- 考虑使用更宽泛的过滤条件

### 3. 定位器匹配多个元素但期望单个元素

**问题**：使用定位器时遇到 `AmbiguousLocatorError` 异常。
**解决方案**：
- 增加更多过滤器以缩小选择范围
- 使用 `index_in_parent` 过滤器选择特定索引的元素
- 在调用 `select` 方法时设置 `expect_single=False`

### 4. 文档修改后定位器失效

**问题**：在编辑文档后，之前有效的定位器不再能找到正确的元素。
**解决方案**：
- 采用前面提到的稳定性策略
- 在文档修改后重新生成定位器
- 使用基于内容的定位而非基于位置的定位

### 5. 索引从0开始

**提示**：所有索引值在定位器中都是从0开始的，而不是从1开始。

### 6. 大小写敏感性

**提示**：大多数过滤器值是区分大小写的，特别是样式名称。

通过正确理解和使用定位器，您可以在 Word 文档中精确定位和操作各种元素，实现复杂的文档处理任务。