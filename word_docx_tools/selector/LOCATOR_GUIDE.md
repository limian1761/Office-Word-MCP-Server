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
- [AI生成定位器参数的约束规范](#ai生成定位器参数的约束规范)
  - [约束目标](#约束目标)
  - [参数类型约束](#参数类型约束)
  - [必填字段约束](#必填字段约束)
  - [过滤器参数规范](#过滤器参数规范)
  - [锚点关系约束](#锚点关系约束)
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

## 元素类型

### 基本元素类型

1. **paragraph** - 段落元素
   根据段落位置或内容定位。
   - 示例: `paragraph:3` (第3个段落)
   - 示例: `paragraph:"标题文本"` (包含"标题文本"的段落)
   
   **重要说明**：当段落定位器包含索引参数时（如`paragraph:5`），系统会强制返回单个对象，确保操作只针对特定段落执行，避免误操作影响整个文档。

2. **table** - 表格元素
3. **text** - 文本元素（会被转换为段落搜索）
4. **inline_shape** 或 **image** - 图片元素
5. **cell** - 表格单元格
6. **run** - 文本运行（单词）
7. **document_start** - 文档开始位置
8. **document_end** - 文档结束位置
9. **range** - 范围元素
10. **bookmark** - 书签
11. **comment** - 注释

## 过滤器

过滤器用于进一步缩小定位范围，常用的过滤器包括：

### 文本相关过滤器
- **contains_text**: 元素包含指定的文本
  - 示例: `[contains_text=重要信息]`
- **text_matches_regex**: 元素文本与指定的正则表达式匹配
  - 示例: `[text_matches_regex=^第[0-9]+章]`
- **is_list_item**: 元素是列表项
  - 示例: `[is_list_item=true]`

### 位置相关过滤器
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

### 样式相关过滤器
- **style**: 元素具有指定的样式
  - 示例: `[style=标题1]`
- **is_bold**: 元素字体为粗体
  - 示例: `[is_bold=true]`

### 形状相关过滤器
- **shape_type**: 形状类型
  - 示例: `[shape_type=Picture]`

## 关系类型

当使用带锚点的定位器时，可以指定以下关系类型：

1. **all_occurrences_within** - 查找锚点元素范围内的所有目标元素
2. **first_occurrence_after** - 查找锚点元素之后的第一个目标元素
3. **parent_of** - 查找锚点元素的父元素
4. **immediately_following** - 查找锚点元素之后紧接着的目标元素
5. **immediately_preceding** - 查找锚点元素之前紧接着的目标元素

## 定位器稳定性策略

为确保定位器在文档修改后仍然有效，建议采用以下策略：

1. **使用相对位置**：优先使用相对位置而非绝对索引或位置
2. **使用内容特征**：基于元素的文本内容、样式等不易受修改影响的特征进行定位
3. **使用书签或自定义标记**：在关键点位置插入书签或自定义标记作为定位锚点
4. **先查找后操作模式**：在执行修改操作前，先使用定位器查找元素并获取其实时位置信息
5. **批量执行策略**：将相关操作分组执行，以减少定位器失效的影响

## 使用示例

以下是一些定位器的使用示例，展示如何在实际应用中精确定位文档元素。

### 定位特定段落

```json
{
  "type": "paragraph",
  "value": "3"
}
```

这将定位文档中的第3个段落。

### 定位包含特定文本的段落

```json
{
  "type": "paragraph",
  "filters": [
    {"contains_text": "示例"}
  ]
}
```

这将定位包含"示例"文本的所有段落。

### 定位文档中的第一个表格

```python
# 定位器字符串形式
locator_str = "table[index_in_parent=0]"
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

### 相对定位示例

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

## AI生成定位器参数的约束规范

为提高定位器的准确性和稳定性，系统对AI生成的定位器参数实施了严格的约束机制。以下是约束规范的详细说明，完全基于代码实现。

### 约束目标

1. **提高定位准确性**：确保AI生成的定位器参数能够精确识别目标元素
2. **增强稳定性**：减少因参数格式错误或逻辑矛盾导致的定位失败
3. **防止歧义**：避免定位器匹配多个不相关元素的情况
4. **优化性能**：通过严格约束减少不必要的搜索和验证开销

### 参数类型约束

定位器参数必须满足以下类型约束：

1. **基础类型验证**：
   - 定位器必须是字符串或JSON对象格式
   - 对象类型字段必须是预定义的有效值
   - 数值类型参数必须是有效的数字格式

2. **支持的对象类型**：
   - `paragraph` - 段落元素
   - `table` - 表格元素
   - `cell` - 表格单元格
   - `inline_shape` / `image` - 图片元素
   - `comment` - 注释
   - `range` - 范围元素
   - `selection` - 选择区域
   - `document` - 整个文档
   - `document_start` - 文档开始位置
   - `document_end` - 文档结束位置

3. **特定元素类型约束**：
   - 段落（paragraph）：当treat_as_index为True时，索引值必须是正整数
   - 表格（table）：索引值必须是正整数
   - 文档开始/结束（document_start/document_end）：不能有value或filters参数
   - 段落和表格类型：必须指定value或filters确保确定性选择

### 必填字段约束

为确保定位器的有效性，以下字段为必填项：

1. **对象类型（type）**：必须指定明确的元素类型
2. **JSON对象格式**：必须包含type字段，且type必须是非空字符串
3. **带锚点的定位器**：必须同时指定anchor和relation字段
4. **过滤器**：每个过滤器必须是单键值对字典

### 过滤器参数规范

系统支持以下过滤器类型，每个过滤器都有严格的参数类型约束：

1. **支持的过滤器类型**：
   - `index` - 元素索引位置
   - `contains_text` - 元素包含指定文本
   - `text_matches_regex` - 元素文本匹配正则表达式
   - `shape_type` - 形状类型
   - `style` - 元素样式
   - `is_bold` - 字体是否加粗
   - `row_index` - 单元格行索引
   - `column_index` - 单元格列索引
   - `table_index` - 表格索引
   - `is_list_item` - 是否为列表项
   - `range_start` - 范围起始位置
   - `range_end` - 范围结束位置
   - `has_style` - 具有指定样式

2. **参数类型验证**：
   - `index`、`row_index`、`column_index`、`table_index`、`range_start`、`range_end`：必须是整数
   - `contains_text`、`text_matches_regex`、`shape_type`、`style`：必须是字符串
   - `is_bold`、`is_list_item`：必须是布尔值

3. **参数范围约束**：
   - 索引值必须在合理范围内（不能超出当前列表大小）
   - 索引值在定位器中是从0开始计数

### 锚点关系约束

使用带锚点的定位器时，必须满足以下约束：

1. **支持的关系类型**：
   - `all_occurrences_within` - 查找锚点元素范围内的所有目标元素
   - `first_occurrence_after` - 查找锚点元素之后的第一个目标元素
   - `parent_of` - 查找锚点元素的父元素
   - `immediately_following` - 查找锚点元素之后紧接着的目标元素

2. **锚点-关系对验证**：
   - 必须同时指定anchor和relation
   - anchor必须是字典类型
   - relation必须是支持的关系类型之一

### 约束实现细节

所有约束通过以下方法在代码中实现：

1. `_validate_locator()` - 验证定位器的基本结构、必填字段、对象类型和过滤器
2. `enhance_parse_locator()` - 对AI生成的定位器参数进行增强验证，包括类型检查、必填字段验证、锚点关系验证和确定性检查
3. `apply_filters()` - 对过滤器参数进行严格验证，确保参数类型和范围符合要求

当定位器参数违反这些约束时，系统将抛出`LocatorSyntaxError`异常，并提供详细的错误信息。

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

### 7. 参数验证失败

**问题**：使用AI生成的定位器时遇到 `LocatorValidationError` 异常。
**解决方案**：
- 检查定位器参数是否符合类型约束
- 确保所有必填字段都已提供
- 验证过滤器参数是否在有效范围内
- 确认锚点和关系类型的组合是否有效

通过正确理解和使用定位器，您可以在 Word 文档中精确定位和操作各种元素，实现复杂的文档处理任务。遵循本指南中的约束规范，可以显著提高AI生成定位器的准确性和稳定性。