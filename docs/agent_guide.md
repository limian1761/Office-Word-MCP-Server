## Office-Word-MCP-Server 使用指南

如何使用Office-Word-MCP-Server进行Word文档操作，包括工具调用方法、参数说明和使用示例。

### 使用流程概述

1. **打开文档**：使用`open_document`工具打开一个Word文档
2. **操作文档**：使用各种功能工具对文档进行编辑、格式化等操作
3. **保存更改**：工具会自动保存文档更改
4. **关闭文档**：使用`shutdown_word`工具关闭Word应用程序

### Locator 使用指南

Locator是Word Document MCP Server中用于精确定位文档元素的查询语言，通过SelectorEngine类实现。它支持直接查找和基于锚点的相对定位，并提供多种过滤器进行精确筛选。

#### 基本结构
有效的Locator必须包含以下字段：
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

#### 支持的元素类型
- `paragraph`: 段落
- `table`: 表格
- `heading`: 标题段落
- `cell`: 表格单元格
- `run`: 文本运行
- `inline_shape`/`image`: 内联形状/图片

#### 常用过滤器
- **文本过滤器**: `contains_text`（包含指定文本）、`text_matches_regex`（匹配正则表达式）
- **样式过滤器**: `style`（特定样式）、`is_bold`（粗体）
- **位置过滤器**: `index_in_parent`（父元素中索引位置）、`row_index`/`column_index`（表格行列）、`table_index`（表格索引）
- **其他过滤器**: `is_list_item`（列表项）、`shape_type`（形状类型）

#### 相对定位
通过`anchor`和`relation`字段实现相对定位，支持关系类型如：
- `all_occurrences_within`: 锚点范围内的所有匹配元素
- `first_occurrence_after`: 锚点之后的第一个匹配元素
- `parent_of`: 锚点的父元素
- `immediately_following`: 紧跟锚点之后的元素

#### 示例
查找包含特定文本的段落：
```python
{
    "target": {
        "type": "paragraph",
        "filters": [{"contains_text": "研究方法"}]
    }
}
```

查找表格中的特定单元格：
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

### 最佳实践

1. **会话管理**：在一个会话中完成一组相关操作，但不要自动调用`shutdown_word`工具，应由人类用户自己接受修订和关闭文档
2. **修订模式**：文档编辑前打开修订模式，以便追踪和管理所有修改
3. **错误处理**：检查工具返回值是否包含错误信息，根据错误类型采取相应措施。特别注意处理COM组件错误，可能需要重启Word应用
4. **Locator 最佳实践**：
   - 使用精确的Locator避免误操作，特别是在批量操作时
   - 尽量使用具体的元素类型和过滤器条件，避免过于宽泛的查询
   - 对于复杂查询，先测试基础部分再逐步添加条件
   - 当使用正则表达式时，确保模式正确且考虑文档中的特殊字符
   - 对于相对定位，确保锚点唯一且易于识别
   - 处理可能的异常，特别是`ElementNotFoundError`和`AmbiguousLocatorError`
5. **路径处理**：确保提供的文件路径是绝对路径
6. **文档检查**：使用`open_document`前确认文件存在且格式正确
7. **样式应用**：对于段落样式应用，优先使用`apply_format`工具，提供更稳定的样式应用体验
8. **批量操作**：当需要对多个元素应用不同格式时，使用`batch_apply_format`工具提高效率
9. **样式验证**：在应用样式前，可以使用`get_document_styles`获取文档中可用的样式列表，避免使用不存在的样式
10. **错误重试**：对于临时性错误，实现适当的重试机制
11. **组件访问**：当访问特定Word组件（如Paragraphs、Comments、InlineShapes、Tables）时，确保文档包含相应组件并使用适当的访问方式
12. **表格操作**：
    - 在操作表格前，确保表格存在且结构完整
    - 使用精确的表格索引、行索引和列索引定位单元格
    - 处理表格操作时，注意文档保护状态和表格锁定状态
    - 对于表格错误，建议关闭并重新打开文档后重试
13. **参数验证**：
    - 在调用工具前验证参数的有效性（如行数、列数必须为正整数）
    - 使用适当的错误消息帮助用户快速定位问题
14. **异常处理**：
    - 实现嵌套异常处理结构，区分不同类型的错误
    - 对COM相关错误提供具体的解决建议
    - 保持错误消息的一致性和可操作性