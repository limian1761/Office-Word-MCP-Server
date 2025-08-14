# Word COM接口集成指南

## 概述

本项目现已集成Word COM接口作为可选的增强功能，提供比python-docx更精确的Word文档处理能力。

## COM接口 vs python-docx对比

### COM接口优势
- **格式精确度**：100%还原Word格式，包括Word特有的格式属性
- **功能完整性**：支持Word的所有功能，包括高级格式、样式、批注等
- **统计准确性**：使用Word内置统计，数据更准确
- **兼容性**：与Microsoft Word完全兼容

### COM接口限制
- **平台限制**：仅支持Windows系统
- **软件依赖**：需要安装Microsoft Word
- **性能开销**：启动Word进程有一定开销
- **并发限制**：COM接口不适合高并发场景

### python-docx优势
- **跨平台**：支持Windows、Linux、macOS
- **轻量级**：无需安装Word，纯Python实现
- **高性能**：适合批量处理和服务器环境
- **易部署**：无外部依赖

## 使用方法

### 检查COM接口可用性

```python
from word_document_server.tools.com_document_tools import check_com_availability_tool

availability = await check_com_availability_tool()
print(f"COM接口可用: {availability['available']}")
```

### 使用COM接口工具

```python
from word_document_server.tools.com_document_tools import (
    get_all_paragraphs_com_tool,
    get_paragraphs_by_range_com_tool,
    get_paragraphs_by_page_com_tool,
    analyze_paragraph_distribution_com_tool,
    get_document_properties_com_tool
)

# 获取所有段落（COM接口）
all_paragraphs = await get_all_paragraphs_com_tool("document.docx")

# 获取指定范围段落
range_paragraphs = await get_paragraphs_by_range_com_tool("document.docx", 0, 10)

# 分页获取段落
page_paragraphs = await get_paragraphs_by_page_com_tool("document.docx", 1, 50)

# 分析段落分布
analysis = await analyze_paragraph_distribution_com_tool("document.docx")

# 获取详细文档属性
properties = await get_document_properties_com_tool("document.docx")
```

### 自动选择最佳接口

```python
async def smart_get_paragraphs(filename: str):
    """智能选择最佳接口"""
    from word_document_server.tools.com_document_tools import check_com_availability_tool
    
    # 检查COM接口可用性
    com_available = await check_com_availability_tool()
    
    if com_available['available']:
        # Windows环境，使用COM接口
        return await get_all_paragraphs_com_tool(filename)
    else:
        # 非Windows环境，使用python-docx
        from word_document_server.tools.document_tools import get_all_paragraphs_tool
        return await get_all_paragraphs_tool(filename)
```

## 新增COM接口工具

### 1. get_all_paragraphs_com_tool
获取所有段落，提供更精确的格式信息

**返回字段：**
- `text`: 段落文本
- `style`: 样式名称
- `runs`: 文本运行级别的格式信息
- `alignment`: 对齐方式
- `indent`: 缩进信息
- `spacing`: 间距信息

### 2. get_document_properties_com_tool
获取详细文档属性，包括Word内置统计

**返回字段：**
- `title`: 文档标题
- `author`: 作者
- `page_count`: 页数
- `word_count`: 词数
- `character_count`: 字符数
- `paragraph_count`: 段落数
- `table_count`: 表格数
- `section_count`: 节数

### 3. 其他COM工具
- `get_paragraphs_by_range_com_tool`: 范围获取段落
- `get_paragraphs_by_page_com_tool`: 分页获取段落
- `analyze_paragraph_distribution_com_tool`: 分布分析

## 安装和配置

### Windows系统
1. 确保已安装Microsoft Word
2. 安装pywin32：
   ```bash
   pip install pywin32
   ```

### 非Windows系统
COM接口不可用，系统会自动回退到python-docx实现

## 测试COM接口

运行测试脚本：
```bash
python test_com_tools.py
```

## 集成策略

### 推荐方案：智能选择
- Windows + Word安装：使用COM接口
- 其他情况：使用python-docx

### 性能优化
- 批量处理时使用python-docx
- 需要精确格式时使用COM接口
- 服务器环境使用python-docx

## 错误处理

### COM接口错误
```python
try:
    result = await get_all_paragraphs_com_tool("document.docx")
    if "error" in result:
        # 处理COM错误
        if "COM接口不可用" in result["error"]:
            # 回退到python-docx
            pass
        else:
            # 其他错误处理
            pass
except Exception as e:
    # 异常处理
    pass
```

## 兼容性说明

- **Windows**: 支持COM接口和python-docx
- **Linux/macOS**: 仅支持python-docx
- **无Word环境**: 仅支持python-docx

## 未来扩展

- [ ] 支持Word批注和修订
- [ ] 支持文档保护功能
- [ ] 支持邮件合并
- [ ] 支持VBA宏执行