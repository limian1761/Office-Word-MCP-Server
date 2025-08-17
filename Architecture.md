
### **架构设计文档：下一代Word MCP服务 (COM接口版)**

**版本:** 2.1
**代号:** "Live Control" (实时控制)
**日期:** 2025年8月15日

---

### **1. 概述**

#### **1.1. 愿景与目标**
本文档旨在设计一个先进的、专为AI调用的Word文档自动化服务。此服务将严格遵守 [modelcontextprotocol.io](https://modelcontextprotocol.io/) 定义的**模型上下文协议 (Model Context Protocol, MCP)** 规范，作为一个标准的MCP服务器运行。

服务的目标是：通过实现一系列由 `mcp-config.json` 文件定义的**工具 (Tools)**，为大语言模型 (LLM) 提供对本地Microsoft Word文档进行复杂、精确、上下文感知操作的能力。

通过采用Microsoft Office的COM接口作为底层技术，本服务将具备对运行中的Word应用程序进行实时、精细化控制的能力，实现真正的“所见即所得”自动化。

#### **1.2. 核心理念**
架构的核心是**“关注点分离 (Separation of Concerns)”**，将复杂的自动化任务分解为三个独立且清晰的层次：

*   **“做什么 (What)”**: 由顶层的**MCP服务层**定义。此层接口语义清晰，直接反映AI的业务意图（如“插入段落”、“应用格式”）。
*   **“在哪里做 (Where)”**”: 由一个独立的、强大的**选择器引擎**负责。它通过一种声明式的查询语言（`Locator`对象）来精确定位文档中的任何元素。
*   **“怎么做 (How)”**”: 由**选择集抽象层**和**COM后端适配层**负责。它们封装了所有与底层COM接口交互的复杂细节，将具体操作转化为对Word应用的指令。

#### **1.3. 架构指导原则**
1.  **MCP合规性:** 服务器的外部接口、工具定义和上下文报告必须严格遵守MCP规范。
2.  **解耦与模块化:** 各组件职责单一、高度独立，便于独立开发、测试和维护。
3.  **声明式定位:** AI调用者只需通过`Locator`对象**声明**“要找什么”，而无需关心**如何找到**它的具体算法。
4.  **抽象化后端:** 必须将所有与`pywin32`相关的COM对象操作严格封装在后端适配层，防止COM的复杂性泄漏到上层业务逻辑中。
5.  **健壮的错误处理与资源管理:** COM对象需要显式管理其生命周期。必须确保在操作完成后，所有COM对象都被正确释放，以防止Word进程挂起或内存泄漏。

---

### **2. MCP服务器实现 (基于官方Python SDK)**

本服务将使用官方的 `python-sdk` (`mcp`库) 来构建，以确保协议的正确实现和简化开发。

*   **MCP Server**: 由 `word_document_server/app.py` 实现。我们将实例化 `mcp.FastMCP` 类来创建服务器应用，而不是手动搭建Flask/FastAPI。
*   **MCP Tools**: 我们将在 `word_document_server/tools/content_tools.py` (以及其他工具文件中) 定义具体的工具函数。这些函数将使用 `@mcp.tool()` 装饰器进行注册。这样做比在 `McpService` 中定义方法更符合SDK的惯例。
*   **MCP Context Provider**: SDK会自动处理上下文报告。我们需要定义一个 `@mcp.resource` 来报告当前活动文档的状态。
*   **`mcp-config.json`**: 此文件依然是服务的清单文件。SDK可以自动生成此文件，或者我们可以手动维护它，以确保其内容与我们用 `@mcp.tool()` 装饰器定义的工具严格同步。

---

### **3. 内部体系架构**

#### **3.1. 高层架构图**

```
+-------------------------------------------------------------------------+
|                      AI / MCP Client (e.g., LLM Agent)                  |
+---------------------------------+---------------------------------------+
                                  |
                                  | (1. MCP Request via python-sdk transport)
                                  v
+-------------------------------------------------------------------------+
| MCP 服务层 (`app.py` using `mcp.FastMCP`)                               |
|                                                                         |
| @mcp.tool()                                                             |
| def insert_paragraph(ctx: mcp.Context, locator: dict, ...):             |
|   # 1. Get active document path from context                            |
|   # 2. Use WordBackend context manager                                  |
|   with WordBackend(file_path) as backend:                               |
|     # 3. Call Selector Engine                                           |
|     selection = selector.select(backend, locator)                       |
|     # 4. Perform action on selection                                    |
|     selection.insert_text(...)                                          |
|                                                                         |
| (由`python-sdk`驱动, 工具函数直接调用内部核心逻辑)                      |
+---------------------------------+---------------------------------------+
                                  |
                                  | (Internal Calls)
                                  v
+-------------------------------------------------------------------------+
| 选择器引擎 (selector.py)                                                |
| (定位逻辑, 负责接收Locator, 解析并执行查询, 返回一个或多个Selection对象)  |
+---------------------------------+---------------------------------------+
                                  |
                                  v
+-------------------------------------------------------------------------+
| 选择集抽象层 (selection.py)                                             |
| (原子操作抽象, 封装了对底层元素的原子操作)                              |
+---------------------------------+---------------------------------------+
                                  |
                                  v
+-------------------------------------------------------------------------+
| COM后端适配层 (com_backend.py)                                          |
| (封装所有pywin32调用, 管理Word应用生命周期)                             |
+---------------------------------+---------------------------------------+
                                  |
                                  v
+-------------------------------------------------------------------------+
|                      Microsoft Word Application (COM Server)            |
+-------------------------------------------------------------------------+
```

### **2. 体系架构**

#### **2.1. 高层架构图**

```
+-------------------------------------------------------------------------+
|                      AI / 外部调用者 (User of the Service)              |
+---------------------------------+---------------------------------------+
                                  |
                                  | (1. MCP函数调用, 携带Locator)
                                  v
+-------------------------------------------------------------------------+
| MCP 服务层 (mcp_service.py)                                             |
| - insert_paragraph(locator, ...)                                        |
| - apply_format(locator, ...)                                            |
| - get_text(locator)                                                     |
| (高层业务逻辑, 负责解析意图，调用选择器引擎)                            |
+---------------------------------+---------------------------------------+
                                  |
                                  | (2. 调用select(locator))
                                  v
+-------------------------------------------------------------------------+
| 选择器引擎 (selector.py)                                                |
|                                                                         |
|  +----------------+      +------------------+      +------------------+ |
|  | Locator Parser |----->|  Document Walker |----->|  Filter Engine   | |
|  +----------------+      +------------------+      +------------------+ |
| (定位逻辑, 负责接收Locator, 解析并执行查询, 返回一个或多个Selection对象)  |
+---------------------------------+---------------------------------------+
                                  |
                                  | (3. 返回Selection对象)
                                  v
+-------------------------------------------------------------------------+
| 选择集抽象层 (selection.py)                                             |
| - class Selection:                                                      |
|   - .apply_format(...)                                                  |
|   - .insert_after(...)                                                  |
|   - .delete()                                                           |
| (原子操作抽象, 封装了对底层元素的原子操作)                              |
+---------------------------------+---------------------------------------+
                                  |
                                  | (4. 调用COM接口)
                                  v
+-------------------------------------------------------------------------+
| COM后端适配层 (com_backend.py)                                          |
| (封装所有pywin32调用, 管理Word应用生命周期)                             |
+---------------------------------+---------------------------------------+
                                  |
                                  | (5. 与Word应用实时交互)
                                  v
+-------------------------------------------------------------------------+
|                      Microsoft Word Application (COM Server)            |
+-------------------------------------------------------------------------+
```

---

### **3. 核心组件详细设计**

#### **3.1. COM后端适配层 (`com_backend.py`)**

这是整个架构的基石，负责所有与Word COM接口的直接交互。

*   **职责:**
    1.  **生命周期管理:** 提供启动、附加、关闭Word应用及打开/关闭文档的方法。
    2.  **API封装:** 将庞大复杂的Word COM API封装成一组简洁、Pythonic的函数（如`set_bold_for_range`, `get_all_tables`）。
    3.  **资源管理:** 强制使用`try...finally`结构，确保COM对象在使用后被妥善释放，防止资源泄漏。
    4.  **对象获取:** 提供获取文档顶层集合（如`document.Paragraphs`, `document.Tables`）的方法。

*   **核心类:** `WordBackend`
    ```python
    # com_backend.py
    import win32com.client

    class WordBackend:
        def __init__(self, visible=True):
            try:
                self.word_app = win32com.client.Dispatch("Word.Application")
                self.word_app.Visible = visible
                self.document = None
            except Exception as e:
                raise RuntimeError(f"Failed to start Word Application: {e}")

        def open_document(self, file_path):
            # ... 打开文档实现 ...
            pass

        # --- 示例API封装 ---
        def get_all_paragraphs(self):
            return list(self.document.Paragraphs)
            
        def set_bold_for_range(self, com_range_obj, is_bold: bool):
            com_range_obj.Font.Bold = is_bold

        def insert_paragraph_after(self, com_range_obj, text):
            # ... COM插入逻辑 ...
            pass

        def cleanup(self):
            # 确保资源被释放
            if self.document:
                self.document.Close(SaveChanges=False)
            if self.word_app:
                self.word_app.Quit()
            self.document = None
            self.word_app = None
    ```

#### **3.2. 选择集抽象层 (`selection.py`)**

该层提供一个统一的接口来操作被“选择”的文档元素，无论它是一个段落、一个表格还是一个单元格。

*   **职责:**
    1.  持有由选择器引擎找到的一组底层COM对象。
    2.  持有一个`WordBackend`实例的引用，以便执行操作。
    3.  定义一组标准化的原子操作方法（`apply_format`, `delete`, `get_text`等），并将这些操作委托给后端执行。

*   **核心类:** `Selection`
    ```python
    # selection.py
    from com_backend import WordBackend

    class Selection:
        def __init__(self, raw_com_elements: list, backend: WordBackend):
            if not raw_com_elements:
                raise ValueError("Selection cannot be empty.")
            self._elements = raw_com_elements
            self._backend = backend

        def apply_format(self, options: dict):
            for element in self._elements:
                # 假设element是一个Range兼容的对象
                if options.get("bold"):
                    self._backend.set_bold_for_range(element.Range, True)
                # ... 其他格式化选项
    ```

#### **3.3. 选择器引擎 (`selector.py`)**

系统的“大脑”，负责解析`Locator`查询，并在文档中找到匹配的元素。

*   **职责:**
    1.  **解析`Locator`:** 验证`Locator`查询语言的语法，并将其分解为查询计划。
    2.  **遍历与筛选:** 调用`WordBackend`获取候选元素集合，并根据`Locator`中的`filters`进行精确筛选。
    3.  **关系处理:** 处理`anchor`和`target`之间的复杂关系（如“之内”、“之后”）。
    4.  **返回`Selection`:** 将找到的所有COM对象封装成一个`Selection`对象并返回。

*   **核心类:** `SelectorEngine`
*   **核心方法:** `select(backend: WordBackend, locator: dict) -> Selection`

#### **3.4. MCP 服务层 (`mcp_service.py`)**

系统的公共API入口，定义了AI可以调用的所有高级功能。

*   **职责:**
    1.  提供语义化的功能接口（如`insert_paragraph`, `apply_format`）。
    2.  管理`WordBackend`实例的生命周期，通常通过上下文管理器 (`with`语句) 来确保`cleanup`被调用。
    3.  将AI的调用翻译成对`SelectorEngine`和`Selection`对象的调用。

*   **示例实现:**
    ```python
    # mcp_service.py
    from selector import SelectorEngine
    from com_backend import WordBackend

    class McpService:
        def __init__(self):
            self.selector = SelectorEngine()

        def run_task(self, task_function, *args, **kwargs):
            backend = None
            try:
                backend = WordBackend()
                # 将backend实例传递给任务函数
                return task_function(backend, *args, **kwargs)
            finally:
                if backend:
                    backend.cleanup()

    # --- 示例任务 ---
    def task_apply_bold_to_first_paragraph(backend: WordBackend):
        service = McpService()
        locator = {"target": {"type": "paragraph", "filters": [{"index_in_parent": 0}]}}
        selection = service.selector.select(backend, locator)
        selection.apply_format({"bold": True})

    # --- 调用入口 ---
    # McpService().run_task(task_apply_bold_to_first_paragraph)
    ```

---

### **4. 核心数据结构：`Locator`查询语言**

`Locator`是驱动整个系统的声明式查询语言，其完整规范如下：

```json
{
  "anchor": {
    "type": "heading" | "bookmark" | "table" | "image" | "start_of_document",
    "identifier": {
      "text": "string",
      "level": "integer",
      "name": "string",
      "caption": "string",
      "index": "integer"
    }
  },
  "target": {
    "type": "paragraph" | "table" | "list_item" | "cell" | "image" | "run" | "end_of_document",
    "filters": [
      { "contains_text": "string" },
      { "text_matches_regex": "string" },
      { "style": "string" },
      { "is_bold": "boolean" },
      { "index_in_parent": "integer" },
      { "row_index": "integer" },
      { "column_index": "integer" }
    ]
  },
  "relation": {
    "type": "first_occurrence_after" | "all_occurrences_within" | "parent_of" | "immediately_following",
    "scope": "section" | "table" | "list" | "document"
  }
}
```

---

### **5. 错误处理与异常**

系统定义了一套自定义异常，以便于问题定位：
*   **`RuntimeError`:** 在`com_backend`中，当无法启动或连接到Word应用时抛出。
*   **`LocatorSyntaxError(ValueError)`:** 当`locator`对象本身格式或语法错误时由 **Locator Parser** 抛出。
*   **`ElementNotFoundError(LookupError)`:** 当`locator`语法正确，但在文档中找不到任何匹配元素时由 **Selector Engine** 抛出。
*   **`OperationUnsupportedError(TypeError)`:** 当对一个`Selection`对象执行其不支持的操作时抛出。

---

### **6. 迁移与部署**

#### **6.1. 现有项目迁移路径**
1.  **并行开发:** 优先创建并彻底测试`com_backend.py`和`selection.py`。
2.  **创建新接口:** 为现有功能创建新的、使用`locator`的函数版本。旧函数标记为`@deprecated`。
3.  **逐个迁移:** 逐个功能地用对新模块的调用来重写其实现。
4.  **逐步淘汰:** 在所有功能迁移完成后，移除旧的、不精确的函数。

#### **6.2. 部署要求**
1.  **操作系统:** Windows。
2.  **软件依赖:** 已安装Microsoft Office Word。
3.  **Python库:** `pywin32`。

---
这份文档为构建一个健壮、精确且可扩展的Word自动化服务提供了全面的蓝图。通过严格遵循分层和解耦的原则，即使面对COM接口的复杂性，项目也能保持清晰的结构和高度的可维护性。