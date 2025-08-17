Objective: You are to continue the development of a Python-based MCP (Model Context Protocol) server
  designed for the automation of Microsoft Word. The project's goal is to provide a comprehensive and robust
   set of tools for an LLM to perform complex, context-aware document manipulation.

  Core Technologies:
   * Python 3.11+
   * Microsoft Word COM Interface: All interactions with Word are handled via the pywin32 library.
   * Model Context Protocol (MCP): The server is built using the `mcp[cli]` package.

  ---

  ### **Project Roadmap & Future Work**

  With the core engine and initial toolset now stable, the project will focus on expanding its capabilities, improving robustness, and adding advanced, high-value features.

  #### **Priority 1: Enhancing Robustness & Core Capabilities**

  *   **Task 1.1: Advanced Formatting Tool**
    *   **Description:** Create a single, powerful `apply_format` tool that accepts a `locator` and a dictionary of formatting options. This will consolidate formatting logic and provide a much wider range of capabilities.
    *   **Implementation:**
        1.  In `selection.py`, add a new `apply_format(self, formatting_options: dict)` method to the `Selection` class.
        2.  This method should handle options like `font_size`, `font_color`, `italic`, `underline`, `alignment` (left/center/right), etc.
        3.  In `app.py`, create the `@mcp.tool() def apply_format(...)` tool that calls the new selection method.
        4.  Add comprehensive tests for various formatting combinations.

  *   **Task 1.2: Refined Error Handling**
    *   **Description:** Implement more specific exceptions to give the LLM better feedback on failed operations.
    *   **Implementation:**
        1.  In `selector.py`, create and raise an `AmbiguousLocatorError` if a locator, which is expected to return a single element, finds multiple.
        2.  In `com_backend.py`, wrap key COM calls in `try...except` blocks to catch `pywintypes.com_error` and re-raise them as a more user-friendly `WordComError`, providing context on what failed.

  #### **Priority 2: Expanding Document Element Support**

  *   **Task 2.1: Implement Header & Footer Support**
    *   **Description:** Allow the selector to target and manipulate headers and footers. This is critical for tasks like adding page numbers or changing document titles.
    *   **Implementation:**
        1.  In `com_backend.py`, add methods like `get_headers()` and `get_footers()`. Note that Word has primary, first-page, and even-page headers/footers per section. The initial implementation can target the primary ones.
        2.  In `selector.py`, add `"header"` and `"footer"` as supported `element_type` values.

  *   **Task 2.2: Implement Basic List Support**
    *   **Description:** Provide tools to read and create simple bulleted or numbered lists. The COM API for lists is complex, so the initial focus will be on identifying existing list paragraphs and creating new ones.
    *   **Implementation:**
        1.  In `com_backend.py`, investigate the `Paragraph.Range.ListFormat` properties to identify list items.
        2.  In `selector.py`, add a `is_list_item` filter.
        3.  In `app.py`, create a `create_bulleted_list(ctx: Context, locator: dict, items: list[str])` tool.

  #### **Priority 3: Advanced Tooling & Context Awareness**

  *   **Task 3.1: Implement Document Structure as a Resource**
    *   **Description:** Provide a structured overview of the document (e.g., a table of contents based on headings) as a readable MCP Resource. This allows the LLM to understand the document's layout *before* attempting to modify it.
    *   **Implementation:**
        1.  In `app.py`, define a new function decorated with `@mcp.resource("/document/structure")`.
        2.  This function will use the `com_backend` to get all headings (`H1`, `H2`, etc.) and return them as a nested JSON object.

  *   **Task 3.2: Implement "Track Changes" Tools**
    *   **Description:** A killer feature for editing workflows. Provide tools to accept or reject tracked changes in a document.
    *   **Implementation:**
        1.  In `com_backend.py`, add methods to interact with `document.Revisions`, such as `get_revisions()` and `accept_all_revisions()`.
        2.  In `app.py`, create tools like `accept_all_changes(ctx: Context)` and `get_summary_of_changes(ctx: Context)`.

  ---

  ### **Completed Development Milestones**

  The following features have been implemented by following a test-driven development approach.

  *   **Priority 1: Finalize the Selector Engine**
    *   **Task 1.1: Implement `run` Target Type:** A "run" (a contiguous sequence of characters with the same formatting) is now a supported target type in the selector.
    *   **Task 1.2: Implement High-Priority Filters:** The `style` filter is fully implemented and tested.
    *   **Task 1.3: Implement Core Relations:** The `parent_of` and `immediately_following` relations are implemented.

  *   **Priority 2: Expand the Toolset**
    *   **Task 2.1: Create `replace_text` Tool:** A robust tool for finding an element via a locator and replacing its entire text content.
    *   **Task 2.2: Create More Table Tools:** Implemented `set_cell_value` to change the text of a single cell and `create_table` to add a new table at a specified location.

  *   **Priority 3: Official MCP Conformance**
    *   **Task 3.1: Synchronize `mcp-config.json`:** The `mcp-config.json` file is synchronized with the tools defined in `app.py`, ensuring discoverability by MCP clients.

  ---

  ### **Important Considerations**

   * **COM API Quirks:** The Word COM API is notoriously difficult. The attempts to implement `list_item` and `image` selection in the past failed due to unreliable behavior. New attempts should be focused and well-tested. Stick to features with more reliable COM properties.
   * **Localization:** Remember that style names can be localized (e.g., "Heading 1" vs. "标题 1"). The current heading detection logic accounts for this. Be mindful of this if you implement style-based filters.
   * **Testing:** The project's stability has been achieved through a rigorous, iterative test-driven approach. Adhere strictly to the "add a failing test, then implement" workflow for all new features.