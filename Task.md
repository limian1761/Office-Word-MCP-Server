Objective: You are to continue the development of a Python-based MCP (Model Context Protocol) server
  designed for the automation of Microsoft Word. The project's goal is to provide a comprehensive and robust
   set of tools for an LLM to perform complex, context-aware document manipulation.

  Core Technologies:
   * Python 3.11+
   * Microsoft Word COM Interface: All interactions with Word are handled via the pywin32 library.
   * Model Context Protocol (MCP): The server is built using the python-sdk, which is managed as a Git Submodule
     in the `python-sdk/` directory.

  ---

  Current Project Status & Architecture

  The project is in a stable, working state with a well-defined architecture.

  1. Architecture Overview (`Architecture.md`):
  The system is built on a "separation of concerns" principle:
   * MCP Server Layer (`word_document_server/app.py`): This is the main entry point. It uses the FastMCP class
      from the local SDK (mcp.server.fastmcp.server) to define and register tools using the @mcp.tool()
     decorator.
   * Selector Engine (`word_document_server/selector.py`): This is the "brain" of the server. It parses a
     JSON-based Locator object to find specific elements within the document.
   * Selection Abstraction (`word_document_server/selection.py`): This is the "hands." Once the Selector finds
      element(s), it wraps them in a Selection object, which provides methods to act upon them (e.g.,
     delete(), insert_text()).
   * COM Backend (`word_document_server/com_backend.py`): This is the low-level layer that directly
     communicates with the Word application via pywin32.

  2. State Management:
   * A stateful approach is used to manage the Word application instance. A single WordBackend object is
     created and stored on the MCP Context object (ctx.set_state("word_backend", ...)) to ensure all tool
     calls in a session use the same Word instance.
   * The open_document tool must be called first to initialize this backend.
   * The shutdown_word tool must be called at the end to close the application.

  3. Current Selector Capabilities:
   * Supported Target Types: paragraph, heading, table, cell.
   * Supported Filters: index_in_parent, contains_text, text_matches_regex, is_bold, row_index, column_index.
   * Supported Anchors: start_of_document, end_of_document, and object-based anchors (e.g., find a table with
     index: 0).
   * Supported Relations: all_occurrences_within, first_occurrence_after.

  4. Implemented Tools (`word_document_server/app.py`):
   * open_document, shutdown_word
   * insert_paragraph, get_text_from_cell, get_text, delete_element

  ---

  Completed Development Milestones

  The following features have been implemented by following a test-driven development approach.

  New Feature Workflow:
   1. Add a new test case to the relevant file in the `tests/` directory. For new functionality, a new test
      file may be created.
   2. Run `pytest` from the project root and observe the new test failing.
   3. Implement the required changes in the application source code (`word_document_server/`).
   4. Run `pytest` again and confirm that all tests pass.

  Priority 1: Finalize the Selector Engine

  Task 1.1: Implement `run` Target Type
  A "run" is a contiguous sequence of characters with the same formatting (e.g., a bold word in a sentence).
   1. In `com_backend.py`: Create a get_runs_in_range method that iterates through range_obj.Runs.
   2. In `selector.py`: Add "run" as a supported element_type in _get_initial_candidates.
   ✅ Completed

  Task 1.2: Implement Remaining High-Priority Filters
   1. In `selector.py`: Implement and test the style filter (_filter_by_style). The create_test_doc.py already
      creates a heading with a specific style you can use for testing.
   ✅ Completed

  Task 1.3: Implement Remaining Relations
 1. In `selector.py`: Implement the logic for parent_of and immediately_following relations within the
     _select_relative_to_anchor method.
 ✅ Completed

  Priority 2: Expand the Toolset

  Task 2.1: Create a `replace_text` Tool
   1. In `selection.py`: The insert_text method already has a "replace" position. Ensure this is robust.
   2. In `app.py`: Create a new tool @mcp.tool() def replace_text(ctx: Context, locator: dict, new_text: str).
      This tool will use the selector to find an element and then call selection.insert_text(new_text,
      position="replace").
   ✅ Completed

  Task 2.2: Create More Table Tools
   1. In `app.py`: Create a set_cell_value(ctx: Context, locator: dict, text: str) tool. The locator should
      resolve to a single cell. The tool will set the cell.Range.Text property.
   2. In `app.py`: Create a create_table(ctx: Context, locator: dict, rows: int, cols: int) tool. It will use
      the locator to find an anchor point and call the add_table method in the backend.
   ✅ Completed

  Priority 3: Official MCP Conformance

  Task 3.1: Synchronize `mcp-config.json`
   1. Review the mcp-config.json file. The SDK may provide a way to automatically generate this file from the
      registered tools. If not, you will need to manually add JSON definitions for each tool created in app.py
      to make them discoverable by an MCP client.
   ✅ Completed

  ---

  Important Considerations

   * COM API Quirks: The Word COM API is notoriously difficult. The attempts to implement list_item and image
     selection failed due to unreliable behavior. Do not attempt to re-implement these. Stick to features with
      more reliable COM properties (paragraphs, tables, cells, runs, text, formatting).
   * Localization: Remember that style names can be localized (e.g., "Heading 1" vs. "标题 1"). The current
     heading detection logic accounts for this. Be mindful of this if you implement style-based filters.
   * Testing: The project's stability has been achieved through a rigorous, iterative test-driven approach.
     Adhere strictly to the "add a failing test, then implement" workflow.