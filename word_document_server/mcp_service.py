# mcp_service.py
from typing import Any, Dict, List

from word_document_server.com_backend import WordBackend
from word_document_server.selector import ElementNotFoundError, SelectorEngine


class McpService:
    """
    Provides a high-level API for interacting with Word documents.
    This is the main entry point for the MCP service.
    """
    def __init__(self):
        self.selector = SelectorEngine()

    def run_task(self, file_path: str, task_function, *args, **kwargs):
        """
        Executes a task within a managed WordBackend context.

        Args:
            file_path: The path to the document to open.
            task_function: A function that takes a WordBackend instance
                           as its first argument.
            *args: Additional arguments for the task function.
            **kwargs: Additional keyword arguments for the task function.
        """
        with WordBackend(file_path=file_path, visible=False) as backend:
            # Pass the service instance itself to the task function
            return task_function(self, backend, *args, **kwargs)

    def apply_format(self, backend: WordBackend, locator: Dict[str, Any], options: Dict[str, Any]):
        """
        Applies formatting to the elements specified by the locator.

        Args:
            backend: The active WordBackend instance.
            locator: The locator query to find the elements.
            options: A dictionary of formatting options.
        """
        selection = self.selector.select(backend, locator)
        selection.apply_format(options)
        print(f"Formatting applied to elements matching: {locator}")

    def insert_paragraph(self, backend: WordBackend, locator: Dict[str, Any], text: str, position: str = "after"):
        """
        Inserts a paragraph relative to the elements specified by the locator.

        Args:
            backend: The active WordBackend instance.
            locator: The locator query to find the anchor element.
            text: The text of the paragraph to insert.
            position: "before" or "after" the located element.
        """
        selection = self.selector.select(backend, locator)
        # The insert_text method in selection is a placeholder and needs to be
        # fully implemented to support 'before' and 'after' properly.
        # For now, we assume it inserts after.
        selection.insert_text(text, position=position)
        print(f"Paragraph inserted '{text}' near elements matching: {locator}")

    def delete(self, backend: WordBackend, locator: Dict[str, Any]):
        """
        Deletes the elements specified by the locator.

        Args:
            backend: The active WordBackend instance.
            locator: The locator query to find the elements to delete.
        """
        selection = self.selector.select(backend, locator)
        selection.delete()
        print(f"Elements deleted matching: {locator}")

    def add_table(self, backend: WordBackend, locator: Dict[str, Any], rows: int, cols: int, data: List[List[str]] = None, position: str = "after"):
        """
        Adds a table with optional data relative to a located element.

        Args:
            backend: The active WordBackend instance.
            locator: The locator for the element to position the table near.
            rows: Number of rows in the table.
            cols: Number of columns in the table.
            data: Optional list of lists of strings to populate the table.
            position: Currently only "after" is supported.
        """
        if position != "after":
            raise NotImplementedError("Position must be 'after'.")

        selection = self.selector.select(backend, locator)
        if not selection._elements:
            raise ElementNotFoundError(f"Could not find element for locator: {locator}")

        # Use the last element in the selection as the anchor
        anchor_element = selection._elements[-1]
        
        new_table = backend.add_table(anchor_element.Range, rows, cols)
        
        if data:
            for i, row_data in enumerate(data):
                if i < rows:
                    for j, cell_text in enumerate(row_data):
                        if j < cols:
                            new_table.Cell(Row=i + 1, Column=j + 1).Range.Text = cell_text
        
        print(f"Table added after element matching: {locator}")
        return new_table


# --- Example Task ---
def example_task_runner(service: McpService, backend: WordBackend):
    """
    An example task that demonstrates the use of the McpService.
    """
    # 1. Make the first paragraph bold
    print("1. Making the first paragraph bold...")
    first_p_locator = {"target": {"type": "paragraph", "filters": [{"index_in_parent": 0}]}}
    service.apply_format(backend, first_p_locator, {"bold": True})

    # 2. Insert a new paragraph after the (now bold) first paragraph
    print("\n2. Inserting a new paragraph...")
    service.insert_paragraph(backend, first_p_locator, "This is a newly inserted paragraph.")

    # 3. Delete the second paragraph (which was originally the second)
    print("\n3. Deleting the original second paragraph...")
    second_p_locator = {"target": {"type": "paragraph", "filters": [{"contains_text": "substring search"}]}}
    service.delete(backend, second_p_locator)
    
    print("\nExample task finished.")


if __name__ == '__main__':
    # This requires a test document to run.
    # You would need to provide a valid path to a docx file.
    # For example:
    # import os
    # TEST_DOC_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'tests', 'test_docs', 'test_document.docx'))
    # service = McpService()
    # service.run_task(TEST_DOC_PATH, example_task_runner)
    # print("Example task executed successfully.")
    pass
