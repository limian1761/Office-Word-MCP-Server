# create_test_doc.py
import os
import win32com.client
from win32com.client import gencache

def create_test_document(file_path):
    """
    Creates a Word document with varied formatting for testing purposes.
    This version is more robust and uses language-independent constants.
    """
    word_app = None
    doc = None
    try:
        dir_name = os.path.dirname(file_path)
        if not os.path.exists(dir_name):
            os.makedirs(dir_name)
            
        # Use EnsureDispatch to generate support for constants if not present
        word_app = gencache.EnsureDispatch("Word.Application")
        word_app.Visible = False
        doc = word_app.Documents.Add()
        
        # Get constants from the application's type library
        constants = win32com.client.constants

        # Start at the beginning of the document
        current_range = doc.Range(0, 0)

        def add_paragraph(text, style_const=None, is_bold=False):
            """Helper to add a formatted paragraph at the current range."""
            nonlocal current_range
            current_range.InsertAfter(text + '\n')
            # The new paragraph is the one we just inserted.
            # Its range is from the start of the insertion to the end.
            new_para_range = doc.Range(current_range.End - len(text) - 1, current_range.End -1)
            
            if style_const:
                new_para_range.Style = doc.Styles(style_const)
            if is_bold:
                new_para_range.Font.Bold = True
            
            # Move current_range to the end of the document for the next insertion
            current_range = doc.Range(doc.Content.End - 1, doc.Content.End - 1)


        # --- Add Content in Order ---
        doc.TrackRevisions = True # Turn on Track Changes
        add_paragraph("First paragraph.")
        
        # Add a paragraph and then delete it to create a revision
        current_range.InsertAfter("This sentence will be deleted.\n")
        deleted_range = doc.Range(current_range.End - len("This sentence will be deleted.\n"), current_range.End - 1)
        deleted_range.Delete()
        
        doc.TrackRevisions = False # Turn off Track Changes for the rest
        
        add_paragraph("A paragraph for substring search.")
        add_paragraph("This is a bold paragraph.", is_bold=True)
        add_paragraph("This is a heading.", style_const=constants.wdStyleHeading1)
        add_paragraph("Paragraph with unique_word_123 for regex.")

        # --- Add a Table ---
        table_range = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        table = doc.Tables.Add(table_range, 3, 2) # 3 rows, 2 columns
        table.Cell(1, 1).Range.Text = "Table Cell 1-1"
        table.Cell(1, 2).Range.Text = "Table Cell 1-2"
        table.Cell(2, 1).Range.Text = "Paragraph inside table."
        table.Cell(2, 2).Range.Text = "Another paragraph in table."
        table.Cell(3, 1).Range.Text = "Last cell."
        
        # Move current_range after the table
        current_range = doc.Range(doc.Content.End - 1, doc.Content.End - 1)

        # --- Add a Bulleted List ---
        # Insert and format the first list item
        current_range.InsertAfter("List item 1\n")
        list_para1_range = doc.Range(current_range.End - len("List item 1\n"), current_range.End - 1)
        list_para1_range.ListFormat.ApplyBulletDefault()
        current_range = doc.Range(doc.Content.End - 1, doc.Content.End - 1)

        # Insert and format the second list item
        current_range.InsertAfter("List item 2\n")
        list_para2_range = doc.Range(current_range.End - len("List item 2\n"), current_range.End - 1)
        # Word will automatically continue the list, but we call it again for robustness
        list_para2_range.ListFormat.ApplyBulletDefault()
        current_range = doc.Range(doc.Content.End - 1, doc.Content.End - 1)

        add_paragraph("This paragraph is outside the table.")
        add_paragraph("The last paragraph.")

        doc.SaveAs(file_path)
        print(f"Test document created at: {file_path}")
        
    except Exception as e:
        print(f"Error creating test document: {e}")
    finally:
        if doc:
            doc.Close(SaveChanges=False)
        if word_app:
            word_app.Quit()

if __name__ == "__main__":
    current_dir = os.path.dirname(os.path.abspath(__file__))
    test_doc_path = os.path.join(current_dir, 'tests', 'test_docs', 'test_document.docx')
    create_test_document(os.path.abspath(test_doc_path))
