import pytest
from word_document_server.tools.comment import add_comment, get_comments, delete_comment, delete_all_comments
from word_document_server.errors import WordDocumentError

@pytest.fixture
def setup_document():
    # Setup code to create a test document
    from word_document_server.tools.document import open_document
    from word_document_server.core_utils import get_session_context
    ctx = get_session_context()
    open_document(ctx, file_path="tests/test_docs/test_document.docx")
    yield ctx
    # Teardown code
    from word_document_server.tools.document import close_document
    close_document(ctx)


def test_add_comment(setup_document):
    ctx = setup_document
    result = add_comment(ctx, text="Test comment", author="Test Author")
    assert "successfully" in result.lower()


def test_get_comments(setup_document):
    ctx = setup_document
    add_comment(ctx, text="Test comment for retrieval", author="Test Author")
    comments = get_comments(ctx)
    assert isinstance(comments, str)
    assert "Test comment for retrieval" in comments


def test_delete_comment(setup_document):
    ctx = setup_document
    add_comment(ctx, text="Test comment for deletion", author="Test Author")
    comments = get_comments(ctx)
    assert len(eval(comments)) > 0
    result = delete_comment(ctx, comment_index=0)
    assert "successfully" in result.lower()
    comments_after_delete = get_comments(ctx)
    assert len(eval(comments_after_delete)) == 0


def test_delete_all_comments(setup_document):
    ctx = setup_document
    add_comment(ctx, text="Test comment 1", author="Test Author")
    add_comment(ctx, text="Test comment 2", author="Test Author")
    result = delete_all_comments(ctx)
    assert "successfully" in result.lower()
    comments_after_delete = get_comments(ctx)
    assert len(eval(comments_after_delete)) == 0


def test_add_comment_error_handling(setup_document):
    ctx = setup_document
    with pytest.raises(WordDocumentError):
        add_comment(ctx, text=None, author="Test Author")