"""Operations package initialization."""

# 文档操作
# 评论操作
from .comment_ops import (add_comment, delete_all_comments, delete_comment,
                          edit_comment, get_comment_thread, get_comments,
                          reply_to_comment)
from .document_ops import (close_document, create_document,
                           get_document_outline, open_document,
                           save_document)
# 图片操作
from .image_ops import (add_caption, get_image_info, insert_image,
                        resize_image, set_image_color_type)
# 文档对象操作（书签、引用）
from .objects_ops import (create_bookmark, create_citation, create_hyperlink,
                          delete_bookmark, get_bookmark)
# 其他操作
from .others_ops import (compare_documents, convert_document_format,
                         export_to_pdf, get_document_statistics,
                         print_document, protect_document, unprotect_document)
# 段落操作
from .paragraphs_ops import (get_all_paragraphs, get_paragraphs_in_range,
                             get_paragraphs_info)
# 元素选择操作
from .range_ops import (batch_apply_formatting, batch_select_objects,
                        delete_object_by_locator, get_object_by_id,
                        select_objects)
# 样式操作
from .styles_ops import apply_formatting, set_font
# 导航工具操作
# 专注于上下文管理和活动对象设置
from .navigate_tools import (
    set_active_context,
    set_active_object
)
# 表格操作
from .table_ops import (create_table, get_cell_text, get_table_info,
                        insert_column, insert_row, set_cell_text)
# 文本格式操作
from .text_format_ops import (set_alignment_for_range, set_bold_for_range,
                              set_font_color_for_range,
                              set_font_name_for_range, set_font_size_for_range,
                              set_italic_for_range, set_paragraph_style)
# 文本操作
from .text_operations import (get_character_count, get_object_text, insert_text,
                              insert_text_after_range, insert_text_before_range,
                              replace_object_text)

__all__ = [
    # document_ops
    "create_document",
    "open_document",
    "close_document",
    "save_document",
    "get_document_outline",
    # text_ops
    "get_character_count",
    "get_object_text",
    "insert_text_before_range",
    "insert_text_after_range",
    "replace_object_text",
    # text_format_ops
    "set_bold_for_range",
    "set_italic_for_range",
    "set_font_size_for_range",
    "set_font_name_for_range",
    "set_font_color_for_range",
    "set_alignment_for_range",
    "set_paragraph_style",
    # paragraphs_ops
    "get_paragraphs_in_range",
    "get_paragraphs_info",
    "get_all_paragraphs",
    # comment_ops
    "add_comment",
    "get_comments",
    "get_comment_thread",
    "delete_comment",
    "delete_all_comments",
    "edit_comment",
    "reply_to_comment",
    # object_selection_ops
    "select_objects",
    "get_object_by_id",
    "batch_select_objects",
    "batch_apply_formatting",
    "delete_object_by_locator",
    # table_ops
    "create_table",
    "get_cell_text",
    "set_cell_text",
    "get_table_info",
    "insert_row",
    "insert_column",
    # image_ops
    "get_image_info",
    "insert_image",
    "add_caption",
    "resize_image",
    "set_image_color_type",
    # objects_ops
    "create_bookmark",
    "get_bookmark",
    "delete_bookmark",
    "create_citation",
    "create_hyperlink",
    # styles_ops
    "apply_formatting",
    "set_font",
    "set_paragraph_style",
    # navigate_tools
    "set_active_context",
    "set_active_object",
    # others_ops
    "get_document_statistics",
    "compare_documents",
    "convert_document_format",
    "export_to_pdf",
    "print_document",
    "protect_document",
    "unprotect_document",
]

# Version information
__version__ = "1.1.9"
