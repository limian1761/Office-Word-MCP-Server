"""Operations package initialization."""

# 文档操作
from .document_ops import (
    create_document,
    open_document,
    close_document,
    save_document,
    get_document_structure,
)

# 文本操作
from .text_ops import (
    get_character_count,
    get_element_text,
    insert_text_before_range,
    insert_text_after_range,
    replace_element_text,
)

# 文本格式操作
from .text_format_ops import (
    set_bold_for_range,
    set_italic_for_range,
    set_font_size_for_range,
    set_font_name_for_range,
    set_font_color_for_range,
    set_alignment_for_range,
    set_paragraph_style,
)

# 段落操作
from .paragraphs_ops import (
    get_paragraphs_in_range,
    get_paragraphs_info,
    get_all_paragraphs,
)

# 评论操作
from .comment_ops import (
    add_comment,
    get_comments,
    get_comment_thread,
    delete_comment,
    delete_all_comments,
    edit_comment,
    reply_to_comment,
)

# 元素选择操作
from .element_selection_ops import (
    select_elements,
    get_element_by_id,
    batch_select_elements,
    batch_apply_formatting,
    delete_element_by_locator,
)

# 表格操作
from .table_ops import (
    create_table,
    get_cell_text,
    set_cell_text,
    get_table_info,
    insert_row,
    insert_column,
)

# 图片操作
from .image_ops import (
    get_image_info,
    insert_image,
    add_caption,
    resize_image,
    set_image_color_type,
)

# 文档对象操作（书签、引用）
from .document_objects_ops import (
    create_bookmark,
    get_bookmark,
    delete_bookmark,
    create_citation,
    create_hyperlink,
)

# 样式操作
from .styles_ops import (
    apply_formatting,
    set_font,
)

# 其他操作
from .others_ops import (
    get_document_statistics,
    compare_documents,
    convert_document_format,
    export_to_pdf,
    print_document,
    protect_document,
    unprotect_document,
)

__all__ = [
    # document_ops
    "create_document",
    "open_document",
    "close_document",
    "save_document",
    "get_document_structure",
    
    # text_ops
    "get_character_count",
    "get_element_text",
    "insert_text_before_range",
    "insert_text_after_range",
    "replace_element_text",
    
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
    
    # element_selection_ops
    "select_elements",
    "get_element_by_id",
    "batch_select_elements",
    "batch_apply_formatting",
    "delete_element_by_locator",
    
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
    
    # document_objects_ops
    "create_bookmark",
    "get_bookmark",
    "delete_bookmark",
    "create_citation",
    "create_hyperlink",
    
    # styles_ops
    "apply_formatting",
    "set_font",
    "set_paragraph_style",
    
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