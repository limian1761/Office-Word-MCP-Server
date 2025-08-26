import pytest
import os
import json
from word_document_server.tools.image import (
    insert_inline_picture, set_image_size, set_image_color_type,
    delete_image, get_image_info, add_picture_caption
)
from word_document_server.errors import WordDocumentError

@pytest.fixture
def image_test_setup():
    """创建测试环境并确保测试后正确清理"""
    from mcp.server.fastmcp.server import Context
    from word_document_server.tools.document import open_document, close_document, shutdown_word
    ctx = Context()
    
    # 确保测试前环境干净
    try:
        close_document(ctx)
    except:
        pass
    
    # 创建测试文档
    test_file = "tests/test_docs/image_test_doc.docx"
    open_document(ctx, file_path=test_file)
    
    # 准备测试SVG图片路径
    test_image = "tests/test_assets/test_image.svg"
    
    yield ctx, test_image
    
    # 测试后清理
    try:
        save_document(ctx)
        close_document(ctx)
    except:
        pass
    shutdown_word(ctx)


def test_insert_inline_picture(image_test_setup):
    """测试插入SVG图片"""
    ctx, test_image = image_test_setup
    result = insert_inline_picture(ctx, image_path=test_image)
    assert "successfully" in result.lower()


def test_set_image_size(image_test_setup):
    """测试调整图片大小"""
    ctx, test_image = image_test_setup
    insert_inline_picture(ctx, image_path=test_image)
    result = set_image_size(ctx, width=200, height=150, lock_aspect_ratio=True)
    assert "successfully" in result.lower()


def test_set_image_color_type(image_test_setup):
    """测试设置图片颜色类型"""
    ctx, test_image = image_test_setup
    insert_inline_picture(ctx, image_path=test_image)
    result = set_image_color_type(ctx, color_type="Grayscale")
    assert "successfully" in result.lower()


def test_delete_image(image_test_setup):
    """测试删除图片"""
    ctx, test_image = image_test_setup
    insert_inline_picture(ctx, image_path=test_image)
    result = delete_image(ctx)
    assert "successfully" in result.lower()


def test_get_image_info(image_test_setup):
    """测试获取图片信息"""
    ctx, test_image = image_test_setup
    insert_inline_picture(ctx, image_path=test_image)
    info = get_image_info(ctx)
    info_data = json.loads(info)
    
    assert isinstance(info_data, list)
    assert len(info_data) > 0
    assert info_data[0]["filename"] == "test_image.svg"


def test_add_picture_caption(image_test_setup):
    """测试添加图片标题"""
    ctx, test_image = image_test_setup
    insert_inline_picture(ctx, image_path=test_image)
    result = add_picture_caption(ctx, filename="test_image.svg", caption_text="测试图片标题")
    assert "successfully" in result.lower()


def test_image_operation_errors(image_test_setup):
    """测试图片操作错误处理"""
    ctx, test_image = image_test_setup
    # 关闭文档后尝试操作
    close_document(ctx)
    
    with pytest.raises(WordDocumentError):
        insert_inline_picture(ctx, image_path=test_image)
    
    with pytest.raises(WordDocumentError):
        get_image_info(ctx)