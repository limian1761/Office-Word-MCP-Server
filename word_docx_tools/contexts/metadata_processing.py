import time
import hashlib
from typing import Dict, List, Optional, Any, Set, Tuple, Union
from datetime import datetime


class MetadataProcessor:
    """元数据处理器，负责元数据的创建、更新、验证和查询"""
    def __init__(self):
        # 支持的元数据类型及其默认验证器
        self._supported_metadata_types = {
            "string": self._validate_string,
            "number": self._validate_number,
            "boolean": self._validate_boolean,
            "datetime": self._validate_datetime,
            "array": self._validate_array,
            "object": self._validate_object,
            "null": self._validate_null
        }
        # 保留的元数据字段
        self._reserved_fields = {
            "id", "type", "title", "created_time", "last_updated", "created_by", "last_modified_by"
        }
        # 默认的元数据字段
        self._default_metadata_fields = {
            "created_time": lambda: time.time(),
            "last_updated": lambda: time.time()
        }

    def create_metadata(self, object_type: str, **kwargs) -> Dict[str, Any]:
        """
        创建标准化的元数据字典
        
        Args:
            object_type: 对象类型
            **kwargs: 其他元数据字段
        
        Returns:
            包含标准化元数据的字典
        """
        metadata = {}
        
        # 添加默认字段
        for field, value_provider in self._default_metadata_fields.items():
            if field not in kwargs:
                metadata[field] = value_provider()
        
        # 添加传入的字段，但覆盖默认字段
        for field, value in kwargs.items():
            if field in self._reserved_fields:
                metadata[field] = value
            else:
                # 非保留字段添加到自定义元数据
                if "custom_metadata" not in metadata:
                    metadata["custom_metadata"] = {}
                metadata["custom_metadata"][field] = value
        
        # 确保type字段存在
        if "type" not in metadata:
            metadata["type"] = object_type
        
        # 如果没有ID，生成一个
        if "id" not in metadata:
            metadata["id"] = self._generate_id(metadata)
        
        return metadata

    def update_metadata(self, 
                        existing_metadata: Dict[str, Any], 
                        update_fields: Dict[str, Any],
                        update_timestamp: bool = True) -> Dict[str, Any]:
        """
        更新现有元数据
        
        Args:
            existing_metadata: 现有元数据
            update_fields: 要更新的字段
            update_timestamp: 是否更新last_updated字段
        
        Returns:
            更新后的元数据
        """
        updated_metadata = existing_metadata.copy()
        
        # 更新字段
        for field, value in update_fields.items():
            if field in self._reserved_fields:
                # 对于保留字段，直接更新
                updated_metadata[field] = value
            else:
                # 对于非保留字段，更新到custom_metadata中
                if "custom_metadata" not in updated_metadata:
                    updated_metadata["custom_metadata"] = {}
                updated_metadata["custom_metadata"][field] = value
        
        # 更新时间戳
        if update_timestamp:
            updated_metadata["last_updated"] = time.time()
        
        return updated_metadata

    def validate_metadata(self, 
                         metadata: Dict[str, Any], 
                         schema: Optional[Dict[str, Any]] = None) -> Tuple[bool, List[str]]:
        """
        验证元数据是否符合指定的模式
        
        Args:
            metadata: 要验证的元数据
            schema: 元数据模式
        
        Returns:
            验证结果和错误消息列表
        """
        errors = []
        
        # 验证基本字段
        if "type" not in metadata:
            errors.append("Missing required field: type")
        
        # 如果提供了模式，进行更详细的验证
        if schema:
            for field_name, field_schema in schema.items():
                field_type = field_schema.get("type", "any")
                is_required = field_schema.get("required", False)
                
                # 检查必需字段
                if is_required and field_name not in metadata:
                    errors.append(f"Missing required field: {field_name}")
                
                # 如果字段存在，验证类型
                if field_name in metadata:
                    validator = self._supported_metadata_types.get(field_type)
                    if validator and not validator(metadata[field_name]):
                        errors.append(f"Field {field_name} must be of type {field_type}")
        
        return len(errors) == 0, errors

    def extract_metadata(self, 
                        metadata: Dict[str, Any], 
                        fields: List[str]) -> Dict[str, Any]:
        """
        从元数据中提取指定的字段
        
        Args:
            metadata: 源元数据
            fields: 要提取的字段列表
        
        Returns:
            包含指定字段的元数据子集
        """
        result = {}
        
        for field in fields:
            if field in metadata:
                result[field] = metadata[field]
            elif "custom_metadata" in metadata and field in metadata["custom_metadata"]:
                result[field] = metadata["custom_metadata"][field]
        
        return result

    def search_metadata(self, 
                       metadata_list: List[Dict[str, Any]], 
                       query: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        根据查询条件搜索元数据
        
        Args:
            metadata_list: 元数据列表
            query: 查询条件
        
        Returns:
            匹配的元数据列表
        """
        results = []
        
        for metadata in metadata_list:
            match = True
            
            for key, value in query.items():
                # 检查是否匹配字段
                if key in metadata:
                    if metadata[key] != value:
                        match = False
                        break
                elif "custom_metadata" in metadata and key in metadata["custom_metadata"]:
                    if metadata["custom_metadata"][key] != value:
                        match = False
                        break
                else:
                    match = False
                    break
            
            if match:
                results.append(metadata)
        
        return results

    def merge_metadata(self, 
                      metadata1: Dict[str, Any], 
                      metadata2: Dict[str, Any],
                      prefer_first: bool = True) -> Dict[str, Any]:
        """
        合并两个元数据字典
        
        Args:
            metadata1: 第一个元数据
            metadata2: 第二个元数据
            prefer_first: 是否优先保留第一个元数据的值
        
        Returns:
            合并后的元数据
        """
        merged = metadata1.copy()
        
        # 合并基础字段
        for field, value in metadata2.items():
            if field == "custom_metadata":
                # 特殊处理custom_metadata字段
                if "custom_metadata" not in merged:
                    merged["custom_metadata"] = {}
                self._merge_custom_metadata(
                    merged["custom_metadata"], 
                    value, 
                    prefer_first
                )
            elif field in merged and prefer_first:
                # 如果字段已存在且prefer_first为True，则保留第一个的值
                continue
            else:
                # 其他情况直接覆盖或添加
                merged[field] = value
        
        # 更新last_updated字段
        merged["last_updated"] = time.time()
        
        return merged

    def calculate_metadata_hash(self, metadata: Dict[str, Any]) -> str:
        """
        计算元数据的哈希值，用于检测变化
        
        Args:
            metadata: 元数据
        
        Returns:
            元数据的哈希值
        """
        # 排除变化频繁的字段
        metadata_copy = metadata.copy()
        if "created_time" in metadata_copy:
            del metadata_copy["created_time"]
        if "last_updated" in metadata_copy:
            del metadata_copy["last_updated"]
        
        # 转换为字符串并计算哈希
        metadata_str = str(sorted(metadata_copy.items()))
        return hashlib.md5(metadata_str.encode()).hexdigest()

    def format_metadata_for_display(self, metadata: Dict[str, Any]) -> Dict[str, Any]:
        """
        格式化元数据以便显示
        
        Args:
            metadata: 原始元数据
        
        Returns:
            格式化后的元数据
        """
        formatted = {}
        
        for key, value in metadata.items():
            if key in ["created_time", "last_updated"] and isinstance(value, (int, float)):
                # 格式化时间戳
                try:
                    formatted[key] = datetime.fromtimestamp(value).strftime('%Y-%m-%d %H:%M:%S')
                except:
                    formatted[key] = value
            elif key == "custom_metadata" and isinstance(value, dict):
                # 格式化自定义元数据
                formatted[key] = self.format_metadata_for_display(value)
            else:
                # 其他字段保持不变
                formatted[key] = value
        
        return formatted

    def _merge_custom_metadata(self, 
                              target: Dict[str, Any], 
                              source: Dict[str, Any],
                              prefer_first: bool = True) -> None:
        """
        合并自定义元数据
        
        Args:
            target: 目标元数据
            source: 源元数据
            prefer_first: 是否优先保留目标的值
        """
        for field, value in source.items():
            if field in target and prefer_first:
                # 如果字段已存在且prefer_first为True，则保留目标的值
                continue
            else:
                # 其他情况直接覆盖或添加
                target[field] = value

    def _generate_id(self, metadata: Dict[str, Any]) -> str:
        """
        为元数据生成唯一ID
        
        Args:
            metadata: 元数据
        
        Returns:
            生成的唯一ID
        """
        # 使用时间戳和类型信息生成ID
        base_id = f"{metadata.get('type', 'unknown')}_{int(time.time() * 1000)}"
        # 如果有标题，也加入ID生成
        if 'title' in metadata:
            title_part = hashlib.md5(metadata['title'].encode()).hexdigest()[:6]
            base_id = f"{base_id}_{title_part}"
        return base_id

    def _validate_string(self, value: Any) -> bool:
        """验证值是否为字符串"""
        return isinstance(value, str)

    def _validate_number(self, value: Any) -> bool:
        """验证值是否为数字"""
        return isinstance(value, (int, float)) and not isinstance(value, bool)

    def _validate_boolean(self, value: Any) -> bool:
        """验证值是否为布尔值"""
        return isinstance(value, bool)

    def _validate_datetime(self, value: Any) -> bool:
        """验证值是否为日期时间"""
        return isinstance(value, (int, float, datetime))

    def _validate_array(self, value: Any) -> bool:
        """验证值是否为数组"""
        return isinstance(value, list)

    def _validate_object(self, value: Any) -> bool:
        """验证值是否为对象"""
        return isinstance(value, dict)

    def _validate_null(self, value: Any) -> bool:
        """验证值是否为null"""
        return value is None


# 创建全局元数据处理器实例
global_metadata_processor = MetadataProcessor()


def get_metadata_processor() -> MetadataProcessor:
    """
    获取全局元数据处理器实例
    """
    return global_metadata_processor


def create_document_metadata(document_title: str = "", document_path: str = "", **kwargs) -> Dict[str, Any]:
    """
    创建文档元数据
    
    Args:
        document_title: 文档标题
        document_path: 文档路径
        **kwargs: 其他元数据字段
    
    Returns:
        文档元数据
    """
    processor = get_metadata_processor()
    return processor.create_metadata(
        "document",
        title=document_title or "Untitled Document",
        path=document_path,
        **kwargs
    )


def create_section_metadata(section_index: int, **kwargs) -> Dict[str, Any]:
    """
    创建节元数据
    
    Args:
        section_index: 节索引
        **kwargs: 其他元数据字段
    
    Returns:
        节元数据
    """
    processor = get_metadata_processor()
    return processor.create_metadata(
        "section",
        title=f"Section {section_index}",
        index=section_index,
        **kwargs
    )


def create_paragraph_metadata(paragraph_id: str, text_preview: str = "", style: str = "Normal", **kwargs) -> Dict[str, Any]:
    """
    创建段落元数据
    
    Args:
        paragraph_id: 段落ID
        text_preview: 文本预览
        style: 段落样式
        **kwargs: 其他元数据字段
    
    Returns:
        段落元数据
    """
    processor = get_metadata_processor()
    return processor.create_metadata(
        "paragraph",
        id=paragraph_id,
        text_preview=text_preview,
        style=style,
        **kwargs
    )


def create_table_metadata(table_id: str, rows: int = 0, columns: int = 0, **kwargs) -> Dict[str, Any]:
    """
    创建表格元数据
    
    Args:
        table_id: 表格ID
        rows: 行数
        columns: 列数
        **kwargs: 其他元数据字段
    
    Returns:
        表格元数据
    """
    processor = get_metadata_processor()
    return processor.create_metadata(
        "table",
        id=table_id,
        rows=rows,
        columns=columns,
        cell_count=rows * columns if rows > 0 and columns > 0 else 0,
        **kwargs
    )


def create_image_metadata(image_id: str, width: int = 0, height: int = 0, **kwargs) -> Dict[str, Any]:
    """
    创建图片元数据
    
    Args:
        image_id: 图片ID
        width: 宽度
        height: 高度
        **kwargs: 其他元数据字段
    
    Returns:
        图片元数据
    """
    processor = get_metadata_processor()
    return processor.create_metadata(
        "image",
        id=image_id,
        width=width,
        height=height,
        **kwargs
    )