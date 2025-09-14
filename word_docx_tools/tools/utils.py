"""
Utility functions for Word Document MCP Server tools.
This module provides common functionality used across different tool implementations.
"""
from typing import Optional, Dict, Any

class LocatorValidationError(ValueError):
    """当locator验证失败时抛出的异常"""


def validate_locator(parsed_locator):
    """验证locator字典的有效性"""
    # 检查是否为字典
    if not isinstance(parsed_locator, dict):
        raise LocatorValidationError(f"Locator must be a dictionary, got {type(parsed_locator).__name__}.")

    # 验证基本结构和必需字段
    required_fields = ["type"]
    for field in required_fields:
        if field not in parsed_locator:
            raise LocatorValidationError(f"Required field '{field}' missing in locator.")

    object_type = parsed_locator["type"]

    # 检查必需的type字段
    if not object_type:
        raise LocatorValidationError("Locator must specify an object type.")

    # 验证对象类型
    valid_object_types = [
        "paragraph", "table", "cell", "inline_shape", 
        "image", "comment", "range", "selection", 
        "document", "document_start", "document_end"
    ]
    if object_type not in valid_object_types:
        raise LocatorValidationError(
            f"Invalid object type '{object_type}'. Valid types are: {', '.join(valid_object_types)}"
        )

    # 如果指定了anchor，验证relation
    if parsed_locator.get("anchor") is not None:
        valid_relations = [
            "all_occurrences_within",
            "first_occurrence_after",
            "parent_of",
            "immediately_following",
        ]
        if (
            parsed_locator.get("relation") is not None
            and parsed_locator["relation"] not in valid_relations
        ):
            raise LocatorValidationError(
                f"Invalid relation '{parsed_locator['relation']}'. Valid relations are: {', '.join(valid_relations)}"
            )
        
        # 确保anchor必须提供relation
        if "relation" not in parsed_locator:
            raise LocatorValidationError("Locator with 'anchor' must also specify a 'relation'.")

    # 如果有filters，验证其格式
    if "filters" in parsed_locator:
        if not isinstance(parsed_locator["filters"], list):
            raise LocatorValidationError("'filters' must be a list.")
        
        # 验证filters列表中的每个filter
        valid_filter_types = [
            "index", "contains_text", "text_matches_regex", 
            "shape_type", "style", "is_bold", "row_index", 
            "column_index", "table_index", "is_list_item", 
            "range_start", "range_end", "has_style"
        ]
        
        for i, filter_item in enumerate(parsed_locator["filters"]):
            if not isinstance(filter_item, dict) or len(filter_item) != 1:
                raise LocatorValidationError(f"Filter at index {i} must be a single key-value pair dictionary.")
            
            filter_name = next(iter(filter_item.keys()))
            if filter_name not in valid_filter_types:
                raise LocatorValidationError(
                    f"Invalid filter type '{filter_name}' at index {i}. "
                    f"Valid filter types are: {', '.join(valid_filter_types)}"
                )

    # 类型特定的验证
    if object_type == "paragraph":
        # 对于段落，如果值是数字且treat_as_index为True，确保它是正整数
        value = parsed_locator.get("value", "")
        treat_as_index = parsed_locator.get("treat_as_index", False)
        
        if treat_as_index and value:
            try:
                index_value = int(str(value))
                if index_value <= 0:
                    raise LocatorValidationError(f"Paragraph index must be a positive integer, got {index_value}.")
            except ValueError:
                raise LocatorValidationError(
                    f"Cannot treat paragraph value '{value}' as index - must be a number."
                )

    elif object_type == "table":
        # 对于表格，如果值是数字，确保它是正整数
        value = parsed_locator.get("value", "")
        if value and str(value).isdigit():
            table_index = int(str(value))
            if table_index <= 0:
                raise LocatorValidationError(f"Table index must be a positive integer, got {table_index}.")

    elif object_type in ["document_start", "document_end"]:
        # 对于document_start和document_end，确保没有冲突的参数
        if parsed_locator.get("value") or parsed_locator.get("filters"):
            raise LocatorValidationError(
                f"'{object_type}' cannot have 'value' or 'filters' parameters."
            )

def check_locator_param(locator_value: Optional[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """Check and validate a locator parameter.

    This function verifies that the locator is a dictionary and validates its structure.
    It returns the validated locator dictionary.

    Args:
        locator_value: The locator parameter to check

    Returns:
        The validated locator dictionary

    Raises:
        TypeError: If locator is not a dictionary
        ValueError: If locator format is invalid
    """
    if locator_value is not None:
        # Check if it's a dictionary type
        if not isinstance(locator_value, dict):
            raise TypeError("locator parameter must be a dictionary")
        
        # Use internal validation function to validate locator structure
        try:
            validate_locator(locator_value)
        except LocatorValidationError as e:
            # Prompt user to refer to locator guide
            raise ValueError(f"Invalid locator format: {str(e)}. Please refer to the locator guide for proper syntax.")
    
    return locator_value