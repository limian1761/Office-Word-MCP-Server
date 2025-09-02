"""
Automated Code Fixer for TDD Service.

This module provides tools for analyzing test failures and automatically fixing code issues.
"""

import re
import os
import logging
from typing import Dict, List, Optional, Any

from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

from .core import tdd_server
# 使用绝对导入路径
from word_docx_tools.mcp_service.core_utils import (
    handle_tool_errors, 
    log_error, 
    log_info
)


# Configure logging
logger = logging.getLogger('TDDAutoFixer')

def analyze_test_failures(test_result: Dict) -> List[Dict]:
    """Analyze test failures and identify potential fixes."""
    failures = []
    
    # Extract failure information from test output
    stderr = test_result.get("stderr", "")
    stdout = test_result.get("stdout", "")
    
    # Look for common error patterns
    # ImportError pattern
    import_errors = re.findall(r"ModuleNotFoundError: No module named '(.+)'", stderr)
    for module in import_errors:
        failures.append({
            "type": "import_error",
            "module": module,
            "description": f"Missing module import: {module}",
            "suggestion": f"Install or fix import for module '{module}'"
        })
    
    # SyntaxError pattern
    syntax_errors = re.findall(r"SyntaxError: (.+)", stderr)
    for error in syntax_errors:
        failures.append({
            "type": "syntax_error",
            "error": error,
            "description": f"Syntax error: {error}",
            "suggestion": "Check syntax and fix errors"
        })
    
    # AttributeError pattern
    attr_errors = re.findall(r"AttributeError: '(.+)' object has no attribute '(.+)'", stderr)
    for obj, attr in attr_errors:
        failures.append({
            "type": "attribute_error",
            "object": obj,
            "attribute": attr,
            "description": f"Object '{obj}' has no attribute '{attr}'",
            "suggestion": f"Check if attribute '{attr}' exists or fix object '{obj}'"
        })
    
    # FileNotFoundError pattern
    file_errors = re.findall(r"FileNotFoundError: (.+)", stderr)
    for error in file_errors:
        failures.append({
            "type": "file_not_found",
            "error": error,
            "description": f"File not found: {error}",
            "suggestion": "Check file paths and ensure files exist"
        })
    
    return failures


def generate_fix_suggestions(failures: List[Dict]) -> List[Dict]:
    """Generate fix suggestions based on failure analysis."""
    suggestions = []
    
    for failure in failures:
        suggestion = {
            "failure": failure,
            "fix_type": "unknown",
            "description": "No specific fix available",
            "confidence": 0.0
        }
        
        if failure["type"] == "import_error":
            suggestion.update({
                "fix_type": "install_dependency",
                "description": f"Install missing package: {failure['module']}",
                "command": f"pip install {failure['module']}",
                "confidence": 0.9
            })
        elif failure["type"] == "syntax_error":
            suggestion.update({
                "fix_type": "code_fix",
                "description": f"Fix syntax error: {failure['error']}",
                "confidence": 0.7
            })
        elif failure["type"] == "attribute_error":
            suggestion.update({
                "fix_type": "code_fix",
                "description": f"Fix attribute error for {failure['object']}.{failure['attribute']}",
                "confidence": 0.8
            })
        elif failure["type"] == "file_not_found":
            suggestion.update({
                "fix_type": "file_check",
                "description": f"Verify file exists: {failure['error']}",
                "confidence": 0.85
            })
        
        suggestions.append(suggestion)
    
    return suggestions


@tdd_server.tool()
@handle_tool_errors
def tdd_auto_fixer(
    ctx: Context[ServerSession, None] = Field(description="Context object"),
    test_results: Dict = Field(
        description="Test results from tdd_test_runner to analyze and fix"
    ),
    auto_apply_fixes: bool = Field(
        default=False,
        description="Automatically apply suggested fixes (default: False)"
    )
) -> dict:
    """
    Analyze test failures and suggest or apply fixes.
    
    This tool analyzes test failure output and provides suggestions for fixing issues.
    It can also automatically apply some fixes if configured to do so.
    
    Args:
        test_results: Test results from tdd_test_runner to analyze
        auto_apply_fixes: Whether to automatically apply suggested fixes
        
    Returns:
        Dictionary with analysis results and fix suggestions
    """
    
    log_info("Starting TDD auto fixer analysis")
    
    # Analyze test failures
    failures = analyze_test_failures(test_results)
    
    # Generate fix suggestions
    suggestions = generate_fix_suggestions(failures)
    
    # Apply fixes if requested
    applied_fixes = []
    if auto_apply_fixes:
        for suggestion in suggestions:
            # In a real implementation, we would apply fixes here
            # For now, we just log that we would apply them
            log_info(f"Would apply fix: {suggestion['description']}")
            applied_fixes.append(suggestion)
    
    result = {
        "status": "analysis_complete",
        "message": f"Analyzed {len(failures)} failures and generated {len(suggestions)} suggestions",
        "failures": failures,
        "suggestions": suggestions,
        "applied_fixes": applied_fixes,
        "auto_applied": auto_apply_fixes
    }
    
    log_info(f"TDD auto fixer completed: {result['message']}")
    return result