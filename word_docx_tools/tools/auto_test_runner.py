"""
Automated Test Runner Tool for Word Document MCP Server.

This module provides a tool that automatically runs tests and continues 
until all tests pass or maximum attempts are reached.
"""

import gc
import json
import logging
import os
import subprocess
import time
from pathlib import Path
from typing import Dict, List, Optional

import psutil
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

from ..mcp_service.core import mcp_server
from ..mcp_service.core_utils import (ErrorCode, WordDocumentError,
                                      format_error_response,
                                      handle_tool_errors, log_error, log_info)

# Configure logging
logger = logging.getLogger("AutoTestRunner")


def send_heartbeat():
    """Send heartbeat signal to keep the tool alive."""
    logger.info("Heartbeat signal sent")


def monitor_resources():
    """Monitor system resources to prevent excessive memory usage."""
    try:
        process = psutil.Process(os.getpid())
        memory_usage = process.memory_info().rss / 1024 / 1024  # MB

        if memory_usage > 500:  # If memory usage exceeds 500MB
            gc.collect()  # Force garbage collection
            logger.info(
                f"Memory usage: {memory_usage} MB, performed garbage collection"
            )

        return memory_usage
    except Exception as e:
        logger.warning(f"Failed to monitor resources: {e}")
        return 0


def run_tests(test_command: str) -> Dict:
    """Run tests and return results."""
    try:
        # Monitor resources before test execution
        monitor_resources()

        # Run tests with JSON output if possible
        result = subprocess.run(
            test_command,
            shell=True,
            capture_output=True,
            text=True,
            timeout=600,  # 10 minute timeout
        )

        return {
            "returncode": result.returncode,
            "stdout": result.stdout,
            "stderr": result.stderr,
            "passed": result.returncode == 0,
        }
    except subprocess.TimeoutExpired:
        return {
            "returncode": -1,
            "stdout": "",
            "stderr": "Test execution timed out",
            "passed": False,
            "timeout": True,
        }
    except Exception as e:
        return {"returncode": -1, "stdout": "", "stderr": str(e), "passed": False}


def parse_test_results(output: str) -> Dict:
    """Parse test output to extract test results."""
    lines = output.strip().split("\n")

    # Look for common pytest patterns
    total_tests = 0
    passed_tests = 0
    failed_tests = 0
    error_tests = 0

    for line in lines:
        if "collected" in line and "items" in line:
            # Pattern: collected 10 items
            try:
                parts = line.split()
                for i, part in enumerate(parts):
                    if part == "collected":
                        total_tests = int(parts[i + 1])
                        break
            except:
                pass
        elif "failed" in line and "passed" in line:
            # Pattern: 5 failed, 15 passed
            try:
                parts = line.replace(",", "").split()
                for i, part in enumerate(parts):
                    if part == "failed":
                        failed_tests = int(parts[i - 1])
                    elif part == "passed":
                        passed_tests = int(parts[i - 1])
            except:
                pass
        elif "ERRORS" in line or "FAILURES" in line:
            # Count errors and failures
            pass

    return {
        "total": total_tests,
        "passed": passed_tests,
        "failed": failed_tests,
        "errors": error_tests,
    }


@mcp_server.tool()
@handle_tool_errors
def auto_test_runner(
    ctx: Context[ServerSession, None] = Field(description="Context object"),
    test_command: str = Field(
        default="python -m pytest tests/ -v",
        description="Command to run tests (default: python -m pytest tests/ -v)",
    ),
    max_attempts: int = Field(
        default=10,
        description="Maximum number of test attempts before giving up (default: 10)",
        ge=1,
        le=100,
    ),
    delay_between_attempts: int = Field(
        default=5,
        description="Delay in seconds between test attempts (default: 5 seconds)",
        ge=0,
        le=300,
    ),
    stop_on_success: bool = Field(
        default=True, description="Stop testing once all tests pass (default: True)"
    ),
    send_heartbeats: bool = Field(
        default=True,
        description="Send heartbeat signals to keep tool alive (default: True)",
    ),
) -> dict:
    """
    Automatically run tests until all tests pass or maximum attempts reached.

    This tool repeatedly executes test commands and monitors the results.
    It can be configured to stop when all tests pass or continue for a
    fixed number of attempts.

    Args:
        test_command: The command to execute tests (e.g., "python -m pytest tests/")
        max_attempts: Maximum number of times to run tests
        delay_between_attempts: Seconds to wait between test runs
        stop_on_success: Whether to stop when all tests pass
        send_heartbeats: Whether to send heartbeat signals to keep tool alive

    Returns:
        Dictionary with test execution results and summary
    """

    log_info(f"Starting auto test runner with command: {test_command}")
    log_info(f"Max attempts: {max_attempts}, Delay: {delay_between_attempts}s")

    attempt = 0
    results_history = []

    while attempt < max_attempts:
        attempt += 1
        log_info(f"Test attempt {attempt}/{max_attempts}")

        # Send heartbeat if enabled
        if send_heartbeats:
            send_heartbeat()

        # Run tests
        test_result = run_tests(test_command)
        results_history.append(
            {"attempt": attempt, "timestamp": time.time(), "result": test_result}
        )

        # Log result
        if test_result["passed"]:
            log_info(f"Test attempt {attempt} PASSED")
            if stop_on_success:
                break
        else:
            log_info(f"Test attempt {attempt} FAILED")
            # Log error details
            if test_result["stderr"]:
                log_error(f"Test error: {test_result['stderr']}")

        # Wait before next attempt (unless it's the last one)
        if attempt < max_attempts:
            log_info(f"Waiting {delay_between_attempts} seconds before next attempt")
            time.sleep(delay_between_attempts)

    # Final summary
    final_passed = results_history[-1]["result"]["passed"] if results_history else False

    summary = {
        "status": "success" if final_passed else "failed",
        "message": f"Auto test runner completed after {attempt} attempts",
        "final_result": results_history[-1]["result"] if results_history else None,
        "attempts_made": attempt,
        "max_attempts": max_attempts,
        "all_results": results_history,
    }

    log_info(f"Auto test runner finished. Final status: {summary['status']}")
    return summary
