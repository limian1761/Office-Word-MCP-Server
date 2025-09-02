"""
Watch and Execute Tool for Word Document MCP Server.

This module provides a tool that watches a target file for changes,
executes specified commands when changes are detected, and 
automatically commits and pushes changes to a remote repository.
"""

import os
import time
import hashlib
import subprocess
from typing import Optional, List
from pathlib import Path

from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

from ..mcp_service.core import mcp_server
from ..mcp_service.core_utils import (
    ErrorCode, 
    WordDocumentError, 
    format_error_response, 
    handle_tool_errors, 
    log_error, 
    log_info
)


def get_file_hash(file_path: str) -> str:
    """Calculate MD5 hash of a file."""
    hash_md5 = hashlib.md5()
    try:
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    except Exception as e:
        raise WordDocumentError(
            ErrorCode.SERVER_ERROR, 
            f"Failed to calculate file hash: {str(e)}"
        )


def execute_commands(commands: List[str]) -> List[dict]:
    """Execute a list of shell commands and return results."""
    results = []
    for command in commands:
        try:
            result = subprocess.run(
                command, 
                shell=True, 
                capture_output=True, 
                text=True,
                timeout=300  # 5 minute timeout for each command
            )
            results.append({
                "command": command,
                "status": "success",
                "returncode": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr
            })
        except subprocess.TimeoutExpired:
            results.append({
                "command": command,
                "status": "timeout",
                "error": "Command timed out after 5 minutes"
            })
        except Exception as e:
            results.append({
                "command": command,
                "status": "error",
                "error": str(e)
            })
    return results


def git_commit_and_push(commit_message: str, branch: str = "main") -> dict:
    """Commit changes and push to remote repository."""
    try:
        # Add all changes
        subprocess.run(["git", "add", "."], check=True, capture_output=True)
        
        # Check if there are changes to commit
        status_result = subprocess.run(["git", "status", "--porcelain"], 
                                     capture_output=True, text=True)
        
        if not status_result.stdout.strip():
            return {
                "status": "success",
                "message": "No changes to commit",
                "stdout": "",
                "stderr": ""
            }
        
        # Commit changes
        commit_result = subprocess.run(["git", "commit", "-m", commit_message], 
                                     capture_output=True, text=True)
        
        if commit_result.returncode != 0 and "nothing to commit" not in commit_result.stderr:
            return {
                "status": "error",
                "message": "Git commit failed",
                "stdout": commit_result.stdout,
                "stderr": commit_result.stderr
            }
        
        # Push to remote
        push_result = subprocess.run(["git", "push", "origin", branch], 
                                   capture_output=True, text=True)
        
        return {
            "status": "success",
            "message": f"Changes committed and pushed to {branch} branch",
            "stdout": push_result.stdout,
            "stderr": push_result.stderr
        }
    except subprocess.CalledProcessError as e:
        return {
            "status": "error",
            "message": f"Git operation failed: {e}",
            "stdout": e.stdout.decode() if e.stdout else "",
            "stderr": e.stderr.decode() if e.stderr else ""
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Unexpected error during git operations: {str(e)}",
            "stdout": "",
            "stderr": str(e)
        }


@mcp_server.tool()
@handle_tool_errors
def watch_and_execute(
    ctx: Context[ServerSession, None] = Field(description="Context object"),
    target_file: str = Field(
        description="Path to the target file to watch for changes"
    ),
    commands: List[str] = Field(
        default=[],
        description="List of shell commands to execute when file changes are detected"
    ),
    refresh_interval: int = Field(
        default=5,
        description="Refresh interval in seconds (default: 5 seconds)",
        ge=1,
        le=300
    ),
    timeout: int = Field(
        default=1800,  # 30 minutes
        description="Timeout in seconds before auto commit and push (default: 1800 seconds / 30 minutes)",
        ge=60,
        le=7200
    ),
    commit_message: str = Field(
        default="Auto commit: Changes detected",
        description="Commit message for git commit"
    ),
    branch: str = Field(
        default="main",
        description="Git branch to push to (default: main)"
    )
) -> dict:
    """
    Watch a target file for changes, execute commands when changes are detected,
    and automatically commit and push changes to a remote repository.
    
    This tool monitors a specified file for updates. When changes are detected, it:
    1. Executes the specified commands
    2. Commits all changes in the repository
    3. Pushes the changes to the remote repository
    
    If no changes are detected within the timeout period, it will:
    1. Execute the specified commands (if any)
    2. Commit any pending changes with a timeout message
    3. Push to the remote repository
    4. Exit the monitoring process
    
    Args:
        target_file: Path to the file to monitor for changes
        commands: List of shell commands to execute when file changes are detected
        refresh_interval: How often to check for file changes (in seconds, 1-300)
        timeout: Maximum time to wait before auto commit and push (in seconds, 60-7200)
        commit_message: Commit message for git operations
        branch: Git branch to push to
        
    Returns:
        Dictionary with monitoring results and executed operations
    """
    
    # Validate target file exists
    if not os.path.exists(target_file):
        raise WordDocumentError(
            ErrorCode.FILE_NOT_FOUND,
            f"Target file not found: {target_file}"
        )
    
    # Get initial file hash
    try:
        initial_hash = get_file_hash(target_file)
    except Exception as e:
        raise WordDocumentError(
            ErrorCode.SERVER_ERROR,
            f"Failed to read target file: {str(e)}"
        )
    
    log_info(f"Started watching file: {target_file}")
    log_info(f"Refresh interval: {refresh_interval}s, Timeout: {timeout}s")
    
    start_time = time.time()
    last_hash = initial_hash
    
    while True:
        # Check if timeout has been reached
        elapsed_time = time.time() - start_time
        if elapsed_time >= timeout:
            log_info("Timeout reached. Executing final operations.")
            
            # Execute commands even if no file change detected
            command_results = []
            if commands:
                command_results = execute_commands(commands)
                log_info(f"Executed {len(commands)} commands on timeout")
            
            # Commit and push changes
            commit_result = git_commit_and_push(
                f"{commit_message} (timeout after {int(elapsed_time)}s)", 
                branch
            )
            
            return {
                "status": "timeout",
                "message": f"Monitoring completed after {int(elapsed_time)} seconds",
                "initial_file_hash": initial_hash,
                "final_file_hash": last_hash,
                "file_changed": initial_hash != last_hash,
                "commands_executed": command_results,
                "git_operation": commit_result,
                "elapsed_time": int(elapsed_time)
            }
        
        # Check for file changes
        try:
            current_hash = get_file_hash(target_file)
        except Exception as e:
            log_error(f"Failed to read target file: {str(e)}")
            time.sleep(refresh_interval)
            continue
        
        # If file has changed
        if current_hash != last_hash:
            log_info(f"File change detected in {target_file}")
            
            # Execute commands
            command_results = []
            if commands:
                command_results = execute_commands(commands)
                log_info(f"Executed {len(commands)} commands after file change")
            
            # Commit and push changes
            commit_result = git_commit_and_push(
                f"{commit_message} (file changed)", 
                branch
            )
            
            return {
                "status": "file_changed",
                "message": "File change detected and operations completed",
                "initial_file_hash": initial_hash,
                "previous_file_hash": last_hash,
                "current_file_hash": current_hash,
                "commands_executed": command_results,
                "git_operation": commit_result,
                "elapsed_time": int(elapsed_time)
            }
        
        # Update last hash and wait for next check
        last_hash = current_hash
        time.sleep(refresh_interval)