# Test-Driven Development (TDD) Service

The TDD Service is a specialized MCP service designed to automate the testing and fixing workflow for the Word Document Tools project.

## Overview

The TDD Service provides two main tools:

1. **TDD Test Runner** - Automatically runs tests repeatedly until they pass or maximum attempts are reached
2. **TDD Auto Fixer** - Analyzes test failures and suggests or applies fixes

## Tools

### TDD Test Runner

Automatically runs tests in a loop until all tests pass or the maximum number of attempts is reached.

#### Parameters

- `test_command` (string, default: "python -m pytest tests/ -v") - The command to execute tests
- `max_attempts` (integer, default: 10) - Maximum number of test attempts before giving up
- `delay_between_attempts` (integer, default: 5) - Delay in seconds between test attempts
- `stop_on_success` (boolean, default: True) - Stop testing once all tests pass

#### Usage Example

```json
{
  "tool": "tdd_test_runner",
  "arguments": {
    "test_command": "python -m pytest tests/test_selector.py -v",
    "max_attempts": 5,
    "delay_between_attempts": 10
  }
}
```

### TDD Auto Fixer

Analyzes test failures and suggests or applies fixes to resolve issues.

#### Parameters

- `test_results` (object) - Test results from tdd_test_runner to analyze and fix
- `auto_apply_fixes` (boolean, default: False) - Automatically apply suggested fixes

#### Usage Example

```json
{
  "tool": "tdd_auto_fixer",
  "arguments": {
    "test_results": {...},  // Results from tdd_test_runner
    "auto_apply_fixes": true
  }
}
```

## Running the TDD Service

To run the TDD service as a standalone MCP server:

```bash
word_docx_tools_tdd
```

Or using Python module syntax:

```bash
python -m word_docx_tools.tdd_service.main
```

## Workflow

A typical TDD workflow would be:

1. Run `tdd_test_runner` to execute tests
2. If tests fail, pass the results to `tdd_auto_fixer` for analysis
3. Apply suggested fixes manually or automatically
4. Repeat until all tests pass

## Extending the Service

The TDD service can be extended by adding new tools to the [tdd_service](file:///C:/Users/lichao/Office-Word-MCP-Server/word_docx_tools/tdd_service/__init__.py) module:

1. Create a new tool file in the [tdd_service](file:///C:/Users/lichao/Office-Word-MCP-Server/word_docx_tools/tdd_service/__init__.py) directory
2. Import the tool in [tdd_service/main.py](file:///C:/Users/lichao/Office-Word-MCP-Server/word_docx_tools/tdd_service/main.py)
3. The tool will automatically be registered with the TDD server

## Error Handling

The TDD service includes comprehensive error handling:

- Test timeouts (10-minute default)
- Process execution errors
- Invalid parameters
- Resource monitoring

All errors are logged and reported back through the MCP protocol.