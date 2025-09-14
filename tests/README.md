# Tests

This directory contains tests for the Word Document MCP Server.

## Test Structure

```
tests/
├── conftest.py              # pytest configuration and fixtures
├── test_main.py             # Tests for main module
├── test_mcp_service.py      # Tests for MCP service components
├── test_app_context.py      # Tests for AppContext class
├── test_document_ops.py     # Tests for document operations
├── test_document_tools.py   # Tests for document tools
├── test_text_operations.py  # Tests for text operations
└── ...
```

## Running Tests

To run all tests:

```bash
python -m pytest
```

To run tests with coverage:

```bash
python -m pytest --cov=word_docx_tools --cov-report=html
```

To run specific test files:

```bash
python -m pytest tests/test_main.py
```

To run tests in verbose mode:

```bash
python -m pytest -v
```

## Test Categories

Tests are organized into the following categories:

1. **Unit Tests** - Test individual functions and classes in isolation
2. **Integration Tests** - Test interactions between components
3. **Functional Tests** - Test complete features and user workflows

## Writing Tests

When writing new tests:

1. Use descriptive test function names that follow the pattern `test_what_is_being_tested_what_is_the_expected_result`
2. Use pytest fixtures for setup and teardown
3. Mock external dependencies (especially COM objects)
4. Focus on testing one behavior per test
5. Use assertions to verify expected outcomes

## Fixtures

Common fixtures are defined in [conftest.py](file:///C:/Users/lichao/Office-Word-MCP-Server/tests/conftest.py):

- `mock_word_app` - Mock Word application COM object
- `mock_document` - Mock Word document COM object
- `mock_app_context` - Mock AppContext instance

## Test Dependencies

Test dependencies are defined in [pyproject.toml](file:///C:/Users/lichao/Office-Word-MCP-Server/pyproject.toml):

- `pytest` - Testing framework
- `pytest-asyncio` - Async support for pytest
- `pytest-cov` - Coverage reporting (if added)

Install test dependencies with:

```bash
pip install -e .[test]
```

Or for development:

```bash
pip install -e .[dev]
```