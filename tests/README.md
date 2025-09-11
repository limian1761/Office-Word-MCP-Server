# Word Document Server Test Suite

This directory contains the test suite for the Word Document Server project. Tests are organized into logical categories to facilitate maintenance and execution.

## Test Organization

Tests are grouped into the following categories:

### Unit Tests
- `test_selector.py` - Tests for the document selector engine
- `test_text_tools.py` - Tests for text manipulation tools

### Integration Tests
- `test_com_cache_clearing.py` - Tests for COM object cache management
- `test_precise_operations.py` - Tests for precise document operations

### End-to-End Tests
- `test_e2e_integration.py` - End-to-end integration tests
- `test_e2e_scenarios.py` - End-to-end scenario tests
- `test_real_e2e.py` - Real end-to-end tests with actual Word documents

### Functional Tests
- `test_complex_document_operations.py` - Tests for complex document operations
- `test_image_tools.py` - Tests for image manipulation tools

## Running Tests

### Prerequisites
- Windows operating system
- Microsoft Word installed
- Python 3.11+
- Required dependencies installed

### Using the Test Runner Script
The project includes a custom test runner script that provides a more user-friendly interface for running tests:

```bash
# Run all tests
python tests/run_tests.py

# Run with verbose output
python tests/run_tests.py -v

# Run tests matching a specific pattern
python tests/run_tests.py -t text
```

### Using pytest Directly
You can also run tests using pytest directly:

```bash
# Run all tests
python -m pytest tests/ -v

# Run specific test category
python -m pytest tests/test_selector.py tests/test_text_tools.py -v

# Run end-to-end tests
python -m pytest tests/test_e2e*.py tests/test_real_e2e.py -v
```

### Run Individual Test Files
```bash
python -m pytest tests/test_selector.py -v
```

### Code Coverage
The project includes coverage configuration to measure test coverage:

```bash
# Run tests with coverage
coverage run -m pytest tests/

# Generate coverage report
coverage report

# Generate HTML coverage report
coverage html
```

## Test Environment

Tests require:
1. A working Microsoft Word installation
2. Proper COM permissions
3. Write access to the test directory for temporary files

Some tests create temporary Word documents during execution. These files are automatically cleaned up after the tests complete.

## Adding New Tests

When adding new tests:
1. Place them in the appropriate category file
2. Follow the existing naming conventions
3. Include both positive and negative test cases
4. Ensure tests clean up any created resources
5. Add descriptive test names and docstrings

## Continuous Integration

Tests are automatically run as part of the CI pipeline. All tests must pass before changes can be merged.

## Test Document Resources

The `test_docs` directory contains sample Word documents used in testing:
- `valid_test_document_v2.docx` - Standard test document with various elements
- `additional_test_document.docx` - Additional test document with complex formatting