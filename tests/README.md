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

### Run All Tests
```bash
python -m pytest tests/ -v
```

### Run Specific Test Category
```bash
# Run unit tests
python -m pytest tests/test_selector.py tests/test_text_tools.py -v

# Run integration tests
python -m pytest tests/test_com_cache_clearing.py tests/test_precise_operations.py -v

# Run end-to-end tests
python -m pytest tests/test_e2e*.py tests/test_real_e2e.py -v
```

### Run Individual Test Files
```bash
python -m pytest tests/test_selector.py -v
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