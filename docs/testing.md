# Testing Guidelines

## Test Organization

The tests are organized into the following categories:

1. **Unit Tests** - Test individual functions and methods in isolation
2. **Integration Tests** - Test the interaction between different components
3. **End-to-End Tests** - Test complete workflows simulating real user scenarios

## Unit Tests

Unit tests are located in files with the pattern `test_*.py` and focus on testing individual functions. These tests use mocks to isolate the function under test from its dependencies.

## Integration Tests

Integration tests verify that different components of the system work together correctly. These tests may use a combination of real and mocked components.

## End-to-End Tests

End-to-end (E2E) tests simulate real-world usage scenarios by testing complete workflows. These tests:

1. Use real Word COM objects when possible
2. Perform actual operations on real document files
3. Test complete sequences of operations that a user might perform
4. Verify that the system behaves correctly across multiple tool invocations

### E2E Test Files

- [test_e2e_integration.py](file:///C:/Users/lichao/Office-Word-MCP-Server/tests/test_e2e_integration.py) - General integration tests covering various tool workflows
- [test_e2e_scenarios.py](file:///C:/Users/lichao/Office-Word-MCP-Server/tests/test_e2e_scenarios.py) - Specific scenario-based tests mimicking real usage patterns
- [test_real_e2e.py](file:///C:/Users/lichao/Office-Word-MCP-Server/tests/test_real_e2e.py) - Tests using actual COM operations with real documents

### Running E2E Tests

To run the end-to-end tests:

```bash
python -m pytest tests/test_e2e_integration.py -v
python -m pytest tests/test_e2e_scenarios.py -v
python -m pytest tests/test_real_e2e.py -v
```

Or to run all tests including E2E tests:

```bash
python -m pytest tests/ -v
```

Note: E2E tests require a Windows environment with Microsoft Word installed, as they interact with the actual Word COM API.

## Test Data

Test documents are located in the [tests/test_docs](file:///C:/Users/lichao/Office-Word-MCP-Server/tests/test_docs) directory. These documents are used by various tests to verify functionality.

## Writing New Tests

When adding new tests, consider the following:

1. **Unit tests** for new functions and methods
2. **Integration tests** when adding new components that interact with existing ones
3. **E2E tests** when implementing new workflows or user-facing features

Follow the existing patterns in the test files, and ensure that tests are properly isolated and do not depend on each other.

## Test Environment

Tests should be able to run in isolation and clean up after themselves. Use temporary directories for file operations and ensure that any COM objects are properly closed at the end of tests.