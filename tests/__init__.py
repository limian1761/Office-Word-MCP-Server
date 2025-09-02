"""
Test suite for Word Document Server.

This module organizes tests into logical groups for better maintainability.
"""

# Test organization
TEST_CATEGORIES = {
    'unit': [
        'test_selector',
        'test_text_tools',
    ],
    'integration': [
        'test_com_cache_clearing',
        'test_precise_operations',
    ],
    'e2e': [
        'test_e2e_integration',
        'test_e2e_scenarios',
        'test_real_e2e',
    ],
    'functional': [
        'test_complex_document_operations',
        'test_image_tools',
    ]
}

# Test descriptions
TEST_DESCRIPTIONS = {
    'test_selector': 'Tests for the document selector engine',
    'test_text_tools': 'Tests for text manipulation tools',
    'test_com_cache_clearing': 'Tests for COM object cache management',
    'test_precise_operations': 'Tests for precise document operations',
    'test_e2e_integration': 'End-to-end integration tests',
    'test_e2e_scenarios': 'End-to-end scenario tests',
    'test_real_e2e': 'Real end-to-end tests with actual Word documents',
    'test_complex_document_operations': 'Tests for complex document operations',
    'test_image_tools': 'Tests for image manipulation tools',
}