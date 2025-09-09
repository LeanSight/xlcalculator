"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class OffsetStringReferencesTest(testing.FunctionalTestCase):
    """
    OFFSET Function integration tests.
    
    Tests: 5 cases across 1 levels
    Category: offset_string_references
    Source: REFERENCE_OBJECTS_DESIGN.md
    """
    filename = "offset_string_references.xlsx"

    def test_2a(self):
        """OFFSET Function - Reference Arithmetic: OFFSET function with string references"""
        
        # OFFSET basic with string reference
        value = self.evaluator.evaluate('Tests!C1')
        self.assertEqual('Value_B2', value, "=OFFSET(\"Data!A1\", 1, 1) should return 'Value_B2'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # OFFSET no movement
        value = self.evaluator.evaluate('Tests!C2')
        self.assertEqual('Value_A1', value, "=OFFSET(\"Data!A1\", 0, 0) should return 'Value_A1'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # OFFSET negative movement
        value = self.evaluator.evaluate('Tests!C3')
        self.assertEqual('Value_A1', value, "=OFFSET(\"Data!B2\", -1, -1) should return 'Value_A1'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # OFFSET row movement
        value = self.evaluator.evaluate('Tests!C4')
        self.assertEqual('Value_A3', value, "=OFFSET(\"Data!A1\", 2, 0) should return 'Value_A3'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # OFFSET column movement
        value = self.evaluator.evaluate('Tests!C5')
        self.assertEqual('Value_C1', value, "=OFFSET(\"Data!A1\", 0, 2) should return 'Value_C1'")
        self.assertIsInstance(value, (str, Text), "Should be text")

    def test_data_integrity(self):
        """Verify test data integrity."""
        # Auto-generated data validation
        self.assertEqual('Value_A1', self.evaluator.evaluate('Data!A2'))
        self.assertEqual('Value_B1', self.evaluator.evaluate('Data!B2'))
        self.assertEqual('Value_C1', self.evaluator.evaluate('Data!C2'))

    def test_type_consistency(self):
        """Verify data type consistency across test cases."""
        # Auto-generated type validation
        # text validation
        text_value = self.evaluator.evaluate('Tests!C1')
        self.assertIsInstance(text_value, (str, Text))
