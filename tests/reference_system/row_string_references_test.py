"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from tests import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class RowStringReferencesTest(testing.FunctionalTestCase):
    """
    ROW Function integration tests.
    
    Tests: 5 cases across 1 levels
    Category: row_string_references
    Source: REFERENCE_OBJECTS_DESIGN.md
    """
    filename = "row_string_references.xlsx"

    def test_1a(self):
        """ROW Function - String References: ROW function with string reference parameters"""
        
        # ROW with string reference A1
        value = self.evaluator.evaluate('Tests!A1')
        self.assertEqual(1, value, "=ROW(\"A1\") should return 1")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # ROW with string reference A100
        value = self.evaluator.evaluate('Tests!A2')
        self.assertEqual(100, value, "=ROW(\"A100\") should return 100")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # ROW with sheet reference
        value = self.evaluator.evaluate('Tests!A3')
        self.assertEqual(5, value, "=ROW(\"Data!A5\") should return 5")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # ROW with column Z
        value = self.evaluator.evaluate('Tests!A4')
        self.assertEqual(1, value, "=ROW(\"Z1\") should return 1")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # ROW with absolute reference
        value = self.evaluator.evaluate('Tests!A5')
        self.assertEqual(1, value, "=ROW(\"$A$1\") should return 1")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

    def test_data_integrity(self):
        """Verify test data integrity."""
        # Auto-generated data validation
        self.assertEqual('Value_A1', self.evaluator.evaluate('Data!A2'))
        self.assertEqual('Value_B1', self.evaluator.evaluate('Data!B2'))
        self.assertEqual('Value_C1', self.evaluator.evaluate('Data!C2'))

    def test_type_consistency(self):
        """Verify data type consistency across test cases."""
        # Auto-generated type validation
        # number validation
        number_value = self.evaluator.evaluate('Tests!A1')
        self.assertIsInstance(number_value, (int, float, Number))
