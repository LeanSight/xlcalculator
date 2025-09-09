"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from tests import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class ColumnStringReferencesTest(testing.FunctionalTestCase):
    """
    COLUMN Function integration tests.
    
    Tests: 5 cases across 1 levels
    Category: column_string_references
    Source: REFERENCE_OBJECTS_DESIGN.md
    """
    filename = "column_string_references.xlsx"

    def test_1b(self):
        """COLUMN Function - String References: COLUMN function with string reference parameters"""
        
        # COLUMN with string reference A1
        value = self.evaluator.evaluate('Tests!B1')
        self.assertEqual(1, value, "=COLUMN(\"A1\") should return 1")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # COLUMN with column Z
        value = self.evaluator.evaluate('Tests!B2')
        self.assertEqual(26, value, "=COLUMN(\"Z1\") should return 26")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # COLUMN with double letter column
        value = self.evaluator.evaluate('Tests!B3')
        self.assertEqual(27, value, "=COLUMN(\"AA1\") should return 27")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # COLUMN with sheet reference
        value = self.evaluator.evaluate('Tests!B4')
        self.assertEqual(2, value, "=COLUMN(\"Data!B1\") should return 2")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # COLUMN with absolute reference
        value = self.evaluator.evaluate('Tests!B5')
        self.assertEqual(2, value, "=COLUMN(\"$B$1\") should return 2")
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
        number_value = self.evaluator.evaluate('Tests!B1')
        self.assertIsInstance(number_value, (int, float, Number))
