"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class OffsetIndirectCombinationsTest(testing.FunctionalTestCase):
    """
    OFFSET + INDIRECT integration tests.
    
    Tests: 3 cases across 1 levels
    Category: offset_indirect_combinations
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "offset_indirect_combinations.xlsx"

    def test_3m(self):
        """OFFSET + INDIRECT: Combinaciones OFFSET+INDIRECT"""
        
        # Combinación OFFSET+INDIRECT
        value = self.evaluator.evaluate('Tests!M1')
        self.assertEqual(25, value, "=OFFSET(INDIRECT(\"Data!A1\"), 1, 1) should return 25")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # Combinación OFFSET+INDIRECT desde B2
        value = self.evaluator.evaluate('Tests!M2')
        self.assertEqual('LA', value, "=OFFSET(INDIRECT(\"Data!B2\"), 1, 1) should return 'LA'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # Combinación OFFSET+INDIRECT array
        value = self.evaluator.evaluate('Tests!M3')
        self.assertIsInstance(value, Array, "=OFFSET(INDIRECT(\"Data!A1\"), 1, 1, 2, 2) should return Array")

    def test_data_integrity(self):
        """Verify test data integrity."""
        # Auto-generated data validation
        self.assertEqual('Alice', self.evaluator.evaluate('Data!A2'))
        self.assertEqual(25, self.evaluator.evaluate('Data!B2'))
        self.assertEqual('NYC', self.evaluator.evaluate('Data!C2'))

    def test_type_consistency(self):
        """Verify data type consistency across test cases."""
        # Auto-generated type validation
        # number validation
        number_value = self.evaluator.evaluate('Tests!M1')
        self.assertIsInstance(number_value, (int, float, Number))

        # text validation
        text_value = self.evaluator.evaluate('Tests!M2')
        self.assertIsInstance(text_value, (str, Text))

        # array validation
        array_value = self.evaluator.evaluate('Tests!M3')
        self.assertIsInstance(array_value, Array)
