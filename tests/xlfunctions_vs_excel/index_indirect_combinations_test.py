"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class IndexIndirectCombinationsTest(testing.FunctionalTestCase):
    """
    INDEX + INDIRECT integration tests.
    
    Tests: 4 cases across 1 levels
    Category: index_indirect_combinations
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "index_indirect_combinations.xlsx"

    def test_3l(self):
        """INDEX + INDIRECT: Combinaciones INDEX+INDIRECT"""
        
        # Combinación INDEX+INDIRECT
        value = self.evaluator.evaluate('Tests!L1')
        self.assertEqual(25, value, "=INDEX(INDIRECT(\"Data!A1:E6\"), 2, 2) should return 25")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # Combinación INDEX+INDIRECT array
        value = self.evaluator.evaluate('Tests!L2')
        self.assertIsInstance(value, Array, "=INDEX(INDIRECT(\"Data!A1:E6\"), 0, 2) should return Array")

        # Combinación INDEX+INDIRECT subrange
        value = self.evaluator.evaluate('Tests!L3')
        self.assertEqual('LA', value, "=INDEX(INDIRECT(\"Data!A2:C4\"), 2, 3) should return 'LA'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # Combinación INDEX+INDIRECT columna completa
        value = self.evaluator.evaluate('Tests!L4')
        self.assertEqual('Bob', value, "=INDEX(INDIRECT(\"Data!A:A\"), 3) should return 'Bob'")
        self.assertIsInstance(value, (str, Text), "Should be text")

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
        number_value = self.evaluator.evaluate('Tests!L1')
        self.assertIsInstance(number_value, (int, float, Number))

        # array validation
        array_value = self.evaluator.evaluate('Tests!L2')
        self.assertIsInstance(array_value, Array)

        # text validation
        text_value = self.evaluator.evaluate('Tests!L3')
        self.assertIsInstance(text_value, (str, Text))
