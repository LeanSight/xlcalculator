"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class IndirectConstructedReferencesTest(testing.FunctionalTestCase):
    """
    INDIRECT integration tests.
    
    Tests: 4 cases across 1 levels
    Category: indirect_constructed_references
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "indirect_constructed_references.xlsx"

    def test_2h(self):
        """INDIRECT - Referencias Dinámicas: Referencias construidas"""
        
        # INDIRECT dinámico - concatenación
        value = self.evaluator.evaluate('Tests!H1')
        self.assertEqual('Alice', value, "=INDIRECT(\"Data!A\" & 2) should return 'Alice'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # INDIRECT dinámico - CHAR
        value = self.evaluator.evaluate('Tests!H2')
        self.assertEqual(30, value, "=INDIRECT(\"Data!\" & CHAR(66) & \"3\") should return 30")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # INDIRECT dinámico - ROW (ROW() from H3 returns 3, Data!A3='Bob')
        value = self.evaluator.evaluate('Tests!H3')
        self.assertEqual('Bob', value, "=INDIRECT(\"Data!A\" & ROW()) should return 'Bob'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # INDIRECT dinámico - COLUMN (COLUMN() from H4 returns 8, CHAR(73)='I', Data!I1=0)
        value = self.evaluator.evaluate('Tests!H4')
        self.assertEqual(0, value, "=INDIRECT(\"Data!\" & CHAR(65+COLUMN()) & \"1\") should return 0")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

    def test_data_integrity(self):
        """Verify test data integrity."""
        # Auto-generated data validation
        self.assertEqual('Alice', self.evaluator.evaluate('Data!A2'))
        self.assertEqual(25, self.evaluator.evaluate('Data!B2'))
        self.assertEqual('NYC', self.evaluator.evaluate('Data!C2'))

    def test_type_consistency(self):
        """Verify data type consistency across test cases."""
        # Auto-generated type validation
        # text validation
        text_value = self.evaluator.evaluate('Tests!H1')
        self.assertIsInstance(text_value, (str, Text))

        # number validation
        number_value = self.evaluator.evaluate('Tests!H2')
        self.assertIsInstance(number_value, (int, float, Number))
