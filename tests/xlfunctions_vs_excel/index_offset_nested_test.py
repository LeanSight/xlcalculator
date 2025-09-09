"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class IndexOffsetNestedTest(testing.FunctionalTestCase):
    """
    Combinaciones Complejas integration tests.
    
    Tests: 4 cases across 1 levels
    Category: index_offset_nested
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "index_offset_nested.xlsx"

    def test_3n(self):
        """Combinaciones Complejas: Funciones anidadas complejas"""
        
        # Combinación INDEX+OFFSET
        value = self.evaluator.evaluate('Tests!N1')
        self.assertEqual(25, value, "=INDEX(OFFSET(Data!A1, 0, 0, 3, 3), 2, 2) should return 25")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # Combinación OFFSET+INDEX
        value = self.evaluator.evaluate('Tests!N2')
        self.assertEqual(30, value, "=OFFSET(INDEX(Data!A1:E6, 2, 1), 1, 1) should return 30")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # Referencia dinámica - concatenación simple
        value = self.evaluator.evaluate('Tests!N3')
        self.assertEqual('Alice', value, "=INDIRECT(\"Data!A\" & 2) should return 'Alice'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # Referencia dinámica - CHAR function
        value = self.evaluator.evaluate('Tests!N4')
        self.assertEqual(25, value, "=INDIRECT(\"Data!\" & CHAR(66) & \"2\") should return 25")
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
        # number validation
        number_value = self.evaluator.evaluate('Tests!N1')
        self.assertIsInstance(number_value, (int, float, Number))

        # text validation
        text_value = self.evaluator.evaluate('Tests!N3')
        self.assertIsInstance(text_value, (str, Text))
