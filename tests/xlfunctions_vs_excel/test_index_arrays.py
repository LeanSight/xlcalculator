"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class IndexArraysTest(testing.FunctionalTestCase):
    """
    INDEX integration tests.
    
    Tests: 4 cases across 1 levels
    Category: index_arrays
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "index_arrays.xlsx"

    def test_1b(self):
        """INDEX - Arrays Completos: Filas y columnas completas"""
        
        # INDEX array - columna completa Age
        value = self.evaluator.evaluate('Tests!B1')
        self.assertIsInstance(value, Array, "=INDEX(Data!A1:E6, 0, 2) should return Array")

        # INDEX array - fila completa Alice
        value = self.evaluator.evaluate('Tests!B2')
        self.assertIsInstance(value, Array, "=INDEX(Data!A1:E6, 2, 0) should return Array")

        # INDEX array - primera columna Name
        value = self.evaluator.evaluate('Tests!B3')
        self.assertIsInstance(value, Array, "=INDEX(Data!A1:E6, 0, 1) should return Array")

        # INDEX array - columna Active
        value = self.evaluator.evaluate('Tests!B4')
        self.assertIsInstance(value, Array, "=INDEX(Data!A1:E6, 0, 5) should return Array")

    def test_data_integrity(self):
        """Verify test data integrity."""
        # Auto-generated data validation
        self.assertEqual('Alice', self.evaluator.evaluate('Data!A2'))
        self.assertEqual(25, self.evaluator.evaluate('Data!B2'))
        self.assertEqual('NYC', self.evaluator.evaluate('Data!C2'))

    def test_type_consistency(self):
        """Verify data type consistency across test cases."""
        # Auto-generated type validation
        # array validation
        array_value = self.evaluator.evaluate('Tests!B1')
        self.assertIsInstance(array_value, Array)
