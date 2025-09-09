"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class IndexOffsetArraysTest(testing.FunctionalTestCase):
    """
    Arrays Dinámicos integration tests.
    
    Tests: 2 cases across 1 levels
    Category: index_offset_arrays
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "index_offset_arrays.xlsx"

    def test_5r(self):
        """Arrays Dinámicos: Comportamiento con arrays dinámicos"""
        
        # INDEX con array de filas
        value = self.evaluator.evaluate('Tests!R1')
        self.assertIsInstance(value, Array, "=INDEX(Data!A1:E6, ROW(A1:A3), 1) should return Array")

        # OFFSET con array de offsets
        value = self.evaluator.evaluate('Tests!R2')
        self.assertIsInstance(value, Array, "=OFFSET(Data!A1, ROW(A1:A2)-1, 0) should return Array")

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
        array_value = self.evaluator.evaluate('Tests!R1')
        self.assertIsInstance(array_value, Array)
