"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class IndirectArraysTest(testing.FunctionalTestCase):
    """
    INDIRECT integration tests.
    
    Tests: 4 cases across 1 levels
    Category: indirect_arrays
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "indirect_arrays.xlsx"

    def test_2i(self):
        """INDIRECT - Arrays de Referencias: Rangos y arrays"""
        
        # INDIRECT array - headers
        value = self.evaluator.evaluate('Tests!I1')
        self.assertIsInstance(value, Array, "=INDIRECT(\"Data!A1:C1\") should return Array")

        # INDIRECT array - columna nombres
        value = self.evaluator.evaluate('Tests!I2')
        self.assertIsInstance(value, Array, "=INDIRECT(\"Data!A2:A6\") should return Array")

        # INDIRECT array - columna edad
        value = self.evaluator.evaluate('Tests!I3')
        self.assertIsInstance(value, Array, "=INDIRECT(\"Data!B1:B6\") should return Array")

        # INDIRECT array - desde celda
        value = self.evaluator.evaluate('Tests!I4')
        self.assertIsInstance(value, Array, "=INDIRECT(P3) should return Array")

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
        array_value = self.evaluator.evaluate('Tests!I1')
        self.assertIsInstance(array_value, Array)
