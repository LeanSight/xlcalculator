"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class IndirectColumnsTest(testing.FunctionalTestCase):
    """
    INDIRECT integration tests.
    
    Tests: 4 cases across 1 levels
    Category: indirect_columns
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "indirect_columns.xlsx"

    def test_2j(self):
        """INDIRECT - Referencias de Columna/Fila Completa: Referencias completas"""
        
        # INDIRECT columna completa A
        value = self.evaluator.evaluate('Tests!J1')
        self.assertIsInstance(value, Array, "=INDIRECT(\"Data!A:A\") should return Array")

        # INDIRECT columna completa B
        value = self.evaluator.evaluate('Tests!J2')
        self.assertIsInstance(value, Array, "=INDIRECT(\"Data!B:B\") should return Array")

        # INDIRECT fila completa 1
        value = self.evaluator.evaluate('Tests!J3')
        self.assertIsInstance(value, Array, "=INDIRECT(\"Data!1:1\") should return Array")

        # INDIRECT fila completa 2
        value = self.evaluator.evaluate('Tests!J4')
        self.assertIsInstance(value, Array, "=INDIRECT(\"Data!2:2\") should return Array")

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
        array_value = self.evaluator.evaluate('Tests!J1')
        self.assertIsInstance(array_value, Array)
