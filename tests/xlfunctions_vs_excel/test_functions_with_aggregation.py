"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class FunctionsWithAggregationTest(testing.FunctionalTestCase):
    """
    Funciones con Agregación integration tests.
    
    Tests: 4 cases across 1 levels
    Category: functions_with_aggregation
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "functions_with_aggregation.xlsx"

    def test_4o(self):
        """Funciones con Agregación: Uso con funciones de agregado"""
        
        # SUM con INDEX array - correct calculation
        value = self.evaluator.evaluate('Tests!O1')
        self.assertEqual(140, value, "=SUM(INDEX(Data!A1:E6, 0, 2)) should return 140")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # AVERAGE con OFFSET array
        value = self.evaluator.evaluate('Tests!O2')
        self.assertEqual(28, value, "=AVERAGE(OFFSET(Data!B1, 1, 0, 5, 1)) should return 28")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # COUNT con INDIRECT columna
        value = self.evaluator.evaluate('Tests!O3')
        self.assertEqual(5, value, "=COUNT(INDIRECT(\"Data!B:B\")) should return 5")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # MAX con INDEX array
        value = self.evaluator.evaluate('Tests!O4')
        self.assertEqual(95, value, "=MAX(INDEX(Data!A1:E6, 0, 4)) should return 95")
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
        number_value = self.evaluator.evaluate('Tests!O1')
        self.assertIsInstance(number_value, (int, float, Number))
