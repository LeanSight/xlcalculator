"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class IndexErrorsTest(testing.FunctionalTestCase):
    """
    INDEX integration tests.
    
    Tests: 5 cases across 1 levels
    Category: index_errors
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "index_errors.xlsx"

    def test_1c(self):
        """INDEX - Casos de Error Estructurales: Manejo de errores"""
        
        # INDEX error - fila fuera de rango
        value = self.evaluator.evaluate('Tests!C1')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=INDEX(Data!A1:E6, 7, 1) should return REF_ERROR")

        # INDEX error - columna fuera de rango
        value = self.evaluator.evaluate('Tests!C2')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=INDEX(Data!A1:E6, 1, 7) should return REF_ERROR")

        # INDEX error - ambos cero
        value = self.evaluator.evaluate('Tests!C3')
        self.assertIsInstance(value, xlerrors.ValueExcelError,
                            "=INDEX(Data!A1:E6, 0, 0) should return VALUE_ERROR")

        # INDEX error - fila negativa
        value = self.evaluator.evaluate('Tests!C4')
        self.assertIsInstance(value, xlerrors.ValueExcelError,
                            "=INDEX(Data!A1:E6, -1, 1) should return VALUE_ERROR")

        # INDEX error - columna negativa
        value = self.evaluator.evaluate('Tests!C5')
        self.assertIsInstance(value, xlerrors.ValueExcelError,
                            "=INDEX(Data!A1:E6, 1, -1) should return VALUE_ERROR")

    def test_data_integrity(self):
        """Verify test data integrity."""
        # Auto-generated data validation
        self.assertEqual('Alice', self.evaluator.evaluate('Data!A2'))
        self.assertEqual(25, self.evaluator.evaluate('Data!B2'))
        self.assertEqual('NYC', self.evaluator.evaluate('Data!C2'))

    def test_type_consistency(self):
        """Verify data type consistency across test cases."""
        # Auto-generated type validation

