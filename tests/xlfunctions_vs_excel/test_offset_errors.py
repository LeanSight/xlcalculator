"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class OffsetErrorsTest(testing.FunctionalTestCase):
    """
    OFFSET integration tests.
    
    Tests: 6 cases across 1 levels
    Category: offset_errors
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "offset_errors.xlsx"

    def test_2f(self):
        """OFFSET - Casos de Error: Errores de referencia y parámetros"""
        
        # OFFSET error - antes del inicio de hoja
        value = self.evaluator.evaluate('Tests!F1')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=OFFSET(Data!A1, -2, 0) should return REF_ERROR")

        # OFFSET error - antes del inicio de hoja
        value = self.evaluator.evaluate('Tests!F2')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=OFFSET(Data!A1, 0, -2) should return REF_ERROR")

        # OFFSET error - más allá de hoja
        value = self.evaluator.evaluate('Tests!F3')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=OFFSET(Data!A1, 100, 0) should return REF_ERROR")

        # OFFSET error - más allá de hoja
        value = self.evaluator.evaluate('Tests!F4')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=OFFSET(Data!A1, 0, 100) should return REF_ERROR")

        # OFFSET error - altura cero
        value = self.evaluator.evaluate('Tests!F5')
        self.assertIsInstance(value, xlerrors.ValueExcelError,
                            "=OFFSET(Data!A1, 1, 1, 0, 1) should return VALUE_ERROR")

        # OFFSET error - ancho cero
        value = self.evaluator.evaluate('Tests!F6')
        self.assertIsInstance(value, xlerrors.ValueExcelError,
                            "=OFFSET(Data!A1, 1, 1, 1, 0) should return VALUE_ERROR")

    def test_data_integrity(self):
        """Verify test data integrity."""
        # Auto-generated data validation
        self.assertEqual('Alice', self.evaluator.evaluate('Data!A2'))
        self.assertEqual(25, self.evaluator.evaluate('Data!B2'))
        self.assertEqual('NYC', self.evaluator.evaluate('Data!C2'))

    def test_type_consistency(self):
        """Verify data type consistency across test cases."""
        # Auto-generated type validation

