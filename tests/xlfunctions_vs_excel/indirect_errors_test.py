"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class IndirectErrorsTest(testing.FunctionalTestCase):
    """
    INDIRECT integration tests.
    
    Tests: 5 cases across 1 levels
    Category: indirect_errors
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "indirect_errors.xlsx"

    def test_2k(self):
        """INDIRECT - Casos de Error: Referencias inválidas"""
        
        # INDIRECT error - hoja inexistente
        value = self.evaluator.evaluate('Tests!K1')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=INDIRECT(\"InvalidSheet!A1\") should return REF_ERROR")

        # INDIRECT error - celda vacía
        value = self.evaluator.evaluate('Tests!K2')
        self.assertEqual(0, value, "=INDIRECT(\"Data!Z99\") should return 0")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # INDIRECT error - referencia vacía
        value = self.evaluator.evaluate('Tests!K3')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=INDIRECT(\"\") should return REF_ERROR")

        # INDIRECT error - texto inválido
        value = self.evaluator.evaluate('Tests!K4')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=INDIRECT(\"NotAReference\") should return REF_ERROR")

        # INDIRECT error - hoja inválida desde celda
        value = self.evaluator.evaluate('Tests!K5')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=INDIRECT(P4) should return REF_ERROR")

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
        number_value = self.evaluator.evaluate('Tests!K2')
        self.assertIsInstance(number_value, (int, float, Number))
