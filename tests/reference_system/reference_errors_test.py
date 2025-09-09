"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class ReferenceErrorsTest(testing.FunctionalTestCase):
    """
    Reference Parsing Errors integration tests.
    
    Tests: 5 cases across 1 levels
    Category: reference_errors
    Source: REFERENCE_OBJECTS_DESIGN.md
    """
    filename = "reference_errors.xlsx"

    def test_3a(self):
        """Reference Parsing Errors: Error handling for invalid references"""
        
        # ROW with invalid reference
        value = self.evaluator.evaluate('Tests!D1')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=ROW(\"InvalidRef\") should return REF_ERROR")

        # ROW with empty string
        value = self.evaluator.evaluate('Tests!D2')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=ROW(\"\") should return REF_ERROR")

        # COLUMN with incomplete reference
        value = self.evaluator.evaluate('Tests!D3')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=COLUMN(\"A\") should return REF_ERROR")

        # OFFSET out of bounds
        value = self.evaluator.evaluate('Tests!D4')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=OFFSET(\"Data!A1\", -1, 0) should return REF_ERROR")

        # OFFSET with zero height
        value = self.evaluator.evaluate('Tests!D5')
        self.assertIsInstance(value, xlerrors.ValueExcelError,
                            "=OFFSET(\"Data!A1\", 0, 0, 0, 1) should return VALUE_ERROR")

    def test_data_integrity(self):
        """Verify test data integrity."""
        # Auto-generated data validation
        self.assertEqual('Value_A1', self.evaluator.evaluate('Data!A2'))
        self.assertEqual('Value_B1', self.evaluator.evaluate('Data!B2'))
        self.assertEqual('Value_C1', self.evaluator.evaluate('Data!C2'))

    def test_type_consistency(self):
        """Verify data type consistency across test cases."""
        # Auto-generated type validation

