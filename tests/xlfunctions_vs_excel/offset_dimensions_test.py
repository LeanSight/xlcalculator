"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class OffsetDimensionsTest(testing.FunctionalTestCase):
    """
    OFFSET integration tests.
    
    Tests: 5 cases across 1 levels
    Category: offset_dimensions
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "offset_dimensions.xlsx"

    def test_2e(self):
        """OFFSET - Arrays con Dimensiones: Rangos con height/width"""
        
        # OFFSET dimensiones - 1x1
        value = self.evaluator.evaluate('Tests!E1')
        self.assertEqual(25, value, "=OFFSET(Data!A1, 1, 1, 1, 1) should return 25")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # OFFSET dimensiones - 2x2
        value = self.evaluator.evaluate('Tests!E2')
        self.assertIsInstance(value, Array, "=OFFSET(Data!A1, 1, 1, 2, 2) should return Array")

        # OFFSET dimensiones - 3x3
        value = self.evaluator.evaluate('Tests!E3')
        self.assertIsInstance(value, Array, "=OFFSET(Data!A1, 0, 0, 3, 3) should return Array")

        # OFFSET dimensiones - 1x3
        value = self.evaluator.evaluate('Tests!E4')
        self.assertIsInstance(value, Array, "=OFFSET(Data!A1, 2, 1, 1, 3) should return Array")

        # OFFSET dimensiones - 3x1
        value = self.evaluator.evaluate('Tests!E5')
        self.assertIsInstance(value, Array, "=OFFSET(Data!A1, 1, 0, 3, 1) should return Array")

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
        number_value = self.evaluator.evaluate('Tests!E1')
        self.assertIsInstance(number_value, (int, float, Number))

        # array validation
        array_value = self.evaluator.evaluate('Tests!E2')
        self.assertIsInstance(array_value, Array)
