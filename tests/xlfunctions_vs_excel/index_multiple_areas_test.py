"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class IndexMultipleAreasTest(testing.FunctionalTestCase):
    """
    Forma de Referencia vs Array integration tests.
    
    Tests: 2 cases across 1 levels
    Category: index_multiple_areas
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "index_multiple_areas.xlsx"

    def test_5s(self):
        """Forma de Referencia vs Array: Casos edge de formas de referencia"""
        
        # INDEX forma referencia área 1
        value = self.evaluator.evaluate('Tests!S1')
        self.assertEqual('Alice', value, "=INDEX((Data!A1:A5, Data!C1:C5), 2, 1, 1) should return 'Alice'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # INDEX forma referencia área 2
        value = self.evaluator.evaluate('Tests!S2')
        self.assertEqual('NYC', value, "=INDEX((Data!A1:A5, Data!C1:C5), 2, 1, 2) should return 'NYC'")
        self.assertIsInstance(value, (str, Text), "Should be text")

    def test_data_integrity(self):
        """Verify test data integrity."""
        # Auto-generated data validation
        self.assertEqual('Alice', self.evaluator.evaluate('Data!A2'))
        self.assertEqual(25, self.evaluator.evaluate('Data!B2'))
        self.assertEqual('NYC', self.evaluator.evaluate('Data!C2'))

    def test_type_consistency(self):
        """Verify data type consistency across test cases."""
        # Auto-generated type validation
        # text validation
        text_value = self.evaluator.evaluate('Tests!S1')
        self.assertIsInstance(text_value, (str, Text))
