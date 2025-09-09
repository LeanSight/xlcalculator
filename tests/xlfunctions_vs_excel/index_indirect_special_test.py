"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class IndexIndirectSpecialTest(testing.FunctionalTestCase):
    """
    Referencias Especiales integration tests.
    
    Tests: 3 cases across 1 levels
    Category: index_indirect_special
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "index_indirect_special.xlsx"

    def test_5q(self):
        """Referencias Especiales: Referencias especiales y complejas"""
        
        # INDIRECT misma hoja
        value = self.evaluator.evaluate('Tests!Q1')
        self.assertEqual('Test Value', value, "=INDIRECT(\"Tests!Z1\") should return 'Test Value'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # INDEX con columna completa
        value = self.evaluator.evaluate('Tests!Q2')
        self.assertEqual('Alice', value, "=INDEX(Data!A:A, 2) should return 'Alice'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # OFFSET con columna completa
        value = self.evaluator.evaluate('Tests!Q3')
        self.assertIsInstance(value, Array, "=OFFSET(Data!A:A, 1, 0, 3, 1) should return Array")

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
        text_value = self.evaluator.evaluate('Tests!Q1')
        self.assertIsInstance(text_value, (str, Text))

        # array validation
        array_value = self.evaluator.evaluate('Tests!Q3')
        self.assertIsInstance(array_value, Array)
