"""
Test de aceptación para CellReference dataclass.

Este test define el comportamiento esperado del nuevo CellReference
que debe reemplazar las tuplas (sheet, address) en todo el código.
"""
import unittest
from xlcalculator.utils import CellReference


class CellReferenceTest(unittest.TestCase):
    """Test de aceptación para CellReference dataclass."""

    def test_parse_explicit_sheet_reference(self):
        """Test parsing explicit sheet references like 'Sheet1!A1'."""
        ref = CellReference.parse('Sheet1!A1', current_sheet='Sheet2')
        
        self.assertEqual(ref.sheet, 'Sheet1')
        self.assertEqual(ref.address, 'A1')
        self.assertTrue(ref.is_sheet_explicit)
        self.assertEqual(str(ref), 'Sheet1!A1')

    def test_parse_implicit_sheet_reference(self):
        """Test parsing implicit sheet references like 'A1' using current sheet."""
        ref = CellReference.parse('A1', current_sheet='Sheet3')
        
        self.assertEqual(ref.sheet, 'Sheet3')
        self.assertEqual(ref.address, 'A1')
        self.assertFalse(ref.is_sheet_explicit)
        self.assertEqual(str(ref), 'Sheet3!A1')

    def test_parse_range_reference(self):
        """Test parsing range references like 'Data!B2:C3'."""
        ref = CellReference.parse('Data!B2:C3', current_sheet='Sheet1')
        
        self.assertEqual(ref.sheet, 'Data')
        self.assertEqual(ref.address, 'B2:C3')
        self.assertTrue(ref.is_sheet_explicit)
        self.assertEqual(str(ref), 'Data!B2:C3')

    def test_parse_implicit_range_reference(self):
        """Test parsing implicit range references using current sheet."""
        ref = CellReference.parse('B2:C3', current_sheet='Analysis')
        
        self.assertEqual(ref.sheet, 'Analysis')
        self.assertEqual(ref.address, 'B2:C3')
        self.assertFalse(ref.is_sheet_explicit)
        self.assertEqual(str(ref), 'Analysis!B2:C3')

    def test_is_same_sheet_as_context(self):
        """Test checking if reference is in same sheet as context."""
        ref1 = CellReference.parse('Sheet1!A1', current_sheet='Sheet2')
        ref2 = CellReference.parse('A1', current_sheet='Sheet2')
        
        self.assertFalse(ref1.is_same_sheet_as_context('Sheet2'))
        self.assertTrue(ref2.is_same_sheet_as_context('Sheet2'))

    def test_immutability(self):
        """Test that CellReference is immutable."""
        ref = CellReference.parse('Sheet1!A1', current_sheet='Sheet2')
        
        # Should not be able to modify attributes
        with self.assertRaises(AttributeError):
            ref.sheet = 'NewSheet'
        
        with self.assertRaises(AttributeError):
            ref.address = 'B1'

    def test_equality(self):
        """Test equality comparison between CellReference objects."""
        ref1 = CellReference.parse('Sheet1!A1', current_sheet='Sheet2')
        ref2 = CellReference.parse('Sheet1!A1', current_sheet='Sheet3')
        ref3 = CellReference.parse('A1', current_sheet='Sheet1')
        
        self.assertEqual(ref1, ref2)  # Same resolved reference
        self.assertNotEqual(ref1, ref3)  # Different explicit vs implicit


if __name__ == '__main__':
    unittest.main()