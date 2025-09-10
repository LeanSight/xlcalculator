"""
Unit tests for reference objects.

Tests the CellReference and RangeReference classes
for Excel-compatible behavior.
"""

import unittest
from xlcalculator.references import CellReference, RangeReference
from xlcalculator.xlfunctions import xlerrors


class TestCellReference(unittest.TestCase):
    """Test CellReference class functionality."""
    
    def test_basic_cell_reference(self):
        """Test basic cell reference creation."""
        ref = CellReference(sheet="Sheet1", row=1, column=1)
        self.assertEqual(ref.sheet, "Sheet1")
        self.assertEqual(ref.row, 1)
        self.assertEqual(ref.column, 1)
        self.assertFalse(ref.absolute_row)
        self.assertFalse(ref.absolute_column)
    
    def test_address_property(self):
        """Test address property formatting."""
        # Basic reference
        ref = CellReference(sheet="Sheet1", row=1, column=1)
        self.assertEqual(ref.address, "Sheet1!A1")
        
        # Absolute reference
        ref = CellReference(sheet="Sheet1", row=1, column=1, absolute_row=True, absolute_column=True)
        self.assertEqual(ref.address, "Sheet1!$A$1")
        
        # No sheet
        ref = CellReference(sheet="", row=5, column=26)
        self.assertEqual(ref.address, "Z5")
        
        # Sheet with spaces
        ref = CellReference(sheet="Sheet 2", row=1, column=1)
        self.assertEqual(ref.address, "'Sheet 2'!A1")
    
    def test_coordinate_property(self):
        """Test coordinate property."""
        ref = CellReference(sheet="Sheet1", row=5, column=3)
        self.assertEqual(ref.coordinate, (5, 3))
    
    def test_column_to_letter_conversion(self):
        """Test column number to letter conversion."""
        self.assertEqual(CellReference._column_to_letter(1), "A")
        self.assertEqual(CellReference._column_to_letter(26), "Z")
        self.assertEqual(CellReference._column_to_letter(27), "AA")
        self.assertEqual(CellReference._column_to_letter(52), "AZ")
        self.assertEqual(CellReference._column_to_letter(702), "ZZ")
        self.assertEqual(CellReference._column_to_letter(703), "AAA")
    
    def test_letter_to_column_conversion(self):
        """Test letter to column number conversion."""
        self.assertEqual(CellReference._letter_to_column("A"), 1)
        self.assertEqual(CellReference._letter_to_column("Z"), 26)
        self.assertEqual(CellReference._letter_to_column("AA"), 27)
        self.assertEqual(CellReference._letter_to_column("AZ"), 52)
        self.assertEqual(CellReference._letter_to_column("ZZ"), 702)
        self.assertEqual(CellReference._letter_to_column("AAA"), 703)
    
    def test_parse_basic_references(self):
        """Test parsing basic cell references."""
        # Simple reference
        ref = CellReference.parse("A1")
        self.assertEqual(ref.sheet, "")
        self.assertEqual(ref.row, 1)
        self.assertEqual(ref.column, 1)
        self.assertFalse(ref.absolute_row)
        self.assertFalse(ref.absolute_column)
        
        # Column Z
        ref = CellReference.parse("Z1")
        self.assertEqual(ref.column, 26)
        
        # Double letter column
        ref = CellReference.parse("AA1")
        self.assertEqual(ref.column, 27)
        
        # High row number
        ref = CellReference.parse("A100")
        self.assertEqual(ref.row, 100)
    
    def test_parse_absolute_references(self):
        """Test parsing absolute references."""
        # Fully absolute
        ref = CellReference.parse("$A$1")
        self.assertTrue(ref.absolute_row)
        self.assertTrue(ref.absolute_column)
        
        # Column absolute only
        ref = CellReference.parse("$A1")
        self.assertFalse(ref.absolute_row)
        self.assertTrue(ref.absolute_column)
        
        # Row absolute only
        ref = CellReference.parse("A$1")
        self.assertTrue(ref.absolute_row)
        self.assertFalse(ref.absolute_column)
    
    def test_parse_sheet_references(self):
        """Test parsing references with sheet names."""
        # Basic sheet reference
        ref = CellReference.parse("Sheet1!A1")
        self.assertEqual(ref.sheet, "Sheet1")
        self.assertEqual(ref.row, 1)
        self.assertEqual(ref.column, 1)
        
        # Sheet with spaces
        ref = CellReference.parse("'Sheet 2'!A1")
        self.assertEqual(ref.sheet, "Sheet 2")
        
        # Sheet with absolute reference
        ref = CellReference.parse("Data!$B$5")
        self.assertEqual(ref.sheet, "Data")
        self.assertEqual(ref.row, 5)
        self.assertEqual(ref.column, 2)
        self.assertTrue(ref.absolute_row)
        self.assertTrue(ref.absolute_column)
    
    def test_parse_case_insensitive(self):
        """Test that parsing is case insensitive."""
        ref1 = CellReference.parse("a1")
        ref2 = CellReference.parse("A1")
        self.assertEqual(ref1.row, ref2.row)
        self.assertEqual(ref1.column, ref2.column)
    
    def test_parse_invalid_references(self):
        """Test parsing invalid references raises errors."""
        with self.assertRaises(xlerrors.RefExcelError):
            CellReference.parse("")
        
        with self.assertRaises(xlerrors.RefExcelError):
            CellReference.parse("InvalidRef")
        
        with self.assertRaises(xlerrors.RefExcelError):
            CellReference.parse("A")
        
        with self.assertRaises(xlerrors.RefExcelError):
            CellReference.parse("1A")
        
        with self.assertRaises(xlerrors.RefExcelError):
            CellReference.parse(None)
    
    def test_offset_operations(self):
        """Test reference offset operations."""
        ref = CellReference.parse("B2")
        
        # Basic offset
        offset_ref = ref.offset(1, 1)
        self.assertEqual(offset_ref.address, "C3")
        
        # No offset
        offset_ref = ref.offset(0, 0)
        self.assertEqual(offset_ref.address, "B2")
        
        # Negative offset
        offset_ref = ref.offset(-1, -1)
        self.assertEqual(offset_ref.address, "A1")
        
        # Large offset
        offset_ref = ref.offset(98, 24)
        self.assertEqual(offset_ref.address, "Z100")
    
    def test_offset_bounds_checking(self):
        """Test offset bounds checking."""
        ref = CellReference.parse("A1")
        
        # Out of bounds - negative row
        with self.assertRaises(xlerrors.RefExcelError):
            ref.offset(-1, 0)
        
        # Out of bounds - negative column
        with self.assertRaises(xlerrors.RefExcelError):
            ref.offset(0, -1)
        
        # Out of bounds - row too high
        with self.assertRaises(xlerrors.RefExcelError):
            ref.offset(1048576, 0)
        
        # Out of bounds - column too high
        with self.assertRaises(xlerrors.RefExcelError):
            ref.offset(0, 16384)
    
    def test_excel_bounds_validation(self):
        """Test Excel bounds validation during creation."""
        # Valid bounds
        ref = CellReference(sheet="", row=1, column=1)
        self.assertEqual(ref.row, 1)
        
        ref = CellReference(sheet="", row=1048576, column=16384)
        self.assertEqual(ref.row, 1048576)
        
        # Invalid bounds
        with self.assertRaises(xlerrors.RefExcelError):
            CellReference(sheet="", row=0, column=1)
        
        with self.assertRaises(xlerrors.RefExcelError):
            CellReference(sheet="", row=1048577, column=1)
        
        with self.assertRaises(xlerrors.RefExcelError):
            CellReference(sheet="", row=1, column=0)
        
        with self.assertRaises(xlerrors.RefExcelError):
            CellReference(sheet="", row=1, column=16385)


class TestRangeReference(unittest.TestCase):
    """Test RangeReference class functionality."""
    
    def test_basic_range_reference(self):
        """Test basic range reference creation."""
        start = CellReference(sheet="Sheet1", row=1, column=1)
        end = CellReference(sheet="Sheet1", row=2, column=2)
        range_ref = RangeReference(start_cell=start, end_cell=end)
        
        self.assertEqual(range_ref.start_cell, start)
        self.assertEqual(range_ref.end_cell, end)
    
    def test_range_address_property(self):
        """Test range address formatting."""
        start = CellReference(sheet="Sheet1", row=1, column=1)
        end = CellReference(sheet="Sheet1", row=2, column=2)
        range_ref = RangeReference(start_cell=start, end_cell=end)
        
        self.assertEqual(range_ref.address, "Sheet1!A1:B2")
    
    def test_range_dimensions(self):
        """Test range dimensions calculation."""
        start = CellReference(sheet="Sheet1", row=1, column=1)
        end = CellReference(sheet="Sheet1", row=3, column=2)
        range_ref = RangeReference(start_cell=start, end_cell=end)
        
        self.assertEqual(range_ref.dimensions, (3, 2))
    
    def test_range_parse_basic(self):
        """Test parsing basic range references."""
        range_ref = RangeReference.parse("A1:B2")
        self.assertEqual(range_ref.start_cell.address, "A1")
        self.assertEqual(range_ref.end_cell.address, "B2")
        
        # Single cell as range
        range_ref = RangeReference.parse("A1")
        self.assertEqual(range_ref.start_cell.address, "A1")
        self.assertEqual(range_ref.end_cell.address, "A1")
    
    def test_range_parse_with_sheet(self):
        """Test parsing range references with sheet names."""
        range_ref = RangeReference.parse("Sheet1!A1:B2")
        self.assertEqual(range_ref.start_cell.sheet, "Sheet1")
        self.assertEqual(range_ref.end_cell.sheet, "Sheet1")
        self.assertEqual(range_ref.address, "Sheet1!A1:B2")


if __name__ == '__main__':
    unittest.main()