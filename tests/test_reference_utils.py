"""
Tests for Reference Resolution Utilities

Comprehensive test suite for the reference resolution utilities used by
dynamic range functions.
"""

import unittest
from xlcalculator.xlfunctions.reference_utils import ReferenceResolver
from xlcalculator.xlfunctions import xlerrors


class TestReferenceResolver(unittest.TestCase):
    """Test suite for ReferenceResolver utility class"""
    
    def test_parse_cell_reference_basic(self):
        """Test basic cell reference parsing"""
        test_cases = [
            ("A1", (1, 1)),
            ("B2", (2, 2)),
            ("C3", (3, 3)),
            ("Z26", (26, 26)),
            ("AA1", (1, 27)),
            ("AB2", (2, 28)),
            ("BA1", (1, 53)),
            ("AAA1", (1, 703)),
        ]
        
        for ref, expected in test_cases:
            with self.subTest(ref=ref):
                result = ReferenceResolver.parse_cell_reference(ref)
                self.assertEqual(result, expected)
    
    def test_parse_cell_reference_with_sheet(self):
        """Test cell reference parsing with sheet names"""
        test_cases = [
            ("Sheet1!A1", (1, 1)),
            ("Data!B2", (2, 2)),
            ("'Sheet Name'!C3", (3, 3)),
        ]
        
        for ref, expected in test_cases:
            with self.subTest(ref=ref):
                result = ReferenceResolver.parse_cell_reference(ref)
                self.assertEqual(result, expected)
    
    def test_parse_cell_reference_absolute(self):
        """Test cell reference parsing with absolute references"""
        test_cases = [
            ("$A$1", (1, 1)),
            ("$B2", (2, 2)),
            ("A$3", (3, 1)),
            ("Sheet1!$A$1", (1, 1)),
        ]
        
        for ref, expected in test_cases:
            with self.subTest(ref=ref):
                result = ReferenceResolver.parse_cell_reference(ref)
                self.assertEqual(result, expected)
    
    def test_parse_cell_reference_errors(self):
        """Test cell reference parsing error cases"""
        invalid_refs = [
            "",           # Empty string
            "1A",         # Numbers before letters
            "A",          # Missing row number
            "1",          # Missing column letters
            "A0",         # Row 0 (invalid)
            "A-1",        # Negative row
            "!A1",        # Invalid sheet reference
            123,          # Non-string input
        ]
        
        for ref in invalid_refs:
            with self.subTest(ref=ref):
                with self.assertRaises((xlerrors.ValueExcelError, xlerrors.RefExcelError)):
                    ReferenceResolver.parse_cell_reference(ref)
    
    def test_cell_to_string_basic(self):
        """Test basic coordinate to string conversion"""
        test_cases = [
            ((1, 1), "A1"),
            ((2, 2), "B2"),
            ((3, 3), "C3"),
            ((26, 26), "Z26"),
            ((1, 27), "AA1"),
            ((2, 28), "AB2"),
            ((1, 53), "BA1"),
            ((1, 703), "AAA1"),
        ]
        
        for coords, expected in test_cases:
            with self.subTest(coords=coords):
                result = ReferenceResolver.cell_to_string(*coords)
                self.assertEqual(result, expected)
    
    def test_cell_to_string_errors(self):
        """Test coordinate to string conversion error cases"""
        invalid_coords = [
            (0, 1),      # Row 0
            (1, 0),      # Column 0
            (-1, 1),     # Negative row
            (1, -1),     # Negative column
            (1048577, 1), # Row too large
            (1, 16385),  # Column too large
        ]
        
        for row, col in invalid_coords:
            with self.subTest(coords=(row, col)):
                with self.assertRaises(xlerrors.RefExcelError):
                    ReferenceResolver.cell_to_string(row, col)
    
    def test_roundtrip_conversion(self):
        """Test that parse -> convert -> parse gives same result"""
        test_refs = ["A1", "B2", "Z26", "AA1", "AB2", "BA1", "AAA1"]
        
        for ref in test_refs:
            with self.subTest(ref=ref):
                # Parse to coordinates
                coords = ReferenceResolver.parse_cell_reference(ref)
                # Convert back to string
                result = ReferenceResolver.cell_to_string(*coords)
                # Should match original
                self.assertEqual(result, ref)
    
    def test_parse_range_reference_basic(self):
        """Test basic range reference parsing"""
        test_cases = [
            ("A1", ((1, 1), (1, 1))),           # Single cell
            ("A1:B2", ((1, 1), (2, 2))),        # Simple range
            ("B2:D4", ((2, 2), (4, 4))),        # Larger range
            ("A1:A10", ((1, 1), (10, 1))),      # Column range
            ("A1:Z1", ((1, 1), (1, 26))),       # Row range
        ]
        
        for ref, expected in test_cases:
            with self.subTest(ref=ref):
                result = ReferenceResolver.parse_range_reference(ref)
                self.assertEqual(result, expected)
    
    def test_parse_range_reference_reversed(self):
        """Test range parsing with reversed coordinates"""
        # Excel allows B2:A1 (should normalize to A1:B2)
        test_cases = [
            ("B2:A1", ((1, 1), (2, 2))),
            ("D4:B2", ((2, 2), (4, 4))),
        ]
        
        for ref, expected in test_cases:
            with self.subTest(ref=ref):
                result = ReferenceResolver.parse_range_reference(ref)
                self.assertEqual(result, expected)
    
    def test_parse_range_reference_with_sheet(self):
        """Test range parsing with sheet names"""
        test_cases = [
            ("Sheet1!A1:B2", ((1, 1), (2, 2))),
            ("Data!C3:D4", ((3, 3), (4, 4))),
        ]
        
        for ref, expected in test_cases:
            with self.subTest(ref=ref):
                result = ReferenceResolver.parse_range_reference(ref)
                self.assertEqual(result, expected)
    
    def test_range_to_string(self):
        """Test coordinate pairs to range string conversion"""
        test_cases = [
            (((1, 1), (1, 1)), "A1"),           # Single cell
            (((1, 1), (2, 2)), "A1:B2"),        # Simple range
            (((2, 2), (4, 4)), "B2:D4"),        # Larger range
            (((1, 1), (10, 1)), "A1:A10"),      # Column range
            (((1, 1), (1, 26)), "A1:Z1"),       # Row range
        ]
        
        for coords, expected in test_cases:
            with self.subTest(coords=coords):
                result = ReferenceResolver.range_to_string(*coords)
                self.assertEqual(result, expected)
    
    def test_offset_reference_basic(self):
        """Test basic reference offsetting"""
        test_cases = [
            ("A1", 1, 1, None, None, "B2"),     # Basic offset
            ("A1", 0, 1, None, None, "B1"),     # Column only
            ("A1", 1, 0, None, None, "A2"),     # Row only
            ("A1", -1, -1, None, None, None),   # Negative (should error)
            ("B2", 1, 1, None, None, "C3"),     # From different start
        ]
        
        for ref, rows, cols, height, width, expected in test_cases:
            with self.subTest(ref=ref, rows=rows, cols=cols):
                if expected is None:
                    # Expect error
                    with self.assertRaises(xlerrors.RefExcelError):
                        ReferenceResolver.offset_reference(ref, rows, cols, height, width)
                else:
                    result = ReferenceResolver.offset_reference(ref, rows, cols, height, width)
                    self.assertEqual(result, expected)
    
    def test_offset_reference_with_dimensions(self):
        """Test reference offsetting with explicit dimensions"""
        test_cases = [
            ("A1", 1, 1, 1, 1, "B2"),           # Single cell result
            ("A1", 1, 1, 2, 2, "B2:C3"),        # Range result
            ("A1", 0, 0, 3, 3, "A1:C3"),        # No offset, just resize
            ("B2", 1, 1, 1, 3, "C3:E3"),        # Row result
            ("B2", 1, 1, 3, 1, "C3:C5"),        # Column result
        ]
        
        for ref, rows, cols, height, width, expected in test_cases:
            with self.subTest(ref=ref, rows=rows, cols=cols, height=height, width=width):
                result = ReferenceResolver.offset_reference(ref, rows, cols, height, width)
                self.assertEqual(result, expected)
    
    def test_offset_reference_range_input(self):
        """Test offsetting range references"""
        test_cases = [
            ("A1:B2", 1, 1, None, None, "B2:C3"),   # Preserve dimensions
            ("A1:C3", 0, 1, None, None, "B1:D3"),   # Column shift
            ("A1:B2", 1, 1, 1, 1, "B2"),            # Resize to single cell
            ("A1:B2", 0, 0, 3, 3, "A1:C3"),         # Resize larger
        ]
        
        for ref, rows, cols, height, width, expected in test_cases:
            with self.subTest(ref=ref, rows=rows, cols=cols, height=height, width=width):
                result = ReferenceResolver.offset_reference(ref, rows, cols, height, width)
                self.assertEqual(result, expected)
    
    def test_validate_bounds(self):
        """Test bounds validation"""
        # Valid coordinates
        valid_coords = [
            (1, 1),
            (1000, 1000),
            (1048576, 16384),  # Excel limits
        ]
        
        for row, col in valid_coords:
            with self.subTest(coords=(row, col)):
                # Should not raise
                ReferenceResolver.validate_bounds(row, col)
        
        # Invalid coordinates
        invalid_coords = [
            (0, 1),
            (1, 0),
            (-1, 1),
            (1, -1),
            (1048577, 1),
            (1, 16385),
        ]
        
        for row, col in invalid_coords:
            with self.subTest(coords=(row, col)):
                with self.assertRaises(xlerrors.RefExcelError):
                    ReferenceResolver.validate_bounds(row, col)
    
    def test_get_range_dimensions(self):
        """Test range dimension calculation"""
        test_cases = [
            ("A1", (1, 1)),
            ("A1:A1", (1, 1)),
            ("A1:B2", (2, 2)),
            ("A1:C3", (3, 3)),
            ("A1:A10", (10, 1)),
            ("A1:Z1", (1, 26)),
        ]
        
        for ref, expected in test_cases:
            with self.subTest(ref=ref):
                result = ReferenceResolver.get_range_dimensions(ref)
                self.assertEqual(result, expected)
    
    def test_normalize_reference(self):
        """Test reference normalization"""
        test_cases = [
            ("$A$1", "A1"),
            ("$a$1", "A1"),
            ("Sheet1!$A$1", "Sheet1!A1"),
            ("sheet1!$a$1:$b$2", "sheet1!A1:B2"),
            ("A1:B2", "A1:B2"),  # Already normalized
        ]
        
        for ref, expected in test_cases:
            with self.subTest(ref=ref):
                result = ReferenceResolver.normalize_reference(ref)
                self.assertEqual(result, expected)


if __name__ == '__main__':
    unittest.main(verbosity=2)