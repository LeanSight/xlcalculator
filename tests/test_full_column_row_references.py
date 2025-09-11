"""
Comprehensive tests for full column and row references (A:A, 1:1).

Tests the new FullColumnReference and FullRowReference classes and their
integration with INDEX, OFFSET, INDIRECT functions.
"""

import unittest
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean
from xlcalculator.references import CellReference, FullColumnReference, FullRowReference
from tests import testing


class FullColumnReferenceTest(unittest.TestCase):
    """Test FullColumnReference class functionality."""
    
    def test_parse_simple_column_reference(self):
        """Test parsing simple column references like A:A."""
        ref = FullColumnReference.parse("A:A")
        self.assertEqual(ref.column, 1)
        self.assertEqual(ref.sheet, "")
        self.assertFalse(ref.absolute_column)
        self.assertFalse(ref.is_sheet_explicit)
    
    def test_parse_absolute_column_reference(self):
        """Test parsing absolute column references like $A:$A."""
        ref = FullColumnReference.parse("$A:$A")
        self.assertEqual(ref.column, 1)
        self.assertTrue(ref.absolute_column)
    
    def test_parse_sheet_column_reference(self):
        """Test parsing sheet column references like Sheet1!A:A."""
        ref = FullColumnReference.parse("Sheet1!A:A")
        self.assertEqual(ref.column, 1)
        self.assertEqual(ref.sheet, "Sheet1")
        self.assertTrue(ref.is_sheet_explicit)
    
    def test_parse_quoted_sheet_column_reference(self):
        """Test parsing quoted sheet column references like 'Sheet Name'!A:A."""
        ref = FullColumnReference.parse("'Sheet Name'!A:A")
        self.assertEqual(ref.column, 1)
        self.assertEqual(ref.sheet, "Sheet Name")
        self.assertTrue(ref.is_sheet_explicit)
    
    def test_parse_multi_letter_column(self):
        """Test parsing multi-letter column references like AA:AA."""
        ref = FullColumnReference.parse("AA:AA")
        self.assertEqual(ref.column, 27)  # AA is column 27
    
    def test_parse_invalid_multi_column_range(self):
        """Test that multi-column ranges like A:B raise error."""
        with self.assertRaises(xlerrors.RefExcelError):
            FullColumnReference.parse("A:B")
    
    def test_parse_inconsistent_absolute_markers(self):
        """Test that inconsistent absolute markers like $A:A raise error."""
        with self.assertRaises(xlerrors.RefExcelError):
            FullColumnReference.parse("$A:A")
    
    def test_address_property(self):
        """Test address property formatting."""
        ref = FullColumnReference.parse("A:A")
        self.assertEqual(ref.address, "A:A")
        
        ref_abs = FullColumnReference.parse("$A:$A")
        self.assertEqual(ref_abs.address, "$A:$A")
    
    def test_full_address_property(self):
        """Test full_address property formatting."""
        ref = FullColumnReference.parse("Sheet1!A:A")
        self.assertEqual(ref.full_address, "Sheet1!A:A")
        
        ref_quoted = FullColumnReference.parse("'Sheet Name'!A:A")
        self.assertEqual(ref_quoted.full_address, "'Sheet Name'!A:A")
    
    def test_get_cell_at_row(self):
        """Test getting cell reference at specific row."""
        ref = FullColumnReference.parse("A:A")
        cell = ref.get_cell_at_row(5)
        self.assertEqual(cell.row, 5)
        self.assertEqual(cell.column, 1)
    
    def test_to_range_reference(self):
        """Test converting to range reference."""
        ref = FullColumnReference.parse("A:A")
        range_ref = ref.to_range_reference(1, 10)
        self.assertEqual(range_ref.start_cell.row, 1)
        self.assertEqual(range_ref.end_cell.row, 10)
        self.assertEqual(range_ref.start_cell.column, 1)
        self.assertEqual(range_ref.end_cell.column, 1)


class FullRowReferenceTest(unittest.TestCase):
    """Test FullRowReference class functionality."""
    
    def test_parse_simple_row_reference(self):
        """Test parsing simple row references like 1:1."""
        ref = FullRowReference.parse("1:1")
        self.assertEqual(ref.row, 1)
        self.assertEqual(ref.sheet, "")
        self.assertFalse(ref.absolute_row)
        self.assertFalse(ref.is_sheet_explicit)
    
    def test_parse_absolute_row_reference(self):
        """Test parsing absolute row references like $1:$1."""
        ref = FullRowReference.parse("$1:$1")
        self.assertEqual(ref.row, 1)
        self.assertTrue(ref.absolute_row)
    
    def test_parse_sheet_row_reference(self):
        """Test parsing sheet row references like Sheet1!1:1."""
        ref = FullRowReference.parse("Sheet1!1:1")
        self.assertEqual(ref.row, 1)
        self.assertEqual(ref.sheet, "Sheet1")
        self.assertTrue(ref.is_sheet_explicit)
    
    def test_parse_invalid_multi_row_range(self):
        """Test that multi-row ranges like 1:2 raise error."""
        with self.assertRaises(xlerrors.RefExcelError):
            FullRowReference.parse("1:2")
    
    def test_address_property(self):
        """Test address property formatting."""
        ref = FullRowReference.parse("1:1")
        self.assertEqual(ref.address, "1:1")
        
        ref_abs = FullRowReference.parse("$1:$1")
        self.assertEqual(ref_abs.address, "$1:$1")
    
    def test_get_cell_at_column(self):
        """Test getting cell reference at specific column."""
        ref = FullRowReference.parse("1:1")
        cell = ref.get_cell_at_column(5)
        self.assertEqual(cell.row, 1)
        self.assertEqual(cell.column, 5)
    
    def test_to_range_reference(self):
        """Test converting to range reference."""
        ref = FullRowReference.parse("1:1")
        range_ref = ref.to_range_reference(1, 10)
        self.assertEqual(range_ref.start_cell.row, 1)
        self.assertEqual(range_ref.end_cell.row, 1)
        self.assertEqual(range_ref.start_cell.column, 1)
        self.assertEqual(range_ref.end_cell.column, 10)


class FullReferenceIntegrationTest(unittest.TestCase):
    """Integration tests for full column/row references with functions."""
    
    def setUp(self):
        """Set up test data."""
        from xlcalculator.model import Model
        from xlcalculator.evaluator import Evaluator
        
        # Create a simple model without Excel file dependency
        self.model = Model()
        self.evaluator = Evaluator(self.model)
        
        # Create test data in columns A and B, rows 1-5
        test_data = {
            'Data!A1': 'Header1',
            'Data!A2': 'Alice',
            'Data!A3': 'Bob',
            'Data!A4': 'Charlie',
            'Data!A5': 'David',
            'Data!B1': 'Header2',
            'Data!B2': 10,
            'Data!B3': 20,
            'Data!B4': 30,
            'Data!B5': 40,
            'Data!C1': 'Header3',
            'Data!C2': 100,
            'Data!C3': 200,
            'Data!C4': 300,
            'Data!C5': 400,
        }
        
        # Set up the model with test data
        for addr, value in test_data.items():
            self.evaluator.set_cell_value(addr, value)
    
    def test_index_with_full_column_reference(self):
        """Test INDEX function with full column references."""
        # Test INDEX function directly with full column reference
        from xlcalculator.xlfunctions.dynamic_range import INDEX
        from xlcalculator.ast_nodes import EvalContext
        
        # Create evaluation context
        context = EvalContext(ref='Tests!E1')
        context.evaluator = self.evaluator
        context.sheet = 'Tests'
        
        # Test INDEX(Data!A:A, 2) should return 'Alice'
        result = INDEX("Data!A:A", 2, _context=context)
        self.assertEqual('Alice', result)
        
        # Test INDEX(Data!B:B, 3) should return 20
        result = INDEX("Data!B:B", 3, _context=context)
        self.assertEqual(20, result)
    
    def test_index_with_full_row_reference(self):
        """Test INDEX function with full row references."""
        # Test INDEX function directly with full row reference
        from xlcalculator.xlfunctions.dynamic_range import INDEX
        from xlcalculator.ast_nodes import EvalContext
        
        # Create evaluation context
        context = EvalContext(ref='Tests!E1')
        context.evaluator = self.evaluator
        context.sheet = 'Tests'
        
        # Test INDEX(Data!1:1, 1, 2) should return 'Header2'
        result = INDEX("Data!1:1", 1, 2, _context=context)
        self.assertEqual('Header2', result)
        
        # Test INDEX(Data!2:2, 1, 3) should return 100
        result = INDEX("Data!2:2", 1, 3, _context=context)
        self.assertEqual(100, result)
    
    def test_indirect_with_full_column_reference(self):
        """Test INDIRECT function with full column references."""
        # Test INDIRECT function directly with full column reference
        from xlcalculator.xlfunctions.dynamic_range import INDIRECT
        from xlcalculator.ast_nodes import EvalContext
        
        # Create evaluation context
        context = EvalContext(ref='Tests!E1')
        context.evaluator = self.evaluator
        context.sheet = 'Tests'
        
        # Test INDIRECT("Data!A:A") should return Array
        result = INDIRECT("Data!A:A", _context=context)
        self.assertIsInstance(result, Array)
        
        # Test INDIRECT("Data!B:B") should return Array
        result = INDIRECT("Data!B:B", _context=context)
        self.assertIsInstance(result, Array)
    
    def test_indirect_with_full_row_reference(self):
        """Test INDIRECT function with full row references."""
        # Test INDIRECT function directly with full row reference
        from xlcalculator.xlfunctions.dynamic_range import INDIRECT
        from xlcalculator.ast_nodes import EvalContext
        
        # Create evaluation context
        context = EvalContext(ref='Tests!E1')
        context.evaluator = self.evaluator
        context.sheet = 'Tests'
        
        # Test INDIRECT("Data!1:1") should return Array
        result = INDIRECT("Data!1:1", _context=context)
        self.assertIsInstance(result, Array)
        
        # Test INDIRECT("Data!2:2") should return Array
        result = INDIRECT("Data!2:2", _context=context)
        self.assertIsInstance(result, Array)
    
    def test_combined_index_indirect_full_references(self):
        """Test combined INDEX + INDIRECT with full references."""
        # Test combined INDEX + INDIRECT functions directly
        from xlcalculator.xlfunctions.dynamic_range import INDEX, INDIRECT
        from xlcalculator.ast_nodes import EvalContext
        
        # Create evaluation context
        context = EvalContext(ref='Tests!E1')
        context.evaluator = self.evaluator
        context.sheet = 'Tests'
        
        # Test INDEX(INDIRECT("Data!A:A"), 3) should return 'Bob'
        indirect_result = INDIRECT("Data!A:A", _context=context)
        result = INDEX(indirect_result, 3, _context=context)
        self.assertEqual('Bob', result)
        
        # Test INDEX(INDIRECT("Data!B:B"), 4) should return 30
        indirect_result = INDIRECT("Data!B:B", _context=context)
        result = INDEX(indirect_result, 4, _context=context)
        self.assertEqual(30, result)


class FullReferenceParserTest(unittest.TestCase):
    """Test parser integration with full column/row references."""
    
    def test_parser_recognizes_full_column_reference(self):
        """Test that parser recognizes A:A as full column reference."""
        from xlcalculator.parser import FormulaParser
        
        parser = FormulaParser()
        # Test that _is_full_column_or_row_reference works correctly
        self.assertTrue(parser._is_full_column_or_row_reference("A:A"))
        self.assertTrue(parser._is_full_column_or_row_reference("Sheet!A:A"))
        self.assertTrue(parser._is_full_column_or_row_reference("1:1"))
        self.assertTrue(parser._is_full_column_or_row_reference("Sheet!1:1"))
        
        # Test that regular ranges are not recognized as full references
        self.assertFalse(parser._is_full_column_or_row_reference("A1:B2"))
        self.assertFalse(parser._is_full_column_or_row_reference("A1"))


class FullReferenceBoundsTest(unittest.TestCase):
    """Test bounds validation for full column/row references."""
    
    def test_column_bounds_validation(self):
        """Test column bounds validation."""
        from xlcalculator.constants import EXCEL_MAX_COLUMNS
        
        # Valid column
        ref = FullColumnReference.parse("A:A")
        self.assertEqual(ref.column, 1)
        
        # Test bounds in __post_init__
        with self.assertRaises(xlerrors.RefExcelError):
            FullColumnReference(sheet="", column=0)  # Below minimum
        
        with self.assertRaises(xlerrors.RefExcelError):
            FullColumnReference(sheet="", column=EXCEL_MAX_COLUMNS + 1)  # Above maximum
    
    def test_row_bounds_validation(self):
        """Test row bounds validation."""
        from xlcalculator.constants import EXCEL_MAX_ROWS
        
        # Valid row
        ref = FullRowReference.parse("1:1")
        self.assertEqual(ref.row, 1)
        
        # Test bounds in __post_init__
        with self.assertRaises(xlerrors.RefExcelError):
            FullRowReference(sheet="", row=0)  # Below minimum
        
        with self.assertRaises(xlerrors.RefExcelError):
            FullRowReference(sheet="", row=EXCEL_MAX_ROWS + 1)  # Above maximum


if __name__ == '__main__':
    unittest.main()