"""
Acceptance Tests for Hierarchical Model - Workbook Operations

Tests the core Workbook functionality including worksheet management,
cross-sheet references, and Excel compatibility.
"""
import pytest
from xlcalculator.hierarchical_model import Workbook, Worksheet, Cell
from xlcalculator.model import ModelCompiler


class TestWorkbookBasicOperations:
    """Test basic workbook operations and worksheet management."""
    
    def test_create_empty_workbook(self):
        """Should create an empty workbook with no worksheets."""
        workbook = Workbook(name="TestWorkbook")
        
        assert workbook.name == "TestWorkbook"
        assert len(workbook.worksheets) == 0
        assert workbook.active_sheet is None
    
    def test_add_worksheet(self):
        """Should add a new worksheet to the workbook."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = workbook.add_worksheet("Sheet1")
        
        assert worksheet.name == "Sheet1"
        assert worksheet.workbook is workbook
        assert "Sheet1" in workbook.worksheets
        assert workbook.worksheets["Sheet1"] is worksheet
        assert workbook.active_sheet == "Sheet1"  # First sheet becomes active
    
    def test_get_worksheet(self):
        """Should retrieve existing worksheet by name."""
        workbook = Workbook(name="TestWorkbook")
        sheet1 = workbook.add_worksheet("Sheet1")
        sheet2 = workbook.add_worksheet("Sheet2")
        
        assert workbook.get_worksheet("Sheet1") is sheet1
        assert workbook.get_worksheet("Sheet2") is sheet2
    
    def test_get_nonexistent_worksheet_raises_error(self):
        """Should raise KeyError for non-existent worksheet."""
        workbook = Workbook(name="TestWorkbook")
        
        with pytest.raises(KeyError, match="Worksheet 'NonExistent' not found"):
            workbook.get_worksheet("NonExistent")
    
    def test_remove_worksheet(self):
        """Should remove worksheet from workbook."""
        workbook = Workbook(name="TestWorkbook")
        workbook.add_worksheet("Sheet1")
        workbook.add_worksheet("Sheet2")
        
        workbook.remove_worksheet("Sheet1")
        
        assert "Sheet1" not in workbook.worksheets
        assert "Sheet2" in workbook.worksheets
        assert workbook.active_sheet == "Sheet2"  # Active sheet updates


class TestWorkbookCellOperations:
    """Test workbook-level cell operations with cross-sheet access."""
    
    def test_set_cell_value_with_full_address(self):
        """Should set cell value using full sheet!cell address."""
        workbook = Workbook(name="TestWorkbook")
        workbook.add_worksheet("Sheet1")
        
        workbook.set_cell_value("Sheet1!A1", 42)
        
        cell = workbook.get_cell("Sheet1!A1")
        assert cell.value == 42
        assert cell.address == "A1"
        assert cell.worksheet.name == "Sheet1"
    
    def test_get_cell_with_full_address(self):
        """Should retrieve cell using full sheet!cell address."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = workbook.add_worksheet("Sheet1")
        worksheet.set_cell_value("A1", "Hello")
        
        cell = workbook.get_cell("Sheet1!A1")
        
        assert cell.value == "Hello"
        assert cell.address == "A1"
        assert cell.worksheet is worksheet
    
    def test_cross_sheet_references(self):
        """Should handle references between different worksheets."""
        workbook = Workbook(name="TestWorkbook")
        sheet1 = workbook.add_worksheet("Sheet1")
        sheet2 = workbook.add_worksheet("Sheet2")
        
        # Set value in Sheet1
        sheet1.set_cell_value("A1", 10)
        
        # Reference from Sheet2
        sheet2.set_cell_value("B1", "=Sheet1!A1")
        
        # Verify cross-sheet reference
        cell_b1 = sheet2.get_cell("B1")
        assert cell_b1.formula is not None
        assert "Sheet1!A1" in cell_b1.formula.terms
    
    def test_get_cell_creates_empty_cell_if_not_exists(self):
        """Should create empty cell if it doesn't exist."""
        workbook = Workbook(name="TestWorkbook")
        workbook.add_worksheet("Sheet1")
        
        cell = workbook.get_cell("Sheet1!Z99")
        
        assert cell.value is None
        assert cell.address == "Z99"
        assert cell.worksheet.name == "Sheet1"


class TestWorkbookExcelFileIntegration:
    """Test integration with Excel file reading through ModelCompiler."""
    
    def test_model_compiler_creates_hierarchical_model(self):
        """Should create hierarchical model from Excel file."""
        # This test will use a simple dictionary input to simulate Excel file
        compiler = ModelCompiler()
        
        input_dict = {
            "Sheet1!A1": 10,
            "Sheet1!B1": "=A1*2",
            "Sheet2!A1": 20,
            "Sheet2!B1": "=Sheet1!A1+A1"
        }
        
        model = compiler.read_and_parse_dict_hierarchical(input_dict)
        
        # Verify workbook structure
        assert isinstance(model, Workbook)
        assert len(model.worksheets) == 2
        assert "Sheet1" in model.worksheets
        assert "Sheet2" in model.worksheets
        
        # Verify Sheet1 cells
        sheet1 = model.get_worksheet("Sheet1")
        assert sheet1.get_cell("A1").value == 10
        assert sheet1.get_cell("B1").formula.formula == "=A1*2"
        
        # Verify Sheet2 cells
        sheet2 = model.get_worksheet("Sheet2")
        assert sheet2.get_cell("A1").value == 20
        assert sheet2.get_cell("B1").formula.formula == "=Sheet1!A1+A1"
    
    def test_backward_compatibility_with_flat_model(self):
        """Should maintain backward compatibility with existing Model class."""
        compiler = ModelCompiler()
        
        input_dict = {
            "A1": 10,
            "B1": "=A1*2"
        }
        
        # Test both old and new methods work
        flat_model = compiler.read_and_parse_dict(input_dict)
        hierarchical_model = compiler.read_and_parse_dict_hierarchical(input_dict)
        
        # Verify both models have same data
        assert flat_model.get_cell_value("Sheet1!A1") == 10
        assert hierarchical_model.get_cell("Sheet1!A1").value == 10
        
        # Verify conversion between models
        converted_flat = hierarchical_model.to_flat_model()
        assert converted_flat.get_cell_value("Sheet1!A1") == 10


class TestWorkbookDefinedNames:
    """Test workbook-level defined names and named ranges."""
    
    def test_add_defined_name_for_cell(self):
        """Should add defined name pointing to a cell."""
        workbook = Workbook(name="TestWorkbook")
        sheet1 = workbook.add_worksheet("Sheet1")
        sheet1.set_cell_value("A1", 42)
        
        workbook.add_defined_name("MyCell", "Sheet1!A1")
        
        assert "MyCell" in workbook.defined_names
        cell = workbook.get_cell_by_name("MyCell")
        assert cell.value == 42
        assert cell.address == "A1"
    
    def test_add_defined_name_for_range(self):
        """Should add defined name pointing to a range."""
        workbook = Workbook(name="TestWorkbook")
        sheet1 = workbook.add_worksheet("Sheet1")
        
        workbook.add_defined_name("MyRange", "Sheet1!A1:B2")
        
        assert "MyRange" in workbook.defined_names
        range_obj = workbook.get_range_by_name("MyRange")
        assert range_obj.address == "A1:B2"
        assert range_obj.worksheet is sheet1
    
    def test_use_defined_name_in_formula(self):
        """Should use defined names in formulas."""
        workbook = Workbook(name="TestWorkbook")
        sheet1 = workbook.add_worksheet("Sheet1")
        
        sheet1.set_cell_value("A1", 10)
        workbook.add_defined_name("InputValue", "Sheet1!A1")
        sheet1.set_cell_value("B1", "=InputValue*2")
        
        cell_b1 = sheet1.get_cell("B1")
        assert "InputValue" in cell_b1.formula.terms


class TestWorkbookPerformance:
    """Test performance characteristics of hierarchical model."""
    
    def test_worksheet_lookup_performance(self):
        """Should provide O(1) worksheet lookup performance."""
        workbook = Workbook(name="TestWorkbook")
        
        # Add many worksheets
        for i in range(100):
            workbook.add_worksheet(f"Sheet{i}")
        
        # Lookup should be fast regardless of number of sheets
        import time
        start_time = time.time()
        
        for i in range(100):
            sheet = workbook.get_worksheet(f"Sheet{i}")
            assert sheet.name == f"Sheet{i}"
        
        elapsed = time.time() - start_time
        assert elapsed < 0.1  # Should be very fast
    
    def test_cell_access_within_sheet(self):
        """Should provide efficient cell access within a sheet."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        # Add many cells
        for row in range(1, 101):
            for col in ['A', 'B', 'C', 'D', 'E']:
                sheet.set_cell_value(f"{col}{row}", f"Value{col}{row}")
        
        # Access should be efficient
        import time
        start_time = time.time()
        
        for row in range(1, 101):
            for col in ['A', 'B', 'C', 'D', 'E']:
                cell = sheet.get_cell(f"{col}{row}")
                assert cell.value == f"Value{col}{row}"
        
        elapsed = time.time() - start_time
        assert elapsed < 0.5  # Should be reasonably fast