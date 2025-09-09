"""
Acceptance Tests for Hierarchical Model - Worksheet Operations

Tests worksheet-level functionality including cell management,
range operations, and local address handling.
"""
import pytest
from xlcalculator.hierarchical_model import Workbook, Worksheet, Cell, Range


class TestWorksheetBasicOperations:
    """Test basic worksheet operations and cell management."""
    
    def test_create_worksheet(self):
        """Should create worksheet with proper initialization."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        
        assert worksheet.name == "Sheet1"
        assert worksheet.workbook is workbook
        assert len(worksheet.cells) == 0
        assert len(worksheet.ranges) == 0
        assert worksheet.visible is True
    
    def test_set_cell_value_with_local_address(self):
        """Should set cell value using local address like 'A1'."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        
        worksheet.set_cell_value("A1", 42)
        
        assert "A1" in worksheet.cells
        cell = worksheet.cells["A1"]
        assert cell.value == 42
        assert cell.address == "A1"
        assert cell.worksheet is worksheet
    
    def test_get_cell_with_local_address(self):
        """Should retrieve cell using local address."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        worksheet.set_cell_value("B2", "Hello")
        
        cell = worksheet.get_cell("B2")
        
        assert cell.value == "Hello"
        assert cell.address == "B2"
        assert cell.worksheet is worksheet
    
    def test_get_nonexistent_cell_creates_empty_cell(self):
        """Should create empty cell if it doesn't exist."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        
        cell = worksheet.get_cell("Z99")
        
        assert cell.value is None
        assert cell.address == "Z99"
        assert cell.worksheet is worksheet
        assert "Z99" in worksheet.cells
    
    def test_full_address_property(self):
        """Should provide full address for cells."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        worksheet.set_cell_value("C3", 100)
        
        cell = worksheet.get_cell("C3")
        
        assert cell.full_address == "Sheet1!C3"
        assert worksheet.get_full_address("C3") == "Sheet1!C3"


class TestWorksheetRangeOperations:
    """Test worksheet range operations and management."""
    
    def test_create_range(self):
        """Should create range within worksheet."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        
        # Set some values
        worksheet.set_cell_value("A1", 1)
        worksheet.set_cell_value("A2", 2)
        worksheet.set_cell_value("B1", 3)
        worksheet.set_cell_value("B2", 4)
        
        range_obj = worksheet.get_range("A1:B2")
        
        assert range_obj.address == "A1:B2"
        assert range_obj.worksheet is worksheet
        assert range_obj.full_address == "Sheet1!A1:B2"
    
    def test_range_cell_access(self):
        """Should access cells within a range."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        
        # Set values in 2x2 range
        worksheet.set_cell_value("A1", 1)
        worksheet.set_cell_value("A2", 2)
        worksheet.set_cell_value("B1", 3)
        worksheet.set_cell_value("B2", 4)
        
        range_obj = worksheet.get_range("A1:B2")
        
        # Test cell access by position
        assert range_obj.get_cell(0, 0).value == 1  # A1
        assert range_obj.get_cell(0, 1).value == 3  # B1
        assert range_obj.get_cell(1, 0).value == 2  # A2
        assert range_obj.get_cell(1, 1).value == 4  # B2
    
    def test_range_cells_property(self):
        """Should provide 2D array of cells in range."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        
        # Set values
        worksheet.set_cell_value("A1", "a1")
        worksheet.set_cell_value("A2", "a2")
        worksheet.set_cell_value("B1", "b1")
        worksheet.set_cell_value("B2", "b2")
        
        range_obj = worksheet.get_range("A1:B2")
        cells = range_obj.cells
        
        assert len(cells) == 2  # 2 rows
        assert len(cells[0]) == 2  # 2 columns
        assert cells[0][0].value == "a1"  # A1
        assert cells[0][1].value == "b1"  # B1
        assert cells[1][0].value == "a2"  # A2
        assert cells[1][1].value == "b2"  # B2
    
    def test_single_cell_range(self):
        """Should handle single cell as a range."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        worksheet.set_cell_value("C5", 99)
        
        range_obj = worksheet.get_range("C5:C5")
        
        assert range_obj.address == "C5:C5"
        assert len(range_obj.cells) == 1
        assert len(range_obj.cells[0]) == 1
        assert range_obj.cells[0][0].value == 99


class TestWorksheetFormulaHandling:
    """Test worksheet formula handling and dependencies."""
    
    def test_set_formula_in_cell(self):
        """Should set formula in cell with proper parsing."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        
        worksheet.set_cell_value("A1", 10)
        worksheet.set_cell_value("B1", "=A1*2")
        
        cell_b1 = worksheet.get_cell("B1")
        assert cell_b1.formula is not None
        assert cell_b1.formula.formula == "=A1*2"
        assert "A1" in cell_b1.formula.terms
    
    def test_formula_with_range_reference(self):
        """Should handle formulas with range references."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        
        # Set values for SUM
        worksheet.set_cell_value("A1", 1)
        worksheet.set_cell_value("A2", 2)
        worksheet.set_cell_value("A3", 3)
        worksheet.set_cell_value("B1", "=SUM(A1:A3)")
        
        cell_b1 = worksheet.get_cell("B1")
        assert cell_b1.formula is not None
        assert "A1:A3" in cell_b1.formula.terms
    
    def test_cross_sheet_formula_reference(self):
        """Should handle formulas referencing other sheets."""
        workbook = Workbook(name="TestWorkbook")
        sheet1 = Worksheet(name="Sheet1", workbook=workbook)
        sheet2 = Worksheet(name="Sheet2", workbook=workbook)
        workbook.worksheets["Sheet1"] = sheet1
        workbook.worksheets["Sheet2"] = sheet2
        
        sheet1.set_cell_value("A1", 100)
        sheet2.set_cell_value("B1", "=Sheet1!A1+10")
        
        cell_b1 = sheet2.get_cell("B1")
        assert cell_b1.formula is not None
        assert "Sheet1!A1" in cell_b1.formula.terms
    
    def test_local_vs_full_address_in_formulas(self):
        """Should properly handle local vs full addresses in formulas."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="MySheet", workbook=workbook)
        
        worksheet.set_cell_value("A1", 5)
        worksheet.set_cell_value("B1", "=A1*3")  # Local reference
        
        cell_b1 = worksheet.get_cell("B1")
        # Formula should store local reference but be aware of sheet context
        assert "A1" in cell_b1.formula.terms
        assert cell_b1.formula.sheet_name == "MySheet"


class TestWorksheetCellProperties:
    """Test cell properties and metadata within worksheets."""
    
    def test_cell_row_and_column_properties(self):
        """Should provide correct row and column properties."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        worksheet.set_cell_value("C5", "test")
        
        cell = worksheet.get_cell("C5")
        
        assert cell.row == 5
        assert cell.column == "C"
        assert cell.column_index == 3  # C is the 3rd column
    
    def test_cell_address_parsing(self):
        """Should correctly parse various cell address formats."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        
        test_cases = [
            ("A1", 1, "A", 1),
            ("Z26", 26, "Z", 26),
            ("AA1", 1, "AA", 27),
            ("AB100", 100, "AB", 28),
        ]
        
        for address, expected_row, expected_col, expected_col_idx in test_cases:
            worksheet.set_cell_value(address, f"value_{address}")
            cell = worksheet.get_cell(address)
            
            assert cell.row == expected_row, f"Row mismatch for {address}"
            assert cell.column == expected_col, f"Column mismatch for {address}"
            assert cell.column_index == expected_col_idx, f"Column index mismatch for {address}"
    
    def test_cell_type_handling(self):
        """Should handle different cell value types correctly."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        
        test_values = [
            ("A1", 42, int),
            ("A2", 3.14, float),
            ("A3", "Hello", str),
            ("A4", True, bool),
            ("A5", None, type(None)),
        ]
        
        for address, value, expected_type in test_values:
            worksheet.set_cell_value(address, value)
            cell = worksheet.get_cell(address)
            
            assert cell.value == value
            assert type(cell.value) == expected_type


class TestWorksheetVisibilityAndProperties:
    """Test worksheet visibility and other properties."""
    
    def test_worksheet_visibility(self):
        """Should handle worksheet visibility property."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        
        assert worksheet.visible is True  # Default visible
        
        worksheet.visible = False
        assert worksheet.visible is False
        
        worksheet.visible = True
        assert worksheet.visible is True
    
    def test_worksheet_name_property(self):
        """Should maintain worksheet name property."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="MySpecialSheet", workbook=workbook)
        
        assert worksheet.name == "MySpecialSheet"
        
        # Test name change
        worksheet.name = "RenamedSheet"
        assert worksheet.name == "RenamedSheet"
    
    def test_worksheet_workbook_reference(self):
        """Should maintain proper reference to parent workbook."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        
        assert worksheet.workbook is workbook
        assert worksheet.workbook.name == "TestWorkbook"


class TestWorksheetPerformance:
    """Test performance characteristics of worksheet operations."""
    
    def test_cell_access_performance(self):
        """Should provide efficient cell access within worksheet."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        
        # Create many cells
        for row in range(1, 1001):  # 1000 rows
            worksheet.set_cell_value(f"A{row}", row)
        
        # Access should be O(1)
        import time
        start_time = time.time()
        
        for row in range(1, 1001):
            cell = worksheet.get_cell(f"A{row}")
            assert cell.value == row
        
        elapsed = time.time() - start_time
        assert elapsed < 1.0  # Should be fast
    
    def test_range_creation_performance(self):
        """Should efficiently create ranges of various sizes."""
        workbook = Workbook(name="TestWorkbook")
        worksheet = Worksheet(name="Sheet1", workbook=workbook)
        
        # Set up a large grid
        for row in range(1, 101):
            for col_idx in range(1, 11):  # A through J
                col = chr(ord('A') + col_idx - 1)
                worksheet.set_cell_value(f"{col}{row}", f"{col}{row}")
        
        # Create various range sizes
        import time
        start_time = time.time()
        
        ranges = [
            "A1:J10",    # 10x10
            "A1:J50",    # 10x50
            "A1:J100",   # 10x100
        ]
        
        for range_addr in ranges:
            range_obj = worksheet.get_range(range_addr)
            assert range_obj.address == range_addr
        
        elapsed = time.time() - start_time
        assert elapsed < 0.5  # Should be reasonably fast