"""
Acceptance Tests for Hierarchical Model - Evaluator Integration

Tests the integration between the hierarchical model and the evaluator,
ensuring formulas work correctly with the new structure.
"""
import pytest
from xlcalculator.hierarchical_model import Workbook
from xlcalculator.evaluator import Evaluator
from xlcalculator.model import ModelCompiler


class TestEvaluatorHierarchicalIntegration:
    """Test evaluator integration with hierarchical model."""
    
    def test_evaluate_simple_formula_in_hierarchical_model(self):
        """Should evaluate simple formulas in hierarchical model."""
        compiler = ModelCompiler()
        
        input_dict = {
            "A1": 10,
            "B1": "=A1*2"
        }
        
        workbook = compiler.read_and_parse_dict_hierarchical(input_dict)
        evaluator = Evaluator(workbook)
        result = evaluator.evaluate("Sheet1!B1")
        
        assert result == 20
    
    def test_evaluate_cross_sheet_formula(self):
        """Should evaluate formulas with cross-sheet references."""
        workbook = Workbook(name="TestWorkbook")
        sheet1 = workbook.add_worksheet("Sheet1")
        sheet2 = workbook.add_worksheet("Sheet2")
        
        sheet1.set_cell_value("A1", 15)
        sheet2.set_cell_value("B1", "=Sheet1!A1+5")
        
        evaluator = Evaluator(workbook)
        result = evaluator.evaluate("Sheet2!B1")
        
        assert result == 20
    
    def test_evaluate_range_formula(self):
        """Should evaluate formulas with range references."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        # Set up range values
        sheet.set_cell_value("A1", 1)
        sheet.set_cell_value("A2", 2)
        sheet.set_cell_value("A3", 3)
        sheet.set_cell_value("B1", "=SUM(A1:A3)")
        
        evaluator = Evaluator(workbook)
        result = evaluator.evaluate("Sheet1!B1")
        
        assert result == 6
    
    def test_evaluate_with_defined_names(self):
        """Should evaluate formulas using defined names."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        sheet.set_cell_value("A1", 25)
        workbook.add_defined_name("MyValue", "Sheet1!A1")
        sheet.set_cell_value("B1", "=MyValue*2")
        
        evaluator = Evaluator(workbook)
        result = evaluator.evaluate("Sheet1!B1")
        
        assert result == 50
    
    def test_evaluate_nested_formulas(self):
        """Should evaluate nested formula dependencies."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        sheet.set_cell_value("A1", 5)
        sheet.set_cell_value("B1", "=A1*2")
        sheet.set_cell_value("C1", "=B1+10")
        sheet.set_cell_value("D1", "=C1*3")
        
        evaluator = Evaluator(workbook)
        
        # Test intermediate results
        assert evaluator.evaluate("Sheet1!B1") == 10
        assert evaluator.evaluate("Sheet1!C1") == 20
        assert evaluator.evaluate("Sheet1!D1") == 60
    
    def test_evaluate_circular_reference_detection(self):
        """Should detect circular references in hierarchical model."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        sheet.set_cell_value("A1", "=B1+1")
        sheet.set_cell_value("B1", "=A1+1")
        
        evaluator = Evaluator(workbook)
        
        with pytest.raises(RuntimeError, match="Cycle detected"):
            evaluator.evaluate("Sheet1!A1")


class TestEvaluatorContextWithHierarchy:
    """Test evaluator context handling with hierarchical model."""
    
    def test_context_sheet_awareness(self):
        """Should maintain sheet context during evaluation."""
        workbook = Workbook(name="TestWorkbook")
        sheet1 = workbook.add_worksheet("Data")
        sheet2 = workbook.add_worksheet("Calculations")
        
        # Set up data in Data sheet
        sheet1.set_cell_value("A1", 100)
        sheet1.set_cell_value("A2", 200)
        
        # Reference from Calculations sheet
        sheet2.set_cell_value("B1", "=Data!A1+Data!A2")
        
        evaluator = Evaluator(workbook)
        result = evaluator.evaluate("Calculations!B1")
        
        assert result == 300
    
    def test_context_local_references(self):
        """Should handle local references within sheet context."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("MySheet")
        
        sheet.set_cell_value("A1", 10)
        sheet.set_cell_value("A2", 20)
        sheet.set_cell_value("B1", "=A1+A2")  # Local references
        
        evaluator = Evaluator(workbook)
        result = evaluator.evaluate("MySheet!B1")
        
        assert result == 30
    
    def test_context_mixed_references(self):
        """Should handle mix of local and cross-sheet references."""
        workbook = Workbook(name="TestWorkbook")
        sheet1 = workbook.add_worksheet("Sheet1")
        sheet2 = workbook.add_worksheet("Sheet2")
        
        sheet1.set_cell_value("A1", 5)
        sheet2.set_cell_value("A1", 10)
        sheet2.set_cell_value("B1", "=A1+Sheet1!A1")  # Local A1 + Sheet1 A1
        
        evaluator = Evaluator(workbook)
        result = evaluator.evaluate("Sheet2!B1")
        
        assert result == 15  # 10 + 5


class TestEvaluatorReferenceAwareFunctions:
    """Test reference-aware functions with hierarchical model."""
    
    def test_row_function_with_hierarchical_model(self):
        """Should evaluate ROW function correctly in hierarchical model."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        sheet.set_cell_value("A5", "=ROW()")
        sheet.set_cell_value("B5", "=ROW(A1)")
        
        evaluator = Evaluator(workbook)
        
        assert evaluator.evaluate("Sheet1!A5") == 5
        assert evaluator.evaluate("Sheet1!B5") == 1
    
    def test_column_function_with_hierarchical_model(self):
        """Should evaluate COLUMN function correctly in hierarchical model."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        sheet.set_cell_value("C3", "=COLUMN()")
        sheet.set_cell_value("D3", "=COLUMN(A1)")
        
        evaluator = Evaluator(workbook)
        
        assert evaluator.evaluate("Sheet1!C3") == 3
        assert evaluator.evaluate("Sheet1!D3") == 1
    
    def test_offset_function_with_hierarchical_model(self):
        """Should evaluate OFFSET function correctly in hierarchical model."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        # Set up grid
        sheet.set_cell_value("A1", 10)
        sheet.set_cell_value("B1", 20)
        sheet.set_cell_value("A2", 30)
        sheet.set_cell_value("B2", 40)
        
        sheet.set_cell_value("D1", "=OFFSET(A1,1,1,1,1)")  # Should get B2
        
        evaluator = Evaluator(workbook)
        result = evaluator.evaluate("Sheet1!D1")
        
        assert result == 40
    
    def test_indirect_function_with_hierarchical_model(self):
        """Should evaluate INDIRECT function correctly in hierarchical model."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        sheet.set_cell_value("A1", 100)
        sheet.set_cell_value("B1", "A1")
        sheet.set_cell_value("C1", "=INDIRECT(B1)")
        
        evaluator = Evaluator(workbook)
        result = evaluator.evaluate("Sheet1!C1")
        
        assert result == 100


class TestEvaluatorBackwardCompatibility:
    """Test backward compatibility with existing evaluator patterns."""
    
    def test_flat_model_compatibility(self):
        """Should maintain compatibility with flat model evaluation."""
        # Create both flat and hierarchical models with same data
        compiler = ModelCompiler()
        
        input_dict = {
            "A1": 10,
            "B1": "=A1*2",
            "C1": "=SUM(A1:B1)"
        }
        
        flat_model = compiler.read_and_parse_dict(input_dict)
        hierarchical_model = compiler.read_and_parse_dict_hierarchical(input_dict)
        
        flat_evaluator = Evaluator(flat_model)
        hierarchical_evaluator = Evaluator(hierarchical_model)
        
        # Both should give same results
        assert flat_evaluator.evaluate("Sheet1!B1") == hierarchical_evaluator.evaluate("Sheet1!B1")
        assert flat_evaluator.evaluate("Sheet1!C1") == hierarchical_evaluator.evaluate("Sheet1!C1")
    
    def test_address_resolution_compatibility(self):
        """Should resolve addresses consistently between models."""
        compiler = ModelCompiler()
        
        input_dict = {
            "Sheet1!A1": 5,
            "Sheet2!A1": 10,
            "Sheet2!B1": "=Sheet1!A1+A1"
        }
        
        hierarchical_model = compiler.read_and_parse_dict_hierarchical(input_dict)
        evaluator = Evaluator(hierarchical_model)
        
        result = evaluator.evaluate("Sheet2!B1")
        assert result == 15  # 5 + 10
    
    def test_defined_names_compatibility(self):
        """Should handle defined names consistently."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        sheet.set_cell_value("A1", 42)
        workbook.add_defined_name("MyConstant", "Sheet1!A1")
        sheet.set_cell_value("B1", "=MyConstant*2")
        
        evaluator = Evaluator(workbook)
        result = evaluator.evaluate("Sheet1!B1")
        
        assert result == 84


class TestEvaluatorPerformance:
    """Test evaluator performance with hierarchical model."""
    
    def test_large_workbook_evaluation_performance(self):
        """Should efficiently evaluate formulas in large workbooks."""
        workbook = Workbook(name="TestWorkbook")
        
        # Create multiple sheets with many cells
        for sheet_num in range(1, 4):  # 3 sheets
            sheet = workbook.add_worksheet(f"Sheet{sheet_num}")
            
            # Add base values
            for row in range(1, 101):  # 100 rows
                sheet.set_cell_value(f"A{row}", row)
                sheet.set_cell_value(f"B{row}", f"=A{row}*2")
        
        evaluator = Evaluator(workbook)
        
        # Evaluate many formulas
        import time
        start_time = time.time()
        
        for sheet_num in range(1, 4):
            for row in range(1, 101):
                result = evaluator.evaluate(f"Sheet{sheet_num}!B{row}")
                assert result == row * 2
        
        elapsed = time.time() - start_time
        assert elapsed < 5.0  # Should complete in reasonable time
    
    def test_cross_sheet_evaluation_performance(self):
        """Should efficiently handle cross-sheet evaluations."""
        workbook = Workbook(name="TestWorkbook")
        data_sheet = workbook.add_worksheet("Data")
        calc_sheet = workbook.add_worksheet("Calculations")
        
        # Set up data
        for row in range(1, 51):  # 50 rows
            data_sheet.set_cell_value(f"A{row}", row)
            calc_sheet.set_cell_value(f"B{row}", f"=Data!A{row}*3")
        
        evaluator = Evaluator(workbook)
        
        import time
        start_time = time.time()
        
        for row in range(1, 51):
            result = evaluator.evaluate(f"Calculations!B{row}")
            assert result == row * 3
        
        elapsed = time.time() - start_time
        assert elapsed < 2.0  # Should be efficient