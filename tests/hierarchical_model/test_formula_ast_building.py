"""
ATDD Tests for Formula AST Building in Hierarchical Model

These tests define the expected behavior for automatic AST building
when formulas are created in the hierarchical model.

Test-Driven Development approach:
1. Write failing tests that define expected behavior
2. Implement the minimum code to make tests pass
3. Refactor while keeping tests green
"""
import pytest
from xlcalculator.hierarchical_model import Workbook
from xlcalculator.evaluator import Evaluator
from xlcalculator.model import ModelCompiler


class TestFormulaASTBuildingBasic:
    """Test basic AST building functionality for hierarchical model formulas."""
    
    def test_formula_ast_built_automatically_on_creation(self):
        """Should automatically build AST when formula is set in worksheet."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        # Set a formula value
        sheet.set_cell_value("B1", "=A1*2")
        
        # AST should be built automatically
        cell = sheet.get_cell("B1")
        assert cell.formula is not None
        assert cell.formula.formula == "=A1*2"
        assert cell.formula.ast is not None, "AST should be built automatically"
        assert str(cell.formula.ast) == "(A1) * (2)", "AST should be correctly parsed"
    
    def test_formula_ast_built_for_complex_expressions(self):
        """Should build AST for complex formula expressions."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        # Set complex formula
        sheet.set_cell_value("C1", "=SUM(A1:A3)+B1*2")
        
        cell = sheet.get_cell("C1")
        assert cell.formula is not None
        assert cell.formula.ast is not None, "Complex formula AST should be built"
    
    def test_formula_ast_built_for_cross_sheet_references(self):
        """Should build AST for formulas with cross-sheet references."""
        workbook = Workbook(name="TestWorkbook")
        sheet1 = workbook.add_worksheet("Sheet1")
        sheet2 = workbook.add_worksheet("Sheet2")
        
        # Set cross-sheet formula
        sheet2.set_cell_value("A1", "=Sheet1!B1+10")
        
        cell = sheet2.get_cell("A1")
        assert cell.formula is not None
        assert cell.formula.ast is not None, "Cross-sheet formula AST should be built"
    
    def test_non_formula_values_dont_create_ast(self):
        """Should not create AST for non-formula values."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        # Set non-formula values
        sheet.set_cell_value("A1", 42)
        sheet.set_cell_value("B1", "Hello")
        sheet.set_cell_value("C1", "=NotAFormula")  # String starting with = but not a formula
        
        # No AST should be created for non-formulas
        assert sheet.get_cell("A1").formula is None
        assert sheet.get_cell("B1").formula is None
        # Note: "=NotAFormula" will be treated as a formula and should have AST
        assert sheet.get_cell("C1").formula is not None
        assert sheet.get_cell("C1").formula.ast is not None
    
    def test_formula_update_rebuilds_ast(self):
        """Should rebuild AST when formula is updated."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        # Set initial formula
        sheet.set_cell_value("A1", "=B1+1")
        original_ast = sheet.get_cell("A1").formula.ast
        
        # Update formula
        sheet.set_cell_value("A1", "=B1*2")
        updated_ast = sheet.get_cell("A1").formula.ast
        
        # AST should be different
        assert updated_ast is not None
        assert str(updated_ast) != str(original_ast)
        assert str(updated_ast) == "(B1) * (2)"


class TestFormulaASTBuildingEvaluation:
    """Test that formulas with built AST can be evaluated correctly."""
    
    def test_formula_evaluation_works_with_auto_built_ast(self):
        """Should evaluate formulas correctly when AST is auto-built."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        # Set up data and formula
        sheet.set_cell_value("A1", 10)
        sheet.set_cell_value("B1", "=A1*2")
        
        # Evaluation should work
        evaluator = Evaluator(workbook)
        result = evaluator.evaluate("Sheet1!B1")
        
        assert result == 20
    
    def test_complex_formula_evaluation(self):
        """Should evaluate complex formulas with auto-built AST."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        # Set up data
        sheet.set_cell_value("A1", 1)
        sheet.set_cell_value("A2", 2)
        sheet.set_cell_value("A3", 3)
        sheet.set_cell_value("B1", "=SUM(A1:A3)")
        
        # Evaluation should work
        evaluator = Evaluator(workbook)
        result = evaluator.evaluate("Sheet1!B1")
        
        assert result == 6
    
    def test_cross_sheet_formula_evaluation(self):
        """Should evaluate cross-sheet formulas with auto-built AST."""
        workbook = Workbook(name="TestWorkbook")
        sheet1 = workbook.add_worksheet("Sheet1")
        sheet2 = workbook.add_worksheet("Sheet2")
        
        # Set up data
        sheet1.set_cell_value("A1", 15)
        sheet2.set_cell_value("B1", "=Sheet1!A1+5")
        
        # Evaluation should work
        evaluator = Evaluator(workbook)
        result = evaluator.evaluate("Sheet2!B1")
        
        assert result == 20
    
    def test_nested_formula_evaluation(self):
        """Should evaluate nested formulas with auto-built AST."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        # Set up nested formulas
        sheet.set_cell_value("A1", 5)
        sheet.set_cell_value("B1", "=A1*2")
        sheet.set_cell_value("C1", "=B1+10")
        sheet.set_cell_value("D1", "=C1*3")
        
        # All evaluations should work
        evaluator = Evaluator(workbook)
        
        assert evaluator.evaluate("Sheet1!B1") == 10
        assert evaluator.evaluate("Sheet1!C1") == 20
        assert evaluator.evaluate("Sheet1!D1") == 60


class TestFormulaASTBuildingWithDefinedNames:
    """Test AST building with defined names in hierarchical model."""
    
    def test_formula_with_defined_names_builds_ast(self):
        """Should build AST for formulas using defined names."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        # Set up defined name and formula
        sheet.set_cell_value("A1", 25)
        workbook.add_defined_name("MyValue", "Sheet1!A1")
        sheet.set_cell_value("B1", "=MyValue*2")
        
        # AST should be built
        cell = sheet.get_cell("B1")
        assert cell.formula is not None
        assert cell.formula.ast is not None
        
        # Evaluation should work
        evaluator = Evaluator(workbook)
        result = evaluator.evaluate("Sheet1!B1")
        assert result == 50


class TestFormulaASTBuildingCompatibility:
    """Test compatibility between different model creation methods."""
    
    def test_direct_vs_compiler_ast_consistency(self):
        """Should produce consistent AST whether created directly or via compiler."""
        # Direct creation
        workbook1 = Workbook(name="TestWorkbook")
        sheet1 = workbook1.add_worksheet("Sheet1")
        sheet1.set_cell_value("A1", 10)
        sheet1.set_cell_value("B1", "=A1*2")
        
        # Compiler creation
        compiler = ModelCompiler()
        input_dict = {"A1": 10, "B1": "=A1*2"}
        workbook2 = compiler.read_and_parse_dict_hierarchical(input_dict)
        
        # Both should have AST
        cell1 = sheet1.get_cell("B1")
        cell2 = workbook2.get_worksheet("Sheet1").get_cell("B1")
        
        assert cell1.formula.ast is not None
        assert cell2.formula.ast is not None
        
        # Both should evaluate to same result
        evaluator1 = Evaluator(workbook1)
        evaluator2 = Evaluator(workbook2)
        
        result1 = evaluator1.evaluate("Sheet1!B1")
        result2 = evaluator2.evaluate("Sheet1!B1")
        
        assert result1 == result2 == 20
    
    def test_hierarchical_vs_flat_model_consistency(self):
        """Should produce consistent results between hierarchical and flat models."""
        compiler = ModelCompiler()
        input_dict = {"A1": 10, "B1": "=A1*2", "C1": "=B1+5"}
        
        # Create both model types
        flat_model = compiler.read_and_parse_dict(input_dict)
        hierarchical_model = compiler.read_and_parse_dict_hierarchical(input_dict)
        
        # Both should evaluate consistently
        flat_evaluator = Evaluator(flat_model)
        hierarchical_evaluator = Evaluator(hierarchical_model)
        
        test_addresses = ["Sheet1!A1", "Sheet1!B1", "Sheet1!C1"]
        
        for addr in test_addresses:
            flat_result = flat_evaluator.evaluate(addr)
            hierarchical_result = hierarchical_evaluator.evaluate(addr)
            assert flat_result == hierarchical_result, f"Inconsistent results for {addr}"


class TestFormulaASTBuildingErrorHandling:
    """Test error handling in AST building process."""
    
    def test_invalid_formula_handling(self):
        """Should handle invalid formulas gracefully."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        # Set invalid formula
        sheet.set_cell_value("A1", "=INVALID_FUNCTION()")
        
        # Should create formula object but AST might be None or handle error
        cell = sheet.get_cell("A1")
        assert cell.formula is not None
        assert cell.formula.formula == "=INVALID_FUNCTION()"
        # AST building might fail, but should not crash the application
    
    def test_circular_reference_detection(self):
        """Should detect circular references in hierarchical model."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        # Create circular reference
        sheet.set_cell_value("A1", "=B1+1")
        sheet.set_cell_value("B1", "=A1+1")
        
        # Both should have AST
        assert sheet.get_cell("A1").formula.ast is not None
        assert sheet.get_cell("B1").formula.ast is not None
        
        # Evaluation should detect cycle
        evaluator = Evaluator(workbook)
        with pytest.raises(RuntimeError, match="Cycle detected"):
            evaluator.evaluate("Sheet1!A1")


class TestFormulaASTBuildingPerformance:
    """Test performance characteristics of AST building."""
    
    def test_bulk_formula_creation_performance(self):
        """Should efficiently build AST for many formulas."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        # Create many formulas
        import time
        start_time = time.time()
        
        for i in range(1, 101):  # 100 formulas
            sheet.set_cell_value(f"A{i}", i)
            sheet.set_cell_value(f"B{i}", f"=A{i}*2")
        
        elapsed = time.time() - start_time
        
        # Should complete in reasonable time
        assert elapsed < 5.0, f"AST building took too long: {elapsed}s"
        
        # All should have AST
        for i in range(1, 101):
            cell = sheet.get_cell(f"B{i}")
            assert cell.formula.ast is not None, f"Missing AST for B{i}"
    
    def test_ast_building_doesnt_slow_evaluation(self):
        """Should not significantly impact evaluation performance."""
        workbook = Workbook(name="TestWorkbook")
        sheet = workbook.add_worksheet("Sheet1")
        
        # Set up test data
        for i in range(1, 51):
            sheet.set_cell_value(f"A{i}", i)
            sheet.set_cell_value(f"B{i}", f"=A{i}*2")
        
        # Evaluation should be fast
        evaluator = Evaluator(workbook)
        
        import time
        start_time = time.time()
        
        for i in range(1, 51):
            result = evaluator.evaluate(f"Sheet1!B{i}")
            assert result == i * 2
        
        elapsed = time.time() - start_time
        assert elapsed < 2.0, f"Evaluation took too long: {elapsed}s"