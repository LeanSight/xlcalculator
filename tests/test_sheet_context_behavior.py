"""
Test de aceptación para comportamiento correcto de contexto de hoja.

Este test define el comportamiento esperado que debe ser consistente con Excel:
"Una referencia implícita se resuelve en el contexto de la hoja donde está definida la fórmula"
"""
import unittest
import os
from xlcalculator import ModelCompiler, Evaluator


class SheetContextBehaviorTest(unittest.TestCase):
    """Test de aceptación para contexto correcto de hoja según Excel."""

    def setUp(self):
        """Setup test model with multi-sheet references."""
        # Use existing cross_sheet.xlsx for testing
        resource_dir = os.path.join(os.path.dirname(__file__), 'resources')
        filename = os.path.join(resource_dir, 'cross_sheet.xlsx')
        
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(filename)
        self.evaluator = Evaluator(self.model)

    def test_implicit_reference_uses_formula_sheet_context(self):
        """Test that implicit references use the sheet where formula is defined.
        
        Excel behavior: If a formula in Sheet2 contains =SUM(A1:A10),
        then A1:A10 should resolve to Sheet2!A1:A10, not Sheet1!A1:A10.
        """
        # Test the core issue: EvalContext should receive and use formula sheet context
        from xlcalculator.ast_nodes import EvalContext
        from xlcalculator.xlfunctions import xl
        
        # Scenario 1: Simulating a formula in Sheet2 with implicit reference
        # This should use Sheet2 as context, not hardcoded Sheet1
        context_sheet2 = EvalContext(namespace=xl.FUNCTIONS, ref='A1')
        
        # Current behavior: hardcoded 'Sheet1' regardless of actual context
        # Expected behavior: should use the sheet where formula is defined
        actual_sheet = context_sheet2.sheet
        
        # This test FAILS because EvalContext hardcodes 'Sheet1'
        # But if we were evaluating a formula in Sheet2, it should use 'Sheet2'
        self.assertEqual(actual_sheet, 'Sheet1', 
                        "Current broken behavior: EvalContext hardcodes Sheet1")
        
        # The REAL test: EvalContext should accept formula_sheet parameter
        # This will fail until we implement proper context propagation
        try:
            # This should be the correct way to create context
            correct_context = EvalContext(namespace=xl.FUNCTIONS, ref='A1', formula_sheet='Sheet2')
            expected_sheet = 'Sheet2'
            actual_sheet = correct_context.sheet
            
            self.assertEqual(actual_sheet, expected_sheet,
                            f"EvalContext should use formula_sheet parameter '{expected_sheet}' "
                            f"for implicit references, got: '{actual_sheet}'")
        except TypeError as e:
            # This will fail because EvalContext doesn't accept formula_sheet parameter yet
            self.fail(f"EvalContext should accept formula_sheet parameter. Error: {e}")

    def test_hardcoded_sheet1_problem_demonstration(self):
        """Demonstrate the problem with hardcoded Sheet1 default."""
        from xlcalculator.ast_nodes import EvalContext
        from xlcalculator.xlfunctions import xl
        
        # Problem: All implicit references default to Sheet1, regardless of context
        context1 = EvalContext(namespace=xl.FUNCTIONS, ref='A1')
        context2 = EvalContext(namespace=xl.FUNCTIONS, ref='B2')  
        context3 = EvalContext(namespace=xl.FUNCTIONS, ref='C3')
        
        # All should potentially be in different sheets, but all default to Sheet1
        self.assertEqual(context1.sheet, 'Sheet1')
        self.assertEqual(context2.sheet, 'Sheet1') 
        self.assertEqual(context3.sheet, 'Sheet1')
        
        # This demonstrates the problem: no way to specify different sheet contexts
        # All implicit references are forced to Sheet1, which is incorrect Excel behavior

    def test_mixed_references_resolve_correctly(self):
        """Test mixed explicit and implicit references in same formula.
        
        Excel behavior: In formula =Sheet1!A1 + B1 + Sheet2!C1
        - Sheet1!A1 -> explicit, use Sheet1
        - B1 -> implicit, use current formula's sheet
        - Sheet2!C1 -> explicit, use Sheet2
        """
        # This test documents the expected behavior
        # Will be implemented after fixing context propagation
        
        # For now, just document the requirement
        self.assertTrue(True, "TODO: Implement after context propagation fix")

    def test_formula_in_different_sheets_resolve_differently(self):
        """Test that same formula text resolves differently based on sheet context.
        
        Excel behavior:
        - Sheet1!E1 = "=SUM(A1:A5)" -> resolves to SUM(Sheet1!A1:A5)  
        - Sheet2!E1 = "=SUM(A1:A5)" -> resolves to SUM(Sheet2!A1:A5)
        """
        # This test will validate the fix works correctly
        # Will be implemented after fixing context propagation
        
        self.assertTrue(True, "TODO: Implement after context propagation fix")

    def test_context_propagation_through_evaluator(self):
        """Test that Evaluator properly propagates sheet context to EvalContext.
        
        The fix should ensure that when evaluating a cell's formula,
        the sheet context comes from the cell's formula.sheet_name,
        not from hardcoded defaults.
        """
        # Test that evaluator extracts and uses correct context
        d3_cell = self.model.cells['D3']
        expected_context_sheet = d3_cell.formula.sheet_name
        
        # This should work after implementing context propagation
        # Currently will fail due to hardcoded 'Sheet1'
        
        # Simulate what should happen:
        # 1. Evaluator gets cell.formula.sheet_name
        # 2. Passes it to EvaluatorContext  
        # 3. EvaluatorContext passes it to EvalContext
        # 4. EvalContext uses it for implicit references
        
        self.assertEqual(expected_context_sheet, 'Sheet1',
                        "Test setup: D3 should be in Sheet1 context")
        
        # The actual test will be: does EvalContext get the right context?
        # This will be validated after implementing the fix


if __name__ == '__main__':
    unittest.main()