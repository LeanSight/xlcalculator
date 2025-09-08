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
        # Use proper sheet context test file
        resource_dir = os.path.join(os.path.dirname(__file__), 'resources')
        filename = os.path.join(resource_dir, 'sheet_context_test.xlsx')
        
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(filename)
        self.evaluator = Evaluator(self.model)



    def test_mixed_references_resolve_correctly(self):
        """Test mixed explicit and implicit references in same formula.
        
        Excel behavior: In formula =A1 + Sheet2!A1
        - A1 -> implicit, use current formula's sheet
        - Sheet2!A1 -> explicit, use Sheet2
        """
        # Test Sheet1!C2 = A1 + Sheet2!A1
        # Should resolve to Sheet1!A1 + Sheet2!A1 = 10 + 100 = 110
        result_sheet1 = self.evaluator.evaluate('Sheet1!C2')
        self.assertEqual(result_sheet1, 110, 
                        "Sheet1!C2 should resolve A1 to Sheet1!A1 (10) + Sheet2!A1 (100) = 110")
        
        # Test Sheet2!C2 = A1 + Sheet1!A1  
        # Should resolve to Sheet2!A1 + Sheet1!A1 = 100 + 10 = 110
        result_sheet2 = self.evaluator.evaluate('Sheet2!C2')
        self.assertEqual(result_sheet2, 110,
                        "Sheet2!C2 should resolve A1 to Sheet2!A1 (100) + Sheet1!A1 (10) = 110")

    def test_formula_in_different_sheets_resolve_differently(self):
        """Test that same formula text resolves differently based on sheet context.
        
        Excel behavior:
        - Sheet1!C1 = "=SUM(A1:A3)" -> resolves to SUM(Sheet1!A1:A3) = 60
        - Sheet2!C1 = "=SUM(A1:A3)" -> resolves to SUM(Sheet2!A1:A3) = 600
        """
        # Test Sheet1!C1 = SUM(A1:A3) should sum Sheet1 values
        result_sheet1 = self.evaluator.evaluate('Sheet1!C1')
        self.assertEqual(result_sheet1, 60,
                        "Sheet1!C1 should resolve SUM(A1:A3) to Sheet1 context: 10+20+30=60")
        
        # Test Sheet2!C1 = SUM(A1:A3) should sum Sheet2 values  
        result_sheet2 = self.evaluator.evaluate('Sheet2!C1')
        self.assertEqual(result_sheet2, 600,
                        "Sheet2!C1 should resolve SUM(A1:A3) to Sheet2 context: 100+200+300=600")

    def test_context_propagation_through_evaluator(self):
        """Test that Evaluator properly propagates sheet context to EvalContext.
        
        The fix should ensure that when evaluating a cell's formula,
        the sheet context comes from the cell's formula.sheet_name,
        not from hardcoded defaults.
        """
        # Test cross-sheet references work correctly
        # Sheet1!D1 = Sheet2!C1 (should reference Sheet2's SUM result = 600)
        result_d1 = self.evaluator.evaluate('Sheet1!D1')
        self.assertEqual(result_d1, 600,
                        "Sheet1!D1 should reference Sheet2!C1 = 600")
        
        # Sheet2!D1 = Sheet1!C1 (should reference Sheet1's SUM result = 60)  
        result_d2 = self.evaluator.evaluate('Sheet2!D1')
        self.assertEqual(result_d2, 60,
                        "Sheet2!D1 should reference Sheet1!C1 = 60")
        
        # Verify the formulas have correct sheet context
        sheet1_c1_formula = self.model.cells['Sheet1!C1'].formula
        sheet2_c1_formula = self.model.cells['Sheet2!C1'].formula
        
        self.assertEqual(sheet1_c1_formula.sheet_name, 'Sheet1',
                        "Sheet1!C1 formula should have Sheet1 context")
        self.assertEqual(sheet2_c1_formula.sheet_name, 'Sheet2', 
                        "Sheet2!C1 formula should have Sheet2 context")


if __name__ == '__main__':
    unittest.main()