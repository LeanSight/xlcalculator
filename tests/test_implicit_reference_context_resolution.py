"""
Test de consistencia con Excel para contexto de hoja.

Verifica que xlcalculator se comporta exactamente como Excel
en la resolución de contexto de hoja para referencias implícitas.
"""
import unittest
from xlcalculator.ast_nodes import EvalContext
from xlcalculator.xlfunctions import xl


class ExcelConsistencyTest(unittest.TestCase):
    """Test de consistencia con comportamiento de Excel."""

    def test_excel_behavior_implicit_references_use_formula_sheet(self):
        """Test que replica el comportamiento exacto de Excel.
        
        En Excel:
        - Si una fórmula está en Sheet2 y contiene =SUM(A1:A5)
        - Excel resuelve A1:A5 como Sheet2!A1:A5
        - NO como Sheet1!A1:A5
        """
        # Simular fórmula en Sheet2 con referencia implícita
        context = EvalContext(
            namespace=xl.FUNCTIONS, 
            ref='A1', 
            formula_sheet='Sheet2'
        )
        
        # Verificar que usa Sheet2 como contexto (comportamiento Excel)
        self.assertEqual(context.sheet, 'Sheet2')
        self.assertEqual(context.cell_ref.sheet, 'Sheet2')
        self.assertEqual(context.cell_ref.address, 'A1')
        self.assertFalse(context.cell_ref.is_sheet_explicit)

    def test_excel_behavior_explicit_references_override_context(self):
        """Test que referencias explícitas ignoran contexto de fórmula.
        
        En Excel:
        - Si una fórmula en Sheet2 contiene =Sheet1!A1
        - Excel usa Sheet1, no Sheet2
        """
        # Simular fórmula en Sheet2 con referencia explícita a Sheet1
        context = EvalContext(
            namespace=xl.FUNCTIONS,
            ref='Sheet1!A1',
            formula_sheet='Sheet2'  # Contexto de fórmula
        )
        
        # Verificar que usa Sheet1 (explícito), no Sheet2 (contexto)
        self.assertEqual(context.sheet, 'Sheet1')
        self.assertEqual(context.cell_ref.sheet, 'Sheet1')
        self.assertEqual(context.cell_ref.address, 'A1')
        self.assertTrue(context.cell_ref.is_sheet_explicit)

    def test_excel_behavior_mixed_references_in_formula(self):
        """Test comportamiento Excel con referencias mixtas.
        
        En Excel, en fórmula: =Sheet1!A1 + B1 + Sheet3!C1
        - Sheet1!A1 -> usa Sheet1 (explícito)
        - B1 -> usa hoja de la fórmula (implícito)  
        - Sheet3!C1 -> usa Sheet3 (explícito)
        """
        formula_sheet = 'Sheet2'
        
        # Referencia explícita
        context1 = EvalContext(xl.FUNCTIONS, 'Sheet1!A1', formula_sheet=formula_sheet)
        self.assertEqual(context1.sheet, 'Sheet1')
        self.assertTrue(context1.cell_ref.is_sheet_explicit)
        
        # Referencia implícita
        context2 = EvalContext(xl.FUNCTIONS, 'B1', formula_sheet=formula_sheet)
        self.assertEqual(context2.sheet, 'Sheet2')  # Usa contexto de fórmula
        self.assertFalse(context2.cell_ref.is_sheet_explicit)
        
        # Otra referencia explícita
        context3 = EvalContext(xl.FUNCTIONS, 'Sheet3!C1', formula_sheet=formula_sheet)
        self.assertEqual(context3.sheet, 'Sheet3')
        self.assertTrue(context3.cell_ref.is_sheet_explicit)

    def test_excel_behavior_fallback_when_no_context(self):
        """Test comportamiento cuando no hay contexto de fórmula.
        
        Esto puede ocurrir en casos edge o durante inicialización.
        Debe usar fallback seguro.
        """
        # Sin contexto de fórmula
        context = EvalContext(xl.FUNCTIONS, 'A1')  # No formula_sheet
        
        # Debe usar fallback seguro (Sheet1)
        self.assertEqual(context.sheet, 'Sheet1')
        self.assertFalse(context.cell_ref.is_sheet_explicit)

    def test_excel_behavior_range_references(self):
        """Test que rangos también usan contexto correcto."""
        # Rango implícito en contexto de Sheet3
        context = EvalContext(xl.FUNCTIONS, 'A1:B5', formula_sheet='Sheet3')
        
        self.assertEqual(context.sheet, 'Sheet3')
        self.assertEqual(context.cell_ref.address, 'A1:B5')
        self.assertFalse(context.cell_ref.is_sheet_explicit)


if __name__ == '__main__':
    unittest.main()