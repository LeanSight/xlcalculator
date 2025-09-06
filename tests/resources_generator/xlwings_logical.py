#!/usr/bin/env python3
"""
Generate logical.xlsx using xlwings with Excel calculations.
This ensures Excel calculates the formula values for proper integration testing.
"""

import xlwings as xw
import os


def create_logical_excel_with_xlwings(filepath):
    """Create logical.xlsx with AND, OR, TRUE, FALSE tests using xlwings."""
    
    # Start Excel application
    app = xw.App(visible=False)
    try:
        wb = app.books.add()
        ws = wb.sheets[0]
        ws.name = "Sheet1"
        
        # Test data
        ws['A1'].value = True
        ws['B1'].value = False
        ws['C1'].value = 5
        ws['D1'].value = 10
        ws['E1'].value = 0
        
        # AND function tests
        ws['A3'].formula = '=AND(TRUE, TRUE)'
        ws['A4'].formula = '=AND(TRUE, FALSE)'
        ws['A5'].formula = '=AND(FALSE, FALSE)'
        ws['A6'].formula = '=AND(A1, B1)'  # Reference to cells
        ws['A7'].formula = '=AND(C1>0, D1>5)'  # Logical expressions
        ws['A8'].formula = '=AND(C1>0, D1>5, E1=0)'  # Multiple conditions
        
        # OR function tests
        ws['B3'].formula = '=OR(TRUE, TRUE)'
        ws['B4'].formula = '=OR(TRUE, FALSE)'
        ws['B5'].formula = '=OR(FALSE, FALSE)'
        ws['B6'].formula = '=OR(A1, B1)'  # Reference to cells
        ws['B7'].formula = '=OR(C1>10, D1>5)'  # Logical expressions
        ws['B8'].formula = '=OR(C1>10, D1>15, E1>0)'  # Multiple conditions
        
        # TRUE and FALSE constants
        ws['C3'].formula = '=TRUE()'
        ws['C4'].formula = '=FALSE()'
        
        # Nested logical functions
        ws['D3'].formula = '=AND(OR(A1, B1), NOT(E1>0))'
        ws['D4'].formula = '=OR(AND(A1, B1), AND(C1>0, D1>0))'
        
        # Edge cases
        ws['E3'].formula = '=AND()'  # Empty AND (should be TRUE)
        ws['E4'].formula = '=OR()'   # Empty OR (should be FALSE)
        
        # Force calculation to ensure all formulas are evaluated
        wb.app.calculate()
        
        # Save the workbook
        wb.save(filepath)
        print(f"✅ Created {filepath} with Excel calculations")
        
    finally:
        # Clean up
        wb.close()
        app.quit()


if __name__ == "__main__":
    output_path = "logical.xlsx"
    create_logical_excel_with_xlwings(output_path)
    print(f"logical.xlsx created successfully at {output_path}")