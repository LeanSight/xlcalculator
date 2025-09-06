#!/usr/bin/env python3
"""
Generate MATH.xlsx using xlwings with Excel calculations.
This ensures Excel calculates the formula values for proper integration testing.
"""

import xlwings as xw
import os


def create_math_excel_with_xlwings(filepath):
    """Create MATH.xlsx with mathematical function tests using xlwings."""
    
    # Start Excel application
    app = xw.App(visible=False)
    try:
        wb = app.books.add()
        ws = wb.sheets[0]
        ws.name = "Sheet1"
        
        # Test data
        ws['A1'].value = 3.7
        ws['A2'].value = -2.3
        ws['A3'].value = 0
        ws['A4'].value = 100
        ws['A5'].value = 2.71828  # Approximately e
        
        # FLOOR tests
        ws['B1'].formula = '=FLOOR(A1, 1)'     # 3.7 -> 3
        ws['B2'].formula = '=FLOOR(A2, 1)'     # -2.3 -> -3
        ws['B3'].formula = '=FLOOR(A1, 0.5)'   # 3.7 -> 3.5
        ws['B4'].formula = '=FLOOR(A4, 10)'    # 100 -> 100
        
        # TRUNC tests
        ws['C1'].formula = '=TRUNC(A1)'        # 3.7 -> 3
        ws['C2'].formula = '=TRUNC(A2)'        # -2.3 -> -2
        ws['C3'].formula = '=TRUNC(A1, 1)'     # 3.7 -> 3.7
        ws['C4'].formula = '=TRUNC(A4, -1)'    # 100 -> 100
        
        # SIGN tests
        ws['D1'].formula = '=SIGN(A1)'         # Positive -> 1
        ws['D2'].formula = '=SIGN(A2)'         # Negative -> -1
        ws['D3'].formula = '=SIGN(A3)'         # Zero -> 0
        
        # LOG tests
        ws['E1'].formula = '=LOG(A4)'          # LOG base 10
        ws['E2'].formula = '=LOG(A4, 2)'       # LOG base 2
        ws['E3'].formula = '=LOG(A5, EXP(1))'  # Natural log
        
        # LOG10 tests
        ws['F1'].formula = '=LOG10(A4)'        # LOG10(100) = 2
        ws['F2'].formula = '=LOG10(1000)'      # LOG10(1000) = 3
        
        # EXP tests
        ws['G1'].formula = '=EXP(0)'           # e^0 = 1
        ws['G2'].formula = '=EXP(1)'           # e^1 = e
        ws['G3'].formula = '=EXP(2)'           # e^2
        
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
    output_path = "MATH.xlsx"
    create_math_excel_with_xlwings(output_path)
    print(f"MATH.xlsx created successfully at {output_path}")