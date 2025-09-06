#!/usr/bin/env python3
"""
Generate TEXT.xlsx using xlwings with Excel calculations.
This ensures Excel calculates the formula values for proper integration testing.
"""

import xlwings as xw
import os


def create_text_excel_with_xlwings(filepath):
    """Create TEXT.xlsx with text function tests using xlwings."""
    
    # Start Excel application
    app = xw.App(visible=False)
    try:
        wb = app.books.add()
        ws = wb.sheets[0]
        ws.name = "Sheet1"
        
        # Test data
        ws['A1'].value = "Hello World"
        ws['A2'].value = "  Excel Functions  "
        ws['A3'].value = "UPPERCASE"
        ws['A4'].value = "lowercase"
        ws['A5'].value = "MiXeD cAsE"
        ws['A6'].value = "Replace This Text"
        
        # LEFT tests
        ws['B1'].formula = '=LEFT(A1, 5)'      # "Hello"
        ws['B2'].formula = '=LEFT(A1, 1)'      # "H"
        ws['B3'].formula = '=LEFT(A1)'         # Default 1 char
        
        # UPPER tests
        ws['C3'].formula = '=UPPER(A3)'        # Already uppercase
        ws['C4'].formula = '=UPPER(A4)'        # Convert lowercase
        ws['C5'].formula = '=UPPER(A5)'        # Convert mixed case
        
        # LOWER tests
        ws['D3'].formula = '=LOWER(A3)'        # Convert uppercase
        ws['D4'].formula = '=LOWER(A4)'        # Already lowercase
        ws['D5'].formula = '=LOWER(A5)'        # Convert mixed case
        
        # TRIM tests
        ws['E2'].formula = '=TRIM(A2)'         # Remove leading/trailing spaces
        ws['E1'].formula = '=TRIM(A1)'         # No extra spaces
        
        # REPLACE tests
        ws['F6'].formula = '=REPLACE(A6, 9, 4, "That")'  # Replace "This" with "That"
        ws['F1'].formula = '=REPLACE(A1, 1, 5, "Hi")'    # Replace "Hello" with "Hi"
        
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
    output_path = "TEXT.xlsx"
    create_text_excel_with_xlwings(output_path)
    print(f"TEXT.xlsx created successfully at {output_path}")