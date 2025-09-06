#!/usr/bin/env python3
"""
Generate INFORMATION.xlsx using xlwings with Excel calculations.
This ensures Excel calculates the formula values for proper integration testing.
"""

import xlwings as xw
import os


def create_information_excel_with_xlwings(filepath):
    """Create INFORMATION.xlsx with IS* function tests using xlwings."""
    
    # Start Excel application
    app = xw.App(visible=False)
    try:
        wb = app.books.add()
        ws = wb.sheets[0]
        ws.name = "Sheet1"
        
        # Test data with various types
        ws['A1'].value = 123        # Number
        ws['A2'].value = "Hello"    # Text
        ws['A3'].value = True       # Boolean
        ws['A4'].value = None       # Empty/Blank
        ws['A5'].formula = '=1/0'   # Error formula
        ws['A6'].formula = '=NA()'  # #N/A error
        ws['A7'].value = 4          # Even number
        ws['A8'].value = 5          # Odd number
        ws['A9'].value = 0          # Zero
        ws['A10'].value = -3        # Negative odd
        
        # ISNUMBER tests
        ws['B1'].formula = '=ISNUMBER(A1)'   # Number
        ws['B2'].formula = '=ISNUMBER(A2)'   # Text
        ws['B3'].formula = '=ISNUMBER(A3)'   # Boolean
        ws['B4'].formula = '=ISNUMBER(A4)'   # Blank
        ws['B5'].formula = '=ISNUMBER(123)'  # Direct number
        
        # ISTEXT tests
        ws['C1'].formula = '=ISTEXT(A1)'     # Number
        ws['C2'].formula = '=ISTEXT(A2)'     # Text
        ws['C3'].formula = '=ISTEXT(A3)'     # Boolean
        ws['C4'].formula = '=ISTEXT(A4)'     # Blank
        ws['C5'].formula = '=ISTEXT("test")' # Direct text
        
        # ISBLANK tests
        ws['D1'].formula = '=ISBLANK(A1)'    # Number
        ws['D2'].formula = '=ISBLANK(A2)'    # Text
        ws['D4'].formula = '=ISBLANK(A4)'    # Blank
        ws['D5'].formula = '=ISBLANK("")'    # Empty string
        
        # ISERROR tests
        ws['E1'].formula = '=ISERROR(A1)'    # Number
        ws['E5'].formula = '=ISERROR(A5)'    # Error
        ws['E6'].formula = '=ISERROR(A6)'    # #N/A
        ws['E7'].formula = '=ISERROR(1/0)'   # Direct error
        
        # ISNA tests
        ws['F6'].formula = '=ISNA(A6)'       # #N/A error
        ws['F5'].formula = '=ISNA(A5)'       # Other error
        ws['F1'].formula = '=ISNA(A1)'       # Number
        
        # ISERR tests (errors except #N/A)
        ws['G5'].formula = '=ISERR(A5)'      # #DIV/0! error
        ws['G6'].formula = '=ISERR(A6)'      # #N/A error
        ws['G1'].formula = '=ISERR(A1)'      # Number
        
        # ISEVEN tests
        ws['H7'].formula = '=ISEVEN(A7)'     # Even number
        ws['H8'].formula = '=ISEVEN(A8)'     # Odd number
        ws['H9'].formula = '=ISEVEN(A9)'     # Zero
        ws['H10'].formula = '=ISEVEN(A10)'   # Negative odd
        
        # ISODD tests
        ws['I7'].formula = '=ISODD(A7)'      # Even number
        ws['I8'].formula = '=ISODD(A8)'      # Odd number
        ws['I9'].formula = '=ISODD(A9)'      # Zero
        ws['I10'].formula = '=ISODD(A10)'    # Negative odd
        
        # NA function test
        ws['J1'].formula = '=NA()'           # Generate #N/A error
        
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
    output_path = "INFORMATION.xlsx"
    create_information_excel_with_xlwings(output_path)
    print(f"INFORMATION.xlsx created successfully at {output_path}")