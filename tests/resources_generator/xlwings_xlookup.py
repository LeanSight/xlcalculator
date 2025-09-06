#!/usr/bin/env python3
"""
Generate XLOOKUP.xlsx using xlwings with Excel calculations.
This ensures Excel calculates the formula values for proper integration testing.
"""

import xlwings as xw
import os


def create_xlookup_excel_with_xlwings(filepath):
    """Create XLOOKUP.xlsx with comprehensive test scenarios using xlwings."""
    
    # Start Excel application
    app = xw.App(visible=False)
    try:
        wb = app.books.add()
        ws = wb.sheets[0]
        ws.name = "Sheet1"
        
        # Test data setup
        ws['A1'].value = 'Fruit'
        ws['B1'].value = 'Price'
        ws['A2'].value = 'Apple'
        ws['B2'].value = 10
        ws['A3'].value = 'Banana'
        ws['B3'].value = 20
        ws['A4'].value = 'Cherry'
        ws['B4'].value = 30
        ws['A5'].value = 'Date'
        ws['B5'].value = 40
        
        # Sorted numbers for approximate match testing
        ws['D1'].value = 'Score'
        ws['E1'].value = 'Grade'
        ws['D2'].value = 10
        ws['E2'].value = 'F'
        ws['D3'].value = 20
        ws['E3'].value = 'D'
        ws['D4'].value = 30
        ws['E4'].value = 'C'
        ws['D5'].value = 40
        ws['E5'].value = 'B'
        ws['D6'].value = 50
        ws['E6'].value = 'A'
        
        # Duplicate values for reverse search testing
        ws['G1'].value = 'Item'
        ws['H1'].value = 'Position'
        ws['G2'].value = 'A'
        ws['H2'].value = 1
        ws['G3'].value = 'B'
        ws['H3'].value = 2
        ws['G4'].value = 'A'
        ws['H4'].value = 3
        ws['G5'].value = 'C'
        ws['H5'].value = 4
        ws['G6'].value = 'A'
        ws['H6'].value = 5
        
        # Basic XLOOKUP tests
        ws['A8'].formula = '=XLOOKUP("Apple", A2:A5, B2:B5)'
        ws['A9'].formula = '=XLOOKUP("Orange", A2:A5, B2:B5, "Not Found")'
        ws['A10'].formula = '=XLOOKUP("Cherry", A2:A5, B2:B5)'
        
        # Approximate match tests
        ws['A12'].formula = '=XLOOKUP(25, D2:D6, E2:E6, , -1)'  # Next smallest
        ws['A13'].formula = '=XLOOKUP(15, D2:D6, E2:E6, , 1)'   # Next largest
        ws['A14'].formula = '=XLOOKUP(30, D2:D6, E2:E6, , 0)'   # Exact match
        
        # Wildcard match tests
        ws['A16'].formula = '=XLOOKUP("App*", A2:A5, B2:B5, , 2)'
        ws['A17'].formula = '=XLOOKUP("Ban?na", A2:A5, B2:B5, , 2)'
        ws['A18'].formula = '=XLOOKUP("*erry", A2:A5, B2:B5, , 2)'
        
        # Reverse search tests
        ws['A20'].formula = '=XLOOKUP("A", G2:G6, H2:H6, , 0, 1)'   # First occurrence
        ws['A21'].formula = '=XLOOKUP("A", G2:G6, H2:H6, , 0, -1)'  # Last occurrence
        
        # Binary search tests (sorted data)
        ws['A23'].formula = '=XLOOKUP(30, D2:D6, E2:E6, , 0, 2)'
        ws['A24'].formula = '=XLOOKUP(20, D2:D6, E2:E6, , 0, 2)'
        
        # Error cases
        ws['A26'].formula = '=XLOOKUP("Grape", A2:A5, B2:B5)'  # Should return #N/A
        
        # Horizontal array test
        ws['A28'].value = 'Apple'
        ws['B28'].value = 'Banana'
        ws['C28'].value = 'Cherry'
        ws['A29'].value = 100
        ws['B29'].value = 200
        ws['C29'].value = 300
        ws['A30'].formula = '=XLOOKUP("Banana", A28:C28, A29:C29)'
        
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
    output_path = "XLOOKUP.xlsx"
    create_xlookup_excel_with_xlwings(output_path)
    print(f"XLOOKUP.xlsx created successfully at {output_path}")