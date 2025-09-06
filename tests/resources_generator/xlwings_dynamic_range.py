#!/usr/bin/env python3
"""
Generate DYNAMIC_RANGE.xlsx using xlwings with Excel calculations.
This ensures Excel calculates the formula values for proper integration testing.
"""

import xlwings as xw
import os


def create_dynamic_range_excel_with_xlwings(filepath):
    """Create DYNAMIC_RANGE.xlsx with INDEX, OFFSET, and INDIRECT test scenarios using xlwings."""
    
    # Start Excel application
    app = xw.App(visible=False)
    try:
        wb = app.books.add()
        ws = wb.sheets[0]
        ws.name = "Sheet1"
        
        # Test data grid (A1:E5) for INDEX and OFFSET tests
        ws['A1'].value = 'Name'
        ws['B1'].value = 'Age'
        ws['C1'].value = 'City'
        ws['D1'].value = 'Score'
        ws['E1'].value = 'Active'
        
        ws['A2'].value = 'Alice'
        ws['B2'].value = 25
        ws['C2'].value = 'NYC'
        ws['D2'].value = 85
        ws['E2'].value = True
        
        ws['A3'].value = 'Bob'
        ws['B3'].value = 30
        ws['C3'].value = 'LA'
        ws['D3'].value = 92
        ws['E3'].value = False
        
        ws['A4'].value = 'Charlie'
        ws['B4'].value = 35
        ws['C4'].value = 'Chicago'
        ws['D4'].value = 78
        ws['E4'].value = True
        
        ws['A5'].value = 'Diana'
        ws['B5'].value = 28
        ws['C5'].value = 'Miami'
        ws['D5'].value = 95
        ws['E5'].value = True
        
        # INDEX function tests
        ws['G1'].formula = '=INDEX(A1:E5, 2, 2)'  # Alice's age (25)
        ws['G2'].formula = '=INDEX(A1:E5, 3, 1)'  # Bob
        ws['G3'].formula = '=INDEX(A1:E5, 1, 3)'  # City
        ws['G4'].formula = '=INDEX(A1:E5, 4, 4)'  # Charlie's score (78)
        ws['G5'].formula = '=INDEX(A1:E5, 5, 5)'  # Diana's active status (True)
        
        # INDEX with entire row/column
        ws['G7'].formula = '=INDEX(A1:E5, 0, 2)'  # Entire column B (ages)
        ws['G8'].formula = '=INDEX(A1:E5, 2, 0)'  # Entire row 2 (Alice's data)
        
        # INDEX error cases
        ws['G10'].formula = '=INDEX(A1:E5, 6, 1)'  # #REF! (row out of bounds)
        ws['G11'].formula = '=INDEX(A1:E5, 1, 6)'  # #REF! (column out of bounds)
        
        # OFFSET function tests
        ws['I1'].formula = '=OFFSET(A1, 1, 1)'     # B2 reference (25)
        ws['I2'].formula = '=OFFSET(B2, 1, 1)'     # C3 reference ("LA")
        ws['I3'].formula = '=OFFSET(A1, 0, 2)'     # C1 reference ("City")
        ws['I4'].formula = '=OFFSET(A1, 2, 3)'     # D3 reference (92)
        
        # OFFSET with height and width
        ws['I6'].formula = '=OFFSET(A1, 1, 1, 2, 2)'  # B2:C3 range
        ws['I7'].formula = '=OFFSET(A1, 0, 0, 3, 3)'  # A1:C3 range
        
        # OFFSET error cases
        ws['I9'].formula = '=OFFSET(A1, -1, 0)'    # #REF! (out of bounds)
        ws['I10'].formula = '=OFFSET(A1, 0, -1)'   # #REF! (out of bounds)
        
        # INDIRECT function tests - reference strings
        ws['K1'].value = 'B2'
        ws['K2'].value = 'C3'
        ws['K3'].value = 'D4'
        ws['K4'].value = 'A1:C3'
        ws['K5'].value = 'InvalidRef'
        
        # INDIRECT formulas
        ws['M1'].formula = '=INDIRECT(K1)'         # Value at B2 (25)
        ws['M2'].formula = '=INDIRECT(K2)'         # Value at C3 ("LA")
        ws['M3'].formula = '=INDIRECT(K3)'         # Value at D4 (78)
        ws['M4'].formula = '=INDIRECT("B2")'       # 25 (direct reference)
        ws['M5'].formula = '=INDIRECT("C3")'       # "LA" (direct reference)
        
        # INDIRECT with ranges
        ws['M7'].formula = '=INDIRECT(K4)'         # A1:C3 range
        ws['M8'].formula = '=INDIRECT("A1:B2")'    # A1:B2 range
        
        # INDIRECT error cases
        ws['M10'].formula = '=INDIRECT(K5)'        # #NAME! (invalid reference)
        ws['M11'].formula = '=INDIRECT("")'        # #NAME! (empty reference)
        
        # Complex combinations
        ws['O1'].formula = '=INDEX(INDIRECT("A1:E5"), 2, 2)'  # Nested: INDEX with INDIRECT
        ws['O2'].formula = '=INDIRECT(OFFSET("K1", 1, 0))'    # Nested: INDIRECT with OFFSET result
        
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
    output_path = "DYNAMIC_RANGE.xlsx"
    create_dynamic_range_excel_with_xlwings(output_path)
    print(f"DYNAMIC_RANGE.xlsx created successfully at {output_path}")