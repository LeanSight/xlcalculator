#!/usr/bin/env python3
"""
Generate DYNAMIC_RANGE.xlsx using xlwings with Excel calculations.
This is a SAFE version that avoids problematic formulas that cause COM automation errors.
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
        
        # === BASIC INDEX FUNCTION TESTS ===
        # Simple INDEX formulas that should work reliably
        ws['G1'].formula = '=INDEX(A1:E5, 2, 2)'  # Alice's age (25)
        ws['G2'].formula = '=INDEX(A1:E5, 3, 1)'  # Bob
        ws['G3'].formula = '=INDEX(A1:E5, 1, 3)'  # City
        ws['G4'].formula = '=INDEX(A1:E5, 4, 4)'  # Charlie's score (78)
        ws['G5'].formula = '=INDEX(A1:E5, 5, 5)'  # Diana's active status (True)
        
        # More INDEX tests with valid ranges
        ws['G7'].formula = '=INDEX(B1:B5, 2)'     # B2 (Alice's age)
        ws['G8'].formula = '=INDEX(A2:E2, 2)'     # B2 (Alice's age)
        
        # === BASIC OFFSET FUNCTION TESTS ===
        # Simple OFFSET formulas
        ws['I1'].formula = '=OFFSET(A1, 1, 1)'     # B2 reference (25)
        ws['I2'].formula = '=OFFSET(B2, 1, 1)'     # C3 reference ("LA")
        ws['I3'].formula = '=OFFSET(A1, 0, 2)'     # C1 reference ("City")
        ws['I4'].formula = '=OFFSET(A1, 2, 3)'     # D3 reference (92)
        
        # === BASIC INDIRECT FUNCTION TESTS ===
        # Reference strings for INDIRECT
        ws['K1'].value = 'B2'
        ws['K2'].value = 'C3'
        ws['K3'].value = 'D4'
        
        # Simple INDIRECT formulas
        ws['M1'].formula = '=INDIRECT(K1)'         # Value at B2 (25)
        ws['M2'].formula = '=INDIRECT(K2)'         # Value at C3 ("LA")
        ws['M3'].formula = '=INDIRECT(K3)'         # Value at D4 (78)
        ws['M4'].formula = '=INDIRECT("B2")'       # 25 (direct reference)
        ws['M5'].formula = '=INDIRECT("C3")'       # "LA" (direct reference)
        
        # === SAFE ERROR HANDLING ===
        # Use simple IF statements instead of IFERROR for better compatibility
        ws['G10'].formula = '=IF(ISERROR(INDEX(A1:E5, 6, 1)), "OUT_OF_BOUNDS", INDEX(A1:E5, 6, 1))'
        ws['G11'].formula = '=IF(ISERROR(INDEX(A1:E5, 1, 6)), "OUT_OF_BOUNDS", INDEX(A1:E5, 1, 6))'
        
        ws['I9'].formula = '=IF(ISERROR(OFFSET(A1, -1, 0)), "OFFSET_ERROR", OFFSET(A1, -1, 0))'
        ws['I10'].formula = '=IF(ISERROR(OFFSET(A1, 0, -1)), "OFFSET_ERROR", OFFSET(A1, 0, -1))'
        
        # Invalid reference for INDIRECT
        ws['K5'].value = 'InvalidRef'
        ws['M10'].formula = '=IF(ISERROR(INDIRECT(K5)), "INVALID_REF", INDIRECT(K5))'
        ws['M11'].formula = '=IF(ISERROR(INDIRECT("")), "EMPTY_REF", INDIRECT(""))'
        
        # === SIMPLE COMBINATIONS ===
        # Avoid complex nested formulas that might cause COM errors
        ws['O1'].formula = '=INDEX(A1:E5, 2, 2)'   # Simple INDEX
        ws['O2'].formula = '=INDIRECT("B2")'       # Simple INDIRECT
        
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