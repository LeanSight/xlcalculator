#!/usr/bin/env python3
"""
Generate DYNAMIC_RANGE.xlsx using xlwings with Excel calculations.
This ensures Excel calculates the formula values for proper integration testing.
"""

import xlwings as xw
import os


def create_dynamic_range_excel_with_xlwings(filepath):
    """Create DYNAMIC_RANGE.xlsx with INDEX, OFFSET, and INDIRECT test scenarios using xlwings."""
    
    # Start Excel application with more robust settings
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    
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
        
        # INDIRECT function tests - reference strings (set up data for INDIRECT)
        ws['K1'].value = 'B2'
        ws['K2'].value = 'C3'
        ws['K3'].value = 'D4'
        ws['K4'].value = 'A1:C3'
        ws['K5'].value = 'InvalidRef'  # Invalid reference for error testing

        
        # Add formulas incrementally with error checking
        formulas_to_add = [
            # Basic INDEX formulas
            ('G1', '=INDEX(A1:E5, 2, 2)'),
            ('G2', '=INDEX(A1:E5, 3, 1)'),
            ('G3', '=INDEX(A1:E5, 1, 3)'),
            ('G4', '=INDEX(A1:E5, 4, 4)'),
            ('G5', '=INDEX(A1:E5, 5, 5)'),
            ('G7', '=INDEX(A1:E5, 0, 2)'),  # Entire column 2 (ages) - CRITICAL for tests
            ('G8', '=INDEX(A1:E5, 2, 0)'),  # Entire row 2 (Alice's data) - CRITICAL for tests
            ('G10', '=INDEX(A1:E5, 6, 1)'),  # Row out of bounds - CRITICAL for error testing
            ('G11', '=INDEX(A1:E5, 1, 6)'),  # Column out of bounds - CRITICAL for error testing
            
            # Basic OFFSET formulas
            ('I1', '=OFFSET(A1, 1, 1)'),
            ('I2', '=OFFSET(B2, 1, 1)'),
            ('I3', '=OFFSET(A1, 0, 2)'),
            ('I4', '=OFFSET(A1, 2, 3)'),
            ('I6', '=OFFSET(A1, 1, 1, 2, 2)'),  # B2:C3 range - CRITICAL for range testing
            ('I7', '=OFFSET(A1, 0, 0, 3, 3)'),  # A1:C3 range - CRITICAL for range testing
            ('I9', '=OFFSET(A1, -1, 0)'),   # Negative row - CRITICAL for error testing
            ('I10', '=OFFSET(A1, 0, -1)'),  # Negative column - CRITICAL for error testing
            
            # Basic INDIRECT formulas
            ('M1', '=INDIRECT(K1)'),
            ('M2', '=INDIRECT(K2)'),
            ('M3', '=INDIRECT(K3)'),
            ('M4', '=INDIRECT("B2")'),
            ('M5', '=INDIRECT("C3")'),
            ('M7', '=INDIRECT(K4)'),         # Range reference - CRITICAL for range testing
            ('M8', '=INDIRECT("A1:B2")'),    # Direct range reference - CRITICAL for range testing
            ('M10', '=INDIRECT(K5)'),        # Invalid reference - CRITICAL for error testing
            ('M11', '=INDIRECT("")'),          # Empty reference - CRITICAL for error testing
            
            # Simple combinations
            ('O1', '=INDEX(INDIRECT("A1:E5"), 2, 2)'),
            ('O2', '=INDIRECT("K2")'),  # Simplified: Direct reference instead of OFFSET("K1", 1, 0) - tests INDIRECT functionality
            ('O3', '=OFFSET(K1, 1, 0)'),  # Separate OFFSET test - tests OFFSET with cell reference
        ]
        
        # Add formulas one by one - fail fast on any error
        print(f"📝 Adding {len(formulas_to_add)} formulas to Excel...")
        for i, (cell, formula) in enumerate(formulas_to_add, 1):
            try:
                print(f"   {i:2d}/{len(formulas_to_add)}: {cell} = {formula}")
                ws[cell].formula = formula
                # Test calculation immediately to catch formula errors early
                calculated_value = ws[cell].value
                print(f"       ✅ Calculated: {calculated_value}")
            except Exception as e:
                print(f"       ❌ FAILED: {e}")
                print(f"\n❌ GENERATION FAILED at formula {i}/{len(formulas_to_add)}")
                print(f"   Cell: {cell}")
                print(f"   Formula: {formula}")
                print(f"   Error: {e}")
                print(f"\nThis formula is not compatible with Excel COM automation.")
                print(f"The formula must be fixed or simplified, not worked around.")
                raise Exception(f"Excel formula generation failed for {cell}: {formula}")
        
        print(f"✅ All formulas successfully added and calculated")
        
        # Force calculation to ensure all formulas are evaluated
        try:
            wb.app.calculate()
        except Exception as e:
            print(f"⚠️  Calculation warning: {e}")
        
        # Save the workbook
        wb.save(filepath)
        print(f"✅ Created {filepath} with Excel calculations")
        print(f"✅ All {len(formulas_to_add)} formulas successfully added and calculated by Excel")
        
    except Exception as e:
        print(f"❌ Failed to create {filepath}: {e}")
        raise
    finally:
        # Clean up with error handling
        try:
            if 'wb' in locals():
                wb.close()
        except:
            pass
        try:
            app.quit()
        except:
            pass


if __name__ == "__main__":
    output_path = "DYNAMIC_RANGE.xlsx"
    create_dynamic_range_excel_with_xlwings(output_path)
    print(f"DYNAMIC_RANGE.xlsx created successfully at {output_path}")