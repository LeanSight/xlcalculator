#!/usr/bin/env python3
"""
Verify that the restored DYNAMIC_RANGE.xlsx formulas match the integration test expectations.
This cross-references the generated formulas with the test requirements.
"""

import re
from pathlib import Path


def extract_test_expectations():
    """Extract expected formulas and behaviors from the test file."""
    test_expectations = {
        # INDEX tests
        'G1': {'formula': 'INDEX(A1:E5, 2, 2)', 'expected': 25, 'description': "Alice's age"},
        'G2': {'formula': 'INDEX(A1:E5, 3, 1)', 'expected': "Bob", 'description': "Bob's name"},
        'G3': {'formula': 'INDEX(A1:E5, 1, 3)', 'expected': "City", 'description': "City header"},
        'G4': {'formula': 'INDEX(A1:E5, 4, 4)', 'expected': 78, 'description': "Charlie's score"},
        'G5': {'formula': 'INDEX(A1:E5, 5, 5)', 'expected': True, 'description': "Diana's active status"},
        'G7': {'formula': 'INDEX(A1:E5, 0, 2)', 'expected': 'Array', 'description': "Entire column 2 - CRITICAL"},
        'G8': {'formula': 'INDEX(A1:E5, 2, 0)', 'expected': 'Array', 'description': "Entire row 2 - CRITICAL"},
        'G10': {'formula': 'INDEX(A1:E5, 6, 1)', 'expected': '#REF!', 'description': "Row out of bounds - CRITICAL"},
        'G11': {'formula': 'INDEX(A1:E5, 1, 6)', 'expected': '#REF!', 'description': "Column out of bounds - CRITICAL"},
        
        # OFFSET tests
        'I1': {'formula': 'OFFSET(A1, 1, 1)', 'expected': 25, 'description': "B2 reference"},
        'I2': {'formula': 'OFFSET(B2, 1, 1)', 'expected': "LA", 'description': "C3 reference"},
        'I3': {'formula': 'OFFSET(A1, 0, 2)', 'expected': "City", 'description': "C1 reference"},
        'I4': {'formula': 'OFFSET(A1, 2, 3)', 'expected': 92, 'description': "D3 reference"},
        'I6': {'formula': 'OFFSET(A1, 1, 1, 2, 2)', 'expected': 'ValueExcelError', 'description': "Range B2:C3 - CRITICAL"},
        'I7': {'formula': 'OFFSET(A1, 0, 0, 3, 3)', 'expected': 'ValueExcelError', 'description': "Range A1:C3 - CRITICAL"},
        'I9': {'formula': 'OFFSET(A1, -1, 0)', 'expected': 'ValueExcelError', 'description': "Negative row - CRITICAL"},
        'I10': {'formula': 'OFFSET(A1, 0, -1)', 'expected': 'ValueExcelError', 'description': "Negative column - CRITICAL"},
        
        # INDIRECT tests
        'M1': {'formula': 'INDIRECT(K1)', 'expected': 25, 'description': "Value at B2"},
        'M2': {'formula': 'INDIRECT(K2)', 'expected': "LA", 'description': "Value at C3"},
        'M3': {'formula': 'INDIRECT(K3)', 'expected': 78, 'description': "Value at D4"},
        'M4': {'formula': 'INDIRECT("B2")', 'expected': 25, 'description': "Direct reference"},
        'M5': {'formula': 'INDIRECT("C3")', 'expected': "LA", 'description': "Direct reference"},
        'M7': {'formula': 'INDIRECT(K4)', 'expected': "A1:C3", 'description': "Range reference - CRITICAL"},
        'M8': {'formula': 'INDIRECT("A1:B2")', 'expected': "A1:B2", 'description': "Direct range - CRITICAL"},
        'M10': {'formula': 'INDIRECT(K5)', 'expected': '#NAME!', 'description': "Invalid reference - CRITICAL"},
        'M11': {'formula': 'INDIRECT("")', 'expected': '#NAME!', 'description': "Empty reference - CRITICAL"},
        
        # Complex combinations
        'O1': {'formula': 'INDEX(INDIRECT("A1:E5"), 2, 2)', 'expected': 25, 'description': "Nested INDEX/INDIRECT"},
        'O2': {'formula': 'INDIRECT(OFFSET("K1", 1, 0))', 'expected': "C3", 'description': "Nested INDIRECT/OFFSET - CRITICAL"},
    }
    return test_expectations


def extract_generated_formulas():
    """Extract formulas from the generated script."""
    script_path = Path("xlwings_dynamic_range.py")
    formulas = {}
    
    with open(script_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Extract formulas from the formulas_to_add list
    formula_pattern = r"\('([^']+)',\s*'([^']+)'\)"
    matches = re.findall(formula_pattern, content)
    
    for cell, formula in matches:
        # Remove leading = if present
        if formula.startswith('='):
            formula = formula[1:]
        formulas[cell] = formula
    
    return formulas


def compare_formulas():
    """Compare generated formulas with test expectations."""
    test_expectations = extract_test_expectations()
    generated_formulas = extract_generated_formulas()
    
    print("🔍 DYNAMIC_RANGE Test Compatibility Verification")
    print("=" * 70)
    
    matches = 0
    mismatches = 0
    missing = 0
    
    for cell, expected in test_expectations.items():
        if cell not in generated_formulas:
            print(f"❌ {cell}: MISSING - Expected: {expected['formula']}")
            missing += 1
            continue
        
        generated = generated_formulas[cell]
        expected_formula = expected['formula']
        
        # Check if the core formula logic matches (ignoring error handling wrappers)
        if expected_formula in generated or core_formula_matches(expected_formula, generated):
            print(f"✅ {cell}: MATCH - {expected['description']}")
            print(f"   Expected: {expected_formula}")
            print(f"   Generated: {generated}")
            matches += 1
        else:
            print(f"❌ {cell}: MISMATCH - {expected['description']}")
            print(f"   Expected: {expected_formula}")
            print(f"   Generated: {generated}")
            mismatches += 1
        print()
    
    # Summary
    total = len(test_expectations)
    print("=" * 70)
    print("📊 COMPATIBILITY SUMMARY")
    print("=" * 70)
    print(f"✅ Matches:    {matches}/{total} ({matches/total*100:.1f}%)")
    print(f"❌ Mismatches: {mismatches}/{total} ({mismatches/total*100:.1f}%)")
    print(f"⚠️  Missing:    {missing}/{total} ({missing/total*100:.1f}%)")
    
    compatibility_score = matches / total * 100
    print(f"🎯 Compatibility Score: {compatibility_score:.1f}%")
    
    if compatibility_score >= 90:
        print("\n🎉 EXCELLENT: High compatibility with integration tests!")
    elif compatibility_score >= 75:
        print("\n✅ GOOD: Acceptable compatibility with integration tests")
    else:
        print("\n⚠️  POOR: Low compatibility - many tests may fail")
    
    return matches, mismatches, missing


def core_formula_matches(expected, generated):
    """Check if the core formula logic matches, ignoring error handling wrappers."""
    # Remove IF(ISERROR(...)) wrappers to check core formula
    if 'IF(ISERROR(' in generated:
        # Extract the core formula from IF(ISERROR(FORMULA, ERROR_VALUE, FORMULA))
        pattern = r'IF\(ISERROR\(([^,]+)\),[^,]+,\1\)'
        match = re.search(pattern, generated)
        if match:
            core_formula = match.group(1)
            return expected == core_formula
    
    return False


if __name__ == "__main__":
    matches, mismatches, missing = compare_formulas()
    
    if mismatches > 0 or missing > 0:
        print(f"\n🔧 ACTION REQUIRED:")
        print(f"   - Fix {mismatches} mismatched formulas")
        print(f"   - Add {missing} missing formulas")
        print(f"   - Ensure error handling preserves core Excel behavior")
    else:
        print(f"\n🎉 ALL FORMULAS MATCH TEST EXPECTATIONS!")