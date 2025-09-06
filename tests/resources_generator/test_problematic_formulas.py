#!/usr/bin/env python3
"""
Test individual formulas to identify which ones cause COM automation errors.
This helps isolate the problematic formulas in DYNAMIC_RANGE.xlsx generation.
"""

import re


def analyze_formula_complexity(formula):
    """Analyze a formula for potential COM automation issues."""
    issues = []
    
    # Remove leading = if present
    if formula.startswith('='):
        formula = formula[1:]
    
    # Check for nested functions (high complexity)
    nesting_level = 0
    max_nesting = 0
    for char in formula:
        if char == '(':
            nesting_level += 1
            max_nesting = max(max_nesting, nesting_level)
        elif char == ')':
            nesting_level -= 1
    
    if max_nesting > 3:
        issues.append(f"High nesting level: {max_nesting}")
    
    # Check for IFERROR usage (can be problematic)
    if 'IFERROR' in formula.upper():
        issues.append("Uses IFERROR function")
    
    # Check for range operations that return arrays
    range_patterns = [
        r'OFFSET\([^,]+,[^,]+,[^,]+,\s*\d+\s*,\s*\d+\s*\)',  # OFFSET with height/width
        r'INDEX\([^,]+,\s*0\s*,',  # INDEX with row 0 (entire column)
        r'INDEX\([^,]+,[^,]+,\s*0\s*\)',  # INDEX with column 0 (entire row)
    ]
    
    for pattern in range_patterns:
        if re.search(pattern, formula):
            issues.append(f"Range-returning formula pattern: {pattern}")
    
    # Check for potentially invalid references
    if 'INDIRECT("")' in formula:
        issues.append("Empty INDIRECT reference")
    
    # Check for negative OFFSET values
    if re.search(r'OFFSET\([^,]+,\s*-\d+', formula):
        issues.append("Negative OFFSET row")
    if re.search(r'OFFSET\([^,]+,[^,]+,\s*-\d+', formula):
        issues.append("Negative OFFSET column")
    
    return issues


def categorize_formulas():
    """Categorize formulas by risk level for COM automation."""
    
    # Extract formulas from the original file
    original_formulas = [
        # INDEX formulas
        '=INDEX(A1:E5, 2, 2)',
        '=INDEX(A1:E5, 3, 1)',
        '=INDEX(A1:E5, 1, 3)',
        '=INDEX(A1:E5, 4, 4)',
        '=INDEX(A1:E5, 5, 5)',
        '=INDEX(B1:B5, 2)',
        '=INDEX(A2:E2, 1, 2)',
        '=IFERROR(INDEX(A1:E5, 6, 1), "OUT_OF_BOUNDS")',
        '=IFERROR(INDEX(A1:E5, 1, 6), "OUT_OF_BOUNDS")',
        
        # OFFSET formulas
        '=OFFSET(A1, 1, 1)',
        '=OFFSET(B2, 1, 1)',
        '=OFFSET(A1, 0, 2)',
        '=OFFSET(A1, 2, 3)',
        '=SUM(OFFSET(A1, 1, 1, 2, 2))',
        '=COUNT(OFFSET(A1, 0, 0, 3, 3))',
        '=IFERROR(OFFSET(A1, -1, 0), "OFFSET_ERROR")',
        '=IFERROR(OFFSET(A1, 0, -1), "OFFSET_ERROR")',
        
        # INDIRECT formulas
        '=INDIRECT(K1)',
        '=INDIRECT(K2)',
        '=INDIRECT(K3)',
        '=INDIRECT("B2")',
        '=INDIRECT("C3")',
        '=SUM(INDIRECT(K4))',
        '=COUNT(INDIRECT("A1:B2"))',
        '=IFERROR(INDIRECT(K5), "INVALID_REF")',
        '=IFERROR(INDIRECT(""), "EMPTY_REF")',
        
        # Complex combinations
        '=INDEX(INDIRECT("A1:E5"), 2, 2)',
        '=IFERROR(INDIRECT(OFFSET(K1, 1, 0)), "COMPLEX_ERROR")',
    ]
    
    print("🔍 Formula Risk Analysis for COM Automation")
    print("=" * 60)
    
    low_risk = []
    medium_risk = []
    high_risk = []
    
    for formula in original_formulas:
        issues = analyze_formula_complexity(formula)
        
        if not issues:
            risk_level = "LOW"
            low_risk.append(formula)
        elif len(issues) == 1 and "Uses IFERROR function" in issues:
            risk_level = "MEDIUM"
            medium_risk.append(formula)
        else:
            risk_level = "HIGH"
            high_risk.append(formula)
        
        print(f"\n{risk_level:6} | {formula}")
        for issue in issues:
            print(f"       | - {issue}")
    
    print("\n" + "=" * 60)
    print("📊 SUMMARY")
    print("=" * 60)
    print(f"✅ LOW RISK:    {len(low_risk)} formulas")
    print(f"⚠️  MEDIUM RISK: {len(medium_risk)} formulas")
    print(f"❌ HIGH RISK:   {len(high_risk)} formulas")
    
    print("\n🔧 RECOMMENDATIONS:")
    print("1. Start with LOW RISK formulas only")
    print("2. Replace IFERROR with IF(ISERROR(...)) for better compatibility")
    print("3. Avoid range-returning formulas (use SUM/COUNT wrappers)")
    print("4. Simplify complex nested formulas")
    print("5. Test incrementally, adding formulas one by one")
    
    return low_risk, medium_risk, high_risk


if __name__ == "__main__":
    low_risk, medium_risk, high_risk = categorize_formulas()
    
    print(f"\n📝 SAFE FORMULAS TO START WITH ({len(low_risk)} formulas):")
    for formula in low_risk:
        print(f"   {formula}")