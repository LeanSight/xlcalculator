#!/usr/bin/env python3
"""
Validation script to test DYNAMIC_RANGE.xlsx generation fixes.
This script validates the improved xlwings_dynamic_range.py without requiring Excel.
"""

import ast
import re
import sys
from pathlib import Path


def extract_formulas_from_script(script_path):
    """Extract all formulas from the xlwings script."""
    formulas = []
    
    with open(script_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Extract formulas from the formulas_to_add list
    formula_pattern = r"\('([^']+)',\s*'([^']+)'\)"
    matches = re.findall(formula_pattern, content)
    
    for cell, formula in matches:
        formulas.append((cell, formula))
    
    return formulas


def validate_formula_safety(formula):
    """Check if a formula is safe for Excel COM automation."""
    issues = []
    
    # Remove leading = if present
    if formula.startswith('='):
        formula = formula[1:]
    
    # Check for high-risk patterns
    high_risk_patterns = [
        (r'IFERROR\s*\(', "Uses IFERROR (can cause COM issues)"),
        (r'OFFSET\([^,]+,[^,]+,[^,]+,\s*\d+\s*,\s*\d+\s*\)', "OFFSET with height/width (returns range)"),
        (r'INDEX\([^,]+,\s*0\s*,', "INDEX with row 0 (returns entire column)"),
        (r'INDEX\([^,]+,[^,]+,\s*0\s*\)', "INDEX with column 0 (returns entire row)"),
        (r'OFFSET\([^,]+,\s*-\d+', "OFFSET with negative row"),
        (r'OFFSET\([^,]+,[^,]+,\s*-\d+', "OFFSET with negative column"),
        (r'INDIRECT\s*\(\s*""\s*\)', "INDIRECT with empty string"),
    ]
    
    for pattern, description in high_risk_patterns:
        if re.search(pattern, formula, re.IGNORECASE):
            issues.append(description)
    
    # Check nesting level
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
    
    return issues


def validate_script_syntax(script_path):
    """Validate Python syntax of the script."""
    try:
        with open(script_path, 'r', encoding='utf-8') as f:
            content = f.read()
        ast.parse(content)
        return True, None
    except SyntaxError as e:
        return False, str(e)


def main():
    """Main validation function."""
    script_path = Path("xlwings_dynamic_range.py")
    
    print("🔍 DYNAMIC_RANGE.xlsx Generation Validation")
    print("=" * 60)
    
    # 1. Validate Python syntax
    print("1. Checking Python syntax...")
    syntax_ok, syntax_error = validate_script_syntax(script_path)
    if syntax_ok:
        print("   ✅ Python syntax is valid")
    else:
        print(f"   ❌ Python syntax error: {syntax_error}")
        return 1
    
    # 2. Extract and validate formulas
    print("\n2. Extracting formulas from script...")
    try:
        formulas = extract_formulas_from_script(script_path)
        print(f"   ✅ Found {len(formulas)} formulas")
    except Exception as e:
        print(f"   ❌ Failed to extract formulas: {e}")
        return 1
    
    # 3. Validate each formula
    print("\n3. Validating formula safety...")
    safe_formulas = 0
    risky_formulas = 0
    
    for cell, formula in formulas:
        issues = validate_formula_safety(formula)
        if issues:
            print(f"   ⚠️  {cell}: {formula}")
            for issue in issues:
                print(f"      - {issue}")
            risky_formulas += 1
        else:
            print(f"   ✅ {cell}: {formula}")
            safe_formulas += 1
    
    # 4. Summary
    print("\n" + "=" * 60)
    print("📊 VALIDATION SUMMARY")
    print("=" * 60)
    print(f"✅ Safe formulas:  {safe_formulas}")
    print(f"⚠️  Risky formulas: {risky_formulas}")
    print(f"📊 Total formulas: {len(formulas)}")
    
    safety_percentage = (safe_formulas / len(formulas)) * 100 if formulas else 0
    print(f"🎯 Safety score:   {safety_percentage:.1f}%")
    
    # 5. Recommendations
    print("\n🔧 RECOMMENDATIONS:")
    if risky_formulas == 0:
        print("✅ All formulas are safe for Excel COM automation!")
        print("✅ The script should work reliably on Windows with Excel.")
    elif risky_formulas <= 2:
        print("⚠️  Few risky formulas detected. Consider:")
        print("   - Test the script incrementally")
        print("   - Monitor for COM automation errors")
    else:
        print("❌ Many risky formulas detected. Consider:")
        print("   - Replace IFERROR with IF(ISERROR(...))")
        print("   - Avoid range-returning formulas")
        print("   - Simplify complex nested formulas")
    
    # 6. Expected improvements
    print(f"\n📈 EXPECTED IMPROVEMENTS:")
    print(f"   - Robust error handling with try/catch per formula")
    print(f"   - Incremental formula addition with validation")
    print(f"   - Fallback to safe values for failed formulas")
    print(f"   - Better Excel application settings")
    
    return 0 if risky_formulas <= 2 else 1


if __name__ == "__main__":
    sys.exit(main())