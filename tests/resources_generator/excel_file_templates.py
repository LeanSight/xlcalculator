#!/usr/bin/env python3
"""
DEPRECATED: Excel File Template Generator for Integration Tests

⚠️  THIS FILE IS DEPRECATED ⚠️

This file has been replaced by xlwings-based generators that create Excel files
with real Excel calculations. The openpyxl-based approach in this file creates
formulas but no calculated values, which defeats the purpose of Excel compatibility testing.

REPLACEMENT FILES:
- xlwings_information.py    (replaces create_information_excel)
- xlwings_logical.py        (replaces create_logical_excel)
- xlwings_math.py          (replaces create_math_excel)
- xlwings_text.py          (replaces create_text_excel)
- xlwings_xlookup.py       (replaces create_xlookup_excel)
- xlwings_dynamic_range.py (replaces create_dynamic_range_excel)

USE INSTEAD:
- generate_all_xlwings.py  (generates all Excel files with real calculations)

WHY DEPRECATED:
The goal of integration tests is to compare xlcalculator against Excel's actual
behavior. This file creates formulas without calculated values, so tests compare
against None instead of Excel's real results.

The xlwings generators use Excel's calculation engine to create files with both
formulas AND Excel's calculated values, ensuring proper compatibility testing.

MIGRATION:
1. Use xlwings generators on Windows with Excel installed
2. Copy generated Excel files to tests/resources/
3. Run integration tests to verify xlcalculator matches Excel behavior

This file is kept for reference only and should not be used for new Excel file generation.
"""

# Original implementation preserved for reference but marked as deprecated
import openpyxl
from openpyxl import Workbook
import os

def create_xlookup_excel():
    """DEPRECATED: Use xlwings_xlookup.py instead."""
    raise DeprecationWarning(
        "create_xlookup_excel() is deprecated. Use xlwings_xlookup.py instead. "
        "This function creates formulas without calculated values, which defeats "
        "the purpose of Excel compatibility testing."
    )

def create_logical_excel():
    """DEPRECATED: Use xlwings_logical.py instead."""
    raise DeprecationWarning(
        "create_logical_excel() is deprecated. Use xlwings_logical.py instead. "
        "This function creates formulas without calculated values, which defeats "
        "the purpose of Excel compatibility testing."
    )

def create_information_excel():
    """DEPRECATED: Use xlwings_information.py instead."""
    raise DeprecationWarning(
        "create_information_excel() is deprecated. Use xlwings_information.py instead. "
        "This function creates formulas without calculated values, which defeats "
        "the purpose of Excel compatibility testing."
    )

def create_math_excel():
    """DEPRECATED: Use xlwings_math.py instead."""
    raise DeprecationWarning(
        "create_math_excel() is deprecated. Use xlwings_math.py instead. "
        "This function creates formulas without calculated values, which defeats "
        "the purpose of Excel compatibility testing."
    )

def create_text_excel():
    """DEPRECATED: Use xlwings_text.py instead."""
    raise DeprecationWarning(
        "create_text_excel() is deprecated. Use xlwings_text.py instead. "
        "This function creates formulas without calculated values, which defeats "
        "the purpose of Excel compatibility testing."
    )

def create_dynamic_range_excel():
    """DEPRECATED: Use xlwings_dynamic_range.py instead."""
    raise DeprecationWarning(
        "create_dynamic_range_excel() is deprecated. Use xlwings_dynamic_range.py instead. "
        "This function creates formulas without calculated values, which defeats "
        "the purpose of Excel compatibility testing."
    )

def save_excel_files():
    """DEPRECATED: Use generate_all_xlwings.py instead."""
    raise DeprecationWarning(
        "save_excel_files() is deprecated. Use generate_all_xlwings.py instead. "
        "This function creates formulas without calculated values, which defeats "
        "the purpose of Excel compatibility testing."
    )

if __name__ == "__main__":
    print("❌ DEPRECATED: This file is deprecated")
    print("✅ USE INSTEAD: generate_all_xlwings.py")
    print("")
    print("This file creates Excel files with formulas but no calculated values,")
    print("which defeats the purpose of Excel compatibility testing.")
    print("")
    print("The xlwings generators create files with real Excel calculations:")
    print("  python generate_all_xlwings.py")
    print("")
    print("See README_XLWINGS.md for complete usage instructions.")