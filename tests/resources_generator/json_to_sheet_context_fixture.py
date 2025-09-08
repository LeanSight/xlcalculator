#!/usr/bin/env python3
"""
Generate Excel test files for sheet context behavior testing.
Creates Excel workbooks with multiple sheets for testing implicit reference resolution.

Follows the json_to_excel_fixture.py model for consistent test file generation.
"""

import argparse
import os
from pathlib import Path
from typing import Dict, List, Any

try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False

# Fallback to openpyxl if xlwings not available
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

from json_to_tests_utils import (
    load_json_config, extract_metadata, extract_generation_config,
    validate_json_and_output_dir
)


def check_excel_requirements() -> str:
    """Check available Excel libraries and return preferred method."""
    if XLWINGS_AVAILABLE:
        try:
            app = xw.App(visible=False)
            app.quit()
            return "xlwings"
        except Exception:
            pass
    
    if OPENPYXL_AVAILABLE:
        return "openpyxl"
    
    raise ImportError("No Excel library available. Install xlwings or openpyxl")


def create_workbook_xlwings(workbook_config: Dict[str, Any], output_path: Path) -> None:
    """Create Excel workbook using xlwings (preferred for formula calculation)."""
    app = xw.App(visible=False)
    try:
        wb = app.books.add()
        
        # Remove default sheets and create configured sheets
        for sheet in wb.sheets:
            sheet.delete()
        
        for sheet_config in workbook_config["sheets"]:
            sheet = wb.sheets.add(sheet_config["name"])
            
            # Add data values
            for cell_ref, value in sheet_config["data"].items():
                sheet.range(cell_ref).value = value
            
            # Add formulas
            for cell_ref, formula in sheet_config["formulas"].items():
                sheet.range(cell_ref).formula = formula
        
        # Save and close
        wb.save(str(output_path))
        wb.close()
        
    finally:
        app.quit()


def create_workbook_openpyxl(workbook_config: Dict[str, Any], output_path: Path) -> None:
    """Create Excel workbook using openpyxl (fallback method)."""
    wb = openpyxl.Workbook()
    
    # Remove default sheet
    wb.remove(wb.active)
    
    for sheet_config in workbook_config["sheets"]:
        sheet = wb.create_sheet(sheet_config["name"])
        
        # Add data values
        for cell_ref, value in sheet_config["data"].items():
            sheet[cell_ref] = value
        
        # Add formulas
        for cell_ref, formula in sheet_config["formulas"].items():
            sheet[cell_ref] = formula
    
    wb.save(str(output_path))


def generate_excel_file(config: Dict[str, Any], output_path: Path) -> None:
    """Generate Excel file from JSON configuration."""
    workbook_config = config["workbook_config"]
    excel_method = check_excel_requirements()
    
    print(f"Creating Excel file using {excel_method}...")
    
    if excel_method == "xlwings":
        create_workbook_xlwings(workbook_config, output_path)
    else:
        create_workbook_openpyxl(workbook_config, output_path)
    
    print(f"Created Excel file: {output_path}")


def print_expected_behavior(config: Dict[str, Any]) -> None:
    """Print expected behavior documentation."""
    expected = config.get("expected_behavior", {})
    
    print("\nExpected Excel Behavior:")
    for behavior, description in expected.items():
        print(f"- {behavior.replace('_', ' ').title()}: {description}")
    
    print("\nTest Cases Summary:")
    for test_level in config["test_cases"]:
        print(f"Level {test_level['level']}: {test_level['title']} ({len(test_level['cases'])} cases)")


def main(json_path: str, output_dir: str) -> None:
    """Generate Excel test file from JSON configuration."""
    json_file, output_path_dir = validate_json_and_output_dir(json_path, output_dir)
    
    config = load_json_config(str(json_file))
    gen_config = extract_generation_config(config)
    metadata = extract_metadata(config)
    
    # Get Excel filename from config
    excel_filename = gen_config.get("excel_filename", "sheet_context_test.xlsx")
    excel_path = output_path_dir / excel_filename
    
    # Generate Excel file
    generate_excel_file(config, excel_path)
    
    # Print summary
    print(f"\nGenerated: {excel_filename}")
    print(f"Title: {metadata['title']}")
    print(f"Description: {metadata['description']}")
    print(f"Total test cases: {metadata['total_cases']}")
    
    print_expected_behavior(config)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Generate Excel test files for sheet context behavior testing"
    )
    parser.add_argument("json_path", help="Path to JSON test configuration file")
    parser.add_argument("output_dir", help="Output directory for generated Excel file")
    
    args = parser.parse_args()
    main(args.json_path, args.output_dir)