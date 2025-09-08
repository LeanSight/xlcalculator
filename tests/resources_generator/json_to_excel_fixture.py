#!/usr/bin/env python3
"""
Generate Excel test files from JSON test configuration.
Creates Excel workbooks with Data and Tests sheets for dynamic ranges testing.
"""

import argparse
from pathlib import Path
from typing import Dict, List, Any

try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False

from json_to_tests_utils import (
    load_json_config, extract_test_levels, extract_data_config,
    extract_auxiliary_data, extract_metadata, extract_generation_config,
    validate_json_and_output_dir, count_total_test_cases, get_excel_filename_from_config, TestLevel
)


def check_xlwings_requirements() -> None:
    """Verify xlwings is available and Excel is accessible."""
    if not XLWINGS_AVAILABLE:
        raise ImportError("xlwings not available. Install with: pip install xlwings")
    
    try:
        app = xw.App(visible=False)
        app.quit()
    except Exception as e:
        raise RuntimeError(f"Excel not accessible: {e}")


def create_data_sheet(wb, data_config: Dict[str, Any]) -> None:
    """Create and populate Data sheet with test data."""
    data_sheet = wb.sheets[0]
    data_sheet.name = "Data"
    
    # Add headers
    headers = data_config["headers"]
    for i, header in enumerate(headers, 1):
        data_sheet.cells(1, i).value = header
    
    # Add data rows
    for row_idx, row_data in enumerate(data_config["rows"], 2):
        for col_idx, value in enumerate(row_data, 1):
            data_sheet.cells(row_idx, col_idx).value = value


def create_auxiliary_data(tests_sheet, aux_data: Dict[str, Any]) -> None:
    """Add auxiliary data for INDIRECT tests."""
    for cell, value in aux_data.items():
        tests_sheet[cell].value = value


def add_formula_to_sheet(tests_sheet, cell: str, formula: str) -> None:
    """Add formula to specific cell with error handling."""
    try:
        tests_sheet[cell].formula = formula
    except Exception as e:
        print(f"Warning: Failed to add formula {cell}: {formula} - {e}")


def populate_test_formulas(tests_sheet, levels: List[TestLevel]) -> int:
    """Add all test formulas to Tests sheet."""
    formula_count = 0
    
    for level in levels:
        for case in level.test_cases:
            add_formula_to_sheet(tests_sheet, case.cell, case.formula)
            formula_count += 1
    
    return formula_count


def add_level_labels(tests_sheet, levels: List[TestLevel], gen_config: Dict[str, Any]) -> None:
    """Add descriptive labels for each test level."""
    label_row = gen_config.get("label_row", 20)
    
    for level in levels:
        # Extract column from first test case
        if level.test_cases:
            first_cell = level.test_cases[0].cell
            column = first_cell.rstrip('0123456789')
            label_cell = f"{column}{label_row}"
            tests_sheet[label_cell].value = f"{level.level}: {level.title}"


def create_tests_sheet(wb, levels: List[TestLevel], aux_data: Dict[str, Any], gen_config: Dict[str, Any]) -> int:
    """Create and populate Tests sheet with formulas."""
    tests_sheet = wb.sheets.add("Tests")
    
    # Add auxiliary data first
    create_auxiliary_data(tests_sheet, aux_data)
    
    # Add test formulas
    formula_count = populate_test_formulas(tests_sheet, levels)
    
    # Add descriptive labels
    add_level_labels(tests_sheet, levels, gen_config)
    
    return formula_count


def force_calculation(wb) -> None:
    """Force Excel to calculate all formulas."""
    try:
        wb.app.calculate()
        print("Excel calculation completed")
    except Exception as e:
        print(f"Calculation warning: {e}")


def create_excel_workbook(config: Dict[str, Any], output_path: Path) -> None:
    """Create complete Excel workbook from configuration."""
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    
    try:
        wb = app.books.add()
        
        # Create sheets
        data_config = extract_data_config(config)
        gen_config = extract_generation_config(config)
        create_data_sheet(wb, data_config)
        
        levels = extract_test_levels(config)
        aux_data = extract_auxiliary_data(config)
        formula_count = create_tests_sheet(wb, levels, aux_data, gen_config)
        
        # Calculate and save
        force_calculation(wb)
        wb.save(str(output_path))
        
        print(f"Excel created: {output_path}")
        print(f"Added {formula_count} formulas across {len(levels)} levels")
        
    except Exception as e:
        print(f"Excel creation failed: {e}")
        raise
    finally:
        try:
            if 'wb' in locals():
                wb.close()
        except:
            pass
        try:
            app.quit()
        except:
            pass


def generate_summary_report(config: Dict[str, Any]) -> None:
    """Print summary of Excel generation."""
    metadata = extract_metadata(config)
    gen_config = extract_generation_config(config)
    levels = extract_test_levels(config)
    total_cases = count_total_test_cases(levels)
    
    title = metadata.get('title', gen_config.get('class_name', 'Test Suite'))
    
    print("\n" + "="*60)
    print("EXCEL GENERATION SUMMARY")
    print("="*60)
    print(f"Title: {title}")
    print(f"Total test cases: {total_cases}")
    print(f"Test levels: {len(levels)}")
    print(f"Source: {metadata.get('source', 'JSON configuration')}")
    
    print("\nTEST CATEGORIES:")
    for level in levels:
        case_count = len(level.test_cases)
        print(f"   - {level.title}: {case_count} cases")


def main(json_path: str, output_dir: str) -> None:
    """Generate Excel file from JSON configuration."""
    check_xlwings_requirements()
    
    json_file, output_path = validate_json_and_output_dir(json_path, output_dir)
    
    config = load_json_config(str(json_file))
    
    # Generate filename from config
    excel_filename = get_excel_filename_from_config(config)
    excel_file_path = output_path / excel_filename
    
    print(f"Starting Excel generation...")
    print(f"Output: {excel_file_path}")
    
    create_excel_workbook(config, excel_file_path)
    generate_summary_report(config)
    
    print("\nNEXT STEPS:")
    print("1. Use generated Excel for integration testing")
    print("2. Run tests to verify Excel compatibility")
    print("3. Implement functions using red-green-refactor strategy")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Generate Excel test files from JSON configuration"
    )
    parser.add_argument("json_path", help="Path to JSON test configuration file")
    parser.add_argument("output_dir", help="Output directory for generated Excel file")
    
    args = parser.parse_args()
    main(args.json_path, args.output_dir)