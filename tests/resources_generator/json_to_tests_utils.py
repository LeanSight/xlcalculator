#!/usr/bin/env python3
"""
Test configuration parser for dynamic ranges test generation.
Shared utilities for parsing JSON test configurations and validation.
"""

import json
import os
from pathlib import Path
from typing import Dict, List, Any, NamedTuple


class TestCase(NamedTuple):
    cell: str
    formula: str
    expected_value: Any
    expected_type: str
    description: str


class TestLevel(NamedTuple):
    level: str
    title: str
    description: str
    cell_range: str
    category: str
    test_cases: List[TestCase]


def load_json_config(json_path: str) -> Dict[str, Any]:
    """Load and parse JSON test configuration."""
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def extract_data_config(config: Dict[str, Any]) -> Dict[str, Any]:
    """Extract data sheet configuration from JSON."""
    return config["data_sheet"]


def extract_auxiliary_data(config: Dict[str, Any]) -> Dict[str, Any]:
    """Extract auxiliary test data from JSON."""
    return config["auxiliary_data"]


def extract_metadata(config: Dict[str, Any]) -> Dict[str, Any]:
    """Extract test metadata from JSON."""
    return config["metadata"]


def extract_generation_config(config: Dict[str, Any]) -> Dict[str, Any]:
    """Extract generation configuration from JSON."""
    return config.get("generation_config", {})


def parse_test_case(case_data: Dict[str, Any]) -> TestCase:
    """Parse individual test case from JSON data."""
    return TestCase(
        cell=case_data["cell"],
        formula=case_data["formula"],
        expected_value=case_data["expected_value"],
        expected_type=case_data["expected_type"],
        description=case_data["description"]
    )


def parse_test_level(level_data: Dict[str, Any]) -> TestLevel:
    """Parse test level containing multiple test cases."""
    test_cases = [parse_test_case(case) for case in level_data["test_cases"]]
    
    return TestLevel(
        level=level_data["level"],
        title=level_data["title"],
        description=level_data["description"],
        cell_range=level_data["cell_range"],
        category=level_data["category"],
        test_cases=test_cases
    )


def extract_test_levels(config: Dict[str, Any]) -> List[TestLevel]:
    """Extract all test levels from JSON configuration."""
    return [parse_test_level(level) for level in config["levels"]]


def map_excel_type_to_python(excel_type: str) -> str:
    """Map Excel data types to Python type assertions."""
    type_mapping = {
        "number": "(int, float, Number)",
        "text": "(str, Text)", 
        "boolean": "(bool, Boolean)",
        "array": "Array",
        "ref_error": "xlerrors.RefExcelError",
        "value_error": "xlerrors.ValueExcelError",
        "name_error": "xlerrors.NameExcelError",
        "num_error": "xlerrors.NumExcelError",
        "na_error": "xlerrors.NaExcelError"
    }
    return type_mapping.get(excel_type, "object")


def validate_json_path(json_path: str) -> Path:
    """Validate JSON file exists and is readable."""
    path = Path(json_path)
    if not path.exists():
        raise FileNotFoundError(f"JSON file not found: {json_path}")
    if not path.is_file():
        raise ValueError(f"Path is not a file: {json_path}")
    if path.suffix.lower() != '.json':
        raise ValueError(f"File is not JSON: {json_path}")
    return path


def validate_output_path(output_path: str, extension: str) -> Path:
    """Validate output file path and create parent directory if needed."""
    path = Path(output_path)
    if not path.suffix.lower() == extension.lower():
        raise ValueError(f"Output file must have {extension} extension: {output_path}")
    
    # Create parent directory if it doesn't exist
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def validate_json_and_output(json_path: str, output_path: str, extension: str) -> tuple[Path, Path]:
    """Validate input JSON and output file path."""
    json_file = validate_json_path(json_path)
    output_file = validate_output_path(output_path, extension)
    return json_file, output_file


def get_test_filename_from_config(config: Dict[str, Any]) -> str:
    """Generate test filename from configuration."""
    metadata = extract_metadata(config)
    title = metadata.get("title", "test_suite")
    # Convert to snake_case and add suffix
    base_name = title.replace(" ", "_").replace("-", "_").lower()
    return f"{base_name}_test.py"


def get_excel_filename_from_config(config: Dict[str, Any]) -> str:
    """Generate Excel filename from configuration."""
    metadata = extract_metadata(config)
    title = metadata.get("title", "test_suite")
    # Convert to snake_case and add extension
    base_name = title.replace(" ", "_").replace("-", "_").lower()
    return f"{base_name}.xlsx"


def setup_output_directory(output_dir: str) -> Path:
    """Create output directory if it doesn't exist."""
    path = Path(output_dir)
    path.mkdir(parents=True, exist_ok=True)
    return path


def validate_json_and_output_dir(json_path: str, output_dir: str) -> tuple[Path, Path]:
    """Validate input JSON and setup output directory."""
    json_file = validate_json_path(json_path)
    output_path = setup_output_directory(output_dir)
    return json_file, output_path


def count_total_test_cases(levels: List[TestLevel]) -> int:
    """Count total number of test cases across all levels."""
    return sum(len(level.test_cases) for level in levels)