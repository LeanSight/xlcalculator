#!/usr/bin/env python3
"""
Generate Python test files from JSON test configuration.
Creates FunctionalTestCase classes for dynamic ranges testing.
"""

import argparse
from pathlib import Path
from typing import List, Dict, Any

from json_to_tests_utils import (
    load_json_config, extract_test_levels, extract_metadata, extract_generation_config,
    extract_data_config, validate_json_and_output_dir, map_excel_type_to_python,
    get_test_filename_from_config,
    get_excel_filename_from_config,
    TestLevel, TestCase, count_total_test_cases
)


def generate_test_method_from_case(case: TestCase) -> str:
    """Generate individual test assertion from test case."""
    expected_repr = repr(case.expected_value) if case.expected_value is not None else "None"
    
    if case.expected_type in ["ref_error", "value_error", "name_error", "num_error", "na_error"]:
        return f"""        # {case.description}
        value = self.evaluator.evaluate('{case.cell}')
        self.assertIsInstance(value, {map_excel_type_to_python(case.expected_type)},
                            '{case.formula} should return {case.expected_type.upper()}')"""
    
    elif case.expected_type == "array":
        return f"""        # {case.description}
        value = self.evaluator.evaluate('{case.cell}')
        self.assertIsInstance(value, Array, '{case.formula} should return Array')"""
    
    else:
        python_type = map_excel_type_to_python(case.expected_type)
        return f"""        # {case.description}
        value = self.evaluator.evaluate('{case.cell}')
        self.assertEqual({expected_repr}, value, '{case.formula} should return {expected_repr}')
        self.assertIsInstance(value, {python_type}, 'Should be {case.expected_type}')"""


def generate_test_method_from_level(level: TestLevel) -> str:
    """Generate complete test method from test level."""
    method_name = f"test_{level.level.lower().replace('-', '_').replace(' ', '_')}"
    
    test_assertions = "\n\n".join(
        generate_test_method_from_case(case) for case in level.test_cases
    )
    
    return f"""    def {method_name}(self):
        \"\"\"{level.title}: {level.description}\"\"\"
        
{test_assertions}"""


def generate_data_integrity_method(data_config: Dict[str, Any], gen_config: Dict[str, Any]) -> str:
    """Generate data integrity test method from data configuration."""
    method_name = gen_config.get("integrity_method_name", "test_data_integrity")
    description = gen_config.get("integrity_method_description", "Verify test data integrity")
    
    # Generate assertions from first data row
    assertions = []
    if data_config.get("rows"):
        first_row = data_config["rows"][0]
        for col_idx, value in enumerate(first_row[:3], 1):  # First 3 values
            cell_ref = f"Data!{chr(64 + col_idx)}2"  # A2, B2, C2
            assertions.append(f'        self.assertEqual({repr(value)}, self.evaluator.evaluate(\'{cell_ref}\'))')
    
    assertions_code = "\n".join(assertions)
    
    return f'''    def {method_name}(self):
        """{description}."""
        # Auto-generated data validation
{assertions_code}'''


def generate_type_consistency_method(levels: List[TestLevel], gen_config: Dict[str, Any]) -> str:
    """Generate type consistency test method from test levels."""
    method_name = gen_config.get("consistency_method_name", "test_type_consistency")
    description = gen_config.get("consistency_method_description", "Verify data type consistency")
    
    # Find first test case of each type
    type_samples = {}
    for level in levels:
        for case in level.test_cases:
            if case.expected_type not in type_samples:
                type_samples[case.expected_type] = case
    
    # Generate type checks
    type_checks = []
    for expected_type, case in type_samples.items():
        if expected_type not in ["ref_error", "value_error", "name_error", "num_error", "na_error"]:
            python_type = map_excel_type_to_python(expected_type)
            type_checks.append(f"""        # {expected_type} validation
        {expected_type}_value = self.evaluator.evaluate('{case.cell}')
        self.assertIsInstance({expected_type}_value, {python_type})""")
    
    type_checks_code = "\n\n".join(type_checks)
    
    return f'''    def {method_name}(self):
        """{description}."""
        # Auto-generated type validation
{type_checks_code}'''


def generate_imports() -> str:
    """Generate import statements for test file."""
    return '''"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean'''


def generate_test_class(levels: List[TestLevel], data_config: Dict[str, Any], gen_config: Dict[str, Any], metadata: dict, excel_filename: str) -> str:
    """Generate complete test class with all methods."""
    class_name = gen_config.get("class_name", "ComprehensiveTest")
    class_description = gen_config.get("class_docstring", "Comprehensive integration tests")
    
    test_methods = "\n\n".join(
        generate_test_method_from_level(level) for level in levels
    )
    
    integrity_method = generate_data_integrity_method(data_config, gen_config)
    consistency_method = generate_type_consistency_method(levels, gen_config)
    
    docstring = f"""
    {class_description}.
    
    Tests: {count_total_test_cases(levels)} cases across {len(levels)} levels
    Source: {metadata.get('source', 'JSON configuration')}
    """
    
    return f'''

class {class_name}(testing.FunctionalTestCase):
    """{docstring}"""
    filename = "{excel_filename}"

{test_methods}

{integrity_method}

{consistency_method}'''


def generate_complete_test_file(levels: List[TestLevel], data_config: Dict[str, Any], gen_config: Dict[str, Any], metadata: dict, excel_filename: str) -> str:
    """Generate complete Python test file."""
    imports = generate_imports()
    test_class = generate_test_class(levels, data_config, gen_config, metadata, excel_filename)
    
    return f"{imports}{test_class}\n"


def write_test_file(content: str, output_path: Path) -> None:
    """Write generated test content to file."""
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"Generated test file: {output_path}")


def main(json_path: str, output_dir: str) -> None:
    """Generate Python test file from JSON configuration."""
    json_file, output_path = validate_json_and_output_dir(json_path, output_dir)
    
    config = load_json_config(str(json_file))
    levels = extract_test_levels(config)
    data_config = extract_data_config(config)
    gen_config = extract_generation_config(config)
    metadata = extract_metadata(config)
    
    # Generate filenames from config
    test_filename = get_test_filename_from_config(config)
    excel_filename = get_excel_filename_from_config(config)
    
    test_content = generate_complete_test_file(levels, data_config, gen_config, metadata, excel_filename)
    
    test_file_path = output_path / test_filename
    write_test_file(test_content, test_file_path)
    
    print(f"Generated {count_total_test_cases(levels)} test cases across {len(levels)} levels")
    print(f"Output: {test_file_path}")
    print(f"Excel expected: {excel_filename}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Generate Python test files from JSON configuration"
    )
    parser.add_argument("json_path", help="Path to JSON test configuration file")
    parser.add_argument("output_dir", help="Output directory for generated test file")
    
    args = parser.parse_args()
    main(args.json_path, args.output_dir)