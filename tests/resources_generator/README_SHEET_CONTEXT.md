# Sheet Context Test Generation

## Overview

Sheet context tests use the existing `json_to_excel_fixture.py` infrastructure with a properly formatted JSON configuration file.

## Usage

### Generate Excel File

```bash
# Using json_to_excel_fixture.py (requires xlwings)
cd tests/resources_generator
python json_to_excel_fixture.py sheet_context_test_cases.json ../resources

# Alternative: Manual generation using openpyxl (fallback)
cd tests/resources
python -c "
import openpyxl
# ... manual creation script ...
"
```

### JSON Structure

The `sheet_context_test_cases.json` follows the established pattern:

```json
{
  "metadata": { /* title, description, etc. */ },
  "generation_config": { /* class_name, filenames, etc. */ },
  "data_sheet": { /* Sheet1 data and formulas */ },
  "auxiliary_data": { /* Sheet2 data and formulas */ },
  "levels": [ /* Test cases organized by level */ ]
}
```

### Generated Structure

The Excel file contains:
- **Sheet1**: Base data (A1:A3=10,20,30) and formulas (C1=SUM, C2=mixed, D1=cross-ref)
- **Sheet2**: Base data (A1:A3=100,200,300) and formulas (C1=SUM, C2=mixed, D1=cross-ref)  
- **Tests**: Test formulas referencing the other sheets (A1:D2 with 8 test cases)

### Test Cases

1. **Level 1**: Implicit reference context (2 cases)
2. **Level 2**: Mixed references (2 cases)
3. **Level 3**: Cross-sheet references (2 cases)
4. **Level 4**: Data integrity (2 cases)

## Integration with Tests

The generated Excel file works with:
- `test_sheet_context_integration.py` (3 integration tests)
- `test_sheet_context_unit.py` (5 unit tests)

## Key Principles

1. **Use existing infrastructure**: No custom generators needed
2. **JSON-driven configuration**: Maintainable and explicit test definitions
3. **Multiple generation methods**: xlwings (preferred) or openpyxl (fallback)
4. **Backward compatibility**: Same test data and expected results