# Excel File Generation with xlwings

This directory contains xlwings-based generators for creating Excel files with **real Excel calculations** for integration testing.

## Purpose

The goal is to create Excel files that contain:
1. **Formulas** (for xlcalculator to evaluate)
2. **Excel's actual calculated values** (for comparison)

This ensures xlcalculator integration tests compare against Excel's real behavior, not manual calculations.

## Complete xlwings Generation Suite

These generators create ALL Excel files needed for xlcalculator integration testing:

### Individual Generators
- `xlwings_information.py` - Creates `INFORMATION.xlsx` with IS* function tests
- `xlwings_logical.py` - Creates `logical.xlsx` with AND, OR, TRUE, FALSE tests  
- `xlwings_math.py` - Creates `MATH.xlsx` with FLOOR, TRUNC, SIGN, LOG, EXP tests
- `xlwings_text.py` - Creates `TEXT.xlsx` with LEFT, UPPER, LOWER, TRIM, REPLACE tests
- `xlwings_xlookup.py` - Creates `XLOOKUP.xlsx` with XLOOKUP function tests
- `xlwings_dynamic_range.py` - Creates `DYNAMIC_RANGE.xlsx` with INDEX, OFFSET, INDIRECT tests

### Master Generator
- `generate_all_xlwings.py` - Runs all generators and creates all 6 Excel files

### Deprecated
- `excel_file_templates.py` - ⚠️ DEPRECATED - Use xlwings generators instead

## Requirements

**Windows Environment:**
- Microsoft Excel installed
- Python with xlwings

**Installation:**
```bash
# Option 1: Install xlwings directly
pip install xlwings

# Option 2: Install as optional dependency
pip install xlcalculator[excel_generation]

# Option 3: Install from this project with excel generation support
pip install -e .[excel_generation]
```

## Usage

### Option 1: Generate All Files
```bash
cd tests/resources_generator

# Generate to default directory (generated_excel_files)
python generate_all_xlwings.py

# Generate directly to tests/resources
python generate_all_xlwings.py ../resources

# Generate to custom directory
python generate_all_xlwings.py C:\temp\excel_files

# Check requirements only
python generate_all_xlwings.py --check-only
```

### Option 2: Generate Individual Files
```bash
cd tests/resources_generator
python xlwings_information.py
python xlwings_logical.py
python xlwings_math.py
python xlwings_text.py
python xlwings_xlookup.py
python xlwings_dynamic_range.py
```

## Output

Generated files will be created in the specified output directory:
- `INFORMATION.xlsx` - Information function tests
- `logical.xlsx` - Logical function tests
- `MATH.xlsx` - Math function tests  
- `TEXT.xlsx` - Text function tests
- `XLOOKUP.xlsx` - XLOOKUP function tests
- `DYNAMIC_RANGE.xlsx` - Dynamic range function tests

## Integration

After generation on Windows:

**If generated to default directory:**
1. Copy generated Excel files: `copy generated_excel_files\*.xlsx ..\resources\`
2. Run integration tests: `python -m pytest tests/xlfunctions_vs_excel/ -v`

**If generated directly to tests/resources:**
1. Run integration tests: `python -m pytest tests/xlfunctions_vs_excel/ -v`

## What Each File Tests

### INFORMATION.xlsx
- ISNUMBER, ISTEXT, ISBLANK, ISERROR, ISERR, ISNA, ISEVEN, ISODD, NA
- Tests with numbers, text, booleans, blanks, errors

### logical.xlsx  
- AND, OR, TRUE, FALSE functions
- Nested logical operations
- Edge cases (empty AND/OR)

### MATH.xlsx
- FLOOR, TRUNC, SIGN, LOG, LOG10, EXP
- Various numeric inputs and edge cases

### TEXT.xlsx
- LEFT, UPPER, LOWER, TRIM, REPLACE
- String manipulation and formatting

### XLOOKUP.xlsx
- XLOOKUP function with all match modes
- Exact, approximate, wildcard, reverse, binary search
- Error handling and edge cases

### DYNAMIC_RANGE.xlsx
- INDEX, OFFSET, INDIRECT functions
- Dynamic range references and calculations
- Complex nested function combinations

## Why xlwings?

- **Real Excel calculations** - Uses actual Excel engine
- **Formula + Value storage** - Excel calculates and stores both
- **Exact compatibility** - Matches Excel behavior precisely
- **No manual calculation errors** - Excel does the math

This ensures integration tests validate xlcalculator against Excel's actual behavior, not assumptions.