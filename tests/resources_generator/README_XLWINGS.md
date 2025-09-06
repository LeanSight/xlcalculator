# Excel File Generation with xlwings

This directory contains xlwings-based generators for creating Excel files with **real Excel calculations** for integration testing.

## Purpose

The goal is to create Excel files that contain:
1. **Formulas** (for xlcalculator to evaluate)
2. **Excel's actual calculated values** (for comparison)

This ensures xlcalculator integration tests compare against Excel's real behavior, not manual calculations.

## Files Created This Week

These generators create Excel files for features added/tested this week:

### Individual Generators
- `xlwings_information.py` - Creates `INFORMATION.xlsx` with IS* function tests
- `xlwings_logical.py` - Creates `logical.xlsx` with AND, OR, TRUE, FALSE tests  
- `xlwings_math.py` - Creates `MATH.xlsx` with FLOOR, TRUNC, SIGN, LOG, EXP tests
- `xlwings_text.py` - Creates `TEXT.xlsx` with LEFT, UPPER, LOWER, TRIM, REPLACE tests

### Master Generator
- `generate_all_xlwings.py` - Runs all generators and creates all Excel files

## Requirements

**Windows Environment:**
- Microsoft Excel installed
- Python with xlwings: `pip install xlwings`

## Usage

### Option 1: Generate All Files
```bash
cd tests/resources_generator
python generate_all_xlwings.py
```

### Option 2: Generate Individual Files
```bash
cd tests/resources_generator
python xlwings_information.py
python xlwings_logical.py
python xlwings_math.py
python xlwings_text.py
```

## Output

Generated files will be created in `generated_excel_files/` directory:
- `INFORMATION.xlsx` - Information function tests
- `logical.xlsx` - Logical function tests
- `MATH.xlsx` - Math function tests  
- `TEXT.xlsx` - Text function tests

## Integration

After generation on Windows:
1. Copy generated Excel files to `tests/resources/`
2. Run integration tests: `python -m pytest tests/xlfunctions_vs_excel/ -v`

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

## Why xlwings?

- **Real Excel calculations** - Uses actual Excel engine
- **Formula + Value storage** - Excel calculates and stores both
- **Exact compatibility** - Matches Excel behavior precisely
- **No manual calculation errors** - Excel does the math

This ensures integration tests validate xlcalculator against Excel's actual behavior, not assumptions.