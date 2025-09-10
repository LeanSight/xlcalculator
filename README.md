# Excel Calculator - Modern Python Fork

[![Python 3.12+](https://img.shields.io/badge/Python-3.12+-blue.svg)](https://github.com/LeanSight/xlcalculator)
[![NumPy 1.24+ & 2.x](https://img.shields.io/badge/NumPy-1.24%2B%20%26%202.x-green.svg)](https://github.com/LeanSight/xlcalculator)
[![Status: Unofficial Fork](https://img.shields.io/badge/Status-Unofficial%20Fork-orange.svg)](https://github.com/LeanSight/xlcalculator)
[![Tests: 962/963 Pass](https://img.shields.io/badge/Tests-962%2F963%20Pass-brightgreen.svg)](https://github.com/LeanSight/xlcalculator)

> **⚠️ ALPHA SOFTWARE WARNING**
> 
> This code is in **ALPHA** stage and has been developed with AI assistance. While comprehensive testing has been performed (962/963 tests pass), this software requires thorough verification and validation before use in production environments.
> 
> **For Production Use:**
> - Perform comprehensive testing with your specific Excel files and use cases
> - Validate all calculations against Excel for accuracy
> - Review and test error handling for your scenarios
> - Consider this a development/testing tool until further validation
> 
> **Use at your own risk** - Always verify results against Excel before relying on calculations for critical applications.

## **UNOFFICIAL FORK** - Modern NumPy and Python Compatible

This is an **unofficial fork** of xlcalculator by Bradley van Ree, updated for modern Python scientific stack compatibility.

**Original repository:** https://github.com/bradbase/xlcalculator

## Fork Improvements

✅ **NumPy 1.24+ and 2.x Support** - Compatible with both NumPy 1.x and 2.x series

✅ **Python 3.12+ Validated** - Tested on Python 3.12.1, compatible with 3.13+

✅ **Modern Dependencies** - Updated to latest scientific Python stack

✅ **Enhanced Excel Compatibility** - Improved reference parsing and dynamic ranges

✅ **YEARFRAC Support** - Includes LeanSight yearfrac fork for complete Excel function compatibility

✅ **Comprehensive Testing** - 962/963 tests pass (99.9% success rate)

## Recent Enhancements (2025-09-10)

🚀 **Enhanced Reference Parsing**
   - Column references (A:A, B:B) 
   - Row references (1:1, 2:2)
   - Range references (A1:B5)
   - Dynamic INDIRECT construction

🚀 **Improved Error Handling**
   - Excel-compatible error propagation
   - ISERROR/IFERROR integration
   - Bounds checking validation

🚀 **Infrastructure Improvements**
   - Reference normalization system
   - Array parameter support in OFFSET
   - Comprehensive refactoring with test validation

## Installation

**Basic Installation:**
```bash
pip install git+https://github.com/LeanSight/xlcalculator.git
```

**With Excel Functions (including YEARFRAC):**
```bash
pip install git+https://github.com/LeanSight/xlcalculator.git[excel_functions]
```

**Development Installation:**
```bash
git clone https://github.com/LeanSight/xlcalculator.git
cd xlcalculator
pip install -e .[test,excel_functions]
```

## Requirements

- **Python:** 3.12+ (validated on 3.12.1)
- **NumPy:** 1.24+ (tested with 1.26.4 and 2.3.3)
- **pandas:** 2.3.0+
- **scipy:** 1.14.1+

## Validation Status

This fork has been thoroughly tested with:

| Component | Version | Status |
|-----------|---------|--------|
| Python | 3.12.1 | ✅ Validated |
| NumPy | 1.26.4 | ✅ All tests pass |
| NumPy | 2.3.3 | ✅ All tests pass |
| pandas | 2.3.2 | ✅ Compatible |
| scipy | 1.16.1 | ✅ Compatible |
| Test Suite | 962/963 | ✅ 99.9% pass rate |
| Excel Functions | All | ✅ Working |
| YEARFRAC | All methods | ✅ Working |

## About xlcalculator

xlcalculator is a Python library that reads MS Excel files and, to the extent
of supported functions, can translate the Excel functions into Python code and
subsequently evaluate the generated Python code. Essentially doing the Excel
calculations without the need for Excel.

xlcalculator is a modernization of the [koala2](https://github.com/vallettea/koala) library.

`xlcalculator` currently supports:

* **Loading an Excel file into a Python compatible state** - [Example](examples/common_use_case/)
* **Saving Python compatible state** - [Example](examples/common_use_case/)
* **Loading Python compatible state** - [Example](examples/common_use_case/)
* **Ignore worksheets** - [Example](examples/ignore_worksheets/)
* **Extracting sub-portions of a model** - [Example](examples/model_focusing/) - "focussing" on provided cell addresses or defined names
* **Evaluating:**

    * **Individual cells** - [Example](examples/common_use_case/)
    * **Defined Names** (a "named cell" or range) - [Example](examples/common_use_case/)
    * **Ranges** - Basic range support available
    * **Shared formulas** - [not an Array Formula](https://stackoverflow.com/questions/1256359/what-is-the-difference-between-a-shared-formula-and-an-array-formula)
    * **Operands** (+, -, /, \\*, ==, <>, <=, >=) - Basic arithmetic and comparison operators
    * **Set cell value** - [Example](examples/common_use_case/)
    * **Get cell value** - [Example](examples/common_use_case/)
    * **Parsing a dict into the Model object** - [Example](examples/third_party_datastructure/)

## Enhanced Excel Function Support

This fork includes enhanced support for:

**Dynamic Range Functions:** - Basic support available
    * INDEX - Basic implementation with standard references
    * OFFSET - Basic implementation available
    * INDIRECT - Basic implementation available

**Reference Types:** - Standard Excel reference support
    * Column references: A:A, B:B, $A:$A
    * Row references: 1:1, 2:2, $1:$1  
    * Range references: A1:B5, $A$1:$B$5
    * Dynamic construction: Basic INDIRECT support

**Error Handling:** - Basic error handling available
    * Standard error propagation
    * ISERROR/IFERROR basic support
    * Standard bounds checking

**Mathematical Functions:** - Core mathematical functions available
    * LN - Python Math.log() differs from Excel LN. Currently returning Math.log()
    * VLOOKUP - Exact match only
    * YEARFRAC - All daycount methods supported (see Supported Functions section for details)

**Not currently supported:**

  * Array Formulas or CSE Formulas (not a shared formula)
  * Functions required to complete testing as per Microsoft Office Help
    website for SQRT and LN
  * EXP, DB functions

## Migration from Original

This fork is a **drop-in replacement**. Simply change your installation:

```bash
# Before (original):
pip install xlcalculator

# After (fork):
pip install git+https://github.com/LeanSight/xlcalculator.git
```

**No code changes required** - all APIs remain identical.

**Enhanced Features Available:**
    * Improved NumPy 1.24+ and 2.x compatibility
    * Enhanced reference parsing (column/row references)
    * Better error handling and bounds checking
    * YEARFRAC function support with all daycount methods

## Usage Example

```python
from xlcalculator import ModelCompiler
from xlcalculator import Model

# Load Excel file
compiler = ModelCompiler()
model = compiler.read_and_parse_archive("example.xlsx")

# Evaluate cells
result = model.evaluate("Sheet1!A1")

# Evaluate ranges
range_result = model.evaluate("Sheet1!A1:C3")

# Evaluate with column references
column_result = model.evaluate("Sheet1!A:A")
```

## Examples

Working examples are available in the [examples/](examples/) directory:

**Core Functionality:**
    * [Basic Usage](examples/common_use_case/) - Loading, saving, evaluating Excel files
    * [Third-party Data](examples/third_party_datastructure/) - Working with Python dictionaries
    * [Rounding Operations](examples/rounding_example/) - Precision handling and floating-point behavior

**Performance & Optimization:**
    * [Model Focusing](examples/model_focusing/) - Focus on specific model portions using ignore_sheets
    * [Ignore Worksheets](examples/ignore_worksheets/) - Selective sheet loading for performance optimization

Each example includes:
    * **Working Code** - Fully functional demonstrations with real Excel files
    * **ATDD Tests** - Test-driven development approach with comprehensive test cases (where applicable)
    * **Documentation** - Clear explanations of functionality and usage patterns

## Run Example

From the examples/common_use_case directory:

```bash
python use_case_01.py
```

This will demonstrate basic Excel file loading, evaluation, and saving functionality.

## Run Tests

Setup your environment:

```bash
python -m venv ve
ve\Scripts\activate  # Windows
pip install -e .[test]
```

From the root xlcalculator directory:

```bash
python -m pytest tests/
```

Or use `tox` (if available):

```bash
tox
```

**Test Coverage:**

```bash
# Total test coverage
python -m pytest tests/ --collect-only
# Result: 963 tests collected

# Run with coverage
python -m pytest tests/ -v
# Result: 962 passed, 1 skipped (99.9% success rate)
```

## Adding/Registering Excel Functions

Excel function support can be easily added.

Fundamental function support is found in the xlfunctions directory. The
functions are thematically organised in modules.

Excel functions can be added by any code using the
`xlfunctions.xl.register()` decorator. Here is a simple example:

```python
from xlcalculator.xlfunctions import xl

@xl.register()
@xl.validate_args
def ADDONE(num: xl.Number):
    return num + 1
```

The `@xl.validate_args` decorator will ensure that the annotated arguments are
converted and validated. For example, even if you pass in a string, it is
converted to a number (in typical Excel fashion):

```python
>>> ADDONE(1):
2
>>> ADDONE('1'):
2
```

If you would like to contribute functions, please create a pull request. All
new functions should be accompanied by sufficient tests to cover the
functionality. Tests need to be written for both the Python implementation of
the function (tests/xlfunctions) and a comparison with Excel
(tests/xlfunctions_vs_excel).

## Excel number precision

Excel number precision is a complex discussion.

It has been discussed in a [Wikipedia page](https://en.wikipedia.org/wiki/Numeric_precision_in_Microsoft_Excel).

The fundamentals come down to floating point numbers and a contention between
how they are represented in memory Vs how they are stored on disk Vs how they
are presented on screen. A [Microsoft article](https://www.microsoft.com/en-us/microsoft-365/blog/2008/04/10/understanding-floating-point-precision-aka-why-does-excel-give-me-seemingly-wrong-answers/)
explains the contention.

This project is attempting to take care while reading numbers from the Excel
file to try and remove a variety of representation errors.

Further work will be required to keep numbers in-line with Excel throughout
different transformations.

From what I can determine this requires a low-level implementation of a
numeric datatype (C or C++, Cython??) to replicate its behaviour. Python
built-in numeric types don't replicate behaviours appropriately.

## Unit testing Excel formulas directly from the workbook

If you are interested in unit testing formulas in your workbook, you can use
[FlyingKoala](https://github.com/bradbase/flyingkoala). An example on how can
be found
[here](https://github.com/bradbase/flyingkoala/tree/master/flyingkoala/unit_testing_formulas).

## Dependencies

This fork includes these updated dependencies:

**Core Dependencies:**
    * `numpy>=1.24.0` (supports both 1.x and 2.x series)
    * `pandas>=2.3.0`
    * `scipy>=1.14.1`
    * `openpyxl` (latest)
    * `numpy-financial` (latest)
    * `jsonpickle` (latest)

**Excel Functions (Optional):**
    * `git+https://github.com/LeanSight/yearfrac.git` (NumPy 1.24+ and 2.x compatible fork)

## Related Forks

This xlcalculator fork depends on:

* **LeanSight yearfrac fork:** https://github.com/LeanSight/yearfrac
  - Adds NumPy 1.24+ and 2.x compatibility to yearfrac
  - Enables YEARFRAC Excel function support

## Known Limitations

* **Python Support:** Validated on Python 3.12.1, compatible with 3.13+
* **Platform:** Primarily validated on Linux, should work on Windows/macOS
* **Excel Functions:** Some advanced Excel functions may not be supported (same as original)

## Support

**For Fork-Specific Issues:**
    * **Issues:** https://github.com/LeanSight/xlcalculator/issues
    * **Discussions:** Use GitHub Discussions on the fork repo

**For Original Functionality:**
    * **Documentation:** Refer to original xlcalculator documentation
    * **Excel Functions:** Check original function support list

## Contributing

Contributions welcome! Please:

1. Fork this repository (not the original)
2. Create feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open Pull Request

## TODO

- Do not treat ranges as a granular AST node it instead as an operation ":" of
  two cell references to create the range. That will make implementing
  features like `A1:OFFSET(...)` easy to implement.

- Support for alternative range evaluation: by ref (pointer), by expr (lazy
  eval) and current eval mode.

    * Pointers would allow easy implementations of functions like OFFSET().

    * Lazy evals will allow efficient implementation of IF() since execution
      of true and false expressions can be delayed until it is decided which
      expression is needed.

- Implement array functions. It is really not that hard once a proper
  RangeData class has been implemented on which one can easily act with scalar
  functions.

- Improve testing

- Refactor model and evaluator to use pass-by-object-reference for values of
  cells which then get "used"/referenced by ranges, defined names and formulas

- Handle multi-file addresses

- Improve integration with pyopenxl for reading and writing files

## Supported Functions

For the complete list of supported functions, see the original documentation.
This fork maintains full compatibility with all original functions plus:

* **Enhanced dynamic range support** - Improved INDEX, OFFSET, INDIRECT functions
* **YEARFRAC support** - All daycount methods supported via LeanSight yearfrac fork
  
  * Basis 1, Actual/actual, is within 3 decimal places of Excel
  * All other basis methods match Excel exactly

* **Improved reference parsing** - Column references (A:A), row references (1:1), range references
* **Better error handling** - Excel-compatible error propagation and bounds checking

## Credits

**Original Author:** Bradley van Ree

**Fork Maintainer:** LeanSight

**License:** MIT

**Original Repository:** https://github.com/bradbase/xlcalculator

**Fork Repository:** https://github.com/LeanSight/xlcalculator

**Last Updated:** 2025-09-10

**Validation Date:** 2025-09-10