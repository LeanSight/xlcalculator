# xlcalculator Examples

This directory contains working examples demonstrating xlcalculator functionality.

## 📁 Available Examples

### Core Functionality

#### [common_use_case/](common_use_case/)
**Basic Excel file operations**
- Loading Excel files (.xlsm format)
- Evaluating cells and formulas
- Working with defined names
- Setting cell values
- Saving and loading model state

**Files:**
- `use_case_01.py` - Main example script
- `use_case_01.xlsm` - Sample Excel file
- `use_case_01.json` - Saved model state

#### [third_party_datastructure/](third_party_datastructure/)
**Working with Python dictionaries**
- Creating models from Python dictionaries
- Evaluating formulas without Excel files
- Dynamic model creation

**Files:**
- `third_party_datastructure.py` - Dictionary-based model example

#### [rounding_example/](rounding_example/)
**Precision and rounding behavior**
- Floating-point precision handling
- Excel vs Python rounding differences
- Numerical accuracy considerations

**Files:**
- `rounding_example.py` - Precision demonstration

### Performance & Optimization

#### [ignore_worksheets/](ignore_worksheets/)
**Selective sheet loading**
- Excluding specific worksheets during loading
- Performance optimization for large files
- Memory usage reduction

**Files:**
- `ignore_worksheets_example.py` - Main example
- `test_ignore_worksheets.py` - ATDD test suite

#### [model_focusing/](model_focusing/)
**Model focusing and optimization**
- Using `ignore_sheets` parameter for performance
- Working with large financial models
- Scenario analysis with focused models

**Files:**
- `model_focusing_example.py` - Comprehensive focusing example
- `test_model_focusing.py` - ATDD test suite
- `README.md` - Detailed documentation

## 🧪 Running Examples

### Individual Examples
```bash
# Basic usage
cd examples/common_use_case
python use_case_01.py

# Dictionary-based models
cd examples/third_party_datastructure
python third_party_datastructure.py

# Rounding behavior
cd examples/rounding_example
python rounding_example.py

# Ignore worksheets
cd examples/ignore_worksheets
python ignore_worksheets_example.py

# Model focusing
cd examples/model_focusing
python model_focusing_example.py
```

### Running Tests
```bash
# Test ignore worksheets functionality
cd examples/ignore_worksheets
python -m pytest test_ignore_worksheets.py -v

# Test model focusing functionality
cd examples/model_focusing
python -m pytest test_model_focusing.py -v
```

### Quick Verification
```bash
# Verify all examples work (run from examples directory)
cd examples/common_use_case && python use_case_01.py > /dev/null && echo "✅ common_use_case"
cd ../third_party_datastructure && python third_party_datastructure.py > /dev/null && echo "✅ third_party_datastructure"  
cd ../rounding_example && python rounding_example.py > /dev/null && echo "✅ rounding_example"
cd ../ignore_worksheets && python ignore_worksheets_example.py > /dev/null && echo "✅ ignore_worksheets"
cd ../model_focusing && python model_focusing_example.py > /dev/null && echo "✅ model_focusing"
```

## 📊 Example Status

| Example | Status | Tests | Files | Description |
|---------|--------|-------|-------|-------------|
| common_use_case | ✅ Working | No tests | 3 files | Basic Excel operations |
| third_party_datastructure | ✅ Working | No tests | 1 file | Dictionary-based models |
| rounding_example | ✅ Working | No tests | 1 file | Precision handling |
| ignore_worksheets | ✅ Working | ✅ 3 tests pass | 2 files | Performance optimization |
| model_focusing | ✅ Working | ✅ 4 tests pass | 3 files | Advanced focusing techniques |

## 🎯 Key Features Demonstrated

### Excel File Operations
- Loading Excel files with `ModelCompiler.read_and_parse_archive()`
- Evaluating cells, ranges, and defined names
- Setting cell values dynamically
- Saving and loading model state

### Performance Optimization
- Using `ignore_sheets` parameter to exclude unnecessary worksheets
- Reducing memory usage for large Excel files
- Improving loading and evaluation performance

### Model Management
- Working with complex financial models
- Scenario analysis and sensitivity testing
- Dependency management and calculation chains

### Data Integration
- Creating models from Python dictionaries
- Working without Excel files
- Dynamic model construction

## ⚠️ Important Notes

- All examples use **only verified functionality** that exists in xlcalculator
- Examples with ATDD tests have been thoroughly validated
- Only functionality that has been tested and confirmed working is documented
- Always test examples in your environment before using in production
- Some examples require specific file paths - run from their respective directories

## 🔗 Related Documentation

- [Main README](../README.rst) - Complete xlcalculator documentation
- [Model Focusing](model_focusing/README.md) - Detailed focusing documentation