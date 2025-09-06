# Error Handling and Validation Strategy

## Overview

This document outlines the comprehensive error handling and validation strategy for dynamic range functions, ensuring Excel compatibility and robust error reporting.

## Excel Error Types

### Standard Excel Errors
```python
# xlcalculator.xlfunctions.xlerrors

#REF!    - RefExcelError      - Reference errors (out of bounds, invalid range)
#VALUE!  - ValueExcelError    - Value errors (wrong type, invalid parameter)
#NAME?   - NameExcelError     - Name errors (invalid reference text, unknown function)
#N/A     - NaExcelError       - Not available (lookup failures, no match found)
#NUM!    - NumExcelError      - Number errors (invalid numeric operations)
#DIV/0!  - DivZeroExcelError  - Division by zero
#NULL!   - NullExcelError     - Null intersection (range operator errors)
```

### Error Mapping for Dynamic Range Functions

| Function | Common Errors | Excel Error Type | Trigger Conditions |
|----------|---------------|------------------|-------------------|
| OFFSET   | Out of bounds | #REF! | Negative coordinates, beyond Excel limits |
|          | Invalid params | #VALUE! | Non-numeric rows/cols, invalid reference |
| INDEX    | Out of bounds | #REF! | Row/col index beyond array dimensions |
|          | Invalid array | #VALUE! | Empty array, inconsistent row lengths |
|          | Invalid params | #VALUE! | Negative indices, both row/col = 0 |
| INDIRECT | Invalid ref | #NAME? | Malformed reference text, unknown sheet |
|          | Invalid params | #VALUE! | Empty text, unsupported R1C1 style |

## Validation Layers

### Layer 1: Type Validation (Automatic)
```python
@xl.validate_args
def FUNCTION_NAME(
    param1: func_xltypes.XlNumber,    # Automatic number validation
    param2: func_xltypes.XlText,      # Automatic text validation
    param3: func_xltypes.XlArray,     # Automatic array validation
) -> func_xltypes.XlAnything:
```

**Benefits**:
- Automatic type conversion and validation
- Consistent error messages
- Reduces boilerplate code

**Limitations**:
- Cannot handle complex validation rules
- Limited customization of error messages

### Layer 2: Parameter Validation (Function-Specific)
```python
def OFFSET(reference, rows, cols, height=None, width=None):
    try:
        # Convert and validate parameters
        rows_int = int(rows)
        cols_int = int(cols)
        
        if height is not None:
            height_int = int(height)
            if height_int < 1:
                raise xlerrors.ValueExcelError("Height must be positive")
        
        if width is not None:
            width_int = int(width)
            if width_int < 1:
                raise xlerrors.ValueExcelError("Width must be positive")
                
    except ValueError:
        raise xlerrors.ValueExcelError("Parameters must be numeric")
```

### Layer 3: Reference Validation (Utility-Based)
```python
def validate_reference_bounds(row, col):
    """Centralized bounds checking"""
    if row < 1 or row > 1048576:
        raise xlerrors.RefExcelError(f"Row {row} is out of bounds")
    if col < 1 or col > 16384:
        raise xlerrors.RefExcelError(f"Column {col} is out of bounds")
```

### Layer 4: Excel Compatibility Validation
```python
def validate_excel_compatibility(function_name, *args):
    """Ensure behavior matches Excel exactly"""
    # Check for Excel-specific edge cases
    # Validate against known Excel limitations
    # Handle Excel quirks and inconsistencies
```

## Error Handling Patterns

### Pattern 1: Graceful Degradation
```python
def FUNCTION_NAME(*args):
    try:
        # Primary logic
        return calculate_result(*args)
    except SpecificError as e:
        # Handle known error cases gracefully
        return handle_specific_error(e)
    except Exception as e:
        # Convert unexpected errors to Excel errors
        raise xlerrors.ValueExcelError(f"FUNCTION_NAME error: {str(e)}")
```

### Pattern 2: Early Validation
```python
def FUNCTION_NAME(*args):
    # Validate all parameters before processing
    validate_parameters(*args)
    validate_references(*args)
    validate_bounds(*args)
    
    # Process with confidence that inputs are valid
    return process_validated_inputs(*args)
```

### Pattern 3: Contextual Error Messages
```python
def OFFSET(reference, rows, cols, height=None, width=None):
    try:
        result = calculate_offset(reference, rows, cols, height, width)
    except xlerrors.RefExcelError:
        # Add context to reference errors
        raise xlerrors.RefExcelError(
            f"OFFSET({reference}, {rows}, {cols}) result is out of bounds"
        )
```

## Validation Functions

### Reference Validation
```python
def validate_cell_reference(ref: str) -> None:
    """Validate cell reference format and bounds"""
    try:
        row, col = ReferenceResolver.parse_cell_reference(ref)
        ReferenceResolver.validate_bounds(row, col)
    except Exception as e:
        raise xlerrors.ValueExcelError(f"Invalid cell reference: {ref}")

def validate_range_reference(ref: str) -> None:
    """Validate range reference format and bounds"""
    try:
        start, end = ReferenceResolver.parse_range_reference(ref)
        ReferenceResolver.validate_bounds(start[0], start[1])
        ReferenceResolver.validate_bounds(end[0], end[1])
    except Exception as e:
        raise xlerrors.ValueExcelError(f"Invalid range reference: {ref}")
```

### Array Validation
```python
def validate_array_structure(array) -> None:
    """Validate array has consistent structure"""
    if not hasattr(array, 'values') or not array.values:
        raise xlerrors.ValueExcelError("Array is empty or invalid")
    
    # Check for consistent row lengths
    first_row_len = len(array.values[0])
    for i, row in enumerate(array.values):
        if len(row) != first_row_len:
            raise xlerrors.ValueExcelError(
                f"Array row {i+1} has {len(row)} columns, expected {first_row_len}"
            )

def validate_array_bounds(array, row_num: int, col_num: int) -> None:
    """Validate array access is within bounds"""
    num_rows, num_cols = len(array.values), len(array.values[0])
    
    if row_num < 0 or (row_num > num_rows and row_num != 0):
        raise xlerrors.RefExcelError(f"Row {row_num} is out of range (0-{num_rows})")
    
    if col_num < 0 or (col_num > num_cols and col_num != 0):
        raise xlerrors.RefExcelError(f"Column {col_num} is out of range (0-{num_cols})")
```

### Parameter Validation
```python
def validate_positive_integer(value, param_name: str) -> int:
    """Validate parameter is positive integer"""
    try:
        int_value = int(value)
        if int_value < 1:
            raise xlerrors.ValueExcelError(f"{param_name} must be positive")
        return int_value
    except ValueError:
        raise xlerrors.ValueExcelError(f"{param_name} must be a number")

def validate_non_negative_integer(value, param_name: str) -> int:
    """Validate parameter is non-negative integer"""
    try:
        int_value = int(value)
        if int_value < 0:
            raise xlerrors.ValueExcelError(f"{param_name} must be non-negative")
        return int_value
    except ValueError:
        raise xlerrors.ValueExcelError(f"{param_name} must be a number")
```

## Error Message Standards

### Consistency Rules
1. **Function name prefix**: Include function name in error messages
2. **Parameter identification**: Clearly identify which parameter caused the error
3. **Expected vs actual**: Show what was expected vs what was provided
4. **Actionable guidance**: Suggest how to fix the error when possible

### Message Templates
```python
# Reference errors
"OFFSET result is out of bounds: row {row} exceeds maximum {max_row}"
"Invalid cell reference format: '{ref}' (expected format: A1, B2, etc.)"

# Parameter errors  
"OFFSET height must be positive, got {height}"
"INDEX row_num must be between 0 and {max_rows}, got {row_num}"

# Array errors
"INDEX array is empty or invalid"
"Array row {row} has {actual} columns, expected {expected}"
```

## Testing Strategy for Error Handling

### Unit Tests for Each Error Type
```python
class TestDynamicRangeErrors(unittest.TestCase):
    def test_offset_ref_errors(self):
        """Test OFFSET #REF! errors"""
        with self.assertRaises(xlerrors.RefExcelError):
            OFFSET("A1", -1, 0)  # Negative row
        
        with self.assertRaises(xlerrors.RefExcelError):
            OFFSET("A1", 0, -1)  # Negative column
    
    def test_offset_value_errors(self):
        """Test OFFSET #VALUE! errors"""
        with self.assertRaises(xlerrors.ValueExcelError):
            OFFSET("A1", "invalid", 1)  # Non-numeric rows
        
        with self.assertRaises(xlerrors.ValueExcelError):
            OFFSET("InvalidRef", 1, 1)  # Invalid reference
    
    def test_index_ref_errors(self):
        """Test INDEX #REF! errors"""
        array = create_test_array(3, 3)
        
        with self.assertRaises(xlerrors.RefExcelError):
            INDEX(array, 4, 1)  # Row out of bounds
        
        with self.assertRaises(xlerrors.RefExcelError):
            INDEX(array, 1, 4)  # Column out of bounds
    
    def test_indirect_name_errors(self):
        """Test INDIRECT #NAME? errors"""
        with self.assertRaises(xlerrors.NameExcelError):
            INDIRECT("InvalidRef123")  # Invalid reference format
        
        with self.assertRaises(xlerrors.NameExcelError):
            INDIRECT("")  # Empty reference
```

### Excel Compatibility Tests
```python
class TestExcelCompatibility(unittest.TestCase):
    def test_error_types_match_excel(self):
        """Verify error types match Excel exactly"""
        # Test cases derived from actual Excel behavior
        excel_test_cases = [
            ("OFFSET(A1, -1, 0)", xlerrors.RefExcelError),
            ("INDEX(A1:C3, 4, 1)", xlerrors.RefExcelError),
            ("INDIRECT('Invalid')", xlerrors.NameExcelError),
        ]
        
        for formula, expected_error in excel_test_cases:
            with self.subTest(formula=formula):
                with self.assertRaises(expected_error):
                    evaluate_formula(formula)
```

### Edge Case Testing
```python
class TestEdgeCases(unittest.TestCase):
    def test_boundary_conditions(self):
        """Test behavior at Excel limits"""
        # Test maximum row/column references
        # Test empty arrays and single-cell arrays
        # Test very large offset values
        # Test special characters in references
```

## Performance Considerations

### Validation Optimization
1. **Early exit**: Validate cheapest conditions first
2. **Caching**: Cache validation results for repeated references
3. **Lazy validation**: Only validate when needed
4. **Batch validation**: Validate multiple parameters together

### Error Object Reuse
```python
# Pre-create common error objects to avoid repeated string formatting
COMMON_ERRORS = {
    'negative_row': xlerrors.RefExcelError("Row number must be positive"),
    'negative_col': xlerrors.RefExcelError("Column number must be positive"),
    'empty_array': xlerrors.ValueExcelError("Array cannot be empty"),
}
```

## Integration with xlcalculator

### Error Propagation
- Errors should propagate correctly through formula evaluation
- Nested function calls should preserve original error context
- Error messages should be helpful for debugging

### Logging and Debugging
```python
import logging

logger = logging.getLogger('xlcalculator.dynamic_range')

def OFFSET(*args):
    logger.debug(f"OFFSET called with args: {args}")
    try:
        result = calculate_offset(*args)
        logger.debug(f"OFFSET result: {result}")
        return result
    except Exception as e:
        logger.error(f"OFFSET error: {e}")
        raise
```

## Summary

This error handling strategy ensures:

1. **Excel Compatibility**: Error types and messages match Excel behavior
2. **Robustness**: Comprehensive validation at multiple layers
3. **Usability**: Clear, actionable error messages
4. **Maintainability**: Consistent patterns and reusable validation functions
5. **Performance**: Efficient validation with minimal overhead
6. **Testability**: Comprehensive test coverage for all error conditions

The strategy balances strict validation with performance, ensuring that dynamic range functions behave exactly like their Excel counterparts while providing helpful error messages for debugging.