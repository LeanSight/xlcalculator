# XLCalculator Deduplication Strategy

## Implementation Roadmap

### Phase 1: Foundation Utilities (Week 1)

#### 1.1 Create Core Validation Module
**File**: `xlcalculator/utils/validation.py`

```python
"""Core validation utilities for Excel functions."""
import xlerrors

def validate_integer_parameter(value, param_name, min_value=None, max_value=None):
    """Validate and convert parameter to integer with bounds checking."""
    try:
        int_value = int(value)
    except (ValueError, TypeError):
        raise xlerrors.ValueExcelError(f"Invalid {param_name}: {value}")
    
    if min_value is not None and int_value < min_value:
        raise xlerrors.ValueExcelError(f"{param_name} must be >= {min_value}")
    
    if max_value is not None and int_value > max_value:
        raise xlerrors.ValueExcelError(f"{param_name} must be <= {max_value}")
    
    return int_value

def validate_positive_integer(value, param_name):
    """Validate parameter is a positive integer (>= 1)."""
    return validate_integer_parameter(value, param_name, min_value=1)

def validate_array_bounds(array_data, row_idx, col_idx, row_name="row", col_name="column"):
    """Validate array bounds for row and column indices."""
    if not array_data:
        raise xlerrors.RefExcelError("Array data is empty")
    
    if row_idx < 0 or row_idx >= len(array_data):
        raise xlerrors.RefExcelError(f"{row_name.title()} index out of range")
    
    if col_idx < 0 or col_idx >= len(array_data[0]):
        raise xlerrors.RefExcelError(f"{col_name.title()} index out of range")

def validate_dimension_parameter(value, param_name):
    """Validate parameter is a positive dimension (> 0)."""
    return validate_integer_parameter(value, param_name, min_value=1)
```

#### 1.2 Create Context Handling Decorator
**File**: `xlcalculator/utils/decorators.py`

```python
"""Decorators for Excel function standardization."""
import functools
import xlerrors

def require_context(func):
    """Decorator to ensure function has evaluator context."""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        context = kwargs.get('_context')
        if context is None:
            raise xlerrors.ValueExcelError(f"{func.__name__} requires evaluator context")
        return func(*args, **kwargs)
    return wrapper

def excel_function(func):
    """Decorator for standardized Excel function behavior."""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            # Log error for debugging while preserving Excel error types
            if not isinstance(e, (xlerrors.ExcelError,)):
                # Convert unexpected errors to ValueExcelError
                raise xlerrors.ValueExcelError(f"Error in {func.__name__}: {str(e)}")
            raise
    return wrapper
```

#### 1.3 Create Array Processing Utilities
**File**: `xlcalculator/utils/arrays.py`

```python
"""Array processing utilities for Excel functions."""
import xlerrors

class ArrayProcessor:
    """Utility class for common array operations."""
    
    @staticmethod
    def extract_array_data(reference, evaluator):
        """Extract array data from various reference types."""
        if hasattr(reference, 'value'):
            # Handle range objects
            return evaluator.evaluate(reference)
        elif isinstance(reference, (list, tuple)):
            # Handle direct array data
            return reference
        else:
            # Handle single values
            return [[reference]]
    
    @staticmethod
    def ensure_2d_array(data):
        """Ensure data is a 2D array."""
        if not isinstance(data, (list, tuple)):
            return [[data]]
        
        if not data:
            return [[]]
        
        # Check if it's already 2D
        if isinstance(data[0], (list, tuple)):
            return data
        
        # Convert 1D to 2D
        return [data]
    
    @staticmethod
    def get_array_dimensions(array_data):
        """Get dimensions of 2D array."""
        if not array_data:
            return 0, 0
        
        rows = len(array_data)
        cols = len(array_data[0]) if array_data[0] else 0
        return rows, cols
    
    @staticmethod
    def is_single_value(data):
        """Check if data represents a single value."""
        if not isinstance(data, (list, tuple)):
            return True
        
        if len(data) == 1 and isinstance(data[0], (list, tuple)) and len(data[0]) == 1:
            return True
        
        return False
```

### Phase 2: Function Refactoring (Week 2)

#### 2.1 Refactor INDEX Function
**Target**: Replace 15+ lines of validation with 3-4 utility calls

```python
# Before (15+ lines)
try:
    row_num_int = int(row_num)
except (ValueError, TypeError):
    raise xlerrors.ValueExcelError(f"Invalid row number: {row_num}")

if row_num_int < 1:
    raise xlerrors.ValueExcelError("Row number must be >= 1")

# After (2 lines)
from xlcalculator.utils.validation import validate_positive_integer
row_num_int = validate_positive_integer(row_num, "row number")
```

#### 2.2 Refactor Bounds Checking
**Target**: Replace 10+ identical bounds checks

```python
# Before (4 lines each occurrence)
if row_idx < 0 or row_idx >= len(array_data):
    raise xlerrors.RefExcelError("Row index out of range")
if col_idx < 0 or col_idx >= len(array_data[0]):
    raise xlerrors.RefExcelError("Column index out of range")

# After (1 line)
from xlcalculator.utils.validation import validate_array_bounds
validate_array_bounds(array_data, row_idx, col_idx)
```

#### 2.3 Apply Context Decorator
**Target**: Standardize context handling across 6+ functions

```python
# Before (3-4 lines each function)
if _context is None:
    raise xlerrors.ValueExcelError("INDEX requires evaluator context")
evaluator = _context.evaluator

# After (1 decorator)
from xlcalculator.utils.decorators import require_context

@require_context
def INDEX(array, row_num, col_num=None, *, _context=None):
    evaluator = _context.evaluator
    # Function logic...
```

### Phase 3: Advanced Consolidation (Week 3)

#### 3.1 Create Type Conversion Module
**File**: `xlcalculator/utils/types.py`

```python
"""Type conversion utilities for Excel functions."""
import xlerrors

class ExcelTypeConverter:
    """Handles Excel-specific type conversions."""
    
    @staticmethod
    def to_number(value, param_name="value"):
        """Convert value to number with Excel semantics."""
        if isinstance(value, (int, float)):
            return value
        
        if isinstance(value, bool):
            return 1 if value else 0
        
        if isinstance(value, str):
            try:
                return float(value)
            except ValueError:
                raise xlerrors.ValueExcelError(f"Cannot convert {param_name} to number: {value}")
        
        raise xlerrors.ValueExcelError(f"Invalid {param_name} type: {type(value)}")
    
    @staticmethod
    def to_boolean(value):
        """Convert value to boolean with Excel semantics."""
        if isinstance(value, bool):
            return value
        
        if isinstance(value, (int, float)):
            return value != 0
        
        if isinstance(value, str):
            return value.upper() in ('TRUE', '1')
        
        return False
```

#### 3.2 Create Reference Processing Module
**File**: `xlcalculator/utils/references.py`

```python
"""Reference processing utilities for Excel functions."""
import xlerrors
from .arrays import ArrayProcessor

def parse_excel_reference(reference, context, allow_single_value=True):
    """Parse and validate Excel reference with context."""
    evaluator = context.evaluator
    
    if isinstance(reference, str):
        # Parse reference string
        try:
            parsed_ref = evaluator.parse_reference(reference)
            return evaluator.evaluate(parsed_ref)
        except Exception as e:
            raise xlerrors.RefExcelError(f"Invalid reference: {reference}")
    
    elif hasattr(reference, 'address'):
        # Handle range object
        return evaluator.evaluate(reference)
    
    elif allow_single_value:
        # Handle single values
        return ArrayProcessor.ensure_2d_array(reference)
    
    else:
        raise xlerrors.RefExcelError("Reference must be a range or cell address")
```

### Phase 4: Testing and Validation (Week 4)

#### 4.1 Create Comprehensive Test Suite
**File**: `tests/utils/test_validation.py`

```python
"""Tests for validation utilities."""
import pytest
import xlerrors
from xlcalculator.utils.validation import (
    validate_integer_parameter,
    validate_positive_integer,
    validate_array_bounds
)

class TestValidation:
    def test_validate_integer_parameter_valid(self):
        assert validate_integer_parameter("5", "test") == 5
        assert validate_integer_parameter(5.0, "test") == 5
    
    def test_validate_integer_parameter_invalid(self):
        with pytest.raises(xlerrors.ValueExcelError):
            validate_integer_parameter("abc", "test")
    
    def test_validate_positive_integer(self):
        assert validate_positive_integer("5", "test") == 5
        
        with pytest.raises(xlerrors.ValueExcelError):
            validate_positive_integer("0", "test")
    
    def test_validate_array_bounds(self):
        array_data = [[1, 2, 3], [4, 5, 6]]
        
        # Valid bounds
        validate_array_bounds(array_data, 0, 0)
        validate_array_bounds(array_data, 1, 2)
        
        # Invalid bounds
        with pytest.raises(xlerrors.RefExcelError):
            validate_array_bounds(array_data, 2, 0)  # Row out of bounds
        
        with pytest.raises(xlerrors.RefExcelError):
            validate_array_bounds(array_data, 0, 3)  # Column out of bounds
```

#### 4.2 Migration Testing Strategy
1. **Parallel Testing**: Run old and new implementations side-by-side
2. **Regression Testing**: Ensure all existing tests pass
3. **Performance Testing**: Verify no performance degradation
4. **Edge Case Testing**: Test boundary conditions and error cases

### Phase 5: Deployment and Cleanup (Week 5)

#### 5.1 Gradual Migration Plan
1. **Function-by-function migration**: Start with INDEX, COLUMN, ROW
2. **Backward compatibility**: Keep old code until migration complete
3. **Monitoring**: Track for any behavioral changes
4. **Documentation**: Update function documentation

#### 5.2 Code Cleanup
1. **Remove duplicate code**: Delete old validation logic
2. **Update imports**: Standardize utility imports
3. **Code review**: Ensure consistency across all functions
4. **Performance optimization**: Profile and optimize hot paths

## Success Metrics

### Quantitative Goals
- **Code reduction**: 40-60% reduction in duplicate code
- **Line count**: Reduce dynamic_range.py by 500-800 lines
- **Test coverage**: Maintain 95%+ test coverage
- **Performance**: No degradation in function execution time

### Qualitative Goals
- **Maintainability**: Easier to add new Excel functions
- **Consistency**: Standardized error handling and validation
- **Reliability**: Reduced bugs through shared, tested utilities
- **Developer experience**: Faster development of new functions

## Risk Mitigation

### Technical Risks
1. **Breaking changes**: Comprehensive testing before deployment
2. **Performance impact**: Profile critical paths
3. **Compatibility issues**: Gradual migration with fallbacks

### Process Risks
1. **Timeline delays**: Prioritize high-impact, low-risk changes first
2. **Resource constraints**: Focus on automated testing and validation
3. **Integration issues**: Continuous integration testing

## Conclusion

This strategy provides a systematic approach to eliminating code duplication while maintaining functionality and improving code quality. The phased approach minimizes risk while delivering incremental value throughout the implementation process.