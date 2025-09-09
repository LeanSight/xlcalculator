# Context Injection System Guide

## Overview

The xlcalculator context injection system provides Excel functions with access to their execution context (current cell coordinates, evaluator instance, etc.) without relying on global variables. This system is thread-safe, performant, and easy to extend.

## Architecture

### Core Components

1. **CellContext**: Data class containing cell coordinates and evaluator access
2. **Context Creation**: Factory functions for creating context objects
3. **Context Injection**: Automatic parameter injection during function execution
4. **Function Registration**: System for marking functions that need context

### Key Benefits

- **Thread Safety**: No global state variables
- **Performance**: Cached context creation and fast function lookup
- **Maintainability**: Clear separation of concerns
- **Extensibility**: Easy to add context to new functions

## Using Context in Functions

### Basic Pattern

```python
from xlcalculator.xlfunctions import xl
from xlcalculator.context import context_aware

@xl.register()
@context_aware  # Automatically registers for context injection
def MY_FUNCTION(arg1, arg2, *, _context=None):
    """Example function using context injection."""
    if _context is None:
        raise ValueError("Function requires context")
    
    # Access current cell information
    current_row = _context.row
    current_col = _context.column
    current_address = _context.address
    
    # Access evaluator for additional operations
    evaluator = _context.evaluator
    other_cell_value = evaluator.get_cell_value("A1")
    
    return f"Cell {current_address} at row {current_row}, col {current_col}"
```

### Manual Registration

```python
from xlcalculator.context import register_context_function

@xl.register()
def ANOTHER_FUNCTION(value, *, _context=None):
    """Function with manual context registration."""
    # Function implementation
    pass

# Register manually
register_context_function('ANOTHER_FUNCTION')
```

## Context Properties

### CellContext Class

```python
@dataclass
class CellContext:
    cell: 'XLCell'                    # Current cell being evaluated
    evaluator: 'Evaluator'            # Evaluator instance
    
    @property
    def row(self) -> int:             # Row number (1-based)
    
    @property  
    def column(self) -> int:          # Column number (1-based)
    
    @property
    def address(self) -> str:         # Full address (e.g., 'Sheet1!A1')
    
    @property
    def sheet(self) -> str:           # Sheet name
    
    def get_cell_value(self, address: str)  # Get value of any cell
    def evaluate(self, address: str)        # Evaluate any cell
```

### Available Properties

| Property | Type | Description | Example |
|----------|------|-------------|---------|
| `row` | int | 1-based row number | `3` for row 3 |
| `column` | int | 1-based column number | `2` for column B |
| `address` | str | Full cell address | `"Sheet1!B3"` |
| `sheet` | str | Sheet name | `"Sheet1"` |
| `cell` | XLCell | Raw cell object | Cell instance |
| `evaluator` | Evaluator | Evaluator instance | Evaluator instance |

## Implementation Examples

### ROW() Function

```python
@xl.register()
@xl.validate_args
def ROW(reference: func_xltypes.XlAnything = None, *, _context=None) -> func_xltypes.XlAnything:
    """Returns the row number of a reference."""
    
    if reference is None:
        # Return row number of current cell
        if _context is not None:
            return _context.row
        else:
            raise xlerrors.ValueExcelError("ROW() without reference requires current cell context")
    
    # Handle explicit reference parameter
    # ... implementation for explicit references
```

### COLUMN() Function

```python
@xl.register()
@xl.validate_args
def COLUMN(reference: func_xltypes.XlAnything = None, *, _context=None) -> func_xltypes.XlNumber:
    """Returns the column number of a reference."""
    
    if reference is None:
        # Return column number of current cell
        if _context is not None:
            return _context.column
        else:
            raise xlerrors.ValueExcelError("COLUMN() without reference requires current cell context")
    
    # Handle explicit reference parameter
    # ... implementation for explicit references
```

### INDEX() Function

```python
@xl.register()
@xl.validate_args
def INDEX(array, row_num, col_num=1, area_num=1, *, _context=None):
    """Returns value at intersection of row/column in array."""
    
    if _context is None:
        raise xlerrors.ValueExcelError("INDEX function requires evaluator context")
    
    evaluator = _context.evaluator
    
    # Use evaluator to resolve array references
    if hasattr(array, 'values'):
        array_data = array.values.tolist()
    else:
        array_data = evaluator.get_range_values(str(array))
    
    # ... rest of implementation
```

## Performance Optimizations

### Fast Function Lookup

The system uses O(1) set lookup instead of expensive signature inspection:

```python
# Fast lookup by function name
if needs_context_by_name(func.__name__):
    # Inject context
```

### Context Caching

Context objects are cached to avoid repeated creation:

```python
# Cached context creation
context = create_context_cached(cell, evaluator)
```

### LRU Caching

Function context detection is cached:

```python
@lru_cache(maxsize=256)
def needs_context(func) -> bool:
    # Cached signature inspection
```

## Error Handling

### Best Practices

1. **Always check for context availability**:
   ```python
   if _context is None:
       raise xlerrors.ValueExcelError("Function requires context")
   ```

2. **Use appropriate Excel error types**:
   ```python
   # For missing context
   raise xlerrors.ValueExcelError("Context required")
   
   # For invalid references
   raise xlerrors.RefExcelError("Invalid reference")
   ```

3. **Provide clear error messages**:
   ```python
   raise xlerrors.ValueExcelError("ROW() without reference requires current cell context")
   ```

## Testing Context-Aware Functions

### Unit Testing

```python
def test_my_function_with_context():
    # Create test context
    from xlcalculator.context import create_context
    from xlcalculator.model import Model
    
    model = Model()
    model.set_cell_value('Sheet1!A1', 'test')
    evaluator = Evaluator(model)
    test_cell = evaluator.model.cells['Sheet1!A1']
    
    context = create_context(test_cell, evaluator)
    
    # Test function with context
    result = MY_FUNCTION(arg1, arg2, _context=context)
    
    assert result == expected_value
```

### Integration Testing

```python
def test_function_in_evaluator():
    # Test through evaluator (automatic context injection)
    compiler = ModelCompiler()
    model = compiler.read_and_parse_dict({
        'Sheet1': {
            'A1': '=MY_FUNCTION("test", 123)'
        }
    })
    
    evaluator = Evaluator(model)
    result = evaluator.evaluate('Sheet1!A1')
    
    assert result == expected_value
```

## Migration Guide

### From Global Context

If you have functions using the old global context system:

**Before:**
```python
def OLD_FUNCTION():
    evaluator = _get_evaluator_context()
    current_cell = _get_current_cell_context()
    # ... implementation
```

**After:**
```python
@context_aware
def NEW_FUNCTION(*, _context=None):
    if _context is None:
        raise xlerrors.ValueExcelError("Function requires context")
    
    evaluator = _context.evaluator
    current_address = _context.address
    # ... implementation
```

### Registration Changes

**Before:**
```python
# Manual global context setup
_set_evaluator_context(evaluator, cell_address)
```

**After:**
```python
# Automatic context injection
register_context_function('FUNCTION_NAME')
# or use @context_aware decorator
```

## Advanced Usage

### Custom Context Types

For specialized functions, you can extend the context system:

```python
@dataclass
class SpecializedContext(CellContext):
    """Extended context for specialized functions."""
    
    def custom_method(self):
        # Custom functionality
        pass
```

### Conditional Context

Some functions may optionally use context:

```python
@xl.register()
def FLEXIBLE_FUNCTION(value, *, _context=None):
    """Function that works with or without context."""
    
    if _context is not None:
        # Enhanced behavior with context
        current_sheet = _context.sheet
        # ... context-aware implementation
    else:
        # Basic behavior without context
        # ... fallback implementation
```

### Performance Monitoring

Monitor context system performance:

```python
from xlcalculator.context import get_registered_context_functions

# Check registered functions
context_functions = get_registered_context_functions()
print(f"Functions using context: {len(context_functions)}")

# Clear cache when needed
from xlcalculator.context import clear_context_cache
clear_context_cache()
```

## Troubleshooting

### Common Issues

1. **Missing Context Error**:
   ```
   ValueError: Function requires context
   ```
   **Solution**: Ensure function is registered with `@context_aware` or `register_context_function()`

2. **Context Not Injected**:
   ```
   _context is None in function
   ```
   **Solution**: Check function signature has `*, _context=None` parameter

3. **Performance Issues**:
   **Solution**: Use `clear_context_cache()` periodically for long-running evaluations

### Debug Tips

1. **Check Registration**:
   ```python
   from xlcalculator.context import needs_context_by_name
   print(needs_context_by_name('MY_FUNCTION'))  # Should be True
   ```

2. **Verify Context Properties**:
   ```python
   def debug_context(*, _context=None):
       if _context:
           print(f"Row: {_context.row}, Col: {_context.column}")
           print(f"Address: {_context.address}")
   ```

## Best Practices

### Function Design

1. **Use keyword-only context parameter**: `*, _context=None`
2. **Check context availability early**: First thing in function
3. **Use appropriate error types**: Excel-compatible error handling
4. **Document context requirements**: Clear docstrings

### Performance

1. **Register functions properly**: Use `@context_aware` decorator
2. **Cache context when possible**: Use `create_context_cached()`
3. **Clear cache periodically**: For long-running operations
4. **Avoid unnecessary context access**: Only access what you need

### Maintainability

1. **Follow consistent patterns**: Use established function templates
2. **Separate context logic**: Keep context handling separate from business logic
3. **Test with and without context**: Ensure proper error handling
4. **Document context usage**: Clear examples and patterns

## Conclusion

The context injection system provides a robust, performant, and maintainable way to give Excel functions access to their execution context. By following the patterns and best practices outlined in this guide, you can create context-aware functions that are thread-safe, efficient, and easy to maintain.