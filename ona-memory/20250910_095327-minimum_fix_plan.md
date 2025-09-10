# Minimum Fix Plan for ArrayProcessor.extract_array_data()

## ATDD Approach: Minimum Changes to Turn RED → GREEN

### Current Broken Code
```python
@staticmethod
def extract_array_data(reference, evaluator):
    if hasattr(reference, 'value'):           # BUG: Should be 'values'
        return evaluator.evaluate(reference)  # BUG: Wrong method for pandas
    elif hasattr(reference, 'address'):
        return evaluator.evaluate(reference)  # OK for range objects
    elif isinstance(reference, (list, tuple)):
        return ArrayProcessor.ensure_2d_array(reference)  # OK
    else:
        return [[reference]]                   # BUG: Strings as single values
```

### Minimum Fix Implementation
```python
@staticmethod
def extract_array_data(reference, evaluator):
    # Fix 1: Correct pandas DataFrame detection and handling
    if hasattr(reference, 'values'):
        return reference.values.tolist()
    # Keep working range object handling
    elif hasattr(reference, 'address'):
        return evaluator.evaluate(reference)
    # Fix 2: Add string detection for range references
    elif isinstance(reference, str):
        return evaluator.get_range_values(reference)
    # Keep working array handling
    elif isinstance(reference, (list, tuple)):
        return ArrayProcessor.ensure_2d_array(reference)
    # Keep working single value handling (for non-strings)
    else:
        return [[reference]]
```

### Changes Made
1. **Line 19**: `hasattr(reference, 'value')` → `hasattr(reference, 'values')`
2. **Line 20**: `evaluator.evaluate(reference)` → `reference.values.tolist()`
3. **Added**: String detection before line 29: `isinstance(reference, str)`
4. **Added**: `evaluator.get_range_values(reference)` for strings

### Tests That Should Pass After Fix
- ✅ test_string_range_reference_extraction (Fix 2)
- ✅ test_pandas_dataframe_extraction (Fix 1)
- ✅ test_integration_with_index_function (Fix 2)
- ✅ test_direct_array_preservation (already working)
- ✅ test_range_object_evaluation (already working)
- ✅ test_single_value_handling (already working)

### Preserved Working Logic
- Range objects with address attribute
- Direct arrays and tuples
- Single non-string values
- All existing helper methods unchanged