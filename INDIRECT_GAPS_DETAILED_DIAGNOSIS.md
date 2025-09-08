# INDIRECT Function Gaps - Detailed Diagnosis

## Root Cause Analysis

### Gap 1: K1 - `INDIRECT("InvalidSheet!A1")` returns BLANK instead of #REF!

**Location**: `_resolve_indirect_reference()` function, line ~235-240

**Current Behavior**:
```python
try:
    return evaluator.evaluate(ref_string)  # Fails silently
except Exception:
    try:
        return evaluator.get_cell_value(ref_string)  # Also fails silently
    except Exception:
        raise xlerrors.RefExcelError(f"Invalid reference: {ref_string}")  # Should reach here
```

**Problem**: The exception is being caught somewhere and converted to BLANK instead of propagating the RefExcelError.

**Expected**: Should return `xlerrors.RefExcelError` which displays as `#REF!`

### Gap 2: I4 - `INDIRECT(P3)` where P3="Sheet Error" returns Array instead of #REF!

**Location**: `_resolve_indirect_reference()` function, lines 225-228

**Current Behavior**:
```python
elif ref_string == "Sheet Error":
    # Special case for P3 test - return placeholder Array
    # This is a workaround for test compatibility when IFERROR is not implemented
    return func_xltypes.Array([[0]])
```

**Problem**: Hardcoded workaround that returns Array instead of proper error handling.

**Expected**: "Sheet Error" is not a valid Excel reference, should return `xlerrors.RefExcelError`

## Technical Issues Identified

### 1. Hardcoded Workarounds
The function contains several hardcoded special cases that mask proper error handling:

```python
# Line 220-223
if ref_string in ["Not Found", ""]:
    return 25

# Line 225-228  
elif ref_string == "Sheet Error":
    return func_xltypes.Array([[0]])
```

### 2. Missing Reference Format Validation
No validation that the reference string is a valid Excel reference format before attempting resolution.

### 3. Error Propagation Issues
RefExcelError exceptions are being caught and converted to BLANK somewhere in the evaluation chain.

### 4. Sheet Existence Validation Missing
No check if referenced sheets exist in the model before attempting to resolve references.

## Specific Fixes Required

### Fix 1: Remove Hardcoded Workarounds
```python
def _resolve_indirect_reference(ref_string, evaluator):
    # Remove these hardcoded cases:
    # if ref_string in ["Not Found", ""]: return 25
    # elif ref_string == "Sheet Error": return func_xltypes.Array([[0]])
    
    # Add proper validation instead
    if not _is_valid_reference_format(ref_string):
        return xlerrors.RefExcelError("Invalid reference format")
```

### Fix 2: Add Sheet Validation
```python
def _validate_sheet_reference(ref_string, evaluator):
    if '!' in ref_string:
        sheet_name = ref_string.split('!')[0]
        # Check if sheet exists in model
        available_sheets = [cell.split('!')[0] for cell in evaluator.model.cells.keys() if '!' in cell]
        if sheet_name not in available_sheets:
            return xlerrors.RefExcelError("Sheet does not exist")
    return None
```

### Fix 3: Add Reference Format Validation
```python
def _is_valid_reference_format(ref_string):
    """Check if string is a valid Excel reference format"""
    import re
    
    # Basic patterns for Excel references
    cell_pattern = r'^[A-Z]+[0-9]+$'  # A1, B2, etc.
    range_pattern = r'^[A-Z]+[0-9]+:[A-Z]+[0-9]+$'  # A1:B2
    sheet_cell_pattern = r'^[^!]+![A-Z]+[0-9]+$'  # Sheet!A1
    sheet_range_pattern = r'^[^!]+![A-Z]+[0-9]+:[A-Z]+[0-9]+$'  # Sheet!A1:B2
    
    return (re.match(cell_pattern, ref_string) or 
            re.match(range_pattern, ref_string) or
            re.match(sheet_cell_pattern, ref_string) or
            re.match(sheet_range_pattern, ref_string))
```

### Fix 4: Ensure Error Propagation
```python
def INDIRECT(ref_text, a1=True):
    # Remove the BLANK handling workarounds
    # Let errors propagate properly
    
    ref_string = str(ref_text)
    
    # Validate format first
    if not _is_valid_reference_format(ref_string):
        return xlerrors.RefExcelError("Invalid reference format")
    
    # Validate sheet exists
    sheet_error = _validate_sheet_reference(ref_string, evaluator)
    if sheet_error:
        return sheet_error
    
    # Resolve reference
    return _resolve_indirect_reference(ref_string, evaluator)
```

## Test Cases to Verify Fixes

```python
def test_indirect_error_cases():
    # Invalid sheet reference
    result = evaluator.evaluate('INDIRECT("InvalidSheet!A1")')
    assert isinstance(result, xlerrors.RefExcelError)
    assert str(result) == "#REF!"
    
    # Invalid reference format
    result = evaluator.evaluate('INDIRECT("Sheet Error")')
    assert isinstance(result, xlerrors.RefExcelError)
    assert str(result) == "#REF!"
    
    # Empty reference
    result = evaluator.evaluate('INDIRECT("")')
    assert isinstance(result, xlerrors.RefExcelError)
    assert str(result) == "#REF!"
```

## Implementation Priority

1. **Immediate**: Remove hardcoded "Sheet Error" workaround (fixes I4)
2. **Immediate**: Add sheet existence validation (fixes K1)  
3. **High**: Add reference format validation
4. **Medium**: Ensure proper error propagation throughout evaluation chain

## Impact Assessment

- **Breaking Changes**: None for valid references
- **Test Compatibility**: Will require updating tests that expect workaround behavior
- **Excel Fidelity**: Significant improvement in error handling accuracy