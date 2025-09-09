# Evaluator Behavior Documentation

## Critical Behavior: evaluator.evaluate() Requires Full Cell Addresses

**IMPORTANT**: The `evaluator.evaluate()` method ALWAYS requires FULL cell addresses with sheet prefix.

### Examples

✅ **Correct Usage**:
```python
evaluator.evaluate("Tests!P1")    # Returns cell content
evaluator.evaluate("Data!B2")     # Returns cell content
```

❌ **Incorrect Usage**:
```python
evaluator.evaluate("P1")          # Returns <BLANK> (invalid reference)
evaluator.evaluate("B2")          # Returns <BLANK> (invalid reference)
```

### Impact on Function Implementation

When implementing Excel functions that receive cell references as parameters, you must:

1. **Check if the parameter is a cell reference without sheet prefix**
2. **Construct the full address using current sheet context**
3. **Use evaluator.evaluate() with the full address**

### Example Implementation (INDIRECT function)

```python
def INDIRECT(ref_text, *, _context=None):
    evaluator = _context.evaluator
    ref_string = str(ref_text)
    
    # Handle cell references without sheet prefix
    if _is_valid_excel_reference(ref_string) and '!' not in ref_string:
        current_sheet = getattr(_context, 'sheet', None)
        if current_sheet:
            full_ref = f"{current_sheet}!{ref_string}"
            try:
                cell_content = evaluator.evaluate(full_ref)
                ref_string = str(cell_content)
            except Exception:
                pass  # Treat as literal string if evaluation fails
    
    # Continue with ref_string processing...
```

### Context Object

The `_context` parameter provides access to:
- `_context.evaluator`: The evaluator instance
- `_context.sheet`: Current sheet name
- `_context.address`: Current cell address
- Other evaluation context information

This behavior is consistent across all xlcalculator function implementations.