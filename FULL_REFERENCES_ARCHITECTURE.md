# Full Column/Row References Architecture Design

## Current Flow (Working)
```
"Data!A:A" 
    ↓ Tokenizer
RangeNode(tvalue='Data!A:A', ttype='operand', tsubtype='range')
    ↓ RangeNode.eval()
DataFrame with column A data: ['Name', 'Alice', 'Bob', 'Charlie', 'Diana', 'Eve']
    ↓ OFFSET function
❌ FAILS: Tries to parse DataFrame as cell reference
```

## Problem Analysis

### ✅ What Works
1. **Tokenizer/Parser**: Correctly recognizes `Data!A:A` as RangeNode
2. **RangeNode.eval()**: Returns correct DataFrame with full column data
3. **ExcelCompliantLazyRange**: Properly resolves full column references

### ❌ What Fails
**OFFSET function logic**:
```python
# Current problematic logic in OFFSET
if isinstance(reference, (str, func_xltypes.Text)):
    ref_string = str(reference)
    try:
        start_ref = CellReference.parse(ref_string)  # ❌ FAILS for "Data!A:A"
    except xlerrors.RefExcelError:
        # ❌ WRONG: Treats "Data!A:A" as a value to search for
        found_address = _find_cell_address_for_value(ref_string, evaluator)
```

## Solution Design

### 1. Enhanced Reference Detection in OFFSET
```python
def OFFSET(reference, rows, cols, height=None, width=None, *, _context=None):
    # Handle different reference types
    if isinstance(reference, func_xltypes.Array):
        # DataFrame from RangeNode.eval() - this is a full column/row reference
        start_ref = _create_reference_from_dataframe(reference, _context)
    elif isinstance(reference, (str, func_xltypes.Text)):
        ref_string = str(reference)
        
        # Check if it's a full column/row reference pattern
        if _is_full_column_or_row_reference(ref_string):
            start_ref = _parse_full_reference(ref_string)
        else:
            # Try normal cell reference
            try:
                start_ref = CellReference.parse(ref_string)
            except xlerrors.RefExcelError:
                # Value from INDEX - existing logic
                found_address = _find_cell_address_for_value(ref_string, evaluator)
                start_ref = CellReference.parse(found_address)
```

### 2. Full Reference Types
```python
class FullColumnReference(ReferenceBase):
    def __init__(self, sheet: str, column: str):
        self.sheet = sheet
        self.column = column
        self.row = 1  # Start at row 1
        
    def offset(self, rows_offset: int, cols_offset: int):
        # Calculate new position
        new_col_num = _column_letter_to_number(self.column) + cols_offset
        new_row = self.row + rows_offset
        
        # Return regular CellReference for the offset position
        new_col_letter = _number_to_column_letter(new_col_num)
        return CellReference(self.sheet, new_row, new_col_num)

class FullRowReference(ReferenceBase):
    def __init__(self, sheet: str, row: int):
        self.sheet = sheet
        self.row = row
        self.column = 1  # Start at column A
        
    def offset(self, rows_offset: int, cols_offset: int):
        # Similar logic for row references
        pass
```

### 3. DataFrame to Reference Conversion
```python
def _create_reference_from_dataframe(dataframe, context):
    """
    Convert DataFrame from RangeNode.eval() back to reference.
    
    This is needed because OFFSET receives the evaluated DataFrame
    but needs to know the original reference to calculate offsets.
    """
    # Get the original range address from context
    # This requires tracking the original reference through evaluation
    original_ref = context.get_original_reference()  # Need to implement
    
    if _is_full_column_or_row_reference(original_ref):
        return _parse_full_reference(original_ref)
    else:
        # Regular range reference
        return RangeReference.parse(original_ref)
```

## Expected Behavior
```
OFFSET(Data!A:A, 1, 0, 3, 1)
    ↓
reference = DataFrame(['Name', 'Alice', 'Bob', ...])
original_ref = "Data!A:A" (from context)
start_ref = FullColumnReference(sheet='Data', column='A')
offset_ref = start_ref.offset(1, 0) → CellReference('Data', 2, 1) = "Data!A2"
target_range = build_range("Data!A2", height=3, width=1) = "Data!A2:A4"
result = ['Alice', 'Bob', 'Charlie']
```

## Implementation Strategy

### Phase 1: Quick Fix (Low Risk)
- Detect DataFrame input in OFFSET
- Add pattern detection for full references
- Handle `Data!A:A` pattern specifically

### Phase 2: Proper Architecture (Higher Risk)
- Implement FullColumnReference/FullRowReference classes
- Add context tracking for original references
- Full reference system integration

## Risk Assessment
- **Phase 1**: Low risk, targeted fix for specific pattern
- **Phase 2**: Medium risk, requires context system changes
- **Backward Compatibility**: Maintained in both phases

## Recommendation
Start with **Phase 1** for immediate fix, then consider Phase 2 for completeness.