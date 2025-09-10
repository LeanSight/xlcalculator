# Multiple Areas Architecture Design

## Current Flow (Working)
```
"(Data!A1:A5, Data!C1:C5)" 
    ↓ Tokenizer
[subexpression_start, range1, union_operator, range2, subexpression_stop]
    ↓ Parser  
OperatorNode(tvalue=',', left=RangeNode1, right=RangeNode2)
    ↓ OP_UNION
(DataFrame1, DataFrame2)  # Tuple of pandas DataFrames
```

## Problem
INDEX function receives tuple but doesn't handle it correctly:

```python
# Current INDEX logic
if (hasattr(array, '__iter__') and 
    not isinstance(array, (str, func_xltypes.Array)) and 
    not hasattr(array, 'values') and
    len(array) > 0 and
    isinstance(array[0], str)):  # ❌ FAILS HERE
```

The condition `isinstance(array[0], str)` fails because `array[0]` is a pandas DataFrame, not a string.

## Solution Design

### 1. Update Multiple Areas Detection
```python
def _is_multiple_areas(array):
    """Detect if array parameter represents multiple areas."""
    return (isinstance(array, tuple) and 
            len(array) > 0 and
            all(hasattr(area, 'values') or isinstance(area, str) for area in array))
```

### 2. Update INDEX Function Logic
```python
def INDEX(array, row_num, col_num=1, area_num=1, *, _context=None):
    if _is_multiple_areas(array):
        # Handle multiple areas (tuple of DataFrames or strings)
        areas = array
        
        # Validate area_num
        area_num_int = int(area_num)
        if area_num_int < 1 or area_num_int > len(areas):
            raise xlerrors.RefExcelError("Area number out of range")
        
        # Select the specified area (1-based index)
        selected_area = areas[area_num_int - 1]
        
        # Convert to array data
        if hasattr(selected_area, 'values'):
            array_data = selected_area.values.tolist()
        else:
            array_data = evaluator.get_range_values(str(selected_area))
    else:
        # Handle single area (existing logic)
        # ...
```

## Expected Behavior
```
INDEX((Data!A1:A5, Data!C1:C5), 2, 1, 2)
    ↓
areas = (DataFrame1, DataFrame2)
area_num = 2 → select DataFrame2 
DataFrame2.values = [['City'], ['NYC'], ['LA'], ['Chicago'], ['Miami']]
row=2, col=1 → array_data[1][0] = 'NYC'
    ↓
Result: 'NYC'
```

## Implementation Strategy
1. ✅ No parser changes needed (already working)
2. ✅ No AST changes needed (OP_UNION works)
3. ❌ Fix INDEX function multiple areas detection
4. ❌ Update INDEX function area selection logic

## Risk Assessment
- **Low Risk**: Only changes INDEX function logic
- **No Breaking Changes**: Existing single-area functionality unchanged
- **Backward Compatible**: All current tests should still pass