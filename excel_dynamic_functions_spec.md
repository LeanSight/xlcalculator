# Excel Dynamic Range Functions Specification

## Core Dynamic Range Functions

### 1. OFFSET Function
**Syntax**: `OFFSET(reference, rows, cols, [height], [width])`
**Purpose**: Returns a reference to a range offset from a starting reference

**Parameters**:
- `reference`: Starting cell/range reference
- `rows`: Number of rows to offset (positive = down, negative = up)
- `cols`: Number of columns to offset (positive = right, negative = left)  
- `height`: [Optional] Height of returned range (default = height of reference)
- `width`: [Optional] Width of returned range (default = width of reference)

**Examples**:
- `OFFSET(A1, 1, 1)` → Returns reference to B2
- `OFFSET(A1:B2, 1, 1)` → Returns reference to B2:C3
- `OFFSET(A1, 1, 1, 2, 2)` → Returns reference to B2:C3

**Special Syntax**: `:OFFSET(...)` - Excel allows colon prefix for range operations

### 2. INDEX Function
**Syntax**: `INDEX(array, row_num, [col_num])`
**Purpose**: Returns value/reference at intersection of row and column

**Parameters**:
- `array`: Range of cells or array
- `row_num`: Row number (1-based, 0 = entire column)
- `col_num`: [Optional] Column number (1-based, 0 = entire row)

**Examples**:
- `INDEX(A1:C3, 2, 2)` → Returns value at B2
- `INDEX(A1:C3, 0, 2)` → Returns entire column B as array
- `INDEX(A1:C3, 2, 0)` → Returns entire row 2 as array

**Special Syntax**: `:INDEX(...)` - Similar to OFFSET

### 3. INDIRECT Function  
**Syntax**: `INDIRECT(ref_text, [a1])`
**Purpose**: Returns reference specified by text string

**Parameters**:
- `ref_text`: Text string containing cell reference
- `a1`: [Optional] TRUE = A1 style, FALSE = R1C1 style (default TRUE)

**Examples**:
- `INDIRECT("B2")` → Returns reference to B2
- `INDIRECT("Sheet2!A1")` → Returns reference to A1 on Sheet2
- `INDIRECT(A1)` → If A1 contains "B2", returns reference to B2

## Implementation Patterns

### Pattern 1: Reference Resolution
All dynamic functions need to:
1. Parse cell/range references (A1, B2:C3, etc.)
2. Convert to internal coordinates (row, col)
3. Apply transformations (offset, index, etc.)
4. Return new reference or value

### Pattern 2: Array Handling
Functions can return:
- **Single values**: `INDEX(A1:C3, 2, 2)` → single cell value
- **Range references**: `OFFSET(A1, 1, 1, 2, 2)` → range B2:C3
- **Array values**: `INDEX(A1:C3, 0, 2)` → column array

### Pattern 3: Error Handling
Common errors:
- `#REF!`: Reference out of bounds
- `#VALUE!`: Invalid parameters
- `#NAME?`: Invalid reference text (INDIRECT)

## Architecture Considerations

### Option A: Parser-Level Resolution (Current OFFSET/INDEX approach)
- **Pros**: Can resolve references at parse time
- **Cons**: Complex parser logic, hard to extend
- **Current State**: Partially implemented for :OFFSET/:INDEX

### Option B: Function-Level Resolution (Standard approach)
- **Pros**: Clean separation, easy to extend, standard pattern
- **Cons**: Need reference resolution utilities
- **Pattern**: Like current CHOOSE, MATCH, VLOOKUP

### Option C: Hybrid Approach
- **Simple cases**: Function-level (INDIRECT, basic INDEX)
- **Complex syntax**: Parser-level (:OFFSET, :INDEX)
- **Pros**: Best of both worlds
- **Cons**: Complexity in maintaining two patterns

## Recommendation: Option B (Function-Level)
**Rationale**: 
1. **Consistency**: Matches existing VLOOKUP, MATCH, CHOOSE pattern
2. **Maintainability**: Clear separation of concerns
3. **Extensibility**: Easy to add new functions
4. **Testability**: Each function can be tested independently
5. **Simplicity**: Remove complex parser preprocessing

**Implementation Strategy**:
1. Create reference resolution utilities
2. Implement functions using standard @register pattern
3. Remove/simplify parser preprocessing for :OFFSET/:INDEX
4. Add comprehensive test coverage