# Reference System Excel Analysis

**Document Version**: 1.0  
**Date**: 2025-09-09  
**Phase**: ATDD Red Phase - Excel Documentation Analysis  
**Application**: Reference Object System implementation for xlcalculator

---

## 🎯 Objective

Analyze official Microsoft Excel documentation and behavior for reference handling to establish the foundation for ATDD implementation of the Reference Object System.

## 📚 Official Excel Documentation Analysis

### CellReference Behavior

**Source**: [Microsoft Excel Cell References](https://support.microsoft.com/en-us/office/create-or-change-a-cell-reference-c7b8b95d-c594-4488-947e-c835649fd8d3)

#### **Core Behaviors**:
1. **A1 Reference Style**: Column letter + row number (A1, B2, Z100)
2. **Absolute References**: $ prefix for fixed references ($A$1, $A1, A$1)
3. **Sheet References**: Sheet name + ! + cell reference (Sheet1!A1)
4. **Workbook References**: [Workbook.xlsx]Sheet1!A1

#### **Coordinate System**:
- **Rows**: 1-based indexing (1 to 1,048,576)
- **Columns**: 1-based indexing (A=1, B=2, ..., XFD=16,384)
- **Case Insensitive**: A1 = a1 = A1

#### **Error Conditions**:
- **#REF!**: Reference to deleted cells or out-of-bounds
- **#NAME?**: Invalid sheet name or syntax
- **#VALUE!**: Invalid reference format

### ROW Function Analysis

**Source**: [ROW Function Documentation](https://support.microsoft.com/en-us/office/row-function-3a63b74a-c4d0-4093-b49a-e76eb49a6d8d)

#### **Syntax**: `ROW([reference])`

#### **Behaviors**:
1. **No Parameter**: Returns row number of current cell
2. **Cell Reference**: Returns row number of referenced cell
3. **Range Reference**: Returns array of row numbers
4. **Text Reference**: Accepts "A1" string format

#### **Test Cases from Excel**:
```excel
ROW()        → Current cell row (context-dependent)
ROW(A1)      → 1
ROW("A1")    → 1
ROW(A1:A3)   → {1; 2; 3}
ROW("A1:A3") → {1; 2; 3}
ROW(Sheet2!A1) → 1
```

#### **Error Cases**:
```excel
ROW("InvalidRef") → #REF!
ROW(#REF!)        → #REF!
```

### COLUMN Function Analysis

**Source**: [COLUMN Function Documentation](https://support.microsoft.com/en-us/office/column-function-44e8c754-711c-4df3-9da4-47a55042554b)

#### **Syntax**: `COLUMN([reference])`

#### **Behaviors**:
1. **No Parameter**: Returns column number of current cell
2. **Cell Reference**: Returns column number of referenced cell
3. **Range Reference**: Returns array of column numbers
4. **Text Reference**: Accepts "A1" string format

#### **Test Cases from Excel**:
```excel
COLUMN()        → Current cell column (context-dependent)
COLUMN(A1)      → 1
COLUMN("A1")    → 1
COLUMN(A1:C1)   → {1, 2, 3}
COLUMN("A1:C1") → {1, 2, 3}
COLUMN(Z1)      → 26
```

### OFFSET Function Analysis

**Source**: [OFFSET Function Documentation](https://support.microsoft.com/en-us/office/offset-function-c8de19ae-dd79-4b9b-a14e-b4d906d11b66)

#### **Syntax**: `OFFSET(reference, rows, cols, [height], [width])`

#### **Behaviors**:
1. **Single Cell Result**: When height/width omitted
2. **Range Result**: When height/width specified
3. **Reference Arithmetic**: Offset from starting reference
4. **Dynamic References**: Can reference any cell/range

#### **Test Cases from Excel**:
```excel
OFFSET(A1, 1, 1)        → B2 value
OFFSET(A1, 0, 0)        → A1 value
OFFSET(A1, 1, 1, 2, 2)  → B2:C3 range
OFFSET("A1", 1, 1)      → B2 value
```

#### **Error Cases**:
```excel
OFFSET(A1, -1, 0)       → #REF! (if goes out of bounds)
OFFSET(A1, 0, 0, 0, 1)  → #VALUE! (height cannot be 0)
OFFSET(A1, 0, 0, 1, 0)  → #VALUE! (width cannot be 0)
```

### INDIRECT Function Analysis

**Source**: [INDIRECT Function Documentation](https://support.microsoft.com/en-us/office/indirect-function-474b3a3a-8a26-4f44-b491-92b6306fa261)

#### **Syntax**: `INDIRECT(ref_text, [a1])`

#### **Behaviors**:
1. **Text to Reference**: Converts text string to reference
2. **Dynamic References**: Reference can be calculated
3. **A1 vs R1C1**: Style parameter (default A1)
4. **Cross-Sheet**: Supports sheet references

#### **Test Cases from Excel**:
```excel
INDIRECT("A1")          → A1 value
INDIRECT("A" & "1")     → A1 value (dynamic)
INDIRECT("Sheet2!A1")   → Sheet2 A1 value
INDIRECT("A1:B2")       → A1:B2 range
```

#### **Error Cases**:
```excel
INDIRECT("InvalidRef")  → #REF!
INDIRECT("")            → #REF!
INDIRECT("Sheet99!A1")  → #REF! (if sheet doesn't exist)
```

## 🔍 Excel Behavior Patterns Identified

### Pattern 1: Reference String Parsing
- Excel functions accept both direct references (A1) and string references ("A1")
- String references are parsed at function level, not AST level
- Invalid reference strings generate #REF! errors

### Pattern 2: Context-Aware Evaluation
- ROW() and COLUMN() without parameters use current cell context
- Context must be available during function execution
- Functions must have access to current cell coordinates

### Pattern 3: Reference Arithmetic
- OFFSET performs coordinate-based arithmetic
- Results can be single cells or ranges
- Bounds checking generates #REF! errors

### Pattern 4: Lazy Reference Resolution
- References preserve coordinate information
- Values are resolved only when needed
- Dynamic references support calculated coordinates

## 📋 Implementation Requirements

### Core Reference Classes Needed
1. **CellReference**: Single cell with sheet, row, column
2. **RangeReference**: Multi-cell range with start/end cells
3. **NamedReference**: Named range resolution

### Function Enhancement Requirements
1. **String Reference Parsing**: Functions must parse "A1" format
2. **Context Injection**: Access to current cell coordinates
3. **Reference Arithmetic**: Coordinate-based calculations
4. **Error Handling**: Excel-compatible error types

### AST Integration Requirements
1. **Parameter Type Detection**: Distinguish reference strings from cell values
2. **Lazy Evaluation**: Preserve reference information
3. **Context Propagation**: Pass cell context to functions

## 🎯 Success Criteria

### Functional Requirements
- ✅ ROW("A1") returns 1 (not BLANK)
- ✅ COLUMN("A1") returns 1 (not BLANK)
- ✅ OFFSET works with any Excel file without hardcoded mappings
- ✅ INDIRECT handles dynamic references correctly

### Technical Requirements
- ✅ Reference objects preserve coordinate information
- ✅ Functions receive reference strings, not evaluated values
- ✅ Context injection provides current cell coordinates
- ✅ Excel-compatible error handling (#REF!, #VALUE!, #NAME!)

### Performance Requirements
- ✅ ≤10% overhead compared to current implementation
- ✅ Lazy evaluation for large ranges
- ✅ Thread-safe context management

---

**Next Phase**: Create design document with structured test cases based on this Excel behavior analysis.