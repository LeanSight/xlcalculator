# Dynamic Range Functions Architecture

## Overview

This document outlines the clean architecture for implementing Excel dynamic range functions in xlcalculator using the Function-Level approach.

## Current State Analysis

**Implemented (3/12)**: CHOOSE, MATCH, VLOOKUP  
**Missing Core (3)**: OFFSET, INDEX, INDIRECT  
**Missing Advanced (2)**: HLOOKUP, TRANSPOSE  
**Missing Modern (4)**: XLOOKUP, FILTER, SORT, UNIQUE  

## Architecture Principles

### 1. Consistency
- Follow existing xlcalculator patterns (`@xl.register()`, `@xl.validate_args`)
- Use standard `func_xltypes` for type validation
- Maintain consistent error handling with `xlerrors`

### 2. Separation of Concerns
```
┌─────────────────────────────────────────────────────────────┐
│                    Function Layer                           │
│  ┌─────────────┐ ┌─────────────┐ ┌─────────────┐           │
│  │   OFFSET    │ │    INDEX    │ │  INDIRECT   │           │
│  └─────────────┘ └─────────────┘ └─────────────┘           │
└─────────────────────────────────────────────────────────────┘
┌─────────────────────────────────────────────────────────────┐
│                   Utilities Layer                          │
│  ┌─────────────┐ ┌─────────────┐ ┌─────────────┐           │
│  │ Reference   │ │   Range     │ │   Error     │           │
│  │  Parser     │ │ Calculator  │ │  Handler    │           │
│  └─────────────┘ └─────────────┘ └─────────────┘           │
└─────────────────────────────────────────────────────────────┘
┌─────────────────────────────────────────────────────────────┐
│                 xlcalculator Core                          │
│        (Evaluator, Parser, AST, Type System)               │
└─────────────────────────────────────────────────────────────┘
```

### 3. Modularity
- **Reference utilities**: Reusable across all dynamic functions
- **Function implementations**: Independent, testable units
- **Error handling**: Centralized validation and error reporting

## Component Design

### Reference Resolution Utilities

```python
# xlcalculator/xlfunctions/reference_utils.py

class ReferenceResolver:
    """Centralized reference resolution for dynamic range functions"""
    
    @staticmethod
    def parse_cell_reference(ref: str) -> Tuple[int, int]:
        """Parse A1 notation to (row, col) coordinates"""
        
    @staticmethod
    def parse_range_reference(ref: str) -> Tuple[Tuple[int, int], Tuple[int, int]]:
        """Parse A1:B2 notation to coordinate pairs"""
        
    @staticmethod
    def cell_to_string(row: int, col: int) -> str:
        """Convert coordinates to A1 notation"""
        
    @staticmethod
    def range_to_string(start: Tuple[int, int], end: Tuple[int, int]) -> str:
        """Convert coordinate pairs to A1:B2 notation"""
        
    @staticmethod
    def offset_reference(ref: str, rows: int, cols: int, 
                        height: int = None, width: int = None) -> str:
        """Apply offset to reference and return new reference"""
        
    @staticmethod
    def validate_bounds(row: int, col: int, max_row: int = 1048576, 
                       max_col: int = 16384) -> None:
        """Validate reference is within Excel bounds"""
```

### Function Interface Design

```python
# Standard function signature pattern

@xl.register()
@xl.validate_args
def FUNCTION_NAME(
    param1: func_xltypes.XlType,
    param2: func_xltypes.XlType = default_value,
    ...
) -> func_xltypes.XlAnything:
    """
    Excel-compatible docstring with examples
    
    Args:
        param1: Description
        param2: Optional description
        
    Returns:
        Result type description
        
    Raises:
        RefExcelError: When reference is out of bounds
        ValueExcelError: When parameters are invalid
        
    Examples:
        FUNCTION_NAME(arg1, arg2) → expected_result
    """
    try:
        # 1. Parameter validation and conversion
        # 2. Core logic using utilities
        # 3. Result formatting and return
    except SpecificError:
        # 4. Specific error handling
    except Exception as e:
        # 5. Generic error handling
        raise xlerrors.ValueExcelError(f"FUNCTION_NAME error: {str(e)}")
```

## Implementation Strategy

### Phase 1: Foundation (Priority: HIGH)
**Goal**: Core reference utilities and OFFSET function

**Components**:
1. `reference_utils.py` - Reference resolution utilities
2. `OFFSET` function implementation
3. Comprehensive unit tests
4. Integration with existing evaluator

**Success Criteria**:
- `OFFSET(A1, 1, 1)` → `"B2"`
- `OFFSET(A1:B2, 1, 1)` → `"B2:C3"`
- Proper error handling for out-of-bounds references
- 100% test coverage for utilities

### Phase 2: Core Functions (Priority: HIGH)
**Goal**: INDEX and INDIRECT functions

**Components**:
1. `INDEX` function with array support
2. `INDIRECT` function with text reference resolution
3. Integration tests with complex formulas
4. Performance validation

**Success Criteria**:
- `INDEX(A1:C3, 2, 2)` → value at B2
- `INDEX(A1:C3, 0, 2)` → entire column B
- `INDIRECT("B2")` → reference to B2
- `INDIRECT("Sheet2!A1")` → cross-sheet reference

### Phase 3: Advanced Functions (Priority: MEDIUM)
**Goal**: HLOOKUP and TRANSPOSE

**Components**:
1. `HLOOKUP` function (horizontal lookup)
2. `TRANSPOSE` function (array transposition)
3. Extended test coverage
4. Documentation updates

### Phase 4: Modern Functions (Priority: LOW)
**Goal**: Excel 365 functions

**Components**:
1. `XLOOKUP` (modern lookup)
2. `FILTER`, `SORT`, `UNIQUE` (array functions)
3. Performance optimization
4. Advanced Excel compatibility

## File Organization

```
xlcalculator/xlfunctions/
├── __init__.py
├── xl.py                    # Function registry (existing)
├── xlerrors.py             # Error types (existing)
├── func_xltypes.py         # Type system (existing)
├── lookup.py               # VLOOKUP, MATCH, CHOOSE (existing)
├── reference_utils.py      # NEW: Reference resolution utilities
├── dynamic_range.py        # NEW: OFFSET, INDEX, INDIRECT
├── advanced_lookup.py      # NEW: HLOOKUP, TRANSPOSE
└── modern_functions.py     # NEW: XLOOKUP, FILTER, SORT, UNIQUE
```

## Error Handling Strategy

### Error Types and Usage
```python
# Reference errors - out of bounds, invalid format
raise xlerrors.RefExcelError("Reference out of bounds")

# Value errors - invalid parameters, type mismatches  
raise xlerrors.ValueExcelError("Invalid parameter value")

# Name errors - invalid reference text
raise xlerrors.NameExcelError("Invalid reference name")

# Not available errors - lookup failures
raise xlerrors.NaExcelError("Value not found")
```

### Validation Layers
1. **Type validation**: `@xl.validate_args` decorator
2. **Range validation**: Reference bounds checking
3. **Parameter validation**: Function-specific rules
4. **Excel compatibility**: Match Excel error behavior

## Testing Strategy

### Unit Tests
```python
class TestReferenceUtils(unittest.TestCase):
    def test_parse_cell_reference(self)
    def test_parse_range_reference(self)
    def test_cell_to_string(self)
    def test_offset_reference(self)
    def test_validate_bounds(self)

class TestDynamicRangeFunctions(unittest.TestCase):
    def test_offset_basic(self)
    def test_offset_range(self)
    def test_offset_errors(self)
    def test_index_single_value(self)
    def test_index_array_return(self)
    def test_indirect_basic(self)
    def test_indirect_sheet_reference(self)
```

### Integration Tests
```python
class TestDynamicRangeIntegration(unittest.TestCase):
    def test_nested_functions(self)      # INDEX(OFFSET(...))
    def test_formula_evaluation(self)    # =SUM(OFFSET(...))
    def test_cross_sheet_references(self)
    def test_complex_scenarios(self)
```

### Excel Compatibility Tests
- Compare results with actual Excel calculations
- Test edge cases and boundary conditions
- Validate error messages and types

## Performance Considerations

### Optimization Strategies
1. **Lazy evaluation**: Only resolve references when needed
2. **Caching**: Cache parsed references for repeated use
3. **Bounds checking**: Early validation to avoid expensive operations
4. **Memory efficiency**: Minimize object creation in hot paths

### Benchmarking
- Measure impact on existing function performance
- Profile reference resolution utilities
- Compare with Excel calculation speed
- Memory usage analysis

## Migration and Compatibility

### Breaking Changes
- Remove complex parser preprocessing for `:OFFSET`/`:INDEX`
- Simplify parser by removing lines 126-185 in `parser.py`

### Backward Compatibility
- All existing functions remain unchanged
- Existing formula syntax continues to work
- No impact on current test suite

### Migration Path
1. Implement new functions alongside existing code
2. Update failing `:OFFSET` test to use standard `OFFSET` syntax
3. Remove parser preprocessing after validation
4. Update documentation and examples

## Success Metrics

### Functional Metrics
- All core functions (OFFSET, INDEX, INDIRECT) implemented
- 100% compatibility with Excel behavior
- Comprehensive error handling
- Integration with existing evaluator

### Quality Metrics
- 100% test coverage for new code
- Zero regressions in existing functionality
- Performance impact < 5%
- Code review approval

### Maintenance Metrics
- Clear, documented code
- Modular, extensible design
- Consistent with existing patterns
- Easy to debug and troubleshoot

## Risk Mitigation

### Technical Risks
- **Complex reference resolution**: Mitigate with comprehensive testing
- **Excel compatibility edge cases**: Research Excel documentation thoroughly
- **Performance impact**: Benchmark and optimize critical paths
- **Integration issues**: Incremental implementation with validation

### Process Risks
- **Scope creep**: Focus on core functions first, advanced features later
- **Breaking changes**: Careful migration strategy with backward compatibility
- **Testing complexity**: Automated test suite with Excel validation

This architecture provides a solid foundation for implementing Excel dynamic range functions while maintaining code quality, performance, and maintainability.