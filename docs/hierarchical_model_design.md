# Hierarchical Model Structure Design

## Current State Analysis

### Current Flat Model Structure
- **Model class**: Single flat dictionary for all cells (`cells: dict`)
- **Cell addressing**: Full addresses like "Sheet1!A1" as dictionary keys
- **Sheet handling**: Implicit through address prefixes
- **Memory usage**: All cells loaded into single dictionary
- **Cross-sheet references**: Handled through string parsing

### Issues with Current Approach
1. **Performance**: O(n) lookups for sheet operations
2. **Memory inefficiency**: No lazy loading of sheets
3. **Limited Excel compatibility**: No proper sheet object model
4. **Maintenance complexity**: String parsing for sheet operations
5. **Scalability**: Poor performance with large workbooks

## Proposed Hierarchical Structure

### Excel Object Model Hierarchy
```
Workbook
├── Worksheets (collection)
│   ├── Worksheet
│   │   ├── Cells (collection)
│   │   │   └── Cell
│   │   ├── Ranges (collection)
│   │   │   └── Range
│   │   └── Properties (name, visibility, etc.)
│   └── ...
├── DefinedNames (collection)
└── Properties (active_sheet, etc.)
```

### Implementation Design

#### 1. Workbook Class
```python
@dataclass
class Workbook:
    name: str = ""
    worksheets: Dict[str, 'Worksheet'] = field(default_factory=dict)
    defined_names: Dict[str, Any] = field(default_factory=dict)
    active_sheet: Optional[str] = None
    
    def get_worksheet(self, name: str) -> 'Worksheet'
    def add_worksheet(self, name: str) -> 'Worksheet'
    def remove_worksheet(self, name: str) -> None
    def get_cell(self, address: str) -> 'Cell'  # Cross-sheet access
    def set_cell_value(self, address: str, value: Any) -> None
```

#### 2. Worksheet Class
```python
@dataclass
class Worksheet:
    name: str
    workbook: 'Workbook'
    cells: Dict[str, 'Cell'] = field(default_factory=dict)  # "A1" format
    ranges: Dict[str, 'Range'] = field(default_factory=dict)
    visible: bool = True
    
    def get_cell(self, address: str) -> 'Cell'  # Local address like "A1"
    def set_cell_value(self, address: str, value: Any) -> None
    def get_range(self, address: str) -> 'Range'
    def get_full_address(self, local_address: str) -> str  # "A1" -> "Sheet1!A1"
```

#### 3. Enhanced Cell Class
```python
@dataclass
class Cell:
    address: str  # Local address like "A1"
    worksheet: 'Worksheet'
    value: Any = None
    formula: Optional['XLFormula'] = None
    
    @property
    def full_address(self) -> str  # "Sheet1!A1"
    @property
    def row(self) -> int
    @property
    def column(self) -> str
    @property
    def column_index(self) -> int
```

#### 4. Enhanced Range Class
```python
@dataclass
class Range:
    address: str  # Local range like "A1:B2"
    worksheet: 'Worksheet'
    
    @property
    def full_address(self) -> str  # "Sheet1!A1:B2"
    @property
    def cells(self) -> List[List['Cell']]
    def get_cell(self, row: int, col: int) -> 'Cell'
```

### Migration Strategy

#### Phase 1: Create New Classes
1. Implement Workbook, Worksheet, Cell, Range classes
2. Maintain backward compatibility with existing Model class
3. Add conversion methods between old and new structures

#### Phase 2: Update ModelCompiler
1. Modify ModelCompiler to create hierarchical structure
2. Update Excel file reading to populate worksheets
3. Maintain existing API for backward compatibility

#### Phase 3: Update Evaluator
1. Modify Evaluator to work with hierarchical model
2. Update address resolution to use worksheet lookups
3. Optimize cross-sheet reference handling

#### Phase 4: Update Functions
1. Update reference-aware functions to use new model
2. Enhance context injection for worksheet awareness
3. Optimize range operations

### Benefits of Hierarchical Model

#### Performance Improvements
- **O(1) sheet lookups**: Direct worksheet access
- **Lazy loading**: Load sheets on demand
- **Memory efficiency**: Better memory management per sheet
- **Optimized ranges**: Sheet-local range operations

#### Excel Compatibility
- **Proper object model**: Matches Excel's structure
- **Sheet operations**: Native support for sheet-level operations
- **Cross-workbook support**: Foundation for multiple workbooks
- **Hidden sheet handling**: Proper visibility management

#### Maintainability
- **Clear separation**: Sheet logic separated from global logic
- **Type safety**: Better type hints and validation
- **Extensibility**: Easy to add new Excel features
- **Testing**: Easier unit testing of components

### Backward Compatibility

#### API Preservation
- Keep existing Model.get_cell_value() and set_cell_value() methods
- Maintain flat dictionary access through properties
- Preserve existing ModelCompiler interface

#### Migration Path
- Gradual migration of internal code
- Deprecation warnings for old patterns
- Clear upgrade documentation

### Implementation Priority

1. **Core Classes**: Workbook, Worksheet, Cell (high priority)
2. **ModelCompiler Integration**: Update file reading (high priority)
3. **Evaluator Updates**: Address resolution (medium priority)
4. **Function Updates**: Reference-aware functions (medium priority)
5. **Performance Optimization**: Lazy loading (low priority)
6. **Advanced Features**: Cross-workbook support (future)

This design provides a solid foundation for Excel compatibility while maintaining performance and extensibility.