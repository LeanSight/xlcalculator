# Reference System Design

**Document Version**: 1.0  
**Last Updated**: 2025-09-09  
**Application**: Excel-compatible reference objects for xlcalculator

---

## 🎯 Design Objectives

### Primary Goals
- **Excel Compatibility**: Match Excel's reference behavior exactly
- **Lazy Evaluation**: Preserve reference information through evaluation pipeline
- **Dynamic Operations**: Enable proper reference arithmetic (OFFSET, INDIRECT)
- **Context Awareness**: Provide direct access to cell coordinates and properties

### Success Criteria
- ✅ OFFSET() works with any Excel file without hardcoded mappings
- ✅ ROW() and COLUMN() return actual cell coordinates
- ✅ Reference objects preserve sheet, row, and column information
- ✅ Lazy evaluation maintains Excel's calculation semantics

---

## 🏗️ Reference Object Architecture

### Core Reference Classes

#### **CellReference - Single Cell Reference**
```python
@dataclass
class CellReference:
    """Excel-compatible single cell reference."""
    
    sheet: str
    row: int           # 1-based row index (Excel convention)
    column: int        # 1-based column index (Excel convention)
    absolute_row: bool = False      # $ prefix for row
    absolute_column: bool = False   # $ prefix for column
    
    @property
    def address(self) -> str:
        """Get Excel-style address (e.g., 'Sheet1!$A$1')."""
        col_letter = self._column_to_letter(self.column)
        row_prefix = '$' if self.absolute_row else ''
        col_prefix = '$' if self.absolute_column else ''
        return f"{self.sheet}!{col_prefix}{col_letter}{row_prefix}{self.row}"
    
    @property
    def coordinate(self) -> tuple:
        """Get (row, column) coordinate tuple."""
        return (self.row, self.column)
    
    def offset(self, rows: int, cols: int) -> 'CellReference':
        """Excel-style reference arithmetic."""
        new_row = self.row + rows
        new_col = self.column + cols
        
        # Validate bounds (Excel limits: 1048576 rows, 16384 columns)
        if new_row < 1 or new_row > 1048576:
            raise RefExcelError("Row index out of Excel bounds")
        if new_col < 1 or new_col > 16384:
            raise RefExcelError("Column index out of Excel bounds")
            
        return CellReference(
            sheet=self.sheet,
            row=new_row,
            column=new_col,
            absolute_row=self.absolute_row,
            absolute_column=self.absolute_column
        )
    
    def resolve(self, evaluator) -> Any:
        """Get actual cell value through evaluator."""
        return evaluator.get_cell_value(self.address)
    
    @classmethod
    def parse(cls, address: str, current_sheet: str = None) -> 'CellReference':
        """Parse Excel address string to CellReference."""
        # Handle sheet prefix
        if '!' in address:
            sheet, cell_part = address.split('!', 1)
        else:
            sheet = current_sheet or 'Sheet1'
            cell_part = address
        
        # Parse absolute references
        absolute_col = cell_part.startswith('$')
        if absolute_col:
            cell_part = cell_part[1:]
        
        # Extract column and row parts
        col_match = ''
        row_match = ''
        for i, char in enumerate(cell_part):
            if char.isalpha():
                col_match += char
            elif char == '$':
                absolute_row = True
            elif char.isdigit():
                row_match = cell_part[i:]
                break
        
        absolute_row = '$' in row_match
        row_num = int(row_match.replace('$', ''))
        col_num = cls._letter_to_column(col_match)
        
        return cls(
            sheet=sheet,
            row=row_num,
            column=col_num,
            absolute_row=absolute_row,
            absolute_column=absolute_col
        )
    
    @staticmethod
    def _column_to_letter(col_num: int) -> str:
        """Convert 1-based column number to Excel letter."""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord('A')) + result
            col_num //= 26
        return result
    
    @staticmethod
    def _letter_to_column(letters: str) -> int:
        """Convert Excel column letters to 1-based number."""
        result = 0
        for char in letters:
            result = result * 26 + (ord(char.upper()) - ord('A') + 1)
        return result
```

#### **RangeReference - Multi-Cell Range Reference**
```python
@dataclass
class RangeReference:
    """Excel-compatible range reference."""
    
    start_cell: CellReference
    end_cell: CellReference
    
    @property
    def address(self) -> str:
        """Get Excel-style range address (e.g., 'Sheet1!A1:C3')."""
        if self.start_cell.sheet == self.end_cell.sheet:
            start_addr = self.start_cell.address.split('!')[1]
            end_addr = self.end_cell.address.split('!')[1]
            return f"{self.start_cell.sheet}!{start_addr}:{end_addr}"
        else:
            return f"{self.start_cell.address}:{self.end_cell.address}"
    
    @property
    def dimensions(self) -> tuple:
        """Get (rows, columns) dimensions."""
        rows = self.end_cell.row - self.start_cell.row + 1
        cols = self.end_cell.column - self.start_cell.column + 1
        return (rows, cols)
    
    def offset(self, rows: int, cols: int) -> 'RangeReference':
        """Offset entire range by specified rows/columns."""
        return RangeReference(
            start_cell=self.start_cell.offset(rows, cols),
            end_cell=self.end_cell.offset(rows, cols)
        )
    
    def resize(self, rows: int, cols: int) -> 'RangeReference':
        """Resize range to specified dimensions."""
        new_end_row = self.start_cell.row + rows - 1
        new_end_col = self.start_cell.column + cols - 1
        
        new_end_cell = CellReference(
            sheet=self.start_cell.sheet,
            row=new_end_row,
            column=new_end_col
        )
        
        return RangeReference(
            start_cell=self.start_cell,
            end_cell=new_end_cell
        )
    
    def get_cell(self, row_offset: int, col_offset: int) -> CellReference:
        """Get specific cell within range by offset."""
        target_row = self.start_cell.row + row_offset
        target_col = self.start_cell.column + col_offset
        
        if (target_row > self.end_cell.row or 
            target_col > self.end_cell.column):
            raise RefExcelError("Cell offset outside range bounds")
        
        return CellReference(
            sheet=self.start_cell.sheet,
            row=target_row,
            column=target_col
        )
    
    def resolve(self, evaluator) -> List[List[Any]]:
        """Get 2D array of values from range."""
        return evaluator.get_range_values(self.address)
    
    @classmethod
    def parse(cls, address: str, current_sheet: str = None) -> 'RangeReference':
        """Parse Excel range address to RangeReference."""
        if ':' not in address:
            # Single cell treated as 1x1 range
            cell_ref = CellReference.parse(address, current_sheet)
            return cls(start_cell=cell_ref, end_cell=cell_ref)
        
        start_addr, end_addr = address.split(':', 1)
        start_cell = CellReference.parse(start_addr, current_sheet)
        end_cell = CellReference.parse(end_addr, current_sheet)
        
        return cls(start_cell=start_cell, end_cell=end_cell)
```

#### **NamedReference - Named Range Reference**
```python
@dataclass
class NamedReference:
    """Excel-compatible named range reference."""
    
    name: str
    workbook_scope: bool = True  # True for workbook-level, False for sheet-level
    sheet: str = None           # Sheet name for sheet-level names
    
    def resolve_to_reference(self, evaluator) -> Union[CellReference, RangeReference]:
        """Resolve named reference to actual cell/range reference."""
        definition = evaluator.get_defined_name(self.name, self.sheet)
        if definition is None:
            raise NameExcelError(f"Name '{self.name}' not found")
        
        # Parse the definition to get actual reference
        if ':' in definition:
            return RangeReference.parse(definition)
        else:
            return CellReference.parse(definition)
    
    def resolve(self, evaluator) -> Any:
        """Get actual value(s) from named reference."""
        reference = self.resolve_to_reference(evaluator)
        return reference.resolve(evaluator)
```

---

## 🔄 Context Integration System

### CellContext - Function Execution Context

```python
@dataclass
class CellContext:
    """Execution context for Excel functions."""
    
    cell: CellReference          # Current cell being evaluated
    evaluator: 'Evaluator'       # Evaluator instance
    formula_sheet: str = None    # Sheet containing the formula
    
    @property
    def current_row(self) -> int:
        """Get current cell row (1-based)."""
        return self.cell.row
    
    @property
    def current_column(self) -> int:
        """Get current cell column (1-based)."""
        return self.cell.column
    
    @property
    def current_sheet(self) -> str:
        """Get current sheet name."""
        return self.cell.sheet
    
    def get_cell_value(self, reference: Union[str, CellReference]) -> Any:
        """Get value from any cell reference."""
        if isinstance(reference, str):
            reference = CellReference.parse(reference, self.current_sheet)
        return reference.resolve(self.evaluator)
    
    def get_range_values(self, reference: Union[str, RangeReference]) -> List[List[Any]]:
        """Get values from range reference."""
        if isinstance(reference, str):
            reference = RangeReference.parse(reference, self.current_sheet)
        return reference.resolve(self.evaluator)
```

### Function Registration with Context

```python
def register_context_function(func):
    """Decorator for functions requiring cell context."""
    
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        # Extract context from evaluator if available
        if '_context' not in kwargs and hasattr(args[0], '_current_context'):
            kwargs['_context'] = args[0]._current_context
        return func(*args, **kwargs)
    
    return wrapper

# Usage example
@xl.register()
@xl.validate_args
@register_context_function
def ROW(reference: func_xltypes.XlAnything = None, *, _context: CellContext = None) -> int:
    """Returns the row number of a reference."""
    if reference is None:
        if _context is None:
            raise RuntimeError("ROW() requires context when called without reference")
        return _context.current_row
    
    # Handle reference parameter
    if isinstance(reference, str):
        ref = CellReference.parse(reference, _context.current_sheet if _context else None)
        return ref.row
    elif hasattr(reference, 'row'):
        return reference.row
    else:
        raise ValueExcelError("Invalid reference for ROW function")
```

---

## 🔧 Function Implementation Patterns

### Pattern 1: Context-Aware Functions (ROW, COLUMN)

```python
@xl.register()
@xl.validate_args
@register_context_function
def COLUMN(reference: func_xltypes.XlAnything = None, *, _context: CellContext = None) -> int:
    """Returns the column number of a reference."""
    if reference is None:
        if _context is None:
            raise RuntimeError("COLUMN() requires context when called without reference")
        return _context.current_column
    
    # Parse reference to get column
    if isinstance(reference, str):
        ref = CellReference.parse(reference, _context.current_sheet if _context else None)
        return ref.column
    elif hasattr(reference, 'column'):
        return reference.column
    else:
        raise ValueExcelError("Invalid reference for COLUMN function")

@xl.register()
@xl.validate_args
@register_context_function
def ROW(reference: func_xltypes.XlAnything = None, *, _context: CellContext = None) -> int:
    """Returns the row number of a reference."""
    if reference is None:
        if _context is None:
            raise RuntimeError("ROW() requires context when called without reference")
        return _context.current_row
    
    # Parse reference to get row
    if isinstance(reference, str):
        ref = CellReference.parse(reference, _context.current_sheet if _context else None)
        return ref.row
    elif hasattr(reference, 'row'):
        return reference.row
    else:
        raise ValueExcelError("Invalid reference for ROW function")
```

### Pattern 2: Reference Arithmetic Functions (OFFSET)

```python
@xl.register()
@xl.validate_args
@register_context_function
def OFFSET(reference: func_xltypes.XlAnything, rows: int, cols: int, 
          height: int = None, width: int = None, *, _context: CellContext = None):
    """Returns a reference offset from a starting reference."""
    
    # Parse starting reference
    if isinstance(reference, str):
        start_ref = CellReference.parse(reference, _context.current_sheet if _context else None)
    elif hasattr(reference, 'offset'):
        start_ref = reference
    else:
        # Try to resolve reference through context
        if _context:
            start_ref = CellReference.parse(str(reference), _context.current_sheet)
        else:
            raise ValueExcelError("Invalid reference for OFFSET function")
    
    # Calculate offset reference
    try:
        offset_ref = start_ref.offset(rows, cols)
    except RefExcelError as e:
        raise e
    
    # Handle height/width parameters for range result
    if height is not None or width is not None:
        height = height or 1
        width = width or 1
        
        if height <= 0 or width <= 0:
            raise ValueExcelError("Height and width must be positive")
        
        # Create range reference
        end_ref = offset_ref.offset(height - 1, width - 1)
        range_ref = RangeReference(start_cell=offset_ref, end_cell=end_ref)
        
        # Return range values
        return range_ref.resolve(_context.evaluator)
    else:
        # Return single cell value
        return offset_ref.resolve(_context.evaluator)
```

### Pattern 3: Dynamic Reference Functions (INDIRECT)

```python
@xl.register()
@xl.validate_args
@register_context_function
def INDIRECT(ref_text: str, a1_style: bool = True, *, _context: CellContext = None):
    """Returns the reference specified by a text string."""
    
    if not isinstance(ref_text, str):
        raise ValueExcelError("INDIRECT requires text reference")
    
    if not a1_style:
        # R1C1 style not implemented yet
        raise ValueExcelError("R1C1 reference style not supported")
    
    try:
        # Parse the text as a reference
        if ':' in ref_text:
            # Range reference
            range_ref = RangeReference.parse(ref_text, _context.current_sheet if _context else None)
            return range_ref.resolve(_context.evaluator)
        else:
            # Single cell reference
            cell_ref = CellReference.parse(ref_text, _context.current_sheet if _context else None)
            return cell_ref.resolve(_context.evaluator)
    
    except Exception:
        raise RefExcelError(f"Invalid reference text: {ref_text}")
```

---

## 🔗 Evaluator Integration

### Enhanced Evaluator with Reference Support

```python
class Evaluator:
    """Enhanced evaluator with reference system support."""
    
    def __init__(self, model, namespace=None):
        self.model = model
        self.namespace = namespace if namespace is not None else xl.FUNCTIONS.copy()
        self._current_context = None
    
    def evaluate(self, addr, context=None):
        """Evaluate cell with proper context injection."""
        
        # Create cell context
        cell_ref = CellReference.parse(addr)
        self._current_context = CellContext(
            cell=cell_ref,
            evaluator=self,
            formula_sheet=cell_ref.sheet
        )
        
        try:
            # Set global context for legacy functions
            from .xlfunctions.dynamic_range import _set_evaluator_context
            _set_evaluator_context(self, addr)
            
            # Evaluate with context
            if addr in self.model.cells:
                cell = self.model.cells[addr]
                if cell.formula:
                    # Inject context into evaluation
                    eval_context = self._get_context(addr, cell_ref.sheet)
                    return cell.formula.ast.eval(eval_context)
                else:
                    return cell.value
            else:
                return None
                
        finally:
            # Clean up context
            self._current_context = None
    
    def get_cell_reference(self, addr: str) -> CellReference:
        """Get CellReference object for address."""
        return CellReference.parse(addr)
    
    def get_range_reference(self, addr: str) -> RangeReference:
        """Get RangeReference object for address."""
        return RangeReference.parse(addr)
```

---

## 📊 Performance Considerations

### Lazy Evaluation Strategy

```python
class LazyReference:
    """Lazy evaluation wrapper for references."""
    
    def __init__(self, reference: Union[CellReference, RangeReference]):
        self.reference = reference
        self._cached_value = None
        self._is_evaluated = False
    
    def resolve(self, evaluator) -> Any:
        """Lazy evaluation with caching."""
        if not self._is_evaluated:
            self._cached_value = self.reference.resolve(evaluator)
            self._is_evaluated = True
        return self._cached_value
    
    def invalidate(self):
        """Invalidate cache when dependencies change."""
        self._is_evaluated = False
        self._cached_value = None
```

### Memory Optimization

```python
@dataclass
class CompactCellReference:
    """Memory-optimized cell reference for large workbooks."""
    
    sheet_id: int      # Sheet index instead of name
    row: int          # 1-based row
    column: int       # 1-based column
    flags: int = 0    # Bit flags for absolute references
    
    @property
    def absolute_row(self) -> bool:
        return bool(self.flags & 0x01)
    
    @property
    def absolute_column(self) -> bool:
        return bool(self.flags & 0x02)
```

---

## 🧪 Testing Strategy

### Reference Object Testing

```python
class TestCellReference(unittest.TestCase):
    
    def test_address_parsing(self):
        """Test Excel address parsing."""
        ref = CellReference.parse("Sheet1!$A$1")
        self.assertEqual(ref.sheet, "Sheet1")
        self.assertEqual(ref.row, 1)
        self.assertEqual(ref.column, 1)
        self.assertTrue(ref.absolute_row)
        self.assertTrue(ref.absolute_column)
    
    def test_offset_arithmetic(self):
        """Test reference offset calculations."""
        ref = CellReference("Sheet1", 5, 3)  # C5
        offset_ref = ref.offset(2, -1)       # B7
        self.assertEqual(offset_ref.row, 7)
        self.assertEqual(offset_ref.column, 2)
    
    def test_bounds_validation(self):
        """Test Excel bounds validation."""
        ref = CellReference("Sheet1", 1048576, 16384)  # Max Excel cell
        with self.assertRaises(RefExcelError):
            ref.offset(1, 0)  # Beyond row limit

class TestRangeReference(unittest.TestCase):
    
    def test_range_parsing(self):
        """Test Excel range parsing."""
        range_ref = RangeReference.parse("Sheet1!A1:C3")
        self.assertEqual(range_ref.dimensions, (3, 3))
    
    def test_range_operations(self):
        """Test range offset and resize."""
        range_ref = RangeReference.parse("A1:B2")
        offset_range = range_ref.offset(1, 1)
        self.assertEqual(offset_range.address, "Sheet1!B2:C3")
```

### Integration Testing with Functions

```python
class TestReferenceIntegration(FunctionalTestCase):
    filename = "reference_test.xlsx"
    
    def test_row_column_functions(self):
        """Test ROW/COLUMN with reference system."""
        # Test without parameters (current cell)
        row_result = self.evaluator.evaluate("Sheet1!C5")  # =ROW()
        self.assertEqual(row_result, 5)
        
        col_result = self.evaluator.evaluate("Sheet1!D6")  # =COLUMN()
        self.assertEqual(col_result, 4)
    
    def test_offset_function(self):
        """Test OFFSET with reference system."""
        # Test dynamic offset calculation
        result = self.evaluator.evaluate("Sheet1!E1")  # =OFFSET(A1,2,1)
        expected = self.evaluator.get_cell_value("Sheet1!B3")
        self.assertEqual(result, expected)
```

---

## 🚀 Migration Plan

### Phase 1: Core Reference Classes (2-3 days)
- Implement CellReference and RangeReference classes
- Add parsing and address generation methods
- Create comprehensive unit tests
- Validate Excel compatibility

### Phase 2: Context Integration (2-3 days)
- Implement CellContext class
- Add context injection to evaluator
- Update function registration system
- Test context propagation

### Phase 3: Function Updates (3-4 days)
- Update ROW, COLUMN, OFFSET, INDIRECT functions
- Remove hardcoded mappings
- Implement reference arithmetic
- Comprehensive integration testing

### Phase 4: Performance Optimization (1-2 days)
- Add lazy evaluation support
- Implement reference caching
- Memory optimization for large workbooks
- Performance benchmarking

---

## 📈 Success Validation

### Functional Validation
- ✅ ROW() returns actual cell row numbers
- ✅ COLUMN() returns actual cell column numbers  
- ✅ OFFSET() works with any Excel file
- ✅ INDIRECT() handles dynamic references correctly

### Performance Validation
- ✅ Reference operations complete within 10ms
- ✅ Memory usage increase ≤20% for reference objects
- ✅ Evaluation performance maintains current benchmarks
- ✅ Large workbook handling remains efficient

### Compatibility Validation
- ✅ All existing tests continue passing
- ✅ Excel behavior matching for edge cases
- ✅ Error handling matches Excel exactly
- ✅ Reference parsing handles all Excel formats

---

**Related Documents**:
- [Architecture Analysis](ARCHITECTURE_ANALYSIS.md) - Current architecture gaps and recommendations
- [Excel Compliance Project Goals](PROJECT_GOALS_EXCEL_COMPLIANCE.md) - Project objectives and success criteria
- [Function Implementation Guide](FUNCTION_IMPLEMENTATION_GUIDE.md) *(Coming Soon)* - Guidelines for implementing Excel functions