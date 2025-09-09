# xlcalculator Development Guidelines

**Document Version**: 1.0  
**Created**: 2025-09-09  
**Application**: Universal development guidelines for xlcalculator and similar projects

---

## 🎯 Core Development Philosophy

### Primary Principles
- **ATDD Strict Compliance**: Implementation must follow expected behavior exactly as defined by acceptance tests
- **Excel Fidelity First**: Match Excel exactly, including quirks and edge cases
- **Architecture-First Problem Solving**: Fix architectural foundations that make problems automatic
- **Evidence-Based Development**: Use official documentation to verify legitimate vs problematic patterns
- **Multiplicative Impact**: Prefer solutions that fix multiple issues simultaneously

---

## 🔄 ATDD Methodology Framework

### Double Nested Cycle Approach

#### **🔄 Outer Cycle (ATDD) - Outside-In**
- **Primary Rule**: Implementation must follow expected behavior exactly as defined by acceptance tests
- **Test-First**: Acceptance tests define business behavior, implementation follows
- **No Test Bypassing**: Never implement functionality that circumvents acceptance test expectations

#### **🔄 Inner Cycle (TDD) - Inside-Out**
- **Unit-Level TDD**: For each acceptance test failure, decompose into unit tests following Red-Green-Refactor
- **Granular Development**: Build components incrementally through unit test cycles
- **Integration Focus**: Unit tests support acceptance test fulfillment

### Implementation Phases

#### **🔴 Red Phase (Failing Tests)**
- **Acceptance Level**: Write failing acceptance test based on business requirements
- **Unit Level**: Write failing unit tests for required components
- **No Implementation**: Code only after test fails

#### **🟢 Green Phase (Passing Tests)**
- **Minimal Implementation**: Write simplest code to make current test pass
- **Return Actual Data**: Return real data or proper error responses, no hardcoded fallbacks
- **No Premature Optimization**: Focus on making test pass first
- **📝 Immediate Commit**: Every green test triggers immediate commit and push

#### **🔵 Refactor Phase (Code Improvement)**
- **Maintain Behavior**: Improve code structure without changing test outcomes
- **Eliminate ALL Duplicate Logic**: Remove every instance of code duplication
- **Functional Style Priority**: Prefer functional style over object-oriented when simpler
- **Idiomatic Code**: Use most idiomatic code for the language and version
- **📝 Immediate Commit**: Every refactor completion triggers immediate commit

### NEVER MOVE FORWARD TO A NEW ACCEPTANCE TEST WITHOUT ALL TESTS IN GREEN
---

## 🧩 Function Implementation Methodology

### Outside-In ATDD Implementation Flow

```
Excel Official Documentation
         ↓
Function Design Document
         ↓
JSON Test Cases (67+ cases)
         ↓
Excel File Generation
         ↓
Failing Acceptance Tests (RED phase)
         ↓
Minimal Implementation (GREEN phase)
         ↓
Refactoring (BLUE phase)
```

### Test Structure Requirements

#### **5-Level Test Organization**
1. **NIVEL 1: CASOS ESTRUCTURALES** - Comportamiento Core (10-15 cases)
2. **NIVEL 2: CASOS INTERMEDIOS** - Funciones Individuales (20-30 cases)
3. **NIVEL 3: CASOS AVANZADOS** - Combinaciones (8-12 cases)
4. **NIVEL 4: CASOS DE CONTEXTO** - Uso con Otras Funciones (5-8 cases)
5. **NIVEL 5: CASOS EDGE** - Comportamientos Límite (5-10 cases)

#### **JSON Test Case Structure**
```json
{
  "metadata": {
    "title": "[Function Group Name]",
    "total_cases": 67,
    "source": "[FUNCTION_GROUP]_DESIGN.md"
  },
  "generation_config": {
    "class_name": "[FunctionGroup]ComprehensiveTest",
    "excel_filename": "[function_group].xlsx"
  },
  "levels": [
    {
      "level": "1A",
      "title": "[Function] - Casos Fundamentales",
      "test_cases": [
        {
          "cell": "A1",
          "formula": "=[EXCEL_FORMULA]",
          "expected_value": [expected_result],
          "expected_type": "number|text|boolean|array|ref_error|value_error"
        }
      ]
    }
  ]
}
```

---

## 💻 Code Quality Standards

### Code Style Requirements

#### **Comments Policy**
- **API Comments**: Document the "why," not the "what"
- **Avoid Redundant Comments**: No line or block comments that restate code
- **Clarification Only**: Only add comments to clarify non-obvious logic

#### **Code Structure**
- **Self-Documenting**: Use descriptive variable names that explain intent
- **Functional Style Priority**: Prefer functional style when it results in simpler code
- **Idiomatic Code**: Use most idiomatic code for the specific language and version
- **Zero Duplication**: Eliminate all duplicate logic through extraction
- **Explicit Logic**: Prefer explicit conditionals over clever shortcuts

### File Operations Standards

#### **File Reading Rules**
- **Always read files before editing** to understand structure and conventions
- **For large files**: Read relevant sections rather than entire file
- **Explore codebase**: Start with entry points and configuration files
- **No assumptions**: Check dependencies (package.json/requirements.txt) first

#### **File Editing Process**
1. **Understand Context**: Read file's conventions (style, imports, patterns)
2. **Maintain Consistency**: Match existing code style exactly
3. **Check Dependencies**: Never assume libraries are available
4. **Follow Patterns**: Apply project's established patterns
5. **Intentional Edits**: All file edits are intentional and MUST NOT be reverted

---

## 🎨 Code Tidying Techniques

### The 15 Tidying Techniques

#### **1. Guard Clauses**
Reduce nesting by handling edge cases early:
```python
def process_user(user):
    if not user: return None
    if not user.is_active: return None
    if not user.has_permission: return None
    
    return do_something(user)
```

#### **2. Dead Code**
Remove unused functions, variables, imports, and commented-out code.

#### **3. Normalize Symmetries**
Make similar operations follow the same pattern:
```python
status_map = {"active": True, "inactive": False}
if status not in status_map:
    raise Exception("Invalid status")
return status_map[status]
```

#### **4. New Interface, Old Implementation**
Create the interface you wish you had:
```python
def calculate_user_score(user):
    return legacy_score_calculation(user.data, user.preferences, user.history)
```

#### **5. Reading Order**
Arrange code so it reads naturally from top to bottom.

#### **6. Cohesion Order**
Group related elements together.

#### **7. Move Declaration and Initialization Together**
Keep variable declaration close to usage.

#### **8. Explaining Variables**
Extract complex expressions into well-named variables:
```python
is_adult = user.age >= 18
is_supported_country = user.country in ['US', 'CA', 'UK']
is_verified_user = user.verified

if is_adult and is_supported_country and is_verified_user:
    allow_access()
```

#### **9. Explaining Constants**
Replace magic numbers and strings with named constants.

#### **10. Explicit Parameters**
Make implicit dependencies explicit through parameters.

#### **11. Chunk Statements**
Use blank lines to group related statements.

#### **12. Extract Helper**
Create helper functions to name and isolate logical operations.

#### **13. One Pile**
Sometimes bring scattered code together to see the full picture.

#### **14. Explaining Comments**
Add comments when code logic is complex or non-obvious.

#### **15. Delete Redundant Comments**
Remove comments that don't add value.

### Tidying Application Guidelines

#### **Tidy First When:**
- You need to understand code before changing it
- The change will be easier after structural improvements
- The tidying takes less time than the benefit it provides

#### **Tidy After When:**
- You're going to change the same area again soon
- You want to leave the code better than you found it

#### **Don't Tidy When:**
- You'll never touch the code again
- Time pressure is extreme
- The code works and changes are risky

---

## 🏗️ Architecture Patterns

### Context-Aware Function Pattern

```python
@xl.register()
@xl.validate_args
def EXCEL_FUNCTION(reference: func_xltypes.XlAnything = None, *, _context=None):
    """Excel-compatible function with context injection."""
    
    if reference is None:
        if _context is None:
            raise RuntimeError("Function requires context when called without reference")
        return _context.property
    
    # Handle reference parameter
    if isinstance(reference, str):
        ref = CellReference.parse(reference, _context.current_sheet if _context else None)
        return ref.property
    
    raise ValueExcelError("Invalid reference parameter")

# Registration
@context_aware
def MY_FUNCTION(arg1, *, _context=None):
    # Function implementation
    pass
```

### Reference Object System

```python
@dataclass
class CellReference:
    """Excel-compatible single cell reference."""
    
    sheet: str
    row: int           # 1-based row index
    column: int        # 1-based column index
    absolute_row: bool = False
    absolute_column: bool = False
    
    def offset(self, rows: int, cols: int) -> 'CellReference':
        """Excel-style reference arithmetic."""
        new_row = self.row + rows
        new_col = self.column + cols
        
        if new_row < 1 or new_row > 1048576:
            raise RefExcelError("Row index out of Excel bounds")
        if new_col < 1 or new_col > 16384:
            raise RefExcelError("Column index out of Excel bounds")
            
        return CellReference(self.sheet, new_row, new_col, 
                           self.absolute_row, self.absolute_column)
    
    def resolve(self, evaluator) -> Any:
        """Get actual cell value through evaluator."""
        return evaluator.get_cell_value(self.address)
    
    @classmethod
    def parse(cls, address: str, current_sheet: str = None) -> 'CellReference':
        """Parse Excel address string to CellReference."""
        # Implementation details...

@dataclass
class RangeReference:
    """Excel-compatible range reference."""
    
    start_cell: CellReference
    end_cell: CellReference
    
    def offset(self, rows: int, cols: int) -> 'RangeReference':
        """Offset entire range by specified rows/columns."""
        return RangeReference(
            start_cell=self.start_cell.offset(rows, cols),
            end_cell=self.end_cell.offset(rows, cols)
        )
    
    def resize(self, rows: int, cols: int) -> 'RangeReference':
        """Resize range to specified dimensions."""
        # Implementation details...
```

### Hierarchical Excel Model

```python
@dataclass
class Workbook:
    name: str = ""
    worksheets: Dict[str, 'Worksheet'] = field(default_factory=dict)
    defined_names: Dict[str, Any] = field(default_factory=dict)
    active_sheet: Optional[str] = None
    
    def get_worksheet(self, name: str) -> 'Worksheet'
    def add_worksheet(self, name: str) -> 'Worksheet'
    def get_cell(self, address: str) -> 'Cell'

@dataclass
class Worksheet:
    name: str
    workbook: 'Workbook'
    cells: Dict[str, 'Cell'] = field(default_factory=dict)
    ranges: Dict[str, 'Range'] = field(default_factory=dict)
    
    def get_cell(self, address: str) -> 'Cell'
    def get_range(self, address: str) -> 'Range'

@dataclass
class Cell:
    address: str  # Local address like "A1"
    worksheet: 'Worksheet'
    value: Any = None
    formula: Optional['XLFormula'] = None
    
    @property
    def row(self) -> int
    @property
    def column_index(self) -> int
```

### Excel Error Handling Pattern

```python
from xlcalculator.xlfunctions.xlerrors import (
    ValueExcelError,    # #VALUE! - Invalid argument type or value
    RefExcelError,      # #REF! - Invalid cell reference
    NameExcelError,     # #NAME? - Unrecognized function or name
    NumExcelError,      # #NUM! - Invalid numeric value
    DivExcelError,      # #DIV/0! - Division by zero
    NaExcelError,       # #N/A - Value not available
    NullExcelError      # #NULL! - Null intersection
)

def EXCEL_FUNCTION(param1, param2):
    """Function with proper Excel error handling."""
    
    # Parameter validation
    if param1 is None:
        raise ValueExcelError("Parameter 1 cannot be empty")
    
    # Type validation
    try:
        numeric_param = float(param1)
    except (ValueError, TypeError):
        raise ValueExcelError("Parameter 1 must be numeric")
    
    # Range validation
    if numeric_param < 0:
        raise NumExcelError("Parameter 1 must be non-negative")
    
    # Division by zero check
    if param2 == 0:
        raise DivExcelError("Cannot divide by zero")
    
    # Reference validation
    if isinstance(param2, str) and '!' in param2:
        try:
            CellReference.parse(param2)
        except Exception:
            raise RefExcelError("Invalid cell reference")
    
    return numeric_param / param2
```

---

## 📋 Task Management Framework

### Todo System Usage Rules

**For ALL tasks beyond trivial one-liners, MUST use Todo system:**

1. **New Unrelated Tasks**: Start with `todo_clear` for clean slate
2. **Immediate Analysis**: Create comprehensive todo list using `todo_write`
3. **Todo Item Quality**:
   - Specific and actionable
   - Logically sequenced with dependencies
   - Granular (3-6 items for most tasks, max 10 for complex)
   - Include verification steps
4. **Processing Method**: Use `todo_next` to advance through items
5. **Completion Rule**: Only summarize when ALL items completed

### Task Classification
- **New unrelated task**: `todo_reset` (clean slate)
- **Continuation of current work**: `todo_write` (append)
- **Complex tasks**: Always include cleanup step

### Todo System Examples

```bash
# New Task Pattern
todo_reset ["Read existing code to understand structure", "Check dependencies", "Create functionality following conventions", "Test functionality", "Run existing tests", "Document findings", "Clean up tmp/ directory", "Commit changes"]

# Documentation Pattern  
todo_write ["Document findings in ona-memory/[timestamp]-analysis.md", "Clean up all files in tmp/ directory", "Verify no temporary artifacts remain"]
```

---

## 🔧 Git Operations & Version Control

### ATDD Git Flow Integration

```bash
# After every Green phase
git add . && git commit -m "🟢 Make [test description] pass" && git push

# After every Refactor phase  
git add . && git commit -m "🔵 Refactor [component] - [improvement]" && git push
```

### Commit Process
1. Run `git status` to see all changes
2. Run `git diff` to review modifications
3. Run `git log --oneline -5` to understand message style
4. Only stage files relevant to current task
5. Add co-author: `Co-authored-by: Ona <no-reply@ona.com>`
6. Follow repository's commit message conventions

### Commit Rules
- **Never commit or push** changes unless explicitly asked
- **One-time Permission**: Each commit permission is explicit and one-time only
- **Phase Tracking**: Every commit linked to specific ATDD phase

---

## 🧪 Testing Standards

### Testing Approach
- **ATDD Driven**: Tests define expected behavior, implementation follows
- **Test-First**: Write failing tests before any implementation
- **Complete Coverage**: Both acceptance and unit tests required
- **Continuous Verification**: Run tests frequently during development

### Test Structure Requirements

#### **Integration Test Pattern**
```python
from tests.testing import FunctionalTestCase

class FunctionComprehensiveTest(FunctionalTestCase):
    """Comprehensive integration tests for [function group].
    
    Test cases derived from [DESIGN].md and validated against Excel.
    Total test cases: 67+
    """
    
    filename = "[function_group].xlsx"
    
    def test_1a_fundamentals_a1(self):
        """[Function] básico: =[FORMULA] → [expected]"""
        excel_value = self.evaluator.get_cell_value('Tests!A1')
        calculated_value = self.evaluator.evaluate('Tests!A1')
        self.assertEqual(calculated_value, excel_value)
```

#### **Unit Test Pattern**
```python
def test_context_aware_function():
    # Create test context
    context = create_context(test_cell, test_evaluator)
    
    # Test function with context
    result = my_function(args, _context=context)
    
    assert result == expected
```

### Quality Verification
- **Run Project Tests**: Use project's existing test commands
- **Linting**: Run linting if available
- **Coverage**: Verify test coverage meets standards
- **No Regression**: New features cannot break existing test suite

---

## 🚀 Performance Guidelines

### Optimization Strategies

#### **Context System Performance**
```python
# Fast O(1) lookup vs O(n) signature inspection
_CONTEXT_REQUIRED_FUNCTIONS: Set[str] = set()

def needs_context_by_name(func_name: str) -> bool:
    return func_name in _CONTEXT_REQUIRED_FUNCTIONS

# Context caching
@lru_cache(maxsize=1000)
def _cached_calculation(param1, param2):
    return complex_calculation(param1, param2)
```

#### **Lazy Evaluation**
```python
class LazyReference:
    def __init__(self, reference):
        self.reference = reference
        self._cached_value = None
        self._is_evaluated = False
    
    def resolve(self, evaluator):
        if not self._is_evaluated:
            self._cached_value = self.reference.resolve(evaluator)
            self._is_evaluated = True
        return self._cached_value
```

### Performance Monitoring
```python
import time
import functools

def performance_monitor(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        
        execution_time = end_time - start_time
        if execution_time > 0.1:
            print(f"Slow function: {func.__name__} took {execution_time:.3f}s")
        
        return result
    return wrapper
```

---

## 📁 File Management Guidelines

### Document Storage in ona-memory/

#### **Naming Convention**
- **Format**: `[timestamp]-[descriptive-name].md`
- **Timestamp**: YYYYMMDD-HHMMSS
- **Extension**: Always `.md` for markdown

#### **Content Guidelines**
- **Structured Format**: Use markdown headers
- **ATDD Context**: Include phase information
- **Actionable Insights**: Document recommendations
- **Evidence-Based**: Include data and observations

### Temporary Files in tmp/

#### **Usage Rules**
- **Session Scope**: Files exist only for current session
- **Auto-cleanup**: May be removed between sessions
- **No Permanent Data**: Never store important info only in tmp/

#### **Cleanup Management**
- **Track Creation**: Track all temporary files created
- **Pre-completion Cleanup**: Remove temp files before task completion
- **Preserve Documentation**: Keep analysis docs in ona-memory/
- **Final Step**: Include cleanup in todo list

---

## 🎯 Excel Compliance Standards

### Core Principles
1. **Test-Driven Compliance**: All Excel behavior verified through tests
2. **No Assumptions**: Never assume Excel behavior without testing
3. **Exact Replication**: Match Excel's exact output, including error types
4. **Documentation**: Document Excel quirks discovered

### Implementation Guidelines
- Use ATDD for all Excel function implementations
- Write tests that compare xlcalculator output directly with Excel
- Handle Excel's specific error types properly
- Maintain Excel's precedence rules and calculation order
- Preserve Excel's edge case handling

### Excel Documentation References
- [Microsoft Excel Function Reference](https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188)
- [Excel Error Values](https://support.microsoft.com/en-us/office/excel-error-values-3ecf8b8b-dc34-4a47-8712-c688b8f8a0a3)
- Function-specific Microsoft documentation for each implementation

---

## ✅ Quality Gates

### Each Phase Must Meet These Criteria:

#### **Design Quality (Phases 1-3)**
- ✅ All Excel behaviors documented with official references
- ✅ 60+ test cases covering fundamental to edge scenarios
- ✅ JSON specification complete and validated

#### **Test Infrastructure Quality (Phases 4-5)**
- ✅ Excel file generates expected results
- ✅ All integration tests fail initially (RED phase confirmed)
- ✅ Test infrastructure validated and working

#### **Implementation Quality (Phases 6-7)**
- ✅ All tests pass (100% GREEN phase)
- ✅ Code follows established patterns
- ✅ Performance meets benchmarks

#### **Production Quality (Phases 8-10)**
- ✅ Excel fidelity validated (exact behavior matching)
- ✅ Integration testing complete
- ✅ Documentation complete and accurate
- ✅ Ready for production deployment

---

## 🔍 Success Metrics

### **Technical Metrics**
- **Function Coverage**: Target 100% for implemented function groups
- **Excel Behavior Matching**: Verified against official documentation
- **Performance Overhead**: ≤10% compared to baseline
- **Test Coverage**: 100% integration tests with real Excel files

### **Quality Metrics**
- **ATDD Compliance**: All implementations follow test-driven approach
- **Documentation Coverage**: All Excel behaviors documented with references
- **Error Compatibility**: All error types match Excel exactly
- **Edge Case Handling**: All Excel edge cases properly implemented

---

**Document Status**: Living document - update with new patterns and guidelines as they are established.