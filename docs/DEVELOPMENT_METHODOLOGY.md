# Development Methodology & Problem Resolution Framework

This document captures the development methodology, problem resolution approach, and rules extracted from the Excel compliance project conversation.

## 🎯 Core Development Principles

### ATDD (Acceptance Test Driven Development) Compliance
- **Primary Rule**: Implementation must follow Excel behavior exactly as defined by tests
- **No Fallbacks**: Avoid fallbacks that violate Excel behavior
- **Return Actual Data**: Return actual Excel data or proper Excel errors
- **No Hardcoded Data**: Performance optimization without compromising compatibility
- **Test-First**: Tests define the expected behavior, implementation follows

### Tone and Communication Standards
- **Be Concise**: Direct and technical communication, avoid conversational pleasantries
- **No Pleasantries**: Never start responses with "Great", "Certainly", "Okay", "Sure"
- **No False Assertions**: Never assert user is "absolutely right" unless certain
- **Output Necessity**: Output only what's necessary to accomplish the task
- **Command Explanation**: Briefly state what commands do and why you're running them
- **Emoji Policy**: No emojis unless explicitly requested (✅, ❌, ⚠️ allowed without permission)

## 📋 Task Management Framework

### Todo System Usage Rules
**For ALL tasks beyond trivial one-liners, MUST use Todo system:**

1. **New Unrelated Tasks**: Start with `todo_clear` for clean slate
2. **Immediate Analysis**: Create comprehensive todo list using `todo_write`
3. **Todo Item Quality**:
   - Specific and actionable (e.g., "Read package.json to check dependencies" not "Understand project")
   - Logically sequenced with dependencies considered
   - Granular (3-6 items for most tasks, no more than 10 for complex ones)
   - Include verification steps (e.g., "Run tests to verify changes")
4. **Processing Method**:
   - Use `todo_next` to efficiently advance through items (preferred)
   - Alternative: Mark as in_progress when starting, mark as done when completed
   - Update list if plan needs adjustment
   - Produce 1-2 lines explaining what you're doing when beginning work
   - Use `todo_write` to append new items if needed
5. **Completion Rule**: Only summarize work when ALL items are completed

**Important**: When all todos completed and starting something else (even trivial), always reset todo list.

### Todo System Examples

**New Task Pattern**:
```
todo_reset ["Read server.js to understand current structure", "Check package.json for installed dependencies", "Identify existing API endpoint patterns", "Create new endpoint following conventions", "Test endpoint with curl command", "Run any existing tests", "Clean up any temporary files or scripts created", "Commit changes with appropriate message"]
```

**Continuation vs New Task**:
- **New unrelated task**: `todo_reset` (clean slate)
- **Continuation of current work**: `todo_write` (append to existing)

## 🔧 Tool Usage Guidelines

### General Tool Principles
- Execute tools without verbose explanations
- Batch related operations when possible
- Use `todo_clear` only for new unrelated tasks
- Always read files before editing to understand context
- Prefer complex one-liners over creating scripts
- Chain commands with `&&` to minimize interruptions
- Use appropriate flags (`-y`, `-f`) to avoid interactive prompts
- **Use relative paths** when executing commands

### File Operations Rules

#### Reading Files
- Always read files before editing to understand structure and conventions
- For large files, read relevant sections rather than entire file
- When exploring codebase, start with entry points and configuration files
- Do not assume programming language, understand languages used first

#### Editing Files
**Process**:
1. First understand file's conventions (style, imports, patterns)
2. Maintain consistency with existing code
3. Never assume libraries are available - check package.json/requirements.txt first
4. Follow project's established patterns for similar components
5. Apply edits using appropriate tool for scope of changes

**Important**: Always read file first to understand context before making edits. All file edits are intentional and MUST NOT be reverted unless explicitly requested by user.

#### Comments Policy
- API comments must document the "why," not the "what"
- Avoid redundant line or block comments that restate code
- Only add line or block comments to clarify non-obvious logic or performance trade-offs

#### Code Style
- Match existing code style exactly
- Use project's preferred formatting and naming conventions

### Preview and Web Server Rules
- Only ever use `exec_preview` to run development servers or preview applications
- NEVER tell user about localhost URLs or ports (they cannot access them)
- When modifying .tsx, .jsx, .ts, or .js files, seek to run development server using `exec_preview`
- Always tell user the preview URL where they can see changes
- Use proper markdown link syntax: `[actual_url](actual_url)` - NEVER use bold formatting

## 🔄 Git Operations Protocol

### Commit Process
1. Run `git status` to see all changes
2. Run `git diff` to review modifications
3. Run `git log --oneline -5` to understand commit message style
4. Only stage files relevant to current task
5. Do not commit files modified before task began unless directly related
6. Add co-author: `Co-authored-by: Ona <no-reply@ona.com>`
7. Follow repository's commit message conventions

### Commit Rules
- **Never commit or push** changes unless explicitly asked
- If user asked to commit once, that does NOT give permission to do it again
- Each commit permission is explicit and one-time only

## 🚨 Error Handling Framework

### Error Resolution Process
1. Read error messages carefully
2. Check for common issues (missing dependencies, syntax errors, configuration)
3. Verify changes against project conventions
4. If stuck after 3 attempts, explain issue clearly and ask for guidance

### Web Reading Error Policy
- If you get 4xx or 5xx HTTP errors three times in a row, stop trying
- Consider it a failure - do not keep retrying same or different URLs
- This prevents infinite loops

## ✅ Best Practices Framework

### 1. Completeness Standard
**Ensure all code is immediately runnable**:
- Include all necessary imports
- Add required dependencies to package files
- Provide complete, not partial, implementations

### 2. Testing Approach
**When applicable, suggest running tests to verify changes**:
- Use project's existing test commands
- Run linting if available
- Verify changes work as expected

### 3. Security Standards
- Never expose or log secrets, API keys, or sensitive data

### 4. Documentation Policy
- Only document when explicitly asked or for genuinely complex logic

### 5. Cleanup Management
**Always clean up temporary artifacts**:
- Track all temporary files, scripts, and artifacts created during task execution
- Before task completion, remove all temporary files not part of deliverable
- Examples: migration scripts, backup files, temporary configs, test artifacts
- Use final cleanup step in todo list for complex tasks
- Only leave files explicitly requested or part of final solution

## 🏗️ Architecture and Design Principles

### ATDD Methodology Rules
- **Test-Driven**: Implementation follows test expectations exactly
- **Excel Behavior**: Match Excel behavior as defined by official documentation
- **No Hardcoded Values**: Dynamic context-based calculation
- **Proper Error Handling**: Meaningful errors when context unavailable
- **No Fallbacks**: Avoid fallbacks that violate Excel behavior
- **Return Actual Data**: Return actual Excel data or proper Excel errors
- **Performance Without Compromise**: Optimization without compromising compatibility

### ATDD Strict Compliance
**❌ NOT Permitted**:
- Hardcoding specific values for test cases
- Changing tests to match incorrect implementation
- Modifying data to make formulas work
- Fallbacks that mask Excel errors

**✅ Permitted**:
- Correcting implementation if there's a real bug
- Implementing missing Excel functionality
- Documenting limitations for incorrect formula design
- Using Excel's pre-calculated values for compatibility

### Problem Resolution Approach
1. **Root Cause Analysis**: Identify fundamental architectural gaps, not symptoms
2. **Architecture-First**: Fix architectural foundations that make individual problems automatic
3. **Evidence-Based**: Use official documentation to verify legitimate vs problematic patterns
4. **Multiplicative Impact**: Prefer solutions that fix multiple issues simultaneously
5. **Gap Analysis**: Distinguish between function bugs and evaluator architecture limitations
6. **Context Isolation**: Identify when functions lack necessary calling context

### Solution Selection Criteria
**Primary Criteria**:
1. **Cleanliness**: Minimal, focused changes that address core issue
2. **Self-Documentation**: Code that clearly expresses intent without extensive comments
3. **Low Risk**: Changes that minimize chance of introducing new bugs
4. **Immediate Impact**: Solutions that directly fix identified problems
5. **Maintainability**: Code that is easy to understand and modify in future

### Design Pattern Priorities
1. **Context-Aware Function Execution**: Functions receive proper context, not global state
2. **Reference Object Preservation**: Maintain Excel's lazy evaluation semantics
3. **Hierarchical Model Structure**: Mirror Excel's actual object model
4. **Coordinate-First API Design**: Work with coordinate objects, not strings
5. **Error Propagation Consistency**: Maintain error types through evaluation chain
6. **Parameter Evaluation Pipeline**: Proper handling of nested function calls

## 📊 Analysis and Documentation Standards

### Code Analysis Framework
1. **Search for Patterns**: Identify all instances of problematic patterns
2. **Categorize by Legitimacy**: Distinguish between Excel-compliant and ATDD violations
3. **Evidence-Based Verification**: Use official documentation to confirm legitimacy
4. **Document with Context**: Provide clear explanations and recommendations
5. **Prioritize by Impact**: Focus on architectural changes over individual fixes
6. **Gap Analysis**: Distinguish between function implementation and evaluator architecture issues
7. **Impact Assessment**: Evaluate working vs limited functionality

### Root Cause Analysis Framework
1. **Information Flow Analysis**: Track how data flows through evaluation chain
2. **Context Loss Identification**: Find where necessary context information is lost
3. **Error Propagation Tracking**: Verify errors maintain type through evaluation
4. **Parameter Evaluation Pipeline**: Analyze how nested functions are processed
5. **Architecture vs Implementation**: Distinguish architectural gaps from function bugs

### Documentation Requirements
- **Evidence-Based**: Include official documentation references
- **Actionable**: Provide concrete implementation steps
- **Prioritized**: Clear priority levels and dependencies
- **Measurable**: Define success criteria and metrics
- **Timeline**: Realistic estimates with resource allocation
- **Gap Classification**: Clearly identify type of issue (architecture vs implementation)
- **Impact Scope**: Define what functionality is affected

## 🎯 Strategic Planning Framework

### Phase-Based Implementation
1. **Architecture Foundation**: Build proper foundations first
2. **Function Implementation**: Individual fixes become automatic
3. **Testing & Validation**: Comprehensive verification
4. **Optimization & Enhancement**: Performance and additional features

### Implementation Strategy Types
1. **Hybrid Targeted Fixes**: Combine targeted fixes without major architectural changes
2. **Architecture-First Approach**: Implement foundations that make function fixes automatic
3. **Incremental Migration**: Gradual transition with backward compatibility
4. **Collaborative Integration**: Work with upstream dependencies (e.g., openpyxl)

### Success Metrics Definition
- **Immediate Benefits**: What improves right after each phase
- **Short-term Benefits**: What improves in following phases
- **Long-term Benefits**: Strategic advantages and scalability
- **Measurable Criteria**: Specific, testable success conditions
- **Functional Success**: Core functionality working correctly
- **Code Quality Success**: Maintainable, self-documenting changes
- **Compatibility Success**: Excel behavior matching

### Risk Management Approach
- **Technical Risks**: Backward compatibility, performance, integration
- **Implementation Risks**: Scope creep, timeline, resource allocation
- **Mitigation Strategies**: Specific actions to address each risk category
- **Rollback Plans**: Phase-by-phase and emergency rollback procedures
- **Risk Classification**: Low/Medium/High risk areas with specific mitigations

## 🧪 Integration Testing Framework

### Integration Test Architecture
**Purpose**: Validate xlcalculator functions against actual Excel behavior by comparing results from real Excel files.

**Test Structure Pattern**:
```python
from .. import testing

class FunctionNameTest(testing.FunctionalTestCase):
    filename = "FUNCTION_NAME.xlsx"
    
    def test_evaluation_cellref(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual(excel_value, value)
```

### Excel File Requirements
1. **Formula Cells**: Contain Excel formulas to be tested
2. **Data Cells**: Provide input data for formulas
3. **Result Storage**: Excel calculates and stores expected results
4. **Multiple Scenarios**: Cover edge cases, data types, error conditions

### Test Design Templates

#### Template 1: Simple Function Test
- Basic functionality validation
- Edge cases and error conditions
- Single scenario per test method

#### Template 2: Multi-Scenario Test
- Matrix testing across multiple inputs
- Comprehensive scenario coverage
- Loop-based validation

#### Template 3: Data Type Validation
- Numeric inputs testing
- Text inputs testing
- Error inputs testing
- Type conversion validation

### Excel File Design Patterns

#### Pattern 1: Basic Function Testing
```
A1: =FUNCTION(parameter1, parameter2)
A2: =FUNCTION(edge_case_param)
A3: =FUNCTION(error_case_param)
B1: Input data 1
B2: Input data 2
B3: Input data 3
```

#### Pattern 2: Comprehensive Matrix Testing
```
    A        B        C        D        E
1   Input1   Input2   Input3   Input4   Input5
2   =FUNC(A1) =FUNC(B1) =FUNC(C1) =FUNC(D1) =FUNC(E1)
3   =FUNC(A1,B1) =FUNC(B1,C1) =FUNC(C1,D1) =FUNC(D1,E1) =FUNC(E1,A1)
```

#### Pattern 3: Error Condition Testing
```
A1: =FUNCTION(valid_input)     → Expected result
A2: =FUNCTION(#DIV/0!)         → Error handling
A3: =FUNCTION("")              → Empty string handling
A4: =FUNCTION(text_input)      → Type conversion
A5: =FUNCTION(large_number)    → Boundary testing
```

### Priority Classification
- **Priority 1**: Critical missing functions (immediate)
- **Priority 2**: Information & text functions (high)
- **Priority 3**: Advanced functions (medium)
- **Priority 4**: Specialized functions (low)

### Coverage Targets
- **Phase 1**: 70% integration test coverage
- **Phase 2**: 85% integration test coverage
- **Phase 3**: 95% integration test coverage
- **Phase 4**: 100% integration test coverage

## 🔧 Implementation Patterns

### Red-Green-Refactor Cycle
**Standard TDD approach for all implementations**:
1. **Red**: Write failing test that defines expected behavior
2. **Green**: Implement minimal code to make test pass
3. **Refactor**: Improve code quality while maintaining test passage

### Function Implementation Patterns

#### Pattern 1: Simple Function Implementation
```python
@xl.register()
@xl.validate_args
def FUNCTION_NAME(param1: func_xltypes.XlType, param2: func_xltypes.XlType = None) -> func_xltypes.XlType:
    """Function description with Excel documentation link."""
    # Validation logic
    # Core implementation
    # Return proper Excel type
```

#### Pattern 2: Context-Aware Function Implementation
```python
@xl.register()
@xl.validate_args
def CONTEXT_FUNCTION(reference: func_xltypes.XlAnything = None, *, _context: CellContext = None) -> func_xltypes.XlType:
    """Context-dependent function (ROW, COLUMN, etc.)."""
    if reference is None:
        return _context.cell.property  # Use context for current cell
    # Handle reference parameter
```

#### Pattern 3: Error Handling Implementation
```python
def FUNCTION_WITH_ERRORS(param):
    """Function with proper Excel error handling."""
    try:
        # Validation
        if invalid_condition:
            raise xlerrors.ValueExcelError("Specific error message")
        # Implementation
        return result
    except Exception as e:
        # Convert to appropriate Excel error
        return self._convert_to_excel_error(e)
```

### Evaluator Integration Patterns

#### Pattern 1: Parameter Evaluation with Fallback
```python
def _eval_with_fallback(self, pitem, context):
    """Evaluate parameter with fallback to stored cell values."""
    result = pitem.eval(context)
    
    # If evaluation returns BLANK, try fallback strategies
    if isinstance(result, Blank) and hasattr(pitem, 'tvalue'):
        # Strategy 1: Try to get stored cell value
        cell_addr = pitem.tvalue
        if hasattr(context, 'evaluator') and cell_addr in context.evaluator.model.cells:
            cell = context.evaluator.model.cells[cell_addr]
            if cell.value and str(cell.value) != 'BLANK':
                return func_xltypes.ExcelType.cast_from_native(cell.value)
    
    return result
```

#### Pattern 2: Context Propagation
```python
def _get_context(self, ref, formula_sheet=None):
    """Create context with proper sheet information."""
    return EvaluatorContext(self, ref, formula_sheet)

def __init__(self, evaluator, ref, formula_sheet=None):
    """Initialize context with formula sheet context."""
    super().__init__(evaluator.namespace, ref, formula_sheet)
    self.evaluator = evaluator
```

#### Pattern 3: Error Propagation Enhancement
```python
def evaluate(self, addr, context=None):
    """Enhanced evaluation with proper error handling."""
    try:
        value = cell.formula.ast.eval(context)
        
        # Enhanced error handling
        if isinstance(value, xlerrors.ExcelError):
            # Preserve error types instead of converting to BLANK
            return value
        elif value is None:
            return func_xltypes.BLANK
        
        return value
    except Exception as e:
        # Improved exception handling
        if isinstance(e, xlerrors.ExcelError):
            return e
        # Convert other exceptions to appropriate Excel errors
        return self._convert_exception_to_excel_error(e)
```

### Code Quality Patterns

#### Self-Documenting Code
- Use descriptive variable names that explain intent
- Prefer explicit conditionals over clever shortcuts
- Structure code to read like the problem domain
- Minimize comments by making code self-explanatory

#### Minimal Change Principle
- Make smallest possible change to fix issue
- Prefer single-word changes when possible (e.g., `is not None`)
- Avoid refactoring unrelated code in same change
- Focus changes on core issue being addressed

#### Backward Compatibility
- Maintain existing API signatures
- Add new parameters as optional with defaults
- Use feature flags for behavior changes if needed
- Provide migration path for breaking changes

## 🔍 Gitpod Knowledge Rules

### Documentation Priority
- When asked about Gitpod features, configuration, or usage, ALWAYS use `gitpod_docs` tool FIRST
- Embedded documentation is authoritative source for current Gitpod functionality
- Only rely on general knowledge if documentation doesn't contain relevant information

### Gitpod CLI Commands
**Environments**: Use `gitpod environment` commands for lifecycle management
**Automations**: Use `gitpod automations` for workflow management

## 📝 Example Workflows

### New Task Workflow
```
User: "Add a new API endpoint to the Express app"
Response: 
todo_reset ["Read server.js to understand current structure", "Check package.json for installed dependencies", "Identify existing API endpoint patterns", "Create new endpoint following conventions", "Test endpoint with curl command", "Run any existing tests", "Clean up any temporary files or scripts created", "Commit changes with appropriate message"]
[Execute each todo item methodically using todo_next]
```

### Continuation Workflow
```
User: "Also add error handling to that endpoint"
Response:
todo_write ["Add error handling to endpoint", "Update tests for error cases"]
[Continue with existing todo list]
```

### New Unrelated Task Workflow
```
User: "Now update the documentation"
Response:
todo_reset ["Read existing docs", "Add API endpoint documentation", "Test documentation links"]
[NEW UNRELATED TASK = RESET LIST]
```

## 🎯 Goal and Focus

**Primary Goal**: Accomplish user's task efficiently and correctly, not engage in conversation. Focus on action and results.

**Core Philosophy**: 
- Architecture-first problem solving
- Evidence-based decision making
- ATDD compliance over workarounds
- Multiplicative impact over individual fixes
- Proper foundations enable automatic solutions

---

**Document Version**: 1.0  
**Source**: Excel Compliance Project Conversation  
**Last Updated**: 2025-01-09  
**Application**: All xlcalculator development work