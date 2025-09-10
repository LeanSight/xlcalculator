# 🎯 ATDD DECISION PROTOCOL FOR ONA

## 🚨 MANDATORY PROTOCOL: FAILING TESTS

### STEP 1: FAILING TEST DETECTION
```
Is any test failing?
├── YES → ACTIVATE ATDD PROTOCOL (go to STEP 2)
└── NO → Continue normal development
```

### STEP 2: TEST CASE VALIDATION
```
Does this test represent legitimate Excel behavior?

VALIDATE AGAINST:
✅ Official Excel documentation
✅ Behavior in real Excel (if possible)
✅ Project design specifications
✅ Function standards in question

RESULT:
├── Test is CORRECT (99% cases) → go to STEP 3A
└── Test is INCORRECT (1% cases) → go to STEP 3B
```

### STEP 3A: CORRECT TEST - FIX IMPLEMENTATION
```
MANDATORY ACTION: Fix implementation

1. ✅ Identify why current implementation fails
2. ✅ Design minimal change to make test pass
3. ✅ Implement ONLY what's needed for current test
4. ✅ Verify test passes
5. ✅ Verify other tests didn't break
6. ✅ Immediate commit: "🟢 Make [test_name] pass"
7. ✅ Immediate push

PROHIBITED:
❌ Create new test
❌ Modify existing test
❌ Implement extra functionality
❌ Premature optimization
```

### STEP 3B: INCORRECT TEST - CORRECT TEST
```
EXCEPTIONAL ACTION: Correct test (only if documentedly incorrect)

1. ✅ Document why test is incorrect
2. ✅ Reference official Excel documentation
3. ✅ Update source design document
4. ✅ Correct test to reflect real Excel behavior
5. ✅ Verify consistency with related tests
6. ✅ Commit: "🔧 Fix test [test_name] - align with Excel behavior"
7. ✅ Document change in ona-memory/

REQUIRES:
📝 Clear evidence of discrepancy with Excel
📝 References to official documentation
📝 Detailed justification for change
```

## 🔄 STRICT ATDD WORKFLOW

### STATE: TEST FAILING (🔴 RED)
```
OBJECTIVE: Understand what implementation is needed
ACTION: Analyze failing test
PROHIBITED: Write implementation code
RESULT: Clear plan to make test pass
```

### STATE: IMPLEMENTING FIX (🟡 YELLOW)
```
OBJECTIVE: Make test pass with minimal code
ACTION: Implement ONLY what's necessary
PROHIBITED: Add extra functionality
RESULT: Test passes
```

### STATE: TEST PASSING (🟢 GREEN)
```
OBJECTIVE: Confirm everything works
ACTION: Verify all tests
MANDATORY: Immediate commit
RESULT: Code in repository
```

### STATE: REFACTORING (🔵 BLUE)
```
OBJECTIVE: Improve structure without changing behavior
ACTION: Eliminate duplication, improve readability
PROHIBITED: Change test behavior
RESULT: Cleaner code, tests still passing
```

## 🚫 ANTI-PATTERNS TO ABSOLUTELY AVOID

### ❌ ANTI-PATTERN: "New Test to Avoid Problem"
```python
# SITUATION: test_function_basic() fails
# INCORRECT:
def test_function_alternative():  # ❌ CREATE NEW TEST
    # Test that passes to avoid fixing the failing one
    pass

# CORRECT:
# Fix implementation to make test_function_basic() pass
```

### ❌ ANTI-PATTERN: "Modify Test Without Validation"
```python
# SITUATION: test expects result X, implementation returns Y
# INCORRECT:
def test_function():
    result = function()
    assert result == Y  # ❌ CHANGE EXPECTATION WITHOUT VALIDATION

# CORRECT:
# 1. Validate against Excel: should it be X or Y?
# 2. If Excel says X: fix implementation
# 3. If Excel says Y: document and correct test
```

### ❌ ANTI-PATTERN: "Premature Complete Implementation"
```python
# SITUATION: basic test fails
# INCORRECT:
def function(param1, param2, param3):  # ❌ IMPLEMENT EVERYTHING
    # Complete implementation with all cases
    pass

# CORRECT:
def function(param1):  # ✅ MINIMAL FOR CURRENT TEST
    # Only what's needed to make current test pass
    if param1 == "test_case_value":
        return "expected_result"
    raise NotImplementedError()
```

## ✅ MANDATORY CORRECT PATTERNS

### ✅ PATTERN: "Validation First"
```python
# 1. Test fails
def test_excel_function():
    assert EXCEL_FUNCTION("input") == "expected"  # FAILS

# 2. MANDATORY: Validate against Excel
# What does EXCEL_FUNCTION("input") return in real Excel?
# Consult official documentation

# 3. Fix implementation based on validation
def EXCEL_FUNCTION(input):
    # Implementation based on validated Excel behavior
    pass
```

### ✅ PATTERN: "Incremental Implementation"
```python
# Iteration 1: Make first test pass
def FUNCTION(param):
    if param == "case1":
        return "result1"
    raise NotImplementedError("Not implemented for: " + str(param))

# Iteration 2: Make second test pass
def FUNCTION(param):
    if param == "case1":
        return "result1"
    elif param == "case2":
        return "result2"
    raise NotImplementedError("Not implemented for: " + str(param))

# Iteration 3: Refactor when pattern is clear
def FUNCTION(param):
    # General implementation that handles all tested cases
    return general_logic(param)
```

### ✅ PATTERN: "Immediate Green Commit"
```bash
# After each passing test
git add .
git commit -m "🟢 Make test_specific_function pass

- Implement minimal logic for test case
- Validates against Excel behavior
- All tests passing

Co-authored-by: Ona <no-reply@ona.com>"
git push
```

## 🎯 PRE-ACTION CHECKLIST

### Before Touching Any Code:
- [ ] Are there failing tests?
- [ ] Have I validated the test against Excel documentation?
- [ ] Do I understand exactly why the test fails?
- [ ] Do I have a minimal plan to make the test pass?
- [ ] Am I sure I will NOT create a new test?

### Before Implementing:
- [ ] Is my implementation minimal for current test?
- [ ] Am I not adding extra functionality?
- [ ] Am I not optimizing prematurely?
- [ ] Will my code make ONLY the failing test pass?

### Before Commit:
- [ ] Does the specific test now pass?
- [ ] Did I not break other tests?
- [ ] Are all tests green?
- [ ] Does my commit message reflect which test was fixed?

## 🚨 ATDD VIOLATION ALERTS

### 🚨 RED ALERT: Creating New Test During Red Phase
```
DETECTED: Attempt to create new test while tests are failing
ACTION: STOP immediately
CORRECTION: Fix existing test first
```

### 🚨 RED ALERT: Modifying Test Without Validation
```
DETECTED: Change in test expectation without Excel validation
ACTION: STOP immediately  
CORRECTION: Validate against Excel, then decide
```

### 🚨 RED ALERT: Advancing Without Complete Green
```
DETECTED: Attempt to work on new functionality with red tests
ACTION: STOP immediately
CORRECTION: Make all tests pass first
```

## 📋 MANDATORY VERIFICATION COMMANDS

### Test Status:
```bash
# Check which tests fail
python -m pytest -x --tb=short

# Run specific failing test
python -m pytest path/to/test.py::test_function_name -v

# Verify fix works
python -m pytest path/to/test.py::test_function_name -v
```

### State Validation:
```bash
# All tests must pass before continuing
python -m pytest

# Repository state must be clean after commit
git status
```

---

## 🎯 EXECUTIVE SUMMARY

**GOLDEN RULE**: Test fails = Incorrect implementation (99% cases)

**MANDATORY PROTOCOL**:
1. Test fails → Validate test → Fix implementation → Test passes → Commit

**ABSOLUTE PROHIBITIONS**:
- ❌ Create new test to avoid fixing existing one
- ❌ Modify test without Excel validation
- ❌ Advance without all tests green

**THIS PROTOCOL IS MANDATORY AND NON-NEGOTIABLE**