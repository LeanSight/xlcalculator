# 🔴 CRITICAL ATDD CONFIGURATION FOR ONA AGENT

## ⚠️ ABSOLUTE RULES - MAXIMUM PRIORITY

### 🚨 RULE #1: FAILING TESTS = INCORRECT IMPLEMENTATION

**WHEN A TEST FAILS:**
```
❌ PROHIBITED: Create new test
❌ PROHIBITED: Modify test without Excel validation
❌ PROHIBITED: Avoid fixing implementation
✅ MANDATORY: Fix implementation to make existing test pass
✅ MANDATORY: Validate test against Excel documentation first
```

### 🚨 RULE #2: MANDATORY PRE-ACTION PROTOCOL

**BEFORE ANY CODE CHANGES:**
1. ✅ **VALIDATE**: Does this test represent legitimate Excel behavior?
2. ✅ **CONFIRM**: Is the current implementation incorrect?
3. ✅ **PLAN**: What minimal change makes this test pass?
4. ✅ **IMPLEMENT**: Only then make the change

### 🚨 RULE #3: NEVER ADVANCE WITHOUT COMPLETE GREEN

```
NEVER MOVE FORWARD TO A NEW ACCEPTANCE TEST WITHOUT ALL TESTS IN GREEN
```

## 🎯 MANDATORY DECISION TREE

```
Test Fails?
├── YES → STOP: Apply ATDD Protocol
│   ├── 1. Validate test against official Excel
│   ├── 2. If test correct (99%) → FIX IMPLEMENTATION
│   ├── 3. If test incorrect (1%) → DOCUMENT + CORRECT TEST
│   └── 4. Make existing test pass
└── NO → Continue normal development
```

## 🔄 STRICT ATDD CYCLE

### 🔴 RED PHASE (Test Failing)
- **ACTION**: Identify what implementation is missing/broken
- **PROHIBITED**: Create alternative test
- **MANDATORY**: Understand why current test fails

### 🟢 GREEN PHASE (Make Test Pass)
- **ACTION**: Minimal implementation to make existing test pass
- **PROHIBITED**: Premature optimization
- **MANDATORY**: Immediate commit when test passes

### 🔵 REFACTOR PHASE (Improve Code)
- **ACTION**: Improve structure without changing behavior
- **PROHIBITED**: Add new functionality
- **MANDATORY**: Keep all tests green

## 🚫 COMMON VIOLATIONS TO AVOID

### ❌ ANTI-PATTERN 1: "New Test to Avoid Problem"
```python
# INCORRECT
def test_new_function_that_works():
    # Create test that passes to avoid fixing failing test
    pass

# CORRECT
def fix_existing_failing_test():
    # Fix implementation to make existing test pass
    pass
```

### ❌ ANTI-PATTERN 2: "Modify Test Without Validation"
```python
# INCORRECT
def test_function():
    # Change expected_value to make test pass
    assert result == "new_invented_value"

# CORRECT
def test_function():
    # Validate against Excel: what should it really return?
    # Fix implementation to return correct value
    assert result == "value_validated_against_excel"
```

### ❌ ANTI-PATTERN 3: "Complex Implementation in First Iteration"
```python
# INCORRECT
def complex_implementation_with_all_features():
    # Implement everything at once
    pass

# CORRECT
def minimal_implementation_to_pass_current_test():
    # Only what's needed to make current test pass
    pass
```

## ✅ CORRECT ATDD PATTERNS

### ✅ PATTERN 1: "Validation Before Action"
```python
# 1. Test fails
def test_indirect_function():
    assert INDIRECT("A1") == expected_value  # FAILS

# 2. Validate against Excel
# What does INDIRECT("A1") do in real Excel?

# 3. Fix implementation
def INDIRECT(reference):
    # Minimal implementation to make test pass
    pass
```

### ✅ PATTERN 2: "Incremental Implementation"
```python
# Iteration 1: Make basic test pass
def FUNCTION(param):
    if param == "test_case_1":
        return expected_result_1
    raise NotImplementedError()

# Iteration 2: Make next test pass
def FUNCTION(param):
    if param == "test_case_1":
        return expected_result_1
    elif param == "test_case_2":
        return expected_result_2
    raise NotImplementedError()

# Iteration 3: Refactor to generalize
def FUNCTION(param):
    # General implementation that handles all cases
    return general_implementation(param)
```

### ✅ PATTERN 3: "Commit on Green"
```bash
# After each passing test
git add .
git commit -m "🟢 Make test_specific_case pass"
git push

# After each refactor
git add .
git commit -m "🔵 Refactor FUNCTION - eliminate duplication"
git push
```

## 🎯 ATDD COMPLIANCE METRICS

### ✅ SUCCESS INDICATORS
- Existing tests pass without test modification
- Minimal effective implementation
- Frequent commits in green state
- Zero new tests during red phase
- Excel validation documented

### ❌ FAILURE INDICATORS
- Create new test to avoid fixing existing one
- Modify test without Excel validation
- Complex implementation in first iteration
- Advance without all tests green
- Assume behavior without verification

## 🔧 VERIFICATION COMMANDS

### Before Any Changes:
```bash
# 1. Check test status
python -m pytest path/to/failing/test.py -v

# 2. Identify specific failing test
python -m pytest path/to/failing/test.py::test_specific_function -v

# 3. Understand why it fails
# Read error message, understand expectation vs reality
```

### After Implementing Fix:
```bash
# 1. Verify specific test passes
python -m pytest path/to/test.py::test_specific_function -v

# 2. Verify no other tests broke
python -m pytest path/to/test.py -v

# 3. Immediate commit if all green
git add . && git commit -m "🟢 Make test_specific_function pass" && git push
```

## 🎓 CORRECT APPLICATION EXAMPLES

### Example 1: INDIRECT Test Failing
```python
# EXISTING FAILING TEST
def test_indirect_basic():
    result = INDIRECT("A1")
    assert result == 10  # FAILS: returns None

# CORRECT ATDD PROCESS:
# 1. Validate: Does INDIRECT("A1") in Excel return A1 value? YES
# 2. Problem: Implementation doesn't evaluate reference
# 3. Minimal fix:
def INDIRECT(reference):
    # Minimal implementation to make test pass
    if reference == "A1":
        return evaluator.get_cell_value("A1")
    raise NotImplementedError()

# 4. Test passes → Commit → Next test
```

### Example 2: OFFSET Test Failing
```python
# EXISTING FAILING TEST
def test_offset_basic():
    result = OFFSET("A1", 1, 1)
    assert result == "B2"  # FAILS: returns error

# CORRECT ATDD PROCESS:
# 1. Validate: Does OFFSET("A1", 1, 1) in Excel return "B2"? YES
# 2. Problem: Implementation doesn't calculate offset correctly
# 3. Minimal fix:
def OFFSET(reference, rows, cols):
    # Minimal implementation to make test pass
    if reference == "A1" and rows == 1 and cols == 1:
        return "B2"
    raise NotImplementedError()

# 4. Test passes → Commit → Next test
```

## 🚨 FINAL REMINDER

**THIS CONFIGURATION IS MANDATORY AND NON-NEGOTIABLE**

Every time you see a failing test, your FIRST and ONLY reaction must be:
1. ✅ Validate test against Excel
2. ✅ Fix implementation
3. ✅ Make existing test pass
4. ✅ Commit on green

**NEVER:**
- ❌ Create new test
- ❌ Modify test without validation
- ❌ Advance without complete green
- ❌ Implement functionality not required by current test

---

**This configuration must be consulted BEFORE any action when tests are failing.**