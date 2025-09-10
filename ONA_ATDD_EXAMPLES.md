# 🎯 ATDD EXAMPLES FOR ONA AGENT

## 📋 CASE STUDIES: CORRECT vs INCORRECT BEHAVIOR

### 🔴 CASE 1: INDIRECT TEST FAILING

#### ❌ INCORRECT BEHAVIOR (ATDD VIOLATION)
```python
# SITUATION: test_indirect_basic() fails
def test_indirect_basic():
    result = INDIRECT("A1")
    assert result == 10  # FAILS: returns None

# INCORRECT RESPONSE:
def test_indirect_alternative():  # ❌ CREATE NEW TEST
    result = INDIRECT("A1")
    assert result is None  # ❌ CHANGE EXPECTATION
    
# OR WORSE:
def test_indirect_basic():
    result = INDIRECT("A1") 
    assert result is None  # ❌ MODIFY TEST WITHOUT VALIDATION
```

#### ✅ CORRECT BEHAVIOR (STRICT ATDD)
```python
# SITUATION: test_indirect_basic() fails
def test_indirect_basic():
    result = INDIRECT("A1")
    assert result == 10  # FAILS: returns None

# CORRECT ATDD PROCESS:
# 1. VALIDATE: Does INDIRECT("A1") in Excel return A1 value? YES
# 2. PROBLEM: Implementation doesn't evaluate reference correctly
# 3. MINIMAL FIX:
def INDIRECT(reference):
    if reference == "A1":
        return evaluator.get_cell_value("A1")  # Returns 10
    raise NotImplementedError()

# 4. TEST PASSES → COMMIT: "🟢 Make test_indirect_basic pass"
# 5. NEXT TEST
```

### 🔴 CASE 2: OFFSET TEST FAILING

#### ❌ INCORRECT BEHAVIOR
```python
# EXISTING FAILING TEST
def test_offset_basic():
    result = OFFSET("A1", 1, 1)
    assert result == "B2"  # FAILS: returns RefError

# INCORRECT RESPONSE:
def test_offset_simple():  # ❌ CREATE NEW TEST THAT PASSES
    result = OFFSET("A1", 0, 0)
    assert result == "A1"  # ❌ AVOID THE PROBLEM

# OR:
def test_offset_basic():
    result = OFFSET("A1", 1, 1)
    assert isinstance(result, RefError)  # ❌ CHANGE EXPECTATION
```

#### ✅ CORRECT BEHAVIOR
```python
# EXISTING FAILING TEST
def test_offset_basic():
    result = OFFSET("A1", 1, 1)
    assert result == "B2"  # FAILS: returns RefError

# CORRECT ATDD PROCESS:
# 1. VALIDATE: Does OFFSET("A1", 1, 1) in Excel return "B2"? YES
# 2. PROBLEM: Implementation doesn't calculate offset correctly
# 3. MINIMAL FIX:
def OFFSET(reference, rows, cols):
    if reference == "A1" and rows == 1 and cols == 1:
        # Calculate: A1 + 1 row + 1 column = B2
        return "B2"
    raise NotImplementedError()

# 4. TEST PASSES → COMMIT: "🟢 Make test_offset_basic pass"
```

### 🔴 CASE 3: MULTIPLE TESTS FAILING

#### ❌ INCORRECT BEHAVIOR
```python
# SITUATION: 3 INDEX tests fail
def test_index_basic():
    assert INDEX(range_data, 1, 1) == "A1"  # FAILS

def test_index_row():
    assert INDEX(range_data, 2, 1) == "A2"  # FAILS

def test_index_col():
    assert INDEX(range_data, 1, 2) == "B1"  # FAILS

# INCORRECT RESPONSE:
def test_index_working():  # ❌ CREATE TEST THAT PASSES
    assert INDEX([[1]], 1, 1) == 1  # ❌ AVOID COMPLEX TESTS

# OR:
def INDEX(array, row, col):  # ❌ PREMATURE COMPLEX IMPLEMENTATION
    # Implement all functionality at once
    return complex_implementation(array, row, col)
```

#### ✅ CORRECT BEHAVIOR
```python
# SITUATION: 3 INDEX tests fail
def test_index_basic():
    assert INDEX(range_data, 1, 1) == "A1"  # FAILS

def test_index_row():
    assert INDEX(range_data, 2, 1) == "A2"  # FAILS

def test_index_col():
    assert INDEX(range_data, 1, 2) == "B1"  # FAILS

# CORRECT ATDD PROCESS:
# 1. FOCUS ON FIRST TEST ONLY
# 2. VALIDATE: Does INDEX(range_data, 1, 1) in Excel return "A1"? YES
# 3. MINIMAL FIX FOR FIRST TEST:
def INDEX(array, row, col):
    if row == 1 and col == 1:
        return array[0][0]  # "A1"
    raise NotImplementedError()

# 4. FIRST TEST PASSES → COMMIT
# 5. SECOND TEST:
def INDEX(array, row, col):
    if row == 1 and col == 1:
        return array[0][0]  # "A1"
    elif row == 2 and col == 1:
        return array[1][0]  # "A2"
    raise NotImplementedError()

# 6. SECOND TEST PASSES → COMMIT
# 7. THIRD TEST... and so on
```

### 🔴 CASE 4: TEST WITH DUBIOUS EXCEL BEHAVIOR

#### ❌ INCORRECT BEHAVIOR
```python
# TEST THAT MIGHT BE INCORRECT
def test_vlookup_edge_case():
    result = VLOOKUP("value", data, 0, False)  # col_index = 0
    assert result == "found"  # Is this correct in Excel?

# INCORRECT RESPONSE:
def VLOOKUP(lookup_value, table, col_index, exact):
    if col_index == 0:  # ❌ ASSUME WITHOUT VALIDATION
        return "found"
    # ...
```

#### ✅ CORRECT BEHAVIOR
```python
# TEST THAT MIGHT BE INCORRECT
def test_vlookup_edge_case():
    result = VLOOKUP("value", data, 0, False)  # col_index = 0
    assert result == "found"  # Is this correct in Excel?

# CORRECT ATDD PROCESS:
# 1. VALIDATE AGAINST EXCEL: Is VLOOKUP with col_index=0 valid?
# RESULT: Excel returns #VALUE! error for col_index < 1

# 2. TEST IS INCORRECT - CORRECT:
def test_vlookup_edge_case():
    with pytest.raises(ValueExcelError):  # ✅ REAL EXCEL BEHAVIOR
        VLOOKUP("value", data, 0, False)

# 3. IMPLEMENT CORRECT BEHAVIOR:
def VLOOKUP(lookup_value, table, col_index, exact):
    if col_index < 1:
        raise ValueExcelError("Column index must be >= 1")
    # ...

# 4. DOCUMENT CHANGE IN ona-memory/
```

### 🔴 CASE 5: REFACTORING AFTER GREEN

#### ❌ INCORRECT BEHAVIOR
```python
# SITUATION: 3 tests pass with duplicated implementation
def FUNCTION(param):
    if param == "case1":
        return calculate_result_case1()
    elif param == "case2":
        return calculate_result_case2()
    elif param == "case3":
        return calculate_result_case3()

# INCORRECT RESPONSE:
def FUNCTION(param):  # ❌ CHANGE BEHAVIOR DURING REFACTOR
    # Add new functionality during refactor
    if param == "case4":  # ❌ NEW BEHAVIOR
        return "new_feature"
    return general_implementation(param)
```

#### ✅ CORRECT BEHAVIOR
```python
# SITUATION: 3 tests pass with duplicated implementation
def FUNCTION(param):
    if param == "case1":
        return calculate_result_case1()
    elif param == "case2":
        return calculate_result_case2()
    elif param == "case3":
        return calculate_result_case3()

# CORRECT REFACTOR PROCESS:
# 1. IDENTIFY COMMON PATTERN
# 2. EXTRACT GENERAL LOGIC
def FUNCTION(param):
    # Same cases, cleaner implementation
    case_map = {
        "case1": calculate_result_case1,
        "case2": calculate_result_case2,
        "case3": calculate_result_case3
    }
    if param in case_map:
        return case_map[param]()
    raise NotImplementedError()

# 3. VERIFY ALL TESTS STILL PASS
# 4. COMMIT: "🔵 Refactor FUNCTION - eliminate duplication"
```

## 🎯 AUTOMATIC RESPONSE PATTERNS

### When User Says: "This test fails"
```
AUTOMATIC ONA RESPONSE:
1. "Running test to confirm failure..."
2. "Analyzing error message..."
3. "Validating expected behavior against Excel..."
4. "Implementing minimal fix to make test pass..."
5. "Verifying test passes..."
6. "Commit: 🟢 Make [test_name] pass"

NEVER RESPOND:
- "Should I create a new test?"
- "We could modify the test to make it pass"
- "Let's implement all the functionality"
```

### When User Says: "Fix this test"
```
AUTOMATIC ONA INTERPRETATION:
"Fix implementation to make this test pass"

NEVER INTERPRET AS:
- "Create new test"
- "Modify existing test"
- "Implement complete functionality"

AUTOMATIC ACTION:
1. Validate test against Excel
2. Fix minimal implementation
3. Verify test passes
4. Immediate commit
```

### When Multiple Tests Fail:
```
AUTOMATIC ONA BEHAVIOR:
1. "Detected X failing tests"
2. "Focusing on first test: [test_name]"
3. "Validating expected behavior..."
4. "Implementing fix for first test only..."
5. "Test passes → Commit"
6. "Continuing with second test..."

NEVER DO:
- Implement fix for all tests at once
- Create alternative tests
- Modify multiple tests
```

## 🚨 AUTOMATIC ALERTS FOR ONA

### 🚨 ALERT: ATDD Violation Detection
```
TRIGGER: Ona attempts to create new test while tests are failing
MESSAGE: "❌ ATDD VIOLATION DETECTED: Must fix existing test first"
ACTION: Stop operation, focus on failing test
```

### 🚨 ALERT: Test Modification Without Validation
```
TRIGGER: Ona attempts to change test expectation
MESSAGE: "❌ VALIDATION REQUIRED: Consult Excel before modifying test"
ACTION: Pause, validate against Excel documentation
```

### 🚨 ALERT: Premature Complex Implementation
```
TRIGGER: Ona implements more functionality than required by current test
MESSAGE: "⚠️ EXCESSIVE IMPLEMENTATION: Only minimal for current test"
ACTION: Simplify implementation
```

## 📋 AUTOMATIC CHECKLIST FOR ONA

### Before Any Action:
- [ ] Are tests failing? → If YES, apply ATDD protocol
- [ ] Have I validated test against Excel? → Mandatory before implementing
- [ ] Is my implementation minimal? → Only for current test
- [ ] Am I creating new test? → PROHIBITED during red phase

### After Implementing:
- [ ] Does specific test pass? → Verify before continuing
- [ ] Do other tests still pass? → Don't break existing functionality
- [ ] Are all tests green? → Mandatory before refactor
- [ ] Commit performed? → Immediate when test passes

---

## 🎯 SUMMARY FOR ONA

**YOUR BEHAVIOR MUST BE PREDICTABLE:**
- Test fails → Validate → Fix implementation → Test passes → Commit
- NEVER create new test to avoid problem
- NEVER modify test without Excel validation
- NEVER advance without complete green

**THESE EXAMPLES DEFINE YOUR CORRECT AND INCORRECT BEHAVIOR**