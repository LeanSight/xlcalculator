# Test Case Correction Log

## Date: 2025-09-09

### Issue: OFFSET Function Test Cases Incorrect

**Problem**: Test cases F3 and F4 in OFFSET error testing were expecting #REF! errors for valid Excel references.

**Original Test Cases (INCORRECT)**:
- `F3: =OFFSET(Data!A1, 100, 0)` → Expected #REF! 
- `F4: =OFFSET(Data!A1, 0, 100)` → Expected #REF!

**Issue Analysis**:
- Row 101 and Column 101 are VALID Excel coordinates (within 1-1,048,576 rows and 1-16,384 columns)
- According to Microsoft Excel documentation: "If rows and cols offset reference over the edge of the worksheet, OFFSET returns the #REF! error value"
- "Edge of worksheet" means Excel's sheet limits, not data limits

**Corrected Test Cases (CORRECT)**:
- `F3: =OFFSET(Data!A1, 1048576, 0)` → #REF! (exceeds Excel row limit)
- `F4: =OFFSET(Data!A1, 0, 16384)` → #REF! (exceeds Excel column limit)

**Excel Documentation Reference**:
- Source: https://support.microsoft.com/en-us/office/offset-function-c8de19ae-dd79-4b9b-a14e-b4d906d11b66
- Key Quote: "If rows and cols offset reference over the edge of the worksheet, OFFSET returns the #REF! error value"

**Files Updated**:
1. `tests/resources_generator/DYNAMIC_RANGES_DESIGN.md` - Updated test case specifications
2. `tests/resources_generator/dynamic_range_test_cases.json` - Updated JSON test definitions

**Next Steps**:
1. Regenerate Excel test files with corrected formulas
2. Regenerate Python test classes
3. Verify OFFSET implementation handles Excel limits correctly
4. Run updated tests to ensure compliance

**Lesson Learned**:
Always validate test case expectations against official Excel documentation before implementing fixes. Test cases must represent legitimate Excel behavior, not assumptions about how functions should work.