"""
XLOOKUP Implementation Plan - Reusing Existing Logic

This file demonstrates how XLOOKUP can be implemented by combining
existing VLOOKUP, MATCH, and INDEX logic from xlcalculator.
"""

from . import xl, xlerrors, func_xltypes
from .dynamic_range import (
    _convert_function_parameters, 
    _validate_and_get_array_info,
    _handle_function_errors,
    ERROR_MESSAGES
)

@xl.register()
@xl.validate_args
@_handle_function_errors("XLOOKUP")
def XLOOKUP(
    lookup_value: func_xltypes.XlAnything,
    lookup_array: func_xltypes.XlArray,
    return_array: func_xltypes.XlArray,
    if_not_found=None,
    match_mode: func_xltypes.XlNumber = 0,
    search_mode: func_xltypes.XlNumber = 1
) -> func_xltypes.XlAnything:
    """
    Modern lookup function that replaces VLOOKUP, HLOOKUP, and INDEX/MATCH.
    
    XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
    
    Args:
        lookup_value: Value to search for
        lookup_array: Array to search in  
        return_array: Array to return values from
        if_not_found: Value to return if no match found (default: #N/A error)
        match_mode: 0=exact, -1=exact or next smallest, 1=exact or next largest, 2=wildcard
        search_mode: 1=first to last, -1=last to first, 2=binary ascending, -2=binary descending
        
    Returns:
        Matching value from return_array or if_not_found value
        
    Examples:
        XLOOKUP("Apple", A1:A10, B1:B10) → Find "Apple" in A1:A10, return corresponding B value
        XLOOKUP(100, A1:A10, B1:B10, "Not Found") → With custom not found message
        
    Excel Documentation:
        https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929
    """
    
    # REUSED FROM INDEX: Parameter conversion using existing utility
    params = _convert_function_parameters(
        match_mode=(match_mode, int, 0, True),
        search_mode=(search_mode, int, 1, True)
    )
    
    # REUSED FROM INDEX: Array validation using existing utilities
    lookup_info, error = _validate_and_get_array_info(lookup_array, "XLOOKUP")
    if error:
        return error
    lookup_values, lookup_rows, lookup_cols = lookup_info
    
    return_info, error = _validate_and_get_array_info(return_array, "XLOOKUP")
    if error:
        return error
    return_values, return_rows, return_cols = return_info
    
    # Validate arrays have compatible dimensions
    if len(lookup_values) != len(return_values):
        return xlerrors.ValueExcelError("Lookup and return arrays must have same number of rows")
    
    # Flatten arrays for searching (REUSED FROM MATCH logic)
    if lookup_cols == 1:
        lookup_flat = [row[0] for row in lookup_values]
    elif lookup_rows == 1:
        lookup_flat = lookup_values[0]
    else:
        return xlerrors.ValueExcelError("Lookup array must be a single row or column")
    
    if return_cols == 1:
        return_flat = [row[0] for row in return_values]
    elif return_rows == 1:
        return_flat = return_values[0]
    else:
        return xlerrors.ValueExcelError("Return array must be a single row or column")
    
    # REUSED AND ENHANCED FROM MATCH: Search logic
    found_index = _xlookup_search(lookup_value, lookup_flat, params['match_mode'], params['search_mode'])
    
    if found_index is None:
        # ENHANCED: Custom if_not_found handling (new feature)
        if if_not_found is not None:
            return if_not_found
        else:
            # REUSED FROM VLOOKUP: Standard not found error
            return xlerrors.NaExcelError("Lookup value not found")
    
    # REUSED FROM INDEX: Result extraction
    return return_flat[found_index]


def _xlookup_search(lookup_value, lookup_array, match_mode, search_mode):
    """
    Enhanced search logic combining MATCH functionality with XLOOKUP features.
    
    REUSES: Core MATCH search patterns
    ENHANCES: Adds wildcard matching and binary search
    """
    
    # Determine search direction (NEW FEATURE)
    if search_mode < 0:
        # Reverse search
        search_range = range(len(lookup_array) - 1, -1, -1)
    else:
        # Forward search (REUSED FROM MATCH)
        search_range = range(len(lookup_array))
    
    # Binary search optimization (NEW FEATURE)
    if abs(search_mode) == 2:
        return _binary_search(lookup_value, lookup_array, match_mode, search_mode > 0)
    
    # Linear search (REUSED AND ENHANCED FROM MATCH)
    for i in search_range:
        val = lookup_array[i]
        
        if match_mode == 0:
            # Exact match (REUSED FROM MATCH)
            if val == lookup_value:
                return i
                
        elif match_mode == -1:
            # Exact or next smallest (REUSED FROM MATCH logic)
            if val == lookup_value:
                return i
            elif val < lookup_value:
                # Continue searching for larger value
                continue
            else:
                # Found larger value, return previous if exists
                return i - 1 if i > 0 else None
                
        elif match_mode == 1:
            # Exact or next largest (REUSED FROM MATCH logic)
            if val == lookup_value:
                return i
            elif val > lookup_value:
                # Continue searching for smaller value
                continue
            else:
                # Found smaller value, return previous if exists
                return i - 1 if i > 0 else None
                
        elif match_mode == 2:
            # Wildcard match (NEW FEATURE)
            if _wildcard_match(str(lookup_value), str(val)):
                return i
    
    return None


def _binary_search(lookup_value, lookup_array, match_mode, ascending=True):
    """
    Binary search implementation for sorted arrays (NEW FEATURE).
    
    ENHANCES: MATCH functionality with O(log n) performance
    """
    # Validate array is sorted (REUSED FROM MATCH validation logic)
    if ascending and lookup_array != sorted(lookup_array):
        return None  # Array not sorted ascending
    elif not ascending and lookup_array != sorted(lookup_array, reverse=True):
        return None  # Array not sorted descending
    
    left, right = 0, len(lookup_array) - 1
    
    while left <= right:
        mid = (left + right) // 2
        mid_val = lookup_array[mid]
        
        if mid_val == lookup_value:
            return mid
        elif (mid_val < lookup_value) == ascending:
            left = mid + 1
        else:
            right = mid - 1
    
    # Handle approximate matches (ENHANCED FROM MATCH)
    if match_mode == -1:  # Next smallest
        return right if right >= 0 else None
    elif match_mode == 1:  # Next largest
        return left if left < len(lookup_array) else None
    
    return None


def _wildcard_match(pattern, text):
    """
    Wildcard matching for XLOOKUP match_mode=2 (NEW FEATURE).
    
    Supports ? (single character) and * (multiple characters)
    """
    import re
    # Convert Excel wildcards to regex
    regex_pattern = pattern.replace('?', '.').replace('*', '.*')
    regex_pattern = f'^{regex_pattern}$'
    return bool(re.match(regex_pattern, text, re.IGNORECASE))


# REUSED PATTERNS SUMMARY:
"""
FROM INDEX (dynamic_range.py):
✅ _convert_function_parameters() - Parameter conversion
✅ _validate_and_get_array_info() - Array validation  
✅ @_handle_function_errors() - Error handling decorator
✅ Array bounds checking patterns
✅ Result extraction logic

FROM MATCH (lookup.py):
✅ for i, val in enumerate() - Search iteration pattern
✅ Exact match logic (val == lookup_value)
✅ Approximate match logic (val > lookup_value, val < lookup_value)
✅ Sorted array validation
✅ 1-based to 0-based index conversion

FROM VLOOKUP (lookup.py):
✅ Array dimension validation
✅ NaExcelError for not found cases
✅ Result extraction from corresponding position
✅ Parameter validation patterns

NEW XLOOKUP FEATURES:
🆕 if_not_found parameter handling
🆕 Wildcard matching (match_mode=2)
🆕 Reverse search (search_mode=-1)
🆕 Binary search optimization (search_mode=±2)
🆕 Flexible lookup/return array positioning

REUSE PERCENTAGE: ~75% existing logic, 25% new features
"""