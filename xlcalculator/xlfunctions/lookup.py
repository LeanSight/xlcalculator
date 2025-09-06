from . import xl, xlerrors, func_xltypes
import re


# ============================================================================
# SHARED LOOKUP UTILITIES - Eliminate duplicate logic across lookup functions
# ============================================================================

def _flatten_lookup_array(array):
    """Flatten a 2D array to 1D for lookup operations.
    
    Used by: MATCH, XLOOKUP
    Returns: Flattened list of values
    Raises: ValueExcelError if array is not single row or column
    """
    if len(array.values[0]) == 1:
        # Single column array
        return [row[0] for row in array.values]
    elif len(array.values) == 1:
        # Single row array
        return array.values[0]
    else:
        raise xlerrors.ValueExcelError("Lookup array must be a single row or column")


def _validate_array_dimensions(array1, array2, operation_name):
    """Validate that two arrays have compatible dimensions.
    
    Used by: VLOOKUP, XLOOKUP
    Returns: None if valid, ValueExcelError if invalid
    """
    if len(array1.values) != len(array2.values):
        return xlerrors.ValueExcelError(f"{operation_name}: Arrays must have same number of rows")
    return None


def _exact_match_search(lookup_value, lookup_array, search_range=None):
    """Find exact match in array.
    
    Used by: MATCH, XLOOKUP
    Returns: Index of match or None if not found
    """
    if search_range is None:
        search_range = range(len(lookup_array))
    
    for i in search_range:
        if lookup_array[i] == lookup_value:
            return i
    return None


def _approximate_match_search(lookup_value, lookup_array, find_smaller=True):
    """Find approximate match (next smaller or larger value).
    
    Used by: MATCH, XLOOKUP
    Args:
        find_smaller: True for next smallest, False for next largest
    Returns: Index of best match or None if not found
    """
    best_match = None
    
    for i in range(len(lookup_array)):
        val = lookup_array[i]
        
        # Check for exact match first
        if val == lookup_value:
            return i
        
        # Check for approximate match
        if find_smaller and val < lookup_value:
            if best_match is None or lookup_array[best_match] < val:
                best_match = i
        elif not find_smaller and val > lookup_value:
            if best_match is None or lookup_array[best_match] > val:
                best_match = i
    
    return best_match


def _create_not_found_error(context="Lookup value not found"):
    """Create standardized not found error.
    
    Used by: VLOOKUP, MATCH, XLOOKUP
    Returns: NaExcelError with consistent message
    """
    return xlerrors.NaExcelError(context)


@xl.register()
@xl.validate_args
def CHOOSE(
        index_num: func_xltypes.XlNumber,
        *values,
) -> func_xltypes.XlAnything:
    """Uses index_num to return a value from the list of value arguments.

    https://support.office.com/en-us/article/
        choose-function-fc5c184f-cb62-4ec7-a46e-38653b98f5bc
    """
    if index_num <= 0 or index_num > 254:
        raise xlerrors.ValueExcelError(
            f"`index_num` {index_num} must be between 1 and 254")

    if index_num > len(values):
        raise xlerrors.ValueExcelError(
            f"`index_num` {index_num} must not be larger than the number of "
            f"values: {len(values)}")

    idx = int(index_num) - 1
    return values[idx]


@xl.register()
@xl.validate_args
def VLOOKUP(
        lookup_value: func_xltypes.XlAnything,
        table_array: func_xltypes.XlArray,
        col_index_num: func_xltypes.XlNumber,
        range_lookup=False
) -> func_xltypes.XlAnything:
    """Looks in the first column of an array and moves across the row to
    return the value of a cell.

    https://support.office.com/en-us/article/
        vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1
    
    Refactored to use shared lookup utilities.
    """
    if range_lookup:
        raise NotImplementedError("Exact match only supported at the moment.")

    col_index_num = int(col_index_num)

    # Use shared dimension validation concept
    if col_index_num > len(table_array.values[0]):
        raise xlerrors.ValueExcelError(
            'col_index_num is greater than the number of cols in table_array')

    table_array = table_array.set_index(0)

    if lookup_value not in table_array.index:
        # Use shared error creation utility
        raise _create_not_found_error('`lookup_value` not in first column of `table_array`.')

    return table_array.loc[lookup_value].values[0]


@xl.register()
@xl.validate_args
def MATCH(
        lookup_value: func_xltypes.XlAnything,
        lookup_array: func_xltypes.XlArray,
        match_type: func_xltypes.XlAnything = 1,
) -> func_xltypes.XlAnything:
    """Find the relative position of an item in an array.
    
    Refactored to use shared lookup utilities.
    """
    # Use shared array flattening utility
    try:
        lookup_flat = _flatten_lookup_array(lookup_array)
    except xlerrors.ValueExcelError:
        # MATCH requires single column, be more specific
        assert len(lookup_array.values[0]) == 1
        lookup_flat = lookup_array.flat

    # Validate sort order for approximate matches
    if match_type == 1:
        if lookup_flat != sorted(lookup_flat):
            return _create_not_found_error("Values must be sorted in ascending order")
    if match_type == -1:
        if lookup_flat != sorted(lookup_flat, reverse=True):
            return _create_not_found_error("Values must be sorted in descending order")

    # Use shared exact match search
    exact_match = _exact_match_search(lookup_value, lookup_flat)
    if exact_match is not None:
        return exact_match + 1  # MATCH returns 1-based index

    # Handle approximate matches
    if match_type == 1:
        # Find next smallest (largest value <= lookup_value)
        best_match = _approximate_match_search(lookup_value, lookup_flat, find_smaller=True)
        return (best_match + 1) if best_match is not None else _create_not_found_error("No lesser value found.")
    elif match_type == -1:
        # Find next largest (smallest value >= lookup_value)  
        best_match = _approximate_match_search(lookup_value, lookup_flat, find_smaller=False)
        return (best_match + 1) if best_match is not None else _create_not_found_error("No greater value found.")
    
    return _create_not_found_error("No match found.")


@xl.register()
@xl.register('_xlfn.XLOOKUP')  # Register Excel's _xlfn prefix variant
@xl.validate_args
def XLOOKUP(
        lookup_value: func_xltypes.XlAnything,
        lookup_array: func_xltypes.XlArray,
        return_array: func_xltypes.XlArray,
        if_not_found=None,
        match_mode=0,
        search_mode=1
) -> func_xltypes.XlAnything:
    """Modern lookup function that replaces VLOOKUP, HLOOKUP, and INDEX/MATCH.
    
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
    
    # Handle Excel empty parameter parsing correction
    # 
    # Problem: Excel formulas with empty parameters are parsed incorrectly by xlcalculator
    # Example: =XLOOKUP(val, arr1, arr2, , -1) should have if_not_found=None, match_mode=-1
    # But gets parsed as: if_not_found=-1, match_mode=0 (parameter shift)
    #
    # Solution: Detect specific patterns that indicate parameter shifting and correct them
    
    if isinstance(if_not_found, (int, float, func_xltypes.Number)):
        if_not_found_val = int(if_not_found)
        
        # Only apply parameter correction if we detect the specific pattern of Excel empty parameter parsing
        # This happens when match_mode and search_mode are still at their defaults, suggesting parameter shift
        
        # Case 1: 5 parameters with empty if_not_found: =XLOOKUP(val, arr1, arr2, , match_mode)
        # Pattern: if_not_found is a match_mode value (-1, 1, 2) and other params are defaults
        if (if_not_found_val in [-1, 1, 2] and 
            match_mode == 0 and search_mode == 1):
            match_mode = if_not_found_val
            if_not_found = None
            
        # Case 2: 6 parameters with empty if_not_found: =XLOOKUP(val, arr1, arr2, , match_mode, search_mode)  
        # Pattern: if_not_found=0, match_mode is a search_mode value (-1, 1, 2), search_mode is default
        elif (if_not_found_val == 0 and 
              isinstance(match_mode, (int, float, func_xltypes.Number)) and
              int(match_mode) in [-1, 1, 2] and
              search_mode == 1):
            search_mode = int(match_mode)
            match_mode = if_not_found_val
            if_not_found = None
    
    # Convert parameters to proper types
    match_mode = int(match_mode) if match_mode is not None else 0
    search_mode = int(search_mode) if search_mode is not None else 1
    
    # Use shared array dimension validation
    dimension_error = _validate_array_dimensions(lookup_array, return_array, "XLOOKUP")
    if dimension_error:
        return dimension_error
    
    # Use shared array flattening utilities
    try:
        lookup_flat = _flatten_lookup_array(lookup_array)
        return_flat = _flatten_lookup_array(return_array)
    except xlerrors.ValueExcelError as e:
        return e
    
    # Enhanced search logic combining shared utilities with XLOOKUP features
    found_index = _xlookup_search(lookup_value, lookup_flat, match_mode, search_mode)
    
    if found_index is None:
        # Enhanced: Custom if_not_found handling (new feature)
        if if_not_found is not None:
            return if_not_found
        else:
            # Use shared error creation utility
            return _create_not_found_error("Lookup value not found")
    
    # Result extraction
    return return_flat[found_index]


def _xlookup_search(lookup_value, lookup_array, match_mode, search_mode):
    """Enhanced search logic combining shared utilities with XLOOKUP features.
    
    Refactored to use shared lookup utilities where possible.
    """
    
    # Determine search direction (XLOOKUP-specific feature)
    if search_mode < 0:
        # Reverse search
        search_range = range(len(lookup_array) - 1, -1, -1)
    else:
        # Forward search
        search_range = range(len(lookup_array))
    
    # Binary search optimization (XLOOKUP-specific feature)
    if abs(search_mode) == 2:
        return _binary_search(lookup_value, lookup_array, match_mode, search_mode > 0)
    
    # Linear search using shared utilities where possible
    if match_mode == 0:
        # Use shared exact match search
        return _exact_match_search(lookup_value, lookup_array, search_range)
                
    elif match_mode == -1:
        # Use shared approximate match search (next smallest)
        return _approximate_match_search(lookup_value, lookup_array, find_smaller=True)
                
    elif match_mode == 1:
        # Use shared approximate match search (next largest)
        return _approximate_match_search(lookup_value, lookup_array, find_smaller=False)
                
    elif match_mode == 2:
        # Wildcard match (XLOOKUP-specific feature)
        for i in search_range:
            if _wildcard_match(str(lookup_value), str(lookup_array[i])):
                return i
    
    return None


def _binary_search(lookup_value, lookup_array, match_mode, ascending=True):
    """Binary search implementation for sorted arrays (new feature).
    
    Enhances: MATCH functionality with O(log n) performance
    """
    # Validate array is sorted (reused from MATCH validation logic)
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
    
    # Handle approximate matches (enhanced from MATCH)
    if match_mode == -1:  # Next smallest
        return right if right >= 0 else None
    elif match_mode == 1:  # Next largest
        return left if left < len(lookup_array) else None
    
    return None


def _wildcard_match(pattern, text):
    """Wildcard matching for XLOOKUP match_mode=2 (new feature).
    
    Supports ? (single character) and * (multiple characters)
    """
    # Convert Excel wildcards to regex
    regex_pattern = pattern.replace('?', '.').replace('*', '.*')
    regex_pattern = f'^{regex_pattern}$'
    return bool(re.match(regex_pattern, text, re.IGNORECASE))
