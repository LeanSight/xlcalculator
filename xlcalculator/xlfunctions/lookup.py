from . import xl, xlerrors, func_xltypes
import re


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
    """
    if range_lookup:
        raise NotImplementedError("Excact match only supported at the moment.")

    col_index_num = int(col_index_num)

    if col_index_num > len(table_array.values[0]):
        raise xlerrors.ValueExcelError(
            'col_index_num is greater than the number of cols in table_array')

    table_array = table_array.set_index(0)

    if lookup_value not in table_array.index:
        raise xlerrors.NaExcelError(
            '`lookup_value` not in first column of `table_array`.')

    return table_array.loc[lookup_value].values[0]


@xl.register()
@xl.validate_args
def MATCH(
        lookup_value: func_xltypes.XlAnything,
        lookup_array: func_xltypes.XlArray,
        match_type: func_xltypes.XlAnything = 1,
) -> func_xltypes.XlAnything:
    assert len(lookup_array.values[0]) == 1

    lookup_array = lookup_array.flat

    if match_type == 1:
        if lookup_array != sorted(lookup_array):
            return xlerrors.NaExcelError(
                "Values must be sorted in ascending order"
            )
    if match_type == -1:
        if lookup_array != sorted(lookup_array, reverse=True):
            return xlerrors.NaExcelError(
                "Values must be sorted in descending order"
            )

    for i, val in enumerate(lookup_array):
        if val == lookup_value:
            return i + 1
        if match_type == 1 and val > lookup_value:
            return i or xlerrors.NaExcelError(
                "No lesser value found."
            )
        if match_type == -1 and val < lookup_value:
            return i or xlerrors.NaExcelError(
                "No greater value found."
            )
    return xlerrors.NaExcelError("No match found.")


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
    
    # Handle parameter type conversion and empty parameter detection
    # Excel formulas with empty parameters like =XLOOKUP(val, arr1, arr2, , mode) get parsed incorrectly
    # We need to detect and fix parameter shifting caused by empty parameter parsing
    
    if isinstance(if_not_found, (int, float, func_xltypes.Number)):
        if_not_found_val = int(if_not_found)
        
        # Case 1: 5 parameters with empty if_not_found: =XLOOKUP(val, arr1, arr2, , match_mode)
        if if_not_found_val in [-1, 1, 2] and match_mode == 0 and search_mode == 1:
            match_mode = if_not_found_val
            if_not_found = None
            
        # Case 2: 6 parameters with empty if_not_found: =XLOOKUP(val, arr1, arr2, , match_mode, search_mode)  
        elif if_not_found_val == 0 and isinstance(match_mode, (int, float, func_xltypes.Number)):
            search_mode = int(match_mode)
            match_mode = if_not_found_val
            if_not_found = None
    
    # Convert parameters to proper types
    match_mode = int(match_mode) if match_mode is not None else 0
    search_mode = int(search_mode) if search_mode is not None else 1
    
    # Validate arrays have compatible dimensions (reused from VLOOKUP validation)
    if len(lookup_array.values) != len(return_array.values):
        return xlerrors.ValueExcelError("Lookup and return arrays must have same number of rows")
    
    # Flatten arrays for searching (reused from MATCH logic)
    if len(lookup_array.values[0]) == 1:
        # Single column array
        lookup_flat = [row[0] for row in lookup_array.values]
    elif len(lookup_array.values) == 1:
        # Single row array
        lookup_flat = lookup_array.values[0]
    else:
        return xlerrors.ValueExcelError("Lookup array must be a single row or column")
    
    if len(return_array.values[0]) == 1:
        # Single column array
        return_flat = [row[0] for row in return_array.values]
    elif len(return_array.values) == 1:
        # Single row array
        return_flat = return_array.values[0]
    else:
        return xlerrors.ValueExcelError("Return array must be a single row or column")
    
    # Enhanced search logic combining MATCH functionality with XLOOKUP features
    found_index = _xlookup_search(lookup_value, lookup_flat, match_mode, search_mode)
    
    if found_index is None:
        # Enhanced: Custom if_not_found handling (new feature)
        if if_not_found is not None:
            return if_not_found
        else:
            # Reused from VLOOKUP: Standard not found error
            return xlerrors.NaExcelError("Lookup value not found")
    
    # Reused from INDEX: Result extraction
    return return_flat[found_index]


def _xlookup_search(lookup_value, lookup_array, match_mode, search_mode):
    """Enhanced search logic combining MATCH functionality with XLOOKUP features.
    
    Reuses: Core MATCH search patterns
    Enhances: Adds wildcard matching and binary search
    """
    
    # Determine search direction (new feature)
    if search_mode < 0:
        # Reverse search
        search_range = range(len(lookup_array) - 1, -1, -1)
    else:
        # Forward search (reused from MATCH)
        search_range = range(len(lookup_array))
    
    # Binary search optimization (new feature)
    if abs(search_mode) == 2:
        return _binary_search(lookup_value, lookup_array, match_mode, search_mode > 0)
    
    # Linear search (reused and enhanced from MATCH)
    if match_mode == 0:
        # Exact match (reused from MATCH)
        for i in search_range:
            if lookup_array[i] == lookup_value:
                return i
                
    elif match_mode == -1:
        # Exact or next smallest - find largest value <= lookup_value
        best_match = None
        for i in range(len(lookup_array)):
            val = lookup_array[i]
            if val == lookup_value:
                return i  # Exact match
            elif val < lookup_value:
                if best_match is None or lookup_array[best_match] < val:
                    best_match = i
        return best_match
                
    elif match_mode == 1:
        # Exact or next largest - find smallest value >= lookup_value
        best_match = None
        for i in range(len(lookup_array)):
            val = lookup_array[i]
            if val == lookup_value:
                return i  # Exact match
            elif val > lookup_value:
                if best_match is None or lookup_array[best_match] > val:
                    best_match = i
        return best_match
                
    elif match_mode == 2:
        # Wildcard match (new feature)
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
