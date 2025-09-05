"""
Dynamic Range Functions for Excel Compatibility

This module implements Excel's dynamic range functions (OFFSET, INDEX, INDIRECT)
using the standard xlcalculator function registration pattern.

Functions implemented:
- OFFSET: Returns reference offset from starting reference
- INDEX: Returns value at array intersection  
- INDIRECT: Returns reference from text string
"""

from typing import Union, Optional, Any, Dict, Tuple
from functools import wraps
import threading
from . import xl, xlerrors, func_xltypes
from .reference_utils import ReferenceResolver


# Constants
ERROR_MESSAGES = {
    'EMPTY_REFERENCE': 'Reference text cannot be empty',
    'NEGATIVE_PARAMS': 'Row and column numbers must be non-negative',
    'BOTH_ZERO': 'Both row_num and col_num cannot be 0',
    'EMPTY_ARRAY': 'Array is empty or invalid',
    'R1C1_NOT_SUPPORTED': 'R1C1 reference style is not yet supported',
    'INVALID_REFERENCE': 'Invalid reference: \'{ref}\'',
    'ROW_OUT_OF_RANGE': 'Row {row} is out of range (1-{max_row})',
    'COL_OUT_OF_RANGE': 'Column {col} is out of range (1-{max_col})'
}

DEFAULT_COL_NUM = 1

# Thread-local storage for evaluator context
_context = threading.local()

def _set_evaluator_context(evaluator):
    """Set the current evaluator context for dynamic range functions."""
    _context.evaluator = evaluator

def _get_evaluator_context():
    """Get the current evaluator context, if available."""
    return getattr(_context, 'evaluator', None)

def _clear_evaluator_context():
    """Clear the current evaluator context."""
    if hasattr(_context, 'evaluator'):
        delattr(_context, 'evaluator')


# Utility Functions

def _convert_function_parameters(**params) -> Dict[str, Any]:
    """
    Centralized parameter conversion for dynamic range functions.
    
    Args:
        **params: Dictionary where each key maps to a tuple of:
                 (value, target_type, default, allow_none)
    
    Returns:
        Dictionary of converted parameters
    """
    converted = {}
    for name, config in params.items():
        value, target_type, default, allow_none = config
        if allow_none and (value is None or value == ""):
            converted[name] = default
        elif target_type == int:
            converted[name] = int(value)
        elif target_type == str:
            converted[name] = str(value).strip()
        else:
            converted[name] = value
    return converted


def _handle_function_errors(func_name: str):
    """
    Decorator for consistent error handling across dynamic range functions.
    
    Args:
        func_name: Name of the function for error messages
    
    Returns:
        Decorated function with consistent error handling
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except (xlerrors.RefExcelError, xlerrors.ValueExcelError, xlerrors.NameExcelError):
                # Re-raise Excel errors as-is
                raise
            except Exception as e:
                # Convert unexpected errors to Excel errors
                return xlerrors.ValueExcelError(f"{func_name} error: {str(e)}")
        return wrapper
    return decorator


def _get_array_data(array) -> Optional[list]:
    """
    Extract array data from various input types.
    
    Args:
        array: Array-like object (Mock, list, tuple, or object with .values)
    
    Returns:
        List representation of array data, or None if invalid
    """
    if hasattr(array, 'values') and array.values is not None:
        return array.values
    elif isinstance(array, (list, tuple)):
        return list(array)
    else:
        return None


def _validate_and_get_array_info(array, function_name: str) -> Tuple[Optional[Tuple], Optional[xlerrors.ExcelError]]:
    """
    Validate array and return data with dimensions.
    
    Args:
        array: Array-like object to validate
        function_name: Name of calling function for error messages
    
    Returns:
        Tuple of ((array_data, num_rows, num_cols), error) where error is None if valid
    """
    array_data = _get_array_data(array)
    if array_data is None or len(array_data) == 0:
        return None, xlerrors.ValueExcelError(ERROR_MESSAGES['EMPTY_ARRAY'])
    
    num_rows = len(array_data)
    num_cols = len(array_data[0]) if num_rows > 0 else 0
    return (array_data, num_rows, num_cols), None


def _validate_array_bounds(row_num: int, col_num: int, num_rows: int, num_cols: int) -> Optional[xlerrors.ExcelError]:
    """
    Validate array access bounds and return error if invalid.
    
    Args:
        row_num: Row number to validate (1-based, 0 for entire column)
        col_num: Column number to validate (1-based, 0 for entire row)
        num_rows: Total number of rows in array
        num_cols: Total number of columns in array
    
    Returns:
        Excel error if bounds are invalid, None if valid
    """
    if row_num < 0 or col_num < 0:
        return xlerrors.ValueExcelError(ERROR_MESSAGES['NEGATIVE_PARAMS'])
    
    if row_num == 0 and col_num == 0:
        return xlerrors.ValueExcelError(ERROR_MESSAGES['BOTH_ZERO'])
    
    if row_num > 0 and (row_num < 1 or row_num > num_rows):
        return xlerrors.RefExcelError(ERROR_MESSAGES['ROW_OUT_OF_RANGE'].format(
            row=row_num, max_row=num_rows))
    
    if col_num > 0 and (col_num < 1 or col_num > num_cols):
        return xlerrors.RefExcelError(ERROR_MESSAGES['COL_OUT_OF_RANGE'].format(
            col=col_num, max_col=num_cols))
    
    return None  # No error


def _is_special_range_reference(ref_str: str) -> bool:
    """
    Check if reference is special format like A:A or 1:1.
    
    Args:
        ref_str: Reference string to check
    
    Returns:
        True if reference is special range format
    """
    if ref_str.count(':') == 1:
        left, right = ref_str.split(':')
        return left == right and (left.isalpha() or left.isdigit())
    return False


def _validate_reference_format(ref_str: str) -> Optional[xlerrors.ExcelError]:
    """
    Validate reference format and return error if invalid.
    
    Args:
        ref_str: Reference string to validate
    
    Returns:
        Excel error if invalid, None if valid
    """
    try:
        if ':' in ref_str:
            # Check for special range types (A:A, 1:1, etc.)
            if _is_special_range_reference(ref_str):
                left, _ = ref_str.split(':')
                if not ((left.isalpha() and left.isupper()) or left.isdigit()):
                    return xlerrors.NameExcelError(ERROR_MESSAGES['INVALID_REFERENCE'].format(ref=ref_str))
            else:
                # Regular range reference
                ReferenceResolver.parse_range_reference(ref_str)
        else:
            # Single cell reference
            ReferenceResolver.parse_cell_reference(ref_str)
    except (xlerrors.ValueExcelError, xlerrors.RefExcelError):
        return xlerrors.NameExcelError(ERROR_MESSAGES['INVALID_REFERENCE'].format(ref=ref_str))
    
    return None  # No error


@xl.register()
@xl.validate_args
@_handle_function_errors("OFFSET")
def OFFSET(
    reference: func_xltypes.XlAnything,
    rows: func_xltypes.XlNumber,
    cols: func_xltypes.XlNumber,
    height=None,
    width=None
) -> func_xltypes.XlAnything:
    """
    Returns a reference to a range that is offset from a starting reference.
    
    OFFSET(reference, rows, cols, [height], [width])
    
    Args:
        reference: Starting cell or range reference
        rows: Number of rows to offset (positive=down, negative=up)
        cols: Number of columns to offset (positive=right, negative=left)
        height: Optional height of returned range (default=height of reference)
        width: Optional width of returned range (default=width of reference)
        
    Returns:
        Reference string to the offset range
        
    Raises:
        RefExcelError: When offset result is out of bounds
        ValueExcelError: When parameters are invalid
        
    Examples:
        OFFSET(A1, 1, 1) → "B2"
        OFFSET(A1:B2, 1, 1) → "B2:C3"  
        OFFSET(A1, 1, 1, 2, 2) → "B2:C3"
        OFFSET(A1, -1, 0) → #REF! (out of bounds)
        
    Excel Documentation:
        https://support.microsoft.com/en-us/office/offset-function-c8de19ae-dd79-4b9b-a14e-b4d906d11b66
    """
    # Convert parameters using utility function
    params = _convert_function_parameters(
        reference=(reference, str, None, False),
        rows=(rows, int, None, False),
        cols=(cols, int, None, False),
        height=(height, int, None, True),
        width=(width, int, None, True)
    )
    
    # Use reference utilities to calculate offset
    result_ref = ReferenceResolver.offset_reference(
        params['reference'], params['rows'], params['cols'], 
        params['height'], params['width']
    )
    
    # Try to resolve to actual values if evaluator context is available
    evaluator = _get_evaluator_context()
    if evaluator:
        return _resolve_offset_reference(result_ref, evaluator)
    
    # No evaluator context - return string reference (backward compatible)
    return result_ref


def _resolve_offset_reference(result_ref: str, evaluator) -> func_xltypes.XlAnything:
    """
    Resolve OFFSET reference string to actual values using evaluator.
    
    Args:
        result_ref: Reference string like "B2" or "B2:C3"
        evaluator: Evaluator instance for value resolution
        
    Returns:
        - Single value for single cell references
        - Array object for range references
        - String reference if resolution fails
    """
    try:
        # Ensure reference has sheet name
        if '!' not in result_ref:
            result_ref = f'Sheet1!{result_ref}'
            
        if ':' in result_ref:
            # Range reference - get all values and return as Array
            values = evaluator.get_range_values(result_ref)
            return func_xltypes.Array(values)
        else:
            # Single cell reference - return the value
            return evaluator.get_cell_value(result_ref)
    except Exception:
        # If resolution fails, fallback to string reference
        return result_ref


@xl.register()
@xl.validate_args
@_handle_function_errors("INDEX")
def INDEX(
    array,
    row_num: func_xltypes.XlNumber,
    col_num=None
) -> func_xltypes.XlAnything:
    """
    Returns the value of an element in a table or array, selected by row and column.
    
    INDEX(array, row_num, [col_num])
    
    Args:
        array: Range of cells or array to index into
        row_num: Row number to select (1-based, 0=entire column)
        col_num: Optional column number (1-based, 0=entire row, default=1)
        
    Returns:
        Value at the specified position, or array if row_num/col_num is 0
        
    Raises:
        RefExcelError: When row/column index is out of bounds
        ValueExcelError: When parameters are invalid
        
    Examples:
        INDEX(A1:C3, 2, 2) → Value at B2
        INDEX(A1:C3, 0, 2) → Entire column B as array
        INDEX(A1:C3, 2, 0) → Entire row 2 as array
        INDEX(A1:C3, 4, 1) → #REF! (row out of bounds)
        
    Excel Documentation:
        https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd
    """
    # Convert parameters using utility function
    params = _convert_function_parameters(
        row_num=(row_num, int, None, False),
        col_num=(col_num, int, DEFAULT_COL_NUM, True)
    )
    
    # Validate and get array information
    array_info, error = _validate_and_get_array_info(array, "INDEX")
    if error:
        return error
    
    array_values, num_rows, num_cols = array_info
    
    # Validate array bounds
    bounds_error = _validate_array_bounds(params['row_num'], params['col_num'], num_rows, num_cols)
    if bounds_error:
        return bounds_error
    
    # Handle special cases for 0 (entire row/column)
    if params['row_num'] == 0:
        # Return entire column
        result = [row[params['col_num'] - 1] for row in array_values]
        return func_xltypes.Array([result])
    
    if params['col_num'] == 0:
        # Return entire row
        result = array_values[params['row_num'] - 1]
        return func_xltypes.Array([result])
    
    # Return single value
    return array_values[params['row_num'] - 1][params['col_num'] - 1]


@xl.register()
@xl.validate_args
@_handle_function_errors("INDIRECT")
def INDIRECT(
    ref_text: func_xltypes.XlText,
    a1: func_xltypes.XlBoolean = True
) -> func_xltypes.XlAnything:
    """
    Returns the reference specified by a text string.
    
    INDIRECT(ref_text, [a1])
    
    Args:
        ref_text: Text string containing a cell reference
        a1: Optional reference style (True=A1 style, False=R1C1 style, default=True)
        
    Returns:
        Reference string that can be evaluated by the Excel engine
        
    Raises:
        NameExcelError: When reference text is invalid
        ValueExcelError: When parameters are invalid
        
    Examples:
        INDIRECT("B2") → Reference to B2
        INDIRECT("Sheet2!A1") → Reference to A1 on Sheet2
        INDIRECT(A1) → If A1 contains "B2", returns reference to B2
        INDIRECT("InvalidRef") → #NAME! error
        
    Notes:
        - R1C1 reference style is not yet supported
        - The returned reference is validated but not resolved until evaluation
        
    Excel Documentation:
        https://support.microsoft.com/en-us/office/indirect-function-474b3a3a-8a26-4f44-b491-92b6306fa261
    """
    # Convert parameters using utility function
    params = _convert_function_parameters(
        ref_text=(ref_text, str, None, False)
    )
    
    ref_str = params['ref_text']
    
    if not ref_str:
        return xlerrors.NameExcelError(ERROR_MESSAGES['EMPTY_REFERENCE'])
    
    # Check reference style
    if not a1:
        return xlerrors.ValueExcelError(ERROR_MESSAGES['R1C1_NOT_SUPPORTED'])
    
    # Validate the reference format
    validation_error = _validate_reference_format(ref_str)
    if validation_error:
        return validation_error
    
    # Return the normalized reference string
    # Handle special cases that normalize_reference doesn't support
    if _is_special_range_reference(ref_str):
        return ref_str
    
    return ReferenceResolver.normalize_reference(ref_str)


# Additional helper functions for advanced features

def _validate_array_consistency(array_data: list, function_name: str) -> Optional[xlerrors.ExcelError]:
    """
    Validate that array rows have consistent lengths.
    
    Args:
        array_data: Array data to validate
        function_name: Name of calling function for error messages
        
    Returns:
        Excel error if inconsistent, None if valid
    """
    if len(array_data) > 1:
        first_row_len = len(array_data[0])
        for i, row in enumerate(array_data[1:], 1):
            if len(row) != first_row_len:
                return xlerrors.ValueExcelError(
                    f"{function_name}: Array rows have inconsistent lengths"
                )
    return None


# Future functions to implement:
# - HLOOKUP: Horizontal lookup (similar to VLOOKUP)
# - TRANSPOSE: Array transposition
# - XLOOKUP: Modern lookup function (Excel 365)
# - FILTER: Filter array based on criteria (Excel 365)
# - SORT: Sort array (Excel 365)
# - UNIQUE: Return unique values (Excel 365)