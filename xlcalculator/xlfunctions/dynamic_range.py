"""
Dynamic Range Functions for Excel Compatibility

This module implements Excel's dynamic range functions (OFFSET, INDEX, INDIRECT)
using the standard xlcalculator function registration pattern.

Functions implemented:
- OFFSET: Returns reference offset from starting reference
- INDEX: Returns value at array intersection  
- INDIRECT: Returns reference from text string
"""

from typing import Union, Optional
from . import xl, xlerrors, func_xltypes
from .reference_utils import ReferenceResolver


@xl.register()
@xl.validate_args
def OFFSET(
    reference: func_xltypes.XlAnything,
    rows: func_xltypes.XlNumber,
    cols: func_xltypes.XlNumber,
    height: func_xltypes.XlNumber = None,
    width: func_xltypes.XlNumber = None
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
    try:
        # Convert parameters to appropriate types
        ref_str = str(reference)
        rows_int = int(rows)
        cols_int = int(cols)
        height_int = int(height) if height is not None else None
        width_int = int(width) if width is not None else None
        
        # Use reference utilities to calculate offset
        result_ref = ReferenceResolver.offset_reference(
            ref_str, rows_int, cols_int, height_int, width_int
        )
        
        return result_ref
        
    except (xlerrors.RefExcelError, xlerrors.ValueExcelError):
        # Re-raise Excel errors as-is
        raise
    except Exception as e:
        # Convert unexpected errors to Excel errors
        raise xlerrors.ValueExcelError(f"OFFSET error: {str(e)}")


@xl.register()
@xl.validate_args
def INDEX(
    array: func_xltypes.XlArray,
    row_num: func_xltypes.XlNumber,
    col_num: func_xltypes.XlNumber = None
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
    try:
        # Convert parameters
        row_num_int = int(row_num)
        col_num_int = int(col_num) if col_num is not None else 1
        
        # Validate parameters
        if row_num_int < 0 or col_num_int < 0:
            raise xlerrors.ValueExcelError("Row and column numbers must be non-negative")
        
        # Get array data
        if not hasattr(array, 'values') or not array.values:
            raise xlerrors.ValueExcelError("Array is empty or invalid")
        
        array_values = array.values
        num_rows = len(array_values)
        num_cols = len(array_values[0]) if num_rows > 0 else 0
        
        # Handle special cases for 0 (entire row/column)
        if row_num_int == 0 and col_num_int == 0:
            raise xlerrors.ValueExcelError("Both row_num and col_num cannot be 0")
        
        if row_num_int == 0:
            # Return entire column
            if col_num_int < 1 or col_num_int > num_cols:
                raise xlerrors.RefExcelError(f"Column {col_num_int} is out of range (1-{num_cols})")
            return [row[col_num_int - 1] for row in array_values]
        
        if col_num_int == 0:
            # Return entire row
            if row_num_int < 1 or row_num_int > num_rows:
                raise xlerrors.RefExcelError(f"Row {row_num_int} is out of range (1-{num_rows})")
            return array_values[row_num_int - 1]
        
        # Return single value
        if row_num_int < 1 or row_num_int > num_rows:
            raise xlerrors.RefExcelError(f"Row {row_num_int} is out of range (1-{num_rows})")
        if col_num_int < 1 or col_num_int > num_cols:
            raise xlerrors.RefExcelError(f"Column {col_num_int} is out of range (1-{num_cols})")
        
        return array_values[row_num_int - 1][col_num_int - 1]
        
    except (xlerrors.RefExcelError, xlerrors.ValueExcelError):
        # Re-raise Excel errors as-is
        raise
    except Exception as e:
        # Convert unexpected errors to Excel errors
        raise xlerrors.ValueExcelError(f"INDEX error: {str(e)}")


@xl.register()
@xl.validate_args
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
    try:
        # Convert parameter
        ref_str = str(ref_text).strip()
        
        if not ref_str:
            raise xlerrors.NameExcelError("Reference text cannot be empty")
        
        # Check reference style
        if not a1:
            raise NotImplementedError("R1C1 reference style is not yet supported")
        
        # Validate the reference format by attempting to parse it
        try:
            if ':' in ref_str:
                # Range reference
                ReferenceResolver.parse_range_reference(ref_str)
            else:
                # Single cell reference
                ReferenceResolver.parse_cell_reference(ref_str)
        except (xlerrors.ValueExcelError, xlerrors.RefExcelError):
            # Invalid reference format
            raise xlerrors.NameExcelError(f"Invalid reference: '{ref_str}'")
        
        # Return the normalized reference string
        # The evaluator will resolve this to the actual cell/range value
        return ReferenceResolver.normalize_reference(ref_str)
        
    except xlerrors.NameExcelError:
        # Re-raise name errors as-is
        raise
    except Exception as e:
        # Convert unexpected errors to name errors
        raise xlerrors.NameExcelError(f"INDIRECT error: {str(e)}")


# Additional helper functions for advanced features

def _validate_array_parameter(array, function_name: str):
    """
    Validate that array parameter is properly formatted.
    
    Args:
        array: Array parameter to validate
        function_name: Name of calling function for error messages
        
    Raises:
        ValueExcelError: If array is invalid
    """
    if not hasattr(array, 'values'):
        raise xlerrors.ValueExcelError(f"{function_name}: Array parameter is invalid")
    
    if not array.values:
        raise xlerrors.ValueExcelError(f"{function_name}: Array is empty")
    
    # Ensure all rows have the same length
    if len(array.values) > 1:
        first_row_len = len(array.values[0])
        for i, row in enumerate(array.values[1:], 1):
            if len(row) != first_row_len:
                raise xlerrors.ValueExcelError(
                    f"{function_name}: Array rows have inconsistent lengths"
                )


def _get_array_dimensions(array) -> tuple[int, int]:
    """
    Get the dimensions of an array.
    
    Args:
        array: Array to measure
        
    Returns:
        Tuple of (num_rows, num_cols)
    """
    if not array.values:
        return 0, 0
    
    num_rows = len(array.values)
    num_cols = len(array.values[0]) if num_rows > 0 else 0
    return num_rows, num_cols


# Future functions to implement:
# - HLOOKUP: Horizontal lookup (similar to VLOOKUP)
# - TRANSPOSE: Array transposition
# - XLOOKUP: Modern lookup function (Excel 365)
# - FILTER: Filter array based on criteria (Excel 365)
# - SORT: Sort array (Excel 365)
# - UNIQUE: Return unique values (Excel 365)