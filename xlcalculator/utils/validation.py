"""Core validation utilities for Excel functions."""

from ..xlfunctions import xlerrors
from ..constants import EXCEL_MAX_ROWS, EXCEL_MAX_COLUMNS


def validate_integer_parameter(value, param_name, min_value=None, max_value=None):
    """Validate and convert parameter to integer with bounds checking.
    
    Args:
        value: Value to validate and convert
        param_name: Name of parameter for error messages
        min_value: Minimum allowed value (inclusive)
        max_value: Maximum allowed value (inclusive)
        
    Returns:
        int: Validated integer value
        
    Raises:
        ValueExcelError: If value cannot be converted or is out of bounds
    """
    try:
        int_value = int(value)
    except (ValueError, TypeError):
        raise xlerrors.ValueExcelError(f"Invalid {param_name}: {value}")
    
    if min_value is not None and int_value < min_value:
        raise xlerrors.ValueExcelError(f"{param_name} must be >= {min_value}")
    
    if max_value is not None and int_value > max_value:
        raise xlerrors.ValueExcelError(f"{param_name} must be <= {max_value}")
    
    return int_value


def validate_positive_integer(value, param_name):
    """Validate parameter is a positive integer (>= 1).
    
    Args:
        value: Value to validate
        param_name: Name of parameter for error messages
        
    Returns:
        int: Validated positive integer
        
    Raises:
        ValueExcelError: If value is not a positive integer
    """
    return validate_integer_parameter(value, param_name, min_value=1)


def validate_array_bounds(array_data, row_idx, col_idx, row_name="row", col_name="column"):
    """Validate array bounds for row and column indices.
    
    Args:
        array_data: 2D array to validate bounds against
        row_idx: Row index to validate (0-based)
        col_idx: Column index to validate (0-based)
        row_name: Name for row in error messages
        col_name: Name for column in error messages
        
    Raises:
        RefExcelError: If indices are out of bounds
    """
    if not array_data:
        raise xlerrors.RefExcelError("Array data is empty")
    
    if row_idx < 0 or row_idx >= len(array_data):
        raise xlerrors.RefExcelError(f"{row_name.title()} index out of range")
    
    if not array_data[0]:
        raise xlerrors.RefExcelError("Array row is empty")
    
    if col_idx < 0 or col_idx >= len(array_data[0]):
        raise xlerrors.RefExcelError(f"{col_name.title()} index out of range")


def validate_dimension_parameter(value, param_name):
    """Validate parameter is a positive dimension (> 0).
    
    Args:
        value: Value to validate
        param_name: Name of parameter for error messages
        
    Returns:
        int: Validated dimension value
        
    Raises:
        ValueExcelError: If value is not a positive dimension
    """
    int_value = validate_integer_parameter(value, param_name, min_value=1)
    if int_value <= 0:
        raise xlerrors.ValueExcelError(f"{param_name} must be > 0")
    return int_value


def validate_area_number(area_num, total_areas, param_name="area number"):
    """Validate area number is within valid range.
    
    Args:
        area_num: Area number to validate (1-based)
        total_areas: Total number of areas available
        param_name: Name of parameter for error messages
        
    Returns:
        int: Validated area number
        
    Raises:
        RefExcelError: If area number is out of range
    """
    area_num_int = validate_positive_integer(area_num, param_name)
    
    if area_num_int > total_areas:
        raise xlerrors.RefExcelError("Area number out of range")
    
    return area_num_int


def validate_excel_bounds(row, col, param_prefix=""):
    """Validate row and column are within Excel bounds.
    
    Args:
        row: Row number to validate (1-based)
        col: Column number to validate (1-based)
        param_prefix: Prefix for parameter names in error messages
        
    Raises:
        RefExcelError: If coordinates are out of Excel bounds
    """
    if row < 1 or row > EXCEL_MAX_ROWS:
        raise xlerrors.RefExcelError(f"{param_prefix}Row {row} is out of Excel bounds (1-{EXCEL_MAX_ROWS})")
    
    if col < 1 or col > EXCEL_MAX_COLUMNS:
        raise xlerrors.RefExcelError(f"{param_prefix}Column {col} is out of Excel bounds (1-{EXCEL_MAX_COLUMNS})")


def validate_offset_parameters(rows_offset, cols_offset):
    """Validate OFFSET function row and column offset parameters.
    
    Args:
        rows_offset: Row offset value
        cols_offset: Column offset value
        
    Returns:
        tuple: (validated_rows_int, validated_cols_int)
        
    Raises:
        ValueExcelError: If parameters cannot be converted to integers
    """
    try:
        rows_int = int(rows_offset)
        cols_int = int(cols_offset)
        return rows_int, cols_int
    except (ValueError, TypeError):
        raise xlerrors.ValueExcelError("Row and column offsets must be numbers")


def validate_offset_bounds(base_row, base_col, rows_offset, cols_offset):
    """Validate OFFSET operation stays within Excel bounds.
    
    Args:
        base_row: Starting row (1-based)
        base_col: Starting column (1-based)
        rows_offset: Row offset to apply
        cols_offset: Column offset to apply
        
    Raises:
        RefExcelError: If offset results in coordinates outside Excel bounds
    """
    target_row = base_row + rows_offset
    target_col = base_col + cols_offset
    
    # Check if target is before sheet start
    if target_row < 1 or target_col < 1:
        raise xlerrors.RefExcelError("Reference before sheet start")
    
    # Check Excel bounds
    validate_excel_bounds(target_row, target_col, "Target ")


def validate_range_dimensions(height, width):
    """Validate range height and width parameters.
    
    Args:
        height: Range height (None or positive integer)
        width: Range width (None or positive integer)
        
    Returns:
        tuple: (validated_height_int, validated_width_int) with defaults of 1
        
    Raises:
        ValueExcelError: If parameters are invalid
    """
    try:
        height_int = int(height) if height is not None else 1
        width_int = int(width) if width is not None else 1
    except (ValueError, TypeError):
        raise xlerrors.ValueExcelError("Height and width must be numbers")
    
    validate_dimension_parameter(height_int, "height")
    validate_dimension_parameter(width_int, "width")
    
    return height_int, width_int