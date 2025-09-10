"""Utility modules for xlcalculator."""

# Import legacy utils functions for backward compatibility
try:
    from ..range import (
        resolve_ranges,
        resolve_sheet,
        parse_sheet_and_address,
        resolve_address,
        is_full_range,
        CellReference,
        ParsedAddress,
        RangeReference,
        MAX_COL,
        MAX_ROW
    )
except ImportError:
    # Fallback for circular import issues
    resolve_ranges = None
    resolve_sheet = None
    parse_sheet_and_address = None
    resolve_address = None
    is_full_range = None
    CellReference = None
    ParsedAddress = None
    RangeReference = None
    MAX_COL = None
    MAX_ROW = None

from .validation import (
    validate_integer_parameter,
    validate_positive_integer,
    validate_array_bounds,
    validate_dimension_parameter,
    validate_area_number
)

from .decorators import (
    require_context,
    excel_function,
    validate_parameters
)

from .arrays import ArrayProcessor

from .types import ExcelTypeConverter

from .reference_parsing import (
    parse_excel_reference,
    extract_reference_data,
    validate_reference_dimensions,
    get_reference_areas
)

__all__ = [
    # Legacy utils functions
    'resolve_ranges',
    'resolve_sheet',
    'parse_sheet_and_address',
    'resolve_address',
    'is_full_range',
    'CellReference',
    'ParsedAddress',
    'RangeReference',
    'MAX_COL',
    'MAX_ROW',
    # New utility functions
    'validate_integer_parameter',
    'validate_positive_integer', 
    'validate_array_bounds',
    'validate_dimension_parameter',
    'validate_area_number',
    'require_context',
    'excel_function',
    'validate_parameters',
    'ArrayProcessor',
    'ExcelTypeConverter',
    'parse_excel_reference',
    'extract_reference_data',
    'validate_reference_dimensions',
    'get_reference_areas'
]