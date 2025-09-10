"""Utility modules for xlcalculator."""

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

from .references import (
    parse_excel_reference,
    extract_reference_data,
    validate_reference_dimensions,
    get_reference_areas
)

__all__ = [
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