# DEPRECATED: This module's range functionality has been moved to xlcalculator.range
# Import from the new range module for backward compatibility

from .range import (
    # Core dataclasses
    CellReference,
    ParsedAddress, 
    RangeReference,
    
    # Utility functions
    resolve_sheet,
    parse_sheet_and_address,
    resolve_address,
    resolve_ranges,
    is_full_range,
    
    # Constants
    MAX_COL,
    MAX_ROW,
    DEFAULT_MAX_DATA_ROW,
    DEFAULT_MAX_DATA_COL,
    SMART_RANGE_ENABLED,
)

# Re-export everything for backward compatibility
__all__ = [
    'CellReference',
    'ParsedAddress',
    'RangeReference', 
    'resolve_sheet',
    'parse_sheet_and_address',
    'resolve_address',
    'resolve_ranges',
    'is_full_range',
    'MAX_COL',
    'MAX_ROW',
    'DEFAULT_MAX_DATA_ROW',
    'DEFAULT_MAX_DATA_COL',
    'SMART_RANGE_ENABLED',
]
