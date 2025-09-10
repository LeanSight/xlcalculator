"""
Excel Constants Module

Centralized constants for Excel compatibility and limits.
All magic numbers and Excel-specific values should be defined here.
"""

# Excel worksheet limits (Excel 2007+)
EXCEL_MAX_ROWS = 1048576
EXCEL_MAX_COLUMNS = 16384
EXCEL_MAX_COLUMN_INDEX = 18278  # XFD column

# Excel cell content limits
EXCEL_CELL_CHARACTER_LIMIT = 32767

# Excel function limits
EXCEL_CONCAT_MAX_STRINGS = 254

# Excel error values (for reference)
EXCEL_ERROR_DIV_ZERO = "#DIV/0!"
EXCEL_ERROR_NA = "#N/A"
EXCEL_ERROR_NAME = "#NAME?"
EXCEL_ERROR_NULL = "#NULL!"
EXCEL_ERROR_NUM = "#NUM!"
EXCEL_ERROR_REF = "#REF!"
EXCEL_ERROR_VALUE = "#VALUE!"