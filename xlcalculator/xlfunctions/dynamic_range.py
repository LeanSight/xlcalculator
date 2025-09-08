"""
Dynamic Range Functions: INDEX, OFFSET, INDIRECT

Implementation following ATDD strict methodology.
Each function implemented incrementally based on failing acceptance tests.

Functions:
- INDEX: Returns value at intersection of row/column in array
- OFFSET: Returns reference offset from starting reference  
- INDIRECT: Returns reference specified by text string

Architecture:
- Context injection system for evaluator access
- Dynamic reference parsing and evaluation
- Excel-compatible error handling
"""

from . import xl, xlerrors, func_xltypes


# ============================================================================
# CONTEXT INJECTION SYSTEM - Access to evaluator during function execution
# ============================================================================

# Global evaluator context - set by evaluator before function calls
_EVALUATOR_CONTEXT = None


def _set_evaluator_context(evaluator):
    """Set evaluator context for dynamic range functions.
    
    Called by evaluator before executing dynamic range functions.
    Provides access to model, cells, and evaluation capabilities.
    """
    global _EVALUATOR_CONTEXT
    _EVALUATOR_CONTEXT = evaluator


def _get_evaluator_context():
    """Get current evaluator context.
    
    Returns evaluator instance for accessing model and evaluation.
    Raises RuntimeError if no context available.
    """
    if _EVALUATOR_CONTEXT is None:
        raise RuntimeError("No evaluator context available for dynamic range function")
    return _EVALUATOR_CONTEXT


# ============================================================================
# SHARED UTILITIES - Eliminate duplicate logic across dynamic range functions
# ============================================================================

def _convert_to_python_int(xl_number):
    """Convert XL Number to Python int, eliminating duplication.
    
    Used by: INDEX, OFFSET (future)
    Returns: Python integer
    """
    return int(xl_number)


def _resolve_array_data(array, evaluator):
    """Resolve array parameter to Python list data structure.
    
    Used by: INDEX, OFFSET (future)
    Returns: 2D list of values
    """
    if hasattr(array, 'values'):
        # It's a pandas DataFrame from xlcalculator
        return array.values.tolist()
    else:
        # It's a string reference, use get_range_values
        return evaluator.get_range_values(str(array))


# ============================================================================
# DYNAMIC RANGE FUNCTIONS - Implemented via ATDD strict methodology
# ============================================================================

# Functions will be implemented incrementally following ATDD cycles:
# 1. RED: Test fails
# 2. GREEN: Minimal implementation to pass test
# 3. REFACTOR: Eliminate duplication
# 4. COMMIT: Save progress

@xl.register()
@xl.validate_args
def INDEX(
    array: func_xltypes.XlAnything,
    row_num: func_xltypes.XlNumber,
    col_num: func_xltypes.XlNumber = 1
) -> func_xltypes.XlAnything:
    """Returns value at intersection of row/column in array.
    
    CICLO 2.1: INDEX(Data!A1:E6, 2, 2) = 25
    Refactored to use shared utilities.
    """
    evaluator = _get_evaluator_context()
    
    # Convert parameters using shared utility
    row_idx = _convert_to_python_int(row_num) - 1  # Convert to 0-based index
    col_idx = _convert_to_python_int(col_num) - 1  # Convert to 0-based index
    
    # Resolve array data using shared utility
    array_data = _resolve_array_data(array, evaluator)
    
    # Return value at specified position
    return array_data[row_idx][col_idx]