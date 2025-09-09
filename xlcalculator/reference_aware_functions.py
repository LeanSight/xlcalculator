"""
Reference-aware function registry.

This module maintains a registry of functions that need string references
instead of evaluated cell values.

ATDD Implementation: Solves the core issue where ROW("A1") returns BLANK
because AST evaluates "A1" as a cell reference instead of passing the string.
"""

from typing import Set

# Set of function names that need reference strings, not evaluated values
_REFERENCE_AWARE_FUNCTIONS: Set[str] = {
    'ROW',
    'COLUMN', 
    'OFFSET',
    'INDIRECT'
}


def is_reference_aware_function(func_name: str) -> bool:
    """
    Check if a function needs reference strings instead of evaluated values.
    
    Args:
        func_name: Name of the function (case insensitive)
        
    Returns:
        True if function needs reference strings
    """
    return func_name.upper() in _REFERENCE_AWARE_FUNCTIONS


def register_reference_aware_function(func_name: str) -> None:
    """
    Register a function as needing reference strings.
    
    Args:
        func_name: Name of the function to register
    """
    _REFERENCE_AWARE_FUNCTIONS.add(func_name.upper())


def get_reference_aware_functions() -> Set[str]:
    """Get set of all reference-aware function names."""
    return _REFERENCE_AWARE_FUNCTIONS.copy()


def is_string_reference_parameter(pitem, func_name: str, param_index: int = 0) -> bool:
    """
    Check if a parameter should be treated as a string reference.
    
    Args:
        pitem: Parameter AST node
        func_name: Name of the function being called
        param_index: Index of the parameter (0-based)
        
    Returns:
        True if parameter should be passed as string reference
    """
    # Only apply to reference-aware functions
    if not is_reference_aware_function(func_name):
        return False
    
    # Check if parameter looks like a string literal
    if (hasattr(pitem, 'tvalue') and 
        hasattr(pitem, 'ttype') and 
        pitem.ttype == 'operand' and
        isinstance(pitem.tvalue, str)):
        
        # For ROW and COLUMN, first parameter can be reference string
        if func_name.upper() in ['ROW', 'COLUMN'] and param_index == 0:
            return True
            
        # For OFFSET, first parameter can be reference string
        if func_name.upper() == 'OFFSET' and param_index == 0:
            return True
            
        # For INDIRECT, first parameter is always reference string
        if func_name.upper() == 'INDIRECT' and param_index == 0:
            return True
    
    return False