"""Decorators for Excel function standardization."""

import functools
from ..xlfunctions import xlerrors


def require_context(func):
    """Decorator to ensure function has evaluator context.
    
    Args:
        func: Function to decorate
        
    Returns:
        Decorated function that validates context presence
        
    Raises:
        ValueExcelError: If _context parameter is None
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        context = kwargs.get('_context')
        if context is None:
            raise xlerrors.ValueExcelError(f"{func.__name__} requires evaluator context")
        return func(*args, **kwargs)
    return wrapper


def excel_function(func):
    """Decorator for standardized Excel function behavior.
    
    Provides consistent error handling and logging for Excel functions.
    
    Args:
        func: Function to decorate
        
    Returns:
        Decorated function with standardized error handling
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            # Preserve Excel error types, convert others to ValueExcelError
            if not isinstance(e, (xlerrors.ExcelError,)):
                raise xlerrors.ValueExcelError(f"Error in {func.__name__}: {str(e)}")
            raise
    return wrapper


def validate_parameters(**validators):
    """Decorator to validate function parameters using provided validators.
    
    Args:
        **validators: Mapping of parameter names to validation functions
        
    Returns:
        Decorator function
        
    Example:
        @validate_parameters(
            row_num=lambda x: validate_positive_integer(x, "row number"),
            col_num=lambda x: validate_positive_integer(x, "column number")
        )
        def my_function(row_num, col_num):
            pass
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            # Get function signature to map args to parameter names
            import inspect
            sig = inspect.signature(func)
            bound_args = sig.bind(*args, **kwargs)
            bound_args.apply_defaults()
            
            # Validate specified parameters
            for param_name, validator in validators.items():
                if param_name in bound_args.arguments:
                    value = bound_args.arguments[param_name]
                    if value is not None:  # Skip None values
                        bound_args.arguments[param_name] = validator(value)
            
            return func(*bound_args.args, **bound_args.kwargs)
        return wrapper
    return decorator