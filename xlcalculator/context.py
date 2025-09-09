"""
Context system for Excel function execution.

Provides context injection for functions that need access to cell coordinates,
evaluator, and other execution context without relying on global variables.

ATDD Implementation: Replaces global context system with parameter injection.
"""

import inspect
from dataclasses import dataclass
from typing import Optional, TYPE_CHECKING, Set
from functools import lru_cache

if TYPE_CHECKING:
    from .xltypes import XLCell
    from .evaluator import Evaluator


# ============================================================================
# PERFORMANCE OPTIMIZATION - Fast context function lookup
# ============================================================================

# Set of function names that require context injection for fast lookup
_CONTEXT_REQUIRED_FUNCTIONS: Set[str] = set()


@dataclass
class CellContext:
    """
    Context for function execution with direct cell access.
    
    Provides functions with access to:
    - Current cell being evaluated (with coordinates)
    - Evaluator instance for additional operations
    - Thread-safe execution context
    
    Replaces global variables with parameter injection pattern.
    """
    
    cell: 'XLCell'                    # Current cell being evaluated
    evaluator: 'Evaluator'            # Evaluator instance
    
    @property
    def row(self) -> int:
        """Get row number of current cell (1-based)."""
        return self.cell.row_index
    
    @property  
    def column(self) -> int:
        """Get column number of current cell (1-based)."""
        return self.cell.column_index
    
    @property
    def address(self) -> str:
        """Get full address of current cell (e.g., 'Sheet1!A1')."""
        return self.cell.address
    
    @property
    def sheet(self) -> str:
        """Get sheet name of current cell."""
        return self.cell.sheet
    
    def get_cell_value(self, address: str):
        """Get value of any cell through evaluator."""
        return self.evaluator.get_cell_value(address)
    
    def evaluate(self, address: str):
        """Evaluate any cell through evaluator."""
        return self.evaluator.evaluate(address)


def create_context(cell: 'XLCell', evaluator: 'Evaluator') -> CellContext:
    """
    Factory function to create CellContext.
    
    Args:
        cell: XLCell object being evaluated
        evaluator: Evaluator instance
        
    Returns:
        CellContext with cell and evaluator access
    """
    return CellContext(cell=cell, evaluator=evaluator)


# ============================================================================
# PERFORMANCE OPTIMIZATION - Context creation caching
# ============================================================================

# Cache for context objects to avoid repeated creation
_CONTEXT_CACHE = {}

def create_context_cached(cell: 'XLCell', evaluator: 'Evaluator') -> CellContext:
    """
    Optimized context creation with caching.
    
    Caches context objects by cell address to avoid repeated creation
    for the same cell during evaluation cycles.
    
    Args:
        cell: XLCell object being evaluated
        evaluator: Evaluator instance
        
    Returns:
        CellContext with cell and evaluator access (cached if possible)
    """
    # Use cell address as cache key
    cache_key = (cell.address, id(evaluator))
    
    if cache_key not in _CONTEXT_CACHE:
        _CONTEXT_CACHE[cache_key] = CellContext(cell=cell, evaluator=evaluator)
    
    return _CONTEXT_CACHE[cache_key]


def clear_context_cache():
    """Clear the context cache. Call this when evaluation is complete."""
    global _CONTEXT_CACHE
    _CONTEXT_CACHE.clear()


@lru_cache(maxsize=256)
def needs_context(func) -> bool:
    """
    Check if a function needs context injection.
    
    Functions that have a parameter named '_context' with CellContext annotation
    will receive context injection during evaluation.
    
    Args:
        func: Function to check
        
    Returns:
        True if function needs context injection
    """
    sig = inspect.signature(func)
    return '_context' in sig.parameters


def inject_context(func, args: list, context: CellContext) -> list:
    """
    Inject context into function arguments if needed.
    
    Args:
        func: Function being called
        args: Current function arguments
        context: CellContext to inject
        
    Returns:
        Modified arguments list with context injected
    """
    if not needs_context(func):
        return args
    
    # Add context as keyword argument
    sig = inspect.signature(func)
    bound = sig.bind_partial(*args)
    bound.arguments['_context'] = context
    
    # Return positional args + keyword args as expected by function
    return list(bound.args) + [context]


def needs_context_by_name(func_name: str) -> bool:
    """
    Fast lookup to check if function needs context by name.
    
    Args:
        func_name: Name of the function to check
        
    Returns:
        True if function needs context injection
    """
    return func_name in _CONTEXT_REQUIRED_FUNCTIONS


def register_context_function(func_name: str) -> None:
    """
    Register a function as requiring context injection.
    
    Args:
        func_name: Name of the function that needs context
    """
    _CONTEXT_REQUIRED_FUNCTIONS.add(func_name)


def context_aware(func):
    """
    Decorator to automatically register a function for context injection.
    
    Usage:
        @xl.register()
        @context_aware
        def MY_FUNCTION(arg1, arg2, *, _context=None):
            # Function implementation with context access
            pass
    
    Args:
        func: Function to register for context injection
        
    Returns:
        The original function (decorator doesn't modify behavior)
    """
    register_context_function(func.__name__)
    return func


def get_registered_context_functions() -> set:
    """
    Get the set of all functions registered for context injection.
    
    Returns:
        Set of function names that require context injection
    """
    return _CONTEXT_REQUIRED_FUNCTIONS.copy()