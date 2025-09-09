"""
Context system for Excel function execution.

Provides context injection for functions that need access to cell coordinates,
evaluator, and other execution context without relying on global variables.

ATDD Implementation: Replaces global context system with parameter injection.
"""

import inspect
from dataclasses import dataclass
from typing import Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from .xltypes import XLCell
    from .evaluator import Evaluator


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