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
# DYNAMIC RANGE FUNCTIONS - Implemented via ATDD strict methodology
# ============================================================================

# Functions will be implemented incrementally following ATDD cycles:
# 1. RED: Test fails
# 2. GREEN: Minimal implementation to pass test
# 3. REFACTOR: Eliminate duplication
# 4. COMMIT: Save progress

# INDEX, OFFSET, INDIRECT functions will be added here as tests demand them