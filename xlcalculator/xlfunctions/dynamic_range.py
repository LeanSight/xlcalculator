"""
Dynamic Range Functions - ATDD Placeholder

This module will implement Excel's dynamic range functions (INDEX, OFFSET, INDIRECT)
following strict ATDD (Acceptance Test-Driven Development).

Implementation will be driven by acceptance tests from DYNAMIC_RANGES_COMPREHENSIVE.xlsx

Currently: No implementation - waiting for acceptance tests to drive development.
"""

from . import xl, xlerrors, func_xltypes


# Minimal context functions required by evaluator (legacy compatibility)
def _set_evaluator_context(evaluator):
    """Set evaluator context - placeholder for system compatibility."""
    pass

def _get_evaluator_context():
    """Get evaluator context - placeholder for system compatibility."""
    return None

def _clear_evaluator_context():
    """Clear evaluator context - placeholder for system compatibility."""
    pass


# Placeholder functions - no implementation until acceptance tests drive development
@xl.register()
def INDEX(array, row_num, col_num=1):
    """INDEX function placeholder - no implementation until acceptance test fails."""
    return xlerrors.ValueExcelError("INDEX: Not implemented - awaiting acceptance test")


@xl.register()
def OFFSET(reference, rows, cols, height=None, width=None):
    """OFFSET function placeholder - no implementation until acceptance test fails."""
    return xlerrors.ValueExcelError("OFFSET: Not implemented - awaiting acceptance test")


@xl.register()
def INDIRECT(ref_text, a1=True):
    """INDIRECT function placeholder - no implementation until acceptance test fails."""
    return xlerrors.ValueExcelError("INDIRECT: Not implemented - awaiting acceptance test")