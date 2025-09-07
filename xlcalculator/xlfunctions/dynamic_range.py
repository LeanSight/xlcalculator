"""
Dynamic Range Functions - ATDD Implementation

This module implements Excel's dynamic range functions (INDEX, OFFSET, INDIRECT)
following strict ATDD (Acceptance Test-Driven Development).

Implementation will be driven by acceptance tests from DYNAMIC_RANGES_COMPREHENSIVE.xlsx
"""

from . import xl, xlerrors, func_xltypes


# Minimal context functions required by evaluator
def _set_evaluator_context(evaluator):
    """Set evaluator context - minimal implementation for ATDD."""
    pass

def _get_evaluator_context():
    """Get evaluator context - minimal implementation for ATDD."""
    return None

def _clear_evaluator_context():
    """Clear evaluator context - minimal implementation for ATDD."""
    pass


@xl.register()
def INDEX(array, row_num, col_num=1):
    """
    Returns the value of an element in a table or array, selected by row and column.
    
    ATDD: Minimal implementation for specific test case INDEX(Data!A1:E6, 2, 2) -> 25
    """
    # ATDD: Implementación mínima para pasar el caso específico
    # Solo maneja el caso exacto: array con datos, row=2, col=2
    
    # Extraer datos del array (func_xltypes.Array)
    if isinstance(array, func_xltypes.Array) and hasattr(array, 'values'):
        # Convertir tipos Excel a valores nativos
        array_data = []
        for row in array.values:
            native_row = []
            for cell in row:
                if hasattr(cell, 'value'):
                    native_row.append(cell.value)
                else:
                    native_row.append(cell)
            array_data.append(native_row)
        
        # Acceso directo para el caso específico
        row_idx = int(row_num) - 1  # Convertir a 0-based
        col_idx = int(col_num) - 1  # Convertir a 0-based
        
        return array_data[row_idx][col_idx]
    
    # Fallback para otros casos
    return xlerrors.ValueExcelError("INDEX: Only specific test case implemented")


@xl.register()
@xl.validate_args
def OFFSET(
    reference,
    rows: func_xltypes.XlNumber,
    cols: func_xltypes.XlNumber,
    height=None,
    width=None
) -> func_xltypes.XlAnything:
    """
    Returns a reference to a range that is offset from a starting reference.
    
    ATDD: Implementation driven by acceptance tests.
    """
    # ATDD: Esta función debe fallar hasta que tengamos un test de aceptación específico
    return xlerrors.ValueExcelError("OFFSET: Not implemented - waiting for acceptance test")


@xl.register()
@xl.validate_args
def INDIRECT(
    ref_text: func_xltypes.XlText,
    a1: func_xltypes.XlBoolean = True
) -> func_xltypes.XlAnything:
    """
    Returns the reference specified by a text string.
    
    ATDD: Implementation driven by acceptance tests.
    """
    # ATDD: Esta función debe fallar hasta que tengamos un test de aceptación específico
    return xlerrors.ValueExcelError("INDIRECT: Not implemented - waiting for acceptance test")