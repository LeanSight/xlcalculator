from . import xl, xlerrors, func_xltypes


@xl.register()
@xl.validate_args
def OP_MUL(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlNumber:
    return left * right


@xl.register()
@xl.validate_args
def OP_DIV(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlNumber:
    if right == 0:
        raise xlerrors.DivZeroExcelError()
    return left / right


@xl.register()
@xl.validate_args
def OP_ADD(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlAnything:
    # Handle Array arithmetic for dynamic arrays
    if isinstance(left, func_xltypes.Array) and not isinstance(right, func_xltypes.Array):
        # Array + scalar: add scalar to each element
        result_values = []
        for row in left.values:
            result_row = []
            for cell in row:
                try:
                    result_row.append(func_xltypes.Number(float(cell) + float(right)))
                except (ValueError, TypeError):
                    result_row.append(xlerrors.ValueExcelError("Cannot add"))
            result_values.append(result_row)
        return func_xltypes.Array(result_values)
    
    elif isinstance(right, func_xltypes.Array) and not isinstance(left, func_xltypes.Array):
        # Scalar + Array: add each element to scalar
        result_values = []
        for row in right.values:
            result_row = []
            for cell in row:
                try:
                    result_row.append(func_xltypes.Number(float(left) + float(cell)))
                except (ValueError, TypeError):
                    result_row.append(xlerrors.ValueExcelError("Cannot add"))
            result_values.append(result_row)
        return func_xltypes.Array(result_values)
    
    elif isinstance(left, func_xltypes.Array) and isinstance(right, func_xltypes.Array):
        # Array + Array: element-wise addition
        if len(left.values) != len(right.values):
            raise xlerrors.ValueExcelError("Array dimensions must match")
        
        result_values = []
        for i, (left_row, right_row) in enumerate(zip(left.values, right.values)):
            if len(left_row) != len(right_row):
                raise xlerrors.ValueExcelError("Array dimensions must match")
            
            result_row = []
            for left_cell, right_cell in zip(left_row, right_row):
                try:
                    result_row.append(func_xltypes.Number(float(left_cell) + float(right_cell)))
                except (ValueError, TypeError):
                    result_row.append(xlerrors.ValueExcelError("Cannot add"))
            result_values.append(result_row)
        return func_xltypes.Array(result_values)
    
    else:
        # Regular scalar arithmetic
        return left + right


@xl.register()
@xl.validate_args
def OP_SUB(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlAnything:
    # Handle Array arithmetic for dynamic arrays
    if isinstance(left, func_xltypes.Array) and not isinstance(right, func_xltypes.Array):
        # Array - scalar: subtract scalar from each element
        result_values = []
        for row in left.values:
            result_row = []
            for cell in row:
                try:
                    result_row.append(func_xltypes.Number(float(cell) - float(right)))
                except (ValueError, TypeError):
                    result_row.append(xlerrors.ValueExcelError("Cannot subtract"))
            result_values.append(result_row)
        return func_xltypes.Array(result_values)
    
    elif isinstance(right, func_xltypes.Array) and not isinstance(left, func_xltypes.Array):
        # Scalar - Array: subtract each element from scalar
        result_values = []
        for row in right.values:
            result_row = []
            for cell in row:
                try:
                    result_row.append(func_xltypes.Number(float(left) - float(cell)))
                except (ValueError, TypeError):
                    result_row.append(xlerrors.ValueExcelError("Cannot subtract"))
            result_values.append(result_row)
        return func_xltypes.Array(result_values)
    
    elif isinstance(left, func_xltypes.Array) and isinstance(right, func_xltypes.Array):
        # Array - Array: element-wise subtraction
        if len(left.values) != len(right.values):
            raise xlerrors.ValueExcelError("Array dimensions must match")
        
        result_values = []
        for i, (left_row, right_row) in enumerate(zip(left.values, right.values)):
            if len(left_row) != len(right_row):
                raise xlerrors.ValueExcelError("Array dimensions must match")
            
            result_row = []
            for left_cell, right_cell in zip(left_row, right_row):
                try:
                    result_row.append(func_xltypes.Number(float(left_cell) - float(right_cell)))
                except (ValueError, TypeError):
                    result_row.append(xlerrors.ValueExcelError("Cannot subtract"))
            result_values.append(result_row)
        return func_xltypes.Array(result_values)
    
    else:
        # Regular scalar arithmetic
        return left - right


@xl.register()
def OP_EQ(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlBoolean:
    return left == right


@xl.register()
def OP_NE(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlBoolean:
    return left != right


@xl.register()
@xl.validate_args
def OP_GT(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlBoolean:
    if isinstance(left, func_xltypes.Blank) or \
            isinstance(right, func_xltypes.Blank):
        return False
    return left > right


@xl.register()
@xl.validate_args
def OP_LT(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlBoolean:
    if isinstance(left, func_xltypes.Blank) or \
            isinstance(right, func_xltypes.Blank):
        return False
    return left < right


@xl.register()
@xl.validate_args
def OP_GE(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlBoolean:
    if isinstance(left, func_xltypes.Blank) or \
            isinstance(right, func_xltypes.Blank):
        return False
    return left >= right


@xl.register()
@xl.validate_args
def OP_LE(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlBoolean:
    if isinstance(left, func_xltypes.Blank) or \
            isinstance(right, func_xltypes.Blank):
        return False
    return left <= right


@xl.register()
@xl.validate_args
def OP_NEG(
        right: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    return -1 * right


@xl.register()
@xl.validate_args
def OP_PERCENT(
        left: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    return left * 0.01


@xl.register()
def OP_UNION(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlAnything:
    """Handle Excel union operator (comma).
    
    Returns a tuple containing both operands for functions like INDEX
    that need to handle multiple areas.
    """
    # Return a tuple that can be handled by functions like INDEX
    return (left, right)
