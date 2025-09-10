from . import xl, xlerrors, func_xltypes


@xl.register()
@xl.validate_args
def ISBLANK(cell: func_xltypes.XlAnything) -> func_xltypes.Boolean:
    """Returns TRUE if the cell is empty.
    """
    return isinstance(cell, func_xltypes.Blank) or cell.value == ''


@xl.register()
def ISERR(value: func_xltypes.XlAnything) -> func_xltypes.Boolean:
    """Value refers to error values
    (#VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!)
    And NOT error #N/A
    """
    if isinstance(value, xlerrors.ExcelError) \
            and not isinstance(value, xlerrors.NaExcelError):
        return True
    else:
        return False


@xl.register()
def ISERROR(value: func_xltypes.XlAnything) -> func_xltypes.Boolean:
    """Value refers to any error value
    (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!).
    """
    return isinstance(value, xlerrors.ExcelError)


def _check_even_odd(num: func_xltypes.XlNumber, check_even: bool = True) -> func_xltypes.Boolean:
    """Helper function to check if a number is even or odd.
    
    Args:
        num: Number to check
        check_even: If True, returns True for even numbers. If False, returns True for odd numbers.
    """
    int_num = int(num)
    is_even = (int_num % 2) == 0
    return is_even if check_even else not is_even


@xl.register()
@xl.validate_args
def ISEVEN(num: func_xltypes.XlNumber) -> func_xltypes.Boolean:
    """Returns TRUE if number is even, or FALSE if number is odd.

    https://support.microsoft.com/en-us/office/
        iseven-function-aa15929a-d77b-4fbb-92f4-2f479af55356
    """
    return _check_even_odd(num, check_even=True)


@xl.register()
@xl.validate_args
def ISNUMBER(cell: func_xltypes.XlAnything) -> func_xltypes.Boolean:
    """Returns True if the cell is number.
    """
    return isinstance(cell, func_xltypes.Number)


@xl.register()
@xl.validate_args
def ISTEXT(cell: func_xltypes.XlAnything) -> func_xltypes.Boolean:
    """Returns True if the cell is text.
    """
    return isinstance(cell, func_xltypes.Text)


@xl.register()
@xl.validate_args
def NA() -> xlerrors.ExcelError:
    return xlerrors.NaExcelError()


@xl.register()
def ISNA(cell) -> func_xltypes.Boolean:
    """Returns True if the cell is #N/A.

    Don't call validate_args here because we allow errors to be
    passed in.

    """
    return isinstance(cell, xlerrors.NaExcelError)


@xl.register()
@xl.validate_args
def ISODD(num: func_xltypes.XlNumber) -> func_xltypes.Boolean:
    """Returns TRUE if number is odd, or FALSE if number is even.

    https://support.microsoft.com/en-us/office/
        isodd-function-1208a56d-4f10-4f44-a5fc-648cafd6c07a
    """
    return _check_even_odd(num, check_even=False)
