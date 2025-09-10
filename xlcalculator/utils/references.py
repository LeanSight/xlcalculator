"""Reference processing utilities for Excel functions."""

from ..xlfunctions import xlerrors
from .arrays import ArrayProcessor


def parse_excel_reference(reference, context, allow_single_value=True):
    """Parse and validate Excel reference with context.
    
    Args:
        reference: Reference to parse (string, range object, or value)
        context: Evaluator context
        allow_single_value: Whether to allow single values
        
    Returns:
        Evaluated reference data
        
    Raises:
        RefExcelError: If reference is invalid
        ValueExcelError: If single values not allowed but provided
    """
    evaluator = context.evaluator
    
    if isinstance(reference, str):
        # Parse reference string
        try:
            # Let evaluator handle the reference parsing
            return evaluator.evaluate(reference)
        except Exception as e:
            raise xlerrors.RefExcelError(f"Invalid reference: {reference}")
    
    elif hasattr(reference, 'address') or hasattr(reference, 'value'):
        # Handle range objects
        try:
            return evaluator.evaluate(reference)
        except Exception as e:
            raise xlerrors.RefExcelError(f"Cannot evaluate reference: {reference}")
    
    elif allow_single_value:
        # Handle single values
        return ArrayProcessor.ensure_2d_array(reference)
    
    else:
        raise xlerrors.RefExcelError("Reference must be a range or cell address")


def extract_reference_data(reference, context):
    """Extract array data from reference using context.
    
    Args:
        reference: Reference to extract data from
        context: Evaluator context
        
    Returns:
        2D array of reference data
    """
    evaluator = context.evaluator
    return ArrayProcessor.extract_array_data(reference, evaluator)


def validate_reference_dimensions(reference_data, min_rows=None, min_cols=None, 
                                max_rows=None, max_cols=None, param_name="reference"):
    """Validate reference dimensions.
    
    Args:
        reference_data: 2D array reference data
        min_rows: Minimum required rows
        min_cols: Minimum required columns
        max_rows: Maximum allowed rows
        max_cols: Maximum allowed columns
        param_name: Parameter name for error messages
        
    Raises:
        ValueExcelError: If dimensions don't meet requirements
    """
    rows, cols = ArrayProcessor.get_array_dimensions(reference_data)
    
    if min_rows is not None and rows < min_rows:
        raise xlerrors.ValueExcelError(f"{param_name} must have at least {min_rows} rows")
    
    if min_cols is not None and cols < min_cols:
        raise xlerrors.ValueExcelError(f"{param_name} must have at least {min_cols} columns")
    
    if max_rows is not None and rows > max_rows:
        raise xlerrors.ValueExcelError(f"{param_name} cannot have more than {max_rows} rows")
    
    if max_cols is not None and cols > max_cols:
        raise xlerrors.ValueExcelError(f"{param_name} cannot have more than {max_cols} columns")


def get_reference_areas(reference, context):
    """Get areas from a reference (for multi-area references).
    
    Args:
        reference: Reference that may contain multiple areas
        context: Evaluator context
        
    Returns:
        list: List of area data arrays
    """
    evaluator = context.evaluator
    
    # Check if reference has multiple areas
    if hasattr(reference, 'areas') and reference.areas:
        return [evaluator.evaluate(area) for area in reference.areas]
    else:
        # Single area reference
        return [extract_reference_data(reference, context)]