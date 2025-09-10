"""Array processing utilities for Excel functions."""

from ..xlfunctions import xlerrors


class ArrayProcessor:
    """Utility class for common array operations in Excel functions."""
    
    @staticmethod
    def extract_array_data(reference, evaluator):
        """Extract array data from various reference types.
        
        Args:
            reference: Reference to extract data from (range, cell, or value)
            evaluator: Excel evaluator instance
            
        Returns:
            2D array of values
        """
        if hasattr(reference, 'value'):
            # Handle range objects with value attribute
            return evaluator.evaluate(reference)
        elif hasattr(reference, 'address'):
            # Handle range objects with address attribute
            return evaluator.evaluate(reference)
        elif isinstance(reference, (list, tuple)):
            # Handle direct array data
            return ArrayProcessor.ensure_2d_array(reference)
        else:
            # Handle single values
            return [[reference]]
    
    @staticmethod
    def ensure_2d_array(data):
        """Ensure data is a 2D array.
        
        Args:
            data: Data to convert to 2D array
            
        Returns:
            2D array representation of data
        """
        if not isinstance(data, (list, tuple)):
            return [[data]]
        
        if not data:
            return [[]]
        
        # Check if it's already 2D
        if isinstance(data[0], (list, tuple)):
            return data
        
        # Convert 1D to 2D (single row)
        return [data]
    
    @staticmethod
    def get_array_dimensions(array_data):
        """Get dimensions of 2D array.
        
        Args:
            array_data: 2D array to measure
            
        Returns:
            tuple: (rows, columns) dimensions
        """
        if not array_data:
            return 0, 0
        
        rows = len(array_data)
        cols = len(array_data[0]) if array_data[0] else 0
        return rows, cols
    
    @staticmethod
    def is_single_value(data):
        """Check if data represents a single value.
        
        Args:
            data: Data to check
            
        Returns:
            bool: True if data represents a single value
        """
        if not isinstance(data, (list, tuple)):
            return True
        
        if len(data) == 1 and isinstance(data[0], (list, tuple)) and len(data[0]) == 1:
            return True
        
        return False
    
    @staticmethod
    def get_single_value(data):
        """Extract single value from array data.
        
        Args:
            data: Array data that should contain a single value
            
        Returns:
            Single value from the array
            
        Raises:
            ValueExcelError: If data doesn't contain exactly one value
        """
        if not ArrayProcessor.is_single_value(data):
            raise xlerrors.ValueExcelError("Expected single value, got array")
        
        if not isinstance(data, (list, tuple)):
            return data
        
        return data[0][0]
    
    @staticmethod
    def flatten_array(array_data):
        """Flatten 2D array to 1D list.
        
        Args:
            array_data: 2D array to flatten
            
        Returns:
            list: Flattened array values
        """
        if not array_data:
            return []
        
        result = []
        for row in array_data:
            if isinstance(row, (list, tuple)):
                result.extend(row)
            else:
                result.append(row)
        
        return result
    
    @staticmethod
    def validate_array_not_empty(array_data, param_name="array"):
        """Validate that array is not empty.
        
        Args:
            array_data: Array to validate
            param_name: Parameter name for error messages
            
        Raises:
            ValueExcelError: If array is empty
        """
        if not array_data or (isinstance(array_data, (list, tuple)) and len(array_data) == 0):
            raise xlerrors.ValueExcelError(f"{param_name} cannot be empty")
        
        if isinstance(array_data, (list, tuple)) and len(array_data) > 0:
            if not array_data[0] or (isinstance(array_data[0], (list, tuple)) and len(array_data[0]) == 0):
                raise xlerrors.ValueExcelError(f"{param_name} cannot be empty")