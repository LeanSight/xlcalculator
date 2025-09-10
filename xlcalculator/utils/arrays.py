"""Array processing utilities for Excel functions."""

from typing import Any, List, Union, Tuple
from ..xlfunctions import xlerrors


class ArrayProcessor:
    """Utility class for common array operations in Excel functions."""
    
    @staticmethod
    def extract_array_data(reference: Any, evaluator: Any) -> List[List[Any]]:
        """Extract array data from various reference types.
        
        This method provides a unified interface for extracting 2D array data
        from different Excel reference types, ensuring consistent data format
        for all Excel functions.
        
        Args:
            reference: Reference to extract data from. Can be:
                      - pandas DataFrame (has .values attribute)
                      - Range object (has .address attribute)  
                      - String range reference (e.g., 'Data!A1:E6')
                      - Direct array data (list/tuple)
                      - Single scalar value
            evaluator: Excel evaluator instance for range resolution
            
        Returns:
            list: 2D array of values (list of lists)
            
        Examples:
            >>> # String range reference
            >>> extract_array_data('Data!A1:C2', evaluator)
            [[1, 2, 3], [4, 5, 6]]
            
            >>> # pandas DataFrame
            >>> extract_array_data(dataframe, evaluator)
            [['Alice', 25], ['Bob', 30]]
            
            >>> # Direct array
            >>> extract_array_data([1, 2, 3], evaluator)
            [[1, 2, 3]]
            
            >>> # Single value
            >>> extract_array_data(42, evaluator)
            [[42]]
        """
        # Priority order is important: most specific to most general
        
        if ArrayProcessor._is_pandas_dataframe(reference):
            return ArrayProcessor._extract_from_dataframe(reference)
            
        elif ArrayProcessor._is_range_object(reference):
            return ArrayProcessor._extract_from_range_object(reference, evaluator)
            
        elif isinstance(reference, str):
            return ArrayProcessor._extract_from_string_reference(reference, evaluator)
            
        elif isinstance(reference, (list, tuple)):
            return ArrayProcessor.ensure_2d_array(reference)
            
        else:
            # Single scalar value: wrap as 2D array
            return [[reference]]
    
    @staticmethod
    def _is_pandas_dataframe(reference: Any) -> bool:
        """Check if reference is a pandas DataFrame."""
        return hasattr(reference, 'values')
    
    @staticmethod
    def _is_range_object(reference: Any) -> bool:
        """Check if reference is a range object."""
        return hasattr(reference, 'address')
    
    @staticmethod
    def _extract_from_dataframe(dataframe: Any) -> List[List[Any]]:
        """Extract 2D array from pandas DataFrame."""
        return dataframe.values.tolist()
    
    @staticmethod
    def _extract_from_range_object(range_obj: Any, evaluator: Any) -> List[List[Any]]:
        """Extract 2D array from range object."""
        return evaluator.evaluate(range_obj)
    
    @staticmethod
    def _extract_from_string_reference(reference: str, evaluator: Any) -> List[List[Any]]:
        """Extract 2D array from string range reference."""
        return evaluator.get_range_values(reference)
    
    @staticmethod
    def ensure_2d_array(data: Any) -> List[List[Any]]:
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
    def get_array_dimensions(array_data: List[List[Any]]) -> Tuple[int, int]:
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
    def is_single_value(data: Any) -> bool:
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
    def get_single_value(data: Any) -> Any:
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
    def flatten_array(array_data: List[List[Any]]) -> List[Any]:
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
    def validate_array_not_empty(array_data: List[List[Any]], param_name: str = "array") -> None:
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