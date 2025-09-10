"""Type conversion utilities for Excel functions."""

from ..xlfunctions import xlerrors


class ExcelTypeConverter:
    """Handles Excel-specific type conversions."""
    
    @staticmethod
    def to_number(value, param_name="value"):
        """Convert value to number with Excel semantics.
        
        Args:
            value: Value to convert
            param_name: Parameter name for error messages
            
        Returns:
            float or int: Numeric value
            
        Raises:
            ValueExcelError: If value cannot be converted to number
        """
        if isinstance(value, (int, float)):
            return value
        
        if isinstance(value, bool):
            return 1 if value else 0
        
        if isinstance(value, str):
            # Handle empty string
            if not value.strip():
                return 0
            
            try:
                # Try integer first to preserve type
                if '.' not in value and 'e' not in value.lower():
                    return int(value)
                return float(value)
            except ValueError:
                raise xlerrors.ValueExcelError(f"Cannot convert {param_name} to number: {value}")
        
        raise xlerrors.ValueExcelError(f"Invalid {param_name} type: {type(value).__name__}")
    
    @staticmethod
    def to_integer(value, param_name="value"):
        """Convert value to integer with Excel semantics.
        
        Args:
            value: Value to convert
            param_name: Parameter name for error messages
            
        Returns:
            int: Integer value
            
        Raises:
            ValueExcelError: If value cannot be converted to integer
        """
        number = ExcelTypeConverter.to_number(value, param_name)
        
        try:
            return int(number)
        except (ValueError, OverflowError):
            raise xlerrors.ValueExcelError(f"Cannot convert {param_name} to integer: {value}")
    
    @staticmethod
    def to_boolean(value):
        """Convert value to boolean with Excel semantics.
        
        Args:
            value: Value to convert
            
        Returns:
            bool: Boolean value
        """
        if isinstance(value, bool):
            return value
        
        if isinstance(value, (int, float)):
            return value != 0
        
        if isinstance(value, str):
            upper_value = value.upper().strip()
            if upper_value in ('TRUE', '1'):
                return True
            elif upper_value in ('FALSE', '0', ''):
                return False
            else:
                # Try to convert to number first
                try:
                    return ExcelTypeConverter.to_number(value) != 0
                except xlerrors.ValueExcelError:
                    return bool(value)  # Non-empty string is True
        
        return bool(value)
    
    @staticmethod
    def to_string(value):
        """Convert value to string with Excel semantics.
        
        Args:
            value: Value to convert
            
        Returns:
            str: String representation
        """
        if isinstance(value, str):
            return value
        
        if isinstance(value, bool):
            return "TRUE" if value else "FALSE"
        
        if isinstance(value, (int, float)):
            return str(value)
        
        if value is None:
            return ""
        
        return str(value)
    
    @staticmethod
    def coerce_to_common_type(values):
        """Coerce list of values to common type for comparison.
        
        Args:
            values: List of values to coerce
            
        Returns:
            list: Values coerced to common type
        """
        if not values:
            return values
        
        # Check if all values are already the same type
        first_type = type(values[0])
        if all(isinstance(v, first_type) for v in values):
            return values
        
        # Try to convert all to numbers
        try:
            return [ExcelTypeConverter.to_number(v) for v in values]
        except xlerrors.ValueExcelError:
            pass
        
        # Fall back to strings
        return [ExcelTypeConverter.to_string(v) for v in values]