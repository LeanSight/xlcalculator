"""
Cell Class Implementation

Represents an Excel cell with proper address parsing,
formula handling, and value management.
"""
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Optional
from openpyxl.utils.cell import column_index_from_string

if TYPE_CHECKING:
    from .worksheet import Worksheet
    from ..xltypes import XLFormula


@dataclass
class Cell:
    """Excel Cell with hierarchical structure and proper address handling."""
    
    address: str  # Local address like "A1"
    worksheet: 'Worksheet'
    value: Any = None
    formula: Optional['XLFormula'] = None
    
    # Computed properties (cached)
    _row: Optional[int] = field(init=False, default=None, repr=False)
    _column: Optional[str] = field(init=False, default=None, repr=False)
    _column_index: Optional[int] = field(init=False, default=None, repr=False)
    
    def __post_init__(self):
        """Initialize computed properties after creation."""
        self._parse_address()
    
    def _parse_address(self) -> None:
        """Parse the cell address to extract row and column information."""
        try:
            from ..range import ParsedAddress
            # Parse using full address for consistency
            full_address = f"{self.worksheet.name}!{self.address}"
            parsed = ParsedAddress.parse(full_address)
            
            self._row = parsed.row
            self._column = parsed.column
            self._column_index = column_index_from_string(parsed.column)
        except Exception as e:
            raise ValueError(f"Invalid cell address '{self.address}': {e}")
    
    @property
    def row(self) -> int:
        """Get the row number (1-based)."""
        if self._row is None:
            self._parse_address()
        return self._row
    
    @property
    def column(self) -> str:
        """Get the column letter(s)."""
        if self._column is None:
            self._parse_address()
        return self._column
    
    @property
    def column_index(self) -> int:
        """Get the column index (1-based)."""
        if self._column_index is None:
            self._parse_address()
        return self._column_index
    
    @property
    def full_address(self) -> str:
        """Get the full address including sheet name."""
        return f"{self.worksheet.name}!{self.address}"
    
    @property
    def has_formula(self) -> bool:
        """Check if cell contains a formula."""
        return self.formula is not None
    
    @property
    def has_value(self) -> bool:
        """Check if cell contains a value."""
        return self.value is not None
    
    @property
    def is_empty(self) -> bool:
        """Check if cell is empty (no value and no formula)."""
        return self.value is None and self.formula is None
    
    def set_value(self, value: Any) -> None:
        """Set cell value, clearing any existing formula.
        
        Args:
            value: Value to set in the cell
        """
        self.value = value
        self.formula = None
    
    def set_formula(self, formula: 'XLFormula') -> None:
        """Set cell formula, clearing any existing value.
        
        Args:
            formula: XLFormula instance to set
        """
        self.formula = formula
        self.value = None  # Formula cells don't have direct values
    
    def clear(self) -> None:
        """Clear both value and formula from the cell."""
        self.value = None
        self.formula = None
    
    def copy_to(self, target_cell: 'Cell') -> None:
        """Copy this cell's content to another cell.
        
        Args:
            target_cell: Target cell to copy to
        """
        target_cell.value = self.value
        if self.formula:
            # TODO: Adjust formula references for new location
            target_cell.formula = self.formula
    
    def get_display_value(self) -> str:
        """Get the display value of the cell.
        
        Returns:
            String representation of the cell's content
        """
        if self.has_formula:
            return self.formula.formula
        elif self.has_value:
            return str(self.value)
        else:
            return ""
    
    def get_data_type(self) -> str:
        """Get the data type of the cell's value.
        
        Returns:
            String representing the data type
        """
        if self.has_formula:
            return "formula"
        elif self.value is None:
            return "empty"
        elif isinstance(self.value, bool):
            return "boolean"
        elif isinstance(self.value, int):
            return "integer"
        elif isinstance(self.value, float):
            return "float"
        elif isinstance(self.value, str):
            return "string"
        else:
            return "unknown"
    
    def is_numeric(self) -> bool:
        """Check if cell contains a numeric value.
        
        Returns:
            True if cell contains int or float, False otherwise
        """
        return isinstance(self.value, (int, float)) and not isinstance(self.value, bool)
    
    def is_text(self) -> bool:
        """Check if cell contains text.
        
        Returns:
            True if cell contains string, False otherwise
        """
        return isinstance(self.value, str)
    
    def is_boolean(self) -> bool:
        """Check if cell contains a boolean value.
        
        Returns:
            True if cell contains boolean, False otherwise
        """
        return isinstance(self.value, bool)
    
    def to_number(self) -> Optional[float]:
        """Convert cell value to number if possible.
        
        Returns:
            Float value if conversion is possible, None otherwise
        """
        if self.is_numeric():
            return float(self.value)
        elif self.is_text():
            try:
                return float(self.value)
            except (ValueError, TypeError):
                return None
        else:
            return None
    
    def to_string(self) -> str:
        """Convert cell value to string.
        
        Returns:
            String representation of the cell value
        """
        if self.value is None:
            return ""
        elif isinstance(self.value, bool):
            return "TRUE" if self.value else "FALSE"
        else:
            return str(self.value)
    
    def validate_address_format(self) -> bool:
        """Validate that the cell address is in correct format.
        
        Returns:
            True if address is valid, False otherwise
        """
        try:
            self._parse_address()
            return True
        except Exception:
            return False
    
    def get_relative_address(self, row_offset: int = 0, col_offset: int = 0) -> str:
        """Get address relative to this cell.
        
        Args:
            row_offset: Row offset (positive = down, negative = up)
            col_offset: Column offset (positive = right, negative = left)
            
        Returns:
            Address of the relative cell
            
        Raises:
            ValueError: If resulting address is invalid
        """
        new_row = self.row + row_offset
        new_col_index = self.column_index + col_offset
        
        if new_row < 1:
            raise ValueError(f"Row offset {row_offset} results in invalid row {new_row}")
        if new_col_index < 1:
            raise ValueError(f"Column offset {col_offset} results in invalid column {new_col_index}")
        
        from openpyxl.utils.cell import get_column_letter
        new_col_letter = get_column_letter(new_col_index)
        
        return f"{new_col_letter}{new_row}"
    
    def is_adjacent_to(self, other_cell: 'Cell') -> bool:
        """Check if this cell is adjacent to another cell.
        
        Args:
            other_cell: Cell to check adjacency with
            
        Returns:
            True if cells are adjacent (horizontally or vertically)
        """
        if self.worksheet != other_cell.worksheet:
            return False
        
        row_diff = abs(self.row - other_cell.row)
        col_diff = abs(self.column_index - other_cell.column_index)
        
        # Adjacent means exactly one cell away in one direction
        return (row_diff == 1 and col_diff == 0) or (row_diff == 0 and col_diff == 1)
    
    def __float__(self) -> float:
        """Convert cell to float for numeric operations."""
        if self.is_numeric():
            return float(self.value)
        else:
            raise TypeError(f"Cannot convert cell value '{self.value}' to float")
    
    def __str__(self) -> str:
        """String representation of cell value."""
        return self.to_string()
    
    def __eq__(self, other) -> bool:
        """Compare cells for equality."""
        if not isinstance(other, Cell):
            return False
        
        return (
            self.address == other.address
            and self.worksheet == other.worksheet
            and self.value == other.value
            and self.formula == other.formula
        )
    
    def __hash__(self) -> int:
        """Hash for cell (based on worksheet and address)."""
        return hash((self.worksheet.name, self.address))
    
    def __repr__(self) -> str:
        """Detailed string representation of cell."""
        content = ""
        if self.has_formula:
            content = f"formula='{self.formula.formula}'"
        elif self.has_value:
            content = f"value={repr(self.value)}"
        else:
            content = "empty"
        
        return f"Cell('{self.full_address}', {content})"