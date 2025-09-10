"""
Unified Reference Objects for Excel-compatible cell and range references.

This module provides the unified implementation of CellReference and RangeReference
that combines the best features from both range.py and reference_objects.py.

Replaces the duplicate implementations with a single, comprehensive solution.
"""

import re
from dataclasses import dataclass
from typing import Any, Optional, Union, TYPE_CHECKING

from .xlfunctions import xlerrors
from .constants import EXCEL_MAX_ROWS, EXCEL_MAX_COLUMNS

if TYPE_CHECKING:
    from .evaluator import Evaluator


@dataclass
class CellReference:
    """
    Unified Excel-compatible cell reference.
    
    Combines string-based parsing with coordinate-based arithmetic.
    Supports both explicit sheet references (Sheet1!A1) and coordinate operations.
    
    Attributes:
        sheet: Sheet name (empty string for current sheet)
        row: 1-based row index (Excel convention)
        column: 1-based column index (Excel convention)
        absolute_row: True if row has $ prefix ($A1)
        absolute_column: True if column has $ prefix (A$1)
        is_sheet_explicit: True if sheet was explicitly specified in original reference
    """
    
    sheet: str
    row: int
    column: int
    absolute_row: bool = False
    absolute_column: bool = False
    is_sheet_explicit: bool = True
    
    def __post_init__(self):
        """Validate Excel bounds after initialization."""
        if self.row < 1 or self.row > EXCEL_MAX_ROWS:
            raise xlerrors.RefExcelError(f"Row {self.row} is out of Excel bounds (1-{EXCEL_MAX_ROWS})")
        if self.column < 1 or self.column > EXCEL_MAX_COLUMNS:
            raise xlerrors.RefExcelError(f"Column {self.column} is out of Excel bounds (1-{EXCEL_MAX_COLUMNS})")
    
    @property
    def address(self) -> str:
        """Get Excel-style address (e.g., 'A1' or '$A$1')."""
        col_letter = self._column_to_letter(self.column)
        row_prefix = '$' if self.absolute_row else ''
        col_prefix = '$' if self.absolute_column else ''
        
        return f"{col_prefix}{col_letter}{row_prefix}{self.row}"
    
    @property
    def full_address(self) -> str:
        """Get full sheet!address format."""
        if self.sheet:
            # Handle sheet names with spaces or special characters
            if ' ' in self.sheet or "'" in self.sheet:
                sheet_part = f"'{self.sheet}'"
            else:
                sheet_part = self.sheet
            return f"{sheet_part}!{self.address}"
        else:
            return self.address
    
    @property
    def coordinate(self) -> tuple:
        """Get (row, column) coordinate tuple."""
        return (self.row, self.column)
    
    @classmethod
    def parse(cls, ref: str, current_sheet: str = None) -> 'CellReference':
        """
        Parse reference string with comprehensive support.
        
        Handles both coordinate parsing and sheet context.
        Supports formats like: A1, $A$1, Sheet1!A1, 'Sheet Name'!$A$1
        
        Args:
            ref: Reference string to parse
            current_sheet: Current sheet context for implicit references
            
        Returns:
            CellReference object with parsed coordinates and sheet info
            
        Raises:
            RefExcelError: If reference format is invalid
        """
        if not ref or not isinstance(ref, str):
            raise xlerrors.RefExcelError("Invalid reference string")
        
        ref = ref.strip()
        
        # Handle sheet references
        if '!' in ref:
            # Explicit sheet reference
            sheet_part, cell_part = ref.split('!', 1)
            sheet = cls._resolve_sheet_name(sheet_part)
            is_explicit = True
        else:
            # Implicit reference - use current sheet context
            cell_part = ref
            sheet = current_sheet or ""
            is_explicit = False
        
        # Parse cell part (e.g., A1, $A$1)
        row, column, absolute_row, absolute_col = cls._parse_cell_address(cell_part)
        
        return cls(
            sheet=sheet,
            row=row,
            column=column,
            absolute_row=absolute_row,
            absolute_column=absolute_col,
            is_sheet_explicit=is_explicit
        )
    
    def offset(self, rows: int, cols: int) -> 'CellReference':
        """
        Excel-style reference arithmetic.
        
        Args:
            rows: Number of rows to offset (can be negative)
            cols: Number of columns to offset (can be negative)
            
        Returns:
            New CellReference with offset coordinates
            
        Raises:
            RefExcelError: If offset results in coordinates outside Excel bounds
        """
        new_row = self.row + rows
        new_col = self.column + cols
        
        # Validate bounds
        if new_row < 1 or new_row > EXCEL_MAX_ROWS:
            raise xlerrors.RefExcelError(f"Row offset results in row {new_row}, outside Excel bounds")
        if new_col < 1 or new_col > EXCEL_MAX_COLUMNS:
            raise xlerrors.RefExcelError(f"Column offset results in column {new_col}, outside Excel bounds")
        
        return CellReference(
            sheet=self.sheet,
            row=new_row,
            column=new_col,
            absolute_row=self.absolute_row,
            absolute_column=self.absolute_column,
            is_sheet_explicit=self.is_sheet_explicit
        )
    
    def resolve(self, evaluator: 'Evaluator') -> Any:
        """
        Get actual cell value through evaluator.
        
        Args:
            evaluator: Evaluator instance to resolve cell value
            
        Returns:
            Cell value from the evaluator
        """
        return evaluator.evaluate(self.full_address)
    
    def is_same_sheet_as_context(self, context_sheet: str) -> bool:
        """Check if reference is in the same sheet as given context."""
        return self.sheet == context_sheet
    
    def to_tuple(self) -> tuple:
        """Convert to tuple format for backward compatibility."""
        return (self.sheet, self.address)
    
    def __str__(self) -> str:
        """Return full sheet!address format."""
        return self.full_address
    
    @staticmethod
    def _column_to_letter(col_num: int) -> str:
        """Convert column number to Excel letter(s)."""
        result = ""
        while col_num > 0:
            col_num -= 1  # Make it 0-based
            result = chr(65 + (col_num % 26)) + result
            col_num //= 26
        return result
    
    @staticmethod
    def _letter_to_column(letters: str) -> int:
        """Convert Excel column letter(s) to number."""
        result = 0
        for char in letters.upper():
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result
    
    @staticmethod
    def _resolve_sheet_name(sheet_str: str) -> str:
        """Resolve sheet name from sheet string, handling quoted names."""
        sheet_str = sheet_str.strip()
        
        # Handle quoted sheet names
        if sheet_str.startswith("'") and sheet_str.endswith("'"):
            return sheet_str[1:-1]  # Remove quotes
        
        return sheet_str
    
    @staticmethod
    def _parse_cell_address(cell_part: str) -> tuple:
        """
        Parse cell address part (e.g., A1, $A$1) into components.
        
        Returns:
            tuple: (row, column, absolute_row, absolute_column)
        """
        if not cell_part:
            raise xlerrors.RefExcelError("Empty cell address")
        
        # Pattern to match cell addresses like A1, $A$1, $A1, A$1
        pattern = r'^(\$?)([A-Z]+)(\$?)(\d+)$'
        match = re.match(pattern, cell_part.upper())
        
        if not match:
            raise xlerrors.RefExcelError(f"Invalid cell address: {cell_part}")
        
        col_absolute = bool(match.group(1))  # $ before column
        col_letters = match.group(2)
        row_absolute = bool(match.group(3))  # $ before row
        row_num = int(match.group(4))
        
        column = CellReference._letter_to_column(col_letters)
        
        return row_num, column, row_absolute, col_absolute


@dataclass
class RangeReference:
    """
    Unified Excel-compatible range reference.
    
    Represents a range of cells with start and end coordinates.
    Supports range operations and resolution.
    
    Attributes:
        start_cell: Starting cell of the range
        end_cell: Ending cell of the range
    """
    
    start_cell: CellReference
    end_cell: CellReference
    
    def __post_init__(self):
        """Validate range after initialization."""
        if self.start_cell.sheet != self.end_cell.sheet:
            raise xlerrors.RefExcelError("Range cannot span multiple sheets")
    
    @property
    def address(self) -> str:
        """Get Excel-style range address (e.g., 'A1:B2')."""
        if self.start_cell.sheet:
            # Include sheet name only once for the range
            if ' ' in self.start_cell.sheet or "'" in self.start_cell.sheet:
                sheet_part = f"'{self.start_cell.sheet}'"
            else:
                sheet_part = self.start_cell.sheet
            return f"{sheet_part}!{self.start_cell.address}:{self.end_cell.address}"
        else:
            return f"{self.start_cell.address}:{self.end_cell.address}"
    
    @property
    def dimensions(self) -> tuple:
        """Get (rows, columns) dimensions of the range."""
        rows = self.end_cell.row - self.start_cell.row + 1
        cols = self.end_cell.column - self.start_cell.column + 1
        return (rows, cols)
    
    @classmethod
    def parse(cls, ref: str, current_sheet: str = None) -> 'RangeReference':
        """
        Parse range reference string.
        
        Supports formats like: A1:B2, Sheet1!A1:B2, 'Sheet Name'!$A$1:$B$2
        
        Args:
            ref: Range reference string to parse
            current_sheet: Current sheet context for implicit references
            
        Returns:
            RangeReference object with parsed start and end cells
        """
        if ':' not in ref:
            raise xlerrors.RefExcelError("Range reference must contain ':'")
        
        # Handle sheet references
        if '!' in ref:
            sheet_part, range_part = ref.split('!', 1)
            sheet = CellReference._resolve_sheet_name(sheet_part)
            is_explicit = True
        else:
            range_part = ref
            sheet = current_sheet or ""
            is_explicit = False
        
        # Split range part
        start_addr, end_addr = range_part.split(':', 1)
        
        # Parse start and end cells
        start_cell = CellReference.parse(f"{sheet}!{start_addr}" if sheet else start_addr, current_sheet)
        end_cell = CellReference.parse(f"{sheet}!{end_addr}" if sheet else end_addr, current_sheet)
        
        # Ensure both cells have same sheet context
        start_cell.is_sheet_explicit = is_explicit
        end_cell.is_sheet_explicit = is_explicit
        
        return cls(start_cell=start_cell, end_cell=end_cell)
    
    def offset(self, rows: int, cols: int) -> 'RangeReference':
        """
        Offset entire range by specified rows/columns.
        
        Args:
            rows: Number of rows to offset
            cols: Number of columns to offset
            
        Returns:
            New RangeReference with offset coordinates
        """
        return RangeReference(
            start_cell=self.start_cell.offset(rows, cols),
            end_cell=self.end_cell.offset(rows, cols)
        )
    
    def resize(self, rows: int, cols: int) -> 'RangeReference':
        """
        Resize range to specified dimensions.
        
        Args:
            rows: New number of rows
            cols: New number of columns
            
        Returns:
            New RangeReference with specified dimensions
        """
        if rows < 1 or cols < 1:
            raise xlerrors.ValueExcelError("Range dimensions must be positive")
        
        end_cell = CellReference(
            sheet=self.start_cell.sheet,
            row=self.start_cell.row + rows - 1,
            column=self.start_cell.column + cols - 1,
            absolute_row=self.end_cell.absolute_row,
            absolute_column=self.end_cell.absolute_column,
            is_sheet_explicit=self.start_cell.is_sheet_explicit
        )
        
        return RangeReference(start_cell=self.start_cell, end_cell=end_cell)
    
    def resolve(self, evaluator: 'Evaluator') -> list:
        """
        Get actual range values through evaluator.
        
        Args:
            evaluator: Evaluator instance to resolve range values
            
        Returns:
            2D list of cell values
        """
        return evaluator.get_range_values(self.address)
    
    def __str__(self) -> str:
        """Return range address string."""
        return self.address