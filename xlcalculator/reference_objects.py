"""
Reference Objects for Excel-compatible cell and range references.

This module implements Excel-compatible reference objects that preserve
coordinate information and support reference arithmetic operations.

ATDD Implementation: Based on REFERENCE_OBJECTS_DESIGN.md and Excel documentation.
"""

import re
from dataclasses import dataclass
from typing import Any, Optional, Union, TYPE_CHECKING

from .xlfunctions import xlerrors

if TYPE_CHECKING:
    from .evaluator import Evaluator


@dataclass
class CellReference:
    """
    Excel-compatible single cell reference.
    
    Represents a single cell with sheet, row, and column coordinates.
    Supports Excel-style reference arithmetic and parsing.
    
    Attributes:
        sheet: Sheet name (empty string for current sheet)
        row: 1-based row index (Excel convention)
        column: 1-based column index (Excel convention)
        absolute_row: True if row has $ prefix ($A1)
        absolute_column: True if column has $ prefix (A$1)
    """
    
    sheet: str
    row: int
    column: int
    absolute_row: bool = False
    absolute_column: bool = False
    
    def __post_init__(self):
        """Validate Excel bounds after initialization."""
        if self.row < 1 or self.row > 1048576:
            raise xlerrors.RefExcelError(f"Row {self.row} is out of Excel bounds (1-1048576)")
        if self.column < 1 or self.column > 16384:
            raise xlerrors.RefExcelError(f"Column {self.column} is out of Excel bounds (1-16384)")
    
    @property
    def address(self) -> str:
        """Get Excel-style address (e.g., 'Sheet1!$A$1')."""
        col_letter = self._column_to_letter(self.column)
        row_prefix = '$' if self.absolute_row else ''
        col_prefix = '$' if self.absolute_column else ''
        
        cell_part = f"{col_prefix}{col_letter}{row_prefix}{self.row}"
        
        if self.sheet:
            # Handle sheet names with spaces
            if ' ' in self.sheet or "'" in self.sheet:
                sheet_part = f"'{self.sheet}'"
            else:
                sheet_part = self.sheet
            return f"{sheet_part}!{cell_part}"
        else:
            return cell_part
    
    @property
    def coordinate(self) -> tuple:
        """Get (row, column) coordinate tuple."""
        return (self.row, self.column)
    
    def offset(self, rows: int, cols: int) -> 'CellReference':
        """
        Excel-style reference arithmetic.
        
        Args:
            rows: Number of rows to offset (can be negative)
            cols: Number of columns to offset (can be negative)
            
        Returns:
            New CellReference with offset coordinates
            
        Raises:
            RefExcelError: If offset goes out of Excel bounds
        """
        new_row = self.row + rows
        new_col = self.column + cols
        
        # Validate bounds (Excel limits: 1048576 rows, 16384 columns)
        if new_row < 1 or new_row > 1048576:
            raise xlerrors.RefExcelError("Row index out of Excel bounds")
        if new_col < 1 or new_col > 16384:
            raise xlerrors.RefExcelError("Column index out of Excel bounds")
            
        return CellReference(
            sheet=self.sheet,
            row=new_row,
            column=new_col,
            absolute_row=self.absolute_row,
            absolute_column=self.absolute_column
        )
    
    def resolve(self, evaluator: 'Evaluator') -> Any:
        """Get actual cell value through evaluator."""
        return evaluator.get_cell_value(self.address)
    
    @classmethod
    def parse(cls, address: str, current_sheet: str = None) -> 'CellReference':
        """
        Parse Excel address string to CellReference.
        
        Args:
            address: Excel address string (e.g., "A1", "Sheet1!B2", "$A$1")
            current_sheet: Default sheet name if not specified in address
            
        Returns:
            CellReference object
            
        Raises:
            RefExcelError: If address format is invalid
        """
        if not address or not isinstance(address, str):
            raise xlerrors.RefExcelError("Invalid reference: empty or non-string")
        
        address = address.strip()
        if not address:
            raise xlerrors.RefExcelError("Invalid reference: empty string")
        
        # Handle sheet prefix
        sheet = ""
        cell_part = address
        
        if '!' in address:
            sheet_part, cell_part = address.split('!', 1)
            # Remove quotes from sheet name if present
            if sheet_part.startswith("'") and sheet_part.endswith("'"):
                sheet = sheet_part[1:-1]
            else:
                sheet = sheet_part
        elif current_sheet:
            sheet = current_sheet
        
        # Parse cell part using regex
        # Pattern: optional $ + letters + optional $ + digits
        pattern = r'^(\$?)([A-Za-z]+)(\$?)(\d+)$'
        match = re.match(pattern, cell_part)
        
        if not match:
            raise xlerrors.RefExcelError(f"Invalid reference format: {address}")
        
        col_absolute = bool(match.group(1))  # $ before column
        col_letters = match.group(2).upper()
        row_absolute = bool(match.group(3))  # $ before row
        row_digits = match.group(4)
        
        try:
            row_num = int(row_digits)
            col_num = cls._letter_to_column(col_letters)
        except (ValueError, OverflowError):
            raise xlerrors.RefExcelError(f"Invalid reference format: {address}")
        
        return cls(
            sheet=sheet,
            row=row_num,
            column=col_num,
            absolute_row=row_absolute,
            absolute_column=col_absolute
        )
    
    @staticmethod
    def _column_to_letter(col_num: int) -> str:
        """Convert 1-based column number to Excel letter."""
        if col_num < 1:
            raise ValueError("Column number must be >= 1")
        
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord('A')) + result
            col_num //= 26
        return result
    
    @staticmethod
    def _letter_to_column(letters: str) -> int:
        """Convert Excel column letters to 1-based number."""
        if not letters or not letters.isalpha():
            raise ValueError("Invalid column letters")
        
        result = 0
        for char in letters.upper():
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result


@dataclass
class RangeReference:
    """
    Excel-compatible range reference.
    
    Represents a range of cells with start and end coordinates.
    Supports range operations and Excel-style addressing.
    """
    
    start_cell: CellReference
    end_cell: CellReference
    
    def __post_init__(self):
        """Validate range after initialization."""
        if self.start_cell.sheet != self.end_cell.sheet:
            raise xlerrors.RefExcelError("Range cannot span multiple sheets")
    
    @property
    def address(self) -> str:
        """Get Excel-style range address (e.g., 'Sheet1!A1:C3')."""
        if self.start_cell == self.end_cell:
            # Single cell range
            return self.start_cell.address
        
        start_addr = self.start_cell.address
        end_addr = self.end_cell.address
        
        # If same sheet, optimize format
        if self.start_cell.sheet == self.end_cell.sheet and self.start_cell.sheet:
            start_cell_part = start_addr.split('!')[1]
            end_cell_part = end_addr.split('!')[1]
            return f"{self.start_cell.sheet}!{start_cell_part}:{end_cell_part}"
        else:
            return f"{start_addr}:{end_addr}"
    
    @property
    def dimensions(self) -> tuple:
        """Get (rows, columns) dimensions."""
        rows = self.end_cell.row - self.start_cell.row + 1
        cols = self.end_cell.column - self.start_cell.column + 1
        return (rows, cols)
    
    def offset(self, rows: int, cols: int) -> 'RangeReference':
        """Offset entire range by specified rows/columns."""
        return RangeReference(
            start_cell=self.start_cell.offset(rows, cols),
            end_cell=self.end_cell.offset(rows, cols)
        )
    
    def resize(self, rows: int, cols: int) -> 'RangeReference':
        """Resize range to specified dimensions."""
        if rows <= 0 or cols <= 0:
            raise xlerrors.ValueExcelError("Range dimensions must be positive")
        
        new_end_row = self.start_cell.row + rows - 1
        new_end_col = self.start_cell.column + cols - 1
        
        new_end_cell = CellReference(
            sheet=self.start_cell.sheet,
            row=new_end_row,
            column=new_end_col
        )
        
        return RangeReference(
            start_cell=self.start_cell,
            end_cell=new_end_cell
        )
    
    def get_cell(self, row_offset: int, col_offset: int) -> CellReference:
        """Get specific cell within range by offset."""
        target_row = self.start_cell.row + row_offset
        target_col = self.start_cell.column + col_offset
        
        if (target_row > self.end_cell.row or 
            target_col > self.end_cell.column or
            target_row < self.start_cell.row or
            target_col < self.start_cell.column):
            raise xlerrors.RefExcelError("Cell offset outside range bounds")
        
        return CellReference(
            sheet=self.start_cell.sheet,
            row=target_row,
            column=target_col
        )
    
    def resolve(self, evaluator: 'Evaluator') -> list:
        """Get 2D array of values from range."""
        return evaluator.get_range_values(self.address)
    
    @classmethod
    def parse(cls, address: str, current_sheet: str = None) -> 'RangeReference':
        """
        Parse Excel range address to RangeReference.
        
        Args:
            address: Excel range address (e.g., "A1:B2", "Sheet1!A1:C3")
            current_sheet: Default sheet name if not specified
            
        Returns:
            RangeReference object
        """
        if ':' not in address:
            # Single cell treated as 1x1 range
            cell_ref = CellReference.parse(address, current_sheet)
            return cls(start_cell=cell_ref, end_cell=cell_ref)
        
        start_addr, end_addr = address.split(':', 1)
        
        # Parse start cell
        start_cell = CellReference.parse(start_addr, current_sheet)
        
        # For end cell, inherit sheet from start if not specified
        if '!' not in end_addr and start_cell.sheet:
            end_addr = f"{start_cell.sheet}!{end_addr}"
        
        end_cell = CellReference.parse(end_addr, current_sheet)
        
        return cls(start_cell=start_cell, end_cell=end_cell)


@dataclass
class NamedReference:
    """
    Excel-compatible named range reference.
    
    Represents a named range that can be resolved to actual cell/range references.
    """
    
    name: str
    workbook_scope: bool = True
    sheet: str = None
    
    def resolve_to_reference(self, evaluator: 'Evaluator') -> Union[CellReference, RangeReference]:
        """Resolve named reference to actual cell/range reference."""
        definition = evaluator.get_defined_name(self.name, self.sheet)
        if definition is None:
            raise xlerrors.NameExcelError(f"Name '{self.name}' not found")
        
        # Parse the definition to get actual reference
        if ':' in definition:
            return RangeReference.parse(definition)
        else:
            return CellReference.parse(definition)
    
    def resolve(self, evaluator: 'Evaluator') -> Any:
        """Get actual value(s) from named reference."""
        reference = self.resolve_to_reference(evaluator)
        return reference.resolve(evaluator)