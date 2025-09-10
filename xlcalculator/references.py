"""
Unified Reference Objects for Excel-compatible cell and range references.

This module provides the unified implementation of CellReference and RangeReference
that combines the best features from range.py and the previous reference_objects.py.

Replaces the duplicate implementations with a single, comprehensive solution.
"""

import re
from dataclasses import dataclass
from typing import Any, TYPE_CHECKING

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
    row: int = None  # None for column references like A:A
    column: int = None  # None for row references like 1:1
    absolute_row: bool = False
    absolute_column: bool = False
    is_sheet_explicit: bool = True
    is_column_reference: bool = False  # True for A:A type references
    is_row_reference: bool = False  # True for 1:1 type references
    is_range_reference: bool = False  # True for A1:B5 type references
    original_range: str = None  # Store original range string for range references
    
    def __post_init__(self):
        """Validate Excel bounds after initialization."""
        if self.row is not None and (self.row < 1 or self.row > EXCEL_MAX_ROWS):
            raise xlerrors.RefExcelError(f"Row {self.row} is out of Excel bounds (1-{EXCEL_MAX_ROWS})")
        if self.column is not None and (self.column < 1 or self.column > EXCEL_MAX_COLUMNS):
            raise xlerrors.RefExcelError(f"Column {self.column} is out of Excel bounds (1-{EXCEL_MAX_COLUMNS})")
    
    @property
    def cell_address(self) -> str:
        """Get Excel-style cell address (e.g., 'A1', '$A$1', 'A:A')."""
        if self.is_range_reference and self.original_range:
            # Range reference like A1:B5
            return self.original_range
        elif self.is_column_reference:
            # Column reference like A:A
            col_letter = self._column_to_letter(self.column)
            col_prefix = '$' if self.absolute_column else ''
            return f"{col_prefix}{col_letter}:{col_prefix}{col_letter}"
        elif self.is_row_reference:
            # Row reference like 1:1
            row_prefix = '$' if self.absolute_row else ''
            return f"{row_prefix}{self.row}:{row_prefix}{self.row}"
        else:
            # Regular cell reference
            col_letter = self._column_to_letter(self.column)
            row_prefix = '$' if self.absolute_row else ''
            col_prefix = '$' if self.absolute_column else ''
            return f"{col_prefix}{col_letter}{row_prefix}{self.row}"
    
    @property
    def address(self) -> str:
        """Get cell address part only (e.g., 'A1', 'A1:B5')."""
        return self.cell_address
    
    @property
    def full_address(self) -> str:
        """Get full sheet!address format (e.g., 'Sheet1!A1')."""
        if self.sheet:
            # Handle sheet names with spaces or special characters
            if ' ' in self.sheet or "'" in self.sheet:
                sheet_part = f"'{self.sheet}'"
            else:
                sheet_part = self.sheet
            return f"{sheet_part}!{self.cell_address}"
        else:
            return self.cell_address
    
    @property
    def coordinate(self) -> tuple:
        """Get (row, column) coordinate tuple."""
        return (self.row, self.column)
    
    @classmethod
    def parse(cls, ref: str, current_sheet: str | None = None) -> 'CellReference':
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
        
        # Parse cell part (e.g., A1, $A$1, A:A)
        row, column, absolute_row, absolute_col, is_column_ref, is_row_ref, is_range_ref, original_range = cls._parse_cell_address(cell_part)
        
        return cls(
            sheet=sheet,
            row=row,
            column=column,
            absolute_row=absolute_row,
            absolute_column=absolute_col,
            is_sheet_explicit=is_explicit,
            is_column_reference=is_column_ref,
            is_row_reference=is_row_ref,
            is_range_reference=is_range_ref,
            original_range=original_range
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
    
    def to_tuple(self) -> tuple[str, str]:
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
    def _parse_cell_address(cell_part: str) -> tuple[int, int, bool, bool, bool, bool, bool, str]:
        """
        Parse cell address part (e.g., A1, $A$1, A:A) into components.
        
        Returns:
            tuple: (row, column, absolute_row, absolute_column, is_column_ref, is_row_ref, is_range_ref, original_range)
        """
        if not cell_part:
            raise xlerrors.RefExcelError("Empty cell address")
        
        # Check for column reference (A:A, $A:$A, etc.)
        column_pattern = r'^(\$?)([A-Z]+):(\$?)([A-Z]+)$'
        column_match = re.match(column_pattern, cell_part.upper())
        
        if column_match:
            # Column reference like A:A
            col1_absolute = bool(column_match.group(1))
            col1_letters = column_match.group(2)
            col2_absolute = bool(column_match.group(3))
            col2_letters = column_match.group(4)
            
            # For now, handle single column references (A:A)
            if col1_letters == col2_letters:
                column = CellReference._letter_to_column(col1_letters)
                return None, column, False, col1_absolute, True, False, False, None
            else:
                raise xlerrors.RefExcelError(f"Multi-column ranges not supported yet: {cell_part}")
        
        # Check for row reference (1:1, $1:$1, etc.)
        row_pattern = r'^(\$?)(\d+):(\$?)(\d+)$'
        row_match = re.match(row_pattern, cell_part.upper())
        
        if row_match:
            # Row reference like 1:1
            row1_absolute = bool(row_match.group(1))
            row1_num = int(row_match.group(2))
            row2_absolute = bool(row_match.group(3))
            row2_num = int(row_match.group(4))
            
            # For now, handle single row references (1:1)
            if row1_num == row2_num:
                return row1_num, None, row1_absolute, False, False, True, False, None
            else:
                raise xlerrors.RefExcelError(f"Multi-row ranges not supported yet: {cell_part}")
        
        # Check for cell range (A1:B5, $A$1:$B$5, etc.)
        range_pattern = r'^(\$?)([A-Z]+)(\$?)(\d+):(\$?)([A-Z]+)(\$?)(\d+)$'
        range_match = re.match(range_pattern, cell_part.upper())
        
        if range_match:
            # Cell range like A1:B5
            # For compatibility, return the first cell of the range
            col_absolute = bool(range_match.group(1))  # $ before column
            col_letters = range_match.group(2)
            row_absolute = bool(range_match.group(3))  # $ before row
            row_num = int(range_match.group(4))
            
            column = CellReference._letter_to_column(col_letters)
            
            return row_num, column, row_absolute, col_absolute, False, False, True, cell_part
        
        # Pattern to match cell addresses like A1, $A$1, $A1, A$1
        cell_pattern = r'^(\$?)([A-Z]+)(\$?)(\d+)$'
        cell_match = re.match(cell_pattern, cell_part.upper())
        
        if cell_match:
            col_absolute = bool(cell_match.group(1))  # $ before column
            col_letters = cell_match.group(2)
            row_absolute = bool(cell_match.group(3))  # $ before row
            row_num = int(cell_match.group(4))
            
            column = CellReference._letter_to_column(col_letters)
            
            return row_num, column, row_absolute, col_absolute, False, False, False, None
        
        raise xlerrors.RefExcelError(f"Invalid cell address: {cell_part}")


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
            return f"{sheet_part}!{self.start_cell.cell_address}:{self.end_cell.cell_address}"
        else:
            return f"{self.start_cell.cell_address}:{self.end_cell.cell_address}"
    
    @property
    def dimensions(self) -> tuple[int, int]:
        """Get (rows, columns) dimensions of the range."""
        rows = self.end_cell.row - self.start_cell.row + 1
        cols = self.end_cell.column - self.start_cell.column + 1
        return (rows, cols)
    
    @classmethod
    def parse(cls, ref: str, current_sheet: str | None = None) -> 'RangeReference':
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
            # Single cell treated as 1x1 range
            cell_ref = CellReference.parse(ref, current_sheet)
            return cls(start_cell=cell_ref, end_cell=cell_ref)
        
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