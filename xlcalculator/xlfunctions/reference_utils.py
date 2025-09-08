"""
Reference Resolution Utilities for Dynamic Range Functions

This module provides centralized utilities for parsing, manipulating, and validating
Excel cell and range references. Used by OFFSET, INDEX, INDIRECT and other dynamic
range functions.
"""

import re
from typing import Tuple, Union, Optional
from . import xlerrors


class ReferenceResolver:
    """
    Centralized reference resolution for dynamic range functions.
    
    Handles conversion between Excel A1 notation and internal coordinates,
    reference validation, and range calculations.
    """
    
    # Excel limits (Excel 2019/365)
    MAX_ROWS = 1048576
    MAX_COLS = 16384
    
    @staticmethod
    def parse_cell_reference(ref: str) -> Tuple[int, int]:
        """
        Parse A1 notation to (row, col) coordinates.
        
        Args:
            ref: Cell reference in A1 notation (e.g., "A1", "B2", "AA10")
            
        Returns:
            Tuple of (row, col) where both are 1-based
            
        Raises:
            ValueExcelError: If reference format is invalid
            
        Examples:
            parse_cell_reference("A1") → (1, 1)
            parse_cell_reference("B2") → (2, 2)
            parse_cell_reference("AA1") → (1, 27)
        """
        if not isinstance(ref, str):
            raise xlerrors.ValueExcelError(f"Reference must be string, got {type(ref)}")
        
        # Remove sheet name if present (Sheet1!A1 → A1)
        if '!' in ref:
            sheet_part, ref = ref.split('!', 1)
            if not sheet_part:  # Empty sheet name like "!A1"
                raise xlerrors.ValueExcelError(f"Invalid sheet reference: '{ref}'")
            if not ref:  # Empty cell part like "Sheet1!"
                raise xlerrors.ValueExcelError(f"Invalid cell reference: '{ref}'")
        
        # Remove $ signs for absolute references ($A$1 → A1)
        ref = ref.replace('$', '')
        
        # Match pattern: letters followed by numbers
        match = re.match(r'^([A-Z]+)(\d+)$', ref.upper().strip())
        if not match:
            raise xlerrors.ValueExcelError(f"Invalid cell reference format: '{ref}'")
        
        col_str, row_str = match.groups()
        
        # Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)
        col = 0
        for char in col_str:
            col = col * 26 + (ord(char) - ord('A') + 1)
        
        row = int(row_str)
        
        # Validate bounds
        ReferenceResolver.validate_bounds(row, col)
        
        return row, col
    
    @staticmethod
    def cell_to_string(row: int, col: int) -> str:
        """
        Convert (row, col) coordinates to A1 notation.
        
        Args:
            row: Row number (1-based)
            col: Column number (1-based)
            
        Returns:
            Cell reference in A1 notation
            
        Raises:
            RefExcelError: If coordinates are out of bounds
            
        Examples:
            cell_to_string(1, 1) → "A1"
            cell_to_string(2, 2) → "B2"
            cell_to_string(1, 27) → "AA1"
        """
        ReferenceResolver.validate_bounds(row, col)
        
        # Convert column number to letters
        col_str = ""
        temp_col = col
        while temp_col > 0:
            temp_col -= 1
            col_str = chr(ord('A') + temp_col % 26) + col_str
            temp_col //= 26
        
        return f"{col_str}{row}"
    
    @staticmethod
    def parse_range_reference(ref: str) -> Tuple[Tuple[int, int], Tuple[int, int]]:
        """
        Parse A1:B2 notation to coordinate pairs.
        
        Args:
            ref: Range reference (e.g., "A1:B2", "A1", "Sheet1!A1:B2")
            
        Returns:
            Tuple of ((start_row, start_col), (end_row, end_col))
            
        Raises:
            ValueExcelError: If reference format is invalid
            
        Examples:
            parse_range_reference("A1:B2") → ((1, 1), (2, 2))
            parse_range_reference("A1") → ((1, 1), (1, 1))
        """
        if not isinstance(ref, str):
            raise xlerrors.ValueExcelError(f"Reference must be string, got {type(ref)}")
        
        # Handle sheet names
        sheet_name = None
        if '!' in ref:
            sheet_name, ref = ref.split('!', 1)
        
        if ':' not in ref:
            # Single cell reference
            row, col = ReferenceResolver.parse_cell_reference(ref)
            return (row, col), (row, col)
        
        # Range reference
        start_ref, end_ref = ref.split(':', 1)
        if not start_ref or not end_ref:
            raise xlerrors.ValueExcelError(f"Invalid range reference: '{ref}'")
        
        start = ReferenceResolver.parse_cell_reference(start_ref)
        end = ReferenceResolver.parse_cell_reference(end_ref)
        
        # Ensure start is top-left, end is bottom-right
        start_row, start_col = min(start[0], end[0]), min(start[1], end[1])
        end_row, end_col = max(start[0], end[0]), max(start[1], end[1])
        
        return (start_row, start_col), (end_row, end_col)
    
    @staticmethod
    def range_to_string(start: Tuple[int, int], end: Tuple[int, int]) -> str:
        """
        Convert coordinate pairs to A1:B2 notation.
        
        Args:
            start: (start_row, start_col) coordinates
            end: (end_row, end_col) coordinates
            
        Returns:
            Range reference in A1:B2 notation
            
        Examples:
            range_to_string((1, 1), (2, 2)) → "A1:B2"
            range_to_string((1, 1), (1, 1)) → "A1"
        """
        start_ref = ReferenceResolver.cell_to_string(start[0], start[1])
        
        if start == end:
            # Single cell
            return start_ref
        
        end_ref = ReferenceResolver.cell_to_string(end[0], end[1])
        return f"{start_ref}:{end_ref}"
    
    @staticmethod
    def offset_reference(ref: str, rows: int, cols: int, 
                        height: Optional[int] = None, 
                        width: Optional[int] = None) -> str:
        """
        Apply offset to reference and return new reference.
        
        Args:
            ref: Original reference (cell or range)
            rows: Number of rows to offset (positive = down, negative = up)
            cols: Number of columns to offset (positive = right, negative = left)
            height: Optional height of result range
            width: Optional width of result range
            
        Returns:
            New reference string after applying offset
            
        Raises:
            RefExcelError: If result is out of bounds
            ValueExcelError: If parameters are invalid
            
        Examples:
            offset_reference("A1", 1, 1) → "B2"
            offset_reference("A1:B2", 1, 1) → "B2:C3"
            offset_reference("A1", 1, 1, 2, 2) → "B2:C3"
        """
        # Parse original reference
        if ':' in ref:
            # Range reference
            start, end = ReferenceResolver.parse_range_reference(ref)
            original_height = end[0] - start[0] + 1
            original_width = end[1] - start[1] + 1
        else:
            # Single cell reference
            start = ReferenceResolver.parse_cell_reference(ref)
            original_height = 1
            original_width = 1
        
        # Apply offset to start position
        new_start_row = start[0] + rows
        new_start_col = start[1] + cols
        
        # Determine result dimensions
        result_height = height if height is not None else original_height
        result_width = width if width is not None else original_width
        
        if result_height < 1 or result_width < 1:
            raise xlerrors.ValueExcelError("Height and width must be positive")
        
        # Calculate end position
        new_end_row = new_start_row + result_height - 1
        new_end_col = new_start_col + result_width - 1
        
        # Validate bounds
        ReferenceResolver.validate_bounds(new_start_row, new_start_col)
        ReferenceResolver.validate_bounds(new_end_row, new_end_col)
        
        # Return appropriate reference format
        if result_height == 1 and result_width == 1:
            return ReferenceResolver.cell_to_string(new_start_row, new_start_col)
        else:
            return ReferenceResolver.range_to_string(
                (new_start_row, new_start_col), 
                (new_end_row, new_end_col)
            )
    
    @staticmethod
    def validate_bounds(row: int, col: int, 
                       max_row: int = None, max_col: int = None) -> None:
        """
        Validate that coordinates are within Excel bounds.
        
        Args:
            row: Row number (1-based)
            col: Column number (1-based)
            max_row: Maximum row (default: Excel limit)
            max_col: Maximum column (default: Excel limit)
            
        Raises:
            RefExcelError: If coordinates are out of bounds
        """
        if max_row is None:
            max_row = ReferenceResolver.MAX_ROWS
        if max_col is None:
            max_col = ReferenceResolver.MAX_COLS
        
        if row < 1 or row > max_row:
            raise xlerrors.RefExcelError(
                f"Row {row} is out of bounds (1-{max_row})")
        
        if col < 1 or col > max_col:
            raise xlerrors.RefExcelError(
                f"Column {col} is out of bounds (1-{max_col})")
    
    @staticmethod
    def get_range_dimensions(ref: str) -> Tuple[int, int]:
        """
        Get the dimensions (height, width) of a range reference.
        
        Args:
            ref: Range reference string
            
        Returns:
            Tuple of (height, width)
            
        Examples:
            get_range_dimensions("A1") → (1, 1)
            get_range_dimensions("A1:C3") → (3, 3)
        """
        start, end = ReferenceResolver.parse_range_reference(ref)
        height = end[0] - start[0] + 1
        width = end[1] - start[1] + 1
        return height, width
    
    @staticmethod
    def normalize_reference(ref: str) -> str:
        """
        Normalize a reference to standard format.
        
        Args:
            ref: Reference string (may have $ signs, mixed case, etc.)
            
        Returns:
            Normalized reference string
            
        Examples:
            normalize_reference("$a$1") → "A1"
            normalize_reference("sheet1!$A$1:$B$2") → "Sheet1!A1:B2"
        """
        if '!' in ref:
            sheet, cell_part = ref.split('!', 1)
            # Keep original sheet name case, normalize cell part
            if ':' in cell_part:
                start_ref, end_ref = cell_part.split(':', 1)
                start = ReferenceResolver.parse_cell_reference(start_ref)
                end = ReferenceResolver.parse_cell_reference(end_ref)
                normalized_cell = ReferenceResolver.range_to_string(start, end)
            else:
                coords = ReferenceResolver.parse_cell_reference(cell_part)
                normalized_cell = ReferenceResolver.cell_to_string(*coords)
            return f"{sheet}!{normalized_cell}"
        else:
            if ':' in ref:
                start, end = ReferenceResolver.parse_range_reference(ref)
                return ReferenceResolver.range_to_string(start, end)
            else:
                coords = ReferenceResolver.parse_cell_reference(ref)
                return ReferenceResolver.cell_to_string(*coords)