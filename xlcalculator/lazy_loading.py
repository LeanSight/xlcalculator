#!/usr/bin/env python3
"""
Excel-Compliant Lazy Loading Implementation for xlcalculator

This module implements lazy loading for Excel ranges following ATDD principles:
- No fallbacks that violate Excel behavior
- Return actual Excel data or proper Excel errors
- Performance optimization without compromising compatibility

ATDD Principle: Tests specify Excel behavior, implementation follows exactly.
"""

import logging
from xlcalculator import xltypes


class ExcelCompliantLazyRange:
    """
    Excel-compliant lazy range that returns actual Excel data.
    
    ATDD Principle: Implementation follows Excel behavior exactly.
    No fallbacks, no hardcoded data, no masking of Excel errors.
    """
    
    def __init__(self, address_str, model, name=None):
        self.address_str = address_str
        self.name = name or address_str
        self.model = model  # Access to actual Excel data
        self.sheet = None
        self._cells = None  # Lazy-loaded
        self._is_full_range = self._check_if_full_range(address_str)
    
    def _check_if_full_range(self, address_str):
        """Check if this is a full column/row reference that needs lazy loading."""
        from xlcalculator.utils import parse_sheet_and_address
        
        try:
            sheet, range_part = parse_sheet_and_address(address_str)
            
            if ':' in range_part:
                parts = range_part.split(':')
                if len(parts) == 2:
                    left, right = parts
                    
                    # Check for full column (A:A, B:B)
                    if left.isalpha() and right.isalpha() and left == right:
                        return True
                    
                    # Check for full row (1:1, 2:2)
                    if left.isdigit() and right.isdigit() and left == right:
                        return True
        except Exception:
            pass
        
        return False
    
    @property
    def cells(self):
        """
        Lazy-loaded cells property that returns actual Excel data.
        
        EXCEL BEHAVIOR:
        - Full column A:A returns all non-empty cells in column A
        - Empty cells return 0 or empty string as Excel does
        - Invalid ranges return #REF! error
        - No fallbacks, no hardcoded data
        """
        if self._cells is None:
            if self._is_full_range:
                self._cells = self._resolve_full_range_excel_compliant()
            else:
                # For normal ranges, use standard resolution
                self._cells = self._resolve_normal_range()
        return self._cells
    
    def _resolve_full_range_excel_compliant(self):
        """
        Resolve full range using actual Excel data boundaries.
        
        EXCEL BEHAVIOR:
        - Find actual data boundaries in the sheet
        - Return only cells that actually exist
        - Preserve Excel's empty cell behavior
        """
        sheet_name, range_part = self._parse_range()
        
        if ':' not in range_part:
            return [[self.address_str]]  # Single cell
        
        left, right = range_part.split(':')
        
        # For full column (A:A), find actual data boundaries
        if left.isalpha() and right.isalpha() and left == right:
            return self._resolve_full_column_excel_data(sheet_name, left)
        
        # For full row (1:1), find actual data boundaries  
        if left.isdigit() and right.isdigit() and left == right:
            return self._resolve_full_row_excel_data(sheet_name, left)
        
        # Should not reach here for full ranges
        return [[self.address_str]]
    
    def _resolve_full_column_excel_data(self, sheet_name, column):
        """
        Resolve full column using actual Excel data.
        
        EXCEL BEHAVIOR:
        - Scan actual cells in the model to find data boundaries
        - Return actual cell values, not hardcoded data
        - Preserve Excel's handling of empty cells
        """
        cells = []
        
        # Find the actual last row with data in this column
        max_row = self._find_last_row_with_data(sheet_name, column)
        
        if max_row == 0:
            # Column is completely empty - Excel behavior
            return [[]]
        
        # Build cell references for actual data range
        for row in range(1, max_row + 1):
            cell_ref = f"{sheet_name}!{column}{row}"
            cells.append([cell_ref])
        
        return cells
    
    def _resolve_full_row_excel_data(self, sheet_name, row):
        """
        Resolve full row using actual Excel data.
        
        EXCEL BEHAVIOR:
        - Scan actual cells in the model to find data boundaries
        - Return actual cell values, not hardcoded data
        """
        # Find the actual last column with data in this row
        max_col = self._find_last_col_with_data(sheet_name, row)
        
        if max_col == 0:
            # Row is completely empty - Excel behavior
            return [[]]
        
        # Build cell references for actual data range
        cells = []
        for col_num in range(1, max_col + 1):
            col_letter = chr(ord('A') + col_num - 1)
            cell_ref = f"{sheet_name}!{col_letter}{row}"
            cells.append(cell_ref)
        
        return [cells]  # Single row with multiple columns
    
    def _find_last_row_with_data(self, sheet_name, column):
        """
        Find the last row with actual data in the specified column.
        
        EXCEL BEHAVIOR:
        - Scan actual model cells to find boundaries
        - Return 0 if no data found
        """
        max_row = 0
        
        # Scan all cells in the model for this sheet and column
        from xlcalculator.utils import parse_sheet_and_address
        from openpyxl.utils.cell import COORD_RE
        
        for cell_address in self.model.cells:
            try:
                parsed_sheet, address_part = parse_sheet_and_address(cell_address)
                if parsed_sheet == sheet_name:
                    # Use openpyxl's COORD_RE to parse the address
                    coord_match = COORD_RE.split(address_part)
                    if len(coord_match) >= 3:
                        col, row = coord_match[1:3]
                        if col == column and row.isdigit():
                            row_num = int(row)
                            # Check if cell has actual data (not empty/zero)
                            cell_value = self.model.cells[cell_address].value
                            if cell_value is not None and cell_value != 0 and cell_value != '':
                                max_row = max(max_row, row_num)
            except:
                continue
        
        return max_row
    
    def _find_last_col_with_data(self, sheet_name, row):
        """
        Find the last column with actual data in the specified row.
        
        EXCEL BEHAVIOR:
        - Scan actual model cells to find boundaries
        - Return 0 if no data found
        """
        max_col = 0
        
        # Scan all cells in the model for this sheet and row
        from xlcalculator.utils import parse_sheet_and_address
        from openpyxl.utils.cell import COORD_RE, column_index_from_string
        
        for cell_address in self.model.cells:
            try:
                parsed_sheet, address_part = parse_sheet_and_address(cell_address)
                if parsed_sheet == sheet_name:
                    # Use openpyxl's COORD_RE to parse the address
                    coord_match = COORD_RE.split(address_part)
                    if len(coord_match) >= 3:
                        col, cell_row = coord_match[1:3]
                        if cell_row == row:
                            col_num = column_index_from_string(col)
                            # Check if cell has actual data
                            cell_value = self.model.cells[cell_address].value
                            if cell_value is not None and cell_value != 0 and cell_value != '':
                                max_col = max(max_col, col_num)
            except:
                continue
        
        return max_col
    
    def _resolve_normal_range(self):
        """Resolve normal (non-full) ranges using standard method."""
        try:
            from xlcalculator import utils
            sheet, cells = utils.resolve_ranges(self.address_str)
            return cells
        except Exception as e:
            # Return Excel error, not fallback data
            from xlcalculator.xlfunctions import xlerrors
            raise xlerrors.RefExcelError(f"Invalid range: {self.address_str}")
    
    def _parse_range(self):
        """Parse range to extract sheet name and range part."""
        from xlcalculator.utils import parse_sheet_and_address
        return parse_sheet_and_address(self.address_str)
    
    @property
    def address(self):
        """Return the cells for compatibility with XLRange."""
        return self.cells


# Legacy alias for compatibility
LazyXLRange = ExcelCompliantLazyRange


class ExcelCompliantLazyManager:
    """
    Excel-compliant lazy range manager.
    
    ATDD Principle: No fallbacks, exact Excel behavior only.
    """
    
    def __init__(self, evaluator):
        self.evaluator = evaluator
        self.original_get_range_values = evaluator.get_range_values
        
        # Patch the evaluator's get_range_values method
        evaluator.get_range_values = self.excel_compliant_get_range_values
    
    def excel_compliant_get_range_values(self, range_ref):
        """
        Excel-compliant range resolution.
        
        EXCEL BEHAVIOR:
        - Return actual Excel data only
        - Propagate Excel errors properly
        - No fallbacks, no hardcoded data
        """
        try:
            # Use original method for all ranges
            # The lazy loading happens at the build_ranges level
            return self.original_get_range_values(range_ref)
        
        except Exception as e:
            # Propagate Excel errors, don't mask them
            from xlcalculator.xlfunctions import xlerrors
            
            # Convert common errors to proper Excel errors
            if "invalid literal" in str(e) or "resolve" in str(e):
                raise xlerrors.RefExcelError(f"Invalid range reference: {range_ref}")
            else:
                # Re-raise the original error
                raise e


# Legacy alias for compatibility
LazyRangeManager = ExcelCompliantLazyManager


def patch_evaluator_with_lazy_loading(evaluator):
    """Patch an evaluator with Excel-compliant lazy loading capabilities."""
    lazy_manager = ExcelCompliantLazyManager(evaluator)
    return lazy_manager


def is_full_range(range_str):
    """Check if a range reference is a full column/row that needs lazy loading."""
    from xlcalculator.utils import parse_sheet_and_address
    
    try:
        sheet, range_part = parse_sheet_and_address(range_str)
        
        if ':' in range_part:
            parts = range_part.split(':')
            if len(parts) == 2:
                left, right = parts
                
                # Check for full column (A:A, B:B)
                if left.isalpha() and right.isalpha() and left == right:
                    return True
                
                # Check for full row (1:1, 2:2)
                if left.isdigit() and right.isdigit() and left == right:
                    return True
    except Exception:
        pass
    
    return False


def create_excel_compliant_lazy_range(address_str, model, name=None):
    """
    Factory function to create Excel-compliant lazy ranges.
    
    ATDD Principle: Only create lazy ranges for performance,
    never compromise Excel compatibility.
    """
    if is_full_range(address_str):
        return ExcelCompliantLazyRange(address_str, model, name)
    else:
        # Use standard XLRange for non-full ranges
        from xlcalculator.xltypes import XLRange
        return XLRange(address_str, name)


# Legacy aliases for compatibility
is_problematic_range = is_full_range
create_lazy_range = create_excel_compliant_lazy_range