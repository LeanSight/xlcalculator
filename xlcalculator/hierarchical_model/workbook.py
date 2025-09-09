"""
Workbook Class Implementation

Represents an Excel workbook with proper worksheet management,
defined names, and cross-sheet operations.
"""
from dataclasses import dataclass, field
from typing import Dict, Optional, Any, Union
import logging

from .worksheet import Worksheet
from .cell import Cell
from .range import Range
from ..xltypes import XLCell, XLRange


@dataclass
class Workbook:
    """Excel Workbook with hierarchical worksheet structure."""
    
    name: str = ""
    worksheets: Dict[str, Worksheet] = field(default_factory=dict)
    defined_names: Dict[str, Any] = field(default_factory=dict)
    active_sheet: Optional[str] = None
    
    def add_worksheet(self, name: str) -> Worksheet:
        """Add a new worksheet to the workbook.
        
        Args:
            name: Name of the worksheet to create
            
        Returns:
            The created Worksheet instance
            
        Raises:
            ValueError: If worksheet with same name already exists
        """
        if name in self.worksheets:
            raise ValueError(f"Worksheet '{name}' already exists")
        
        worksheet = Worksheet(name=name, workbook=self)
        self.worksheets[name] = worksheet
        
        # Set as active sheet if it's the first one
        if self.active_sheet is None:
            self.active_sheet = name
            
        return worksheet
    
    def get_worksheet(self, name: str) -> Worksheet:
        """Get worksheet by name.
        
        Args:
            name: Name of the worksheet to retrieve
            
        Returns:
            The Worksheet instance
            
        Raises:
            KeyError: If worksheet doesn't exist
        """
        if name not in self.worksheets:
            raise KeyError(f"Worksheet '{name}' not found")
        
        return self.worksheets[name]
    
    def remove_worksheet(self, name: str) -> None:
        """Remove worksheet from workbook.
        
        Args:
            name: Name of the worksheet to remove
            
        Raises:
            KeyError: If worksheet doesn't exist
        """
        if name not in self.worksheets:
            raise KeyError(f"Worksheet '{name}' not found")
        
        del self.worksheets[name]
        
        # Update active sheet if removed sheet was active
        if self.active_sheet == name:
            if self.worksheets:
                self.active_sheet = next(iter(self.worksheets.keys()))
            else:
                self.active_sheet = None
    
    def get_cell(self, address: str) -> Cell:
        """Get cell by full address (Sheet!Cell format).
        
        Args:
            address: Full cell address like "Sheet1!A1"
            
        Returns:
            The Cell instance
            
        Raises:
            ValueError: If address format is invalid
            KeyError: If worksheet doesn't exist
        """
        sheet_name, cell_address = self._parse_full_address(address)
        worksheet = self.get_worksheet(sheet_name)
        return worksheet.get_cell(cell_address)
    
    def set_cell_value(self, address: str, value: Any) -> None:
        """Set cell value by full address.
        
        Args:
            address: Full cell address like "Sheet1!A1"
            value: Value to set in the cell
        """
        sheet_name, cell_address = self._parse_full_address(address)
        worksheet = self.get_worksheet(sheet_name)
        worksheet.set_cell_value(cell_address, value)
    
    def get_cell_value(self, address: str) -> Any:
        """Get cell value by full address.
        
        Args:
            address: Full cell address like "Sheet1!A1"
            
        Returns:
            The cell value
        """
        cell = self.get_cell(address)
        return cell.value
    
    def add_defined_name(self, name: str, reference: str) -> None:
        """Add a defined name pointing to a cell or range.
        
        Args:
            name: Name of the defined name
            reference: Reference like "Sheet1!A1" or "Sheet1!A1:B2"
        """
        if ':' in reference:
            # Range reference
            sheet_name, range_address = self._parse_full_address(reference)
            worksheet = self.get_worksheet(sheet_name)
            range_obj = worksheet.get_range(range_address)
            self.defined_names[name] = range_obj
        else:
            # Cell reference
            cell = self.get_cell(reference)
            self.defined_names[name] = cell
        
        # Rebuild AST for all formulas since defined names context changed
        self.build_all_formula_ast()
    
    def get_cell_by_name(self, name: str) -> Cell:
        """Get cell by defined name.
        
        Args:
            name: Defined name
            
        Returns:
            The Cell instance
            
        Raises:
            KeyError: If defined name doesn't exist
            TypeError: If defined name doesn't point to a cell
        """
        if name not in self.defined_names:
            raise KeyError(f"Defined name '{name}' not found")
        
        obj = self.defined_names[name]
        if not isinstance(obj, Cell):
            raise TypeError(f"Defined name '{name}' does not point to a cell")
        
        return obj
    
    def get_range_by_name(self, name: str) -> Range:
        """Get range by defined name.
        
        Args:
            name: Defined name
            
        Returns:
            The Range instance
            
        Raises:
            KeyError: If defined name doesn't exist
            TypeError: If defined name doesn't point to a range
        """
        if name not in self.defined_names:
            raise KeyError(f"Defined name '{name}' not found")
        
        obj = self.defined_names[name]
        if not isinstance(obj, Range):
            raise TypeError(f"Defined name '{name}' does not point to a range")
        
        return obj
    
    def to_flat_model(self):
        """Convert hierarchical model to flat model for backward compatibility.
        
        Returns:
            A Model instance with flat dictionary structure
        """
        from ..model import Model
        from ..xltypes import XLCell, XLFormula
        
        flat_model = Model()
        
        # Convert all cells to flat structure
        for sheet_name, worksheet in self.worksheets.items():
            for cell_address, cell in worksheet.cells.items():
                full_address = f"{sheet_name}!{cell_address}"
                
                # Create XLCell for flat model
                xl_cell = XLCell(full_address, cell.value)
                if cell.formula:
                    xl_cell.formula = cell.formula
                
                flat_model.cells[full_address] = xl_cell
                
                # Add formula to formulae dict if present
                if cell.formula:
                    flat_model.formulae[full_address] = cell.formula
        
        # Convert defined names
        for name, obj in self.defined_names.items():
            if isinstance(obj, Cell):
                flat_model.defined_names[name] = flat_model.cells[obj.full_address]
            elif isinstance(obj, Range):
                # Convert range to XLRange
                xl_range = XLRange(obj.full_address, name)
                flat_model.defined_names[name] = xl_range
                flat_model.ranges[obj.full_address] = xl_range
        
        return flat_model
    
    def _parse_full_address(self, address: str) -> tuple[str, str]:
        """Parse full address into sheet name and cell address.
        
        Args:
            address: Full address like "Sheet1!A1" or "Sheet1!A1:B2"
            
        Returns:
            Tuple of (sheet_name, cell_or_range_address)
            
        Raises:
            ValueError: If address format is invalid
        """
        if '!' not in address:
            raise ValueError(f"Invalid address format: '{address}'. Expected 'Sheet!Cell' format")
        
        try:
            from ..range import CellReference
            # Try parsing as cell reference first
            cell_ref = CellReference.parse(address, current_sheet='Sheet1')
            return cell_ref.sheet, cell_ref.address
        except Exception:
            # If that fails, use manual parsing for ranges
            pass
        
        parts = address.split('!', 1)
        if len(parts) != 2:
            raise ValueError(f"Invalid address format: '{address}'. Expected 'Sheet!Cell' format")
        
        sheet_name, cell_address = parts
        
        if not sheet_name or not cell_address:
            raise ValueError(f"Invalid address format: '{address}'. Sheet name and cell address cannot be empty")
        
        return sheet_name, cell_address
    
    def build_all_formula_ast(self):
        """Build AST for all formulas in all worksheets."""
        for worksheet in self.worksheets.values():
            worksheet.build_all_formula_ast()
    
    def __eq__(self, other) -> bool:
        """Compare workbooks for equality."""
        if not isinstance(other, Workbook):
            return False
        
        return (
            self.name == other.name
            and self.worksheets == other.worksheets
            and self.defined_names == other.defined_names
            and self.active_sheet == other.active_sheet
        )
    
    def __repr__(self) -> str:
        """String representation of workbook."""
        sheet_names = list(self.worksheets.keys())
        return f"Workbook(name='{self.name}', sheets={sheet_names}, active='{self.active_sheet}')"