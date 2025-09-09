import sys
from functools import lru_cache

from xlcalculator.xlfunctions import xl, func_xltypes

from . import ast_nodes, xltypes


class EvaluatorContext(ast_nodes.EvalContext):

    def __init__(self, evaluator, ref, formula_sheet=None):
        super().__init__(evaluator.namespace, ref, formula_sheet=formula_sheet)
        self.evaluator = evaluator

    @property
    def cells(self):
        # Support both flat model and hierarchical model
        if hasattr(self.evaluator.model, 'cells'):
            return self.evaluator.model.cells
        else:
            # Hierarchical model - create flat view
            return self.evaluator._get_flat_cells_view()

    @property
    def ranges(self):
        # Support both flat model and hierarchical model
        if hasattr(self.evaluator.model, 'ranges'):
            return self.evaluator.model.ranges
        else:
            # Hierarchical model - create flat view
            return self.evaluator._get_flat_ranges_view()

    @lru_cache(maxsize=None)
    def eval_cell(self, addr):
        # Check for a cycle.
        if addr in self.seen:
            raise RuntimeError(
                f'Cycle detected for {addr}:\n- ' + '\n- '.join(self.seen))
        self.seen.append(addr)

        return self.evaluator.evaluate(addr, None)


class Evaluator:
    """Traverses and evaluates a given model."""

    def __init__(self, model, namespace=None):
        self.model = model
        self.namespace = namespace \
            if namespace is not None else xl.FUNCTIONS.copy()
        self.cache_count = 0
        self._lazy_manager = None
        self._is_hierarchical = self._detect_hierarchical_model()

    def _detect_hierarchical_model(self):
        """Detect if model is hierarchical (Workbook) or flat (Model)."""
        return hasattr(self.model, 'worksheets')
    
    def _get_flat_cells_view(self):
        """Create flat cells view from hierarchical model."""
        if not self._is_hierarchical:
            return {}
        
        flat_cells = {}
        for sheet_name, worksheet in self.model.worksheets.items():
            for cell_address, cell in worksheet.cells.items():
                full_address = f"{sheet_name}!{cell_address}"
                # Create XLCell-like object for compatibility
                xl_cell = xltypes.XLCell(full_address, cell.value)
                xl_cell.formula = cell.formula
                flat_cells[full_address] = xl_cell
        
        return flat_cells
    
    def _get_flat_ranges_view(self):
        """Create flat ranges view from hierarchical model."""
        if not self._is_hierarchical:
            return {}
        
        flat_ranges = {}
        for sheet_name, worksheet in self.model.worksheets.items():
            for range_address, range_obj in worksheet.ranges.items():
                full_address = f"{sheet_name}!{range_address}"
                # Create XLRange-like object for compatibility
                xl_range = xltypes.XLRange(full_address, range_obj.address)
                flat_ranges[full_address] = xl_range
        
        return flat_ranges
    
    def _get_hierarchical_cell(self, addr):
        """Get cell from hierarchical model by full address."""
        try:
            from .range import CellReference
            cell_ref = CellReference.parse(addr, current_sheet='Sheet1')
            
            if cell_ref.sheet in self.model.worksheets:
                worksheet = self.model.worksheets[cell_ref.sheet]
                if cell_ref.address in worksheet.cells:
                    return worksheet.cells[cell_ref.address]
                else:
                    # Create empty cell if it doesn't exist
                    return worksheet.get_cell(cell_ref.address)
            return None
        except Exception:
            return None
    
    def _get_context(self, ref, formula_sheet=None):
        return EvaluatorContext(self, ref, formula_sheet)

    def resolve_names(self, addr):
        # Although defined names have been resolved in Model.create_node()
        # we need to attempt to resolve defined names as we might have been
        # given one in argument addr.
        if addr not in self.model.defined_names:
            return addr

        defn = self.model.defined_names[addr]

        if isinstance(defn, xltypes.XLCell):
            return defn.address

        if isinstance(defn, xltypes.XLRange):
            raise ValueError(
                f"I can't resolve {addr} to a cell. It's a "
                f"range and they aren't supported yet.")

        if isinstance(defn, xltypes.XLFormula):
            raise ValueError(
                f"I can't resolve {addr} to a cell. It's a "
                f"formula and they aren't supported as a cell "
                f"reference.")

    def evaluate(self, addr, context=None):
        # 1. Resolve the address to a cell.
        addr = self.resolve_names(addr)
        
        # Handle hierarchical model
        if self._is_hierarchical:
            cell = self._get_hierarchical_cell(addr)
            if cell is None:
                return func_xltypes.BLANK
        else:
            # Handle flat model
            if addr not in self.model.cells:
                # Blank cell that has no stored value in the model.
                return func_xltypes.BLANK
            cell = self.model.cells[addr]

        # 2. If there is no formula, we simply return the cell value.
        if (cell.formula is None or cell.formula.evaluate is False):
            if self._is_hierarchical:
                return func_xltypes.ExcelType.cast_from_native(cell.value)
            else:
                return func_xltypes.ExcelType.cast_from_native(
                    self.model.cells[addr].value)

        # 3. Prepare the execution environment and evaluate the formula.
        #    Extract formula sheet context for proper Excel behavior
        formula_sheet = cell.formula.sheet_name if cell.formula else None
        context = context if context is not None else self._get_context(addr, formula_sheet)
        
        # Context injection now handles evaluator access for dynamic range functions
        
        try:
            value = cell.formula.ast.eval(context)
        except Exception as err:
            # Handle Excel errors as return values, not exceptions
            from xlcalculator.xlfunctions import xlerrors
            if isinstance(err, (xlerrors.RefExcelError, xlerrors.ValueExcelError, 
                              xlerrors.NameExcelError, xlerrors.NumExcelError, 
                              xlerrors.NaExcelError, xlerrors.DivZeroExcelError,
                              xlerrors.NullExcelError)):
                value = err
            else:
                raise RuntimeError(
                    f"Problem evaluating cell {addr} formula "
                    f"{cell.formula.formula}: {repr(err)}"
                ).with_traceback(sys.exc_info()[2])

        # 4. Update the cell value.
        #    Note for later: If an array is returned, we should distribute the
        #    values to the respective cell (known as spilling).
        cell.value = value
        cell.need_update = False

        return value

    def set_cell_value(self, address, value):
        """Sets the value of a cell in the model."""
        self.model.set_cell_value(address, value)

    def get_cell_value(self, address):
        """Gets the value of a cell in the model."""
        return self.model.get_cell_value(address)
    
    def get_range_values(self, range_ref):
        """Gets the values of a range in the model."""
        # Parse range reference like "Sheet1!B2:C3" or "B2:C3"
        if ':' not in range_ref:
            # Single cell
            return [[self.get_cell_value(range_ref)]]
        
        # Handle sheet prefix
        sheet_prefix = ""
        if '!' in range_ref:
            sheet_prefix, range_part = range_ref.split('!', 1)
            sheet_prefix += '!'
        else:
            range_part = range_ref
            sheet_prefix = 'Sheet1!'
        
        # Parse range part
        start_ref, end_ref = range_part.split(':')
        
        # Simple parsing for basic ranges (A1:B2 format)
        # Extract column and row from start reference
        start_col_letter = ''.join(c for c in start_ref if c.isalpha())
        start_row = int(''.join(c for c in start_ref if c.isdigit()))
        start_col = ord(start_col_letter) - ord('A') + 1
        
        # Extract column and row from end reference  
        end_col_letter = ''.join(c for c in end_ref if c.isalpha())
        end_row = int(''.join(c for c in end_ref if c.isdigit()))
        end_col = ord(end_col_letter) - ord('A') + 1
        
        values = []
        for row in range(start_row, end_row + 1):
            row_values = []
            for col in range(start_col, end_col + 1):
                col_letter = chr(ord('A') + col - 1)
                cell_ref = f'{sheet_prefix}{col_letter}{row}'
                value = self.get_cell_value(cell_ref)
                row_values.append(value)
            values.append(row_values)
        
        return values
    
    def clear_context_cache(self):
        """Clear context cache to free memory after evaluation cycles."""
        from .context import clear_context_cache
        clear_context_cache()
    
    def enable_lazy_loading(self):
        """Enable lazy loading for this evaluator."""
        from .lazy_loading import patch_evaluator_with_lazy_loading
        self._lazy_manager = patch_evaluator_with_lazy_loading(self)
        return self._lazy_manager


def create_evaluator_with_lazy_loading(excel_file_path, ignore_sheets=[], ignore_hidden=False):
    """Factory function to create an evaluator with Excel-compliant lazy loading enabled."""
    from .model import ModelCompiler
    
    compiler = ModelCompiler()
    model = compiler.read_and_parse_archive(
        excel_file_path, 
        ignore_sheets=ignore_sheets, 
        ignore_hidden=ignore_hidden
    )
    
    evaluator = Evaluator(model)
    evaluator.enable_lazy_loading()
    
    return evaluator
