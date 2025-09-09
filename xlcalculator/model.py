import copy
import gzip
import jsonpickle
import logging
import os
from dataclasses import dataclass, field

from . import xltypes, reader, parser, tokenizer


@dataclass
class Model():

    cells: dict = field(
        init=False, default_factory=dict, compare=True, hash=True, repr=True)
    formulae: dict = field(
        init=False, default_factory=dict, compare=True, hash=True, repr=True)
    ranges: dict = field(
        init=False, default_factory=dict, compare=True, hash=True, repr=True)
    defined_names: dict = field(
        init=False, default_factory=dict, compare=True, hash=True, repr=True)

    def set_cell_value(self, address, value):
        """Sets a new value for a specified cell."""
        if address in self.defined_names:
            if isinstance(self.defined_names[address], xltypes.XLCell):
                address = self.defined_names[address].address

        if isinstance(address, str):
            if address in self.cells:
                self.cells[address].value = copy.copy(value)
            else:
                self.cells[address] = xltypes.XLCell(address, copy.copy(value))

        elif isinstance(address, xltypes.XLCell):
            if address.address in self.cells:
                self.cells[address.address].value = value
            else:
                self.cells[address.address] = xltypes.XLCell
                (address.address, value)

        else:
            raise TypeError(
                f"Cannot set the cell value for an address of type "
                f"{address}. XLCell or a string is needed."
            )

    def get_cell_value(self, address):
        if address in self.defined_names:
            if isinstance(self.defined_names[address], xltypes.XLCell):
                address = self.defined_names[address].address

        if isinstance(address, str):
            if address in self.cells:
                return self.cells[address].value
            else:
                logging.debug(
                    "Trying to get value for cell {address} but that cell "
                    "doesn't exist.")
                return 0

        elif isinstance(address, xltypes.XLCell):
            if address.address in self.cells:
                return self.cells[address.address].value
            else:
                logging.debug(
                    "Trying to get value for cell {address.address} but "
                    "that cell doesn't exist")
                return 0

        else:
            raise TypeError(
                f"Cannot set the cell value for an address of type "
                f"{address}. XLCell or a string is needed."
            )

    def persist_to_json_file(self, fname):
        """Writes the state to disk.

        Doesn't write the graph directly, but persist all the things that
        provide the ability to re-create the graph.
        """
        output = {
            'cells': self.cells,
            'defined_names': self.defined_names,
            'formulae': self.formulae,
            'ranges': self.ranges,
        }

        file_open = gzip.GzipFile \
            if os.path.splitext(fname)[-1].lower() in ['.gzip', '.gz'] \
            else open

        with file_open(fname, 'wb') as fp:
            fp.write(jsonpickle.encode(output, keys=True).encode())

    def construct_from_json_file(self, fname, build_code=False):
        """Constructs a graph from a state persisted to disk."""

        file_open = gzip.GzipFile \
            if os.path.splitext(fname)[-1].lower() in ['.gzip', '.gz'] \
            else open

        with file_open(fname, "rb") as fp:
            json_bytes = fp.read()

        data = jsonpickle.decode(
            json_bytes, keys=True,
            classes=(
                xltypes.XLCell, xltypes.XLFormula, xltypes.XLRange,
                tokenizer.f_token
            )
        )
        self.cells = data['cells']

        self.defined_names = data['defined_names']
        self.ranges = data['ranges']
        self.formulae = data['formulae']

        if build_code:
            self.build_code()

    def build_code(self):
        """Define the Python code for all cells in the dict of cells."""

        for cell in self.cells:
            if self.cells[cell].formula is not None:
                defined_names = {
                    name: defn.address
                    for name, defn in self.defined_names.items()}
                self.cells[cell].formula.ast = parser.FormulaParser().parse(
                    self.cells[cell].formula.formula, defined_names)

    def __eq__(self, other):

        cells_comparison = []
        for self_cell in self.cells:
            cells_comparison.append(
                self.cells[self_cell] == other.cells[self_cell])

        defined_names_comparison = []
        for self_defined_names in self.defined_names:
            defined_names_comparison.append(
                self.defined_names[self_defined_names]
                    == other.defined_names[self_defined_names])

        return (
            self.__class__ == other.__class__
            and all(cells_comparison)
            and all(defined_names_comparison)
        )


class ModelCompiler:
    """Excel Workbook Data Model Compiler

    Factory class responsible for taking Microsoft Excel cells and named_range
    and create a model represented by a network graph that can be serialized
    to disk, and executed independently of Excel.
    """

    def __init__(self):
        self.model = Model()

    def read_excel_file(self, file_name):
        archive = reader.Reader(file_name)
        archive.read()
        return archive

    def parse_archive(self, archive, ignore_sheets=[], ignore_hidden=False):
        self.model.cells, self.model.formulae, self.model.ranges = \
            archive.read_cells(ignore_sheets, ignore_hidden)
        self.defined_names = archive.read_defined_names(
            ignore_sheets, ignore_hidden)
        self.build_defined_names()
        self.link_cells_to_defined_names()
        self.build_ranges()

    def read_and_parse_archive(
            self, file_name=None, ignore_sheets=[], ignore_hidden=False,
            build_code=True
    ):
        archive = self.read_excel_file(file_name)
        self.parse_archive(
            archive, ignore_sheets=ignore_sheets, ignore_hidden=ignore_hidden)
        
        # Add default sheet aliases for active sheet
        self._add_default_sheet_aliases(archive)

        if build_code:
            self.model.build_code()

        return self.model
    
    def read_and_parse_dict_hierarchical(
            self, input_dict, default_sheet="Sheet1", build_code=True):
        """Create hierarchical model from dictionary input.
        
        Args:
            input_dict: Dictionary of cell addresses to values
            default_sheet: Default sheet name for cells without sheet prefix
            build_code: Whether to build AST code for formulas
            
        Returns:
            Workbook instance with hierarchical structure
        """
        from .hierarchical_model import Workbook
        
        workbook = Workbook(name="Workbook")
        
        # Group cells by sheet
        sheet_cells = {}
        for item, value in input_dict.items():
            from .range import CellReference
            cell_ref = CellReference.parse(item, current_sheet=default_sheet)
            sheet_name = cell_ref.sheet
            cell_address = cell_ref.address
            
            if sheet_name not in sheet_cells:
                sheet_cells[sheet_name] = {}
            sheet_cells[sheet_name][cell_address] = value
        
        # Create worksheets and populate cells
        for sheet_name, cells in sheet_cells.items():
            worksheet = workbook.add_worksheet(sheet_name)
            
            for cell_address, value in cells.items():
                worksheet.set_cell_value(cell_address, value)
        
        # Build code if requested
        if build_code:
            self._build_hierarchical_code(workbook)
        
        return workbook
    
    def _build_hierarchical_code(self, workbook):
        """Build AST code for formulas in hierarchical model.
        
        Args:
            workbook: Workbook instance to build code for
        """
        # Use the workbook's built-in AST building capability
        workbook.build_all_formula_ast()
    
    def parse_archive_hierarchical(self, archive, ignore_sheets=[], ignore_hidden=False):
        """Parse archive into hierarchical model structure.
        
        Args:
            archive: Excel archive reader
            ignore_sheets: List of sheet names to ignore
            ignore_hidden: Whether to ignore hidden sheets
            
        Returns:
            Workbook instance with hierarchical structure
        """
        from .hierarchical_model import Workbook
        
        # Read cells, formulae, and ranges from archive
        cells, formulae, ranges = archive.read_cells(ignore_sheets, ignore_hidden)
        defined_names = archive.read_defined_names(ignore_sheets, ignore_hidden)
        
        workbook = Workbook(name=getattr(archive, 'filename', 'Workbook'))
        
        # Group cells by sheet
        sheet_cells = {}
        for full_address, cell in cells.items():
            from .range import CellReference
            cell_ref = CellReference.parse(full_address, current_sheet="Sheet1")
            sheet_name = cell_ref.sheet
            cell_address = cell_ref.address
            
            if sheet_name not in sheet_cells:
                sheet_cells[sheet_name] = {}
            sheet_cells[sheet_name][cell_address] = cell
        
        # Create worksheets and populate cells
        for sheet_name, sheet_cell_dict in sheet_cells.items():
            if sheet_name in ignore_sheets:
                continue
                
            worksheet = workbook.add_worksheet(sheet_name)
            
            for cell_address, xl_cell in sheet_cell_dict.items():
                # Create hierarchical cell
                hierarchical_cell = worksheet.get_cell(cell_address)
                hierarchical_cell.value = xl_cell.value
                
                # Handle formula
                if xl_cell.formula:
                    hierarchical_cell.formula = xl_cell.formula
        
        # Handle defined names
        for name, reference in defined_names.items():
            try:
                workbook.add_defined_name(name, reference)
            except Exception as e:
                logging.warning(f"Failed to add defined name '{name}': {e}")
        
        return workbook
    
    def read_and_parse_archive_hierarchical(
            self, file_name=None, ignore_sheets=[], ignore_hidden=False,
            build_code=True
    ):
        """Read Excel file and create hierarchical model.
        
        Args:
            file_name: Path to Excel file
            ignore_sheets: List of sheet names to ignore
            ignore_hidden: Whether to ignore hidden sheets
            build_code: Whether to build AST code for formulas
            
        Returns:
            Workbook instance with hierarchical structure
        """
        archive = self.read_excel_file(file_name)
        workbook = self.parse_archive_hierarchical(
            archive, ignore_sheets=ignore_sheets, ignore_hidden=ignore_hidden)
        
        if build_code:
            self._build_hierarchical_code(workbook)
        
        return workbook
    
    def _add_default_sheet_aliases(self, archive):
        """Add aliases for active sheet cells to support relative references.
        
        This allows tests to use 'A1' instead of 'Sheet!A1' for the active sheet.
        """
        if hasattr(archive, 'book') and archive.book.active:
            active_sheet_name = archive.book.active.title
            
            # Create aliases for active sheet cells
            active_cells = {k: v for k, v in self.model.cells.items() 
                          if k.startswith(f'{active_sheet_name}!')}
            
            for full_addr, cell in active_cells.items():
                short_addr = full_addr.replace(f'{active_sheet_name}!', '')
                if short_addr not in self.model.cells:  # Don't overwrite existing cells
                    self.model.cells[short_addr] = cell
                    
            # Also create aliases for formulae
            active_formulae = {k: v for k, v in self.model.formulae.items() 
                             if k.startswith(f'{active_sheet_name}!')}
            
            for full_addr, formula in active_formulae.items():
                short_addr = full_addr.replace(f'{active_sheet_name}!', '')
                if short_addr not in self.model.formulae:  # Don't overwrite existing formulae
                    self.model.formulae[short_addr] = formula

    def read_and_parse_dict(
            self, input_dict, default_sheet="Sheet1", build_code=True):
        for item in input_dict:
            if "!" in item:
                cell_address = item
            else:
                cell_address = "{}!{}".format(default_sheet, item)

            if (
                    not isinstance(input_dict[item], (float, int))
                    and input_dict[item][0] == '='
            ):
                formula = xltypes.XLFormula(
                    input_dict[item],
                    sheet_name=default_sheet
                )
                cell = xltypes.XLCell(
                    cell_address, None,
                    formula=formula)
                self.model.cells[cell_address] = cell
                self.model.formulae[cell_address] = cell.formula

            else:
                self.model.cells[cell_address] = xltypes.XLCell(
                    cell_address, input_dict[item])

        self.build_ranges(default_sheet=default_sheet)

        if build_code:
            self.model.build_code()

        return self.model
    
    def read_and_parse_dict_hierarchical(
            self, input_dict, default_sheet="Sheet1", build_code=True):
        """Create hierarchical model from dictionary input.
        
        Args:
            input_dict: Dictionary of cell addresses to values
            default_sheet: Default sheet name for cells without sheet prefix
            build_code: Whether to build AST code for formulas
            
        Returns:
            Workbook instance with hierarchical structure
        """
        from .hierarchical_model import Workbook
        
        workbook = Workbook(name="Workbook")
        
        # Group cells by sheet
        sheet_cells = {}
        for item, value in input_dict.items():
            from .range import CellReference
            cell_ref = CellReference.parse(item, current_sheet=default_sheet)
            sheet_name = cell_ref.sheet
            cell_address = cell_ref.address
            
            if sheet_name not in sheet_cells:
                sheet_cells[sheet_name] = {}
            sheet_cells[sheet_name][cell_address] = value
        
        # Create worksheets and populate cells
        for sheet_name, cells in sheet_cells.items():
            worksheet = workbook.add_worksheet(sheet_name)
            
            for cell_address, value in cells.items():
                worksheet.set_cell_value(cell_address, value)
        
        # Build code if requested
        if build_code:
            self._build_hierarchical_code(workbook)
        
        return workbook
    
    def _build_hierarchical_code(self, workbook):
        """Build AST code for formulas in hierarchical model.
        
        Args:
            workbook: Workbook instance to build code for
        """
        # Use the workbook's built-in AST building capability
        workbook.build_all_formula_ast()
    
    def parse_archive_hierarchical(self, archive, ignore_sheets=[], ignore_hidden=False):
        """Parse archive into hierarchical model structure.
        
        Args:
            archive: Excel archive reader
            ignore_sheets: List of sheet names to ignore
            ignore_hidden: Whether to ignore hidden sheets
            
        Returns:
            Workbook instance with hierarchical structure
        """
        from .hierarchical_model import Workbook
        
        # Read cells, formulae, and ranges from archive
        cells, formulae, ranges = archive.read_cells(ignore_sheets, ignore_hidden)
        defined_names = archive.read_defined_names(ignore_sheets, ignore_hidden)
        
        workbook = Workbook(name=getattr(archive, 'filename', 'Workbook'))
        
        # Group cells by sheet
        sheet_cells = {}
        for full_address, cell in cells.items():
            from .range import CellReference
            cell_ref = CellReference.parse(full_address, current_sheet="Sheet1")
            sheet_name = cell_ref.sheet
            cell_address = cell_ref.address
            
            if sheet_name not in sheet_cells:
                sheet_cells[sheet_name] = {}
            sheet_cells[sheet_name][cell_address] = cell
        
        # Create worksheets and populate cells
        for sheet_name, sheet_cell_dict in sheet_cells.items():
            if sheet_name in ignore_sheets:
                continue
                
            worksheet = workbook.add_worksheet(sheet_name)
            
            for cell_address, xl_cell in sheet_cell_dict.items():
                # Create hierarchical cell
                hierarchical_cell = worksheet.get_cell(cell_address)
                hierarchical_cell.value = xl_cell.value
                
                # Handle formula
                if xl_cell.formula:
                    hierarchical_cell.formula = xl_cell.formula
        
        # Handle defined names
        for name, reference in defined_names.items():
            try:
                workbook.add_defined_name(name, reference)
            except Exception as e:
                logging.warning(f"Failed to add defined name '{name}': {e}")
        
        return workbook
    
    def read_and_parse_archive_hierarchical(
            self, file_name=None, ignore_sheets=[], ignore_hidden=False,
            build_code=True
    ):
        """Read Excel file and create hierarchical model.
        
        Args:
            file_name: Path to Excel file
            ignore_sheets: List of sheet names to ignore
            ignore_hidden: Whether to ignore hidden sheets
            build_code: Whether to build AST code for formulas
            
        Returns:
            Workbook instance with hierarchical structure
        """
        archive = self.read_excel_file(file_name)
        workbook = self.parse_archive_hierarchical(
            archive, ignore_sheets=ignore_sheets, ignore_hidden=ignore_hidden)
        
        if build_code:
            self._build_hierarchical_code(workbook)
        
        return workbook

    def build_defined_names(self):
        """Add defined ranges to model."""
        for name in self.defined_names:
            cell_address = self.defined_names[name]
            cell_address = cell_address.replace('$', '')

            # a cell has an address like; Sheet1!A1
            if ':' not in cell_address:
                if cell_address not in self.model.cells:
                    logging.warning(
                        f"Defined name {name} refers to empty cell "
                        f"{cell_address}. Is not being loaded.")
                    continue

                else:
                    if self.model.cells[cell_address] is not None:
                        self.model.defined_names[name] = \
                            self.model.cells[cell_address]

            else:
                from .lazy_loading import create_excel_compliant_lazy_range
                self.model.defined_names[name] = create_excel_compliant_lazy_range(
                    cell_address, self.model, name=name)
                self.model.ranges[cell_address] = \
                    self.model.defined_names[name]

            if (
                    cell_address in self.model.formulae
                    and name not in self.model.formulae
            ):
                self.model.formulae[name] = \
                    self.model.cells[cell_address].formula

    def link_cells_to_defined_names(self):
        for name in self.model.defined_names:
            defn = self.model.defined_names[name]

            if isinstance(defn, xltypes.XLCell):
                self.model.cells[defn.address].defined_names.append(name)

            elif isinstance(defn, xltypes.XLRange):
                if any(isinstance(el, list) for el in defn.cells):
                    for column in defn.cells:
                        for row_address in column:
                            self.model.cells[row_address].defined_names.append(
                                name)
                else:
                    # programmer error
                    message = "This isn't a dim2 array. {}".format(name)
                    logging.error(message)
                    raise Exception(message)
            else:
                message = (
                    f"Trying to link cells for {name}, but got unkown "
                    f"type {type(defn)}"
                )
                logging.error(message)
                raise ValueError(message)

    def build_ranges(self, default_sheet=None):
        from .lazy_loading import create_excel_compliant_lazy_range, is_full_range
        
        for formula in self.model.formulae:
            associated_cells = set()
            for range in self.model.formulae[formula].terms:
                if ":" in range:
                    if "!" not in range:
                        range = "{}!{}".format(default_sheet, range)
                    
                    # Use Excel-compliant lazy loading for full ranges
                    if is_full_range(range):
                        logging.info(f"Using Excel-compliant lazy loading for full range: {range}")
                        self.model.ranges[range] = create_excel_compliant_lazy_range(range, self.model, range)
                        # For lazy ranges, add actual cells based on Excel data
                        lazy_range = self.model.ranges[range]
                        for row in lazy_range.cells:
                            for cell in row:
                                associated_cells.add(cell)
                    else:
                        # Use standard XLRange for normal ranges
                        self.model.ranges[range] = xltypes.XLRange(range, range)
                        associated_cells.update([
                            cell
                            for row in self.model.ranges[range].cells
                                for cell in row  # noqa: E131
                        ])
                else:
                    associated_cells.add(range)

                if range in self.model.ranges:
                    for row in self.model.ranges[range].cells:
                        for cell_address in row:
                            if cell_address not in self.model.cells.keys():
                                self.model.cells[cell_address] = \
                                    xltypes.XLCell(cell_address, '')

            if formula in self.model.cells:
                self.model.cells[formula].formula.associated_cells = \
                    associated_cells

            if formula in self.model.defined_names:
                self.model.defined_names[formula].formula.associated_cells = \
                    associated_cells

            self.model.formulae[formula].associated_cells = associated_cells

    @staticmethod
    def extract(model, focus):
        extracted_model = Model()

        for address in focus:
            if isinstance(address, str) and address in model.cells:
                extracted_model.cells[address] = copy.deepcopy(
                    model.cells[address])

            elif isinstance(address, str) and address in model.defined_names:

                extracted_model.defined_names[address] = defn = copy.deepcopy(
                    model.defined_names[address])

                if isinstance(defn, xltypes.XLCell):
                    extracted_model.cells[defn.address] = copy.deepcopy(
                        model.cells[defn.address])

                elif isinstance(defn, xltypes.XLRange):
                    for row in defn.cells:
                        for column in row:
                            extracted_model.cells[column] = copy.deepcopy(
                                model.cells[column])

        terms_to_copy = []
        for addr, cell in extracted_model.cells.items():
            if cell.formula is not None:
                for term in cell.formula.terms:
                    if (term in extracted_model.cells
                            and cell.formula != model.cells[addr].formula):
                        cell.formula = copy.deepcopy(model.cells[addr].formula)

                    elif term not in extracted_model.cells:
                        terms_to_copy.append(term)

        for term in terms_to_copy:
            extracted_model.cells[term] = copy.deepcopy(model.cells[term])

        extracted_model.build_code()

        return extracted_model
