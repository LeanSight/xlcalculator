import jsonpickle
import unittest

from xlcalculator import reader, xltypes, tokenizer
from . import testing


class ReaderTest(unittest.TestCase):

    def setUp(self):
        infile = open(testing.get_resource("reader.json"), "rb")
        json_bytes = infile.read()
        infile.close()
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

    def test_read_cells(self):
        archive = reader.Reader(testing.get_resource("reader.xlsm"))
        archive.read()
        cells, formulae, ranges = \
            archive.read_cells(ignore_sheets=['Eleventh'])

        self.assertEqual(sorted(self.cells.keys()), sorted(cells.keys()))

    def test_read_formulae(self):
        archive = reader.Reader(testing.get_resource("reader.xlsm"))
        archive.read()
        cells, formulae, ranges = \
            archive.read_cells(ignore_sheets=['Eleventh'])

        self.assertEqual(sorted(self.formulae.keys()), sorted(formulae.keys()))

    def test_read_defined_names(self):
        archive = reader.Reader(testing.get_resource("reader.xlsm"))
        archive.read()
        defined_names = archive.read_defined_names()

        # Test that defined names match exactly
        self.assertEqual(
            sorted(defined_names.keys()),
            sorted(self.defined_names.keys())
        )
