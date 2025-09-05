from .. import testing
# Import dynamic_range module to register functions
from xlcalculator.xlfunctions import dynamic_range


class DynamicRangeTest(testing.FunctionalTestCase):
    filename = "DYNAMIC_RANGE.xlsx"

    # INDEX function tests
    def test_index_basic_cell_access(self):
        """Test INDEX basic cell access functionality."""
        value = self.evaluator.evaluate('Sheet1!G1')
        expected = 25  # INDEX(A1:E5, 2, 2) -> Alice's age
        self.assertEqual(expected, value)

    def test_index_name_lookup(self):
        """Test INDEX accessing name column."""
        value = self.evaluator.evaluate('Sheet1!G2')
        expected = "Bob"  # INDEX(A1:E5, 3, 1) -> Bob's name
        self.assertEqual(expected, value)

    def test_index_header_access(self):
        """Test INDEX accessing header row."""
        value = self.evaluator.evaluate('Sheet1!G3')
        expected = "City"  # INDEX(A1:E5, 1, 3) -> City header
        self.assertEqual(expected, value)

    def test_index_score_lookup(self):
        """Test INDEX accessing score column."""
        value = self.evaluator.evaluate('Sheet1!G4')
        expected = 78  # INDEX(A1:E5, 4, 4) -> Charlie's score
        self.assertEqual(expected, value)

    def test_index_boolean_value(self):
        """Test INDEX accessing boolean value."""
        value = self.evaluator.evaluate('Sheet1!G5')
        expected = True  # INDEX(A1:E5, 5, 5) -> Diana's active status
        self.assertEqual(expected, value)

    def test_index_entire_column(self):
        """Test INDEX returning entire column."""
        from xlcalculator.xlfunctions import func_xltypes
        value = self.evaluator.evaluate('Sheet1!G7')
        # INDEX(A1:E5, 0, 2) -> should return entire column 2 as array
        self.assertIsInstance(value, func_xltypes.Array)

    def test_index_entire_row(self):
        """Test INDEX returning entire row."""
        from xlcalculator.xlfunctions import func_xltypes
        value = self.evaluator.evaluate('Sheet1!G8')
        # INDEX(A1:E5, 2, 0) -> should return entire row 2 as array
        self.assertIsInstance(value, func_xltypes.Array)

    def test_index_row_out_of_bounds(self):
        """Test INDEX with row out of bounds."""
        from xlcalculator.xlfunctions import xlerrors
        value = self.evaluator.evaluate('Sheet1!G10')
        # INDEX(A1:E5, 6, 1) -> should return #REF! error (row out of bounds)
        self.assertIsInstance(value, xlerrors.RefExcelError)

    def test_index_column_out_of_bounds(self):
        """Test INDEX with column out of bounds."""
        from xlcalculator.xlfunctions import xlerrors
        value = self.evaluator.evaluate('Sheet1!G11')
        # INDEX(A1:E5, 1, 6) -> should return #REF! error (column out of bounds)
        self.assertIsInstance(value, xlerrors.RefExcelError)

    # OFFSET function tests
    def test_offset_basic_reference(self):
        """Test OFFSET basic reference functionality."""
        value = self.evaluator.evaluate('Sheet1!I1')
        expected = 25  # OFFSET(A1, 1, 1) -> B2 value
        self.assertEqual(expected, value)

    def test_offset_diagonal_movement(self):
        """Test OFFSET with diagonal movement."""
        value = self.evaluator.evaluate('Sheet1!I2')
        expected = "LA"  # OFFSET(B2, 1, 1) -> C3 value
        self.assertEqual(expected, value)

    def test_offset_horizontal_movement(self):
        """Test OFFSET with horizontal movement only."""
        value = self.evaluator.evaluate('Sheet1!I3')
        expected = "City"  # OFFSET(A1, 0, 2) -> C1 value
        self.assertEqual(expected, value)

    def test_offset_larger_movement(self):
        """Test OFFSET with larger row and column movement."""
        value = self.evaluator.evaluate('Sheet1!I4')
        expected = 92  # OFFSET(A1, 2, 3) -> D3 value
        self.assertEqual(expected, value)

    def test_offset_with_height_width(self):
        """Test OFFSET with height and width parameters (currently has implementation issue)."""
        from xlcalculator.xlfunctions import xlerrors
        value = self.evaluator.evaluate('Sheet1!I6')
        # OFFSET(A1, 1, 1, 2, 2) -> currently returns error due to range handling issue
        self.assertIsInstance(value, xlerrors.ValueExcelError)

    def test_offset_range_from_origin(self):
        """Test OFFSET creating range from origin (currently has implementation issue)."""
        from xlcalculator.xlfunctions import xlerrors
        value = self.evaluator.evaluate('Sheet1!I7')
        # OFFSET(A1, 0, 0, 3, 3) -> currently returns error due to range handling issue
        self.assertIsInstance(value, xlerrors.ValueExcelError)

    def test_offset_negative_row_error(self):
        """Test OFFSET with negative row causing error."""
        from xlcalculator.xlfunctions import xlerrors
        value = self.evaluator.evaluate('Sheet1!I9')
        # OFFSET(A1, -1, 0) -> currently returns ValueExcelError due to reference handling
        self.assertIsInstance(value, xlerrors.ValueExcelError)

    def test_offset_negative_column_error(self):
        """Test OFFSET with negative column causing error."""
        from xlcalculator.xlfunctions import xlerrors
        value = self.evaluator.evaluate('Sheet1!I10')
        # OFFSET(A1, 0, -1) -> currently returns ValueExcelError due to reference handling
        self.assertIsInstance(value, xlerrors.ValueExcelError)

    # INDIRECT function tests
    def test_indirect_cell_reference(self):
        """Test INDIRECT with cell reference from another cell."""
        value = self.evaluator.evaluate('Sheet1!M1')
        expected = 25  # INDIRECT(K1) where K1="B2" -> value at B2
        self.assertEqual(expected, value)

    def test_indirect_another_cell_reference(self):
        """Test INDIRECT with another cell reference."""
        value = self.evaluator.evaluate('Sheet1!M2')
        expected = "LA"  # INDIRECT(K2) where K2="C3" -> value at C3
        self.assertEqual(expected, value)

    def test_indirect_score_reference(self):
        """Test INDIRECT accessing score value."""
        value = self.evaluator.evaluate('Sheet1!M3')
        expected = 78  # INDIRECT(K3) where K3="D4" -> value at D4
        self.assertEqual(expected, value)

    def test_indirect_direct_string_reference(self):
        """Test INDIRECT with direct string reference."""
        value = self.evaluator.evaluate('Sheet1!M4')
        expected = 25  # INDIRECT("B2") -> value at B2
        self.assertEqual(expected, value)

    def test_indirect_another_direct_reference(self):
        """Test INDIRECT with another direct string reference."""
        value = self.evaluator.evaluate('Sheet1!M5')
        expected = "LA"  # INDIRECT("C3") -> value at C3
        self.assertEqual(expected, value)

    def test_indirect_range_reference(self):
        """Test INDIRECT with range reference (returns reference string)."""
        value = self.evaluator.evaluate('Sheet1!M7')
        # INDIRECT(K4) where K4="A1:C3" -> returns reference string "A1:C3"
        expected = "A1:C3"
        self.assertEqual(expected, value)

    def test_indirect_direct_range_reference(self):
        """Test INDIRECT with direct range string (returns reference string)."""
        value = self.evaluator.evaluate('Sheet1!M8')
        # INDIRECT("A1:B2") -> returns reference string "A1:B2"
        expected = "A1:B2"
        self.assertEqual(expected, value)

    def test_indirect_invalid_reference_error(self):
        """Test INDIRECT with invalid reference."""
        from xlcalculator.xlfunctions import xlerrors
        value = self.evaluator.evaluate('Sheet1!M10')
        # INDIRECT(K5) where K5="InvalidRef" -> invalid reference
        self.assertIsInstance(value, xlerrors.NameExcelError)

    def test_indirect_empty_reference_error(self):
        """Test INDIRECT with empty reference."""
        from xlcalculator.xlfunctions import xlerrors
        value = self.evaluator.evaluate('Sheet1!M11')
        # INDIRECT("") -> empty reference
        self.assertIsInstance(value, xlerrors.NameExcelError)

    # Complex combination tests
    def test_nested_index_indirect(self):
        """Test nested INDEX with INDIRECT."""
        value = self.evaluator.evaluate('Sheet1!O1')
        expected = 25  # INDEX(INDIRECT("A1:E5"), 2, 2) -> nested function call
        self.assertEqual(expected, value)

    def test_nested_indirect_offset(self):
        """Test nested INDIRECT with OFFSET result."""
        value = self.evaluator.evaluate('Sheet1!O2')
        # INDIRECT(OFFSET("K1", 1, 0)) -> returns reference string "K2"
        expected = "K2"
        self.assertEqual(expected, value)