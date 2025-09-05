from .. import testing


class XLookupTest(testing.FunctionalTestCase):
    filename = "XLOOKUP.xlsx"

    def test_basic_exact_match(self):
        """Test basic exact match functionality."""
        # Since XLOOKUP is newly implemented, we test against expected values
        # rather than Excel-calculated values (which don't exist yet)
        value = self.evaluator.evaluate('Sheet1!A8')
        expected = 10  # XLOOKUP("Apple", A2:A5, B2:B5) should return 10
        self.assertEqual(expected, value)

    def test_with_if_not_found(self):
        """Test XLOOKUP with custom if_not_found value."""
        value = self.evaluator.evaluate('Sheet1!A9')
        expected = "Not Found"  # XLOOKUP("Orange", A2:A5, B2:B5, "Not Found")
        self.assertEqual(expected, value)

    def test_another_exact_match(self):
        """Test another exact match case."""
        value = self.evaluator.evaluate('Sheet1!A10')
        expected = 30  # XLOOKUP("Cherry", A2:A5, B2:B5) should return 30
        self.assertEqual(expected, value)

    def test_approximate_match_next_smallest(self):
        """Test approximate match with next smallest."""
        value = self.evaluator.evaluate('Sheet1!A12')
        expected = "D"  # XLOOKUP(25, D2:D6, E2:E6, , -1) -> next smallest is 20 -> "D"
        self.assertEqual(expected, value)

    def test_approximate_match_next_largest(self):
        """Test approximate match with next largest."""
        value = self.evaluator.evaluate('Sheet1!A13')
        expected = "D"  # XLOOKUP(15, D2:D6, E2:E6, , 1) -> next largest is 20 -> "D"
        self.assertEqual(expected, value)

    def test_approximate_match_exact(self):
        """Test exact match in approximate mode."""
        value = self.evaluator.evaluate('Sheet1!A14')
        expected = "C"  # XLOOKUP(30, D2:D6, E2:E6, , 0) -> exact match 30 -> "C"
        self.assertEqual(expected, value)

    def test_wildcard_match_asterisk(self):
        """Test wildcard matching with asterisk."""
        value = self.evaluator.evaluate('Sheet1!A16')
        expected = 10  # XLOOKUP("App*", A2:A5, B2:B5, , 2) -> matches "Apple" -> 10
        self.assertEqual(expected, value)

    def test_wildcard_match_question(self):
        """Test wildcard matching with question mark."""
        value = self.evaluator.evaluate('Sheet1!A17')
        expected = 20  # XLOOKUP("Ban?na", A2:A5, B2:B5, , 2) -> matches "Banana" -> 20
        self.assertEqual(expected, value)

    def test_wildcard_match_prefix(self):
        """Test wildcard matching with prefix asterisk."""
        value = self.evaluator.evaluate('Sheet1!A18')
        expected = 30  # XLOOKUP("*erry", A2:A5, B2:B5, , 2) -> matches "Cherry" -> 30
        self.assertEqual(expected, value)

    def test_forward_search(self):
        """Test forward search (first occurrence)."""
        value = self.evaluator.evaluate('Sheet1!A20')
        expected = 1  # XLOOKUP("A", G2:G6, H2:H6, , 0, 1) -> first "A" at position 1
        self.assertEqual(expected, value)

    def test_reverse_search(self):
        """Test reverse search (last occurrence)."""
        value = self.evaluator.evaluate('Sheet1!A21')
        expected = 5  # XLOOKUP("A", G2:G6, H2:H6, , 0, -1) -> last "A" at position 5
        self.assertEqual(expected, value)

    def test_binary_search_1(self):
        """Test binary search functionality."""
        value = self.evaluator.evaluate('Sheet1!A23')
        expected = "C"  # XLOOKUP(30, D2:D6, E2:E6, , 0, 2) -> binary search for 30 -> "C"
        self.assertEqual(expected, value)

    def test_binary_search_2(self):
        """Test binary search functionality."""
        value = self.evaluator.evaluate('Sheet1!A24')
        expected = "D"  # XLOOKUP(20, D2:D6, E2:E6, , 0, 2) -> binary search for 20 -> "D"
        self.assertEqual(expected, value)

    def test_not_found_error(self):
        """Test not found error case."""
        from xlcalculator.xlfunctions import xlerrors
        value = self.evaluator.evaluate('Sheet1!A26')
        # XLOOKUP("Grape", A2:A5, B2:B5) -> not found, should return #N/A error
        self.assertIsInstance(value, xlerrors.NaExcelError)

    def test_horizontal_arrays(self):
        """Test XLOOKUP with horizontal arrays."""
        value = self.evaluator.evaluate('Sheet1!A30')
        expected = 200  # XLOOKUP("Banana", A28:C28, A29:C29) -> "Banana" at position 2 -> 200
        self.assertEqual(expected, value)