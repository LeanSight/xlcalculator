from .. import testing


class InformationComprehensiveTest(testing.FunctionalTestCase):
    filename = "INFORMATION.xlsx"

    def test_isnumber_with_number(self):
        """Test ISNUMBER with numeric value."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B1')
        value = self.evaluator.evaluate('Sheet1!B1')
        self.assertEqual(excel_value, value)

    def test_isnumber_with_text(self):
        """Test ISNUMBER with text value."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B2')
        value = self.evaluator.evaluate('Sheet1!B2')
        self.assertEqual(excel_value, value)

    def test_isnumber_with_boolean(self):
        """Test ISNUMBER with boolean value."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B3')
        value = self.evaluator.evaluate('Sheet1!B3')
        self.assertEqual(excel_value, value)

    def test_isnumber_with_blank(self):
        """Test ISNUMBER with blank cell."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B4')
        value = self.evaluator.evaluate('Sheet1!B4')
        self.assertEqual(excel_value, value)

    def test_isnumber_direct(self):
        """Test ISNUMBER with direct number."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B5')
        value = self.evaluator.evaluate('Sheet1!B5')
        self.assertEqual(excel_value, value)

    def test_istext_with_number(self):
        """Test ISTEXT with numeric value."""
        excel_value = self.evaluator.get_cell_value('Sheet1!C1')
        value = self.evaluator.evaluate('Sheet1!C1')
        self.assertEqual(excel_value, value)

    def test_istext_with_text(self):
        """Test ISTEXT with text value."""
        excel_value = self.evaluator.get_cell_value('Sheet1!C2')
        value = self.evaluator.evaluate('Sheet1!C2')
        self.assertEqual(excel_value, value)

    def test_istext_with_boolean(self):
        """Test ISTEXT with boolean value."""
        excel_value = self.evaluator.get_cell_value('Sheet1!C3')
        value = self.evaluator.evaluate('Sheet1!C3')
        self.assertEqual(excel_value, value)

    def test_istext_with_blank(self):
        """Test ISTEXT with blank cell."""
        excel_value = self.evaluator.get_cell_value('Sheet1!C4')
        value = self.evaluator.evaluate('Sheet1!C4')
        self.assertEqual(excel_value, value)

    def test_istext_direct(self):
        """Test ISTEXT with direct text."""
        excel_value = self.evaluator.get_cell_value('Sheet1!C5')
        value = self.evaluator.evaluate('Sheet1!C5')
        self.assertEqual(excel_value, value)

    def test_isblank_with_number(self):
        """Test ISBLANK with numeric value."""
        excel_value = self.evaluator.get_cell_value('Sheet1!D1')
        value = self.evaluator.evaluate('Sheet1!D1')
        self.assertEqual(excel_value, value)

    def test_isblank_with_text(self):
        """Test ISBLANK with text value."""
        excel_value = self.evaluator.get_cell_value('Sheet1!D2')
        value = self.evaluator.evaluate('Sheet1!D2')
        self.assertEqual(excel_value, value)

    def test_isblank_with_blank(self):
        """Test ISBLANK with blank cell."""
        excel_value = self.evaluator.get_cell_value('Sheet1!D4')
        value = self.evaluator.evaluate('Sheet1!D4')
        self.assertEqual(excel_value, value)

    def test_isblank_empty_string(self):
        """Test ISBLANK with empty string."""
        excel_value = self.evaluator.get_cell_value('Sheet1!D5')
        value = self.evaluator.evaluate('Sheet1!D5')
        self.assertEqual(excel_value, value)

    def test_iserror_with_number(self):
        """Test ISERROR with numeric value."""
        excel_value = self.evaluator.get_cell_value('Sheet1!E1')
        value = self.evaluator.evaluate('Sheet1!E1')
        self.assertEqual(excel_value, value)

    def test_iserror_with_error(self):
        """Test ISERROR with error value."""
        excel_value = self.evaluator.get_cell_value('Sheet1!E5')
        value = self.evaluator.evaluate('Sheet1!E5')
        self.assertEqual(excel_value, value)

    def test_iserror_with_na(self):
        """Test ISERROR with #N/A error."""
        excel_value = self.evaluator.get_cell_value('Sheet1!E6')
        value = self.evaluator.evaluate('Sheet1!E6')
        self.assertEqual(excel_value, value)

    def test_iserror_direct(self):
        """Test ISERROR with direct error."""
        excel_value = self.evaluator.get_cell_value('Sheet1!E7')
        value = self.evaluator.evaluate('Sheet1!E7')
        self.assertEqual(excel_value, value)

    def test_isna_with_na(self):
        """Test ISNA with #N/A error."""
        excel_value = self.evaluator.get_cell_value('Sheet1!F6')
        value = self.evaluator.evaluate('Sheet1!F6')
        self.assertEqual(excel_value, value)

    def test_isna_with_other_error(self):
        """Test ISNA with other error."""
        excel_value = self.evaluator.get_cell_value('Sheet1!F5')
        value = self.evaluator.evaluate('Sheet1!F5')
        self.assertEqual(excel_value, value)

    def test_isna_with_number(self):
        """Test ISNA with numeric value."""
        excel_value = self.evaluator.get_cell_value('Sheet1!F1')
        value = self.evaluator.evaluate('Sheet1!F1')
        self.assertEqual(excel_value, value)

    def test_iserr_with_div_error(self):
        """Test ISERR with #DIV/0! error."""
        excel_value = self.evaluator.get_cell_value('Sheet1!G5')
        value = self.evaluator.evaluate('Sheet1!G5')
        self.assertEqual(excel_value, value)

    def test_iserr_with_na_error(self):
        """Test ISERR with #N/A error (should be FALSE)."""
        excel_value = self.evaluator.get_cell_value('Sheet1!G6')
        value = self.evaluator.evaluate('Sheet1!G6')
        self.assertEqual(excel_value, value)

    def test_iserr_with_number(self):
        """Test ISERR with numeric value."""
        excel_value = self.evaluator.get_cell_value('Sheet1!G1')
        value = self.evaluator.evaluate('Sheet1!G1')
        self.assertEqual(excel_value, value)

    def test_iseven_with_even(self):
        """Test ISEVEN with even number."""
        excel_value = self.evaluator.get_cell_value('Sheet1!H7')
        value = self.evaluator.evaluate('Sheet1!H7')
        self.assertEqual(excel_value, value)

    def test_iseven_with_odd(self):
        """Test ISEVEN with odd number."""
        excel_value = self.evaluator.get_cell_value('Sheet1!H8')
        value = self.evaluator.evaluate('Sheet1!H8')
        self.assertEqual(excel_value, value)

    def test_iseven_with_zero(self):
        """Test ISEVEN with zero."""
        excel_value = self.evaluator.get_cell_value('Sheet1!H9')
        value = self.evaluator.evaluate('Sheet1!H9')
        self.assertEqual(excel_value, value)

    def test_iseven_with_negative_odd(self):
        """Test ISEVEN with negative odd number."""
        excel_value = self.evaluator.get_cell_value('Sheet1!H10')
        value = self.evaluator.evaluate('Sheet1!H10')
        self.assertEqual(excel_value, value)

    def test_isodd_with_even(self):
        """Test ISODD with even number."""
        excel_value = self.evaluator.get_cell_value('Sheet1!I7')
        value = self.evaluator.evaluate('Sheet1!I7')
        self.assertEqual(excel_value, value)

    def test_isodd_with_odd(self):
        """Test ISODD with odd number."""
        excel_value = self.evaluator.get_cell_value('Sheet1!I8')
        value = self.evaluator.evaluate('Sheet1!I8')
        self.assertEqual(excel_value, value)

    def test_isodd_with_zero(self):
        """Test ISODD with zero."""
        excel_value = self.evaluator.get_cell_value('Sheet1!I9')
        value = self.evaluator.evaluate('Sheet1!I9')
        self.assertEqual(excel_value, value)

    def test_isodd_with_negative_odd(self):
        """Test ISODD with negative odd number."""
        excel_value = self.evaluator.get_cell_value('Sheet1!I10')
        value = self.evaluator.evaluate('Sheet1!I10')
        self.assertEqual(excel_value, value)

    def test_na_function(self):
        """Test NA() function."""
        excel_value = self.evaluator.get_cell_value('Sheet1!J1')
        value = self.evaluator.evaluate('Sheet1!J1')
        self.assertEqual(excel_value, value)