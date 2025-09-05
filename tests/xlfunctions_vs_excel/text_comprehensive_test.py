from .. import testing


class TextComprehensiveTest(testing.FunctionalTestCase):
    filename = "TEXT.xlsx"

    def test_left_five_chars(self):
        """Test LEFT with 5 characters."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B1')
        value = self.evaluator.evaluate('Sheet1!B1')
        self.assertEqual(excel_value, value)

    def test_left_one_char(self):
        """Test LEFT with 1 character."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B2')
        value = self.evaluator.evaluate('Sheet1!B2')
        self.assertEqual(excel_value, value)

    def test_left_default(self):
        """Test LEFT with default parameter."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B3')
        value = self.evaluator.evaluate('Sheet1!B3')
        self.assertEqual(excel_value, value)

    def test_upper_already_uppercase(self):
        """Test UPPER with already uppercase text."""
        excel_value = self.evaluator.get_cell_value('Sheet1!C3')
        value = self.evaluator.evaluate('Sheet1!C3')
        self.assertEqual(excel_value, value)

    def test_upper_convert_lowercase(self):
        """Test UPPER converting lowercase."""
        excel_value = self.evaluator.get_cell_value('Sheet1!C4')
        value = self.evaluator.evaluate('Sheet1!C4')
        self.assertEqual(excel_value, value)

    def test_upper_convert_mixed(self):
        """Test UPPER converting mixed case."""
        excel_value = self.evaluator.get_cell_value('Sheet1!C5')
        value = self.evaluator.evaluate('Sheet1!C5')
        self.assertEqual(excel_value, value)

    def test_lower_convert_uppercase(self):
        """Test LOWER converting uppercase."""
        excel_value = self.evaluator.get_cell_value('Sheet1!D3')
        value = self.evaluator.evaluate('Sheet1!D3')
        self.assertEqual(excel_value, value)

    def test_lower_already_lowercase(self):
        """Test LOWER with already lowercase text."""
        excel_value = self.evaluator.get_cell_value('Sheet1!D4')
        value = self.evaluator.evaluate('Sheet1!D4')
        self.assertEqual(excel_value, value)

    def test_lower_convert_mixed(self):
        """Test LOWER converting mixed case."""
        excel_value = self.evaluator.get_cell_value('Sheet1!D5')
        value = self.evaluator.evaluate('Sheet1!D5')
        self.assertEqual(excel_value, value)

    def test_trim_with_spaces(self):
        """Test TRIM removing leading/trailing spaces."""
        excel_value = self.evaluator.get_cell_value('Sheet1!E2')
        value = self.evaluator.evaluate('Sheet1!E2')
        self.assertEqual(excel_value, value)

    def test_trim_no_spaces(self):
        """Test TRIM with no extra spaces."""
        excel_value = self.evaluator.get_cell_value('Sheet1!E1')
        value = self.evaluator.evaluate('Sheet1!E1')
        self.assertEqual(excel_value, value)

    def test_replace_this_with_that(self):
        """Test REPLACE changing 'This' to 'That'."""
        excel_value = self.evaluator.get_cell_value('Sheet1!F6')
        value = self.evaluator.evaluate('Sheet1!F6')
        self.assertEqual(excel_value, value)

    def test_replace_hello_with_hi(self):
        """Test REPLACE changing 'Hello' to 'Hi'."""
        excel_value = self.evaluator.get_cell_value('Sheet1!F1')
        value = self.evaluator.evaluate('Sheet1!F1')
        self.assertEqual(excel_value, value)