from .. import testing


class LogicalTest(testing.FunctionalTestCase):
    filename = "logical.xlsx"

    def test_and_true_true(self):
        """Test AND with two TRUE values."""
        excel_value = self.evaluator.get_cell_value('Sheet1!A3')
        value = self.evaluator.evaluate('Sheet1!A3')
        self.assertEqual(excel_value, value)

    def test_and_true_false(self):
        """Test AND with TRUE and FALSE."""
        excel_value = self.evaluator.get_cell_value('Sheet1!A4')
        value = self.evaluator.evaluate('Sheet1!A4')
        self.assertEqual(excel_value, value)

    def test_and_false_false(self):
        """Test AND with two FALSE values."""
        excel_value = self.evaluator.get_cell_value('Sheet1!A5')
        value = self.evaluator.evaluate('Sheet1!A5')
        self.assertEqual(excel_value, value)

    def test_and_cell_references(self):
        """Test AND with cell references."""
        excel_value = self.evaluator.get_cell_value('Sheet1!A6')
        value = self.evaluator.evaluate('Sheet1!A6')
        self.assertEqual(excel_value, value)

    def test_and_logical_expressions(self):
        """Test AND with logical expressions."""
        excel_value = self.evaluator.get_cell_value('Sheet1!A7')
        value = self.evaluator.evaluate('Sheet1!A7')
        self.assertEqual(excel_value, value)

    def test_and_multiple_conditions(self):
        """Test AND with multiple conditions."""
        excel_value = self.evaluator.get_cell_value('Sheet1!A8')
        value = self.evaluator.evaluate('Sheet1!A8')
        self.assertEqual(excel_value, value)

    def test_or_true_true(self):
        """Test OR with two TRUE values."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B3')
        value = self.evaluator.evaluate('Sheet1!B3')
        self.assertEqual(excel_value, value)

    def test_or_true_false(self):
        """Test OR with TRUE and FALSE."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B4')
        value = self.evaluator.evaluate('Sheet1!B4')
        self.assertEqual(excel_value, value)

    def test_or_false_false(self):
        """Test OR with two FALSE values."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B5')
        value = self.evaluator.evaluate('Sheet1!B5')
        self.assertEqual(excel_value, value)

    def test_or_cell_references(self):
        """Test OR with cell references."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B6')
        value = self.evaluator.evaluate('Sheet1!B6')
        self.assertEqual(excel_value, value)

    def test_or_logical_expressions(self):
        """Test OR with logical expressions."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B7')
        value = self.evaluator.evaluate('Sheet1!B7')
        self.assertEqual(excel_value, value)

    def test_or_multiple_conditions(self):
        """Test OR with multiple conditions."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B8')
        value = self.evaluator.evaluate('Sheet1!B8')
        self.assertEqual(excel_value, value)

    def test_true_constant(self):
        """Test TRUE() function."""
        excel_value = self.evaluator.get_cell_value('Sheet1!C3')
        value = self.evaluator.evaluate('Sheet1!C3')
        self.assertEqual(excel_value, value)

    def test_false_constant(self):
        """Test FALSE() function."""
        excel_value = self.evaluator.get_cell_value('Sheet1!C4')
        value = self.evaluator.evaluate('Sheet1!C4')
        self.assertEqual(excel_value, value)

    def test_nested_logical_1(self):
        """Test nested logical functions."""
        excel_value = self.evaluator.get_cell_value('Sheet1!D3')
        value = self.evaluator.evaluate('Sheet1!D3')
        self.assertEqual(excel_value, value)

    def test_nested_logical_2(self):
        """Test nested logical functions."""
        excel_value = self.evaluator.get_cell_value('Sheet1!D4')
        value = self.evaluator.evaluate('Sheet1!D4')
        self.assertEqual(excel_value, value)

    def test_and_empty(self):
        """Test AND with no arguments."""
        excel_value = self.evaluator.get_cell_value('Sheet1!E3')
        value = self.evaluator.evaluate('Sheet1!E3')
        self.assertEqual(excel_value, value)

    def test_or_empty(self):
        """Test OR with no arguments."""
        excel_value = self.evaluator.get_cell_value('Sheet1!E4')
        value = self.evaluator.evaluate('Sheet1!E4')
        self.assertEqual(excel_value, value)