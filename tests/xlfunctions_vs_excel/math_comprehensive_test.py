from .. import testing


class MathComprehensiveTest(testing.FunctionalTestCase):
    filename = "MATH.xlsx"

    def test_floor_positive_decimal(self):
        """Test FLOOR with positive decimal."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B1')
        value = self.evaluator.evaluate('Sheet1!B1')
        self.assertEqual(excel_value, value)

    def test_floor_negative_decimal(self):
        """Test FLOOR with negative decimal."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B2')
        value = self.evaluator.evaluate('Sheet1!B2')
        self.assertEqual(excel_value, value)

    def test_floor_with_significance(self):
        """Test FLOOR with custom significance."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B3')
        value = self.evaluator.evaluate('Sheet1!B3')
        self.assertEqual(excel_value, value)

    def test_floor_exact_multiple(self):
        """Test FLOOR with exact multiple."""
        excel_value = self.evaluator.get_cell_value('Sheet1!B4')
        value = self.evaluator.evaluate('Sheet1!B4')
        self.assertEqual(excel_value, value)

    def test_trunc_positive_decimal(self):
        """Test TRUNC with positive decimal."""
        excel_value = self.evaluator.get_cell_value('Sheet1!C1')
        value = self.evaluator.evaluate('Sheet1!C1')
        self.assertEqual(excel_value, value)

    def test_trunc_negative_decimal(self):
        """Test TRUNC with negative decimal."""
        excel_value = self.evaluator.get_cell_value('Sheet1!C2')
        value = self.evaluator.evaluate('Sheet1!C2')
        self.assertEqual(excel_value, value)

    def test_trunc_with_precision(self):
        """Test TRUNC with precision."""
        excel_value = self.evaluator.get_cell_value('Sheet1!C3')
        value = self.evaluator.evaluate('Sheet1!C3')
        self.assertEqual(excel_value, value)

    def test_trunc_negative_precision(self):
        """Test TRUNC with negative precision."""
        excel_value = self.evaluator.get_cell_value('Sheet1!C4')
        value = self.evaluator.evaluate('Sheet1!C4')
        self.assertEqual(excel_value, value)

    def test_sign_positive(self):
        """Test SIGN with positive number."""
        excel_value = self.evaluator.get_cell_value('Sheet1!D1')
        value = self.evaluator.evaluate('Sheet1!D1')
        self.assertEqual(excel_value, value)

    def test_sign_negative(self):
        """Test SIGN with negative number."""
        excel_value = self.evaluator.get_cell_value('Sheet1!D2')
        value = self.evaluator.evaluate('Sheet1!D2')
        self.assertEqual(excel_value, value)

    def test_sign_zero(self):
        """Test SIGN with zero."""
        excel_value = self.evaluator.get_cell_value('Sheet1!D3')
        value = self.evaluator.evaluate('Sheet1!D3')
        self.assertEqual(excel_value, value)

    def test_log_base_10(self):
        """Test LOG with base 10."""
        excel_value = self.evaluator.get_cell_value('Sheet1!E1')
        value = self.evaluator.evaluate('Sheet1!E1')
        self.assertEqual(excel_value, value)

    def test_log_base_2(self):
        """Test LOG with base 2."""
        excel_value = self.evaluator.get_cell_value('Sheet1!E2')
        value = self.evaluator.evaluate('Sheet1!E2')
        self.assertEqual(excel_value, value)

    def test_log_natural(self):
        """Test LOG with natural base."""
        excel_value = self.evaluator.get_cell_value('Sheet1!E3')
        value = self.evaluator.evaluate('Sheet1!E3')
        self.assertEqual(excel_value, value)

    def test_log10_100(self):
        """Test LOG10 with 100."""
        excel_value = self.evaluator.get_cell_value('Sheet1!F1')
        value = self.evaluator.evaluate('Sheet1!F1')
        self.assertEqual(excel_value, value)

    def test_log10_1000(self):
        """Test LOG10 with 1000."""
        excel_value = self.evaluator.get_cell_value('Sheet1!F2')
        value = self.evaluator.evaluate('Sheet1!F2')
        self.assertEqual(excel_value, value)

    def test_exp_zero(self):
        """Test EXP with 0."""
        excel_value = self.evaluator.get_cell_value('Sheet1!G1')
        value = self.evaluator.evaluate('Sheet1!G1')
        self.assertEqual(excel_value, value)

    def test_exp_one(self):
        """Test EXP with 1."""
        excel_value = self.evaluator.get_cell_value('Sheet1!G2')
        value = self.evaluator.evaluate('Sheet1!G2')
        self.assertEqual(excel_value, value)

    def test_exp_two(self):
        """Test EXP with 2."""
        excel_value = self.evaluator.get_cell_value('Sheet1!G3')
        value = self.evaluator.evaluate('Sheet1!G3')
        self.assertEqual(excel_value, value)