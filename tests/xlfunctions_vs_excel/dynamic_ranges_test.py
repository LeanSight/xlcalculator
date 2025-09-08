"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class DynamicRangesComprehensiveTest(testing.FunctionalTestCase):
    """
    These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.
    
    Tests: 75 cases across 19 levels
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "dynamic_ranges.xlsx"

    def test_1a(self):
        """INDEX - Casos Fundamentales: Valores individuales básicos"""
        
        # INDEX básico - valor numérico
        value = self.evaluator.evaluate('A1')
        self.assertEqual(25, value, "=INDEX(Data!A1:E6, 2, 2) should return 25")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # INDEX básico - texto
        value = self.evaluator.evaluate('A2')
        self.assertEqual('Bob', value, "=INDEX(Data!A1:E6, 3, 1) should return 'Bob'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # INDEX básico - boolean
        value = self.evaluator.evaluate('A3')
        self.assertEqual(True, value, "=INDEX(Data!A1:E6, 4, 5) should return True")
        self.assertIsInstance(value, (bool, Boolean), "Should be boolean")

        # INDEX básico - última fila
        value = self.evaluator.evaluate('A4')
        self.assertEqual('Eve', value, "=INDEX(Data!A1:E6, 6, 1) should return 'Eve'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # INDEX básico - primera fila
        value = self.evaluator.evaluate('A5')
        self.assertEqual('Active', value, "=INDEX(Data!A1:E6, 1, 5) should return 'Active'")
        self.assertIsInstance(value, (str, Text), "Should be text")

    def test_1b(self):
        """INDEX - Arrays Completos: Filas y columnas completas"""
        
        # INDEX array - columna completa Age
        value = self.evaluator.evaluate('B1')
        self.assertIsInstance(value, Array, "=INDEX(Data!A1:E6, 0, 2) should return Array")

        # INDEX array - fila completa Alice
        value = self.evaluator.evaluate('B2')
        self.assertIsInstance(value, Array, "=INDEX(Data!A1:E6, 2, 0) should return Array")

        # INDEX array - primera columna Name
        value = self.evaluator.evaluate('B3')
        self.assertIsInstance(value, Array, "=INDEX(Data!A1:E6, 0, 1) should return Array")

        # INDEX array - columna Active
        value = self.evaluator.evaluate('B4')
        self.assertIsInstance(value, Array, "=INDEX(Data!A1:E6, 0, 5) should return Array")

    def test_1c(self):
        """INDEX - Casos de Error Estructurales: Manejo de errores"""
        
        # INDEX error - fila fuera de rango
        value = self.evaluator.evaluate('C1')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=INDEX(Data!A1:E6, 7, 1) should return REF_ERROR")

        # INDEX error - columna fuera de rango
        value = self.evaluator.evaluate('C2')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=INDEX(Data!A1:E6, 1, 7) should return REF_ERROR")

        # INDEX error - ambos cero
        value = self.evaluator.evaluate('C3')
        self.assertIsInstance(value, xlerrors.ValueExcelError,
                            "=INDEX(Data!A1:E6, 0, 0) should return VALUE_ERROR")

        # INDEX error - fila negativa
        value = self.evaluator.evaluate('C4')
        self.assertIsInstance(value, xlerrors.ValueExcelError,
                            "=INDEX(Data!A1:E6, -1, 1) should return VALUE_ERROR")

        # INDEX error - columna negativa
        value = self.evaluator.evaluate('C5')
        self.assertIsInstance(value, xlerrors.ValueExcelError,
                            "=INDEX(Data!A1:E6, 1, -1) should return VALUE_ERROR")

    def test_2d(self):
        """OFFSET - Casos Fundamentales: Valores individuales básicos"""
        
        # OFFSET básico - B2
        value = self.evaluator.evaluate('D1')
        self.assertEqual(25, value, "=OFFSET(Data!A1, 1, 1) should return 25")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # OFFSET básico - desde B2
        value = self.evaluator.evaluate('D2')
        self.assertEqual('LA', value, "=OFFSET(Data!B2, 1, 1) should return 'LA'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # OFFSET básico - horizontal
        value = self.evaluator.evaluate('D3')
        self.assertEqual('City', value, "=OFFSET(Data!A1, 0, 2) should return 'City'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # OFFSET básico - esquina
        value = self.evaluator.evaluate('D4')
        self.assertEqual(False, value, "=OFFSET(Data!A1, 5, 4) should return False")
        self.assertIsInstance(value, (bool, Boolean), "Should be boolean")

        # OFFSET básico - offset negativo
        value = self.evaluator.evaluate('D5')
        self.assertEqual(30, value, "=OFFSET(Data!C3, -1, 1) should return 30")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

    def test_2e(self):
        """OFFSET - Arrays con Dimensiones: Rangos con height/width"""
        
        # OFFSET dimensiones - 1x1
        value = self.evaluator.evaluate('E1')
        self.assertEqual(25, value, "=OFFSET(Data!A1, 1, 1, 1, 1) should return 25")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # OFFSET dimensiones - 2x2
        value = self.evaluator.evaluate('E2')
        self.assertIsInstance(value, Array, "=OFFSET(Data!A1, 1, 1, 2, 2) should return Array")

        # OFFSET dimensiones - 3x3
        value = self.evaluator.evaluate('E3')
        self.assertIsInstance(value, Array, "=OFFSET(Data!A1, 0, 0, 3, 3) should return Array")

        # OFFSET dimensiones - 1x3
        value = self.evaluator.evaluate('E4')
        self.assertIsInstance(value, Array, "=OFFSET(Data!A1, 2, 1, 1, 3) should return Array")

        # OFFSET dimensiones - 3x1
        value = self.evaluator.evaluate('E5')
        self.assertIsInstance(value, Array, "=OFFSET(Data!A1, 1, 0, 3, 1) should return Array")

    def test_2f(self):
        """OFFSET - Casos de Error: Errores de referencia y parámetros"""
        
        # OFFSET error - antes del inicio de hoja
        value = self.evaluator.evaluate('F1')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=OFFSET(Data!A1, -2, 0) should return REF_ERROR")

        # OFFSET error - antes del inicio de hoja
        value = self.evaluator.evaluate('F2')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=OFFSET(Data!A1, 0, -2) should return REF_ERROR")

        # OFFSET error - más allá de hoja
        value = self.evaluator.evaluate('F3')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=OFFSET(Data!A1, 100, 0) should return REF_ERROR")

        # OFFSET error - más allá de hoja
        value = self.evaluator.evaluate('F4')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=OFFSET(Data!A1, 0, 100) should return REF_ERROR")

        # OFFSET error - altura cero
        value = self.evaluator.evaluate('F5')
        self.assertIsInstance(value, xlerrors.ValueExcelError,
                            "=OFFSET(Data!A1, 1, 1, 0, 1) should return VALUE_ERROR")

        # OFFSET error - ancho cero
        value = self.evaluator.evaluate('F6')
        self.assertIsInstance(value, xlerrors.ValueExcelError,
                            "=OFFSET(Data!A1, 1, 1, 1, 0) should return VALUE_ERROR")

    def test_2g(self):
        """INDIRECT - Casos Fundamentales: Referencias directas"""
        
        # INDIRECT básico - valor numérico
        value = self.evaluator.evaluate('G1')
        self.assertEqual(25, value, "=INDIRECT(\"Data!B2\") should return 25")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # INDIRECT básico - texto
        value = self.evaluator.evaluate('G2')
        self.assertEqual('LA', value, "=INDIRECT(\"Data!C3\") should return 'LA'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # INDIRECT básico - boolean
        value = self.evaluator.evaluate('G3')
        self.assertEqual(True, value, "=INDIRECT(\"Data!E4\") should return True")
        self.assertIsInstance(value, (bool, Boolean), "Should be boolean")

        # INDIRECT básico - desde celda
        value = self.evaluator.evaluate('G4')
        self.assertEqual(25, value, "=INDIRECT(P1) should return 25")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

    def test_2h(self):
        """INDIRECT - Referencias Dinámicas: Referencias construidas"""
        
        # INDIRECT dinámico - concatenación
        value = self.evaluator.evaluate('H1')
        self.assertEqual('Alice', value, "=INDIRECT(\"Data!A\" & 2) should return 'Alice'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # INDIRECT dinámico - CHAR
        value = self.evaluator.evaluate('H2')
        self.assertEqual(30, value, "=INDIRECT(\"Data!\" & CHAR(66) & \"3\") should return 30")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # INDIRECT dinámico - ROW
        value = self.evaluator.evaluate('H3')
        self.assertEqual('Charlie', value, "=INDIRECT(\"Data!A\" & ROW()) should return 'Charlie'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # INDIRECT dinámico - COLUMN
        value = self.evaluator.evaluate('H4')
        self.assertEqual('Score', value, "=INDIRECT(\"Data!\" & CHAR(65+COLUMN()) & \"1\") should return 'Score'")
        self.assertIsInstance(value, (str, Text), "Should be text")

    def test_2i(self):
        """INDIRECT - Arrays de Referencias: Rangos y arrays"""
        
        # INDIRECT array - headers
        value = self.evaluator.evaluate('I1')
        self.assertIsInstance(value, Array, "=INDIRECT(\"Data!A1:C1\") should return Array")

        # INDIRECT array - columna nombres
        value = self.evaluator.evaluate('I2')
        self.assertIsInstance(value, Array, "=INDIRECT(\"Data!A2:A6\") should return Array")

        # INDIRECT array - columna edad
        value = self.evaluator.evaluate('I3')
        self.assertIsInstance(value, Array, "=INDIRECT(\"Data!B1:B6\") should return Array")

        # INDIRECT array - desde celda
        value = self.evaluator.evaluate('I4')
        self.assertIsInstance(value, Array, "=INDIRECT(P3) should return Array")

    def test_2j(self):
        """INDIRECT - Referencias de Columna/Fila Completa: Referencias completas"""
        
        # INDIRECT columna completa A
        value = self.evaluator.evaluate('J1')
        self.assertIsInstance(value, Array, "=INDIRECT(\"Data!A:A\") should return Array")

        # INDIRECT columna completa B
        value = self.evaluator.evaluate('J2')
        self.assertIsInstance(value, Array, "=INDIRECT(\"Data!B:B\") should return Array")

        # INDIRECT fila completa 1
        value = self.evaluator.evaluate('J3')
        self.assertIsInstance(value, Array, "=INDIRECT(\"Data!1:1\") should return Array")

        # INDIRECT fila completa 2
        value = self.evaluator.evaluate('J4')
        self.assertIsInstance(value, Array, "=INDIRECT(\"Data!2:2\") should return Array")

    def test_2k(self):
        """INDIRECT - Casos de Error: Referencias inválidas"""
        
        # INDIRECT error - hoja inexistente
        value = self.evaluator.evaluate('K1')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=INDIRECT(\"InvalidSheet!A1\") should return REF_ERROR")

        # INDIRECT error - celda vacía
        value = self.evaluator.evaluate('K2')
        self.assertEqual(0, value, "=INDIRECT(\"Data!Z99\") should return 0")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # INDIRECT error - referencia vacía
        value = self.evaluator.evaluate('K3')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=INDIRECT(\"\") should return REF_ERROR")

        # INDIRECT error - texto inválido
        value = self.evaluator.evaluate('K4')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=INDIRECT(\"NotAReference\") should return REF_ERROR")

        # INDIRECT error - hoja inválida desde celda
        value = self.evaluator.evaluate('K5')
        self.assertIsInstance(value, xlerrors.RefExcelError,
                            "=INDIRECT(P4) should return REF_ERROR")

    def test_3l(self):
        """INDEX + INDIRECT: Combinaciones INDEX+INDIRECT"""
        
        # Combinación INDEX+INDIRECT
        value = self.evaluator.evaluate('L1')
        self.assertEqual(25, value, "=INDEX(INDIRECT(\"Data!A1:E6\"), 2, 2) should return 25")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # Combinación INDEX+INDIRECT array
        value = self.evaluator.evaluate('L2')
        self.assertIsInstance(value, Array, "=INDEX(INDIRECT(\"Data!A1:E6\"), 0, 2) should return Array")

        # Combinación INDEX+INDIRECT subrange
        value = self.evaluator.evaluate('L3')
        self.assertEqual('Chicago', value, "=INDEX(INDIRECT(\"Data!A2:C4\"), 2, 3) should return 'Chicago'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # Combinación INDEX+INDIRECT columna completa
        value = self.evaluator.evaluate('L4')
        self.assertEqual('Bob', value, "=INDEX(INDIRECT(\"Data!A:A\"), 3) should return 'Bob'")
        self.assertIsInstance(value, (str, Text), "Should be text")

    def test_3m(self):
        """OFFSET + INDIRECT: Combinaciones OFFSET+INDIRECT"""
        
        # Combinación OFFSET+INDIRECT
        value = self.evaluator.evaluate('M1')
        self.assertEqual(25, value, "=OFFSET(INDIRECT(\"Data!A1\"), 1, 1) should return 25")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # Combinación OFFSET+INDIRECT desde B2
        value = self.evaluator.evaluate('M2')
        self.assertEqual('LA', value, "=OFFSET(INDIRECT(\"Data!B2\"), 1, 1) should return 'LA'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # Combinación OFFSET+INDIRECT array
        value = self.evaluator.evaluate('M3')
        self.assertIsInstance(value, Array, "=OFFSET(INDIRECT(\"Data!A1\"), 1, 1, 2, 2) should return Array")

    def test_3n(self):
        """Combinaciones Complejas: Funciones anidadas complejas"""
        
        # Combinación INDEX+OFFSET
        value = self.evaluator.evaluate('N1')
        self.assertEqual(25, value, "=INDEX(OFFSET(Data!A1, 0, 0, 3, 3), 2, 2) should return 25")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # Combinación OFFSET+INDEX
        value = self.evaluator.evaluate('N2')
        self.assertEqual(30, value, "=OFFSET(INDEX(Data!A1:E6, 2, 1), 1, 1) should return 30")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # Referencia dinámica compleja
        value = self.evaluator.evaluate('N3')
        self.assertEqual('Alice', value, "=INDIRECT(\"Data!\" & \"A\" & INDEX(Data!B1:B6, 2, 1)) should return 'Alice'")
        self.assertIsInstance(value, (str, Text), "Should be text")

    def test_4o(self):
        """Funciones con Agregación: Uso con funciones de agregado"""
        
        # SUM con INDEX array
        value = self.evaluator.evaluate('O1')
        self.assertEqual(130, value, "=SUM(INDEX(Data!A1:E6, 0, 2)) should return 130")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # AVERAGE con OFFSET array
        value = self.evaluator.evaluate('O2')
        self.assertEqual(28, value, "=AVERAGE(OFFSET(Data!B1, 1, 0, 5, 1)) should return 28")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # COUNT con INDIRECT columna
        value = self.evaluator.evaluate('O3')
        self.assertEqual(5, value, "=COUNT(INDIRECT(\"Data!B:B\")) should return 5")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # MAX con INDEX array
        value = self.evaluator.evaluate('O4')
        self.assertEqual(95, value, "=MAX(INDEX(Data!A1:E6, 0, 4)) should return 95")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

    def test_4p(self):
        """Manejo de Errores: Manejo de errores con IFERROR/ISERROR"""
        
        # Manejo errores IFERROR+INDEX
        value = self.evaluator.evaluate('P1')
        self.assertEqual('Not Found', value, "=IFERROR(INDEX(Data!A1:E6, 10, 1), \"Not Found\") should return 'Not Found'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # Detección errores IF+ISERROR+OFFSET
        value = self.evaluator.evaluate('P2')
        self.assertEqual('Error', value, "=IF(ISERROR(OFFSET(Data!A1, -1, 0)), \"Error\", \"OK\") should return 'Error'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # Manejo errores IFERROR+INDIRECT
        value = self.evaluator.evaluate('P3')
        self.assertEqual('Sheet Error', value, "=IFERROR(INDIRECT(\"InvalidSheet!A1\"), \"Sheet Error\") should return 'Sheet Error'")
        self.assertIsInstance(value, (str, Text), "Should be text")

    def test_5q(self):
        """Referencias Especiales: Referencias especiales y complejas"""
        
        # INDIRECT misma hoja
        value = self.evaluator.evaluate('Q1')
        self.assertEqual('Test Value', value, "=INDIRECT(\"Tests!O1\") should return 'Test Value'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # INDEX con columna completa
        value = self.evaluator.evaluate('Q2')
        self.assertEqual('Alice', value, "=INDEX(Data!A:A, 2) should return 'Alice'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # OFFSET con columna completa
        value = self.evaluator.evaluate('Q3')
        self.assertIsInstance(value, Array, "=OFFSET(Data!A:A, 1, 0, 3, 1) should return Array")

    def test_5r(self):
        """Arrays Dinámicos: Comportamiento con arrays dinámicos"""
        
        # INDEX con array de filas
        value = self.evaluator.evaluate('R1')
        self.assertIsInstance(value, Array, "=INDEX(Data!A1:E6, ROW(A1:A3), 1) should return Array")

        # OFFSET con array de offsets
        value = self.evaluator.evaluate('R2')
        self.assertIsInstance(value, Array, "=OFFSET(Data!A1, ROW(A1:A2)-1, 0) should return Array")

    def test_5s(self):
        """Forma de Referencia vs Array: Casos edge de formas de referencia"""
        
        # INDEX forma referencia área 1
        value = self.evaluator.evaluate('S1')
        self.assertEqual('Alice', value, "=INDEX((Data!A1:A5, Data!C1:C5), 2, 1, 1) should return 'Alice'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # INDEX forma referencia área 2
        value = self.evaluator.evaluate('S2')
        self.assertEqual('NYC', value, "=INDEX((Data!A1:A5, Data!C1:C5), 2, 1, 2) should return 'NYC'")
        self.assertIsInstance(value, (str, Text), "Should be text")

    def test_data_integrity(self):
        """Verify test data integrity."""
        # Auto-generated data validation
        self.assertEqual('Alice', self.evaluator.evaluate('Data!A2'))
        self.assertEqual(25, self.evaluator.evaluate('Data!B2'))
        self.assertEqual('NYC', self.evaluator.evaluate('Data!C2'))

    def test_type_consistency(self):
        """Verify data type consistency across test cases."""
        # Auto-generated type validation
        # number validation
        number_value = self.evaluator.evaluate('A1')
        self.assertIsInstance(number_value, (int, float, Number))

        # text validation
        text_value = self.evaluator.evaluate('A2')
        self.assertIsInstance(text_value, (str, Text))

        # boolean validation
        boolean_value = self.evaluator.evaluate('A3')
        self.assertIsInstance(boolean_value, (bool, Boolean))

        # array validation
        array_value = self.evaluator.evaluate('B1')
        self.assertIsInstance(array_value, Array)
