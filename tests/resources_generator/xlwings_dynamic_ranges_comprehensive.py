#!/usr/bin/env python3
"""
Generador comprehensivo de Excel para rangos dinámicos usando xlwings.
Este archivo genera un Excel que captura FIELMENTE el comportamiento de Excel
para todas las funciones de rangos dinámicos.

Ejecutar en Windows con Excel instalado.
"""

import xlwings as xw
import os


def create_comprehensive_dynamic_ranges_excel(filepath):
    """Crear Excel comprehensivo para rangos dinámicos con comportamiento fiel a Excel."""
    
    # Iniciar Excel con configuración robusta
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    
    try:
        wb = app.books.add()
        
        # === HOJA 1: DATA ===
        data_sheet = wb.sheets[0]
        data_sheet.name = "Data"
        
        print("📊 Creando hoja de datos...")
        
        # Headers
        data_sheet['A1'].value = 'Name'
        data_sheet['B1'].value = 'Age'
        data_sheet['C1'].value = 'City'
        data_sheet['D1'].value = 'Score'
        data_sheet['E1'].value = 'Active'
        data_sheet['F1'].value = 'Notes'
        
        # Datos de prueba
        data_sheet['A2'].value = 'Alice'
        data_sheet['B2'].value = 25
        data_sheet['C2'].value = 'NYC'
        data_sheet['D2'].value = 85
        data_sheet['E2'].value = True
        data_sheet['F2'].value = 'Good'
        
        data_sheet['A3'].value = 'Bob'
        data_sheet['B3'].value = 30
        data_sheet['C3'].value = 'LA'
        data_sheet['D3'].value = 92
        data_sheet['E3'].value = False
        data_sheet['F3'].value = 'Great'
        
        data_sheet['A4'].value = 'Charlie'
        data_sheet['B4'].value = 35
        data_sheet['C4'].value = 'Chicago'
        data_sheet['D4'].value = 78
        data_sheet['E4'].value = True
        data_sheet['F4'].value = 'OK'
        
        data_sheet['A5'].value = 'Diana'
        data_sheet['B5'].value = 28
        data_sheet['C5'].value = 'Miami'
        data_sheet['D5'].value = 95
        data_sheet['E5'].value = True
        data_sheet['F5'].value = 'Excellent'
        
        data_sheet['A6'].value = 'Eve'
        data_sheet['B6'].value = 22
        data_sheet['C6'].value = 'Boston'
        data_sheet['D6'].value = 88
        data_sheet['E6'].value = False
        data_sheet['F6'].value = 'Average'
        
        # === HOJA 2: TESTS ===
        tests_sheet = wb.sheets.add("Tests")
        
        print("🧪 Creando casos de prueba...")
        
        # Referencias auxiliares para INDIRECT
        tests_sheet['P1'].value = 'Data.B2'
        tests_sheet['P2'].value = 'Data.C3'
        tests_sheet['P3'].value = 'Data.A1:C3'
        tests_sheet['P4'].value = 'InvalidRef'
        tests_sheet['P5'].value = ''
        
        # Valores esperados para validación
        tests_sheet['Q1'].value = 25
        tests_sheet['Q2'].value = 'Bob'
        tests_sheet['Q3'].value = True
        tests_sheet['Q4'].value = '#REF!'
        tests_sheet['Q5'].value = '#VALUE!'
        
        # Valor para referencia circular
        tests_sheet['O1'].value = 'Test Value'
        
        # Definir todas las fórmulas organizadas por nivel
        formulas = [
            # NIVEL 1: CASOS ESTRUCTURALES
            # A. INDEX - Casos Fundamentales
            ('A1', '=INDEX(Data.A1:E6, 2, 2)', 'INDEX básico - valor numérico'),
            ('A2', '=INDEX(Data.A1:E6, 3, 1)', 'INDEX básico - texto'),
            ('A3', '=INDEX(Data.A1:E6, 4, 5)', 'INDEX básico - boolean'),
            ('A4', '=INDEX(Data.A1:E6, 6, 1)', 'INDEX básico - última fila'),
            ('A5', '=INDEX(Data.A1:E6, 1, 5)', 'INDEX básico - primera fila'),
            
            # B. INDEX - Casos de Error Estructurales
            ('B1', '=INDEX(Data.A1:E6, 7, 1)', 'INDEX error - fila fuera de rango'),
            ('B2', '=INDEX(Data.A1:E6, 1, 7)', 'INDEX error - columna fuera de rango'),
            ('B3', '=INDEX(Data.A1:E6, 0, 0)', 'INDEX error - ambos cero'),
            ('B4', '=INDEX(Data.A1:E6, -1, 1)', 'INDEX error - fila negativa'),
            ('B5', '=INDEX(Data.A1:E6, 1, -1)', 'INDEX error - columna negativa'),
            
            # C. INDEX - Casos de Fila/Columna Completa
            ('C1', '=INDEX(Data.A1:E6, 0, 2)', 'INDEX array - columna completa'),
            ('C2', '=INDEX(Data.A1:E6, 2, 0)', 'INDEX array - fila completa'),
            ('C3', '=INDEX(Data.A1:E6, 0, 1)', 'INDEX array - primera columna'),
            
            # NIVEL 2: CASOS INTERMEDIOS
            # D. OFFSET - Casos Fundamentales
            ('D1', '=OFFSET(Data.A1, 1, 1)', 'OFFSET básico - B2'),
            ('D2', '=OFFSET(Data.B2, 1, 1)', 'OFFSET básico - desde B2'),
            ('D3', '=OFFSET(Data.A1, 0, 2)', 'OFFSET básico - horizontal'),
            ('D4', '=OFFSET(Data.A1, 5, 4)', 'OFFSET básico - esquina'),
            
            # E. OFFSET - Casos con Dimensiones
            ('E1', '=OFFSET(Data.A1, 1, 1, 1, 1)', 'OFFSET dimensiones - 1x1'),
            ('E2', '=OFFSET(Data.A1, 1, 1, 2, 2)', 'OFFSET dimensiones - 2x2'),
            ('E3', '=OFFSET(Data.A1, 0, 0, 3, 3)', 'OFFSET dimensiones - 3x3'),
            ('E4', '=OFFSET(Data.A1, 2, 1, 1, 3)', 'OFFSET dimensiones - 1x3'),
            
            # F. OFFSET - Casos de Error
            ('F1', '=OFFSET(Data.A1, -1, 0)', 'OFFSET error - fila negativa'),
            ('F2', '=OFFSET(Data.A1, 0, -1)', 'OFFSET error - columna negativa'),
            ('F3', '=OFFSET(Data.A1, 10, 0)', 'OFFSET error - fuera de hoja'),
            ('F4', '=OFFSET(Data.A1, 0, 10)', 'OFFSET error - fuera de hoja'),
            ('F5', '=OFFSET(Data.A1, 1, 1, 0, 1)', 'OFFSET error - altura cero'),
            ('F6', '=OFFSET(Data.A1, 1, 1, 1, 0)', 'OFFSET error - ancho cero'),
            
            # G. INDIRECT - Casos Fundamentales
            ('G1', '=INDIRECT("Data.B2")', 'INDIRECT básico - valor numérico'),
            ('G2', '=INDIRECT("Data.C3")', 'INDIRECT básico - texto'),
            ('G3', '=INDIRECT("Data.E4")', 'INDIRECT básico - boolean'),
            
            # H. INDIRECT - Referencias Dinámicas
            ('H1', '=INDIRECT("Data.A" & 2)', 'INDIRECT dinámico - concatenación'),
            ('H2', '=INDIRECT("Data." & CHAR(66) & "3")', 'INDIRECT dinámico - CHAR'),
            ('H3', '=INDIRECT("Data.A1:C1")', 'INDIRECT rango - headers'),
            ('H4', '=INDIRECT("Data.A2:A6")', 'INDIRECT rango - columna'),
            
            # I. INDIRECT - Casos de Error
            ('I1', '=INDIRECT("InvalidSheet.A1")', 'INDIRECT error - hoja inexistente'),
            ('I2', '=INDIRECT("Data.Z99")', 'INDIRECT error - celda inválida'),
            ('I3', '=INDIRECT("")', 'INDIRECT error - referencia vacía'),
            ('I4', '=INDIRECT("NotAReference")', 'INDIRECT error - texto inválido'),
            
            # NIVEL 3: CASOS AVANZADOS
            # J. INDEX + INDIRECT
            ('J1', '=INDEX(INDIRECT("Data.A1:E6"), 2, 2)', 'Combinación INDEX+INDIRECT'),
            ('J2', '=INDEX(INDIRECT("Data.A1:E6"), 0, 2)', 'Combinación INDEX+INDIRECT array'),
            ('J3', '=INDEX(INDIRECT("Data.A2:C4"), 2, 3)', 'Combinación INDEX+INDIRECT subrange'),
            
            # K. OFFSET + INDIRECT
            ('K1', '=OFFSET(INDIRECT("Data.A1"), 1, 1)', 'Combinación OFFSET+INDIRECT'),
            ('K2', '=OFFSET(INDIRECT("Data.B2"), 1, 1)', 'Combinación OFFSET+INDIRECT desde B2'),
            
            # L. Combinaciones Complejas
            ('L1', '=INDEX(OFFSET(Data.A1, 0, 0, 3, 3), 2, 2)', 'Combinación INDEX+OFFSET'),
            
            # NIVEL 4: CASOS EDGE
            # M. Rangos Especiales
            ('M1', '=INDEX(Data.A:A, 2)', 'INDEX columna completa'),
            ('M2', '=INDEX(Data.1:1, 1, 2)', 'INDEX fila completa'),
            
            # N. Referencias Complejas
            ('N1', '=INDIRECT("Tests.O1")', 'INDIRECT misma hoja'),
            
            # O. Casos de Compatibilidad
            ('O2', '=IFERROR(INDEX(Data.A1:E6, 10, 1), "Not Found")', 'Manejo errores IFERROR'),
            ('O3', '=IF(ISERROR(OFFSET(Data.A1, -1, 0)), "Error", "OK")', 'Detección errores IF+ISERROR'),
        ]
        
        # Agregar fórmulas una por una con validación
        print(f"📝 Agregando {len(formulas)} fórmulas de prueba...")
        
        for i, (cell, formula, description) in enumerate(formulas, 1):
            try:
                print(f"   {i:2d}/{len(formulas)}: {cell} = {formula}")
                tests_sheet[cell].formula = formula
                
                # Intentar calcular inmediatamente para detectar errores
                calculated_value = tests_sheet[cell].value
                print(f"       ✅ Calculado: {repr(calculated_value)}")
                
            except Exception as e:
                print(f"       ❌ FALLO: {e}")
                print(f"\\n❌ GENERACIÓN FALLÓ en fórmula {i}/{len(formulas)}")
                print(f"   Celda: {cell}")
                print(f"   Fórmula: {formula}")
                print(f"   Descripción: {description}")
                print(f"   Error: {e}")
                raise Exception(f"Fallo en generación de Excel para {cell}: {formula}")
        
        print(f"✅ Todas las fórmulas agregadas exitosamente")
        
        # Forzar cálculo completo
        try:
            wb.app.calculate()
            print("✅ Cálculo completo realizado")
        except Exception as e:
            print(f"⚠️  Advertencia en cálculo: {e}")
        
        # Guardar el archivo
        wb.save(filepath)
        print(f"✅ Excel guardado: {filepath}")
        print(f"✅ {len(formulas)} fórmulas de prueba creadas exitosamente")
        
        # Mostrar resumen
        print("\\n📋 RESUMEN DEL EXCEL GENERADO:")
        print("   - Hoja 'Data': Datos de prueba (6 filas x 6 columnas)")
        print("   - Hoja 'Tests': Casos de prueba organizados por nivel")
        print("   - Nivel 1: Casos estructurales (INDEX básico y errores)")
        print("   - Nivel 2: Casos intermedios (OFFSET e INDIRECT)")
        print("   - Nivel 3: Casos avanzados (combinaciones)")
        print("   - Nivel 4: Casos edge (comportamientos límite)")
        
    except Exception as e:
        print(f"❌ Error en creación del Excel: {e}")
        raise
    finally:
        # Limpiar recursos
        try:
            if 'wb' in locals():
                wb.close()
        except:
            pass
        try:
            app.quit()
        except:
            pass


if __name__ == "__main__":
    output_path = "DYNAMIC_RANGES_COMPREHENSIVE.xlsx"
    print("🚀 Iniciando generación de Excel comprehensivo para rangos dinámicos...")
    print("📋 Este Excel captura el comportamiento FIEL de Excel para:")
    print("   - INDEX: Valores, arrays, errores")
    print("   - OFFSET: Referencias, dimensiones, errores")
    print("   - INDIRECT: Referencias dinámicas, errores")
    print("   - Combinaciones: Funciones anidadas")
    print("   - Edge cases: Comportamientos límite")
    print()
    
    create_comprehensive_dynamic_ranges_excel(output_path)
    print(f"\\n🎉 Excel comprehensivo creado exitosamente: {output_path}")
    print("\\n📋 PRÓXIMOS PASOS:")
    print("1. Copiar el archivo a tests/resources/")
    print("2. Ejecutar tests de integración")
    print("3. Implementar funciones usando red-green-refactor")
    print("4. Validar comportamiento fiel a Excel")