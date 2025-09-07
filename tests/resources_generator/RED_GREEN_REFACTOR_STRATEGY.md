# Estrategia Red-Green-Refactor para Rangos Dinámicos

## Objetivo
Implementar funciones de rangos dinámicos (INDEX, OFFSET, INDIRECT) que representen FIELMENTE el comportamiento de Excel, usando una estrategia estructural → incremental.

## Filosofía de Implementación

### Principio Central: Comportamiento Fiel a Excel
- **Siempre devolver el mismo tipo que Excel devuelve**
- **Siempre devolver el mismo valor que Excel devuelve**
- **Siempre devolver el mismo error que Excel devuelve**
- **No comprometer la fidelidad por conveniencia de implementación**

### Estrategia de Tipos Consistente
```python
# REGLA FUNDAMENTAL: Funciones dinámicas siempre devuelven valores resueltos
INDEX(array, row, col) → Valor de celda o Array (nunca referencia string)
OFFSET(ref, rows, cols) → Valor de celda o Array (nunca referencia string)  
INDIRECT(ref_text) → Valor de celda o Array (nunca referencia string)
```

## Orden de Implementación (Estructural → Incremental)

### FASE 1: Fundamentos Estructurales (INDEX)
**Objetivo**: Establecer la base sólida del comportamiento de arrays

#### 1.1 INDEX - Valores Simples (A1-A5)
```python
def test_level1_index_fundamentals():
    # RED: Fallar tests básicos
    # GREEN: Implementar INDEX básico que devuelve valores
    # REFACTOR: Limpiar implementación
```

**Implementación mínima**:
```python
def INDEX(array, row_num, col_num=None):
    # Validar parámetros
    # Extraer datos del array
    # Devolver valor específico
    return array.values[row_num-1][col_num-1]
```

#### 1.2 INDEX - Manejo de Errores (B1-B5)
```python
def test_level1_index_errors():
    # RED: Fallar con errores apropiados
    # GREEN: Agregar validación de bounds y parámetros
    # REFACTOR: Centralizar manejo de errores
```

**Implementación**:
```python
def INDEX(array, row_num, col_num=None):
    # Validar bounds
    if row_num < 0 or col_num < 0:
        return ValueExcelError("Negative indices not allowed")
    if row_num > len(array.values):
        return RefExcelError("Row out of bounds")
    # ... resto de validaciones
```

#### 1.3 INDEX - Arrays Completos (C1-C3)
```python
def test_level1_index_arrays():
    # RED: Fallar con arrays de fila/columna completa
    # GREEN: Implementar lógica de row=0 y col=0
    # REFACTOR: Optimizar extracción de arrays
```

### FASE 2: Funciones Individuales (OFFSET, INDIRECT)

#### 2.1 OFFSET - Casos Fundamentales (D1-D4)
```python
def test_level2_offset_fundamentals():
    # RED: Fallar tests básicos de OFFSET
    # GREEN: Implementar OFFSET que devuelve valores
    # REFACTOR: Reutilizar lógica de resolución
```

**Implementación**:
```python
def OFFSET(reference, rows, cols, height=None, width=None):
    # Calcular referencia destino
    target_ref = calculate_target_reference(reference, rows, cols)
    
    # SIEMPRE devolver valores, no referencias
    evaluator = get_evaluator_context()
    if height and width:
        return evaluator.get_range_values(target_ref)  # Array
    else:
        return evaluator.get_cell_value(target_ref)    # Valor
```

#### 2.2 OFFSET - Dimensiones (E1-E4)
#### 2.3 OFFSET - Errores (F1-F6)
#### 2.4 INDIRECT - Fundamentales (G1-G3)
#### 2.5 INDIRECT - Dinámico (H1-H4)
#### 2.6 INDIRECT - Errores (I1-I4)

### FASE 3: Combinaciones Avanzadas

#### 3.1 INDEX + INDIRECT (J1-J3)
```python
def test_level3_index_indirect_combinations():
    # RED: Fallar combinaciones
    # GREEN: Asegurar que INDIRECT devuelve Arrays para INDEX
    # REFACTOR: Optimizar interoperabilidad
```

#### 3.2 OFFSET + INDIRECT (K1-K2)
#### 3.3 Combinaciones Complejas (L1)

### FASE 4: Casos Edge

#### 4.1 Rangos Especiales (M1-M2)
#### 4.2 Referencias Complejas (N1)
#### 4.3 Compatibilidad (O2-O3)

## Implementación de Tests Unitarios

### Estrategia Paralela
Por cada test de integración que pase, crear tests unitarios correspondientes:

```python
# Test de integración pasa
def test_level1_index_fundamentals():
    value = self.evaluator.evaluate('Tests!A1')
    self.assertEqual(25, value)

# Crear test unitario correspondiente
def test_index_basic_value():
    array = Array([[1, 2], [3, 4]])
    result = INDEX(array, 2, 2)
    self.assertEqual(4, result)
```

## Criterios de Éxito por Fase

### FASE 1 - Fundamentos
- ✅ INDEX devuelve valores correctos para celdas individuales
- ✅ INDEX maneja errores con tipos correctos (#REF!, #VALUE!)
- ✅ INDEX devuelve Arrays para filas/columnas completas
- ✅ Todos los tipos coinciden con Excel

### FASE 2 - Funciones Individuales  
- ✅ OFFSET devuelve valores de celdas destino
- ✅ OFFSET devuelve Arrays para rangos con dimensiones
- ✅ OFFSET maneja errores apropiadamente
- ✅ INDIRECT resuelve referencias a valores
- ✅ INDIRECT maneja referencias dinámicas
- ✅ INDIRECT maneja errores correctamente

### FASE 3 - Combinaciones
- ✅ INDEX(INDIRECT(...)) funciona correctamente
- ✅ OFFSET(INDIRECT(...)) funciona correctamente
- ✅ Combinaciones complejas funcionan

### FASE 4 - Edge Cases
- ✅ Rangos especiales (A:A, 1:1) funcionan
- ✅ Referencias circulares manejadas
- ✅ Compatibilidad con IFERROR/ISERROR

## Principios de Refactoring

### 1. Nunca Romper Tests Existentes
- Cada refactor debe mantener todos los tests verdes
- Si un refactor rompe tests, revertir y repensar

### 2. Extraer Utilidades Comunes
```python
# Después de implementar INDEX y OFFSET
def _resolve_to_values(reference_or_array, evaluator):
    """Utilidad común para resolver referencias a valores."""
    
def _validate_array_bounds(array, row, col):
    """Utilidad común para validar bounds."""
```

### 3. Mantener Consistencia de Tipos
```python
# SIEMPRE devolver el mismo tipo para el mismo input
INDEX(array, 1, 1) → Number/Text/Boolean (nunca string)
OFFSET(ref, 1, 1) → Number/Text/Boolean (nunca string)
INDIRECT("A1") → Number/Text/Boolean (nunca string)
```

## Manejo de Errores Consistente

### Jerarquía de Errores
1. **ValueExcelError**: Parámetros inválidos (negativos, cero, etc.)
2. **RefExcelError**: Referencias fuera de bounds
3. **NameExcelError**: Referencias inválidas (INDIRECT)

### Patrones de Error
```python
# Parámetros negativos
if row_num < 0:
    return ValueExcelError("Row cannot be negative")

# Fuera de bounds
if row_num > array_height:
    return RefExcelError("Row out of bounds")

# Referencias inválidas
if not is_valid_reference(ref_text):
    return NameExcelError("Invalid reference")
```

## Validación Continua

### Tests de Regresión
- Ejecutar todos los tests después de cada cambio
- Mantener cobertura de tests al 100%
- Validar tipos de retorno en cada test

### Comparación con Excel
- Cada celda del Excel debe coincidir exactamente
- Tipos de datos deben ser idénticos
- Errores deben ser del tipo correcto

## Resultado Final Esperado

### Comportamiento Consistente
```python
# Todas las funciones devuelven valores, nunca referencias
INDEX(array, 2, 2) → 25 (Number)
OFFSET("A1", 1, 1) → 25 (Number) 
INDIRECT("B2") → 25 (Number)

# Arrays cuando corresponde
INDEX(array, 0, 2) → Array([...])
OFFSET("A1", 0, 0, 2, 2) → Array([...])
INDIRECT("A1:B2") → Array([...])

# Errores apropiados
INDEX(array, 10, 1) → RefExcelError
OFFSET("A1", -1, 0) → ValueExcelError
INDIRECT("InvalidRef") → NameExcelError
```

### Interoperabilidad Perfecta
```python
# Combinaciones funcionan naturalmente
INDEX(INDIRECT("A1:E5"), 2, 2) → 25
OFFSET(INDIRECT("A1"), 1, 1) → 25
INDEX(OFFSET("A1", 0, 0, 3, 3), 2, 2) → 25
```

Esta estrategia garantiza que implementemos funciones que representen fielmente el comportamiento de Excel, manteniendo consistencia y predictibilidad en todos los casos de uso.