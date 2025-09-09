# Reference Objects Design Document

**Document Version**: 1.0  
**Date**: 2025-09-09  
**Phase**: ATDD Red Phase - Design Document  
**Application**: Reference Object System for xlcalculator Excel compliance

---

## 🎯 Objetivo

Crear un sistema de objetos de referencia que capture FIELMENTE el comportamiento de Excel para referencias de celdas, rangos y nombres, organizando los casos de más estructurales a menos estructurales siguiendo la metodología ATDD.

## 📋 Estructura del Sistema de Referencias

### Componente 1: "CellReference" - Referencias de Celdas Individuales
**Objetivo**: Manejar referencias a celdas individuales con coordenadas exactas de Excel

### Componente 2: "RangeReference" - Referencias de Rangos
**Objetivo**: Manejar referencias a rangos de celdas con aritmética de referencias

### Componente 3: "NamedReference" - Referencias con Nombre
**Objetivo**: Manejar rangos con nombre y resolución dinámica

### Componente 4: "ReferenceParser" - Análisis de Cadenas de Referencia
**Objetivo**: Convertir cadenas de texto a objetos de referencia

## 📊 NIVEL 1: CASOS ESTRUCTURALES (Comportamiento Core)

### Grupo 1.1: CellReference - Parsing Básico
**Objetivo**: Verificar que las referencias de celdas se analizan correctamente

| Caso | Entrada | Resultado Esperado | Descripción |
|------|---------|-------------------|-------------|
| CR01 | "A1" | CellReference(sheet="", row=1, col=1) | Referencia básica |
| CR02 | "Z1" | CellReference(sheet="", row=1, col=26) | Columna Z |
| CR03 | "AA1" | CellReference(sheet="", row=1, col=27) | Columna doble letra |
| CR04 | "A100" | CellReference(sheet="", row=100, col=1) | Fila alta |
| CR05 | "$A$1" | CellReference(sheet="", row=1, col=1, abs_row=True, abs_col=True) | Referencia absoluta |

### Grupo 1.2: CellReference - Referencias con Hoja
**Objetivo**: Verificar referencias entre hojas

| Caso | Entrada | Resultado Esperado | Descripción |
|------|---------|-------------------|-------------|
| CR06 | "Sheet1!A1" | CellReference(sheet="Sheet1", row=1, col=1) | Referencia con hoja |
| CR07 | "'Sheet 2'!A1" | CellReference(sheet="Sheet 2", row=1, col=1) | Hoja con espacios |
| CR08 | "Data!$B$5" | CellReference(sheet="Data", row=5, col=2, abs_row=True, abs_col=True) | Hoja + absoluta |

### Grupo 1.3: CellReference - Aritmética de Referencias
**Objetivo**: Verificar operaciones de offset en referencias

| Caso | Referencia Base | Offset (rows, cols) | Resultado Esperado | Descripción |
|------|-----------------|--------------------|--------------------|-------------|
| CR09 | A1 | (1, 1) | B2 | Offset básico |
| CR10 | A1 | (0, 0) | A1 | Sin offset |
| CR11 | B2 | (-1, -1) | A1 | Offset negativo |
| CR12 | A1 | (0, 25) | Z1 | Offset columna Z |

### Grupo 1.4: RangeReference - Parsing de Rangos
**Objetivo**: Verificar análisis de rangos de celdas

| Caso | Entrada | Resultado Esperado | Descripción |
|------|---------|-------------------|-------------|
| RR01 | "A1:B2" | RangeReference(A1, B2) | Rango básico |
| RR02 | "A1:A3" | RangeReference(A1, A3) | Rango columna |
| RR03 | "A1:C1" | RangeReference(A1, C1) | Rango fila |
| RR04 | "Sheet1!A1:B2" | RangeReference(Sheet1!A1, Sheet1!B2) | Rango con hoja |

### Grupo 1.5: RangeReference - Operaciones de Rango
**Objetivo**: Verificar operaciones en rangos

| Caso | Rango Base | Operación | Resultado Esperado | Descripción |
|------|------------|-----------|-------------------|-------------|
| RR05 | A1:B2 | offset(1,1) | B2:C3 | Offset de rango |
| RR06 | A1:B2 | resize(3,3) | A1:C3 | Redimensionar |
| RR07 | A1:B2 | get_cell(0,0) | A1 | Celda en rango |
| RR08 | A1:B2 | dimensions() | (2,2) | Dimensiones |

## 📊 NIVEL 2: CASOS INTERMEDIOS (Integración con Funciones)

### Grupo 2.1: ROW Function - Referencias de Cadena
**Objetivo**: Verificar que ROW acepta referencias como cadenas

| Caso | Fórmula | Resultado Esperado | Descripción |
|------|---------|-------------------|-------------|
| RF01 | ROW("A1") | 1 | Cadena de referencia |
| RF02 | ROW("A100") | 100 | Fila alta |
| RF03 | ROW("Sheet1!A5") | 5 | Con hoja |
| RF04 | ROW("Z1") | 1 | Columna Z |

### Grupo 2.2: COLUMN Function - Referencias de Cadena
**Objetivo**: Verificar que COLUMN acepta referencias como cadenas

| Caso | Fórmula | Resultado Esperado | Descripción |
|------|---------|-------------------|-------------|
| CF01 | COLUMN("A1") | 1 | Cadena de referencia |
| CF02 | COLUMN("Z1") | 26 | Columna Z |
| CF03 | COLUMN("AA1") | 27 | Columna doble |
| CF04 | COLUMN("Sheet1!B1") | 2 | Con hoja |

### Grupo 2.3: OFFSET Function - Aritmética de Referencias
**Objetivo**: Verificar que OFFSET funciona con cualquier referencia

| Caso | Fórmula | Resultado Esperado | Descripción |
|------|---------|-------------------|-------------|
| OF01 | OFFSET("A1", 1, 1) | Valor en B2 | Offset básico |
| OF02 | OFFSET("B2", -1, -1) | Valor en A1 | Offset negativo |
| OF03 | OFFSET("A1", 0, 0, 2, 2) | Rango A1:B2 | Con dimensiones |
| OF04 | OFFSET("Sheet1!A1", 1, 0) | Valor en Sheet1!A2 | Con hoja |

### Grupo 2.4: INDIRECT Function - Referencias Dinámicas
**Objetivo**: Verificar resolución dinámica de referencias

| Caso | Fórmula | Resultado Esperado | Descripción |
|------|---------|-------------------|-------------|
| IF01 | INDIRECT("A1") | Valor en A1 | Referencia directa |
| IF02 | INDIRECT("A" & "1") | Valor en A1 | Referencia calculada |
| IF03 | INDIRECT("Sheet1!A1") | Valor en Sheet1!A1 | Con hoja |
| IF04 | INDIRECT("A1:B2") | Rango A1:B2 | Rango dinámico |

## 📊 NIVEL 3: CASOS AVANZADOS (Contexto y Evaluación)

### Grupo 3.1: Context-Aware Functions
**Objetivo**: Verificar funciones que usan contexto de celda actual

| Caso | Celda Actual | Fórmula | Resultado Esperado | Descripción |
|------|--------------|---------|-------------------|-------------|
| CA01 | A1 | ROW() | 1 | Fila actual |
| CA02 | B5 | ROW() | 5 | Fila actual |
| CA03 | A1 | COLUMN() | 1 | Columna actual |
| CA04 | Z1 | COLUMN() | 26 | Columna actual |

### Grupo 3.2: Lazy Reference Resolution
**Objetivo**: Verificar que las referencias se resuelven solo cuando es necesario

| Caso | Referencia | Operación | Resultado Esperado | Descripción |
|------|------------|-----------|-------------------|-------------|
| LR01 | CellReference("A1") | .address | "A1" | Dirección sin resolver |
| LR02 | CellReference("A1") | .resolve(evaluator) | Valor real | Resolución lazy |
| LR03 | RangeReference("A1:B2") | .address | "A1:B2" | Rango sin resolver |
| LR04 | RangeReference("A1:B2") | .resolve(evaluator) | Array 2x2 | Resolución lazy |

## 📊 NIVEL 4: CASOS DE ERROR (Manejo de Errores Excel)

### Grupo 4.1: Reference Parsing Errors
**Objetivo**: Verificar errores de análisis de referencias

| Caso | Entrada | Error Esperado | Descripción |
|------|---------|----------------|-------------|
| PE01 | "InvalidRef" | #REF! | Referencia inválida |
| PE02 | "" | #REF! | Cadena vacía |
| PE03 | "A" | #REF! | Referencia incompleta |
| PE04 | "1A" | #REF! | Formato incorrecto |

### Grupo 4.2: Bounds Checking Errors
**Objetivo**: Verificar errores de límites de Excel

| Caso | Operación | Error Esperado | Descripción |
|------|-----------|----------------|-------------|
| BE01 | OFFSET("A1", -1, 0) | #REF! | Fila fuera de límites |
| BE02 | OFFSET("A1", 0, -1) | #REF! | Columna fuera de límites |
| BE03 | OFFSET("A1", 1048577, 0) | #REF! | Fila máxima excedida |
| BE04 | OFFSET("A1", 0, 16385) | #REF! | Columna máxima excedida |

### Grupo 4.3: Function Parameter Errors
**Objetivo**: Verificar errores de parámetros de función

| Caso | Fórmula | Error Esperado | Descripción |
|------|---------|----------------|-------------|
| FE01 | OFFSET("A1", 0, 0, 0, 1) | #VALUE! | Height = 0 |
| FE02 | OFFSET("A1", 0, 0, 1, 0) | #VALUE! | Width = 0 |
| FE03 | INDIRECT("Sheet99!A1") | #REF! | Hoja inexistente |
| FE04 | ROW("InvalidRef") | #REF! | Referencia inválida |

## 🔧 Arquitectura de Implementación

### Componentes Principales

#### 1. Reference Objects
```python
@dataclass
class CellReference:
    sheet: str
    row: int
    column: int
    absolute_row: bool = False
    absolute_column: bool = False
    
    def offset(self, rows: int, cols: int) -> 'CellReference'
    def resolve(self, evaluator) -> Any
    @classmethod
    def parse(cls, address: str) -> 'CellReference'
```

#### 2. AST Integration
```python
# Modificar ast_nodes.py para detectar funciones que necesitan referencias
def _eval_parameter_with_reference_support(self, param, evaluator):
    if function_needs_reference_strings(current_function):
        return param.value  # Pasar cadena sin evaluar
    else:
        return param.eval(evaluator)  # Evaluación normal
```

#### 3. Function Enhancement
```python
@xl.register()
@xl.validate_args
def ROW(reference=None, *, _context: CellContext = None) -> int:
    if reference is None:
        return _context.current_row
    ref = CellReference.parse(reference, _context.current_sheet)
    return ref.row
```

## 📋 Plan de Implementación ATDD

### Fase RED (Tests Failing)
1. Crear JSON test cases con todos los casos definidos
2. Generar tests de aceptación que fallen
3. Verificar que todos los tests fallan por las razones correctas

### Fase GREEN (Minimal Implementation)
1. Implementar CellReference con parsing básico
2. Implementar RangeReference con operaciones básicas
3. Modificar AST para pasar cadenas de referencia
4. Actualizar funciones ROW, COLUMN, OFFSET, INDIRECT

### Fase REFACTOR (Code Improvement)
1. Optimizar parsing de referencias
2. Implementar lazy evaluation
3. Mejorar manejo de errores
4. Optimizar rendimiento

## 🎯 Criterios de Éxito

### Funcionales
- ✅ Todos los 67+ casos de prueba pasan
- ✅ Comportamiento idéntico a Excel
- ✅ Manejo correcto de errores Excel
- ✅ Soporte completo para referencias dinámicas

### Técnicos
- ✅ Integración sin regresiones
- ✅ Rendimiento ≤10% overhead
- ✅ Arquitectura extensible
- ✅ Código auto-documentado

---

**Siguiente Fase**: Crear casos de prueba JSON basados en este diseño estructurado.