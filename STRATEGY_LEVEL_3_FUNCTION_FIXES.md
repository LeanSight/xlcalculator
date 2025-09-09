# NIVEL 3: FUNCIONES ESPECÍFICAS - Array Parameters & Error Handling

## Problema A: Array Parameters in OFFSET
```excel
=OFFSET(Data!A1, ROW(A1:A2)-1, 0)
```
Error actual: `ValueExcelError('Row and column offsets must be numbers')`

## Problema B: Error Boundary Validation
```excel
=IF(ISERROR(OFFSET(Data!A1, -1, 0)), "Error", "OK")
```
Resultado actual: `'OK'` (esperado: `'Error'`)

---

## SOLUCIÓN A: Array Parameters Support

### 1. OFFSET Function Enhancement
**Ubicación**: `xlcalculator/xlfunctions/dynamic_range.py`

**Problema**: OFFSET no maneja parámetros array para rows/cols

**Solución**:
```python
def OFFSET(reference, rows, cols, height=None, width=None, *, _context=None):
    # ... código existente para reference parsing ...
    
    # Manejar parámetros array
    if isinstance(rows, func_xltypes.Array):
        # Array de offsets - retornar array de resultados
        results = []
        for row_offset in rows.values:
            if isinstance(row_offset, list):
                row_offset = row_offset[0]  # Extraer valor escalar
            
            # Calcular offset individual
            try:
                offset_ref = start_ref.offset(int(row_offset), int(cols))
                result_value = offset_ref.resolve(evaluator)
                results.append([result_value])
            except Exception:
                results.append([xlerrors.RefExcelError("Invalid offset")])
        
        return func_xltypes.Array(results)
    
    # Código existente para casos escalares
    rows_int = int(rows)
    cols_int = int(cols)
    # ...
```

### 2. Parameter Type Detection
**Ubicación**: `xlcalculator/xlfunctions/dynamic_range.py`

**Solución**:
```python
def _handle_array_parameter(param, param_name):
    """Manejar parámetros que pueden ser escalares o arrays."""
    if isinstance(param, func_xltypes.Array):
        return param.values  # Lista de valores
    elif isinstance(param, (list, tuple)):
        return param
    else:
        return [param]  # Convertir escalar a lista de un elemento
```

---

## SOLUCIÓN B: Error Boundary Validation

### 1. OFFSET Bounds Checking
**Ubicación**: `xlcalculator/xlfunctions/dynamic_range.py`

**Problema**: OFFSET no valida límites correctamente

**Solución**:
```python
def OFFSET(reference, rows, cols, height=None, width=None, *, _context=None):
    # ... código existente ...
    
    # Validar bounds ANTES de calcular offset
    try:
        # Verificar que el offset no vaya fuera de los límites de Excel
        target_row = start_ref.row + rows_int
        target_col = start_ref.column + cols_int
        
        # Excel limits: 1,048,576 rows × 16,384 columns
        if target_row < 1 or target_row > 1048576:
            raise xlerrors.RefExcelError("Row offset out of bounds")
        if target_col < 1 or target_col > 16384:
            raise xlerrors.RefExcelError("Column offset out of bounds")
        
        offset_ref = start_ref.offset(rows_int, cols_int)
    except Exception as e:
        # Asegurar que errores de bounds se propaguen correctamente
        raise xlerrors.RefExcelError("Offset results in invalid reference")
```

### 2. Error Propagation Testing
**Ubicación**: Test validation

**Verificación**:
```python
# Test que OFFSET(-1, 0) efectivamente lance error
def test_offset_negative_bounds():
    result = evaluator.evaluate('OFFSET(Data!A1, -1, 0)')
    assert isinstance(result, xlerrors.RefExcelError)
    
    # Test que ISERROR detecte el error correctamente
    result = evaluator.evaluate('ISERROR(OFFSET(Data!A1, -1, 0))')
    assert result == True
```

---

## Estimación de Esfuerzo

### Solución A: Array Parameters
- **Complejidad**: MEDIA
- **Tiempo estimado**: 4-6 horas
- **Archivos afectados**: 1 archivo
- **Riesgo de regresión**: BAJO

### Solución B: Error Validation
- **Complejidad**: BAJA
- **Tiempo estimado**: 2-3 horas
- **Archivos afectados**: 1 archivo
- **Riesgo de regresión**: MUY BAJO

## Dependencias
- Ninguna dependencia externa
- Cambios aislados en funciones específicas
- Testing directo y simple

## Prioridad
🟢 **ALTA** - Fixes rápidos con alto impacto en compatibilidad