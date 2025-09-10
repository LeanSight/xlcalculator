# NIVEL 2: SISTEMA DE REFERENCIAS - Full Column/Row References

## Problema
```excel
=OFFSET(Data!A:A, 1, 0, 3, 1)
```
Error actual: `RefExcelError('Cannot find cell containing value: Data!A:A')`

## Análisis del Sistema de Referencias

### 1. Reference Objects Enhancement
**Ubicación**: `xlcalculator/reference_objects.py`

**Problema**: El sistema no distingue entre referencias de celda y referencias completas

**Solución**:
```python
class FullColumnReference(ReferenceBase):
    def __init__(self, sheet: str, column: str):
        self.sheet = sheet
        self.column = column
    
    def resolve(self, evaluator):
        # Retornar toda la columna como Array
        # Implementar límites de Excel (1,048,576 filas)
        pass

class FullRowReference(ReferenceBase):
    def __init__(self, sheet: str, row: int):
        self.sheet = sheet
        self.row = row
    
    def resolve(self, evaluator):
        # Retornar toda la fila como Array
        # Implementar límites de Excel (16,384 columnas)
        pass
```

### 2. Reference Parser Enhancement
**Ubicación**: `xlcalculator/reference_objects.py`

**Problema**: El parser no reconoce patrones como `A:A`, `1:1`

**Solución**:
```python
@classmethod
def parse(cls, ref_string: str):
    # Detectar patrones de referencia completa
    if re.match(r'^[A-Z]+:[A-Z]+$', ref_string):
        # Columna completa: A:A, B:B
        return FullColumnReference.parse(ref_string)
    elif re.match(r'^\d+:\d+$', ref_string):
        # Fila completa: 1:1, 2:2
        return FullRowReference.parse(ref_string)
    # ... resto del código existente
```

### 3. OFFSET Function Fix
**Ubicación**: `xlcalculator/xlfunctions/dynamic_range.py`

**Problema**: OFFSET no reconoce referencias completas como válidas

**Solución**:
```python
def OFFSET(reference, rows, cols, height=None, width=None, *, _context=None):
    # Mejorar detección de tipos de referencia
    if isinstance(reference, (str, func_xltypes.Text)):
        ref_string = str(reference)
        
        # Detectar referencias completas ANTES de buscar valores
        if _is_full_column_or_row_reference(ref_string):
            # Parsear como referencia completa
            start_ref = parse_full_reference(ref_string)
        else:
            # Intentar como referencia de celda normal
            try:
                start_ref = CellReference.parse(ref_string)
            except xlerrors.RefExcelError:
                # Buscar como valor (código existente)
                found_address = _find_cell_address_for_value(ref_string, evaluator)
                # ...
```

## Estimación de Esfuerzo
- **Complejidad**: MEDIA-ALTA
- **Tiempo estimado**: 1-2 días
- **Archivos afectados**: 2-3 archivos
- **Riesgo de regresión**: MEDIO

## Dependencias
- Extensión del sistema de referencias
- Nuevos tipos de referencia
- Validación de límites de Excel

## Prioridad
🟡 **MEDIA** - Funcionalidad común en Excel, moderadamente importante