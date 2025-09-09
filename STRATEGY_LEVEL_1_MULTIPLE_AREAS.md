# NIVEL 1: ARQUITECTURAL - Multiple Areas Support

## Problema
```excel
=INDEX((Data!A1:A5, Data!C1:C5), 2, 1, 2)
```
Error actual: `RefExcelError('Row index 2 out of range (1-1)')`

## Análisis Arquitectural

### 1. Parser/Tokenizer Changes
**Ubicación**: `xlcalculator/tokenizer.py`, `xlcalculator/parser.py`

**Problema**: El parser no reconoce correctamente las tuplas de rangos `(range1, range2)`

**Solución**:
```python
# Nuevo tipo de nodo AST para múltiples áreas
class MultipleAreasNode(ASTNode):
    def __init__(self, areas: List[RangeNode]):
        self.areas = areas
    
    def eval(self, context):
        # Evaluar cada área y retornar lista de rangos
        return [area.eval(context) for area in self.areas]
```

### 2. Function Parameter Handling
**Ubicación**: `xlcalculator/ast_nodes.py`

**Problema**: Los parámetros de función no manejan estructuras complejas

**Solución**:
```python
def _eval_parameter_with_excel_fallback(self, param, context, func_name, param_index):
    # Detectar si el parámetro es una tupla de áreas
    if isinstance(param, MultipleAreasNode):
        return param.eval(context)  # Lista de rangos evaluados
    # ... resto del código existente
```

### 3. INDEX Function Enhancement
**Ubicación**: `xlcalculator/xlfunctions/dynamic_range.py`

**Problema**: INDEX no maneja correctamente múltiples áreas

**Solución**:
```python
def INDEX(array, row_num, col_num=1, area_num=1, *, _context=None):
    # Detectar múltiples áreas
    if isinstance(array, list) and all(isinstance(area, str) for area in array):
        # Es una lista de referencias de rango
        areas = array
        
        # Validar area_num
        if area_num < 1 or area_num > len(areas):
            raise xlerrors.RefExcelError("Area number out of range")
        
        # Seleccionar el área específica
        selected_area = areas[area_num - 1]
        array_data = evaluator.get_range_values(selected_area)
    # ... resto del código
```

## Estimación de Esfuerzo
- **Complejidad**: ALTA
- **Tiempo estimado**: 2-3 días
- **Archivos afectados**: 4-5 archivos core
- **Riesgo de regresión**: ALTO

## Dependencias
- Cambios en parser/tokenizer
- Nuevos tipos de nodos AST
- Modificaciones en evaluador de parámetros
- Testing extensivo requerido

## Prioridad
🔴 **BAJA** - Funcionalidad avanzada, no crítica para uso básico