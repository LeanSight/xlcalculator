# Diseño Comprehensivo de Excel para Rangos Dinámicos

## Objetivo
Crear un Excel que capture FIELMENTE el comportamiento de Excel para todas las funciones de rangos dinámicos, organizando los casos de más estructurales a menos estructurales.

## Estructura del Excel

### Hoja 1: "Data" - Datos de Prueba
```
    A       B       C       D       E       F
1   Name    Age     City    Score   Active  Notes
2   Alice   25      NYC     85      TRUE    Good
3   Bob     30      LA      92      FALSE   Great
4   Charlie 35      Chicago 78      TRUE    OK
5   Diana   28      Miami   95      TRUE    Excellent
6   Eve     22      Boston  88      FALSE   Average
```

### Hoja 2: "Tests" - Casos de Prueba Organizados

## NIVEL 1: CASOS ESTRUCTURALES (Comportamiento Core)

### A. INDEX - Casos Fundamentales (Valores Individuales)
```
A1: =INDEX(Data!A1:E6, 2, 2)     → 25 (valor numérico)
A2: =INDEX(Data!A1:E6, 3, 1)     → "Bob" (texto)
A3: =INDEX(Data!A1:E6, 4, 5)     → TRUE (boolean)
A4: =INDEX(Data!A1:E6, 6, 1)     → "Eve" (última fila)
A5: =INDEX(Data!A1:E6, 1, 5)     → TRUE (primera fila)
```

### B. INDEX - Arrays Completos (row=0 o col=0)
```
B1: =INDEX(Data!A1:E6, 0, 2)     → Array completo columna Age
B2: =INDEX(Data!A1:E6, 2, 0)     → Array completo fila Alice
B3: =INDEX(Data!A1:E6, 0, 1)     → Array completo columna Name
B4: =INDEX(Data!A1:E6, 0, 5)     → Array completo columna Active
```

### C. INDEX - Casos de Error Estructurales
```
C1: =INDEX(Data!A1:E6, 7, 1)     → #REF! (fila fuera de rango)
C2: =INDEX(Data!A1:E6, 1, 7)     → #REF! (columna fuera de rango)
C3: =INDEX(Data!A1:E6, 0, 0)     → #VALUE! (ambos cero)
C4: =INDEX(Data!A1:E6, -1, 1)    → #VALUE! (fila negativa)
C5: =INDEX(Data!A1:E6, 1, -1)    → #VALUE! (columna negativa)
```

## NIVEL 2: CASOS INTERMEDIOS (Funciones Individuales)

### D. OFFSET - Casos Fundamentales (Valores Individuales)
```
D1: =OFFSET(Data!A1, 1, 1)       → 25 (B2)
D2: =OFFSET(Data!B2, 1, 1)       → "LA" (C3)
D3: =OFFSET(Data!A1, 0, 2)       → "City" (C1)
D4: =OFFSET(Data!A1, 5, 4)       → FALSE (E6)
D5: =OFFSET(Data!C3, -1, 1)      → 30 (D2)
```

### E. OFFSET - Arrays con Dimensiones
```
E1: =OFFSET(Data!A1, 1, 1, 1, 1) → 25 (B2 como array 1x1)
E2: =OFFSET(Data!A1, 1, 1, 2, 2) → Array 2x2 desde B2
E3: =OFFSET(Data!A1, 0, 0, 3, 3) → Array 3x3 desde A1
E4: =OFFSET(Data!A1, 2, 1, 1, 3) → Array 1x3 desde B3
E5: =OFFSET(Data!A1, 1, 0, 3, 1) → Array 3x1 desde A2
```

### F. OFFSET - Casos de Error
```
F1: =OFFSET(Data!A1, -2, 0)      → #REF! (antes del inicio de hoja)
F2: =OFFSET(Data!A1, 0, -2)      → #REF! (antes del inicio de hoja)
F3: =OFFSET(Data!A1, 100, 0)     → #REF! (más allá de hoja)
F4: =OFFSET(Data!A1, 0, 100)     → #REF! (más allá de hoja)
F5: =OFFSET(Data!A1, 1, 1, 0, 1) → #VALUE! (altura cero)
F6: =OFFSET(Data!A1, 1, 1, 1, 0) → #VALUE! (ancho cero)
```

### G. INDIRECT - Casos Fundamentales (Valores Individuales)
```
G1: =INDIRECT("Data!B2")         → 25 (valor numérico)
G2: =INDIRECT("Data!C3")         → "LA" (texto)
G3: =INDIRECT("Data!E4")         → TRUE (boolean)
G4: =INDIRECT(P1)                → 25 (desde celda con "Data!B2")
```

### H. INDIRECT - Referencias Dinámicas
```
H1: =INDIRECT("Data!A" & 2)      → "Alice" (concatenación)
H2: =INDIRECT("Data!" & CHAR(66) & "3") → 30 (usando CHAR)
H3: =INDIRECT("Data!A" & ROW())  → Referencia dinámica por fila actual
H4: =INDIRECT("Data!" & CHAR(65+COLUMN()) & "1") → Referencia por columna
```

### I. INDIRECT - Arrays de Referencias
```
I1: =INDIRECT("Data!A1:C1")      → Array de headers (3 elementos)
I2: =INDIRECT("Data!A2:A6")      → Array de nombres (5 elementos)
I3: =INDIRECT("Data!B1:B6")      → Array de edad (6 elementos)
I4: =INDIRECT(P3)                → Array desde celda (A1:C3)
```

### J. INDIRECT - Referencias de Columna/Fila Completa
```
J1: =INDIRECT("Data!A:A")        → Columna completa A
J2: =INDIRECT("Data!B:B")        → Columna completa B
J3: =INDIRECT("Data!1:1")        → Fila completa 1
J4: =INDIRECT("Data!2:2")        → Fila completa 2
```

### K. INDIRECT - Casos de Error
```
K1: =INDIRECT("InvalidSheet!A1") → #REF! (hoja inexistente)
K2: =INDIRECT("Data!Z99")        → Valor de celda vacía o 0
K3: =INDIRECT("")                → #REF! (referencia vacía)
K4: =INDIRECT("NotAReference")   → #REF! (texto inválido)
K5: =INDIRECT(P4)                → #REF! (hoja inválida desde celda)
```

## NIVEL 3: CASOS AVANZADOS (Combinaciones)

### L. INDEX + INDIRECT
```
L1: =INDEX(INDIRECT("Data!A1:E6"), 2, 2)    → 25
L2: =INDEX(INDIRECT("Data!A1:E6"), 0, 2)    → Array columna Age
L3: =INDEX(INDIRECT("Data!A2:C4"), 2, 3)    → "Chicago"
L4: =INDEX(INDIRECT("Data!A:A"), 3)         → "Bob"
```

### M. OFFSET + INDIRECT
```
M1: =OFFSET(INDIRECT("Data!A1"), 1, 1)      → 25
M2: =OFFSET(INDIRECT("Data!B2"), 1, 1)      → "LA"
M3: =OFFSET(INDIRECT("Data!A1"), 1, 1, 2, 2) → Array 2x2
```

### N. Combinaciones Complejas
```
N1: =INDEX(OFFSET(Data!A1, 0, 0, 3, 3), 2, 2)     → 25
N2: =OFFSET(INDEX(Data!A1:E6, 2, 1), 1, 1)        → 30
N3: =INDIRECT("Data!" & "A" & INDEX(Data!B1:B6, 2, 1)) → Ref dinámica
```

## NIVEL 4: CASOS DE CONTEXTO (Uso con Otras Funciones)

### O. Funciones con Agregación
```
O1: =SUM(INDEX(Data!A1:E6, 0, 2))            → Suma columna Age
O2: =AVERAGE(OFFSET(Data!B1, 1, 0, 5, 1))    → Promedio de edades
O3: =COUNT(INDIRECT("Data!B:B"))              → Contar números en col B
O4: =MAX(INDEX(Data!A1:E6, 0, 4))            → Máximo de scores
```

### P. Manejo de Errores
```
P1: =IFERROR(INDEX(Data!A1:E6, 10, 1), "Not Found") → Manejo con IFERROR
P2: =IF(ISERROR(OFFSET(Data!A1, -1, 0)), "Error", "OK") → Detección errores
P3: =IFERROR(INDIRECT("InvalidSheet!A1"), "Sheet Error") → Error de hoja
```

## NIVEL 5: CASOS AVANZADOS (Edge Cases)

### Q. Referencias Especiales
```
Q1: =INDIRECT("Tests!O1")        → Referencia misma hoja
Q2: =INDEX(Data!A:A, 2)          → INDEX con columna completa
Q3: =OFFSET(Data!A:A, 1, 0, 3, 1) → OFFSET con columna completa
```

### R. Arrays Dinámicos (Excel 365)
```
R1: =INDEX(Data!A1:E6, ROW(A1:A3), 1)       → Array con múltiples filas
R2: =OFFSET(Data!A1, ROW(A1:A2)-1, 0)       → Array con offset múltiple
```

### S. Forma de Referencia vs Array
```
S1: =INDEX((Data!A1:A5, Data!C1:C5), 2, 1, 1) → Forma referencia área 1
S2: =INDEX((Data!A1:A5, Data!C1:C5), 2, 1, 2) → Forma referencia área 2
```

## Datos de Referencia para Validación

### Referencias Auxiliares (Columna P)
```
P1: "Data!B2"     → Para INDIRECT básico
P2: "Data!C3"     → Para INDIRECT texto
P3: "Data!A1:C3"  → Para INDIRECT con rango
P4: "InvalidSheet!A1" → Para error testing
P5: ""            → Para error testing vacío
P6: "Data!A:A"    → Para columna completa
P7: "Data!1:1"    → Para fila completa
```

### Valores Esperados (Columna Q)
```
Q1: 25            → Valor esperado para A1
Q2: "Bob"         → Valor esperado para A2
Q3: TRUE          → Valor esperado para A3
Q4: "#REF!"       → Error esperado para C1
Q5: "#VALUE!"     → Error esperado para C3
```

## Estrategia de Testing

### Orden de Implementación (Estructural → Incremental)
1. **INDEX valores individuales** (A1-A5): Valores simples
2. **INDEX arrays** (B1-B4): Filas/columnas completas
3. **INDEX errores** (C1-C5): Manejo de errores
4. **OFFSET valores** (D1-D5): Referencias simples
5. **OFFSET arrays** (E1-E5): Rangos con dimensiones
6. **OFFSET errores** (F1-F6): Casos límite
7. **INDIRECT valores** (G1-G4): Referencias directas
8. **INDIRECT dinámico** (H1-H4): Referencias construidas
9. **INDIRECT arrays** (I1-I4): Rangos y arrays
10. **INDIRECT columnas** (J1-J4): Referencias completas
11. **INDIRECT errores** (K1-K5): Casos inválidos
12. **Combinaciones** (L1-N3): Funciones anidadas
13. **Contexto** (O1-P3): Uso con agregación/errores
14. **Avanzados** (Q1-S2): Edge cases y formas especiales

### Criterios de Éxito
- Cada celda debe devolver exactamente el mismo valor/error que Excel
- Los tipos de datos deben coincidir (Number, Text, Boolean, Array, Error)
- Los arrays deben tener las mismas dimensiones y valores
- Los errores deben ser del tipo correcto (#REF!, #VALUE!, #NAME!)

### Comportamientos Clave Validados
- **INDEX**: Valores vs arrays según parámetros row=0/col=0
- **OFFSET**: Valores individuales vs arrays con height/width
- **INDIRECT**: Conversión texto→referencia→valor/array
- **Combinaciones**: Compatibilidad entre funciones anidadas
- **Errores**: Propagación correcta de tipos de error

### Compatibilidad
- **Excel 365**: Dynamic array spilling
- **Excel Legacy**: Array formulas con CSE
- **Todas las versiones**: Comportamiento básico consistente