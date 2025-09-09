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

### Hoja 2: "Tests" - Casos de Prueba Organizados (76 Test Cases)

## NIVEL 1: CASOS ESTRUCTURALES (Comportamiento Core) - 14 Cases

### 1A. INDEX - Casos Fundamentales (Valores Individuales) - 5 Cases
```
A1: =INDEX(Data!A1:E6, 2, 2)     → 25 (valor numérico)
A2: =INDEX(Data!A1:E6, 3, 1)     → "Bob" (texto)
A3: =INDEX(Data!A1:E6, 4, 5)     → TRUE (boolean)
A4: =INDEX(Data!A1:E6, 6, 1)     → "Eve" (última fila)
A5: =INDEX(Data!A1:E6, 1, 5)     → "Active" (primera fila)
```

### 1B. INDEX - Arrays Completos (row=0 o col=0) - 4 Cases
```
B1: =INDEX(Data!A1:E6, 0, 2)     → Array completo columna Age
B2: =INDEX(Data!A1:E6, 2, 0)     → Array completo fila Alice
B3: =INDEX(Data!A1:E6, 0, 1)     → Array completo columna Name
B4: =INDEX(Data!A1:E6, 0, 5)     → Array completo columna Active
```

### 1C. INDEX - Casos de Error Estructurales - 5 Cases
```
C1: =INDEX(Data!A1:E6, 7, 1)     → #REF! (fila fuera de rango)
C2: =INDEX(Data!A1:E6, 1, 7)     → #REF! (columna fuera de rango)
C3: =INDEX(Data!A1:E6, 0, 0)     → #VALUE! (ambos cero)
C4: =INDEX(Data!A1:E6, -1, 1)    → #VALUE! (fila negativa)
C5: =INDEX(Data!A1:E6, 1, -1)    → #VALUE! (columna negativa)
```

## NIVEL 2: CASOS INTERMEDIOS (Funciones Individuales) - 37 Cases

### 2D. OFFSET - Casos Fundamentales (Valores Individuales) - 5 Cases
```
D1: =OFFSET(Data!A1, 1, 1)       → 25 (B2)
D2: =OFFSET(Data!B2, 1, 1)       → "LA" (C3)
D3: =OFFSET(Data!A1, 0, 2)       → "City" (C1)
D4: =OFFSET(Data!A1, 5, 4)       → FALSE (E6)
D5: =OFFSET(Data!C3, -1, 1)      → 30 (D2)
```

### 2E. OFFSET - Arrays con Dimensiones - 5 Cases
```
E1: =OFFSET(Data!A1, 1, 1, 1, 1) → 25 (B2 como array 1x1)
E2: =OFFSET(Data!A1, 1, 1, 2, 2) → Array 2x2 desde B2
E3: =OFFSET(Data!A1, 0, 0, 3, 3) → Array 3x3 desde A1
E4: =OFFSET(Data!A1, 2, 1, 1, 3) → Array 1x3 desde B3
E5: =OFFSET(Data!A1, 1, 0, 3, 1) → Array 3x1 desde A2
```

### 2F. OFFSET - Casos de Error - 6 Cases
```
F1: =OFFSET(Data!A1, -2, 0)      → #REF! (antes del inicio de hoja)
F2: =OFFSET(Data!A1, 0, -2)      → #REF! (antes del inicio de hoja)
F3: =OFFSET(Data!A1, 100, 0)     → #REF! (más allá de hoja)
F4: =OFFSET(Data!A1, 0, 100)     → #REF! (más allá de hoja)
F5: =OFFSET(Data!A1, 1, 1, 0, 1) → #VALUE! (altura cero)
F6: =OFFSET(Data!A1, 1, 1, 1, 0) → #VALUE! (ancho cero)
```

### 2G. INDIRECT - Casos Fundamentales (Valores Individuales) - 4 Cases
```
G1: =INDIRECT("Data!B2")         → 25 (valor numérico)
G2: =INDIRECT("Data!C3")         → "LA" (texto)
G3: =INDIRECT("Data!E4")         → TRUE (boolean)
G4: =INDIRECT(P1)                → 25 (desde celda con "Data!B2")
```

### 2H. INDIRECT - Referencias Dinámicas - 4 Cases
```
H1: =INDIRECT("Data!A" & 2)      → "Alice" (concatenación)
H2: =INDIRECT("Data!" & CHAR(66) & "3") → 30 (usando CHAR)
H3: =INDIRECT("Data!A" & ROW())  → "Bob" (referencia dinámica por fila actual - ROW() from H3 returns 3, Data!A3="Bob")
H4: =INDIRECT("Data!" & CHAR(65+COLUMN()) & "1") → 0 (referencia por columna - COLUMN() from H4 returns 8, CHAR(73)="I", Data!I1=0)
```

**CORRECCIÓN IMPORTANTE - Comportamiento Oficial de Excel:**
- **ROW()**: Retorna el número de fila de la celda donde aparece la función
  - ROW() en H3 retorna 3 (no 4)
  - Por tanto: "Data!A" & ROW() = "Data!A3" = "Bob"
- **COLUMN()**: Retorna el número de columna de la celda donde aparece la función  
  - COLUMN() en H4 retorna 8 (columna H)
  - Por tanto: CHAR(65+8) = CHAR(73) = "I", "Data!I1" = 0

**Fuente**: [Microsoft Excel ROW Function Documentation](https://support.microsoft.com/en-us/office/row-function-3a63b74a-c4d0-4093-b49a-e76eb49a6d8d)
**Fuente**: [Microsoft Excel COLUMN Function Documentation](https://support.microsoft.com/en-us/office/column-function-44e8c754-711c-4df3-9da4-47a55042554b)

### 2I. INDIRECT - Arrays de Referencias - 4 Cases
```
I1: =INDIRECT("Data!A1:C1")      → Array de headers (3 elementos)
I2: =INDIRECT("Data!A2:A6")      → Array de nombres (5 elementos)
I3: =INDIRECT("Data!B1:B6")      → Array de edad (6 elementos)
I4: =INDIRECT(P3)                → Array desde celda (A1:C3)
```

### 2J. INDIRECT - Referencias de Columna/Fila Completa - 4 Cases
```
J1: =INDIRECT("Data!A:A")        → Columna completa A
J2: =INDIRECT("Data!B:B")        → Columna completa B
J3: =INDIRECT("Data!1:1")        → Fila completa 1
J4: =INDIRECT("Data!2:2")        → Fila completa 2
```

### 2K. INDIRECT - Casos de Error - 5 Cases
```
K1: =INDIRECT("InvalidSheet!A1") → #REF! (hoja inexistente)
K2: =INDIRECT("Data!Z99")        → 0 (celda vacía)
K3: =INDIRECT("")                → #REF! (referencia vacía)
K4: =INDIRECT("NotAReference")   → #REF! (texto inválido)
K5: =INDIRECT(P4)                → #REF! (hoja inválida desde celda)
```

## NIVEL 3: CASOS AVANZADOS (Combinaciones) - 10 Cases

### 3L. INDEX + INDIRECT - 4 Cases
```
L1: =INDEX(INDIRECT("Data!A1:E6"), 2, 2)    → 25
L2: =INDEX(INDIRECT("Data!A1:E6"), 0, 2)    → Array columna Age
L3: =INDEX(INDIRECT("Data!A2:C4"), 2, 3)    → "Chicago"
L4: =INDEX(INDIRECT("Data!A:A"), 3)         → "Bob"
```

### 3M. OFFSET + INDIRECT - 3 Cases
```
M1: =OFFSET(INDIRECT("Data!A1"), 1, 1)      → 25
M2: =OFFSET(INDIRECT("Data!B2"), 1, 1)      → "LA"
M3: =OFFSET(INDIRECT("Data!A1"), 1, 1, 2, 2) → Array 2x2
```

### 3N. Combinaciones Complejas - 4 Cases
```
N1: =INDEX(OFFSET(Data!A1, 0, 0, 3, 3), 2, 2)     → 25 (INDEX+OFFSET)
N2: =OFFSET(INDEX(Data!A1:E6, 2, 1), 1, 1)        → 30 (OFFSET+INDEX)
N3: =INDIRECT("Data!A" & 2)                       → "Alice" (Dynamic reference)
N4: =INDIRECT("Data!" & CHAR(66) & "2")           → 25 (CHAR-based reference)
```

## NIVEL 4: CASOS DE CONTEXTO (Uso con Otras Funciones) - 7 Cases

### 4O. Funciones con Agregación - 4 Cases
```
O1: =SUM(INDEX(Data!A1:E6, 0, 2))            → 140 (SUM+INDEX array - correct sum)
O2: =AVERAGE(OFFSET(Data!B1, 1, 0, 5, 1))    → 28 (promedio de edades)
O3: =COUNT(INDIRECT("Data!B:B"))              → 5 (contar números en col B)
O4: =MAX(INDEX(Data!A1:E6, 0, 4))            → 95 (máximo de scores)
```

### 4P. Manejo de Errores - 3 Cases
```
P1: =IFERROR(INDEX(Data!A1:E6, 10, 1), "Not Found") → "Not Found"
P2: =IF(ISERROR(OFFSET(Data!A1, -1, 0)), "Error", "OK") → "Error"
P3: =IFERROR(INDIRECT("InvalidSheet!A1"), "Sheet Error") → "Sheet Error"
```

## NIVEL 5: CASOS EDGE (Comportamientos Límite) - 7 Cases

### 5Q. Referencias Especiales - 3 Cases
```
Q1: =INDIRECT("Tests!O1")        → "Test Value" (referencia misma hoja)
Q2: =INDEX(Data!A:A, 2)          → "Alice" (INDEX con columna completa)
Q3: =OFFSET(Data!A:A, 1, 0, 3, 1) → Array (OFFSET con columna completa)
```

### 5R. Arrays Dinámicos (Excel 365) - 2 Cases
```
R1: =INDEX(Data!A1:E6, ROW(A1:A3), 1)       → Array ["Name", "Alice", "Bob"] (3 filas)
R2: =OFFSET(Data!A1, ROW(A1:A2)-1, 0)       → Array ["Name", "Alice"] (2 filas)
```

**Comportamiento Esperado:**
- ROW(A1:A3) debe retornar array [1, 2, 3]
- ROW(A1:A2) debe retornar array [1, 2]  
- ROW(A1:A2)-1 debe retornar array [0, 1]
- INDEX y OFFSET con arrays de entrada deben retornar arrays de salida
- Este es comportamiento válido de Excel 365 dynamic arrays

### 5S. Forma de Referencia vs Array - 2 Cases
```
S1: =INDEX((Data!A1:A5, Data!C1:C5), 2, 1, 1) → "Alice" (área 1)
S2: =INDEX((Data!A1:A5, Data!C1:C5), 2, 1, 2) → "NYC" (área 2)
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
O1: "Test Value"  → Para referencias circulares
```

## Resumen de Cobertura (75 Test Cases Total)

### Por Función:
- **INDEX**: 14 casos (fundamentales, arrays, errores)
- **OFFSET**: 16 casos (fundamentales, dimensiones, errores)
- **INDIRECT**: 21 casos (fundamentales, dinámico, arrays, errores)
- **Combinaciones**: 10 casos (INDEX+INDIRECT, OFFSET+INDIRECT, complejas)
- **Contexto**: 7 casos (agregación, manejo errores)
- **Edge Cases**: 7 casos (referencias especiales, arrays dinámicos)

### Por Comportamiento:
- **Valores simples**: 19 casos
- **Arrays/Rangos**: 28 casos
- **Errores**: 16 casos
- **Combinaciones**: 10 casos
- **Edge cases**: 7 casos

### Estrategia de Testing (Estructural → Incremental)
1. **INDEX valores individuales** (1A): Base fundamental
2. **INDEX arrays y errores** (1B, 1C): Comportamiento completo INDEX
3. **OFFSET valores y arrays** (2D, 2E): Referencias dinámicas básicas
4. **OFFSET errores** (2F): Manejo límites OFFSET
5. **INDIRECT todas las formas** (2G-2K): Conversión texto→referencia
6. **Combinaciones** (3L-3N): Interoperabilidad entre funciones
7. **Contexto** (4O-4P): Uso con otras funciones Excel
8. **Edge cases** (5Q-5S): Casos límite y compatibilidad moderna

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
- **Edge Cases**: Referencias especiales y arrays dinámicos

### Compatibilidad
- **Excel 365**: Dynamic array spilling y funciones modernas
- **Excel Legacy**: Array formulas con CSE
- **Todas las versiones**: Comportamiento básico consistente

Los 75 casos proporcionan cobertura exhaustiva y fiel del comportamiento de Excel para rangos dinámicos, asegurando implementación robusta y compatible.