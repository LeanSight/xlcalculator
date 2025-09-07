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

### A. INDEX - Casos Fundamentales
```
A1: =INDEX(Data.A1:E6, 2, 2)     → 25 (valor simple)
A2: =INDEX(Data.A1:E6, 3, 1)     → "Bob" (texto)
A3: =INDEX(Data.A1:E6, 4, 5)     → TRUE (boolean)
A4: =INDEX(Data.A1:E6, 6, 1)     → "Eve" (última fila)
A5: =INDEX(Data.A1:E6, 1, 5)     → TRUE (primera fila)
```

### B. INDEX - Casos de Error Estructurales
```
B1: =INDEX(Data.A1:E6, 7, 1)     → #REF! (fila fuera de rango)
B2: =INDEX(Data.A1:E6, 1, 7)     → #REF! (columna fuera de rango)
B3: =INDEX(Data.A1:E6, 0, 0)     → #VALUE! (ambos cero)
B4: =INDEX(Data.A1:E6, -1, 1)    → #VALUE! (fila negativa)
B5: =INDEX(Data.A1:E6, 1, -1)    → #VALUE! (columna negativa)
```

### C. INDEX - Casos de Fila/Columna Completa
```
C1: =INDEX(Data.A1:E6, 0, 2)     → Array completo columna Age
C2: =INDEX(Data.A1:E6, 2, 0)     → Array completo fila Alice
C3: =INDEX(Data.A1:E6, 0, 1)     → Array completo columna Name
```

## NIVEL 2: CASOS INTERMEDIOS (Funciones Individuales)

### D. OFFSET - Casos Fundamentales
```
D1: =OFFSET(Data.A1, 1, 1)       → 25 (B2)
D2: =OFFSET(Data.B2, 1, 1)       → "LA" (C3)
D3: =OFFSET(Data.A1, 0, 2)       → "City" (C1)
D4: =OFFSET(Data.A1, 5, 4)       → FALSE (E6)
```

### E. OFFSET - Casos con Dimensiones
```
E1: =OFFSET(Data.A1, 1, 1, 1, 1) → 25 (B2 como valor único)
E2: =OFFSET(Data.A1, 1, 1, 2, 2) → Array 2x2 desde B2
E3: =OFFSET(Data.A1, 0, 0, 3, 3) → Array 3x3 desde A1
E4: =OFFSET(Data.A1, 2, 1, 1, 3) → Array 1x3 desde B3
```

### F. OFFSET - Casos de Error
```
F1: =OFFSET(Data.A1, -1, 0)      → #VALUE! (fila negativa)
F2: =OFFSET(Data.A1, 0, -1)      → #VALUE! (columna negativa)
F3: =OFFSET(Data.A1, 10, 0)      → #REF! (fuera de hoja)
F4: =OFFSET(Data.A1, 0, 10)      → #REF! (fuera de hoja)
F5: =OFFSET(Data.A1, 1, 1, 0, 1) → #VALUE! (altura cero)
F6: =OFFSET(Data.A1, 1, 1, 1, 0) → #VALUE! (ancho cero)
```

### G. INDIRECT - Casos Fundamentales
```
G1: =INDIRECT("Data.B2")         → 25 (valor de celda)
G2: =INDIRECT("Data.C3")         → "LA" (texto)
G3: =INDIRECT("Data.E4")         → TRUE (boolean)
```

### H. INDIRECT - Referencias Dinámicas
```
H1: =INDIRECT("Data.A" & 2)      → "Alice" (referencia construida)
H2: =INDIRECT("Data." & CHAR(66) & "3") → 30 (columna B, fila 3)
H3: =INDIRECT("Data.A1:C1")      → Array de headers
H4: =INDIRECT("Data.A2:A6")      → Array de nombres
```

### I. INDIRECT - Casos de Error
```
I1: =INDIRECT("InvalidSheet.A1") → #REF! (hoja inexistente)
I2: =INDIRECT("Data.Z99")        → 0 o #REF! (celda vacía/inválida)
I3: =INDIRECT("")                → #REF! (referencia vacía)
I4: =INDIRECT("NotAReference")   → #REF! (texto inválido)
```

## NIVEL 3: CASOS AVANZADOS (Combinaciones)

### J. INDEX + INDIRECT
```
J1: =INDEX(INDIRECT("Data.A1:E6"), 2, 2)    → 25
J2: =INDEX(INDIRECT("Data.A1:E6"), 0, 2)    → Array columna Age
J3: =INDEX(INDIRECT("Data.A2:C4"), 2, 3)    → "Chicago"
```

### K. OFFSET + INDIRECT
```
K1: =OFFSET(INDIRECT("Data.A1"), 1, 1)      → 25
K2: =OFFSET(INDIRECT("Data.B2"), 1, 1)      → "LA"
K3: =INDIRECT(OFFSET("Data.A1", 1, 0))      → "Alice" (si OFFSET devuelve ref)
```

### L. Combinaciones Complejas
```
L1: =INDEX(OFFSET(Data.A1, 0, 0, 3, 3), 2, 2)     → 25
L2: =OFFSET(INDEX(Data.A1:E6, 0, 2), 1, 0)        → 25 (si INDEX devuelve ref)
L3: =INDIRECT("Data.A" & INDEX(Data.B1:B6, 2, 1)) → Referencia dinámica
```

## NIVEL 4: CASOS EDGE (Comportamientos Límite)

### M. Rangos Especiales
```
M1: =INDEX(Data.A:A, 2)          → "Alice" (columna completa)
M2: =INDEX(Data.1:1, 1, 2)       → "Age" (fila completa)
M3: =OFFSET(Data.A:A, 1, 0, 3, 1) → Array de 3 celdas
```

### N. Referencias Circulares y Complejas
```
N1: =INDIRECT("Tests.O1")        → Referencia a celda en misma hoja
N2: =INDEX(Data.A1:E6, ROW(), 1) → Referencia dinámica por fila actual
```

### O. Casos de Compatibilidad
```
O1: "Test Value"                 → Para N1
O2: =IFERROR(INDEX(Data.A1:E6, 10, 1), "Not Found") → Manejo de errores
O3: =IF(ISERROR(OFFSET(Data.A1, -1, 0)), "Error", "OK") → Detección errores
```

## Datos de Referencia para Validación

### Referencias Auxiliares (Columna P)
```
P1: "Data.B2"     → Para INDIRECT
P2: "Data.C3"     → Para INDIRECT  
P3: "Data.A1:C3"  → Para INDIRECT con rango
P4: "InvalidRef"  → Para error testing
P5: ""            → Para error testing
```

### Valores Esperados (Columna Q)
```
Q1: 25            → Valor esperado para A1
Q2: "Bob"         → Valor esperado para A2
Q3: TRUE          → Valor esperado para A3
Q4: "#REF!"       → Error esperado para B1
Q5: "#VALUE!"     → Error esperado para B3
```

## Estrategia de Testing

### Orden de Implementación (Estructural → Incremental)
1. **INDEX básico** (A1-A5): Valores simples
2. **INDEX errores** (B1-B5): Manejo de errores
3. **INDEX arrays** (C1-C3): Filas/columnas completas
4. **OFFSET básico** (D1-D4): Referencias simples
5. **OFFSET dimensiones** (E1-E4): Rangos con tamaño
6. **OFFSET errores** (F1-F6): Casos límite
7. **INDIRECT básico** (G1-G3): Referencias directas
8. **INDIRECT dinámico** (H1-H4): Referencias construidas
9. **INDIRECT errores** (I1-I4): Casos inválidos
10. **Combinaciones** (J1-L3): Funciones anidadas
11. **Edge cases** (M1-O3): Comportamientos límite

### Criterios de Éxito
- Cada celda debe devolver exactamente el mismo valor/error que Excel
- Los tipos de datos deben coincidir (Number, Text, Boolean, Array, Error)
- Los arrays deben tener las mismas dimensiones y valores
- Los errores deben ser del tipo correcto (#REF!, #VALUE!, #NAME!)