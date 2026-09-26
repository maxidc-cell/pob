# Cómo validar el Excel contra la referencia en Python

## 1. Generar los resultados de referencia

Con Python instalado (no hace falta en tu PC corporativa, esto se corre acá
o en cualquier máquina con Python):

```bash
pip install pandas openpyxl pytest
python3 validacion/referencia.py
```

Esto lee `data/POB.xlsx` (la muestra: 88.008 fichadas, del 1 al 25/9/2026) y
genera `validacion/resultados_esperados.xlsx` con las mismas tablas que va a
producir tu Excel: `POB_Diario`, `POB_Horario`, `POB_Franjas`, `Comedor`,
`Control_Calidad`, `Estadias`.

Para correr las pruebas unitarias de las reglas de negocio (fin de mes,
salida sin ingreso previo, movimiento entre segmentos, etc.):

```bash
python3 -m pytest validacion/test_logica_pob.py -v
```

Los 15 casos tienen que pasar en verde.

## 2. Cargar la misma muestra en tu Excel

Subí `data/POB.xlsx` a la carpeta de OneDrive que configuraste en
`tParametros` (con el nombre igual, empezando con el prefijo `POB`), refrescá
todas las consultas (Datos > Actualizar todo) y compará las cifras de abajo.

> Los números de esta sección salen de correr `referencia.py` el
> 2026-09-26. La regla de fin de mes se evalúa "en vivo" (contra el momento
> del cálculo, ver README): si la comparás muchos meses después de esa
> fecha, **Descartes de fin de mes** puede dar distinto de 0 aunque el resto
> coincida — no es un error, es exactamente lo esperado (ver `README.md`,
> punto 3 de "Supuestos y decisiones tomadas").

## 3. Cifras a comparar

### POB_Diario (filtrando Segmento = TOTAL, Empresa = TODAS)

| Fecha | Presentes_en_el_dia | Maximo_simultaneo | Hora_del_maximo |
|---|---|---|---|
| 2026-09-05 | 311 | 290 | 2026-09-05 13:34 |
| 2026-09-15 | 439 | 368 | 2026-09-15 13:23 |
| 2026-09-24 | 420 | 385 | 2026-09-24 13:12 |

### POB_Horario (Segmento = TOTAL, Empresa = TODAS)

| Fecha | Hora | POB |
|---|---|---|
| 2026-09-05 | 08:00 | 280 |
| 2026-09-05 | 14:00 | 287 |
| 2026-09-05 | 20:00 | 272 |
| 2026-09-15 | 08:00 | 325 |
| 2026-09-15 | 14:00 | 350 |
| 2026-09-15 | 20:00 | 303 |
| 2026-09-24 | 08:00 | 349 |
| 2026-09-24 | 14:00 | 376 |
| 2026-09-24 | 20:00 | 319 |

### POB_Franjas (Segmento = TOTAL, Empresa = TODAS)

| Fecha | Franja | POB |
|---|---|---|
| 2026-09-05 | 07-14 | 287 |
| 2026-09-05 | 19-24 | 286 |
| 2026-09-15 | 07-14 | 350 |
| 2026-09-15 | 19-24 | 314 |
| 2026-09-24 | 07-14 | 376 |
| 2026-09-24 | 19-24 | 327 |

### Comedor (Segmento = TOTAL, Empresa = TODAS)

| Fecha | Turno | Personas_unicas |
|---|---|---|
| 2026-09-05 | Desayuno | 79 |
| 2026-09-05 | Almuerzo | 161 |
| 2026-09-05 | Merienda | 25 |
| 2026-09-05 | Cena | 141 |
| 2026-09-15 | Desayuno | 90 |
| 2026-09-15 | Almuerzo | 182 |
| 2026-09-15 | Merienda | 10 |
| 2026-09-15 | Cena | 134 |
| 2026-09-24 | Desayuno | 107 |
| 2026-09-24 | Almuerzo | 190 |
| 2026-09-24 | Merienda | 16 |
| 2026-09-24 | Cena | 125 |

### Control_Calidad (con sólo `data/POB.xlsx` cargado)

| Indicador | Valor |
|---|---|
| Archivos leídos | 1 |
| Filas totales (antes de deduplicar) | 88008 |
| Duplicados eliminados | 1324 |
| Lectores sin clasificar | (ninguno) |
| Marcas sin IGG identificadas por credencial | 836 |
| Personas con más de una credencial | 39 |
| Salidas sin ingreso previo | 118 |
| Descartes de fin de mes | 0 (ver nota arriba si validás después de octubre) |
| Personas SIN EMPRESA | 307 |
| Personas --- A DEFINIR --- | 21 |
| Marcas descartadas por evento: Invalid Badge | 691 |
| Marcas descartadas por evento: Invalid Access Level | 454 |
| Marcas descartadas por evento: Invalid Card Format | 225 |
| Marcas descartadas por evento: Inactive Badge | 101 |
| Marcas descartadas por evento: Access Granted No Entry Made | 3 |

## 4. Si algo no coincide

1. Primero revisá `Control_Calidad`: si "Filas totales" o "Duplicados
   eliminados" ya difieren, el problema está en la carga del archivo (ruta
   de OneDrive, nombre del archivo, o el archivo se cargó dos veces).
2. Si Control de Calidad coincide pero algún reporte no, mirá la hoja oculta
   `Estadias` de tu Excel y compará el detalle de una persona puntual contra
   la hoja `Estadias` de `resultados_esperados.xlsx` (filtrando por
   `IdPersona`).
3. Repasá `README.md` → "Supuestos y decisiones tomadas": varias diferencias
   posibles ya están documentadas ahí.
