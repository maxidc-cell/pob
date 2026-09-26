# -*- coding: utf-8 -*-
"""
Script de referencia: procesa los archivos POB*.xlsx de la carpeta `data/`
y genera `validacion/resultados_esperados.xlsx` con las mismas tablas que
va a producir el modelo de Power Query, más una hoja de Estadías (detalle
de auditoría) y el resumen de Control de Calidad.

Uso:
    python -m validacion.referencia [carpeta_data] [archivo_salida.xlsx]

Por defecto usa data/ como carpeta de entrada y
validacion/resultados_esperados.xlsx como salida.
"""
import sys
from pathlib import Path

import pandas as pd

if __package__ in (None, ""):
    # Permite ejecutar el archivo directamente (python validacion/referencia.py)
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from validacion import logica_pob as lp
else:
    from . import logica_pob as lp


def main(carpeta_data: str = "data", archivo_salida: str = "validacion/resultados_esperados.xlsx"):
    resultado = lp.procesar_todo(carpeta_data)

    with pd.ExcelWriter(archivo_salida, engine="openpyxl") as writer:
        resultado["POB_Diario"].to_excel(writer, sheet_name="POB_Diario", index=False)
        resultado["POB_Horario"].to_excel(writer, sheet_name="POB_Horario", index=False)
        resultado["POB_Franjas"].to_excel(writer, sheet_name="POB_Franjas", index=False)
        resultado["Comedor"].to_excel(writer, sheet_name="Comedor", index=False)
        resultado["Control_Calidad"].to_excel(writer, sheet_name="Control_Calidad", index=False)
        resultado["estadias"].to_excel(writer, sheet_name="Estadias", index=False)

    print(f"Archivos leídos: {resultado['n_archivos']}")
    print(f"Filas totales: {resultado['control_limpieza']['filas_leidas']}")
    print(f"Duplicados eliminados: {resultado['control_limpieza']['duplicados_eliminados']}")
    print(f"Personas distintas (IdPersona): {resultado['df_todo']['IdPersona'].nunique()}")
    print(f"Estadías construidas: {len(resultado['estadias'])}")
    print(f"Salidas sin ingreso previo: {resultado['control_estadias']['salidas_sin_ingreso_previo']}")
    print(f"Descartes de fin de mes: {resultado['control_estadias']['descartes_fin_de_mes']}")
    print(f"Lectores sin clasificar: {resultado['control_limpieza']['lectores_sin_clasificar']}")
    print(f"\nGenerado: {archivo_salida}")

    return resultado


if __name__ == "__main__":
    carpeta = sys.argv[1] if len(sys.argv) > 1 else "data"
    salida = sys.argv[2] if len(sys.argv) > 2 else "validacion/resultados_esperados.xlsx"
    main(carpeta, salida)
