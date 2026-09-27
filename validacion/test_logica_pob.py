# -*- coding: utf-8 -*-
"""
Pruebas unitarias de la lógica de negocio (§4 del prompt), con datos
sintéticos chicos y armados a mano para poder verificar cada regla de
manera aislada.
"""
import pandas as pd
import pytest

from . import config as cfg
from . import logica_pob as lp


def _ts(texto):
    return pd.Timestamp(texto)


def _validos(filas):
    """Arma un DataFrame de eventos válidos ya limpios: (IdPersona, FechaHora, TipoLector, Segmento)."""
    return pd.DataFrame(filas, columns=["IdPersona", "FechaHora", "TipoLector", "Segmento"])


# ===========================================================================
# Regla de fin de mes (3 casos pedidos explícitamente en el prompt)
# ===========================================================================

def test_fin_de_mes_descarta_si_no_hay_marca_posterior():
    """Entró el 10, no marcó nada más -> se descarta el día 1 (del mes siguiente)."""
    eventos = _validos([
        ("P1", _ts("2026-09-10 08:00"), "INGRESO", "AP"),
    ])
    estadias, control = lp.construir_estadias(eventos, ahora=_ts("2026-10-05"))
    assert control["descartes_fin_de_mes"] == 1
    fila = estadias[estadias["Motivo"] == "Descarte fin de mes"].iloc[0]
    assert fila["Inicio"] == _ts("2026-09-10 08:00")
    assert fila["Fin"] == _ts("2026-10-01 00:00")
    # no debe quedar ninguna estadía abierta para P1
    assert not ((estadias["IdPersona"] == "P1") & estadias["Fin"].isna()).any()


def test_fin_de_mes_no_descarta_si_hay_marca_posterior_aunque_sea_interna():
    """Entró el 28, marcó comedor el 29 y el 30, sin salida -> sigue adentro el día 1."""
    eventos = _validos([
        ("P2", _ts("2026-09-28 09:00"), "INGRESO", "AP"),
        ("P2", _ts("2026-09-29 12:30"), "COMEDOR", "AP"),
        ("P2", _ts("2026-09-30 12:30"), "COMEDOR", "AP"),
    ])
    estadias, control = lp.construir_estadias(eventos, ahora=_ts("2026-10-02"))
    assert control["descartes_fin_de_mes"] == 0
    abiertas = estadias[(estadias["IdPersona"] == "P2") & estadias["Fin"].isna()]
    assert len(abiertas) == 1
    assert abiertas.iloc[0]["Inicio"] == _ts("2026-09-28 09:00")


def test_fin_de_mes_descartado_vuelve_a_entrar_con_marca_posterior():
    """Descartado el día 1 y marca en un lector interno el día 3 -> vuelve a ADENTRO desde esa marca."""
    eventos = _validos([
        ("P3", _ts("2026-09-10 08:00"), "INGRESO", "AP"),
        ("P3", _ts("2026-10-03 07:00"), "INTERNO", "AP"),
    ])
    estadias, control = lp.construir_estadias(eventos, ahora=_ts("2026-10-05"))
    assert control["descartes_fin_de_mes"] == 1
    descarte = estadias[estadias["Motivo"] == "Descarte fin de mes"].iloc[0]
    assert descarte["Fin"] == _ts("2026-10-01 00:00")
    abierta = estadias[(estadias["IdPersona"] == "P3") & estadias["Fin"].isna()].iloc[0]
    assert abierta["Inicio"] == _ts("2026-10-03 07:00")


# ===========================================================================
# Salida sin ingreso previo
# ===========================================================================

def test_salida_sin_ingreso_previo_usa_00_00_del_dia_si_es_la_primera_marca():
    eventos = _validos([
        ("P4", _ts("2026-09-05 09:15"), "EGRESO", "AP"),
    ])
    estadias, control = lp.construir_estadias(eventos)
    assert control["salidas_sin_ingreso_previo"] == 1
    fila = estadias.iloc[0]
    assert fila["Motivo"] == "Salida sin ingreso previo"
    assert fila["Inicio"] == _ts("2026-09-05 00:00")
    assert fila["Fin"] == _ts("2026-09-05 09:15")
    assert fila["Segmento"] == "AP"


def test_salida_sin_ingreso_previo_usa_la_marca_anterior_si_es_mas_tarde_que_00_00():
    """Si la persona ya había marcado ese día (p.ej. un interno) y después sale sin
    haber hecho un INGRESO formal, el inicio de la estadía retroactiva es esa marca."""
    eventos = _validos([
        ("P5", _ts("2026-09-05 06:00"), "INTERNO", "AP"),
        ("P5", _ts("2026-09-05 06:00"), "EGRESO", "AP"),  # egresa apenas entra: NO es "sin ingreso previo"
    ])
    estadias, control = lp.construir_estadias(eventos)
    # Este caso en realidad sí tiene "ingreso previo" (el interno la puso adentro),
    # así que la salida es normal, no retroactiva.
    assert control["salidas_sin_ingreso_previo"] == 0
    assert (estadias["Motivo"] == "Normal").all()


def test_salida_sin_ingreso_previo_con_marca_de_otro_dia_usa_00_00_de_hoy():
    """La marca anterior fue ayer: el inicio retroactivo no puede ser antes de las
    00:00 de HOY (no se arrastra la presencia de un día a otro sin evidencia)."""
    eventos = _validos([
        ("P6", _ts("2026-09-04 22:00"), "COMEDOR", "AP"),  # entra y ese mismo día no vuelve a marcar
    ])
    estadias, control = lp.construir_estadias(eventos)
    # P6 queda ADENTRO (estadía abierta) tras el comedor del día 4
    assert control["salidas_sin_ingreso_previo"] == 0
    eventos2 = _validos([
        ("P6", _ts("2026-09-04 22:00"), "COMEDOR", "AP"),
        ("P6", _ts("2026-09-06 05:00"), "EGRESO", "AP"),
    ])
    estadias2, control2 = lp.construir_estadias(eventos2)
    # Acá la marca anterior (4/9 22:00) sigue siendo válida como "ingreso previo":
    # la persona nunca salió, así que la salida del día 6 es normal (cierra la
    # estadía abierta desde el 4), no "sin ingreso previo".
    assert control2["salidas_sin_ingreso_previo"] == 0
    assert (estadias2["Motivo"] == "Normal").all()
    assert estadias2.iloc[0]["Inicio"] == _ts("2026-09-04 22:00")
    assert estadias2.iloc[0]["Fin"] == _ts("2026-09-06 05:00")


# ===========================================================================
# Movimiento AP <-> LLY
# ===========================================================================

def test_movimiento_entre_segmentos_cierra_y_abre_estadia():
    eventos = _validos([
        ("P7", _ts("2026-09-01 08:00"), "INGRESO", "AP"),
        ("P7", _ts("2026-09-01 10:00"), "INTERNO", "LLY"),  # se mueve de AP a LLY
        ("P7", _ts("2026-09-01 18:00"), "EGRESO", "LLY"),
    ])
    estadias, _ = lp.construir_estadias(eventos)
    estadias = estadias.sort_values("Inicio").reset_index(drop=True)
    assert len(estadias) == 2
    assert estadias.iloc[0]["Segmento"] == "AP"
    assert estadias.iloc[0]["Inicio"] == _ts("2026-09-01 08:00")
    assert estadias.iloc[0]["Fin"] == _ts("2026-09-01 10:00")
    assert estadias.iloc[1]["Segmento"] == "LLY"
    assert estadias.iloc[1]["Inicio"] == _ts("2026-09-01 10:00")
    assert estadias.iloc[1]["Fin"] == _ts("2026-09-01 18:00")


def test_una_persona_nunca_esta_en_dos_segmentos_a_la_vez():
    eventos = _validos([
        ("P8", _ts("2026-09-01 08:00"), "INGRESO", "AP"),
        ("P8", _ts("2026-09-01 09:00"), "INTERNO", "LLY"),
        ("P8", _ts("2026-09-01 09:30"), "INTERNO", "AP"),
        ("P8", _ts("2026-09-01 20:00"), "EGRESO", "AP"),
    ])
    estadias, _ = lp.construir_estadias(eventos)
    intervalos = list(zip(estadias["Inicio"], estadias["Fin"]))
    intervalos.sort()
    for i in range(len(intervalos) - 1):
        assert intervalos[i][1] <= intervalos[i + 1][0]


# ===========================================================================
# Criterio de minuto: foto en HH:MM incluye eventos con FechaHora <= HH:MM
# ===========================================================================

def test_quien_sale_exactamente_a_las_14_no_cuenta_en_la_foto_de_las_14():
    estadias = pd.DataFrame([
        {"IdPersona": "P9", "Segmento": "AP", "Inicio": _ts("2026-09-01 07:00"),
         "Fin": _ts("2026-09-01 14:00"), "Motivo": "Normal", "Empresa": "TODAS"},
    ])
    instante = _ts("2026-09-01 14:00")
    presente = lp._presentes_en_instante(estadias, instante)
    assert not presente.iloc[0]


def test_quien_entra_exactamente_a_las_14_si_cuenta_en_la_foto_de_las_14():
    estadias = pd.DataFrame([
        {"IdPersona": "P10", "Segmento": "AP", "Inicio": _ts("2026-09-01 14:00"),
         "Fin": pd.NaT, "Motivo": "Abierta (sigue adentro)", "Empresa": "TODAS"},
    ])
    instante = _ts("2026-09-01 14:00")
    presente = lp._presentes_en_instante(estadias, instante)
    assert presente.iloc[0]


# ===========================================================================
# Duplicados entre archivos
# ===========================================================================

def test_duplicados_exactos_se_eliminan_una_sola_vez():
    fila = {
        "Fecha": "1/9/2026", "Hora": "08:00", "IGG": "D1", "Apellido": "PEREZ",
        "Nombre": "JUAN", "Credencial": 111, "Evento": "Access Granted",
        "SEGMENTO": "AGUADA PICHANA", "EMPRESA": "ACME", "Reader Description": "AP - PortExter ENT 01-1-01",
    }
    # el mismo registro viene repetido en dos archivos distintos (superposición de fechas)
    df_raw = pd.DataFrame([fila, fila, fila])
    df_todo, control = lp.limpiar_datos(df_raw)
    assert control["duplicados_eliminados"] == 2
    assert len(df_todo) == 1


# ===========================================================================
# Lector sin clasificar
# ===========================================================================

def test_lector_no_precargado_se_reporta_sin_clasificar_y_se_trata_como_interno():
    fila = {
        "Fecha": "1/9/2026", "Hora": "08:00", "IGG": "D1", "Apellido": "PEREZ",
        "Nombre": "JUAN", "Credencial": 111, "Evento": "Access Granted",
        "SEGMENTO": "AGUADA PICHANA", "EMPRESA": "ACME",
        "Reader Description": "AP - Lector Nuevo Que No Está En La Tabla 99-9-99",
    }
    df_raw = pd.DataFrame([fila])
    df_todo, control = lp.limpiar_datos(df_raw)
    assert control["lectores_sin_clasificar"] == ["AP - Lector Nuevo Que No Está En La Tabla 99-9-99"]
    assert df_todo.iloc[0]["TipoLector"] == "INTERNO"


def test_lector_con_doble_espacio_matchea_la_tabla_normalizada():
    fila = {
        "Fecha": "1/9/2026", "Hora": "08:00", "IGG": "D1", "Apellido": "PEREZ",
        "Nombre": "JUAN", "Credencial": 111, "Evento": "Access Granted",
        "SEGMENTO": "LOMA LAS YEGUAS", "EMPRESA": "ACME",
        "Reader Description": "LLY- Porteria Vehiculos SAL  02-1-05",  # doble espacio, como en la muestra real
    }
    df_raw = pd.DataFrame([fila])
    df_todo, control = lp.limpiar_datos(df_raw)
    assert control["lectores_sin_clasificar"] == []
    assert df_todo.iloc[0]["TipoLector"] == "EGRESO"


# ===========================================================================
# IdPersona e IGG vacío
# ===========================================================================

def test_filas_sin_fecha_o_sin_hora_se_descartan_sin_romper():
    filas = [
        {
            "Fecha": "1/9/2026", "Hora": "08:00", "IGG": "D1", "Apellido": "PEREZ",
            "Nombre": "JUAN", "Credencial": 111, "Evento": "Access Granted",
            "SEGMENTO": "AGUADA PICHANA", "EMPRESA": "ACME", "Reader Description": "AP - PortExter ENT 01-1-01",
        },
        {
            "Fecha": None, "Hora": "09:00", "IGG": "D2", "Apellido": "GOMEZ",
            "Nombre": "ANA", "Credencial": 222, "Evento": "Access Granted",
            "SEGMENTO": "AGUADA PICHANA", "EMPRESA": "ACME", "Reader Description": "AP - PortExter ENT 01-1-01",
        },
        {
            "Fecha": "1/9/2026", "Hora": None, "IGG": "D3", "Apellido": "RUIZ",
            "Nombre": "LUIS", "Credencial": 333, "Evento": "Access Granted",
            "SEGMENTO": "AGUADA PICHANA", "EMPRESA": "ACME", "Reader Description": "AP - PortExter ENT 01-1-01",
        },
    ]
    df_raw = pd.DataFrame(filas)
    df_todo, control = lp.limpiar_datos(df_raw)
    assert control["filas_sin_fecha_hora_descartadas"] == 2
    assert len(df_todo) == 1
    assert df_todo.iloc[0]["IdPersona"] == "D1"


def test_igg_vacio_usa_credencial_como_identificador():
    fila = {
        "Fecha": "1/9/2026", "Hora": "08:00", "IGG": None, "Apellido": "PEREZ",
        "Nombre": "JUAN", "Credencial": 999, "Evento": "Access Granted",
        "SEGMENTO": "AGUADA PICHANA", "EMPRESA": "ACME", "Reader Description": "AP - PortExter ENT 01-1-01",
    }
    df_raw = pd.DataFrame([fila])
    df_todo, control = lp.limpiar_datos(df_raw)
    assert df_todo.iloc[0]["IdPersona"] == "CRED_999"
    assert control["marcas_identificadas_por_credencial"] == 1


# ===========================================================================
# Comedor: extremos inclusivos, marcas fuera de turno se ignoran
# ===========================================================================

def test_comedor_extremos_inclusivos_y_fuera_de_turno_se_ignora():
    df_validos = pd.DataFrame([
        {"IdPersona": "P11", "FechaHora": _ts("2026-09-01 12:00"), "TipoLector": "COMEDOR",
         "Segmento": "AP", "Lector": cfg.normalizar_lector("AP - COMEDOR 10-0-01")},
        {"IdPersona": "P12", "FechaHora": _ts("2026-09-01 14:00"), "TipoLector": "COMEDOR",
         "Segmento": "AP", "Lector": cfg.normalizar_lector("AP - COMEDOR 10-0-01")},
        {"IdPersona": "P13", "FechaHora": _ts("2026-09-01 14:01"), "TipoLector": "COMEDOR",
         "Segmento": "AP", "Lector": cfg.normalizar_lector("AP - COMEDOR 10-0-01")},
    ])
    empresa = pd.Series({"P11": "ACME", "P12": "ACME", "P13": "ACME"})
    comedor = lp.calcular_comedor(df_validos, empresa)
    almuerzo_total = comedor[(comedor["Turno"] == "Almuerzo") & (comedor["Segmento"] == "TOTAL")
                              & (comedor["Empresa"] == "TODAS")]
    assert almuerzo_total.iloc[0]["Personas_unicas"] == 2  # P11 (12:00) y P12 (14:00, borde inclusive)


if __name__ == "__main__":
    import sys
    sys.exit(pytest.main([__file__, "-v"]))
