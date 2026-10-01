# -*- coding: utf-8 -*-
"""
Lógica de negocio para los reportes de POB (personal a bordo) y Comedor,
a partir de fichadas RFID. Es la implementación de referencia en Python
(pandas) de las reglas descriptas en el prompt del proyecto; las consultas
Power Query (carpeta pq/) tienen que replicar exactamente este
comportamiento.

Todas las funciones son puras (reciben DataFrames, devuelven DataFrames)
para poder testearlas con pytest usando datos sintéticos chicos.
"""
from __future__ import annotations

import glob
import os
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Optional

import numpy as np
import pandas as pd

from . import config as cfg

COLUMNAS_ORIGINALES = [
    "Fecha", "Hora", "IGG", "Apellido", "Nombre", "Credencial",
    "Evento", "SEGMENTO", "EMPRESA", "Reader Description",
]


# ===========================================================================
# 1) Carga de archivos
# ===========================================================================

def listar_archivos_origen(carpeta: str, prefijo: str = cfg.PREFIJO_ARCHIVO) -> list[str]:
    """Devuelve, ordenados, los .xlsx de la carpeta cuyo nombre empieza con el prefijo."""
    patron = os.path.join(carpeta, f"{prefijo}*.xlsx")
    return sorted(glob.glob(patron))


def cargar_archivos(carpeta: str, prefijo: str = cfg.PREFIJO_ARCHIVO,
                     hoja: str = cfg.HOJA_ORIGEN) -> pd.DataFrame:
    """Lee y concatena todos los archivos de origen. Agrega columna ArchivoOrigen."""
    archivos = listar_archivos_origen(carpeta, prefijo)
    if not archivos:
        raise FileNotFoundError(
            f"No se encontraron archivos '{prefijo}*.xlsx' en '{carpeta}'."
        )
    partes = []
    for ruta in archivos:
        df = pd.read_excel(ruta, sheet_name=hoja, dtype={"IGG": str, "Credencial": str})
        df["ArchivoOrigen"] = os.path.basename(ruta)
        partes.append(df)
    return pd.concat(partes, ignore_index=True)


# ===========================================================================
# 2) Limpieza y normalización
# ===========================================================================

def _parsear_fecha(serie: pd.Series) -> pd.Series:
    """Fecha viene como texto d/m/yyyy (cultura es-AR). Robusto a que ya sea datetime."""
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie.dt.normalize()
    return pd.to_datetime(serie, format="%d/%m/%Y", dayfirst=True, errors="coerce")


def _parsear_hora(serie: pd.Series) -> pd.Series:
    """Hora viene como texto HH:mm. Devuelve un timedelta desde las 00:00.

    Una celda vacía (None/NaN) devuelve NaT (no 00:00): así la fila queda
    detectable como incompleta más adelante, en vez de fecharse en
    silencio a medianoche.
    """
    if pd.api.types.is_datetime64_any_dtype(serie):
        return pd.to_timedelta(serie.dt.strftime("%H:%M:%S"))
    if hasattr(serie, "dtype") and str(serie.dtype).startswith("timedelta"):
        return serie
    vacio = serie.isna()
    texto = serie.astype(str).str.strip()
    partes = texto.str.split(":", expand=True)
    horas = pd.to_numeric(partes[0], errors="coerce").fillna(0).astype(int)
    minutos = pd.to_numeric(partes[1], errors="coerce").fillna(0).astype(int) if partes.shape[1] > 1 else 0
    resultado = pd.to_timedelta(horas, unit="h") + pd.to_timedelta(minutos, unit="m")
    resultado[vacio] = pd.NaT
    return resultado


def limpiar_datos(df_raw: pd.DataFrame):
    """Limpia, normaliza y clasifica las fichadas crudas.

    Devuelve (df_todo, control) donde:
    - df_todo: todas las filas después de sacar duplicados, con las columnas
      derivadas (FechaHora, IdPersona, Lector, TipoLector, Segmento, EventoValido).
    - control: dict con contadores para Control de Calidad.
    """
    control: dict = {}

    df = df_raw.copy()
    control["filas_leidas"] = len(df)

    # --- Duplicados: se consideran todas las columnas originales ---
    antes = len(df)
    df = df.drop_duplicates(subset=COLUMNAS_ORIGINALES, keep="first").reset_index(drop=True)
    control["duplicados_eliminados"] = antes - len(df)

    # --- Fecha/Hora -> FechaHora (resolución de minuto) ---
    fecha = _parsear_fecha(df["Fecha"])
    hora = _parsear_hora(df["Hora"])
    df["FechaHora"] = fecha + hora
    # Filas sin Fecha u Hora (blancos en un archivo armado a mano, por
    # ejemplo) no se pueden fichar: se descartan, igual que en pq/03_Limpieza.pq.
    control["filas_sin_fecha_hora_descartadas"] = int(df["FechaHora"].isna().sum())
    df = df.loc[df["FechaHora"].notna()].reset_index(drop=True)

    # Filas con un año disparatado (typo de fecha, o un serial de Excel mal
    # interpretado) también se descartan: además de ser un dato basura, si
    # llegaran a construir_estadias() podrían disparar miles de checkpoints
    # de fin de mes entre esa fecha rota y hoy (ver misma protección en
    # pq/03_Limpieza.pq y el límite de MaximoControlesFinDeMes en
    # pq/06_Estadias.pq / _primeros_dias_de_mes más abajo).
    anio_actual = pd.Timestamp.now().year
    fuera_de_rango = ~df["FechaHora"].dt.year.between(2015, anio_actual + 1)
    control["filas_fecha_fuera_de_rango_descartadas"] = int(fuera_de_rango.sum())
    df = df.loc[~fuera_de_rango].reset_index(drop=True)

    # --- Normalización de lector ---
    df["Lector"] = df["Reader Description"].map(cfg.normalizar_lector)

    # --- Clasificación de lector ---
    df["TipoLector"] = df["Lector"].map(cfg.LECTORES).fillna(cfg.TIPO_POR_DEFECTO)
    lectores_sin_clasificar = sorted(
        set(df.loc[~df["Lector"].isin(cfg.LECTORES), "Lector"].unique())
    )
    control["lectores_sin_clasificar"] = lectores_sin_clasificar

    # --- Segmento corto ---
    df["Segmento"] = df["SEGMENTO"].map(cfg.SEGMENTOS)
    segmentos_desconocidos = sorted(set(df.loc[df["Segmento"].isna(), "SEGMENTO"].dropna().unique()))
    control["segmentos_desconocidos"] = segmentos_desconocidos

    # --- IdPersona: IGG sin espacios, o CRED_<credencial> si no hay IGG ---
    igg_limpio = df["IGG"].astype(str).str.strip()
    igg_vacio = df["IGG"].isna() | (igg_limpio == "") | (igg_limpio.str.lower() == "nan")
    df["IdPersona"] = np.where(igg_vacio, "CRED_" + df["Credencial"].astype(str).str.strip(), igg_limpio)
    control["marcas_identificadas_por_credencial"] = int(igg_vacio.sum())

    # --- Evento válido ---
    df["EventoValido"] = df["Evento"].isin(cfg.EVENTOS_VALIDOS)
    control["marcas_por_evento"] = df["Evento"].value_counts().to_dict()
    control["marcas_descartadas_por_evento"] = (
        df.loc[~df["EventoValido"], "Evento"].value_counts().to_dict()
    )

    # --- Personas con más de una credencial (mismo IdPersona, distintas credenciales) ---
    con_igg = df.loc[~igg_vacio]
    personas_multi_credencial = (
        con_igg.groupby("IdPersona")["Credencial"].nunique()
    )
    control["personas_con_mas_de_una_credencial"] = int((personas_multi_credencial > 1).sum())

    return df.sort_values(["FechaHora"], kind="stable").reset_index(drop=True), control


def calcular_datos_persona(df_todo: pd.DataFrame, empresa_por_persona: pd.Series) -> pd.DataFrame:
    """Último Apellido/Nombre informados por persona (por FechaHora), más su
    Empresa (de calcular_empresa_por_persona). Devuelve un DataFrame con
    columnas IdPersona, Apellido, Nombre, Empresa — pensado para cruzarse
    (merge on="IdPersona") contra tablas de detalle como Foto_13_22 o
    POB_Periodos.
    """
    df = df_todo.sort_values("FechaHora")
    datos = df.groupby("IdPersona")[["Apellido", "Nombre"]].last().reset_index()
    datos["Empresa"] = datos["IdPersona"].map(empresa_por_persona).fillna("SIN EMPRESA")
    return datos


def calcular_empresa_por_persona(df_todo: pd.DataFrame) -> pd.Series:
    """Última empresa informada por persona en el período (por FechaHora).

    Se usa el historial completo de fichadas (incluidas las inválidas), porque
    el dato de empresa viaja con la credencial más allá de que el acceso haya
    sido otorgado o no. Si la persona nunca tiene empresa, se usa 'SIN EMPRESA'.
    '--- A DEFINIR ---' se respeta tal cual viene.
    """
    df = df_todo.sort_values("FechaHora")
    ultima = df.groupby("IdPersona")["EMPRESA"].last()
    ultima = ultima.fillna("SIN EMPRESA")
    ultima = ultima.replace("", "SIN EMPRESA")
    return ultima


# ===========================================================================
# 3) Máquina de estados: fichadas -> estadías
# ===========================================================================

@dataclass
class _Estado:
    adentro: bool = False
    segmento: Optional[str] = None
    apertura: Optional[pd.Timestamp] = None
    ultima_marca: Optional[pd.Timestamp] = None


def _primeros_dias_de_mes(fecha_min: pd.Timestamp, ahora: pd.Timestamp) -> list[pd.Timestamp]:
    """Boundaries de fin de mes (00:00 del día 1) estrictamente dentro de
    [fecha_min, ahora]. La regla de fin de mes es "en vivo": se dispara cada
    vez que el reloj cruza un 1° de mes, exista o no ya el archivo del mes
    siguiente cargado. Por eso el límite superior es `ahora` (el momento en
    que se corre el cálculo), no el último dato cargado."""
    boundaries = []
    cursor = pd.Timestamp(year=fecha_min.year, month=fecha_min.month, day=1)
    if cursor <= fecha_min:
        # avanzar al primer día 1 posterior a fecha_min
        if cursor.month == 12:
            cursor = pd.Timestamp(cursor.year + 1, 1, 1)
        else:
            cursor = pd.Timestamp(cursor.year, cursor.month + 1, 1)
    MAXIMO_CONTROLES_FIN_DE_MES = 240  # 20 años; ver misma protección en pq/06_Estadias.pq
    while cursor <= ahora and len(boundaries) < MAXIMO_CONTROLES_FIN_DE_MES:
        boundaries.append(cursor)
        if cursor.month == 12:
            cursor = pd.Timestamp(cursor.year + 1, 1, 1)
        else:
            cursor = pd.Timestamp(cursor.year, cursor.month + 1, 1)
    return boundaries


def construir_estadias(df_validos: pd.DataFrame, ahora: Optional[pd.Timestamp] = None) -> tuple[pd.DataFrame, dict]:
    """Aplica la máquina de estados descripta en el prompt (§4) a las marcas
    válidas de cada persona, en orden cronológico.

    df_validos debe tener: IdPersona, FechaHora, TipoLector, Segmento (todas
    ya limpias). `ahora` es el instante de referencia contra el que se evalúa
    la regla de fin de mes (por defecto, el momento real en que se corre el
    cálculo — igual que haría la consulta M al refrescarse en Excel). Se
    puede fijar explícitamente para pruebas reproducibles.

    Devuelve (estadias, control) donde estadias tiene columnas:
    IdPersona, Segmento, Inicio, Fin (NaT si sigue adentro), Motivo.
    """
    if df_validos.empty:
        return (
            pd.DataFrame(columns=["IdPersona", "Segmento", "Inicio", "Fin", "Motivo"]),
            {"salidas_sin_ingreso_previo": 0, "descartes_fin_de_mes": 0},
        )

    if ahora is None:
        ahora = pd.Timestamp.now()

    fecha_min = df_validos["FechaHora"].min()
    boundaries = _primeros_dias_de_mes(fecha_min, ahora)

    estadias = []
    control = {"salidas_sin_ingreso_previo": 0, "descartes_fin_de_mes": 0}

    df = df_validos.sort_values(["IdPersona", "FechaHora"], kind="stable")

    for id_persona, grupo in df.groupby("IdPersona", sort=False):
        estado = _Estado()
        marca_previa: Optional[pd.Timestamp] = None  # última marca real (de cualquier tipo) antes de la actual
        eventos = list(grupo[["FechaHora", "TipoLector", "Segmento"]].itertuples(index=False, name=None))

        # Mezclar los boundaries de fin de mes que caen dentro del rango de
        # este período con los eventos reales, ordenados por tiempo. Los
        # boundaries son "eventos virtuales" (tipo especial CHECKPOINT).
        linea_tiempo = [(t, tipo, seg, False) for (t, tipo, seg) in eventos]
        linea_tiempo += [(b, "CHECKPOINT", None, True) for b in boundaries]
        # Si un checkpoint (00:00 del día 1) empata en el tiempo con una marca real,
        # el checkpoint se evalúa primero (la marca ya pertenece al mes nuevo).
        linea_tiempo.sort(key=lambda x: (x[0], 0 if x[3] else 1))

        for tiempo, tipo, segmento, es_checkpoint in linea_tiempo:
            if es_checkpoint:
                if estado.adentro and estado.ultima_marca == estado.apertura:
                    estadias.append({
                        "IdPersona": id_persona, "Segmento": estado.segmento,
                        "Inicio": estado.apertura, "Fin": tiempo,
                        "Motivo": "Descarte fin de mes",
                    })
                    control["descartes_fin_de_mes"] += 1
                    estado = _Estado()
                continue

            if tipo == "INGRESO":
                if not estado.adentro:
                    estado.adentro = True
                    estado.segmento = segmento
                    estado.apertura = tiempo
                    estado.ultima_marca = tiempo
                else:
                    estado.ultima_marca = tiempo
                    if segmento != estado.segmento:
                        estadias.append({
                            "IdPersona": id_persona, "Segmento": estado.segmento,
                            "Inicio": estado.apertura, "Fin": tiempo, "Motivo": "Normal",
                        })
                        estado.segmento = segmento
                        estado.apertura = tiempo
                    # si es el mismo segmento, no pasa nada más que refrescar ultima_marca

            elif tipo in ("INTERNO", "COMEDOR"):
                if not estado.adentro:
                    estado.adentro = True
                    estado.segmento = segmento
                    estado.apertura = tiempo
                    estado.ultima_marca = tiempo
                else:
                    estado.ultima_marca = tiempo
                    if segmento != estado.segmento:
                        estadias.append({
                            "IdPersona": id_persona, "Segmento": estado.segmento,
                            "Inicio": estado.apertura, "Fin": tiempo, "Motivo": "Normal",
                        })
                        estado.segmento = segmento
                        estado.apertura = tiempo
                    # mismo segmento estando adentro: sólo refresca ultima_marca

            elif tipo == "EGRESO":
                if estado.adentro:
                    estadias.append({
                        "IdPersona": id_persona, "Segmento": estado.segmento,
                        "Inicio": estado.apertura, "Fin": tiempo, "Motivo": "Normal",
                    })
                    estado = _Estado()
                else:
                    inicio_dia = tiempo.normalize()
                    inicio = max(inicio_dia, marca_previa) if marca_previa is not None else inicio_dia
                    estadias.append({
                        "IdPersona": id_persona, "Segmento": segmento,
                        "Inicio": inicio, "Fin": tiempo, "Motivo": "Salida sin ingreso previo",
                    })
                    control["salidas_sin_ingreso_previo"] += 1
                    estado = _Estado()

            marca_previa = tiempo

        # fin del período: si sigue adentro, queda una estadía abierta (Fin=NaT)
        if estado.adentro:
            estadias.append({
                "IdPersona": id_persona, "Segmento": estado.segmento,
                "Inicio": estado.apertura, "Fin": pd.NaT, "Motivo": "Abierta (sigue adentro)",
            })

    out = pd.DataFrame(estadias, columns=["IdPersona", "Segmento", "Inicio", "Fin", "Motivo"])
    return out, control


# ===========================================================================
# 4) Reportes de salida
# ===========================================================================

def _rango_dias(estadias: pd.DataFrame) -> pd.DatetimeIndex:
    inicio = estadias["Inicio"].min().normalize()
    fin_validos = estadias["Fin"].dropna()
    ultimo = fin_validos.max() if not fin_validos.empty else estadias["Inicio"].max()
    fin = max(estadias["Inicio"].max(), ultimo).normalize()
    return pd.date_range(inicio, fin, freq="D")


def _presentes_en_instante(estadias: pd.DataFrame, instante: pd.Timestamp) -> pd.Series:
    """Máscara booleana: estadías que cubren el instante (Inicio <= instante < Fin)."""
    cubre_inicio = estadias["Inicio"] <= instante
    cubre_fin = estadias["Fin"].isna() | (estadias["Fin"] > instante)
    return cubre_inicio & cubre_fin


def _maximo_simultaneo(estadias: pd.DataFrame, day_start: pd.Timestamp, day_end: pd.Timestamp):
    """Pico de personas ADENTRO simultáneamente durante [day_start, day_end).

    Devuelve (maximo, hora_del_maximo). Si no hay nadie, (0, None).
    """
    solapan = (estadias["Inicio"] < day_end) & (estadias["Fin"].isna() | (estadias["Fin"] > day_start))
    sub = estadias.loc[solapan]
    if sub.empty:
        return 0, None

    inicio_clip = sub["Inicio"].clip(lower=day_start)
    fin_clip = sub["Fin"].fillna(day_end).clip(upper=day_end)

    eventos = list(zip(inicio_clip, [1] * len(sub))) + list(zip(fin_clip, [-1] * len(sub)))
    # A igual tiempo, primero las salidas (-1) y después las entradas (+1),
    # coherente con el intervalo semi-abierto [Inicio, Fin).
    eventos.sort(key=lambda e: (e[0], e[1]))

    contador = 0
    maximo = 0
    hora_maximo = day_start
    for tiempo, delta in eventos:
        contador += delta
        if contador > maximo:
            maximo = contador
            hora_maximo = tiempo
    return maximo, hora_maximo


def _combos_segmento_empresa(segmentos, empresas):
    return (
        [([s], s) for s in segmentos] + [(list(segmentos), "TOTAL")],
        [([e], e) for e in empresas] + [(list(empresas), "TODAS")],
    )


def calcular_pob_diario(estadias: pd.DataFrame) -> pd.DataFrame:
    segmentos = sorted(estadias["Segmento"].dropna().unique())
    empresas = sorted(estadias["Empresa"].dropna().unique())
    combos_seg, combos_emp = _combos_segmento_empresa(segmentos, empresas)
    dias = _rango_dias(estadias)

    filas = []
    for dia in dias:
        day_start, day_end = dia, dia + pd.Timedelta(days=1)
        solapan_dia = (estadias["Inicio"] < day_end) & (estadias["Fin"].isna() | (estadias["Fin"] > day_start))
        base_dia = estadias.loc[solapan_dia]

        for seg_list, seg_label in combos_seg:
            for emp_list, emp_label in combos_emp:
                m = base_dia["Segmento"].isin(seg_list) & base_dia["Empresa"].isin(emp_list)
                sub = base_dia.loc[m]
                maximo, hora_max = _maximo_simultaneo(sub, day_start, day_end)
                filas.append({
                    "Fecha": dia, "Segmento": seg_label, "Empresa": emp_label,
                    "Presentes_en_el_dia": sub["IdPersona"].nunique(),
                    "Maximo_simultaneo": maximo,
                    "Hora_del_maximo": hora_max,
                })
    return pd.DataFrame(filas)


def calcular_pob_horario(estadias: pd.DataFrame, horas_foto=cfg.HORAS_FOTO) -> pd.DataFrame:
    segmentos = sorted(estadias["Segmento"].dropna().unique())
    empresas = sorted(estadias["Empresa"].dropna().unique())
    combos_seg, combos_emp = _combos_segmento_empresa(segmentos, empresas)
    dias = _rango_dias(estadias)

    filas = []
    for dia in dias:
        for hora_txt in horas_foto:
            h, m = (int(x) for x in hora_txt.split(":"))
            instante = dia + pd.Timedelta(hours=h, minutes=m)
            base = estadias.loc[_presentes_en_instante(estadias, instante)]
            for seg_list, seg_label in combos_seg:
                for emp_list, emp_label in combos_emp:
                    m2 = base["Segmento"].isin(seg_list) & base["Empresa"].isin(emp_list)
                    filas.append({
                        "Fecha": dia, "Hora": hora_txt, "Segmento": seg_label, "Empresa": emp_label,
                        "POB": base.loc[m2, "IdPersona"].nunique(),
                    })
    return pd.DataFrame(filas)


def calcular_pob_franjas(estadias: pd.DataFrame, franjas=cfg.FRANJAS) -> pd.DataFrame:
    segmentos = sorted(estadias["Segmento"].dropna().unique())
    empresas = sorted(estadias["Empresa"].dropna().unique())
    combos_seg, combos_emp = _combos_segmento_empresa(segmentos, empresas)
    dias = _rango_dias(estadias)

    filas = []
    for dia in dias:
        for franja in franjas:
            hc, mc = (int(x) for x in franja["HoraCierre"].split(":"))
            instante = dia + pd.Timedelta(hours=hc, minutes=mc)
            base = estadias.loc[_presentes_en_instante(estadias, instante)]
            for seg_list, seg_label in combos_seg:
                for emp_list, emp_label in combos_emp:
                    m2 = base["Segmento"].isin(seg_list) & base["Empresa"].isin(emp_list)
                    filas.append({
                        "Fecha": dia, "Franja": franja["Franja"], "Segmento": seg_label, "Empresa": emp_label,
                        "POB": base.loc[m2, "IdPersona"].nunique(),
                    })
    return pd.DataFrame(filas)


def _presentes_en_ventana(estadias: pd.DataFrame, ventana_inicio: pd.Timestamp, ventana_fin: pd.Timestamp) -> pd.Series:
    """Máscara booleana: estadías que SOLAPAN la ventana [ventana_inicio, ventana_fin]
    (estuvo adentro en algún momento de la ventana), a diferencia de
    _presentes_en_instante que es una foto puntual."""
    return (estadias["Inicio"] <= ventana_fin) & (estadias["Fin"].isna() | (estadias["Fin"] > ventana_inicio))


def calcular_foto_detalle(estadias: pd.DataFrame, datos_persona: pd.DataFrame,
                           horas=cfg.HORAS_FOTO_DETALLE) -> pd.DataFrame:
    """Detalle (una fila por persona) de quién está presente en cada hora
    puntual de `horas`, para armar una tabla dinámica. A diferencia de
    calcular_pob_horario (que sólo cuenta), acá se listan IdPersona,
    Apellido, Nombre, Segmento y Empresa de cada presente.
    """
    dias = _rango_dias(estadias)
    partes = []
    for dia in dias:
        for hora_txt in horas:
            h, m = (int(x) for x in hora_txt.split(":"))
            instante = dia + pd.Timedelta(hours=h, minutes=m)
            presentes = estadias.loc[_presentes_en_instante(estadias, instante), ["IdPersona", "Segmento"]]
            presentes = presentes.drop_duplicates().merge(datos_persona, on="IdPersona", how="left")
            presentes.insert(0, "Hora", hora_txt)
            presentes.insert(0, "Fecha", dia)
            partes.append(presentes)
    columnas = ["Fecha", "Hora", "IdPersona", "Apellido", "Nombre", "Segmento", "Empresa"]
    if not partes:
        return pd.DataFrame(columns=columnas)
    out = pd.concat(partes, ignore_index=True)
    out["Empresa"] = out["Empresa"].fillna("SIN EMPRESA")
    return out[columnas].sort_values(["Fecha", "Hora", "Segmento", "Apellido"]).reset_index(drop=True)


def calcular_periodos_detalle(estadias: pd.DataFrame, datos_persona: pd.DataFrame,
                               estadios=cfg.ESTADIOS) -> pd.DataFrame:
    """Detalle (una fila por persona) de quién estuvo presente en algún
    momento de cada estadio del día (Mañana/Tarde/Noche/Madrugada), para
    armar una tabla dinámica. Usa solapamiento de ventana (_presentes_en_ventana),
    no una foto puntual: "estuvo adentro en algún instante de esa franja".
    """
    dias = _rango_dias(estadias)
    partes = []
    for dia in dias:
        for estadio in estadios:
            hi, mi = (int(x) for x in estadio["HoraInicio"].split(":"))
            hf, mf = (int(x) for x in estadio["HoraFin"].split(":"))
            ventana_inicio = dia + pd.Timedelta(hours=hi, minutes=mi)
            ventana_fin = dia + pd.Timedelta(hours=hf, minutes=mf)
            presentes = estadias.loc[
                _presentes_en_ventana(estadias, ventana_inicio, ventana_fin), ["IdPersona", "Segmento"]
            ]
            presentes = presentes.drop_duplicates().merge(datos_persona, on="IdPersona", how="left")
            presentes.insert(0, "Orden", estadio["Orden"])
            presentes.insert(0, "Estadio", estadio["Estadio"])
            presentes.insert(0, "Fecha", dia)
            partes.append(presentes)
    columnas = ["Fecha", "Estadio", "Orden", "IdPersona", "Apellido", "Nombre", "Segmento", "Empresa"]
    if not partes:
        return pd.DataFrame(columns=columnas)
    out = pd.concat(partes, ignore_index=True)
    out["Empresa"] = out["Empresa"].fillna("SIN EMPRESA")
    return out[columnas].sort_values(["Fecha", "Orden", "Segmento", "Apellido"]).reset_index(drop=True)


def calcular_comedor(df_validos: pd.DataFrame, empresa_por_persona: pd.Series,
                      turnos=cfg.TURNOS_COMEDOR) -> pd.DataFrame:
    """Cuenta personas distintas con marca de COMEDOR dentro de cada turno (extremos inclusivos).

    Las marcas fuera de turno se ignoran. El segmento del comedor sale del
    lector (AP - COMEDOR / LLY - COMEDOR), no de la columna SEGMENTO de la
    persona (que podría diferir si alguien cruzó de segmento antes de comer).
    """
    marcas = df_validos.loc[df_validos["TipoLector"] == "COMEDOR"].copy()
    marcas["SegmentoComedor"] = marcas["Lector"].map(cfg.SEGMENTO_COMEDOR)
    marcas["Empresa"] = marcas["IdPersona"].map(empresa_por_persona).fillna("SIN EMPRESA")
    marcas["Hora_min"] = marcas["FechaHora"].dt.hour * 60 + marcas["FechaHora"].dt.minute
    marcas["Fecha"] = marcas["FechaHora"].dt.normalize()

    segmentos = sorted(marcas["SegmentoComedor"].dropna().unique())
    empresas = sorted(empresa_por_persona.unique())
    combos_seg, combos_emp = _combos_segmento_empresa(segmentos, empresas)
    dias = sorted(marcas["Fecha"].unique())

    filas = []
    for turno in turnos:
        hd, md = (int(x) for x in turno["Desde"].split(":"))
        hh, mh = (int(x) for x in turno["Hasta"].split(":"))
        desde_min, hasta_min = hd * 60 + md, hh * 60 + mh
        en_turno = marcas[(marcas["Hora_min"] >= desde_min) & (marcas["Hora_min"] <= hasta_min)]

        for dia in dias:
            base_dia = en_turno[en_turno["Fecha"] == dia]
            for seg_list, seg_label in combos_seg:
                for emp_list, emp_label in combos_emp:
                    m = base_dia["SegmentoComedor"].isin(seg_list) & base_dia["Empresa"].isin(emp_list)
                    filas.append({
                        "Fecha": dia, "Turno": turno["Turno"], "Segmento": seg_label, "Empresa": emp_label,
                        "Personas_unicas": base_dia.loc[m, "IdPersona"].nunique(),
                    })
    return pd.DataFrame(filas)


def calcular_control_calidad(control_limpieza: dict, control_estadias: dict,
                              empresa_por_persona: pd.Series, n_archivos: int) -> pd.DataFrame:
    filas = [
        ("Archivos leídos", n_archivos),
        ("Filas totales (antes de deduplicar)", control_limpieza["filas_leidas"]),
        ("Duplicados eliminados", control_limpieza["duplicados_eliminados"]),
        ("Filas sin Fecha/Hora descartadas", control_limpieza["filas_sin_fecha_hora_descartadas"]),
        ("Filas con Fecha fuera de rango descartadas", control_limpieza["filas_fecha_fuera_de_rango_descartadas"]),
        ("Lectores sin clasificar", ", ".join(control_limpieza["lectores_sin_clasificar"]) or "(ninguno)"),
        ("Marcas sin IGG identificadas por credencial", control_limpieza["marcas_identificadas_por_credencial"]),
        ("Personas con más de una credencial", control_limpieza["personas_con_mas_de_una_credencial"]),
        ("Salidas sin ingreso previo", control_estadias["salidas_sin_ingreso_previo"]),
        ("Descartes de fin de mes", control_estadias["descartes_fin_de_mes"]),
        ("Personas SIN EMPRESA", int((empresa_por_persona == "SIN EMPRESA").sum())),
        ("Personas --- A DEFINIR ---", int((empresa_por_persona == "--- A DEFINIR ---").sum())),
    ]
    for evento, cantidad in control_limpieza["marcas_descartadas_por_evento"].items():
        filas.append((f"Marcas descartadas por evento: {evento}", cantidad))
    return pd.DataFrame(filas, columns=["Indicador", "Valor"])


# ===========================================================================
# 5) Orquestación completa
# ===========================================================================

def procesar_todo(carpeta: str, ahora: Optional[pd.Timestamp] = None):
    """Corre todo el pipeline sobre los archivos de `carpeta` y devuelve un dict
    con todas las tablas de salida más las estructuras intermedias (para poder
    inspeccionar/testear). `ahora` se pasa tal cual a construir_estadias
    (ver ese docstring); por defecto es el momento real de ejecución."""
    df_raw = cargar_archivos(carpeta)
    n_archivos = df_raw["ArchivoOrigen"].nunique()

    df_todo, control_limpieza = limpiar_datos(df_raw)
    empresa_por_persona = calcular_empresa_por_persona(df_todo)

    df_validos = df_todo.loc[df_todo["EventoValido"]].copy()
    estadias, control_estadias = construir_estadias(df_validos, ahora=ahora)
    estadias["Empresa"] = estadias["IdPersona"].map(empresa_por_persona).fillna("SIN EMPRESA")
    datos_persona = calcular_datos_persona(df_todo, empresa_por_persona)

    return {
        "df_todo": df_todo,
        "df_validos": df_validos,
        "empresa_por_persona": empresa_por_persona,
        "datos_persona": datos_persona,
        "estadias": estadias,
        "control_limpieza": control_limpieza,
        "control_estadias": control_estadias,
        "n_archivos": n_archivos,
        "POB_Diario": calcular_pob_diario(estadias),
        "POB_Horario": calcular_pob_horario(estadias),
        "POB_Franjas": calcular_pob_franjas(estadias),
        "Foto_13_22": calcular_foto_detalle(estadias, datos_persona),
        "POB_Periodos": calcular_periodos_detalle(estadias, datos_persona),
        "Comedor": calcular_comedor(df_validos, empresa_por_persona),
        "Control_Calidad": calcular_control_calidad(
            control_limpieza, control_estadias, empresa_por_persona, n_archivos
        ),
    }
