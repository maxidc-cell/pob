# -*- coding: utf-8 -*-
"""
Configuración de referencia para el procesamiento de fichadas POB.

Este módulo es el equivalente en Python de la hoja "Config" del Excel
(tParametros, tHorasFoto, tFranjas, tTurnosComedor, tEventosValidos,
tLectores). Si el usuario cambia algo en esa hoja, hay que reflejar el
mismo cambio acá para que la validación siga comparando manzanas con
manzanas.
"""
import re

# ---------------------------------------------------------------------------
# tParametros
# ---------------------------------------------------------------------------
PREFIJO_ARCHIVO = "POB"
HOJA_ORIGEN = "04 - Fichadas entre fechas"

# ---------------------------------------------------------------------------
# tEventosValidos: únicos eventos que prueban presencia física.
# ---------------------------------------------------------------------------
EVENTOS_VALIDOS = [
    "Access Granted",
    "Access Granted: Reader Unlocked",
]

# ---------------------------------------------------------------------------
# tHorasFoto: horas del día en que se toma la "foto" horaria del POB.
# ---------------------------------------------------------------------------
HORAS_FOTO = [f"{h:02d}:00" for h in range(24)]

# ---------------------------------------------------------------------------
# tFranjas
# ---------------------------------------------------------------------------
FRANJAS = [
    {"Franja": "07-14", "HoraInicio": "07:00", "HoraCierre": "14:00"},
    {"Franja": "19-24", "HoraInicio": "19:00", "HoraCierre": "23:59"},
]

# ---------------------------------------------------------------------------
# tHorasFotoDetalle: horas puntuales para el reporte de detalle "Foto_13_22"
# (listado de personas, no sólo el conteo de tHorasFoto/POB_Horario).
# ---------------------------------------------------------------------------
HORAS_FOTO_DETALLE = ["13:00", "22:00"]

# ---------------------------------------------------------------------------
# tEstadios: franjas de 4 estadios del día para el reporte POB_Periodos.
# "Presente en el estadio" = estuvo dentro del yacimiento en algún momento de
# esa ventana (solapamiento), no una foto puntual.
# ---------------------------------------------------------------------------
ESTADIOS = [
    {"Estadio": "Mañana", "HoraInicio": "07:00", "HoraFin": "12:00", "Orden": 1},
    {"Estadio": "Tarde", "HoraInicio": "12:01", "HoraFin": "19:00", "Orden": 2},
    {"Estadio": "Noche", "HoraInicio": "19:01", "HoraFin": "23:59", "Orden": 3},
    {"Estadio": "Madrugada", "HoraInicio": "00:00", "HoraFin": "06:59", "Orden": 4},
]

# ---------------------------------------------------------------------------
# tTurnosComedor (extremos inclusivos)
# ---------------------------------------------------------------------------
TURNOS_COMEDOR = [
    {"Turno": "Desayuno", "Desde": "06:00", "Hasta": "08:00"},
    {"Turno": "Almuerzo", "Desde": "12:00", "Hasta": "14:00"},
    {"Turno": "Merienda", "Desde": "17:00", "Hasta": "19:00"},
    {"Turno": "Cena", "Desde": "20:30", "Hasta": "22:00"},
]


def normalizar_lector(nombre: str) -> str:
    """Recorta espacios y colapsa espacios dobles a uno solo.

    En la muestra hay lectores con doble espacio (p. ej.
    'LLY- Porteria Vehiculos SAL  02-1-05'); hay que aplicar esta misma
    normalización tanto a los datos de origen como a la tabla de
    clasificación de lectores para que matcheen.
    """
    return re.sub(r"\s+", " ", str(nombre).strip())


# ---------------------------------------------------------------------------
# tLectores: clasificación de cada lector. Todo lector no listado se trata
# como INTERNO (y se reporta en Control de Calidad como "sin clasificar").
# ---------------------------------------------------------------------------
_INGRESO = [
    "AP - PortExter ENT 01-1-01",
    "AP - PortPasVehic ENT 01-1-03",
    "LLY- Porteria Exterior ENT 02-1-02",
    "LLY- Porteria Vehiculos ENT 02-1-04",
    # Accesos secundarios a planta: personal que entra/sale EXCLUSIVAMENTE
    # por acá (confirmado contra datos reales de producción, no eran
    # lectores internos como se asumió al precargar la tabla por primera
    # vez). Ver README, "Supuestos y decisiones tomadas".
    "AP - PlantaAcc Puerta 4 ENT 08-0-00",
    "AP - PlantaAcc Puerta 6 ENT 09-0-00",
    "AP - PlantaAcc Puerta 9 ENT 07-0-00",
    "AP - PlantaAccPeato-ENT 04-1-01",
    "LLY - Proc - KM 0 ENT 09-0-02",
    "LLY - Proc - KM 0 ENT DESCONEC 09-0-00",
    "LLY - Proc - Porton - Mpl - ENT 01-0-01",
    "LLY - Proc - Porton 1 Zona GEKs ENT 07-0-00",
    "LLY - Proc - Porton 2 Zona Tanque ENT 08-0-02",
    "LLY - Proc - Puerta 1 Acceso Peatonal ENT 08-0-00",
]

_EGRESO = [
    "AP - PortExter SAL 01-0-05",
    "AP - PortPasVehic SAL 01-1-00",
    "LLY- Porteria Exterior SAL 02-1-03",
    "LLY- Porteria Vehiculos SAL 02-1-05",
    # Contraparte de salida de los accesos secundarios de arriba.
    "AP - PlantaAcc Puerta 4 SAL 08-0-01",
    "AP - PlantaAcc Puerta 6 SAL 09-0-01",
    "AP - PlantaAcc Puerta 9 SAL 07-0-01",
    "AP - PlantaAccPeato- SAL 04-1-00",
    "LLY - Proc - KM 0 SAL 09-0-01",
    "LLY - Proc - Porton - Mpl - SAL 01-0-02",
    "LLY - Proc - Porton 1 Zona GEKs SAL 07-0-01",
    "LLY - Proc - Porton 2 Zona Tanque SAL 08-0-03",
    "LLY - Proc - Puerta 1 Acceso Peatonal SAL 08-0-01",
    # Puertas de emergencia (MP_...): en el uso real sólo se marcan al
    # evacuar, nunca para entrar. Tratarlas como INTERNO dejaba a la gente
    # "presente" indefinidamente después de una evacuación real (verificado
    # con datos de producción: un grupo marcó varias puertas de emergencia
    # de LLY la noche del 5/9 y sin esta corrección seguían apareciendo
    # presentes 9 días después). Se tratan como EGRESO.
    "MP_AP - Emerg*Brigada 02-0-01",
    "MP_AP - Emerg*Cancha Basket 10-1-03",
    "MP_AP - Emerg*Comedor 13-0-00",
    "MP_AP - Emerg*Enfermeria 02-0-00",
    "MP_AP - Emerg*Hotel 03-0-01",
    "MP_AP - Emerg*Oficinas 03-0-00",
    "MP_AP - Emerg*Porteria 01-1-04",
    "MP_AP - Emerg*Puerta 6 09-0-02",
    "MP_AP - Emerg*Puerta 9 07-0-02",
    "MP_AP - Emerg*SalaCtrl 04-1-02",
    "MP_LLY- Emerg* Brigada 06-0-00",
    "MP_LLY- Emerg* Enfermeria 02-2-01",
    "MP_LLY- Emerg* Hotel 03-1-06",
    "MP_LLY- Emerg* Oficina PCS 05-1-05",
    "MP_LLY- Emerg* Porteria 02-2-00",
    "MP_LLY- Emerg* Sala control 05-1-04",
    "MP_LLY- Emerg* WorkShop 04-2-02",
]

_COMEDOR = [
    "AP - COMEDOR 10-0-01",
    "LLY - COMEDOR 03-1-07",
]

# Resto de los lectores de la muestra, precargados como INTERNO: habitaciones,
# oficinas y salas dentro del predio (no son accesos al yacimiento ni al
# comedor). Los accesos secundarios a planta y las puertas de emergencia
# están arriba, en _INGRESO/_EGRESO (ver comentarios ahí).
_INTERNO = [
    "AP - ALA ENSI ENT 05-0-00",
    "AP - ALA ENSI SAL 05-0-01",
    "AP - ALA TOTAL ENT 06-0-00",
    "AP - ALA TOTAL SAL 06-0-01",
    "AP - GIM ENT 10-0-00",
    "AP - GIM SAL 10-0-01",
    "AP - Isla Impr ENT 12-0-00",
    "AP - Isla Impr SAL 12-0-01",
    "AP - LABORATORIO ELECTRICO ENT 04-1-06",
    "AP - LABORATORIO ELECTRICO SAL 04-1-07",
    "AP - Oficinas Ent 02-0-02",
    "AP - Oficinas Sal 02-0-03",
    "AP - PlantaSalaCtrl- ENT 04-1-03",
    "AP - PlantaSalaCtrl- SAL 04-1-04",
    "AP - Sala Técnica LLP ENT 17-0-00",
    "AP - Sala Técnica LLP SAL 17-0-01",
    "AP - Sala Técnica MP/LP ENT 16-0-00",
    "AP - Sala Técnica MP/LP SAL 16-0-01",
    "AP - Salas Eléctricas GENERADOR ENT 14-1-04",
    "AP - Salas Eléctricas GENERADOR SAL 14-1-05",
    "AP - Salas Eléctricas Oeste ENT 14-1-00",
    "AP - Salas Eléctricas Oeste SAL 14-1-01",
    "AP - Taller Entrada 11-0-00",
    "AP - Taller Salida 11-0-01",
    "LLY- Deposito ENT 04-1-04",
    "LLY- Deposito SAL 04-1-05",
    "LLY- Gimnasio ENT 02-1-06",
    "LLY- Gimnasio SAL 02-1-07",
    "LLY- Hotel ALA ENSI ENT 03-1-02",
    "LLY- Hotel ALA ENSI SAL 03-1-03",
    "LLY- Hotel ALA TOTAL ENT 03-1-00",
    "LLY- Hotel ALA TOTAL SAL 03-1-01",
    "LLY- Laboratorio ENT 04-2-00",
    "LLY- Laboratorio SAL 04-2-01",
    "LLY- Oficinas ENT 05-1-00",
    "LLY- Oficinas SAL 05-1-01",
    "LLY- Sala Control Botinero ENT 05-1-02",
    "LLY- Sala Control Botinero SAL 05-1-03",
    "LLY- Sala Control ENT 05-1-06",
    "LLY- Sala Control SAL 05-1-07",
    "LLY- Sala IT ENT 04-1-00",
    "LLY- Sala IT SAL 04-1-01",
    "LLY- WorkShop ENT 04-1-02",
    "LLY- WorkShop SAL 04-1-03",
    "LLY-Entrada Pasillo SR1 ENT 10-0-02",
    "LLY-Sala Eléctrica SR1 ENT 10-0-00",
    "LLY-Sala Eléctrica SR1 SAL 10-0-01",
    "LLY-Sala Generador MP ENT 11-1-02",
    "LLY-Sala Generador MP SAL 11-1-03",
    "LLY-Sala HVAC MP ENT 11-1-04",
    "LLY-Sala HVAC MP SAL 11-1-05",
    "LLY-Sala Técnica MP ENT 11-1-00",
    "LLY-Sala Técnica MP SAL 11-1-01",
    "LLY-Salida Pasillo SR1 SAL 10-0-03",
]

LECTORES = {}
for _r in _INGRESO:
    LECTORES[normalizar_lector(_r)] = "INGRESO"
for _r in _EGRESO:
    LECTORES[normalizar_lector(_r)] = "EGRESO"
for _r in _COMEDOR:
    LECTORES[normalizar_lector(_r)] = "COMEDOR"
for _r in _INTERNO:
    LECTORES[normalizar_lector(_r)] = "INTERNO"

# Tipo asignado a un lector que aparece en los datos pero NO está precargado
# en tLectores (lector nuevo, todavía no clasificado por el usuario).
TIPO_POR_DEFECTO = "INTERNO"

# Mapeo de lector de comedor -> segmento del comedor (para el reporte Comedor)
SEGMENTO_COMEDOR = {
    normalizar_lector("AP - COMEDOR 10-0-01"): "AP",
    normalizar_lector("LLY - COMEDOR 03-1-07"): "LLY",
}

# ---------------------------------------------------------------------------
# Segmentos: etiqueta larga (columna SEGMENTO) -> etiqueta corta de salida
# ---------------------------------------------------------------------------
SEGMENTOS = {
    "AGUADA PICHANA": "AP",
    "LOMA LAS YEGUAS": "LLY",
}
