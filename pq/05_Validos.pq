// =============================================================================
// 05_Validos
// -----------------------------------------------------------------------------
// Sólo las fichadas con evento en tEventosValidos (por defecto: Access
// Granted y Access Granted: Reader Unlocked). Es la base para la máquina de
// estados (06_Estadias) y para el reporte de Comedor (11_Comedor).
// Table.Buffer fuerza el cálculo una sola vez en memoria, porque esta tabla
// se usa varias veces más adelante.
// =============================================================================
let
    Origen = #"03_Limpieza",
    SoloValidos = Table.SelectRows(Origen, each [EventoValido] = true),
    Columnas = Table.SelectColumns(SoloValidos, {"IdPersona", "FechaHora", "TipoLector", "Segmento", "Lector"}),
    Resultado = Table.Buffer(Columnas)
in
    Resultado
