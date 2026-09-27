// =============================================================================
// 09_POB_Horario
// -----------------------------------------------------------------------------
// Fecha | Hora | Segmento | Empresa | POB. Es la "foto" en cada hora de
// tHorasFoto. Equivalente a calcular_pob_horario() en referencia.py.
// =============================================================================
let
    Estadias = #"06_Estadias",
    HorasFotoTexto = Excel.CurrentWorkbook(){[Name="tHorasFoto"]}[Content][Hora],

    ColumnasSalida = type table [Fecha = date, Hora = text, Segmento = text, Empresa = text, POB = Int64.Type],

    Salida =
        // Sin estadías todavía (mes sin datos, o ninguna marca válida): no
        // hay rango de fechas que armar, se devuelve la tabla vacía.
        if Table.IsEmpty(Estadias) then
            #table(ColumnasSalida, {})
        else
        let
            FechaMinDatos = List.Min(Estadias[Inicio]),
            FinesNoNulos = List.RemoveNulls(Estadias[Fin]),
            TopeFin = if List.IsEmpty(FinesNoNulos) then List.Max(Estadias[Inicio]) else List.Max(FinesNoNulos),
            FechaMaxDatos = List.Max({List.Max(Estadias[Inicio]), TopeFin}),
            DiaInicio = DateTime.Date(FechaMinDatos),
            DiaFin = DateTime.Date(FechaMaxDatos),
            CantidadDias = Duration.Days(DiaFin - DiaInicio) + 1,
            Dias = List.Transform({0 .. CantidadDias - 1}, each Date.AddDays(DiaInicio, _)),

            TablaDias = Table.FromList(Dias, Splitter.SplitByNothing(), {"Fecha"}),
            TablaHoras = Table.FromList(HorasFotoTexto, Splitter.SplitByNothing(), {"Hora"}),
            ConHoras = Table.AddColumn(TablaDias, "H", each TablaHoras),
            Grilla = Table.ExpandTableColumn(ConHoras, "H", {"Hora"}),
            ConInstante = Table.AddColumn(Grilla, "Instante", (fila) =>
                let
                    partes = Text.Split(fila[Hora], ":"),
                    h = Number.FromText(partes{0}),
                    m = Number.FromText(partes{1})
                in
                    fila[Fecha] & #time(h, m, 0)
            ),

            // Por cada instante de la grilla, quiénes están adentro (Inicio <= instante < Fin).
            ConPresentes = Table.AddColumn(ConInstante, "Presentes", (fila) =>
                Table.SelectRows(Estadias, (fe) => fe[Inicio] <= fila[Instante] and (fe[Fin] = null or fe[Fin] > fila[Instante]))
            ),
            Expandido = Table.ExpandTableColumn(ConPresentes, "Presentes", {"IdPersona", "Segmento", "Empresa"}),
            Presencia = Table.SelectColumns(Expandido, {"Fecha", "Hora", "IdPersona", "Segmento", "Empresa"}),

            Resultado = #"07_fnExpandirGrupos"(Presencia, {"Fecha", "Hora"}, "POB")
        in
            Resultado
in
    Salida
