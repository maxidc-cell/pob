// =============================================================================
// 10_POB_Franjas
// -----------------------------------------------------------------------------
// Fecha | Franja | Segmento | Empresa | POB. Cuenta a quienes están adentro
// en la HoraCierre de cada franja (tFranjas), hayan marcado o no dentro de
// la franja. Equivalente a calcular_pob_franjas() en referencia.py.
// =============================================================================
let
    Estadias = #"06_Estadias",
    Franjas = Table.TransformColumns(
        Excel.CurrentWorkbook(){[Name="tFranjas"]}[Content],
        {{"Franja", Text.Trim, type text}, {"HoraInicio", Text.Trim, type text}, {"HoraCierre", Text.Trim, type text}}
    ),

    FechaMinDatos = List.Min(Estadias[Inicio]),
    FinesNoNulos = List.RemoveNulls(Estadias[Fin]),
    TopeFin = if List.IsEmpty(FinesNoNulos) then List.Max(Estadias[Inicio]) else List.Max(FinesNoNulos),
    FechaMaxDatos = List.Max({List.Max(Estadias[Inicio]), TopeFin}),
    DiaInicio = DateTime.Date(FechaMinDatos),
    DiaFin = DateTime.Date(FechaMaxDatos),
    CantidadDias = Duration.Days(DiaFin - DiaInicio) + 1,
    Dias = List.Transform({0 .. CantidadDias - 1}, each Date.AddDays(DiaInicio, _)),

    TablaDias = Table.FromList(Dias, Splitter.SplitByNothing(), {"Fecha"}),
    ConFranjas = Table.AddColumn(TablaDias, "F", each Franjas),
    Grilla = Table.ExpandTableColumn(ConFranjas, "F", {"Franja", "HoraCierre"}),
    ConInstante = Table.AddColumn(Grilla, "Instante", (fila) =>
        let
            partes = Text.Split(fila[HoraCierre], ":"),
            h = Number.FromText(partes{0}),
            m = Number.FromText(partes{1})
        in
            fila[Fecha] & #time(h, m, 0)
    ),

    ConPresentes = Table.AddColumn(ConInstante, "Presentes", (fila) =>
        Table.SelectRows(Estadias, (fe) => fe[Inicio] <= fila[Instante] and (fe[Fin] = null or fe[Fin] > fila[Instante]))
    ),
    Expandido = Table.ExpandTableColumn(ConPresentes, "Presentes", {"IdPersona", "Segmento", "Empresa"}),
    Presencia = Table.SelectColumns(Expandido, {"Fecha", "Franja", "IdPersona", "Segmento", "Empresa"}),

    Resultado = #"07_fnExpandirGrupos"(Presencia, {"Fecha", "Franja"}, "POB")
in
    Resultado
