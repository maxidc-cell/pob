// =============================================================================
// 11_Comedor
// -----------------------------------------------------------------------------
// Fecha | Turno | Segmento | Empresa | Personas_unicas. Cuenta personas
// distintas con marca en un lector COMEDOR dentro del turno (tTurnosComedor,
// extremos inclusivos). Las marcas fuera de turno se ignoran. Equivalente a
// calcular_comedor() en referencia.py.
//
// El segmento del comedor sale de la misma columna Segmento de la fichada
// (AP/LLY): en la muestra, todos los lectores "AP - COMEDOR..." tienen
// SEGMENTO = AGUADA PICHANA y todos los "LLY - COMEDOR..." tienen SEGMENTO =
// LOMA LAS YEGUAS (verificado 1 a 1), así que no hace falta una tabla de
// mapeo aparte.
// =============================================================================
let
    Origen = #"03_Limpieza",
    SoloComedor = Table.SelectRows(Origen, each [TipoLector] = "COMEDOR" and [EventoValido] = true),

    ConEmpresaJoin = Table.NestedJoin(SoloComedor, {"IdPersona"}, #"04_Empresas", {"IdPersona"}, "_emp", JoinKind.LeftOuter),
    ConEmpresa = Table.AddColumn(ConEmpresaJoin, "EmpresaPersona", (fila) =>
        if fila[_emp] = null or Table.IsEmpty(fila[_emp]) then "SIN EMPRESA" else fila[_emp]{0}[Empresa],
        type text
    ),

    Marcas = Table.SelectColumns(ConEmpresa, {"IdPersona", "FechaHora", "Segmento", "EmpresaPersona"}),
    MarcasRenombradas = Table.RenameColumns(Marcas, {"EmpresaPersona", "Empresa"}),
    ConFecha = Table.AddColumn(MarcasRenombradas, "Fecha", each DateTime.Date([FechaHora]), type date),
    ConMinutos = Table.AddColumn(ConFecha, "MinutosDia", each Time.Hour(DateTime.Time([FechaHora])) * 60 + Time.Minute(DateTime.Time([FechaHora])), Int64.Type),

    Turnos = Table.TransformColumns(
        Excel.CurrentWorkbook(){[Name="tTurnosComedor"]}[Content],
        {{"Turno", Text.Trim, type text}, {"Desde", Text.Trim, type text}, {"Hasta", Text.Trim, type text}}
    ),
    fnMinutos = (txt as text) as number =>
        let partes = Text.Split(txt, ":") in Number.FromText(partes{0}) * 60 + Number.FromText(partes{1}),

    ConMarcasPorTurno = Table.AddColumn(Turnos, "Marcas", (turno) =>
        let
            desdeMin = fnMinutos(turno[Desde]),
            hastaMin = fnMinutos(turno[Hasta])
        in
            Table.SelectRows(ConMinutos, (m) => m[MinutosDia] >= desdeMin and m[MinutosDia] <= hastaMin)
    ),
    Expandido = Table.ExpandTableColumn(ConMarcasPorTurno, "Marcas", {"Fecha", "IdPersona", "Segmento", "Empresa"}),
    Presencia = Table.SelectColumns(Expandido, {"Fecha", "Turno", "IdPersona", "Segmento", "Empresa"}),

    Resultado = #"07_fnExpandirGrupos"(Presencia, {"Fecha", "Turno"}, "Personas_unicas")
in
    Resultado
