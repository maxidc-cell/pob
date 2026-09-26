// =============================================================================
// 12_ControlCalidad
// -----------------------------------------------------------------------------
// Resumen de indicadores de calidad de datos. Equivalente a
// calcular_control_calidad() en referencia.py — mismos indicadores, mismos
// nombres de fila, para poder comparar línea por línea contra
// validacion/resultados_esperados.xlsx.
// =============================================================================
let
    Origen = #"02_Origen",
    Limpio = #"03_Limpieza",
    Estadias = #"06_Estadias",
    Empresas = #"04_Empresas",

    TablaLectores = Table.TransformColumns(
        Excel.CurrentWorkbook(){[Name="tLectores"]}[Content],
        {{"Lector", each Text.Combine(List.Select(Text.Split(Text.Trim(_), " "), each _ <> ""), " "), type text}}
    ),

    ArchivosLeidos = List.Count(List.Distinct(Origen[ArchivoOrigen])),
    FilasTotales = Table.RowCount(Origen),
    DuplicadosEliminados = FilasTotales - Table.RowCount(Limpio),

    LectoresUsados = List.Distinct(Limpio[Lector]),
    LectoresConocidos = TablaLectores[Lector],
    LectoresSinClasificar = List.Sort(List.Difference(LectoresUsados, LectoresConocidos)),
    TextoLectoresSinClasificar = if List.IsEmpty(LectoresSinClasificar) then "(ninguno)" else Text.Combine(LectoresSinClasificar, ", "),

    MarcasPorCredencial = Table.RowCount(Table.SelectRows(Limpio, each Text.StartsWith([IdPersona], "CRED_"))),

    ConIgg = Table.SelectRows(Limpio, each not Text.StartsWith([IdPersona], "CRED_")),
    PorPersona = Table.Group(ConIgg, {"IdPersona"}, {{"NCred", each List.Count(List.Distinct(_[Credencial])), Int64.Type}}),
    PersonasMultiCredencial = Table.RowCount(Table.SelectRows(PorPersona, each [NCred] > 1)),

    SalidasSinIngreso = Table.RowCount(Table.SelectRows(Estadias, each [Motivo] = "Salida sin ingreso previo")),
    DescartesFinDeMes = Table.RowCount(Table.SelectRows(Estadias, each [Motivo] = "Descarte fin de mes")),

    PersonasSinEmpresa = Table.RowCount(Table.SelectRows(Empresas, each [Empresa] = "SIN EMPRESA")),
    PersonasADefinir = Table.RowCount(Table.SelectRows(Empresas, each [Empresa] = "--- A DEFINIR ---")),

    Descartados = Table.SelectRows(Limpio, each [EventoValido] = false),
    PorEvento = Table.Group(Descartados, {"Evento"}, {{"Cantidad", each Table.RowCount(_), Int64.Type}}),
    FilasEvento = Table.AddColumn(PorEvento, "Indicador", each "Marcas descartadas por evento: " & [Evento]),
    FilasEventoFinal = Table.RenameColumns(Table.SelectColumns(FilasEvento, {"Indicador", "Cantidad"}), {"Cantidad", "Valor"}),

    Indicadores = #table(
        {"Indicador", "Valor"},
        {
            {"Archivos leídos", ArchivosLeidos},
            {"Filas totales (antes de deduplicar)", FilasTotales},
            {"Duplicados eliminados", DuplicadosEliminados},
            {"Lectores sin clasificar", TextoLectoresSinClasificar},
            {"Marcas sin IGG identificadas por credencial", MarcasPorCredencial},
            {"Personas con más de una credencial", PersonasMultiCredencial},
            {"Salidas sin ingreso previo", SalidasSinIngreso},
            {"Descartes de fin de mes", DescartesFinDeMes},
            {"Personas SIN EMPRESA", PersonasSinEmpresa},
            {"Personas --- A DEFINIR ---", PersonasADefinir}
        }
    ),
    Final = Table.Combine({Indicadores, FilasEventoFinal})
in
    Final
