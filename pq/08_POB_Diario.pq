// =============================================================================
// 08_POB_Diario
// -----------------------------------------------------------------------------
// Fecha | Segmento | Empresa | Presentes_en_el_dia | Maximo_simultaneo | Hora_del_maximo
// Equivalente a calcular_pob_diario() en referencia.py. A diferencia de
// POB_Horario/POB_Franjas/Comedor, acá no alcanza con contar personas
// distintas: además hay que barrer la línea de tiempo del día para
// encontrar el pico de gente adentro al mismo tiempo, por eso no usa
// 07_fnExpandirGrupos y arma el cálculo combo por combo (Fecha x Segmento x
// Empresa), del mismo modo que el loop de Python.
// =============================================================================
let
    Estadias = #"06_Estadias",

    // --- Pico de concurrencia dentro de [dayStart, dayEnd) para un
    // subconjunto de estadías ya filtrado. Intervalo semi-abierto
    // [Inicio, Fin): quien sale justo a las 14:00 no cuenta en ese instante,
    // quien entra a esa hora sí. ---
    fnMaximoSimultaneo = (sub as table, dayStart as datetime, dayEnd as datetime) as record =>
        if Table.IsEmpty(sub) then [Maximo = 0, HoraMaximo = null] else
        let
            conClip = Table.AddColumn(
                Table.AddColumn(sub, "InicioClip", (fe) => if fe[Inicio] < dayStart then dayStart else fe[Inicio]),
                "FinClip", (fe) => if fe[Fin] = null then dayEnd else if fe[Fin] > dayEnd then dayEnd else fe[Fin]
            ),
            entradas = Table.RenameColumns(Table.AddColumn(Table.SelectColumns(conClip, {"InicioClip"}), "Delta", each 1), {"InicioClip", "Tiempo"}),
            salidas = Table.RenameColumns(Table.AddColumn(Table.SelectColumns(conClip, {"FinClip"}), "Delta", each -1), {"FinClip", "Tiempo"}),
            // A igual instante, primero las salidas y después las entradas
            // (semi-abierto: el que se va ya no cuenta, el que llega sí).
            ordenados = Table.Sort(Table.Combine({entradas, salidas}), {{"Tiempo", Order.Ascending}, {"Delta", Order.Ascending}}),
            resultado = List.Accumulate(
                Table.ToRecords(ordenados),
                [Contador = 0, Maximo = 0, HoraMaximo = dayStart],
                (estado, ev) =>
                    let nuevoContador = estado[Contador] + ev[Delta]
                    in
                        if nuevoContador > estado[Maximo] then
                            [Contador = nuevoContador, Maximo = nuevoContador, HoraMaximo = ev[Tiempo]]
                        else
                            [Contador = nuevoContador, Maximo = estado[Maximo], HoraMaximo = estado[HoraMaximo]]
            )
        in
            [Maximo = resultado[Maximo], HoraMaximo = resultado[HoraMaximo]],

    Segmentos = List.Distinct(List.RemoveNulls(Estadias[Segmento])),
    Empresas = List.Distinct(Estadias[Empresa]),

    FechaMinDatos = List.Min(Estadias[Inicio]),
    FinesNoNulos = List.RemoveNulls(Estadias[Fin]),
    TopeFin = if List.IsEmpty(FinesNoNulos) then List.Max(Estadias[Inicio]) else List.Max(FinesNoNulos),
    FechaMaxDatos = List.Max({List.Max(Estadias[Inicio]), TopeFin}),
    DiaInicio = DateTime.Date(FechaMinDatos),
    DiaFin = DateTime.Date(FechaMaxDatos),
    CantidadDias = Duration.Days(DiaFin - DiaInicio) + 1,
    Dias = List.Transform({0 .. CantidadDias - 1}, each Date.AddDays(DiaInicio, _)),

    // Combos de Segmento: cada valor real + TOTAL (todos juntos).
    CombosSegmento = List.Combine({
        List.Transform(Segmentos, (s) => [Etiqueta = s, Valores = {s}]),
        {[Etiqueta = "TOTAL", Valores = Segmentos]}
    }),
    // Combos de Empresa: cada valor real + TODAS (todas juntas).
    CombosEmpresa = List.Combine({
        List.Transform(Empresas, (e) => [Etiqueta = e, Valores = {e}]),
        {[Etiqueta = "TODAS", Valores = Empresas]}
    }),

    TablaDias = Table.FromList(Dias, Splitter.SplitByNothing(), {"Fecha"}),
    TablaSeg = Table.FromRecords(CombosSegmento),
    TablaEmp = Table.FromRecords(CombosEmpresa),

    ConSeg = Table.AddColumn(TablaDias, "Seg", each TablaSeg),
    ExpSeg = Table.ExpandTableColumn(ConSeg, "Seg", {"Etiqueta", "Valores"}, {"SegmentoLabel", "SegmentoValores"}),
    ConEmp = Table.AddColumn(ExpSeg, "Emp", each TablaEmp),
    ExpEmp = Table.ExpandTableColumn(ConEmp, "Emp", {"Etiqueta", "Valores"}, {"EmpresaLabel", "EmpresaValores"}),

    ConResultado = Table.AddColumn(ExpEmp, "Resultado", (filaActual) =>
        let
            dayStart = DateTime.From(filaActual[Fecha]),
            dayEnd = DateTime.From(Date.AddDays(filaActual[Fecha], 1)),
            solapan = Table.SelectRows(Estadias, (fe) =>
                List.Contains(filaActual[SegmentoValores], fe[Segmento])
                and List.Contains(filaActual[EmpresaValores], fe[Empresa])
                and fe[Inicio] < dayEnd
                and (fe[Fin] = null or fe[Fin] > dayStart)
            ),
            presentes = Table.RowCount(Table.Distinct(Table.SelectColumns(solapan, {"IdPersona"}))),
            maxSim = fnMaximoSimultaneo(solapan, dayStart, dayEnd)
        in
            [Presentes = presentes, Maximo = maxSim[Maximo], HoraMaximo = maxSim[HoraMaximo]]
    ),

    Expandido = Table.ExpandRecordColumn(ConResultado, "Resultado", {"Presentes", "Maximo", "HoraMaximo"}, {"Presentes_en_el_dia", "Maximo_simultaneo", "Hora_del_maximo"}),
    Renombrado = Table.RenameColumns(Expandido, {{"SegmentoLabel", "Segmento"}, {"EmpresaLabel", "Empresa"}}),
    Final = Table.SelectColumns(Renombrado, {"Fecha", "Segmento", "Empresa", "Presentes_en_el_dia", "Maximo_simultaneo", "Hora_del_maximo"})
in
    Final
