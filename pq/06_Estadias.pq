// =============================================================================
// 06_Estadias
// -----------------------------------------------------------------------------
// El corazón del modelo: convierte las fichadas válidas de cada persona en
// "estadías" (Segmento, Inicio, Fin, Motivo) aplicando la máquina de estados
// de la sección 4 del prompt. Es el equivalente exacto de
// construir_estadias() en validacion/logica_pob.py — cualquier duda sobre
// una regla puntual, comparar función por función con ese archivo.
//
// Estrategia (pensada para performance con ~90.000 filas/mes):
//   1) Se arma una lista de "checkpoints" de fin de mes (00:00 del día 1 de
//      cada mes, desde la primera fichada hasta el momento de refrescar la
//      consulta) y se cruzan con cada persona -> eventos virtuales.
//   2) Se agrupan TODOS los eventos (reales + checkpoints) por persona en
//      una lista chica y ordenada.
//   3) Por persona se corre List.Accumulate (equivalente al loop de Python)
//      sobre esa lista chica. Nunca se comparan estadías entre sí a lo
//      ancho de toda la tabla (evita productos cartesianos grandes).
// =============================================================================
let
    EventosValidos = #"05_Validos",

    ColumnasSalida = type table [
        IdPersona = text, Segmento = text, Inicio = datetime, Fin = nullable datetime,
        Motivo = text, Empresa = text
    ],

    Salida =
        // Sin ninguna fichada válida (archivo de prueba vacío, o ningún
        // evento de tEventosValidos todavía) no hay nada que procesar: se
        // devuelve la tabla vacía en vez de calcular sobre listas vacías
        // (List.Min de una lista vacía da null y rompe todo lo que sigue).
        if Table.IsEmpty(EventosValidos) then
            #table(ColumnasSalida, {})
        else
        let
            // --- 1) Checkpoints de fin de mes ("en vivo": hasta el momento del refresco) ---
            FechaMin = List.Min(EventosValidos[FechaHora]),
    Ahora = DateTime.LocalNow(),
    PrimerCandidato = Date.AddMonths(Date.StartOfMonth(DateTime.Date(FechaMin)), 1),
    // Tope de seguridad: en uso normal (archivos cargados mes a mes) esto
    // nunca genera más que un puñado de controles. Si por algún dato
    // corrupto FechaMin quedara mal (un typo de año, por ejemplo), este
    // límite evita que la consulta intente generar miles de controles y se
    // cuelgue; 03_Limpieza además descarta fechas con un año disparatado
    // antes de que lleguen hasta acá.
    MaximoControlesFinDeMes = 240, // 20 años
    Boundaries = List.FirstN(
        List.Generate(
            () => PrimerCandidato,
            each DateTime.From(_) <= Ahora,
            each Date.AddMonths(_, 1)
        ),
        MaximoControlesFinDeMes
    ),
    BoundariesDT = List.Transform(Boundaries, each DateTime.From(_)),

    Personas = List.Distinct(EventosValidos[IdPersona]),
    TablaPersonas = Table.FromList(Personas, Splitter.SplitByNothing(), {"IdPersona"}),
    TablaBoundaries = Table.FromList(BoundariesDT, Splitter.SplitByNothing(), {"FechaHora"}),
    CruzadoCheckpoints = Table.AddColumn(TablaPersonas, "FechaHora", each TablaBoundaries),
    CheckpointsExpandido = Table.ExpandTableColumn(CruzadoCheckpoints, "FechaHora", {"FechaHora"}),
    Checkpoints = Table.AddColumn(
        Table.AddColumn(
            Table.AddColumn(CheckpointsExpandido, "TipoLector", each "CHECKPOINT"),
            "Segmento", each null
        ),
        "EsCheckpoint", each true
    ),

    // --- 2) Eventos reales + checkpoints, ordenados por persona y tiempo ---
    // A igual FechaHora, el checkpoint se procesa primero (OrdenEmpate = 0),
    // porque la marca real ya pertenece al mes nuevo.
    EventosReales = Table.AddColumn(
        Table.SelectColumns(EventosValidos, {"IdPersona", "FechaHora", "TipoLector", "Segmento"}),
        "EsCheckpoint", each false
    ),
    Todos = Table.Combine({EventosReales, Table.SelectColumns(Checkpoints, {"IdPersona", "FechaHora", "TipoLector", "Segmento", "EsCheckpoint"})}),
    ConOrdenEmpate = Table.AddColumn(Todos, "OrdenEmpate", each if [EsCheckpoint] then 0 else 1),
    Ordenados = Table.Sort(ConOrdenEmpate, {{"IdPersona", Order.Ascending}, {"FechaHora", Order.Ascending}, {"OrdenEmpate", Order.Ascending}}),

    Agrupado = Table.Group(Ordenados, {"IdPersona"}, {
        {"Eventos", each Table.ToRecords(Table.SelectColumns(_, {"FechaHora", "TipoLector", "Segmento", "EsCheckpoint"}))}
    }),

    // --- 3) Máquina de estados por persona ---
    // Estado: Estadias (lista acumulada), Adentro, Segmento actual, Apertura
    // (inicio de la estadía abierta), UltimaMarca (para la regla de fin de
    // mes) y MarcaPrevia (para "salida sin ingreso previo").
    fnProcesarPersona = (eventos as list) as list =>
        let
            resultado = List.Accumulate(
                eventos,
                [Estadias = {}, Adentro = false, Segmento = null, Apertura = null, UltimaMarca = null, MarcaPrevia = null],
                (estado, evento) =>
                    let
                        tiempo = evento[FechaHora],
                        tipo = evento[TipoLector],
                        segmento = evento[Segmento]
                    in
                        if evento[EsCheckpoint] then
                            // Fin de mes: descarta sólo si la marca que abrió la
                            // estadía sigue siendo la última marca de la persona
                            // (nadie volvió a marcar desde entonces, ni siquiera
                            // en el mismo segmento).
                            if estado[Adentro] and estado[UltimaMarca] = estado[Apertura] then
                                [
                                    Estadias = estado[Estadias] & {[Segmento = estado[Segmento], Inicio = estado[Apertura], Fin = tiempo, Motivo = "Descarte fin de mes"]},
                                    Adentro = false, Segmento = null, Apertura = null, UltimaMarca = null,
                                    MarcaPrevia = estado[MarcaPrevia]
                                ]
                            else
                                estado
                        else if tipo = "INGRESO" or tipo = "INTERNO" or tipo = "COMEDOR" then
                            if not estado[Adentro] then
                                [Estadias = estado[Estadias], Adentro = true, Segmento = segmento, Apertura = tiempo, UltimaMarca = tiempo, MarcaPrevia = tiempo]
                            else if segmento <> estado[Segmento] then
                                // Se mueve de segmento: cierra la estadía anterior y abre una nueva.
                                [
                                    Estadias = estado[Estadias] & {[Segmento = estado[Segmento], Inicio = estado[Apertura], Fin = tiempo, Motivo = "Normal"]},
                                    Adentro = true, Segmento = segmento, Apertura = tiempo, UltimaMarca = tiempo, MarcaPrevia = tiempo
                                ]
                            else
                                // Mismo segmento: sigue adentro, sólo refresca UltimaMarca.
                                [Estadias = estado[Estadias], Adentro = true, Segmento = estado[Segmento], Apertura = estado[Apertura], UltimaMarca = tiempo, MarcaPrevia = tiempo]
                        else if tipo = "EGRESO" then
                            if estado[Adentro] then
                                [
                                    Estadias = estado[Estadias] & {[Segmento = estado[Segmento], Inicio = estado[Apertura], Fin = tiempo, Motivo = "Normal"]},
                                    Adentro = false, Segmento = null, Apertura = null, UltimaMarca = null, MarcaPrevia = tiempo
                                ]
                            else
                                // Salida sin ingreso previo: se asume adentro desde
                                // lo último entre las 00:00 de hoy y su marca anterior.
                                let
                                    inicioDia = DateTime.From(DateTime.Date(tiempo)),
                                    inicio = if estado[MarcaPrevia] <> null and estado[MarcaPrevia] > inicioDia then estado[MarcaPrevia] else inicioDia
                                in
                                    [
                                        Estadias = estado[Estadias] & {[Segmento = segmento, Inicio = inicio, Fin = tiempo, Motivo = "Salida sin ingreso previo"]},
                                        Adentro = false, Segmento = null, Apertura = null, UltimaMarca = null, MarcaPrevia = tiempo
                                    ]
                        else
                            estado
            ),
            // Si al final del período sigue adentro, queda una estadía abierta (Fin = null).
            conAbierta =
                if resultado[Adentro] then
                    resultado[Estadias] & {[Segmento = resultado[Segmento], Inicio = resultado[Apertura], Fin = null, Motivo = "Abierta (sigue adentro)"]}
                else
                    resultado[Estadias]
        in
            conAbierta,

    ConEstadias = Table.AddColumn(Agrupado, "Estadias", each fnProcesarPersona([Eventos])),
    Expandido1 = Table.ExpandListColumn(ConEstadias, "Estadias"),
    Expandido2 = Table.ExpandRecordColumn(Expandido1, "Estadias", {"Segmento", "Inicio", "Fin", "Motivo"}),
    SoloNecesarias = Table.SelectColumns(Expandido2, {"IdPersona", "Segmento", "Inicio", "Fin", "Motivo"}),

    // --- Empresa de la persona (última informada, ver 04_Empresas) ---
    ConEmpresaJoin = Table.NestedJoin(SoloNecesarias, {"IdPersona"}, #"04_Empresas", {"IdPersona"}, "_emp", JoinKind.LeftOuter),
    ConEmpresa = Table.AddColumn(ConEmpresaJoin, "Empresa", each
        if [_emp] = null or Table.IsEmpty([_emp]) then "SIN EMPRESA" else [_emp]{0}[Empresa],
        type text
    ),
    Final = Table.Buffer(Table.RemoveColumns(ConEmpresa, {"_emp"}))
        in
            Final
in
    Salida
