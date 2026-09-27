// =============================================================================
// 03_Limpieza
// -----------------------------------------------------------------------------
// A partir de 02_Origen: saca duplicados exactos, arma FechaHora, normaliza
// el lector, lo clasifica (INGRESO/EGRESO/COMEDOR/INTERNO), resuelve el
// segmento corto (AP/LLY) y arma IdPersona. Es el equivalente de
// `limpiar_datos()` en validacion/logica_pob.py: revisar ese archivo función
// por función si hay dudas sobre una regla puntual.
//
// Deja UNA fila por fichada (no filtra por evento válido todavía: eso lo
// hace 05_Validos, porque 04_Empresas necesita también las fichadas
// inválidas para conocer la empresa de la persona).
// =============================================================================
let
    // --- Duplicados: se consideran las 10 columnas originales (no
    // ArchivoOrigen, porque el mismo registro puede venir repetido en dos
    // archivos que se superponen en fechas). Se guarda la primera ocurrencia.
    ColumnasOriginales = {
        "Fecha", "Hora", "IGG", "Apellido", "Nombre", "Credencial",
        "Evento", "SEGMENTO", "EMPRESA", "Reader Description"
    },
    Origen = Table.Buffer(Table.Distinct(#"02_Origen", ColumnasOriginales)),

    // Filas sin Fecha o sin Hora no se pueden fichar (puede pasar en un
    // archivo de prueba armado a mano, con alguna fila vacía o incompleta
    // dentro del rango usado de la hoja). Se descartan acá, antes de
    // intentar armar FechaHora, para no romper con "cannot convert null".
    SinFilasIncompletas = Table.SelectRows(Origen, each [Fecha] <> null and [Hora] <> null),

    // --- FechaHora: soporta Fecha/Hora como texto "d/m/yyyy" y "HH:mm", y
    // también como fecha/hora nativa si Excel ya las autodetectó así. ---
    ConFechaHora = Table.AddColumn(SinFilasIncompletas, "FechaHora", each
        let
            valorFecha = [Fecha],
            fechaBase =
                if Value.Is(valorFecha, type date) or Value.Is(valorFecha, type datetime) then
                    DateTime.Date(valorFecha)
                else
                    let partes = Text.Split(Text.Trim(valorFecha), "/")
                    in #date(Number.FromText(partes{2}), Number.FromText(partes{1}), Number.FromText(partes{0})),
            valorHora = [Hora],
            horaBase =
                if Value.Is(valorHora, type time) or Value.Is(valorHora, type datetime) then
                    if Value.Is(valorHora, type datetime) then DateTime.Time(valorHora) else valorHora
                else
                    let partes = Text.Split(Text.Trim(valorHora), ":")
                    in #time(Number.FromText(partes{0}), Number.FromText(partes{1}), 0)
        in
            fechaBase & horaBase,
        type datetime
    ),

    // --- Lector normalizado: Trim + colapsar espacios dobles a uno solo ---
    // "?? """ antes de Text.Trim/Text.Split: esas funciones no toleran null
    // (tiran "cannot convert null to type Text"), y una fila con el lector
    // en blanco no tiene por qué frenar toda la consulta.
    ConLector = Table.AddColumn(ConFechaHora, "Lector", each
        Text.Combine(List.Select(Text.Split(Text.Trim([#"Reader Description"] ?? ""), " "), each _ <> ""), " "),
        type text
    ),

    // --- Clasificación de lector: join contra tLectores (Lector | Tipo) ---
    TablaLectores = Table.TransformColumns(
        Excel.CurrentWorkbook(){[Name="tLectores"]}[Content],
        {{"Lector", each Text.Combine(List.Select(Text.Split(Text.Trim(_ ?? ""), " "), each _ <> ""), " "), type text}}
    ),
    ConTipoLector = Table.NestedJoin(ConLector, {"Lector"}, TablaLectores, {"Lector"}, "_lectorRef", JoinKind.LeftOuter),
    ConTipoLector2 = Table.AddColumn(ConTipoLector, "TipoLector", each
        if [_lectorRef] = null or Table.IsEmpty([_lectorRef]) then "INTERNO" else [_lectorRef]{0}[Tipo],
        type text
    ),
    SinRefLector = Table.RemoveColumns(ConTipoLector2, {"_lectorRef"}),

    // --- Segmento corto: AGUADA PICHANA -> AP, LOMA LAS YEGUAS -> LLY ---
    ConSegmento = Table.AddColumn(SinRefLector, "Segmento", each
        if [SEGMENTO] = "AGUADA PICHANA" then "AP"
        else if [SEGMENTO] = "LOMA LAS YEGUAS" then "LLY"
        else null,
        type text
    ),

    // --- IdPersona: IGG sin espacios, o "CRED_" & Credencial si no hay IGG ---
    ConIdPersona = Table.AddColumn(ConSegmento, "IdPersona", each
        let iggLimpio = if [IGG] = null then "" else Text.Trim([IGG])
        in if iggLimpio = "" then "CRED_" & Text.Trim([Credencial] ?? "") else iggLimpio,
        type text
    ),

    // --- Evento válido: está en tEventosValidos ---
    TablaEventos = Excel.CurrentWorkbook(){[Name="tEventosValidos"]}[Content],
    ListaEventosValidos = TablaEventos[Evento],
    ConEventoValido = Table.AddColumn(ConIdPersona, "EventoValido", each
        List.Contains(ListaEventosValidos, [Evento]), type logical
    )
in
    ConEventoValido
