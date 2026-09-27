// =============================================================================
// 02_Origen
// -----------------------------------------------------------------------------
// Se conecta a la carpeta de OneDrive for Business por conexión directa
// (SharePoint.Files), toma todos los .xlsx cuyo nombre empieza con el
// prefijo configurado, lee de cada uno la hoja "04 - Fichadas entre fechas"
// y los apila en una sola tabla. Agrega la columna ArchivoOrigen para
// trazabilidad (de qué archivo vino cada fila).
//
// IMPORTANTE: no usa rutas locales. La URL y la carpeta salen de la hoja
// Config (tParametros), así que si el sitio o la carpeta cambian, el
// usuario los edita ahí y no toca esta consulta.
// =============================================================================
let
    Parametros = #"01_Parametros",
    UrlSitio = Parametros[URL_Sitio_OneDrive],
    RutaCarpeta = Parametros[Ruta_Carpeta],
    Prefijo = Parametros[Prefijo_Archivo],
    Hoja = Parametros[Hoja_Origen],

    // SharePoint.Files devuelve TODOS los archivos del sitio (recursivo);
    // filtramos por carpeta, prefijo de nombre y extensión. La comparación
    // ignora mayúsculas/minúsculas (Ruta_Carpeta no necesita coincidir
    // exactamente con cómo SharePoint guarda el path). Si Ruta_Carpeta
    // queda vacía, no filtra por carpeta (sirve para ver todos los
    // archivos del sitio y confirmar el [Folder Path] real la primera vez).
    Origen = SharePoint.Files(UrlSitio, [ApiVersion = 15]),
    FiltradoCarpeta = Table.SelectRows(
        Origen,
        each Text.Contains([Folder Path], RutaCarpeta, Comparer.OrdinalIgnoreCase)
            and Text.StartsWith([Name], Prefijo, Comparer.OrdinalIgnoreCase)
            and Text.EndsWith([Name], ".xlsx", Comparer.OrdinalIgnoreCase)
    ),

    // Por cada archivo, abrimos el workbook y extraemos la hoja de fichadas.
    ConHojas = Table.AddColumn(FiltradoCarpeta, "HojaDatos", each
        let
            libro = Excel.Workbook([Content], null, true),
            hojaBuscada = Table.SelectRows(libro, each [Item] = Hoja and [Kind] = "Sheet")
        in
            if Table.IsEmpty(hojaBuscada) then null else hojaBuscada{0}[Data]
    ),

    SinArchivosVacios = Table.SelectRows(ConHojas, each [HojaDatos] <> null),

    // Cada [HojaDatos] es una tabla con la primera fila como encabezado real.
    ConEncabezados = Table.AddColumn(SinArchivosVacios, "HojaConEncabezado", each
        Table.PromoteHeaders([HojaDatos], [PromoteAllScalars = true])
    ),

    SoloNecesarias = Table.SelectColumns(ConEncabezados, {"Name", "HojaConEncabezado"}),
    Renombrada = Table.RenameColumns(SoloNecesarias, {"Name", "ArchivoOrigen"}),

    // Table.ExpandTableColumn necesita conocer las columnas de destino; las
    // tomamos de la primera hoja no vacía como referencia.
    ColumnasFichadas = {
        "Fecha", "Hora", "IGG", "Apellido", "Nombre", "Credencial",
        "Evento", "SEGMENTO", "EMPRESA", "Reader Description"
    },
    Expandida = Table.ExpandTableColumn(Renombrada, "HojaConEncabezado", ColumnasFichadas),

    // Fecha y Hora se dejan tal cual llegan (en la muestra son texto, pero
    // puede haber archivos donde Excel las autodetecte como fecha/hora
    // nativa): 03_Limpieza contempla ambos casos. El resto se fuerza a texto
    // para que los joins y comparaciones de más adelante sean consistentes.
    ColumnasTexto = List.RemoveItems(ColumnasFichadas, {"Fecha", "Hora"}),
    TiposTexto = Table.TransformColumnTypes(Expandida, List.Transform(ColumnasTexto, each {_, type text}))
in
    TiposTexto
