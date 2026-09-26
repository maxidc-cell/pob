// =============================================================================
// 01_Parametros
// -----------------------------------------------------------------------------
// Lee la tabla tParametros de la hoja Config y la expone como un record, para
// poder escribir Parametros[URL_Sitio_OneDrive] en lugar de buscar la fila
// cada vez. Si el usuario agrega/renombra parámetros, esta consulta no hay
// que tocarla: solo hay que usar el nombre nuevo en el record resultante.
// =============================================================================
let
    Origen = Excel.CurrentWorkbook(){[Name="tParametros"]}[Content],
    TiposCambiados = Table.TransformColumnTypes(Origen, {{"Parametro", type text}, {"Valor", type text}}),
    // Convierte la tabla (Parametro | Valor) en un record { Parametro = Valor, ... }
    Registro = Record.FromList(TiposCambiados[Valor], TiposCambiados[Parametro])
in
    Registro
