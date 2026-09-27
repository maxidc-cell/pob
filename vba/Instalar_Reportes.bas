Attribute VB_Name = "Instalar_Reportes"
Option Explicit

' =============================================================================
' Instalar_Reportes
' -----------------------------------------------------------------------------
' Instala (o reinstala) todo el modelo de reportes POB/Comedor en el libro
' activo: crea las 12 consultas Power Query desde el código M embebido acá
' abajo (generado desde pq/*.pq por generar_macro.py), las carga como tabla
' en su hoja correspondiente, y arma un dashboard básico con tablas
' dinámicas, segmentaciones y gráficos.
'
' Es IDEMPOTENTE: se puede correr de nuevo las veces que haga falta (por
' ejemplo después de editar algo en la hoja Config) y reemplaza lo que ya
' existía en lugar de duplicarlo.
'
' Si algún paso del dashboard no se puede crear automáticamente en tu
' versión de Excel, la macro sigue con el resto y al final avisa qué faltó;
' GUIA_INSTALACION.md explica cómo armar esa parte a mano.
' =============================================================================

Sub Instalar_Reportes()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim advertencias As String
    advertencias = ""

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next

    ' --- 1) Crear/actualizar las 12 consultas ---
    CrearOActualizarConsulta wb, "01_Parametros", Codigo_01_Parametros()
    CrearOActualizarConsulta wb, "02_Origen", Codigo_02_Origen()
    CrearOActualizarConsulta wb, "03_Limpieza", Codigo_03_Limpieza()
    CrearOActualizarConsulta wb, "04_Empresas", Codigo_04_Empresas()
    CrearOActualizarConsulta wb, "05_Validos", Codigo_05_Validos()
    CrearOActualizarConsulta wb, "06_Estadias", Codigo_06_Estadias()
    CrearOActualizarConsulta wb, "07_fnExpandirGrupos", Codigo_07_fnExpandirGrupos()
    CrearOActualizarConsulta wb, "08_POB_Diario", Codigo_08_POB_Diario()
    CrearOActualizarConsulta wb, "09_POB_Horario", Codigo_09_POB_Horario()
    CrearOActualizarConsulta wb, "10_POB_Franjas", Codigo_10_POB_Franjas()
    CrearOActualizarConsulta wb, "11_Comedor", Codigo_11_Comedor()
    CrearOActualizarConsulta wb, "12_ControlCalidad", Codigo_12_ControlCalidad()

    ' --- 2) Cargar como tabla las que van a una hoja ---
    CargarComoTabla wb, "06_Estadias", "Estadias"
    CargarComoTabla wb, "08_POB_Diario", "POB_Diario"
    CargarComoTabla wb, "09_POB_Horario", "POB_Horario"
    CargarComoTabla wb, "10_POB_Franjas", "POB_Franjas"
    CargarComoTabla wb, "11_Comedor", "Comedor"
    CargarComoTabla wb, "12_ControlCalidad", "Control_Calidad"

    ' --- 3) Refrescar todo antes de armar el dashboard ---
    wb.RefreshAll
    On Error Resume Next
    Application.CalculateUntilAsyncQueriesDone
    On Error Resume Next

    ' --- 4) Dashboard: dinámicas, segmentaciones y gráficos ---
    If Not ArmarDashboard(wb) Then
        advertencias = advertencias & "- El dashboard no se pudo armar por completo automáticamente." & vbCrLf
    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0

    If advertencias = "" Then
        MsgBox "Instalación completa: consultas, tablas y dashboard armados.", vbInformation, "Reportes POB"
    Else
        MsgBox "Instalación completa, con observaciones:" & vbCrLf & vbCrLf & advertencias & _
            vbCrLf & "Revisá GUIA_INSTALACION.md (Camino B) para completar esas partes a mano.", _
            vbExclamation, "Reportes POB"
    End If
End Sub

' -----------------------------------------------------------------------------
' Crea la consulta `nombre` con el código M `formula`. Si ya existe, la
' reemplaza (idempotente) en vez de duplicarla.
' -----------------------------------------------------------------------------
Private Sub CrearOActualizarConsulta(wb As Workbook, nombre As String, formula As String)
    Dim q As WorkbookQuery
    On Error Resume Next
    Set q = wb.Queries(nombre)
    On Error GoTo 0
    If Not q Is Nothing Then
        q.Formula = formula
    Else
        wb.Queries.Add Name:=nombre, Formula:=formula
    End If
End Sub

' -----------------------------------------------------------------------------
' Carga la consulta `nombreConsulta` como tabla en la hoja `nombreHoja`,
' reemplazando cualquier tabla/conexión previa con ese destino (idempotente).
' -----------------------------------------------------------------------------
Private Sub CargarComoTabla(wb As Workbook, nombreConsulta As String, nombreHoja As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nombreHoja)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim eraOculta As Boolean
    eraOculta = (ws.Visible <> xlSheetVisible)
    ws.Visible = xlSheetVisible

    Dim i As Integer
    For i = ws.ListObjects.Count To 1 Step -1
        ws.ListObjects(i).Delete
    Next i
    ws.Cells.UnMerge
    ws.Cells.ClearContents

    BorrarConexionSiExiste wb, "Consulta - " & nombreConsulta
    BorrarConexionSiExiste wb, nombreConsulta

    Dim connString As String
    connString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & _
        nombreConsulta & ";Extended Properties="""""

    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(SourceType:=xlSrcExternal, Source:=connString, Destination:=ws.Range("$A$1"))
    lo.Name = "tbl_" & Replace(nombreConsulta, " ", "_")
    On Error Resume Next
    lo.TableStyle = "TableStyleMedium2"
    On Error GoTo 0

    If eraOculta Then ws.Visible = xlSheetVeryHidden
End Sub

Private Sub BorrarConexionSiExiste(wb As Workbook, nombre As String)
    Dim c As WorkbookConnection
    On Error Resume Next
    Set c = wb.Connections(nombre)
    If Not c Is Nothing Then c.Delete
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' Dashboard básico: una tabla dinámica por reporte (POB_Diario, POB_Horario,
' POB_Franjas, Comedor), segmentaciones de Segmento (selección única) y
' Empresa conectadas a las cuatro, y un gráfico por dinámica. Devuelve False
' si algo no se pudo crear (la instalación de consultas/tablas ya se hizo
' igual, esto es sólo la parte visual).
' -----------------------------------------------------------------------------
Private Function ArmarDashboard(wb As Workbook) As Boolean
    Dim ok As Boolean
    ok = True
    On Error Resume Next

    Dim wsDash As Worksheet
    Set wsDash = wb.Worksheets("Dashboard")
    If wsDash Is Nothing Then
        ArmarDashboard = False
        Exit Function
    End If

    ' Limpiar dinámicas y gráficos previos para poder recrear sin duplicar.
    Dim pt As PivotTable
    Do While wsDash.PivotTables.Count > 0
        wsDash.PivotTables(1).TableRange2.Clear
    Loop
    Dim co As ChartObject
    For Each co In wsDash.ChartObjects
        co.Delete
    Next co
    Dim sc As SlicerCache
    For Each sc In wb.SlicerCaches
        sc.Delete
    Next sc

    Dim ptDiario As PivotTable, ptHorario As PivotTable, ptFranjas As PivotTable, ptComedor As PivotTable
    Set ptDiario = CrearDinamica(wb, wsDash, "POB_Diario", "tbl_08_POB_Diario", wsDash.Range("A5"), "Segmento", "Empresa", "Fecha", "Presentes_en_el_dia")
    Set ptHorario = CrearDinamica(wb, wsDash, "POB_Horario", "tbl_09_POB_Horario", wsDash.Range("A25"), "Segmento", "Empresa", "Hora", "POB")
    Set ptFranjas = CrearDinamica(wb, wsDash, "POB_Franjas", "tbl_10_POB_Franjas", wsDash.Range("N5"), "Segmento", "Empresa", "Franja", "POB")
    Set ptComedor = CrearDinamica(wb, wsDash, "Comedor", "tbl_11_Comedor", wsDash.Range("N25"), "Segmento", "Empresa", "Turno", "Personas_unicas")

    If ptDiario Is Nothing Or ptHorario Is Nothing Or ptFranjas Is Nothing Or ptComedor Is Nothing Then
        ok = False
    End If

    ' Segmentaciones conectadas a las 4 dinámicas.
    Dim pts As Variant
    pts = Array(ptDiario, ptHorario, ptFranjas, ptComedor)
    AgregarSegmentacion wb, wsDash, pts, "Segmento", wsDash.Range("A1")
    AgregarSegmentacion wb, wsDash, pts, "Empresa", wsDash.Range("F1")

    ' NOTA: la segmentación de Segmento tiene que quedar en "Selección
    ' única" (para no mezclar filas TOTAL con AP/LLY), pero esa opción no
    ' tiene una propiedad simple y estable en el modelo de objetos de VBA
    ' entre versiones de Excel, así que queda como paso manual: click
    ' derecho sobre la segmentación "Segmento" > Configuración de
    ' segmentación de datos > tildar "Selección única" (ver GUIA_INSTALACION.md).

    ' Gráficos, uno por dinámica.
    AgregarGrafico wsDash, ptDiario, wsDash.Range("A45"), xlLine, "POB diario (presentes)"
    AgregarGrafico wsDash, ptHorario, wsDash.Range("A65"), xlLine, "POB horario"
    AgregarGrafico wsDash, ptFranjas, wsDash.Range("N45"), xlColumnClustered, "POB por franja"
    AgregarGrafico wsDash, ptComedor, wsDash.Range("N65"), xlColumnStacked, "Comedor por turno"

    On Error GoTo 0
    ArmarDashboard = ok
End Function

Private Function CrearDinamica(wb As Workbook, wsDash As Worksheet, nombrePT As String, nombreTabla As String, _
    destino As Range, campoSegmento As String, campoEmpresa As String, campoEje As String, campoValor As String) As PivotTable
    On Error Resume Next
    Dim loOrigen As ListObject
    Set loOrigen = Nothing
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Not loOrigen Is Nothing Then Exit For
        On Error Resume Next
        Set loOrigen = ws.ListObjects(nombreTabla)
        On Error Resume Next
    Next ws
    If loOrigen Is Nothing Then Exit Function

    Dim pc As PivotCache
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=loOrigen.Range)

    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=destino, TableName:="pt_" & nombrePT)

    pt.PivotFields(campoEje).Orientation = xlRowField
    pt.PivotFields(campoValor).Orientation = xlDataField
    pt.PivotFields(campoSegmento).Orientation = xlPageField
    pt.PivotFields(campoEmpresa).Orientation = xlPageField

    ' Filtro por defecto: Segmento = TOTAL, Empresa = TODAS.
    On Error Resume Next
    pt.PivotFields(campoSegmento).CurrentPage = "TOTAL"
    pt.PivotFields(campoEmpresa).CurrentPage = "TODAS"
    On Error GoTo 0

    Set CrearDinamica = pt
End Function

Private Sub AgregarSegmentacion(wb As Workbook, wsDash As Worksheet, pts As Variant, campo As String, destino As Range)
    On Error Resume Next
    Dim primerPt As PivotTable
    Set primerPt = pts(0)
    If primerPt Is Nothing Then Exit Sub

    Dim sc As SlicerCache
    Set sc = wb.SlicerCaches.Add2(primerPt, campo)
    sc.Slicers.Add wsDash, , "Slicer_" & campo, campo, destino.Top, destino.Left, 140, 150

    Dim i As Integer
    For i = 1 To UBound(pts)
        If Not pts(i) Is Nothing Then sc.PivotTables.AddPivotTable pts(i)
    Next i
    On Error GoTo 0
End Sub

Private Sub AgregarGrafico(wsDash As Worksheet, pt As PivotTable, destino As Range, tipo As XlChartType, titulo As String)
    On Error Resume Next
    If pt Is Nothing Then Exit Sub
    Dim co As ChartObject
    Set co = wsDash.ChartObjects.Add(destino.Left, destino.Top, 380, 220)
    co.Chart.SetSourceData pt.TableRange2
    co.Chart.ChartType = tipo
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = titulo
    On Error GoTo 0
End Sub

' =============================================================================
' Código M embebido de cada consulta (generado desde pq/*.pq).
' Si editás un archivo .pq, volvé a correr generar_macro.py para que esto
' quede sincronizado.
' =============================================================================


Private Function Codigo_01_Parametros() As String
    Dim s As String
    s = ""
    s = s & "// =============================================================================" & vbCrLf
    s = s & "// 01_Parametros" & vbCrLf
    s = s & "// -----------------------------------------------------------------------------" & vbCrLf
    s = s & "// Lee la tabla tParametros de la hoja Config y la expone como un record, para" & vbCrLf
    s = s & "// poder escribir Parametros[URL_Sitio_OneDrive] en lugar de buscar la fila" & vbCrLf
    s = s & "// cada vez. Si el usuario agrega/renombra parámetros, esta consulta no hay" & vbCrLf
    s = s & "// que tocarla: solo hay que usar el nombre nuevo en el record resultante." & vbCrLf
    s = s & "// =============================================================================" & vbCrLf
    s = s & "let" & vbCrLf
    s = s & "    Origen = Excel.CurrentWorkbook(){[Name=""tParametros""]}[Content]," & vbCrLf
    s = s & "    TiposCambiados = Table.TransformColumnTypes(Origen, {{""Parametro"", type text}, {""Valor"", type text}})," & vbCrLf
    s = s & "    // Convierte la tabla (Parametro | Valor) en un record { Parametro = Valor, ... }" & vbCrLf
    s = s & "    Registro = Record.FromList(TiposCambiados[Valor], TiposCambiados[Parametro])" & vbCrLf
    s = s & "in" & vbCrLf
    s = s & "    Registro" & vbCrLf
    Codigo_01_Parametros = s
End Function

Private Function Codigo_02_Origen() As String
    Dim s As String
    s = ""
    s = s & "// =============================================================================" & vbCrLf
    s = s & "// 02_Origen" & vbCrLf
    s = s & "// -----------------------------------------------------------------------------" & vbCrLf
    s = s & "// Se conecta a la carpeta de OneDrive for Business por conexión directa" & vbCrLf
    s = s & "// (SharePoint.Files), toma todos los .xlsx cuyo nombre empieza con el" & vbCrLf
    s = s & "// prefijo configurado, lee de cada uno la hoja ""04 - Fichadas entre fechas""" & vbCrLf
    s = s & "// y los apila en una sola tabla. Agrega la columna ArchivoOrigen para" & vbCrLf
    s = s & "// trazabilidad (de qué archivo vino cada fila)." & vbCrLf
    s = s & "//" & vbCrLf
    s = s & "// IMPORTANTE: no usa rutas locales. La URL y la carpeta salen de la hoja" & vbCrLf
    s = s & "// Config (tParametros), así que si el sitio o la carpeta cambian, el" & vbCrLf
    s = s & "// usuario los edita ahí y no toca esta consulta." & vbCrLf
    s = s & "// =============================================================================" & vbCrLf
    s = s & "let" & vbCrLf
    s = s & "    Parametros = #""01_Parametros""," & vbCrLf
    s = s & "    UrlSitio = Parametros[URL_Sitio_OneDrive]," & vbCrLf
    s = s & "    RutaCarpeta = Parametros[Ruta_Carpeta]," & vbCrLf
    s = s & "    Prefijo = Parametros[Prefijo_Archivo]," & vbCrLf
    s = s & "    Hoja = Parametros[Hoja_Origen]," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // SharePoint.Files devuelve TODOS los archivos del sitio (recursivo);" & vbCrLf
    s = s & "    // filtramos por carpeta, prefijo de nombre y extensión. La comparación" & vbCrLf
    s = s & "    // ignora mayúsculas/minúsculas (Ruta_Carpeta no necesita coincidir" & vbCrLf
    s = s & "    // exactamente con cómo SharePoint guarda el path). Si Ruta_Carpeta" & vbCrLf
    s = s & "    // queda vacía, no filtra por carpeta (sirve para ver todos los" & vbCrLf
    s = s & "    // archivos del sitio y confirmar el [Folder Path] real la primera vez)." & vbCrLf
    s = s & "    Origen = SharePoint.Files(UrlSitio, [ApiVersion = 15])," & vbCrLf
    s = s & "    FiltradoCarpeta = Table.SelectRows(" & vbCrLf
    s = s & "        Origen," & vbCrLf
    s = s & "        each Text.Contains([Folder Path], RutaCarpeta, Comparer.OrdinalIgnoreCase)" & vbCrLf
    s = s & "            and Text.StartsWith([Name], Prefijo, Comparer.OrdinalIgnoreCase)" & vbCrLf
    s = s & "            and Text.EndsWith([Name], "".xlsx"", Comparer.OrdinalIgnoreCase)" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // Por cada archivo, abrimos el workbook y extraemos la hoja de fichadas." & vbCrLf
    s = s & "    ConHojas = Table.AddColumn(FiltradoCarpeta, ""HojaDatos"", each" & vbCrLf
    s = s & "        let" & vbCrLf
    s = s & "            libro = Excel.Workbook([Content], null, true)," & vbCrLf
    s = s & "            hojaBuscada = Table.SelectRows(libro, each [Item] = Hoja and [Kind] = ""Sheet"")" & vbCrLf
    s = s & "        in" & vbCrLf
    s = s & "            if Table.IsEmpty(hojaBuscada) then null else hojaBuscada{0}[Data]" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SinArchivosVacios = Table.SelectRows(ConHojas, each [HojaDatos] <> null)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // Cada [HojaDatos] es una tabla con la primera fila como encabezado real." & vbCrLf
    s = s & "    ConEncabezados = Table.AddColumn(SinArchivosVacios, ""HojaConEncabezado"", each" & vbCrLf
    s = s & "        Table.PromoteHeaders([HojaDatos], [PromoteAllScalars = true])" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SoloNecesarias = Table.SelectColumns(ConEncabezados, {""Name"", ""HojaConEncabezado""})," & vbCrLf
    s = s & "    Renombrada = Table.RenameColumns(SoloNecesarias, {""Name"", ""ArchivoOrigen""})," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // Table.ExpandTableColumn necesita conocer las columnas de destino; las" & vbCrLf
    s = s & "    // tomamos de la primera hoja no vacía como referencia." & vbCrLf
    s = s & "    ColumnasFichadas = {" & vbCrLf
    s = s & "        ""Fecha"", ""Hora"", ""IGG"", ""Apellido"", ""Nombre"", ""Credencial""," & vbCrLf
    s = s & "        ""Evento"", ""SEGMENTO"", ""EMPRESA"", ""Reader Description""" & vbCrLf
    s = s & "    }," & vbCrLf
    s = s & "    Expandida = Table.ExpandTableColumn(Renombrada, ""HojaConEncabezado"", ColumnasFichadas)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // Fecha y Hora se dejan tal cual llegan (en la muestra son texto, pero" & vbCrLf
    s = s & "    // puede haber archivos donde Excel las autodetecte como fecha/hora" & vbCrLf
    s = s & "    // nativa): 03_Limpieza contempla ambos casos. El resto se fuerza a texto" & vbCrLf
    s = s & "    // para que los joins y comparaciones de más adelante sean consistentes." & vbCrLf
    s = s & "    ColumnasTexto = List.RemoveItems(ColumnasFichadas, {""Fecha"", ""Hora""})," & vbCrLf
    s = s & "    TiposTexto = Table.TransformColumnTypes(Expandida, List.Transform(ColumnasTexto, each {_, type text}))" & vbCrLf
    s = s & "in" & vbCrLf
    s = s & "    TiposTexto" & vbCrLf
    Codigo_02_Origen = s
End Function

Private Function Codigo_03_Limpieza() As String
    Dim s As String
    s = ""
    s = s & "// =============================================================================" & vbCrLf
    s = s & "// 03_Limpieza" & vbCrLf
    s = s & "// -----------------------------------------------------------------------------" & vbCrLf
    s = s & "// A partir de 02_Origen: saca duplicados exactos, arma FechaHora, normaliza" & vbCrLf
    s = s & "// el lector, lo clasifica (INGRESO/EGRESO/COMEDOR/INTERNO), resuelve el" & vbCrLf
    s = s & "// segmento corto (AP/LLY) y arma IdPersona. Es el equivalente de" & vbCrLf
    s = s & "// `limpiar_datos()` en validacion/logica_pob.py: revisar ese archivo función" & vbCrLf
    s = s & "// por función si hay dudas sobre una regla puntual." & vbCrLf
    s = s & "//" & vbCrLf
    s = s & "// Deja UNA fila por fichada (no filtra por evento válido todavía: eso lo" & vbCrLf
    s = s & "// hace 05_Validos, porque 04_Empresas necesita también las fichadas" & vbCrLf
    s = s & "// inválidas para conocer la empresa de la persona)." & vbCrLf
    s = s & "// =============================================================================" & vbCrLf
    s = s & "let" & vbCrLf
    s = s & "    // --- Duplicados: se consideran las 10 columnas originales (no" & vbCrLf
    s = s & "    // ArchivoOrigen, porque el mismo registro puede venir repetido en dos" & vbCrLf
    s = s & "    // archivos que se superponen en fechas). Se guarda la primera ocurrencia." & vbCrLf
    s = s & "    ColumnasOriginales = {" & vbCrLf
    s = s & "        ""Fecha"", ""Hora"", ""IGG"", ""Apellido"", ""Nombre"", ""Credencial""," & vbCrLf
    s = s & "        ""Evento"", ""SEGMENTO"", ""EMPRESA"", ""Reader Description""" & vbCrLf
    s = s & "    }," & vbCrLf
    s = s & "    Origen = Table.Buffer(Table.Distinct(#""02_Origen"", ColumnasOriginales))," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // Filas sin Fecha o sin Hora no se pueden fichar (puede pasar en un" & vbCrLf
    s = s & "    // archivo de prueba armado a mano, con alguna fila vacía o incompleta" & vbCrLf
    s = s & "    // dentro del rango usado de la hoja). Se descartan acá, antes de" & vbCrLf
    s = s & "    // intentar armar FechaHora, para no romper con ""cannot convert null""." & vbCrLf
    s = s & "    SinFilasIncompletas = Table.SelectRows(Origen, each [Fecha] <> null and [Hora] <> null)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // --- FechaHora: soporta Fecha/Hora como texto ""d/m/yyyy"" y ""HH:mm"", y" & vbCrLf
    s = s & "    // también como fecha/hora nativa si Excel ya las autodetectó así. ---" & vbCrLf
    s = s & "    ConFechaHora = Table.AddColumn(SinFilasIncompletas, ""FechaHora"", each" & vbCrLf
    s = s & "        let" & vbCrLf
    s = s & "            valorFecha = [Fecha]," & vbCrLf
    s = s & "            fechaBase =" & vbCrLf
    s = s & "                if Value.Is(valorFecha, type date) or Value.Is(valorFecha, type datetime) then" & vbCrLf
    s = s & "                    DateTime.Date(valorFecha)" & vbCrLf
    s = s & "                else" & vbCrLf
    s = s & "                    let partes = Text.Split(Text.Trim(valorFecha), ""/"")" & vbCrLf
    s = s & "                    in #date(Number.FromText(partes{2}), Number.FromText(partes{1}), Number.FromText(partes{0}))," & vbCrLf
    s = s & "            valorHora = [Hora]," & vbCrLf
    s = s & "            horaBase =" & vbCrLf
    s = s & "                if Value.Is(valorHora, type time) or Value.Is(valorHora, type datetime) then" & vbCrLf
    s = s & "                    if Value.Is(valorHora, type datetime) then DateTime.Time(valorHora) else valorHora" & vbCrLf
    s = s & "                else" & vbCrLf
    s = s & "                    let partes = Text.Split(Text.Trim(valorHora), "":"")" & vbCrLf
    s = s & "                    in #time(Number.FromText(partes{0}), Number.FromText(partes{1}), 0)" & vbCrLf
    s = s & "        in" & vbCrLf
    s = s & "            fechaBase & horaBase," & vbCrLf
    s = s & "        type datetime" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // --- Lector normalizado: Trim + colapsar espacios dobles a uno solo ---" & vbCrLf
    s = s & "    // ""?? """""" antes de Text.Trim/Text.Split: esas funciones no toleran null" & vbCrLf
    s = s & "    // (tiran ""cannot convert null to type Text""), y una fila con el lector" & vbCrLf
    s = s & "    // en blanco no tiene por qué frenar toda la consulta." & vbCrLf
    s = s & "    ConLector = Table.AddColumn(ConFechaHora, ""Lector"", each" & vbCrLf
    s = s & "        Text.Combine(List.Select(Text.Split(Text.Trim([#""Reader Description""] ?? """"), "" ""), each _ <> """"), "" "")," & vbCrLf
    s = s & "        type text" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // --- Clasificación de lector: join contra tLectores (Lector | Tipo) ---" & vbCrLf
    s = s & "    TablaLectores = Table.TransformColumns(" & vbCrLf
    s = s & "        Excel.CurrentWorkbook(){[Name=""tLectores""]}[Content]," & vbCrLf
    s = s & "        {{""Lector"", each Text.Combine(List.Select(Text.Split(Text.Trim(_ ?? """"), "" ""), each _ <> """"), "" ""), type text}}" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "    ConTipoLector = Table.NestedJoin(ConLector, {""Lector""}, TablaLectores, {""Lector""}, ""_lectorRef"", JoinKind.LeftOuter)," & vbCrLf
    s = s & "    ConTipoLector2 = Table.AddColumn(ConTipoLector, ""TipoLector"", each" & vbCrLf
    s = s & "        if [_lectorRef] = null or Table.IsEmpty([_lectorRef]) then ""INTERNO"" else [_lectorRef]{0}[Tipo]," & vbCrLf
    s = s & "        type text" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "    SinRefLector = Table.RemoveColumns(ConTipoLector2, {""_lectorRef""})," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // --- Segmento corto: AGUADA PICHANA -> AP, LOMA LAS YEGUAS -> LLY ---" & vbCrLf
    s = s & "    ConSegmento = Table.AddColumn(SinRefLector, ""Segmento"", each" & vbCrLf
    s = s & "        if [SEGMENTO] = ""AGUADA PICHANA"" then ""AP""" & vbCrLf
    s = s & "        else if [SEGMENTO] = ""LOMA LAS YEGUAS"" then ""LLY""" & vbCrLf
    s = s & "        else null," & vbCrLf
    s = s & "        type text" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // --- IdPersona: IGG sin espacios, o ""CRED_"" & Credencial si no hay IGG ---" & vbCrLf
    s = s & "    ConIdPersona = Table.AddColumn(ConSegmento, ""IdPersona"", each" & vbCrLf
    s = s & "        let iggLimpio = if [IGG] = null then """" else Text.Trim([IGG])" & vbCrLf
    s = s & "        in if iggLimpio = """" then ""CRED_"" & Text.Trim([Credencial] ?? """") else iggLimpio," & vbCrLf
    s = s & "        type text" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // --- Evento válido: está en tEventosValidos ---" & vbCrLf
    s = s & "    TablaEventos = Excel.CurrentWorkbook(){[Name=""tEventosValidos""]}[Content]," & vbCrLf
    s = s & "    ListaEventosValidos = TablaEventos[Evento]," & vbCrLf
    s = s & "    ConEventoValido = Table.AddColumn(ConIdPersona, ""EventoValido"", each" & vbCrLf
    s = s & "        List.Contains(ListaEventosValidos, [Evento]), type logical" & vbCrLf
    s = s & "    )" & vbCrLf
    s = s & "in" & vbCrLf
    s = s & "    ConEventoValido" & vbCrLf
    Codigo_03_Limpieza = s
End Function

Private Function Codigo_04_Empresas() As String
    Dim s As String
    s = ""
    s = s & "// =============================================================================" & vbCrLf
    s = s & "// 04_Empresas" & vbCrLf
    s = s & "// -----------------------------------------------------------------------------" & vbCrLf
    s = s & "// Última empresa informada por cada persona en todo el período (usa TODAS" & vbCrLf
    s = s & "// las fichadas, válidas o no: el dato de empresa viaja con la credencial" & vbCrLf
    s = s & "// más allá de si el acceso fue otorgado). Si la persona nunca tiene" & vbCrLf
    s = s & "// empresa, queda ""SIN EMPRESA"". ""--- A DEFINIR ---"" se respeta tal cual" & vbCrLf
    s = s & "// viene. Equivalente a calcular_empresa_por_persona() en referencia.py." & vbCrLf
    s = s & "// =============================================================================" & vbCrLf
    s = s & "let" & vbCrLf
    s = s & "    Origen = #""03_Limpieza""," & vbCrLf
    s = s & "    Agrupado = Table.Group(Origen, {""IdPersona""}, {" & vbCrLf
    s = s & "        {""Empresa"", each" & vbCrLf
    s = s & "            let" & vbCrLf
    s = s & "                ordenado = Table.Sort(_, {{""FechaHora"", Order.Ascending}})," & vbCrLf
    s = s & "                valor = ordenado{Table.RowCount(ordenado) - 1}[EMPRESA]" & vbCrLf
    s = s & "            in" & vbCrLf
    s = s & "                if valor = null or Text.Trim(valor) = """" then ""SIN EMPRESA"" else valor," & vbCrLf
    s = s & "            type text" & vbCrLf
    s = s & "        }" & vbCrLf
    s = s & "    })" & vbCrLf
    s = s & "in" & vbCrLf
    s = s & "    Agrupado" & vbCrLf
    Codigo_04_Empresas = s
End Function

Private Function Codigo_05_Validos() As String
    Dim s As String
    s = ""
    s = s & "// =============================================================================" & vbCrLf
    s = s & "// 05_Validos" & vbCrLf
    s = s & "// -----------------------------------------------------------------------------" & vbCrLf
    s = s & "// Sólo las fichadas con evento en tEventosValidos (por defecto: Access" & vbCrLf
    s = s & "// Granted y Access Granted: Reader Unlocked). Es la base para la máquina de" & vbCrLf
    s = s & "// estados (06_Estadias) y para el reporte de Comedor (11_Comedor)." & vbCrLf
    s = s & "// Table.Buffer fuerza el cálculo una sola vez en memoria, porque esta tabla" & vbCrLf
    s = s & "// se usa varias veces más adelante." & vbCrLf
    s = s & "// =============================================================================" & vbCrLf
    s = s & "let" & vbCrLf
    s = s & "    Origen = #""03_Limpieza""," & vbCrLf
    s = s & "    SoloValidos = Table.SelectRows(Origen, each [EventoValido] = true)," & vbCrLf
    s = s & "    Columnas = Table.SelectColumns(SoloValidos, {""IdPersona"", ""FechaHora"", ""TipoLector"", ""Segmento"", ""Lector""})," & vbCrLf
    s = s & "    Resultado = Table.Buffer(Columnas)" & vbCrLf
    s = s & "in" & vbCrLf
    s = s & "    Resultado" & vbCrLf
    Codigo_05_Validos = s
End Function

Private Function Codigo_06_Estadias() As String
    Dim s As String
    s = ""
    s = s & "// =============================================================================" & vbCrLf
    s = s & "// 06_Estadias" & vbCrLf
    s = s & "// -----------------------------------------------------------------------------" & vbCrLf
    s = s & "// El corazón del modelo: convierte las fichadas válidas de cada persona en" & vbCrLf
    s = s & "// ""estadías"" (Segmento, Inicio, Fin, Motivo) aplicando la máquina de estados" & vbCrLf
    s = s & "// de la sección 4 del prompt. Es el equivalente exacto de" & vbCrLf
    s = s & "// construir_estadias() en validacion/logica_pob.py — cualquier duda sobre" & vbCrLf
    s = s & "// una regla puntual, comparar función por función con ese archivo." & vbCrLf
    s = s & "//" & vbCrLf
    s = s & "// Estrategia (pensada para performance con ~90.000 filas/mes):" & vbCrLf
    s = s & "//   1) Se arma una lista de ""checkpoints"" de fin de mes (00:00 del día 1 de" & vbCrLf
    s = s & "//      cada mes, desde la primera fichada hasta el momento de refrescar la" & vbCrLf
    s = s & "//      consulta) y se cruzan con cada persona -> eventos virtuales." & vbCrLf
    s = s & "//   2) Se agrupan TODOS los eventos (reales + checkpoints) por persona en" & vbCrLf
    s = s & "//      una lista chica y ordenada." & vbCrLf
    s = s & "//   3) Por persona se corre List.Accumulate (equivalente al loop de Python)" & vbCrLf
    s = s & "//      sobre esa lista chica. Nunca se comparan estadías entre sí a lo" & vbCrLf
    s = s & "//      ancho de toda la tabla (evita productos cartesianos grandes)." & vbCrLf
    s = s & "// =============================================================================" & vbCrLf
    s = s & "let" & vbCrLf
    s = s & "    EventosValidos = #""05_Validos""," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ColumnasSalida = type table [" & vbCrLf
    s = s & "        IdPersona = text, Segmento = text, Inicio = datetime, Fin = nullable datetime," & vbCrLf
    s = s & "        Motivo = text, Empresa = text" & vbCrLf
    s = s & "    ]," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Salida =" & vbCrLf
    s = s & "        // Sin ninguna fichada válida (archivo de prueba vacío, o ningún" & vbCrLf
    s = s & "        // evento de tEventosValidos todavía) no hay nada que procesar: se" & vbCrLf
    s = s & "        // devuelve la tabla vacía en vez de calcular sobre listas vacías" & vbCrLf
    s = s & "        // (List.Min de una lista vacía da null y rompe todo lo que sigue)." & vbCrLf
    s = s & "        if Table.IsEmpty(EventosValidos) then" & vbCrLf
    s = s & "            #table(ColumnasSalida, {})" & vbCrLf
    s = s & "        else" & vbCrLf
    s = s & "        let" & vbCrLf
    s = s & "            // --- 1) Checkpoints de fin de mes (""en vivo"": hasta el momento del refresco) ---" & vbCrLf
    s = s & "            FechaMin = List.Min(EventosValidos[FechaHora])," & vbCrLf
    s = s & "    Ahora = DateTime.LocalNow()," & vbCrLf
    s = s & "    PrimerCandidato = Date.AddMonths(Date.StartOfMonth(DateTime.Date(FechaMin)), 1)," & vbCrLf
    s = s & "    Boundaries = List.Generate(" & vbCrLf
    s = s & "        () => PrimerCandidato," & vbCrLf
    s = s & "        each DateTime.From(_) <= Ahora," & vbCrLf
    s = s & "        each Date.AddMonths(_, 1)" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "    BoundariesDT = List.Transform(Boundaries, each DateTime.From(_))," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Personas = List.Distinct(EventosValidos[IdPersona])," & vbCrLf
    s = s & "    TablaPersonas = Table.FromList(Personas, Splitter.SplitByNothing(), {""IdPersona""})," & vbCrLf
    s = s & "    TablaBoundaries = Table.FromList(BoundariesDT, Splitter.SplitByNothing(), {""FechaHora""})," & vbCrLf
    s = s & "    CruzadoCheckpoints = Table.AddColumn(TablaPersonas, ""FechaHora"", each TablaBoundaries)," & vbCrLf
    s = s & "    CheckpointsExpandido = Table.ExpandTableColumn(CruzadoCheckpoints, ""FechaHora"", {""FechaHora""})," & vbCrLf
    s = s & "    Checkpoints = Table.AddColumn(" & vbCrLf
    s = s & "        Table.AddColumn(" & vbCrLf
    s = s & "            Table.AddColumn(CheckpointsExpandido, ""TipoLector"", each ""CHECKPOINT"")," & vbCrLf
    s = s & "            ""Segmento"", each null" & vbCrLf
    s = s & "        )," & vbCrLf
    s = s & "        ""EsCheckpoint"", each true" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // --- 2) Eventos reales + checkpoints, ordenados por persona y tiempo ---" & vbCrLf
    s = s & "    // A igual FechaHora, el checkpoint se procesa primero (OrdenEmpate = 0)," & vbCrLf
    s = s & "    // porque la marca real ya pertenece al mes nuevo." & vbCrLf
    s = s & "    EventosReales = Table.AddColumn(" & vbCrLf
    s = s & "        Table.SelectColumns(EventosValidos, {""IdPersona"", ""FechaHora"", ""TipoLector"", ""Segmento""})," & vbCrLf
    s = s & "        ""EsCheckpoint"", each false" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "    Todos = Table.Combine({EventosReales, Table.SelectColumns(Checkpoints, {""IdPersona"", ""FechaHora"", ""TipoLector"", ""Segmento"", ""EsCheckpoint""})})," & vbCrLf
    s = s & "    ConOrdenEmpate = Table.AddColumn(Todos, ""OrdenEmpate"", each if [EsCheckpoint] then 0 else 1)," & vbCrLf
    s = s & "    Ordenados = Table.Sort(ConOrdenEmpate, {{""IdPersona"", Order.Ascending}, {""FechaHora"", Order.Ascending}, {""OrdenEmpate"", Order.Ascending}})," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Agrupado = Table.Group(Ordenados, {""IdPersona""}, {" & vbCrLf
    s = s & "        {""Eventos"", each Table.ToRecords(Table.SelectColumns(_, {""FechaHora"", ""TipoLector"", ""Segmento"", ""EsCheckpoint""}))}" & vbCrLf
    s = s & "    })," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // --- 3) Máquina de estados por persona ---" & vbCrLf
    s = s & "    // Estado: Estadias (lista acumulada), Adentro, Segmento actual, Apertura" & vbCrLf
    s = s & "    // (inicio de la estadía abierta), UltimaMarca (para la regla de fin de" & vbCrLf
    s = s & "    // mes) y MarcaPrevia (para ""salida sin ingreso previo"")." & vbCrLf
    s = s & "    fnProcesarPersona = (eventos as list) as list =>" & vbCrLf
    s = s & "        let" & vbCrLf
    s = s & "            resultado = List.Accumulate(" & vbCrLf
    s = s & "                eventos," & vbCrLf
    s = s & "                [Estadias = {}, Adentro = false, Segmento = null, Apertura = null, UltimaMarca = null, MarcaPrevia = null]," & vbCrLf
    s = s & "                (estado, evento) =>" & vbCrLf
    s = s & "                    let" & vbCrLf
    s = s & "                        tiempo = evento[FechaHora]," & vbCrLf
    s = s & "                        tipo = evento[TipoLector]," & vbCrLf
    s = s & "                        segmento = evento[Segmento]" & vbCrLf
    s = s & "                    in" & vbCrLf
    s = s & "                        if evento[EsCheckpoint] then" & vbCrLf
    s = s & "                            // Fin de mes: descarta sólo si la marca que abrió la" & vbCrLf
    s = s & "                            // estadía sigue siendo la última marca de la persona" & vbCrLf
    s = s & "                            // (nadie volvió a marcar desde entonces, ni siquiera" & vbCrLf
    s = s & "                            // en el mismo segmento)." & vbCrLf
    s = s & "                            if estado[Adentro] and estado[UltimaMarca] = estado[Apertura] then" & vbCrLf
    s = s & "                                [" & vbCrLf
    s = s & "                                    Estadias = estado[Estadias] & {[Segmento = estado[Segmento], Inicio = estado[Apertura], Fin = tiempo, Motivo = ""Descarte fin de mes""]}," & vbCrLf
    s = s & "                                    Adentro = false, Segmento = null, Apertura = null, UltimaMarca = null," & vbCrLf
    s = s & "                                    MarcaPrevia = estado[MarcaPrevia]" & vbCrLf
    s = s & "                                ]" & vbCrLf
    s = s & "                            else" & vbCrLf
    s = s & "                                estado" & vbCrLf
    s = s & "                        else if tipo = ""INGRESO"" or tipo = ""INTERNO"" or tipo = ""COMEDOR"" then" & vbCrLf
    s = s & "                            if not estado[Adentro] then" & vbCrLf
    s = s & "                                [Estadias = estado[Estadias], Adentro = true, Segmento = segmento, Apertura = tiempo, UltimaMarca = tiempo, MarcaPrevia = tiempo]" & vbCrLf
    s = s & "                            else if segmento <> estado[Segmento] then" & vbCrLf
    s = s & "                                // Se mueve de segmento: cierra la estadía anterior y abre una nueva." & vbCrLf
    s = s & "                                [" & vbCrLf
    s = s & "                                    Estadias = estado[Estadias] & {[Segmento = estado[Segmento], Inicio = estado[Apertura], Fin = tiempo, Motivo = ""Normal""]}," & vbCrLf
    s = s & "                                    Adentro = true, Segmento = segmento, Apertura = tiempo, UltimaMarca = tiempo, MarcaPrevia = tiempo" & vbCrLf
    s = s & "                                ]" & vbCrLf
    s = s & "                            else" & vbCrLf
    s = s & "                                // Mismo segmento: sigue adentro, sólo refresca UltimaMarca." & vbCrLf
    s = s & "                                [Estadias = estado[Estadias], Adentro = true, Segmento = estado[Segmento], Apertura = estado[Apertura], UltimaMarca = tiempo, MarcaPrevia = tiempo]" & vbCrLf
    s = s & "                        else if tipo = ""EGRESO"" then" & vbCrLf
    s = s & "                            if estado[Adentro] then" & vbCrLf
    s = s & "                                [" & vbCrLf
    s = s & "                                    Estadias = estado[Estadias] & {[Segmento = estado[Segmento], Inicio = estado[Apertura], Fin = tiempo, Motivo = ""Normal""]}," & vbCrLf
    s = s & "                                    Adentro = false, Segmento = null, Apertura = null, UltimaMarca = null, MarcaPrevia = tiempo" & vbCrLf
    s = s & "                                ]" & vbCrLf
    s = s & "                            else" & vbCrLf
    s = s & "                                // Salida sin ingreso previo: se asume adentro desde" & vbCrLf
    s = s & "                                // lo último entre las 00:00 de hoy y su marca anterior." & vbCrLf
    s = s & "                                let" & vbCrLf
    s = s & "                                    inicioDia = DateTime.From(DateTime.Date(tiempo))," & vbCrLf
    s = s & "                                    inicio = if estado[MarcaPrevia] <> null and estado[MarcaPrevia] > inicioDia then estado[MarcaPrevia] else inicioDia" & vbCrLf
    s = s & "                                in" & vbCrLf
    s = s & "                                    [" & vbCrLf
    s = s & "                                        Estadias = estado[Estadias] & {[Segmento = segmento, Inicio = inicio, Fin = tiempo, Motivo = ""Salida sin ingreso previo""]}," & vbCrLf
    s = s & "                                        Adentro = false, Segmento = null, Apertura = null, UltimaMarca = null, MarcaPrevia = tiempo" & vbCrLf
    s = s & "                                    ]" & vbCrLf
    s = s & "                        else" & vbCrLf
    s = s & "                            estado" & vbCrLf
    s = s & "            )," & vbCrLf
    s = s & "            // Si al final del período sigue adentro, queda una estadía abierta (Fin = null)." & vbCrLf
    s = s & "            conAbierta =" & vbCrLf
    s = s & "                if resultado[Adentro] then" & vbCrLf
    s = s & "                    resultado[Estadias] & {[Segmento = resultado[Segmento], Inicio = resultado[Apertura], Fin = null, Motivo = ""Abierta (sigue adentro)""]}" & vbCrLf
    s = s & "                else" & vbCrLf
    s = s & "                    resultado[Estadias]" & vbCrLf
    s = s & "        in" & vbCrLf
    s = s & "            conAbierta," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ConEstadias = Table.AddColumn(Agrupado, ""Estadias"", each fnProcesarPersona([Eventos]))," & vbCrLf
    s = s & "    Expandido1 = Table.ExpandListColumn(ConEstadias, ""Estadias"")," & vbCrLf
    s = s & "    Expandido2 = Table.ExpandRecordColumn(Expandido1, ""Estadias"", {""Segmento"", ""Inicio"", ""Fin"", ""Motivo""})," & vbCrLf
    s = s & "    SoloNecesarias = Table.SelectColumns(Expandido2, {""IdPersona"", ""Segmento"", ""Inicio"", ""Fin"", ""Motivo""})," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // --- Empresa de la persona (última informada, ver 04_Empresas) ---" & vbCrLf
    s = s & "    ConEmpresaJoin = Table.NestedJoin(SoloNecesarias, {""IdPersona""}, #""04_Empresas"", {""IdPersona""}, ""_emp"", JoinKind.LeftOuter)," & vbCrLf
    s = s & "    ConEmpresa = Table.AddColumn(ConEmpresaJoin, ""Empresa"", each" & vbCrLf
    s = s & "        if [_emp] = null or Table.IsEmpty([_emp]) then ""SIN EMPRESA"" else [_emp]{0}[Empresa]," & vbCrLf
    s = s & "        type text" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "    Final = Table.Buffer(Table.RemoveColumns(ConEmpresa, {""_emp""}))" & vbCrLf
    s = s & "        in" & vbCrLf
    s = s & "            Final" & vbCrLf
    s = s & "in" & vbCrLf
    s = s & "    Salida" & vbCrLf
    Codigo_06_Estadias = s
End Function

Private Function Codigo_07_fnExpandirGrupos() As String
    Dim s As String
    s = ""
    s = s & "// =============================================================================" & vbCrLf
    s = s & "// 07_fnExpandirGrupos" & vbCrLf
    s = s & "// -----------------------------------------------------------------------------" & vbCrLf
    s = s & "// Función reutilizada por los reportes POB_Horario, POB_Franjas y Comedor." & vbCrLf
    s = s & "// A partir de una tabla con columnas [...columnasGrupo, IdPersona, Segmento," & vbCrLf
    s = s & "// Empresa], arma las filas de salida con Segmento en {valores..., TOTAL} x" & vbCrLf
    s = s & "// Empresa en {valores..., TODAS}, contando personas DISTINTAS (nunca" & vbCrLf
    s = s & "// sumando filas). Es el equivalente vectorizado de _combos_segmento_empresa" & vbCrLf
    s = s & "// + los nunique() de referencia.py." & vbCrLf
    s = s & "//" & vbCrLf
    s = s & "// NOTA para quien audite el M contra Python: a diferencia de" & vbCrLf
    s = s & "// validacion/logica_pob.py (que en algunos lugares itera combo por combo" & vbCrLf
    s = s & "// con nunique()), acá se usa Table.Group una sola vez por variante para que" & vbCrLf
    s = s & "// la consulta escale bien con ~90.000 filas/mes. El resultado numérico es" & vbCrLf
    s = s & "// idéntico; sólo cambia la forma de calcularlo." & vbCrLf
    s = s & "//" & vbCrLf
    s = s & "// Filas con conteo 0 no se generan (igual que en referencia.py): una" & vbCrLf
    s = s & "// combinación Fecha/Hora/Segmento/Empresa sin nadie presente simplemente no" & vbCrLf
    s = s & "// aparece en la tabla. Para una tabla dinámica esto no es un problema" & vbCrLf
    s = s & "// (Power Pivot no necesita la fila en cero para graficar bien un total)." & vbCrLf
    s = s & "// =============================================================================" & vbCrLf
    s = s & "let" & vbCrLf
    s = s & "    fnExpandirGrupos = (tabla as table, columnasGrupo as list, nombreValor as text) as table =>" & vbCrLf
    s = s & "        let" & vbCrLf
    s = s & "            contarDistintos = (t as table) as number => Table.RowCount(Table.Distinct(Table.SelectColumns(t, {""IdPersona""})))," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            PorSegmentoYEmpresa = Table.Group(tabla, columnasGrupo & {""Segmento"", ""Empresa""}, {{nombreValor, each contarDistintos(_), Int64.Type}})," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            PorSegmentoTodas = Table.Group(tabla, columnasGrupo & {""Segmento""}, {{nombreValor, each contarDistintos(_), Int64.Type}})," & vbCrLf
    s = s & "            PorSegmentoTodas2 = Table.AddColumn(PorSegmentoTodas, ""Empresa"", each ""TODAS"", type text)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            TotalPorEmpresa = Table.Group(tabla, columnasGrupo & {""Empresa""}, {{nombreValor, each contarDistintos(_), Int64.Type}})," & vbCrLf
    s = s & "            TotalPorEmpresa2 = Table.AddColumn(TotalPorEmpresa, ""Segmento"", each ""TOTAL"", type text)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            TotalTodas = Table.Group(tabla, columnasGrupo, {{nombreValor, each contarDistintos(_), Int64.Type}})," & vbCrLf
    s = s & "            TotalTodas2 = Table.AddColumn(Table.AddColumn(TotalTodas, ""Segmento"", each ""TOTAL"", type text), ""Empresa"", each ""TODAS"", type text)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            OrdenColumnas = columnasGrupo & {""Segmento"", ""Empresa"", nombreValor}," & vbCrLf
    s = s & "            Combinado = Table.Combine({" & vbCrLf
    s = s & "                Table.ReorderColumns(PorSegmentoYEmpresa, OrdenColumnas)," & vbCrLf
    s = s & "                Table.ReorderColumns(PorSegmentoTodas2, OrdenColumnas)," & vbCrLf
    s = s & "                Table.ReorderColumns(TotalPorEmpresa2, OrdenColumnas)," & vbCrLf
    s = s & "                Table.ReorderColumns(TotalTodas2, OrdenColumnas)" & vbCrLf
    s = s & "            })" & vbCrLf
    s = s & "        in" & vbCrLf
    s = s & "            Combinado" & vbCrLf
    s = s & "in" & vbCrLf
    s = s & "    fnExpandirGrupos" & vbCrLf
    Codigo_07_fnExpandirGrupos = s
End Function

Private Function Codigo_08_POB_Diario() As String
    Dim s As String
    s = ""
    s = s & "// =============================================================================" & vbCrLf
    s = s & "// 08_POB_Diario" & vbCrLf
    s = s & "// -----------------------------------------------------------------------------" & vbCrLf
    s = s & "// Fecha | Segmento | Empresa | Presentes_en_el_dia | Maximo_simultaneo | Hora_del_maximo" & vbCrLf
    s = s & "// Equivalente a calcular_pob_diario() en referencia.py. A diferencia de" & vbCrLf
    s = s & "// POB_Horario/POB_Franjas/Comedor, acá no alcanza con contar personas" & vbCrLf
    s = s & "// distintas: además hay que barrer la línea de tiempo del día para" & vbCrLf
    s = s & "// encontrar el pico de gente adentro al mismo tiempo, por eso no usa" & vbCrLf
    s = s & "// 07_fnExpandirGrupos y arma el cálculo combo por combo (Fecha x Segmento x" & vbCrLf
    s = s & "// Empresa), del mismo modo que el loop de Python." & vbCrLf
    s = s & "// =============================================================================" & vbCrLf
    s = s & "let" & vbCrLf
    s = s & "    Estadias = #""06_Estadias""," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // --- Pico de concurrencia dentro de [dayStart, dayEnd) para un" & vbCrLf
    s = s & "    // subconjunto de estadías ya filtrado. Intervalo semi-abierto" & vbCrLf
    s = s & "    // [Inicio, Fin): quien sale justo a las 14:00 no cuenta en ese instante," & vbCrLf
    s = s & "    // quien entra a esa hora sí. ---" & vbCrLf
    s = s & "    fnMaximoSimultaneo = (sub as table, dayStart as datetime, dayEnd as datetime) as record =>" & vbCrLf
    s = s & "        if Table.IsEmpty(sub) then [Maximo = 0, HoraMaximo = null] else" & vbCrLf
    s = s & "        let" & vbCrLf
    s = s & "            conClip = Table.AddColumn(" & vbCrLf
    s = s & "                Table.AddColumn(sub, ""InicioClip"", (fe) => if fe[Inicio] < dayStart then dayStart else fe[Inicio])," & vbCrLf
    s = s & "                ""FinClip"", (fe) => if fe[Fin] = null then dayEnd else if fe[Fin] > dayEnd then dayEnd else fe[Fin]" & vbCrLf
    s = s & "            )," & vbCrLf
    s = s & "            entradas = Table.RenameColumns(Table.AddColumn(Table.SelectColumns(conClip, {""InicioClip""}), ""Delta"", each 1), {""InicioClip"", ""Tiempo""})," & vbCrLf
    s = s & "            salidas = Table.RenameColumns(Table.AddColumn(Table.SelectColumns(conClip, {""FinClip""}), ""Delta"", each -1), {""FinClip"", ""Tiempo""})," & vbCrLf
    s = s & "            // A igual instante, primero las salidas y después las entradas" & vbCrLf
    s = s & "            // (semi-abierto: el que se va ya no cuenta, el que llega sí)." & vbCrLf
    s = s & "            ordenados = Table.Sort(Table.Combine({entradas, salidas}), {{""Tiempo"", Order.Ascending}, {""Delta"", Order.Ascending}})," & vbCrLf
    s = s & "            resultado = List.Accumulate(" & vbCrLf
    s = s & "                Table.ToRecords(ordenados)," & vbCrLf
    s = s & "                [Contador = 0, Maximo = 0, HoraMaximo = dayStart]," & vbCrLf
    s = s & "                (estado, ev) =>" & vbCrLf
    s = s & "                    let nuevoContador = estado[Contador] + ev[Delta]" & vbCrLf
    s = s & "                    in" & vbCrLf
    s = s & "                        if nuevoContador > estado[Maximo] then" & vbCrLf
    s = s & "                            [Contador = nuevoContador, Maximo = nuevoContador, HoraMaximo = ev[Tiempo]]" & vbCrLf
    s = s & "                        else" & vbCrLf
    s = s & "                            [Contador = nuevoContador, Maximo = estado[Maximo], HoraMaximo = estado[HoraMaximo]]" & vbCrLf
    s = s & "            )" & vbCrLf
    s = s & "        in" & vbCrLf
    s = s & "            [Maximo = resultado[Maximo], HoraMaximo = resultado[HoraMaximo]]," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ColumnasSalida = type table [" & vbCrLf
    s = s & "        Fecha = date, Segmento = text, Empresa = text," & vbCrLf
    s = s & "        Presentes_en_el_dia = Int64.Type, Maximo_simultaneo = Int64.Type, Hora_del_maximo = nullable datetime" & vbCrLf
    s = s & "    ]," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Resultado =" & vbCrLf
    s = s & "        // Sin estadías (mes sin datos todavía, o ninguna marca válida) no hay" & vbCrLf
    s = s & "        // rango de fechas que armar: se devuelve la tabla vacía directamente," & vbCrLf
    s = s & "        // en vez de romper en el List.Min/Max sobre una lista vacía." & vbCrLf
    s = s & "        if Table.IsEmpty(Estadias) then" & vbCrLf
    s = s & "            #table(ColumnasSalida, {})" & vbCrLf
    s = s & "        else" & vbCrLf
    s = s & "        let" & vbCrLf
    s = s & "            Segmentos = List.Distinct(List.RemoveNulls(Estadias[Segmento]))," & vbCrLf
    s = s & "            Empresas = List.Distinct(Estadias[Empresa])," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            FechaMinDatos = List.Min(Estadias[Inicio])," & vbCrLf
    s = s & "            FinesNoNulos = List.RemoveNulls(Estadias[Fin])," & vbCrLf
    s = s & "            TopeFin = if List.IsEmpty(FinesNoNulos) then List.Max(Estadias[Inicio]) else List.Max(FinesNoNulos)," & vbCrLf
    s = s & "            FechaMaxDatos = List.Max({List.Max(Estadias[Inicio]), TopeFin})," & vbCrLf
    s = s & "            DiaInicio = DateTime.Date(FechaMinDatos)," & vbCrLf
    s = s & "            DiaFin = DateTime.Date(FechaMaxDatos)," & vbCrLf
    s = s & "            CantidadDias = Duration.Days(DiaFin - DiaInicio) + 1," & vbCrLf
    s = s & "            Dias = List.Transform({0 .. CantidadDias - 1}, each Date.AddDays(DiaInicio, _))," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            // Combos de Segmento: cada valor real + TOTAL (todos juntos)." & vbCrLf
    s = s & "            CombosSegmento = List.Combine({" & vbCrLf
    s = s & "                List.Transform(Segmentos, (s) => [Etiqueta = s, Valores = {s}])," & vbCrLf
    s = s & "                {[Etiqueta = ""TOTAL"", Valores = Segmentos]}" & vbCrLf
    s = s & "            })," & vbCrLf
    s = s & "            // Combos de Empresa: cada valor real + TODAS (todas juntas)." & vbCrLf
    s = s & "            CombosEmpresa = List.Combine({" & vbCrLf
    s = s & "                List.Transform(Empresas, (e) => [Etiqueta = e, Valores = {e}])," & vbCrLf
    s = s & "                {[Etiqueta = ""TODAS"", Valores = Empresas]}" & vbCrLf
    s = s & "            })," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            // --- Optimización clave: filtrar Estadias por combo Segmento x" & vbCrLf
    s = s & "            // Empresa UNA sola vez (y guardarlo en memoria con Table.Buffer)," & vbCrLf
    s = s & "            // en vez de volver a filtrar toda la tabla por cada día. Con" & vbCrLf
    s = s & "            // ~70 empresas y un mes de datos, filtrar la tabla completa" & vbCrLf
    s = s & "            // (Fecha x Segmento x Empresa) en vez de (Segmento x Empresa)" & vbCrLf
    s = s & "            // multiplica el trabajo por la cantidad de días y hacía que la" & vbCrLf
    s = s & "            // consulta tardara varios minutos. ---" & vbCrLf
    s = s & "            TablaSeg = Table.FromRecords(CombosSegmento)," & vbCrLf
    s = s & "            TablaEmp = Table.FromRecords(CombosEmpresa)," & vbCrLf
    s = s & "            ConEmpCombo = Table.AddColumn(TablaSeg, ""Emp"", each TablaEmp)," & vbCrLf
    s = s & "            CombosExpandido = Table.ExpandTableColumn(ConEmpCombo, ""Emp"", {""Etiqueta"", ""Valores""}, {""EmpresaLabel"", ""EmpresaValores""})," & vbCrLf
    s = s & "            CombosRenombrado = Table.RenameColumns(CombosExpandido, {{""Etiqueta"", ""SegmentoLabel""}, {""Valores"", ""SegmentoValores""}})," & vbCrLf
    s = s & "            ConSubEstadias = Table.AddColumn(CombosRenombrado, ""SubEstadias"", (combo) =>" & vbCrLf
    s = s & "                Table.Buffer(Table.SelectRows(Estadias, (fe) =>" & vbCrLf
    s = s & "                    List.Contains(combo[SegmentoValores], fe[Segmento]) and List.Contains(combo[EmpresaValores], fe[Empresa])" & vbCrLf
    s = s & "                ))" & vbCrLf
    s = s & "            )," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            TablaDias = Table.FromList(Dias, Splitter.SplitByNothing(), {""Fecha""})," & vbCrLf
    s = s & "            ConDias = Table.AddColumn(ConSubEstadias, ""Dias"", each TablaDias)," & vbCrLf
    s = s & "            Grilla = Table.ExpandTableColumn(ConDias, ""Dias"", {""Fecha""})," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            ConResultado = Table.AddColumn(Grilla, ""Resultado"", (filaActual) =>" & vbCrLf
    s = s & "                let" & vbCrLf
    s = s & "                    dayStart = DateTime.From(filaActual[Fecha])," & vbCrLf
    s = s & "                    dayEnd = DateTime.From(Date.AddDays(filaActual[Fecha], 1))," & vbCrLf
    s = s & "                    // filaActual[SubEstadias] ya viene filtrado por Segmento/Empresa" & vbCrLf
    s = s & "                    // y es chico: acá sólo se recorta por el día." & vbCrLf
    s = s & "                    solapan = Table.SelectRows(filaActual[SubEstadias], (fe) =>" & vbCrLf
    s = s & "                        fe[Inicio] < dayEnd and (fe[Fin] = null or fe[Fin] > dayStart)" & vbCrLf
    s = s & "                    )," & vbCrLf
    s = s & "                    presentes = Table.RowCount(Table.Distinct(Table.SelectColumns(solapan, {""IdPersona""})))," & vbCrLf
    s = s & "                    maxSim = fnMaximoSimultaneo(solapan, dayStart, dayEnd)" & vbCrLf
    s = s & "                in" & vbCrLf
    s = s & "                    [Presentes = presentes, Maximo = maxSim[Maximo], HoraMaximo = maxSim[HoraMaximo]]" & vbCrLf
    s = s & "            )," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            Expandido = Table.ExpandRecordColumn(ConResultado, ""Resultado"", {""Presentes"", ""Maximo"", ""HoraMaximo""}, {""Presentes_en_el_dia"", ""Maximo_simultaneo"", ""Hora_del_maximo""})," & vbCrLf
    s = s & "            Renombrado = Table.RenameColumns(Expandido, {{""SegmentoLabel"", ""Segmento""}, {""EmpresaLabel"", ""Empresa""}})," & vbCrLf
    s = s & "            Final = Table.SelectColumns(Renombrado, {""Fecha"", ""Segmento"", ""Empresa"", ""Presentes_en_el_dia"", ""Maximo_simultaneo"", ""Hora_del_maximo""})" & vbCrLf
    s = s & "        in" & vbCrLf
    s = s & "            Final" & vbCrLf
    s = s & "in" & vbCrLf
    s = s & "    Resultado" & vbCrLf
    Codigo_08_POB_Diario = s
End Function

Private Function Codigo_09_POB_Horario() As String
    Dim s As String
    s = ""
    s = s & "// =============================================================================" & vbCrLf
    s = s & "// 09_POB_Horario" & vbCrLf
    s = s & "// -----------------------------------------------------------------------------" & vbCrLf
    s = s & "// Fecha | Hora | Segmento | Empresa | POB. Es la ""foto"" en cada hora de" & vbCrLf
    s = s & "// tHorasFoto. Equivalente a calcular_pob_horario() en referencia.py." & vbCrLf
    s = s & "// =============================================================================" & vbCrLf
    s = s & "let" & vbCrLf
    s = s & "    Estadias = #""06_Estadias""," & vbCrLf
    s = s & "    HorasFotoTexto = Excel.CurrentWorkbook(){[Name=""tHorasFoto""]}[Content][Hora]," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ColumnasSalida = type table [Fecha = date, Hora = text, Segmento = text, Empresa = text, POB = Int64.Type]," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Salida =" & vbCrLf
    s = s & "        // Sin estadías todavía (mes sin datos, o ninguna marca válida): no" & vbCrLf
    s = s & "        // hay rango de fechas que armar, se devuelve la tabla vacía." & vbCrLf
    s = s & "        if Table.IsEmpty(Estadias) then" & vbCrLf
    s = s & "            #table(ColumnasSalida, {})" & vbCrLf
    s = s & "        else" & vbCrLf
    s = s & "        let" & vbCrLf
    s = s & "            FechaMinDatos = List.Min(Estadias[Inicio])," & vbCrLf
    s = s & "            FinesNoNulos = List.RemoveNulls(Estadias[Fin])," & vbCrLf
    s = s & "            TopeFin = if List.IsEmpty(FinesNoNulos) then List.Max(Estadias[Inicio]) else List.Max(FinesNoNulos)," & vbCrLf
    s = s & "            FechaMaxDatos = List.Max({List.Max(Estadias[Inicio]), TopeFin})," & vbCrLf
    s = s & "            DiaInicio = DateTime.Date(FechaMinDatos)," & vbCrLf
    s = s & "            DiaFin = DateTime.Date(FechaMaxDatos)," & vbCrLf
    s = s & "            CantidadDias = Duration.Days(DiaFin - DiaInicio) + 1," & vbCrLf
    s = s & "            Dias = List.Transform({0 .. CantidadDias - 1}, each Date.AddDays(DiaInicio, _))," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            TablaDias = Table.FromList(Dias, Splitter.SplitByNothing(), {""Fecha""})," & vbCrLf
    s = s & "            TablaHoras = Table.FromList(HorasFotoTexto, Splitter.SplitByNothing(), {""Hora""})," & vbCrLf
    s = s & "            ConHoras = Table.AddColumn(TablaDias, ""H"", each TablaHoras)," & vbCrLf
    s = s & "            Grilla = Table.ExpandTableColumn(ConHoras, ""H"", {""Hora""})," & vbCrLf
    s = s & "            ConInstante = Table.AddColumn(Grilla, ""Instante"", (fila) =>" & vbCrLf
    s = s & "                let" & vbCrLf
    s = s & "                    partes = Text.Split(fila[Hora], "":"")," & vbCrLf
    s = s & "                    h = Number.FromText(partes{0})," & vbCrLf
    s = s & "                    m = Number.FromText(partes{1})" & vbCrLf
    s = s & "                in" & vbCrLf
    s = s & "                    fila[Fecha] & #time(h, m, 0)" & vbCrLf
    s = s & "            )," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            // Por cada instante de la grilla, quiénes están adentro (Inicio <= instante < Fin)." & vbCrLf
    s = s & "            ConPresentes = Table.AddColumn(ConInstante, ""Presentes"", (fila) =>" & vbCrLf
    s = s & "                Table.SelectRows(Estadias, (fe) => fe[Inicio] <= fila[Instante] and (fe[Fin] = null or fe[Fin] > fila[Instante]))" & vbCrLf
    s = s & "            )," & vbCrLf
    s = s & "            Expandido = Table.ExpandTableColumn(ConPresentes, ""Presentes"", {""IdPersona"", ""Segmento"", ""Empresa""})," & vbCrLf
    s = s & "            Presencia = Table.SelectColumns(Expandido, {""Fecha"", ""Hora"", ""IdPersona"", ""Segmento"", ""Empresa""})," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            Resultado = #""07_fnExpandirGrupos""(Presencia, {""Fecha"", ""Hora""}, ""POB"")" & vbCrLf
    s = s & "        in" & vbCrLf
    s = s & "            Resultado" & vbCrLf
    s = s & "in" & vbCrLf
    s = s & "    Salida" & vbCrLf
    Codigo_09_POB_Horario = s
End Function

Private Function Codigo_10_POB_Franjas() As String
    Dim s As String
    s = ""
    s = s & "// =============================================================================" & vbCrLf
    s = s & "// 10_POB_Franjas" & vbCrLf
    s = s & "// -----------------------------------------------------------------------------" & vbCrLf
    s = s & "// Fecha | Franja | Segmento | Empresa | POB. Cuenta a quienes están adentro" & vbCrLf
    s = s & "// en la HoraCierre de cada franja (tFranjas), hayan marcado o no dentro de" & vbCrLf
    s = s & "// la franja. Equivalente a calcular_pob_franjas() en referencia.py." & vbCrLf
    s = s & "// =============================================================================" & vbCrLf
    s = s & "let" & vbCrLf
    s = s & "    Estadias = #""06_Estadias""," & vbCrLf
    s = s & "    Franjas = Table.TransformColumns(" & vbCrLf
    s = s & "        Excel.CurrentWorkbook(){[Name=""tFranjas""]}[Content]," & vbCrLf
    s = s & "        {{""Franja"", Text.Trim, type text}, {""HoraInicio"", Text.Trim, type text}, {""HoraCierre"", Text.Trim, type text}}" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ColumnasSalida = type table [Fecha = date, Franja = text, Segmento = text, Empresa = text, POB = Int64.Type]," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Salida =" & vbCrLf
    s = s & "        // Sin estadías todavía (mes sin datos, o ninguna marca válida): no" & vbCrLf
    s = s & "        // hay rango de fechas que armar, se devuelve la tabla vacía." & vbCrLf
    s = s & "        if Table.IsEmpty(Estadias) then" & vbCrLf
    s = s & "            #table(ColumnasSalida, {})" & vbCrLf
    s = s & "        else" & vbCrLf
    s = s & "        let" & vbCrLf
    s = s & "            FechaMinDatos = List.Min(Estadias[Inicio])," & vbCrLf
    s = s & "            FinesNoNulos = List.RemoveNulls(Estadias[Fin])," & vbCrLf
    s = s & "            TopeFin = if List.IsEmpty(FinesNoNulos) then List.Max(Estadias[Inicio]) else List.Max(FinesNoNulos)," & vbCrLf
    s = s & "            FechaMaxDatos = List.Max({List.Max(Estadias[Inicio]), TopeFin})," & vbCrLf
    s = s & "            DiaInicio = DateTime.Date(FechaMinDatos)," & vbCrLf
    s = s & "            DiaFin = DateTime.Date(FechaMaxDatos)," & vbCrLf
    s = s & "            CantidadDias = Duration.Days(DiaFin - DiaInicio) + 1," & vbCrLf
    s = s & "            Dias = List.Transform({0 .. CantidadDias - 1}, each Date.AddDays(DiaInicio, _))," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            TablaDias = Table.FromList(Dias, Splitter.SplitByNothing(), {""Fecha""})," & vbCrLf
    s = s & "            ConFranjas = Table.AddColumn(TablaDias, ""F"", each Franjas)," & vbCrLf
    s = s & "            Grilla = Table.ExpandTableColumn(ConFranjas, ""F"", {""Franja"", ""HoraCierre""})," & vbCrLf
    s = s & "            ConInstante = Table.AddColumn(Grilla, ""Instante"", (fila) =>" & vbCrLf
    s = s & "                let" & vbCrLf
    s = s & "                    partes = Text.Split(fila[HoraCierre], "":"")," & vbCrLf
    s = s & "                    h = Number.FromText(partes{0})," & vbCrLf
    s = s & "                    m = Number.FromText(partes{1})" & vbCrLf
    s = s & "                in" & vbCrLf
    s = s & "                    fila[Fecha] & #time(h, m, 0)" & vbCrLf
    s = s & "            )," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            ConPresentes = Table.AddColumn(ConInstante, ""Presentes"", (fila) =>" & vbCrLf
    s = s & "                Table.SelectRows(Estadias, (fe) => fe[Inicio] <= fila[Instante] and (fe[Fin] = null or fe[Fin] > fila[Instante]))" & vbCrLf
    s = s & "            )," & vbCrLf
    s = s & "            Expandido = Table.ExpandTableColumn(ConPresentes, ""Presentes"", {""IdPersona"", ""Segmento"", ""Empresa""})," & vbCrLf
    s = s & "            Presencia = Table.SelectColumns(Expandido, {""Fecha"", ""Franja"", ""IdPersona"", ""Segmento"", ""Empresa""})," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            Resultado = #""07_fnExpandirGrupos""(Presencia, {""Fecha"", ""Franja""}, ""POB"")" & vbCrLf
    s = s & "        in" & vbCrLf
    s = s & "            Resultado" & vbCrLf
    s = s & "in" & vbCrLf
    s = s & "    Salida" & vbCrLf
    Codigo_10_POB_Franjas = s
End Function

Private Function Codigo_11_Comedor() As String
    Dim s As String
    s = ""
    s = s & "// =============================================================================" & vbCrLf
    s = s & "// 11_Comedor" & vbCrLf
    s = s & "// -----------------------------------------------------------------------------" & vbCrLf
    s = s & "// Fecha | Turno | Segmento | Empresa | Personas_unicas. Cuenta personas" & vbCrLf
    s = s & "// distintas con marca en un lector COMEDOR dentro del turno (tTurnosComedor," & vbCrLf
    s = s & "// extremos inclusivos). Las marcas fuera de turno se ignoran. Equivalente a" & vbCrLf
    s = s & "// calcular_comedor() en referencia.py." & vbCrLf
    s = s & "//" & vbCrLf
    s = s & "// El segmento del comedor sale de la misma columna Segmento de la fichada" & vbCrLf
    s = s & "// (AP/LLY): en la muestra, todos los lectores ""AP - COMEDOR..."" tienen" & vbCrLf
    s = s & "// SEGMENTO = AGUADA PICHANA y todos los ""LLY - COMEDOR..."" tienen SEGMENTO =" & vbCrLf
    s = s & "// LOMA LAS YEGUAS (verificado 1 a 1), así que no hace falta una tabla de" & vbCrLf
    s = s & "// mapeo aparte." & vbCrLf
    s = s & "// =============================================================================" & vbCrLf
    s = s & "let" & vbCrLf
    s = s & "    Origen = #""03_Limpieza""," & vbCrLf
    s = s & "    SoloComedor = Table.SelectRows(Origen, each [TipoLector] = ""COMEDOR"" and [EventoValido] = true)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ConEmpresaJoin = Table.NestedJoin(SoloComedor, {""IdPersona""}, #""04_Empresas"", {""IdPersona""}, ""_emp"", JoinKind.LeftOuter)," & vbCrLf
    s = s & "    ConEmpresa = Table.AddColumn(ConEmpresaJoin, ""EmpresaPersona"", (fila) =>" & vbCrLf
    s = s & "        if fila[_emp] = null or Table.IsEmpty(fila[_emp]) then ""SIN EMPRESA"" else fila[_emp]{0}[Empresa]," & vbCrLf
    s = s & "        type text" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Marcas = Table.SelectColumns(ConEmpresa, {""IdPersona"", ""FechaHora"", ""Segmento"", ""EmpresaPersona""})," & vbCrLf
    s = s & "    MarcasRenombradas = Table.RenameColumns(Marcas, {""EmpresaPersona"", ""Empresa""})," & vbCrLf
    s = s & "    ConFecha = Table.AddColumn(MarcasRenombradas, ""Fecha"", each DateTime.Date([FechaHora]), type date)," & vbCrLf
    s = s & "    ConMinutos = Table.AddColumn(ConFecha, ""MinutosDia"", each Time.Hour(DateTime.Time([FechaHora])) * 60 + Time.Minute(DateTime.Time([FechaHora])), Int64.Type)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Turnos = Table.TransformColumns(" & vbCrLf
    s = s & "        Excel.CurrentWorkbook(){[Name=""tTurnosComedor""]}[Content]," & vbCrLf
    s = s & "        {{""Turno"", Text.Trim, type text}, {""Desde"", Text.Trim, type text}, {""Hasta"", Text.Trim, type text}}" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "    fnMinutos = (txt as text) as number =>" & vbCrLf
    s = s & "        let partes = Text.Split(txt, "":"") in Number.FromText(partes{0}) * 60 + Number.FromText(partes{1})," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ConMarcasPorTurno = Table.AddColumn(Turnos, ""Marcas"", (turno) =>" & vbCrLf
    s = s & "        let" & vbCrLf
    s = s & "            desdeMin = fnMinutos(turno[Desde])," & vbCrLf
    s = s & "            hastaMin = fnMinutos(turno[Hasta])" & vbCrLf
    s = s & "        in" & vbCrLf
    s = s & "            Table.SelectRows(ConMinutos, (m) => m[MinutosDia] >= desdeMin and m[MinutosDia] <= hastaMin)" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "    Expandido = Table.ExpandTableColumn(ConMarcasPorTurno, ""Marcas"", {""Fecha"", ""IdPersona"", ""Segmento"", ""Empresa""})," & vbCrLf
    s = s & "    Presencia = Table.SelectColumns(Expandido, {""Fecha"", ""Turno"", ""IdPersona"", ""Segmento"", ""Empresa""})," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Resultado = #""07_fnExpandirGrupos""(Presencia, {""Fecha"", ""Turno""}, ""Personas_unicas"")" & vbCrLf
    s = s & "in" & vbCrLf
    s = s & "    Resultado" & vbCrLf
    Codigo_11_Comedor = s
End Function

Private Function Codigo_12_ControlCalidad() As String
    Dim s As String
    s = ""
    s = s & "// =============================================================================" & vbCrLf
    s = s & "// 12_ControlCalidad" & vbCrLf
    s = s & "// -----------------------------------------------------------------------------" & vbCrLf
    s = s & "// Resumen de indicadores de calidad de datos. Equivalente a" & vbCrLf
    s = s & "// calcular_control_calidad() en referencia.py — mismos indicadores, mismos" & vbCrLf
    s = s & "// nombres de fila, para poder comparar línea por línea contra" & vbCrLf
    s = s & "// validacion/resultados_esperados.xlsx." & vbCrLf
    s = s & "// =============================================================================" & vbCrLf
    s = s & "let" & vbCrLf
    s = s & "    Origen = #""02_Origen""," & vbCrLf
    s = s & "    Limpio = #""03_Limpieza""," & vbCrLf
    s = s & "    Estadias = #""06_Estadias""," & vbCrLf
    s = s & "    Empresas = #""04_Empresas""," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    TablaLectores = Table.TransformColumns(" & vbCrLf
    s = s & "        Excel.CurrentWorkbook(){[Name=""tLectores""]}[Content]," & vbCrLf
    s = s & "        {{""Lector"", each Text.Combine(List.Select(Text.Split(Text.Trim(_), "" ""), each _ <> """"), "" ""), type text}}" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ArchivosLeidos = List.Count(List.Distinct(Origen[ArchivoOrigen]))," & vbCrLf
    s = s & "    FilasTotales = Table.RowCount(Origen)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // Recalcula el dedup por separado (mismas 10 columnas que 03_Limpieza)" & vbCrLf
    s = s & "    // para no mezclar ""duplicados"" con ""filas sin Fecha/Hora"": 03_Limpieza" & vbCrLf
    s = s & "    // saca ambas cosas, pero acá interesa reportarlas por separado." & vbCrLf
    s = s & "    ColumnasOriginales = {" & vbCrLf
    s = s & "        ""Fecha"", ""Hora"", ""IGG"", ""Apellido"", ""Nombre"", ""Credencial""," & vbCrLf
    s = s & "        ""Evento"", ""SEGMENTO"", ""EMPRESA"", ""Reader Description""" & vbCrLf
    s = s & "    }," & vbCrLf
    s = s & "    Deduplicado = Table.Distinct(Origen, ColumnasOriginales)," & vbCrLf
    s = s & "    DuplicadosEliminados = FilasTotales - Table.RowCount(Deduplicado)," & vbCrLf
    s = s & "    FilasSinFechaHora = Table.RowCount(Table.SelectRows(Deduplicado, each [Fecha] = null or [Hora] = null))," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    LectoresUsados = List.Distinct(Limpio[Lector])," & vbCrLf
    s = s & "    LectoresConocidos = TablaLectores[Lector]," & vbCrLf
    s = s & "    LectoresSinClasificar = List.Sort(List.Difference(LectoresUsados, LectoresConocidos))," & vbCrLf
    s = s & "    TextoLectoresSinClasificar = if List.IsEmpty(LectoresSinClasificar) then ""(ninguno)"" else Text.Combine(LectoresSinClasificar, "", "")," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    MarcasPorCredencial = Table.RowCount(Table.SelectRows(Limpio, each Text.StartsWith([IdPersona], ""CRED_"")))," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ConIgg = Table.SelectRows(Limpio, each not Text.StartsWith([IdPersona], ""CRED_""))," & vbCrLf
    s = s & "    PorPersona = Table.Group(ConIgg, {""IdPersona""}, {{""NCred"", each List.Count(List.Distinct(_[Credencial])), Int64.Type}})," & vbCrLf
    s = s & "    PersonasMultiCredencial = Table.RowCount(Table.SelectRows(PorPersona, each [NCred] > 1))," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SalidasSinIngreso = Table.RowCount(Table.SelectRows(Estadias, each [Motivo] = ""Salida sin ingreso previo""))," & vbCrLf
    s = s & "    DescartesFinDeMes = Table.RowCount(Table.SelectRows(Estadias, each [Motivo] = ""Descarte fin de mes""))," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PersonasSinEmpresa = Table.RowCount(Table.SelectRows(Empresas, each [Empresa] = ""SIN EMPRESA""))," & vbCrLf
    s = s & "    PersonasADefinir = Table.RowCount(Table.SelectRows(Empresas, each [Empresa] = ""--- A DEFINIR ---""))," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Descartados = Table.SelectRows(Limpio, each [EventoValido] = false)," & vbCrLf
    s = s & "    PorEvento = Table.Group(Descartados, {""Evento""}, {{""Cantidad"", each Table.RowCount(_), Int64.Type}})," & vbCrLf
    s = s & "    FilasEvento = Table.AddColumn(PorEvento, ""Indicador"", each ""Marcas descartadas por evento: "" & [Evento])," & vbCrLf
    s = s & "    FilasEventoFinal = Table.RenameColumns(Table.SelectColumns(FilasEvento, {""Indicador"", ""Cantidad""}), {""Cantidad"", ""Valor""})," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Indicadores = #table(" & vbCrLf
    s = s & "        {""Indicador"", ""Valor""}," & vbCrLf
    s = s & "        {" & vbCrLf
    s = s & "            {""Archivos leídos"", ArchivosLeidos}," & vbCrLf
    s = s & "            {""Filas totales (antes de deduplicar)"", FilasTotales}," & vbCrLf
    s = s & "            {""Duplicados eliminados"", DuplicadosEliminados}," & vbCrLf
    s = s & "            {""Filas sin Fecha/Hora descartadas"", FilasSinFechaHora}," & vbCrLf
    s = s & "            {""Lectores sin clasificar"", TextoLectoresSinClasificar}," & vbCrLf
    s = s & "            {""Marcas sin IGG identificadas por credencial"", MarcasPorCredencial}," & vbCrLf
    s = s & "            {""Personas con más de una credencial"", PersonasMultiCredencial}," & vbCrLf
    s = s & "            {""Salidas sin ingreso previo"", SalidasSinIngreso}," & vbCrLf
    s = s & "            {""Descartes de fin de mes"", DescartesFinDeMes}," & vbCrLf
    s = s & "            {""Personas SIN EMPRESA"", PersonasSinEmpresa}," & vbCrLf
    s = s & "            {""Personas --- A DEFINIR ---"", PersonasADefinir}" & vbCrLf
    s = s & "        }" & vbCrLf
    s = s & "    )," & vbCrLf
    s = s & "    Final = Table.Combine({Indicadores, FilasEventoFinal})" & vbCrLf
    s = s & "in" & vbCrLf
    s = s & "    Final" & vbCrLf
    Codigo_12_ControlCalidad = s
End Function
