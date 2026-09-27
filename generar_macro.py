# -*- coding: utf-8 -*-
"""
Genera vba/Instalar_Reportes.bas a partir de los archivos pq/*.pq: embebe el
código M de cada consulta como una función VBA que arma el string línea por
línea (evita problemas de comillas/longitud de línea), y arma el resto del
módulo (creación idempotente de consultas, carga como tabla, y un dashboard
básico de tablas dinámicas + segmentaciones + gráficos).

Se corre una sola vez para generar el .bas; si se edita algún .pq, hay que
volver a correr este script para que la macro quede sincronizada.
"""
import re
from pathlib import Path

PQ_DIR = Path(__file__).parent / "pq"

# (nombre_consulta, hoja_destino_o_None)
CONSULTAS = [
    ("01_Parametros", None),
    ("02_Origen", None),
    ("03_Limpieza", None),
    ("04_Empresas", None),
    ("05_Validos", None),
    ("06_Estadias", "Estadias"),
    ("07_fnExpandirGrupos", None),
    ("08_POB_Diario", "POB_Diario"),
    ("09_POB_Horario", "POB_Horario"),
    ("10_POB_Franjas", "POB_Franjas"),
    ("11_Comedor", "Comedor"),
    ("12_ControlCalidad", "Control_Calidad"),
]


def nombre_funcion(nombre_consulta: str) -> str:
    return "Codigo_" + re.sub(r"[^0-9A-Za-z_]", "_", nombre_consulta)


def vba_escape(linea: str) -> str:
    return linea.replace('"', '""')


def generar_funcion_codigo(nombre_consulta: str) -> str:
    ruta = PQ_DIR / f"{nombre_consulta}.pq"
    texto = ruta.read_text(encoding="utf-8")
    lineas = texto.splitlines()
    fn = nombre_funcion(nombre_consulta)
    cuerpo = [f"Private Function {fn}() As String", "    Dim s As String", '    s = ""']
    for linea in lineas:
        cuerpo.append(f'    s = s & "{vba_escape(linea)}" & vbCrLf')
    cuerpo.append(f"    {fn} = s")
    cuerpo.append("End Function")
    return "\n".join(cuerpo)


def generar_modulo():
    partes_codigo = [generar_funcion_codigo(nombre) for nombre, _ in CONSULTAS]

    lineas_instalar = []
    for nombre, _hoja in CONSULTAS:
        lineas_instalar.append(
            f'    CrearOActualizarConsulta wb, "{nombre}", {nombre_funcion(nombre)}()'
        )

    lineas_cargar = []
    for nombre, hoja in CONSULTAS:
        if hoja:
            lineas_cargar.append(f'    CargarComoTabla wb, "{nombre}", "{hoja}"')

    modulo = f'''Attribute VB_Name = "Instalar_Reportes"
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
{chr(10).join(lineas_instalar)}

    ' --- 2) Cargar como tabla las que van a una hoja ---
{chr(10).join(lineas_cargar)}

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

{chr(10).join(f'{chr(10)}{codigo}' for codigo in partes_codigo)}
'''
    return modulo


def main():
    salida = Path(__file__).parent / "vba" / "Instalar_Reportes.bas"
    salida.write_text(generar_modulo(), encoding="utf-8")
    print(f"Generado: {salida}")


if __name__ == "__main__":
    main()
