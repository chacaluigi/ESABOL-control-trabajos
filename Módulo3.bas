Attribute VB_Name = "M�dulo3"
Option Explicit

Public Const YEAR_REF As Long = 2026
Public Const COL_DIA_INICIO As Long = 7   ' G = d�a 1 en la hoja de control (solo referencia)
Public Const COL_DIA_FIN As Long = 37     ' AK = d�a 31
Public Const SHEET_CONTROL As String = "tabla_control"
Public Const SHEET_TAREAS As String = "tareas"
Public Const TABLE_TAREAS_NAME As String = "tareas"
' --- Hoja / tabla puente (personal_tareas)
Public Const SHEET_PT As String = "personal_tareas"      ' nombre de la hoja donde est� la tabla puente
Public Const TABLE_PT_NAME As String = "personal_tareas" ' nombre del ListObject de la tabla puente


' --- Actualizar FECHA INICIO / FECHA FINAL / PORCENTAJE en la tabla "tareas"
Public Sub ActualizarTareaOrigen(tareaId As Long, fechaIni As Variant, fechaFin As Variant, sumaPorc As Double)
    Dim wsT As Worksheet
    Dim tbl As ListObject
    Dim rngId As Range
    Dim filaTbl As Range
    Dim idxFechaIni As Long, idxFechaFin As Long, idxPorc As Long

    On Error GoTo ErrHandler
    Set wsT = ThisWorkbook.Worksheets(SHEET_TAREAS)
    Set tbl = wsT.ListObjects(TABLE_TAREAS_NAME)

    Set rngId = tbl.ListColumns("tarea_id").DataBodyRange
    Set filaTbl = rngId.Find(What:=tareaId, LookAt:=xlWhole)

    If filaTbl Is Nothing Then Exit Sub

    idxFechaIni = tbl.ListColumns("FECHA INICIO").Index
    idxFechaFin = tbl.ListColumns("FECHA FINAL").Index
    idxPorc = tbl.ListColumns("PORCENTAJE").Index

    filaTbl.Offset(0, idxFechaIni - 1).Value = IIf(IsEmpty(fechaIni), vbNullString, fechaIni)
    filaTbl.Offset(0, idxFechaFin - 1).Value = IIf(IsEmpty(fechaFin), vbNullString, fechaFin)
    filaTbl.Offset(0, idxPorc - 1).Value = Application.WorksheetFunction.Min(sumaPorc / 100#, 1#)

    Exit Sub
ErrHandler:
    MsgBox "Error en ActualizarTareaOrigen: " & Err.Description, vbExclamation
End Sub

' --- Actualiza la columna del d�a (1..31) en la tabla "tareas"
Public Sub ActualizarDiaEnTablaOrigen(tareaId As Long, dia As Long, valor As Variant)
    Dim wsT As Worksheet
    Dim tbl As ListObject
    Dim rngId As Range
    Dim filaTbl As Range
    Dim colName As String
    Dim idxCol As Long

    If dia < 1 Or dia > 31 Then Exit Sub

    On Error GoTo ErrHandler
    Set wsT = ThisWorkbook.Worksheets(SHEET_TAREAS)
    Set tbl = wsT.ListObjects(TABLE_TAREAS_NAME)

    Set rngId = tbl.ListColumns("tarea_id").DataBodyRange
    Set filaTbl = rngId.Find(What:=tareaId, LookAt:=xlWhole)

    If filaTbl Is Nothing Then Exit Sub

    colName = CStr(dia) ' columnas "1","2",...
    idxCol = tbl.ListColumns(colName).Index

    With filaTbl.Offset(0, idxCol - 1)
        If IsMissing(valor) Or IsEmpty(valor) Or Trim(CStr(valor)) = "" Then
            .ClearContents
        Else
            .Value = CDbl(valor)
        End If
    End With

    Exit Sub
ErrHandler:
    MsgBox "Error en ActualizarDiaEnTablaOrigen: " & Err.Description, vbExclamation
End Sub


' --- Recalcula FECHA INICIO / FECHA FINAL / PORCENTAJE leyendo directamente la fila de la tabla "tareas"
' devuelve por referencia fechaIni, fechaFin, sumaPorc (suma en porcentaje 0..100)
Public Sub RecalcularTareaEnTabla( _
    tareaId As Long, _
    ByRef fechaIni As Variant, _
    ByRef fechaFin As Variant, _
    ByRef sumaPorc As Double)

    Dim wsT As Worksheet
    Dim tbl As ListObject
    Dim filaTbl As Range
    Dim rowIndex As Long
    Dim dia As Long
    Dim idxCol As Long
    Dim val As Variant
    Dim ultimaFecha As Variant

    fechaIni = Empty
    fechaFin = Empty
    sumaPorc = 0
    ultimaFecha = Empty

    Set wsT = ThisWorkbook.Worksheets(SHEET_TAREAS)
    Set tbl = wsT.ListObjects(TABLE_TAREAS_NAME)

    Set filaTbl = tbl.ListColumns("tarea_id").DataBodyRange.Find( _
                    What:=tareaId, LookAt:=xlWhole)

    If filaTbl Is Nothing Then Exit Sub

    rowIndex = filaTbl.Row - tbl.DataBodyRange.Row + 1

    For dia = 1 To 31
        idxCol = tbl.ListColumns(CStr(dia)).Index
        val = tbl.DataBodyRange.Cells(rowIndex, idxCol).Value

        If IsNumeric(val) And val > 0 Then
            If IsEmpty(fechaIni) Then
                fechaIni = DateSerial(YEAR_REF, 1, dia)
            End If

            ultimaFecha = DateSerial(YEAR_REF, 1, dia)
            sumaPorc = sumaPorc + CDbl(val)
        End If
    Next dia

    If sumaPorc >= 100 Then
        sumaPorc = 100
        fechaFin = ultimaFecha
    Else
        fechaFin = Empty
    End If
End Sub


' Recalcula fecha inicio, fecha final y porcentaje (en la hoja de control) para una fila dada
' y retorna los valores por referencia
Public Sub RecalcularFilaControl( _
    ws As Worksheet, _
    fila As Long, _
    ByRef fechaIni As Variant, _
    ByRef fechaFin As Variant, _
    ByRef sumaPorc As Double)

    Dim rngDias As Range, c As Range
    Dim colInicio As Long
    Dim ultimaFecha As Variant

    colInicio = COL_DIA_INICIO

    Set rngDias = ws.Cells(fila, colInicio).Resize(1, 31)

    fechaIni = Empty
    fechaFin = Empty
    sumaPorc = 0
    ultimaFecha = Empty

    For Each c In rngDias
        If IsNumeric(c.Value) And c.Value > 0 Then
            If IsEmpty(fechaIni) Then
                fechaIni = DateSerial(YEAR_REF, 1, c.Column - colInicio + 1)
            End If

            ultimaFecha = DateSerial(YEAR_REF, 1, c.Column - colInicio + 1)
            sumaPorc = sumaPorc + CDbl(c.Value)
        End If
    Next c

    If sumaPorc >= 100 Then
        sumaPorc = 100
        fechaFin = ultimaFecha
    Else
        fechaFin = Empty
    End If
End Sub


' ------------------------
' Color utilities for tasks
' ------------------------
Public Function ColorFromName(colorName As String) As Long
    Select Case LCase(Trim(colorName))
        Case "amarillo"
            ColorFromName = RGB(255, 255, 0)        ' D�as de trabajo
        Case "rojo"
            ColorFromName = RGB(255, 0, 0)          ' Guardia entrante
        Case "naranja"
            ColorFromName = RGB(255, 192, 0)        ' Guardia saliente
        Case "celeste"
            ColorFromName = RGB(0, 176, 240)      ' Vacaci�n (light blue)
        Case "verde oscuro"
            ColorFromName = RGB(196, 215, 155)          ' Comisi�n Vuelo (dark green)
        Case "gris"
            ColorFromName = RGB(221, 297, 196)      ' Comisi�n varios (grey)
        Case "verde claro"
            ColorFromName = RGB(0, 255, 0)      ' D�a de permiso (light green)
        Case "caf�", "cafe", "caf� "
            ColorFromName = RGB(151, 71, 6)        ' Otros (brown)
        Case Else
            ColorFromName = xlNone                  ' sin color por defecto
    End Select
End Function

' Aplica color de fondo a la celda del d�a correspondiente en la tabla "tareas"
Public Sub AplicarColorDiaEnTablaOrigen(tareaId As Long, dia As Long, colorLong As Long)
    Dim wsT As Worksheet
    Dim tbl As ListObject
    Dim rngId As Range
    Dim filaTbl As Range
    Dim idxCol As Long
    Dim cel As Range

    On Error GoTo ErrHandler
    If dia < 1 Or dia > 31 Then Exit Sub

    Set wsT = ThisWorkbook.Worksheets(SHEET_TAREAS)
    Set tbl = wsT.ListObjects(TABLE_TAREAS_NAME)

    Set rngId = tbl.ListColumns("tarea_id").DataBodyRange
    Set filaTbl = rngId.Find(What:=tareaId, LookAt:=xlWhole)

    If filaTbl Is Nothing Then Exit Sub

    idxCol = tbl.ListColumns(CStr(dia)).Index
    Set cel = filaTbl.Offset(0, idxCol - 1)

    If colorLong = xlNone Then
        ' opcional: quitar color
        cel.Interior.Pattern = xlNone
    Else
        cel.Interior.Pattern = xlSolid
        cel.Interior.Color = colorLong
    End If

    Exit Sub
ErrHandler:
    ' Silencioso o mostrar mensaje si lo deseas:
    ' MsgBox "Error AplicarColorDiaEnTablaOrigen: " & Err.Description, vbExclamation
End Sub


' Devuelve persona_id por nombre exacto (o 0 si no encuentra)
Public Function GetPersonaIDByName(personName As String) As Long
    Dim wsP As Worksheet, tblP As ListObject, foundP As Range
    On Error GoTo ErrHandler
    Set wsP = ThisWorkbook.Worksheets("personal")
    Set tblP = wsP.ListObjects("personal")
    Set foundP = tblP.ListColumns("Apellidos y Nombres").DataBodyRange.Find(What:=personName, LookAt:=xlWhole, MatchCase:=False)
    If Not foundP Is Nothing Then
        GetPersonaIDByName = CLng(tblP.DataBodyRange.Cells(foundP.Row - tblP.DataBodyRange.Row + 1, tblP.ListColumns("persona_id").Index).Value)
    Else
        GetPersonaIDByName = 0
    End If
    Exit Function
ErrHandler:
    GetPersonaIDByName = 0
End Function

