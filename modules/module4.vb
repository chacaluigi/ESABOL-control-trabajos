Option Explicit

' Ajusta estos nombres si tus hojas/tablas se llaman distinto
Private Const SHEET_CONTROL As String = "tabla_control"    ' hoja donde está C5/A6 y B8 vista
Private Const SHEET_TAREAS As String = "tareas"           ' hoja que contiene la tabla "tareas"
Private Const TABLE_TAREAS_NAME As String = "tareas"      ' nombre del ListObject con datos origen
Private Const SHEET_PT As String = "personal_tareas"      ' hoja con la tabla puente
Private Const TABLE_PT_NAME As String = "personal_tareas" ' nombre del ListObject puente

' Lugar donde pegar la vista (primera celda de la primera fila de datos filtrados)
Private Const DEST_FIRST_CELL As String = "B8"

' Máximo número de filas a limpiar al final (por seguridad)
Private Const MAX_CLEAN_ROWS As Long = 1000

Public Sub RefreshTablaControl()
    Dim wb As Workbook
    Dim wsControl As Worksheet, wsT As Worksheet, wsPT As Worksheet, wsPersonal As Worksheet
    Dim tblT As ListObject, tblPT As ListObject, tblP As ListObject
    Dim idPersona As Variant
    Dim dict As Object
    Dim i As Long, j As Long
    Dim idVal As Variant
    Dim srcRow As Range, destCell As Range
    Dim destStartRow As Long, destStartCol As Long
    Dim colCount As Long
    Dim copied As Long

    On Error GoTo ErrHandler
    Set wb = ThisWorkbook
    Set wsControl = wb.Worksheets(SHEET_CONTROL)
    Set wsT = wb.Worksheets(SHEET_TAREAS)
    Set wsPT = wb.Worksheets(SHEET_PT)
    Set tblT = wsT.ListObjects(TABLE_TAREAS_NAME)
    Set tblPT = wsPT.ListObjects(TABLE_PT_NAME)

    ' --- Leer el nombre desde C5 (NO usar A6)
    Dim personaName As String
    personaName = Trim(CStr(wsControl.Range("C5").Value))
    If personaName = "" Then
        Call LimpiarAreaDestino(wsControl, DEST_FIRST_CELL, tblT.ListColumns.Count)
        MsgBox "No hay persona seleccionada en C5.", vbInformation
        Exit Sub
    End If

    ' --- Buscar el nro (id) de la persona en la tabla "personal"
    Set wsPersonal = wb.Worksheets("personal")
    Set tblP = wsPersonal.ListObjects("personal")

    Dim foundP As Range
    Dim idxP_name As Long, idxP_nro As Long
    idxP_name = tblP.ListColumns("Apellidos y Nombres").Index
    ' Columna que contiene el identificador de persona; AJUSTA si tu encabezado es distinto (ej. "Nro")
    idxP_nro = tblP.ListColumns("persona_id").Index

    Set foundP = tblP.ListColumns("Apellidos y Nombres").DataBodyRange.Find(What:=personaName, LookAt:=xlWhole, MatchCase:=False)
    If foundP Is Nothing Then
        Call LimpiarAreaDestino(wsControl, DEST_FIRST_CELL, tblT.ListColumns.Count)
        MsgBox "La persona '" & personaName & "' no se encontró en la tabla personal.", vbExclamation
        Exit Sub
    End If

    ' Obtener el id (nro) desde la fila encontrada
    idPersona = foundP.Offset(0, idxP_nro - idxP_name).Value

    ' Si idPersona está vacío -> salir
    If Trim(CStr(idPersona)) = "" Then
        Call LimpiarAreaDestino(wsControl, DEST_FIRST_CELL, tblT.ListColumns.Count)
        MsgBox "No se pudo determinar el identificador de la persona.", vbExclamation
        Exit Sub
    End If

    ' --- Construir diccionario de tareas asociadas desde la tabla puente personal_tareas
    Set dict = CreateObject("Scripting.Dictionary")
    dict.RemoveAll

    Dim idxPT_nro_persona As Long, idxPT_nro_tarea As Long
    idxPT_nro_persona = tblPT.ListColumns("nro_persona").Index
    idxPT_nro_tarea = tblPT.ListColumns("nro_tarea").Index

    For i = 1 To tblPT.DataBodyRange.Rows.Count
        If CStr(tblPT.DataBodyRange.Cells(i, idxPT_nro_persona).Value) = CStr(idPersona) Then
            idVal = tblPT.DataBodyRange.Cells(i, idxPT_nro_tarea).Value
            If Not dict.Exists(CStr(idVal)) Then dict.Add CStr(idVal), True
        End If
    Next i

    ' Si no hay tareas, limpiar y salir
    If dict.Count = 0 Then
        Call LimpiarAreaDestino(wsControl, DEST_FIRST_CELL, tblT.ListColumns.Count)
        MsgBox "No se encontraron tareas para la persona seleccionada.", vbInformation
        Exit Sub
    End If

    ' Preparar destino y limpiar área
    destStartRow = wsControl.Range(DEST_FIRST_CELL).Row
    destStartCol = wsControl.Range(DEST_FIRST_CELL).Column
    colCount = tblT.ListColumns.Count
    Call LimpiarAreaDestino(wsControl, DEST_FIRST_CELL, colCount)

    ' Copiar filas desde la tabla "tareas" que estén en dict (manteniendo colores y formatos)
    copied = 0
    Dim idxT_tareaID As Long
    idxT_tareaID = tblT.ListColumns("tarea_id").Index

    For i = 1 To tblT.DataBodyRange.Rows.Count
        idVal = tblT.DataBodyRange.Cells(i, idxT_tareaID).Value
        If dict.Exists(CStr(idVal)) Then
            Set srcRow = tblT.DataBodyRange.Rows(i)
            Set destCell = wsControl.Cells(destStartRow + copied, destStartCol)
            srcRow.Copy Destination:=destCell
            copied = copied + 1
        End If
    Next i

    ' Ajustar alturas (opcional)
    For j = 1 To copied
        wsControl.Rows(destStartRow + j - 1).RowHeight = tblT.DataBodyRange.Rows(j).RowHeight
    Next j

    ' Limpiar resto debajo para evitar residuos
    If copied < MAX_CLEAN_ROWS Then
        Dim startClear As Long
        startClear = destStartRow + copied
        wsControl.Range(wsControl.Cells(startClear, destStartCol), _
                        wsControl.Cells(startClear + MAX_CLEAN_ROWS - copied, destStartCol + colCount - 1)).Clear
    End If

    MsgBox "Vista actualizada: " & copied & " tarea(s) copiadas.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error en RefreshTablaControl: " & Err.Description, vbExclamation
End Sub


' Limpia el área destino (valores, formatos e interior) antes de pegar
Private Sub LimpiarAreaDestino(ws As Worksheet, firstCellAddress As String, colCount As Long)
    Dim startRow As Long, startCol As Long
    startRow = ws.Range(firstCellAddress).Row
    startCol = ws.Range(firstCellAddress).Column
    ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + MAX_CLEAN_ROWS, startCol + colCount - 1)).Clear
End Sub


