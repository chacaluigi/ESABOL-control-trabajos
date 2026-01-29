Option Explicit

Private Sub UserForm_Initialize()
    ' Inicializaciones
    gSelectedPersonID = 0
    gSelectedPersonName = ""
    gBuscarPersona_WriteToC5 = False ' por defecto al abrir aquí, no sobrescribimos C5
    Me.lstAssigned.Clear
End Sub

Private Sub btnBuscarPersona_Click()
    ' Abrir el buscador en modo "picker" (no escribe en C5)
    gBuscarPersona_WriteToC5 = False
    gSelectedPersonID = 0
    gSelectedPersonName = ""
    frmBuscarPersona.Show vbModal

    ' Al volver, gSelectedPersonID/gSelectedPersonName pueden tener datos
    If Len(Trim(gSelectedPersonName)) > 0 Then
        ' Añadir a la lista si no existe
        Dim i As Long, exists As Boolean
        exists = False
        For i = 0 To Me.lstAssigned.ListCount - 1
            If Me.lstAssigned.List(i, 0) = gSelectedPersonName Then
                exists = True
                Exit For
            End If
        Next i

        If Not exists Then
            ' Guardar como "persona_id | nombre" usando el .ItemData para el id no es práctico en VBA ListBox,
            ' así que almacenamos el texto "id - nombre" para referencia, pero mostramos solo el nombre.
            Dim displayText As String
            If gSelectedPersonID > 0 Then
                displayText = CStr(gSelectedPersonID) & " - " & gSelectedPersonName
            Else
                displayText = "0 - " & gSelectedPersonName
            End If
            Me.lstAssigned.AddItem displayText
        Else
            MsgBox "La persona ya está en la lista.", vbInformation
        End If
    End If
End Sub

Private Sub btnAgregarPersona_Click()
    ' Mismo comportamiento que btnBuscarPersona para compatibilidad
    Call btnBuscarPersona_Click
End Sub

Private Sub btnEliminarPersona_Click()
    Dim i As Long
    ' eliminar los seleccionados
    For i = Me.lstAssigned.ListCount - 1 To 0 Step -1
        If Me.lstAssigned.Selected(i) Then Me.lstAssigned.RemoveItem i
    Next i
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnCrearTarea_Click()
    Dim taskName As String
    taskName = Trim(Me.txtNewTaskName.Value)
    If taskName = "" Then
        MsgBox "Ingrese el nombre de la tarea.", vbExclamation
        Exit Sub
    End If

    ' Referencias a tablas
    Dim wsT As Worksheet, wsPT As Worksheet
    Dim tblT As ListObject, tblPT As ListObject
    Set wsT = ThisWorkbook.Worksheets(SHEET_TAREAS)
    Set wsPT = ThisWorkbook.Worksheets(SHEET_PT)
    Set tblT = wsT.ListObjects(TABLE_TAREAS_NAME)
    Set tblPT = wsPT.ListObjects(TABLE_PT_NAME)

    ' Calcular nuevo tarea_id (mayor + 1) - si tabla vacía, empezar en 1
    Dim newID As Long
    On Error Resume Next
    newID = Application.WorksheetFunction.Max(tblT.ListColumns("tarea_id").DataBodyRange) + 1
    If Err.Number <> 0 Then
        Err.Clear
        newID = 1
    End If
    On Error GoTo 0

    ' Agregar nueva fila a la tabla tareas
    Dim newRow As ListRow
    Set newRow = tblT.ListRows.Add
    
    ' asignar valores por columnas (usar índices dinámicos)
    Dim idx_id As Long, idx_tarea As Long
    idx_id = tblT.ListColumns("tarea_id").Index
    idx_tarea = tblT.ListColumns("TAREA").Index

    newRow.Range.Cells(1, idx_id).Value = newID
    newRow.Range.Cells(1, idx_tarea).Value = taskName
    ' resto columnas (FECHA INICIO, FECHA FINAL, PORCENTAJE, 1..31) quedan vacías por diseño
    
    ' -------------------------
    ' Quitar solo el relleno (Interior) heredado de la fila superior
    ' -------------------------
    Dim cell As Range
    For Each cell In newRow.Range.Cells
        cell.Interior.Pattern = xlNone
    Next cell
    
    ' -------------------------
    ' (Opcional/seguro) Asegurar que las columnas de días 1..31 estén vacías y sin relleno
    ' -------------------------
    Dim dia As Long, idxDia As Long
    For dia = 1 To 31
        On Error Resume Next
        idxDia = tblT.ListColumns(CStr(dia)).Index
        On Error GoTo 0
        If idxDia > 0 Then
            With newRow.Range.Cells(1, idxDia)
                .ClearContents
                .Interior.Pattern = xlNone
            End With
        End If
    Next dia
    
    
    ' Ahora agregar entradas en la tabla puente personal_tareas para cada persona de la lista
    Dim i As Long, personaID As Long, itemText As String
    For i = 0 To Me.lstAssigned.ListCount - 1
        itemText = Me.lstAssigned.List(i)
        ' esperar formato "id - nombre" (si no se encontró id, id=0)
        personaID = CLng(Split(itemText, " - ")(0))
        If personaID = 0 Then
            ' intentar resolver por nombre
            Dim pName As String
            pName = Trim(Mid(itemText, InStr(itemText, " - ") + 3))
            personaID = GetPersonaIDByName(pName) ' función utilitaria (ver más abajo)
        End If

        If personaID > 0 Then
            ' evitar duplicados: verificar si ya existe en tblPT
            Dim found As Range
            Set found = tblPT.DataBodyRange.Columns(tblPT.ListColumns("nro_persona").Index).Find(What:=personaID, LookAt:=xlWhole)
            If Not found Is Nothing Then
                ' hay fila(s) para persona; comprobar si para el mismo nro_tarea ya existe
                Dim r As Range, existsPT As Boolean
                existsPT = False
                For Each r In tblPT.DataBodyRange.Rows
                    If CLng(r.Cells(1, tblPT.ListColumns("nro_persona").Index).Value) = personaID Then
                        If CLng(r.Cells(1, tblPT.ListColumns("nro_tarea").Index).Value) = newID Then
                            existsPT = True
                            Exit For
                        End If
                    End If
                Next r
                If Not existsPT Then
                    Dim newRowPT As ListRow
                    Set newRowPT = tblPT.ListRows.Add
                    newRowPT.Range.Cells(1, tblPT.ListColumns("nro_persona").Index).Value = personaID
                    newRowPT.Range.Cells(1, tblPT.ListColumns("nro_tarea").Index).Value = newID
                End If
            Else
                ' no hay filas para esa persona aún: añadir directamente
                Dim newRowPT2 As ListRow
                Set newRowPT2 = tblPT.ListRows.Add
                newRowPT2.Range.Cells(1, tblPT.ListColumns("nro_persona").Index).Value = personaID
                newRowPT2.Range.Cells(1, tblPT.ListColumns("nro_tarea").Index).Value = newID
            End If
        Else
            ' Si no se pudo resolver personaID, avisar (pero continuar)
            MsgBox "No se encontró el ID de la persona: " & itemText, vbExclamation
        End If
    Next i

    ' Refrescar vista general (si quieres)
    On Error Resume Next
    Application.ScreenUpdating = False
    RefreshTablaControl
    Application.ScreenUpdating = True
    On Error GoTo 0

    MsgBox "Tarea '" & taskName & "' creada con ID " & newID & " y " & Me.lstAssigned.ListCount & " persona(s) asignada(s).", vbInformation
    Unload Me
End Sub

