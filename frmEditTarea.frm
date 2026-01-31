VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditTarea 
   Caption         =   "EDITAR TAREA"
   ClientHeight    =   5076
   ClientLeft      =   0
   ClientTop       =   12
   ClientWidth     =   4416
   OleObjectBlob   =   "frmEditTarea.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmEditTarea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fraEdit_Click()

End Sub


' Formulario: frmEditTarea
' Frame fraCargar -> pedir ID
' Frame fraEdit -> edición (igual que frmAddTarea layout)

Private Sub UserForm_Initialize()
    ' Mostrar solo la sección de carga
    Me.fraCargar.Visible = True
    Me.fraEdit.Visible = False

    ' Inicializar lista
    Me.lstAssigned.Clear
    gBuscarPersona_WriteToC5 = False ' cuando se llame desde aquí, no escribir en C5
    gSelectedPersonID = 0
    gSelectedPersonName = ""
End Sub

' --- Botón cerrar en carga
Private Sub btnCancelarCarga_Click()
    Unload Me
End Sub

' --- Botón Cargar: trae datos de la tarea y poblamos la UI de edición
Private Sub btnCargarTarea_Click()
    Dim tareaID As Long
    If Trim(Me.txtEditTaskID.Value) = "" Then
        MsgBox "Ingrese un ID de tarea.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(Me.txtEditTaskID.Value) Then
        MsgBox "ID inválido.", vbExclamation
        Exit Sub
    End If

    tareaID = CLng(Me.txtEditTaskID.Value)

    ' Buscar tarea en la tabla tareas
    Dim wsT As Worksheet, tblT As ListObject, found As Range
    Set wsT = ThisWorkbook.Worksheets(SHEET_TAREAS)
    Set tblT = wsT.ListObjects(TABLE_TAREAS_NAME)

    On Error Resume Next
    Set found = tblT.ListColumns("tarea_id").DataBodyRange.Find(What:=tareaID, LookAt:=xlWhole)
    On Error GoTo 0

    If found Is Nothing Then
        MsgBox "No se encontró tarea con ID " & tareaID, vbExclamation
        Exit Sub
    End If

    ' Mostrar frame de edición y rellenar campos
    Me.fraCargar.Visible = False
    Me.fraEdit.Visible = True

    ' Rellenar nombre de tarea
    Dim rowIndex As Long
    rowIndex = found.Row - tblT.DataBodyRange.Row + 1
    Me.txtNewTaskName.Value = tblT.DataBodyRange.Cells(rowIndex, tblT.ListColumns("TAREA").Index).Value

    ' Limpiar lista asignados e insertar las personas asignadas actualmente
    Me.lstAssigned.Clear
    Call LoadAssignedPersonsIntoList(tareaID)

    ' Guardamos el id en un Tag o en txtEditTaskID para referencia
    Me.txtEditTaskID.Tag = CStr(tareaID)
End Sub

' Cargar personas asignadas en la lista desde tabla puente
Private Sub LoadAssignedPersonsIntoList(tareaID As Long)
    Dim wsPT As Worksheet, tblPT As ListObject
    Dim wsP As Worksheet, tblP As ListObject
    Dim i As Long, pid As Long, pName As String
    Set wsPT = ThisWorkbook.Worksheets(SHEET_PT)
    Set tblPT = wsPT.ListObjects(TABLE_PT_NAME)
    Set wsP = ThisWorkbook.Worksheets("personal")
    Set tblP = wsP.ListObjects("personal")

    ' Recorremos tabla puente y añadimos las personas que tengan nro_tarea = tareaID
    Dim nRows As Long
    nRows = tblPT.ListRows.Count
    If nRows = 0 Then Exit Sub

    For i = 1 To tblPT.DataBodyRange.Rows.Count
        If CLng(tblPT.DataBodyRange.Cells(i, tblPT.ListColumns("nro_tarea").Index).Value) = tareaID Then
            pid = CLng(tblPT.DataBodyRange.Cells(i, tblPT.ListColumns("nro_persona").Index).Value)
            pName = GetPersonaNameByID(pid) ' utilidad en module3
            If Len(Trim(pName)) = 0 Then pName = "ID:" & pid
            ' Añadir "id - nombre"
            Me.lstAssigned.AddItem CStr(pid) & " - " & pName
        End If
    Next i
End Sub

' Botón buscar persona (reusa frmBuscarPersona)
Private Sub btnBuscarPersona_Click()
    gBuscarPersona_WriteToC5 = False
    gSelectedPersonID = 0
    gSelectedPersonName = ""
    frmBuscarPersona.Show vbModal

    If Len(Trim(gSelectedPersonName)) > 0 Then
        Dim displayText As String
        If gSelectedPersonID > 0 Then
            displayText = CStr(gSelectedPersonID) & " - " & gSelectedPersonName
        Else
            displayText = "0 - " & gSelectedPersonName
        End If

        ' evitar duplicados
        Dim i As Long, exists As Boolean
        exists = False
        For i = 0 To Me.lstAssigned.ListCount - 1
            If Me.lstAssigned.List(i) = displayText Then
                exists = True
                Exit For
            End If
        Next i
        If Not exists Then Me.lstAssigned.AddItem displayText
    End If
End Sub

Private Sub btnAgregarPersona_Click()
    Call btnBuscarPersona_Click
End Sub

Private Sub btnEliminarPersona_Click()
    Dim i As Long
    For i = Me.lstAssigned.ListCount - 1 To 0 Step -1
        If Me.lstAssigned.Selected(i) Then Me.lstAssigned.RemoveItem i
    Next i
End Sub

' Guardar cambios: actualizar nombre en tabla tareas y reemplazar relaciones en personal_tareas
Private Sub btnGuardarCambios_Click()
    Dim tareaID As Long
    tareaID = CLng(Me.txtEditTaskID.Tag)
    If tareaID = 0 Then
        MsgBox "No hay tarea cargada.", vbExclamation
        Exit Sub
    End If

    Dim newName As String
    newName = Trim(Me.txtNewTaskName.Value)
    If newName = "" Then
        MsgBox "Ingrese el nombre de la tarea.", vbExclamation
        Exit Sub
    End If

    Dim wsT As Worksheet, tblT As ListObject
    Set wsT = ThisWorkbook.Worksheets(SHEET_TAREAS)
    Set tblT = wsT.ListObjects(TABLE_TAREAS_NAME)

    ' Encontrar fila en tabla tareas
    Dim found As Range, rowIndex As Long
    Set found = tblT.ListColumns("tarea_id").DataBodyRange.Find(What:=tareaID, LookAt:=xlWhole)
    If found Is Nothing Then
        MsgBox "Tarea no encontrada en tabla origen.", vbCritical
        Exit Sub
    End If
    rowIndex = found.Row - tblT.DataBodyRange.Row + 1

    ' Actualizar nombre
    tblT.DataBodyRange.Cells(rowIndex, tblT.ListColumns("TAREA").Index).Value = newName

    ' --------- Reemplazar relaciones en tabla puente ----------
    Dim wsPT As Worksheet, tblPT As ListObject
    Set wsPT = ThisWorkbook.Worksheets(SHEET_PT)
    Set tblPT = wsPT.ListObjects(TABLE_PT_NAME)

    ' Eliminar filas existentes con nro_tarea = tareaID (iterar hacia atrás)
    Dim i As Long
    For i = tblPT.ListRows.Count To 1 Step -1
        If CLng(tblPT.DataBodyRange.Cells(i, tblPT.ListColumns("nro_tarea").Index).Value) = tareaID Then
            tblPT.ListRows(i).Delete
        End If
    Next i

    ' Añadir nuevas relaciones desde lstAssigned
    Dim itemText As String, pid As Long, newRowPT As ListRow
    For i = 0 To Me.lstAssigned.ListCount - 1
        itemText = Me.lstAssigned.List(i)
        pid = CLng(Split(itemText, " - ")(0))
        If pid = 0 Then
            ' intentar resolver por nombre
            Dim pName As String
            pName = Trim(Mid(itemText, InStr(itemText, " - ") + 3))
            pid = GetPersonaIDByName(pName)
        End If
        If pid > 0 Then
            Set newRowPT = tblPT.ListRows.Add
            newRowPT.Range.Cells(1, tblPT.ListColumns("nro_persona").Index).Value = pid
            newRowPT.Range.Cells(1, tblPT.ListColumns("nro_tarea").Index).Value = tareaID
        End If
    Next i

    ' Refrescar vista
    On Error Resume Next
    Application.ScreenUpdating = False
    RefreshTablaControl
    Application.ScreenUpdating = True
    On Error GoTo 0

    MsgBox "Tarea guardada correctamente.", vbInformation
    Unload Me
End Sub

Private Sub btnCancelarEdit_Click()
    Unload Me
End Sub

