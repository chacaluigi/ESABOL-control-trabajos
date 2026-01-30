VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditarTarea 
   Caption         =   "EDITAR TAREA"
   ClientHeight    =   3696
   ClientLeft      =   -12
   ClientTop       =   -216
   ClientWidth     =   3276
   OleObjectBlob   =   "frmEditarTarea.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frmEditarTarea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCerrar_Click()
    Unload Me
End Sub

' Asume que module3 tiene YEAR_REF, RecalcularTareaEnTabla, ActualizarDiaEnTablaOrigen, ActualizarTareaOrigen, ColorFromName, AplicarColorDiaEnTablaOrigen

Private Sub txtPorcentaje_Change()

End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Left = (Application.Width - Me.Width) / 2
    Me.Top = (Application.Height - Me.Height) / 2
    
    ' Poblar lista de colores semanticamente
    With Me.cboColor
        .Clear
        .AddItem "Amarillo"
        .AddItem "Rojo"
        .AddItem "Naranja"
        .AddItem "Celeste"
        .AddItem "Verde oscuro"
        .AddItem "Gris"
        .AddItem "Verde claro"
        .AddItem "Café"
        .ListIndex = 0 ' selecciona Amarillo por defecto
    End With

    ' Default UI
    Me.chkTerminado.Value = False
    Me.lblTotalPorc.Caption = "Total: 0%"
End Sub

' Llamar desde la hoja antes de .Show, o llamar SetupFromSheet después de asignar txtId, txtTarea, etc.
Public Sub SetupFromSheet()
    Dim tareaId As Long
    Dim fechaIni As Variant, fechaFin As Variant, sumaPorc As Double

    On Error Resume Next
    tareaId = CLng(Me.txtId.Value)
    On Error GoTo 0

    If tareaId = 0 Then Exit Sub

    ' Recalcular desde la tabla origen (valor real)
    RecalcularTareaEnTabla tareaId, fechaIni, fechaFin, sumaPorc

    ' Mostrar porcentaje total
    Me.lblTotalPorc.Caption = "Total: " & Format(Round(sumaPorc, 0), "0") & "%"

    ' Si la tarea alcanzó 100% o ya tiene fecha final -> marcar como terminada
    If Not IsEmpty(fechaFin) Or sumaPorc >= 100 Then
        Me.chkTerminado.Value = True
        ' Mostrar fecha final si hay
        If Not IsEmpty(fechaFin) Then
            Me.txtFinal.Value = fechaFin
        Else
            Me.txtFinal.Value = "" ' o Date si prefieres
        End If
    Else
        Me.chkTerminado.Value = False
        ' Ocultar o limpiar fecha final si no está terminada
        Me.txtFinal.Value = ""
    End If

    ' Ajustar spinner por defecto: mostrar 0 (usuario elegirá)
    Me.spnPorcentaje.Value = 0
    Me.txtPorcentaje.Value = 0
End Sub

Private Sub spnPorcentaje_Change()
    txtPorcentaje.Value = spnPorcentaje.Value
End Sub

Private Sub btnElegirFecha_Click()
    Dim s As String
    s = InputBox("Introduce la fecha (ej: 12-ene-2026) o escribe DD/MM/AAAA:", "Elegir fecha")
    If Trim(s) = "" Then Exit Sub
    If Not IsDate(s) Then
        MsgBox "Fecha inválida. Usa formato DD/MM/AAAA o 12-ene-2026", vbExclamation
        Exit Sub
    End If
    Dim d As Date
    d = CDate(s)
    If Year(d) <> YEAR_REF Then
        If MsgBox("La fecha no pertenece al año " & YEAR_REF & ". ¿Continuar?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If
    txtFecha.Value = Format(d, "d-mmm-yyyy")
    Call CargarValorDiaExistente(d)
End Sub

Private Sub CargarValorDiaExistente(d As Date)
    ' intenta precargar el valor existente en la hoja de control para este tarea_id y día
    On Error Resume Next
    Dim ws As Worksheet
    Dim tareaId As Long
    Dim fila As Range
    Dim diaNum As Long
    Set ws = ThisWorkbook.Worksheets(SHEET_CONTROL)
    tareaId = CLng(txtId.Value)
    Set fila = ws.Columns(2).Find(What:=tareaId, LookAt:=xlWhole)
    If fila Is Nothing Then Exit Sub
    diaNum = Day(d)
    Dim col As Long
    col = COL_DIA_INICIO + diaNum - 1
    If IsNumeric(ws.Cells(fila.Row, col).Value) Then
        spnPorcentaje.Value = ws.Cells(fila.Row, col).Value
        txtPorcentaje.Value = spnPorcentaje.Value
    Else
        spnPorcentaje.Value = 0
        txtPorcentaje.Value = 0
    End If
End Sub

Private Sub chkTerminado_Click()
    ' Si se marca como terminado y no hay fecha seleccionada, asigna hoy como fecha objetivo
    If Me.chkTerminado.Value = True Then
        If Trim(Me.txtFecha.Value) = "" Then
            Me.txtFecha.Value = Format(Date, "d-mmm-yyyy")
        End If
    Else
        ' Si se desmarca, elimina txtFinal (no persistimos hasta Guardar)
        Me.txtFinal.Value = ""
    End If
End Sub

Private Sub btnAgregarDia_Click()
    ' Valida fecha
    If Trim(txtFecha.Value) = "" Then
        MsgBox "Debe elegir una fecha primero (btn 'Elegir Fecha').", vbExclamation
        Exit Sub
    End If

    If Not IsDate(txtFecha.Value) Then
        MsgBox "Fecha inválida.", vbExclamation
        Exit Sub
    End If

    Dim d As Date
    d = CDate(txtFecha.Value)

    If Year(d) <> YEAR_REF Then
        If MsgBox("La fecha no pertenece al año " & YEAR_REF & ". ¿Desea continuar?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If

    ' Preparar variables
    Dim tareaId As Long
    tareaId = CLng(txtId.Value)
    Dim diaNum As Long
    diaNum = Day(d)

    ' Obtener suma actual antes de agregar (desde tabla)
    Dim fechaIni As Variant, fechaFin As Variant, sumaPorc As Double
    RecalcularTareaEnTabla tareaId, fechaIni, fechaFin, sumaPorc

    ' Si se marca terminado -> calcular cuanto falta para llegar a 100
    Dim pctToWrite As Variant
    If Me.chkTerminado.Value = True Then
        pctToWrite = Application.WorksheetFunction.Max(0, 100 - sumaPorc)
    Else
        ' Si no está marcado terminado, usar spinner
        If Not IsNumeric(txtPorcentaje.Value) Then
            MsgBox "Porcentaje inválido.", vbExclamation
            Exit Sub
        End If
        pctToWrite = CDbl(txtPorcentaje.Value)
        If pctToWrite < 0 Or pctToWrite > 100 Then
            MsgBox "El porcentaje debe estar entre 0 y 100.", vbExclamation
            Exit Sub
        End If
    End If

    ' Si color elegido NO es Amarillo -> no escribir valor, solo aplicar color
    Dim colorName As String, colorLong As Long
    colorName = Me.cboColor.Value
    colorLong = ColorFromName(colorName)

    If LCase(Trim(colorName)) = "amarillo" Then
        ' Escribe porcentaje (si es 0 entonces dejar 0)
        Call ActualizarDiaEnTablaOrigen(tareaId, diaNum, pctToWrite)
    Else
        ' No escribir valor, dejar celda vacía
        Call ActualizarDiaEnTablaOrigen(tareaId, diaNum, Empty)
    End If

    ' Aplicar color en la tabla origen
    Call AplicarColorDiaEnTablaOrigen(tareaId, diaNum, colorLong)

    ' Si se marcó terminado, establecer FECHA FINAL (usar la fecha elegida)
    If Me.chkTerminado.Value = True Then
        fechaFin = d
    End If

    ' Recalcular desde tabla y persistir FECHAS y PORCENTAJE
    RecalcularTareaEnTabla tareaId, fechaIni, fechaFin, sumaPorc
    Call ActualizarTareaOrigen(tareaId, fechaIni, fechaFin, sumaPorc)

    ' Forzar refresh de la vista pegada
    On Error Resume Next
    Application.ScreenUpdating = False
    RefreshTablaControl
    Application.ScreenUpdating = True
    On Error GoTo 0

    MsgBox "Avance del día " & diaNum & " procesado (color: " & colorName & ").", vbInformation

    ' Actualizar UI con nuevos totales
    Me.lblTotalPorc.Caption = "Total: " & Format(Round(sumaPorc, 0), "0") & "%"
    If sumaPorc >= 100 Then
        Me.chkTerminado.Value = True
        Me.txtFinal.Value = fechaFin
    End If

    ' Opcional: actualizar spinner con suma truncada (para info)
    Me.spnPorcentaje.Value = Application.WorksheetFunction.Min(100, Application.WorksheetFunction.RoundDown(sumaPorc, 0))
    Me.txtPorcentaje.Value = Me.spnPorcentaje.Value
End Sub

Private Sub btnGuardar_Click()

    Dim tbl As ListObject
    Dim fila As Range

    Set tbl = ThisWorkbook.Sheets(SHEET_TAREAS).ListObjects(TABLE_TAREAS_NAME)

    Set fila = tbl.ListColumns("tarea_id").DataBodyRange.Find( _
        What:=CLng(txtId.Value), LookAt:=xlWhole)

    If fila Is Nothing Then
        MsgBox "Tarea no encontrada", vbCritical
        Exit Sub
    End If

    ' Guardar campos básicos
    fila.Offset(0, 1).Value = txtTarea.Value

    

    ' Si está marcada como terminada -> fijar fecha final y porcentaje 100%
    If Me.chkTerminado.Value = True Then
        Dim finalDate As Variant
        If IsDate(Me.txtFinal.Value) Then
            finalDate = CDate(Me.txtFinal.Value)
        ElseIf IsDate(Me.txtFecha.Value) Then
            finalDate = CDate(Me.txtFecha.Value)
        Else
            finalDate = Date
        End If
        fila.Offset(0, 3).Value = finalDate
        fila.Offset(0, 4).Value = 1 ' 100%
    Else
        ' No terminada: mantener/limpiar fecha final según campo txtFinal
        If IsDate(txtFinal.Value) Then
            fila.Offset(0, 3).Value = CDate(txtFinal.Value)
        Else
            fila.Offset(0, 3).ClearContents
        End If

        ' Recalcular porcentajes reales a partir de columnas 1..31
        Dim fechaIni As Variant, fechaFi As Variant, sumaPorc As Double
        RecalcularTareaEnTabla CInt(txtId.Value), fechaIni, fechaFi, sumaPorc
        fila.Offset(0, 4).Value = Application.WorksheetFunction.Min(sumaPorc / 100#, 1#)
    End If

    ' REFRESH: actualizar la vista pegada
    On Error Resume Next
    Application.ScreenUpdating = False
    RefreshTablaControl
    Application.ScreenUpdating = True
    On Error GoTo 0

    Unload Me
    MsgBox "Tarea actualizada correctamente", vbInformation

End Sub




