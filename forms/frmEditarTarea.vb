Option Explicit

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
    ' opcional: ajustar spinner al valor existente del día (si existe)
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

    ' Valida porcentaje
    Dim pct As Double
    If Not IsNumeric(txtPorcentaje.Value) Then
        MsgBox "Porcentaje inválido.", vbExclamation
        Exit Sub
    End If
    pct = CDbl(txtPorcentaje.Value)
    If pct < 0 Or pct > 100 Then
        MsgBox "El porcentaje debe estar entre 0 y 100.", vbExclamation
        Exit Sub
    End If

    ' Preparar variables
    Dim tareaId As Long
    tareaId = CLng(txtId.Value)
    Dim diaNum As Long
    diaNum = Day(d)

    ' --- 1) Escribir directamente en la TABLA origen (columna "diaNum")
    Call ActualizarDiaEnTablaOrigen(tareaId, diaNum, pct)

    ' --- 2) Recalcular usando la fila de la TABLA (no la vista)
    Dim fechaIni As Variant, fechaFin As Variant, sumaPorc As Double
    RecalcularTareaEnTabla tareaId, fechaIni, fechaFin, sumaPorc

    ' --- 3) Guardar FECHA INICIO / FECHA FINAL / PORCENTAJE en la tabla origen
    Call ActualizarTareaOrigen(tareaId, fechaIni, fechaFin, sumaPorc)
    

    ' --- NUEVO: aplicar color seleccionado a la celda del día en la tabla origen
    Dim colorName As String, colorLong As Long
    colorName = Me.cboColor.Value
    colorLong = ColorFromName(colorName)
    Call AplicarColorDiaEnTablaOrigen(tareaId, diaNum, colorLong)

    
    ' --- 4) Informar y salir (la hoja de control se actualizará por FILTRAR automáticamente)
    
    ' Pero forzamos un Refresh para actualizar inmediatamente la vista pegada
    On Error Resume Next
    Application.ScreenUpdating = False
    RefreshTablaControl
    Application.ScreenUpdating = True
    On Error GoTo 0

    MsgBox "Avance del día " & diaNum & " = " & pct & "% guardado en la tabla.", vbInformation

    ' Opcional: actualizar los controles del formulario con los nuevos valores
    If Not IsEmpty(fechaIni) Then Me.txtInicio.Value = fechaIni
    If Not IsEmpty(fechaFin) Then Me.txtFinal.Value = fechaFin
    Me.spnPorcentaje.Value = Application.WorksheetFunction.RoundDown(sumaPorc, 0)
    Me.txtPorcentaje.Value = spnPorcentaje.Value

End Sub

