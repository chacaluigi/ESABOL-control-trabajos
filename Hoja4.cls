VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    If Target.Column <> 2 Or Target.Row < 8 Then Exit Sub
    Cancel = True

    Dim pct As Double
    pct = Target.Offset(0, 4).Value * 100

    With frmEditarTarea
        .txtId.Value = Target.Value
        .txtTarea.Value = Target.Offset(0, 1).Value
        .txtInicio.Value = Target.Offset(0, 2).Value
        .txtFinal.Value = Target.Offset(0, 3).Value
        .spnPorcentaje.Value = pct
        .txtPorcentaje.Value = pct
        ' inicializar campo fecha vacío (usuario elegirá)
        .txtFecha.Value = ""
        ' cargar totales y estado desde la tabla origen
        .SetupFromSheet
        .Show
    End With

End Sub


Private Sub Worksheet_Change(ByVal Target As Range)

    Dim fila As Long
    Dim fechaIni As Variant, fechaFin As Variant
    Dim sumaPorc As Double
    Dim tareaId As Long

    ' Columnas de días
    If Target.Column < COL_DIA_INICIO Or Target.Column > COL_DIA_FIN Then Exit Sub
    If Target.Row < 8 Then Exit Sub

    Application.EnableEvents = False

    fila = Target.Row
    tareaId = CLng(Cells(fila, "B").Value)

    ' recalcular la fila de la vista
    RecalcularFilaControl Me, fila, fechaIni, fechaFin, sumaPorc

    ' actualizar vista
    Cells(fila, "D").Value = IIf(IsEmpty(fechaIni), vbNullString, fechaIni)
    Cells(fila, "E").Value = IIf(IsEmpty(fechaFin), vbNullString, fechaFin)
    Cells(fila, "F").Value = sumaPorc / 100
    Cells(fila, "F").NumberFormat = "0%"

    ' guardar: actualiza FECHAS y PORCENTAJE en tabla origen
    ActualizarTareaOrigen tareaId, fechaIni, fechaFin, sumaPorc

    ' y también actualizar cada día en la tabla origen (sin iterar todo; actualizar solo la celda cambiada)
    ' si Target es una sola celda y contiene número >0 o 0, la guardamos
    If Target.Cells.Count = 1 Then
        If IsNumeric(Target.Value) Then
            Dim diaNum As Long
            diaNum = Target.Column - COL_DIA_INICIO + 1
            ' guardar valor (0..100) tal cual
            ActualizarDiaEnTablaOrigen tareaId, diaNum, CDbl(Target.Value)
        End If
    Else
        ' Si se pegaron varios días, sincronizamos todos (recorrer rango Target)
        Dim cel As Range
        For Each cel In Intersect(Target, Range(Cells(8, COL_DIA_INICIO), Cells(Rows.Count, COL_DIA_FIN)))
            If IsNumeric(cel.Value) Then
                diaNum = cel.Column - COL_DIA_INICIO + 1
                ActualizarDiaEnTablaOrigen tareaId, diaNum, CDbl(cel.Value)
            End If
        Next cel
    End If

    Application.EnableEvents = True
    
End Sub


