Dim arrPersonas As Variant

Private Sub btnCerrar_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()

    Dim rng As Range

    ' Cargar nombres desde la tabla personal
    Set rng = Sheets("personal").ListObjects("personal") _
              .ListColumns("Apellidos y Nombres").DataBodyRange

    arrPersonas = rng.Value

    lstPersonas.List = arrPersonas

    ' Centrar formulario
    Me.StartUpPosition = 0
    Me.Left = (Application.Width - Me.Width) / 2
    Me.Top = (Application.Height - Me.Height) / 2

End Sub

Private Sub txtBuscar_Change()

    Dim i As Long
    lstPersonas.Clear

    For i = LBound(arrPersonas) To UBound(arrPersonas)
        If InStr(1, arrPersonas(i, 1), txtBuscar.Text, vbTextCompare) > 0 Then
            lstPersonas.AddItem arrPersonas(i, 1)
        End If
    Next i

End Sub

Private Sub lstPersonas_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call SeleccionarPersona
End Sub

Private Sub btnAceptar_Click()
    Call SeleccionarPersona
End Sub

Private Sub SeleccionarPersona()

    If lstPersonas.ListIndex = -1 Then
        MsgBox "Seleccione una persona", vbExclamation
        Exit Sub
    End If

    Dim s As String
    s = lstPersonas.Value

    ' Guardar en variables globales para que otros formularios lo usen
    gSelectedPersonName = s
    gSelectedPersonID = 0

    ' Intentar obtener el persona_id desde la tabla "personal"
    On Error Resume Next
    Dim wsP As Worksheet
    Dim tblP As ListObject
    Dim foundP As Range
    Set wsP = ThisWorkbook.Worksheets("personal")
    Set tblP = wsP.ListObjects("personal")
    On Error GoTo 0

    If Not tblP Is Nothing Then
        Set foundP = tblP.ListColumns("Apellidos y Nombres").DataBodyRange.Find(What:=s, LookAt:=xlWhole, MatchCase:=False)
        If Not foundP Is Nothing Then
            ' suponer que la columna de id se llama "persona_id"
            gSelectedPersonID = CLng(tblP.DataBodyRange.Cells(foundP.Row - tblP.DataBodyRange.Row + 1, tblP.ListColumns("persona_id").Index).Value)
        End If
    End If

    ' Escribir en C5 / refrescar solo si la bandera lo permite (comportamiento anterior)
    If gBuscarPersona_WriteToC5 <> False Then
        With ThisWorkbook.Worksheets("tabla_control")
            .Range("C5").Value = s
        End With

        On Error Resume Next
        Application.ScreenUpdating = False
        RefreshTablaControl
        Application.ScreenUpdating = True
        On Error GoTo 0
    End If

    Unload Me

End Sub

