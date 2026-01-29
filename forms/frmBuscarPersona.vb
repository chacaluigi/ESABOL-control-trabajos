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

    ' Escribir sólo en C5 (no tocar A6 para no sobrescribir la fórmula)
    With ThisWorkbook.Worksheets("tabla_control")
        .Range("C5").Value = s
    End With

    Unload Me

    ' Ejecutar refresco de la vista (llama a la macro que copia desde la tabla origen)
    On Error Resume Next
    Application.ScreenUpdating = False
    RefreshTablaControl
    Application.ScreenUpdating = True
    On Error GoTo 0

End Sub


