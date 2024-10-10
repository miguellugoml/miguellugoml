Private Sub btn_Cancelar_Click()
End
Unload Me
End Sub

Private Sub UserForm_Activate()
Dim Conteo As Long
Dim nFilas As Long
Dim nColumnas As Long
Dim f As Long
Dim c As Long
Dim Porcentaje As Double
 carga_mes
 Sheets(el_mes).Activate
 Conteo = WorksheetFunction.Count(Sheets("JORNADAS").Range("B1:B500"))
 MsgBox Conteo
    'Cells.Clear
    Conteo = 1
    nFilas = 5000
    nColumnas = 100
    
        For f = 1 To nFilas
            For c = 1 To nColumnas
                'Cells(f, c) = Conteo
                Conteo = Conteo + 1
            Next c
                Porcentaje = Conteo / (nFilas * nColumnas)
                Me.Caption = Format(Porcentaje, "0%")
                Me.Label1.Width = Porcentaje * Me.Frame1.Width
                DoEvents
        Next f
        Unload Me
End Sub

