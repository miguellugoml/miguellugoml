Private Sub UserForm_Activate()
Dim Conteo, cuantos As Long
Dim act_code As Variant
Dim linea As Integer
Dim sebusca As Boolean
Dim Porcentaje As Double
 Application.ScreenUpdating = False
 carga_mes
 'Sheets(el_mes).Activate
 cuantos = WorksheetFunction.Count(Sheets("Fichaje").Range("A1:A2000")) * n_semanas * 7
 Conteo = 1
 Sheets("Fichaje").Activate
 Range("A1").Select
 linea = 1
 Do While cuantos > 0
   ActiveCell.Offset(1, 0).Activate
   act_code = ActiveCell.Value
   If IsNumeric(act_code) And act_code <> "" Then
     sebusca = True
     act_code = ActiveCell.Value
   Else
     sebusca = False
   End If
   'Debug.Print linea & " " &
   mapeo (sebusca)
   linea = linea + 1
   Conteo = Conteo + 1
   Porcentaje = Conteo / cuantos
   Me.Caption = "Espere, por favor... " & Format(Porcentaje, "0%")
   Me.Label1.Width = Porcentaje * Me.Frame1.Width
   DoEvents
   If ActiveCell.Offset(0, 4).Value = "" Or ActiveCell.Row >= 2000 Then
     Exit Do
   End If
 Loop
 Range("A1").Select
 Application.ScreenUpdating = True
 Unload Me
 
End Sub
