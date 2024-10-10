Private Sub UserForm_Activate()
'UserForm4 - JORNADAS
Dim ndia, d1, d2, d3, d4, d5, d6, dia_fecha, ndia_g, dia_final, linea As Integer
Dim semanas_mes, dini As Integer
Dim Conteo, cuantos As Long
Dim act_code As Variant
Dim sebusca As Boolean
Dim Porcentaje As Double
Dim d(6) As Integer
Dim Fecha, fecha_1, fcha_ing As Date
Dim letras, msg, act_Hoja, cel_back, txt_Fecha As String
Dim x(31) As String

Application.ScreenUpdating = False
carga_mes
txt_Fecha = Format(Day(WorksheetFunction.EoMonth(CVDate("01/" & el_mes & "/" & el_anho), 0)) & "/" & el_mes & "/" & el_anho, "dd \de MMMM \de YYYY")
Sheets("JORNADAS").Activate
Range(Range("A1:C7").Find("FECHA:").Address).Select
ActiveCell.Offset(0, 1).Value = txt_Fecha
cuantos = WorksheetFunction.Count(Sheets("JORNADAS").Range("B1:B500")) * 5.2
If cuantos <> 0 Then
 act_Hoja = ActiveSheet.Name
 Sheets("JORNADAS").Activate
 Range("B1").Select
 Do While True
   If IsNumeric(ActiveCell.Text) Then
     Exit Do
   End If
   ActiveCell.Offset(1, 0).Select
 Loop
 cel_back = Replace(ActiveCell.Address, "$", "")
 Rows(ActiveCell.Row & ":" & Range("Total_JORNADAS").Row - 1).Select
 Selection.Delete Shift:=xlUp
 Selection.ClearContents
 Rows(Range(cel_back).Row & ":" & Range(cel_back).Row).Select
 Selection.ClearContents
End If
LosQueTrabajaron "JORNADAS"
cuantos = WorksheetFunction.Count(Sheets("JORNADAS").Range("B1:B500")) * 5.2
If cuantos > 0 Then
  Conteo = 1
  act_code = 0
  n_semanas = 0
  carga_mes
  letras = "CFILOR"
  linea = 2
  Fecha = CVDate("01/" & Format((InStr(1, meses, mes_hoja) + 9) / 10, "00") & "/2024")
  fecha_ing = Mid(Fecha, 4, 3) & Mid(Fecha, 1, 3) & Right(Fecha, 4)
  fecha_1 = CVDate("01/" & Format(Fecha, "MM/YYYY"))
  dia_fecha = Day(Fecha)
  ndia = Weekday(fecha_1)
  ndia_g = IIf(Weekday(Fecha) = 1, 8, Weekday(Fecha))
  dia_final = Day(CVDate(Application.EoMonth(IIf(dia_fecha > 12, fecha_ing, Fecha), 0)))
  
  d(1) = IIf(ndia = 1, 1, 1 + 8 - ndia)
  d(2) = d(1) + 7
  d(3) = d(2) + 7
  d(4) = d(3) + 7
  If d(4) >= dia_final Then
    d(4) = dia_final
  End If
  If d(4) < dia_final Then
    d(5) = d(4) + 7
  End If
  If d(5) >= dia_final Then
    d(5) = dia_final
  End If
  If d(5) < dia_final Then
    d(6) = dia_final
  End If
  For i = 1 To 6
    msg = msg & d(i) & " -"
  Next i
  'MsgBox msg
  Sheets("JORNADAS").Activate
  Range("B1").Select
  Do While True
    ActiveCell.Offset(1, 0).Select
    If IsNumeric(ActiveCell.Text) Then
      act_code = ActiveCell.Value
      Exit Do
    End If
  Loop
 ' MsgBox "Comienza a trabajar con el codigo: " & act_code & " para " & n_semanas
 Do While IsNumeric(ActiveCell.Text)
  For i = 1 To n_semanas
   Sheets("SEMANA_" & Mid(mes_hoja, 1, 3) & "_" & i).Activate
   Range(Range("B1:B1000").Find(act_code).Address).Select
   dini = (-4 * d(1)) + 32
   d1 = 1
   If i = 1 Then
    msg = "semana " & i & " comienza por "
    For j = dini To 28 Step 4
      x(d1) = IIf(ActiveCell.Offset(0, j).Value <> 0, "x", "")
      d1 = d1 + 1
      msg = msg & j & " - "
    Next j
   Else
    msg = "semana " & i & " comienza por "
    For j = 4 To (d(i) - d(i - 1)) * 4 Step 4
      x(d(i - 1) + (j / 4)) = IIf(ActiveCell.Offset(0, j).Value <> 0, "x", "")
      msg = msg & j & " - "
    Next j
   End If
   Conteo = Conteo + 1
   Porcentaje = Conteo / cuantos
   Me.Caption = "Espere, por favor... " & Format(Porcentaje, "0%")
   Me.Label1.Width = Porcentaje * Me.Frame1.Width
   DoEvents
   Sheets("JORNADAS").Activate
   For j = 1 To 31
    ActiveCell.Offset(0, j + 1).Value = x(j)
   Next j
  Next i
  ActiveCell.Offset(1, 0).Select
  act_code = ActiveCell.Value
 Loop
End If
Range("A1").Select
Application.ScreenUpdating = True
Unload Me

End Sub

