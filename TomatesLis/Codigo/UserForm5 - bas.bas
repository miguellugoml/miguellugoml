Private Sub UserForm_Activate()
'UserForm5 - VARIABLES
Dim Conteo, cuantos As Long
Dim Porcentaje As Double
  Dim act_code, linea, linea_Mes As Integer
  Dim normal, mv, pp, su_color As Double
  Dim act_nombre, voler_a, txt_Fecha As String
  
  If el_mes = "" Then
    carga_mes
  End If
  Application.ScreenUpdating = False
  Sheets("VARIABLES").Activate
  txt_Fecha = Format(Day(WorksheetFunction.EoMonth(CVDate("01/" & el_mes & "/" & el_anho), 0)) & "/" & el_mes & "/" & el_anho, "dd \de MMMM \de YYYY")
  Range(Range("A1:C6").Find("FECHA:").Address).Select
  ActiveCell.Offset(0, 1).Value = txt_Fecha
  Range("B1").Select
  Selection.End(xlDown).Select
  ActiveCell.Offset(2, 0).Select
  volver_a = Replace(ActiveCell.Address, "$", "")
    Range("B9:J700").Select
    ActiveWindow.SmallScroll Down:=-236
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
  LosQueTrabajaron "VARIABLES"
  Range(volver_a).Select
  cuantos = WorksheetFunction.Count(Sheets("VARIABLES").Range("B1:B500"))
  Conteo = 1
  Sheets("VARIABLES").Activate
  Range("B1").Select
  Do While True
    ActiveCell.Offset(1, 0).Select
    If IsNumeric(ActiveCell.Text) Then
      act_code = ActiveCell.Value
      Exit Do
    End If
  Loop
  Do While IsNumeric(ActiveCell.Text)
    Sheets(el_mes).Activate
    Range(Range("A1:A150").Find(act_code).Address).Select
    linea = ActiveCell.Row
    normal = Range("X" & linea).Value
    mv = Range("Y" & linea).Value
    pp = Range("Z" & linea).Value
    'MsgBox act_code & " - N: " & Range("X" & linea).Value & " - MV: " & Range("Y" & linea).Value & " - PP: " & Range("Z" & linea).Value
    Sheets("VARIABLES").Activate
    ActiveCell.Offset(0, 2).Value = normal
    ActiveCell.Offset(0, 3).Value = mv
    ActiveCell.Offset(0, 4).Value = pp
    ActiveCell.Offset(1, 0).Select
    act_code = ActiveCell.Value
   Conteo = Conteo + 1
   Porcentaje = Conteo / cuantos
   Me.Caption = "Espere, por favor... " & Format(Porcentaje, "0%")
   Me.Label1.Width = Porcentaje * Me.Frame1.Width
   DoEvents
  Loop
  'COPIAR CODIGOS CON DATOS A VARIABLS PARA COMPARAR Y SABER QUIEN FALTA
  Range("H9").Select
  Sheets(el_mes).Activate
  Range("A1").Select
  Do While True
    ActiveCell.Offset(1, 0).Select
    If IsNumeric(ActiveCell.Text) Then
      ActiveCell.Offset(-1, 0).Select
      linea = ActiveCell.Row
      Exit Do
    End If
    If ActiveCell.Row > 150 Then
      Exit Do
    End If
  Loop
  
  Do While linea < 500
    ActiveCell.Offset(1, 0).Select
    linea = ActiveCell.Row
    If IsNumeric(ActiveCell.Value) Then
     If ActiveCell.Value > 500 Then
      act_code = ActiveCell.Value
      act_nombre = ActiveCell.Offset(0, 1).Value
      su_color = ActiveCell.Offset(0, 1).Font.Color
      linea = ActiveCell.Row
 '     If Range("X" & linea).Value <> 0 Or Range("Y" & linea).Value <> 0 Or Range("Z" & linea).Value <> 0 Then
        Sheets("VARIABLES").Activate
        ActiveCell.Value = act_code
        ActiveCell.Offset(0, 1).Value = act_nombre
        ActiveCell.Offset(0, 1).Font.Color = su_color
        ActiveCell.Offset(1, 0).Select
        Sheets(el_mes).Activate
      End If
    End If
  Loop
  Sheets("VARIABLES").Activate
  Resalta_Duplicados
  AgregarNoEncontradosVariables
  Application.ScreenUpdating = True
  Sheets(el_mes).Activate
   Unload Me
End Sub




