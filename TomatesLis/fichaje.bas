  Global codmobi As Integer
  Global dni As String
  Global nom_tra As String
  Global cod_nom As Integer

Function mapeo(buscar As Boolean) As String
  On Error GoTo NO_ENCONTRADO
  Dim minutos As Integer
  Dim horas As Double
  Dim t_fecha As String
  
 If buscar Then
  codmobi = ActiveCell.Value
  dni = Mid(Replace(ActiveCell.Offset(0, 3), "-", ""), 1, 9)
  nom_tra = ActiveCell.Offset(0, 1).Text
  cod_nom = Application.VLookup(dni, ThisWorkbook.Sheets("NOMINA").Range("$B$2:$T$111"), 2, False)
 End If
  horas = Val(Format(ActiveCell.Offset(0, 5).Text, "HH"))
  minutos = Val(Right(Format(ActiveCell.Offset(0, 5).Value, "HH:MM"), 2))
  If minutos >= 30 Then
    horas = horas + (5 / 10)
  End If
  t_fecha = Mid(ActiveCell.Offset(0, 4).Text, 1, 10)
  t_fecha = Mid(t_fecha, 4, 3) & Mid(t_fecha, 1, 3) & Right(t_fecha, 4)
  
  mapeo = "MOBIBUK: " & codmobi & " - " & nom_tra & _
          " - " & IIf(IsNumeric(Mid(dni, 1, 1)), "DNI", "NIE") & ": " & _
          dni & " - FECHA: " & ActiveCell.Offset(0, 4).Text & " - HORAS: " & Format(horas, "#.0") & " - " & minutos & _
          semana_grabar(t_fecha) & _
          " - Cod Nomina: "
  mapeo = mapeo & cod_nom
  Exit Function
NO_ENCONTRADO:
  mapeo = mapeo & "NO ENCONTRADO "
End Function

Sub PRUEBA2()
  Dim act_code As Variant
  Dim linea As Integer
  Dim sebusca As Boolean
  
  Sheets("Fichaje").Activate
  Range("A1").Select
  linea = 1
  Do While True
    ActiveCell.Offset(1, 0).Activate
    act_code = ActiveCell.Value
    If IsNumeric(act_code) And act_code <> "" Then
      sebusca = True
      act_code = ActiveCell.Value
    Else
      sebusca = False
    End If
      Debug.Print linea & " " & mapeo(sebusca)
    linea = linea + 1
    If ActiveCell.Offset(0, 4).Value = "" Or ActiveCell.Row >= 2000 Then
      Exit Do
    End If
  Loop
  Range("A1").Select
End Sub
Function semana_grabar(fecha) As String
  Dim ndia, d1, d2, d3, d4, d5, d_grabar, dia_fecha, ndia_g As Integer
  Dim semana As String
  Dim fecha_1 As Date ', fecha
  'fecha = CVDate("01/09/2024")
  fecha_1 = CVDate("01/" & Format(fecha, "MM/YYYY"))
  dia_fecha = Day(fecha)
  ndia = Weekday(fecha_1)
  ndia_g = IIf(Weekday(fecha) = 1, 8, Weekday(fecha))
  semana = "SEMANA_"
  d1 = IIf(ndia = 1, 1, 1 + 8 - ndia)
  d2 = d1 + 7
  d3 = d2 + 7
  d4 = d3 + 7
  d5 = Day(CVDate(Application.EoMonth(fecha, 0)))
  Select Case dia_fecha
    Case 1 To d1
      semana = semana & "1"
    Case d1 To d2
      semana = semana & "2"
    Case d2 To d3
      semana = semana & "3"
    Case d3 To d4
      semana = semana & "4"
    Case d4 To d5
      semana = semana & "5"
  End Select
  d_grabar = (4 * ndia_g) - 4
  'MsgBox fecha & dia_fecha & " semana: " & semana & "," & ndia & "," & d_grabar
  semana_grabar = semana & "-" & d_grabar
End Function
Sub p2()
  MsgBox semana_grabar(CVDate("15/08/2024"))
End Sub

RGB(0,0,0) NEGRO
RGB(0,255,0) VERDE