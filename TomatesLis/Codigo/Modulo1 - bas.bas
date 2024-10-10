Global c_nom(4) As String
Global n_veces(4)  As Integer
Global el_mes, el_anho As String
Global la_semana As String
Global n_semana As Integer
Global n_semanas As Integer
Global accion As String
Global siBuscas As Boolean
Global codmobi As Integer
Global dni As String
Global nom_tra As String
Global cod_Nom As Integer
Global meses As String
Global mes_hoja As String
Global Nombre_Archivo As String

Sub Inserta_Col_Codigo()
Sheets(la_semana).Activate
If Range("B2").Text = "Cod Empleado" Then
  'MsgBox "Ya se ha insertado la columna de Código de Empleado", vbOKOnly, "Inacción"
Else
  Range("B2:B100").Select
  Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
  Range("B2").Select
  ActiveCell = "Cod Empleado"
  Columns("B").AutoFit
  Columns("C").AutoFit
End If
End Sub
Sub Act_Dias()
 Dim celdas(7, 2), Hojas_Sem(6) As String
 Dim fec1 As Date
 Dim le_suma, inMes As Integer
 celdas(1, 1) = "D"
 celdas(1, 2) = "LUNES"
 celdas(2, 1) = "H"
 celdas(2, 2) = "MARTES"
 celdas(3, 1) = "L"
 celdas(3, 2) = "MIERCOLES"
 celdas(4, 1) = "P"
 celdas(4, 2) = "JUEVES"
 celdas(5, 1) = "T"
 celdas(5, 2) = "VIERNES"
 celdas(6, 1) = "X"
 celdas(6, 2) = "SABADO"
 celdas(7, 1) = "AB"
 celdas(7, 2) = "DOMINGO"
 For Each h In ThisWorkbook.Sheets
  If Mid(UCase(h.Name), 1, 6) = "SEMANA" Then
    Sheets(h.Name).Activate
    fec1 = Primer_dia_sem_mes_n(Range("AI2").Value, Mes_Sem_H("M"), Mes_Sem_H("S"))
    inMes = Mes_Sem_H("M")
    le_suma = 0
    For i = 1 To 7
      Range(celdas(i, 1) & "2").Value = celdas(i, 2) & " " & Format(fec1 + le_suma, "DD")
      If Val(Format(fec1 + le_suma, "MM")) <> inMes Then
        Range(celdas(i, 1) & "2").Font.Color = RGB(205, 205, 205)
      Else
        Range(celdas(i, 1) & "2").Font.Color = RGB(0, 0, 0)
      End If
      le_suma = le_suma + 1
    Next i
   End If
 Next
 If el_mes = "" Then
   carga_mes
 End If
 Sheets("MES").Activate
End Sub
Sub Carga_Codigo()
Dim c_nomb As String
Dim n_pos As Integer
Dim n_lin As Integer
Dim range_back_mes As String
siBuscas = False
Application.ScreenUpdating = False
If el_mes = "" Then
  carga_mes
End If
'la_semana = "SEMANA_3"
accion = "Procesar"
datos
Sheets(el_mes).Activate
range_back_mes = Replace(ActiveCell.AddressLocal, "$", "")
If accion = "Cancelar" Then
  Exit Sub
End If
If UCase(la_semana) = "TODAS" Then
  sem_ini = 1
  sem_fin = n_semanas
Else
  sem_ini = Mid(la_semana, Len(la_semana), 1)
  sem_fin = sem_ini
End If
For s = sem_ini To sem_fin
 n_semana = s 'Mid(la_semana, Len(la_semana), 1)
 la_semana = "SEMANA_" & Mid(el_mes, 1, 3) & "_" & s
' Inserta_Col_Codigo
 n_lin = 1
 Sheets(la_semana).Activate
 Range("C2").Activate
 ActiveCell.Offset(1, 0).Activate
' Do While n_lin < 20000
' c_nomb = ActiveCell.Text
' For i = 1 To 4
'  'MsgBox c_nomb
'  n_pos = InStr(1, c_nomb, " ")
'  If n_pos = 0 Then
'   c_nom(i) = c_nomb
'  Else
'   c_nom(i) = Mid(c_nomb, 1, n_pos - 1)
'  End If
'  c_nomb = Mid(c_nomb, n_pos + 1, Len(c_nomb) - n_pos)
' Next i
' BuscaCodigo
' n_lin = n_lin + 1
' ActiveCell.Offset(1, 1).Activate
' If ActiveCell.Text = "" Then
'   Do While ActiveCell.Text = "" And n_lin < 20000
'     ActiveCell.Offset(1, 0).Activate
'     If ActiveCell.Column = 1 Then
'       ActiveCell.Offset(0, 1).Activate
'     End If
'     n_lin = n_lin + 1
'   Loop
' End If
 'Loop
 Marca_Duplicados "B:B"
 Sheets(la_semana).Activate
 Range("C2").Activate
 Obten_horas
 Sheets(el_mes).Activate
Next s
Range(range_back_mes).Activate
siBuscas = True
Application.ScreenUpdating = True
End Sub
Sub BuscaCodigo()
Dim coincide, n_cod As Integer
Dim lineas As Integer

coincide = 0
Sheets(el_mes).Activate
Range("B6").Select
Do While n_cod = 0
For i = 1 To 4
 If InStr(1, ActiveCell.Text, c_nom(i)) <> 0 Then
   coincide = coincide + 1
 End If
Next i
If coincide > 1 Then
  ActiveCell.Offset(0, -1).Activate
  n_code = ActiveCell.Value
  Sheets(la_semana).Activate
  ActiveCell.Offset(0, -1).Activate
  If ActiveCell.Text = "" Then
    ActiveCell.Value = n_code
  End If
  Exit Do
Else
  coincide = 0
  ActiveCell.Offset(1, 0).Activate
  linea = linea + 1
  If linea = 20000 Then
    Exit Do
  End If
  If ActiveCell.Text = "Código" Then
    ActiveCell.Offset(0, 1).Activate
  End If
End If
Loop
End Sub
Sub buscacodigo_s()
la_semana = ActiveSheet.Name
'el_mes = "AGOSTO"
c_nomb = ActiveCell.Text
For i = 1 To 4
 'MsgBox c_nomb
 n_pos = InStr(1, c_nomb, " ")
 If n_pos = 0 Then
  c_nom(i) = c_nomb
 Else
  c_nom(i) = Mid(c_nomb, 1, n_pos - 1)
 End If
 c_nomb = Mid(c_nomb, n_pos + 1, Len(c_nomb) - n_pos)
Next i
BuscaCodigo
End Sub
Sub Marca_Duplicados(el_rango As String)

    Columns(el_rango).Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

Sub que_color_tiene()
  'MsgBox ActiveCell.Value & " - " & ActiveCell.Interior.Color
  MsgBox ActiveCell.Formula = ""
End Sub

Sub Obten_horas()
  Dim act_code, la_linea, cuantos As Integer
  Dim HN, mv, pp, horas_dia As Double
  Dim act_range, range_back, celda_a_volver As String
  Dim debe_salir, unavez As Boolean
  Application.ScreenUpdating = False
  unavez = False
  range_back = "B1"
  Range(range_back).Activate
  Do Until IsNumeric(ActiveCell.Text)
    ActiveCell.Offset(1, 0).Activate
  Loop
  act_code = ActiveCell.Value
  Do While True
    HN = 0
    mv = 0
    pp = 0
    debe_salir = False
    Sheets(la_semana).Activate
    'Range(range_back).Activate
    cuantos = 0
    If unavez Then
    Do Until IsNumeric(ActiveCell.Text) And ActiveCell.Value <> act_code
      ActiveCell.Offset(1, 0).Activate
      If cuantos > 10 Then
        Application.ScreenUpdating = True
        Sheets(el_mes).Activate
        Range("A1").Activate
        Exit Sub
      Else
        cuantos = cuantos + 1
      End If
    Loop
    End If
    unavez = True
    act_code = ActiveCell.Value
    range_back = Replace(ActiveCell.AddressLocal, "$", "")
    Do While True
      act_range = Replace(ActiveCell.AddressLocal, "$", "")
      For i = 1 To 7
        ActiveCell.Offset(0, 4).Activate
        celda_a_volver = Replace(ActiveCell.AddressLocal, "$", "")
        horas_dia = 0
        Do While ActiveCell.Value <> ""
          If ActiveCell.Interior.Color = 49407 Then
            If ActiveCell.Value <> "VACACIONES" Then
              pp = pp + ActiveCell.Value
            End If
          Else
            horas_dia = horas_dia + IIf(ActiveCell.Value = "VACACIONES", 8, ActiveCell.Value)
          End If
          ActiveCell.Offset(1, 0).Activate
        Loop
        Range(celda_a_volver).Activate
        If horas_dia > 8 Then
          mv = mv + horas_dia - 8
          horas_dia = 8
        End If
        HN = HN + horas_dia
      Next i
      Range(act_range).Activate
      ActiveCell.Offset(1, 0).Activate
      Exit Do
    Loop
    'MsgBox "HN=" & HN & " y " & "MV=" & MV & "PP=" & PP & " - ACT_CODE=" & act_code & " - N_SEMANA=" & n_semana
    Sheets(el_mes).Activate
    Range("A1").Select
    Do Until ActiveCell.Value = act_code Or la_linea = 50000
      ActiveCell.Offset(1, 0).Activate
      la_linea = la_linea + 1
    Loop
    ActiveCell.Offset(0, (3 * n_semana) - 1).Activate
    ActiveCell.Value = HN
    ActiveCell.Offset(0, 1).Activate
    ActiveCell.Value = mv
    ActiveCell.Offset(0, 1).Activate
    ActiveCell.Value = pp
  Loop
  Application.ScreenUpdating = True
  Sheets(el_mes).Activate
  Range("A1").Activate
End Sub

Sub datos()
  Dim h As Object
  
  Dim Semanas, s_meses As String
  Dim userf As UserForm1
  'userf = userfrom1
  For Each h In ThisWorkbook.Sheets
    If InStr(1, meses, UCase(h.Name)) > 0 Then
      s_meses = s_meses & h.Name & Space(10 - Len(h.Name))
    Else
      If Mid(UCase(h.Name), 1, 6) = "SEMANA" Then
        Semanas = Semanas & h.Name
      End If
    End If
  Next
  'MsgBox "Los Meses son: " & s_meses & " y las semans son: " & Semanas
  With UserForm1
    For i = 1 To Len(s_meses) / 10
      If i = 1 Then
        .cbo_mes.Value = Trim(Mid(s_meses, (i * 10) - 9, 10))
        el_mes = Trim(Mid(s_meses, (i * 10) - 9, 10))
      End If
      .cbo_mes.AddItem Trim(Mid(s_meses, (i * 10) - 9, 10))
    Next i
    For i = 1 To Len(Semanas) / 12
     If Mid(Mid(Semanas, (i * 12) - 11, 12), 8, 3) = Mid(el_mes, 1, 3) Then
      If i = 1 Then
        .Cbo_Semana.Value = Mid(Semanas, (i * 12) - 11, 12)
        la_semana = Mid(Semanas, (i * 12) - 11, 12)
      End If
      .Cbo_Semana.AddItem Mid(Semanas, (i * 12) - 11, 12)
     End If
    Next i
    .Txt_Las_semanas = Semanas
    .Show
  End With
End Sub
Function Nombre_Hoja() As String
  Dim n As String
   Nombre_Hoja = ActiveSheet.Name
End Function
Sub PRUEBA()
  MsgBox ActiveCell.Offset(0, 4).Value + ActiveCell.Offset(0, 8).Value + _
         ActiveCell.Offset(0, 12).Value + ActiveCell.Offset(0, 16).Value + _
         ActiveCell.Offset(0, 20).Value + ActiveCell.Offset(0, 24).Value + ActiveCell.Offset(0, 26).Value
End Sub
Sub Primer_dia_sem_mes()
 Dim el_l As Integer
 For i = 1 To 12
  'For j = 1 To 7
  ' If Weekday(DateSerial(2024, i, 1)) = j Then
  '  MsgBox "El mes " & _
  '   IIf(i = 1, "Enero", IIf(i = 2, "Febrero", IIf(i = 3, "Marzo", IIf(i = 4, "Abril", IIf(i = 5, "Mayo", IIf(i = 6, "Junio", _
  '   IIf(i = 7, "Julio", IIf(i = 8, "Agosto", IIf(i = 9, "Septiembre", IIf(i = 10, "Octubre", IIf(i = 11, "Noviembre", "Diciembre"))))))))))) & _
  '   " comienza un " & IIf(j = 1, "Domingo", IIf(j = 2, "Lunes", IIf(j = 3, "Martes", IIf(j = 4, "Miércoles", IIf(j = 5, "Jueves", IIf(j = 6, "Viernes", "Sábado"))))))
  ' End If
  'Next j
  For j = 1 To 5
    el_l = Val(Weekday(DateSerial(2024, i, (j * 7) - 6)))
    MsgBox "El lunes " & j & "ra semana de " & _
       IIf(i = 1, "Enero", IIf(i = 2, "Febrero", IIf(i = 3, "Marzo", IIf(i = 4, "Abril", IIf(i = 5, "Mayo", IIf(i = 6, "Junio", _
       IIf(i = 7, "Julio", IIf(i = 8, "Agosto", IIf(i = 9, "Septiembre", IIf(i = 10, "Octubre", IIf(i = 11, "Noviembre", "Diciembre"))))))))))) & _
       " es el " & Format(DateSerial(2024, i, (j * 7) - 6) - IIf(el_l = 1, 6, el_l - 2), "DD")
     Next j
 Next i
End Sub

Function Primer_dia_sem_mes_f(anho As Integer, mes As Integer, sem As Integer) As String
 Dim el_l As Integer

el_l = Val(Weekday(DateSerial(anho, mes, (sem * 7) - 6)))
    Primer_dia_sem_mes_f = "El lunes " & sem & "ra semana de " & _
       IIf(mes = 1, "Enero", IIf(mes = 2, "Febrero", IIf(mes = 3, "Marzo", IIf(mes = 4, "Abril", IIf(mes = 5, "Mayo", IIf(i = 6, "Junio", _
       IIf(mes = 7, "Julio", IIf(mes = 8, "Agosto", IIf(mes = 9, "Septiembre", IIf(mes = 10, "Octubre", IIf(mes = 11, "Noviembre", "Diciembre"))))))))))) & _
       " es el " & Format(DateSerial(anho, mes, (sem * 7) - 6) - IIf(el_l = 1, 6, el_l - 2), "DD")
End Function
Function Primer_dia_sem_mes_n(anho As Integer, mes As Integer, sem As Integer) As String
 Dim el_l As Integer

el_l = Val(Weekday(DateSerial(anho, mes, (sem * 7) - 6)))
Primer_dia_sem_mes_n = DateSerial(anho, mes, (sem * 7) - 6) - IIf(el_l = 1, 6, el_l - 2)
End Function

Function Mes_Sem_H(MoS As String) As Integer
Dim c_sem As String
 Dim n_mes, n_sem As Integer
 n_mes = 0
 c_sem = Nombre_Hoja()
 n_sem = Val(Mid(c_sem, 12, 1))
 If InStr(1, c_sem, "SEMANA") > 0 Then
   Select Case UCase(Mid(c_sem, 8, 3))
     Case "ENE"
       n_mes = 1
     Case "FEB"
       n_mes = 2
     Case "MAR"
       n_mes = 3
     Case "ABR"
       n_mes = 4
     Case "MAY"
       n_mes = 5
     Case "JUN"
       n_mes = 6
     Case "JUL"
       n_mes = 7
     Case "AGO"
       n_mes = 8
     Case "SEP"
       n_mes = 9
     Case "OCT"
       n_mes = 10
     Case "NOV"
       n_mes = 11
     Case "DIC"
       c_mes = 12
    End Select
 End If
 Mes_Sem_H = IIf(MoS = "M", n_mes, n_sem)
End Function

Function mapeo(buscar As Boolean) As String
  On Error GoTo NO_ENCONTRADO
  Dim minutos As Integer
  Dim horas As Double
  Dim t_fecha As String
  Dim datos_g, hoja_g As String
  Dim col_g As Integer
  Dim hoja_volver As String
  
 If buscar Then
  codmobi = ActiveCell.Value
  dni = Mid(Replace(ActiveCell.Offset(0, 3), "-", ""), 1, 9)
  nom_tra = ActiveCell.Offset(0, 1).Text
  cod_Nom = Application.VLookup(dni, ThisWorkbook.Sheets("NOMINA").Range("$B$2:$T$111"), 2, False)
 End If
  horas = Val(Format(ActiveCell.Offset(0, 5).Text, "HH"))
  minutos = Val(Right(Format(ActiveCell.Offset(0, 5).Value, "HH:MM"), 2))
  If minutos >= 30 Then
    horas = horas + (5 / 10)
  End If
  t_fecha = Mid(ActiveCell.Offset(0, 4).Text, 1, 10)
  't_fecha = Mid(t_fecha, 4, 3) & Mid(t_fecha, 1, 3) & Right(t_fecha, 4)
  datos_g = semana_grabar(t_fecha)
  hoja_g = Mid(datos_g, 1, 12)
  col_g = Mid(datos_g, 14, 2)
  mapeo = "MOBIBUK: " & codmobi & " - " & nom_tra & _
          " - " & IIf(IsNumeric(Mid(dni, 1, 1)), "DNI", "NIE") & ": " & _
          dni & " - FECHA: " & ActiveCell.Offset(0, 4).Text & " - HORAS: " & Format(horas, "#.0") & " - " & minutos & _
          datos_g & _
          " - Cod Nomina: "
  mapeo = mapeo & cod_Nom
  hoja_volver = ActiveSheet.Name
  Sheets(hoja_g).Activate
  Range("B2").Select
  Do While True
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.Value = cod_Nom Or ActiveCell.Row = 500 Then
      Exit Do
    End If
  Loop
  If ActiveCell.Value = cod_Nom Then
    'MsgBox "Aqui grabaria en col " & col_g & " las horas " & Format(horas, "#.0")
    ActiveCell.Offset(0, col_g).Value = Format(horas, "#.0")
  End If
  Sheets(hoja_volver).Activate
  Exit Function
NO_ENCONTRADO:
  mapeo = mapeo & "NO ENCONTRADO "
End Function

Sub Carga_Fichaje()
  Dim act_code As Variant
  Dim linea As Integer
  Dim sebusca As Boolean
  carga_mes
  Application.ScreenUpdating = False
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
    'Debug.Print linea & " " &
    mapeo (sebusca)
    linea = linea + 1
    If ActiveCell.Offset(0, 4).Value = "" Or ActiveCell.Row >= 2000 Then
      Exit Do
    End If
  Loop
  Application.ScreenUpdating = True
  Range("A1").Select
End Sub
Function semana_grabar(Fecha) As String
  Dim ndia, d1, d2, d3, d4, d5, d6, d_grabar, dia_fecha, ndia_g, dia_final As Integer
  Dim semana As String
  Dim fecha_1 As Date ', fecha
  Dim fcha_ing As Date
  fecha_ing = Mid(Fecha, 4, 3) & Mid(Fecha, 1, 3) & Right(Fecha, 4)
  'fecha = CVDate("01/09/2024")
  fecha_1 = CVDate("01/" & Format(Fecha, "MM/YYYY"))
  dia_fecha = Day(Fecha)
  ndia = Weekday(fecha_1)
  ndia_g = IIf(Weekday(Fecha) = 1, 8, Weekday(Fecha))
  dia_final = Day(CVDate(Application.EoMonth(IIf(dia_fecha > 12, fecha_ing, Fecha), 0)))
  semana = "SEMANA_" & Mid(mes_hoja, 1, 3) & "_"
  d1 = IIf(ndia = 1, 1, 1 + 8 - ndia)
  d2 = d1 + 7
  d3 = d2 + 7
  d4 = d3 + 7
  If d4 >= dia_final Then
    d4 = dia_final
  End If
  If d4 < dia_final Then
    d5 = d4 + 7
  End If
  If d5 >= dia_final Then
    d5 = dia_final
  End If
  If d5 < dia_final Then
    d6 = dia_final
  End If
  'd5 = Day(CVDate(Application.EoMonth(IIf(dia_fecha > 12, fecha_ing, fecha), 0)))
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
    Case d5 To d6
      semana = semana & "6"
  End Select
  d_grabar = (4 * ndia_g) - 4
  'MsgBox fecha & dia_fecha & " semana: " & semana & "," & ndia & "," & d_grabar
  semana_grabar = semana & "-" & d_grabar
End Function
Sub p2()
  MsgBox semana_grabar(CVDate("15/08/2024"))
End Sub

Sub p()
  carga_mes
  
  MsgBox Format((InStr(1, meses, mes_hoja) + 9) / 10, "00")
End Sub
Sub carga_mes()
 meses = "ENERO     FEBRERO   MARZO     ABRIL     MAYO      JUNIO     JULIO     AGOSTO    SEPTIEMBREOCTUBRE   NOVIEMBRE DICIEMBRE "
 For Each h In ThisWorkbook.Sheets
  If InStr(1, meses, UCase(h.Name)) > 0 Then
    mes_hoja = h.Name
    el_mes = h.Name
    Exit For
  End If
 Next
 If el_mes = "" Then
   el_mes = "MES"
 End If
 el_anho = Sheets(el_mes).Range("ANHO_LIBRO").Value
 n_semanas = 0
 For Each h In ThisWorkbook.Sheets
   If Mid(UCase(h.Name), 1, 6) = "SEMANA" Then
     n_semanas = n_semanas + 1
   End If
 Next
End Sub
Sub Pon_Semana()
  Dim ndia, d1, d2, d3, d4, d5, d6, dia_fecha, ndia_g, dia_final, linea As Integer
  Dim semana(6) As String
  Dim Fecha, fecha_1, fcha_ing As Date
  Dim letras As String
  carga_mes
  letras = "CFILOR"
  linea = 2
  Fecha = CVDate("01/" & Format((InStr(1, meses, mes_hoja) + 9) / 10, "00") & "/" & Sheets(el_mes).Range("ANHO_LIBRO").Value)
  fecha_ing = Mid(Fecha, 4, 3) & Mid(Fecha, 1, 3) & Right(Fecha, 4)
  fecha_1 = CVDate("01/" & Format(Fecha, "MM/YYYY"))
  dia_fecha = Day(Fecha)
  ndia = Weekday(fecha_1)
  ndia_g = IIf(Weekday(Fecha) = 1, 8, Weekday(Fecha))
  dia_final = Day(CVDate(Application.EoMonth(IIf(dia_fecha > 12, fecha_ing, Fecha), 0)))
  
  d1 = IIf(ndia = 1, 1, 1 + 8 - ndia)
  d2 = d1 + 7
  d3 = d2 + 7
  d4 = d3 + 7
  If d4 >= dia_final Then
    d4 = dia_final
  End If
  If d4 < dia_final Then
    d5 = d4 + 7
  End If
  If d5 >= dia_final Then
    d5 = dia_final
  End If
  If d5 < dia_final Then
    d6 = dia_final
  End If
  semana(1) = "DEL 1" & IIf(d1 <> 1, " AL " & d1, "")
  semana(2) = "DEL " & d1 + 1 & " AL " & d2
  semana(3) = "DEL " & d2 + 1 & " AL " & d3
  semana(4) = "DEL " & d3 + 1 & " AL " & d4
  semana(5) = "DEL " & d4 + 1 & IIf(d4 + 1 <> dia_final, " AL " & d5, "")
  semana(6) = IIf(d5 + 1 > dia_final, "", IIf(d5 + 1 < dia_final, " DEL " & d5 + 1 & " AL " & dia_final, "DEL " & dia_final))
'  MsgBox semana(1) & semana(2) & semana(3) & semana(4) & semana(5) & IIf(d5 + 1 <= dia_final, semana(6), "")
If mes_hoja <> "" Then
  Sheets(mes_hoja).Activate
 For j = 1 To 6
  Range(Range(Mid(letras, j, 1) & linea & ":U200").Find("SEMANA " & j).Address).Select
  For i = 1 To 4
    If ActiveCell.Value = "SEMANA " & j Then
      ActiveCell.Offset(1, 0).Select
      ActiveCell.Value = semana(j)
      If i < 4 Then
        Range(Range(Replace(ActiveCell.AddressLocal, "$", "") & ":U200").Find("SEMANA " & j).Address).Select
      End If
      linea = IIf(i = 4, 2, ActiveCell.Offset(1, 0).Row)
    End If
  Next i
 Next j
End If
End Sub
Sub Carga_Variables()
  Dim act_code, linea, linea_Mes As Integer
  Dim normal, mv, pp As Double
  Dim act_nombre As String
  If el_mes = "" Then
    carga_mes
  End If
  Application.ScreenUpdating = False
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
  Loop
  'COPIAR CODIGOS CON DATOS A VARIABLS PARA COMPARAR Y SABER QUIEN FALTA
  Range("H9").Select
  Sheets(el_mes).Activate
  Range("A65").Select
  linea = 65
  Do While linea < 144
    ActiveCell.Offset(1, 0).Select
    linea = ActiveCell.Row
    Do While IsNumeric(ActiveCell.Text)
      act_code = ActiveCell.Value
      act_nombre = ActiveCell.Offset(0, 1)
      linea = ActiveCell.Row
      If Range("X" & linea).Value <> 0 Or Range("Y" & linea).Value <> 0 Or Range("Z" & linea).Value <> 0 Then
        Sheets("VARIABLES").Activate
        ActiveCell.Value = act_code
        ActiveCell.Offset(0, 1).Value = act_nombre
        ActiveCell.Offset(1, 0).Select
        Sheets(el_mes).Activate
      End If
      ActiveCell.Offset(1, 0).Select
      linea = ActiveCell.Row
    Loop
  Loop
  Sheets("VARIABLES").Activate
  Resalta_Duplicados
  Application.ScreenUpdating = True
  Sheets(el_mes).Activate
End Sub
Sub Carga_Jornada()
  Dim ndia, d1, d2, d3, d4, d5, d6, dia_fecha, ndia_g, dia_final, linea As Integer
  Dim act_code, semanas_mes, dini As Integer
  Dim d(6) As Integer
  Dim Fecha, fecha_1, fcha_ing As Date
  Dim letras, msg As String
  Dim x(31) As String
  Application.ScreenUpdating = False
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
  'MsgBox msg
  Sheets("JORNADAS").Activate
  For j = 1 To 31
    ActiveCell.Offset(0, j + 1).Value = x(j)
  Next j
 Next i
 ActiveCell.Offset(1, 0).Select
 act_code = ActiveCell.Value
Loop
Application.ScreenUpdating = True
End Sub

Sub probar()
  Dim ndia, d1, d2, d3, d4, d5, d6, dia_fecha, ndia_g, dia_final, linea As Integer
  Dim semana(6) As String
  Dim Fecha, fecha_1, fcha_ing As Date
  Dim letras As String
  'carga_mes
  letras = "CFILOR"
  linea = 2
For m = 1 To 12
  Fecha = CVDate("12/" & Format(m, "00") & "/2024")
  fecha_ing = Mid(Fecha, 4, 3) & Mid(Fecha, 1, 3) & Right(Fecha, 4)
  fecha_1 = CVDate("01/" & Format(Fecha, "MM/YYYY"))
  dia_fecha = Day(Fecha)
  ndia = Weekday(fecha_1)
  ndia_g = IIf(Weekday(Fecha) = 1, 8, Weekday(Fecha))
  dia_final = Day(CVDate(Application.EoMonth(IIf(dia_fecha > 12, fecha_ing, Fecha), 0)))
  
  d1 = IIf(ndia = 1, 1, 1 + 8 - ndia)
  d2 = d1 + 7
  d3 = d2 + 7
  d4 = d3 + 7
  If d4 >= dia_final Then
    d4 = dia_final
  End If
  If d4 < dia_final Then
    d5 = d4 + 7
  End If
  If d5 >= dia_final Then
    d5 = dia_final
  End If
  If d5 < dia_final Then
    d6 = dia_final
  End If
  semana(1) = "DEL 1" & IIf(d1 <> 1, " AL " & d1, "")
  semana(2) = "DEL " & d1 + 1 & " AL " & d2
  semana(3) = "DEL " & d2 + 1 & " AL " & d3
  semana(4) = "DEL " & d3 + 1 & " AL " & d4
  semana(5) = "DEL " & d4 + 1 & IIf(d4 + 1 <> dia_final, " AL " & d5, "")
  semana(6) = IIf(d5 + 1 > dia_final, "", IIf(d5 + 1 < dia_final, " DEL " & d5 + 1 & " AL " & dia_final, "DEL " & dia_final))
  MsgBox semana(1) & " " & semana(2) & " " & semana(3) & " " & semana(4) & " " & semana(5) & " " & IIf(d5 + 1 <= dia_final, semana(6), "")
Next m
End Sub




