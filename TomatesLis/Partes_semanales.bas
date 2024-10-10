REM Modulo1

Global c_nom(4) As String
Global n_veces(4)  As Integer
Global el_mes As String
Global la_semana As String
Global n_semana As Integer
Global accion As String
Global siBuscas As Boolean
Global codmobi As Integer
Global dni As String
Global nom_tra As String
Global cod_nom As Integer
Global meses As String
Global mes_hoja As String

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
Sub Carga_Codigo()
Dim c_nomb As String
Dim n_pos As Integer
Dim n_lin As Integer
Dim range_back_mes As String
siBuscas = False
Application.ScreenUpdating = False
'el_mes = "AGOSTO"
'la_semana = "SEMANA_3"
accion = "Procesar"
datos
Sheets(el_mes).Activate
range_back_mes = Replace(ActiveCell.AddressLocal, "$", "")
If accion = "Cancelar" Then
  Exit Sub
End If
n_semana = Mid(la_semana, Len(la_semana), 1)
Inserta_Col_Codigo
n_lin = 1
Sheets(la_semana).Activate
Range("C2").Activate
ActiveCell.Offset(1, 0).Activate
Do While n_lin < 20000
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
n_lin = n_lin + 1
ActiveCell.Offset(1, 1).Activate
If ActiveCell.Text = "" Then
  Do While ActiveCell.Text = "" And n_lin < 20000
    ActiveCell.Offset(1, 0).Activate
    If ActiveCell.Column = 1 Then
      ActiveCell.Offset(0, 1).Activate
    End If
    n_lin = n_lin + 1
  Loop
End If
Loop
Marca_Duplicados n_lin
Sheets(la_semana).Activate
Range("C2").Activate
Obten_horas
Sheets(el_mes).Activate
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
el_mes = "AGOSTO"
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
Sub Marca_Duplicados(Cant_lin As Integer)

    Range("B4:B" & Cant_lin).Select
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
  Dim HN, MV, PP, horas_dia As Double
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
    MV = 0
    PP = 0
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
              PP = PP + ActiveCell.Value
            End If
          Else
            horas_dia = horas_dia + IIf(ActiveCell.Value = "VACACIONES", 8, ActiveCell.Value)
          End If
          ActiveCell.Offset(1, 0).Activate
        Loop
        Range(celda_a_volver).Activate
        If horas_dia > 8 Then
          MV = MV + horas_dia - 8
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
    ActiveCell.Value = MV
    ActiveCell.Offset(0, 1).Activate
    ActiveCell.Value = PP
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
        .Cbo_Mes.Value = Trim(Mid(s_meses, (i * 10) - 9, 10))
        el_mes = Trim(Mid(s_meses, (i * 10) - 9, 10))
      End If
      .Cbo_Mes.AddItem Trim(Mid(s_meses, (i * 10) - 9, 10))
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
  cod_nom = Application.VLookup(dni, ThisWorkbook.Sheets("NOMINA").Range("$B$2:$T$111"), 2, False)
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
  mapeo = mapeo & cod_nom
  hoja_volver = ActiveSheet.Name
  Sheets(hoja_g).Activate
  Range("B2").Select
  Do While True
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.Value = cod_nom Or ActiveCell.Row = 500 Then
      Exit Do
    End If
  Loop
  If ActiveCell.Value = cod_nom Then
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
  
  meses = "ENERO     FEBRERO   MARZO     ABRIL     MAYO      JUNIO     JULIO     AGOSTO    SEPTIEMBREOCTUBRE   NOVIEMBRE DICIEMBRE "
  mes_hoja = "AGOSTO"
  
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
  Range("A1").Select
End Sub
Function semana_grabar(fecha) As String
  Dim ndia, d1, d2, d3, d4, d5, d_grabar, dia_fecha, ndia_g As Integer
  Dim semana As String
  Dim fecha_1 As Date ', fecha
  Dim fcha_ing As Date
  fecha_ing = Mid(fecha, 4, 3) & Mid(fecha, 1, 3) & Right(fecha, 4)
  'fecha = CVDate("01/09/2024")
  fecha_1 = CVDate("01/" & Format(fecha, "MM/YYYY"))
  dia_fecha = Day(fecha)
  ndia = Weekday(fecha_1)
  ndia_g = IIf(Weekday(fecha) = 1, 8, Weekday(fecha))
  semana = "SEMANA_" & Mid(mes_hoja, 1, 3) & "_"
  d1 = IIf(ndia = 1, 1, 1 + 8 - ndia)
  d2 = d1 + 7
  d3 = d2 + 7
  d4 = d3 + 7
  d5 = Day(CVDate(Application.EoMonth(IIf(dia_fecha > 12, fecha_ing, fecha), 0)))
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

Sub p()
 Dim h(5) As Integer
 For i = 1 To 5
   h(i) = 10 * i
 Next i
 For i = 1 To 5
 Debug.Print h(i)
 Next i
End Sub

REM ThisWorkbook


Private Sub Workbook_Open()
 siBuscas = True
 meses = "ENERO     FEBRERO   MARZO     ABRIL     MAYO      JUNIO     JULIO     AGOSTO    SEPTIEMBREOCTUBRE   NOVIEMBRE DICIEMBRE "
 For Each h In ThisWorkbook.Sheets
  If InStr(1, meses, UCase(h.Name)) > 0 Then
    mes_hoja = h.Name
    Exit For
  End If
 Next

 With UserForm2
   .Show
 End With

 UserForm2.Enabled = False
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
 On Error GoTo salirbusqueda

 Dim num_semana, cod_buscar As Integer
 Dim los_meses, x_sem As String
 los_meses = "ENERO     FEBRERO   MARZO     ABRIL     MAYO      JUNIO     JULIO     AGOSTO    SEPTIEMBREOCTUBRE   NOVIEMBRE DICIEMBRE "
 If InStr(1, los_meses, Mid(ActiveSheet.Name, 1, 10)) > 0 Then
   If Mid(Replace(ActiveCell.AddressLocal, "$", ""), 1, 1) = "B" And ActiveCell.Value <> "" Then
    cod_buscar = ActiveCell.Offset(0, -1).Value
    num_semana = InputBox("Número de la semana de " & ActiveSheet.Name, "Ingresa el dato, por favor", 1)
    x_sem = "SEMANA_" & Mid(ActiveSheet.Name, 1, 3) & "_" & Format(num_semana, "0")
    If num_semana > 0 And num_semana < 6 Then
      Sheets(x_sem).Activate
      Range(Range("B6:B2000").Find(cod_buscar).Address).Select
    End If
  Else
    Exit Sub
  End If
Else
  Exit Sub
End If
  Exit Sub

salirbusqueda:
  MsgBox "No se consigue"
  siBuscas = InputBox("Quieres seguri buscando", "confirma", "no") = "si"
End Sub

REM UserForm1


Private Sub Btn_Cancelar_Click()
  accion = "Cancelar"
  siBuscas = True
  Unload Me
End Sub

Private Sub btn_salir_Click()
  'MsgBox "El mes es: " & el_mes & " y la semana es: " & la_semana
  siBuscas = False
  Unload Me
End Sub

Private Sub Cbo_Mes_Change()
  Dim una_vez As Boolean
  una_vez = True
  el_mes = Trim(Me.Cbo_Mes.Value)
  Me.Cbo_Semana.Clear
  For i = 1 To Len(Me.Txt_Las_semanas) / 12
   If Mid(Mid(Me.Txt_Las_semanas, (i * 12) - 11, 12), 8, 3) = Mid(el_mes, 1, 3) Then
    If una_vez Then
      Me.Cbo_Semana.Value = Mid(Me.Txt_Las_semanas, (i * 12) - 11, 12)
      la_semana = Mid(Me.Txt_Las_semanas, (i * 12) - 11, 12)
      una_vez = False
    End If
    Me.Cbo_Semana.AddItem Mid(Me.Txt_Las_semanas, (i * 12) - 11, 12)
   End If
  Next i
  
End Sub

Private Sub Cbo_Semana_Change()
  la_semana = Me.Cbo_Semana.Value
End Sub

REM UserForm2

Private Sub UserForm_Activate()
  ActiveWorkbook.Queries("Fichaje").Refresh
  Me.Image1.Visible = True
  Me.Label2.ForeColor = RGB(0, 255, 0)
  ActiveWorkbook.Queries("MOBIBUK").Refresh
  Me.Image2.Visible = True
  Me.Label3.ForeColor = RGB(0, 255, 0)
  ActiveWorkbook.Queries("NOMINA").Refresh
  Me.Image3.Visible = True
  Me.Label4.ForeColor = RGB(0, 255, 0)
  UserForm2.Hide
End Sub
