Sub Resalta_Duplicados()
If Range("B:B,H:H").FormatConditions.Count = 0 Then
  Range("B:B,H:H").Select
  Range("H2").Activate
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
End If
End Sub



Sub d()
  Dim el_Path As String
  el_Path = ActiveWorkbook.FullName
  el_Path = Mid(el_Path, 1, InStr(el_Path, "PARTES") - 1) & "NOMINA_MOBIBUK.xlsm"
  MsgBox el_Path
  
End Sub
Sub AgregarNoEncontradosVariables()
 Dim actcod, encontrados, no_encontrados, linea_insert, linea_ini, linea_fin As Integer
 Dim rg_back As String
 Dim val_encontrado As Variant
 Dim salir_cod_vacio, encontrado As Boolean
 Dim sh_volver As String

 Sheets("VARIABLES").Activate
 Range("B1").Select
 Selection.End(xlDown).Select
 Selection.End(xlDown).Select
 linea_fin = ActiveCell.Row
 ActiveCell.Offset(1, 0).Select
 linea_insert = ActiveCell.Row
    
 salir_cod_vacio = False
 linea_ini = 0
 Range("H1").Select
 actcod = ActiveCell.Value
 encontrados = 0
 no_encontrados = 0
 Do While True
   If ActiveCell.Value <> 0 Then
     salir_cod_vacio = True
     linea_ini = IIf(linea_ini = 0, ActiveCell.Row, linea_ini)
     If ActiveCell.Value <> actcod Then
       actcod = ActiveCell.Value
       rg_back = ActiveCell.Address
       Set val_encontrado = Range("B:B").Find(actcod, LookAt:=xlWhole)
       If Not val_encontrado Is Nothing Then
         'MsgBox "ENCONTRADO"
         encontrado = True
         encontrados = encontrados + 1
       Else
         'MsgBox "NO ENCONTRADO"
         encontrado = False
         no_encontrados = no_encontrados + 1
       End If
       Range(rg_back).Select
       If Not encontrado Then
         Range("B" & linea_insert).Value = actcod
         Range("B" & linea_insert).Offset(0, 1).Value = Range(rg_back).Offset(0, 1).Value
         Range("B" & linea_insert).Offset(0, 1).Font.Color = Range(rg_back).Offset(0, 1).Font.Color
         Range("B" & linea_insert).Offset(0, 1).Select
         With Selection.Interior
           .Pattern = xlSolid
           .PatternColorIndex = xlAutomatic
           .Color = 65535
           .TintAndShade = 0
           .PatternTintAndShade = 0
         End With
         Range(rg_back).Select
         linea_insert = linea_insert + 1
       End If
     End If
   Else
     If salir_cod_vacio Then
       Exit Do
     End If
   End If
   ActiveCell.Offset(1, 0).Select
   If ActiveCell.Row > 100 Then
     Exit Do
   End If
 Loop
 linea_fin = linea_insert
 'MsgBox "Encontrados: " & encontrados & " - No Encontrados: " & no_encontrados & " - Total: " & encontrados + no_encontrados
 If el_mes = "" Then
   carga_mes
 End If
sh_volver = ActiveSheet.Name
With UserForm5
'  .Show
End With
Sheets(sh_volver).Activate
Range("D" & linea_fin + 2).Offset(0, -1).Value = "ALMACEN: "
Range("D" & linea_fin + 2).Formula2 = "=SUMA_POR_COLOR(" & Chr(34) & "D" & linea_ini & ":D" & linea_fin & Chr(34) & ", 13819130, -1)": Range("D" & linea_fin + 2).Interior.Color = 13819130
Range("D" & linea_fin + 3).Offset(0, -1).Value = "LA TORRE: "
Range("D" & linea_fin + 3).Formula2 = "=SUMA_POR_COLOR(" & Chr(34) & "D" & linea_ini & ":D" & linea_fin & Chr(34) & ", 13826780, -1)"
Range("D" & linea_fin + 4).Offset(0, -1).Value = "INV.2.1: "
Range("D" & linea_fin + 4).Formula2 = "=SUMA_POR_COLOR(" & Chr(34) & "D" & linea_ini & ":D" & linea_fin & Chr(34) & ", 16440530, -1)"
Range("D" & linea_fin + 5).Offset(0, -1).Value = "INV.3.1: "
Range("D" & linea_fin + 5).Formula2 = "=SUMA_POR_COLOR(" & Chr(34) & "D" & linea_ini & ":D" & linea_fin & Chr(34) & ", 13172735, -1)"
Range("D" & linea_fin + 6).Offset(0, -1).Value = "GOBERNADORA: "
Range("D" & linea_fin + 6).Formula2 = "=D" & linea_fin + 4 & " + D" & linea_fin + 5
Range("D" & linea_fin + 7).Offset(0, -1).Value = "ENCARGADOS: "
Range("D" & linea_fin + 7).Formula2 = "=SUMA_POR_COLOR(" & Chr(34) & "D" & linea_ini & ":D" & linea_fin & Chr(34) & ", 15790320, -1)"
Range("D" & linea_fin + 8).Offset(0, -1).Value = "TOTAL GENERAL: "
Range("D" & linea_fin + 8).Formula2 = "=D" & linea_fin + 2 & " + D" & linea_fin + 3 & " + D" & linea_fin + 4 & " + D" & linea_fin + 5 & " + D" & linea_fin + 7

Range("E" & linea_fin + 2).Formula2 = "=SUMA_POR_COLOR(" & Chr(34) & "E" & linea_ini & ":E" & linea_fin & Chr(34) & ", 13819130, -2)"
Range("E" & linea_fin + 3).Formula2 = "=SUMA_POR_COLOR(" & Chr(34) & "E" & linea_ini & ":E" & linea_fin & Chr(34) & ", 13826780, -2)"
Range("E" & linea_fin + 4).Formula2 = "=SUMA_POR_COLOR(" & Chr(34) & "E" & linea_ini & ":E" & linea_fin & Chr(34) & ", 16440530, -2)"
Range("E" & linea_fin + 5).Formula2 = "=SUMA_POR_COLOR(" & Chr(34) & "E" & linea_ini & ":E" & linea_fin & Chr(34) & ", 13172735, -2)"
Range("E" & linea_fin + 6).Formula2 = "=E" & linea_fin + 4 & " + E" & linea_fin + 5
Range("E" & linea_fin + 7).Formula2 = "=SUMA_POR_COLOR(" & Chr(34) & "E" & linea_ini & ":E" & linea_fin & Chr(34) & ", 15790320, -2)"
Range("E" & linea_fin + 8).Formula2 = "=E" & linea_fin + 2 & " + E" & linea_fin + 3 & " + E" & linea_fin + 4 & " + E" & linea_fin + 5 & " + E" & linea_fin + 7

Range("F" & linea_fin + 2).Formula2 = "=SUMA_POR_COLOR(" & Chr(34) & "F" & linea_ini & ":E" & linea_fin & Chr(34) & ", 13819130, -3)"
Range("F" & linea_fin + 3).Formula2 = "=SUMA_POR_COLOR(" & Chr(34) & "F" & linea_ini & ":E" & linea_fin & Chr(34) & ", 13826780, -3)"
Range("F" & linea_fin + 4).Formula2 = "=SUMA_POR_COLOR(" & Chr(34) & "F" & linea_ini & ":E" & linea_fin & Chr(34) & ", 16440530, -3)"
Range("F" & linea_fin + 5).Formula2 = "=SUMA_POR_COLOR(" & Chr(34) & "F" & linea_ini & ":E" & linea_fin & Chr(34) & ", 13172735, -3)"
Range("F" & linea_fin + 6).Formula2 = "=F" & linea_fin + 4 & " + F" & linea_fin + 5
Range("F" & linea_fin + 7).Formula2 = "=SUMA_POR_COLOR(" & Chr(34) & "F" & linea_ini & ":E" & linea_fin & Chr(34) & ", 15790320, -3)"
Range("F" & linea_fin + 8).Formula2 = "=F" & linea_fin + 2 & " + F" & linea_fin + 3 & " + F" & linea_fin + 4 & " + F" & linea_fin + 5 & " + F" & linea_fin + 7

Range("C" & linea_fin + 2 & ":F" & linea_fin + 2).Interior.Color = 13819130
Range("C" & linea_fin + 3 & ":F" & linea_fin + 3).Interior.Color = 13826780
Range("C" & linea_fin + 4 & ":F" & linea_fin + 4).Interior.Color = 16440530
Range("C" & linea_fin + 5 & ":F" & linea_fin + 5).Interior.Color = 13172735
Range("C" & linea_fin + 7 & ":F" & linea_fin + 7).Interior.Color = 15790320

Range("C" & linea_fin + 2 & ":C" & linea_fin + 8).Select
With Selection
  .HorizontalAlignment = xlRight
  .VerticalAlignment = xlBottom
  .WrapText = False
  .Orientation = 0
  .AddIndent = False
  .IndentLevel = 0
  .ShrinkToFit = False
  .ReadingOrder = xlContext
  .MergeCells = False
End With
Range("C" & linea_fin + 1).Select
End Sub

Function SUMA_POR_COLOR(rango As String, ColorCell As Double, colizq As Integer) As Double
 Dim ACell As Range
 For Each ACell In Range(rango)
   If ACell.Offset(0, colizq).Interior.Color = ColorCell Then
     SUMA_POR_COLOR = SUMA_POR_COLOR + ACell.Value
   End If
 Next ACell
End Function
Sub ActualizarFactores()
 Dim linea, l_back, codemp, cuantos, proc As Integer
 Dim factorN, factorMV, factorPP As Double
 Dim Formula_Calcul_Mto(3) As String
 If el_mes = "" Then
   carga_mes
 End If
 'voy a probar con JORNADAS
 'Sheets("JORNADAS").Activate
 'cuantos = WorksheetFunction.Count(ActiveSheet.Range("B1:B2000"))
 'Range("B1").Select
 
 Sheets(el_mes).Activate
 cuantos = WorksheetFunction.Count(ActiveSheet.Range("A1:A2000"))
 If cuantos = 0 Then
  MsgBox "No hay empleados aún en la página " & el_mes, vbOKOnly, "Información..."
 Else
  Range("A1").Select
  proc = 0
  Do While True
    If IsNumeric(ActiveCell.Text) Or ActiveCell.Row > 500 Then
      Exit Do
    End If
    ActiveCell.Offset(1, 0).Select
  Loop
  Do While True
   If IsNumeric(ActiveCell.Text) Then
       linea = ActiveCell.Row
      codemp = Range("NOMINA_1").Find(ActiveCell.Value).Offset(0, 3).Value
        proc = proc + 1
     factorN = Range("FACTOR_HORAS").Find(codemp).Offset(0, 1).Value
    factorMV = Range("FACTOR_HORAS").Find(codemp).Offset(0, 2).Value
    factorPP = Range("FACTOR_HORAS").Find(codemp).Offset(0, 3).Value
    Formula_Calcul_Mto(1) = "=U" & linea & "*" & Replace(Format(factorN, "0.00"), ",", ".")
    Formula_Calcul_Mto(2) = "=V" & linea & "*" & Replace(Format(factorMV, "0.00"), ",", ".")
    Formula_Calcul_Mto(3) = "=W" & linea & "*" & Replace(Format(factorPP, "0.00"), ",", ".")
   'MsgBox codemp & " - " & proc & " - " & Formula_Calcul_Mto(1) & " - " & Formula_Calcul_Mto(2) & " - " & Formula_Calcul_Mto(3)
    ActiveCell.Offset(0, 23).Formula2 = Formula_Calcul_Mto(1)
    ActiveCell.Offset(0, 24).Formula2 = Formula_Calcul_Mto(2)
    ActiveCell.Offset(0, 25).Formula2 = Formula_Calcul_Mto(3)
   End If
   If proc > cuantos Or ActiveCell.Row > 500 Then
    Exit Do
   End If
   ActiveCell.Offset(1, 0).Select
   l_back = ActiveCell.Row + 2
  Loop
  Range("A" & l_back).Select
 End If
End Sub
Sub hs()
  MsgBox SUMA_POR_COLOR("D9:D56", 65535, -1)
  MsgBox SUMA_POR_COLOR("E9:E56", 65535, -2)
  MsgBox SUMA_POR_COLOR("F9:F56", 65535, -3)
  MsgBox SUMA_POR_COLOR("D9:D56", 16777215, -1)
  MsgBox SUMA_POR_COLOR("E9:E56", 16777215, -2)
  MsgBox SUMA_POR_COLOR("F9:F56", 16777215, -3)
End Sub

Sub LosQueTrabajaron(queHoja As String)
 Dim cuantos As Double
 Dim en, codemp, lin_volver, lin As Integer
 Dim rosa, verdito, azulito, blanco, amarilli, gris, fiesta, finde, laboral, el_color As Double
 Dim losColores(4) As Double
 Dim nomEmp, celda_back, rng_Sel, lin_Char, actHoja As String
 Dim Festivo, Fecha As Variant
 Dim saltar As Boolean
 azulito = 16440530
 verdito = 13826780
    rosa = 13819130
  blanco = 16777215
amarilli = 13172735
    gris = 15790320
  fiesta = 192
   finde = 8355711
 laboral = 11272191
 losColores(1) = rosa
 losColores(2) = verdito
 losColores(3) = azulito
 losColores(4) = amarilli
 saltar = True
 en = 1
 cuantos = 0
 Sheets(el_mes).Activate
 cuantos = WorksheetFunction.Count(Range("A1:A500"))
 'MsgBox cuantos
 If cuantos > 0 Then
  For j = 1 To 2
   Range("A1").Select
   For i = 1 To 4
    If i = 4 Then
     saltar = False
    End If
    Do While True
     If IsNumeric(ActiveCell.Text) Or ActiveCell.Row > 500 Then
       Exit Do
     End If
     ActiveCell.Offset(1, 0).Select
    Loop
    Do While IsNumeric(ActiveCell.Text)
If (ActiveCell.Offset(0, 20).Value + ActiveCell.Offset(0, 21).Value + ActiveCell.Offset(0, 22).Value) > 0 Then
     codemp = ActiveCell.Value
     nomEmp = ActiveCell.Offset(0, 1).Value
      If queHoja = "VARIABLES" Then
        Poner_Bordes
      End If
'      Range(celda_back).Select
'      Sheets(actHoja).Activate
     End If
      If j = 1 And ActiveCell.Offset(0, 1).Font.Color <> 255 Then
       If (ActiveCell.Offset(0, 20).Value + ActiveCell.Offset(0, 21).Value + ActiveCell.Offset(0, 22).Value) > 0 Then
        'MsgBox ActiveCell.Value & " - " & IIf(i = 1, "ALMACEN", IIf(i = 2, "LA TORRE", IIf(i = 3, "INV.1", "INV.2"))) & " " & ActiveCell.Offset(0, 1).Value & " y NumHoras = " & ActiveCell.Offset(0, 20).Value + ActiveCell.Offset(0, 21).Value + ActiveCell.Offset(0, 22).Value
        If queHoja = "JORNADAS" Then
          inserta_linea_jornada
        End If
        Sheets(queHoja).Activate
        ActiveCell.Value = codemp
        ActiveCell.Offset(0, 1).Value = nomEmp
         ActiveCell.Offset(0, 1).Interior.Color = losColores(i)
        ActiveCell.Offset(1, 0).Select
        Sheets(el_mes).Activate
       End If
      ElseIf j = 2 And ActiveCell.Offset(0, 1).Font.Color = 255 Then
       If (ActiveCell.Offset(0, 20).Value + ActiveCell.Offset(0, 21).Value + ActiveCell.Offset(0, 22).Value) > 0 Then
        'MsgBox ActiveCell.Value & " - " & IIf(i = 1, "ALMACEN", IIf(i = 2, "LA TORRE", IIf(i = 3, "INV.1", "INV.2"))) & " " & ActiveCell.Offset(0, 1).Value & " y NumHoras = " & ActiveCell.Offset(0, 20).Value + ActiveCell.Offset(0, 21).Value + ActiveCell.Offset(0, 22).Value
        If queHoja = "JORNADAS" Then
          inserta_linea_jornada
        End If
        Sheets(queHoja).Activate
        ActiveCell.Value = codemp
        ActiveCell.Offset(0, 1).Value = nomEmp
        ActiveCell.Offset(0, 1).Font.Color = 255
        ActiveCell.Offset(0, 1).Interior.Color = gris
        ActiveCell.Offset(1, 0).Select
        Sheets(el_mes).Activate
       End If
      End If

      ActiveCell.Offset(1, 0).Select
    Loop
   Next i
  Next j
 End If
End Sub

Sub inserta_linea_jornada()
 Dim cuantos As Double
 Dim en, codemp, lin_volver, lin As Integer
 Dim rosa, verdito, azulito, blanco, amarilli, gris, fiesta, finde, laboral, el_color As Double
 Dim losColores(4) As Double
 Dim nomEmp, celda_back, rng_Sel, lin_Char, actHoja As String
 Dim Festivo, Fecha As Variant
 Dim saltar As Boolean
 azulito = 16440530
 verdito = 13826780
    rosa = 13819130
  blanco = 16777215
amarilli = 13172735
    gris = 15790320
  fiesta = 192
   finde = 8355711
 laboral = 11272191
 losColores(1) = rosa
 losColores(2) = verdito
 losColores(3) = azulito
 losColores(4) = amarilli
 saltar = True
      actHoja = ActiveSheet.Name
      Sheets("JORNADAS").Activate
      Range("Total_Jornadas").Select
      lin = ActiveCell.Row
      celda_back = Replace(ActiveCell.Address, "$", "")
      lin_Char = lin & ":" & lin
      Rows(lin_Char).Select
      Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
      Range(celda_back).Select
     'Pon los colores Laboral-Amarillo, Finde-Gris, Festivo-Rojo
      For k = 1 To Day(WorksheetFunction.EoMonth(CVDate("01/" & el_mes & "/" & el_anho), 0))
        el_color = laboral
        Fecha = CVDate(Format(k, "00") & "/" & Format(el_mes, "00") & "/" & el_anho)
        Set Festivo = Range("FESTIVOS").Find(Fecha)
        If Not Festivo Is Nothing Then
         el_color = fiesta
        ElseIf Weekday(Fecha) = 7 Or Weekday(Fecha) = 1 Then
           el_color = finde
        End If
        ActiveCell.Offset(0, k + 1).Interior.Color = el_color
      Next k
    Range("B" & lin & ":C" & lin).Select
    Selection.Font.Bold = False
    Range("C" & lin).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
      
      
      Range(celda_back).Select
      Sheets(actHoja).Activate
End Sub
