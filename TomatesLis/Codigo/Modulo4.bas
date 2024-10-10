Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("NOMINA").Select
    Range("D3").Select
    ActiveWorkbook.Worksheets("NOMINA").ListObjects("NOMINA_1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("NOMINA").ListObjects("NOMINA_1").Sort.SortFields. _
        Add2 Key:=Range("NOMINA_1[FINCA]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("NOMINA").ListObjects("NOMINA_1").Sort.SortFields. _
        Add2 Key:=Range("NOMINA_1[NOMBRE Y APELLIDOS]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("NOMINA").ListObjects("NOMINA_1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.ListObjects("NOMINA").Range.AutoFilter Field:=10, Criteria1:= _
        "<>"
End Sub

Sub Ciclo_Agrega_Empleados()
 Dim cuantos, codNom, TipoContrato As Integer
 Dim factorN, factorMV, factorPP As Double
 Dim actFinca As String
 Dim apeNom As String
 Dim celdaback As String
 Dim fincas(5) As String
 Dim salir As Boolean
 salir = False
 fincas(1) = "ALMACEN"
 fincas(2) = "TORRE"
 fincas(3) = "GOBERNADORA FASE I"
 fincas(4) = "GOBERNADORA FASE II"
 fincas(5) = "OFICINA"
 actFinca = "OFICINA"
 Sheets("NOMINA").Activate
For i = 1 To 4
 actFinca = fincas(i)
 Range("E1").Select
 Do While True
   If ActiveCell.Value = actFinca Then
     Do While ActiveCell.Value = actFinca
       If ActiveCell.Offset(0, 5).Value = "ACTIVO" Then
         codNom = ActiveCell.Offset(0, -2).Value
         apeNom = ActiveCell.Offset(0, -1).Value
         TipoContrato = ActiveCell.Offset(0, 1).Value
         factorN = Range("FACTOR_HORAS").Find(TipoContrato).Offset(0, 1).Value
        factorMV = Range("FACTOR_HORAS").Find(TipoContrato).Offset(0, 2).Value
        factorPP = Range("FACTOR_HORAS").Find(TipoContrato).Offset(0, 3).Value
        'ubica la celda de Totales_(FINCA) en la hoja mes
        'inserta la linea
        'MsgBox actFinca & " - " & codNom & " - " & apeNom & " - " & TipoContrato & " - " & factorN & " - " & factorMV & " - " & factorPP
        celdaback = Replace(ActiveCell.Address, "$", "")
        Inserta_Empleado TipoContrato, actFinca, Val(codNom), apeNom, Val(Replace(factorN, ",", ".")), Val(Replace(factorMV, ",", ".")), Val(Replace(factorPP, ",", "."))
        Range(celdaback).Select
       End If
       Sheets("NOMINA").Activate
       ActiveCell.Offset(1, 0).Select
     Loop
   End If
   If ActiveCell.Value = "" Or ActiveCell.Row = 2000 Then
     Exit Do
   Else
     ActiveCell.Offset(1, 0).Select
   End If
 Loop
Next i
End Sub
Sub Inserta_Empleado(TipoCont As Integer, laFinca As String, cod_Nom As Integer, Ape_Noms As String, Factor_N As Double, Factor_MV As Double, Factor_PP As Double)
  Dim rosado, verde, azul, blanco, linea As Long
  Dim Formula_BRT, n_format1, n_format2 As String
  Dim Formula_SUMA_HORAS(3) As String
  Dim Formula_Calcul_Mto(3) As String
  Dim celdaT_Finca As String
  Dim Colores(3) As Double
  n_format1 = "0.0"
 Select Case laFinca
  Case "ALMACEN"
   celdaT_Finca = "Totales_Almacen"
  Case "TORRE"
   celdaT_Finca = "Totales_La_Torre"
  Case "GOBERNADORA FASE I"
   celdaT_Finca = "Totales_GOB_I"
  Case "GOBERNADORA FASE II"
   celdaT_Finca = "Totales_GOB_II"
 End Select
  Inserta_Linea celdaT_Finca
  linea = ActiveCell.Row
  n_format2 = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & " - " & Chr(34) & "??_);_(@_)"
  Formula_SUMA_HORAS(1) = "=C" & linea & "+F" & linea & "+I" & linea & "+L" & linea & "+O" & linea & "+R" & linea
  Formula_SUMA_HORAS(2) = "=D" & linea & "+G" & linea & "+J" & linea & "+M" & linea & "+P" & linea & "+S" & linea
  Formula_SUMA_HORAS(3) = "=E" & linea & "+H" & linea & "+K" & linea & "+N" & linea & "+Q" & linea & "+T" & linea
  Formula_Calcul_Mto(1) = "=U" & linea & "*" & Replace(Format(Factor_N, "0.00"), ",", ".") '"8,9"
  Formula_Calcul_Mto(2) = "=V" & linea & "*" & Replace(Format(Factor_MV, "0.00"), ",", ".") '"8,9"
  Formula_Calcul_Mto(3) = "=W" & linea & "*" & Replace(Format(Factor_PP, "0.00"), ",", ".") '"11,12"
  Formula_BRT = "=X" & linea & "+Z" & linea
      rosado = 14017275
       verde = 13693658
        azul = 16312794
      blanco = 16777215
  Colores(1) = rosado
  Colores(2) = verde
  Colores(3) = azul
  'ActiveCell.Interior.Color = rosado
  'ActiveCell.NumberFormat = n_format2
  'ActiveCell.Formula2 = formula_N
  ActiveCell.Value = cod_Nom
  ActiveCell.Offset(0, 1).Value = Ape_Noms
  If Val(TipoCont) <> 300 Then
    ActiveCell.Offset(0, 1).Font.Color = 255
  Else
    ActiveCell.Offset(0, 1).Font.Color = 0
  End If
  For i = 1 To 22 Step 3
   For j = 1 To 3
    'MsgBox i + j
    ActiveCell.Offset(0, i + j).Interior.Color = Colores(j)
    If i + j > 19 Then
     ActiveCell.Offset(0, i + j).NumberFormat = n_format2
     If i + j < 23 Then
      'MsgBox "pone formula 1 - " & j
      ActiveCell.Offset(0, i + j).Formula2 = Formula_SUMA_HORAS(j)
     Else
      'MsgBox "pone formula 2 - " & j
      ActiveCell.Offset(0, i + j).Formula2 = Formula_Calcul_Mto(j)
     End If
    Else
     ActiveCell.Offset(0, i + j).NumberFormat = n_format1
    End If
   Next j
  Next i
  'MsgBox "pone la última formula"
  ActiveCell.Offset(0, i + 1).Interior.Color = blanco
  ActiveCell.Offset(0, i + 1).Formula2 = Formula_BRT
  Actualiza_formulas celdaT_Finca
  'Range("Totales_Almacen").Select
End Sub

Sub Inserta_Linea(CeldaTotalesFinca As String)
 Dim lin As Double
 Dim lin_Char, celda_back, rng_Sel  As String
 Sheets("MES").Activate
 Range(CeldaTotalesFinca).Select
 If ActiveCell.Offset(-1, 0).Value = 0 And ActiveCell.Offset(-1, 0).Text = "" Then
  ActiveCell.Offset(-1, 0).Select
 Else
  lin = ActiveCell.Row
  celda_back = Replace(ActiveCell.Address, "$", "")
  rng_Sel = celda_back & ":" & Replace(ActiveCell.Offset(0, 1).Address, "$", "")
  lin_Char = lin & ":" & lin
  Rows(lin_Char).Select
  Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
  Range(celda_back).Select
  Range(rng_Sel).Select
  Poner_Bordes
  Range(celda_back).Select
 End If
End Sub
Sub Poner_Bordes()
  Selection.Borders(xlDiagonalDown).LineStyle = xlNone
  Selection.Borders(xlDiagonalUp).LineStyle = xlNone
  With Selection.Borders(xlEdgeLeft)
   .LineStyle = xlContinuous
   .ColorIndex = 0
   .TintAndShade = 0
   .Weight = xlThin
  End With
  With Selection.Borders(xlEdgeTop)
   .LineStyle = xlContinuous
   .ColorIndex = 0
   .TintAndShade = 0
   .Weight = xlThin
  End With
  With Selection.Borders(xlEdgeBottom)
   .LineStyle = xlContinuous
   .ColorIndex = 0
   .TintAndShade = 0
   .Weight = xlThin
  End With
  With Selection.Borders(xlEdgeRight)
   .LineStyle = xlContinuous
   .ColorIndex = 0
   .TintAndShade = 0
   .Weight = xlThin
  End With
  With Selection.Borders(xlInsideVertical)
   .LineStyle = xlContinuous
   .ColorIndex = 0
   .TintAndShade = 0
   .Weight = xlThin
  End With
  With Selection.Borders(xlInsideHorizontal)
   .LineStyle = xlContinuous
   .ColorIndex = 0
   .TintAndShade = 0
   .Weight = xlThin
  End With
End Sub
Sub Actualiza_formulas(celdaTFinca As String)
 Dim formulacelda As String
 Dim pos As Integer
 Range(celdaTFinca).Select
 For i = 1 To 26
  formulacelda = ActiveCell.Offset(0, i).Formula2
  pos = InStr(formulacelda, ":")
  ActiveCell.Offset(0, i).Formula2 = Mid(formulacelda, 1, pos) & Replace(ActiveCell.Offset(-1, i).Address, "$", "") & ")"
 Next i
End Sub
Sub probando2()
 Dim n_anho, n_sem, sem_ini, sem_fin As Integer
 Dim x As Double
 n_anho = 2025
 For i = 1 To 12
  sem_fin = WorksheetFunction.WeekNum(CVDate(WorksheetFunction.EoMonth(DateSerial(n_anho, i, 1), 0)))
  sem_ini = WorksheetFunction.WeekNum((DateSerial(n_anho, i, 1))) - IIf(Weekday(DateSerial(n_anho, i, 1)) = 1, 1, 0)
  n_sem = (sem_fin - sem_ini) + 1
  MsgBox UCase(Format("01/" & Format(i, "00") & "/2024", "MMMM - tie\ne: ")) & n_sem & " Semanas"
 Next i
End Sub
