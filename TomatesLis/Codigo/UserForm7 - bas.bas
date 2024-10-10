Private gn_anho, gn_mes, gn_sem As Integer
Private gs_mes, gsi_mes As String


Private Sub btn_Cancelar_Click()
End
Unload Me
End Sub

Private Sub btn_generar_Libro_Click()
 Dim sheet_n As Worksheet
 Dim nombre_semana, act_name As String
 Dim celdas(7, 3) As String
 Dim fec1 As Date
 Dim le_suma, inMes, n_lineas, act_code, n_reg As Integer
 Dim btn As Shape
 Dim Almacen, LaTorre, GobernI, GobernII, ColorFondo As Double
 Dim marcar As Boolean
 
  Almacen = 13819130
  LaTorre = 13826780
  GobernI = 16440530
 GobernII = 13172735
 celdas(1, 1) = "D": celdas(1, 2) = "LUNES": celdas(1, 3) = "F"
 celdas(2, 1) = "H": celdas(2, 2) = "MARTES": celdas(2, 3) = "J"
 celdas(3, 1) = "L": celdas(3, 2) = "MIERCOLES": celdas(3, 3) = "N"
 celdas(4, 1) = "P": celdas(4, 2) = "JUEVES": celdas(4, 3) = "R"
 celdas(5, 1) = "T": celdas(5, 2) = "VIERNES": celdas(5, 3) = "V"
 celdas(6, 1) = "X": celdas(6, 2) = "SABADO": celdas(6, 3) = "Z"
 celdas(7, 1) = "AB": celdas(7, 2) = "DOMINGO": celdas(7, 3) = "AD"
 gn_anho = Me.cbo_Anho.Value
 gn_mes = Me.cbo_mes.ListIndex + 1
 gs_mes = Me.cbo_mes.Column(0)
 gsi_mes = Mid(gs_mes, 1, 3)
 Application.ScreenUpdating = False
 gn_sem = (Application.WeekNum(CVDate(Application.EoMonth(CVDate("01/" & Format(gs_mes, "00") & "/" & gn_anho) - 1, 1))) - Application.WeekNum(CVDate("01/" & Format(gs_mes, "00") & "/" & gn_anho), vbMonday)) + 1
 Sheets("MES").Name = gs_mes
 Sheets(gs_mes).Activate
 Range("ANHO_LIBRO").Value = gn_anho
 Ciclo_Agrega_Empleados
 Sheets(gs_mes).Activate
 Marca_Duplicados "A:A"
 Range("A600").Select
 Selection.End(xlUp).Select
 n_lineas = ActiveCell.Row
 ColorFondo = Almacen
 For i = 1 To gn_sem
  nombre_semana = "SEMANA_" & gsi_mes & "_" & i
  'MsgBox "Agrega la " & nombre_semana & " Copiada de la " & IIf(i = 1, " Hoja_Base y le pone el año en la celda AI2", "Semana_" & gsi_mes & "_1")
  Sheets("VARIABLES").Activate
  Set sheet_n = ActiveWorkbook.Sheets.Add
  nom = sheet_n.Name
  Sheets(nom).Name = nombre_semana
  If i = 1 Then
    Sheets("Hoja_Base").Visible = True
    Sheets("Hoja_Base").Select
    Cells.Select
    Selection.Copy
    Sheets(nombre_semana).Select
    Cells.Select
    ActiveSheet.Paste
    Sheets("Hoja_Base").Visible = False
    Sheets(nombre_semana).Select
    Range("AI2").Value = gn_anho
    fec1 = Primer_dia_sem_mes_n(Range("AI2").Value, Mes_Sem_H("M"), Mes_Sem_H("S"))
    inMes = Mes_Sem_H("M")
    le_suma = 0
    For j = 1 To 7
      Range(celdas(j, 1) & "2").Value = celdas(j, 2) & " " & Format(fec1 + le_suma, "DD")
      If Val(Format(fec1 + le_suma, "MM")) <> inMes Then
        Range(celdas(j, 1) & "2").Font.Color = RGB(205, 205, 205)
      Else
        Range(celdas(j, 1) & "2").Font.Color = RGB(0, 0, 0)
      End If
      le_suma = le_suma + 1
    Next j
    n_reg = WorksheetFunction.Count(Sheets(gs_mes).Range("A1:A600"))
    For m = 1 To n_reg
      Rows(m * 4 & ":" & (m * 4) + 3).Select
      Selection.Copy
      Rows((m * 4) + 4 & ":" & (m * 4) + 4).Select
      Selection.Insert Shift:=xlDown
      Application.CutCopyMode = False
    Next m
    Sheets(gs_mes).Activate
    Range("A1").Select
    n_reg = 1
    For k = 1 To n_lineas
      ActiveCell.Offset(1, 0).Select
      If IsNumeric(ActiveCell.Text) Then
        act_code = ActiveCell.Value
        act_name = ActiveCell.Offset(0, 1).Value
        Sheets(nombre_semana).Activate
        Range("A" & n_reg * 4).Value = n_reg
        Range("B" & n_reg * 4).Value = act_code
        Range("C" & n_reg * 4).Value = act_name
        Range("A" & n_reg * 4 & ":AF" & n_reg * 4).Interior.Color = ColorFondo
        Sheets(gs_mes).Activate
        n_reg = n_reg + 1
      End If
      Select Case ActiveCell.Value
        Case "Totales_Almacen"
          ColorFondo = LaTorre
        Case "Totales_La_Torre"
          ColorFondo = GobernI
        Case "Totales_GOB_I"
          ColorFondo = GobernII
      End Select
    Next k
  Else
    Sheets("SEMANA_" & gsi_mes & "_1").Activate
    Cells.Select
    Selection.Copy
    Sheets(nombre_semana).Select
    Cells.Select
    ActiveSheet.Paste
    Sheets("SEMANA_" & gsi_mes & "_1").Activate
    Range("A1").Select
    Sheets(nombre_semana).Select
    fec1 = Primer_dia_sem_mes_n(Range("AI2").Value, Mes_Sem_H("M"), Mes_Sem_H("S"))
    inMes = Mes_Sem_H("M")
    le_suma = 0
    For j = 1 To 7
      Range(celdas(j, 1) & "2").Value = celdas(j, 2) & " " & Format(fec1 + le_suma, "DD")
      If Val(Format(fec1 + le_suma, "MM")) <> inMes Then
        Range(celdas(j, 1) & "2").Font.Color = RGB(205, 205, 205)
      Else
        Range(celdas(j, 1) & "2").Font.Color = RGB(0, 0, 0)
      End If
      le_suma = le_suma + 1
    Next j
  End If
  Range("A1").Select
 Next i
 marcar = False
 For i = 1 To gn_sem
  nombre_semana = "SEMANA_" & gsi_mes & "_" & i
  Sheets(nombre_semana).Select
  For j = 1 To 7
   If j > 5 Then
    marcar = True
   ElseIf Range(celdas(j, 1) & 2).Font.Color = RGB(0, 0, 0) Then
    ndia = Val(Mid(Range(celdas(j, 1) & 2).Value, InStr(1, Range(celdas(j, 1) & 2).Value, " ") + 1, 2))
    fec1 = CVDate(ndia & "/" & Format(gn_mes, "00") & "/" & gn_anho)
    Set Festivo = Range("FESTIVOS").Find(fec1)
    If Not Festivo Is Nothing Then
     If Festivo = fec1 Then
      marcar = True
     End If
    End If
   End If
   If marcar Then
     Columns(celdas(j, 3) & ":" & celdas(j, 3)).Interior.Color = 49407
     marcar = False
   End If
  Next j
 Next i
 Sheets(gs_mes).Activate
 Range("A1").Select
 Application.ScreenUpdating = True
 For Each btn In Sheets(gs_mes).Shapes
  If btn.Name = "btn_Genera_Libro" Then
    btn.Visible = False
  Else
    btn.Visible = True
  End If
 Next
 Pon_Semana
 Sheets(gs_mes).Activate
 Nombre_Archivo = Application.ActiveWorkbook.Path & "/PARTES SEMANALES " & gs_mes & " " & gn_anho & ".xlsm"
 Application.ActiveWorkbook.SaveAs Nombre_Archivo
 Unload Me
End Sub

Private Sub UserForm_Activate()
For i = 2020 To Year(Now()) + 30
  Me.cbo_Anho.AddItem i
Next i
Me.cbo_Anho = Year(Now())
  Me.cbo_mes.AddItem "ENERO"
  Me.cbo_mes.AddItem "FEBRERO"
  Me.cbo_mes.AddItem "MARZO"
  Me.cbo_mes.AddItem "ABRIL"
  Me.cbo_mes.AddItem "MAYO"
  Me.cbo_mes.AddItem "JUNIO"
  Me.cbo_mes.AddItem "JULIO"
  Me.cbo_mes.AddItem "AGOSTO"
  Me.cbo_mes.AddItem "SEPTIEMBRE"
  Me.cbo_mes.AddItem "OCTUBRE"
  Me.cbo_mes.AddItem "NOVIEMBRE"
  Me.cbo_mes.AddItem "DICIEMBRE"
  Me.cbo_mes.Value = UCase(Format(Now(), "MMMM"))
End Sub


