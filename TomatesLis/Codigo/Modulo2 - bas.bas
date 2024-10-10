
Sub Agrega_Paginas()
Dim sheet_n As Worksheet
Dim nom As String
  If el_mes = "" Then
    carga_mes
  End If
  Set sheet_n = ActiveWorkbook.Sheets.Add '(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
  nom = sheet_n.Name
  Sheets(nom).Name = "MIGUEL"
  Sheets("Hoja_Base").Visible = True
  Sheets("Hoja_Base").Select
  Cells.Select
  Selection.Copy
  Sheets("MIGUEL").Select
  Cells.Select
  ActiveSheet.Paste
  Sheets("Hoja_Base").Visible = False
  Sheets("MIGUEL").Select
  Range("A1").Select
End Sub

Sub Prepara_Libro()
'SOLO SI HAY UNA HOJA QUE SE LLAMA MES que contendrá la hoja principal de los trabajadores separados por invernadero
Dim x_anho, x_sem As Integer
Dim s_mes As String
Dim no_tiene_mes As Boolean

For Each sheet In Worksheets
  no_tiene_mes = sheet.Name = "MES"
  If no_tiene_mes Then
    Exit For
  End If
Next
If Not no_tiene_mes Then
'pide año y mes
 
 x_anho = 2024
 x_mes = 9
 s_mes = Mid(meses, InStr(meses, el_mes), 3)
 x_sem = (Application.WeekNum(CVDate(Application.EoMonth(CVDate("01/" & Format(x_mes, "00") & "/" & x_anho) - 1, 1))) - Application.WeekNum(CVDate("01/" & Format(x_mes, "00") & "/" & x_anho), vbMonday)) + 1
 For i = 1 To x_sem
 'si i = 1 agrega la SEMANA desde la hoja base y a esta le agrega todos los trabajadores de la hoja mes
  MsgBox "Agrega la SEMANA_" & s_mes & "_" & i & " Copiada de la " & IIf(i = 1, " Hoja_Base", "Semana_" & s_mes & "_1")
 Next i
End If
End Sub

Sub cx()
  Dim btn As Shape
  For Each btn In Sheets("MES").Shapes
    If btn.Name = "btn_Genera_Libro" Then
      btn.Visible = True
    Else
      btn.Visible = False
    End If
  Next
  ThisWorkbook.Close SaveChanges:=False
End Sub


