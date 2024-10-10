
Private Sub Workbook_Open()
 siBuscas = True
 carga_mes
 With UserForm2
   .Show
 End With
 For Each sh In ThisWorkbook.Sheets
  If sh.Name = "MES" Then
    For Each btn In Sheets("MES").Shapes
      If btn.Name = "btn_Genera_Libro" Then
        btn.Visible = True
      Else
        btn.Visible = False
      End If
    Next
    Exit For
  End If
 Next
 UserForm2.Enabled = False
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)
 On Error GoTo salirbusqueda

 Dim num_semana, cod_buscar, lin As Integer
 Dim los_meses, x_sem As String
 los_meses = "ENERO     FEBRERO   MARZO     ABRIL     MAYO      JUNIO     JULIO     AGOSTO    SEPTIEMBREOCTUBRE   NOVIEMBRE DICIEMBRE "
 If InStr(1, los_meses, Mid(ActiveSheet.Name, 1, 10)) > 0 Then
  x_sem = "SEMANA_" & Mid(el_mes, 1, 3) & "_"
  If Mid(Replace(ActiveCell.AddressLocal, "$", ""), 1, 1) = "B" And ActiveCell.Value <> "" Then
   cod_buscar = ActiveCell.Offset(0, -1).Value
   If siBuscas Then
    num_semana = InputBox("Número de la semana de " & ActiveSheet.Name, "Ingresa el dato, por favor", 1)
   Else
    num_semana = 0
   End If
   If num_semana > 0 And num_semana <= 6 Then
    x_sem = x_sem & num_semana
    Sheets(x_sem).Activate
    Range(Range("B1:B2000").Find(cod_buscar).Address).Select
    lin = ActiveCell.Row
    Range(lin & ":" & lin).Select
   End If
  Else
   Exit Sub
  End If
 Else
  Exit Sub
 End If
 Exit Sub

salirbusqueda:
 If siBuscas Then
  MsgBox "No se consigue"
  siBuscas = UCase(InputBox("Quieres mantener activa la opción de búsqueda", "confirma", "no")) = "SI"
 End If
End Sub
