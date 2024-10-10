
Private Sub Workbook_Open()
 siBuscas = True
 carga_mes
 With UserForm2
   .Show
 End With
 'Act_Dias
 'Pon_Semana
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
    If num_semana > 0 And num_semana <= 6 Then
      Sheets(x_sem).Activate
      Range(Range("B1:B2000").Find(cod_buscar).Address).Select
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
  siBuscas = UCase(InputBox("Quieres seguri buscando", "confirma", "no")) = "SI"
End Sub
