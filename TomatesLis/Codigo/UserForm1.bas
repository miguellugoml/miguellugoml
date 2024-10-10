
Private Sub btn_Cancelar_Click()
  accion = "Cancelar"
  siBuscas = True
  Unload Me
End Sub

Private Sub btn_salir_Click()
  'MsgBox "El mes es: " & el_mes & " y la semana es: " & la_semana
  siBuscas = False
  Unload Me
End Sub

Private Sub cbo_mes_Change()
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

