Private Sub btn_Calcular_Click()
  Carga_Codigo
End Sub

Private Sub btn_Carga_Jornada_Click()
Dim sh_volver As String
sh_volver = ActiveSheet.Name
With UserForm4
  .Show
End With
Sheets(sh_volver).Activate
End Sub

Private Sub btn_Carga_Variables_Click()
Dim sh_volver As String
sh_volver = ActiveSheet.Name
With UserForm5
  .Show
End With
Sheets(sh_volver).Activate
AgregarNoEncontradosVariables
End Sub

Private Sub btn_Cargar_Fichaje_Click()
Dim sh_volver As String
sh_volver = ActiveSheet.Name
With UserForm3
  .Show
End With
Sheets(sh_volver).Activate
End Sub

Private Sub btn_Genera_Libro_Click()
 Dim archivo_cerrar As String
 archivo_cerrar = Application.ThisWorkbook.FullName
 With UserForm7
  .Show
 End With
 If Application.Wait(Now + TimeValue("0:00:05")) Then
   Application.Workbooks.Open Nombre_Archivo
 End If
 'Application.Workbooks(archivo_cerrar).Close SaveChanges:=False
End Sub