Private Sub UserForm_Activate()
 'Agregué
 'LocalPath = Excel.CurrentWorkbook(){[Name="FilePath"]}[Content]{0}[Column1],   Source = Excel.Workbook(File.Contents(LocalPath), null, true),
 'en la formulación de las conexiones de querys para parametrizar la ubicación de los archivos
 Dim No_ExisteFile As Boolean
 Dim la_Ruta As String
 Dim l_path, el_Path As String
 l_path = ActiveWorkbook.FullName
 l_path = Mid(l_path, 1, InStr(l_path, "PARTES") - 1)
 el_Path = l_path & "NOMINA_MOBIBUK.xlsm"
 'MsgBox el_Path
 Application.ScreenUpdating = False
 Sheets("NOMINA").Activate
 No_ExisteFile = Dir(el_Path) <> "NOMINA_MOBIBUK.xlsm"
 If No_ExisteFile Then
   No_ExisteFile = Dir(Range("AA1").Value) <> "NOMINA_MOBIBUK.xlsm"
 Else
   la_Ruta = el_Path
   Range("AA1").Value = la_Ruta
   Sheets("MOBIBUK").Activate
   Range("AA1").Value = la_Ruta
   Sheets("Fichaje").Activate
   Range("AA1").Value = Replace(la_Ruta, "NOMINA_MOBIBUK.xlsm", "Fichaje.xlsx")
 End If
 If No_ExisteFile Then
   MsgBox "Seleccione el archivo NOMINA_MOBIBUK.xlsm donde también debe estra Fichaje.xlsx ", vbCritical, "Para poder actualizar los links, por favor..."
   ChDir (l_path)
   la_Ruta = Application.GetOpenFilename()
   Range("AA1").Value = la_Ruta
   Sheets("MOBIBUK").Activate
   Range("AA1").Value = la_Ruta
   Sheets("Fichaje").Activate
   Range("AA1").Value = Replace(la_Ruta, "NOMINA_MOBIBUK.xlsm", "Fichaje.xlsx")
 End If
 Sheets("NOMINA").Activate
 No_ExisteFile = Dir(ActiveSheet.Range("AA1").Value) <> "NOMINA_MOBIBUK.xlsm"
 If No_ExisteFile Then
   MsgBox "Deben exitir los archivos: Fichaje.xlsx y NOMINA_MOBIBUK.xlsm", vbCritical, "ERROR: El archivo podría arrojar resultados errados"
 End If
 
 If Application.Wait(Now + TimeValue("0:00:03")) Then
  ActiveWorkbook.Queries("Fichaje").Refresh
  Me.Image1.Visible = True
  Me.Label2.ForeColor = RGB(0, 255, 0)
  ActiveWorkbook.Queries("MOBIBUK").Refresh
  Me.Image2.Visible = True
  Me.Label3.ForeColor = RGB(0, 255, 0)
  ActiveWorkbook.Queries("NOMINA").Refresh
  Me.Image3.Visible = True
  Me.Label4.ForeColor = RGB(0, 255, 0)
 End If
 Application.ScreenUpdating = True
 If Application.Wait(Now + TimeValue("0:00:03")) Then
  'UserForm2.Hide
  Unload Me
 End If
End Sub
