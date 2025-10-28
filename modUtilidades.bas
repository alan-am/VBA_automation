Attribute VB_Name = "modUtilidades"

' Muestra el diálogo para seleccionar carpeta y devuelve la ruta
Function SeleccionarCarpeta() As String
    Dim folderDialog As fileDialog
    Set folderDialog = Application.fileDialog(msoFileDialogFolderPicker)
    
    With folderDialog
        .Title = "Selecciona una carpeta para analizar"
        If .Show = -1 Then
            SeleccionarCarpeta = .SelectedItems(1)
        Else
            MsgBox "No se seleccionó ninguna carpeta.", vbExclamation, "Cancelado"
            SeleccionarCarpeta = ""
        End If
    End With
End Function

' Obtiene información de la carpeta y devuelve un diccionario
Function ObtenerInfoCarpeta(folderPath As String) As Object
    Dim fso As Object, carpeta As Object
    Dim info As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpeta = fso.GetFolder(folderPath)
    Set info = CreateObject("Scripting.Dictionary")
    
    
    
    info("Nombre") = carpeta.Name
    info("Ruta") = carpeta.Path
    info("CantidadArchivos") = carpeta.Files.Count
    
    ' seteamos los bytes a KB(/1024) y redondeamos
    info("TamanoTotal") = Round(carpeta.Size / 1024, 1)
    ' definimos que solo quede la fecha y no las horas.
    info("FechaCreacion") = DateValue(carpeta.DateCreated)
    
    Set ObtenerInfoCarpeta = info
End Function

' Limpia todos los campos del formulario
Sub LimpiarFormulario()
    With frmDatosCarpeta
        .txtRutaCarpeta.Value = ""
        .txtNombreCarpeta.Value = ""
        .txtFechaCreacion.Value = ""
        .txtCantidadArchivos.Value = ""
        .txtTamanoTotal.Value = ""
        .cmbSoporte.ListIndex = -1
        .txtObservaciones.Value = ""
    End With
End Sub

