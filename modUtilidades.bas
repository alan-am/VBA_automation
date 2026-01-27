Attribute VB_Name = "modUtilidades"
'modUtilidades
' **************************************************************************
' ! UTILIDADES GENÉRICAS DEL SISTEMA DE ARCHIVOS
' **************************************************************************

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
    Dim fso As Object, carpeta As Object, archivo As Object
    Dim info As Object
    Dim fechaMax As Date
    Dim fechaArchivo As Date
    Dim conteoFojas As Long
    
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpeta = fso.GetFolder(folderPath)
    Set info = CreateObject("Scripting.Dictionary")
    
    ' inicializacion variable de fecha Max
    fechaMax = 0
    conteoFojas = 0
    
    'Busqueda de la fecha de cierre
    For Each archivo In carpeta.Files
        If (UCase(archivo.Path) <> UCase(ThisWorkbook.FullName)) And _
           (Left(archivo.Name, 1) <> "~") And _
           (LCase(Right(archivo.Name, 4)) <> ".tmp") Then
            
            ' AUMENTAMOS # Fojas
            conteoFojas = conteoFojas + 1
            
            ' Seteo de fecha
            fechaArchivo = archivo.DateLastModified
            If fechaArchivo > fechaMax Then
                fechaMax = fechaArchivo
            End If
            
        End If
    Next archivo
    
    
    
    
    'Seteo de demas datos
    info("Nombre") = carpeta.Name
    info("Ruta") = carpeta.Path
    info("CantidadArchivos") = conteoFojas
    
    ' seteamos los bytes a KB(/1024) y redondeamos
    info("TamanoTotal") = Round(carpeta.Size / 1024, 1)
    ' definimos que solo quede la fecha y no las horas.
    info("FechaCreacion") = DateValue(carpeta.DateCreated)
    
    ' Si fechaMax sigue siendo 0 (carpeta vacía), lo dejamos DEFAULT
    If fechaMax > 0 Then
        info("FechaCierre") = DateValue(fechaMax)
    Else
        info("FechaCierre") = "dd/mm/aaaa"
    End If
    
    Set ObtenerInfoCarpeta = info
End Function





