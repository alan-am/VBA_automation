Attribute VB_Name = "modUtilidades"
'modUtilidades

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
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpeta = fso.GetFolder(folderPath)
    Set info = CreateObject("Scripting.Dictionary")
    
    ' inicializacion variable de fecha Max
    fechaMax = 0
    
    'Busqueda de la fecha de cierre
    For Each archivo In carpeta.Files
        ' IMPORTANTE: se excluye el propio archivo Excel si está dentro de la carpeta
        ' Usamos UCase para asegurar que la comparación no falle por mayúsculas/minúsculas
        If (UCase(archivo.Path) <> UCase(ThisWorkbook.FullName)) And _
           (Left(archivo.Name, 2) <> "~$") Then
            
            fechaArchivo = archivo.DateLastModified
            
            ' Si la fecha de este archivo es mayor a la que tenemos guardada, actualizamos
            If fechaArchivo > fechaMax Then
                fechaMax = fechaArchivo
            End If
            
        End If
    Next archivo
    
    
    
    
    'Seteo de demas datos
    info("Nombre") = carpeta.Name
    info("Ruta") = carpeta.Path
    info("CantidadArchivos") = carpeta.Files.Count
    
    ' seteamos los bytes a KB(/1024) y redondeamos
    info("TamanoTotal") = Round(carpeta.Size / 1024, 1)
    ' definimos que solo quede la fecha y no las horas.
    info("FechaCreacion") = DateValue(carpeta.DateCreated)
    
    ' Si fechaMax sigue siendo 0 (carpeta vacía), lo dejamos vacío
    If fechaMax > 0 Then
        info("FechaCierre") = DateValue(fechaMax)
    Else
        info("FechaCierre") = "dd/mm/aaaa"
    End If
    
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
        .txtObservaciones.Value = ""
        .txtFechaCierre.Value = "dd/mm/aaaa"
    End With
End Sub

Private Sub buscarSeccion_Click()

    Dim ws As Worksheet
    
    ' Validar que se haya seleccionado algo
    If Me.ComboBox1.Value = "" Then
        MsgBox "Por favor, selecciona o escribe una opción.", vbExclamation
        Exit Sub
    End If

    ' Usamos la hoja activa (desde donde llamaste al botón)
    Set ws = ActiveSheet

    On Error GoTo ErrorHandler

    ' 2. Escribir SIEMPRE en la celda E5
    ' (Al ser combinada, basta con apuntar a la primera celda del rango, que es C6)
    ws.Range("E5").Value = Me.ComboBox1.Value


    ' Cerrar formulario
    Unload Me
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub


Function GenerarNuevoCodigoExpediente() As String
    ' Esta función calcula el código visualmente para el formulario
    
    Dim wsConfig As Worksheet
    Dim wsInventario As Worksheet
    Dim tbl As ListObject
    Dim codSeleccionado As String
    Dim siguienteNumero As Long
    
    ' Definición de Constantes
    Const FORMATO_PREFIJO As String = "ESPOL-"
    Const CODIGO_DESCONOCIDO As String = "???"
    Const NOMBRE_TABLA As String = "tabla_test89"
    
    ' Referencias
    Set wsInventario = ThisWorkbook.Sheets("Inventario General")
    On Error Resume Next
    Set tbl = wsInventario.ListObjects(NOMBRE_TABLA)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        GenerarNuevoCodigoExpediente = "Error-Tabla"
        Exit Function
    End If
    
    ' Leer Código de Sección
    codSeleccionado = Trim(Hoja4.Range("Q2").Value)
    If codSeleccionado = "" Then codSeleccionado = CODIGO_DESCONOCIDO
    
    ' cantidad filas actual + 1
    siguienteNumero = tbl.ListRows.Count + 1
    
    ' Armar el String
    GenerarNuevoCodigoExpediente = FORMATO_PREFIJO & codSeleccionado & "-" & Format(siguienteNumero, "000")

End Function

