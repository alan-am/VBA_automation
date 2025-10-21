Attribute VB_Name = "modInicio"
Sub AbrirFormularioDatosCarpeta()
    ' Muestra el formulario principal
    frmDatosCarpeta.Show
End Sub

Sub MostrarDatosCarpeta(folderPath As String)
    Dim info As Object
    Set info = ObtenerInfoCarpeta(folderPath)
    
    With frmDatosCarpeta
        .txtRutaCarpeta.Value = info("Ruta")
        .txtNombreCarpeta.Value = info("Nombre")
        .txtFechaCreacion.Value = info("FechaCreacion")
        .txtCantidadArchivos.Value = info("CantidadArchivos")
        .txtTamanoTotal.Value = info("TamanoTotal")
    End With
End Sub

