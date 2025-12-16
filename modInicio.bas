Attribute VB_Name = "modInicio"
'modInicio

Sub AbrirFormularioDatosCarpeta()
    ' Muestra el formulario principal
    frmDatosCarpeta.Show
End Sub

Sub MostrarDatosCarpeta(info As Object)

    'llenamos los datos con la info del argumento
    
    With frmDatosCarpeta
        .txtRutaCarpeta.Value = info("Ruta")
        .txtNombreCarpeta.Value = info("Nombre")
        .txtFechaCreacion.Value = info("FechaCreacion")
        .txtCantidadArchivos.Value = info("CantidadArchivos")
        .txtTamanoTotal.Value = info("TamanoTotal")
        .txtFechaCierre.Value = info("FechaCierre")
        
    End With
End Sub


