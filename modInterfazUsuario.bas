Attribute VB_Name = "modInterfazUsuario"
' Acción del botón Analizar
Sub btnAnalizar_Click()
    Dim folderPath As String
    folderPath = SeleccionarCarpeta()
    If folderPath <> "" Then
        MostrarDatosCarpeta folderPath
    End If
End Sub

' Acción del botón Limpiar
Sub btnLimpiarCampos_Click()
    LimpiarFormulario
End Sub

' Acción del botón Cerrar
Sub btnCerrar_Click()
    Unload frmDatosCarpeta
End Sub

