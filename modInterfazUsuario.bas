Attribute VB_Name = "modInterfazUsuario"
' Acci�n del bot�n Analizar
Sub btnAnalizar_Click()
    Dim folderPath As String
    folderPath = SeleccionarCarpeta()
    If folderPath <> "" Then
        MostrarDatosCarpeta folderPath
    End If
End Sub

' Acci�n del bot�n Limpiar
Sub btnLimpiarCampos_Click()
    LimpiarFormulario
End Sub

' Acci�n del bot�n Cerrar
Sub btnCerrar_Click()
    Unload frmDatosCarpeta
End Sub

