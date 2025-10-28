VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatosCarpeta 
   Caption         =   "Gestor de Carpetas Digitales"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7515
   OleObjectBlob   =   "frmDatosCarpeta.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frmDatosCarpeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCerrar_Click()
    Unload Me
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFormulario
End Sub

Private Sub btnSeleccionarCarpeta_Click()
    ' Llama a la función que abrimos diálogo y llena los TextBox
    Dim folderPath As String
    folderPath = SeleccionarCarpeta()
    
    If folderPath <> "" Then
        MostrarDatosCarpeta folderPath
    End If
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub encabezado_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub logoEspol_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub observaciones_Click()

End Sub

Private Sub seccionDatos_Click()

End Sub

Private Sub serieDocumento_Click()

End Sub

Private Sub soporte_Click()

End Sub

Private Sub tamanio_Click()

End Sub

Private Sub Titulo_Click()

End Sub

Private Sub txtRutaCarpeta_Change()

End Sub

Private Sub UserForm_Click()

End Sub
