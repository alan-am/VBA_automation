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

' Variable a nivel de formulario para guardar los datos de la carpeta
Private pDatosCarpeta As Object

' Metodo de inicializacion del forms
Private Sub UserForm_Initialize()
    ' Carga de las listas dinámicas
    CargarListasDinamicas
    
    ' Seteado valores default de cierto campos
    Me.txtNumCaja.Value = 0
    Me.cmbDestino.Value = "Conservación"
    Me.cmbSoporte.Value = "Digital"
End Sub
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
    ' Obtiene el diccionario y lo guarda en la variable del formulario
        Set pDatosCarpeta = ObtenerInfoCarpeta(folderPath) ' modUtilidades
        
        MostrarDatosCarpeta pDatosCarpeta 'modInicio
    End If
End Sub

' Carga de opciones para los comboBox en el forms
' los datos se cargan a partir de la hoja "Config"
Private Sub CargarListasDinamicas()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' Definicion hoja de configuración
    Set ws = ThisWorkbook.Sheets("Config")
    
    ' Reinicion de los comboBox
    Me.cmbSerieSubserie.Clear
    Me.cmbDestino.Clear
    Me.cmbSoporte.Clear
    
    ' OJO Asumiendo que la Fila 1 es el título
    ' Cargar Serie/Subserie (Columna A)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If Trim(ws.Cells(i, "A").Value) <> "" Then
            Me.cmbSerieSubserie.AddItem ws.Cells(i, "A").Value
        End If
    Next i
    
    ' Cargar Destino Final (Columna B)
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 2 To lastRow
        If Trim(ws.Cells(i, "B").Value) <> "" Then
            Me.cmbDestino.AddItem ws.Cells(i, "B").Value
        End If
    Next i
    
    ' Cargar Soporte (Columna C)
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    For i = 2 To lastRow
        If Trim(ws.Cells(i, "C").Value) <> "" Then
            Me.cmbSoporte.AddItem ws.Cells(i, "C").Value
        End If
    Next i
    
    Exit Sub

ErrorHandler:
    MsgBox "Error al cargar las listas de configuración." & vbCrLf & _
           "Asegúrese que la hoja 'Config' existe y tiene el formato correcto.", _
           vbCritical, "Error de Carga"
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
