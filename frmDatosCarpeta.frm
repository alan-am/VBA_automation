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


Private Sub labelDestino_Click()

End Sub

' Metodo de inicializacion del forms
Private Sub UserForm_Initialize()
    ' Carga de las listas din�micas
    CargarListasDinamicas
    
    ' Seteado valores default de cierto campos
    Me.txtNumCaja.Value = 0
    Me.cmbDestino.Value = "Conservaci�n"
    Me.cmbSoporte.Value = "Digital"
    Me.txtFechaCierre.Value = "dd/mm/aaaa"
End Sub
Private Sub btnCerrar_Click()
    Unload Me
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFormulario 'modUtilidades
End Sub

Private Sub btnSeleccionarCarpeta_Click()
    ' Llama a la funci�n que abrimos di�logo y llena los TextBox
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
    
    ' Definicion hoja de configuraci�n
    Set ws = ThisWorkbook.Sheets("Config")
    
    ' Reinicion de los comboBox
    Me.cmbSerieSubserie.Clear
    Me.cmbDestino.Clear
    Me.cmbSoporte.Clear
    
    ' OJO Asumiendo que la Fila 1 es el t�tulo
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
    MsgBox "Error al cargar las listas de configuraci�n." & vbCrLf & _
           "Aseg�rese que la hoja 'Config' existe y tiene el formato correcto.", _
           vbCritical, "Error de Carga"
End Sub

' Funcion btn Insertar Datos
Private Sub btnInsertar_Click()

    ' Validar que los datos de la carpeta no esten vacios
    If pDatosCarpeta Is Nothing Then
        MsgBox "Error: Primero debe seleccionar una carpeta usando el bot�n 'Examinar...'.", vbCritical, "Acci�n Requerida"
        Me.btnSeleccionarCarpeta.SetFocus ' Sugiere al usuario qu� bot�n presionar
        Exit Sub
    End If
    
    ' Validar que el campo serie no este vacio
    If Trim(Me.cmbSerieSubserie.Value) = "" Then
        MsgBox "El campo 'Serie/Subserie' es obligatorio.", vbCritical, "Dato Faltante"
        Me.cmbSerieSubserie.SetFocus
        Exit Sub
    End If
    
    ' Validar que el campo destino no este vacio
    If Trim(Me.cmbDestino.Value) = "" Then
        MsgBox "El campo 'Destino Final' es obligatorio.", vbCritical, "Dato Faltante"
        Me.cmbDestino.SetFocus
        Exit Sub
    End If
    
    'Validar que el campo soporte no este vacio
    If Trim(Me.cmbSoporte.Value) = "" Then
        MsgBox "El campo 'Soporte' es obligatorio.", vbCritical, "Dato Faltante"
        Me.cmbSoporte.SetFocus
        Exit Sub
    End If
    
    ' Validaci�n opcional para la fecha de cierre(Deshabilitado)
    'If Trim(Me.txtFechaCierre.Value) <> "" And Not IsDate(Me.txtFechaCierre.Value) Then
        'MsgBox "El formato de 'Fecha Cierre' no es v�lido. Use un formato de fecha.", vbExclamation, "Formato Inv�lido"
        'Me.txtFechaCierre.SetFocus
        'Exit Sub
    'End If
    
    
    ' seteo de los datos manuales de la carpeta
    ' ya se tiene de la carpeta:
    ' ("Nombre", "Ruta", "CantidadArchivos", "TamanoTotal", "FechaCreacion")
    
    ' Agregamos los nuevos datos manuales
    pDatosCarpeta("SerieSubserie") = Me.cmbSerieSubserie.Value
    pDatosCarpeta("NumExpediente") = Me.txtNumExpediente.Value
    pDatosCarpeta("Destino") = Me.cmbDestino.Value
    pDatosCarpeta("Soporte") = Me.cmbSoporte.Value
    pDatosCarpeta("Observaciones") = Me.txtObservaciones.Value
    
    ' Campo oculto requerido
    pDatosCarpeta("UbicacionTopografica") = "NN"
    
    ' validar N�mero de Caja (asegurar que sea num�rico)
    pDatosCarpeta("NumCaja") = IIf(IsNumeric(Me.txtNumCaja.Value), CLng(Me.txtNumCaja.Value), 0)
    
    'REVISAR -> lafecha debe estar vacia o ser valida, sino excepcion y focus en fecha.
    ' validar Fecha de Cierre final (asegurar que sea fecha o vacia)
    If IsDate(Me.txtFechaCierre.Value) Then
        pDatosCarpeta("FechaCierre") = CDate(Me.txtFechaCierre.Value)
    Else
        pDatosCarpeta("FechaCierre") = "dd/mm/aaaa"
    End If
    
    
    ' --- ESCRIBIR LOS DATOS
    
    ' Pasamos el diccionario completo a la funci�n de exportaci�n
    If ExportarDatosInventario(pDatosCarpeta) Then
        MsgBox "��xito! Los datos de la carpeta '" & pDatosCarpeta("Nombre") & "' se han guardado en el inventario.", vbInformation, "Exportaci�n Completa"
        
        ' Limpiar formulario para siguiente ingreso
        LimpiarFormulario
        
        ' Resetea el diccionario
        Set pDatosCarpeta = Nothing
    Else
        MsgBox "Ocurri� un error al intentar guardar los datos en la hoja de Excel.", vbCritical, "Error de Exportaci�n"
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
