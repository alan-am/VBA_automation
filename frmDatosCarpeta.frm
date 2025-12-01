VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatosCarpeta 
   Caption         =   "Gestor de Carpetas Digitales"
   ClientHeight    =   7335
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

Private Sub labelSerieDocumento_Click()

End Sub

Private Sub txtNombreCarpeta_Change()

End Sub

' Metodo de inicializacion del forms
Private Sub UserForm_Initialize()
    ' Carga de las listas dinámicas
    CargarListasDinamicas
    
    ' Seteado valores default de cierto campos
    Me.txtNumCaja.Value = 0
    Me.cmbDestino.Value = "Conservación"
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
    Me.cmbSerie.Clear
    Me.cmbSubserie.Clear
    Me.cmbDestino.Clear
    Me.cmbSoporte.Clear
    
    ' OJO Asumiendo que la Fila 1 es el título
    ' Cargar Serie Documental(Columna B)
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 2 To lastRow
        If Trim(ws.Cells(i, "B").Value) <> "" Then
            Me.cmbSerie.AddItem ws.Cells(i, "B").Value
        End If
    Next i
    
    ' Cargar Subserie Documental(Columna C)
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    For i = 2 To lastRow
        If Trim(ws.Cells(i, "C").Value) <> "" Then
            Me.cmbSubserie.AddItem ws.Cells(i, "C").Value
        End If
    Next i
    
    ' Cargar Destino Final (Columna D)
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    For i = 2 To lastRow
        If Trim(ws.Cells(i, "D").Value) <> "" Then
            Me.cmbDestino.AddItem ws.Cells(i, "D").Value
        End If
    Next i
    
    ' Cargar Soporte (Columna E)
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    For i = 2 To lastRow
        If Trim(ws.Cells(i, "E").Value) <> "" Then
            Me.cmbSoporte.AddItem ws.Cells(i, "E").Value
        End If
    Next i
    
    Exit Sub

ErrorHandler:
    MsgBox "Error al cargar las listas de configuración." & vbCrLf & _
           "Asegúrese que la hoja 'Config' existe y tiene el formato correcto.", _
           vbCritical, "Error de Carga"
End Sub

' Funcion btn Insertar Datos
Private Sub btnInsertar_Click()

    'MEJORA -> Deshabilitar boton de click , al presionar, para evitar doble click.
    'Me.btnInsertarDatos.Enabled = False
    'Me.btnInsertarDatos.Enabled = True

    ' Validar que los datos de la carpeta no esten vacios
    If pDatosCarpeta Is Nothing Then
        MsgBox "Error: Primero debe seleccionar una carpeta usando el botón 'Examinar...'.", vbCritical, "Acción Requerida"
        Me.btnSeleccionarCarpeta.SetFocus ' Sugiere al usuario qué botón presionar
        Exit Sub
    End If
    
    ' Validar que el campo serie no este vacio
    If Trim(Me.cmbSerie.Value) = "" Then
        MsgBox "El campo 'Serie' es obligatorio.", vbCritical, "Dato Faltante"
        Me.cmbSerie.SetFocus
        Exit Sub
    End If
    
    ' Validar que el campo Subsserie no este vacio
    If Trim(Me.cmbSubserie.Value) = "" Then
        MsgBox "El campo 'Subserie' es obligatorio.", vbCritical, "Dato Faltante"
        Me.cmbSubserie.SetFocus
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
    
    ' Validación opcional para la fecha de cierre(Deshabilitado)
    'If Trim(Me.txtFechaCierre.Value) <> "" And Not IsDate(Me.txtFechaCierre.Value) Then
        'MsgBox "El formato de 'Fecha Cierre' no es válido. Use un formato de fecha.", vbExclamation, "Formato Inválido"
        'Me.txtFechaCierre.SetFocus
        'Exit Sub
    'End If
    
    
    ' seteo de -todos- los datos manuales de la carpeta
       
    pDatosCarpeta("Nombre") = Me.txtNombreCarpeta.Value
    pDatosCarpeta("Ruta") = Me.txtRutaCarpeta.Value
    pDatosCarpeta("CantidadArchivos") = Me.txtCantidadArchivos.Value
    pDatosCarpeta("TamanoTotal") = Me.txtTamanoTotal.Value
    
    ' Agregamos los nuevos datos manuales
    pDatosCarpeta("Serie") = Me.cmbSerie.Value
    pDatosCarpeta("Subserie") = Me.cmbSubserie.Value
    pDatosCarpeta("NumExpediente") = Me.txtNumExpediente.Value
    pDatosCarpeta("Destino") = Me.cmbDestino.Value
    pDatosCarpeta("Soporte") = Me.cmbSoporte.Value
    pDatosCarpeta("Observaciones") = Me.txtObservaciones.Value
    
    ' Campo oculto requerido
    pDatosCarpeta("UbicacionTopografica") = "NN"
    
    ' validar Número de Caja (asegurar que sea numérico)
    pDatosCarpeta("NumCaja") = IIf(IsNumeric(Me.txtNumCaja.Value), CLng(Me.txtNumCaja.Value), 0)
    
    '  Validaciones de fecha
    'MEJORA -> lafecha debe estar vacia o ser valida, sino excepcion y focus en fecha(bloqueando la escritura en excel hasta que tenga buen formato).
    
    
    ' validar Fecha de Cierre final (asegurar que sea fecha o vacia)
    If IsDate(Me.txtFechaCierre.Value) Then
        pDatosCarpeta("FechaCierre") = CDate(Me.txtFechaCierre.Value)
    Else
        pDatosCarpeta("FechaCierre") = "dd/mm/aaaa"
    End If
    
    
    ' validar Fecha de creacion(Si no se parsea correctamente, se escribe en el excel como "dd/mm/aaaa")
    If IsDate(Me.txtFechaCreacion.Value) Then
        pDatosCarpeta("FechaCreacion") = CDate(Me.txtFechaCreacion.Value)
    Else
        pDatosCarpeta("FechaCreacion") = "dd/mm/aaaa"
    End If
    
    
    
    ' --- ESCRIBIR LOS DATOS
    
    ' Pasamos el diccionario completo a la función de exportación
    If ExportarDatosInventario(pDatosCarpeta) Then
        MsgBox "¡Éxito! Los datos de la carpeta '" & pDatosCarpeta("Nombre") & "' se han guardado en el inventario.", vbInformation, "Exportación Completa"
        
        ' Limpiar formulario para siguiente ingreso
        LimpiarFormulario
        
        ' Resetea el diccionario
        Set pDatosCarpeta = Nothing
    Else
        MsgBox "Ocurrió un error al intentar guardar los datos en la hoja de Excel.", vbCritical, "Error de Exportación"
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
