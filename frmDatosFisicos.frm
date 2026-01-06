VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatosFisicos 
   Caption         =   "Gestor de Carpetas Físicas"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7515
   OleObjectBlob   =   "frmDatosFisicos.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frmDatosFisicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'frmDatosFisicos

' Variable a nivel de formulario para guardar los datos de la carpeta
Private pDatosCarpeta As Object


' Metodo de inicializacion del forms
Private Sub UserForm_Initialize()
    ' Carga de las listas dinámicas
    CargarListasDinamicas
    
    ' Seteado valores default de cierto campos
    Me.txtNumCaja.Value = 0
    Me.cmbDestino.Value = "Conservación"
    Me.cmbSoporte.Value = "Físico"
    Me.txtFechaCierre.Value = "dd/mm/aaaa"
    Me.txtFechaCreacion.Value = "dd/mm/aaaa"
    Me.txtNumExpediente.Value = GenerarNuevoCodigoExpediente()
    
    
End Sub
Private Sub btnCerrar_Click()
    Unload Me
End Sub

Private Sub btnLimpiar_Click()
    ' Limpiar campos manuales
    Me.txtNombreCarpeta.Value = ""
    Me.txtFechaCreacion.Value = "dd/mm/aaaa"
    Me.txtCantidadArchivos.Value = "" ' Fojas
    Me.txtObservaciones.Value = ""
    Me.txtFechaCierre.Value = "dd/mm/aaaa"
    Me.txtNumExpediente.Value = ""
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
    
    ' Reinicio de los comboBox
    Me.cmbSerie.Clear
    Me.cmbSubserie.Clear
    Me.cmbDestino.Clear
    Me.cmbSoporte.Clear
    
    
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
    MsgBox "Error al cargar las listas en configuración." & vbCrLf & _
           "Asegúrese que la hoja 'Config' existe y tiene el formato correcto.", _
           vbCritical, "Error de Carga"
End Sub

' Funcion btn Insertar Datos
Private Sub btnInsertar_Click()

    'MEJORA -> Deshabilitar boton de click , al presionar, para evitar doble click.
    'Me.btnInsertarDatos.Enabled = False
    'Me.btnInsertarDatos.Enabled = True


    ' VALIDACIÓNES DE CAMPOS OBLIGATORIOS
    ' Nombre Carpeta
    If Trim(Me.txtNombreCarpeta.Value) = "" Then
        MsgBox "El campo 'Nombre Carpeta' es obligatorio.", vbCritical, "Dato Faltante"
        Me.txtNombreCarpeta.SetFocus
        Exit Sub
    End If
    
    ' Fojas (CantidadArchivos) - Debe tener valor
    If Trim(Me.txtCantidadArchivos.Value) = "" Then
        MsgBox "El campo 'N° Fojas' es obligatorio.", vbCritical, "Dato Faltante"
        Me.txtCantidadArchivos.SetFocus
        Exit Sub
    End If
    
    ' Fecha Creación - Debe ser fecha válida
    If Not IsDate(Me.txtFechaCreacion.Value) Then
        MsgBox "El campo 'Fecha de Creación' es obligatorio y debe ser una fecha válida.", vbCritical, "Formato Incorrecto"
        Me.txtFechaCreacion.SetFocus
        Exit Sub
    End If
    
    ' Serie y Subserie
    If Trim(Me.cmbSerie.Value) = "" Then
        MsgBox "El campo 'Serie' es obligatorio.", vbCritical, "Dato Faltante"
        Me.cmbSerie.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.cmbSubserie.Value) = "" Then
        MsgBox "El campo 'Subserie' es obligatorio.", vbCritical, "Dato Faltante"
        Me.cmbSubserie.SetFocus
        Exit Sub
    End If
    
    ' N° Caja
    If Trim(Me.txtNumCaja.Value) = "" Then
        MsgBox "El campo 'N° Caja' es obligatorio.", vbCritical, "Dato Faltante"
        Me.txtNumCaja.SetFocus
        Exit Sub
    End If
    
    ' Soporte y Destino
    If Trim(Me.cmbSoporte.Value) = "" Then
        MsgBox "El campo 'Soporte' es obligatorio.", vbCritical, "Dato Faltante"
        Me.cmbSoporte.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.cmbDestino.Value) = "" Then
        MsgBox "El campo 'Destino' es obligatorio.", vbCritical, "Dato Faltante"
        Me.cmbDestino.SetFocus
        Exit Sub
    End If

    ' PREPARACIÓN DE DATOS (y valores por defecto)

    
    Dim datosManuales As Object
    Set datosManuales = CreateObject("Scripting.Dictionary")
    
    ' --- OBLIGATORIOS DIRECTOS ---
    datosManuales("Nombre") = Me.txtNombreCarpeta.Value
    datosManuales("CantidadArchivos") = Val(Me.txtCantidadArchivos.Value) ' Fojas
    datosManuales("FechaCreacion") = CDate(Me.txtFechaCreacion.Value)
    datosManuales("Serie") = Me.cmbSerie.Value
    datosManuales("Subserie") = Me.cmbSubserie.Value
    datosManuales("NumCaja") = Me.txtNumCaja.Value
    datosManuales("Soporte") = Me.cmbSoporte.Value
    datosManuales("Destino") = Me.cmbDestino.Value
        


    
    ' --- NO OBLIGATORIOS CON DEFAULT ---
    
    ' Fecha de Cierre (Default: dd/mm/aaaa)
    If IsDate(Me.txtFechaCierre.Value) Then
        datosManuales("FechaCierre") = CDate(Me.txtFechaCierre.Value)
    Else
        datosManuales("FechaCierre") = "dd/mm/aaaa"
    End If

    
    ' Observaciones (Se permite vacío)
    datosManuales("Observaciones") = Me.txtObservaciones.Value
    datosManuales("NumExpediente") = Me.txtNumExpediente.Value
    
    ' CAMPOS UBICACIÓN TOPOGRÁFICA (Default: NN)

    If Trim(Me.txtZona.Value) = "" Then
        datosManuales("Zona") = "NN"
    Else
        datosManuales("Zona") = Me.txtZona.Value
    End If
    
    ' Estanteria
    If Trim(Me.txtEstanteria.Value) = "" Then
        datosManuales("Estanteria") = "NN"
    Else
        datosManuales("Estanteria") = Me.txtEstanteria.Value
    End If
    
    ' Bandeja
    If Trim(Me.txtBandeja.Value) = "" Then
        datosManuales("Bandeja") = "NN"
    Else
        datosManuales("Bandeja") = Me.txtBandeja.Value
    End If

    

    ' ENVIAR A EXCEL

    If ExportarDatosInventario(datosManuales) Then
        MsgBox "Registro Guardado con éxito.", vbInformation
        btnLimpiar_Click
        Me.txtNumExpediente.Value = GenerarNuevoCodigoExpediente()
    Else
        MsgBox "Error al guardar la carpeta.", vbCritical
    End If
    
End Sub

