VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatosCarpeta 
   Caption         =   "Gestor de Carpetas Digitales"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7455
   OleObjectBlob   =   "frmDatosCarpeta.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frmDatosCarpeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmDatosCarpeta
' **************************************************************************
' ! FORMULARIO DE GESTIÓN DE CARPETAS DIGITALES
' **************************************************************************
' funcionalidades formulario:
' 1. Seleccionar una carpeta digital individualmente.
' 2. Visualizar sus metadatos (Nombre, Fecha, Fojas, Peso).
' 3. Completar información archivística (Serie, Subserie, Caja).
' 4. Generar automáticamente el código de expediente.
' 5. Guardar el registro en la hoja Excel principal.
'
' UX:
' - El botón "Insertar" permanece bloqueado hasta que se selecciona una carpeta válida.
' - Se bloquea visualmente durante el proceso de guardado para evitar duplicados.
' **************************************************************************
Option Explicit

Private pDatosCarpeta As Object    ' info de carpeta en proceso
Private COLOR_BOTON_ACTIVO As Long
Private COLOR_BOTON_INACTIVO As Long

Private Sub UserForm_Initialize()
    COLOR_BOTON_ACTIVO = RGB(31, 73, 125)   ' Azul Oscuro
    COLOR_BOTON_INACTIVO = RGB(160, 160, 160)
    
    
    CargarListasDinamicas
    
    ' Seteado valores default de cierto campos
    Me.txtNumCaja.Value = 0
    Me.cmbDestino.Value = "Conservación"
    Me.cmbSoporte.Value = "Digital"
    Me.txtFechaCierre.Value = "dd/mm/aaaa"
    
    'Pre-llenar el N° Expediente
    Me.txtNumExpediente.Value = GenerarNuevoCodigoExpediente()
    
    Me.btnInsertar.Enabled = False
    Me.btnInsertar.BackColor = COLOR_BOTON_INACTIVO
End Sub

Private Sub btnSeleccionarCarpeta_Click()
    Dim folderPath As String
    
    ' Dialogo de seleccion
    folderPath = SeleccionarCarpeta()
    
    If folderPath <> "" Then
        Set pDatosCarpeta = ObtenerInfoCarpeta(folderPath)
        MostrarDatosCarpeta pDatosCarpeta
        
        Me.txtNumExpediente.Value = GenerarNuevoCodigoExpediente()
        
        Me.btnInsertar.Enabled = True
        Me.btnInsertar.BackColor = COLOR_BOTON_ACTIVO
    End If
End Sub



Private Sub btnInsertar_Click()
    
    ' Validar que los datos de la carpeta no esten vacios
    If pDatosCarpeta Is Nothing Then
        MsgBox "Error: Primero debe seleccionar una carpeta usando el botón 'Examinar...'.", vbCritical, "Acción Requerida"
        Me.btnSeleccionarCarpeta.SetFocus
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
    
    
    ' UX
    Me.btnInsertar.Enabled = False
    Me.btnInsertar.BackColor = COLOR_BOTON_INACTIVO
    Me.Repaint ' actualización visual
    
    
    On Error GoTo ManejoError

    ' PREPARAR DATOS
    pDatosCarpeta("Nombre") = Me.txtNombreCarpeta.Value
    pDatosCarpeta("Ruta") = Me.txtRutaCarpeta.Value
    pDatosCarpeta("CantidadArchivos") = Me.txtCantidadArchivos.Value
    pDatosCarpeta("TamanoTotal") = Me.txtTamanoTotal.Value
    
    pDatosCarpeta("Serie") = Me.cmbSerie.Value
    pDatosCarpeta("Subserie") = Me.cmbSubserie.Value
    pDatosCarpeta("NumExpediente") = Me.txtNumExpediente.Value
    pDatosCarpeta("Destino") = Me.cmbDestino.Value
    pDatosCarpeta("Soporte") = Me.cmbSoporte.Value
    pDatosCarpeta("Observaciones") = Me.txtObservaciones.Value
    pDatosCarpeta("NumCaja") = IIf(IsNumeric(Me.txtNumCaja.Value), CLng(Me.txtNumCaja.Value), 0)
    
    ' Fechas
    If IsDate(Me.txtFechaCreacion.Value) Then
        pDatosCarpeta("FechaCreacion") = CDate(Me.txtFechaCreacion.Value)
    Else
        pDatosCarpeta("FechaCreacion") = "dd/mm/aaaa"
    End If
    
    If IsDate(Me.txtFechaCierre.Value) Then
        pDatosCarpeta("FechaCierre") = CDate(Me.txtFechaCierre.Value)
    Else
        pDatosCarpeta("FechaCierre") = "dd/mm/aaaa"
    End If
    
    ' Ubicación por defecto para carpetas digitales
    pDatosCarpeta("Zona") = "NN"
    pDatosCarpeta("Estanteria") = "NN"
    pDatosCarpeta("Bandeja") = "NN"

    ' GUARDAR
    If ExportarDatosInventario(pDatosCarpeta) Then
        MsgBox "Expediente guardado con éxito.", vbInformation
        
        ' Limpiar para el siguiente ingreso manual
        LimpiarFormulario
        Me.txtNumExpediente.Value = GenerarNuevoCodigoExpediente()
        Set pDatosCarpeta = Nothing
        Me.btnInsertar.Enabled = False
        Me.btnInsertar.BackColor = COLOR_BOTON_INACTIVO
    Else
        GoTo ManejoError
    End If
    Exit Sub
ManejoError:
        MsgBox "Ocurrió un error al intentar guardar los datos en la hoja de Excel.", vbCritical, "Error de Registro" & Err.Description, vbCritical
        Me.btnInsertar.Enabled = True
        Me.btnInsertar.BackColor = COLOR_BOTON_ACTIVO
End Sub

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

    ' Cargar Serie Documental(Columna I)
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    For i = 3 To lastRow
        If Trim(ws.Cells(i, "I").Value) <> "" Then
            Me.cmbSerie.AddItem ws.Cells(i, "I").Value
        End If
    Next i
    
    ' Cargar Subserie Documental(Columna J)
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    For i = 3 To lastRow
        If Trim(ws.Cells(i, "J").Value) <> "" Then
            Me.cmbSubserie.AddItem ws.Cells(i, "J").Value
        End If
    Next i
    
    ' Cargar Destino Final (Columna G)
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    For i = 3 To lastRow
        If Trim(ws.Cells(i, "G").Value) <> "" Then
            Me.cmbDestino.AddItem ws.Cells(i, "G").Value
        End If
    Next i
    
    ' Cargar Soporte (Columna H)
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    For i = 3 To lastRow
        If Trim(ws.Cells(i, "H").Value) <> "" Then
            Me.cmbSoporte.AddItem ws.Cells(i, "H").Value
        End If
    Next i
    
    Exit Sub

ErrorHandler:
    MsgBox "Error al cargar las listas de la hoja de configuración." & vbCrLf & _
           "Asegúrese que la hoja 'Config' existe y tiene el formato correcto.", _
           vbCritical, "Error de Carga"
End Sub

' Método para llenar los campos visuales con la info del diccionario
Private Sub MostrarDatosCarpeta(info As Object)
    Me.txtRutaCarpeta.Value = info("Ruta")
    Me.txtNombreCarpeta.Value = info("Nombre")
    Me.txtFechaCreacion.Value = info("FechaCreacion")
    Me.txtCantidadArchivos.Value = info("CantidadArchivos")
    Me.txtTamanoTotal.Value = info("TamanoTotal")
    Me.txtFechaCierre.Value = info("FechaCierre")
End Sub

Private Sub LimpiarFormulario()
    Me.txtRutaCarpeta.Value = ""
    Me.txtNombreCarpeta.Value = ""
    Me.txtFechaCreacion.Value = ""
    Me.txtCantidadArchivos.Value = ""
    Me.txtTamanoTotal.Value = ""
    Me.txtObservaciones.Value = ""
    Me.txtFechaCierre.Value = "dd/mm/aaaa"
End Sub

Private Sub btnCerrar_Click()
    Unload Me
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFormulario
End Sub




