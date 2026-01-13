VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatosCarpeta 
   Caption         =   "Gestor de Carpetas Digitales"
   ClientHeight    =   7305
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
'frmDatosCarpeta (Digital)

' Variable a nivel de formulario para guardar los datos de la carpeta
Private pDatosCarpeta As Object    ' info de carpeta en proceso
Private ColaCarpetas As Collection ' Cola de rutas pendientes
Private ModoMasivo As Boolean      ' Bandera para modo de flujo



' Metodo de inicializacion del forms
Private Sub UserForm_Initialize()
    ' Carga de las listas dinámicas
    CargarListasDinamicas
    
    ' Seteado valores default de cierto campos
    Me.txtNumCaja.Value = 0
    Me.cmbDestino.Value = "Conservación"
    Me.cmbSoporte.Value = "Digital"
    Me.txtFechaCierre.Value = "dd/mm/aaaa"
    
    'Pre-llenar el N° Expediente
    Me.txtNumExpediente.Value = GenerarNuevoCodigoExpediente()
End Sub
Private Sub btnCerrar_Click()
    Unload Me
End Sub


Private Sub btnSeleccionarCarpeta_Click()
    
    Dim folderPath As String
    Dim fso As Object, carpetaMadre As Object, subCarpeta As Object
    Dim respuesta As VbMsgBoxResult
    
    'muestra del dialogo de seleccion
    folderPath = SeleccionarCarpeta()
    
    If folderPath = "" Then Exit Sub
    
    Set fso = CreateObject("Scripting.FileSystemObject") 'inicializacion objeto Carpeta
    Set carpetaMadre = fso.GetFolder(folderPath)
    
    'Verificacion de disponibilidad de subcarpetas y pregunta
    If carpetaMadre.SubFolders.Count > 0 Then
        respuesta = MsgBox("La carpeta seleccionada contiene " & carpetaMadre.SubFolders.Count & " subcarpetas." & vbCrLf & vbCrLf & _
                           "¿Desea activar el 'Modo Secuencial' para procesarlas continuamente?" & vbCrLf & _
                           "SÍ: Carga la primera subcarpeta y prepara la cola." & vbCrLf & _
                           "NO: Analiza solo la carpeta seleccionada (comportamiento normal).", _
                           vbYesNo + vbQuestion, "Modo de Análisis")
        
        If respuesta = vbYes Then
                   ' Activar bandera modo flujo
                   ModoMasivo = True
                   Set ColaCarpetas = New Collection
                   
                   ' Llenar la cola con las subcarpetas
                   For Each subCarpeta In carpetaMadre.SubFolders
                       ColaCarpetas.Add subCarpeta.Path
                   Next subCarpeta
                   
                   MsgBox "Se han puesto en cola " & ColaCarpetas.Count & " carpetas. Empecemos.", vbInformation
                   
                   ' Cargar la primera de la lista
                   CargarSiguienteDeLaCola
                   Exit Sub
        End If
    End If
    
    'Seteo modo normal y procesado de la carpeta elegida
    ModoMasivo = False
    Set ColaCarpetas = Nothing
    ProcesarCarpetaIndividual folderPath
    
    'If folderPath <> "" Then
    ' Obtiene el diccionario y lo guarda en la variable del formulario
        'Set pDatosCarpeta = ObtenerInfoCarpeta(folderPath) ' modUtilidades
        
        'MostrarDatosCarpeta pDatosCarpeta 'modInicio
    'End If
    
    
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
    MsgBox "Error al cargar las listas en configuración." & vbCrLf & _
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
        MsgBox "Expediente '" & pDatosCarpeta("Nombre") & "' guardado con éxito.", vbInformation, "Exportación Completa"
        
        If ModoMasivo Then
            ' Se mantiene Serie, Subserie, Destino, Soporte, Caja, etc.
            ' Solo limpiamos los datos específicos de la carpeta anterior.
            LimpiarSoloDatosVariables
            
            ' Cargar inmediatamente la siguiente
            CargarSiguienteDeLaCola
        Else
            ' MODO NORMAL
            LimpiarFormulario
            Me.txtNumExpediente.Value = GenerarNuevoCodigoExpediente()
            Set pDatosCarpeta = Nothing
        End If
    Else
        MsgBox "Ocurrió un error al intentar guardar los datos en la hoja de Excel.", vbCritical, "Error de Exportación"
    End If
    
    
    
End Sub

'Funcion de gestion de avance de modo masivo
Private Sub CargarSiguienteDeLaCola()
    Dim siguienteRuta As String
    
    If ColaCarpetas.Count > 0 Then
        'Procesar siguiente ruta(subcarpeta)
        siguienteRuta = ColaCarpetas(1)
        
        ' Actualizar cola
        ColaCarpetas.Remove 1
        
        ' Procesar
        ProcesarCarpetaIndividual siguienteRuta
        
        ' FUTURO UX: Avisar visualmente
        ' Podría poner un Label en el form que diga "Procesando carpeta..."
    Else
        ' Se acabó la cola
        MsgBox "¡Proceso terminado! Se han analizado todas las subcarpetas.", vbInformation
        ModoMasivo = False
        Me.Caption = "Gestor de Carpetas Digitales"
        LimpiarFormulario
        Unload Me
    End If
End Sub


' Funcion auxiliar para procesar una ruta específica
Private Sub ProcesarCarpetaIndividual(ruta As String)
    ' Obtencion de datos y guardado
    Set pDatosCarpeta = ObtenerInfoCarpeta(ruta)
    MostrarDatosCarpeta pDatosCarpeta ' modInicio
    
    ' Generar Código de Expediente Sugerido
    Me.txtNumExpediente.Value = GenerarNuevoCodigoExpediente()
    
    ' Actualizar título para UX
    If ModoMasivo Then
        Me.Caption = "Gestor Digital - Pendientes: " & ColaCarpetas.Count
        Me.btnOmitir.Visible = True
    Else
        Me.Caption = "Gestor de Carpetas Digitales"
        Me.btnOmitir.Visible = False
    End If
End Sub

' --- EN frmDatosCarpeta (Nuevo método) ---

Private Sub LimpiarSoloDatosVariables()
    ' Borramos solo lo que cambia de carpeta a carpeta
    Me.txtNombreCarpeta.Value = ""
    Me.txtRutaCarpeta.Value = ""
    Me.txtCantidadArchivos.Value = ""
    Me.txtTamanoTotal.Value = ""
    Me.txtObservaciones.Value = ""
    Me.txtFechaCreacion.Value = ""
    Me.txtFechaCierre.Value = "dd/mm/aaaa"

    ' lo demas se mantiene(excepto N° expediente)
End Sub

Private Sub btnOmitir_Click()
    If ModoMasivo Then
        LimpiarSoloDatosVariables
        CargarSiguienteDeLaCola
    End If
End Sub


'DESHABILITADO
Private Sub btnLimpiar_Click()
    LimpiarFormulario 'modUtilidades
End Sub
