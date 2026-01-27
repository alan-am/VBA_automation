VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBusqueda 
   Caption         =   "Buscar Secciones"
   ClientHeight    =   2340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6540
   OleObjectBlob   =   "frmBusqueda.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frmBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmBusqueda
' **************************************************************************
' ! FORMULARIO DE SELECCIÓN DE UNIDAD PRODUCTORA (SECCIÓN)
' **************************************************************************
' Este formulario permite:
' 1. Seleccionar la Sección y Subsección desde la lista de configuración.
' 2. Filtrar dinámicamente las subsecciones (Lista en Cascada).
' 3. Identificar el "Código de Sección" (ej. "REC", "FIN") necesario para
'    generar los números de expediente.
' 4. Actualizar los encabezados en la hoja de Excel activa y guardar la
'    configuración global para el uso de los otros formularios.
'
' UX:
' - Soporte para teclado: Enter avanza entre campos y despliega listas.
' **************************************************************************

Option Explicit

' Variable a nivel de formulario
Private wsConfig As Worksheet

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 2
    
    ' hoja de configuración
    Set wsConfig = wskConfig
    ConfigurarCombos
    CargarSeccionesUnicas
End Sub

Private Sub ConfigurarCombos()
    ' Configurar ComboBox de SECCIÓN
    With Me.cmbSeccion
        .Clear
        .MatchEntry = fmMatchEntryComplete
        .Style = fmStyleDropDownCombo
    End With
    
    ' Configurar ComboBox de SUBSECCIÓN
    With Me.cmbSubseccion
        .Clear
        .MatchEntry = fmMatchEntryComplete
        .Style = fmStyleDropDownCombo
        .Enabled = False ' Se mantiene bloqueado hasta elegir Sección
        .BackColor = RGB(240, 240, 240) ' UX
    End With
End Sub

Private Sub CargarSeccionesUnicas()
    Dim lastRow As Long
    Dim i As Long
    Dim seccion As String
    Dim dictUnicos As Object
    
    ' diccionario para evitar duplicados en la lista desplegable
    Set dictUnicos = CreateObject("Scripting.Dictionary")
    
    lastRow = wsConfig.Cells(wsConfig.Rows.Count, "M").End(xlUp).Row
    
    ' Si no hay datos, exit
    If lastRow < 2 Then Exit Sub
    
    'Cargamos
    For i = 3 To lastRow
        seccion = Trim(wsConfig.Cells(i, "M").Value)
        
        If seccion <> "" And Not dictUnicos.Exists(seccion) Then
            dictUnicos.Add seccion, Nothing
            Me.cmbSeccion.AddItem seccion
        End If
    Next i
End Sub

' Detecto de cambios en Seccion para cargar Subsecciones
Private Sub cmbSeccion_Change()
    Dim seccionElegida As String
    Dim lastRow As Long
    Dim i As Long
    
    seccionElegida = Me.cmbSeccion.Value
    
    ' Limpiamos siempre la subsección al cambiar de padre
    Me.cmbSubseccion.Clear
    Me.cmbSubseccion.Value = ""
    
    ' Si sección vacia, bloqueamos subsección
    If seccionElegida = "" Then
        Me.cmbSubseccion.Enabled = False
        Me.cmbSubseccion.BackColor = RGB(240, 240, 240)
        Exit Sub
    End If
    
    ' Habilitamos subsección
    Me.cmbSubseccion.Enabled = True
    Me.cmbSubseccion.BackColor = vbWhite
    
    ' Filtramos las subsecciones correspondientes en la hoja Config
    lastRow = wsConfig.Cells(wsConfig.Rows.Count, "M").End(xlUp).Row
    
    For i = 2 To lastRow
        ' Si la Columna M (Sección) coincide
        If Trim(wsConfig.Cells(i, "M").Value) = seccionElegida Then
            ' Agregamos la Columna N (Subsección) a la lista
            Me.cmbSubseccion.AddItem wsConfig.Cells(i, "N").Value
        End If
    Next i
    
    ' UX/Desplegar lista automáticamente
    On Error Resume Next
    Me.cmbSubseccion.DropDown
    On Error GoTo 0
End Sub


Private Sub btnAceptar_Click()
    ' VALIDACIÓN SECCION ESCOGIDA
    If Trim(Me.cmbSeccion.Value) = "" Then
        MsgBox "Debe seleccionar una Sección.", vbExclamation
        Me.cmbSeccion.SetFocus
        Exit Sub
    End If
    
    Dim wsActiva As Worksheet
    Dim filaConfig As Long
    Dim codSeccion As String
    Dim codSubseccion As String
    Dim codSeleccionado As String
    
    ' Deficion constantes para la lógica de seleccion de codigo EXP
    Const CODIGO_DEFECTO_TABLA As String = "###"
    Const CODIGO_DESCONOCIDO As String = "???"
    
    Set wsActiva = ActiveSheet
    
    ' BUSCAR CÓDIGOS
    filaConfig = BuscarFilaConfig(Me.cmbSeccion.Value, Me.cmbSubseccion.Value)
    
    If filaConfig > 0 Then
        codSeccion = Trim(wsConfig.Cells(filaConfig, "L").Value)
        codSubseccion = Trim(wsConfig.Cells(filaConfig, "O").Value)
    Else
        codSeccion = CODIGO_DEFECTO_TABLA
        codSubseccion = CODIGO_DEFECTO_TABLA
    End If

    ' ORDEN DE PRIORIDAD DE CODIGO
    '  a. Existe Subsección seleccionada
    If Me.cmbSubseccion.Value <> "" Then
        If codSubseccion <> CODIGO_DEFECTO_TABLA And codSubseccion <> "" Then
            codSeleccionado = codSubseccion
        Else
            codSeleccionado = CODIGO_DESCONOCIDO
        End If
        
    ' b. Solo hay Sección
    Else
        If codSeccion <> CODIGO_DEFECTO_TABLA And codSeccion <> "" Then
            codSeleccionado = codSeccion
        Else
            codSeleccionado = CODIGO_DESCONOCIDO
        End If
    End If
    ' ------------------------

    On Error Resume Next
    
    ' Escribir la seleccion en la tabla
    wsActiva.Range("E5").Value = Me.cmbSeccion.Value
    wsActiva.Range("E6").Value = Me.cmbSubseccion.Value
    
    ' Guardado de codigo en la hoja config celda Q2
    wskConfig.Range("Q2").Value = codSeleccionado
    
    If Err.Number <> 0 Then
        MsgBox "Error al escribir los datos.", vbCritical
        Err.Clear
    End If
    On Error GoTo 0
    
    Unload Me
End Sub


' UX eventos de teclado
Private Sub cmbSeccion_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        ' Al dar Enter en Sección, pasamos el foco a Subsección
        Me.cmbSubseccion.SetFocus
        On Error Resume Next
        Me.cmbSubseccion.DropDown
    End If
End Sub

Private Sub cmbSubseccion_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        ' Al dar Enter en Subsección, equivale a Aceptar
        btnAceptar_Click
    End If
End Sub


Private Sub btnCancelar_Click()
    Unload Me
End Sub
