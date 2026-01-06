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
Option Explicit

' Variable a nivel de formulario
Private wsConfig As Worksheet

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 2
    
    ' hoja de configuración
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    ' Configuración visual inicial
    ConfigurarCombos
    
    ' Cargar lista de Secciones(Columna M)
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
        .Enabled = False ' Se mantiene bloqueado hasta que elijan Sección
        .BackColor = RGB(240, 240, 240) ' Gris visual para indicar "inactivo"
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

' BOTONES
Private Sub btnAceptar_Click()
    ' VALIDACIÓN Seccion obligatoria
    If Trim(Me.cmbSeccion.Value) = "" Then
        MsgBox "Debe seleccionar una Sección.", vbExclamation
        Me.cmbSeccion.SetFocus
        Exit Sub
    End If
    
    
    ' ESCRITURA EN LA HOJA
    Dim wsActiva As Worksheet
    Set wsActiva = ActiveSheet
    
    On Error Resume Next
    ' Escribir Sección (Celda E5)
    wsActiva.Range("E5").Value = Me.cmbSeccion.Value
    
    ' Escribir Subsección (Celda E6) - Puede ir vacía
    wsActiva.Range("E6").Value = Me.cmbSubseccion.Value
    
    ' =========================================================================
    ' LÓGICA FUTURA PARA CÓDIGOS DE EXPEDIENTE
    
    ' Aquí irá la lógica para buscar los acrónimos en Columnas M y O
    ' y guardarlos en celdas ocultas (ej. Z5 y Z6).
    
    ' Dim filaEncontrada As Long
    ' filaEncontrada = BuscarFilaConfig(Me.cmbSeccion.Value, Me.cmbSubseccion.Value)
    '
    ' If filaEncontrada > 0 Then
    '     wsActiva.Range("Z5").Value = wsConfig.Cells(filaEncontrada, "M").Value ' Cód Sección
    '     wsActiva.Range("Z6").Value = wsConfig.Cells(filaEncontrada, "O").Value ' Cód Subsección
    ' End If
    ' =========================================================================
    
    If Err.Number <> 0 Then
        MsgBox "Error al escribir en la hoja.", vbCritical
        Err.Clear
    End If
    On Error GoTo 0
    
    Unload Me
End Sub

Private Sub btnCancelar_Click()
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

' ------------------------------------------------------------------------
' FUNCIONES AUXILIARES FUTURAS

Private Function BuscarFilaConfig(sec As String, subSec As String) As Long
    Dim i As Long
    Dim lastRow As Long
    lastRow = wsConfig.Cells(wsConfig.Rows.Count, "M").End(xlUp).Row
    
    For i = 2 To lastRow
        ' Si la subsección está vacía, buscamos solo por sección (la primera coincidencia)
        If subSec = "" Then
             If wsConfig.Cells(i, "M").Value = sec Then
                BuscarFilaConfig = i
                Exit Function
             End If
        Else
            ' Si hay subsección, buscamos coincidencia exacta de ambos
            If wsConfig.Cells(i, "M").Value = sec And wsConfig.Cells(i, "N").Value = subSec Then
                BuscarFilaConfig = i
                Exit Function
            End If
        End If
    Next i
    BuscarFilaConfig = 0
End Function
