VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBusqueda 
   Caption         =   "Buscar Sección"
   ClientHeight    =   1650
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

Private Sub cmbBuscador_Change()

End Sub

Private Sub cmbBuscador_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ' Evita el sonido de "beep"
        KeyCode = 0
        
        ' Llama a la acción del botón aceptar
        btnAceptar_Click
    End If
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    Me.cmbBuscador.DropDown
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 2
    CargarOpciones
End Sub


'Carga las opciones de columna "Seccion Documental" en la hoja config
Private Sub CargarOpciones()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rango As Range
    
    Set ws = ThisWorkbook.Sheets("Config")
    
    ' limpieza del campo por seguridad
    Me.cmbBuscador.Clear
    
    ' Encuentra la última fila en la Columna
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then Exit Sub ' No hay datos
    
    ' Carga los datos al ComboBox
    Me.cmbBuscador.List = ws.Range("A2:A" & lastRow).Value
    
    ' CONFI DE AUTOCOMPLETADO
    With Me.cmbBuscador
        .SetFocus
        .MatchEntry = fmMatchEntryComplete ' Autocompleta mientras se escribe
        .Style = fmStyleDropDownCombo      ' Permite escribir
        '.DropDown                          ' Despliega la lista automáticamente al iniciar
    End With
End Sub

Private Sub btnAceptar_Click()
    ' Validación de seleccion
    If Trim(Me.cmbBuscador.Value) = "" Then
        MsgBox "Por favor seleccione una opción o presione Cancelar.", vbExclamation
        Exit Sub
    End If
    
    
    On Error Resume Next
    ActiveSheet.Range("E5").Value = Me.cmbBuscador.Value 'usa hoja actual
    
    If Err.Number <> 0 Then
        MsgBox "No se pudo escribir en la celda E5, puede haber errores de modificación del archivo.", vbCritical
        Err.Clear
    End If
    On Error GoTo 0
    
    Unload Me
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub


