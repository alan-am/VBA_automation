VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatosMasivos 
   Caption         =   "Procesamiento automático"
   ClientHeight    =   4350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7455
   OleObjectBlob   =   "frmDatosMasivos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmDatosMasivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmDatosMasivos
Option Explicit

' Cola de carpetas a procesar
Private ColaCarpetas As Collection
Private COLOR_BOTON_ACTIVO As Long
Private COLOR_BOTON_INACTIVO As Long

Private Sub UserForm_Initialize()
    'Asignacion colores boton
    COLOR_BOTON_ACTIVO = RGB(31, 73, 125) ' Azul Oscuro
    COLOR_BOTON_INACTIVO = RGB(160, 160, 160) ' Gris

    CargarListasDinamicas
    ' Valores por defecto
    Me.txtNumCaja.Value = 0
    Me.cmbDestino.Value = "Conservación"
    Me.cmbSoporte.Value = "Digital"
    Me.btnProcesarLote.Enabled = False ' Deshabilitado
    Me.btnProcesarLote.BackColor = COLOR_BOTON_INACTIVO
    
End Sub


Private Sub btnSeleccionarCarpeta_Click()
    Dim folderPath As String
    Dim fso As Object, carpetaMadre As Object, subCarpeta As Object
    

    folderPath = SeleccionarCarpeta()    ' modUtilidades
    
    If folderPath = "" Then Exit Sub
    
    Me.txtRutaCarpeta.Value = folderPath
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpetaMadre = fso.GetFolder(folderPath)
    
    ' VERIFICAR SUBCARPETAS
    If carpetaMadre.SubFolders.Count = 0 Then
        MsgBox "La carpeta seleccionada NO contiene subcarpetas." & vbCrLf & _
               "Este modo es para procesar lotes. Use el botón 'Carpeta Digital' para archivos individuales.", vbExclamation
        Me.lblEstado.Caption = "Estado: 0 subcarpetas encontradas."
        Me.lblEstado.ForeColor = RGB(200, 0, 0)
        Me.btnProcesarLote.Enabled = False
        Me.btnProcesarLote.BackColor = COLOR_BOTON_INACTIVO
        Exit Sub
    End If
    
    ' LLENAR COLA
    Set ColaCarpetas = New Collection
    For Each subCarpeta In carpetaMadre.SubFolders
        ColaCarpetas.Add subCarpeta.Path
    Next subCarpeta
    
    ' Actu UX
    Me.lblEstado.Caption = "Estado: " & ColaCarpetas.Count & " subcarpetas encontradas."
    Me.lblEstado.ForeColor = RGB(0, 100, 0) ' Verde oscuro
    Me.btnProcesarLote.Enabled = True
    Me.btnProcesarLote.BackColor = COLOR_BOTON_ACTIVO
    
End Sub

Private Sub btnProcesarLote_Click()

    ' Validaciones
    If Trim(Me.cmbSerie.Value) = "" Or Trim(Me.cmbSubserie.Value) = "" Then
        MsgBox "Serie y Subserie son obligatorias.", vbExclamation
        Exit Sub
    End If
    
    ' Confirmación
    If MsgBox("Se procesarán " & ColaCarpetas.Count & " carpetas con:" & vbCrLf & _
              "Serie: " & Me.cmbSerie.Value & vbCrLf & _
              "¿Continuar?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    'Deshabilitacion boton Empezar
    'UX- cambio a gris oscuro
    Me.btnProcesarLote.BackColor = COLOR_BOTON_INACTIVO
    Me.btnProcesarLote.Enabled = False
    Me.Repaint ' Forzar actualización visual
    
    On Error GoTo ManejoError
    ' Procesamiento
    Dim rutaActual As String
    Dim infoCarpeta As Object
    Dim contador As Long
    Dim i As Long
    

    Application.ScreenUpdating = False
    
    For i = 1 To ColaCarpetas.Count
        rutaActual = ColaCarpetas(i)
        
        ' Obtener datos variables de la carpeta real
        Set infoCarpeta = ObtenerInfoCarpeta(rutaActual)
        
        ' datos fijos
        infoCarpeta("Serie") = Me.cmbSerie.Value
        infoCarpeta("Subserie") = Me.cmbSubserie.Value
        infoCarpeta("Destino") = Me.cmbDestino.Value
        infoCarpeta("Soporte") = Me.cmbSoporte.Value
        infoCarpeta("NumCaja") = Me.txtNumCaja.Value
        infoCarpeta("Observaciones") = " "
        
        ' Generar N° Expediente
        infoCarpeta("NumExpediente") = GenerarNuevoCodigoExpediente()
        
        ' Valores Default
        infoCarpeta("UbicacionTopografica") = "NN"
        infoCarpeta("Zona") = "NN"
        infoCarpeta("Estanteria") = "NN"
        infoCarpeta("Bandeja") = "NN"
        
        ' Guardar
        If ExportarDatosInventario(infoCarpeta) Then
            contador = contador + 1
        End If
        
        Set infoCarpeta = Nothing
        DoEvents
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "Proceso finalizado. Registros creados: " & contador, vbInformation
    Unload Me
    Exit Sub
    
ManejoError:
    Application.ScreenUpdating = True
    Me.btnProcesarLote.BackColor = COLOR_BOTON_ACTIVO
    Me.btnProcesarLote.Enabled = True
    MsgBox "Ocurrió un error inesperado durante el procesamiento: " & Err.Description, vbCritical

End Sub

' CARGA DE LISTAS
Private Sub CargarListasDinamicas()
     Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config")
    
    ' Serie
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    For i = 3 To lastRow
        If Trim(ws.Cells(i, "I").Value) <> "" Then Me.cmbSerie.AddItem ws.Cells(i, "I").Value
    Next i
    ' Subserie
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    For i = 3 To lastRow
        If Trim(ws.Cells(i, "J").Value) <> "" Then Me.cmbSubserie.AddItem ws.Cells(i, "J").Value
    Next i
    ' Destino
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    For i = 3 To lastRow
        If Trim(ws.Cells(i, "G").Value) <> "" Then Me.cmbDestino.AddItem ws.Cells(i, "G").Value
    Next i
    ' Soporte
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    For i = 3 To lastRow
        If Trim(ws.Cells(i, "H").Value) <> "" Then Me.cmbSoporte.AddItem ws.Cells(i, "H").Value
    Next i
End Sub

Private Sub btnCerrar_Click()
    Unload Me
End Sub
