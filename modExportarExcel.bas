Attribute VB_Name = "modExportarExcel"
'modExportarExcel
' **************************************************************************
' ! LÓGICA DE ESCRITURA EN LA HOJA PRINCIPAL
' **************************************************************************

' Funcion que escribe los datos de la carpeta en una hoja excel
Function ExportarDatosInventario(datos As Object) As Boolean
    'VAR NOMBRE TABLA
    Dim nombreTabla As String
    nombreTabla = "tabla_inventario"
    ' Manejador de errores
    On Error GoTo ManejoError

    Dim wsInventario As Worksheet
    Dim wsConfig As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    
    ' --- Variables para formato dinámico(por la hoja Config) ---
    Dim minAltura As Double
    Dim nombreFuente As String
    Dim tamanoFuente As Double
    
    ' Definir hojas y tabla a editar/usar
    Set wsInventario = wskInventario
    Set wsConfig = wskConfig
    Set tbl = wsInventario.ListObjects(nombreTabla)
    
    ' -- Lectura de valores
    minAltura = Val(wsConfig.Range("C3").Value)
    nombreFuente = wsConfig.Range("D3").Value
    tamanoFuente = Val(wsConfig.Range("E3").Value)
    
    ' --- Valores por default si la config está vacía ---
    If minAltura < 15 Then minAltura = 15 ' Un mínimo de 15
    If nombreFuente = "" Then nombreFuente = "Calibri"
    If tamanoFuente < 8 Then tamanoFuente = 8
    
    ' AÑADIR UNA NUEVA FILA A LA TABLA
    Set newRow = tbl.ListRows.Add(AlwaysInsert:=True)
    
    ' Escritura de datos en la nueva fila
    ' Col 1: SERIE DOCUMENTAL
    newRow.Range(1).Value = datos("Serie")
    
    ' Col 2: SUBSERIE DOCUMENTAL
    newRow.Range(2).Value = datos("Subserie")
    
    ' Col 3: N° CAJA
    newRow.Range(3).Value = datos("NumCaja")
    
    ' Col 4: N° DE EXPEDIENTE
    newRow.Range(4).NumberFormat = "@"
    newRow.Range(4).Value = datos("NumExpediente")
    
    ' Col 5: NOMBRE DEL EXPEDIENTE
    newRow.Range(5).Value = datos("Nombre")
    
    ' Col 6: FECHAS EXTREMAS - APERTURA
    newRow.Range(6).Value = datos("FechaCreacion")
    
    ' Col 7: FECHAS EXTREMAS - CIERRE
    newRow.Range(7).Value = datos("FechaCierre")
    
    ' Col 8: FOJAS
    newRow.Range(8).NumberFormat = "0"
    newRow.Range(8).Value = Val(datos("CantidadArchivos"))
    
    ' Col 9: DESTINO FINAL
    newRow.Range(9).Value = datos("Destino")
    
    ' Col 10: SOPORTE
    newRow.Range(10).Value = datos("Soporte")
    
    ' Col 11: UBICACIÓN TOPOGRÁFICA - ZONA
    If datos.Exists("Zona") Then
        newRow.Range(11).Value = datos("Zona")
    Else
        newRow.Range(11).Value = "NN" ' Default digital
    End If
    
    ' Col 12: UBICACIÓN TOPOGRÁFICA - ESTANTE
    If datos.Exists("Estanteria") Then
        newRow.Range(12).Value = datos("Estanteria")
    Else
        newRow.Range(12).Value = "NN"
    End If
    
    ' Col 13: UBICACIÓN TOPOGRÁFICA - BANDEJA
    If datos.Exists("Bandeja") Then
        newRow.Range(13).Value = datos("Bandeja")
    Else
        newRow.Range(13).Value = "NN"
    End If
    
    ' Col 14: OBSERVACIONES
    newRow.Range(14).Value = datos("Observaciones")
    
    ' --- APLICACION DE FORMATO A LA NUEVA FILA ---
    With newRow.Range

        .Interior.Color = RGB(255, 255, 255) ' Blanco
        .Borders.LineStyle = xlContinuous ' Añade bordes
        .Borders.Weight = xlThin ' Borde delgado
        
        ' Formato de texto y alineación
        .WrapText = True
        .Font.Name = nombreFuente
        .Font.Size = tamanoFuente
        
        'Altura Mínima --
        .EntireRow.AutoFit
        
        ' Compara con la variable leída de Config
        If .RowHeight < minAltura Then
            .RowHeight = minAltura
        End If
    End With
    ' ------------------------------------------------
    
    ExportarDatosInventario = True
    Exit Function

'manejo de errores
ManejoError:
    MsgBox "Error al exportar: " & Err.Description & vbCrLf & _
           "Asegúrese que la hoja 'Inventario General' existe " & _
           "y que la tabla se llama" & nombreTabla & ".", vbCritical, "Error en 'modExportarExcel'"
    ExportarDatosInventario = False
End Function
      
Function GenerarNuevoCodigoExpediente() As String
    ' Esta función calcula el código visualmente para el formulario
    
    Dim wsInventario As Worksheet
    Dim tbl As ListObject
    Dim codSeleccionado As String
    Dim siguienteNumero As Long
    
    ' Definición de Constantes
    Const FORMATO_PREFIJO As String = "ESPOL-"
    Const CODIGO_DESCONOCIDO As String = "???"
    Const NOMBRE_TABLA As String = "tabla_inventario"
    
    ' Referencias
    Set wsInventario = wskInventario
    On Error Resume Next
    Set tbl = wsInventario.ListObjects(NOMBRE_TABLA)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        GenerarNuevoCodigoExpediente = "Error-Tabla"
        Exit Function
    End If
    
    ' Leer Código de Sección
    codSeleccionado = Trim(wskConfig.Range("Q2").Value)
    If codSeleccionado = "" Then codSeleccionado = CODIGO_DESCONOCIDO
    
    ' cantidad filas actual + 1
    siguienteNumero = tbl.ListRows.Count + 1
    
    ' Armar el String
    GenerarNuevoCodigoExpediente = FORMATO_PREFIJO & codSeleccionado & "-" & Format(siguienteNumero, "000")

End Function


Public Function BuscarFilaConfig(sec As String, subSec As String) As Long
    Dim wsConfig As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Set wsConfig = wskConfig
    
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
