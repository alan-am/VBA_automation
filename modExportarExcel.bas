Attribute VB_Name = "modExportarExcel"

' Funcion que escribe los datos de la carpeta en una hoja excel
Function ExportarDatosInventario(datos As Object) As Boolean
    'VAR NOMBRE TABLA
    Dim nombreTabla As String
    nombreTabla = "tabla_test89"
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
    Set wsInventario = ThisWorkbook.Sheets("Inventario General")
    Set wsConfig = ThisWorkbook.Sheets("Config")
    Set tbl = wsInventario.ListObjects(nombreTabla)
    
    ' -- Lectura de valores
    minAltura = Val(wsConfig.Range("D2").Value)
    nombreFuente = wsConfig.Range("E2").Value
    tamanoFuente = Val(wsConfig.Range("F2").Value)
    
    ' --- Valores por default si la config está vacía ---
    If minAltura < 15 Then minAltura = 15 ' Un mínimo de 15
    If nombreFuente = "" Then nombreFuente = "Calibri"
    If tamanoFuente < 8 Then tamanoFuente = 8
    
    
    ' AÑADIR UNA NUEVA FILA A LA TABLA
    ' se añade la fila al final de la tabla
    ' y "empuja" cualquier contenido de abajo
    Set newRow = tbl.ListRows.Add(AlwaysInsert:=True)
    
    ' Escritura de datos en la nueva fila
    ' Usamos .Range(columna_numero) para escribir en la celda
    
    ' Col 1: SERIE/SUBSERIE DOCUMENTAL
    newRow.Range(1).Value = datos("SerieSubserie")
    
    ' Col 2: N° CAJA
    newRow.Range(2).Value = datos("NumCaja")
    
    ' Col 3: N° DE EXPEDIENTE
    newRow.Range(3).Value = datos("NumExpediente")
    
    ' Col 4: NOMBRE DEL EXPEDIENTE
    newRow.Range(4).Value = datos("Nombre")
    
    ' Col 5: FECHAS EXTREMAS - APERTURA
    newRow.Range(5).Value = datos("FechaCreacion")
    
    ' Col 6: FECHAS EXTREMAS - CIERRE
    newRow.Range(6).Value = datos("FechaCierre")
    
    ' Col 7: FOJAS
    newRow.Range(7).Value = datos("CantidadArchivos")
    
    ' Col 8: DESTINO FINAL
    newRow.Range(8).Value = datos("Destino")
    
    ' Col 9: SOPORTE
    newRow.Range(9).Value = datos("Soporte")
    
    ' Col 10: UBICACIÓN TOPOGRÁFICA - ZONA
    newRow.Range(10).Value = datos("UbicacionTopografica")
    
    ' Col 11: UBICACIÓN TOPOGRÁFICA - ESTANTE
    newRow.Range(11).Value = datos("UbicacionTopografica")
    
    ' Col 12: UBICACIÓN TOPOGRÁFICA - BANDEJA
    newRow.Range(12).Value = datos("UbicacionTopografica")
    
    ' Col 13: OBSERVACIONES
    newRow.Range(13).Value = datos("Observaciones")
    
    ' --- APLICACION DE FORMATO A LA NUEVA FILA ---
    With newRow.Range

        .Interior.Color = RGB(255, 255, 255) ' Blanco
        .Borders.LineStyle = xlContinuous ' Añade bordes
        .Borders.Weight = xlThin ' Borde delgado
        
        ' Formato de texto y alineación ---
        ' Centra el texto verticalmente en las celdas
        '.VerticalAlignment = xlCenter
        ' Permite que el texto se ajuste y la fila crezca
        .WrapText = True
        
        ' fuente y tamano
        .Font.Name = nombreFuente
        .Font.Size = tamanoFuente
        
        'Altura Mínima ---
        ' 1. Deja que Excel autoajuste la fila según el contenido
        .EntireRow.AutoFit
        
        ' Compara con la variable leída de Config
        If .RowHeight < minAltura Then
            .RowHeight = minAltura
        End If
    End With
    ' ------------------------------------------------
    
    ' 4. Si todo salió bien, devuelve True
    ExportarDatosInventario = True
    Exit Function

'manejo de errores
ManejoError:
    MsgBox "Error al exportar: " & Err.Description & vbCrLf & _
           "Asegúrese que la hoja 'Inventario General' existe " & _
           "y que la tabla se llama" & nombreTabla & ".", vbCritical, "Error en 'modExportarExcel'"
    ExportarDatosInventario = False
End Function

    ' ------ datos de prueba adicionales que no van en la plantilla final ------
    'ws.Cells(lRow, 14).Value = "Ruta: " & datos("Ruta")
    'ws.Cells(lRow, 15).Value = "Tamaño: " & datos("TamanoTotal") & " KB"
