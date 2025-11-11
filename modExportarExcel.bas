Attribute VB_Name = "modExportarExcel"

' Funcion que escribe los datos de la carpeta en una hoja excel
Function ExportarDatosInventario(datos As Object) As Boolean
    'VAR NOMBRE TABLA
    Dim nombreTabla As String
    nombreTabla = "tabla_test8910"
    ' Manejador de errores
    On Error GoTo ManejoError

    Dim wsInventario As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    
    ' Definir hoja y tabla a editar
    Set wsInventario = ThisWorkbook.Sheets("Inventario General")
    Set tbl = wsInventario.ListObjects(nombreTabla)
    
    ' AÑADIR UNA NUEVA FILA A LA TABLA
    ' se añade la fila al final de la tabla
    ' y "empuja" cualquier contenido de abajo
    Set newRow = tbl.ListRows.Add(AlwaysInsert:=True)
    
    With newRow.Range
        .Interior.Color = RGB(255, 255, 255) ' Blanco
        .Borders.LineStyle = xlContinuous ' Añade bordes
        .Borders.Weight = xlThin ' Borde delgado
    End With
    
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
