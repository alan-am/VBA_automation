Attribute VB_Name = "modExportarExcel"

' Funcion que escribe los datos de la carpeta en una hoja excel
Function ExportarDatosInventario(datos As Object) As Boolean
    
    ' Manejador de errores
    On Error GoTo ManejoError

    Dim ws As Worksheet
    Dim lRow As Long
    
    ' Hoja Destino
    Set ws = ThisWorkbook.Sheets("Test")
    
    ' Encuentra la primera fila vacía a partir de la columna A
    lRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    
    ' Mapeamos los datos a las columnas según la plantilla final y con respecto al diccionario Carpeta
    
    ' Col 1: SERIE/SUBSERIE DOCUMENTAL
    ws.Cells(lRow, 1).Value = datos("SerieSubserie")
    
    ' Col 2: N° CAJA
    ws.Cells(lRow, 2).Value = datos("NumCaja")
    
    ' Col 3: N° DE EXPEDIENTE
    ws.Cells(lRow, 3).Value = datos("NumExpediente")
    
    ' Col 4: NOMBRE DEL EXPEDIENTE (Nombre automático de la carpeta)
    ws.Cells(lRow, 4).Value = datos("Nombre")
    
    ' Col 5: FECHAS EXTREMAS - APERTURA (Fecha automática de creación)
    ws.Cells(lRow, 5).Value = datos("FechaCreacion")
    
    ' Col 6: FECHAS EXTREMAS - CIERRE (Fecha manual)
    ws.Cells(lRow, 6).Value = datos("FechaCierre")
    
    ' Col 7: FOJAS (Cantidad automática de archivos)
    ws.Cells(lRow, 7).Value = datos("CantidadArchivos")
    
    ' Col 8: DESTINO FINAL
    ws.Cells(lRow, 8).Value = datos("Destino")
    
    ' Col 9: SOPORTE
    ws.Cells(lRow, 9).Value = datos("Soporte")
    
    ' Col 10: UBICACIÓN TOPOGRÁFICA - ZONA
    ws.Cells(lRow, 10).Value = datos("UbicacionTopografica")
    
    ' Col 13: OBSERVACIONES (Saltamos Col 11 y 12)
    ws.Cells(lRow, 13).Value = datos("Observaciones")
    
    ' ------ datos de prueba adicionales que no van en la plantilla final ------
    ws.Cells(lRow, 14).Value = "Ruta: " & datos("Ruta")
    ws.Cells(lRow, 15).Value = "Tamaño: " & datos("TamanoTotal") & " KB"
    
    ' devoler true
    ExportarDatosInventario = True
    Exit Function

' Bloque de manejo de errores
ManejoError:
    ' Si algo falla (ej: la hoja "Test" no existe), devuelve False
    ExportarDatosInventario = False
End Function
