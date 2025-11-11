Attribute VB_Name = "modExportarExcel"
'revisar bug puntero
' Funcion que escribe los datos de la carpeta en una hoja excel
Function ExportarDatosInventario(datos As Object) As Boolean
    
    ' Manejador de errores
    On Error GoTo ManejoError

    Dim ws As Worksheet
    Dim lRow As Long
    
    ' Hoja Destino
    Set ws = ThisWorkbook.Sheets("Inventario General")
    
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    ' 2. Lee el número de fila desde la celda puntero (D2)
    ' Val() convierte el contenido de la celda a número. Si está vacío, devuelve 0.
    lRow = CLng(Val(wsConfig.Range("D2").Value))
    
    ' 3. Verificación de seguridad
    ' Si la celda D2 está vacía (0) o < 2, empieza en la fila 9
    ' (Asumiendo que la fila 1 siempre es la cabecera)
    If lRow < 2 Then
        lRow = 9
    End If
                                    
  
    ' Mapeamos los datos a las columnas según la plantilla final y con respecto al diccionario Carpeta
    
    'empieza desde la columna 2
    ' Col 1: SERIE/SUBSERIE DOCUMENTAL
    ws.Cells(lRow, 2).Value = datos("SerieSubserie")
    
    ' Col 2: N° CAJA
    ws.Cells(lRow, 3).Value = datos("NumCaja")
    
    ' Col 3: N° DE EXPEDIENTE
    ws.Cells(lRow, 4).Value = datos("NumExpediente")
    
    ' Col 4: NOMBRE DEL EXPEDIENTE (Nombre automático de la carpeta)
    ws.Cells(lRow, 5).Value = datos("Nombre")
    
    ' Col 5: FECHAS EXTREMAS - APERTURA (Fecha automática de creación)
    ws.Cells(lRow, 6).Value = datos("FechaCreacion")
    
    ' Col 6: FECHAS EXTREMAS - CIERRE (Fecha manual)
    ws.Cells(lRow, 7).Value = datos("FechaCierre")
    
    ' Col 7: FOJAS (Cantidad automática de archivos)
    ws.Cells(lRow, 8).Value = datos("CantidadArchivos")
    
    ' Col 8: DESTINO FINAL
    ws.Cells(lRow, 9).Value = datos("Destino")
    
    ' Col 9: SOPORTE
    ws.Cells(lRow, 10).Value = datos("Soporte")
    
    ' Col 10: UBICACIÓN TOPOGRÁFICA - ZONA
    ws.Cells(lRow, 11).Value = datos("UbicacionTopografica") ' Asigna "NN"
    
    ' Col 11: UBICACIÓN TOPOGRÁFICA - ESTANTE
    ws.Cells(lRow, 12).Value = datos("UbicacionTopografica") ' Asigna "NN"
    
    ' Col 12: UBICACIÓN TOPOGRÁFICA - BANDEJA
    ws.Cells(lRow, 13).Value = datos("UbicacionTopografica") ' Asigna "NN"
    
    ' Col 13: OBSERVACIONES
    ws.Cells(lRow, 14).Value = datos("Observaciones")
    
    ' ------ datos de prueba adicionales que no van en la plantilla final ------
    'ws.Cells(lRow, 14).Value = "Ruta: " & datos("Ruta")
    'ws.Cells(lRow, 15).Value = "Tamaño: " & datos("TamanoTotal") & " KB"
    
    ' --- 5. ACTUALIZAR EL PUNTERO ---
    ' Si todo salió bien, actualiza la celda Z1 para la *siguiente* fila
    wsConfig.Range("D2").Value = lRow + 1
    
    ' devolver true
    ExportarDatosInventario = True
    Exit Function

' Bloque de manejo de errores
ManejoError:
    ' Si algo falla (ej: la hoja "Inventario General" no existe), devuelve False
    MsgBox "Error al exportar: " & Err.Description & vbCrLf & _
           "Asegúrese que la hoja 'Inventario General' existe.", vbCritical, "Error en 'modExportarExcel'"
    ExportarDatosInventario = False
End Function
