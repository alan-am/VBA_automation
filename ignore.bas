Attribute VB_Name = "ignore"
' **************************************************************************
' !
' **************************************************************************
' Este código maneja la protección automática y fórmulas del índice.
'
' IMPORTANTE:
' **************************************************************************


Sub ExportarTodoElCodigoVBA()
    Dim componente As Object
    Dim rutaDestino As String
    Dim nombreArchivo As String

    ' Define la ruta donde se guardarán los archivos exportados
    rutaDestino = "C:\Users\alespana\Desktop\Alan\VBA_automation\"  ' <--- ruta

    ' Asegúrar termina con "\"
    If Right(rutaDestino, 1) <> "\" Then rutaDestino = rutaDestino & "\"

    ' Ojo de permitir acceso al modelo de objetos VBA
    ' (en el editor VBA: Herramientas > Referencias > Microsoft Visual Basic for Applications Extensibility)
    ' y también habilitar “Confiar en el acceso al modelo de objetos de proyecto VBA” en las opciones de seguridad.

    For Each componente In Application.VBE.ActiveVBProject.VBComponents
        
        nombreArchivo = rutaDestino & componente.Name
        
        Select Case componente.Type
            Case 1 ' Módulo estándar
                nombreArchivo = nombreArchivo & ".bas"
            Case 2 ' Módulo de clase
                nombreArchivo = nombreArchivo & ".cls"
            Case 3 ' Formulario de usuario
                nombreArchivo = nombreArchivo & ".frm"
            Case Else
                nombreArchivo = nombreArchivo & ".txt"
        End Select
        
        ' Exportar el componente
        componente.Export nombreArchivo
    Next componente

    MsgBox "Exportación completada en: " & rutaDestino, vbInformation
End Sub
