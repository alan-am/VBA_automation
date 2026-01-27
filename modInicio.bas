Attribute VB_Name = "modInicio"
'modInicio
' **************************************************************************
' ! MACROS DE ACCESO (PUNTOS DE ENTRADA DESDE LOS BOTONES)
' **************************************************************************
Option Explicit


Sub AbrirAnalisisRecursivo()
    ' Abre el formulario de analisis secuencial
    frmDatosMasivos.Show
End Sub

Sub AbrirBuscadorSeccion()
    'Abre la Lupa / Buscador de secciones
    frmBusqueda.Show
End Sub


'OCULTO(DISPONIBLE PARA USO)
Sub AbrirFormularioDatosCarpeta()
    ' Abre el formulario de edición individual de carpeta individual.
    frmDatosCarpeta.Show
End Sub

'OCULTO(DISPONIBLE PARA USO)
Sub AbrirFormularioFisico()
    ' Abre el formulario de edición individual de carpeta fisica.
    frmDatosFisicos.Show
End Sub

