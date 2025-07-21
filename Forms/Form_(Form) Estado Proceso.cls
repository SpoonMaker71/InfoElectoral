' ---------------------------------------------------------------
' Formulario: Form_(Form) Estado Proceso
' Autor: Juan Francisco Cucharero Cabezas
' Proyecto: InfoElectoral
' Descripción:
'   Formulario auxiliar que muestra el estado actual del proceso
'   de importación, validación y análisis de datos electorales.
'   Permite al usuario visualizar el progreso, detectar errores
'   y confirmar que cada etapa se ha completado correctamente.
'
'   ⚠️ Este formulario no modifica datos directamente.
'   Su propósito es informativo y de control visual del flujo.
'
' Fecha: [añadir fecha]
' ---------------------------------------------------------------

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_(Form) Estado Proceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_(Form) Estado Proceso
' DateTime  : 02/07/2008 16:20
' Author    : Juan Francisco Cucharero Cabezas
' Purpose   :
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Sub Form_Load()
    On Error GoTo Error_Form_Load

    Screen.MousePointer = 11

    With pbProceso
        .Min = 0
        .Max = 100
    End With

    Screen.MousePointer = 0

    Exit Sub

Error_Form_Load:
    GetError Me.Name & ".Form_Load"
End Sub

