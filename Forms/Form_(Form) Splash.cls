' ---------------------------------------------------------------
' Clase: Form_(Form) Splash.cls
' Autor: Juan Francisco Cucharero Cabezas
' Proyecto: InfoElectoral
' Descripción:
'   Formulario de presentación que se muestra al iniciar la aplicación.
'   Puede incluir animaciones, mensajes de carga, logos o inicialización
'   de variables globales.
'
'   ⚠️ Este formulario no contiene lógica de negocio ni interacción directa
'   con el usuario. Su propósito es visual y técnico: mostrar el arranque
'   del sistema mientras se preparan los recursos.
'
' Fecha: [añadir fecha]
' ---------------------------------------------------------------

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_(Form) Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
    ' 1. Ocultamos la ventana de la aplicación Access
    ShowAccessWindow vbHide

    ' 2. Desactivamos elcompactar al salir de la aplicación
    Application.SetOption "Auto Compact", False

    ' 3. Iniciamos el temporizador
    Me.TimerInterval = 3000
End Sub

Private Sub Form_Timer()
    ' 1. Una vez terminado el tiempo del temporizador,
    '    cerramos el formulario del logo
    modMDB.CloseForm Me.Name

    ' 2. Abrimos el formulario principal de la aplicación
    modMDB.OpenForm "(Form) Consulta de resultados electorales por mesa"
End Sub
