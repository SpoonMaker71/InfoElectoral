' ---------------------------------------------------------------
' Clase: clsDocProperty
' Autor: Juan Francisco Cucharero Cabezas
' Proyecto: InfoElectoral
' Descripción:
'   Clase auxiliar para leer, modificar o establecer propiedades
'   personalizadas en documentos Office (principalmente Excel).
'   Permite añadir metadatos como autor, título, proceso electoral,
'   fecha de exportación, etc.
'
'   ⚠️ Esta clase no contiene lógica de negocio ni interacción directa.
'   Su propósito es técnico: enriquecer los archivos generados con
'   información contextual útil para trazabilidad y auditoría.
'
' Fecha: [añadir fecha]
' ---------------------------------------------------------------

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDocProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_sNombre       As String
Private m_lTipo         As msoDocProperties
Private m_vValor        As Variant

Public Property Get ClassName() As String
    ClassName = "clsDocProperty"
End Property

Public Property Let NOMBRE(ByVal vData As String)
    m_sNombre = vData
End Property

Public Property Get NOMBRE() As String
    NOMBRE = m_sNombre
End Property

Public Property Let Tipo(ByVal vData As msoDocProperties)
    m_lTipo = vData
End Property

Public Property Get Tipo() As msoDocProperties
    Tipo = m_lTipo
End Property

Public Property Let Valor(ByVal vData As Variant)
    m_vValor = vData
End Property

Public Property Get Valor() As Variant
    Valor = m_vValor
End Property
