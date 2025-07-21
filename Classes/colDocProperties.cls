VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colDocProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mCol            As Collection

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Property Get FirstItem() As clsDocProperty
    If HasItems() Then Set FirstItem = mCol(1)
End Property

Public Property Get Item(vntIndexKey As Variant) As clsDocProperty
    Set Item = mCol(vntIndexKey)
End Property

Public Property Get HasItems() As Boolean
    HasItems = Not IsZr(mCol.Count)
End Property

Public Property Get LastItem() As clsDocProperty
    If HasItems() Then Set LastItem = mCol(mCol.Count)
End Property

Public Property Get ClassName() As String
    ClassName = "colDocProperties"
End Property

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Public Function Add(ByVal sNombre As String, _
                    Optional ByVal lTipo As msoDocProperties = msoDocProperties.msoPropertyTypeString, _
                    Optional ByVal vValor As Variant = vbNullString) As clsDocProperty
    Dim m_clsDocProperty     As New clsDocProperty

    On Error GoTo Error_Add

    With m_clsDocProperty
        .NOMBRE = sNombre
        .Tipo = lTipo
        .Valor = vValor
    End With
    mCol.Add m_clsDocProperty, sNombre

    Set Add = m_clsDocProperty
    Set m_clsDocProperty = Nothing

    Exit Function

Error_Add:
    Set m_clsDocProperty = Nothing
    GetError Me.ClassName & ".Add"
End Function

' Método para limpiar la colección de todos sus elementos
Public Function Clear() As Boolean
    Dim m_lCountItem      As Long
    Dim m_lIdx            As Long

    On Error GoTo Error_Clear

    m_lCountItem = mCol.Count
    For m_lIdx = 0 To (m_lCountItem - 1)
        mCol.Remove (m_lCountItem - m_lIdx)
    Next m_lIdx
    Clear = Not HasItems()

    Exit Function

Error_Clear:
    GetError Me.ClassName & ".Clear"
End Function

' Función para validar si existe un elemento en la colección
Public Function Exists(ByVal vData As String) As Boolean
    Dim m_lIdx            As Long

    On Error GoTo Error_Exists

    For m_lIdx = 1 To mCol.Count
        If (mCol(m_lIdx).NOMBRE = vData) Then
            Exists = True
            Exit For
        End If
    Next m_lIdx

    Exit Function

Error_Exists:
    GetError Me.ClassName & ".Exists"
End Function

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Private Sub Class_Terminate()
    If HasItems() Then Clear
    Set mCol = Nothing
End Sub

