VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTemplates
' DateTime  : 02/07/2008 16:22
' Author    : Juan Francisco Cucharero Cabezas
' Purpose   :
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const dbKeyField As String = "Id"

Public Enum enumEstadoFichero
    NoExiste = 0
    SinDatos = 1
    Importado = 2
End Enum

Public Property Get ClassName() As String
    ClassName = "clsTemplates"
End Property

Public Function CreateTableDefs(ByVal lIdtemplate As Long) As Boolean
    Dim m_sSql              As String
    Dim m_TableDef          As TableDef
    Dim m_RstRows           As Recordset
    Dim m_RstFields         As Recordset
    Dim m_Index             As Index
    Dim m_Field             As Field
    Dim m_Relation          As Relation
    Dim m_sParentTableDef   As String

    On Error GoTo Error_CreateTableDefs

    ' 1. Obtenemos los distintos registros de la plantilla
    m_sSql = "SELECT"
    Concat m_sSql, " XML_Rows.Id,"
    Concat m_sSql, " XML_Rows.Type,"
    Concat m_sSql, " XML_Rows.TableDef,"
    Concat m_sSql, " XML_Rows.Description,"
    Concat m_sSql, " XML_Rows.Key"
    Concat m_sSql, " FROM XML_Rows"
    Concat m_sSql, " WHERE (XML_Rows.IdTemplate = " & CSqlDbl(lIdtemplate) & ")"
    Concat m_sSql, " ORDER BY XML_Rows.Id;"
    Set m_RstRows = CurrentDb.OpenRecordset(m_sSql)
    If modMDB.IsRst(m_RstRows) Then
        While Not m_RstRows.EOF
            ' 2. Si no existe la tabla, la creamos
            If Not modMDB.ObjectExists(m_RstRows!TableDef, acTable) Then
                m_sSql = "SELECT"
                Concat m_sSql, " XML_Fields.Id,"
                Concat m_sSql, " XML_Fields.Name,"
                Concat m_sSql, " XML_Fields.Type,"
                Concat m_sSql, " XML_Fields.Length,"
                Concat m_sSql, " XML_Fields.Format,"
                Concat m_sSql, " XML_Fields.Description,"
                Concat m_sSql, " XML_Fields.Caption,"
                Concat m_sSql, " XML_Fields.DBFormat"
                Concat m_sSql, " FROM XML_Fields"
                Concat m_sSql, " WHERE (XML_Fields.IdRow = " & m_RstRows!Id & ")"
                Concat m_sSql, " ORDER BY XML_Fields.Pos;"
                Set m_RstFields = CurrentDb.OpenRecordset(m_sSql)
                If modMDB.IsRst(m_RstFields) Then
                    If m_RstRows!Key Then m_sParentTableDef = m_RstRows!TableDef

                    ' 3. Definimos la nueva tabla
                    Set m_TableDef = CurrentDb.CreateTableDef(m_RstRows!TableDef)
                    With m_TableDef
                        Set m_Field = .CreateField(dbKeyField, 4)
                        m_Field.Required = True
                        .Fields.Append m_Field

                        ' 4. Definimos el campo índice de la tabla
                        Set m_Index = .CreateIndex(dbKeyField & "Index")
                        With m_Index
                            .Fields.Append m_Index.CreateField(dbKeyField)
                            .Unique = m_RstRows!Key
                            .Primary = m_RstRows!Key
                        End With
                        .Indexes.Append m_Index

                        ' 5. Agregamos los campos
                        While Not m_RstFields.EOF
                            .Fields.Append .CreateField(m_RstFields!Name, IIf((m_RstFields!Type = dbTimeStamp), dbDate, m_RstFields!Type), IIf((m_RstFields!Type = dbText), m_RstFields!Length, vbEmpty))
                            m_RstFields.MoveNext
                        Wend
                    End With
                End If

                With CurrentDb
                    ' 6. Agregamoa la nueva tabla a la base de datos
                    .TableDefs.Append m_TableDef

                    ' 7. Una vez que la tabla ha sido agregada a la base de datos,
                    '    es cuando podemos asignarle la propiedad "Description"
                    .TableDefs(m_RstRows!TableDef).Properties.Append .TableDefs(m_RstRows!TableDef).CreateProperty("Description", dbText, m_RstRows!Description)

                    ' 8. Agregamos las propiedades "Description" y "Caption" a cada
                    '    campo de la tabla
                    .TableDefs(m_RstRows!TableDef).Fields(dbKeyField).Properties.Append .TableDefs(m_RstRows!TableDef).CreateProperty("Description", dbText, "Identificador del elemento importado.")

                    m_RstFields.MoveFirst
                    While Not m_RstFields.EOF
                        ' 9. Si se definió, agregamos al campo la propiedad "Description"
                        If Not IsNull(m_RstFields!Description) Then .TableDefs(m_RstRows!TableDef).Fields(m_RstFields!Name).Properties.Append .TableDefs(m_RstRows!TableDef).CreateProperty("Description", dbText, m_RstFields!Description)

                        ' 10. Si se definió, agregamos al campo la propiedad "Caption"
                        If Not IsNull(m_RstFields!Caption) Then .TableDefs(m_RstRows!TableDef).Fields(m_RstFields!Name).Properties.Append .TableDefs(m_RstRows!TableDef).CreateProperty("Caption", dbText, m_RstFields!Caption)

                        ' 11. Si se definió, agregamos al campo la propiedad "Format"
                        If Not IsNull(m_RstFields!DBFormat) Then .TableDefs(m_RstRows!TableDef).Fields(m_RstFields!Name).Properties.Append .TableDefs(m_RstRows!TableDef).CreateProperty("Format", dbText, m_RstFields!DBFormat)

                        m_RstFields.MoveNext
                    Wend
                    m_RstFields.Close
                End With

                ' 12. Si no es la tabla principal, se relaciona con la principal,
                '     permitiendo la actualización en cascada
                If Not m_RstRows!Key Then modMDB.SetRelation m_sParentTableDef, m_RstRows!TableDef, (dbKeyField & ":" & dbKeyField), (dbRelationUpdateCascade + dbRelationDeleteCascade)
            End If
            m_RstRows.MoveNext
        Wend
        m_RstRows.Close
    End If

    Set m_Index = Nothing
    Set m_Field = Nothing
    Set m_Relation = Nothing
    Set m_RstFields = Nothing
    Set m_RstRows = Nothing
    Set m_TableDef = Nothing

    CreateTableDefs = True

    Exit Function

Error_CreateTableDefs:
    Set m_Index = Nothing
    Set m_Field = Nothing
    Set m_Relation = Nothing
    Set m_RstFields = Nothing
    Set m_RstRows = Nothing
    Set m_TableDef = Nothing
    GetError Me.ClassName & ".CreateTableDefs", m_sSql
End Function

Public Function CreateTemplateFile(ByVal lIdtemplate As Long, _
                                   Optional ByVal sFolderPath As String = vbNullString) As Boolean
    Dim m_sSql      As String
    Dim m_Xml       As Object
    Dim m_RstRows   As Recordset
    Dim m_RstFields As Recordset
    Dim m_RstValues As Recordset
    Dim m_Field     As Field
    Dim m_Value     As Field
    Dim m_sFilePath As String

    On Error GoTo Error_CreateTemplateFile

    ' 1. Si no se especificó una carpeta de destino
    If IsZrStr(sFolderPath) Then sFolderPath = modMDB.GetMDBPath("Plantillas")

    ' 2. Construimos la ruta de la carpeta si no existe
    If Not modSystem.FolderExists(sFolderPath) Then modSystem.BuiltPath sFolderPath

    ' 3. Obtenemos el nombre del fichero
    m_sFilePath = (sFolderPath & GetFileName(lIdtemplate))

    ' 4. Obtenemos los Nodos del XML
    m_sSql = "SELECT"
    Concat m_sSql, " XML_Rows.Id,"
    Concat m_sSql, " XML_Rows.Type,"
    Concat m_sSql, " XML_Rows.TableDef,"
    Concat m_sSql, " XML_Rows.Description,"
    Concat m_sSql, " IIf(XML_Rows.Key, 1, 0) AS Key"
    Concat m_sSql, " FROM XML_Rows"
    Concat m_sSql, " WHERE (XML_Rows.IdTemplate = " & CSqlDbl(lIdtemplate) & ")"
    Concat m_sSql, " ORDER BY XML_Rows.Id;"
    Set m_RstRows = CurrentDb.OpenRecordset(m_sSql)
    If modMDB.IsRst(m_RstRows) Then
        Set m_Xml = CreateObject("MSXML2.DOMDocument")
        ' 5. Establecemos la Cabecera del XML
        m_Xml.insertBefore CreateObject("MSXML2.DOMDocument").createProcessingInstruction("xml", "version=""1.0"" encoding=""ISO-8859-1"""), m_Xml.FirstChild

        ' 6. Agregamos la raiz del documento XML
        m_Xml.appendChild CreateObject("MSXML2.DOMDocument").createElement("Root")

        While Not m_RstRows.EOF
            ' 7. Agregamos el Nodo padre
            m_Xml.documentElement.appendChild CreateObject("MSXML2.DOMDocument").createElement("Row")

            ' 8. Agregamos los atributos del campo
            For Each m_Field In m_RstRows.Fields
                If EvalField(m_Field) Then
                    With m_Xml.documentElement.lastChild.Attributes
                        .setNamedItem CreateObject("MSXML2.DOMDocument").createAttribute(m_Field.Name)
                        .getNamedItem(m_Field.Name).Text = m_Field.Value
                    End With
                End If
            Next m_Field

            ' 9. Obtenemos la estructura del nodo
            m_sSql = "SELECT"
            Concat m_sSql, " XML_Fields.Id,"
            Concat m_sSql, " XML_Fields.Name,"
            Concat m_sSql, " XML_Fields.Type,"
            Concat m_sSql, " XML_Fields.Pos,"
            Concat m_sSql, " XML_Fields.Length,"
            Concat m_sSql, " XML_Fields.Format,"
            Concat m_sSql, " XML_Fields.Description,"
            Concat m_sSql, " XML_Fields.Caption,"
            Concat m_sSql, " XML_Fields.DBFormat"
            Concat m_sSql, " FROM XML_Fields"
            Concat m_sSql, " WHERE (XML_Fields.IdRow = " & CSqlDbl(m_RstRows!Id) & ")"
            Concat m_sSql, " ORDER BY XML_Fields.Pos;"
            Set m_RstFields = CurrentDb.OpenRecordset(m_sSql)
            If modMDB.IsRst(m_RstFields) Then
                While Not m_RstFields.EOF
                    ' 10. Agregamos el nodo de definición del campo
                    m_Xml.documentElement.lastChild.appendChild CreateObject("MSXML2.DOMDocument").createElement("Field")

                    ' 11. Agregamos los atributos del campo
                    For Each m_Field In m_RstFields.Fields
                        If EvalField(m_Field) Then
                            With m_Xml.documentElement.lastChild.lastChild.Attributes
                                .setNamedItem CreateObject("MSXML2.DOMDocument").createAttribute(m_Field.Name)
                                .getNamedItem(m_Field.Name).Text = m_Field.Value
                            End With
                        End If
                    Next m_Field

                    ' 12. Agregamos como nodos la lista de elementos/valores asociados a dicho campo
                    m_sSql = "SELECT"
                    Concat m_sSql, " XML_Values.Value,"
                    Concat m_sSql, " XML_Values.Description"
                    Concat m_sSql, " FROM XML_Values"
                    Concat m_sSql, " WHERE (XML_Values.IdField = " & CSqlDbl(m_RstFields!Id) & ")"
                    Concat m_sSql, " ORDER BY XML_Values.Id;"
                    Set m_RstValues = CurrentDb.OpenRecordset(m_sSql)
                    If modMDB.IsRst(m_RstValues) Then
                        While Not m_RstValues.EOF
                            ' 13. Agregamos el elemento/valor
                            m_Xml.documentElement.lastChild.lastChild.appendChild CreateObject("MSXML2.DOMDocument").createElement("Item")

                            For Each m_Value In m_RstValues.Fields
                                ' 14. Agregamos los atributos
                                With m_Xml.documentElement.lastChild.lastChild.lastChild.Attributes
                                    .setNamedItem CreateObject("MSXML2.DOMDocument").createAttribute(m_Value.Name)
                                    .getNamedItem(m_Value.Name).Text = Nz(m_Value.Value, " ")
                                End With
                            Next m_Value
                            m_RstValues.MoveNext
                        Wend
                        m_RstValues.Close
                    End If
                    m_RstFields.MoveNext
                Wend
                m_RstFields.Close
            End If
            m_RstRows.MoveNext
        Wend
        m_RstRows.Close
        m_Xml.Save m_sFilePath
    End If

    Set m_Xml = Nothing
    Set m_Value = Nothing
    Set m_Field = Nothing
    Set m_RstRows = Nothing
    Set m_RstFields = Nothing
    Set m_RstValues = Nothing

    CreateTemplateFile = modSystem.FileExists(m_sFilePath)

    Exit Function

Error_CreateTemplateFile:
    Set m_Xml = Nothing
    Set m_Value = Nothing
    Set m_Field = Nothing
    Set m_RstRows = Nothing
    Set m_RstFields = Nothing
    Set m_RstValues = Nothing
    GetError Me.ClassName & ".CreateTemplateFile", m_sSql
End Function

Private Function EvalField(ByVal oField As Object) As Boolean
    EvalField = Not (IsNull(oField.Value) Or (UCase(oField.Name) = UCase(dbKeyField)))
End Function

Private Function GetDate(ByVal sValue As String, _
                         ByVal sFormat As String) As Date
    GetDate = CDate(GetSrhStr(sValue, sFormat, "y") & "-" & GetSrhStr(sValue, sFormat, "m") & "-" & GetSrhStr(sValue, sFormat, "d"))
End Function

Private Function GetTimeStamp(ByVal sValue As String, _
                              ByVal sFormat As String) As Date
    GetTimeStamp = CDate(GetSrhStr(sValue, sFormat, "y") & "-" & GetSrhStr(sValue, sFormat, "m") & "-" & GetSrhStr(sValue, sFormat, "d") & " " & GetSrhStr(sValue, sFormat, "h") & ":" & GetSrhStr(sValue, sFormat, "n") & ":" & GetSrhStr(sValue, sFormat, "s"))
End Function

Public Function GetFileName(ByVal lIdtemplate As Long) As String
    Dim m_sSql      As String
    Dim m_Rst       As Recordset

    On Error GoTo Error_GetFileName

    m_sSql = "SELECT"
    Concat m_sSql, " XML_Templates.Description,"
    Concat m_sSql, " XML_Templates.FileName"
    Concat m_sSql, " FROM XML_Templates"
    Concat m_sSql, " WHERE (XML_Templates.Id = " & CSqlDbl(lIdtemplate) & ");"
    Set m_Rst = CurrentDb.OpenRecordset(m_sSql)
    If modMDB.IsRst(m_Rst) Then
        GetFileName = m_Rst!FileName
        m_Rst.Close
    End If
    Set m_Rst = Nothing

    Exit Function

Error_GetFileName:
    Set m_Rst = Nothing
    GetError Me.ClassName & ".GetFileName", m_sSql
End Function

Public Function GetParentTableDef(ByVal lIdtemplate As Long) As String
    Dim m_sSql      As String
    Dim m_Rst       As Recordset

    On Error GoTo Error_GetParentTableDef

    m_sSql = "SELECT"
    Concat m_sSql, " XML_Rows.TableDef,"
    Concat m_sSql, " XML_Rows.Key"
    Concat m_sSql, " FROM XML_Rows"
    Concat m_sSql, " WHERE (XML_Rows.IdTemplate = " & CSqlDbl(lIdtemplate) & ")"
    Concat m_sSql, " AND (XML_Rows.Key = True);"
    Set m_Rst = CurrentDb.OpenRecordset(m_sSql)
    If modMDB.IsRst(m_Rst) Then
        GetParentTableDef = m_Rst!TableDef
        m_Rst.Close
    End If
    Set m_Rst = Nothing

    Exit Function

Error_GetParentTableDef:
    Set m_Rst = Nothing
    GetError Me.ClassName & ".GetParentTableDef", m_sSql
End Function

Private Function GetRowKeyTableDef(ByVal oXML As Object) As String
    Dim m_Row      As Object

    On Error GoTo Error_GetRowKeyTableDef

    For Each m_Row In oXML.selectSingleNode("//Root").childNodes
        If IsRowKey(m_Row) Then
            GetRowKeyTableDef = GetXMLAttribute(m_Row, "TableDef")
            Exit For
        End If
    Next m_Row
    Set m_Row = Nothing

    Exit Function

Error_GetRowKeyTableDef:
    Set m_Row = Nothing
    GetError Me.ClassName & ".GetRowKeyTableDef"
End Function

Private Function GetRowNode(ByVal oXML As Object, _
                            ByVal sType As String) As Object
    Dim m_Node      As Object
    Dim m_sType     As String

    On Error GoTo Error_GetRowNode

    For Each m_Node In oXML.selectSingleNode("//Root").childNodes
        m_sType = GetXMLAttribute(m_Node, "Type")
        If Not IsZrStr(m_sType) Then
            If GetXMLAttribute(m_Node, "Type") = sType Then
                Set GetRowNode = m_Node
                Exit For
            End If
        Else
            Set GetRowNode = m_Node
            Exit For
        End If
    Next m_Node
    Set m_Node = Nothing

    Exit Function

Error_GetRowNode:
    Set m_Node = Nothing
    GetError Me.ClassName & ".GetRowNode"
End Function

Public Function GetSQL(ByVal sTableDef As String, _
                       Optional ByVal sCriteria As String = vbNullString, _
                       Optional ByVal sSort As String = vbNullString) As String
    Dim m_sSql      As String
    Dim m_Rst       As Recordset

    On Error GoTo Error_GetSQL

    m_sSql = "SELECT"
    Concat m_sSql, " XML_Fields.Name,"
    Concat m_sSql, " XML_Fields.Caption"
    Concat m_sSql, " FROM XML_Rows INNER JOIN XML_Fields ON XML_Rows.Id = XML_Fields.IdRow"
    Concat m_sSql, " WHERE ((XML_Fields.Visible = True)"
    Concat m_sSql, " AND (XML_Rows.TableDef = " & CSqlTxt(sTableDef) & "))"
    Concat m_sSql, " ORDER BY XML_Fields.IdRow,"
    Concat m_sSql, " XML_Fields.Id;"
    Set m_Rst = CurrentDb.OpenRecordset(m_sSql)
    If modMDB.IsRst(m_Rst) Then
        GetSQL = "SELECT"
        While (Not m_Rst.EOF)
            Concat GetSQL, " [" & sTableDef & "].[" & m_Rst!Name & "]"
            If (Not IsNull(m_Rst!Caption)) Then Concat GetSQL, " AS [" & m_Rst!Caption & "]"
            m_Rst.MoveNext
            If Not m_Rst.EOF Then Concat GetSQL, ","
        Wend
        Concat GetSQL, " FROM [" & sTableDef & "]"
        If (Not IsZrStr(sCriteria)) Then Concat GetSQL, " WHERE (" & sCriteria & ")"
        If (Not IsZrStr(sSort)) Then Concat GetSQL, " ORDER BY " & sSort
        Concat GetSQL, ";"
        m_Rst.Close
    End If
    Set m_Rst = Nothing

    Exit Function

Error_GetSQL:
    Set m_Rst = Nothing
    GetError Me.ClassName & ".GetSQL", m_sSql
End Function

Private Function GetValue(ByVal oNode As Object, _
                          ByVal sLine As String) As String
    Dim m_sLength()     As String
    Dim m_sName         As String
    Dim m_sValue        As String
    Dim m_iType         As Integer
    Dim m_lPos          As Long
    Dim m_sSize         As String
    Dim m_sFormat       As String
    Dim m_sDescription  As String
    Dim m_sCaption      As String
    Dim m_lX            As Long

    On Error GoTo Error_GetValue

    ' 1. Obtenemos las características del campo cuyo valor se va a obtener
    m_sName = GetXMLAttribute(oNode, "Name")
    m_iType = CInt(GetXMLAttribute(oNode, "Type"))
    m_sSize = GetXMLAttribute(oNode, "Length")
    m_lPos = CLng(GetXMLAttribute(oNode, "Pos"))
    m_sFormat = GetXMLAttribute(oNode, "Format")
    m_sDescription = GetXMLAttribute(oNode, "Description")
    m_sCaption = GetXMLAttribute(oNode, "Caption")

    ' 2. Si se ha definido un tamaño de campo
    If Not IsZrStr(m_sSize) Then
        m_sLength = Split(m_sSize, ",")
        Select Case UBound(m_sLength)
            Case 0
                m_sValue = Trim(Mid(sLine, m_lPos, CLng(m_sLength(0))))
            Case 1
                m_sValue = Trim(Mid(sLine, m_lPos, (CLng(m_sLength(0)) + CLng(m_sLength(1)))))
        End Select
    Else
        If (m_lPos <= Len(sLine)) Then m_sValue = Trim(Mid(sLine, m_lPos, (Len(sLine) + 1 - m_lPos)))
    End If

    ' 3. Si se ha definido un valor
    If Not IsZrStr(m_sValue) Then
        Select Case m_iType
            ' Byte, Integer, Long, Currency, Single, Double
            Case dbByte, _
                 dbInteger, _
                 dbLong, _
                 dbCurrency, _
                 dbSingle, _
                 dbDouble
                ' 4. Si el tipo del campo es numérico
                If modSystem.EvalStr(m_sValue, "[0-9\.0-9]") Then
                    m_sValue = Replace(m_sValue, ".", vbNullString)
                    m_lX = IIf(IsZr(InStr(1, m_sValue, "-")), 1, InStr(1, m_sValue, "-"))

                    Select Case UBound(m_sLength)
                        Case 0  ' Si solo tiene parte Entera
                            GetValue = CSql(CDbl(Mid(m_sValue, m_lX, (CInt(m_sLength(0)) - (m_lX - 1)))))

                        Case 1  ' Si tiene parte Decimal
                            GetValue = CSql(CDbl(Mid(m_sValue, m_lX, (CInt(m_sLength(0)) + CInt(m_sLength(1)) - (m_lX - 1)))) / (10 ^ CInt(m_sLength(1))))
                    End Select
                Else
                    GetValue = "Null"
                End If

            ' Fecha
            Case dbDate
                GetValue = CSql(GetDate(m_sValue, m_sFormat))

            ' Fecha y hora
            Case dbTimeStamp
                GetValue = CSql(GetTimeStamp(m_sValue, m_sFormat))

            ' Texto
            Case dbText, _
                 dbMemo
                GetValue = CSql(m_sValue)
        End Select
    Else
        GetValue = "Null"
    End If

    Exit Function

Error_GetValue:
    If (Err.Number = 91) Then
        Err.Clear
        Resume 0
    Else
        Dim m_szDebug As String

        ' 5. Obtenemos los datos de depuración
        m_szDebug = "Name=" & m_sName
        If Not IsZrStr(m_sDescription) Then Concat m_szDebug, vbCrLf & "Description=" & m_sDescription
        If Not IsZrStr(m_sCaption) Then Concat m_szDebug, vbCrLf & "Caption=" & m_sCaption
        Concat m_szDebug, vbCrLf & "Type=" & CStr(m_iType)
        If Not IsZrStr(m_sSize) Then Concat m_szDebug, vbCrLf & "Length=" & m_sSize
        Concat m_szDebug, vbCrLf & "Pos=" & CStr(m_lPos)
        If Not IsZrStr(m_sFormat) Then Concat m_szDebug, vbCrLf & "Format=" & m_sFormat
        If Not IsZrStr(m_sValue) Then Concat m_szDebug, vbCrLf & "Value=" & m_sValue

        GetError Me.ClassName & ".GetValue", m_szDebug
    End If
End Function

' Método para obtener el valor de un atributo de un Nodo de XML.
Public Function GetXMLAttribute(ByVal xmlNode As Object, _
                                ByVal sAttributeName As String) As String
    On Error GoTo Error_GetXMLAttribute

    GetXMLAttribute = xmlNode.Attributes.getNamedItem(sAttributeName).Text

    Exit Function

Error_GetXMLAttribute:
    Err.Clear
End Function

Public Function ImportFile(ByVal lIdtemplate As Long, _
                           ByVal sFilePath As String, _
                           Optional ByVal bUseProgressBar As Boolean = True, _
                           Optional ByVal bCloseBarProgres As Boolean = True, _
                           Optional ByVal lCharCode As enumCharcode = enumCharcode.CdoSystemASCII, _
                           Optional ByRef lNumRows As Long = 0, _
                           Optional ByVal sProgressBarCaption As String = vbNullString) As enumEstadoFichero
    Dim m_sTemplatePath     As String
    Dim m_Xml               As Object
    Dim m_Row               As Object
    Dim m_Field             As Object
    Dim m_sSql              As String
    Dim m_sParentTableDef   As String
    Dim m_sDoc              As String
    Dim m_sLines()          As String
    Dim m_iLine             As Integer

    On Error GoTo Error_ImportFile

    lNumRows = 0

    m_sTemplatePath = (modMDB.GetMDBPath("Plantillas") & GetFileName(lIdtemplate))

    ' 1. Comprobamos que exista la plantilla de importación
    If Not modSystem.FileExists(m_sTemplatePath) Then Err.Raise (vbObjectError + 101), GetMDBName, "No se encuentra la plantilla de importación [" & m_sTemplatePath & "]."

    ' 2. Obtenemos el contenido del fichero a importar
    m_sDoc = modSystem.ReadFile(sFilePath, IIf((lCharCode = enumCharcode.CdoSystemASCII), enumCharcode.CdoISO_8859_1, lCharCode))

    ' 3. Si el fichero no está vacío
    If Not IsZrStr(m_sDoc) Then
        ' 4. Establecemos las líneas de datos del fichero
        m_sLines = Split(m_sDoc, vbCrLf)

        ' 5. Si el salto de línea no corresponde con Windows (CR LF)
        If Not IsZrStr(m_sDoc) And IsZr(UBound(m_sLines)) Then Err.Raise (vbObjectError + 101), modMDB.GetMDBName(), "El fichero [" & modSystem.GetFileFullName(sFilePath) & "] no tiene está codificado para el salto de línea de Windows. Utilice Notepad++ para cambiar la codificación del salto de línea a Windows (CR LF)."

        ' 6. Abrimos el formulario de Progreso y le damos formato a la barra de estado
        If bUseProgressBar Then modMDB.SetProgressBar 0, UBound(m_sLines), IIf(IsZrStr(sProgressBarCaption), "Obteniendo datos del fichero seleccionado.", sProgressBarCaption)

        ' 7. Obtenemos la plantilla de importación
        Set m_Xml = CreateObject("MSXML2.DOMDocument")
        m_Xml.Load m_sTemplatePath
        m_sParentTableDef = GetRowKeyTableDef(m_Xml)

        ' 8. Recorremos el contenido del fichero
        For m_iLine = LBound(m_sLines) To UBound(m_sLines)
            ' 9. Importamos cada línea a la tabla temporal correspondiente
            Set m_Row = GetRowNode(m_Xml, Left(m_sLines(m_iLine), 2))
            If Not m_Row Is Nothing Then
                m_sSql = "INSERT INTO "
                Concat m_sSql, GetXMLAttribute(m_Row, "TableDef") & " (Id"
                For Each m_Field In m_Row.childNodes
                    Concat m_sSql, ", [" & GetXMLAttribute(m_Field, "Name") & "]"
                Next m_Field
                Concat m_sSql, IIf(IsRowKey(m_Row), ") SELECT (Nz(Max(" & m_sParentTableDef & ".Id), 0) + 1) AS Id", ") SELECT Max(" & m_sParentTableDef & ".Id) AS Id")
                For Each m_Field In m_Row.childNodes
                    Concat m_sSql, ", " & IIf((GetXMLAttribute(m_Field, "Name") = "FicheroTXT"), CSqlTxt(modSystem.GetFileFullName(sFilePath)), GetValue(m_Field, m_sLines(m_iLine))) & " AS [" & GetXMLAttribute(m_Field, "Name") & "]"
                Next m_Field
                Concat m_sSql, " FROM " & m_sParentTableDef & ";"
                CurrentDb.Execute m_sSql, dbFailOnError

                Incr lNumRows, 1
            End If
            If bUseProgressBar Then modMDB.IncrProgressBar
        Next m_iLine

        ' 10. Cerramos el formulario de progreso una vez acabada la tarea
        If bUseProgressBar And bCloseBarProgres Then modMDB.CloseProgressBar

        Set m_Field = Nothing
        Set m_Row = Nothing
        Set m_Xml = Nothing

        ImportFile = enumEstadoFichero.Importado
    Else
        ImportFile = enumEstadoFichero.SinDatos
    End If

    Exit Function

Error_ImportFile:
    Set m_Field = Nothing
    Set m_Row = Nothing
    Set m_Xml = Nothing
    If bCloseBarProgres Then modMDB.CloseProgressBar
    GetError Me.ClassName & ".ImportFile", "m_iLine=" & m_iLine & vbTab & "m_sSql=" & m_sSql
End Function

Private Function IsRowKey(ByVal oRow As Object) As Boolean
    On Error GoTo Error_IsRowKey

    IsRowKey = (CInt(GetXMLAttribute(oRow, "Key")) = 1)

    Exit Function

Error_IsRowKey:
    GetError Me.ClassName & ".IsRowKey"
End Function

Public Function TemplateExists(ByVal lIdtemplate As Long) As Boolean
    On Error GoTo Error_TemplateExists

    ' Se comprueba que existan las tablas temporales de importación en la base de datos,
    ' y que exista la plantilla XML. En caso contrario se crean.
    TemplateExists = (CreateTableDefs(lIdtemplate) And CreateTemplateFile(lIdtemplate))

    Exit Function

Error_TemplateExists:
    GetError Me.ClassName & ".TemplateExists"
End Function

