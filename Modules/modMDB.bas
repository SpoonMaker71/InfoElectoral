' ---------------------------------------------------------------
' Módulo: modMDB.bas
' Autor: Juan Francisco Cucharero Cabezas
' Proyecto: InfoElectoral
' Descripción:
'   Este módulo contiene funciones técnicas para la gestión directa
'   de tablas en la base de datos Access (.accdb), incluyendo
'   operaciones de importación, limpieza, verificación y duplicación.
'
'   ⚠️ Importante:
'   Este módulo no realiza validaciones de datos ni confirmaciones
'   al usuario. Toda la lógica interactiva (mensajes, decisiones,
'   validaciones de estructura o confirmaciones de borrado) se gestiona
'   exclusivamente desde los formularios del proyecto.
'
'   Esta separación permite mantener una arquitectura limpia y modular,
'   facilitando la reutilización de funciones técnicas en otros contextos.
'
' Fecha: [puedes añadir la fecha de creación o última modificación]
' ---------------------------------------------------------------

Attribute VB_Name = "modMDB"
Option Compare Database
Option Explicit

Public Enum enumBackgroundStatusColor
    DisabledControl = 15269887
    EnabledControl = -2147483643
End Enum

Public Enum enumTypeLetterCase
    BothCase = 0
    LowCase = 1
    UpperCase = 2
End Enum

' Constantes del objeto Excel.Application ----------------------------------------------
Global Const xlAutomatic                            As Long = -4105      ' (&HFFFFEFF7)
Global Const xlContext                              As Long = -5002      ' (&HFFFFEC76)
Global Const xlNone                                 As Long = -4142      ' (&HFFFFEFD2)
Global Const xlTopToBottom                          As Long = 1

Public Enum xlUnderlineStyle
    xlDouble = -4119                        ' (&HFFFFEFE9)
    xlDoubleAccounting = 5
    xlSingle = 2
    xlSingleAccounting = 4
    xlSolid = 1
    xlUnderlineStyleNone = -4142            ' (&HFFFFEFD2)
End Enum

Public Enum xlLineStyle
    xlContinuous = 1
    xlDash = -4115                          ' (&HFFFFEFED)
    xlDashDot = 4
    xlDashDotDot = 5
    xlDot = -4118                           ' (&HFFFFEFEA)
    xlDouble = -4119                        ' (&HFFFFEFE9)
    xlLineStyleNone = -4142                 ' (&HFFFFEFD2)
    xlSlantDashDot = 13
End Enum

Public Enum xlAlignment
    xlDiagonalDown = 5
    xlDiagonalUp = 6
    xlEdgeBottom = 9
    xlEdgeLeft = 7
    xlEdgeRight = 10
    xlEdgeTop = 8
    xlInsideHorizontal = 12
    xlInsideVertical = 11
    xlContinuous = 1
    xlCenter = -4108                        ' (&HFFFFEFF4)
    xlTop = -4160                           ' (&HFFFFEFC0)
    xlGeneral = 1
    xlBottom = -4107                        ' (&HFFFFEFF5)
    xlLeft = -4131                          ' (&HFFFFEFDD)
    xlRight = -4152                         ' (&HFFFFEFC8)
End Enum

Public Enum xlAxisType
    xlCategory = 1
    xlValue = 2
    xlSeriesAxis = 3
End Enum

Public Enum xlChartLocation
    xlLocationAsNewSheet = 1
    xlLocationAsObject = 2
    xlLocationAutomatic = 3
End Enum

Public Enum xlAxisGroup
    xlPrimary = 1
    xlSecondary = 2
End Enum

Public Enum xlTrendlineType
    xlExponential = 5
    xlLinear = -4132                        ' (&HFFFFEFDC)
    xlLogarithmic = -4133                   ' (&HFFFFEFDB)
    xlMovingAvg = 6
    xlPolynomial = 3
    xlPower = 4
End Enum

' Constantes de Tipos de formato de archivo
Public Enum xlFileFormat
    xlAddIn = 18
    xlCSV = 6
    xlCSVMac = 22
    xlCSVMSDOS = 24
    xlCSVWindows = 23
    xlCurrentPlatformText = -4158
    xlDBF2 = 7
    xlDBF3 = 8
    xlDBF4 = 11
    xlDIF = 9
    xlExcel2 = 16
    xlExcel2FarEast = 27
    xlExcel3 = 29
    xlExcel4 = 33
    xlExcel4Workbook = 35
    xlExcel5 = 39
    xlExcel7 = 39
    xlExcel9795 = 43
    xlHtml = 44
    xlIntlMacro = 25
    xlSYLK = 2
    xlTemplate = 17
    xlTextMac = 19
    xlTextMSDOS = 21
    xlTextPrinter = 36
    xlTextWindows = 20
    xlUnicodeText = 42
    xlWebArchive = 45
    xlWJ2WD1 = 14
    xlWJ3 = 40
    xlWJ3FJ3 = 41
    xlWK1 = 5
    xlWK1ALL = 31
    xlWK1FMT = 30
    xlWK3 = 15
    xlWK3FM3 = 32
    xlWK4 = 38
    xlWKS = 4
    xlWorkbookNormal = -4143
    xlWorks2FarEast = 28
    xlWQ1 = 34
    xlXMLSpreadsheet = 46
End Enum

Public Enum xlPrintLocation
    xlPrintInPlace = 16                     ' (&H10)
    xlPrintNoComments = -4142               ' (&HFFFFEFD2)
    xlPrintSheetEnd = 1
End Enum

Public Enum xlPrintErrors
    xlPrintErrorsDisplayed = 0
    xlPrintErrorsBlank = 1
    xlPrintErrorsDash = 2
    xlPrintErrorsNA = 3
End Enum

Public Enum xlPageOrientation
    xlPortrait = 1
    xlLandscape = 2
End Enum

Public Enum xlOrder
    xlDownThenOver = 1
    xlOverThenDown = 2
End Enum

Public Enum xlPaperSize
    xlPaperLetter = 1
    xlPaperLetterSmall = 2
    xlPaperTabloid = 3
    xlPaperLedger = 4
    xlPaperLegal = 5
    xlPaperStatement = 6
    xlPaperExecutive = 7
    xlPaperA3 = 8
    xlPaperA4 = 9
    xlPaperA4Small = 10
    xlPaperA5 = 11
    xlPaperB4 = 12
    xlPaperB5 = 13
    xlPaperFolio = 14
    xlPaperQuarto = 15
    xlPaper10x14 = 16                       ' (&H10)
    xlPaper11x17 = 17                       ' (&H11)
    xlPaperNote = 18                        ' (&H12)
    xlPaperCsheet = 24                      ' (&H18)
    xlPaperDsheet = 25                      ' (&H19)
    xlPaperUser = 256                       ' (&H100)
End Enum

Public Enum xlBorderWeight
    xlHairline = 1
    xlThin = 2
    xlMedium = -4138                        ' (&HFFFFEFD6)
    xlThick = 4
End Enum

Public Enum xlSortOrder
    xlAscending = 1
    xlDescending = 2
End Enum

Public Enum xlYesNoGuess
    xlGuess = 0
    xlYes = 1
    xlNo = 2
End Enum

Public Enum xlSortDataOption
    xlSortNormal = 0
    xlSortTextAsNumbers = 1
End Enum

Public Enum xlDirection
    xlDown = -4121                          ' (&HFFFFEFE7)
    xlToLeft = -4159                        ' (&HFFFFEFC1)
    xlToRight = -4161                       ' (&HFFFFEFBF)
    xlUp = -4162                            ' (&HFFFFEFBE)
End Enum

Public Enum xlPasteType
    xlPasteAll = -4104                      ' (&HFFFFEFF8)
    xlPasteAllExceptBorders = 7
    xlPasteColumnWidths = 8
    xlPasteComments = -4144                 ' (&HFFFFEFD0)
    xlPasteFormats = -4122                  ' (&HFFFFEFE6)
    xlPasteFormulas = -4123                 ' (&HFFFFEFE5)
    xlPasteFormulasAndNumberFormats = 11
    xlPasteValidation = 6
    xlPasteValues = -4163                   ' (&HFFFFEFBD)
    xlPasteValuesAndNumberFormats = 12
End Enum

Public Enum xlChartType
    xl3DArea = -4098                        ' (&HFFFFEFFE)
    xl3DAreaStacked = 78                    ' (&H4E)
    xl3DAreaStacked100 = 79                 ' (&H4F)
    xl3DBarClustered = 60                   ' (&H3C)
    xl3DBarStacked = 61                     ' (&H3D)
    xl3DBarStacked100 = 62                  ' (&H3E)
    xl3DColumn = -4100                      ' (&HFFFFEFFC)
    xl3DColumnClustered = 54                ' (&H36)
    xl3DColumnStacked = 55                  ' (&H37)
    xl3DColumnStacked100 = 56               ' (&H38)
    xl3DLine = -4101                        ' (&HFFFFEFFB)
    xl3DPie = -4102                         ' (&HFFFFEFFA)
    xl3DPieExploded = 70                    ' (&H46)
    xlArea = 1
    xlAreaStacked = 76                      ' (&H4C)
    xlAreaStacked100 = 77                   ' (&H4D)
    xlBarClustered = 57                     ' (&H39)
    xlBarOfPie = 71                         ' (&H47)
    xlBarStacked = 58                       ' (&H3A)
    xlBarStacked100 = 59                    ' (&H3B)
    xlBubble = 15
    xlBubble3DEffect = 87                   ' (&H57)
    xlColumnClustered = 51                  ' (&H33)
    xlColumnStacked = 52                    ' (&H34)
    xlColumnStacked100 = 53                 ' (&H35)
    xlConeBarClustered = 102                ' (&H66)
    xlConeBarStacked = 103                  ' (&H67)
    xlConeBarStacked100 = 104               ' (&H68)
    xlConeCol = 105                         ' (&H69)
    xlConeColClustered = 99                 ' (&H63)
    xlConeColStacked = 100                  ' (&H64)
    xlConeColStacked100 = 101               ' (&H65)
    xlCylinderBarClustered = 95             ' (&H5F)
    xlCylinderBarStacked = 96               ' (&H60)
    xlCylinderBarStacked100 = 97            ' (&H61)
    xlCylinderCol = 98                      ' (&H62)
    xlCylinderColClustered = 92             ' (&H5C)
    xlCylinderColStacked = 93               ' (&H5D)
    xlCylinderColStacked100 = 94            ' (&H5E)
    xlDoughnut = -4120                      ' (&HFFFFEFE8)
    xlDoughnutExploded = 80                 ' (&H50)
    xlLine = 4
    xlLineMarkers = 65                      ' (&H41)
    xlLineMarkersStacked = 66               ' (&H42)
    xlLineMarkersStacked100 = 67            ' (&H43)
    xlLineStacked = 63                      ' (&H3F)
    xlLineStacked100 = 64                   ' (&H40)
    xlPie = 5
    xlPieExploded = 69                      ' (&H45)
    xlPieOfPie = 68                         ' (&H44)
    xlPyramidBarClustered = 109             ' (&H6D)
    xlPyramidBarStacked = 110               ' (&H6E)
    xlPyramidBarStacked100 = 111            ' (&H6F)
    xlPyramidCol = 112                      ' (&H70)
    xlPyramidColClustered = 106             ' (&H6A)
    xlPyramidColStacked = 107               ' (&H6B)
    xlPyramidColStacked100 = 108            ' (&H6C)
    xlRadar = -4151                         ' (&HFFFFEFC9)
    xlRadarFilled = 82                      ' (&H52)
    xlRadarMarkers = 81                     ' (&H51)
    xlStockHLC = 88                         ' (&H58)
    xlStockOHLC = 89                        ' (&H59)
    xlStockVHLC = 90                        ' (&H5A)
    xlSurface = 83                          ' (&H53)
    xlSurfaceTopView = 85                   ' (&H55)
    xlSurfaceTopViewWireframe = 86          ' (&H56)
    xlSurfaceWireframe = 84                 ' (&H54)
    xlXYScatter = -4169                     ' (&HFFFFEFB7)
    xlXYScatterLines = 74                   ' (&H4A)
    xlXYScatterLinesNoMarkers = 75          ' (&H4B)
    xlXYScatterSmooth = 72                  ' (&H48)
    xlXYScatterSmoothNoMarkers = 73         ' (&H49)
End Enum

Public Enum xlThemeColor
    xlThemeColorDark1 = 1
    xlThemeColorLight1 = 2
    xlThemeColorDark2 = 3
    xlThemeColorLight2 = 4
    xlThemeColorAccent1 = 5
    xlThemeColorAccent2 = 6
    xlThemeColorAccent3 = 7
    xlThemeColorAccent4 = 8
    xlThemeColorAccent5 = 9
    xlThemeColorAccent6 = 10
    xlThemeColorHyperlink = 11
    xlThemeColorFollowedHyperlink = 12
End Enum

Private Enum xlThemeFont
    xlThemeFontNone = 0
    xlThemeFontMajor = 1
    xlThemeFontMinor = 2
End Enum

Public Enum enumMSExcelFileStatus
    xlSuccess = -1
    xlLockedByUser = 0
    xlFileNotReplaced = 1
End Enum

Public Enum msoDocProperties
    msoPropertyTypeNumber = 1
    msoPropertyTypeBoolean = 2
    msoPropertyTypeDate = 3
    msoPropertyTypeString = 4
End Enum

Public Enum msoScaleFrom
    msoScaleFromBottomRight = 2
    msoScaleFromMiddle = 1
    msoScaleFromTopLeft = 0
End Enum

Public Enum msoTriState
    msoCTrue = 1
    msoFalse = 0
    msoTriStateMixed = -2                   ' (&HFFFFFFFE)
    msoTriStateToggle = -3                  ' (&HFFFFFFFD)
    msoTrue = -1                            ' (&HFFFFFFFF)
End Enum

Public Enum enumEMailRecipientType
    eMailTo = 1
    eMailCC = 2
    eMailCCO = 3
End Enum

'-----------------------------------------------------------------------------------------------------------------------
' Función para cerrar un formulario empleando su nombre.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       sFormName   String      Nombre del formulario a cerrar.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function CloseForm(ByVal sFormName As String) As Boolean
    On Error GoTo Error_CloseForm

    DoCmd.Close acForm, sFormName, acSaveNo
    CloseForm = (Not modMDB.IsLoaded(sFormName))

    Exit Function

Error_CloseForm:
    GetError "modMDB.CloseForm", "Form=" & sFormName
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para cerrar el formulario con la barra de progreso.
'-----------------------------------------------------------------------------------------------------------------------
Public Sub CloseProgressBar()
    If modMDB.IsLoaded("(Form) Estado Proceso") Then DoCmd.Close acForm, "(Form) Estado Proceso", acSaveNo: DoEvents
End Sub

'-----------------------------------------------------------------------------------------------------------------------
' Método para concatenar una cadena de texto a una variable.
'-----------------------------------------------------------------------------------------------------------------------
' Parámetros:
'
'   sVariable   String      Variable a la que se concatenará.
'
'   sString     String      Cadena de texto a concatenar.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Sub Concat(ByRef sVariable As String, _
                  ByVal sString As String)
    sVariable = sVariable & sString
End Sub

'-----------------------------------------------------------------------------------------------------------------------
' Función para convertir un valor al formato adecuado para una sentencia SQL.
'-----------------------------------------------------------------------------------------------------------------------
'   NOTA:   Esta función es genérica para cualquier tipo de valor.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       vValue      Variant             Valor a convertir a formato de sentencia SQL.
'
'       bReplace    Boolean (Opcional)  Indicador de si el valor es cadena de texto, hay que reemplazar.
'                                       los carácteres que entren en conficto con la sintaxis SQL.
'                                       (Valor por defecto = True)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function CSql(ByVal vValue As Variant, _
                     Optional ByVal bReplace As Boolean = True) As String
    On Error GoTo Error_CSql

    If (IsDate(vValue) And (VarType(vValue) = vbDate)) Then ' 1. Si es una fecha
        CSql = "#" & Format(vValue, IIf((Int(CDbl(vValue)) < CDbl(vValue)), "mm/dd/yyyy hh:mm:ss", "mm/dd/yyyy")) & "#"
    ElseIf (VarType(vValue) = vbBoolean) Then               ' 2. Si es un Boolean
        CSql = IIf(vValue, "True", "False")
    ElseIf (VarType(vValue) = vbString) Then                ' 3. Si es una cadena de texto
        If bReplace Then
            CSql = "'" & modSystem.ReplStr(modSystem.ReplStr(CStr(vValue), "[#]", "Ñ"), "[']", "''") & "'"
        Else
            CSql = "'" & CStr(vValue) & "'"
        End If
    ElseIf IsNumeric(vValue) Then                           ' 4. Si es un número
        CSql = Replace(CStr(CDbl(Nz(vValue, 0))), ",", ".")
    End If

    Exit Function

Error_CSql:
    GetError "modMDB.CSql"
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para convertir un booleano a formato de sentencia SQL.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       bValue      Boolean             Valor booleano a convertir a formato de sentencia SQL.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function CSqlBool(ByVal bValue As Boolean) As String
    CSqlBool = IIf(bValue, "True", "False")
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para convertir un valor numérico a formato de número de sentencia SQL.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       vValue      Variant             Valor booleano a convertir a formato de sentencia SQL.
'
'       bGetZero    Boolean (Opcional)  Indicador de si se debe devolver valor cero o nulo.
'                                       (Por defecto = False)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function CSqlDbl(ByVal vValue As Variant, _
                        Optional ByVal bGetZero As Boolean = False, _
                        Optional ByVal iDecimals As Integer = 0) As String
    If Not IsZr(vValue) Then
        CSqlDbl = Replace(CStr((Int(CDec(((Nz(vValue, 0) * (10 ^ iDecimals)) + 0.5))) / (10 ^ iDecimals))), ",", ".")
    Else
        CSqlDbl = IIf(((vValue = 0) And bGetZero), "0", "Null")
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para convertir una fecha a formato de sentencia SQL.
'-----------------------------------------------------------------------------------------------------------------------
'   NOTA:   Las fechas se manejan internamente como números de doble coma flotante.
'           Cuando se presentan en las aplicaciones, lo hacen con el formato
'           de la configuración regional del sistema operativo. Microsoft Access
'           interpreta siempre las fechas en sentencias SQL con el formato Inglés
'           (mm/dd/yyyy). Cuando se manejan las fechas como número de doble
'           coma flotante, la parte entera del mismo corresponde con la fecha, y la
'           parte decimal con la hora.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       vValue          Variant             Valor de fecha a convertir a formato de sentencia SQL.
'
'       bIsDateTime     Boolean (Opcional)  Indicador de si se debe devolver la parte correspondiente a la hora.
'                                           (Por defecto = False)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function CSqlDt(ByVal vValue As Variant, _
                       Optional ByVal bIsDateTime As Boolean = False) As String
    If Not IsZrDt(vValue) Then
        vValue = CDate(vValue)
        CSqlDt = "#" & Format(vValue, IIf((Not bIsDateTime) Or IsZr(CDbl(vValue) - Int(CDbl(vValue))), "mm/dd/yyyy", "mm/dd/yyyy hh:nn:ss")) & "#"
    Else
        CSqlDt = "Null"
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para convertir un valor a formato de texto de sentencia SQL.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       vValue      Variant             Valor a convertir a formato de literal en sentencia SQL.
'
'       bStripStr   Boolean (Opcional)  Indicador de si se debe limpiar el literal de caracteres
'                                       conflictivos con el lenguaje SQL. (Por defecto = False)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function CSqlTxt(ByVal vValue As Variant, _
                        Optional ByVal bStripStr As Boolean = False) As String
    If Not IsZrStr(vValue) Then
        If bStripStr Then
            CSqlTxt = "'" & modSystem.StripStr(Trim(CStr(vValue)), enumStripStr.SQL) & "'"
        Else
            CSqlTxt = "'" & Trim(CStr(vValue)) & "'"
        End If
    Else
        CSqlTxt = "Null"
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para vaciar una tabla si tuviera datos.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       sTableDef   String                  Nombre de la tabla a vaciar de registros.
'
'       dbDatabase  Database (Opcional)     Base de datos a la que pertenece la tabla.
'                                           (Por defecto = CurrentDB)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function DeleteIfHasRows(ByVal sTableDef As String) As Boolean
    If modMDB.HasRows(sTableDef, acTable) Then CurrentDb.Execute "DELETE FROM " & sTableDef & ";", dbFailOnError
    DeleteIfHasRows = True
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para obtener la letra de una columna de Microsoft Excel.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       Index        Byte               Índice de la columna (de 0 a n - 1)).
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function GetCell(ByVal Index As Byte) As String
    GetCell = IIf((Int(Index / 26) > 0), Chr(65 + Int((Index / 26) - 1)), vbNullString) & Chr(65 + (Index Mod 26))
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para obtener el último valor insertado en un campo autonumérico.
'-----------------------------------------------------------------------------------------------------------------------
Public Function GetKeyId() As Long
    GetKeyId = CurrentDb.OpenRecordset("SELECT @@IDENTITY AS KeyID;")!KeyID
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para obtener el nombre de fichero de la base de datos.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       dbDatabase     Database  (Opcional)     Base de datos a obtener el nombre del fichero. (Por defecto = CurrentDB)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function GetMDBFileName() As String
    GetMDBFileName = modSystem.GetFileFullName(CurrentDb.Name)
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para obtener el nombre de la base de datos.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       dbDatabase     Database  (Opcional)     Base de datos a obtener el nombre de la base de datos. (Por defecto = CurrentDB)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function GetMDBName() As String
    GetMDBName = modSystem.GetFileName(CurrentDb.Name)
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para devolver la ruta de la carpeta de ubicación de una base de datos o de una subcarpeta.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       sSubFolder      String    (Opcional)    Subcarpeta a obtener. (Por defecto = "")
'
'       dbDatabase      Database  (Opcional)    Base de datos a obtener el nombre su ruta de ubicación.
'                                               (Por defecto = CurrentDB)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function GetMDBPath(Optional ByVal sSubFolder As String = vbNullString) As String
    GetMDBPath = modSystem.GetFolderPath(CurrentDb.Name) & IIf(Not IsZrStr(sSubFolder), sSubFolder & "\", vbNullString)
    If Not modSystem.FolderExists(GetMDBPath) Then modSystem.BuiltPath GetMDBPath
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para obtener el número de registros a los que afecta una consulta, o contenidos en una tabla.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       sSource         String                  Sentencia SQL, o nombre de la tabla de donde comprobar los resgistros devueltos
'                                               o que contiene.
'
'       lObjectType     Long                    Tipo de objeto que es el origen.
'
'                                               Valores posibles:
'
'                                                   Access.AcObjectType.acQuery             Consulta de la base de datos. Access.
'
'                                                   Access.AcObjectType.acStoredProcedure   Sentencia SQl.
'
'                                                   Access.AcObjectType.acTable             Tabla de la base de datos
'
'       dbDatabase      Database  (Opcional)    Base de datos a obtener el nombre su ruta de ubicación.
'                                               (Por defecto = CurrentDB)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function GetNumRecords(ByVal sSource As String, _
                              ByVal lObjectType As Access.AcObjectType) As Long
    On Error GoTo Error_GetNumRecords

    If Not IsZrStr(sSource) Then
        Select Case lObjectType
            Case Access.AcObjectType.acQuery, _
                 Access.AcObjectType.acStoredProcedure

                Dim m_Rst       As Recordset

                Set m_Rst = CurrentDb.OpenRecordset(sSource)
                If modMDB.IsRst(m_Rst, True) Then
                    With m_Rst
                        GetNumRecords = .RecordCount
                        .Close
                    End With
                End If
                Set m_Rst = Nothing

            Case Access.AcObjectType.acTable
                
                With CurrentDb.TableDefs(sSource).OpenRecordset
                    If Not IsZr(.RecordCount) Then .MoveLast
                    GetNumRecords = .RecordCount
                End With
        End Select
    End If

    Exit Function

Error_GetNumRecords:
    GetError "modMDB.GetNumRecords"
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para comprobar si existe un determinado objeto en la base de datos.
'-----------------------------------------------------------------------------------------------------------------------
' Parámetros:
'
'       oControl    Control                 Controla bloquear/desbloquear.
'
'       bLocked     Boolean  (Opcional)     Indicador de si se bloquea/desbloquea. (Por defecto = False)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function ObjectExists(ByVal sObjectName As String, _
                             Optional ByVal lObjectType As Access.AcObjectType = Access.AcObjectType.acTable) As Boolean
    On Error GoTo Error_ObjectExists

    If IsZrStr(sObjectName) Then Err.Raise vbObjectError + 500, GetMDBName, "No se ha especificado el nombre del objeto a comprobar."
    Select Case lObjectType
        Case Access.AcObjectType.acForm     ' Si se trata de un Formulario
            ObjectExists = (Application.CodeProject.AllForms(sObjectName).Name = sObjectName)

        Case Access.AcObjectType.acMacro    ' Si se trata de una Macro
            ObjectExists = (Application.CodeProject.AllMacros(sObjectName).Name = sObjectName)

        Case Access.AcObjectType.acQuery    ' Si se trata de una Consulta
            ObjectExists = (CurrentDb.QueryDefs(sObjectName).Name = sObjectName)

        Case Access.AcObjectType.acReport   ' Si se trata de un Informe
            ObjectExists = (Application.Reports(sObjectName).Name = sObjectName)

        Case Access.AcObjectType.acTable    ' Si se trata de una Tabla
            ObjectExists = (CurrentDb.TableDefs(sObjectName).Name = sObjectName)
    End Select

    Exit Function

Error_ObjectExists:
    ' Si el error corresponde a que no existe el elemento a buscar
    If (Err.Number = 2451) Or _
       (Err.Number = 2467) Or _
       (Err.Number = 3265) Then
        Err.Clear
        Exit Function
    Else
        GetError "modMDB.ObjectExists"
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para determinar si una tabla o una consulta SQL contiene registros.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       sSource         String                          Nombre del origen de los registros.
'
'       lObjectType     Access.AcObjectType (Opcional)  Posición de palabra de la palabra a obtener dentro de la cadena
'                                                   de texto. (Por defecto  = 1)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function HasRows(ByVal sSource As String, _
                        Optional ByVal lObjectType As Access.AcObjectType = Access.AcObjectType.acTable) As Boolean

    On Error GoTo Error_HasRows

    If Not IsZrStr(sSource) Then
        Select Case lObjectType
            Case Access.AcObjectType.acTable
                HasRows = (Not IsZr(modMDB.GetNumRecords(sSource, lObjectType)))

            Case Access.AcObjectType.acQuery
                HasRows = (Not IsZr(CurrentDb.QueryDefs(sSource).ReturnsRecords))

            Case Access.AcObjectType.acStoredProcedure
                Dim m_Rst       As Recordset

                Set m_Rst = CurrentDb.OpenRecordset(sSource)
                If modMDB.IsRst(m_Rst, True) Then
                    With m_Rst
                        HasRows = (Not IsZr(.RecordCount))
                        .Close
                    End With
                End If
                Set m_Rst = Nothing
        End Select
    End If

    Exit Function

Error_HasRows:
    ' Si el error corresponde a que no existe el elemento a buscar
    If (Err.Number = 2451) Or _
       (Err.Number = 2467) Or _
       (Err.Number = 3265) Then
        Err.Clear
        Exit Function
    Else
        GetError "modMDB.HasRows"
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para incrementar la barra de progreso.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       lIncrement      Long        (Opcional)      Incremento a aplicar a la barra de progreso. (Por defecto = 1)
'
'       sProcess        String      (Opcional)      Nombre del proceso. (Por defecto = "")
'
'       bCounter        Boolean     (Opcional)      Indicador de si se emplea en contador propio de la barra de progreso.
'                                                   (Por defecto = False)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Sub IncrProgressBar(Optional ByVal lIncrement As Long = 1, _
                           Optional ByVal sProcess As String = vbNullString, _
                           Optional ByVal bCounter As Boolean = False)
    If modMDB.IsLoaded("(Form) Estado Proceso") Then
        With Forms("(Form) Estado Proceso").Form
            If ((.pbProceso.Value + lIncrement) <= .pbProceso.Max) Then .pbProceso.Value = (.pbProceso.Value + lIncrement)
            .Completado_Etiqueta.Caption = "Completado: " & Round(((.pbProceso.Value / .pbProceso.Max) * 100), 0) & "%"
            If Not IsZrStr(sProcess) Then .Proceso_Etiqueta.Caption = sProcess & IIf(bCounter, " (" & .pbProceso.Value & " de " & .pbProceso.Max & ")", vbNullString)
        End With
        DoEvents
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------
' Función para validar si un formulario ya está en ejecución. (cargado en memoria)
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       sFormName       String                      Nombre del formulario a comprobar.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function IsLoaded(ByVal sFormName As String) As Boolean
    IsLoaded = Application.CodeProject.AllForms(sFormName).IsLoaded
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para validar si un objeto es un recordset valido.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       oRst            Recordset       Objeto Recordset a validar.
'
'       bRefresh        Boolean         Indicador de si se debe refrescar el objeto ADO.
'                                       (Por defecto = False)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function IsRst(ByRef oRst As Recordset, _
                      Optional ByVal bRefresh As Boolean = False) As Boolean
    If Not oRst Is Nothing Then
        IsRst = Not (oRst.BOF And oRst.EOF)
        If (IsRst And bRefresh) Then
            With oRst
                .MoveLast
                .MoveFirst
            End With
        End If
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para bloquear/desbloquear un control de formulario.
'-----------------------------------------------------------------------------------------------------------------------
' Parámetros:
'
'       oControl      Control            Control a bloquear/desbloquear.
'
'       bLocked     Boolean  (Opcional)  Indicador de si se bloquea/desbloquea. (Por defecto = False)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Sub LockControl(ByVal oControl As Control, _
                       Optional ByVal bLocked As Boolean = False)
    With oControl
        ' 1. Bloqueamos el control para que no se pueda editar
        .Locked = bLocked

        ' 2. Establecemos el color de fondo del control
        .BackColor = IIf(bLocked, enumBackgroundStatusColor.DisabledControl, enumBackgroundStatusColor.EnabledControl)
    End With
End Sub

'-----------------------------------------------------------------------------------------------------------------------
' Función para abrir un formulario empleando su nombre y pasándole parámetros como argumentos.
'-----------------------------------------------------------------------------------------------------------------------
' Parámetros:
'
'       sFormName       String              Nombre del formulario a abrir.
'
'       sOpenArgs       String  (Opcional)  Argumentos de apertura. (Por defecto = "")
'
'       bPreview        String  (Opcional)  Indicador del modo de apertura del formulario. (Por defecto = False)
'
'       bCheckIsLoaded  Boolean (Opcional)  Indicador de comprobación de apertura del formulario. (Por defecto = True)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function OpenForm(ByVal sFormName As String, _
                         Optional ByVal sOpenArgs As String = vbNullString, _
                         Optional ByVal bPreview As Boolean = False, _
                         Optional ByVal bCheckIsLoaded As Boolean = True) As Form
    On Error GoTo Error_OpenForm

    ' 1. Si no está abierto el formulario, lo abrimos
    If Not modMDB.IsLoaded(sFormName) Then DoCmd.OpenForm sFormName, IIf(bPreview, acPreview, acNormal), , , , , sOpenArgs

    ' 2. Si se indicó comprobar que esté abierto el formulario
    If bCheckIsLoaded Then
        ' 3. Si está abierto el formulario
        If modMDB.IsLoaded(sFormName) Then
            ' 4. Devolvemos el formulario
            Set OpenForm = Forms(sFormName)
        Else
            ' 5. Generamos error
            'Err.Raise (vbObjectError + 101), modMDB.GetMDBName(), "No se ha podido abrir el formulario [" & sFormName & "]."
        End If
    End If

    Exit Function

Error_OpenForm:
    GetError "modMDB.OpenForm", "Form=" & sFormName & vbTab & "OpenArgs=" & sOpenArgs
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para eliminar un objeto de la base de datos solo si este existe.
'-----------------------------------------------------------------------------------------------------------------------
' Parámetros:
'
'       sObjectName     String                          Nombre del objeto a eliminar.
'
'       lObjectType     Access.AcObjectType (Opcional)  Tipo de objeto a eliminar.
'                                                       (Por defecto = Access.AcObjectType.acTable)
'
'       dbDatabase      Database            (Opcional)  Base de datos de donde eliminar el objeto.
'                                                       (Por defecto = CurrentDB)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function RemoveIfExists(ByVal sObjectName As String, _
                               Optional ByVal lObjectType As Access.AcObjectType = Access.AcObjectType.acTable) As Boolean
    On Error GoTo Error_RemoveIfExists

    If IsZrStr(sObjectName) Then Err.Raise (vbObjectError + 500), GetMDBName, "No se ha especificado el nombre del objeto a eliminar."
    Select Case lObjectType
        Case Access.AcObjectType.acQuery
            If modMDB.ObjectExists(sObjectName, lObjectType) Then CurrentDb.QueryDefs.Delete sObjectName

        Case Access.AcObjectType.acTable
            If modMDB.ObjectExists(sObjectName, lObjectType) Then CurrentDb.TableDefs.Delete sObjectName
    End Select
    RemoveIfExists = True
    
    Exit Function

Error_RemoveIfExists:
    GetError "modMDB.RemoveIfExists", sObjectName
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Método para establecer la máscara de entrada para un campo de tipo texto.
'-----------------------------------------------------------------------------------------------------------------------
' Parámetros:
'
'       oTextBox        TextBox                         Caja de texto a la que aplicar la máscara de entrada
'
'       iLength         Integer             (Opcional)  Longitud de la máscara. (Por defecto 50 caracteres)
'
'       lLetterCase     enumTypeLetterCase  (Opcional)  Indicador de si la máscara emplea mayúsculas, minúsculas, o ambas.
'                                                       (Por defecto BothCase = (Mayúsculas y minúsculas))
'
'-----------------------------------------------------------------------------------------------------------------------
Public Sub SetInputMask(ByVal oTextBox As TextBox, _
                        Optional ByVal iLength As Integer = 50, _
                        Optional ByVal lLetterCase As enumTypeLetterCase = enumTypeLetterCase.BothCase)
    Select Case lLetterCase
        Case enumTypeLetterCase.BothCase
            oTextBox.InputMask = "L" & String((iLength - 1), "?") & ";;_"

        Case enumTypeLetterCase.LowCase
            oTextBox.InputMask = "<L" & String((iLength - 1), "?") & ";;_"

        Case enumTypeLetterCase.UpperCase
            oTextBox.InputMask = ">L" & String((iLength - 1), "?") & ";;_"
    End Select
End Sub

'-----------------------------------------------------------------------------------------------------------------------
' Función para devolver una cadena de texto recortada si supera una determinada longitud.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       sValue          String      Cadena de texto a recortar.
'
'       iMaxLength      Intergee    Logitud máxima de caracteres.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function GetMaxLength(ByVal sValue As String, _
                             ByVal iMaxLength As Integer) As String
    GetMaxLength = IIf((Len(sValue) > iMaxLength), Left(sValue, iMaxLength), sValue)
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Método para establecer el formato de la barra de progreso
'-----------------------------------------------------------------------------------------------------------------------
' Parámetros:
'
'   lMinValue   Long    (Optional)  Valor mínimo de la barra de progreso.
'
'   lMaxValue   Long    (Optional)  Valor máximo de la barra de progreso.
'
'   sTitle      String   (Optional)  Texto de título de la venta de progreso.
'
'   sCaption    String   (Optional)  Texto a mostrar en la barra de progreso.
'
'   lValue      lValue   (Optional)  Valor de la barra de progreso.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Sub SetProgressBar(Optional ByVal lMinValue As Long = 0, _
                          Optional ByVal lMaxValue As Long = 100, _
                          Optional ByVal sTitle As String = vbNullString, _
                          Optional ByVal sCaption As String = vbNullString, _
                          Optional ByVal lValue As Long = 0)
    With modMDB.OpenForm("(Form) Estado Proceso")
        If Not IsZrStr(sTitle) Then .Caption = sTitle
        If Not IsZrStr(sCaption) Then .Proceso_Etiqueta.Caption = sCaption
        With .pbProceso
            .Min = lMinValue
            .Max = lMaxValue
            .Value = lValue
        End With
    End With
    DoEvents
End Sub

'-----------------------------------------------------------------------------------------------------------------------
' Función para establecer una relación entre tablas de una base de datos.
'-----------------------------------------------------------------------------------------------------------------------
' Parámetros:
'
'
'       dbDatabase          Database                Base de datos a la que pertenecen las tablas a relaciones.
'
'       sTableDef           String                  Nombre de la tabla padre.
'
'       sForeignTableDef    String                  Nombre de la tabla hija.
'
'       sFields             String                  Relación de campos que componen la relación.
'
'       lAttributes         Long       (Opcional)   Indicador del tipo de relación a establecer.
'                                                   (POr defecto actualizar y eliminar datos en cascada)
'
'       sAlternateName      String     (Opcional)   Indicador de si se debe establecer la consulta como oculta.
'                                                   (Por defecto False)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function SetRelation(ByVal sTableDef As String, _
                            ByVal sForeignTableDef As String, _
                            ByVal sFields As String, _
                            Optional ByVal lAttributes As Long = 0, _
                            Optional ByVal sAlternateName As String = vbNullString) As Boolean

    Dim m_Rel               As Relation
    Dim m_Fld               As Field
    Dim m_sFields()         As String
    Dim m_lIdx              As Long

    On Error GoTo Error_SetRelation

    ' 1. Si se especificaron los campos de enlace
    If Not IsZrStr(sFields) Then
        ' 2. Si no se ha definido un nombre específico
        '    a la relación de las tablas, le creamos uno
        If IsZrStr(sAlternateName) Then sAlternateName = modMDB.GetMaxLength("PK_" & sTableDef & "_" & sForeignTableDef, 64)

        ' 3. Creamos la relación
        Set m_Rel = CurrentDb.CreateRelation(sAlternateName, sTableDef, sForeignTableDef, lAttributes)

        ' 4. Establecemos los campos de enlace de la relación
        m_sFields = Split(sFields, ";")
        For m_lIdx = LBound(m_sFields) To UBound(m_sFields)
            If Not IsZrStr(m_sFields(m_lIdx)) Then
                Set m_Fld = m_Rel.CreateField(Split(m_sFields(m_lIdx), ":")(0))
                m_Fld.ForeignName = Split(m_sFields(m_lIdx), ":")(1)
                m_Rel.Fields.Append m_Fld
            End If
        Next m_lIdx

        ' 5. Agregamos la relación creada a la base de datos
        If Not IsZr(m_Rel.Fields.Count) Then
            With CurrentDb.Relations
                .Append m_Rel
                .Refresh
            End With
        End If
    End If

    Set m_Fld = Nothing
    Set m_Rel = Nothing

    SetRelation = True

    Exit Function

Error_SetRelation:
    If (Err.Number = 3265) Then
        Err.Clear
        Resume Next
    Else
        Set m_Fld = Nothing
        Set m_Rel = Nothing
        GetError "modMDB.SetRelation", sTableDef
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para exportar los resultados de una consulta a un archivo Excel.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       sSQL                String                      Sentencia SQL.
'
'       sWorkSheet          String                      Nombre de la hoja de datos donde se volcará los datos de
'                                                       la consulta.
'
'       sXLSFilePath        String                      Ruta de destino del archivo Excel a generar.
'
'       bReplace            Boolean (Opcional)          Indicador de si se debe reemplazar el fichero de destino,
'                                                       en el caso de que ya exista en la misma ruta.
'                                                       (Por defecto = False)
'
'       bShow               Boolean (Opcional)          Indicador de si se debe abrir el fichero MS Excel para mostrarlo,
'                                                       una vez generado. (Por defecto = False)
'
'       sTitle              String (Opcional)           Título del fichero MS Excel. (Por defecto = "")
'
'       sSubject            String (Opcional)           Asunto del fichero MS Excel. (Por defecto = "")
'
'       sCategory           String (Opcional)           Categoría del fichero MS Excel. (Por defecto = "")
'
'       sComments           String (Opcional)           Comentarios del fichero MS Excel. (Por defecto = "")
'
'       bAutoFilter         Boolean (Opcional)          Indicador de si se debe aplicar autofiltro a la hoja de cálculo
'                                                       generada. (Por defecto = False)
'
'       colDocProperties    colDocProperties (Opcional) Colección de propiedades de documento a aplicar al nuevo
'                                                       fichero MS Excel. (Por defecto = Nothing)
'
'       bControlDevelopment Boolean          (Opcional) Indicador de si se emplea control de desarrollo.
'                                                       (Por defecto = True)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function ToMSExcel(ByVal sSQL As String, _
                          ByVal sWorkSheet As String, _
                          ByRef sXLSFilePath As String, _
                          Optional ByVal bReplace As Boolean = False, _
                          Optional ByVal bShow As Boolean = False, _
                          Optional ByVal sTitle As String = vbNullString, _
                          Optional ByVal sSubject As String = vbNullString, _
                          Optional ByVal sCategory As String = vbNullString, _
                          Optional ByVal sComments As String = vbNullString, _
                          Optional ByVal bAutoFilter As Boolean = False, _
                          Optional ByVal colDocProperties As colDocProperties = Nothing, _
                          Optional ByVal bShowTotals As Boolean = True) As enumMSExcelFileStatus

    Dim m_sWkst         As String

    On Error GoTo Error_ToMSExcel

    ' 1. Damos formato al nombre de la hoja con los datos de la consulta
    m_sWkst = Mid(Replace(sWorkSheet, " ", "_"), 1, 31)

    ' 2. Si existe el fichero MS Excel a generar
    If modSystem.FileExists(sXLSFilePath) Then
        ' 3. Si no se indicó la posibilidad de reemplazar el fichero anterior con el nuevo
        If Not bReplace Then
            ' 4. Mostramos mensaje al usuario
            MsgBox "Ya existe el fichero MS Excel en la ruta indicada, y no puede ser reemplazado.", (vbOKOnly + vbExclamation), "Fichero MS Excel ya existe"

            ToMSExcel = enumMSExcelFileStatus.xlFileNotReplaced

            Exit Function
        Else
            ' 7. Si no se pudo eliminar el fichero antiguo
            If Not modSystem.DeleteFile(sXLSFilePath) Then
                ' 8. Mostramos mensaje al usuario
                MsgBox "El fichero MS Excel ya se encuentra en la ruta establecida, y no puede ser reemplazado, dado que está abierto por algún usuario." & vbCrLf & vbCrLf & "Compruebe que no sea usted mismo quién lo tenga abierto.", (vbOKOnly + vbExclamation), "Fichero MS Excel bloqueado"

                ToMSExcel = enumMSExcelFileStatus.xlLockedByUser

                Exit Function
            End If
        End If
    End If

    ' 9. Si no existe la consulta para generar el fichero MS Excel
    If Not modMDB.ObjectExists(m_sWkst, acQuery) Then
        Dim m_QueryDef  As New QueryDef

        ' 10. Se genera la consulta en la base de datos
        With m_QueryDef
            .Name = m_sWkst
            .SQL = sSQL
        End With

        With CurrentDb.QueryDefs
            .Append m_QueryDef
            .Refresh
        End With

        Set m_QueryDef = Nothing
    Else
        ' 11. Se establece la sentencia SQL de la consulta
        CurrentDb.QueryDefs(m_sWkst).SQL = sSQL
    End If

    With CurrentDb
        With .QueryDefs(m_sWkst)
            ' 12. Se exporta el resultado de la consulta a un fichero MS Excel
            DoCmd.TransferSpreadsheet acExport, modMDB.GetMSExcelDefaultWorkbookFormat(), .Name, sXLSFilePath, True

            ' 13. Indicamos que se realizó la importación
            If modSystem.FileExists(sXLSFilePath) Then
                Dim m_Excel         As Object
                Dim m_lRows         As Long
                Dim m_lRow          As Long
                Dim m_lCols         As Long

                ' 14. Obtenemos el número total de registros y columnas
                m_lRows = modMDB.GetNumRecords(m_sWkst, acQuery)
                m_lCols = CurrentDb.QueryDefs(m_sWkst).Fields.Count

                ' 15. Abrimos el fichero MS Excel generado
                Set m_Excel = modMDB.GetMSExcelApplication()
                With m_Excel
                    .DisplayAlerts = False
                    .Workbooks.Open sXLSFilePath
                    With .Workbooks(1)
                        With .Worksheets(.Worksheets.Count)
                            .Name = sWorkSheet

                            ' 16. Si se indicó, establecemos el autofiltro
                            If bAutoFilter Then .Rows("1:1").AutoFilter

                            ' 17. Damos formato a la cabecera de resultados
                            With .Range(modMDB.GetCell(0) & 1 & ":" & modMDB.GetCell(m_lCols - 1) & 1)
                                With .Interior
                                    .Pattern = xlSolid
                                    .PatternColorIndex = xlAutomatic
                                    .ThemeColor = xlThemeColor.xlThemeColorAccent1
                                    .TintAndShade = 0.799981688894314
                                    .PatternTintAndShade = 0
                                End With

                                .MergeCells = False
                                .HorizontalAlignment = xlAlignment.xlCenter
                                .VerticalAlignment = xlAlignment.xlTop
                                .WrapText = False
                                .Orientation = 0
                                .ShrinkToFit = False
                                .EntireColumn.AutoFit

                                With .Font
                                    .Name = "Arial"
                                    .FontStyle = "Normal"
                                    .Size = 10
                                    .Strikethrough = False
                                    .Superscript = False
                                    .Subscript = False
                                    .OutlineFont = False
                                    .Shadow = False
                                    .Bold = True
                                    .Underline = xlUnderlineStyle.xlUnderlineStyleNone
                                    .ColorIndex = xlAutomatic
                                End With

                                .Borders(xlAlignment.xlDiagonalDown).LineStyle = xlNone
                                .Borders(xlAlignment.xlDiagonalUp).LineStyle = xlNone

                                With .Borders(xlAlignment.xlEdgeLeft)
                                    .LineStyle = xlLineStyle.xlContinuous
                                    .ColorIndex = xlAutomatic
                                    .TintAndShade = 0
                                    .Weight = xlBorderWeight.xlThin
                                End With

                                With .Borders(xlAlignment.xlEdgeTop)
                                    .LineStyle = xlLineStyle.xlContinuous
                                    .ColorIndex = xlAutomatic
                                    .TintAndShade = 0
                                    .Weight = xlBorderWeight.xlThin
                                End With

                                With .Borders(xlAlignment.xlEdgeBottom)
                                    .LineStyle = xlLineStyle.xlContinuous
                                    .ColorIndex = xlAutomatic
                                    .TintAndShade = 0
                                    .Weight = xlBorderWeight.xlThin
                                End With

                                With .Borders(xlAlignment.xlEdgeRight)
                                    .LineStyle = xlLineStyle.xlContinuous
                                    .ColorIndex = xlAutomatic
                                    .TintAndShade = 0
                                    .Weight = xlBorderWeight.xlThin
                                End With

                                If m_lCols > 1 Then
                                    .Borders(xlAlignment.xlInsideVertical).LineStyle = xlUnderlineStyle.xlSolid
                                    .Borders(xlAlignment.xlInsideHorizontal).LineStyle = xlNone
                                End If
                            End With

                            ' 18. Damos formato al grupo de datos
                            For m_lRow = 3 To (m_lRows + 1) Step 2
                                With .Range(modMDB.GetCell(0) & m_lRow & ":" & modMDB.GetCell(m_lCols - 1) & m_lRow).Interior
                                    .Pattern = xlSolid
                                    .PatternColorIndex = xlAutomatic
                                    .Color = 15658734
                                    .TintAndShade = 0
                                    .PatternTintAndShade = 0
                                End With
                            Next m_lRow
                            
                            With .Range(modMDB.GetCell(0) & "2:" & modMDB.GetCell(m_lCols - 1) & (m_lRows + 1))
                                .Borders(xlAlignment.xlDiagonalDown).LineStyle = xlNone
                                .Borders(xlAlignment.xlDiagonalUp).LineStyle = xlNone

                                With .Borders(xlAlignment.xlEdgeLeft)
                                    .LineStyle = xlLineStyle.xlContinuous
                                    .ColorIndex = xlAutomatic
                                    .TintAndShade = 0
                                    .Weight = xlBorderWeight.xlThin
                                End With

                                With .Borders(xlAlignment.xlEdgeTop)
                                    .LineStyle = xlLineStyle.xlContinuous
                                    .ColorIndex = xlAutomatic
                                    .TintAndShade = 0
                                    .Weight = xlBorderWeight.xlThin
                                End With

                                With .Borders(xlAlignment.xlEdgeBottom)
                                    .LineStyle = xlLineStyle.xlContinuous
                                    .ColorIndex = xlAutomatic
                                    .TintAndShade = 0
                                    .Weight = xlBorderWeight.xlThin
                                End With

                                With .Borders(xlAlignment.xlEdgeRight)
                                    .LineStyle = xlLineStyle.xlContinuous
                                    .ColorIndex = xlAutomatic
                                    .TintAndShade = 0
                                    .Weight = xlBorderWeight.xlThin
                                End With

                                If (m_lCols > 1) Then .Borders(xlAlignment.xlInsideVertical).LineStyle = xlUnderlineStyle.xlSolid
                                If (m_lRows > 1) Then .Borders(xlAlignment.xlInsideHorizontal).LineStyle = xlUnderlineStyle.xlSolid
                            End With

                            ' 19. Damos formato a la columna con los números de votos
                            .Range(modMDB.GetCell(m_lCols - 1) & "2:" & modMDB.GetCell(m_lCols - 1) & (m_lRows + IIf(bShowTotals, 2, 1))).NumberFormat = "#,##0"

                            If bShowTotals Then
                                ' 20. Damos formato a la celda de total de votos
                                With .Range(modMDB.GetCell(m_lCols - 1) & (m_lRows + 2))
                                    .FormulaR1C1 = "=SUM(R[-" & m_lRows & "]C:R[-1]C)"
                                    .Font.Bold = True

                                    With .Interior
                                        .Pattern = xlSolid
                                        .PatternColorIndex = xlAutomatic
                                        .Color = 13434879
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With

                                    .Borders(xlAlignment.xlDiagonalDown).LineStyle = xlNone
                                    .Borders(xlAlignment.xlDiagonalUp).LineStyle = xlNone

                                    With .Borders(xlAlignment.xlEdgeLeft)
                                        .LineStyle = xlLineStyle.xlContinuous
                                        .ColorIndex = xlAutomatic
                                        .TintAndShade = 0
                                        .Weight = xlBorderWeight.xlThin
                                    End With

                                    With .Borders(xlAlignment.xlEdgeTop)
                                        .LineStyle = xlLineStyle.xlContinuous
                                        .ColorIndex = xlAutomatic
                                        .TintAndShade = 0
                                        .Weight = xlBorderWeight.xlThin
                                    End With

                                    With .Borders(xlAlignment.xlEdgeBottom)
                                        .LineStyle = xlLineStyle.xlContinuous
                                        .ColorIndex = xlAutomatic
                                        .TintAndShade = 0
                                        .Weight = xlBorderWeight.xlThin
                                    End With
    
                                    With .Borders(xlAlignment.xlEdgeRight)
                                        .LineStyle = xlLineStyle.xlContinuous
                                        .ColorIndex = xlAutomatic
                                        .TintAndShade = 0
                                        .Weight = xlBorderWeight.xlThin
                                    End With

                                    .Borders(xlAlignment.xlInsideVertical).LineStyle = xlNone
                                    .Borders(xlAlignment.xlInsideHorizontal).LineStyle = xlNone
                                End With

                                ' 21. Damos formato a la línea de total
                                With .Range(modMDB.GetCell(0) & (m_lRows + 2) & ":" & modMDB.GetCell(m_lCols - 2) & (m_lRows + 2))
                                    .HorizontalAlignment = xlAlignment.xlRight
                                    .VerticalAlignment = xlAlignment.xlBottom
                                    .WrapText = False
                                    .Orientation = 0
                                    .AddIndent = False
                                    .IndentLevel = 0
                                    .ShrinkToFit = False
                                    .ReadingOrder = xlContext
                                    .MergeCells = True

                                    With .Interior
                                        .Pattern = xlSolid
                                        .PatternColorIndex = xlAutomatic
                                        .ThemeColor = xlThemeColor.xlThemeColorAccent1
                                        .TintAndShade = 0.799981688894314
                                        .PatternTintAndShade = 0
                                    End With

                                    .FormulaR1C1 = "Total de votos a candidaturas"

                                    With .Font
                                        .Name = "Calibri"
                                        .FontStyle = "Negrita"
                                        .Size = 11
                                        .Strikethrough = False
                                        .Superscript = False
                                        .Subscript = False
                                        .OutlineFont = False
                                        .Shadow = False
                                        .Underline = xlUnderlineStyleNone
                                        .ThemeColor = xlThemeColor.xlThemeColorLight1
                                        .TintAndShade = 0
                                        .ThemeFont = xlThemeFont.xlThemeFontMinor
                                    End With

                                    .Borders(xlAlignment.xlDiagonalDown).LineStyle = xlNone
                                    .Borders(xlAlignment.xlDiagonalUp).LineStyle = xlNone

                                    With .Borders(xlAlignment.xlEdgeLeft)
                                        .LineStyle = xlLineStyle.xlContinuous
                                        .ColorIndex = xlAutomatic
                                        .TintAndShade = 0
                                        .Weight = xlBorderWeight.xlThin
                                    End With

                                    With .Borders(xlAlignment.xlEdgeTop)
                                        .LineStyle = xlLineStyle.xlContinuous
                                        .ColorIndex = xlAutomatic
                                        .TintAndShade = 0
                                        .Weight = xlBorderWeight.xlThin
                                    End With

                                    With .Borders(xlAlignment.xlEdgeBottom)
                                        .LineStyle = xlLineStyle.xlContinuous
                                        .ColorIndex = xlAutomatic
                                        .TintAndShade = 0
                                        .Weight = xlBorderWeight.xlThin
                                    End With

                                    With .Borders(xlAlignment.xlEdgeRight)
                                        .LineStyle = xlLineStyle.xlContinuous
                                        .ColorIndex = xlAutomatic
                                        .TintAndShade = 0
                                        .Weight = xlBorderWeight.xlThin
                                    End With

                                    .Borders(xlAlignment.xlInsideVertical).LineStyle = xlNone
                                    .Borders(xlAlignment.xlInsideHorizontal).LineStyle = xlNone
                                End With
                            Else
                                With .Range("A1:AZ1")
                                    .HorizontalAlignment = xlAlignment.xlCenter
                                    .VerticalAlignment = xlAlignment.xlBottom
                                    .WrapText = True
                                    .Orientation = 0
                                    .AddIndent = False
                                    .IndentLevel = 0
                                    .ShrinkToFit = False
                                    .ReadingOrder = xlContext
                                    .MergeCells = False
                                End With
                                .Columns("T:AI").ColumnWidth = 25
                                .Rows("1:1").RowHeight = 48
                            End If

                            ' 22. Seleccionamos la primera celda, de la primera fila
                            .Range("A1").Select
                        End With

                        ' 23. Establecemos las propiedades del documento
                        .BuiltinDocumentProperties("Title") = IIf(Not IsZrStr(sTitle), sTitle, modSystem.GetFileName(sXLSFilePath))
                        .BuiltinDocumentProperties("Subject") = IIf(Not IsZrStr(sSubject), sSubject, "Consulta en formato MS Excel")
                        .BuiltinDocumentProperties("Author") = modSystem.GetCurrentWinUser()
                        .BuiltinDocumentProperties("Company") = "Plataforma Elecciones Transparentes"
                        .BuiltinDocumentProperties("Category") = IIf(Not IsZrStr(sCategory), sCategory, "Consulta en formato MS Excel")
                        .BuiltinDocumentProperties("Comments") = IIf(Not IsZrStr(sComments), sComments, IIf(Not IsZrStr(sTitle), sTitle, modSystem.GetFileName(sXLSFilePath)))

                        ' 24. Si se han indicado propiedades del documento
                        If Not colDocProperties Is Nothing Then
                            If Not IsZr(colDocProperties.Count) Then
                                Dim m_lIdx  As Long

                                ' 25. Las agregamos al documento
                                For m_lIdx = 1 To colDocProperties.Count
                                    .CustomDocumentProperties.Add colDocProperties.Item(m_lIdx).NOMBRE, False, colDocProperties.Item(m_lIdx).Tipo, colDocProperties.Item(m_lIdx).Valor
                                Next m_lIdx
                            End If
                        End If

                        ' 26. Guardamos el fichero
                        .Save: .Close SaveChanges:=False
                    End With
                End With
                Set m_Excel = Nothing
            End If

            ' 27. Si se indicó, abrimos el fichero resultante
            If bShow Then modSystem.OpenMSExcel sXLSFilePath, vbNormalFocus

            .Close
        End With

        ' 28. Se elimina la consulta
        .QueryDefs.Delete m_sWkst

        ' 29. Indicamos el éxito del proceso
        ToMSExcel = enumMSExcelFileStatus.xlSuccess
    End With

Error_ToMSExcel:
    ' 1. Si no está cerrado el objeto Excel.Application
    If Not m_Excel Is Nothing Then
        Dim m_iWkb  As Integer

        ' 2. Recorremos la colección de libros desde el último al primero para cerrarlos
        For m_iWkb = m_Excel.Workbooks.Count To 1 Step -1
            m_Excel.Workbooks(m_iWkb).Close SaveChanges:=False
        Next m_iWkb
        Set m_Excel = Nothing
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para obtener por el tipo de versión de MS Excel los archivos MS Excel compatibles
'-----------------------------------------------------------------------------------------------------------------------
Public Function GetMSExcelDefaultExtension(Optional ByVal bDefault As Boolean = False) As String
    Select Case modMDB.GetMSOfficeVersion()
        Case 11
            GetMSExcelDefaultExtension = IIf(bDefault, "*", vbNullString) & ".xls"

        Case Is > 11
            GetMSExcelDefaultExtension = IIf(bDefault, "*", vbNullString) & ".xlsx"
    End Select
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para obtener por el tipo de versión de MS Excel los archivos MS Excel compatibles
'-----------------------------------------------------------------------------------------------------------------------
Public Function GetMSExcelDefaultWorkbookFormat(Optional ByVal bDefault As Boolean = False) As Long
    Select Case modMDB.GetMSOfficeVersion()
        Case 11
            GetMSExcelDefaultWorkbookFormat = acSpreadsheetTypeExcel9

        Case Is > 11
            GetMSExcelDefaultWorkbookFormat = acSpreadsheetTypeExcel12Xml
    End Select
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para obtener el número de versión de Microsoft Office
'-----------------------------------------------------------------------------------------------------------------------
Public Function GetMSOfficeVersion() As Integer
    GetMSOfficeVersion = CInt(Split(Trim(Application.Version), ".")(0))
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para instanciar un objeto MS Excel Application dependiendo de la versión de MS Office.
'-----------------------------------------------------------------------------------------------------------------------
Public Function GetMSExcelApplication() As Object
    Set GetMSExcelApplication = CreateObject("Excel.Application." & modMDB.GetMSOfficeVersion())
End Function
