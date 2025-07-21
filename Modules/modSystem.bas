' ---------------------------------------------------------------
' Módulo: modSystem.bas
' Autor: Juan Francisco Cucharero Cabezas
' Proyecto: InfoElectoral
' Descripción:
'   Funciones auxiliares para interacción con el sistema operativo.
'   Incluye descompresión de archivos, gestión de rutas, apertura
'   de carpetas y eliminación de archivos temporales.
'
'   ⚠️ Este módulo no contiene validaciones ni confirmaciones.
'   Toda la lógica interactiva se gestiona desde los formularios.
'
' Fecha: [añadir fecha]
' ---------------------------------------------------------------

Attribute VB_Name = "modSystem"
Option Compare Database
Option Explicit

Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As typOPENFILENAME) As Long
Private Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As typOPENFILENAME) As Long

Private Declare PtrSafe Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare PtrSafe Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare PtrSafe Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As typSYSTEMTIME)

Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" (lpbi As typBrowseInfo) As Long
Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As LongPtr, ByVal lpBuffer As String) As Long
Private Declare PtrSafe Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare PtrSafe Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare PtrSafe Function SndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32
Public Const DM_IN_BUFFER = 8
Public Const DM_OUT_BUFFER = 2
Public Const DM_FORMNAME = &H10000
Public Const DM_COPIES = &H100&
Public Const DM_MODIFY = 8
Public Const DM_COPY = 2

Private Const MAX_PATH                  As Long = 260

Public Enum CSIDL_Enum
    USER_DESKTOP = 0                ' User 'Desktop' Folder (C:\Documents and Settings\[USER]\Desktop)
    USER_PROGRAMS = 2               ' User 'Program Groups' Folder (C:\Documents and Settings\[USER]\Start Menu\Programs)
    USER_PERSONAL = 5               ' User 'My Documents' Folder (C:\Documents and Settings\[USER]\My Documents)
    USER_FAVORITES = 6              ' User 'Favorites' Folder (C:\Documents and Settings\[USER]\Favorites)
    USER_STARTUP = 7                ' User 'Startup Group' Folder (C:\Documents and Settings\[USER]\Start Menu\Programs\Startup)
    USER_RECENT = 8                 ' User 'Recently Used Documents' Folder (C:\Documents and Settings\[USER]\Recent)
    USER_SENDTO = 9                 ' User 'Send To' Folder (C:\Documents and Settings\[USER]\SendTo)
    USER_STARTMENU = 11             ' User 'Start Menu' Folder (C:\Documents and Settings\[USER]\Start Menu)
    USER_MYMUSIC = 13               ' User 'My music' Folder (C:\Documents and Settings\[USER]\My Documents\My Music)
    USER_MYVIDEOS = 14              ' User 'My videos' Folder (C:\Documents and Settings\[USER]\My Documents\My Videos)
    USER_DESKTOPDIRECTORY = 16      ' User 'Desktop' Folder (C:\Documents and Settings\[USER]\Desktop)
    USER_NETHOOD = 19               ' User 'Network Neighborhood' Folder (C:\Documents and Settings\[USER]\NetHood)
    WINDOWS_FONTS = 20              ' Windows 'Fonts' Folder (C:\WINNT\Fonts)
    USER_TEMPLATES = 21             ' User 'Document Templates' Folder (C:\Documents and Settings\[USER]\Templates)
    ALLUSERS_STARTMENU = 22         ' All Users 'Start Menu' Folder (C:\Documents and Settings\All Users\Menú Inicio)
    ALLUSERS_PROGRAMS = 23          ' All Users 'Program Group' Folder (C:\Documents and Settings\All Users\Menú Inicio\programas)
    ALLUSERS_STARTUP = 24           ' All Users 'Startup Group' Folder (C:\Documents and Settings\All Users\Menú Inicio\programas\Inicio)
    ALLUSERS_DESKTOPDIRECTORY = 25  ' All Users 'Desktop' Folder (C:\Documents and Settings\All Users\Desktop)
    USER_APPDATA = 26               ' User 'Application Data' Folder (C:\Documents and Settings\[USER]\Application Data)
    USER_PRINTHOOD = 27             ' User 'Printers' Folder (C:\Documents and Settings\[USER]\PrintHood)
    USER_LOCALSETTINGSAPPDATA = 28  ' User 'Local Settings Application Data' Folder (C:\Documents and Settings\[USER]\Local Settings\Application Data)
    ALLUSERS_FAVORITES = 31         ' All Users 'Favorites' Folder (C:\Documents and Settings\All Users\Favorites)
    USER_INTERNET_CACHE = 32        ' User 'Temporary Internet Files' Folder (C:\Documents and Settings\[USER]\Local Settings\Temporary Internet Files)
    USER_COOKIES = 33               ' User 'Cookies' Folder (C:\Documents and Settings\[USER]\Cookies)
    USER_HISTORY = 34               ' User 'History' Folder (C:\Documents and Settings\[USER]\Local Settings\History)
    ALLUSERS_APPDATA = 35           ' All Users 'Application Data' Folder (C:\Documents and Settings\All Users\Datos de programa)
    WINDOWS_DIRECTORY = 36          ' Windows Folder (C:\WINNT)
    WINDOWS_SYSTEM32 = 37           ' Windows 'System32' Folder (C:\WINNT\system32)
    WINDOWS_PROGRAMFILES = 38       ' Windows 'Program Files' Folder (C:\Program Files)
    USER_MYPICTURES = 39            ' User 'My pictures' Folder (C:\Documents and Settings\[USER]\My Documents\My Pictures)
    USER_PATH = 40                  ' User Personal Path (C:\Documents and Settings\[USER])
    WINDOWS_SYSTEMDIRECTORY = 41    ' Windows 'System32' Folder(C:\WINNT\system32)
    WINDOWS_COMMONFILES = 43        ' Windows 'Common Files' Folder(C:\Program Files\Common Files)
    ALLUSERS_TEMPLATES = 45         ' All Users 'Templates' Folder (C:\Documents and Settings\All Users\Plantillas)
    ALLUSERS_ADMINTOOLS = 47        ' All Users 'Administrative Tools' Folder (C:\Documents and Settings\All Users\Start Menu\Programs\Administrative Tools)
    USER_ADMINTOOLS = 48            ' User 'Administrative Tools' Folder (C:\Documents and Settings\[USER]\Start Menu\Programs\Administrative Tools)
    ALLUSERS_MYMUSIC = 53           ' All Users ' My music' Folder (C:\Documents and Settings\All Users\Documents\My Music)
    ALLUSERS_MYPICTURES = 54        ' All Users 'My pictures' Folder (C:\Documents and Settings\All Users\Documents\My Pictures)
    ALLUSERS_MYVIDEOS = 55          ' All Users 'My videos' Folder (C:\Documents and Settings\All Users\Documents\My Videos)
    WINDOWS_RESOURCES = 56          ' Windows 'Resources' Folder (C:\WINNT\Resources)
    USER_RESOURCES = 57             ' User 'Resources' Folder (C:\WINNT\Resources\0c0a)
    USER_CDBURNING = 59             ' User 'CD Burning' Folder (C:\Documents and Settings\[USER]\Local Settings\Application Data\Microsoft\CD Burning)
End Enum

Public Enum SND_Flags
    SND_APPLICATION = &H80          ' Nombre en la entrada de WIN.INI [sounds]
    SND_ALIAS = &H10000             ' Nombre de entrada identificada a WIN.INI es [sounds]
    SND_ALIAS_ID = &H110000
    SND_ASYNC = &H1                 ' Reproducir Asincronadamente (por defecto)
    SND_SYNC = &H0                  ' Reproducir sincronadamente ¡No recomendado, pues el mci no te devuelve el control hasta
    SND_FILENAME = &H20000
    SND_LOOP = &H8                  ' Reproducir en bucle continuo
    SND_MEMORY = &H4
    SND_NODEFAULT = &H2
    SND_NOSTOP = &H10
    SND_NOWAIT = &H2000
    SND_PURGE = &H40
    SND_RESOURCE = &H40004
End Enum

' Declaración de constantes
Public Enum OFN_Enum
    OFN_ALLOWMULTISELECT = &H200
    OFN_CREATEPROMPT = &H2000
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_EXPLORER = &H80000
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_FILEMUSTEXIST = &H1000
    OFN_HIDEREADONLY = &H4
    OFN_LONGNAMES = &H200000
    OFN_NOCHANGEDIR = &H8
    OFN_NODEREFERENCELINKS = &H100000
    OFN_NOLONGNAMES = &H40000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NOVALIDATE = &H100
    OFN_OVERWRITEPROMPT = &H2
    OFN_PATHMUSTEXIST = &H800
    OFN_READONLY = &H1
    OFN_SHAREAWARE = &H4000
    OFN_SHOWHELP = &H10
End Enum


Public Enum enumTypeDateTreePath
    None = 0
    DayTreePath = 1
    MonthTreePath = 2
    YearTreePath = 3
End Enum

Public Enum enumMsoFileDialogType
    msoFileDialogOpen = 1
    msoFileDialogSaveAs = 2
    msoFileDialogFilePicker = 3
    msoFileDialogFolderPicker = 4
End Enum

Public Enum enumStripStr
    None = 0                        ' Sin limpiar
    BOE = 1                         ' Codificación del formato BOE (ISO-8859-1)
    ISO = 2                         ' Codificación del formato SEPA sin caracteres alfanuméricos
    SEPA = 3                        ' Codificación del formato SEPA (ISO-20022)
    SQL = 4                         ' Codificación compatible con cadenas de texto en lenguaje SQL
End Enum

' Constante de tipo de codificación de fichero
Public Enum enumCharcode
    CdoSystemASCII = 0
    CdoBIG5 = 1
    CdoEUC_JP = 2
    CdoEUC_KR = 3
    CdoGB2312 = 4
    CdoISO_2022_JP = 5
    CdoISO_2022_KR = 6
    CdoISO_8859_1 = 7
    CdoISO_8859_2 = 8
    CdoISO_8859_3 = 9
    CdoISO_8859_4 = 10
    CdoISO_8859_5 = 11
    CdoISO_8859_6 = 12
    CdoISO_8859_7 = 13
    CdoISO_8859_8 = 14
    CdoISO_8859_9 = 15
    CdoKOI8_R = 16
    CdoShift_JIS = 17
    CdoUS_ASCII = 18
    CdoUTF_7 = 19
    CdoUTF_8 = 20
End Enum

Public Enum enumSaveCreate
    adSaveCreateNotExist = 1    ' Default. Creates a new file if the file does not already exist
    adSaveCreateOverWrite = 2   ' Overwrites the file with the data from the currently open Stream object, if the file already exists
End Enum

' Tipos definidos
Public Type typSYSTEMTIME
    wYear                               As Integer
    wMonth                              As Integer
    wDayOfWeek                          As Integer
    wDay                                As Integer
    wHour                               As Integer
    wMinute                             As Integer
    wSecond                             As Integer
    wMilliseconds                       As Integer
End Type

Public Type typOPENFILENAME
    lStructSize                         As Long             ' Tamaño en bytes de esta misma estructura.
    hwndOwner                           As Long             ' Manejador de ventana del formulario padre.
    hInstance                           As Long             ' Manejador de instancia.
    lpstrFilter                         As String           ' Filtro de archivos para abrir.
    lpstrCustomFilter                   As String           ' Filtro personalizado.
    nMaxCustFilter                      As Long             ' Indice del filtro personalizado.
    nFilterIndex                        As Long             ' Indíce del filtro.
    lpstrFile                           As String           ' Nombre del fichero inicial.
    nMaxFile                            As Long             ' Longitud del nombre del Fichero.
    lpstrFileTitle                      As String           ' Nombre y extensión del fichero.
    nMaxFileTitle                       As Long             ' Título de la cadena precedente.
    lpstrInitialDir                     As String           ' Ruta de la carpeta inicial.
    lpstrTitle                          As String           ' Título de la ventana.
    Flags                               As Long             ' Flags de configuración de la ventana.
    nFileOffset                         As Integer          ' Posición del nombre de fichero en la cadena.
    nFileExtension                      As Integer          ' Posición de la extensión del fichero en la cadena.
    lpstrDefExt                         As String           ' Extensión por defecto del fichero.
    lCustData                           As Long
    lpfnHook                            As Long
    lpTemplateName                      As String
End Type

Private Type typBrowseInfo
   hwndOwner                            As Long
   pIDLRoot                             As Long
   pszDisplayName                       As Long
   lpszTitle                            As Long
   ulFlags                              As Long
   lpfnCallback                         As Long
   lParam                               As Long
   iImage                               As Long
End Type

Public Type typDrive
    Drive                               As String
    VolumeName                          As String
    SerialNumber                        As Long
    SerialNumberHex                     As String
    MaximumComponentLength              As Long
    FileSystemFlags                     As Long
End Type

Public Type ShortItemId
    cb                                  As LongPtr
    abID                                As Byte
End Type

Public Type ITEMIDLIST
    mkid                                As ShortItemId
End Type

' Función para validar y construir una ruta de archivo o carpeta.
Public Function BuiltPath(ByVal sPath As String) As Boolean
    Dim m_sPath()           As String
    Dim m_sTempPath         As String
    Dim m_lInit             As Long
    Dim m_lIdx              As Long

    On Error GoTo Error_BuiltPath

    If modSystem.EvalStr(sPath, "^[\\]{2}") Then                                ' Unidad de red
        m_lInit = 2
    ElseIf modSystem.EvalStr(sPath, "^[A-Z]{1}[:]{1}[\\]{1}", , True) Then      ' Unidad física o unidad de red mapeada
        m_lInit = 1
    Else
        Err.Raise vbObjectError, GetMDBName, "No se reconoce el tipo de ruta."
    End If

    m_sPath = Split(sPath, "\")
    For m_lIdx = 0 To UBound(m_sPath)
        Concat m_sTempPath, IIf(Not IsZrStr(m_sTempPath), "\", vbNullString) & m_sPath(m_lIdx)
        If (m_lIdx >= m_lInit) Then
            If Not modSystem.FolderExists(m_sTempPath) Then modSystem.CreateFolder m_sTempPath
        End If
    Next m_lIdx

    BuiltPath = modSystem.FolderExists(sPath)

    Exit Function

Error_BuiltPath:
    GetError "modSystem.BuiltPath"
End Function

' Función para cambiar la extensión de la ruta de un fichero
Public Function ChangeFileExtension(ByVal sFilePath As String, _
                                    ByVal sExtension As String) As String

    On Error GoTo Error_ChangeFileExtension

    If Not (IsZrStr(sFilePath) Or IsZrStr(sExtension)) Then
        ChangeFileExtension = modSystem.GetFolderPath(sFilePath)
        Concat ChangeFileExtension, modSystem.GetFileName(sFilePath)
        Concat ChangeFileExtension, IIf(modSystem.EvalStr(Left(sExtension, 1), "^[.]$", , True), sExtension, "." & sExtension)
    End If

    Exit Function

Error_ChangeFileExtension:
    GetError "modSystem.ChangeFileExtension"
End Function

' Función para copiar un fichero.
Public Function CopyFile(ByVal sFromFilePath As String, _
                         ByVal sToFilePath As String, _
                         Optional ByVal bOverwrite As Boolean = False) As Boolean
    On Error GoTo Error_CopyFile

    If modSystem.FileExists(sToFilePath) Then
        If bOverwrite Then
            If Not modSystem.DeleteFile(sToFilePath) Then Err.Raise vbObjectError, GetMDBName, "No se ha podido reemplazar el fichero de destino. ¿Puede que esté en uso?"
        Else
            Err.Raise vbObjectError, GetMDBName, "No se puede copiar el archivo ya que se encuentra otro con el mismo nombre en la misma ubicación."
        End If
    End If

    VBA.FileSystem.FileCopy sFromFilePath, sToFilePath

    CopyFile = modSystem.FileExists(sToFilePath)

    Exit Function

Error_CopyFile:
    GetError "modSystem.CopyFile"
End Function

' Función para crear un fichero (Archivo).
Public Function CreateFile(ByVal sFilePath As String) As Boolean
    On Error GoTo Error_CreateFile

    Open sFilePath For Append As #1
    Close #1
    CreateFile = modSystem.FileExists(sFilePath)

    Exit Function

Error_CreateFile:
    Close #1
    GetError "modSystem.CreateFile"
End Function

' Método para crear una carpeta (Directorio).
Public Function CreateFolder(ByVal sFolderPath As String) As Boolean
    On Error GoTo Error_CreateFolder

    VBA.FileSystem.MkDir sFolderPath
    CreateFolder = modSystem.FolderExists(sFolderPath)

    Exit Function

Error_CreateFolder:
    GetError "modSystem.CreateFolder"
End Function

' Función para crear un acceso directo o fichero "lnk"
Public Function CreateShorcut(ByVal sShorcutFilePath As String, _
                              ByVal sTargetPath As String, _
                              Optional ByVal sWorkingDirectory As String = vbNullString, _
                              Optional ByVal sIconFilePath As String = vbNullString) As Boolean
    On Error GoTo Error_CreateShorcut

    ' 1. Si existe ya el Acceso directo, lo eliminamos
    modSystem.DeleteIfFileExists sShorcutFilePath

    With CreateObject("WScript.Shell").CreateShortcut(sShorcutFilePath)
        ' 2. Establecemos la ruta de ubicaci´no del Acceso directo
        .TargetPath = sTargetPath

        ' 3. Si no está vacio el Directorio de trabajo, y existe, se establece el pasado como parámetro
        If Not (IsZrStr(sWorkingDirectory) Or Not modSystem.FolderExists(sWorkingDirectory)) Then .WorkingDirectory = sWorkingDirectory

        ' 4. Si se pasó una ruta de icono, y existe la ruta del mismo, se procede a establecerlo para el Acceso directo
        If Not (IsZrStr(sIconFilePath) Or modSystem.FileExists(sIconFilePath)) Then .IconLocation = sIconFilePath

        ' 5. Guardamos el Acceso directo
        .Save
    End With

    ' 6. Comprobamos si existe el Acceso directo
    CreateShorcut = modSystem.FileExists(sShorcutFilePath)

    Exit Function

Error_CreateShorcut:
    GetError "modSystem.CreateShorcut"
End Function

' Función para eliminar un fichero.
Public Function DeleteFile(ByVal sFilePath As String) As Boolean
    On Error Resume Next
    VBA.FileSystem.Kill sFilePath
    DeleteFile = (Not modSystem.FileExists(sFilePath))
End Function

' Función que elimina un fichero si existe
Public Function DeleteIfFileExists(ByVal sFilePath As String) As Boolean
    On Error GoTo Error_DeleteIfFileExists

    If Not IsZrStr(sFilePath) Then
        If Not IsZrStr(VBA.FileSystem.Dir(sFilePath, vbArchive)) Then
            VBA.FileSystem.Kill sFilePath
            DeleteIfFileExists = (Not IsZrStr(VBA.FileSystem.Dir(sFilePath, vbArchive)))
        End If
    End If

    Exit Function

Error_DeleteIfFileExists:
    GetError "modSystem.DeleteIfFileExists", sFilePath
End Function

' Función para validar si existe un fichero (Archivo).
Public Function FileExists(ByVal sFilePath As String) As Boolean
    On Error GoTo Error_FileExists

    If Not IsZrStr(sFilePath) Then FileExists = (Not IsZrStr(VBA.FileSystem.Dir(sFilePath, vbArchive)))

    Exit Function

Error_FileExists:
    GetError "modSystem.FileExists", sFilePath
End Function

' Función para validar si existe una carpeta (Directorio).
Public Function FolderExists(ByVal sFolderPath As String) As Boolean
    On Error GoTo Error_FolderExists

    FolderExists = (Not IsZrStr(VBA.FileSystem.Dir(sFolderPath, vbDirectory)))

    Exit Function

Error_FolderExists:
    Err.Clear
End Function

' Función para obtener el Usuario de la sesión actual de Windows.
Public Function GetCurrentWinUser() As String
    Dim m_sBuffer           As String

    On Error GoTo Error_GetCurrentWinUser

    m_sBuffer = Space(MAX_PATH)
    If Not IsZr(GetUserName(m_sBuffer, 255)) Then GetCurrentWinUser = modSystem.StripNulls(m_sBuffer)

    Exit Function

Error_GetCurrentWinUser:
    GetError "modSystem.GetCurrentWinUser"
End Function

' Función para obtener el nombre del equipo donde se inició la sesión actual de Windows.
Public Function GetCurrentWorkStation() As String
    Dim m_sBuffer           As String

    On Error GoTo Error_GetCurrentWorkStation

    m_sBuffer = Space(MAX_PATH)
    If Not IsZr(GetComputerName(m_sBuffer, 255)) Then GetCurrentWorkStation = modSystem.StripNulls(m_sBuffer)

    Exit Function

Error_GetCurrentWorkStation:
    GetError "modSystem.GetCurrentWorkStation"
End Function

' Función para obtener una ruta de directorio a partir de una fecha
Public Function GetDatePathTree(Optional ByVal dtDate As Date = 0, _
                                Optional ByVal lenumTypeDateTreePath As enumTypeDateTreePath = enumTypeDateTreePath.MonthTreePath) As String
    On Error GoTo Error_GetDatePathTree

    If IsZrDt(dtDate) Then dtDate = Date

    Select Case lenumTypeDateTreePath
        Case enumTypeDateTreePath.DayTreePath
            GetDatePathTree = (Format(dtDate, "yyyy\\mm\ \-\ ") & modSystem.GetLetterCase(Format(dtDate, "mmmm\\dd\\")))

        Case enumTypeDateTreePath.MonthTreePath
            GetDatePathTree = (Format(dtDate, "yyyy\\mm\ \-\ ") & modSystem.GetLetterCase(Format(dtDate, "mmmm\\")))

        Case enumTypeDateTreePath.YearTreePath
            GetDatePathTree = Format(dtDate, "yyyy\\")
    End Select

    Exit Function

Error_GetDatePathTree:
    GetError "modSystem.GetDatePathTree"
End Function

' Función para obtener la informacion de volumen de una unidad.
Public Function GetDriveInfo(ByVal sRootPathName As String) As typDrive
    Dim m_sVolumeNameBuffer         As String * 256
    Dim m_sFileSystemNameBuffer     As String * 256

    On Error GoTo Error_GetDriveInfo

    ' 1. Ponemos la ruta en mayúsculas
    sRootPathName = UCase(sRootPathName)

    ' 2. Validamos que sea correcta la ruta
    If Not modSystem.EvalStr(sRootPathName, "^[A-Z]{1}[:]{1}[\\]{1}$", , True) Then Err.Raise vbObjectError, GetMDBName(), "No ha especificado una ruta de raiz de unidad."

    With GetDriveInfo
        ' 3. Obtenemos la información de la unidad
        GetVolumeInformation sRootPathName, m_sVolumeNameBuffer, 256, .SerialNumber, .MaximumComponentLength, .FileSystemFlags, m_sFileSystemNameBuffer, 256

        ' 4. Obtenemos la Letra de la unidad
        .Drive = Left(sRootPathName, 2)

        ' 5. Obtenemos el Tamaño de buffer del volumen
        If Not IsZrStr(m_sVolumeNameBuffer) Then .VolumeName = modSystem.StripNulls(m_sVolumeNameBuffer)

        ' 6. Obtenemos el Número de serie en formato Hexadecimal
        If Not IsZr(.SerialNumber) Then .SerialNumberHex = Format(Hex(.SerialNumber), "@@@@-@@@@")
    End With

    Exit Function

Error_GetDriveInfo:
    GetError "modSystem.GetDriveInfo"
End Function

' Función para obtener la fecha en la que se creó un fichero.
Public Function GetFileDateCreated(ByVal sFilePath As String, _
                                   Optional ByVal bDateTime As Boolean = False) As Date
    On Error GoTo Error_GetFileDateCreated

    With CreateObject("Scripting.FileSystemObject").GetFile(sFilePath)
        GetFileDateCreated = IIf(bDateTime, .DateCreated, Format(.DateCreated, "dd/mm/yyyy"))
    End With

    Exit Function

Error_GetFileDateCreated:
    GetError "modSystem.GetFileDateCreated"
End Function

' Función para obtener la fecha de la última modificación de un fichero.
Public Function GetFileDateLastModified(ByVal sFilePath As String, _
                                        Optional ByVal bDateTime As Boolean = False) As Date
    On Error GoTo Error_GetFileDateLastModified

    With CreateObject("Scripting.FileSystemObject").GetFile(sFilePath)
        GetFileDateLastModified = IIf(bDateTime, .DateLastModified, Format(.DateLastModified, "dd/mm/yyyy"))
    End With

    Exit Function

Error_GetFileDateLastModified:
    GetError "modSystem.GetFileDateLastModified"
End Function

' Función para obtener la extensión de un fichero a partir de su ruta.
Public Function GetFileExt(ByVal sFilePath As String, _
                           Optional ByVal bGetDot As Boolean = False) As String
    Dim m_sFilePath()   As String
    Dim m_sFileName()   As String

    On Error GoTo Error_GetFileExt

    m_sFilePath = Split(sFilePath, "\")
    m_sFileName = Split(m_sFilePath(UBound(m_sFilePath)), ".")
    If (UBound(m_sFileName) > 0) Then GetFileExt = IIf(bGetDot, ".", vbNullString) & m_sFileName(UBound(m_sFileName))

    Exit Function

Error_GetFileExt:
    GetError "modSystem.GetFileExt"
End Function

' Función para obtener el nombre completo de un fichero (Archivo).
Public Function GetFileFullName(ByVal sFilePath As String, _
                                Optional ByVal sDefault As String = vbNullString) As String
    On Error GoTo Error_GetFileFullName

    If Not IsZrStr(sFilePath) Then
        Dim m_sFileName() As String
        m_sFileName = Split(sFilePath, "\")
        GetFileFullName = m_sFileName(UBound(m_sFileName))
    Else
        GetFileFullName = sDefault
    End If

    Exit Function

Error_GetFileFullName:
    GetError "modSystem.GetFileFullName"
End Function

' Función para obtener el nombre de un fichero (Archivo).
Public Function GetFileName(ByVal sFilePath As String) As String
    On Error GoTo Error_GetFileName

    If Not IsZrStr(sFilePath) Then GetFileName = Mid(modSystem.GetFileFullName(sFilePath), 1, (Len(modSystem.GetFileFullName(sFilePath)) - (Len(modSystem.GetFileExt(sFilePath)) + 1)))

    Exit Function

Error_GetFileName:
    GetError "modSystem.GetFileName"
End Function

' Función para mostrar una ventana de diálogo para obtener la ruta del fichero a abrir.
Public Function GetFileToOpen(ByRef result As Long, _
                              ByVal hwnd As Long, _
                              ByVal Filter As String, _
                              ByVal IdxFilter As Long, _
                              Optional ByVal Flags As Long = 0, _
                              Optional ByVal Title As String = vbNullString, _
                              Optional ByVal InitFileName As String = vbNullString, _
                              Optional ByVal InitDir As String = vbNullString, _
                              Optional ByVal DefExt As String = vbNullString) As String
    Dim m_typOPENFILENAME     As typOPENFILENAME

    On Error GoTo Error_GetFileToOpen

    ' 1. Si no se ha especificado una carpeta de destino por defecto,
    '    se toma la carpeta "Escritorio" del usuario del equipo
    If IsZrStr(InitDir) Then InitDir = modSystem.GetSpecialFolder(hwnd, CSIDL_Enum.USER_DESKTOP)

    With m_typOPENFILENAME
        .lStructSize = Len(m_typOPENFILENAME)
        .hwndOwner = hwnd
        .hInstance = Application.hWndAccessApp 'App.hInstance
        .lpstrFilter = Replace(Filter, "|", vbNullChar) & vbNullChar & vbNullChar
        .lpstrCustomFilter = vbNullString
        .nMaxCustFilter = 0
        .nFilterIndex = IdxFilter
        .lpstrFile = Left(InitFileName & String(1024, vbNullChar), 1024)
        .nMaxFile = Len(.lpstrFile) - 1
        .lpstrFileTitle = .lpstrFile
        .nMaxFileTitle = .nMaxFile
        .lpstrInitialDir = InitDir
        .lpstrTitle = Title
        .Flags = Flags
        .lpstrDefExt = DefExt
        .lCustData = 0
        .lpfnHook = 0
        .lpTemplateName = 0
    End With

    result = GetOpenFileName(m_typOPENFILENAME)

    If Not IsZr(result) Then
        IdxFilter = m_typOPENFILENAME.nFilterIndex
        GetFileToOpen = modSystem.StripNulls(m_typOPENFILENAME.lpstrFile)
    Else
        GetFileToOpen = vbNullString
    End If
    DoEvents

    Exit Function

Error_GetFileToOpen:
    GetError "modSystem.GetFileToOpen"
End Function

' Función para mostrar una ventana de diálogo para obtener la ruta del fichero a guardar.
Public Function GetFileToSave(ByRef result As Long, _
                              ByVal hwnd As Long, _
                              ByVal Filter As String, _
                              ByVal IdxFilter As Long, _
                              Optional ByVal Flags As Long = OFN_EXPLORER + OFN_LONGNAMES + OFN_PATHMUSTEXIST, _
                              Optional ByVal Title As String = vbNullString, _
                              Optional ByVal InitFileName As String = vbNullString, _
                              Optional ByVal InitDir As String = vbNullString, _
                              Optional ByVal DefExt As String = vbNullString) As String
    Dim m_typOPENFILENAME        As typOPENFILENAME

    On Error GoTo Error_GetFileToSave

    ' 1. Si no se ha especificado una carpeta de destino por defecto,
    '    se toma la carpeta "Escritorio" del usuario del equipo
    If IsZrStr(InitDir) Then InitDir = modSystem.GetSpecialFolder(hwnd, CSIDL_Enum.USER_DESKTOP)

    With m_typOPENFILENAME
        .lStructSize = Len(m_typOPENFILENAME)
        .hwndOwner = hwnd
        .hInstance = Application.hWndAccessApp 'App.hInstance
        .lpstrFilter = Replace(Filter, "|", vbNullChar) & vbNullChar & vbNullChar
        .lpstrCustomFilter = vbNullString
        .nMaxCustFilter = 0
        .nFilterIndex = IdxFilter
        .lpstrFile = Left(InitFileName & String(1024, vbNullChar), 1024)
        .nMaxFile = Len(.lpstrFile) - 1
        .lpstrFileTitle = .lpstrFile
        .nMaxFileTitle = .nMaxFile
        .lpstrInitialDir = InitDir
        .lpstrTitle = Title
        .Flags = Flags
        .lpstrDefExt = DefExt
        .lCustData = 0
        .lpfnHook = 0
        .lpTemplateName = 0
    End With

    result = GetSaveFileName(m_typOPENFILENAME)

    If Not IsZr(result) Then
        IdxFilter = m_typOPENFILENAME.nFilterIndex
        GetFileToSave = modSystem.StripNulls(m_typOPENFILENAME.lpstrFile)
    Else
        GetFileToSave = vbNullString
    End If
    DoEvents

    Exit Function

Error_GetFileToSave:
    GetError "modSystem.GetFileToSave"
End Function

' Función para obtener la ruta de ubicación de un fichero (Archivo).
Public Function GetFolderPath(ByVal sFilePath As String) As String
    On Error GoTo Error_GetFolderPath

    If Not IsZrStr(sFilePath) Then
        Dim m_sFileName() As String

        m_sFileName = Split(sFilePath, "\")
        GetFolderPath = Left(sFilePath, Len(sFilePath) - Len(m_sFileName(UBound(m_sFileName))))
    End If

    Exit Function

Error_GetFolderPath:
    GetError "modSystem.GetFolderPath"
End Function

Public Function GetCharCode(Optional ByVal lCharCode As enumCharcode = enumCharcode.CdoUS_ASCII) As String
    Select Case lCharCode
        Case enumCharcode.CdoBIG5
            GetCharCode = "big5"
        Case enumCharcode.CdoEUC_JP
            GetCharCode = "euc-jp"
        Case enumCharcode.CdoEUC_KR
            GetCharCode = "euc-kr"
        Case enumCharcode.CdoGB2312
            GetCharCode = "gb2312"
        Case enumCharcode.CdoISO_2022_JP
            GetCharCode = "iso-2022-jp"
        Case enumCharcode.CdoISO_2022_KR
            GetCharCode = "iso-2022-kr"
        Case enumCharcode.CdoISO_8859_1
            GetCharCode = "iso-8859-1"
        Case enumCharcode.CdoISO_8859_2
            GetCharCode = "iso-8859-2"
        Case enumCharcode.CdoISO_8859_3
            GetCharCode = "iso-8859-3"
        Case enumCharcode.CdoISO_8859_4
            GetCharCode = "iso-8859-4"
        Case enumCharcode.CdoISO_8859_5
            GetCharCode = "iso-8859-5"
        Case enumCharcode.CdoISO_8859_6
            GetCharCode = "iso-8859-6"
        Case enumCharcode.CdoISO_8859_7
            GetCharCode = "iso-8859-7"
        Case enumCharcode.CdoISO_8859_8
            GetCharCode = "iso-8859-8"
        Case enumCharcode.CdoISO_8859_9
            GetCharCode = "iso-8859-9"
        Case enumCharcode.CdoKOI8_R
            GetCharCode = "koi8-r"
        Case enumCharcode.CdoShift_JIS
            GetCharCode = "shift-jis"
        Case enumCharcode.CdoUS_ASCII
            GetCharCode = "us-ascii"
        Case enumCharcode.CdoUTF_7
            GetCharCode = "utf-7"
        Case enumCharcode.CdoUTF_8
            GetCharCode = "utf-8"
    End Select
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Método para mostrar un mensaje de error
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       sSource     String              Origen del error.
'
'       sException  String (Opcional)   Excepción causante del error.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Sub GetError(ByVal sSource As String, _
                    Optional ByVal sException As String = vbNullString)
    If Screen.MousePointer = 11 Then Screen.MousePointer = 0

    MsgBox "[" & Err.Number & "]" & vbTab & sSource & ": " & Err.Description & vbTab & Err.Source & IIf(IsZrStr(sException), vbNullString, vbCrLf & vbCrLf & sException), (vbOKOnly + vbExclamation), "Error"
End Sub

' Función para pasar a mayúsculas la primera letra de un texto
Public Function GetLetterCase(ByVal sValue As String) As String
    On Error GoTo Error_GetLetterCase

    GetLetterCase = (UCase(Left(sValue, 1)) & LCase(Mid(sValue, 2)))

    Exit Function

Error_GetLetterCase:
    GetError "modSystem.GetLetterCase"
End Function

' Función para obtener la ruta de una carpeta especial
Public Function GetSpecialFolder(Optional ByVal hwnd As Long = 0, _
                                 Optional ByVal lCSIDL As CSIDL_Enum = CSIDL_Enum.USER_PATH) As String
    Dim m_lResult       As Long
    Dim m_ITEMIDLIST    As ITEMIDLIST

    On Error GoTo Error_GetSpecialFolder

    m_lResult = SHGetSpecialFolderLocation(hwnd, lCSIDL, m_ITEMIDLIST)

    If IsZr(m_lResult) Then
        GetSpecialFolder = Space(MAX_PATH)

        m_lResult = SHGetPathFromIDList(ByVal m_ITEMIDLIST.mkid.cb, ByVal GetSpecialFolder)

        If m_lResult Then GetSpecialFolder = StripNulls(GetSpecialFolder) & "\"
    End If

    Exit Function

Error_GetSpecialFolder:
    GetError "modSystem.GetSpecialFolder"
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para buscar un literal dentro de una cadena de texto con formato.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       sValue      String              Cadena de texto donde buscar.
'
'       sFormat     String              Formato aplicado.
'
'       sSearch     String              Literal a buscar.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function GetSrhStr(ByVal sValue As String, _
                          ByVal sFormat As String, _
                          ByVal sSearch As String) As String
    Dim m_lPos      As Long

    On Error GoTo Error_GetSrhStr

    m_lPos = InStr(1, sFormat, sSearch, vbBinaryCompare)
    While (m_lPos > 0)
        Concat GetSrhStr, Mid(sValue, m_lPos, 1)
        m_lPos = InStr((m_lPos + 1), sFormat, sSearch, vbBinaryCompare)
    Wend

    Exit Function

Error_GetSrhStr:
    GetError "modSystem.GetSrhStr"
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para validar si una variable es nula o tiene valor zero.
'-----------------------------------------------------------------------------------------------------------------------
' Parámetros:
'
'       vValue      Variant     Variable a validar.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function IsZr(ByVal vValue As Variant) As Boolean
    IsZr = (Not IsNumeric(vValue) Or (vValue = 0) Or IsNull(vValue))
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para validar si una variable es una fecha valida.
'-----------------------------------------------------------------------------------------------------------------------
' Parámetros:
'
'       vValue      Variant             Variable a validar.
'
'       sFormat     String  (Opcional)  Formato a aplicar a la fecha. (Por defecto = "dd/mm/yyyy")
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function IsZrDt(ByVal vValue As Variant, _
                       Optional ByVal sFormat As String = "dd/mm/yyyy") As Boolean
    If ((IsNull(vValue) Or IsZrStr(vValue)) Or Not IsDate(vValue)) Then
        IsZrDt = True
    Else
        If (VarType(vValue) = vbString) Then vValue = DateSerial(CInt(modSystem.GetSrhStr(vValue, sFormat, "y")), CInt(modSystem.GetSrhStr(vValue, sFormat, "m")), CInt(modSystem.GetSrhStr(vValue, sFormat, "d")))
        IsZrDt = IsZr(CDbl(vValue))
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para validar si una variable no es cadena vacía o nula.
'-----------------------------------------------------------------------------------------------------------------------
' Parámetros:
'
'       vValue      Variant     Variable a validar.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function IsZrStr(ByVal vValue As Variant) As Boolean
    If IsNull(vValue) Then
        IsZrStr = True
    Else
        IsZrStr = (Trim(vValue) = vbNullString)
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para validar una cadena de texto contra una serie de expresiones regulares
'-----------------------------------------------------------------------------------------------------------------------
' Parámetros:
'
'       sValue      String              Cadena de texto a validar.
'
'       sRegExp     String              Expresión regular con la que validar. Pueden se más de una expresión regular,
'                                       concatenándolas con el símbolo "|" entre medias.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function EvalRegExp(ByVal sValue As String, _
                           ByVal sRegExp As String) As Boolean
    Dim m_sRegExp()     As String
    Dim m_lIdx          As Long

    On Error GoTo Error_EvalRegExp

    If Not IsZrStr(sRegExp) Then
        m_sRegExp = Split(sRegExp, "|")
        For m_lIdx = LBound(m_sRegExp) To UBound(m_sRegExp)
            If Not IsZrStr(m_sRegExp(m_lIdx)) Then
                If (Left(m_sRegExp(m_lIdx), 1) <> "!") Then
                    EvalRegExp = EvalRegExp Or modSystem.EvalStr(sValue, m_sRegExp(m_lIdx))
                Else
                    EvalRegExp = EvalRegExp And Not modSystem.EvalStr(sValue, Mid(m_sRegExp(m_lIdx), 2, (Len(m_sRegExp(m_lIdx)) - 1)))
                End If
            End If
        Next m_lIdx
    End If

    Exit Function

Error_EvalRegExp:
    GetError "modSystem.EvalRegExp", "sRegExp = " & sRegExp
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para validar una cadena de texto contra una expresión regular.
'-----------------------------------------------------------------------------------------------------------------------
' Parámetros:
'
'       sValue      String              Cadena de texto a validar.
'
'       sRegExp     String              Expresión regular con la que validar.
'
'       bGlobal     Boolean (Opcional)  Indicador de si la búsqueda es global.
'
'       bIgnoreCase Boolean (Opcional)  Indicador de si se ignoran mayúsculas y minúsculas.
'
'       bMultiLine  Boolean (Opcional)  Indicador de si el valor contiene varias lineas.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function EvalStr(ByVal sValue As String, _
                        ByVal sRegExp As String, _
                        Optional ByVal bGlobal As Boolean = True, _
                        Optional ByVal bIgnoreCase As Boolean = False, _
                        Optional ByVal bMultiLine As Boolean = False) As Boolean
    On Error GoTo Error_EvalStr

    With CreateObject("VBScript.RegExp")
        .Global = bGlobal
        .IgnoreCase = bIgnoreCase
        .MultiLine = bMultiLine
        .Pattern = sRegExp
        EvalStr = .Test(sValue)
    End With

    Exit Function

Error_EvalStr:
    GetError "modSystem.EvalStr", "sRegExp = " & sRegExp
End Function

' Función para reproducr un fichero de sonido WAV
Public Function PlaySound(Optional ByVal sSoundFilePath As String = vbNullString) As Boolean
    On Error GoTo Error_PlaySound

    If modSystem.FileExists(sSoundFilePath) Then PlaySound = Not IsZr(SndPlaySound(sSoundFilePath, SND_Flags.SND_ASYNC))

    Exit Function

Error_PlaySound:
     GetError "modSystem.PlaySound"
End Function

' Método para leer como texto el contenido de un fichero (Archivo).
Public Function ReadFileV2(ByVal sFilePath As String, _
                           Optional ByVal lCharCode As enumCharcode = enumCharcode.CdoUS_ASCII) As String
    On Error GoTo Error_ReadFileV2

    With CreateObject("ADODB.Stream")
        .Open
        .Charset = modSystem.GetCharCode(lCharCode)
        .LoadFromFile sFilePath
        ReadFileV2 = .ReadText()
        .Close
    End With

    Exit Function

Error_ReadFileV2:
    GetError "modSystem.ReadFileV2"
End Function

' Método para leer como texto el contenido de un fichero (Archivo). (Ya no se emplea)
'Public Function ReadFile(ByVal sFilePath As String) As String
'    Dim m_sLineText   As String
'
'    On Error GoTo Error_ReadFile
'
'    If modSystem.FileExists(sFilePath) Then
'        Open sFilePath For Input As #1
'        While Not EOF(1)
'            Line Input #1, m_sLineText
'            Concat ReadFile, m_sLineText & IIf(EOF(1), vbNullString, vbCrLf)
'        Wend
'        Close #1
'    End If
'
'    Exit Function
'
'Error_ReadFile:
'    Close #1
'    GetError "modSystem.ReadFile"
'End Function

'-----------------------------------------------------------------------------------------------------------------------
' Método para leer como texto el contenido de un fichero (Archivo).
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       sFilePath       String                      Ruta del fichero de leer.
'
'       lCharCode       enumCharcode (opcional)     Tipo de codificación del fichero a leer.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function ReadFile(ByVal sFilePath As String, _
                         Optional ByVal lCharCode As enumCharcode = enumCharcode.CdoUS_ASCII) As String
    On Error GoTo Error_ReadFile

    With CreateObject("ADODB.Stream")
        .Open
        .Charset = modSystem.GetCharCode(lCharCode)
        .LoadFromFile sFilePath
        ReadFile = .ReadText()
        .Close
    End With

    Exit Function

Error_ReadFile:
    GetError "modSystem.ReadFile", "sFilePath=" & sFilePath & vbCrLf & "CharCode=" & modSystem.GetCharCode(lCharCode)
End Function

' Función para renombrar un fichero.
Public Function RenameFile(ByVal sFromFilePath As String, _
                           ByVal sToFilePath As String, _
                           Optional ByVal bOverwrite As Boolean = False) As Boolean
    On Error GoTo Error_RenameFile

    ' Si existe copia de seguridad previa, se elimina
    If modSystem.FileExists(sToFilePath) Then
        If bOverwrite Then
            If Not modSystem.DeleteFile(sToFilePath) Then Err.Raise vbObjectError, GetMDBName, "No se ha podido sustituir el fichero de destino. ¿Puede que esté en uso?"
        Else
            Err.Raise vbObjectError, GetMDBName, "No se puede renombrar el archivo ya que se encuentra otro con el mismo nombre en la misma ubicación."
        End If
    End If

    Name sFromFilePath As sToFilePath

    RenameFile = modSystem.FileExists(sToFilePath)

    Exit Function

Error_RenameFile:
    GetError "modSystem.RenameFile"
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para reemplazar un literal dentro de una cadena de texto.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       sValue      String              Cadena de texto donde buscar.
'
'       sSearch     String              Literal a reemplazar.
'
'       sReplace    String (Opcional)   Literal por el que reemplazar.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function ReplStr(ByVal sValue As String, _
                        ByVal sSearch As String, _
                        Optional ByVal sReplace As String = vbNullString) As String
    On Error GoTo Error_ReplStr

    With CreateObject("VBScript.RegExp")
        .Pattern = sSearch
        .IgnoreCase = True
        .Global = True
        ReplStr = .Replace(sValue, sReplace)
    End With

    Exit Function

Error_ReplStr:
    GetError "modSystem.ReplStr"
End Function

' Función para consultar la ventana de la aplicación de Access
Public Function ShowAccessWindow(Optional ByVal iShowWindowMode As VBA.VbAppWinStyle = VBA.VbAppWinStyle.vbMaximizedFocus) As Long
    ShowAccessWindow = ShowWindow(Application.hWndAccessApp, iShowWindowMode)
End Function

' Función para eliminar cadenas nulas.
Public Function StripNulls(ByVal Str As String) As String
    On Error GoTo Error_StripNulls

    StripNulls = IIf(InStr(Str, Chr(0)) > 0, Left(Str, InStr(Str, vbNullChar) - 1), Str)

    Exit Function

Error_StripNulls:
    GetError "modSystem.StripNulls"
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para limpiar una cadena de texto de caracteres especiales, según el tipo de codificación empleado
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       sValue      String                      Cadena de texto a limpiar.
'
'       lStripStr   StripStr_Enum (Opcional)    Tipo de limpieza a aplicar. (Por defecto para SQL)
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function StripStr(ByVal sValue As String, _
                         Optional ByVal lStripStr As enumStripStr = enumStripStr.SQL) As String
    On Error GoTo Error_StripStr

    ' 1. Si no se trata de un texto para SQL
    If (lStripStr <> enumStripStr.SQL) Then
        ' 2. Limpiamos las vocales acentuadas o con diéresis
        sValue = Replace(sValue, "á", "a", , , vbBinaryCompare)
        sValue = Replace(sValue, "Á", "A", , , vbBinaryCompare)
        sValue = Replace(sValue, "ä", "a", , , vbBinaryCompare)
        sValue = Replace(sValue, "Ä", "A", , , vbBinaryCompare)
        sValue = Replace(sValue, "é", "e", , , vbBinaryCompare)
        sValue = Replace(sValue, "É", "E", , , vbBinaryCompare)
        sValue = Replace(sValue, "ë", "e", , , vbBinaryCompare)
        sValue = Replace(sValue, "Ë", "E", , , vbBinaryCompare)
        sValue = Replace(sValue, "í", "i", , , vbBinaryCompare)
        sValue = Replace(sValue, "Í", "I", , , vbBinaryCompare)
        sValue = Replace(sValue, "ï", "i", , , vbBinaryCompare)
        sValue = Replace(sValue, "Ï", "I", , , vbBinaryCompare)
        sValue = Replace(sValue, "ó", "o", , , vbBinaryCompare)
        sValue = Replace(sValue, "Ó", "O", , , vbBinaryCompare)
        sValue = Replace(sValue, "ö", "o", , , vbBinaryCompare)
        sValue = Replace(sValue, "Ö", "O", , , vbBinaryCompare)
        sValue = Replace(sValue, "ú", "u", , , vbBinaryCompare)
        sValue = Replace(sValue, "Ú", "U", , , vbBinaryCompare)
        sValue = Replace(sValue, "ü", "u", , , vbBinaryCompare)
        sValue = Replace(sValue, "Ü", "U", , , vbBinaryCompare)
    End If

    ' 3. Si la codificación es ISO o SEPA
    If ((lStripStr = enumStripStr.ISO) Or _
        (lStripStr = enumStripStr.SEPA)) Then
        ' 4. Reemplazamos los caracteres "Ñ" y "Ç"
        sValue = Replace(sValue, "ñ", "n", , , vbBinaryCompare)
        sValue = Replace(sValue, "Ñ", "N", , , vbBinaryCompare)
        sValue = Replace(sValue, "ç", "c", , , vbBinaryCompare)
        sValue = Replace(sValue, "Ç", "C", , , vbBinaryCompare)
    End If

    ' 5. Depuramos caracteres extraños según el tipo de codificación
    Select Case lStripStr
        Case enumStripStr.BOE      ' Codificación del formato BOE (ISO-8859-1)
            ' 6. Limpiamos los caracteres no correspondientes al formato BOE
            sValue = modSystem.ReplStr(sValue, "[^A-Za-z0-9/\ //\Ç//\ç//\Ñ//\ñ//\.//\,/]")

        Case enumStripStr.ISO      ' Codificación ISO
            ' 7. Limpiamos los caracteres no alfanuméricos incluyendo los espacios
            sValue = modSystem.ReplStr(sValue, "[^A-Za-z0-9]")

        Case enumStripStr.SEPA     ' Codificación SEPA
            ' 8. Limpiamos los caracteres no alfanuméricos que no correspondan al ISO-20022
            sValue = modSystem.ReplStr(sValue, "[^A-Za-z0-9/\///\-//\?//\://\(//\)//\.//\,//\'//\+//\ //\&//\<//\>//\""//\'/]")

            ' 9. Para la codificación SEPA sustituimos caracteres especiales
            sValue = Replace(sValue, "&", "&amp;", , , vbBinaryCompare)
            sValue = Replace(sValue, "<", "&lt;", , , vbBinaryCompare)
            sValue = Replace(sValue, ">", "&gt;", , , vbBinaryCompare)
            sValue = Replace(sValue, """", "&quot;", , , vbBinaryCompare)
            sValue = Replace(sValue, "'", "&apos;", , , vbBinaryCompare)

        Case enumStripStr.SQL      ' Codificación SQL
            ' 10. Limpiamos los caracteres no permitidos en cadenas de texto SQL
            sValue = modSystem.ReplStr(sValue, "[^A-Za-z0-9/\á//\Á//\ä//\Ä//\é//\É//\ë//\Ë//\í//\Í//\ï//\Ï//\ó//\Ó//\ö//\Ö//\ú//\Ú//\ü//\Ü//\ //\Ç//\ç//\Ñ//\ñ//\.//\,//\://\;//\_//\'//\""//\&//\+//\-//\*//\///\^//\\//\(//\)//\=//\<//\>//\(//\)//\[//\]/]")

            ' 11. Depuramos las comillas simples y dobles
            sValue = Replace(sValue, "'", "''", , , vbBinaryCompare)
            sValue = Replace(sValue, """", """""", , , vbBinaryCompare)
    End Select

    ' 12. Depuramos espacios en exceso
    StripStr = Replace(sValue, Space(2), Space(1), , , vbBinaryCompare)

    Exit Function

Error_StripStr:
    GetError "modSystem.StripStr"
End Function

Public Function OpenFilePicker() As String
    On Error GoTo Error_OpenFilePicker

    With Application.FileDialog(enumMsoFileDialogType.msoFileDialogFilePicker)
        .Show
        If Not IsZr(.SelectedItems.Count) Then OpenFilePicker = .SelectedItems(1)
    End With

    Exit Function

Error_OpenFilePicker:
    GetError "modSystem.OpenFilePicker", "FilePaths=" & OpenFilePicker
End Function

Public Function GetFileToOpenV2() As String
    On Error GoTo Error_GetFileToOpenV2

    With Application.FileDialog(enumMsoFileDialogType.msoFileDialogOpen)
        .Show
        If Not IsZr(.SelectedItems.Count) Then GetFileToOpenV2 = .SelectedItems(1)
    End With

    Exit Function

Error_GetFileToOpenV2:
    GetError "modSystem.GetFileToOpenV2"
End Function

Public Function GetFileToSaveV2() As String
    On Error GoTo Error_GetFileToSaveV2

    With Application.FileDialog(enumMsoFileDialogType.msoFileDialogSaveAs)
        .Show
        If Not IsZr(.SelectedItems.Count) Then GetFileToSaveV2 = .SelectedItems(1)
    End With

    Exit Function

Error_GetFileToSaveV2:
    GetError "modSystem.GetFileToSaveV2"
End Function

Public Function OpenFolderPicker() As String
    On Error GoTo Error_OpenFolderPicker

    With Application.FileDialog(enumMsoFileDialogType.msoFileDialogFolderPicker)
        ' 1. Mostramos la venta de diálogo
        .Show

        ' 2. Si se ha seleccionado ficheros
        If Not IsZr(.SelectedItems.Count) Then
            Dim m_lIdx      As Long

            ' 3. Recorremos la colección de ficheros seleccionados
            For m_lIdx = 1 To .SelectedItems.Count
                ' 4. Construimos la cadena de rutas de ficheros separando por ";"
                If IsZrStr(OpenFolderPicker) Then
                    OpenFolderPicker = .SelectedItems(m_lIdx) & "\"
                Else
                    Concat OpenFolderPicker, ";" & .SelectedItems(m_lIdx)
                End If
            Next m_lIdx
        End If
    End With

    Exit Function

Error_OpenFolderPicker:
    GetError "modSystem.OpenFolderPicker", "FolderPaths=" & OpenFolderPicker
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Función para abrir un fichero MS documento Excel.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       sMSExcelFilePath    String                      Ruta del fichero MS Excel.
'
'       lWindowStyle        VbAppWinStyle (Opcional)    Modo de apertura de la venta con el fichero Excel.
'
'       lHwnd               Long          (Opcional)    Manejador de ventana.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function OpenMSExcel(ByVal sMSExcelFilePath As String, _
                            Optional ByVal lWindowStyle As VBA.VbAppWinStyle = VBA.VbAppWinStyle.vbNormalFocus, _
                            Optional ByVal lhWnd As Long = 0) As Boolean
    If modSystem.FileExists(sMSExcelFilePath) Then OpenMSExcel = Not IsZr(ShellExecute(IIf(IsZr(lhWnd), Application.hWndAccessApp, lhWnd), "open", sMSExcelFilePath, vbNullString, vbNullString, lWindowStyle))
End Function

' Función para abrir un archivo HTML
Public Function OpenURL(ByVal sURL As String, _
                        Optional ByVal lWindowStyle As VBA.VbAppWinStyle = VBA.VbAppWinStyle.vbNormalFocus) As Boolean
    On Error GoTo Error_OpenURL

    If Not IsZrStr(sURL) Then OpenURL = (Not IsZr(ShellExecute(Application.hWndAccessApp, "open", sURL, "", "", lWindowStyle)))

    Exit Function

Error_OpenURL:
    GetError "modSystem.OpenURL"
End Function


'-----------------------------------------------------------------------------------------------------------------------
' Función para descomprimir el contenido de un fichero Zip.
'-----------------------------------------------------------------------------------------------------------------------
'   Parámetros:
'
'       sZipFilePath        String                  Ruta del fichero Zip a descomprimir.
'
'       sFolderPath         String (Opcional)       Ruta de la carpeta donde descomprimir el fichero Zip.
'
'       sFileNameToExtract  String (Opcional)       Fichero a descomprimir del fichero Zip. Si no se indica,
'                                                   se asume que se extraeran todos los elementos.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Function UnZip(ByVal sZIPFilePath As String, _
                      Optional ByRef sFolderPath As String = vbNullString, _
                      Optional ByVal sFileNameToExtract As String = vbNullString) As Boolean
    On Error GoTo Error_UnZip

    ' 1. Si no se ha indicado la ruta donde descomprimir los ficheros,
    '    se toma por defecto la ruta de una carpeta con el nombre del
    '    mismo fichero Zip a descomprimir, y en del misma ubicación
    If IsZrStr(sFolderPath) Then sFolderPath = (modSystem.GetFolderPath(sZIPFilePath) & modSystem.GetFileName(sZIPFilePath) & "\")

    ' 2. Si no existe la ruta de la carpeta
    '    donde descomprimir, la creamos
    If Not modSystem.FolderExists(sFolderPath) Then modSystem.BuiltPath sFolderPath

    With CreateObject("Shell.Application")
        ' 3. Si no se indicó un único fichero a extraer
        If IsZrStr(sFileNameToExtract) Then
            ' 4. Descomprimimos todos los elementos del fichero Zip
            .NameSpace(Error$ & sFolderPath).CopyHere .NameSpace(Error$ & sZIPFilePath).Items

            UnZip = True
        Else
            ' 5. Descomprimimos ese fichero únicamente
            .NameSpace(Error$ & sFolderPath).CopyHere .NameSpace(Error$ & sZIPFilePath).Items.Item(Error$ & sFileNameToExtract)

            UnZip = modSystem.FileExists(sFolderPath & sFileNameToExtract)
        End If
    End With

    Exit Function

Error_UnZip:
    GetError "modSystem.UnZip"
End Function

' Función para insertar una entrada en un archivo de texto.
Public Function WriteFileV2(ByVal sFilePath As String, _
                          ByVal sText As String, _
                          Optional ByVal lCharCode As enumCharcode = enumCharcode.CdoISO_8859_1, _
                          Optional ByVal lSaveCreate As enumSaveCreate = enumSaveCreate.adSaveCreateNotExist) As Boolean
    On Error GoTo Error_WriteFileV2

    If Not IsZrStr(sFilePath) Then
        With CreateObject("ADODB.Stream")
            .Open
            .Charset = modSystem.GetCharCode(lCharCode)
            .WriteText sText
            .SaveToFile sFilePath
            .Close
        End With
    End If
    WriteFileV2 = modSystem.FileExists(sFilePath)

    Exit Function

Error_WriteFileV2:
    GetError "modSystem.WriteFileV2"
End Function

' Función para insertar una entrada en un archivo de texto. (Ya no se emplea)
Public Function WriteFile(ByVal sFilePath As String, _
                          ByVal sText As String) As Boolean
    On Error GoTo Error_WriteFile

    If Not IsZrStr(sFilePath) Then
        Open sFilePath For Output As #1
        Print #1, sText
        Close #1
    End If
    WriteFile = modSystem.FileExists(sFilePath)

    Exit Function

Error_WriteFile:
    Close #1
    GetError "modSystem.WriteFile"
End Function

