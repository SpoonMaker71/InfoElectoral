' ---------------------------------------------------------------
' Módulo: modTemp.bas
' Autor: Juan Francisco Cucharero Cabezas
' Proyecto: InfoElectoral
' Descripción:
'   Funciones auxiliares para el manejo de datos temporales durante
'   la importación, transformación y validación de microdatos electorales.
'
'   ⚠️ Este módulo no contiene interacción directa con el usuario.
'   Su propósito es servir como capa intermedia entre la lectura de
'   ficheros y la inserción definitiva en tablas del sistema.
'
' Fecha: [añadir fecha]
' ---------------------------------------------------------------

Attribute VB_Name = "modTemp"
Option Compare Database
Option Explicit

Public Sub ImportarFichero()
    Dim m_sRutaFicheroTXT          As String

    ' 1. Solicitamos al usuario la ruta de fichero ZIP a importar
    m_sRutaFicheroTXT = modSystem.OpenFilePicker()

    ' 2. Si existe el fichero TXT con los datos del proceso electoral
    If modSystem.FileExists(m_sRutaFicheroTXT) Then
        Dim m_lIdINEComunidadAutonoma       As Long
        Dim m_lIdINEProvincia               As Long
        Dim m_lIdMunicipio                  As Long
        Dim m_sLine()                       As String
        Dim m_lIdx                          As Long
        Dim m_sSql                          As String

        Screen.MousePointer = 11

        m_lIdINEComunidadAutonoma = CLng(Split(modSystem.GetFileName(m_sRutaFicheroTXT), "-")(0))
        m_lIdINEProvincia = CLng(Split(modSystem.GetFileName(m_sRutaFicheroTXT), "-")(1))

        m_sSql = "DELETE IMP_Colegios_Electorales.*"
        Concat m_sSql, " FROM IMP_Colegios_Electorales"
        Concat m_sSql, " WHERE ((IMP_Colegios_Electorales.IdTipoProcesoElectoral = 2)"
        Concat m_sSql, " AND (IMP_Colegios_Electorales.Año = 2019)"
        Concat m_sSql, " AND (IMP_Colegios_Electorales.Mes = 11)"
        Concat m_sSql, " AND (IMP_Colegios_Electorales.Vuelta = 1)"
        Concat m_sSql, " AND (IMP_Colegios_Electorales.IdINEComunidadAutonoma = " & CSqlDbl(m_lIdINEComunidadAutonoma) & ")"
        Concat m_sSql, " AND (IMP_Colegios_Electorales.IdINEProvincia = " & CSqlDbl(m_lIdINEProvincia) & "));"
        CurrentDb.Execute m_sSql, dbFailOnError

        m_sLine = Split(modSystem.ReadFile(m_sRutaFicheroTXT, enumCharcode.CdoUTF_8), vbCrLf)

        modMDB.SetProgressBar LBound(m_sLine), UBound(m_sLine), "Importando fichero", "Colegios y mesas electorales de " & Split(modSystem.GetFileName(m_sRutaFicheroTXT), "-")(2) & "."

        For m_lIdx = LBound(m_sLine) To UBound(m_sLine)
            modMDB.IncrProgressBar
            If (Mid(m_sLine(m_lIdx), 1, Len("Municipio: ")) = "Municipio: ") Then
                m_lIdMunicipio = CLng(Mid(m_sLine(m_lIdx), 12, 3))
            Else
                m_sSql = "INSERT INTO IMP_Colegios_Electorales (IdTipoProcesoElectoral, Año, Mes, Vuelta, IdINEComunidadAutonoma, IdINEProvincia, IdINEMunicipio, IdDistritoMunicipal, IdSeccion, IdMesa, Tramo, Colegio)"
                Concat m_sSql, " VALUES "
                Concat m_sSql, "(2, "
                Concat m_sSql, " 2019,"
                Concat m_sSql, " 11,"
                Concat m_sSql, " 1,"
                Concat m_sSql, " " & CSqlDbl(m_lIdINEComunidadAutonoma) & ","
                Concat m_sSql, " " & CSqlDbl(m_lIdINEProvincia) & ","
                Concat m_sSql, " " & CSqlDbl(m_lIdMunicipio) & ","
                Concat m_sSql, " " & CSqlDbl(Left(m_sLine(m_lIdx), 2)) & ","
                Concat m_sSql, " " & CSqlTxt(Mid(m_sLine(m_lIdx), 4, 3)) & ","
                Concat m_sSql, " " & CSqlTxt(Mid(m_sLine(m_lIdx), 8, 1)) & ","
                Concat m_sSql, " " & CSqlTxt(Replace(Mid(m_sLine(m_lIdx), 10, 5), Space(1), vbNullString)) & ","
                Concat m_sSql, " " & CSqlTxt(UCase(IIf(IsNumeric(Left(Mid(m_sLine(m_lIdx), 15), 10)), Mid(m_sLine(m_lIdx), 25), Mid(m_sLine(m_lIdx), 15))), True) & ");"
                CurrentDb.Execute m_sSql, dbFailOnError
            End If
        Next m_lIdx

        Debug.Print "Se ha importado: " & modSystem.GetFileName(m_sRutaFicheroTXT)

        modMDB.CloseProgressBar

        Screen.MousePointer = 0
    End If
End Sub

Public Sub ImportarFicherosTXT()
    Dim m_sColFicherosTXT               As New Collection
    Dim m_lFicheroTXT                   As Long
    Dim m_sFolderPath                   As String
    Dim m_sRutaFicheroTXT               As String
    Dim m_lIdINEComunidadAutonoma       As Long
    Dim m_lIdINEProvincia               As Long
    Dim m_lIdMunicipio                  As Long
    Dim m_sLine()                       As String
    Dim m_lIdx                          As Long
    Dim m_sSql                          As String

    ' 1. Solicitamos al usuario la ruta de la carpeta con los ficheros TXT a importar
    m_sFolderPath = modSystem.OpenFolderPicker()

    Screen.MousePointer = 11

    ' 2. Obtenemos la colección de ficheros TXT que contiene la carpeta seleccionada
    m_sRutaFicheroTXT = Dir(m_sFolderPath & "*.txt", vbArchive)
    While Not IsZrStr(m_sRutaFicheroTXT)
        m_sColFicherosTXT.Add m_sFolderPath & m_sRutaFicheroTXT
        m_sRutaFicheroTXT = Dir()
    Wend

    ' 3. Si habia ficheros en la colección
    If Not IsZr(m_sColFicherosTXT.Count) Then
        ' 4. Estableceos la barra de progreso
        modMDB.SetProgressBar 0, m_sColFicherosTXT.Count, "Importando Ficheros TXT"

        ' 5. Recorremos la colección de ficheros
        For m_lFicheroTXT = 1 To m_sColFicherosTXT.Count
            ' 6. Incrementamos la barra de progreso
            modMDB.IncrProgressBar 1, "Importando fichero correspondiente a " & Split(modSystem.GetFileName(m_sColFicherosTXT(m_lFicheroTXT)), "-")(2) & ".", True

            ' 7. Obtenemos los códigos INE de la comunidad autónoma y de la provincia del nombre del fichero
            m_lIdINEComunidadAutonoma = CLng(Split(modSystem.GetFileName(m_sColFicherosTXT(m_lFicheroTXT)), "-")(0))
            m_lIdINEProvincia = CLng(Split(modSystem.GetFileName(m_sColFicherosTXT(m_lFicheroTXT)), "-")(1))

            ' 8. Eliminamos las mesas si ya se hubieran importado previamente
            m_sSql = "DELETE IMP_Colegios_Electorales.*"
            Concat m_sSql, " FROM IMP_Colegios_Electorales"
            Concat m_sSql, " WHERE ((IMP_Colegios_Electorales.IdTipoProcesoElectoral = 2)"
            Concat m_sSql, " AND (IMP_Colegios_Electorales.Año = 2019)"
            Concat m_sSql, " AND (IMP_Colegios_Electorales.Mes = 11)"
            Concat m_sSql, " AND (IMP_Colegios_Electorales.Vuelta = 1)"
            Concat m_sSql, " AND (IMP_Colegios_Electorales.IdINEComunidadAutonoma = " & CSqlDbl(m_lIdINEComunidadAutonoma) & ")"
            Concat m_sSql, " AND (IMP_Colegios_Electorales.IdINEProvincia = " & CSqlDbl(m_lIdINEProvincia) & "));"
            CurrentDb.Execute m_sSql, dbFailOnError

            ' 9. Cargamos en un array los datos del fichero
            m_sLine = Split(modSystem.ReadFile(m_sColFicherosTXT(m_lFicheroTXT), enumCharcode.CdoUTF_8), vbCrLf)

            ' 10. Recorremos las líneas del arrray
            For m_lIdx = LBound(m_sLine) To UBound(m_sLine)
                ' 11. Si la línea actual corresponde al municipio
                If (Mid(m_sLine(m_lIdx), 1, Len("Municipio: ")) = "Municipio: ") Then
                    ' 12. Tomamos el código INE del municipio
                    m_lIdMunicipio = CLng(Mid(m_sLine(m_lIdx), 12, 3))
                Else
                    ' 13. Insertamos en la tabla el registro de la mesa electoral
                    m_sSql = "INSERT INTO IMP_Colegios_Electorales (IdTipoProcesoElectoral, Año, Mes, Vuelta, IdINEComunidadAutonoma, IdINEProvincia, IdINEMunicipio, IdDistritoMunicipal, IdSeccion, IdMesa, Tramo, Colegio)"
                    Concat m_sSql, " VALUES "
                    Concat m_sSql, "(2, "
                    Concat m_sSql, " 2019,"
                    Concat m_sSql, " 11,"
                    Concat m_sSql, " 1,"
                    Concat m_sSql, " " & CSqlDbl(m_lIdINEComunidadAutonoma) & ","
                    Concat m_sSql, " " & CSqlDbl(m_lIdINEProvincia) & ","
                    Concat m_sSql, " " & CSqlDbl(m_lIdMunicipio) & ","
                    Concat m_sSql, " " & CSqlDbl(Left(m_sLine(m_lIdx), 2)) & ","
                    Concat m_sSql, " " & CSqlTxt(Mid(m_sLine(m_lIdx), 4, 3)) & ","
                    Concat m_sSql, " " & CSqlTxt(Mid(m_sLine(m_lIdx), 8, 1)) & ","
                    Concat m_sSql, " " & CSqlTxt(Replace(Mid(m_sLine(m_lIdx), 10, 5), Space(1), vbNullString)) & ","
                    Concat m_sSql, " " & CSqlTxt(UCase(IIf(IsNumeric(Left(Mid(m_sLine(m_lIdx), 15), 10)), Mid(m_sLine(m_lIdx), 25), Mid(m_sLine(m_lIdx), 15))), True) & ");"
                    CurrentDb.Execute m_sSql, dbFailOnError
                End If
            Next m_lIdx
        Next m_lFicheroTXT

        ' 14. Cerramos la barra de progreso
        modMDB.CloseProgressBar
    End If

    Screen.MousePointer = 0
End Sub

Public Sub LimpiarFichero()
    Dim m_sRutaFicheroTXT               As String

    ' 1. Solicitamos al usuario el fichero TXT a importar
    m_sRutaFicheroTXT = modSystem.OpenFilePicker()

    ' 2. Si existe el fichero TXT con los datos del proceso electoral
    If modSystem.FileExists(m_sRutaFicheroTXT) Then
        Dim m_sLine()                       As String
        Dim m_sLinea                        As String
        Dim m_lIdx                          As Long
        Dim m_sRutaFicheroDestinoTXT        As String
        Dim m_sText                         As String

        Screen.MousePointer = 11

        m_sLine = Split(modSystem.ReadFile(m_sRutaFicheroTXT, enumCharcode.CdoUTF_8), vbCrLf)

        modMDB.SetProgressBar LBound(m_sLine), UBound(m_sLine), "Procesando fichero", "Limpiando el fichero de datos no necesarios."

        m_sRutaFicheroDestinoTXT = modSystem.GetFolderPath(m_sRutaFicheroTXT) & modSystem.GetFileName(m_sRutaFicheroTXT) & "_Limpiado." & modSystem.GetFileExt(m_sRutaFicheroTXT)

        modSystem.DeleteIfFileExists m_sRutaFicheroDestinoTXT

        m_sText = vbNullString
        For m_lIdx = LBound(m_sLine) To UBound(m_sLine)
            modMDB.IncrProgressBar

            m_sLinea = Trim(m_sLine(m_lIdx))

            If (Left(Trim(m_sLinea), 10) = "MUNICIPIO:") Or (Left(m_sLinea, 10) = "DISTRITO.:") Or (Left(m_sLinea, 16) = "LOCAL ELECTORAL:") Or (Left(m_sLinea, 10) = "DIRECCION:") Or EvalRegExp(Left(m_sLinea, 5), "^[0-9]{5}$") Or (Left(m_sLinea, 10) = "EN LA MESA") Then Concat m_sText, IIf((m_lIdx = LBound(m_sLine)), vbNullString, vbCrLf) & m_sLinea
        Next m_lIdx

        modSystem.WriteFileV2 m_sRutaFicheroDestinoTXT, m_sText, enumCharcode.CdoUTF_8

        modMDB.CloseProgressBar

        Screen.MousePointer = 0
    End If
End Sub

Public Sub FormatearFicheroV1()
    Dim m_sRutaFicheroTXT               As String

    ' 1. Solicitamos al usuario la ruta de fichero TXT a importar
    m_sRutaFicheroTXT = modSystem.OpenFilePicker()

    ' 2. Si existe el fichero TXT con los datos del proceso electoral
    If modSystem.FileExists(m_sRutaFicheroTXT) Then
        Dim m_sLine()                       As String
        Dim m_sLinea                        As String
        Dim m_lIdx                          As Long
        Dim m_sRutaFicheroDestinoTXT        As String
        Dim m_iNumMesas                     As Integer
        Dim m_sText                         As String
        Dim m_sMunicipio                    As String
        Dim m_sMesa                         As String
        Dim m_sLocal                        As String
        Dim m_sRango                        As String

        Screen.MousePointer = 11

        m_sLine = Split(modSystem.ReadFile(m_sRutaFicheroTXT, enumCharcode.CdoUTF_8), vbCrLf)

        modMDB.SetProgressBar LBound(m_sLine), UBound(m_sLine), "Procesando fichero", "Limpiando el fichero de datos no necesarios."

        m_sRutaFicheroDestinoTXT = modSystem.GetFolderPath(m_sRutaFicheroTXT) & modSystem.GetFileName(m_sRutaFicheroTXT) & "_Formateado." & modSystem.GetFileExt(m_sRutaFicheroTXT)

        modSystem.DeleteIfFileExists m_sRutaFicheroDestinoTXT

        m_sText = vbNullString
        For m_lIdx = LBound(m_sLine) To UBound(m_sLine)
            modMDB.IncrProgressBar

            m_sLinea = Trim(m_sLine(m_lIdx))

            If (Left(Trim(m_sLinea), 10) = "MUNICIPIO:") Then
                If (m_sMunicipio <> m_sLinea) Then
                    m_sMunicipio = m_sLinea
                    Concat m_sText, IIf(IsZrStr(m_sText), vbNullString, vbCrLf) & Replace(UCase(m_sLinea), " ", " - ")
                End If

            ElseIf (Left(m_sLinea, 10) = "DISTRITO.:") Then
                m_iNumMesas = Len(Replace(Split(m_sLinea, ":")(UBound(Split(m_sLinea, ":"))), " ", vbNullString))
                m_sMesa = Replace(Replace(Replace(Replace(Replace(Replace(m_sLinea, "DISTRITO.:", vbNullString), "SECCION.:", vbNullString), "SUBSECCION.:", vbNullString), "SUB", vbNullString), "MESA/S.:", vbNullString), Space(2), Space(1))
                m_sMesa = Trim(Mid(m_sMesa, 1, (Len(m_sMesa) - ((m_iNumMesas * 2) - 1))))

            ElseIf (Left(m_sLinea, 16) = "LOCAL ELECTORAL:") Then
                m_sLocal = Space(1) & Trim(Replace(m_sLinea, "LOCAL ELECTORAL:", vbNullString))

            ElseIf (Left(m_sLinea, 10) = "DIRECCION:") Then
                Concat m_sLocal, Space(1) & Trim(Replace(m_sLinea, "DIRECCION:", vbNullString))

            ElseIf (Left(m_sLinea, 10) = "EN LA MESA") Then
                Concat m_sText, vbCrLf & m_sMesa & Space(1) & Mid(m_sLinea, 12, 1) & Space(1) & Replace(Mid(m_sLinea, (InStr(1, m_sLinea, " Y ") - 1), 5), "Y", "-") & m_sLocal

            Else
                Concat m_sLocal, Space(1) & Trim(m_sLinea)
            End If
        Next m_lIdx

        modSystem.WriteFileV2 m_sRutaFicheroDestinoTXT, m_sText, enumCharcode.CdoUTF_8

        modMDB.CloseProgressBar

        Screen.MousePointer = 0
    End If
End Sub

Public Sub FormatearFicheroV2()
    Dim m_sRutaFicheroTXT               As String

    ' 1. Solicitamos al usuario la ruta de fichero TXT a importar
    m_sRutaFicheroTXT = modSystem.OpenFilePicker()

    ' 2. Si existe el fichero TXT con los datos del proceso electoral
    If modSystem.FileExists(m_sRutaFicheroTXT) Then
        Dim m_sLine()                       As String
        Dim m_sLinea                        As String
        Dim m_lIdx                          As Long
        Dim m_sRutaFicheroDestinoTXT        As String
        Dim m_sText                         As String
        Dim m_sIdMunicipio                  As String * 3
        Dim m_iPos                          As Integer

        Screen.MousePointer = 11

        m_sLine = Split(modSystem.ReadFile(m_sRutaFicheroTXT, enumCharcode.CdoUTF_8), vbCrLf)

        modMDB.SetProgressBar LBound(m_sLine), UBound(m_sLine), "Procesando fichero", "Limpiando el fichero de datos no necesarios."

        m_sRutaFicheroDestinoTXT = modSystem.GetFolderPath(m_sRutaFicheroTXT) & modSystem.GetFileName(m_sRutaFicheroTXT) & "_Formateado." & modSystem.GetFileExt(m_sRutaFicheroTXT)

        modSystem.DeleteIfFileExists m_sRutaFicheroDestinoTXT

        m_sText = vbNullString
        For m_lIdx = LBound(m_sLine) To UBound(m_sLine)
            modMDB.IncrProgressBar

            m_sLinea = Trim(m_sLine(m_lIdx))
            
            If (m_sIdMunicipio <> Left(m_sLinea, 3)) Then
                m_sIdMunicipio = Left(m_sLinea, 3)
                m_iPos = InStr(5, m_sLinea, " 01 ")
                
                Concat m_sText, IIf(IsZrStr(m_sText), vbNullString, vbCrLf) & "MUNICIPIO: " & m_sIdMunicipio & " - " & Mid(m_sLinea, 5, (m_iPos - 5))
            End If
            Concat m_sText, vbCrLf & Mid(m_sLinea, (m_iPos + 1))
        Next m_lIdx

        modSystem.WriteFileV2 m_sRutaFicheroDestinoTXT, m_sText, enumCharcode.CdoUTF_8

        modMDB.CloseProgressBar

        Screen.MousePointer = 0
    End If
End Sub

Public Sub FormatearFicheroV3()
    Dim m_sRutaFicheroTXT               As String

    ' 1. Solicitamos al usuario la ruta de fichero TXT a importar
    m_sRutaFicheroTXT = modSystem.OpenFilePicker()

    ' 2. Si existe el fichero TXT con los datos del proceso electoral
    If modSystem.FileExists(m_sRutaFicheroTXT) Then
        Dim m_sLine()                       As String
        Dim m_sLinea                        As String
        Dim m_lIdx                          As Long
        Dim m_sRutaFicheroDestinoTXT        As String
        Dim m_sText                         As String
        Dim m_sIdDistrito                   As String * 2
        Dim m_sIdSeccion                    As String * 3

        Screen.MousePointer = 11

        m_sLine = Split(modSystem.ReadFile(m_sRutaFicheroTXT, enumCharcode.CdoUTF_8), vbCrLf)

        modMDB.SetProgressBar LBound(m_sLine), UBound(m_sLine), "Procesando fichero", "Limpiando el fichero de datos no necesarios."

        m_sRutaFicheroDestinoTXT = modSystem.GetFolderPath(m_sRutaFicheroTXT) & modSystem.GetFileName(m_sRutaFicheroTXT) & "_Formateado." & modSystem.GetFileExt(m_sRutaFicheroTXT)

        modSystem.DeleteIfFileExists m_sRutaFicheroDestinoTXT

        m_sText = vbNullString
        For m_lIdx = LBound(m_sLine) To UBound(m_sLine)
            modMDB.IncrProgressBar

            m_sLinea = Trim(m_sLine(m_lIdx))

            If (Left(m_sLinea, Len("MUNICIPIO:")) = "MUNICIPIO:") Then
                Concat m_sText, IIf(IsZrStr(m_sText), vbNullString, vbCrLf) & m_sLinea

            ElseIf modSystem.EvalRegExp(Left(m_sLinea, 6), "^[0-9]{2}\ [0-9]{3}$") Then
                m_sIdDistrito = Left(m_sLinea, 2)
                m_sIdSeccion = Mid(m_sLinea, 4, 3)
                Concat m_sText, IIf(IsZrStr(m_sText), vbNullString, vbCrLf) & m_sLinea

            ElseIf modSystem.EvalRegExp(Left(m_sLinea, 3), "^[0-9]{3}$") Then
                m_sIdSeccion = Left(m_sLinea, 3)
                Concat m_sText, IIf(IsZrStr(m_sText), vbNullString, vbCrLf) & m_sIdDistrito & Space(1) & m_sLinea
            Else
                Concat m_sText, IIf(IsZrStr(m_sText), vbNullString, vbCrLf) & m_sIdDistrito & Space(1) & m_sIdSeccion & Space(1) & m_sLinea
            End If
        Next m_lIdx

        modSystem.WriteFileV2 m_sRutaFicheroDestinoTXT, m_sText, enumCharcode.CdoUTF_8

        modMDB.CloseProgressBar

        Screen.MousePointer = 0
    End If
End Sub

Public Function Generar_JSON_Municipios() As Boolean
    Dim m_sJSONFilePath         As String
    Dim m_sSql                  As String
    Dim m_rstProvincias         As Recordset

    ' 1. Establecemos la ruta de destino del fichero JSON
    m_sJSONFilePath = modSystem.GetSpecialFolder(Application.hWndAccessApp, USER_DESKTOP) & "AUX_Provincias_Municipios.json"

    ' 2. Eliminamos el fichero JSON si ya existiera
    modSystem.DeleteIfFileExists m_sJSONFilePath

    ' 3. Obtenemos el listado de provincias
    m_sSql = "SELECT"
    Concat m_sSql, " AUX_INE_Provincias.IdINEComunidadAutonoma,"
    Concat m_sSql, " AUX_INE_Provincias.IdINEProvincia,"
    Concat m_sSql, " AUX_INE_Provincias.Nombre"
    Concat m_sSql, " FROM AUX_INE_Provincias"
    Concat m_sSql, " ORDER BY AUX_INE_Provincias.Nombre;"
    Set m_rstProvincias = CurrentDb.OpenRecordset(m_sSql)
    If modMDB.IsRst(m_rstProvincias, True) Then
        Dim m_rstMunicipios         As Recordset
        Dim m_sJSONText             As String

        ' 4. Establecemos la barra de progreso
        modMDB.SetProgressBar 0, m_rstProvincias.RecordCount, "Generando fichero jSON"

        ' 5. Iniciamos la construcción del fichero JSON
        m_sJSONText = "{"

        ' 6. Recorremos las distintas provincias
        While Not m_rstProvincias.EOF
            ' 7. Actualizamos la barra de progreso
            modMDB.IncrProgressBar 1, "Obteniendo municipios de " & m_rstProvincias!NOMBRE & ".", True

            ' 8. Obtenemos los municipios de la provincia actual
            m_sSql = "SELECT"
            Concat m_sSql, " AUX_INE_Municipios.Nombre"
            Concat m_sSql, " FROM AUX_INE_Municipios"
            Concat m_sSql, " WHERE ((AUX_INE_Municipios.IdINEComunidadAutonoma = " & CSqlDbl(m_rstProvincias!IdINEComunidadAutonoma) & ")"
            Concat m_sSql, " AND (AUX_INE_Municipios.IdINEProvincia = " & CSqlDbl(m_rstProvincias!IdINEProvincia) & "))"
            Concat m_sSql, " ORDER BY AUX_INE_Municipios.Nombre;"
            Set m_rstMunicipios = CurrentDb.OpenRecordset(m_sSql)
            If modMDB.IsRst(m_rstMunicipios, True) Then
                ' 9. Insertamos los datos de la provincia, creando el array de municipios
                Concat m_sJSONText, IIf(IsZrStr(m_sJSONText), vbNullString, vbCrLf) & vbTab & """" & m_rstProvincias!NOMBRE & """:["

                ' 10. Recorremos el recordset de municipios de la provincia
                While Not m_rstMunicipios.EOF
                    ' 11. Agregamos el municipio actual
                    Concat m_sJSONText, """" & m_rstMunicipios!NOMBRE & """"

                    ' 12. Pasamos siguiente al registro de municipios
                    m_rstMunicipios.MoveNext

                    ' 13. Si no estamos al final del recordset, insertamos el separados de elementos del array de municipios
                    If Not m_rstMunicipios.EOF Then Concat m_sJSONText, ","
                Wend
                ' 14. Cerranmos el array de municipios
                Concat m_sJSONText, "]"

                ' 15. Cerramos el recordset
                m_rstMunicipios.Close
            End If
            ' 16. Pasamos al siguiente registro de provincias
            m_rstProvincias.MoveNext

            ' 17. Si no estamos al final del recordset, insertamos el separador de elementos del array de provincias
            If Not m_rstProvincias.EOF Then Concat m_sJSONText, ","
        Wend
        Concat m_sJSONText, IIf(IsZrStr(m_sJSONText), vbNullString, vbCrLf & "}")
        m_rstProvincias.Close

        If Not IsZrStr(m_sJSONText) Then Generar_JSON_Municipios = modSystem.WriteFileV2(m_sJSONFilePath, m_sJSONText, CdoUTF_8)

        modMDB.CloseProgressBar
    End If
    Set m_rstMunicipios = Nothing
    Set m_rstProvincias = Nothing
End Function

Public Function Generar_Ficheros_Excel_Comparativa_Actas_Por_Provincia(ByVal lIdTipoProcesoElectoral As Long, _
                                                                       ByVal iAño As Integer, _
                                                                       ByVal iMes As Integer, _
                                                                       ByVal iVuelta As Integer) As Boolean
    Dim m_sSql      As String
    Dim m_Rst       As Recordset

    m_sSql = "SELECT"
    Concat m_sSql, " AUX_INE_Provincias.IdINEProvincia,"
    Concat m_sSql, " AUX_INE_Provincias.Nombre"
    Concat m_sSql, " FROM AUX_INE_Provincias"
    Concat m_sSql, " ORDER BY AUX_INE_Provincias.IdINEProvincia;"
    Set m_Rst = CurrentDb.OpenRecordset(m_sSql)
    If modMDB.IsRst(m_Rst, True) Then
        Generar_Ficheros_Excel_Comparativa_Actas_Por_Provincia = True
        modMDB.SetProgressBar 0, m_Rst.RecordCount, "Generando ficheros Excel de Resultados por Mesa Electoral", "Generando ficheros Excel de Resultados por Mesa Electoral"
        While Not m_Rst.EOF
            modMDB.IncrProgressBar 1, "Generando resultados por mesa electoral para " & m_Rst!NOMBRE & ".", True
            Generar_Ficheros_Excel_Comparativa_Actas_Por_Provincia = Generar_Ficheros_Excel_Comparativa_Actas_Por_Provincia And modTemp.Exportar_Resultados_Candidaturas_por_Provincia_A_Excel(lIdTipoProcesoElectoral, iAño, iMes, iVuelta, m_Rst!IdINEProvincia)
            m_Rst.MoveNext
        Wend
        modMDB.CloseProgressBar
        m_Rst.Close
    End If
    Set m_Rst = Nothing

End Function

Public Function Exportar_Resultados_Candidaturas_por_Provincia_A_Excel(ByVal lIdTipoProcesoElectoral As Long, _
                                                                       ByVal iAño As Integer, _
                                                                       ByVal iMes As Integer, _
                                                                       ByVal iVuelta As Integer, _
                                                                       ByVal iIdINEProvincia As Integer) As Boolean
    Dim m_sSql      As String
    Dim m_Rst       As Recordset

    ' 1. Si existe, eliminamos la tabla temporal con los resultados de las candidaturas por mesas electorales
    modMDB.RemoveIfExists "TMP_Resultados_Candidaturas_por_Mesa_Electoral", acTable

    ' 2. Obtenemos la relación de candidaturas del proceso electoral
    m_sSql = "SELECT"
    Concat m_sSql, " [02XXAAMM].IdTipoProcesoElectoral,"
    Concat m_sSql, " [02XXAAMM].Año,"
    Concat m_sSql, " [02XXAAMM].Mes,"
    Concat m_sSql, " [02XXAAMM].Vuelta,"
    Concat m_sSql, " Format([02XXAAMM].FechaProcesoElectoral, 'yyyy\ ') & Format([02XXAAMM].FechaProcesoElectoral, 'dd') & Left(UCase(Format([02XXAAMM].FechaProcesoElectoral, 'mmm')), 1) AS Titulo,"
    Concat m_sSql, " [10XXAAMM].IdINEProvincia,"
    Concat m_sSql, " AUX_INE_Provincias.Nombre AS Provincia,"
    Concat m_sSql, " [03XXAAMM].IdCandidatura,"
    Concat m_sSql, " [03XXAAMM].Siglas,"
    Concat m_sSql, " [03XXAAMM].Denominacion"
    Concat m_sSql, " FROM (([02XXAAMM]"
    Concat m_sSql, " INNER JOIN [03XXAAMM] ON ([02XXAAMM].IdTipoProcesoElectoral = [03XXAAMM].IdTipoProcesoElectoral) AND ([02XXAAMM].Año = [03XXAAMM].Año) AND ([02XXAAMM].Mes = [03XXAAMM].Mes))"
    Concat m_sSql, " INNER JOIN [10XXAAMM] ON ([03XXAAMM].IdTipoProcesoElectoral = [10XXAAMM].IdTipoProcesoElectoral) AND ([03XXAAMM].Año = [10XXAAMM].Año) AND ([03XXAAMM].Mes = [10XXAAMM].Mes) AND ([02XXAAMM].Vuelta = [10XXAAMM].Vuelta) AND ([03XXAAMM].IdCandidatura = [10XXAAMM].IdCandidatura))"
    Concat m_sSql, " INNER JOIN AUX_INE_Provincias ON [10XXAAMM].IdINEProvincia = AUX_INE_Provincias.IdINEProvincia"
    Concat m_sSql, " WHERE (([02XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(lIdTipoProcesoElectoral) & ")"
    Concat m_sSql, " AND ([02XXAAMM].Año = " & CSqlDbl(iAño) & ")"
    Concat m_sSql, " AND ([02XXAAMM].Mes = " & CSqlDbl(iMes) & ")"
    Concat m_sSql, " AND ([02XXAAMM].Vuelta = " & CSqlDbl(iVuelta) & ")"
    Concat m_sSql, " AND ([10XXAAMM].IdINEProvincia = " & CSqlDbl(iIdINEProvincia) & "))"
    Concat m_sSql, " GROUP BY [02XXAAMM].IdTipoProcesoElectoral,"
    Concat m_sSql, " [02XXAAMM].Año,"
    Concat m_sSql, " [02XXAAMM].Mes,"
    Concat m_sSql, " [02XXAAMM].Vuelta,"
    Concat m_sSql, " Format([02XXAAMM].FechaProcesoElectoral, 'yyyy\ ') & Format([02XXAAMM].FechaProcesoElectoral, 'dd') & Left(UCase(Format([02XXAAMM].FechaProcesoElectoral, 'mmm')), 1),"
    Concat m_sSql, " [10XXAAMM].IdINEProvincia,"
    Concat m_sSql, " AUX_INE_Provincias.Nombre,"
    Concat m_sSql, " [03XXAAMM].IdCandidatura,"
    Concat m_sSql, " [03XXAAMM].Siglas,"
    Concat m_sSql, " [03XXAAMM].Denominacion"
    Concat m_sSql, " ORDER BY [03XXAAMM].IdCandidatura;"
    Set m_Rst = CurrentDb.OpenRecordset(m_sSql)
    If modMDB.IsRst(m_Rst, True) Then
        Dim m_sXLSFilePath              As String
        Dim m_sTitle                   As String

        m_sTitle = m_Rst!Titulo

        m_sXLSFilePath = modMDB.GetMDBPath("Informes") & Replace(Replace(m_sTitle, " / ", "_"), "/", "_") & " - Congreso - " & Replace(Replace(m_Rst!Provincia, " / ", "_"), "/", "_") & modMDB.GetMSExcelDefaultExtension()

        ' 3. Generamos la tabla temporal con los resultados de las candidaturas por mesa electoral
        m_sSql = "SELECT"
        Concat m_sSql, " [02XXAAMM].IdTipoProcesoElectoral,"
        Concat m_sSql, " [02XXAAMM].Año,"
        Concat m_sSql, " [02XXAAMM].Mes,"
        Concat m_sSql, " [02XXAAMM].Vuelta,"
        Concat m_sSql, " [09XXAAMM].IdINEComunidadAutonoma,"
        Concat m_sSql, " [09XXAAMM].IdINEProvincia,"
        Concat m_sSql, " AUX_INE_Provincias.Nombre,"
        Concat m_sSql, " [09XXAAMM].IdINEMunicipio,"
        Concat m_sSql, " AUX_INE_Municipios.Nombre,"
        Concat m_sSql, " [09XXAAMM].IdDistritoMunicipal,"
        Concat m_sSql, " [09XXAAMM].IdSeccion,"
        Concat m_sSql, " [09XXAAMM].IdMesa,"
        Concat m_sSql, " [09XXAAMM].CensoINE,"
        Concat m_sSql, " [09XXAAMM].CensoEscrutinio,"
        Concat m_sSql, " [09XXAAMM].CensoEscrutinioCERE,"
        Concat m_sSql, " [09XXAAMM].TotalVotantesCERE,"
        Concat m_sSql, " [09XXAAMM].VotosEnBlanco,"
        Concat m_sSql, " [09XXAAMM].VotosNulos,"
        Concat m_sSql, " [09XXAAMM].VotosCandidaturas,"
        Concat m_sSql, " [09XXAAMM].DatosOficiales"
        ' 4. Agregamos las candidaturas como campos con su IdCandidatura como nombre del mismo
        While Not m_Rst.EOF
            Concat m_sSql, ", CDbl(0) AS [Campo_" & m_Rst!IdCandidatura & "]"
            m_Rst.MoveNext
        Wend
        Concat m_sSql, " INTO TMP_Resultados_Candidaturas_por_Mesa_Electoral"
        Concat m_sSql, " FROM (([02XXAAMM]"
        Concat m_sSql, " INNER JOIN [09XXAAMM] ON ([02XXAAMM].IdTipoProcesoElectoral = [09XXAAMM].IdTipoProcesoElectoral) AND ([02XXAAMM].Año = [09XXAAMM].Año) AND ([02XXAAMM].Mes = [09XXAAMM].Mes) AND ([02XXAAMM].Vuelta = [09XXAAMM].Vuelta))"
        Concat m_sSql, " INNER JOIN AUX_INE_Municipios ON ([09XXAAMM].IdINEComunidadAutonoma = AUX_INE_Municipios.IdINEComunidadAutonoma) AND ([09XXAAMM].IdINEProvincia = AUX_INE_Municipios.IdINEProvincia) AND ([09XXAAMM].IdINEMunicipio = AUX_INE_Municipios.IdINEMunicipio))"
        Concat m_sSql, " INNER JOIN AUX_INE_Provincias ON (AUX_INE_Provincias.IdINEComunidadAutonoma = AUX_INE_Municipios.IdINEComunidadAutonoma) AND (AUX_INE_Provincias.IdINEProvincia = AUX_INE_Municipios.IdINEProvincia)"
        Concat m_sSql, " WHERE (([02XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(lIdTipoProcesoElectoral) & ")"
        Concat m_sSql, " AND ([02XXAAMM].Año = " & CSqlDbl(iAño) & ")"
        Concat m_sSql, " AND ([02XXAAMM].Mes = " & CSqlDbl(iMes) & ")"
        Concat m_sSql, " AND ([02XXAAMM].Vuelta = " & CSqlDbl(iVuelta) & ")"
        Concat m_sSql, " AND ([09XXAAMM].IdINEProvincia = " & CSqlDbl(iIdINEProvincia) & ")"
        Concat m_sSql, " AND ([09XXAAMM].DatosOficiales = 'S'))"
        Concat m_sSql, " ORDER BY [09XXAAMM].IdINEProvincia,"
        Concat m_sSql, " [09XXAAMM].IdINEMunicipio,"
        Concat m_sSql, " [09XXAAMM].IdDistritoMunicipal,"
        Concat m_sSql, " [09XXAAMM].IdSeccion,"
        Concat m_sSql, " [09XXAAMM].IdMesa;"
        CurrentDb.Execute m_sSql, dbFailOnError

        ' 5. Obtenemos los resultados para cada candidatura
        m_Rst.MoveFirst
        While Not m_Rst.EOF
            m_sSql = "UPDATE ([10XXAAMM]"
            Concat m_sSql, " INNER JOIN [03XXAAMM] ON ([10XXAAMM].IdTipoProcesoElectoral = [03XXAAMM].IdTipoProcesoElectoral)"
            Concat m_sSql, " AND ([10XXAAMM].Año = [03XXAAMM].Año)"
            Concat m_sSql, " AND ([10XXAAMM].Mes = [03XXAAMM].Mes)"
            Concat m_sSql, " AND ([10XXAAMM].IdCandidatura = [03XXAAMM].IdCandidatura))"
            Concat m_sSql, " INNER JOIN TMP_Resultados_Candidaturas_por_Mesa_Electoral ON ([10XXAAMM].IdTipoProcesoElectoral = TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdTipoProcesoElectoral) AND ([10XXAAMM].Año = TMP_Resultados_Candidaturas_por_Mesa_Electoral.Año) AND ([10XXAAMM].Mes = TMP_Resultados_Candidaturas_por_Mesa_Electoral.Mes) AND ([10XXAAMM].Vuelta = TMP_Resultados_Candidaturas_por_Mesa_Electoral.Vuelta) AND ([10XXAAMM].IdINEComunidadAutonoma = TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdINEComunidadAutonoma) AND ([10XXAAMM].IdINEProvincia = TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdINEProvincia) AND ([10XXAAMM].IdINEMunicipio = TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdINEMunicipio) AND ([10XXAAMM].IdDistritoMunicipal = TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdDistritoMunicipal) AND ([10XXAAMM].IdSeccion = TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdSeccion) AND ([10XXAAMM].IdMesa = TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdMesa)"
            Concat m_sSql, " SET TMP_Resultados_Candidaturas_por_Mesa_Electoral.[Campo_" & m_Rst!IdCandidatura & "] = [10XXAAMM].VotosObtenidos"
            Concat m_sSql, " WHERE (([10XXAAMM].IdINEProvincia <> 99)"
            Concat m_sSql, " AND ([10XXAAMM].IdINEMunicipio <> 999)"
            Concat m_sSql, " AND ([10XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(lIdTipoProcesoElectoral) & ")"
            Concat m_sSql, " AND ([10XXAAMM].Año = " & CSqlDbl(iAño) & ")"
            Concat m_sSql, " AND ([10XXAAMM].Mes = " & CSqlDbl(iMes) & ")"
            Concat m_sSql, " AND ([10XXAAMM].Vuelta = " & CSqlDbl(iVuelta) & ")"
            Concat m_sSql, " AND ([10XXAAMM].IdINEProvincia = " & CSqlDbl(iIdINEProvincia) & ")"
            Concat m_sSql, " AND ([10XXAAMM].IdCandidatura = " & CSqlDbl(m_Rst!IdCandidatura) & "));"
            CurrentDb.Execute m_sSql, dbFailOnError
            m_Rst.MoveNext
        Wend
        
        ' 6. Construimos la consulta para generar el fichero Excel con los resultados de las candidaturas por mesa electoral
        m_sSql = "SELECT"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdINEProvincia,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.AUX_INE_Provincias_Nombre,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdINEMunicipio,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.AUX_INE_Municipios_Nombre,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdDistritoMunicipal,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdSeccion,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdMesa,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.CensoINE,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.CensoEscrutinio,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.CensoEscrutinioCERE,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.TotalVotantesCERE,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.VotosEnBlanco,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.VotosNulos,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.VotosCandidaturas"
        ' 7. Agregamos las candidaturas como campos con su IdCandidatura como nombre del mismo
        m_Rst.MoveFirst
        While Not m_Rst.EOF
            Concat m_sSql, ", TMP_Resultados_Candidaturas_por_Mesa_Electoral.[Campo_" & m_Rst!IdCandidatura & "] AS [" & m_Rst!Denominacion & "]"
            m_Rst.MoveNext
        Wend
        Concat m_sSql, " FROM TMP_Resultados_Candidaturas_por_Mesa_Electoral"
        Concat m_sSql, " WHERE ((TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdTipoProcesoElectoral = " & CSqlDbl(lIdTipoProcesoElectoral) & ")"
        Concat m_sSql, " AND (TMP_Resultados_Candidaturas_por_Mesa_Electoral.Año = " & CSqlDbl(iAño) & ")"
        Concat m_sSql, " AND (TMP_Resultados_Candidaturas_por_Mesa_Electoral.Mes = " & CSqlDbl(iMes) & ")"
        Concat m_sSql, " AND (TMP_Resultados_Candidaturas_por_Mesa_Electoral.Vuelta = " & CSqlDbl(iVuelta) & ")"
        Concat m_sSql, " AND (TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdINEProvincia = " & CSqlDbl(iIdINEProvincia) & "))"
        Concat m_sSql, " ORDER BY TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdINEComunidadAutonoma,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdINEProvincia,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.AUX_INE_Provincias_Nombre,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdINEMunicipio,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.AUX_INE_Municipios_Nombre,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdDistritoMunicipal,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdSeccion,"
        Concat m_sSql, " TMP_Resultados_Candidaturas_por_Mesa_Electoral.IdMesa;"

        ' 8. Generamos el fichero MS Excel
        modMDB.ToMSExcel m_sSql, "Resultados candidaturas", m_sXLSFilePath, True, , m_sTitle, m_sTitle, "Resultados electorales por candidatura y mesa electoral en fichero MS Excel", , , , False

        m_Rst.Close
    End If
    Set m_Rst = Nothing

    Exportar_Resultados_Candidaturas_por_Provincia_A_Excel = modSystem.FileExists(m_sXLSFilePath)
End Function

