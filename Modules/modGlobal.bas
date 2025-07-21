Attribute VB_Name = "modGlobal"
Option Compare Database
Option Explicit

Private m_SafeChar(0 To 255) As Boolean

'-----------------------------------------------------------------------------------------------------------------------
' Método para incrementar una variable numérica.
'-----------------------------------------------------------------------------------------------------------------------
' Parámetros:
'
'       vVariable   Variant     Variable a incrementar.
'
'       vValue      Variant     Valor de incremento.
'
'-----------------------------------------------------------------------------------------------------------------------
Public Sub Incr(ByRef vVariable As Variant, _
                ByVal vValue As Variant)
    On Error GoTo Error_Incr

    vVariable = (vVariable + vValue)

    Exit Sub

Error_Incr:
    GetError "modGlobal.Incr"
End Sub

Public Function Update_02XXAAMM_Tabledef() As Boolean
    Dim m_sSql          As String

    On Error GoTo Error_Update_02XXAAMM_Tabledef

    m_sSql = "UPDATE [02XXAAMM] SET [02XXAAMM].[FechaProcesoElectoral] = DateSerial([02XXAAMM].[AñoProcesoElectoral], [02XXAAMM].[MesProcesoElectoral], [02XXAAMM].[DiaProcesoElectoral]);"
    CurrentDb.Execute m_sSql, dbFailOnError

    Exit Function

Error_Update_02XXAAMM_Tabledef:
    GetError "modTemp.Update_02XXAAMM_Tabledef", "m_sSql=" & m_sSql
End Function

Public Function Add_New_Poblations() As Boolean
    Dim m_sSql          As String

    On Error GoTo Error_Add_New_Poblations

    m_sSql = "INSERT INTO AUX_INE_Municipios (IdINEComunidadAutonoma, IdINEProvincia, IdINEMunicipio, Nombre, Concejales)"
    Concat m_sSql, " SELECT DISTINCT"
    Concat m_sSql, " [05XXAAMM].IdINEComunidadAutonoma,"
    Concat m_sSql, " [05XXAAMM].IdINEProvincia,"
    Concat m_sSql, " [05XXAAMM].IdINEMunicipio,"
    Concat m_sSql, " [05XXAAMM].Nombre,"
    Concat m_sSql, " [05XXAAMM].Escaños"
    Concat m_sSql, " FROM [05XXAAMM]"
    Concat m_sSql, " LEFT JOIN AUX_INE_Municipios ON ([05XXAAMM].IdINEProvincia = AUX_INE_Municipios.IdINEProvincia) AND ([05XXAAMM].IdINEMunicipio = AUX_INE_Municipios.IdINEMunicipio)"
    Concat m_sSql, " WHERE (AUX_INE_Municipios.Nombre Is Null);"
    CurrentDb.Execute m_sSql, dbFailOnError

    Exit Function

Error_Add_New_Poblations:
    GetError "modTemp.Add_New_Poblations", "m_sSql=" & m_sSql
End Function

' Return a URL safe encoding of txt.
Public Function URLEncode(EncodeStr As String) As String
    Dim i As Integer
    Dim erg As String
    
    erg = EncodeStr

    ' *** First replace '%' chr
    erg = Replace(erg, "%", Chr(1))

    ' *** then '+' chr
    erg = Replace(erg, "+", Chr(2))
    
    For i = 0 To 255
        Select Case i
            ' *** Allowed 'regular' characters
            Case 37, 43, 48 To 57, 65 To 90, 97 To 122
            
            Case 1  ' *** Replace original %
                erg = Replace(erg, Chr(i), "%25")
        
            Case 2  ' *** Replace original +
                erg = Replace(erg, Chr(i), "%2B")
                
            Case 32
                erg = Replace(erg, Chr(i), "+")
        
            Case 3 To 15
                erg = Replace(erg, Chr(i), "%0" & Hex(i))
        
            Case Else
                erg = Replace(erg, Chr(i), "%" & Hex(i))
        End Select
    Next
    
    URLEncode = erg
End Function

Public Function ProcesarEscrutinio(ByVal IdTipoProcesoElectoral As Single, _
                                   ByVal Año As Integer, _
                                   ByVal Mes As Integer, _
                                   Optional ByVal IdINEProvincia As Single = 0, _
                                   Optional ByVal bEliminarTablasTemporales As Boolean = True) As Boolean
    Dim m_sSql      As String
    Dim m_Rst       As Recordset

    ' 1. Establecemos la barra de progreso
    modMDB.SetProgressBar 0, (IIf(IsZr(IdINEProvincia), modMDB.GetNumRecords("AUX_INE_Provincias", acTable), 1) + 4), "Calculando escrutinio mediante sistema D'Hont", "Preparando datos previos al escrutinio.", 1

    ' 2. Eliminamos el escrutinio si ya existía antes
    modMDB.IncrProgressBar 1
    m_sSql = "DELETE RES_Escaños_Escrutinio.*"
    Concat m_sSql, " FROM RES_Escaños_Escrutinio"
    Concat m_sSql, " WHERE ((RES_Escaños_Escrutinio.IdTipoProcesoElectoral = " & CSqlDbl(IdTipoProcesoElectoral) & ")"
    Concat m_sSql, " AND (RES_Escaños_Escrutinio.Año = " & CSqlDbl(Año) & ")"
    Concat m_sSql, " AND (RES_Escaños_Escrutinio.Mes = " & CSqlDbl(Mes) & "));"
    CurrentDb.Execute m_sSql, dbFailOnError

    ' 3. Obtenemos la relación de provincias, junto con los escaños
    '    que les corresponde para el Congreso de los Diputados
    modMDB.IncrProgressBar 1
    m_sSql = "SELECT"
    Concat m_sSql, " [05XXAAMM].IdTipoProcesoElectoral,"
    Concat m_sSql, " [05XXAAMM].Año,"
    Concat m_sSql, " [05XXAAMM].Mes,"
    Concat m_sSql, " [05XXAAMM].IdINEComunidadAutonoma,"
    Concat m_sSql, " [05XXAAMM].IdINEProvincia,"
    Concat m_sSql, " AUX_INE_Provincias.Nombre AS Provincia,"
    Concat m_sSql, " Sum([05XXAAMM].VotosCandidaturas) AS VotosCandidaturas,"
    Concat m_sSql, " Sum([05XXAAMM].VotosEnBlanco) AS VotosEnBlanco,"
    Concat m_sSql, " Sum([05XXAAMM].VotosNulos) AS VotosNulos,"
    Concat m_sSql, " [07XXAAMM].Escaños"
    Concat m_sSql, " FROM (AUX_INE_Provincias"
    Concat m_sSql, " INNER JOIN [05XXAAMM] ON (AUX_INE_Provincias.IdINEComunidadAutonoma = [05XXAAMM].IdINEComunidadAutonoma) AND (AUX_INE_Provincias.IdINEProvincia = [05XXAAMM].IdINEprovincia))"
    Concat m_sSql, " INNER JOIN [07XXAAMM] ON ([05XXAAMM].IdTipoProcesoElectoral = [07XXAAMM].IdTipoProcesoElectoral) AND ([05XXAAMM].Año = [07XXAAMM].Año) AND ([05XXAAMM].Mes = [07XXAAMM].Mes) AND ([05XXAAMM].IdINEComunidadAutonoma = [07XXAAMM].IdINEComunidadAutonoma) AND ([05XXAAMM].IdINEprovincia = [07XXAAMM].IdINEProvincia)"
    Concat m_sSql, " WHERE (([05XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(IdTipoProcesoElectoral) & ")"
    Concat m_sSql, " AND ([05XXAAMM].Año = " & CSqlDbl(Año) & ")"
    Concat m_sSql, " AND ([05XXAAMM].Mes = " & CSqlDbl(Mes) & ")"
    If Not IsZr(IdINEProvincia) Then Concat m_sSql, " AND ([05XXAAMM].IdINEprovincia = " & CSqlDbl(IdINEProvincia) & ")"
    Concat m_sSql, " AND ([05XXAAMM].IdDistritoMunicipal = 99))"
    Concat m_sSql, " GROUP BY [05XXAAMM].IdTipoProcesoElectoral,"
    Concat m_sSql, " [05XXAAMM].Año,"
    Concat m_sSql, " [05XXAAMM].Mes,"
    Concat m_sSql, " [05XXAAMM].IdINEComunidadAutonoma,"
    Concat m_sSql, " [05XXAAMM].IdINEProvincia,"
    Concat m_sSql, " AUX_INE_Provincias.Nombre,"
    Concat m_sSql, " [07XXAAMM].Escaños"
    Concat m_sSql, " ORDER BY [05XXAAMM].IdINEComunidadAutonoma,"
    Concat m_sSql, " [05XXAAMM].IdINEprovincia;"
    Set m_Rst = CurrentDb.OpenRecordset(m_sSql)
    If modMDB.IsRst(m_Rst, True) Then
        Dim m_iEscaño       As Single

        ' 4. Eliminamos la tabla temporal si existe
        modMDB.IncrProgressBar 1
        modMDB.RemoveIfExists "TMP_Datos_Escrutinio", acTable

        ' 5. Recorremos la relación de provincias
        While Not m_Rst.EOF
            ' 6. Creamos la tabla temporal la relación de candidaturas que
            '    entran en el escrutinio, junto con sus votos obteneidos
            modMDB.IncrProgressBar 1, "Realizando escrutinio de " & m_Rst!Provincia & "."
            If modMDB.ObjectExists("TMP_Datos_Escrutinio", acTable) Then
                m_sSql = "INSERT INTO TMP_Datos_Escrutinio (IdTipoProcesoElectoral, Año, Mes, IdINEComunidadAutonoma, IdINEProvincia, IdCandidatura, IdCandidaturaNivelAutonomico, IdCandidaturaNivelNacional, IdCandidaturaNivelProvincial, VotosObtenidos)"
                Concat m_sSql, " SELECT"
            Else
                m_sSql = "SELECT"
            End If
            Concat m_sSql, " TMP.IdTipoProcesoElectoral,"
            Concat m_sSql, " TMP.Año,"
            Concat m_sSql, " TMP.Mes,"
            Concat m_sSql, " TMP.IdINEComunidadAutonoma,"
            Concat m_sSql, " TMP.IdINEProvincia,"
            Concat m_sSql, " TMP.IdCandidatura,"
            Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelAutonomico,"
            Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelNacional,"
            Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelProvincial,"
            Concat m_sSql, " TMP.VotosObtenidos"
            If Not modMDB.ObjectExists("TMP_Datos_Escrutinio", acTable) Then Concat m_sSql, " INTO TMP_Datos_Escrutinio"
            Concat m_sSql, " FROM [03XXAAMM]"
            Concat m_sSql, " INNER JOIN (SELECT"
            Concat m_sSql, " [03XXAAMM].IdTipoProcesoElectoral,"
            Concat m_sSql, " [03XXAAMM].Año,"
            Concat m_sSql, " [03XXAAMM].Mes,"
            Concat m_sSql, " [10XXAAMM].IdINEComunidadAutonoma,"
            Concat m_sSql, " [10XXAAMM].IdINEProvincia,"
            Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelProvincial AS IdCandidatura,"
            Concat m_sSql, " Sum([10XXAAMM].VotosObtenidos) As VotosObtenidos"
            Concat m_sSql, " FROM [03XXAAMM]"
            Concat m_sSql, " INNER JOIN [10XXAAMM] ON ([03XXAAMM].IdTipoProcesoElectoral = [10XXAAMM].IdTipoProcesoElectoral) AND ([03XXAAMM].Año = [10XXAAMM].Año) AND ([03XXAAMM].Mes = [10XXAAMM].Mes) AND ([03XXAAMM].IdCandidatura = [10XXAAMM].IdCandidatura)"
            Concat m_sSql, " WHERE (([10XXAAMM].IdINEProvincia <> 99)"
            Concat m_sSql, " AND([10XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(m_Rst!IdTipoProcesoElectoral) & ")"
            Concat m_sSql, " AND ([10XXAAMM].Año = " & CSqlDbl(m_Rst!Año) & ")"
            Concat m_sSql, " AND ([10XXAAMM].Mes = " & CSqlDbl(m_Rst!Mes) & ")"
            Concat m_sSql, " AND ([10XXAAMM].IdINEComunidadAutonoma = " & CSqlDbl(m_Rst!IdINEComunidadAutonoma) & ")"
            Concat m_sSql, " AND ([10XXAAMM].IdINEProvincia = " & CSqlDbl(m_Rst!IdINEProvincia) & "))"
            Concat m_sSql, " GROUP BY [03XXAAMM].IdTipoProcesoElectoral,"
            Concat m_sSql, " [03XXAAMM].Año,"
            Concat m_sSql, " [03XXAAMM].Mes,"
            Concat m_sSql, " [10XXAAMM].IdINEComunidadAutonoma,"
            Concat m_sSql, " [10XXAAMM].IdINEProvincia,"
            Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelProvincial) AS TMP ON ([03XXAAMM].IdTipoProcesoElectoral = TMP.IdTipoProcesoElectoral) AND ([03XXAAMM].Año = TMP.Año) AND ([03XXAAMM].Mes = TMP.Mes) AND ([03XXAAMM].IdCandidatura = TMP.IdCandidatura)"
            Concat m_sSql, " WHERE (Round(((TMP.VotosObtenidos / " & CSqlDbl(m_Rst!VotosCandidaturas + m_Rst!VotosEnBlanco) & ") * 100), 2) >= 3)"
            Concat m_sSql, " ORDER BY TMP.VotosObtenidos DESC;"
            CurrentDb.Execute m_sSql, dbFailOnError

            ' 7. Creamos la tabla temporal la relación de candidaturas que
            '    entran en el escrutinio, junto con sus votos obteneidos
            For m_iEscaño = 1 To m_Rst!Escaños
                If modMDB.ObjectExists("TMP_Datos_Dhont", acTable) Then
                    m_sSql = "INSERT INTO TMP_Datos_Dhont (IdTipoProcesoElectoral, Año, Mes, IdINEComunidadAutonoma, IdINEProvincia, IdCandidatura, IdCandidaturaNivelAutonomico, IdCandidaturaNivelNacional, IdCandidaturaNivelProvincial, Dhont)"
                    Concat m_sSql, " SELECT"
                Else
                    m_sSql = "SELECT"
                End If
                Concat m_sSql, " TMP_Datos_Escrutinio.IdTipoProcesoElectoral,"
                Concat m_sSql, " TMP_Datos_Escrutinio.Año,"
                Concat m_sSql, " TMP_Datos_Escrutinio.Mes,"
                Concat m_sSql, " TMP_Datos_Escrutinio.IdINEComunidadAutonoma,"
                Concat m_sSql, " TMP_Datos_Escrutinio.IdINEProvincia,"
                Concat m_sSql, " TMP_Datos_Escrutinio.IdCandidatura,"
                Concat m_sSql, " TMP_Datos_Escrutinio.IdCandidaturaNivelAutonomico,"
                Concat m_sSql, " TMP_Datos_Escrutinio.IdCandidaturaNivelNacional,"
                Concat m_sSql, " TMP_Datos_Escrutinio.IdCandidaturaNivelProvincial,"
                Concat m_sSql, " (TMP_Datos_Escrutinio.VotosObtenidos \ " & CSqlDbl(m_iEscaño) & ") AS Dhont"
                If Not modMDB.ObjectExists("TMP_Datos_Dhont", acTable) Then Concat m_sSql, " INTO TMP_Datos_Dhont"
                Concat m_sSql, " FROM TMP_Datos_Escrutinio"
                Concat m_sSql, " WHERE ((TMP_Datos_Escrutinio.IdTipoProcesoElectoral = " & CSqlDbl(m_Rst!IdTipoProcesoElectoral) & ")"
                Concat m_sSql, " AND (TMP_Datos_Escrutinio.Año = " & CSqlDbl(m_Rst!Año) & ")"
                Concat m_sSql, " AND (TMP_Datos_Escrutinio.Mes = " & CSqlDbl(m_Rst!Mes) & ")"
                Concat m_sSql, " AND (TMP_Datos_Escrutinio.IdINEProvincia = " & CSqlDbl(m_Rst!IdINEProvincia) & "));"
                CurrentDb.Execute m_sSql, dbFailOnError
            Next m_iEscaño

            m_sSql = "INSERT INTO RES_Escaños_Escrutinio (IdTipoProcesoElectoral, Año, Mes, IdINEComunidadAutonoma, IdINEProvincia, IdCandidatura, IdCandidaturaNivelAutonomico, IdCandidaturaNivelNacional, IdCandidaturaNivelProvincial, Dhont, Escaño)"
            Concat m_sSql, " SELECT TOP " & CSqlDbl(m_Rst!Escaños)
            Concat m_sSql, " TMP_Datos_Dhont.IdTipoProcesoElectoral,"
            Concat m_sSql, " TMP_Datos_Dhont.Año,"
            Concat m_sSql, " TMP_Datos_Dhont.Mes,"
            Concat m_sSql, " TMP_Datos_Dhont.IdINEComunidadAutonoma,"
            Concat m_sSql, " TMP_Datos_Dhont.IdINEProvincia,"
            Concat m_sSql, " TMP_Datos_Dhont.IdCandidatura,"
            Concat m_sSql, " TMP_Datos_Dhont.IdCandidaturaNivelAutonomico,"
            Concat m_sSql, " TMP_Datos_Dhont.IdCandidaturaNivelNacional,"
            Concat m_sSql, " TMP_Datos_Dhont.IdCandidaturaNivelProvincial,"
            Concat m_sSql, " TMP_Datos_Dhont.Dhont,"
            Concat m_sSql, " 1 AS Escaño"
            Concat m_sSql, " FROM TMP_Datos_Dhont"
            Concat m_sSql, " WHERE ((TMP_Datos_Dhont.IdTipoProcesoElectoral = " & CSqlDbl(m_Rst!IdTipoProcesoElectoral) & ")"
            Concat m_sSql, " AND (TMP_Datos_Dhont.Año = " & CSqlDbl(m_Rst!Año) & ")"
            Concat m_sSql, " AND (TMP_Datos_Dhont.Mes =  " & CSqlDbl(m_Rst!Mes) & ")"
            Concat m_sSql, " AND (TMP_Datos_Dhont.IdINEProvincia = " & CSqlDbl(m_Rst!IdINEProvincia) & "))"
            Concat m_sSql, " ORDER BY TMP_Datos_Dhont.Dhont DESC;"
            CurrentDb.Execute m_sSql, dbFailOnError

            m_Rst.MoveNext
        Wend
        m_Rst.Close

        ' 8. Si se indica, eliminamos las tablas temporales intermedias
        modMDB.IncrProgressBar 1
        If bEliminarTablasTemporales Then
            modMDB.RemoveIfExists "TMP_Datos_Escrutinio"
            modMDB.RemoveIfExists "TMP_Datos_Dhont"
        End If

        ' 9. Cerramos la barra de progreso
        modMDB.CloseProgressBar

        ProcesarEscrutinio = modMDB.HasRows("RES_Escaños_Escrutinio", acTable)
    End If
    Set m_Rst = Nothing
End Function

Public Function ValorASCII(ByVal sTexto As String) As Integer
    Dim m_iPos      As Integer

    For m_iPos = 1 To Len(sTexto)
        Incr ValorASCII, Asc(Mid(sTexto, m_iPos, 1))
    Next m_iPos
End Function
