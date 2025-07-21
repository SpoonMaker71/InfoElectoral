VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_(Form) Consulta de resultados electorales por Mesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const cteLeftPane_Left      As Long = 120
Private Const cteRightPane_Left     As Long = 9075

Private Sub cboComunidadAutonoma_Change()
    cboProvincia.Value = Null

    ' 1. Si se ha seleccionado una comunidad autónoma
    If (cboComunidadAutonoma.ListIndex <> -1) Then
        Dim m_sSql          As String

        ' 2. Consultamos la lista de provincias de la comunidad autónoma
        m_sSql = "SELECT"
        Concat m_sSql, " AUX_INE_Provincias.IdINEProvincia,"
        Concat m_sSql, " AUX_INE_Provincias.Nombre"
        Concat m_sSql, " FROM AUX_INE_Provincias"
        Concat m_sSql, " WHERE (AUX_INE_Provincias.IdINEComunidadAutonoma = " & CSqlDbl(cboComunidadAutonoma.Column(0)) & ")"
        Concat m_sSql, " ORDER BY AUX_INE_Provincias.Nombre;"
        With cboProvincia
            ' 3. Cargamos la nueva lista de provincias
            .RowSource = m_sSql
            .Requery
    
            ' 4. Si solo hay una provincia en la lista, la seleccionamos
            .Value = IIf((.ListCount = 1), CurrentDb.OpenRecordset(m_sSql).Fields(0).Value, Null)
        End With
    Else
        ' 5. Vaciamos el combo de provincias
        cboProvincia.RowSource = vbNullString
    End If

    ' 6. Bloqueamos el control si solo hay una provincia en la comunidad autónoma
    modMDB.LockControl cboProvincia, (cboProvincia.ListCount = 1)

    ' 7. Llamamos al evento de cambio de valor del combo de provincias
    cboProvincia_Change
End Sub

Private Sub cboMunicipio_Change()
    ' 1. Mostramos el checkbox de comparar actas si se ha seleccionado un municipio
    chkCompararActasMesasElectorales.Visible = (cboMunicipio.ListIndex <> -1)
End Sub

Private Sub cboProcesoElectoral_Change()
    cmdConsultar.Enabled = (cboProcesoElectoral.ListIndex <> -1)
    cmdEscrutinio.Enabled = (cboProcesoElectoral.ListIndex <> -1)

    modMDB.LockControl cboComunidadAutonoma, (cboProcesoElectoral.ListIndex = -1)
    cboComunidadAutonoma.Value = Null

    cboComunidadAutonoma_Change
End Sub

Private Sub cboProvincia_Change()
    cboMunicipio.Value = Null

    ' 1. Si se ha seleccionado una provincia
    If (cboProvincia.ListIndex <> -1) Then
        Dim m_sSql          As String

        ' 2. Construimos la consulta de municipios de la provincia
        m_sSql = "SELECT"
        Concat m_sSql, " AUX_INE_Municipios.IdINEMunicipio,"
        Concat m_sSql, " AUX_INE_Municipios.Nombre"
        Concat m_sSql, " FROM AUX_INE_Municipios"
        Concat m_sSql, " WHERE (AUX_INE_Municipios.IdINEProvincia = " & CSqlDbl(cboProvincia.Value) & ")"
        Concat m_sSql, " ORDER BY AUX_INE_Municipios.Nombre;"
        With cboMunicipio
            ' 3. Establecemos el origen de datos
            .RowSource = m_sSql
            .Requery

            ' 4. Si solo hay un municipio de ese tipo en la lista, lo seleccionamos
            If (.ListCount = 1) Then .Value = CurrentDb.OpenRecordset(m_sSql).Fields(0).Value
        End With
    Else
        ' 5. Vaciamos el combo de municipios
        cboMunicipio.RowSource = vbNullString
    End If

    ' 6. Bloqueamos el control si solo hay un municipio
    modMDB.LockControl cboMunicipio, (cboMunicipio.ListCount = 1)

    ' 7. Mostramos el checkbox de comparar actas si se ha seleccionado un municipio
    chkCompararActasMesasElectorales.Visible = (cboMunicipio.ListIndex <> -1)

    ' 8. Ocultamos el subformulario de mesas electorales
    With sfMesasElectorales
        .Visible = False
        .Form.RecordSource = vbNullString
    End With

    ' 9. ocultamos el botón de notificación del resultado de revisión de acta
    cmdExportarExcel.Visible = False
    cmdNotificarRevisionActa.Visible = False
    ActaNotificada_Etiqueta.Visible = False

    ' 10. Ocultamos el subformulario de resultados de candidaturas
    With sfResultadosCandidaturas
        .Visible = False
        .Form.RecordSource = vbNullString
    End With
End Sub

Private Sub cboTipoProcesoElectoral_Change()
    If (cboTipoProcesoElectoral.ListIndex <> -1) Then
        Dim m_sSql          As String

        ' 1. Consultamos la lista de procesos electorales del tipo indicado
        m_sSql = "SELECT"
        Concat m_sSql, " Format([02XXAAMM].FechaProcesoElectoral, 'yyyymmff') AS Id,"
        Concat m_sSql, " [02XXAAMM].Año,"
        Concat m_sSql, " [02XXAAMM].Mes,"
        Concat m_sSql, " Format([02XXAAMM].FechaProcesoElectoral, 'yyyy\ ') & Format([02XXAAMM].FechaProcesoElectoral, 'dd') & Left(UCase(Format([02XXAAMM].FechaProcesoElectoral, 'mmm')), 1) AS Titulo,"
        Concat m_sSql, " UCase(Format([02XXAAMM].FechaProcesoElectoral,'mmmm yyyy')) AS ProcesoElectoral"
        Concat m_sSql, " FROM 02XXAAMM"
        Concat m_sSql, " WHERE ([02XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(cboTipoProcesoElectoral.Column(0)) & ")"
        Concat m_sSql, " ORDER BY [02XXAAMM].Año,"
        Concat m_sSql, " [02XXAAMM].Mes;"
        With cboProcesoElectoral
            .Value = Null
            .RowSource = m_sSql
            .Requery

            ' 4. Bloqueamos el control si solo hay un proceso electoral de ese tipo
            modMDB.LockControl cboProcesoElectoral, (.ListCount = 1)

            ' 5. Si solo hay un proceso electoral de ese tipo en la lista, lo seleccionamos
            If (.ListCount = 1) Then .Value = CurrentDb.OpenRecordset(m_sSql).Fields(0).Value
        End With
    End If

    ' 6. Llamamos al evento de cambio de valor del combo de procesos elctorales
    cboProcesoElectoral_Change
End Sub

Private Sub cmdEscrutinio_Click()
    Dim m_sSql          As String
    Dim m_Rst           As Recordset

    ' 1. Comprobamos si ya está realizado el escrutinio para el resultado electoral
    m_sSql = "SELECT"
    Concat m_sSql, " Nz(Sum(RES_Escaños_Escrutinio.Escaño), 0) AS TotalEscaños"
    Concat m_sSql, " FROM RES_Escaños_Escrutinio"
    Concat m_sSql, " WHERE ((RES_Escaños_Escrutinio.IdTipoProcesoElectoral = " & CSqlDbl(cboTipoProcesoElectoral.Column(0)) & ")"
    Concat m_sSql, " AND (RES_Escaños_Escrutinio.Año = " & CSqlDbl(cboProcesoElectoral.Column(1)) & ")"
    Concat m_sSql, " AND (RES_Escaños_Escrutinio.Mes = " & CSqlDbl(cboProcesoElectoral.Column(2)) & "));"
    Set m_Rst = CurrentDb.OpenRecordset(m_sSql)
    If modMDB.IsRst(m_Rst, True) Then
        ' 2. Si no está hecho el escrutinio, lo calculamos
        If IsZr(m_Rst!TotalEscaños) Then modGlobal.ProcesarEscrutinio cboTipoProcesoElectoral.Column(0), cboProcesoElectoral.Column(1), cboProcesoElectoral.Column(2), , True

        ' 3. Abrimos el formulario para visualizar el escrutinio
        modMDB.OpenForm "(Form) Resultados Escrutinio"
    End If
End Sub

Private Sub cmdConsultar_Click()
    ' 1. Si se ha seleccionado una provincia y un municipio
    If ((cboProvincia.ListIndex <> -1) And (cboMunicipio.ListIndex <> -1) And chkCompararActasMesasElectorales.Value) Then
        Dim m_sSql          As String

        ' 2. Construimos la consulta para obtener el listado de mesas elctorales de dicho municipio
        m_sSql = "SELECT"
        Concat m_sSql, " [09XXAAMM].IdTipoProcesoElectoral,"
        Concat m_sSql, " [09XXAAMM].Año,"
        Concat m_sSql, " [09XXAAMM].Mes,"
        Concat m_sSql, " [09XXAAMM].IdINEComunidadAutonoma,"
        Concat m_sSql, " [09XXAAMM].IdINEProvincia,"
        Concat m_sSql, " [09XXAAMM].IdINEMunicipio,"
        Concat m_sSql, " [09XXAAMM].IdDistritoMunicipal,"
        Concat m_sSql, " [09XXAAMM].IdSeccion,"
        Concat m_sSql, " [09XXAAMM].IdMesa,"
        Concat m_sSql, " [09XXAAMM].CensoINE,"
        Concat m_sSql, " [09XXAAMM].CensoEscrutinio,"
        Concat m_sSql, " [09XXAAMM].VotosEnBlanco,"
        Concat m_sSql, " [09XXAAMM].VotosNulos,"
        Concat m_sSql, " [09XXAAMM].VotosCandidaturas,"
        Concat m_sSql, " [09XXAAMM].ActaNotificada,"
        Concat m_sSql, " [09XXAAMM].Correcta,"
        Concat m_sSql, " [09XXAAMM].Incidencia"
        Concat m_sSql, " FROM [09XXAAMM]"
        Concat m_sSql, " WHERE (([09XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(cboTipoProcesoElectoral.Value) & ")"
        Concat m_sSql, " AND ([09XXAAMM].Año = " & CSqlDbl(cboProcesoElectoral.Column(1)) & ")"
        Concat m_sSql, " AND ([09XXAAMM].Mes = " & CSqlDbl(cboProcesoElectoral.Column(2)) & ")"
        Concat m_sSql, " AND ([09XXAAMM].IdINEProvincia = " & CSqlDbl(cboProvincia.Value) & ")"
        Concat m_sSql, " AND ([09XXAAMM].IdINEMunicipio = " & CSqlDbl(cboMunicipio.Value) & "))"
        Concat m_sSql, " ORDER BY [09XXAAMM].IdDistritoMunicipal,"
        Concat m_sSql, " [09XXAAMM].IdSeccion,"
        Concat m_sSql, " [09XXAAMM].IdMesa;"
        With sfMesasElectorales
            ' 3. Cargamos el subformulario con el listado de mesas electorales
            With .Form
                .RecordSource = m_sSql
                .Requery
            End With

            ' 4. Mostramos el subformulario si se han obtenido mesas electorales para el municipio
            .Visible = (Not .Form.Recordset Is Nothing)

            ' 5. ocultamos el botón de notificación del resultado de revisión de acta
            cmdExportarExcel.Visible = False
            cmdNotificarRevisionActa.Visible = False
            ActaNotificada_Etiqueta.Visible = False

            ' 6. Posicionamos y ocultamos el subformulario de resultados de candidaturas
            sfResultadosCandidaturas.Left = cteRightPane_Left
            sfResultadosCandidaturas_Etiqueta.Left = cteRightPane_Left

            ' 7. Si hay datos de mesas electorales, mostramos los
            '    resultados de candidaturas de la primera de la lista
            If (Not .Form.Recordset Is Nothing) Then
                sfMesasElectorales_Etiqueta.Caption = "Mesas electorales - Municipio: " & cboMunicipio.Column(1)

                .Form.Form_Click
            End If
        End With
    Else
        ' 8. Ocultamos el subformulario de mesas electorales
        With sfMesasElectorales
            .Visible = False
            .Form.RecordSource = vbNullString
        End With

        With sfResultadosCandidaturas_Etiqueta
            ' 9. Si no se ha indicado comunidad autónoma, provincia o municipio,
            '    asumimos que son los resultados a nivel nacional
            If ((cboComunidadAutonoma.ListIndex = -1) And (cboProvincia.ListIndex = -1)) Then
                .Caption = "Resultados de candidaturas - Total Nacional"

            ' 10. Si es por comunidad autónoma
            ElseIf ((cboComunidadAutonoma.ListIndex <> -1) And (cboProvincia.ListIndex = -1) And (cboMunicipio.ListIndex = -1)) Or _
                   ((cboComunidadAutonoma.ListIndex <> -1) And (cboProvincia.ListIndex <> -1) And cboProvincia.Locked And (cboMunicipio.ListIndex = -1)) Then
                .Caption = "Resultados de candidaturas - Comunidad autónoma: " & cboComunidadAutonoma.Column(1)

            ' 11. Si es por provincia/circunscripción
            ElseIf ((cboComunidadAutonoma.ListIndex <> -1) And (cboProvincia.ListIndex <> -1) And (cboMunicipio.ListIndex = -1)) Then
                .Caption = "Resultados de candidaturas - Provincia: " & cboProvincia.Column(1)
                
            ' 12. Si es por municipio
            ElseIf ((cboComunidadAutonoma.ListIndex <> -1) And (cboProvincia.ListIndex <> -1) And (cboMunicipio.ListIndex <> -1) And Not chkCompararActasMesasElectorales.Value) Then
                .Caption = "Resultados de candidaturas - Municipio: " & cboMunicipio.Column(1)
            End If

            ' 13. Construimos y posicionamos el título del subformulario
            .Left = cteLeftPane_Left
        End With
        DoEvents

        ' 14. Si no se ha indicado comunidad autónoma, provincia o municipio,
        '     asumimos que son los resultados a nivel nacional
        If ((cboComunidadAutonoma.ListIndex = -1) And (cboProvincia.ListIndex = -1) And (cboMunicipio.ListIndex = -1)) Then
            m_sSql = "SELECT"
            Concat m_sSql, " [03XXAAMM].Siglas,"
            Concat m_sSql, " [03XXAAMM].Denominacion,"
            Concat m_sSql, " TMP.VotosObtenidos"
            Concat m_sSql, " FROM [03XXAAMM]"
            Concat m_sSql, " INNER JOIN (SELECT"
            Concat m_sSql, " [03XXAAMM].IdTipoProcesoElectoral,"
            Concat m_sSql, " [03XXAAMM].Año,"
            Concat m_sSql, " [03XXAAMM].Mes,"
            Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelNacional AS IdCandidatura,"
            Concat m_sSql, " Sum([10XXAAMM].VotosObtenidos) As VotosObtenidos"
            Concat m_sSql, " FROM [03XXAAMM]"
            Concat m_sSql, " INNER JOIN [10XXAAMM] ON ([03XXAAMM].IdTipoProcesoElectoral = [10XXAAMM].IdTipoProcesoElectoral) AND ([03XXAAMM].Año = [10XXAAMM].Año) AND ([03XXAAMM].Mes = [10XXAAMM].Mes) AND ([03XXAAMM].IdCandidatura = [10XXAAMM].IdCandidatura)"
            Concat m_sSql, " WHERE (([10XXAAMM].IdINEProvincia <> 99)"
            Concat m_sSql, " AND([10XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(cboTipoProcesoElectoral.Value) & ")"
            Concat m_sSql, " AND ([10XXAAMM].Año = " & CSqlDbl(cboProcesoElectoral.Column(1)) & ")"
            Concat m_sSql, " AND ([10XXAAMM].Mes = " & CSqlDbl(cboProcesoElectoral.Column(2)) & "))"
            Concat m_sSql, " GROUP BY [03XXAAMM].IdTipoProcesoElectoral,"
            Concat m_sSql, " [03XXAAMM].Año,"
            Concat m_sSql, " [03XXAAMM].Mes,"
            Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelNacional) AS TMP ON ([03XXAAMM].IdTipoProcesoElectoral = TMP.IdTipoProcesoElectoral) AND ([03XXAAMM].Año = TMP.Año) AND ([03XXAAMM].Mes = TMP.Mes) AND ([03XXAAMM].IdCandidatura = TMP.IdCandidatura)"
            Concat m_sSql, " ORDER BY TMP.VotosObtenidos DESC;"

        ' 15. Si es por comunidad autónoma
        ElseIf ((cboComunidadAutonoma.ListIndex <> -1) And ((cboProvincia.ListCount = 1) Or (cboProvincia.ListIndex = -1)) And (cboMunicipio.ListIndex = -1)) Then
            m_sSql = "SELECT"
            Concat m_sSql, " [03XXAAMM].Siglas,"
            Concat m_sSql, " [03XXAAMM].Denominacion,"
            Concat m_sSql, " TMP.VotosObtenidos"
            Concat m_sSql, " FROM [03XXAAMM]"
            Concat m_sSql, " INNER JOIN (SELECT"
            Concat m_sSql, " [03XXAAMM].IdTipoProcesoElectoral,"
            Concat m_sSql, " [03XXAAMM].Año,"
            Concat m_sSql, " [03XXAAMM].Mes,"
            Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelAutonomico AS IdCandidatura,"
            Concat m_sSql, " Sum([10XXAAMM].VotosObtenidos) As VotosObtenidos"
            Concat m_sSql, " FROM [03XXAAMM]"
            Concat m_sSql, " INNER JOIN [10XXAAMM] ON ([03XXAAMM].IdTipoProcesoElectoral = [10XXAAMM].IdTipoProcesoElectoral) AND ([03XXAAMM].Año = [10XXAAMM].Año) AND ([03XXAAMM].Mes = [10XXAAMM].Mes) AND ([03XXAAMM].IdCandidatura = [10XXAAMM].IdCandidatura)"
            Concat m_sSql, " WHERE (([10XXAAMM].IdINEProvincia <> 99)"
            Concat m_sSql, " AND([10XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(cboTipoProcesoElectoral.Value) & ")"
            Concat m_sSql, " AND ([10XXAAMM].Año = " & CSqlDbl(cboProcesoElectoral.Column(1)) & ")"
            Concat m_sSql, " AND ([10XXAAMM].Mes = " & CSqlDbl(cboProcesoElectoral.Column(2)) & ")"
            Concat m_sSql, " AND ([10XXAAMM].IdINEComunidadAutonoma = " & CSqlDbl(cboComunidadAutonoma.Value) & "))"
            Concat m_sSql, " GROUP BY [03XXAAMM].IdTipoProcesoElectoral,"
            Concat m_sSql, " [03XXAAMM].Año,"
            Concat m_sSql, " [03XXAAMM].Mes,"
            Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelAutonomico) AS TMP ON ([03XXAAMM].IdTipoProcesoElectoral = TMP.IdTipoProcesoElectoral) AND ([03XXAAMM].Año = TMP.Año) AND ([03XXAAMM].Mes = TMP.Mes) AND ([03XXAAMM].IdCandidatura = TMP.IdCandidatura)"
            Concat m_sSql, " ORDER BY TMP.VotosObtenidos DESC;"

        ' 16. Si es por provincia/circunscripción
        ElseIf ((cboComunidadAutonoma.ListIndex <> -1) And (cboProvincia.ListIndex <> -1) And (cboMunicipio.ListIndex = -1)) Then
            m_sSql = "SELECT"
            Concat m_sSql, " [03XXAAMM].Siglas,"
            Concat m_sSql, " [03XXAAMM].Denominacion,"
            Concat m_sSql, " TMP.VotosObtenidos"
            Concat m_sSql, " FROM [03XXAAMM]"
            Concat m_sSql, " INNER JOIN (SELECT"
            Concat m_sSql, " [03XXAAMM].IdTipoProcesoElectoral,"
            Concat m_sSql, " [03XXAAMM].Año,"
            Concat m_sSql, " [03XXAAMM].Mes,"
            Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelProvincial AS IdCandidatura,"
            Concat m_sSql, " Sum([10XXAAMM].VotosObtenidos) AS VotosObtenidos"
            Concat m_sSql, " FROM [03XXAAMM]"
            Concat m_sSql, " INNER JOIN [10XXAAMM] ON ([03XXAAMM].IdTipoProcesoElectoral = [10XXAAMM].IdTipoProcesoElectoral) AND ([03XXAAMM].Año = [10XXAAMM].Año) AND ([03XXAAMM].Mes = [10XXAAMM].Mes) AND ([03XXAAMM].IdCandidatura = [10XXAAMM].IdCandidatura)"
            Concat m_sSql, " WHERE (([10XXAAMM].IdINEProvincia <> 99)"
            Concat m_sSql, " AND([10XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(cboTipoProcesoElectoral.Value) & ")"
            Concat m_sSql, " AND ([10XXAAMM].Año = " & CSqlDbl(cboProcesoElectoral.Column(1)) & ")"
            Concat m_sSql, " AND ([10XXAAMM].Mes = " & CSqlDbl(cboProcesoElectoral.Column(2)) & ")"
            Concat m_sSql, " AND ([10XXAAMM].IdINEProvincia = " & CSqlDbl(cboProvincia.Value) & "))"
            Concat m_sSql, " GROUP BY [03XXAAMM].IdTipoProcesoElectoral,"
            Concat m_sSql, " [03XXAAMM].Año,"
            Concat m_sSql, " [03XXAAMM].Mes,"
            Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelProvincial) AS TMP ON ([03XXAAMM].IdTipoProcesoElectoral = TMP.IdTipoProcesoElectoral) AND ([03XXAAMM].Año = TMP.Año) AND ([03XXAAMM].Mes = TMP.Mes) AND ([03XXAAMM].IdCandidatura = TMP.IdCandidatura)"
            Concat m_sSql, " ORDER BY TMP.VotosObtenidos DESC;"

        ' 17. Si es por municipio
        ElseIf ((cboComunidadAutonoma.ListIndex <> -1) And (cboProvincia.ListIndex <> -1) And (cboMunicipio.ListIndex <> -1) And Not chkCompararActasMesasElectorales.Value) Then
            m_sSql = "SELECT"
            Concat m_sSql, " [03XXAAMM].Siglas,"
            Concat m_sSql, " [03XXAAMM].Denominacion,"
            Concat m_sSql, " TMP.VotosObtenidos"
            Concat m_sSql, " FROM [03XXAAMM]"
            Concat m_sSql, " INNER JOIN (SELECT"
            Concat m_sSql, " [03XXAAMM].IdTipoProcesoElectoral,"
            Concat m_sSql, " [03XXAAMM].Año,"
            Concat m_sSql, " [03XXAAMM].Mes,"
            Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelProvincial AS IdCandidatura,"
            Concat m_sSql, " Sum([10XXAAMM].VotosObtenidos) AS VotosObtenidos"
            Concat m_sSql, " FROM [03XXAAMM]"
            Concat m_sSql, " INNER JOIN [10XXAAMM] ON ([03XXAAMM].IdTipoProcesoElectoral = [10XXAAMM].IdTipoProcesoElectoral) AND ([03XXAAMM].Año = [10XXAAMM].Año) AND ([03XXAAMM].Mes = [10XXAAMM].Mes) AND ([03XXAAMM].IdCandidatura = [10XXAAMM].IdCandidatura)"
            Concat m_sSql, " WHERE (([10XXAAMM].IdINEProvincia <> 99)"
            Concat m_sSql, " AND([10XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(cboTipoProcesoElectoral.Value) & ")"
            Concat m_sSql, " AND ([10XXAAMM].Año = " & CSqlDbl(cboProcesoElectoral.Column(1)) & ")"
            Concat m_sSql, " AND ([10XXAAMM].Mes = " & CSqlDbl(cboProcesoElectoral.Column(2)) & ")"
            Concat m_sSql, " AND ([10XXAAMM].IdINEProvincia = " & CSqlDbl(cboProvincia.Value) & ")"
            Concat m_sSql, " AND ([10XXAAMM].IdINEMunicipio = " & CSqlDbl(cboMunicipio.Value) & "))"
            Concat m_sSql, " GROUP BY [03XXAAMM].IdTipoProcesoElectoral,"
            Concat m_sSql, " [03XXAAMM].Año,"
            Concat m_sSql, " [03XXAAMM].Mes,"
            Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelProvincial) AS TMP ON ([03XXAAMM].IdTipoProcesoElectoral = TMP.IdTipoProcesoElectoral) AND ([03XXAAMM].Año = TMP.Año) AND ([03XXAAMM].Mes = TMP.Mes) AND ([03XXAAMM].IdCandidatura = TMP.IdCandidatura)"
            Concat m_sSql, " ORDER BY TMP.VotosObtenidos DESC;"
        End If

        With sfResultadosCandidaturas
            With .Form
                .RecordSource = m_sSql
                .Requery
            End With

            ' 18. Posicionamos el subformulario de resultados de candidaturas
            .Left = cteLeftPane_Left

            ' 19. Ocultamos el botón de notificación del resultado de revisión de acta
            cmdExportarExcel.Visible = True
            cmdNotificarRevisionActa.Visible = False
            ActaNotificada_Etiqueta.Visible = False

            ' 20. Mostramos el subformulario si hay datos
            .Visible = (Not .Form.Recordset Is Nothing)
        End With
    End If
End Sub

Private Sub cmdExportarExcel_Click()
    Dim m_sXLSFilePath      As String
    Dim m_sTitle            As String
    Dim m_sSql              As String

    ' 1. Bloqueamos el botón
    cmdExportarExcel.Enabled = False
    DoEvents

    ' 2. Construimos el título del fichero Excel
    m_sTitle = cboProcesoElectoral.Column(3) & " - " & cboTipoProcesoElectoral.Column(1) & " - Resultados de candidaturas - "

    ' 3. Si no se ha indicado comunidad autónoma, provincia o municipio,
    '    asumimos que son los resultados a nivel nacional
    If ((cboComunidadAutonoma.ListIndex = -1) And (cboProvincia.ListIndex = -1)) Then
        Concat m_sTitle, "Total Nacional"

        m_sSql = "SELECT"
        Concat m_sSql, " [03XXAAMM].Siglas,"
        Concat m_sSql, " [03XXAAMM].Denominacion AS [Denominación],"
        Concat m_sSql, " TMP.[Votos Obtenidos]"
        Concat m_sSql, " FROM [03XXAAMM]"
        Concat m_sSql, " INNER JOIN (SELECT"
        Concat m_sSql, " [03XXAAMM].IdTipoProcesoElectoral,"
        Concat m_sSql, " [03XXAAMM].Año,"
        Concat m_sSql, " [03XXAAMM].Mes,"
        Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelNacional AS IdCandidatura,"
        Concat m_sSql, " Sum([10XXAAMM].VotosObtenidos) AS [Votos Obtenidos]"
        Concat m_sSql, " FROM [03XXAAMM]"
        Concat m_sSql, " INNER JOIN [10XXAAMM] ON ([03XXAAMM].IdTipoProcesoElectoral = [10XXAAMM].IdTipoProcesoElectoral) AND ([03XXAAMM].Año = [10XXAAMM].Año) AND ([03XXAAMM].Mes = [10XXAAMM].Mes) AND ([03XXAAMM].IdCandidatura = [10XXAAMM].IdCandidatura)"
        Concat m_sSql, " WHERE (([10XXAAMM].IdINEProvincia <> 99)"
        Concat m_sSql, " AND ([10XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(cboTipoProcesoElectoral.Value) & ")"
        Concat m_sSql, " AND ([10XXAAMM].Año = " & CSqlDbl(cboProcesoElectoral.Column(1)) & ")"
        Concat m_sSql, " AND ([10XXAAMM].Mes = " & CSqlDbl(cboProcesoElectoral.Column(2)) & "))"
        Concat m_sSql, " GROUP BY [03XXAAMM].IdTipoProcesoElectoral,"
        Concat m_sSql, " [03XXAAMM].Año,"
        Concat m_sSql, " [03XXAAMM].Mes,"
        Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelNacional) AS TMP ON ([03XXAAMM].IdTipoProcesoElectoral = TMP.IdTipoProcesoElectoral) AND ([03XXAAMM].Año = TMP.Año) AND ([03XXAAMM].Mes = TMP.Mes) AND ([03XXAAMM].IdCandidatura = TMP.IdCandidatura)"
        Concat m_sSql, " ORDER BY TMP.[Votos Obtenidos] DESC;"

    ' 4. Si es por comunidad autónoma
    ElseIf ((cboComunidadAutonoma.ListIndex <> -1) And (cboProvincia.ListIndex = -1) And (cboMunicipio.ListIndex = -1)) Or _
           ((cboComunidadAutonoma.ListIndex <> -1) And (cboProvincia.ListIndex <> -1) And cboProvincia.Locked And (cboMunicipio.ListIndex = -1)) Then
        Concat m_sTitle, "Comunidad autónoma - [" & cboComunidadAutonoma.Column(1) & "]"

        m_sSql = "SELECT"
        Concat m_sSql, " [03XXAAMM].Siglas,"
        Concat m_sSql, " [03XXAAMM].Denominacion,"
        Concat m_sSql, " TMP.[Votos Obtenidos]"
        Concat m_sSql, " FROM [03XXAAMM]"
        Concat m_sSql, " INNER JOIN (SELECT"
        Concat m_sSql, " [03XXAAMM].IdTipoProcesoElectoral,"
        Concat m_sSql, " [03XXAAMM].Año,"
        Concat m_sSql, " [03XXAAMM].Mes,"
        Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelAutonomico AS IdCandidatura,"
        Concat m_sSql, " Sum([10XXAAMM].VotosObtenidos) AS [Votos Obtenidos]"
        Concat m_sSql, " FROM [03XXAAMM]"
        Concat m_sSql, " INNER JOIN [10XXAAMM] ON ([03XXAAMM].IdTipoProcesoElectoral = [10XXAAMM].IdTipoProcesoElectoral) AND ([03XXAAMM].Año = [10XXAAMM].Año) AND ([03XXAAMM].Mes = [10XXAAMM].Mes) AND ([03XXAAMM].IdCandidatura = [10XXAAMM].IdCandidatura)"
        Concat m_sSql, " WHERE (([10XXAAMM].IdINEProvincia <> 99)"
        Concat m_sSql, " AND ([10XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(cboTipoProcesoElectoral.Value) & ")"
        Concat m_sSql, " AND ([10XXAAMM].Año = " & CSqlDbl(cboProcesoElectoral.Column(1)) & ")"
        Concat m_sSql, " AND ([10XXAAMM].Mes = " & CSqlDbl(cboProcesoElectoral.Column(2)) & ")"
        Concat m_sSql, " AND ([10XXAAMM].IdINEComunidadAutonoma = " & CSqlDbl(cboComunidadAutonoma.Value) & "))"
        Concat m_sSql, " GROUP BY [03XXAAMM].IdTipoProcesoElectoral,"
        Concat m_sSql, " [03XXAAMM].Año,"
        Concat m_sSql, " [03XXAAMM].Mes,"
        Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelAutonomico) AS TMP ON ([03XXAAMM].IdTipoProcesoElectoral = TMP.IdTipoProcesoElectoral) AND ([03XXAAMM].Año = TMP.Año) AND ([03XXAAMM].Mes = TMP.Mes) AND ([03XXAAMM].IdCandidatura = TMP.IdCandidatura)"
        Concat m_sSql, " ORDER BY TMP.[Votos Obtenidos] DESC;"

    ' 5. Si es por provincia/circunscripción
    ElseIf ((cboComunidadAutonoma.ListIndex <> -1) And (cboProvincia.ListIndex <> -1) And (cboMunicipio.ListIndex = -1)) Then
        Concat m_sTitle, "Provincia - [" & cboProvincia.Column(1) & "]"

        m_sSql = "SELECT"
        Concat m_sSql, " [03XXAAMM].Siglas,"
        Concat m_sSql, " [03XXAAMM].Denominacion,"
        Concat m_sSql, " TMP.[Votos Obtenidos]"
        Concat m_sSql, " FROM [03XXAAMM]"
        Concat m_sSql, " INNER JOIN (SELECT"
        Concat m_sSql, " [03XXAAMM].IdTipoProcesoElectoral,"
        Concat m_sSql, " [03XXAAMM].Año,"
        Concat m_sSql, " [03XXAAMM].Mes,"
        Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelProvincial AS IdCandidatura,"
        Concat m_sSql, " Sum([10XXAAMM].VotosObtenidos) AS [Votos Obtenidos]"
        Concat m_sSql, " FROM [03XXAAMM]"
        Concat m_sSql, " INNER JOIN [10XXAAMM] ON ([03XXAAMM].IdTipoProcesoElectoral = [10XXAAMM].IdTipoProcesoElectoral) AND ([03XXAAMM].Año = [10XXAAMM].Año) AND ([03XXAAMM].Mes = [10XXAAMM].Mes) AND ([03XXAAMM].IdCandidatura = [10XXAAMM].IdCandidatura)"
        Concat m_sSql, " WHERE (([10XXAAMM].IdINEProvincia <> 99)"
        Concat m_sSql, " AND ([10XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(cboTipoProcesoElectoral.Value) & ")"
        Concat m_sSql, " AND ([10XXAAMM].Año = " & CSqlDbl(cboProcesoElectoral.Column(1)) & ")"
        Concat m_sSql, " AND ([10XXAAMM].Mes = " & CSqlDbl(cboProcesoElectoral.Column(2)) & ")"
        Concat m_sSql, " AND ([10XXAAMM].IdINEProvincia = " & CSqlDbl(cboProvincia.Value) & "))"
        Concat m_sSql, " GROUP BY [03XXAAMM].IdTipoProcesoElectoral,"
        Concat m_sSql, " [03XXAAMM].Año,"
        Concat m_sSql, " [03XXAAMM].Mes,"
        Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelProvincial) AS TMP ON ([03XXAAMM].IdTipoProcesoElectoral = TMP.IdTipoProcesoElectoral) AND ([03XXAAMM].Año = TMP.Año) AND ([03XXAAMM].Mes = TMP.Mes) AND ([03XXAAMM].IdCandidatura = TMP.IdCandidatura)"
        Concat m_sSql, " ORDER BY TMP.[Votos Obtenidos] DESC;"

    ' 6. Si es por municipio
    ElseIf ((cboComunidadAutonoma.ListIndex <> -1) And (cboProvincia.ListIndex <> -1) And (cboMunicipio.ListIndex <> -1) And Not chkCompararActasMesasElectorales.Value) Then
        Concat m_sTitle, "Municipio - [" & cboMunicipio.Column(1) & "]"

        m_sSql = "SELECT"
        Concat m_sSql, " [03XXAAMM].Siglas,"
        Concat m_sSql, " [03XXAAMM].Denominacion,"
        Concat m_sSql, " TMP.[Votos Obtenidos]"
        Concat m_sSql, " FROM [03XXAAMM]"
        Concat m_sSql, " INNER JOIN (SELECT"
        Concat m_sSql, " [03XXAAMM].IdTipoProcesoElectoral,"
        Concat m_sSql, " [03XXAAMM].Año,"
        Concat m_sSql, " [03XXAAMM].Mes,"
        Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelProvincial AS IdCandidatura,"
        Concat m_sSql, " Sum([10XXAAMM].VotosObtenidos) AS [Votos Obtenidos]"
        Concat m_sSql, " FROM [03XXAAMM]"
        Concat m_sSql, " INNER JOIN [10XXAAMM] ON ([03XXAAMM].IdTipoProcesoElectoral = [10XXAAMM].IdTipoProcesoElectoral) AND ([03XXAAMM].Año = [10XXAAMM].Año) AND ([03XXAAMM].Mes = [10XXAAMM].Mes) AND ([03XXAAMM].IdCandidatura = [10XXAAMM].IdCandidatura)"
        Concat m_sSql, " WHERE (([10XXAAMM].IdINEProvincia <> 99)"
        Concat m_sSql, " AND([10XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(cboTipoProcesoElectoral.Value) & ")"
        Concat m_sSql, " AND ([10XXAAMM].Año = " & CSqlDbl(cboProcesoElectoral.Column(1)) & ")"
        Concat m_sSql, " AND ([10XXAAMM].Mes = " & CSqlDbl(cboProcesoElectoral.Column(2)) & ")"
        Concat m_sSql, " AND ([10XXAAMM].IdINEProvincia = " & CSqlDbl(cboProvincia.Value) & ")"
        Concat m_sSql, " AND ([10XXAAMM].IdINEMunicipio = " & CSqlDbl(cboMunicipio.Value) & "))"
        Concat m_sSql, " GROUP BY [03XXAAMM].IdTipoProcesoElectoral,"
        Concat m_sSql, " [03XXAAMM].Año,"
        Concat m_sSql, " [03XXAAMM].Mes,"
        Concat m_sSql, " [03XXAAMM].IdCandidaturaNivelProvincial) AS TMP ON ([03XXAAMM].IdTipoProcesoElectoral = TMP.IdTipoProcesoElectoral) AND ([03XXAAMM].Año = TMP.Año) AND ([03XXAAMM].Mes = TMP.Mes) AND ([03XXAAMM].IdCandidatura = TMP.IdCandidatura)"
        Concat m_sSql, " ORDER BY TMP.[Votos Obtenidos] DESC;"
    End If

    ' 7. Construimos la ruta del fichero MS Excel
    m_sXLSFilePath = modMDB.GetMDBPath("Informes") & Replace(Replace(m_sTitle, " / ", "_"), "/", "_") & modMDB.GetMSExcelDefaultExtension()

    ' 8. Generamos el fichero MS Excel
    modMDB.ToMSExcel m_sSql, "Resultados candidaturas", m_sXLSFilePath, True, True, m_sTitle, m_sTitle, "Resultados electorales en fichero MS Excel"

    ' 9. Desbloqueamos el botón
    cmdExportarExcel.Enabled = True
    DoEvents
End Sub

Private Sub cmdImportarFicheroZIP_Click()
    Dim m_sRutaFicheroZIP       As String

    ' 1. Limpiamos los controles del formulario
    cmdLimpiarCriterios_Click

    ' 2. Solicitamos al usuario la ruta de fichero ZIP a importar
    m_sRutaFicheroZIP = modSystem.OpenFilePicker()

    ' 3. Si existe el fichero ZIP con los datos del proceso electoral
    If modSystem.FileExists(m_sRutaFicheroZIP) Then
        Dim m_sFolderPath           As String
        Dim m_sFilePath             As String
        Dim m_sSql                  As String
        Dim m_Rst                   As Recordset

        Screen.MousePointer = 11

        ' 4. Descomprimimos el fichero ZIP
        modSystem.UnZip m_sRutaFicheroZIP, m_sFolderPath

        ' 5. Obtenemos la colección de ficheros que coincidan con la extensión DAT
        m_sFilePath = VBA.FileSystem.Dir(m_sFolderPath & "*.DAT", vbArchive)
        While Not IsZrStr(m_sFilePath)
            ' 6. Si devuelve un fichero DAT, lo renombramos a TXT
            If Not IsZrStr(m_sFilePath) Then
                modSystem.RenameFile m_sFolderPath & m_sFilePath, modSystem.ChangeFileExtension(m_sFolderPath & m_sFilePath, "TXT")
                m_sFilePath = VBA.FileSystem.Dir(m_sFolderPath & "*.DAT")
            Else
                m_sFilePath = VBA.FileSystem.Dir()
            End If
        Wend

        ' 7. Obtenemos la relación de ficheros que obligatoriamente debe tener
        '    el tipo de proceso electoral y las tablas de destino de los datos
        m_sSql = "SELECT"
        Concat m_sSql, " Format(AUX_TiposFicheros.IdTipoFichero, '00') & Format(AUX_FicherosObligatorios.IdTipoProcesoElectoral, '00') & '*.TXT' AS FileName,"
        Concat m_sSql, " AUX_TiposFicheros.Descripcion,"
        Concat m_sSql, " AUX_TiposFicheros.TableDef,"
        Concat m_sSql, " AUX_TiposFicheros.VBScript"
        Concat m_sSql, " FROM AUX_TiposFicheros"
        Concat m_sSql, " INNER JOIN AUX_FicherosObligatorios ON AUX_TiposFicheros.IdTipoFichero = AUX_FicherosObligatorios.IdTipoFichero"
        Concat m_sSql, " WHERE (AUX_FicherosObligatorios.IdTipoProcesoElectoral = " & CSqlDbl(Left(modSystem.GetFileName(m_sRutaFicheroZIP), 2)) & ");"
        Set m_Rst = CurrentDb.OpenRecordset(m_sSql)
        If modMDB.IsRst(m_Rst, True) Then
            ' 8. Establecemos la barra de progreso
            modMDB.SetProgressBar 0, m_Rst.RecordCount, "Importando resultados de proceso electoral desde fichero ZIP", m_Rst!Descripcion

            ' 9. Recorremos los distintos tipos de fichero obligatorio
            While Not m_Rst.EOF
                ' 10. Buscamos los ficheros necesarios
                m_sFilePath = VBA.FileSystem.Dir(m_sFolderPath & m_Rst!FileName, vbArchive)

                ' 11. Si existe el fichero a importar y existe la tabla de destino
                If Not IsZrStr(m_sFilePath) And modMDB.ObjectExists(m_Rst!TableDef) Then
                    ' 12. Importamos los datos del fichero a la tabla
                    DoCmd.TransferText acImportFixed, "IMP_" & m_Rst!TableDef & "_Fichero", m_Rst!TableDef, m_sFolderPath & m_sFilePath, False, , 1145

                    ' 13. Si se ha indicado un proceso Visual Basic a ejecutra, lo procesamos
                    If Not IsZrStr(m_Rst!VBScript) Then Eval m_Rst!VBScript
                End If
                m_Rst.MoveNext

                ' 14. Mostramos que fichero estamos importando en la barra de progreso
                If Not m_Rst.EOF Then
                    modMDB.IncrProgressBar 1, m_Rst!Descripcion
                Else
                    modMDB.IncrProgressBar 1
                End If
            Wend
            m_Rst.Close

            ' 15. Eliminamos el directorio con los archivos ya importados
            CreateObject("Scripting.FileSystemObject").DeleteFolder Left(m_sFolderPath, (Len(m_sFolderPath) - 1))

            ' 16. Cerramos la barra de progreso
            modMDB.CloseProgressBar

            ' 17. mostramos mensaje al usuario
            If (MsgBox("Se han importado correctamente los datos del fichero ZIP con los resultados del proceso electoral, pero se recomienda compactar la base de datos para optimizarla." & vbCrLf & vbCrLf & "AVISO: Esta tarea puede llevar un tiempo, dependiento de la cantidad de resultados de procesos electorales que contenga actualmente la base de datos." & vbCrLf & vbCrLf & "¿Desea salir y compactar la base de datos?", (vbYesNo + vbQuestion), "Fichero ZIP importado") = vbYes) Then
                ' 18. Compactamos al salir de la aplicación
                Application.SetOption "Auto Compact", True

                ' 19. Salimos de la aplicación
                cmdSalir_Click
            End If
        End If
        Set m_Rst = Nothing

        cboTipoProcesoElectoral_Change

        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdLimpiarCriterios_Click()
    cboTipoProcesoElectoral_Change
End Sub

Private Sub cmdNotificarRevisionActa_Click()
    ' 1. Abrimos el formulario de notificación del acta
    modMDB.OpenForm "(Form) Notificacion Acta"
End Sub

Private Sub cmdSalir_Click()
    ' 1. Cerramos el formulario
    modMDB.CloseForm Me.Name

    ' 2. Salimos de la aplicación
    DoCmd.Quit
End Sub

Private Sub cmdVaciarBBDD_Click()
    ' 1. Preguntamos al suaurio se está seguro de querer eliminar los procesos electorales
    If (MsgBox("Va a proceder a vaciar la base de datos, eliminando todos los resultados de procesos electorales que contenga." & vbCrLf & vbCrLf & "¿Está seguro de querer proceder a su eliminación?", (vbYesNo + vbQuestion), "Vaciar BBDD") = vbYes) Then
        ' 2. Vaciamos jerárquicamente las tablas con los datos de lo procesos electorales
        modMDB.DeleteIfHasRows "RES_Escaños_Escrutinio"
        modMDB.DeleteIfHasRows "10XXAAMM"
        modMDB.DeleteIfHasRows "09XXAAMM"
        modMDB.DeleteIfHasRows "08XXAAMM"
        modMDB.DeleteIfHasRows "07XXAAMM"
        modMDB.DeleteIfHasRows "06XXAAMM"
        modMDB.DeleteIfHasRows "05XXAAMM"
        modMDB.DeleteIfHasRows "04XXAAMM"
        modMDB.DeleteIfHasRows "03XXAAMM"
        modMDB.DeleteIfHasRows "02XXAAMM"
        modMDB.DeleteIfHasRows "01XXAAMM"

        ' 3. Limpiamos los controles del formulario
        cmdLimpiarCriterios_Click

        ' 4. Preguntamos al usuario si desea salir de la aplicación para proceder a compactarla
        If (MsgBox("Se ha vaciado la base de datos, pero se recomienda proceder a su compactación para optimizarla." & vbCrLf & vbCrLf & "¿Desea salir y compactar la base de datos?", (vbYesNo + vbQuestion), "Compactar BBDD") = vbYes) Then
            ' 5. Compactamos al salir de la aplicación
            Application.SetOption "Auto Compact", True

            ' 6. Salimos de la aplicación
            cmdSalir_Click
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim m_sSql          As String

    ' 1. Cargamos el combo de tipos de procesos electorales
    m_sSql = "SELECT"
    Concat m_sSql, " AUX_TiposProcesosElectorales.IdTipoProcesoElectoral,"
    Concat m_sSql, " AUX_TiposProcesosElectorales.Abreviatura,"
    Concat m_sSql, " AUX_TiposProcesosElectorales.Descripcion"
    Concat m_sSql, " FROM AUX_TiposProcesosElectorales"
    Concat m_sSql, " ORDER BY AUX_TiposProcesosElectorales.[Descripcion];"
    With cboTipoProcesoElectoral
        .RowSource = m_sSql
        .Requery
    End With

    ' 2. Cargamos el combo de comunidades autónomas
    m_sSql = "SELECT"
    Concat m_sSql, " AUX_INE_ComunidadesAutonomas.IdINEComunidadAutonoma,"
    Concat m_sSql, " AUX_INE_ComunidadesAutonomas.Nombre"
    Concat m_sSql, " FROM AUX_INE_ComunidadesAutonomas"
    Concat m_sSql, " ORDER BY AUX_INE_ComunidadesAutonomas.Nombre;"
    With cboComunidadAutonoma
        .RowSource = m_sSql
        .Requery
    End With

    ' 3. Estableceos el valor por defecto del combo de tipos de procesos electorales y lo bloqueamos
    '
    '    NOTA: Debido a que no se ha podido comprobar el formulario para otros tipos de procesos electorales,
    '          bloqueamos el combo de tipos de procesos electorales, para que tenga indicado permanentemente
    '          el tipo 'Congreso de los Diputados'. Para activar otros tipos de procesos electorales, se debe
    '          eliminar o comentar este código.
    '
    cboTipoProcesoElectoral.Value = 2
    modMDB.LockControl cboTipoProcesoElectoral, True
    '---------------------------------------------------------------------------------------------------------

    ' 4. Llamamos al evento de cambio de valor del combo de tipos de procesos elctorales
    cboTipoProcesoElectoral_Change
End Sub

