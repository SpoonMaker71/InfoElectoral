VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_(Form) Resultados Escrutinio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdExportarExcel_Click()
    ' 1. Si está cargado el formulario de consulta de resultados electorales
    If modMDB.IsLoaded("(Form) Consulta de resultados electorales por mesa") Then
        Dim m_sXLSFilePath      As String
        Dim m_sTitle            As String
        Dim m_sSql              As String

        With Forms("(Form) Consulta de resultados electorales por mesa").Form
            ' 2. Bloqueamos el botón
            cmdExportarExcel.Enabled = False
            DoEvents
        
            ' 3. Construimos el título del fichero Excel
            m_sTitle = .cboProcesoElectoral.Column(3) & " - " & .cboTipoProcesoElectoral.Column(1) & " - Resultados del escrutinio - "

            ' 4. Si no se ha indicado comunidad autónoma, provincia
            '    asumimos que son los resultados a nivel nacional
            If ((.cboComunidadAutonoma.ListIndex = -1) And (.cboProvincia.ListIndex = -1)) Then
                Concat m_sTitle, "Total Nacional"

            ' 5. Si es por comunidad autónoma
            ElseIf ((.cboComunidadAutonoma.ListIndex <> -1) And (.cboProvincia.ListIndex = -1)) Or _
                   ((.cboComunidadAutonoma.ListIndex <> -1) And (.cboProvincia.ListIndex <> -1) And .cboProvincia.Locked) Then
                Concat m_sTitle, "Comunidad autónoma - " & .cboComunidadAutonoma.Column(1)

            ' 6. Si es por provincia/circunscripción
            ElseIf ((.cboComunidadAutonoma.ListIndex <> -1) And (.cboProvincia.ListIndex <> -1)) Then
                Concat m_sTitle, "Provincia - " & .cboProvincia.Column(1)
            End If
            DoEvents

            ' 7. Construimos la consulta para obtener los datos del escrutinio
            m_sSql = "SELECT"
            Concat m_sSql, " [03XXAAMM].Siglas,"
            Concat m_sSql, " [03XXAAMM].Denominacion,"
            Concat m_sSql, " Sum(RES_Escaños_Escrutinio.Escaño) AS Escaños"
            Concat m_sSql, " FROM 03XXAAMM"
            Concat m_sSql, " INNER JOIN RES_Escaños_Escrutinio ON (RES_Escaños_Escrutinio.IdCandidaturaNivelNacional = [03XXAAMM].IdCandidatura) AND ([03XXAAMM].Mes = RES_Escaños_Escrutinio.Mes) AND ([03XXAAMM].Año = RES_Escaños_Escrutinio.Año) AND ([03XXAAMM].IdTipoProcesoElectoral = RES_Escaños_Escrutinio.IdTipoProcesoElectoral)"
            Concat m_sSql, " WHERE (([03XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(.cboTipoProcesoElectoral.Value) & ")"
            Concat m_sSql, " AND ([03XXAAMM].Año = " & CSqlDbl(.cboProcesoElectoral.Column(1)) & ")"
            Concat m_sSql, " AND ([03XXAAMM].Mes = " & CSqlDbl(.cboProcesoElectoral.Column(2)) & ")"
            ' 8. Si es por comunidad autónoma
            If ((.cboComunidadAutonoma.ListIndex <> -1) And (.cboProvincia.ListIndex = -1)) Or _
               ((.cboComunidadAutonoma.ListIndex <> -1) And (.cboProvincia.ListIndex <> -1) And .cboProvincia.Locked) Then
                If (.cboComunidadAutonoma.ListIndex <> -1) Then Concat m_sSql, " AND (RES_Escaños_Escrutinio.IdINEComunidadAutonoma = " & CSqlDbl(.cboComunidadAutonoma.Value) & ")"

            ' 9. Si es por provincia/circunscripción
            ElseIf ((.cboComunidadAutonoma.ListIndex <> -1) And (.cboProvincia.ListIndex <> -1)) Then
                If (.cboProvincia.ListIndex <> -1) Then Concat m_sSql, " AND (RES_Escaños_Escrutinio.IdINEProvincia = " & CSqlDbl(.cboProvincia.Value) & ")"
            End If
            Concat m_sSql, ") GROUP BY [03XXAAMM].Siglas,"
            Concat m_sSql, " [03XXAAMM].Denominacion"
            Concat m_sSql, " ORDER BY Sum(RES_Escaños_Escrutinio.Escaño) DESC,"
            Concat m_sSql, " Sum(RES_Escaños_Escrutinio.Dhont) DESC;"

            With sfResultadosEscrutinio.Form
                ' 8. Establecemos el origen de los datos del escrutinio
                .RecordSource = m_sSql
                .Requery
            End With
        End With
        
        ' 10. Construimos la ruta del fichero MS Excel
        m_sXLSFilePath = modMDB.GetMDBPath("Informes") & Replace(Replace(m_sTitle, " / ", "_"), "/", "_") & modMDB.GetMSExcelDefaultExtension()

        ' 11. Generamos el fichero MS Excel
        modMDB.ToMSExcel m_sSql, "Resultados del Escrutinio", m_sXLSFilePath, True, True, m_sTitle, m_sTitle, "Resultados del escrutinio en fichero MS Excel"

        ' 12. Desbloqueamos el botón
        cmdExportarExcel.Enabled = True
        DoEvents
    End If
End Sub

Private Sub cmdSalir_Click()
    ' 1. Cerramos el formulario
    modMDB.CloseForm Me.Name
End Sub

Private Sub Form_Load()
    Dim m_sSql          As String

    ' 1. Si está cargado el formulario de consulta de resultados electorales
    If modMDB.IsLoaded("(Form) Consulta de resultados electorales por mesa") Then
        With Forms("(Form) Consulta de resultados electorales por mesa").Form
            ' 2. Si no se ha indicado comunidad autónoma, provincia
            '    asumimos que son los resultados a nivel nacional
            If ((.cboComunidadAutonoma.ListIndex = -1) And (.cboProvincia.ListIndex = -1)) Then
                sfResultadosEscrutinio_Etiqueta.Caption = "Total Nacional"

            ' 3. Si es por comunidad autónoma
            ElseIf ((.cboComunidadAutonoma.ListIndex <> -1) And (.cboProvincia.ListIndex = -1)) Or _
                   ((.cboComunidadAutonoma.ListIndex <> -1) And (.cboProvincia.ListIndex <> -1) And .cboProvincia.Locked) Then
                sfResultadosEscrutinio_Etiqueta.Caption = "Comunidad autónoma: " & .cboComunidadAutonoma.Column(1)

            ' 4. Si es por provincia/circunscripción
            ElseIf ((.cboComunidadAutonoma.ListIndex <> -1) And (.cboProvincia.ListIndex <> -1)) Then
                sfResultadosEscrutinio_Etiqueta.Caption = "Provincia: " & .cboProvincia.Column(1)
            End If
            DoEvents

            ' 5. Construimos la consulta para obtener los datos del escrutinio
            m_sSql = "SELECT"
            Concat m_sSql, " [03XXAAMM].Siglas,"
            Concat m_sSql, " [03XXAAMM].Denominacion,"
            Concat m_sSql, " Sum(RES_Escaños_Escrutinio.Escaño) AS Escaños"
            Concat m_sSql, " FROM 03XXAAMM"
            Concat m_sSql, " INNER JOIN RES_Escaños_Escrutinio ON (RES_Escaños_Escrutinio.IdCandidaturaNivelNacional = [03XXAAMM].IdCandidatura) AND ([03XXAAMM].Mes = RES_Escaños_Escrutinio.Mes) AND ([03XXAAMM].Año = RES_Escaños_Escrutinio.Año) AND ([03XXAAMM].IdTipoProcesoElectoral = RES_Escaños_Escrutinio.IdTipoProcesoElectoral)"
            Concat m_sSql, " WHERE (([03XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(.cboTipoProcesoElectoral.Value) & ")"
            Concat m_sSql, " AND ([03XXAAMM].Año = " & CSqlDbl(.cboProcesoElectoral.Column(1)) & ")"
            Concat m_sSql, " AND ([03XXAAMM].Mes = " & CSqlDbl(.cboProcesoElectoral.Column(2)) & ")"
            ' 6. Si es por comunidad autónoma
            If ((.cboComunidadAutonoma.ListIndex <> -1) And (.cboProvincia.ListIndex = -1)) Or _
               ((.cboComunidadAutonoma.ListIndex <> -1) And (.cboProvincia.ListIndex <> -1) And .cboProvincia.Locked) Then
                If (.cboComunidadAutonoma.ListIndex <> -1) Then Concat m_sSql, " AND (RES_Escaños_Escrutinio.IdINEComunidadAutonoma = " & CSqlDbl(.cboComunidadAutonoma.Value) & ")"

            ' 7. Si es por provincia/circunscripción
            ElseIf ((.cboComunidadAutonoma.ListIndex <> -1) And (.cboProvincia.ListIndex <> -1)) Then
                If (.cboProvincia.ListIndex <> -1) Then Concat m_sSql, " AND (RES_Escaños_Escrutinio.IdINEProvincia = " & CSqlDbl(.cboProvincia.Value) & ")"
            End If
            Concat m_sSql, ") GROUP BY [03XXAAMM].Siglas,"
            Concat m_sSql, " [03XXAAMM].Denominacion"
            Concat m_sSql, " ORDER BY Sum(RES_Escaños_Escrutinio.Escaño) DESC;"
            With sfResultadosEscrutinio.Form
                ' 8. Establecemos el origen de los datos del escrutinio
                .RecordSource = m_sSql
                .Requery

                ' 9. Mostramos el bontón de exportar a Excel si hay resultados
                cmdExportarExcel.Visible = (Not .Recordset Is Nothing)
            End With
        End With
    End If
End Sub
