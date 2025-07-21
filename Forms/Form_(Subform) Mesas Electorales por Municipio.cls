VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_(Subform) Mesas Electorales por Municipio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Sub Form_Click()
    ' 1. Si está cargado el formulario padre
    If modMDB.IsLoaded("(Form) Consulta de resultados electorales por mesa") Then
        Dim m_sSql          As String

        With Forms("(Form) Consulta de resultados electorales por mesa").Form
            ' 2. Construimos el título del subformulario
            .sfResultadosCandidaturas_Etiqueta.Caption = "Resultados de candidaturas - Distrito: " & IdDistritoMunicipal.Value & " - Sección: " & IdSeccion.Value & " - Mesa: " & IdMesa.Value

            ' 3. Si se notificó el acta, mostramos el mensaje en el formulario
            With .ActaNotificada_Etiqueta
                .Visible = ActaNotificada.Value

                If ActaNotificada.Value Then
                    .ForeColor = IIf(Correcta.Value, RGB(34, 177, 76), RGB(250, 0, 0))
                    .Caption = "Notificada como " & IIf(Correcta.Value, "correcta", "errónea")
                End If
            End With

            ' 4. Si no se notificó el acta, mostramos el botón de notificar
            .cmdNotificarRevisionActa.Visible = (Not ActaNotificada.Value)
            DoEvents

            ' 5. Construimos la consulta de los resultados por candidatura
            With .Form
                m_sSql = "SELECT"
                Concat m_sSql, " [03XXAAMM].Siglas,"
                Concat m_sSql, " [03XXAAMM].Denominacion,"
                Concat m_sSql, " [10XXAAMM].VotosObtenidos"
                Concat m_sSql, " FROM [03XXAAMM]"
                Concat m_sSql, " INNER JOIN [10XXAAMM] ON ([03XXAAMM].IdTipoProcesoElectoral = [10XXAAMM].IdTipoProcesoElectoral) AND ([03XXAAMM].Mes = [10XXAAMM].Mes) AND ([03XXAAMM].Año = [10XXAAMM].Año) AND ([03XXAAMM].IdCandidatura = [10XXAAMM].IdCandidatura)"
                Concat m_sSql, " WHERE (([10XXAAMM].IdTipoProcesoElectoral = " & CSqlDbl(.cboTipoProcesoElectoral.Value) & ")"
                Concat m_sSql, " AND ([10XXAAMM].Año = " & CSqlDbl(.cboProcesoElectoral.Column(1)) & ")"
                Concat m_sSql, " AND ([10XXAAMM].Mes = " & CSqlDbl(.cboProcesoElectoral.Column(2)) & ")"
                Concat m_sSql, " AND ([10XXAAMM].IdINEProvincia = " & CSqlDbl(.cboProvincia.Value) & ")"
                Concat m_sSql, " AND ([10XXAAMM].IdINEMunicipio = " & CSqlDbl(.cboMunicipio.Value) & ")"
                Concat m_sSql, " AND ([10XXAAMM].IdDistritoMunicipal = " & CSqlDbl(IdDistritoMunicipal.Value) & ")"
                Concat m_sSql, " AND ([10XXAAMM].IdSeccion = " & CSqlTxt(IdSeccion.Value) & ")"
                Concat m_sSql, " AND ([10XXAAMM].IdMesa = " & CSqlTxt(IdMesa.Value) & "))"
                Concat m_sSql, " ORDER BY [10XXAAMM].VotosObtenidos DESC;"
                With .sfResultadosCandidaturas
                    With .Form
                        ' 6. Consultamos los datos
                        .RecordSource = m_sSql
                        .Requery
                    End With

                    ' 7. Mostramos el subformulario si hay datos
                    .Visible = (Not .Form.Recordset Is Nothing)
                End With
            End With
        End With
    End If
End Sub
