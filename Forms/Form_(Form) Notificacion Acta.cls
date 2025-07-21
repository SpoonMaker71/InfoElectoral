VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_(Form) Notificacion Acta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdAceptar_Click()
    ' 1. Si está cargado el formulario padre
    If modMDB.IsLoaded("(Form) Consulta de resultados electorales por mesa") Then
        Dim m_sUrl              As String

        With Forms("(Form) Consulta de resultados electorales por mesa").Form.sfMesasElectorales.Form
                '  2. Actualizamos la mesa electoral con los datos de la notificacion
                '     de si es correcta o no, y la incidencia si la hubiera
                .ActaNotificada.Value = True
                .Correcta.Value = (mrcActaCorrecta.Value = chkSi.OptionValue)
                If (mrcActaCorrecta.Value = chkNo.OptionValue) Then .Incidencia.Value = NotificacionIncidencia.Value
                .Refresh

                ' 3. Si se notificó el acta, mostramos el mensaje en el formulario
                Forms("(Form) Consulta de resultados electorales por mesa").Form.ActaNotificada_Etiqueta.Visible = .ActaNotificada.Value
                If .ActaNotificada.Value Then
                    ' 4. Si no se notificó el acta, ocultamos el botón de notificar
                    Forms("(Form) Consulta de resultados electorales por mesa").Form.sfMesasElectorales.SetFocus
                    Forms("(Form) Consulta de resultados electorales por mesa").Form.cmdNotificarRevisionActa.Visible = False

                    ' 5. Mostramos es estado de la notificación de la mesa
                    Forms("(Form) Consulta de resultados electorales por mesa").Form.ActaNotificada_Etiqueta.ForeColor = IIf(.Correcta.Value, RGB(34, 177, 76), RGB(250, 0, 0))
                    Forms("(Form) Consulta de resultados electorales por mesa").Form.ActaNotificada_Etiqueta.Caption = "Notificada como " & IIf(.Correcta.Value, "correcta", "errónea")
                Else
                    ' 6. Si se notificó el acta, mostramos el botón de notificar
                    Forms("(Form) Consulta de resultados electorales por mesa").Form.cmdNotificarRevisionActa.Visible = True
                End If

                ' 27. Abrimos la página web de Ignacio Baixauli con el formulario de notificación
                '    de actas con entradas de datos ya predeterminadas
                m_sUrl = "https://docs.google.com/forms/d/e/1FAIpQLScRvRJmdbv1c2uYOmCNmFgos5IvD6tnJMkzhL9nNjrV6CkY8w/viewform?usp=pp_url"
                Concat m_sUrl, "&entry.1042724278=" & Forms("(Form) Consulta de resultados electorales por mesa").Form.cboProvincia.Column(1)
                Concat m_sUrl, "&entry.984316794=" & Forms("(Form) Consulta de resultados electorales por mesa").Form.cboMunicipio.Column(1)
                Concat m_sUrl, "&entry.1443333222=" & .IdDistritoMunicipal.Value
                Concat m_sUrl, "&entry.1064490468=" & .IdSeccion.Value
                Concat m_sUrl, "&entry.434886749=" & .IdMesa.Value
                Concat m_sUrl, "&entry.2028588206=" & IIf((mrcActaCorrecta.Value = chkSi.OptionValue), "Sí", "No")
                If ((mrcActaCorrecta.Value = chkNo.OptionValue) And Not IsZrStr(NotificacionIncidencia.Value)) Then Concat m_sUrl, "&entry.1233851010=" & NotificacionIncidencia.Value
                Concat m_sUrl, "#i29"
                modSystem.OpenURL m_sUrl

                ' 8. Cerramos el formulario
                modMDB.CloseForm Me.Name
        End With
    End If
End Sub

Private Sub cmdCancelar_Click()
    ' 1. Cerramos el formulario
    modMDB.CloseForm Me.Name
End Sub

Private Sub mrcActaCorrecta_Click()
    Dim m_bVisible          As Boolean

    ' 1. Determinamos si se ha indicado que el acta no es correcta
    m_bVisible = (mrcActaCorrecta.Value = chkNo.OptionValue)

    ' 2. Mostramos/ocultamos los campos para detallar la incidencia
    TituloIncidenciaActa_Etiqueta.Visible = m_bVisible
    NotificacionIncidencia_Etiqueta.Visible = m_bVisible
    NotificacionIncidencia.Visible = m_bVisible

    ' 3. Activamos el botón de Aceptar
    cmdAceptar.Enabled = ((mrcActaCorrecta.Value = chkSi.OptionValue) Or ((mrcActaCorrecta.Value = chkNo.OptionValue) And Not IsZrStr(NotificacionIncidencia.Value)))
End Sub

Private Sub NotificacionIncidencia_KeyPress(KeyAscii As Integer)
    ' 1. Ponemos el texto en mayúsculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    ' 2. Activamos el botón de Aceptar
    cmdAceptar.Enabled = ((mrcActaCorrecta.Value = chkSi.OptionValue) Or ((mrcActaCorrecta.Value = chkNo.OptionValue) And Not IsZrStr(NotificacionIncidencia.Text)))
End Sub

