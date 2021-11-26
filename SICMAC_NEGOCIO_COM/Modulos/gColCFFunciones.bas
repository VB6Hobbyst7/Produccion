Attribute VB_Name = "gColCFFunciones"
Option Explicit

Public Function fgIniciaAxCuentaCF() As String
    fgIniciaAxCuentaCF = gsCodCMAC & gsCodAge
End Function

Public Function fgEstadoColCFDesc(ByVal pnEstado As Integer) As String
Dim lsDesc As String
    Select Case pnEstado
        Case gColocEstSolic
            lsDesc = "Solicitado"
        Case gColocEstSug
            lsDesc = "Sugerido"
        Case gColocEstAprob
            lsDesc = "Aprobado"
        Case gColocEstVigNorm
            lsDesc = "Emitida"
        Case gColocEstCancelado
            lsDesc = "Cancelada"
        Case gColocEstHonrada
            lsDesc = "Honrada"
        Case gColocEstDevuelta
            lsDesc = "Devuelta"
        Case gColocEstRech
            lsDesc = "Rechazada"
        Case gColocEstRetirado
            lsDesc = "Retirada"
    End Select
    fgEstadoColCFDesc = lsDesc
End Function
'Public Sub ImprimeConsolidadoCred(ByVal psCtaCod As String, ByVal psOrigen As String, ByVal pdFecha As Date)
'Dim Rs As ADODB.Recordset
'Dim oCon As DConecta
'Dim sql  As String
'Set oCon = New DConecta
'Set Rs = New ADODB.Recordset
'
'MsgBox "Referencia al acceso a datos"
'oCon.AbreConexion
'sql = "SELECT GETDATE() "
'Set Rs = oCon.CargaRecordSet(sql)
'oCon.CierraConexion
'Set oCon = Nothing
'With DRSugerencia
'      Set .DataSource = Rs
'      .DataMember = ""
'      .inicio psCtaCod, psOrigen, pdFecha
'      .Refresh
'      .Show vbModal
' End With
'
''DRSugerencia.inicio psCtaCod, psOrigen, pdFecha
''DRSugerencia.Show vbModal
'End Sub

'WIOR 20140130 *****************************************************************************
Public Function obtenerFechaFinMes(ByVal pnMes As Integer, ByVal pnAnio As Integer) As Date
    Dim sFecha  As Date
    sFecha = CDate("01/" & Format(pnMes, "00") & "/" & pnAnio)
    sFecha = DateAdd("m", 1, sFecha)
    sFecha = sFecha - 1
    obtenerFechaFinMes = sFecha
End Function
'WIOR FIN ***********************************************************************************

