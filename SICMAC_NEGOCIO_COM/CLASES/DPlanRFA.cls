VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DPlanRFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function ListaCuotas(ByVal psCodCliente As String) As ADODB.Recordset
    Dim strSQL As String
    Dim oConec As DConecta
    On Error GoTo ErrHandler
        Set oConec = New DConecta
        oConec.AbreConexion
        strSQL = "Select DIF.nCuota as nCuotaDIF,DIF.Estado as EstadoDIF,DIF.dVenc as dVencDIF,DIF.dPago as dPagoDIF,DIF.nMontoPago as nMontoPagoDIF,DIF.nMontoPagado as nMontoPagadoDIF,"
        strSQL = strSQL & " RFC.nCuota as nCuotaRFC,RFC.Estado as EstadoRFC,RFC.dVenc as dVencRFC,RFC.dPago as dPagoRFC,RFC.nMontoPago as nMontoPagoRFC,RFC.nMontoPagado as nMontoPagadoRFC,"
        strSQL = strSQL & " RFA.nCuota as nCuotaRFA,RFA.Estado as EstadoRFA,RFA.dVenc as dVencRFA,RFA.dPago as dPagoRFA,RFA.nMontoPago as nMontoPagoRFA,RFA.nMontoPagado as nMontoPagadoRFA"
        strSQL = strSQL & " From( "
        strSQL = strSQL & " Select 'DIF' as cCredito,CO.nCuota,Estado=Case CO.nColocCalendEstado"
        strSQL = strSQL & " When 0 Then 'PEND.'"
        strSQL = strSQL & " When 1 then 'CANC.'"
        strSQL = strSQL & " End,"
        strSQL = strSQL & " CO.dVenc,dPago,"
        strSQL = strSQL & " isnull(sum(CD.nMonto),0) as nMontoPago,"
        strSQL = strSQL & " isnull(Sum(CD.nMontoPagado), 0) As nMontoPagado"
        strSQL = strSQL & " From ProductoPersona PP"
        strSQL = strSQL & " Inner Join ColocacCred CC on PP.cCtaCod=CC.cCtaCod"
        strSQL = strSQL & " Inner Join ColocCalendario CO on CC.cctaCod=CO.cCtaCod"
        strSQL = strSQL & " inner join ColocCalendDet CD on CD.nCuota=CO.nCuota and CD.nnroCalen=CO.nNroCalen and"
        strSQL = strSQL & " CD.nColocCalendApl = Co.nColocCalendApl And Co.cCtaCod = CD.cCtaCod"
        strSQL = strSQL & " Where PP.cPersCod='" & psCodCliente & "' and CC.cRFA in ('DIF') and CO.nColocCalendApl=1"
        strSQL = strSQL & " Group by CO.nCuota,CO.dVenc,dPago,Case CO.nColocCalendEstado"
        strSQL = strSQL & " When 0 Then 'PEND.'"
        strSQL = strSQL & " When 1 then 'CANC.'"
        strSQL = strSQL & " End)DIF"
        strSQL = strSQL & " Inner Join(Select 'RFC' as cCredito,CO.nCuota,Estado=Case CO.nColocCalendEstado "
        strSQL = strSQL & " When 0 Then 'PEND.'"
        strSQL = strSQL & " When 1 then 'CANC.'"
        strSQL = strSQL & " End,"
        strSQL = strSQL & " CO.dVenc,dPago,"
        strSQL = strSQL & " isnull(sum(CD.nMonto),0) as nMontoPago,"
        strSQL = strSQL & " isnull(Sum(CD.nMontoPagado), 0) As nMontoPagado"
        strSQL = strSQL & " From ProductoPersona PP"
        strSQL = strSQL & " Inner Join ColocacCred CC on PP.cCtaCod=CC.cCtaCod"
        strSQL = strSQL & " Inner Join ColocCalendario CO on CC.cctaCod=CO.cCtaCod"
        strSQL = strSQL & " inner join ColocCalendDet CD on CD.nCuota=CO.nCuota and CD.nnroCalen=CO.nNroCalen and"
        strSQL = strSQL & " CD.nColocCalendApl = Co.nColocCalendApl And Co.cCtaCod = CD.cCtaCod"
        strSQL = strSQL & " Where PP.cPersCod='" & psCodCliente & "' and CC.cRFA in ('RFC') and CO.nColocCalendApl=1"
        strSQL = strSQL & " Group by CO.nCuota,CO.dVenc,dPago,Case CO.nColocCalendEstado"
        strSQL = strSQL & " When 0 Then 'PEND.'"
        strSQL = strSQL & " When 1 then 'CANC.'"
        strSQL = strSQL & " End) RFC on DIF.nCuota=RFC.nCuota"
        strSQL = strSQL & " Inner Join("
        strSQL = strSQL & " Select 'RFA' as cCredito,CO.nCuota,Estado=Case CO.nColocCalendEstado"
        strSQL = strSQL & " When 0 Then 'PEND.'"
        strSQL = strSQL & " When 1 then 'CANC.'"
        strSQL = strSQL & " End,"
        strSQL = strSQL & " CO.dVenc,dPago,"
        strSQL = strSQL & " isnull(sum(CD.nMonto),0) as nMontoPago,"
        strSQL = strSQL & " isnull(Sum(CD.nMontoPagado), 0) As nMontoPagado"
        strSQL = strSQL & " From ProductoPersona PP"
        strSQL = strSQL & " Inner Join ColocacCred CC on PP.cCtaCod=CC.cCtaCod"
        strSQL = strSQL & " Inner Join ColocCalendario CO on CC.cctaCod=CO.cCtaCod"
        strSQL = strSQL & " inner join ColocCalendDet CD on CD.nCuota=CO.nCuota and CD.nnroCalen=CO.nNroCalen and"
        strSQL = strSQL & " CD.nColocCalendApl = Co.nColocCalendApl And Co.cCtaCod = CD.cCtaCod"
        strSQL = strSQL & " Where PP.cPersCod='" & psCodCliente & "' and CC.cRFA in ('RFA') and CO.nColocCalendApl=1"
        strSQL = strSQL & " Group by CO.nCuota,CO.dVenc,dPago,Case CO.nColocCalendEstado"
        strSQL = strSQL & " When 0 Then 'PEND.'"
        strSQL = strSQL & " When 1 then 'CANC.'"
        strSQL = strSQL & " End) RFA on RFA.nCuota=RFC.nCuota"
        strSQL = strSQL & " Order by DIF.nCuota"
        
        Set ListaCuotas = oConec.CargaRecordSet(strSQL)
        oConec.CierraConexion
        Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ListaCuotas = Null
End Function

