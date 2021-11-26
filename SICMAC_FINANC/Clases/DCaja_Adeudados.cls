VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCaja_Adeudados"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A89DAEE0013"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Dim vsConexion As String
Dim vsServerAdmin As String
Dim vsServerComunes As String
Dim vsServerImagenes As Variant
Dim vsServerNegocio As String
Dim vsServerPersonas As String
    
Dim sql As String
Dim oConect As DConecta
Dim rs As ADODB.Recordset

Public Function CargaIndiceVAC(pdFecha As Date) As Double
'    On Error GoTo CargaIndiceVACErr
'    CargaIndiceVAC = 0
'    sql = "Select * from IndiceVac iv Where dIndiceVac = (Select Max(dIndiceVac) FROM IndiceVac iv1 where iv1.dIndiceVac <= '" & Format(pdFecha, gsFormatoFecha) & "') "
'    Set rs = oConect.CargaRecordSet(sql)
'    If Not rs.EOF And Not rs.BOF Then
'        CargaIndiceVAC = Format(rs!nIndiceVac, "#,###.00####")
'    End If
'    RSClose rs
'
'    Exit Function
'CargaIndiceVACErr:
'    Call RaiseError(MyUnhandledError, "DCaja_Adeudados:CargaIndiceVAC Method")

    On Error GoTo CargaIndiceVACErr
    CargaIndiceVAC = 0
    sql = "Select dIndiceVac, ISNULL(nIndiceVac,0) as nIndiceVac from IndiceVac iv Where dIndiceVac = (Select Max(dIndiceVac) FROM IndiceVac iv1 where iv1.dIndiceVac = '" & Format(pdFecha, gsFormatoFecha) & "') "
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        CargaIndiceVAC = Format(rs!nIndiceVac, "#,###.00####")
    End If
    RSClose rs
    
    Exit Function
CargaIndiceVACErr:
    Call RaiseError(MyUnhandledError, "DCaja_Adeudados:CargaIndiceVAC Method")

End Function

'Carga el calendario de pagos de los adeudados de las IF.
Public Function CargaCalendarios(ByVal psCodCtaIf As String) As ADODB.Recordset
    On Error GoTo CargaCalendariosErr
    sql = "SELECT cObjetoCod, nNroCuota ,  dPlanPago  dPagoFecha, nPlanCapital as  nCapital, " _
       & " nPlanInteres nInteres, cPlanTipoCuota cPagoTipo  FROM CtaBancoCal  WHERE COBJETOCOD='" & psCodCtaIf & "' ORDER BY cPagoTipo DESC, dPagoFecha "
    
    Set rs = oConect.CargaRecordSet(sql)
    Set CargaCalendarios = rs
    Exit Function
CargaCalendariosErr:
    Call RaiseError(MyUnhandledError, "DCaja_Adeudados:CargaCalendarios Function")
End Function

Public Function GetAdeudadosProvision(psOpeCod As String, pdFecha As Date, pnMoneda As Moneda, Optional psBuscar As String = "", Optional psEstado As CGEstadoCtaIF = gEstadoCtaIFRegistrada) As ADODB.Recordset
    On Error GoTo GetAdeudadosProvisionErr
Dim sSql As String
''    sql = "SELECT ci.cPersCod, ci.cIFTpo, ci.cCtaIFCod, p.cPersNombre, ci.cCtaIFDesc, Adeud.nSaldoCap, Adeud.dCuotaUltPago, Adeud.cMonedaPago, Cal.nNroCuota, Cal.nCapital nCapitalCuota, Cal.nInteresPagado, ISNULL(cCodLinCred,'') cCodLinCred, " _
''        & "       Int.nInteres, Int.nPeriodo, ci.nCtaIFPlazo , Adeud.nCtaIFCuotas, DATEDIFF(day,ISNULL(Adeud.dCuotaUltPago,ci.dCtaIFAper),'" & Format(pdFecha, "mm/dd/yyyy") & "') AS nDiasUltPago, cal.dVencimiento, Adeud.nTramoConcesion, nComision, " _
''        & "       nSaldoConcesion = ISNULL((SELECT SUM(nCapital) FROM CtaIFCalendario cc " _
''        & "          WHERE cEstado = '0' and cTpoCuota = '6' and Cal.cPersCod = cc.cPersCod AND Cal.cIFTpo = cc.cIFTpo AND Cal.cCtaIFCod = cc.cCtaIFCod ),0), Adeud.nSaldoCapLP, convert(int, bmalpg) bmalpg " _
''        & "FROM CtaIF ci " _
''        & " INNER JOIN CtaIFAdeudados Adeud ON ci.cPersCod = Adeud.cPersCod AND ci.cIFTpo = Adeud.cIFTpo AND ci.cCtaIFCod = Adeud.cCtaIFCod " _
''        & " INNER JOIN (SELECT cPersCod, cIFTpo, cCtaIFCod, nCtaIFIntValor nInteres, nCtaIFIntPeriodo nPeriodo FROM CtaIFInteres cii " _
''        & "             WHERE dCtaIFIntRegistro = (SELECT max(dCtaIFIntRegistro) FROM CtaIFInteres WHERE cPersCod = cii.cPersCod and cIFTpo = cii.cIFTpo and cCtaIFCod = cii.cCtaIFCod and dCtaIFIntRegistro <= '" & Format(pdFecha, gsFormatoFecha) & "' ) " _
''        & "            ) Int ON ci.cPersCod = Int.cPersCod AND ci.cIFTpo = Int.cIFTpo AND ci.cCtaIFCod = Int.cCtaIFCod " _
''        & " INNER JOIN CtaIFCalendario Cal ON Adeud.cPersCod = Cal.cPersCod AND Adeud.cIFTpo = Cal.cIFTpo AND Adeud.cCtaIFCod = Cal.cCtaIFCod " _
''        & " INNER JOIN (SELECT cPersCod, cIFTpo, cCtaIFCod, MIN(nNroCuota) Cuota " _
''        & "             FROM CtaIFCalendario " _
''        & "             WHERE cEstado = '" & gTpoEstCuotaAdeudPend & "' and cTpoCuota <> '6' " _
''        & "             GROUP BY cPersCod, cIFTpo, cCtaIFCod " _
''        & "            ) Cal2 ON Cal.cPersCod = Cal2.cPersCod AND Cal.cIFTpo = Cal2.cIFTpo AND Cal.cCtaIFCod = Cal2.cCtaIFCod AND Cal.nNroCuota = Cal2.Cuota " _
''        & " INNER JOIN Persona p ON ci.cPersCod = p.cPersCod " _
''        & "WHERE ci.cCtaIFCod LIKE '__" & pnMoneda & "%' and ci.cCtaIFEstado = '" & gEstadoCtaIFActiva & "' and cTpoCuota <> '6' " & psBuscar & " ORDER BY p.cPersNombre, ci.cCtaIFCod "

'    sSQL = "Select  CIF.cPersCod, CIF.cIFTpo, CIF.cCtaIFCod, P.cPersNombre, CIF.cCtaIFDesc, CIA.nSaldoCap, CIA.nSaldoCapLP, CIA.dCuotaUltPago, "
'    sSQL = sSQL & " CIA.cMonedaPago, CIC.cTpoCuota, CIC.nNroCuota, CIC.nCapital, ISNULL(CIA.cCodLinCred, '') as cCodLinCred, ISNULL(CLC.cDescripcion, '') as cDesLinCred, "
'    sSQL = sSQL & " CII.nCtaIFIntPeriodo, CII.nCtaIFIntValor, CIA.nCtaIFCuotas, "
'    sSQL = sSQL & " DATEDIFF(day,ISNULL(CIA.dCuotaUltPago,CIF.dCtaIFAper),'" & Format(pdFecha, "MM/dd/YYYY") & "') AS nDiasUltPago, CIA.dCuotaUltPago, Cic.nInteresPagado, "
'    sSQL = sSQL & " CIC.dVencimiento ,CIA.nTramoConcesion,bmalpg ,cic.nCapital nCapitalCuota ,cii.nctaifintperiodo nperiodo,cic.ninteres ninteres,cic.ncomision,cmonedapago,cic.cestado cestadocuota,nInteresPagado_Real nInteresProvisionadoReal ,(cic.ncapital+cic.ninteres+cic.ncomision) nTotalCuota "
'    sSQL = sSQL & " From CtaIF CIF "
'    sSQL = sSQL & "   Inner Join CtaIFAdeudados CIA "
'    sSQL = sSQL & "     On CIF.cPersCod=CIA.cPersCod And CIF.cIFTpo=CIA.cIFTpo And CIF.cCtaIFCod=CIA.cCtaIFCod "
'    sSQL = sSQL & " Inner Join (    SELECT cPersCod, cIFTpo, cCtaIFCod, MIN(nNroCuota) nNroCuota "
'    sSQL = sSQL & "                 From CtaIFCalendario "
'    sSQL = sSQL & "                 WHERE (cEstado = '0' or cEstado='2') and cTpoCuota <> '6' "
'    sSQL = sSQL & "                 GROUP BY cPersCod, cIFTpo, cCtaIFCod "
'    sSQL = sSQL & "             ) CalRef ON CIF.cPersCod=CalRef.cPersCod And CIF.cIFTpo=CalRef.cIFTpo And CIF.cCtaIFCod=CalRef.cCtaIFCod "
'    sSQL = sSQL & " Inner Join CtaIFCalendario CIC "
'    sSQL = sSQL & "     ON CIF.cPersCod=CIC.cPersCod And CIF.cIFTpo=CIC.cIFTpo And CIF.cCtaIFCod=CIC.cCtaIFCod "
'    sSQL = sSQL & "         AND CIC.nNroCuota=CalRef.nNroCuota "
'    sSQL = sSQL & " Inner Join CtaIFInteres CII "
'    sSQL = sSQL & "     On CIF.cPersCod=CII.cPersCod And CIF.cIFTpo=CII.cIFTpo And CIF.cCtaIFCod=CII.cCtaIFCod "
'    sSQL = sSQL & " Inner JOIN Persona P "
'    sSQL = sSQL & "     ON CIF.cPersCod = P.cPersCod "
'    sSQL = sSQL & " Left Join ColocLineaCredito CLC "
'    sSQL = sSQL & "     ON CIA.cCodLinCred=CLC.cLineaCred "
'    sSQL = sSQL & " Where   CIF.cCtaIFCod Like '__" & pnMoneda & "%' and CIF.cCtaIFEstado = '" & gEstadoCtaIFActiva & "' "
'
'    'Solo se provisionaran las que no estan pagadas
'    sSQL = sSQL & " And (CIC.cEstado = '0' or CIC.cEstado='2') and cTpoCuota <> '6' "
'    sSQL = sSQL & " "
'
'    sSQL = sSQL & " " & psBuscar & " ORDER BY P.cPersNombre, CIF.cCtaIFCod "


    sSql = "Select  CIF.cPersCod, CIF.cIFTpo, CIF.cCtaIFCod, P.cPersNombre, CIF.cCtaIFDesc, CIA.nSaldoCap, CIA.nSaldoCapLP, CIA.dCuotaUltPago, "
    sSql = sSql & " CIA.cMonedaPago, CIC.cTpoCuota, CIC.nNroCuota, CIC.nCapital, ISNULL(CIA.cCodLinCred, '') as cCodLinCred, ISNULL(CLC.cDescripcion, '') as cDesLinCred, "
    sSql = sSql & " CII.nCtaIFIntPeriodo, CII.nCtaIFIntValor, CIA.nCtaIFCuotas, "
    sSql = sSql & " DATEDIFF(day,ISNULL(CIA.dCuotaUltModSaldos,CIF.dCtaIFAper),'" & Format(pdFecha, "MM/dd/YYYY") & "') AS nDiasUltPago, CIA.dCuotaUltModSaldos, "
    sSql = sSql & " CIC.dVencimiento , CIA.bConcesional, CIA.nTramoConcesional,CIA.nComisionMonto,CIA.nCtaIFCuotas,CIF.nCtaIFPlazo,isnull(CIC.nSaldoMes,0) nSaldoMes"
    'ALPA20130617***********************************************
    sSql = sSql & " ,nSaldoCapConce=isnull(CIA.nSaldoCapConce,'0') "
    '***********************************************************
    sSql = sSql & " From CtaIF CIF "
    sSql = sSql & "   Inner Join CtaIFAdeudados CIA "
    sSql = sSql & "     On CIF.cPersCod=CIA.cPersCod And CIF.cIFTpo=CIA.cIFTpo And CIF.cCtaIFCod=CIA.cCtaIFCod "
    sSql = sSql & " Inner Join (    SELECT cPersCod, cIFTpo, cCtaIFCod, MIN(nNroCuota) nNroCuota "
    sSql = sSql & "                 From CtaIFCalendario "
    sSql = sSql & "                 WHERE bVigente=1 And (cEstado = '0' or cEstado='2') and cTpoCuota <> '6' "
    sSql = sSql & "                 GROUP BY cPersCod, cIFTpo, cCtaIFCod "
    sSql = sSql & "             ) CalRef ON CIF.cPersCod=CalRef.cPersCod And CIF.cIFTpo=CalRef.cIFTpo And CIF.cCtaIFCod=CalRef.cCtaIFCod "
    sSql = sSql & " Inner Join CtaIFCalendario CIC "
    sSql = sSql & "     ON CIF.cPersCod=CIC.cPersCod And CIF.cIFTpo=CIC.cIFTpo And CIF.cCtaIFCod=CIC.cCtaIFCod "
    sSql = sSql & "         AND CIC.nNroCuota=CalRef.nNroCuota "
    sSql = sSql & " Inner Join CtaIFInteres CII "
    sSql = sSql & "     On CIF.cPersCod=CII.cPersCod And CIF.cIFTpo=CII.cIFTpo And CIF.cCtaIFCod=CII.cCtaIFCod "
    sSql = sSql & " Inner JOIN Persona P "
    sSql = sSql & "     ON CIF.cPersCod = P.cPersCod "
    sSql = sSql & " Left Join ColocLineaCredito CLC "
    sSql = sSql & "     ON CIA.cCodLinCred=CLC.cLineaCred "
    sSql = sSql & " Where   CIF.cCtaIFCod Like '__" & pnMoneda & "%' and CIF.cCtaIFEstado = '" & gEstadoCtaIFActiva & "' "
    
    'Solo se provisionaran las que no estan pagadas
    sSql = sSql & " And CIC.bVigente=1 And (CIC.cEstado = '0' or CIC.cEstado='2') and cTpoCuota <> '6' "
    sSql = sSql & " And CII.nVigente=1 "
    sSql = sSql & " " & psBuscar & " ORDER BY P.cPersNombre, CIF.cCtaIFCod "

    
'AND SUBSTRING(M.cMovnro,1,8 ) >='" & Format(pdDesde, "yyyymmdd") & "' AND substring(M.cMovNro, 1,8) <='" & Format(pdHasta, "yyyymmdd") & "'

    Set GetAdeudadosProvision = oConect.CargaRecordSet(sSql)

    Exit Function
GetAdeudadosProvisionErr:
    Call RaiseError(MyUnhandledError, "DCaja_Adeudados:GetAdeudadosProvision Function")

End Function

Public Function GetCuotaMalPagador(psPersCod As String, psIFTpo As String, psCtaIFCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    
    sSql = "Select cperscod,ciftpo,cctaifcod,nCapital,nInteres,ncomision,nnrocuota from ctaifcalendario"
    sSql = sSql & " where cperscod='" & psPersCod & "' and ciftpo='" & psIFTpo & "' and cctaifcod='" & psCtaIFCod & "'  and nnrocuota in (  select min(nnrocuota)"
    sSql = sSql & "     From ctaifcalendario"
    sSql = sSql & "     where cperscod='" & psPersCod & "' and ciftpo='" & psIFTpo & "' and cctaifcod='" & psCtaIFCod & "'"
    sSql = sSql & "     and ctpocuota=6 and cestado in (0,2)"
    sSql = sSql & "           ) and ctpocuota=6"
'    ocon.AbreConexion
    Set rs = oConect.CargaRecordSet(sSql)
    Set GetCuotaMalPagador = rs
'    ocon.CierraConexion
End Function


Public Function CalculaInteres(pnDias As Long, pnPeriodo As Long, pnTasaInt As Currency, pnCapital As Currency) As Currency
    Dim lnFactor As Double
    Dim lnBaseInt As Double
    If pnPeriodo > 0 Then
        lnFactor = pnDias / pnPeriodo
        lnBaseInt = ((1 + (pnTasaInt / 100)) ^ lnFactor) - 1
        CalculaInteres = lnBaseInt * pnCapital
    Else
        CalculaInteres = 0
    End If
End Function


Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni

Set oIni = New ClasIni
vsServerAdmin = oIni.BaseAdministracion
vsServerComunes = oIni.BaseComunes
vsServerImagenes = oIni.BaseImagenes
vsServerNegocio = oIni.BaseNegocio
vsServerPersonas = oIni.BasePersonas
    
Set oConect = New DConecta
oConect.AbreConexion

End Sub

Private Sub Class_Terminate()
oConect.CierraConexion
Set oConect = Nothing
End Sub

'ALPA 20110928******************************
Public Function RegistrarFONDEMICAB(pdFecCierre As Date, pdFechaIni As Date, pdFechaFin As Date, psPaquete As String, psMoneda As String, pnMontoDesem As Currency, pnMontoSaldo As Currency, psTipoInfo As String)
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    
    sSql = "exec stp_ins_FONDEMICAB '" & Format(pdFecCierre, "YYYY/MM/DD") & "','" & Format(pdFechaIni, "YYYY/MM/DD") & "','" & Format(pdFechaFin, "YYYY/MM/DD") & "','" & psPaquete & "','" & psMoneda & "'," & pnMontoDesem & "," & pnMontoSaldo & ",'" & psTipoInfo & "'"
    oCon.AbreConexion
    Set rs = oConect.CargaRecordSet(sSql)
    
    oCon.CierraConexion
End Function

Public Function RegistrarFONDEMIDET(pdFecCierre As Date, pdFechaIni As Date, pdFechaFin As Date, psPaquete As String, psMoneda As String, psTipoInfo As String)
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    
    sSql = "exec stp_ins_FONDEMIDET '" & Format(pdFecCierre, "YYYY/MM/DD") & "','" & Format(pdFechaIni, "YYYY/MM/DD") & "','" & Format(pdFechaFin, "YYYY/MM/DD") & "','" & psPaquete & "','" & psMoneda & "','" & psTipoInfo & "'"
    oCon.AbreConexion
    Set rs = oConect.CargaRecordSet(sSql)
    
    oCon.CierraConexion
End Function

Public Function ActualizarFONDEMICAB_Estado(pdFecCierre As Date, psPaquete As String)
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    
    sSql = "exec stp_upd_FONDEMICAB_nEstado '" & Format(pdFecCierre, "YYYY/MM/DD") & "','" & psPaquete & "'"
    oCon.AbreConexion
    Set rs = oConect.CargaRecordSet(sSql)
    
    oCon.CierraConexion
End Function

Public Function ObtenerFONDEMICAB(pdFecCierre As Date, psPaquete As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    
    sSql = "exec stp_sel_FONDEMICAB '" & Format(pdFecCierre, "YYYY/MM/DD") & "','" & psPaquete & "'"
    oCon.AbreConexion
    Set rs = oConect.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerFONDEMICAB = rs
End Function

Public Function ObtenerFONDEMIDET(pdFecCierre As Date, psPaquete As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    
    sSql = "exec stp_sel_FONDEMIDET '" & Format(pdFecCierre, "YYYY/MM/DD") & "','" & psPaquete & "'"
    oCon.AbreConexion
    Set rs = oConect.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerFONDEMIDET = rs
End Function
Public Function ObtenerFONDEMICAB_SP() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    
    sSql = "exec stp_sel_FONDEMICAB_SP "
    oCon.AbreConexion
    Set rs = oConect.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerFONDEMICAB_SP = rs
End Function
Public Function ObtenerPaqueteFONDEMI() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    
    sSql = "exec stp_sel_PaqueteFONDEMI "
    oCon.AbreConexion
    Set rs = oConect.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerPaqueteFONDEMI = rs
End Function
'*******************************************


