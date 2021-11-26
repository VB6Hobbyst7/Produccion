VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NAnalisisCtas"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A8AE347029F"
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Public sCentralCom As String

Public Function CargaPendientesHisto(psOpeCod As String, Optional psCtaCod As String = "", Optional psMovNro As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaPendientesHistoErr
   psSql = "SELECT m.nMovNro, cMovNro, cMovDesc,convert(varchar(10),convert(datetime, LEFT(m.cMovNro,8),101),103) Fecha, cCtaContCod, nMovImporte, m.nMovEstado, m.nMovFlag " _
         & "FROM Mov m join MovCta mc on mc.nmovnro = m.nmovnro " _
         & "WHERE m.copecod like '" & psOpeCod & "%' and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & ") and m.nMovEstado = " & gMovEstContabPendiente & " ORDER BY mc.cCtaContCod, m.cMovNro "
   Set CargaPendientesHisto = dbConec.CargaRecordSet(psSql, LockType)
   Exit Function
CargaPendientesHistoErr:
   Call RaiseError(MyUnhandledError, "DAnalisisCtas:CargaPendientesHisto Method")
End Function

Public Sub InsertaPendienteHisto(psMovNro As String, psMovDesc As String, psOpeCod As String, psCtaCod As String, pnMovImporte As Currency)
Dim clsMov   As New DMov
Dim lnMovNro As Long
   On Error GoTo InsertaPendienteHistoErr
   clsMov.BeginTrans
   clsMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabPendiente, gMovFlagVigente
   lnMovNro = clsMov.GetnMovNro(psMovNro)
   clsMov.InsertaMovCta lnMovNro, 1, psCtaCod, pnMovImporte
   clsMov.InsertaMovPendientesRend lnMovNro, psCtaCod, pnMovImporte
   clsMov.CommitTrans
   Set clsMov = Nothing
   Exit Sub
InsertaPendienteHistoErr:
   clsMov.RollbackTrans
   Set clsMov = Nothing
   Call RaiseError(MyUnhandledError, "DAnalisisCtas:InsertaPendienteHisto Method")
End Sub

Public Sub ActualizaPendienteHisto(pnMovNro As Long, psMovDesc As String, psCtaCod As String, pnMovImporte As Currency, psMovEstado As MovEstado, psMovFlag As MovFlag)
Dim clsMov As New DMov
   On Error GoTo ActualizaPendienteHistoErr
   clsMov.BeginTrans
   clsMov.ActualizaMov pnMovNro, psMovDesc, psMovEstado, psMovFlag
   clsMov.ActualizaMovCta pnMovNro, 1, psCtaCod, pnMovImporte
   clsMov.CommitTrans
   Set clsMov = Nothing
   Exit Sub
ActualizaPendienteHistoErr:
   clsMov.RollbackTrans
   Set clsMov = Nothing
   Call RaiseError(MyUnhandledError, "DAnalisisCtas:ActualizaPendienteHisto Method")
End Sub

Public Sub EliminaPendienteHisto(psMovNro As String)
Dim clsMov As New DMov
   On Error GoTo EliminaPendienteHistoErr
   clsMov.EliminaMov psMovNro
   Set clsMov = Nothing
   Exit Sub
EliminaPendienteHistoErr:
   Call RaiseError(MyUnhandledError, "DAnalisisCtas:EliminaPendienteHisto Method")
End Sub

Public Function GetSubsidiosPendientes(psPlaCod As String, pdFechaAl As Date, psOpeCodPlanilla As String, Optional psOpeCod As String = "", Optional psCtaPend As String = "", Optional pnMoneda As Integer = 0) As ADODB.Recordset
Dim lsFiltro As String
Dim lsJoin   As String

 On Error GoTo GetSubsidiosPendientesErr
 If psCtaPend = "" Then
    lsJoin = ", OpeCta oc "
    lsFiltro = " and oc.cOpecod = '" & psOpeCod & "' and cc.cCtaContCod = LEFT(oc.cCtaContCod,2)+'" & pnMoneda & "'+SubString(oc.cCtaContCod,4,22) "
 Else
    lsFiltro = " and cc.cCtaContCod = '" & psCtaPend & "' "
 End If
'psSql = "SELECT 'PLA' cDocAbrev, '' cDocTpo, pla.cPlanillaCod cDocNro, Convert( varchar(10), pd.cRRHHPeriodo) dDocFecha, " _
'     & "       p.cPersNombre, p.cPersCod cCodPers, " _
'     & "       pla.nMovNroest cMovNro, pla.cPlaInsDes cMovDesc, pd.nMonto nMovImporte, ISNULL(mRend.nSaldo, pd.nMonto) nSaldo " _
'     & "FROM rhplanilladetCon pd JOIN RHPlanilla pla ON pla.cRRHHPeriodo = pd.cRRHHPeriodo " _
'     & "     JOIN Mov m ON m.nMovNro = pla.nMovNroEst " _
'     & "     JOIN RHConceptoCta cc ON cc.cRHConceptoCod = pd.cRHConceptoCod " _
'     & "     JOIN Persona p ON p.cPersCod = pd.cPersCod " _
'     & "     LEFT JOIN (SELECT m.nMovNro, nMovNroRef, m.cMovNro FROM MovRef mr JOIN Mov m ON m.nMovNro = mr.nMovNro " _
'     & "                WHERE m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag <> " & gMovFlagVigente & ") ref ON ref.nMovNroRef = pla.nMovNroEst " _
'     & "     LEFT JOIN MovPendientesRend mRend ON mRend.nMovNro = pla.nMovNroEst " _
'     & "WHERE not pla.nMovNroEst is NULL and LEFT(m.cMovNro,8) <= '" & Format(pdFechaAl, gsFormatoMovFecha) & "' and cc.cOpeCod like '" & psOpeCodPlanilla & "' " & lsFiltro _
'     & "      and pd.cPlanillaCod = '" & psPlaCod & "' " _
'     & "and ( ref.nMovNro is NULL or mRend.nSaldo <> 0) "

psSql = "SELECT 0 Ok, 'PLA' cDocAbrev, pla.cPlanillaCod cDocNro, p.cPersNombre cNomPers, Convert( varchar(10), pd.cRRHHPeriodo) dDocFecha, Convert(money, pd.nMonto) nMovImporte, " _
     & "       pla.cPlaInsDes cMovDesc, p.cPersCod, pla.nMovNroEst nMovNro, 0 nDocTpo, ISNULL(mRend.nSaldo, ISNULL(pd.nMonto+ref.nPago ,pd.nMonto)) nSaldo, 0 Rendicion, '' cMovNro " _
     & "FROM RHPlanilladetCon pd JOIN RHPlanilla pla ON pla.cRRHHPeriodo = pd.cRRHHPeriodo AND pla.cPlanillaCod = pd.cPlanillaCod " _
     & "     JOIN RHConceptoCta cc ON cc.cRHConceptoCod = pd.cRHConceptoCod " _
     & "     JOIN Persona p ON p.cPersCod = pd.cPersCod " _
     & "     LEFT JOIN (SELECT m.cMovNro, m.nMovNro, nMovNroRef, SUM(mc.nMovImporte) as nPago FROM MovRef mr JOIN Mov m ON m.nMovNro = mr.nMovNro JOIN MovCta mc ON mc.nMovNro = m.nMovNro and mc.cCtaContCod = '" & psCtaPend & "' " _
     & "                WHERE m.nMovEstado = '" & gMovEstContabMovContable & "' and m.nMovFlag <> '" & gMovFlagEliminado & "' " _
     & "                GROUP BY m.cMovNro, m.nMovNro, nMovNroRef) ref ON ref.nMovNroRef = pla.nMovNroEst " _
     & "     LEFT JOIN MovPendientesRend mRend ON mRend.nMovNro = pla.nMovNroEst " _
     & "WHERE not pla.nMovNroEst is NULL and cc.cOpeCod like '" & psOpeCodPlanilla & "' and cc.cCtaContCod = '" & psCtaPend & "' " _
     & "      and pd.cPlanillaCod = '" & psPlaCod & "' " _
     & " and ( ISNULL(mRend.nSaldo, ISNULL(pd.nMonto+ref.nPago ,pd.nMonto)) <> 0 or mRend.nSaldo <> 0 ) "
Set GetSubsidiosPendientes = dbConec.CargaRecordSet(psSql)

Exit Function
GetSubsidiosPendientesErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oConec As New ClasIni
   sCentralCom = oConec.BaseComunes
   Set oConec = Nothing
   
   Set dbConec = New DConecta
   dbConec.AbreConexion
   
End Sub

Private Sub Class_Terminate()
   dbConec.CierraConexion
   Set dbConec = Nothing
End Sub

Public Function GrabaRendirPendIngresoCheque(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                            ByVal psCtaPendiente As String, ByVal psCtaContCheque As String, _
                                            ByVal psObjetoCodProd As String, ByVal psObjetoArea As String, ByVal psAgeCod As String, _
                                            ByVal pnImporte As Currency, _
                                            ByVal psNumCheque As String, ByVal psPersCodIF As String, ByVal psTipoIF As String, ByVal pnPlaza As ChequePlaza, _
                                            ByVal psCtaCheque As String, ByVal pdFechaReg As Date, ByVal pdFechaVal As Date, ByVal pnMonedaCheque As Moneda, _
                                            prsPend As ADODB.Recordset, _
                                            Optional ByVal pnEstadoCheque As ChequeEstado = gChqEstEnValorizacion, _
                                            Optional ByVal pnEstadoCG As CGEstadosChq = gCGEstadosChqRecibido, _
                                            Optional ByVal pnConfirmaCaja As CGEstadoConfCheque = ChqCGSinConfirmacion, _
                                            Optional ByVal psAreaCodChq As String, Optional ByVal psAgeCodChq As String) As Integer
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim lsSubCtaDebe As String
Dim oContFunct As NContFunciones
Dim lsMsgErr As String

Set oDMov = New DMov
Set oContFunct = New NContFunciones

On Error GoTo GrabaRendirPendIngresoChequeErr

GrabaRendirPendIngresoCheque = 1
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaContCheque, pnImporte
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, psObjetoCodProd
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, psAgeCod, psObjetoArea

'guardamos la cuenta de ContraResta
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaPendiente, pnImporte * -1
If psNumCheque <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocCheque, psNumCheque, Format(pdFechaReg, gsFormatoFecha)
End If
'grabacion dentro de la tabla de cheques
oDMov.InsertaCheque TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pnPlaza, psCtaCheque, pnImporte, pdFechaReg, pdFechaVal, pnEstadoCG, pnConfirmaCaja, pnMonedaCheque, , psAreaCodChq, psAgeCodChq
oDMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIF, psTipoIF, pdFechaReg, gChqEstEnValorizacion, psMovNro, psCtaCheque

If Not prsPend Is Nothing Then
    Do While Not prsPend.EOF
        If prsPend!Ok = "1" Then
            oDMov.InsertaMovRef lnMovNro, prsPend!nMovNro
            oDMov.ActualizaMovPendientesRend prsPend!nMovNro, psCtaPendiente, prsPend!Rendicion
        End If
        prsPend.MoveNext
    Loop
End If
If Mid(psOpeCod, 3, 1) = "2" Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If

oDMov.CommitTrans
lbTrans = False
GrabaRendirPendIngresoCheque = 0

Set oDMov = Nothing
Exit Function
GrabaRendirPendIngresoChequeErr:
  lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, lsMsgErr
End Function

Public Function GrabaRendirPendGiroDocumento(ByVal pnTipoArendir As ArendirTipo, _
                                            ByVal psMovNro As String, ByVal pnMovNroSol As Long, ByVal psMovNroAtenc As String, _
                                            ByVal psOpeCod As String, ByVal psMovDesc As String, ByVal psCtaDebe As String, _
                                            ByVal psCtaHaber As String, ByVal psPersCod As String, ByVal pnImporte As Currency, _
                                            ByVal psDocTpo As TpoDoc, ByVal psDocNro As String, _
                                            ByVal psFechaDoc As String, ByVal psDocNroVoucher As String, _
                                            ByVal psPersCodIF As String, ByVal psTipoIF As String, ByVal psCodCtaIf As String, _
                                            Optional ByVal psAreaCh As String = "", Optional ByVal psAgeCh As String = "", _
                                            Optional pnNroProc As Integer = 0, Optional pnMotivo As Long = -1, Optional psCtaCod As String = "", Optional gbBitCentral As Boolean = False, Optional pbGrabaOpeNegocio As Boolean = False) As Integer
                        
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim oDocRec As NDocRec
Dim lsMsg   As String
Set oDMov = New DMov
Set oDocRec = New NDocRec
On Error GoTo ErrorGrabaRendirPendGiroDocumento

GrabaRendirPendGiroDocumento = 1
oDMov.BeginTrans
lbTrans = True
If pnTipoArendir = gArendirTipoCajaChica Then
    oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
Else
    oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
End If
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1: lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe, Abs(pnImporte)
If pnTipoArendir = gArendirTipoCajaChica And psAreaCh <> "" Then
    oDMov.InsertaMovObj lnMovNro, lnMovItem, "1", ObjCMACAgenciaArea
    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, "1", psAgeCh, psAreaCh
End If
'guardamos la cuenta del haber
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber, Abs(pnImporte) * -1
If psPersCodIF <> "" And psCodCtaIf <> "" Then
    lnMovOrden = lnMovOrden + 1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, Format(lnMovItem, "#0"), Format(lnMovOrden, "#0"), psPersCodIF, psTipoIF, psCodCtaIf
    
    'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
'    oDMov.InsertaMovCtaIF lnMovNro, psPersCodIf, psTipoIF, psCodCtaIf, CGCtaIFConCapital, pnImporte * -1
'    oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
'                                    psPersCodIf, psTipoIF, psCodCtaIf
End If
Select Case Val(psDocTpo)
    Case TpoDocNotaAbono, TpoDocNotaCargo
        oDMov.InsertaNotaAbonoCargo psDocTpo, psDocNro, gNCNARegistrado, pnMotivo, Abs(pnImporte)
        oDMov.InsertaNotaAbonoCargoEst psDocTpo, psDocNro, gNCNARegistrado, psMovNro
        oDMov.InsertaRegDocCuenta psDocTpo, psDocNro, psCtaCod
    
        If pbGrabaOpeNegocio Then
            Dim oCapta As NCapMovimientos
            Dim lnSaldo As Double
            Set oCapta = New NCapMovimientos
            If gbBitCentral Then
               lnSaldo = oCapta.CapAbonoCuentaAho(psCtaCod, pnImporte, gCGPagProvAbonoCent, psMovNro, psMovDesc, TpoDocNotaAbono, psDocNro, , , , , , , , , , True, False, , , oDMov.GetConexion)
            Else
                Dim oCon As New DConecta
                Dim oDis As New NRHProcesosCierre
                
                If oCon.AbreConexion Then 'Remota(Left(psCtaCod, 2))
                    oCon.BeginTrans
                    lnSaldo = oDis.Abono(psCtaCod, pnImporte, gCGPagProvAbonoDist, gCGPagProvAbonoDist, "112" & Left(psCtaCod, 2), Right(psMovNro, 4), psDocNro, "PAGO : " & psMovDesc, oCon, CDate(Format(psFechaDoc, "dd/mm/yyyy")))
                    oCon.CommitTrans
                End If
                oCon.CierraConexion
                Set oCon = Nothing
                Set oDis = Nothing
            End If
        End If
    
    Case TpoDocOrdenPago
        
End Select
If psDocNro <> "" Then
    oDMov.InsertaMovDoc lnMovNro, psDocTpo, psDocNro, psFechaDoc
End If
If psDocNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocNroVoucher, psFechaDoc
End If
oDMov.InsertaMovGasto lnMovNro, psPersCod, "0"
oDMov.InsertaMovRef lnMovNro, psMovNroAtenc
If Mid(psOpeCod, 3, 1) = "2" Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"
GrabaRendirPendGiroDocumento = 0

oDMov.CommitTrans
lbTrans = False
Set oDMov = Nothing
Exit Function
ErrorGrabaRendirPendGiroDocumento:
    lsMsg = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaRendirPendGiroDocumento", lsMsg
End Function

Public Function GetOpePendientesNegocio(pbBitCentral As Boolean, pdFecha As Date, pnMoneda As Integer, psCodAge As String, psCtaPendiente As String) As ADODB.Recordset
If pbBitCentral Then
    psSql = "SELECT distinct 0 Ok, d.cDocAbrev, md.cDocNro, ISNULL(PB.cPersNombre,POV.cPersNombre) cNomPers, md.dDocFecha dDocFecha, " _
          & "      ABS(ov.nMovImporte) nMovImporte, m.cMovDesc, ISNULL(TB.cPersCodIF,MG.cPersCod) cCodPers, " _
          & "      m.nMovNro nMovNro, md.nDocTpo, ABS(ov.nMovImporte) nSaldo, 0 nRendicion, m.cMovNro, m.cOpeCod cCodOpe " _
          & "From Mov M JOIN (SELECT Distinct cOpeCod FROM OpeCtaNeg oc " _
          & "           WHERE '" & Replace(psCtaPendiente, "_", pnMoneda) & "' LIKE REPLACE(oc.cCtaContCod,'M','" & pnMoneda & "') + '%') oc ON oc.cOpeCod = M.cOpeCod  " _
          & " LEFT JOIN MovDoc md ON md.nMovNro = m.nMovNro LEFT JOIN Documento d ON d.nDocTpo = md.nDocTpo " _
          & "     LEFT JOIN MovTransferBco TB ON M.nMovNro = TB.nMovNro LEFT JOIN MovCap mc ON mc.nMovNro = m.nMovNro And substring(MC.cCtaCod,9,1) = " & pnMoneda & " LEFT JOIN Persona  PB ON PB.cPersCod = TB.cPersCodIF " _
          & "     LEFT JOIN MovOpeVarias   OV ON M.nMovNro = OV.nMovNro and ov.nMoneda = " & pnMoneda & " LEFT JOIN MovGasto MG ON mg.nMovNro = M.nMovNro LEFT JOIN Persona POV ON POV.cPersCod = MG.cPersCod " _
          & "WHERE M.nMovNro NOT IN (Select MR.nMovNroRef From Mov M1 JOIN MovRef MR ON M1.nMovNro = MR.nMovNro " _
          & "  And M1.nMovFlag = " & gMovFlagVigente & " ) And M.nMovFlag = " & gMovFlagVigente & " and (not tb.nMovNro is NULL or not ov.nMovNro is NULL) " _
          & "  and SubString(M.cMovNro,18,2) = '" & Right(psCodAge, 2) & "' "
    dbConec.AbreConexion
    Set GetOpePendientesNegocio = dbConec.CargaRecordSet(psSql)
Else
    psSql = "SELECT distinct 0 Ok, '' cDocAbrev, td.cNumDoc cDocNro, ISNULL(p.cNomPers,'') cNomPers, Convert(VarChar(10), td.dFecTran,103) dDocFecha, " _
          & "      ABS(td.nMonTran) nMovImporte, ISNULL(oo.cGlosa,'') cMovDesc, p.cCodPers cCodPers, " _
          & "      td.nNumTran nMovNro, '' nDocTpo, ABS(td.nMonTran) nSaldo, 0 nRendicion, '' cMovNro, tr.cCodOpe " _
          & "FROM TransAho td JOIN TransRef tr ON tr.cNroTranRef = td.nNumTran " _
          & "     LEFT JOIN OtrasOpe oo ON oo.nNumTran = td.nNumTran " _
          & "     LEFT JOIN DBpersona..Persona p ON p.cCodPers = oo.cCodPers " _
          & "     JOIN DBComunes..OpeCuentaN oc ON oc.cCodOpe = td.cCodOpe " _
          & "WHERE DatedIff(d,td.dFecTran,'" & Format(pdFecha, gsFormatoFecha) & "') >= 0 and SubString(td.cCodCta,6,1) = '" & pnMoneda & "' " _
          & "  and td.cFlag is NULL and cNroTran is NULL and oc.cCodCnt LIKE '" & psCtaPendiente & "' " _
          & "ORDER BY nMovNro "
    If dbConec.AbreConexion Then 'Remota(psCodAge, False)
        Set GetOpePendientesNegocio = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
        dbConec.AbreConexion
    End If
End If

End Function

Public Function GetOpePendientesNegocioRef(pbBitCentral As Boolean, pdFecha As Date, pdFechaAl As Date, pnMoneda As Integer, psCodAge As String, psCtaPendiente As String) As ADODB.Recordset
If pbBitCentral Then
Else
    psSql = "SELECT Convert(VarChar(8), td.dFecTran,112) , " _
          & "      left(replace(convert(varchar(14),td.dFecTran,114),':',''),6)+ '-' + td.cCodUsu, ISNULL(oo.cGlosa,'') + ' ' + ISNULL(p.cNomPers,'') cMovDesc, " _
          & "      ABS(td.nMonTran) * " & IIf(Left(psCtaPendiente, 1) = "2", "-1", "1") & " nMovImporte, ABS(td.nMonTran) nSaldo, td.nNumTran nMovNro, td.nNumTran cMovNro, '" & psCodAge & "' cAgeCod " _
          & "FROM TransAho td JOIN TransRef tr ON tr.cNroTranRef = td.nNumTran " _
          & "     LEFT JOIN OtrasOpe oo ON oo.nNumTran = td.nNumTran " _
          & "     LEFT JOIN DBpersona..Persona p ON p.cCodPers = oo.cCodPers " _
          & "     JOIN DBComunes..OpeCuentaN oc ON oc.cCodOpe = td.cCodOpe " _
          & "WHERE td.dFecTran BETWEEN '" & Format(pdFecha, gsFormatoFecha) & "' and '" & Format(pdFechaAl, gsFormatoFecha) & " 23:59:59' and SubString(td.cCodCta,6,1) = '" & pnMoneda & "' " _
          & "  and td.cFlag is NULL and cNroTran is NULL and oc.cCodCnt LIKE '" & psCtaPendiente & "' and oc.cDebeHab = '" & IIf(Left(psCtaPendiente, 1) = "2", "H", "D") & "'" _
          & "ORDER BY nMovNro "
End If
If psSql = "" Then
    Exit Function
End If
If dbConec.AbreConexion Then 'Remota(psCodAge, False)
    Set GetOpePendientesNegocioRef = dbConec.CargaRecordSet(psSql)
    dbConec.CierraConexion
    dbConec.AbreConexion
End If
End Function

Public Function GetOpePendientesMov(pbBitCentral As Boolean, pdFecha As Date, pnMoneda As Integer, psCtaPendiente As String, psTpoCta As String) As ADODB.Recordset
    psSql = "SELECT 0 Ok, ISNULL(d.cDocAbrev,'') cDocAbrev, ISNULL(md.cDocNro,'') cDocNro, ISNULL(p.cPersNombre,'') cNomPers, " _
         & "       convert(varchar(10),convert(datetime,LEFT(m.cMovNro,8)),103) dDocFecha, " _
         & "       " & IIf(pnMoneda = 1, "abs(mc.nMovImporte) ", "abs(ISNULL(me.nMovMEImporte,0))") & " nMovImporte, m.cMovDesc, ISNULL(mop.cPersCod,'') cPersCod, m.nMovNro, ISNULL(md.nDocTpo,'') nDocTpo, mRend.nSaldo, 0 nRendicion, m.cMovNro " _
         & "FROM mov m join MovCta mc on m.nMovNro = mc.nMovNro " _
         & "      LEFT JOIN MovMe  me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
         & "      LEFT JOIN MovDoc md ON md.nMovNro = m.nMovNro  and not md.nDocTpo = " & TpoDocVoucherEgreso & " LEFT JOIN Documento d ON d.nDocTpo = md.nDocTpo " _
         & "      LEFT JOIN MovGasto moP ON moP.nMovNro = mc.nMovNro " _
         & "      LEFT JOIN Persona P ON p.cPersCod = moP.cPersCod " _
         & "           JOIN MovPendientesRend mRend ON mRend.nMovNro = m.nMovNro and mRend.cCtaContCod = mc.cCtaContCod " _
         & "WHERE LEFT(m.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' and m.nMovEstado in (" & gMovEstContabMovContable & "," & gMovEstContabPendiente & ") and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & ") " _
         & "      and nMovImporte " & IIf(psTpoCta = "D", " > ", " < ") & " 0 and mc.cCtaContCod like  '" & psCtaPendiente & "' and m.cOpeCod LIKE '__" & pnMoneda & "%' " _
         & "      and mRend.nSaldo <> 0 and not m.cMovNro like '%XXX_' " _
         & "ORDER BY mc.cCtaContCod, m.cMovNro "
Set GetOpePendientesMov = dbConec.CargaRecordSet(psSql)
End Function

Public Function CargaCtaContPendiente(psMoneda As String, Optional psCtaCont As String = "", Optional pbUltimoNiv As Boolean = False) As ADODB.Recordset
psSql = "SELECT p.cCtaContCod, p.cCtaContDesc, 1 nNivel FROM CtaContPend p " _
      & "WHERE  SubString(p.cCtaContCod,3,1) = '" & psMoneda & "' " & IIf(psCtaCont = "", "", " and p.cCtaContCod = '" & psCtaCont & "'")
If pbUltimoNiv Then
    psSql = psSql & "UNION ALL " _
          & "SELECT c.cCtaContCod, c.cCtaContDesc, 2 nNivel " _
          & "FROM   CtaContPend p JOIN CtaCont c ON c.cCtaContCod LIKE p.cCtaContCod + '_%' " _
          & "WHERE SubString(p.cCtaContCod,3,1) = '" & psMoneda & "' " & IIf(psCtaCont = "", "", " and p.cCtaContCod = '" & psCtaCont & "'")
End If
psSql = psSql & " ORDER BY cCtaContCod "

Set CargaCtaContPendiente = dbConec.CargaRecordSet(psSql)

End Function
'JACA 20110819*************************************************
Public Function GetOpePendientesRegAsienCont(pdFecha As Date, pnMoneda As Integer, psCtaPendiente As String) As ADODB.Recordset
    'con referencia a la funcion GetOpePendientesMov
    psSql = "exec stp_sel_OpePendientesRegAsienCont '" & Format(pdFecha, gsFormatoMovFecha) & "','" & pnMoneda & "','" & psCtaPendiente & "'"
    Set GetOpePendientesRegAsienCont = dbConec.CargaRecordSet(psSql)
End Function
'JACA END******************************************************
