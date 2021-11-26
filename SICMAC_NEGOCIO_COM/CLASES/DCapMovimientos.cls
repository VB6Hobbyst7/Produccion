VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCapMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public dbCmact As Connection
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String
Dim sSql As String
Public bTransaccion As Boolean

Public Function GetDataInteresEspecial(ByVal cPersCod As String, ByVal nEstado As Integer, ByVal nProd As Integer, ByVal nMoneda As Integer, ByVal dFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim RSTEMP As ADODB.Recordset, cFecha1 As String, sqlaux As String
Dim dFecha2 As Date, cFecha2 As String
cFecha1 = Year(dFecha) & String(2 - Len(Month(dFecha)), "0") & CStr(Month(dFecha)) & String(2 - Len(Day(dFecha)), "0") & CStr(Day(dFecha))
dFecha2 = DateAdd("d", -2, dFecha)
cFecha2 = Year(dFecha2) & String(2 - Len(Month(dFecha2)), "0") & CStr(Month(dFecha2)) & String(2 - Len(Day(dFecha2)), "0") & CStr(Day(dFecha2))
If nEstado = 99 Then
    sqlaux = " And nestado in (0,1,3)"
Else
    sqlaux = " And nestado=" & nEstado
End If

Set RSTEMP = New ADODB.Recordset
    sSql = "Select ntasa,cestado=(case when nestado=0 then 'SOLICITADA' "
    sSql = sSql & "                    when nestado=1 then 'APROBADA' "
    sSql = sSql & "                    when nestado=2 then 'APERTURADA'   "
    sSql = sSql & "                    when nestado=3 then 'RECHAZADA'"
    sSql = sSql & "                    when nestado=4 then 'EXTORNADA'  end ),"
    sSql = sSql & " nestado,nmonto from captasaespecial where cPersCod='" & cPersCod & "' and nProducto= " & nProd & " And nMoneda = " & nMoneda & sqlaux
    sSql = sSql & " and cast(left(cmovnro,8) as bigint)<= cast('" & cFecha1 & "' as bigint) and cast(left(cmovnro,8) as bigint)>=cast('" & cFecha2 & "' as bigint) "
    sSql = sSql & " ORDER BY CAST(LEFT(CMOVNRO,14) AS BIGINT) DESC "
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
   ' Set rstemp.ActiveConnection = Nothing
    Set GetDataInteresEspecial = RSTEMP
    Set RSTEMP.ActiveConnection = Nothing
    If RSTEMP.State = 1 Then RSTEMP.Close
    
    Set RSTEMP = Nothing

End Function



Public Function GetnMovNro(ByVal sMovNro As String) As Long
Dim rsMov As Recordset
sSql = "Select nMovNro From Mov Where cMovNro = '" & sMovNro & "'"
Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
If rsMov.EOF And rsMov.BOF Then
    GetnMovNro = 0
Else
    GetnMovNro = rsMov("nMovNro")
End If
rsMov.Close
Set rsMov = Nothing
End Function


Public Function AgregaNuevaCaptacion(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal sAgencia As String, ByVal nTasa As Double, ByVal nSaldo As Double, _
        ByVal dFecha As Date, ByVal nFirmas As Integer, ByVal nPersoneria As PersPersoneria, _
        ByVal nTipoCuenta As ProductoCuentaTipo, ByVal sMovNro As String, _
        ByVal nTipoTasa As CaptacTipoTasa, Optional bCheque As Boolean = False, _
        Optional ByVal bOrdPag As Boolean = False, Optional nPlazo As Long = 0, _
        Optional nFormaRetiro As CaptacPFFormaRetiro = gCapPFFormRetMensual, _
        Optional bCtaAboInt As Boolean = False, Optional sCtaCodAbono As String = "", _
        Optional nPorcRetCTS As Double = 0, Optional sInstitucion As String = "", Optional ByVal cAlias As String = "", Optional ByVal nFirmasMin As Integer = 1, _
        Optional cPersTasaEspecial As String, Optional bEsTasaPermanente As Boolean, Optional ByVal nMonITF As Currency = 0, Optional ByVal nTpoPrograma As Integer = 0) As String

Dim clsGen As DGeneral
Dim sCuenta As String, sFecha As String
Dim sFecUltCierre As String, sOrdPag As String, sCtaAbonoIntPF As String
Dim nSaldoDisp As Double, nSaldRetCTS As Double
Set clsGen = New DGeneral
sCuenta = clsGen.GeneraNuevaCuenta(sAgencia, nProducto, nMoneda)
Set clsGen = Nothing
sFecha = Format$(dFecha, "mm/dd/yyyy") & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
sSql = "Insert Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) " _
    & "Values ('" & sCuenta & "'," & nTasa & "," & nSaldo & "," & gCapEstActiva & ",'" & sFecha & "',1)"
dbCmact.Execute sSql
nSaldoDisp = IIf(bCheque, 0 - nMonITF, nSaldo)
sFecUltCierre = Format$(DateAdd("d", -1, dFecha), "mm/dd/yyyy")
If nProducto = gCapAhorros Then
    sSql = "Insert Captaciones (cCtaCod,nSaldoDisp,nPersoneria,nFirmas,nIntAcum,dUltCierre,dApertura,nPrdCtaTpo,nPrdTasaInteres,cUltimaActualizacion,cAlias,nFirmasMin,nTpoPrograma) " _
        & " Values ('" & sCuenta & "'," & nSaldoDisp & "," & nPersoneria & "," & nFirmas & ",0,'" & sFecUltCierre & "','" & sFecha & "'," & nTipoCuenta & "," _
        & nTipoTasa & ",'" & sMovNro & "','" & cAlias & "'," & nFirmasMin & "," & nTpoPrograma & ")"
    dbCmact.Execute sSql
Else
    sSql = "Insert Captaciones (cCtaCod,nSaldoDisp,nPersoneria,nFirmas,nIntAcum,dUltCierre,dApertura,nPrdCtaTpo,nPrdTasaInteres,cUltimaActualizacion,cAlias,nFirmasMin) " _
        & "Values ('" & sCuenta & "'," & nSaldoDisp & "," & nPersoneria & "," & nFirmas & ",0,'" & sFecUltCierre & "','" & sFecha & "'," & nTipoCuenta & "," _
        & nTipoTasa & ",'" & sMovNro & "','" & cAlias & "'," & nFirmasMin & ")"
    dbCmact.Execute sSql
End If
Select Case nProducto
    Case gCapAhorros
        sOrdPag = IIf(bOrdPag, "1", "0")
        sSql = "Insert CaptacAhorros (cCtaCod,nSaldoAnterior,bOrdPag,nSobregiro,dUltContacto) " _
            & "Values ('" & sCuenta & "',0," & sOrdPag & ",0,'" & sFecha & "')"
    Case gCapPlazoFijo
        sCtaAbonoIntPF = IIf(bCtaAboInt, "1", "0")
        sSql = "Insert CaptacPlazoFijo (cCtaCod,nPlazo,nIntPag,dRenovacion,nApertura,dAuxiliar,nFormaRetiro,nDuplicado,bAbonoIntCtaAho,bBusCredPend,bPermanente) " _
            & "Values ('" & sCuenta & "'," & nPlazo & ",0,'" & sFecha & "'," & nSaldo & ",'" & sFecha & "'," & nFormaRetiro & ",0," & sCtaAbonoIntPF & ",0," & IIf(bEsTasaPermanente, 1, 0) & " )"
    Case gCapCTS
        nSaldRetCTS = Round(nSaldo * nPorcRetCTS / 100, 2)
        sSql = "Insert CaptacCTS (cCtaCod,nSaldRetiro,nIntSaldo,cCodInst) " _
            & "Values ('" & sCuenta & "'," & nSaldRetCTS & ",0,'" & sInstitucion & "')"
End Select
dbCmact.Execute sSql
If nProducto = gCapPlazoFijo And bCtaAboInt Then
    sSql = "Insert CaptacCtaAboIntPF (cCtaCod,cCtaCodAbono) " _
        & "Values ('" & sCuenta & "','" & sCtaCodAbono & "')"
    dbCmact.Execute sSql
End If

'If bEsTasaPermanente Then
'    sSql = "Insert Into captasaespecial(cPersCod,nProducto,nMoneda,cMovNro,ntasa,cCtaCod,nMonto,nPlazo) "
'    sSql = sSql & " values('" & cPersTasaEspecial & "'," & nProducto & "," & nMoneda & ", '" & sMovNro & "'," & nTasa & ",'" & sCuenta & "'," & nSaldo & "," & nPlazo & ")"
'    dbCmact.Execute sSql
'End If

AgregaNuevaCaptacion = sCuenta
End Function

Public Sub AgregaNuevoProdPers(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona, Optional ByVal COBLIGATORIO As String)
If Mid(sCuenta, 6, 3) = "232" Or Mid(sCuenta, 6, 3) = "233" Then
 sSql = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac,cObligatorio) " _
        & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ",'" & COBLIGATORIO & "')"

Else
    sSql = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac,cObligatorio) " _
         & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ",'S')"
End If
dbCmact.Execute sSql
End Sub

Public Function GetDatosCheque(ByVal sNroDoc As String, ByVal scperscod As String, ByVal sTipoIF As String) As Recordset
Dim rsChe As Recordset
sSql = "  select cNroDoc,cPersCod,cIFTpo, bPlaza,cIFCta=isnull(cIFCta,''),nMonto, " _
       & " dValorizacion,cDepIF,nConfCaja,nEstado,nMoneda,cAreaCod,cAgeCod, nProducto  from  docrec " _
       & " where cnrodoc='" & sNroDoc & "' and cperscod='" & scperscod & "' and ciftpo='" & sTipoIF & "'"

Set rsChe = New Recordset
rsChe.CursorLocation = adUseClient
rsChe.Open sSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
Set rsChe.ActiveConnection = Nothing
Set GetDatosCheque = rsChe
Set rsChe = Nothing

End Function

Public Function GetVerificaRetiro(ByVal sCuenta As String) As Boolean
Dim rs As Recordset


GetVerificaRetiro = False
sSql = "  select mc.nmonto from mov m join movcap mc on mc.nmovnro=m.nmovnro "
sSql = sSql & " where mc.cctacod='" & sCuenta & "' and mc.copecod='220303' and m.nmovflag=0 "
       
Set rs = New Recordset
rs.CursorLocation = adUseClient
rs.Open sSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
If Not rs.EOF Then
     If rs.Fields("nmonto") > 0 Then
      GetVerificaRetiro = True
     End If
End If


Set rs.ActiveConnection = Nothing

Set rs = Nothing



End Function


Public Function GetDatosOrdenPago(ByVal sCuenta As String, ByVal sNroDoc As String) As Recordset
Dim rsOP As Recordset
sSql = "SELECT E.cMovNro, D.nTpoDoc, D.cNroDoc, D.cCtaCod, D.nMonto, P.cPersNombre, E.cEstado " _
    & "D.cIFCodPers FROM DocRecOPEst E INNER JOIN DocRecOP D INNER JOIN " & sDBPersona & "Persona P " _
    & "ON D.cIFCodPers = P.cPersCod ON E.nTpoDoc = D.nTpoDoc AND E.cNroDoc = D.cNroDoc AND E.cCtaCod " _
    & "= D.cCtaCod WHERE E.cMovNro IN (SELECT MAX(E1.cMovNro) FROM DocRecOPEst E WHERE E1.nTpoDoc = " _
    & "E.nTpoDoc AND E1.cNroDoc = E.cNroDoc AND E1.cCtaCod = E.cCtaCod)"
Set rsOP = New Recordset
rsOP.CursorLocation = adUseClient
rsOP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsOP.ActiveConnection = Nothing
Set GetDatosOrdenPago = rsOP
Set rsOP = Nothing
End Function

Public Sub AgregaOrdenPago(ByVal sCuenta As String, ByVal sNroDoc As String, _
        ByVal sMovNro As String, ByVal nMonto As Double, ByVal nEstadoOP As CaptacOrdPagoEstado, _
        Optional nTpoDoc As TpoDoc = TpoDocOrdenPago, Optional sCodIF As String = "", _
        Optional sIFTipo As String = "")
sSql = "INSERT DocRecOP (nTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers,nEstado,cMovNro,cIFTpo) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCuenta & "'," & nMonto & ",'" & sCodIF & "'," & nEstadoOP & ",'" & sMovNro & "','" & sIFTipo & "')"
dbCmact.Execute sSql
End Sub
Public Sub EliminaOrdenPago(ByVal sCuenta As String, ByVal sNroDoc As String)
sSql = " delete docrecop where cnrodoc='" & sNroDoc & "' and cctacod='" & sCuenta & "'"
dbCmact.Execute sSql
End Sub




Public Sub AgregaDocRecCapta(ByVal sCuenta As String, ByVal nTpoDoc As TpoDoc, _
        ByVal sNroDoc As String, ByVal sCodIF As String, ByVal nMovNro As Long, _
        ByVal nMonto As Double)
sSql = "INSERT DocRecCapta (nTpoDoc,cNroDoc,cPersCod,cIFTpo,nMovNro,cCtaCod,nMonto) " _
    & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "'," & nMovNro & ",'" & sCuenta & "'," & nMonto & ")"
dbCmact.Execute sSql
End Sub


Public Sub AgregaCuentaDocumento(ByVal sCuenta As String, ByVal nTpoDoc As TpoDoc, _
        ByVal sNroDoc As String, ByVal sCodIF As String, ByVal sMovNro As String, _
        ByVal nMovNro As Long, Optional nMonto As Double = 0, _
        Optional nEstadoOP As CaptacOrdPagoEstado = gCapOPEstCobrada, _
        Optional psCtaIf As String)
Dim sFecha As String
sFecha = Mid(sMovNro, 1, 4) & "/" & Mid(sMovNro, 5, 2) & "/" & Mid(sMovNro, 7, 2) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)


Select Case nTpoDoc
    Case TpoDocCheque
        sSql = "INSERT DocRecCapta (nTpoDoc,cNroDoc,cPersCod,cIFTpo,nMovNro,cCtaCod,nMonto,CIFCTA) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "'," & nMovNro & ",'" & sCuenta & "'," & nMonto & ", '" & psCtaIf & "')"
        dbCmact.Execute sSql
        sSql = "Update DocRec Set nEstado = " & gChqEstEnValorizacion & " Where cNroDoc = '" & sNroDoc & "' And " _
            & "cIFTPo = '" & Format$(gTpoIFBanco, "00") & "' And cPersCod = '" & sCodIF & "' and CIFCTA='" & psCtaIf & "'"
        dbCmact.Execute sSql
        sSql = "INSERT DocRecEst (nTpoDoc,cNroDoc,cPersCod,cIFTpo,cMovNro,nEstado,CIFCTA ) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "','" & sMovNro & "'," & gChqEstEnValorizacion & ",'" & psCtaIf & "')"
        dbCmact.Execute sSql
    Case TpoDocOrdenPago
        sSql = "INSERT DocRecOP (nTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers,nEstado,cMovNro) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCuenta & "'," & nMonto & ",'" & sCodIF & "'," & nEstadoOP & ",'" & sMovNro & "')"
        dbCmact.Execute sSql
End Select
sSql = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
    & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
dbCmact.Execute sSql
End Sub

Public Sub AgregaMov(ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
        ByVal sGlosa As String, Optional nMovEstado As MovEstado = gMovEstContabMovContable, _
        Optional nMovFlag As MovFlag = gMovFlagVigente)
sSql = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sGlosa & "'," & nMovEstado & "," & nMovFlag & ")"
dbCmact.Execute sSql
End Sub
Public Sub AgregaMovCapAutorizacionOpe(ByVal lsMovNroSol As String, ByVal lsCuenta As String, ByVal lsCodCliente As String, _
                                        ByVal lnOperacion As CaptacOperacion, ByVal lnOperacionOri As CaptacOperacion, ByVal lsCodEstado As CapEstAutorizaOpe, ByVal lsObservacion As String, _
                                        ByVal lcUltimaActualizacion As String, ByVal lnMondeda As Long, ByVal nMonto As Double, ByVal lcUserOri As String)
                                                                                

    sSql = "INSERT MovCapAutorizacionOpe (nMovNro, cCtaCod, cPersCodCli, cOpeCod, cOpeCodOri, nCodMon ,nMontoAprobado ,nAutEstado, cAutObs, cUltimaActualizacion,cUserOri) " _
        & "VALUES ('" & lsMovNroSol & "','" & lsCuenta & "','" & lsCodCliente & "'," & lnOperacion & "," & lnOperacionOri & "," & lnMondeda & "," & nMonto & ",'" & _
        lsCodEstado & "','" & lsObservacion & "','" & lcUltimaActualizacion & "','" & lcUserOri & "')"
        
    dbCmact.Execute sSql
    
End Sub

Public Sub ActualizaMovCapAutorizacionOpeEst(ByVal lnCodEstado As CapEstAutorizaOpe, ByVal lnMovNro As Long)
                                        

    sSql = "UPDATE MovCapAutorizacionOpe SET nAutEstado =" & lnCodEstado & _
           "  where  nMovNro =" & lnMovNro
    dbCmact.Execute sSql
    
End Sub


Public Sub AgregaMovDoc(ByVal nMovNro As Long, ByVal nTipoDoc As TpoDoc, _
        ByVal sDocumento As String, ByVal dFecDocumento As Date)
Dim sFecha As String

sFecha = Format$(dFecDocumento, "mm/dd/yyyy")
sSql = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
    & "VALUES (" & nMovNro & "," & nTipoDoc & ",'" & sDocumento & "','" & sFecha & "')"
dbCmact.Execute sSql
End Sub
'ALPA 20081009************************************************************
'Public Sub AgregaMovLavDinero(ByVal nMovNro As Long, ByVal sPersCod As String, Optional sPersCodBen As String = "")
''''ssql = "INSERT MovLavDinero (nMovNro,cPersCod,cPersCodBen) " _
''''& "VALUES (" & nMovNro & ",'" & sPersCod & "','" & IIf(sPersCodBen = "", sPersCod, sPersCodBen) & "' )"
Public Sub AgregaMovLavDinero(ByVal nMovNro As Long, ByVal sPersCod As String, Optional sPersCodBen As String = "", Optional nTipoREU As Integer = 1, Optional nMontoAcu As Double = 0, Optional sOrigen As String = "")

    sSql = "INSERT MovLavDinero (nMovNro,cPersCod,cPersCodBen,nTipoReu,nMontoAcum,cOrigenEfectivo) " _
    & "VALUES (" & nMovNro & ",'" & sPersCod & "','" & IIf(sPersCodBen = "", sPersCod, sPersCodBen) & "'," & nTipoREU & ", " & nMontoAcu & ",'" & sOrigen & "')"
    
dbCmact.Execute sSql

End Sub
'*****************************************************************************

Public Sub AgregaMovTipoCambio(ByVal nMovNro As Long, ByVal pnTipoCambio As Currency)
sSql = " INSERT MovTpoCambio (nMovNro, nMovTpoCambio) " _
     & " VALUES (" & nMovNro & "," & pnTipoCambio & ")"
dbCmact.Execute sSql
End Sub

Public Sub AgregaMovCap(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double)
sSql = "INSERT MovCap (nMovNro,cOpeCod,cCtaCod,nMonto,nSaldoDisponible,nSaldoContable) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nMonto & "," & nSaldoDisp & "," & nSaldoCnt & ")"
dbCmact.Execute sSql
End Sub

Public Sub AgregaMovCapDet(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nConcepto As CaptacConcepto, ByVal nMonto As Double)
        
sSql = "INSERT MovCapDet (nMovNro,cOpeCod,cCtaCod,nConceptoCod,nMonto) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nConcepto & "," & nMonto & ")"
dbCmact.Execute sSql
End Sub

Public Sub AgregaCtaExoneradaITF(ByVal sCuenta As String, ByVal TpoExoneracionITF As Integer, ByVal cMovNro As String)
        
sSql = "Insert ITFExoneracionCta(cCtaCod, nExoTpo, cRegistro, cExtorno) " _
    & "VALUES ('" & sCuenta & "'," & TpoExoneracionITF & ",'" & cMovNro & "',Null)"
dbCmact.Execute sSql
        
sSql = "Insert ITFExoneracionCtaDet(cMovNro, cCtaCod, nExoTpo) " _
    & "VALUES ('" & cMovNro & "','" & sCuenta & "'," & TpoExoneracionITF & ")"
dbCmact.Execute sSql

End Sub

Public Sub ActualizaMovPendientesRend(ByVal nMovNro As Long, ByVal pnMonto As Currency)
    sSql = " UPDATE MovPendientesRend " _
         & " Set nSaldo = nSaldo - " & pnMonto & "" _
         & " Where nMovNro = " & nMovNro & ""
    dbCmact.Execute sSql
End Sub

Public Sub AgregaMovCMAC(ByVal nMovNro As Long, ByVal sPersCod As String, _
        ByVal nMoneda As Moneda, ByVal nOperacion As CaptacOperacion, _
        ByVal nMonto As Double, Optional sDocumento As String = "NULL", _
        Optional sCuenta As String = "NULL", Optional sCliente As String = "")
Dim sIFTipo As String
Dim sCta As String, sDoc As String
'sCta = "NULL"
'sDoc = "NULL"
'If sCuenta = "NULL" Then sCuenta = ""
'If sDocumento = "NULL" Then sDocumento = ""
sIFTipo = Format$(gTpoIFCmac, "00")
If sCuenta <> "NULL" Then sCta = sCuenta
If sDocumento <> "NULL" Then sDoc = sDocumento
sSql = "INSERT MovCMAC (nMovNro,cPersCod,cIFTpo,nMoneda,cCtaIF,cDocumento,cOpeCod,nMonto, cCliente) " _
    & "VALUES (" & nMovNro & ",'" & sPersCod & "','" & sIFTipo & "'," & nMoneda & ",'" & sCta & "','" & sDoc & "','" & Trim(nOperacion) & "'," & Trim(nMonto) & ", '" & sCliente & "')"
dbCmact.Execute sSql
End Sub

Public Sub AgregaMovParamCTS(ByVal nMovNro As Long, ByVal nPorcentaje As Double, Optional ByVal nPorInta As Double = 0, Optional ByVal nPorDisp As Double = 0)

If nPorInta = 0 And nPorDisp = 0 Then
    sSql = "INSERT MovParamCTS (nMovNro,nPorcentajeAbono) " _
         & "VALUES (" & nMovNro & "," & nPorcentaje & ")"
    dbCmact.Execute sSql
ElseIf (nPorInta > 0 Or nPorDisp > 0) And nPorcentaje = 0 Then
    sSql = "INSERT MovParamCTS (nMovNro,nPorcentajeAbono,nPorcentajeInta,nPorcentajeDisp) " _
           & "VALUES (" & nMovNro & ",0," & nPorInta & "," & nPorDisp & ")"
    dbCmact.Execute sSql

End If
 
End Sub

Public Sub AgregaMovRef(ByVal nMovNro As Long, ByVal nMovRef As Long)
sSql = "INSERT MovRef (nMovNro,nMovNroRef) " _
    & "VALUES (" & nMovNro & "," & nMovRef & ")"
dbCmact.Execute sSql
End Sub


Public Function GetEstadoCuenta(ByVal sCuenta As String) As CaptacEstado
Dim rsEst As Recordset
Set rsEst = New Recordset
sSql = "Select nPrdEstado FROM Producto Where cCtaCod = '" & sCuenta & "'"
rsEst.CursorLocation = adUseClient
rsEst.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsEst.ActiveConnection = Nothing
If rsEst.EOF And rsEst.BOF Then
    GetEstadoCuenta = -1
Else
    GetEstadoCuenta = rsEst("nPrdEstado")
End If
rsEst.Close
Set rsEst = Nothing
End Function

Public Function GetCuentaBloqueos(ByVal sCuenta As String, ByVal nTipoBloqueo As CaptacTipoBloqueo) As Recordset
Dim rsEst As Recordset
Set rsEst = New Recordset
sSql = "Select K.cConsDescripcion Motivo, B.nBlqMotivo FROM ProductoBloqueos B INNER JOIN " _
    & sDBComunes & "Constante K ON B.nBlqMotivo = K.nConsValor WHERE " _
    & "B.cMovNroDbl IS NULL AND B.cCtaCod = '" & sCuenta & "'"
If nTipoBloqueo = gCapTpoBlqRetiro Then
    sSql = sSql & "And B.nBlqTpo = " & gCapTpoBlqRetiro & " And K.nConsCod = " & gCaptacMotBloqueoRet
ElseIf nTipoBloqueo = gCapTpoBlqTotal Then
    sSql = sSql & "And B.nBlqTpo = " & gCapTpoBlqTotal & " And K.nConsCod = " & gCaptacMotBloqueoTot
End If
rsEst.CursorLocation = adUseClient
rsEst.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsEst.ActiveConnection = Nothing
Set GetCuentaBloqueos = rsEst
Set rsEst = Nothing
End Function

Public Sub ActualizaSaldoRetiroCTS(ByVal sCuenta As String, ByVal nSaldoRet As Double, nIntSaldo As Double)
sSql = "Update CaptacCTS Set nSaldRetiro = " & nSaldoRet & ", nIntSaldo = " & nIntSaldo & " " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaSaldoAnteriorAho(ByVal sCuenta As String, ByVal nSaldoAnt As Double)
    sSql = "Update CaptacAhorros Set nSaldoAnterior = " & nSaldoAnt & "  WHERE cCtaCod = '" & sCuenta & "'"
    dbCmact.Execute sSql
End Sub

Public Sub ActualizadUltContactoAho(ByVal sCuenta As String, ByVal dFecha As Date)
    sSql = "Update CaptacAhorros Set dUltContacto = '" & Format(dFecha, gsFormatoFecha) & "'  WHERE cCtaCod = '" & sCuenta & "'"
    dbCmact.Execute sSql
End Sub

Public Sub ActualizaInactivaAct(ByVal sCuenta As String)
    sSql = "Update CaptacAhorros Set bInactiva = 0 WHERE cCtaCod = '" & sCuenta & "'"
    dbCmact.Execute sSql
End Sub

Public Sub ActualizaInmovilizadaAct(ByVal sCuenta As String)
    sSql = "Update CaptacAhorros Set bInactiva = 0,  bInmovilizada = 0 WHERE cCtaCod = '" & sCuenta & "'"
    dbCmact.Execute sSql
End Sub

Public Sub ActualizaAbonoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)

If bActExtracto Then
    sSql = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSql = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
dbCmact.Execute sSql
sSql = "Update Captaciones Set nSaldoDisp = nSaldoDisp + " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaCargoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)

If bActExtracto Then
    sSql = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSql = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
dbCmact.Execute sSql


    sSql = "Update Captaciones Set nSaldoDisp = nSaldoDisp - " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
         & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
         & " WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaCargoCTS(ByVal sCuenta As String, nMonto As Double, nIntSaldo As Double)
sSql = "Update CaptacCTS Set nSaldRetiro = nSaldRetiro - " & nMonto & ", " _
   & "nIntSaldo = nIntSaldo + " & nIntSaldo & " WHERE cCtaCod = '" & sCuenta & "'"
 
'****modificaciones SOLO PARA CUADRAR
'sSql = "Update CaptacCTS Set nSaldRetiro =  case when nIntSaldo-" & nMonto & "<0 then nSaldRetiro+(nIntSaldo-" & nMonto & ") else nSaldRetiro end , " _
'   & " nIntSaldo = case when nIntSaldo-" & nMonto & "<0 then 0 else nIntSaldo-" & nMonto & " end + " & nIntSaldo & "  WHERE cCtaCod = '" & sCuenta & "'"
 
 
'  nint=case when @nint-@nmonto<0 then 0 else @nint-@nmonto end,
'  nsaldretiro=case when @nint-@nmonto<0 then @nsaldoretiro+(@nint-@nmonto) else @nsaldoretiro end
dbCmact.Execute sSql
End Sub



Public Sub ActualizaRetiroInteresPF(ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nIntGanado As Double, ByVal dUltMov As Date, ByVal sMovNro As String, _
            ByVal dAuxiliar As Date, Optional bActExtracto As Boolean = True)
If bActExtracto Then
    sSql = "Update Producto Set nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
    dbCmact.Execute sSql
End If
sSql = "Update CaptacPlazoFijo Set dUltCierreAnt = C.dUltCierre, " _
    & "dAuxiliarAnt = PF.dAuxiliar FROM Captaciones C INNER JOIN CaptacPlazoFijo PF " _
    & "ON C.cCtaCod = PF.cCtaCod WHERE C.cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
sSql = "Update Captaciones Set  nIntAcum = nIntAcum + " & nIntGanado - nMonto _
    & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
sSql = "Update CaptacPlazoFijo Set nIntPag = nIntPag + " & nMonto & " " _
    & ", dAuxiliar = '" & Format$(dAuxiliar, "mm/dd/yyyy") & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaExtornoRetiroInteresPF(ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nIntGanado As Double, ByVal dUltMov As Date, ByVal sMovNro As String, _
            ByVal dAuxiliar As Date)
sSql = "Update Captaciones Set  nIntAcum = nIntAcum - " & nIntGanado & " + " & nMonto _
    & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
sSql = "Update CaptacPlazoFijo Set nIntPag = nIntPag - " & nMonto & " " _
    & ", dAuxiliar = '" & Format$(dAuxiliar, "mm/dd/yyyy") & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Function TieneChequesValorizacion(ByVal sCuenta As String) As Boolean
Dim rsChq As Recordset
sSql = "Select C.cCtaCod, D.nTpoDoc, D.cNroDoc, D.cPersCod, P.cPersNombre " _
    & "FROM Captaciones C INNER JOIN DocRecCapta D INNER JOIN DocRecEst E " _
    & "INNER JOIN " & sDBPersona & " Persona P ON E.cPersCod = P.cPersCod ON " _
    & "D.nTpoDoc = E.nTpoDoc AND D.cNroDoc = E.cNroDoc AND D.cPersCod = " _
    & "E.cPersCod AND D.cIFTpo = E.cIFTpo ON C.cCtaCod = D.cCtaCod WHERE C.cCtaCod = '" & sCuenta & "' " _
    & "AND E.cMovNro = (Select MAX(cMovNro) FROM DocRecEst E1 WHERE E1.nTpoDoc " _
    & "= E.nTpoDoc AND E1.cNroDoc = E.cNroDoc AND E1.cPersCod = E.cPersCod) AND " _
    & "E.nEstado = " & gChqEstEnValorizacion
Set rsChq = New Recordset
rsChq.CursorLocation = adUseClient
rsChq.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsChq.ActiveConnection = Nothing
If Not (rsChq.EOF And rsChq.BOF) Then
    TieneChequesValorizacion = True
Else
    TieneChequesValorizacion = False
End If
rsChq.Close
Set rsChq = Nothing
End Function

Public Sub UltimaActualizacionCuenta(ByVal sCuenta As String, ByVal sMovNro As String)
'Actualiza la fecha de ultima actualizacion de la cuenta de captaciones
sSql = "Update Captaciones Set cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaEstadoCuenta(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, _
        ByVal dFecha As Date, ByVal sMovNro As String)
'Actualiza el ultimo estado a la tabla de producto
sSql = "Update Producto Set nPrdEstado = " & nEstado & ", " _
    & "dPrdEstado = '" & Format$(dFecha, "mm/dd/yyyy") & "' WHERE " _
    & "cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
'Agrega un registro mas de estado a la historia de estados de la cuenta
sSql = "INSERT CaptacEstado (cCtaCod,nPrdEstado,cMovNro) " _
    & "VALUES ('" & sCuenta & "'," & nEstado & ",'" & sMovNro & "')"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaEstadoGiro(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, ByVal dFecha As Date)
'Actualiza el ultimo estado a la tabla de producto
sSql = "Update Producto Set nPrdEstado = " & nEstado & ", " _
    & "dPrdEstado = '" & Format$(dFecha, "mm/dd/yyyy") & "' WHERE " _
    & "cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaEstadoDocRecEst(ByVal nTipoDoc As TpoDoc, ByVal sNroDoc As String, _
        ByVal sCodIF As String, ByVal sMovNro As String, ByVal nEstado As ChequeEstado, _
        Optional nIFTpo As CGTipoIF = gTpoIFBanco, Optional sCuenta As String)
        
'elimina docreccapta
sSql = " delete docreccapta  WHERE " _
    & "nTpoDoc = '" & nTipoDoc & "' AND cNroDoc = '" & sNroDoc & "' AND cPersCod = '" & sCodIF & "' and cctacod='" & sCuenta & "'"
dbCmact.Execute sSql
'Actualiza el ultimo estado a la tabla de Documento Recibidos
sSql = "Update DocRec Set nEstado = " & nEstado & " WHERE " _
    & "nTpoDoc = '" & nTipoDoc & "' AND cNroDoc = '" & sNroDoc & "' AND cPersCod = '" & sCodIF & "'"
dbCmact.Execute sSql
'Agrega un registro mas de estado a la historia de estados del documento recibido
sSql = "INSERT DocRecEst (nTpoDoc,cNroDoc,cPersCod,cMovNro,nEstado,cIFTpo) " _
    & "VALUES ('" & nTipoDoc & "','" & sNroDoc & "','" & sCodIF & "','" & sMovNro & "'," & nEstado & ",'" & Format$(nIFTpo, "00") & "')"
dbCmact.Execute sSql
End Sub

Public Function GetMovExtorno(ByVal sDatoBus As String, ByVal dFecha As Date, _
        ByVal nOperacion As CaptacOperacion, psAgecod As String, Optional nTipoBus As Integer = 0) As Recordset
Dim rsMov As Recordset
Dim sWhere As String
If nTipoBus = 0 Then
    sWhere = "M.cMovNro LIKE '%" & sDatoBus & "%'"
ElseIf nTipoBus = 1 Then
    sWhere = "M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And C.cCtaCod = '" & sDatoBus & "'"
End If
sSql = "Select Distinct M.cMovNro, O.cOpeDesc, C.cCtaCod, ISNULL(MD.nDocTpo,'') nDocTpo, ISNULL(MD.cDocNro,'') cDocNro, " _
    & " M.cMovDesc, C.cOpeCod, C.nMonto, M.nMovNro, cPersCod=isnull((select cperscod from docreccapta where cctacod=c.cctacod and nmovnro=c.nmovnro ),0), " _
    & " IsNull(CD.nMonto,0) ITFCargo, ISNULL(CD.cOpeCod,'') ITFOperacion, ISNULL(CD.nConceptoCod,'') ITFConcepto FROM " _
    & "(Select MD.nMovNro, MD.nDocTpo, MD.cDocNro From DocRecop D " _
    & "        RIGHT JOIN MovDoc MD ON d.nTpoDoc = MD.nDocTpo And D.cNroDoc = MD.cDocNro) MD RIGHT JOIN Mov M JOIN MovCap C INNER JOIN OpeTpo O ON C.cOpeCod = O.cOpeCod ON M.nMovNro = C.nMovNro " _
    & "ON MD.nMovNro = M.nMovNro Left Join MovCapDet CD On C.nMovNro = CD.nMovNro And C.cCtaCod = CD.cCtaCod And " _
    & "CD.cOpeCod IN ('" & gITFCobroCargo & "','" & gITFCobroEfectivo & "','" & gITFCobroCMAC & "','" & gITFCobroCMACCargo & "') And CD.nConceptoCod IN " _
    & "(" & gConcITFCliente & "," & gConcITFAsumido & ") WHERE " & sWhere & " AND C.cOpeCod  IN ('" & nOperacion & "') " _
    & "AND M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & "," & gMovFlagDeExtorno & ")" _
    & "AND M.nMovEstado IN (" & gMovEstContabMovContable & ") And SubString(M.cMovNro,18,2) = '" & psAgecod & "' " _
    & "ORDER BY M.nMovNro"
Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetMovExtorno = rsMov
End Function

Public Function GetMovExtornoGiro(ByVal sDatoBus As String, ByVal dFecha As Date, _
        ByVal nOperacion As CaptacOperacion, psAgecod As String, Optional nTipoBus As Integer = 0) As Recordset
Dim rsMov As Recordset
Dim sWhere As String
If nTipoBus = 0 Then
    sWhere = "M.cMovNro LIKE '%" & sDatoBus & "%'"
ElseIf nTipoBus = 1 Then
    sWhere = "M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And C.cNumRecibo = '" & sDatoBus & "'"
End If
sSql = "Select distinct M.cMovNro, O.cOpeDesc, C.cNumRecibo cCtaCod, nDocTpo = '', cDocNro = '', M.cMovDesc, M.cOpeCod, "
sSql = sSql & " C.nMonto, M.nMovNro, PP.cPersCod cPersCod "

'Agregado P
sSql = sSql & ", IsNull(CD.nMonto,0) ITFCargo, ISNULL(CD.cOpeCod,'') ITFOperacion, ISNULL(CD.nPrdConcepto,'') ITFConcepto "
' Fin Agregado

sSql = sSql & " FROM OpeTpo O INNER JOIN Mov M JOIN MovServicios C "
sSql = sSql & " JOIN ProductoPersona PP ON C.cNumRecibo = PP.cCtaCod ON M.nMovNro = C.nMovNro ON O.cOpeCod = M.cOpeCod "

'Agregado P
sSql = sSql & " Left Join MovServiciosDet CD On   CD.nMovNro=C.nMovNro  And  CD.cNumRecibo=C.cNumRecibo  "
sSql = sSql & " And CD.cOpeCod IN ('990101','990102','990302') "
sSql = sSql & " And CD.nPrdConcepto IN (20,21) "
' Fin Agregado

sSql = sSql & " WHERE " & sWhere & " AND M.cOpeCod ='" & nOperacion & "' and C.copecod not like '99%'  AND PP.nPrdPersRelac = " & gGiroRelPersRemitente & " "
sSql = sSql & " AND M.nMovFlag=" & gMovFlagVigente
sSql = sSql & " AND M.nMovEstado=" & gMovEstContabMovContable & "  And SubString(M.cMovNro,18,2) = '" & psAgecod & "' "
sSql = sSql & " ORDER BY M.nMovNro "

Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetMovExtornoGiro = rsMov
End Function

Public Function GetITFMOV(ByVal nMovNroBus As Long) As Double
Dim RSTEMP As Recordset
    
    sSql = "Select nmonto from movcapdet where nconceptocod=20 and nmovnro=" & nMovNroBus
    Set RSTEMP = New Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set RSTEMP.ActiveConnection = Nothing
   GetITFMOV = IIf(Not RSTEMP.EOF, RSTEMP!nMonto, 0)
End Function

Public Function GetIntaRetMov(ByVal nMovNroBus As Long) As Double
Dim RSTEMP As Recordset
    
    sSql = "Select nmonto from movcapdet where nconceptocod=30 and nmovnro=" & nMovNroBus
    Set RSTEMP = New Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set RSTEMP.ActiveConnection = Nothing
   GetIntaRetMov = IIf(Not RSTEMP.EOF, RSTEMP!nMonto, 0)


End Function

Public Function GetMovExtornoCMAC(ByVal sDatoBus As String, ByVal dFecha As Date, _
        ByVal nOperacion As CaptacOperacion, Optional nTipoBus As Integer = 0) As Recordset
Dim rsMov As Recordset
Dim sWhere As String
If nTipoBus = 0 And Trim(sDatoBus) <> "" Then
    'sWhere = "M.cMovNro LIKE '%" & sDatoBus & "%'"
    sWhere = " RIGHT(M.cMovNro,4)= '" & sDatoBus & "' and "
ElseIf nTipoBus = 1 And Trim(sDatoBus) <> "" Then
    sWhere = "M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And MC.CCTAIF  = '" & sDatoBus & "' and "
End If
sSql = "Select cperscod='',M.cMovNro, O.cOpeDesc, ISNULL(MC.cCtaIF,'') cCtaCod, ISNULL(MC.cDocumento,'') cDocNro, " _
    & "M.cMovDesc, MC.cOpeCod, MC.nMonto, M.nMovNro, ISNULL(OD.nDocTpo,'') nDocTpo " _
    & "FROM Mov M INNER JOIN MovCMAC MC INNER JOIN " & sDBComunes & "OpeTpo O LEFT JOIN " _
    & sDBComunes & "OpeDoc OD ON O.cOpeCod = OD.cOpeCod ON MC.cOpeCod = O.cOpeCod ON " _
    & "M.nMovNro = MC.nMovNro WHERE " & sWhere & "  M.nMovFlag NOT IN " _
    & "(" & gMovFlagExtornado & "," & gMovFlagEliminado & "," & gMovFlagDeExtorno & ")" _
    & "AND M.nMovEstado IN (" & gMovEstContabMovContable & ") And M.cOpeCod = '" & nOperacion & "' " _
    & "ORDER BY M.nMovNro"
Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetMovExtornoCMAC = rsMov
End Function
Public Sub ActualizaEstadoMov(ByVal nMovNro As Long, ByVal nMovFlag As MovFlag)
sSql = "Update Mov Set nMovFlag = " & nMovFlag & " Where nMovNro = " & nMovNro
dbCmact.Execute sSql
End Sub
Public Function MovAsociados(ByVal nMovNro As Long) As Recordset
Dim RSTEMP As Recordset
sSql = "SELECT NMOVNROREF FROM MOVREF WHERE NMOVNRO=" & nMovNro
Set RSTEMP = New Recordset
RSTEMP.CursorLocation = adUseClient
RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set RSTEMP.ActiveConnection = Nothing
Set MovAsociados = RSTEMP
Set RSTEMP = Nothing
End Function


Public Sub ActualizaDuplicadoCertPF(ByVal sCuenta As String)
sSql = "Update CaptacPlazoFijo Set nDuplicado = nDuplicado + 1 Where cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Function GetMovTransferencia(ByVal nMovNro As Long) As Recordset
Dim rsTrans As Recordset
sSql = "SELECT M.nMovNro, C.cOpeCod, C.cCtaCod, C.nMonto From Mov M INNER JOIN " _
    & "MovCap C ON M.nMovNro = C.nMovNro Where M.nMovNro = " & nMovNro & " and c.copecod not like '99%' " _
    & "ORDER BY M.cOpeCod DESC"
Set rsTrans = New Recordset
rsTrans.CursorLocation = adUseClient
rsTrans.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTrans.ActiveConnection = Nothing
Set GetMovTransferencia = rsTrans
Set rsTrans = Nothing

End Function

Public Function GetMovTransferenciaITF(ByVal nMovNro As Long, ByVal sCuenta As String) As Recordset
Dim rsTrans As Recordset
sSql = " SELECT M.nMovNro, C.cOpeCod, C.cCtaCod, C.nMonto,c.nconceptocod From Mov M INNER JOIN " _
    & " MovCapdet C ON M.nMovNro = C.nMovNro Where M.nMovNro = " & nMovNro & " and c.CCTACOD='" & sCuenta & "' and c.copecod like '99%'"
    
Set rsTrans = New Recordset
rsTrans.CursorLocation = adUseClient
rsTrans.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTrans.ActiveConnection = Nothing
Set GetMovTransferenciaITF = rsTrans
Set rsTrans = Nothing

End Function

Public Function GetFechaUltimoRetiroIntPF(ByVal sCuenta As String) As Date
Dim rsMov As Recordset
sSql = "Select TOP 1 PF.dRenovacion, M.cMovNro FROM CaptacPlazoFijo PF LEFT JOIN " _
    & "MovCap C INNER JOIN Mov M ON C.nMovNro = M.nMovNro ON PF.cCtaCod = C.cCtaCod " _
    & "WHERE PF.cCtaCod = '" & sCuenta & "' OR (C.cOpeCod IN ('" & gPFRetInt & "','" & gPFRetIntAboAho & "') " _
    & "AND M.nMovFlag NOT IN ('" & gMovFlagExtornado & "') AND PF.cCtaCod = '" & sCuenta & "') " _
    & "ORDER BY M.nMovNro "
Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
If Not (rsMov.EOF And rsMov.BOF) Then
    If IsNull(rsMov("cMovNro")) Then
        GetFechaUltimoRetiroIntPF = rsMov("dRenovacion")
    Else
        GetFechaUltimoRetiroIntPF = CDate(Mid(rsMov("cMovNro"), 7, 2) & "/" & Mid(rsMov("cMovNro"), 5, 2) & "/" & Mid(rsMov("cMovNro"), 1, 4))
    End If
End If
rsMov.Close
Set rsMov = Nothing
End Function

Public Function GetFechaPenultimoRetiroIntPF(ByVal sCuenta As String) As Date
Dim rsMov As Recordset
sSql = "Select TOP 2 PF.dRenovacion, M.cMovNro FROM CaptacPlazoFijo PF LEFT JOIN " _
    & "MovCap C INNER JOIN Mov M ON C.nMovNro = M.nMovNro ON PF.cCtaCod = C.cCtaCod " _
    & "WHERE PF.cCtaCod = '" & sCuenta & "' OR (C.cOpeCod IN ('" & gPFRetInt & "','" & gPFRetIntAboAho & "') " _
    & "AND M.nMovFlag NOT IN ('" & gMovFlagExtornado & "') AND PF.cCtaCod = '" & sCuenta & "') " _
    & "ORDER BY M.nMovNro "
Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
If Not (rsMov.EOF And rsMov.BOF) Then
    If IsNull(rsMov("cMovNro")) Then
        GetFechaPenultimoRetiroIntPF = rsMov("dRenovacion")
    Else
        GetFechaPenultimoRetiroIntPF = CDate(Mid(rsMov("cMovNro"), 7, 2) & "/" & Mid(rsMov("cMovNro"), 5, 2) & "/" & Mid(rsMov("cMovNro"), 1, 4))
    End If
End If
rsMov.Close
Set rsMov = Nothing
End Function


Public Function GetFechaValorChqApePF(ByVal sCuenta As String) As String
Dim rsChq As Recordset
sSql = "SELECT C.dValorizacion FROM DocRecCapta A INNER JOIN DocRec C ON " _
    & "A.nTpoDoc = C.nTpoDoc AND A.cNroDoc = C.cNroDoc AND A.cPersCod = C.cPersCod WHERE " _
    & "C.nEstado = " & gChqEstValorizado & " AND A.cCtaCod = '" & sCuenta & "'"
Set rsChq = New Recordset
rsChq.CursorLocation = adUseClient
rsChq.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsChq.ActiveConnection = Nothing
If rsChq.EOF And rsChq.BOF Then
    GetFechaValorChqApePF = ""
Else
    GetFechaValorChqApePF = Format$(rsChq("dValorizacion"), "dd/mm/yyyy")
End If
rsChq.Close
Set rsChq = Nothing
End Function

Public Function GetCuentaAbonoIntPF(ByVal sCuenta As String) As String
Dim rsCta As Recordset
sSql = "SELECT cCtaCodAbono FROM CaptacAboIntPF WHERE cCtaCod = '" & sCuenta & "'"
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
rsCta.CursorLocation = adUseClient
If rsCta.EOF And rsCta.BOF Then
    GetCuentaAbonoIntPF = ""
Else
    GetCuentaAbonoIntPF = rsCta("cCtaCodAbono")
End If
rsCta.Close
Set rsCta = Nothing
End Function

Public Function EsOrdenPagoEmitida(ByVal sCuenta As String, ByVal nNumOP As Long) As Boolean
Dim rsOP As Recordset
Set rsOP = New Recordset
rsOP.CursorLocation = adUseClient
sSql = "Select cCtaCod, nInicio, nFin FROM CapOrdPagEmision " _
    & "WHERE cCtaCod = '" & sCuenta & "' And " & nNumOP & " Between nInicio And nFin " _
    & "And nEstado = " & gCapTalOrdPagEstEntregado
rsOP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsOP.ActiveConnection = Nothing
If rsOP.EOF And rsOP.BOF Then
    EsOrdenPagoEmitida = False
Else
    EsOrdenPagoEmitida = True
End If
rsOP.Close
Set rsOP = Nothing
End Function

Public Function GetCapMov(ByVal nMovNro As Long, ByVal nConcepto As CaptacConcepto, _
    ByVal nProducto As Producto) As Recordset
Dim rsMov As Recordset
sSql = "Select M.cMovNro, MC.nMovNro, MC.cOpeCod, MC.cCtaCod, MC.nMonto, MC.nSaldoDisponible, " _
    & "MC.nSaldoContable, MD.nMonto nMontoConcepto FROM Mov M INNER JOIN MovCap MC INNER JOIN " _
    & "MovCapDet MD ON MC.nMovNro = MD.nMovNro AND MC.cCtaCod = MD.cCtaCod ON M.nMovNro = MC.nMovNro " _
    & "WHERE M.nMovNro = " & nMovNro & " AND MD.nConceptoCod = " & nConcepto & " AND " _
    & "SUBSTRING(MC.cCtaCod,6,3) = '" & nProducto & "'"
Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetCapMov = rsMov
Set rsMov = Nothing
End Function

Public Function GetCapMovCMAC(ByVal nMovNro As Long) As Recordset
Dim rsMov As Recordset
sSql = "Select M.cMovNro, M.nMovNro, MC.cOpeCod, MC.cCtaIF cCtaCod, MC.nMonto, " _
    & "MC.nMoneda, MC.cDocumento FROM Mov M INNER JOIN MovCMAC MC ON " _
    & "M.nMovNro = MC.nMovNro WHERE M.nMovNro = " & nMovNro
Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetCapMovCMAC = rsMov
Set rsMov = Nothing
End Function

Public Function GetCapMovParamCTS(ByVal nMovNro As Long) As Double
Dim rsCTS As Recordset
sSql = "Select nPorcentajeAbono From MovParamCTS Where nMovNro = " & nMovNro
Set rsCTS = New Recordset
rsCTS.CursorLocation = adUseClient
rsCTS.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCTS.ActiveConnection = Nothing
If rsCTS.EOF And rsCTS.BOF Then
    GetCapMovParamCTS = 0
Else
    GetCapMovParamCTS = rsCTS("nPorcentajeAbono")
End If
rsCTS.Close
Set rsCTS = Nothing
End Function

Public Sub AgregaMovTranferBanco(ByVal pnMovNro As Long, ByVal psPersCodIF As String, psCtaBco As String, ByVal pnMoneda As Moneda)
sSql = " Insert MovTransferBco (nMovNro, cPersCodIF, cCtaBco, nMoneda)" _
     & " Values (" & pnMovNro & ",'" & psPersCodIF & "','" & psCtaBco & "'," & pnMoneda & ")"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaEstadoProducto(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, _
        ByVal dFecha As Date)
'Actualiza el ultimo estado a la tabla de producto
sSql = "Update Producto Set nPrdEstado = " & nEstado & ", " _
    & "dPrdEstado = '" & Format$(dFecha & " " & Time, "mm/dd/yyyy hh:mm:ss") & "' WHERE " _
    & "cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub



Private Sub Class_Initialize()
Dim sConn As String
Dim ClsIni As COMConecta.DCOMClasIni
Set ClsIni = New COMConecta.DCOMClasIni
sConn = ClsIni.CadenaConexion
sDBComunes = ClsIni.BaseComunes
sDBPersona = ClsIni.BasePersonas
sDBImagenes = ClsIni.BaseImagenes
Set ClsIni = Nothing
Set dbCmact = New Connection
dbCmact.CursorLocation = adUseClient
dbCmact.Open sConn
dbCmact.Execute "SET DATEFORMAT MDY"
bTransaccion = False
End Sub

Private Sub Class_Terminate()
If bTransaccion = True Then
Else
    dbCmact.Close
    Set dbCmact = Nothing
End If
End Sub

Public Function CtaConFirmas(psctacod As String) As Boolean
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim clsCon As DConecta
    Set clsCon = New DConecta

    sql = " Select PP.cCtaCod From ProductoPersona PP" _
        & " Left Join PersImagen PI On PP.cPersCod = PI.cPersCod" _
        & " Where PP.cCtaCod = '" & psctacod & "' And PI.iPersFirma IS Null"
    clsCon.AbreConexion
    
    Set rs = clsCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        CtaConFirmas = True
    Else
        CtaConFirmas = False
    End If
    
    Set rs = Nothing
    Set clsCon = Nothing
End Function

Public Sub AgregaMovOpeVarias(ByVal pnMovNro As Long, ByVal pnMovImporte As Double, _
        ByVal psNroDoc As String, ByVal psReferencia As String, ByVal pnMoneda As Moneda, _
        Optional nOperacion As CaptacOperacion)
    Dim sql As String
    
    sql = " Insert MovOpeVarias (nMovNro,nMovImporte,cNroDoc,cReferencia,nMoneda,cOpeCod)" _
        & " Values(" & pnMovNro & "," & pnMovImporte & ",'" & psNroDoc & "','" & psReferencia & "'," & pnMoneda & ",'" & Trim(nOperacion) & "')"
    
    dbCmact.Execute sql
End Sub

Public Sub AgregaMovGasto(pnMovNro As Long, psPersCod As String, psPersCodDestino As String)
    Dim sql As String
    
    sql = " Insert MovGasto (nMovNro,cPersCod,cPersCodDest)" _
        & " Values(" & pnMovNro & ",'" & psPersCod & "','" & psPersCodDestino & "')"
    
    dbCmact.Execute sql
End Sub

Public Function GetTarifa(psOpeCod As CaptacOperacion, pnMoneda As Moneda) As Currency
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select nParValor, nMoneda from Parametro PA" _
        & " Inner Join Tarifas TA On PA.nParCod = TA.nParCod And PA.nParProd = TA.nParProd" _
        & " Where TA.cOpeCod = '" & psOpeCod & "' And nMoneda = " & pnMoneda & ""
    rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.EOF And rs.BOF Then
        GetTarifa = 0
    Else
        GetTarifa = rs.Fields(0)
    End If
    
    Set rs = Nothing
End Function

Public Function GetOtrasOperaciones(pdFecha As Date, Optional cUser As String) As ADODB.Recordset
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim clsCon As DConecta
    Set clsCon = New DConecta
    
    clsCon.AbreConexion
    
    sql = "SELECT M.nMovNro, Convert(DateTime,Left(M.cMovNro,8)) dFecTran, M.cOpeCod, Right(M.cMovNro,4) cCodUsu, OP.cOpeDesc, MOV.nMovImporte, MOV.nMoneda " _
        & "FROM Mov M JOIN MovOpeVarias MOV On M.nMovNro = MOV.nMovNro JOIN OpeTpo OP ON " _
        & "M.cOpeCod = OP.cOpeCod WHERE M.nMovFlag = 0 And Left(M.cOpeCod,4) IN ('" & Left(gOtrOpeAhoOtrosIngresos, 4) & "','" _
        & Left(gOtrOpeAhoOtrosEgresos, 4) & "','" & Left(gOtrOpeDuplicadoTarjeta, 4) & "','" & Left(gOtrOpeDepCtaBcoEfec, 4) & "') " _
        & "And M.cMovNro LIKE '" & Format(pdFecha, gsFormatoMovFecha) & "%' and right(M.cMovNro,4)='" & cUser & "'"
    Set rs = clsCon.CargaRecordSet(sql)
    
    Set GetOtrasOperaciones = rs
    
End Function

Public Function GetOtrasOperacionesDet(psMovNro As Long) As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim clsCon As DConecta
    Set clsCon = New DConecta
    
    clsCon.AbreConexion
    
    sSql = "SELECT M.nMovNro, Convert(DateTime,Left(M.cMovNro,8)) dFecTran, M.cOpeCod, Right(M.cMovNro,4) cCodUsu, OP.cOpeDesc, MOV.nMovImporte, PE.cPersCod , PE.cPersNombre, M.cMovDesc, MOV.cNroDoc " _
        & "FROM Mov M JOIN MovOpeVarias MOV On M.nMovNro = MOV.nMovNro " _
        & "JOIN MovGasto MG On M.nMovNro = MG.nMovNro JOIN Persona PE On MG.cPersCod = PE.cPersCod " _
        & "JOIN OpeTpo OP ON M.cOpeCod = OP.cOpeCod WHERE M.nMovNro = " & psMovNro & " " _
        & "UNION " _
        & "SELECT M.nMovNro, Convert(DateTime,Left(M.cMovNro,8)) dFecTran, M.cOpeCod, Right(M.cMovNro,4) cCodUsu, OP.cOpeDesc, MOV.nMovImporte, PE.cPersCod , PE.cPersNombre, M.cMovDesc, MOV.cNroDoc " _
        & "FROM Mov M JOIN MovOpeVarias MOV On M.nMovNro = MOV.nMovNro " _
        & "JOIN Persona PE On MOV.cReferencia = PE.cPersCod OR SUBSTRING(MOV.cReferencia,4,13) = PE.cPersCod " _
        & "JOIN OpeTpo OP ON M.cOpeCod = OP.cOpeCod WHERE M.nMovNro = " & psMovNro
    Set rs = clsCon.CargaRecordSet(sSql)
    
    Set GetOtrasOperacionesDet = rs
    
    Set rs = Nothing
End Function

Public Function CargaCuentas(ByVal sPersona As String, Optional sProducto As String = "CRE") As ADODB.Recordset
    Dim rsCta As ADODB.Recordset
    Dim nCuentasAho As Long
    Dim nCuentasPrd As Long
    Dim nCuentasCre As Long
    Dim nCuentasJud As Long
    Set rsCta = New ADODB.Recordset
    Dim clsCon As DConecta
    Set clsCon = New DConecta
    Dim sql As String
    
    rsCta.CursorLocation = adUseClient
    nCuentasPrd = 0
    nCuentasCre = 0
    nCuentasJud = 0
    
    clsCon.AbreConexion
    
    Select Case sProducto
        Case "PRD"
            sql = " Select PP.cCtaCod From ProductoPersona PP" _
                & " Inner Join Captaciones CAP On PP.cCtaCod = CAP.cCtaCod" _
                & " Where PP.cPersCod = '" & sPersona & "'" _
                & " Union" _
                & " Select PP.cCtaCod From ProductoPersona PP" _
                & " Inner Join Colocaciones COL On PP.cCtaCod = COL.cCtaCod Where Substring(PP.cCtaCod,6,3) = '" & Producto.gColConsuPrendario & "'" _
                & " And PP.cPersCod = '" & sPersona & "' Order by PP.cCtaCod "
        Case "CRE"
'            sql = " Select PP.cCtaCod From ProductoPersona PP" _
'                & " Inner Join Colocaciones COL On PP.cCtaCod = COL.cCtaCod" _
'                & " Inner Join Producto PRD On PP.cCtaCod = PRD.cCtaCod" _
'                & " Where Substring(PP.cCtaCod,6,3) <> '" & Producto.gColConsuPrendario & "' And LEft(PRD.nPrdEstado,2) <> '" & Left(ColocEstado.gColocEstRecVigJud, 2) & "'" _
'                & "  And PP.cPersCod = '" & sPersona & "' Order by PP.cCtaCod "

              sql = " Select PP.cCtaCod ,Prd.nPrdEstado " _
                    & " From ProductoPersona PP " _
                    & " Inner Join Colocaciones COL On PP.cCtaCod = COL.cCtaCod " _
                    & " Inner Join Producto PRD On PP.cCtaCod = PRD.cCtaCod " _
                    & " Where Substring(PP.cCtaCod,6,3) <> '305' And LEft(PRD.nPrdEstado,2) <> '22'  And " _
                    & " PP.cPersCod = '" & sPersona & "' and PP.nPrdPersRelac=20 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2004)" _
                    & " Order by PP.cCtaCod "


        Case "JUD"
'            sql = " Select PP.cCtaCod From ProductoPersona PP" _
'                & " Inner Join Colocaciones COL On PP.cCtaCod = COL.cCtaCod" _
'                & " Inner Join Producto PRD On PP.cCtaCod = PRD.cCtaCod" _
'                & " Where Left(PRD.nPrdEstado,2) = '" & Left(ColocEstado.gColocEstRecVigJud, 2) & "'" _
'                & " And PP.cPersCod = '" & sPersona & "' Order by PP.cCtaCod "

             sql = " Select PP.cCtaCod " _
                  & " From ProductoPersona PP " _
                  & " Inner Join Colocaciones COL On PP.cCtaCod = COL.cCtaCod " _
                  & " Inner Join Producto PRD On PP.cCtaCod = PRD.cCtaCod " _
                  & " Where Left(PRD.nPrdEstado,2) = '22' And PP.nPrdPersRelac=20 and " _
                  & " PP.cPersCod = '" & sPersona & "' " _
                  & " Order by PP.cCtaCod "

    End Select
    Set rsCta = clsCon.CargaRecordSet(sql)
    
    Set CargaCuentas = rsCta
End Function

Public Function BuscaCreditosPendientesPago(ByVal psCuenta As String) As Boolean
    Dim rsCta As ADODB.Recordset
    Set rsCta = New ADODB.Recordset
    Dim clsCon As DConecta
    Set clsCon = New DConecta
    Dim sql As String
    
    clsCon.AbreConexion
    
    sql = " Select bBusCredPend From CaptacPlazoFijo Where cCtaCod = '" & psCuenta & "'"
    Set rsCta = clsCon.CargaRecordSet(sql)
    
    If (rsCta.EOF And rsCta.BOF) Or IsNull(rsCta("bBusCredPend")) Then
        BuscaCreditosPendientesPago = True
    Else
        BuscaCreditosPendientesPago = rsCta("bBusCredPend")
    End If
    rsCta.Close
    Set rsCta = Nothing
    Set clsCon = Nothing
End Function

Public Function GetFirma(psPersCod As String, Optional ByVal psCodAge As String = "") As ADODB.Recordset
    Dim sql As String
    Dim clsCon As DConecta
    Set clsCon = New DConecta
    Dim oPersona As New DPersona
    
    'clsCon.AbreConexion
    oPersona.sCodage = psCodAge
    Set GetFirma = oPersona.RecuperaFirma(psPersCod)
    
    'sql = " Select cPersCod cCodPers, iPersFirma From PersImagen Where cPersCod = '" & psPersCod & "' And iPersFirma Is Not Null"
    'Set GetFirma = clsCon.CargaRecordSet(sql)
    
    Set clsCon = Nothing
End Function

Public Function GetDatosCrePendiente(psTipo As String, pbVencidos As Boolean, psCuentas As String, pnDiasAtraso As Integer, pgdFecSis As Date) As ADODB.Recordset
    Dim sql As String
    Dim clsCon As DConecta
    Dim sVencidos As String
    Set clsCon = New DConecta
    
    clsCon.AbreConexion
    
    If psTipo = "CRE" Then
        If pbVencidos Then
            sVencidos = " nDiasAtraso > " & pnDiasAtraso & " And "
        Else
            sVencidos = ""
        End If
        
        sql = " Select cCtaCod, nDiasAtraso, cRefinan, cNota1, Convert(Varchar(10),dFecVenc,103) dFecVenc, Count(*) nNroCuo," _
            & " SUM(nMCuota) nCuota, SUM(nMMora) nMora, SUM(nGastos) nGastos From (" _
            & " Select cCtaCod, nDiasAtraso, cRefinan, cNota1, nCuota cNroCuo, dFecVenc, nMCuota, nMMora, SUM(nMGastos) nGastos" _
            & " From (" _
            & "     Select C.cCtaCod, CC.nDiasAtraso, CC.bRefCapInt cRefinan, (Select cNota1 from ColocacSaldo CS Where CS.dFecha = (Select Max(CSA.dFecha) From ColocacSaldo CSA Where CSA.cCtaCod = CS.cCtaCod) And CS.cCtaCod = C.cCtaCod) cNota1, P.nCuota,nMCuota = P.Monto, nMMora = ISNULL(PM.Monto,0),nMGastos = ISNULL(PG.Monto,0)," _
            & "     dFecVenc = (Select Min(dVenc)" _
            & "     From ColocCalendario P1" _
            & "     Where P1.cCtaCod = P.cCtaCod And P.nNroCalen = P1.nNroCalen And P1.nColocCalendEstado = " & ColocCalendEstado.gColocCalendEstadoPendiente & " And P1.nColocCalendApl = " & ColocCalendApl.gColocCalendAplCuota & " And DateDiff(dd,P1.dVenc,'" & Format$(pgdFecSis, gsFormatoFecha) & "') > 0)" _
            & "   From Colocaciones C Inner Join Producto PRD On C.cCtaCod = PRD.cCtaCod Inner Join ColocacCred CC On C.cCtaCod = CC.cCtaCod" _
            & "     Inner Join" _
            & "         (Select CD.cCtaCod, CD.nNroCalen, CD.nColocCalendApl, CD.nCuota, CCAL.nColocCalendEstado, CCAL.dVenc , Sum(nMonto- nMontoPagado) Monto From ColocCalendDet CD" _
            & "             Inner Join ColocCalendario CCAL On CD.cCtaCod = CCAL.cCtaCod And CD.nNroCalen = CCAL.nNroCalen And CD.nColocCalendApl = CCAL.nColocCalendApl And CD.nCuota = CCAL.nCuota" _
            & "           Where CD.nNroCalen = (Select Max(CA.nNroCalen) From ColocCalendario CA Where CD.cCtaCod = CA.cCtaCod)" _
            & "                 And Left(CD.nPrdConceptoCod,2) in  (" & Left(ColocConcepto.gColocConceptoCodInteresMoratorio, 2) & "," & Left(ColocConcepto.gColocConceptoCodCapital, 2) & ") And CD.nPrdConceptoCod <> " & ColocConcepto.gColocConceptoCodInteresMoratorio & " And cFlag = ''" _
            & "         Group By CD.cCtaCod, CD.nNroCalen, CD.nColocCalendApl, CD.nCuota, CCAL.nColocCalendEstado, CCAL.dVenc) As P ON C.cCtaCod = P.cCtaCod" _
            & "     Left Join (Select CD.cCtaCod, CD.nNroCalen, CD.nColocCalendApl, CD.nCuota, Sum(nMonto-nMontoPagado) Monto From ColocCalendDet CD" _
            & "                 Where CD.nNroCalen = (Select Max(CA.nNroCalen) From ColocCalendario CA Where CD.cCtaCod = CA.cCtaCod)" _
            & "                 And CD.nPrdConceptoCod = " & ColocConcepto.gColocConceptoCodInteresMoratorio & " And cFlag = ''" _
            & "                 Group By cCtaCod, nNroCalen, nColocCalendApl, nCuota) As PM On PM.cCtaCod = C.cCtaCod And PM.nNroCalen = P.nNroCalen And P.nCuota = PM.nCuota" _
            & "     Left Join (Select CD.cCtaCod, CD.nNroCalen, CD.nColocCalendApl, CD.nCuota, Sum(nMonto- nMontoPagado) Monto From ColocCalendDet CD" _
            & "                 Where CD.nNroCalen = (Select Max(CA.nNroCalen) From ColocCalendario CA Where CD.cCtaCod = CA.cCtaCod) And Left(CD.nPrdConceptoCod,2) = " & Left(ColocConcepto.gColocConceptoCodGastoCentralRiesgos1, 2) & " And cFlag = ''" _
            & "                 Group By CD.cCtaCod, CD.nNroCalen, CD.nColocCalendApl, CD.nCuota) As PG ON P.cCtaCod = PG.cCtaCod And  PM.nNroCalen = P.nNroCalen And P.nCuota = PG.nCuota" _
            & "   Where  " & sVencidos & " Left(PRD.nPrdEstado,3) = " & Left(gColocEstVigNorm, 3) & " And P.nColocCalendEstado = " & ColocCalendEstado.gColocCalendEstadoPendiente & " And P.nColocCalendApl = " & ColocCalendApl.gColocCalendAplCuota & " And C.cCtaCod In " & psCuentas & " And DateDiff(dd,P.dVenc,'" & Format$(pgdFecSis, gsFormatoFecha) & "') > 0 ) T " _
            & " Group by cCtaCod, nDiasAtraso, cRefinan, cNota1, nCuota, dFecVenc, nMCuota, nMMora ) T1" _
            & " Group by cCtaCod, nDiasAtraso, cRefinan, cNota1, Convert(Varchar(10),dFecVenc,103)"
    ElseIf psTipo = "PRD" Then
        If pbVencidos Then
            sVencidos = "And DateDiff(dd,C.dVenc,'" & Format$(pgdFecSis, "mm/dd/yyyy") & "') > " & pnDiasAtraso
        Else
            sVencidos = ""
        End If
        
        'VER nTasaIntVenc, nCostCusto
        sql = " Select C.cCtaCod cCtaCod, C.dVenc dFecVenc, CPIG.nTasacion nValTasac, (Select nTasaIni from ColocLineaCreditoTasa CL Where CL.cLineaCred = C.cLineaCred And CL.nColocLinCredTasaTpo = " & ColocLineaCredTasas.gColocLineaCredTasasIntMoratNormal & ") nTasaIntVenc, 0 nCostCusto, PRD.nPrdEstado, " _
            & " PRD.nSaldo nSaldoCap, PRD.nTasaInteres nTasaInt, C.nPlazo," _
            & " nTasaCustodia = (Select nParamValor from ColocParametro Where nParamVar = '" & ColocPParametros.gColPParamTasaCustodia & "')," _
            & " nTasaImpuesto = (Select nParamValor from ColocParametro Where nParamVar = '" & ColocPParametros.gColPParamTasaImpuesto & "')," _
            & " nCostoPrepRemate = (Select nParamValor from ColocParametro Where nParamVar = '" & ColocPParametros.gColPParamTasaPreparaRemate & "')" _
            & " From Colocaciones C" _
            & " Inner Join ColocPignoraticio CPIG On C.cCtaCod = CPIG.cCtaCod" _
            & " Inner Join Producto PRD On C.cCtaCod = PRD.cCtaCod" _
            & " Where PRD.nPrdEstado IN (" & ColocEstado.gColPEstDesem & "," & ColocEstado.gColPEstVenci & "," & ColocEstado.gColPEstPRema & "," & ColocEstado.gColPEstRenov & ") " & sVencidos & " " _
            & " And C.cCtaCod IN " & psCuentas & "" _
            & " Order by C.cCtaCod"
    Else
        'VER nTasaIntVenc, nCostCusto
        sql = " Select PRD.cCtaCod From Producto PRD " _
            & " Where PRD.nPrdEstado = '" & ColocEstado.gColocEstJudicial & "' And PRD.cCtaCod IN " & psCuentas

    End If
    
    Set GetDatosCrePendiente = clsCon.CargaRecordSet(sql)
End Function

Public Function TieneMovDespuesApertura(ByVal sCuenta As String) As Boolean
Dim rsMov As Recordset
sSql = "Select ISNULL(Count(*),0) nNumMov From Mov M JOIN MovCap C ON M.nMovNro = C.nMovNro " _
    & "Where M.nMovFlag = " & gMovFlagVigente & " And C.cCtaCod = '" & sCuenta & "' And C.cOpeCod Not In ('" & gITFCobroCargo & "','" & gITFCobroEfectivo & "') " _
    & "And M.nMovEstado = " & gMovEstContabMovContable

Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
If Not (rsMov.EOF And rsMov.BOF) Then
    If rsMov("nNumMov") > 1 Then
        TieneMovDespuesApertura = True
    Else
        TieneMovDespuesApertura = False
    End If
Else
    TieneMovDespuesApertura = False
End If
rsMov.Close
Set rsMov = Nothing
End Function

Private Function ListaCtasCont(ByVal pnMoneda As Moneda, psConstSistCod As Integer) As String
    Dim rsVar As Recordset
    Dim sSql As String
    Dim oCon  As DConecta
    Set oCon = New DConecta
    ListaCtasCont = " "
    
    If oCon.AbreConexion = False Then Exit Function
    sSql = "Select nConsSisDesc, nConsSisValor From ConstSistema where nConsSisCod =" & psConstSistCod & ""
    Set rsVar = New Recordset
    Set rsVar = oCon.CargaRecordSet(sSql)
    
    If Not rsVar.EOF And Not rsVar.BOF Then
        
        ListaCtasCont = Left(Trim(CStr(rsVar("nConsSisValor"))), 2) & Trim(CStr(pnMoneda)) & Mid(Trim(CStr(rsVar("nConsSisValor"))), 4)
    
        rsVar.MoveNext
        
        While Not rsVar.EOF
        
                ListaCtasCont = ListaCtasCont & "','" & Left(Trim(Str(rsVar("nConsSisValor"))), 2) & Trim(Str(pnMoneda)) & Mid(Trim(Str(rsVar("nConsSisValor"))), 4)
        
                rsVar.MoveNext
        Wend
        
        
    End If
    rsVar.Close
    ListaCtasCont = ListaCtasCont
    
    Set rsVar = Nothing
    Set oCon = Nothing
End Function

Public Function GetTranfPendientes(pnMoneda As Moneda) As ADODB.Recordset
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lsCtaCont As String
    
    Dim oCon As NConstSistemas
    Set oCon = New NConstSistemas
    
    'lsCtaCont = oCon.LeeConstSistema(56)
    
    lsCtaCont = ListaCtasCont(pnMoneda, 56)
    
    
     sql = " SELECT     m.nMovNro, m.cMovNro, mif.cPersCod, P.cPersNombre, ISNULL(ABS(mPend.nSaldo),0) ImpSoles, ISNULL(ABS(me.nMovMEImporte),0) ImpDolares, md.cDocNro, m.cMovDesc" _
        & " FROM       Mov m INNER JOIN" _
        & " MovCta mc ON m.nMovNro = mc.nMovNro INNER JOIN" _
        & " MovObjIF mif ON m.nMovNro = mif.nMovNro INNER JOIN" _
        & " Persona p ON mif.cPersCod = p.cPersCod LEFT OUTER JOIN" _
        & " MovDoc md ON m.nMovNro = md.nMovNro left JOIN" _
        & " Documento d ON d.nDocTpo = md.nDocTpo left JOIN" _
        & " MovPendientesRend mPend ON m.nMovNro = mPend.nMovNro LEFT OUTER JOIN" _
        & " MovME me ON mc.nMovNro = me.nMovNro AND mc.nMovItem = me.nMovItem" _
        & " WHERE     (m.nMovEstado = 10) and (m.nMovFlag not in (1,2,3,5)) and (NOT m.cMovNro LIKE '%XXX_')" _
        & " AND (mc.nMovImporte < 0) AND (mc.cCtaContCod like '" & Left(lsCtaCont, 8) & "%' )" _
        & " AND ((mPend.nMovNro IS NULL) OR( mPend.nSaldo > 0))" _
        & " ORDER BY m.cMovNro"
    
    
'    SQL = " SELECT     m.nMovNro, m.cMovNro, mif.cPersCod, P.cPersNombre, ISNULL(ABS(mPend.nSaldo),0) ImpSoles, ISNULL(ABS(me.nMovMEImporte),0) ImpDolares, md.cDocNro, m.cMovDesc" _
'        & " FROM       Mov m INNER JOIN" _
'        & " MovCta mc ON m.nMovNro = mc.nMovNro INNER JOIN" _
'        & " MovObjIF mif ON m.nMovNro = mif.nMovNro INNER JOIN" _
'        & " Persona p ON mif.cPersCod = p.cPersCod LEFT OUTER JOIN" _
'        & " MovDoc md ON m.nMovNro = md.nMovNro left JOIN" _
'        & " Documento d ON d.nDocTpo = md.nDocTpo left JOIN" _
'        & " MovPendientesRend mPend ON m.nMovNro = mPend.nMovNro LEFT OUTER JOIN" _
'        & " MovME me ON mc.nMovNro = me.nMovNro AND mc.nMovItem = me.nMovItem" _
'        & " WHERE     (m.nMovEstado = 10) and (m.nMovFlag not in (1,2,3,5)) and (NOT m.cMovNro LIKE '%XXX_')" _
'        & " AND (mc.nMovImporte < 0) AND (mc.cCtaContCod = '" & Left(lsCtaCont, 2) & Trim(Str(pnMoneda)) & Mid(lsCtaCont, 4) & "')" _
'        & " AND ((mPend.nMovNro IS NULL) OR( mPend.nSaldo > 0))" _
'        & " ORDER BY m.cMovNro"
        
        
    rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set GetTranfPendientes = rs
End Function

Public Function GetCaptaOperacion(psOpeCod As String) As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
     
    sql = "Select IsNull(cOpeDesc,'') cOpeDesc From OpeTpo where copecod = '" & Trim(psOpeCod) & "'"
    rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.EOF And rs.BOF Then
        GetCaptaOperacion = ""
    Else
        GetCaptaOperacion = rs!cOpeDesc
    End If
End Function

Public Function GetApeLoteCuentas(pnMovNro As Long) As ADODB.Recordset
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
     
    sql = " Select MOV.cCtaCod, MOV.nMonto, ISNull(ITF.nMonto,0) ITF From " _
        & " (Select cCtaCod, nMonto From movcap Where nMovNro = " & pnMovNro & " And cOpeCod Not Like '" & Left(gITFCobroCargo, 2) & "%') MOV" _
        & " Left Join (Select cCtaCod, nMonto From movcap Where nMovNro = " & pnMovNro & " And cOpeCod Like '" & Left(gITFCobroCargo, 2) & "%') ITF On MOV.cCtaCod = ITF.cCtaCod"
    
    rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.EOF And rs.BOF Then
        Set GetApeLoteCuentas = rs
    Else
        Set GetApeLoteCuentas = rs
    End If
End Function

Public Sub SetConexion(pCon As ADODB.Connection)
    Set dbCmact = pCon
End Sub





