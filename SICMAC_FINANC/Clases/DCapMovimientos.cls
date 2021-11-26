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
Dim lbTransaccion As Boolean

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
        Optional nPorcRetCTS As Double = 0, Optional sInstitucion As String = "") As String

Dim clsGen As DGeneral
Dim sCuenta As String, sFecha As String
Dim sFecUltCierre As String, sOrdPag As String, sCtaAbonoIntPF As String
Dim nSaldoDisp As Double, nSaldRetCTS As Double
Set clsGen = New DGeneral
sCuenta = clsGen.GeneraNuevaCuenta(sAgencia, nProducto, nMoneda)
Set clsGen = Nothing
sFecha = Format$(dFecha, "mm/dd/yyyy") & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
sSql = "Insert Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) " _
    & "Values ('" & sCuenta & "'," & nTasa & "," & nSaldo & "," & gCapEstActiva & ",'" & sFecha & "',0)"
dbCmact.Execute sSql
nSaldoDisp = IIf(bCheque, 0, nSaldo)
sFecUltCierre = Format$(DateAdd("d", -1, dFecha), "mm/dd/yyyy")
sSql = "Insert Captaciones (cCtaCod,nSaldoDisp,nPersoneria,nFirmas,nIntAcum,dUltCierre,dApertura,nPrdCtaTpo,nPrdTasaInteres,cUltimaActualizacion) " _
    & "Values ('" & sCuenta & "'," & nSaldoDisp & "," & nPersoneria & "," & nFirmas & ",0,'" & sFecUltCierre & "','" & sFecha & "'," & nTipoCuenta & "," _
    & nTipoTasa & ",'" & sMovNro & "')"
dbCmact.Execute sSql
Select Case nProducto
    Case gCapAhorros
        sOrdPag = IIf(bOrdPag, "1", "0")
        sSql = "Insert CaptacAhorros (cCtaCod,nSaldoAnterior,bOrdPag,nSobregiro,dUltContacto) " _
            & "Values ('" & sCuenta & "',0," & sOrdPag & ",0,'" & sFecha & "')"
    Case gCapPlazoFijo
        sCtaAbonoIntPF = IIf(bCtaAboInt, "1", "0")
        sSql = "Insert CaptacPlazoFijo (cCtaCod,nPlazo,nIntPag,dRenovacion,nApertura,dAuxiliar,nFormaRetiro,nDuplicado,bAbonoIntCtaAho) " _
            & "Values ('" & sCuenta & "'," & nPlazo & ",0,'" & sFecha & "'," & nSaldo & ",'" & sFecha & "'," & nFormaRetiro & ",0," & sCtaAbonoIntPF & ")"
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
AgregaNuevaCaptacion = sCuenta
End Function

Public Sub AgregaNuevoProdPers(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona)
sSql = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
    & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ")"
dbCmact.Execute sSql
End Sub

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

Public Sub AgregaOrdenPagoEstado(ByVal sCuenta As String, ByVal sNroDoc As String, _
        ByVal sMovNro As String, ByVal nMonto As Double, ByVal nEstadoOP As CaptacOrdPagoEstado)
sSql = "UPDATE DocRecOP Set nMonto = " & nMonto & " WHERE nTpoDoc = " & TpoDocOrdenPago & " " _
    & "AND cNroDoc = '" & sNroDoc & "' And cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
sSql = "INSERT DocRecOPEst (nTpoDoc,cNroDoc,cCtaCod,cMovNro,nMonto,nEstado) " _
    & "VALUES (" & TpoDocOrdenPago & ",'" & sNroDoc & "','" & sCuenta & "','" & sMovNro & "'," & nMonto & "," & nEstadoOP & ")"
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
        Optional nEstadoOP As CaptacOrdPagoEstado = gCapOPEstCobrada)
Dim sFecha As String
sFecha = Mid(sMovNro, 1, 4) & "/" & Mid(sMovNro, 5, 2) & "/" & Mid(sMovNro, 7, 2) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
Select Case nTpoDoc
    Case TpoDocCheque
        sSql = "INSERT DocRecCapta (nTpoDoc,cNroDoc,cPersCod,cIFTpo,nMovNro,cCtaCod,nMonto) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "'," & nMovNro & ",'" & sCuenta & "'," & nMonto & ")"
        dbCmact.Execute sSql
        sSql = "Update DocRec Set nEstado = " & gChqEstEnValorizacion & " Where cNroDoc = '" & sNroDoc & "' And " _
            & "cIFTPo = '" & Format$(gTpoIFBanco, "00") & "' And cPersCod = '" & sCodIF & "'"
        dbCmact.Execute sSql
        sSql = "INSERT DocRecEst (nTpoDoc,cNroDoc,cPersCod,cIFTpo,cMovNro,nEstado) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "','" & sMovNro & "'," & gChqEstEnValorizacion & ")"
        dbCmact.Execute sSql
    Case TpoDocOrdenPago
        sSql = "INSERT DocRecOP (nTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers,nEstado,cMovNro) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCuenta & "'," & nMonto & ",'" & sCodIF & "'," & nEstadoOP & ",'" & sMovNro & "')"
        dbCmact.Execute sSql
End Select
'sSQL = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
'    & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
'dbCmact.Execute sSQL
End Sub

Public Sub AgregaMov(ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
        ByVal sGlosa As String, Optional nMovEstado As MovEstado = gMovEstContabMovContable, _
        Optional nMovFlag As MovFlag = gMovFlagVigente)
sSql = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sGlosa & "'," & nMovEstado & "," & nMovFlag & ")"
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

Public Sub AgregaMovLavDinero(ByVal nMovNro As Long, ByVal sPersCod As String)
sSql = "INSERT MovLavDinero (nMovNro,cPersCod) " _
    & "VALUES (" & nMovNro & ",'" & sPersCod & "')"
dbCmact.Execute sSql
End Sub


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

Public Sub ActualizaMovPendientesRend(ByVal nMovNro As Long, ByVal pnMonto As Currency)
    sSql = " UPDATE MovPendientesRend " _
         & " Set nSaldo = nSaldo - " & pnMonto & "" _
         & " Where nMovNro = " & nMovNro & ""
    dbCmact.Execute sSql
End Sub

Public Sub AgregaMovCMAC(ByVal nMovNro As Long, ByVal sPersCod As String, _
        ByVal nMoneda As Moneda, ByVal nOperacion As CaptacOperacion, _
        ByVal nMonto As Double, Optional sDocumento As String = "NULL", _
        Optional sCuenta As String = "NULL")
Dim sIFTipo As String
Dim sCta As String, sDoc As String
sCta = "NULL"
sDoc = "NULL"
sIFTipo = Format$(gTpoIFCmac, "00")
If sCuenta <> "NULL" Then sCta = "'" & sCuenta & "'"
If sDocumento <> "NULL" Then sDoc = "'" & sDocumento & "'"
sSql = "INSERT MovCMAC (nMovNro,cPersCod,cIFTpo,nMoneda,cCtaIF,cDocumento,cOpeCod,nMonto) " _
    & "VALUES (" & nMovNro & ",'" & sPersCod & "','" & sIFTipo & "'," & nMoneda & "," & sCta & "," & sDoc & ",'" & Trim(nOperacion) & "'," & Trim(nMonto) & ")"
dbCmact.Execute sSql
End Sub

Public Sub AgregaMovParamCTS(ByVal nMovNro As Long, ByVal nPorcentaje As Double)
sSql = "INSERT MovParamCTS (nMovNro,nPorcentajeAbono) " _
    & "VALUES (" & nMovNro & "," & nPorcentaje & ")"
dbCmact.Execute sSql
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
sSql = "Select K.cConsDescripcion Motivo FROM ProductoBloqueos B INNER JOIN " _
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
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaCargoCTS(ByVal sCuenta As String, nMonto As Double, nIntSaldo As Double)
sSql = "Update CaptacCTS Set nSaldRetiro = nSaldRetiro - " & nMonto & ", " _
    & "nIntSaldo = nIntSaldo + " & nIntSaldo & " WHERE cCtaCod = '" & sCuenta & "'"
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

Public Sub ActualizaEstadoDocRecEst(ByVal nTipoDoc As TpoDoc, ByVal sNroDoc As String, _
        ByVal sCodIF As String, ByVal sMovNro As String, ByVal nEstado As ChequeEstado)
'Actualiza el ultimo estado a la tabla de Documento Recibidos
sSql = "Update DocRec Set nEstado = " & nEstado & " WHERE " _
    & "nTpoDoc = '" & nTipoDoc & "' AND cNroDoc = '" & sNroDoc & "' AND cPersCod = '" & sCodIF & "'"
dbCmact.Execute sSql
'Agrega un registro mas de estado a la historia de estados del documento recibido
sSql = "INSERT DocRecEst (nTpoDoc,cNroDoc,cPersCod,cMovNro,nEstado) " _
    & "VALUES ('" & nTipoDoc & "','" & sNroDoc & "','" & sCodIF & "','" & sMovNro & "'," & nEstado & ")"
dbCmact.Execute sSql
End Sub

Public Function GetMovExtorno(ByVal sDatoBus As String, ByVal dFecha As Date, _
        ByVal nOperacion As CaptacOperacion, Optional nTipoBus As Integer = 0) As Recordset
Dim rsMov As Recordset
Dim sWhere As String
If nTipoBus = 0 Then
    sWhere = "M.cMovNro LIKE '%" & sDatoBus & "%'"
ElseIf nTipoBus = 1 Then
    sWhere = "M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And C.cCtaCod = '" & sDatoBus & "'"
End If
sSql = "Select M.cMovNro, O.cOpeDesc, C.cCtaCod, ISNULL(MD.nDocTpo,'') nDocTpo, ISNULL(MD.cDocNro,'') cDocNro, " _
    & "M.cMovDesc, C.cOpeCod, C.nMonto, M.nMovNro, ISNULL(MD.cPersCod,'') cPersCod FROM (Select MD.nMovNro, MD.nDocTpo, MD.cDocNro, ISNULL(D.cPersCod,'') cPersCod " _
    & "From DocRec D RIGHT JOIN MovDoc MD ON D.nTpoDoc = MD.nDocTpo And D.cNroDoc = MD.cDocNro) MD RIGHT JOIN Mov M " _
    & "INNER JOIN MovCap C INNER JOIN " & sDBComunes & "OpeTpo O ON C.cOpeCod = O.cOpeCod ON M.nMovNro = C.nMovNro " _
    & "ON MD.nMovNro = M.nMovNro WHERE " & sWhere & " AND C.cOpeCod  IN ('" & nOperacion & "') " _
    & "AND M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & "," & gMovFlagDeExtorno & ")" _
    & "AND M.nMovEstado IN (" & gMovEstContabMovContable & ") " _
    & "ORDER BY M.nMovNro"
Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetMovExtorno = rsMov
End Function

Public Function GetMovExtornoCMAC(ByVal sDatoBus As String, ByVal dFecha As Date, _
        ByVal nOperacion As CaptacOperacion, Optional nTipoBus As Integer = 0) As Recordset
Dim rsMov As Recordset
Dim sWhere As String
If nTipoBus = 0 Then
    sWhere = "M.cMovNro LIKE '%" & sDatoBus & "%'"
ElseIf nTipoBus = 1 Then
    sWhere = "M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And C.cCtaCod = '" & sDatoBus & "'"
End If
sSql = "Select M.cMovNro, O.cOpeDesc, ISNULL(MC.cCtaIF,'') cCtaCod, ISNULL(MC.cDocumento,'') cDocNro, " _
    & "M.cMovDesc, MC.cOpeCod, MC.nMonto, M.nMovNro, ISNULL(OD.nDocTpo,'') nDocTpo " _
    & "FROM Mov M INNER JOIN MovCMAC MC INNER JOIN " & sDBComunes & "OpeTpo O LEFT JOIN " _
    & sDBComunes & "OpeDoc OD ON O.cOpeCod = OD.cOpeCod ON MC.cOpeCod = O.cOpeCod ON " _
    & "M.nMovNro = MC.nMovNro WHERE " & sWhere & " AND M.nMovFlag NOT IN " _
    & "(" & gMovFlagExtornado & "," & gMovFlagEliminado & "," & gMovFlagDeExtorno & ")" _
    & "AND M.nMovEstado IN (" & gMovEstContabMovContable & ") And M.cOpeCod = '" & nOperacion & "'" _
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

Public Sub ActualizaDuplicadoCertPF(ByVal sCuenta As String)
sSql = "Update CaptacPlazoFijo Set nDuplicado = nDuplicado + 1 Where cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub
Public Function GetMovTransferencia(ByVal sMovNro As String) As Recordset
Dim rsTrans As Recordset
sSql = "SELECT M.nMovNro, C.cOpeCod, C.cCtaCod, C.nMonto From Mov M INNER JOIN " _
    & "MovCap C ON M.nMovNro = C.nMovNro Where M.cMovNro = '" & sMovNro & "' " _
    & "ORDER BY M.cOpeCod DESC"
Set rsTrans = New Recordset
rsTrans.CursorLocation = adUseClient
rsTrans.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTrans.ActiveConnection = Nothing
Set GetMovTransferencia = rsTrans
Set rsTrans = Nothing
End Function

Public Function GetFechaUltimoRetiroIntPF(ByVal sCuenta As String) As Date
Dim rsMov As Recordset
sSql = "Select TOP 1 PF.dRenovacion, M.cMovNro FROM CaptacPlazoFijo PF LEFT JOIN " _
    & "MovCap C INNER JOIN Mov M ON C.nMovNro = M.nMovNro ON PF.cCtaCod = C.cCtaCod " _
    & "WHERE PF.cCtaCod = '" & sCuenta & "' OR (C.cOpeCod IN ('" & gPFRetInt & "') " _
    & "AND M.nMovFlag NOT IN ('" & gMovFlagExtornado & "') AND PF.cCtaCod = '" & sCuenta & "') " _
    & "ORDER BY M.nMovNro"
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

Public Function AgregaGiro(ByVal sCuenta As String, ByVal dApertura As Date, ByVal nMonto As Double, _
    ByVal nTipoGiro As ProductoCuentaTipo, ByVal sAgenciaDest As String)
sSql = "Insert Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) " _
    & "Values ('" & sCuenta & "',0," & nMonto & "," & gCapEstActiva & ",'" & Format$(dApertura & " " & Time, "mm/dd/yyyy hh:mm:ss") & "',0)"
dbCmact.Execute sSql
sSql = "Insert Giro (cCtaCod,dVigencia,nPrdCtaTpo,cAgenciaDest) " _
    & "Values ('" & sCuenta & "','" & Format$(dApertura, "mm/dd/yyyy") & "'," & nTipoGiro & ",'" & sAgenciaDest & "')"
dbCmact.Execute sSql
End Function

Public Sub AgregaProdPersGiro(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As GiroRelacPersona)
sSql = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
    & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ")"
dbCmact.Execute sSql
End Sub

Public Sub AgregaMovTranferBanco(ByVal pnMovNro As Long, ByVal psPersCodIF As String, psCtaBco As String, ByVal pnMoneda As Moneda)
sSql = " Insert MovTransferBco (nMovNro, cPersCodIF, cCtaBco, nMoneda)" _
     & " Values (" & pnMovNro & ",'" & psPersCodIF & "','" & psCtaBco & "'," & pnMoneda & ")"
dbCmact.Execute sSql
End Sub

Public Sub AgregaGiroDestinatario(ByVal sCuenta As String, ByVal sDestinatario As String, _
        ByVal sReferencia As String)
sSql = "Insert GiroDestinatario (cCtaCod,cDestinatario,cReferencia) " _
    & "Values ('" & sCuenta & "','" & sDestinatario & "','" & sReferencia & "')"
dbCmact.Execute sSql
End Sub

Public Sub AgregaMovGiro(ByVal nMovNro As Long, ByVal sCuenta As String, ByVal nMonto As Double, _
        Optional ByVal nComision As Double = 0)
sSql = "Insert MovGiro(nMovNro,cCtaCod,nMonto,nComision) " _
    & "Values (" & nMovNro & ",'" & sCuenta & "'," & nMonto & "," & nComision & ")"
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
'Dim sConn As String
'Dim ClsIni As DClassIni
'lbTransaccion = True
'Set ClsIni = New DClassIni
'sConn = ClsIni.CadenaConexion(App.Path & "\SicmactCFinan.INI")
'sDBComunes = ClsIni.BaseComunes
'sDBPersona = ClsIni.BasePersonas
'sDBImagenes = ClsIni.BaseImagenes
'Set ClsIni = Nothing
'Set dbCmact = New Connection
'dbCmact.CursorLocation = adUseClient
'dbCmact.Open sConn
'dbCmact.Execute "SET DATEFORMAT MDY"

Dim dbCmact As DConecta
Set dbCmact = New DConecta
lbTransaccion = True
dbCmact.AbreConexion
dbCmact.Ejecutar "SET DATEFORMAT MDY"

End Sub

Private Sub Class_Terminate()
On Error GoTo ERROR
If lbTransaccion Then
    dbCmact.Close
    Set dbCmact = Nothing
End If
Exit Sub
ERROR:
End Sub

Public Function CtaConFirmas(psCtaCod As String) As Boolean
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim clsCon As DConecta
    Set clsCon = New DConecta

    sql = " Select PP.cCtaCod From ProductoPersona PP" _
        & " Left Join PersImagen PI On PP.cPersCod = PI.cPersCod" _
        & " Where PP.cCtaCod = '" & psCtaCod & "' And PI.iPersFirma IS Null"
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

Public Sub AgregaMovOpeVarias(pnMovNro As Long, pnMovImporte As Currency, psNroDoc As String, psReferencia As String, pnMoneda As Moneda)
    Dim sql As String
    
    sql = " Insert MovOpeVarias (nMovNro,nMovImporte,cNroDoc,cReferencia,nMoneda)" _
        & " Values(" & pnMovNro & "," & pnMovImporte & ",'" & psNroDoc & "','" & psReferencia & "'," & pnMoneda & ")"
    
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
        & " Where TA.cOpeCod = '" & psOpeCod & "'"
    rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.EOF And rs.BOF Then
        GetTarifa = 0
        pnMoneda = gMonedaNacional
    Else
        GetTarifa = rs.Fields(0)
        pnMoneda = rs.Fields(1)
    End If
    
    Set rs = Nothing
End Function

Public Function GetOtrasOperaciones(pdFecha As Date) As ADODB.Recordset
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim clsCon As DConecta
    Set clsCon = New DConecta
    
    clsCon.AbreConexion
    
    sql = " SELECT M.nMovNro, Convert(DateTime,Left(M.cMovNro,8)) dFecTran, M.cOpeCod, Right(M.cMovNro,4) cCodUsu, OP.cOpeDesc, MOV.nMovImporte, MOV.nMoneda " _
        & " FROM Mov  M" _
        & " Inner Join MovOpeVarias MOV On M.nMovNro = MOV.nMovNro" _
        & " INNER JOIN OpeTpo OP ON M.cOpeCod = OP.cOpeCod" _
        & " WHERE M.nMovFlag = 0 And Left(M.cOpeCod,4) In ('" & Left(gOtrOpeAhoOtrosIngresos, 4) & "','" & Left(gOtrOpeAhoOtrosEgresos, 4) & "','" & Left(gOtrOpeDuplicadoTarjeta, 4) & "') And M.cMovNro Like '" & Format(pdFecha, gsFormatoMovFecha) & "%'"
    Set rs = clsCon.CargaRecordSet(sql)
    
    Set GetOtrasOperaciones = rs
    
End Function

Public Function GetOtrasOperacionesDet(psMovNro As Long) As ADODB.Recordset
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim clsCon As DConecta
    Set clsCon = New DConecta
    
    clsCon.AbreConexion
    
    sql = " SELECT M.nMovNro, Convert(DateTime,Left(M.cMovNro,8)) dFecTran, M.cOpeCod, Right(M.cMovNro,4) cCodUsu, OP.cOpeDesc, MOV.nMovImporte, PE.cPersCod , PE.cPersNombre, M.cMovDesc, MOV.cNroDoc " _
        & " FROM Mov M" _
        & " Inner Join MovOpeVarias MOV On M.nMovNro = MOV.nMovNro" _
        & " Inner Join MovGasto MG On M.nMovNro = MG.nMovNro" _
        & " Inner Join Persona PE On MG.cPersCod = PE.cPersCod" _
        & " INNER JOIN OpeTpo OP ON M.cOpeCod = OP.cOpeCod" _
        & " WHERE M.nMovNro = " & psMovNro
    Set rs = clsCon.CargaRecordSet(sql)
    
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
            sql = " Select PP.cCtaCod From ProductoPersona PP" _
                & " Inner Join Colocaciones COL On PP.cCtaCod = COL.cCtaCod" _
                & " Inner Join Producto PRD On PP.cCtaCod = PRD.cCtaCod" _
                & " Where Substring(PP.cCtaCod,6,3) <> '" & Producto.gColConsuPrendario & "' And LEft(PRD.nPrdEstado,2) <> '" & Left(ColocEstado.gColocEstRecVigJud, 2) & "'" _
                & "  And PP.cPersCod = '" & sPersona & "' Order by PP.cCtaCod "
        Case "JUD"
            sql = " Select PP.cCtaCod From ProductoPersona PP" _
                & " Inner Join Colocaciones COL On PP.cCtaCod = COL.cCtaCod" _
                & " Inner Join Producto PRD On PP.cCtaCod = PRD.cCtaCod" _
                & " Where Left(PRD.nPrdEstado,2) = '" & Left(ColocEstado.gColocEstRecVigJud, 2) & "'" _
                & " And PP.cPersCod = '" & sPersona & "' Order by PP.cCtaCod "
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

Public Function GetFirma(psPersCod As String) As ADODB.Recordset
    Dim sql As String
    Dim clsCon As DConecta
    Set clsCon = New DConecta
    
    clsCon.AbreConexion
    
    sql = " Select cPersCod cCodPers, iPersFirma iFirma From PersImagen Where cPersCod = '" & psPersCod & "' And iPersFirma Is Not Null"
    Set GetFirma = clsCon.CargaRecordSet(sql)
    
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
sSql = "Select CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' " _
    & "+ SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha, " _
    & "O.cOpeDesc Operacion, ISNULL(MD.cDocNro,'') cDocumento , ABS(CD.nMonto) nMonto, " _
    & "C.nSaldoContable, A.cAgeDescripcion cAgencia " _
    & "FROM MovDoc MD RIGHT JOIN Mov M INNER JOIN MovCap C INNER JOIN MovCapDet CD INNER JOIN " _
    & "OpeTpo O INNER JOIN CapMovTipo CT ON O.cOpeCod = CT.cOpeCod ON " _
    & "CD.cOpeCod = O.cOpeCod ON C.nMovNro = CD.nMovNro And C.cCtaCod = CD.cCtaCod ON M.nMovNro = C.nMovNro " _
    & "ON MD.nMovNro = M.nMovNro INNER JOIN Agencias A ON SUBSTRING(M.cMovNro,18,2) = A.cAgeCod " _
    & "And M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "WHERE C.cCtaCod = '" & sCuenta & "'"

Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
If Not (rsMov.EOF And rsMov.BOF) Then
    rsMov.MoveLast
    If rsMov.RecordCount > 1 Then
        TieneMovDespuesApertura = True
    Else
        TieneMovDespuesApertura = False
    End If
Else
    TieneMovDespuesApertura = False
End If
End Function

Public Function GetTranfPendientes(pnMoneda As Moneda) As ADODB.Recordset
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lsCtaCont As String
    
    Dim oCon As NConstSistemas
    Set oCon = New NConstSistemas
    
    lsCtaCont = oCon.LeeConstSistema(56)
    
    sql = " SELECT     m.nMovNro, m.cMovNro, mif.cPersCod, P.cPersNombre, mPend.nSaldo ImpSoles, ISNULL(me.nMovMEImporte,0) ImpDolares, md.cDocNro, m.cMovDesc" _
        & " FROM       Mov m INNER JOIN" _
        & " MovCta mc ON m.nMovNro = mc.nMovNro INNER JOIN" _
        & " MovObjIF mif ON m.nMovNro = mif.nMovNro INNER JOIN" _
        & " Persona p ON mif.cPersCod = p.cPersCod LEFT OUTER JOIN" _
        & " MovDoc md ON m.nMovNro = md.nMovNro INNER JOIN" _
        & " Documento d ON d.nDocTpo = md.nDocTpo  JOIN" _
        & " MovPendientesRend mPend ON m.nMovNro = mPend.nMovNro LEFT OUTER JOIN" _
        & " MovME me ON mc.nMovNro = me.nMovNro AND mc.nMovItem = me.nMovItem" _
        & " WHERE     (m.nMovEstado = 10) and (m.nMovFlag not in (1,2,3,5)) and (NOT m.cMovNro LIKE '%XXX_')" _
        & " AND (mc.nMovImporte < 0) AND (mc.cCtaContCod = '" & Left(lsCtaCont, 2) & Trim(Str(pnMoneda)) & Mid(lsCtaCont, 4) & "')" _
        & " AND ((mPend.nMovNro IS NULL) OR( mPend.nSaldo > 0))" _
        & " ORDER BY m.cMovNro"
    rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set GetTranfPendientes = rs
End Function
 
Public Sub SetConexion(pCon As ADODB.Connection)
lbTransaccion = False
    Set dbCmact = pCon
End Sub
'ALPA 20090929*******************************************************
Public Function ReporteResumenAhorro(ByVal pdFecha As Date, ByVal pnTipoCambio As Currency) As ADODB.Recordset
    Dim sql As String
    Dim clsCon As DConecta
    Set clsCon = New DConecta
    
    clsCon.AbreConexion
    
    sql = " exec stp_sel_ReporteAhorro '" & Format(pdFecha, "YYYY/MM/DD") & "'," & pnTipoCambio
    Set ReporteResumenAhorro = clsCon.CargaRecordSet(sql)
    
    Set clsCon = Nothing
End Function
'********************************************************************

'ALPA 20110419*******************************************************
Public Function ReporteResumenAhorroCtasCbles(ByVal pdFecha As Date) As ADODB.Recordset
    Dim sql As String
    Dim clsCon As DConecta
    Set clsCon = New DConecta
    
    clsCon.AbreConexion
    sql = " exec stp_CtasBalancexAhorroLCPyMNE '" & Format(pdFecha, "YYYY/MM/DD") & "'"
    Set ReporteResumenAhorroCtasCbles = clsCon.CargaRecordSet(sql)
    Set clsCon = Nothing
    'clsCon.CierraConexion
End Function
'********************************************************************

'ALPA 20110419*******************************************************
Public Function ReporteListaCuentasRestringidas(ByVal pdFecha As Date) As ADODB.Recordset
    Dim sql As String
    Dim clsCon As DConecta
    Set clsCon = New DConecta
    
    clsCon.AbreConexion
    sql = " exec stp_sel_ListaCuentasRestringidas '" & Format(pdFecha, "YYYY/MM/DD") & "'"
    Set ReporteListaCuentasRestringidas = clsCon.CargaRecordSet(sql)
    Set clsCon = Nothing
    'clsCon.CierraConexion
End Function
'********************************************************************




