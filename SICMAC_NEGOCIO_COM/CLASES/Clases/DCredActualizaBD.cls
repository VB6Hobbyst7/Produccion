VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCredActualizaBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'* Modulo de Colocaciones Pignoraticio
'* Clase con todas las Actualizaciones a BD
Option Explicit
Dim csConexion As String
Dim csNegocio As String
Dim csCentralPer As String
Dim csCentralCom As String
Dim csCentralImg As String
Dim csAdminist As String

Public coConex As DConecta

'******************************************************************************************
'************************************** Copiado de Clases de Ahorros **********************
'******************************************************************************************

Public Sub dUpdateColocGarantia(ByVal psNumgarant As String, psCodCta As String, pnEstado As Integer)
Dim sSQL As String

    sSQL = "UPDATE ColocGarantia Set nEstado = " & pnEstado & " WHERE cNumGarant = '" & psNumgarant & "' AND cCtaCod = '" & psCodCta & "'"
    coConex.Ejecutar sSQL

End Sub


Public Function RecuperaMovimientoCapataciones(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSQL As String

    sSQL = "Select M.cMovNro, MC.cOpeCod, MC.cCtaCod, MC.nMonto from MovCap MC "
    sSQL = sSQL & " Inner Join Mov M ON M.nMovNro = MC.nMovNro "
    sSQL = sSQL & " Where MC.nMovNro = " & pnMovNro
    
    Set RecuperaMovimientoCapataciones = New ADODB.Recordset
    
    RecuperaMovimientoCapataciones.CursorLocation = adUseClient
    RecuperaMovimientoCapataciones.Open sSQL, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    RecuperaMovimientoCapataciones.ActiveConnection = Nothing

End Function


Public Function GetDatosCuentaAho(ByVal sCuenta As String) As Recordset
Dim sSQL As String
Dim rsCta As Recordset

sSQL = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, " _
    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.bOrdPag, A.dUltContacto, T.cConsDescripcion cEstado, " _
    & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.bInactiva, A.bInmovilizada, A.nSobregiro FROM Producto P " _
    & "INNER JOIN Captaciones C INNER JOIN CaptacAhorros A ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor WHERE P.cCtaCod = '" & sCuenta & "' AND " _
    & "T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    & "AND T2.nConsCod = " & gCaptacTipoTasa

Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSQL, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaAho = rsCta
Set rsCta = Nothing
End Function



Public Function GetDatosMovExtorno(ByVal sDatoBus As String, ByVal dFecha As Date, _
        ByVal noperacion As CaptacOperacion, Optional nTipoBus As Integer = 0) As Recordset

Set GetDatosMovExtorno = GetMovExtorno(sDatoBus, dFecha, nTipoBus)

End Function


Public Function CapExtornoApertura(ByVal sMovNroCap As String, ByVal pnMovNro As Long, ByVal nMovNroBus As Long, ByVal noperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional bDocumento As Boolean = False, Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String = "", Optional sLpt As String = "")

Dim sMsgProd As String, sCodUser As String
Dim nProd As Producto
Dim nMovNro As Long
Dim dFecSis As Date
Dim oFun As NContFunciones

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
nProd = CLng(Mid(sCuenta, 6, 3))
If nProd = gCapAhorros Then
    sMsgProd = "EXTORNO AHORROS"
ElseIf nProd = gCapPlazoFijo Then
    sMsgProd = "EXTORNO PLAZO FIJO"
ElseIf nProd = gCapCTS Then
    sMsgProd = "EXTORNO CTS"
End If

nMovNro = pnMovNro


ActualizaEstadoMov nMovNroBus, gMovFlagExtornado

'Para evtar que se bloquee la cuenta con alguna operacion de retiro
'hecha anteriormente se crea como una nueva operacion y luego referenciamos con el movref
AgregaMov sMovNroCap, noperacion, sGlosa, , gMovFlagDeExtorno
nMovNro = GetnMovNro(sMovNroCap)

ActualizaEstadoCuenta sCuenta, gCapEstAnulada, dFecSis, sMovNro
AgregaMovCap nMovNro, noperacion, sCuenta, nMonto, 0, 0
AgregaMovCapDet nMovNro, noperacion, sCuenta, gConcCapital, nMonto
dInsertMovRef nMovNro, pnMovNro, False
UltimaActualizacionCuenta sCuenta, sMovNro
If bDocumento Then
    If nTipoDoc = TpoDocCheque Then
        ActualizaEstadoDocRecEst TpoDocCheque, sNroDoc, sCodIF, sMovNro, gsChqEstExtornado
    End If
End If


End Function


Public Function CapExtornoAbonoAho(ByVal sMovNroCap As String, ByVal pnMovNro As Long, ByVal nMovNroBus As Long, ByVal noperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String, Optional sLpt As String = "") As Double

Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nDiasTranscurridos As Long, i As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
Dim oFun As NContFunciones

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo

Set rsCta = GetDatosCuentaAho(sCuenta)

nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, False, True)

nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing
'Inicia la transaccion


bTrans = True
Randomize
For i = 0 To Rnd(2000) * 1000
Next i
'If Not ValidaSaldoCuenta(sCuenta, nMonto) Then
'
'    Err.Raise 1000, "CapExtornoAbonoAho", "Cuenta NO Posee Saldo Suficiente"
'    CapExtornoAbonoAho = 0
'    Exit Function
'End If
sGlosa = sGlosa & "Extorno Abono Cuenta = " & sCuenta
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt - nMonto

bTrans = True

ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp

If sNroDoc <> "" Then
    If nTipoDoc = TpoDocCheque Then
        nSaldoDisp = nSaldoDisp - 0
        ActualizaCargoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
        ActualizaEstadoDocRecEst TpoDocCheque, sNroDoc, sCodIF, sMovNro, gsChqEstExtornado
        sMsgOpe = "Ext. Depósito Chq."
    ElseIf nTipoDoc = TpoDocNotaCargo Then
        nSaldoDisp = nSaldoDisp - nMonto
        sMsgOpe = "Ext. Depósito NC"
    End If
Else
    nSaldoDisp = nSaldoDisp - nMonto
    ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True

End If
ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
nMovNro = pnMovNro
'Para evtar que se bloquee la cuenta con alguna operacion de retiro
'hecha anteriormente se crea como una nueva operacion y luego referenciamos con el movref
AgregaMov sMovNroCap, noperacion, sGlosa, gMovEstContabMovContable, gMovFlagDeExtorno
nMovNro = GetnMovNro(sMovNroCap)
AgregaMovCap nMovNro, noperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
AgregaMovCapDet nMovNro, noperacion, sCuenta, gConcCapital, nMonto
dInsertMovRef nMovNro, pnMovNro, False
UltimaActualizacionCuenta sCuenta, sMovNro

CapExtornoAbonoAho = nSaldoCnt

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(noperacion)
sModDep = sMsgOpe
sTipApe = "EXTORNO AHORROS"

End Function

Public Function CapExtornoCargoAho(ByVal pnMovNro As Long, ByVal nMovNroBus As Long, ByVal noperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String = "", Optional sLpt As String = "") As Double

Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, nExtracto As Long
Dim nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo
Set rsCta = GetDatosCuentaAho(sCuenta)

nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 1, True, False)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion

bTrans = True
If sGlosa = "" Then sGlosa = "Extorno Cargo Cuenta = " & sCuenta
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt + nMonto
nSaldoDisp = nSaldoDisp + nMonto
ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
If sNroDoc = "" Then
    sMsgOpe = "Ext. Retiro Efectivo"
Else
    If nTipoDoc = TpoDocOrdenPago Then
        If noperacion = gAhoExtRetOPCanje Then
            sMsgOpe = "Ext. Retiro OP Canje"
        Else
            sMsgOpe = "Ext. Retiro OP"
        End If
        AgregaOrdenPagoEstado sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstExtornada
    ElseIf nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = " Ext. Retiro NC"
    End If
End If
ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
AgregaMov sMovNro, noperacion, sGlosa, , gMovFlagDeExtorno
nMovNro = GetnMovNro(sMovNro)
'nMovNro = pnMovNro
AgregaMovCap nMovNro, noperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
AgregaMovCapDet nMovNro, noperacion, sCuenta, gConcCapital, nMonto
UltimaActualizacionCuenta sCuenta, sMovNro

CapExtornoCargoAho = nSaldoCnt

End Function


Public Function GetnMovNro(ByVal sMovNro As String) As Long
Dim rsMov As Recordset
Dim sSQL As String

sSQL = "Select nMovNro From Mov Where cMovNro = '" & sMovNro & "'"
Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSQL, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
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
Dim sSQL As String
Set clsGen = New DGeneral
sCuenta = clsGen.GeneraNuevaCuenta(sAgencia, nProducto, nMoneda)
Set clsGen = Nothing
sFecha = Format$(dFecha, "mm/dd/yyyy") & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
sSQL = "Insert Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) " _
    & "Values ('" & sCuenta & "'," & nTasa & "," & nSaldo & "," & gCapEstActiva & ",'" & sFecha & "',0)"
coConex.ConexionActiva.Execute sSQL
nSaldoDisp = IIf(bCheque, 0, nSaldo)
sFecUltCierre = Format$(DateAdd("d", -1, dFecha), "mm/dd/yyyy")
sSQL = "Insert Captaciones (cCtaCod,nSaldoDisp,nPersoneria,nFirmas,nIntAcum,dUltCierre,dApertura,nPrdCtaTpo,nPrdTasaInteres,cUltimaActualizacion) " _
    & "Values ('" & sCuenta & "'," & nSaldoDisp & "," & nPersoneria & "," & nFirmas & ",0,'" & sFecUltCierre & "','" & sFecha & "'," & nTipoCuenta & "," _
    & nTipoTasa & ",'" & sMovNro & "')"
coConex.ConexionActiva.Execute sSQL
Select Case nProducto
    Case gCapAhorros
        sOrdPag = IIf(bOrdPag, "1", "0")
        sSQL = "Insert CaptacAhorros (cCtaCod,nSaldoAnterior,bOrdPag,nSobregiro,dUltContacto) " _
            & "Values ('" & sCuenta & "',0," & sOrdPag & ",0,'" & sFecha & "')"
    Case gCapPlazoFijo
        sCtaAbonoIntPF = IIf(bCtaAboInt, "1", "0")
        sSQL = "Insert CaptacPlazoFijo (cCtaCod,nPlazo,nIntPag,dRenovacion,nApertura,dAuxiliar,nFormaRetiro,nDuplicado,bAbonoIntCtaAho) " _
            & "Values ('" & sCuenta & "'," & nPlazo & ",0,'" & sFecha & "'," & nSaldo & ",'" & sFecha & "'," & nFormaRetiro & ",0," & sCtaAbonoIntPF & ")"
    Case gCapCTS
        nSaldRetCTS = Round(nSaldo * nPorcRetCTS / 100, 2)
        sSQL = "Insert CaptacCTS (cCtaCod,nSaldRetiro,nIntSaldo,cCodInst) " _
            & "Values ('" & sCuenta & "'," & nSaldRetCTS & ",0,'" & sInstitucion & "')"
End Select
coConex.ConexionActiva.Execute sSQL
If nProducto = gCapPlazoFijo And bCtaAboInt Then
    sSQL = "Insert CaptacCtaAboIntPF (cCtaCod,cCtaCodAbono) " _
        & "Values ('" & sCuenta & "','" & sCtaCodAbono & "')"
    coConex.ConexionActiva.Execute sSQL
End If
AgregaNuevaCaptacion = sCuenta
End Function

Public Sub AgregaNuevoProdPers(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona)
Dim sSQL As String
sSQL = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
    & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ")"
coConex.ConexionActiva.Execute sSQL
End Sub

Public Function GetDatosOrdenPago(ByVal sCuenta As String, ByVal sNroDoc As String) As Recordset
Dim rsOP As Recordset
Dim sSQL As String
sSQL = "SELECT E.cMovNro, D.nTpoDoc, D.cNroDoc, D.cCtaCod, D.nMonto, P.cPersNombre, E.cEstado " _
    & "D.cIFCodPers FROM DocRecOPEst E INNER JOIN DocRecOP D INNER JOIN " & csCentralPer & "Persona P " _
    & "ON D.cIFCodPers = P.cPersCod ON E.nTpoDoc = D.nTpoDoc AND E.cNroDoc = D.cNroDoc AND E.cCtaCod " _
    & "= D.cCtaCod WHERE E.cMovNro IN (SELECT MAX(E1.cMovNro) FROM DocRecOPEst E WHERE E1.nTpoDoc = " _
    & "E.nTpoDoc AND E1.cNroDoc = E.cNroDoc AND E1.cCtaCod = E.cCtaCod)"
Set rsOP = New Recordset
rsOP.CursorLocation = adUseClient
rsOP.Open sSQL, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsOP.ActiveConnection = Nothing
Set GetDatosOrdenPago = rsOP
Set rsOP = Nothing
End Function

Public Sub AgregaOrdenPagoEstado(ByVal sCuenta As String, ByVal sNroDoc As String, _
        ByVal sMovNro As String, ByVal nMonto As Double, ByVal nEstadoOP As CaptacOrdPagoEstado)
Dim sSQL As String
sSQL = "UPDATE DocRecOP Set nMonto = " & nMonto & " WHERE nTpoDoc = " & TpoDocOrdenPago & " " _
    & "AND cNroDoc = '" & sNroDoc & "' And cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSQL
sSQL = "INSERT DocRecOPEst (nTpoDoc,cNroDoc,cCtaCod,cMovNro,nMonto,nEstado) " _
    & "VALUES (" & TpoDocOrdenPago & ",'" & sNroDoc & "','" & sCuenta & "','" & sMovNro & "'," & nMonto & "," & nEstadoOP & ")"
coConex.ConexionActiva.Execute sSQL
End Sub

Public Sub AgregaCuentaDocumento(ByVal sCuenta As String, ByVal nTpoDoc As TpoDoc, _
        ByVal sNroDoc As String, ByVal sCodIF As String, ByVal sMovNro As String, _
        ByVal nMovNro As Long, Optional nMonto As Double = 0, _
        Optional nEstadoOP As CaptacOrdPagoEstado = gCapOPEstCobrada)
Dim sFecha As String
Dim sSQL As String
sFecha = Mid(sMovNro, 1, 4) & "/" & Mid(sMovNro, 5, 2) & "/" & Mid(sMovNro, 7, 2) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
Select Case nTpoDoc
    Case TpoDocCheque
        sSQL = "INSERT DocRecCapta (nTpoDoc,cNroDoc,cPersCod,cIFTpo,nMovNro,cCtaCod,nMonto) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "'," & nMovNro & ",'" & sCuenta & "'," & nMonto & ")"
    Case TpoDocOrdenPago
        sSQL = "INSERT DocRecOP (nTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers,nEstado,cMovNro) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCuenta & "'," & nMonto & ",'" & sCodIF & "'," & nEstadoOP & ",'" & sMovNro & "')"
    Case TpoDocNotaAbono, TpoDocNotaCargo
End Select
coConex.ConexionActiva.Execute sSQL
sSQL = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
    & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
coConex.ConexionActiva.Execute sSQL
End Sub

Public Sub AgregaMov(ByVal sMovNro As String, ByVal noperacion As CaptacOperacion, _
        ByVal sGlosa As String, Optional nMovEstado As MovEstado = gMovEstContabMovContable, _
        Optional nMovFlag As MovFlag = gMovFlagVigente)
Dim sSQL As String
sSQL = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & noperacion & "','" & sGlosa & "'," & nMovEstado & "," & nMovFlag & ")"
coConex.ConexionActiva.Execute sSQL
End Sub

Public Sub AgregaMovCap(ByVal nMovNro As Long, ByVal noperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double)
Dim sSQL As String

sSQL = "INSERT MovCap (nMovNro,cOpeCod,cCtaCod,nMonto,nSaldoDisponible,nSaldoContable) " _
    & "VALUES (" & nMovNro & ",'" & noperacion & "','" & sCuenta & "'," & nMonto & "," & nSaldoDisp & "," & nSaldoCnt & ")"
coConex.ConexionActiva.Execute sSQL
End Sub

Public Sub AgregaMovCapDet(ByVal nMovNro As Long, ByVal noperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nConcepto As CaptacConcepto, ByVal nMonto As Double)
Dim sSQL As String

sSQL = "INSERT MovCapDet (nMovNro,cOpeCod,cCtaCod,nConceptoCod,nMonto) " _
    & "VALUES (" & nMovNro & ",'" & noperacion & "','" & sCuenta & "'," & nConcepto & "," & nMonto & ")"
coConex.ConexionActiva.Execute sSQL
End Sub

Public Sub ActualizaSaldoAnteriorAho(ByVal sCuenta As String, ByVal nSaldoAnt As Double)
Dim sSQL As String

sSQL = "Update CaptacAhorros Set nSaldoAnterior = " & nSaldoAnt & "  WHERE cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSQL
End Sub

Public Sub ActualizaAbonoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)
Dim sSQL As String

If bActExtracto Then
    sSQL = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSQL = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
coConex.ConexionActiva.Execute sSQL
sSQL = "Update Captaciones Set nSaldoDisp = nSaldoDisp + " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSQL

End Sub

Public Sub ActualizaCargoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)

Dim sSQL As String

If bActExtracto Then
    sSQL = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSQL = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
coConex.ConexionActiva.Execute sSQL
sSQL = "Update Captaciones Set nSaldoDisp = nSaldoDisp - " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSQL
End Sub

Public Sub UltimaActualizacionCuenta(ByVal sCuenta As String, ByVal sMovNro As String)
'Actualiza la fecha de ultima actualizacion de la cuenta de captaciones
Dim sSQL As String
sSQL = "Update Captaciones Set cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSQL
End Sub

Public Sub ActualizaEstadoCuenta(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, _
        ByVal dFecha As Date, ByVal sMovNro As String)
'Actualiza el ultimo estado a la tabla de producto
Dim sSQL As String
sSQL = "Update Producto Set nPrdEstado = " & nEstado & ", " _
    & "dPrdEstado = '" & Format$(dFecha, "mm/dd/yyyy") & "' WHERE " _
    & "cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSQL
'Agrega un registro mas de estado a la historia de estados de la cuenta
sSQL = "INSERT CaptacEstado (cCtaCod,nPrdEstado,cMovNro) " _
    & "VALUES ('" & sCuenta & "'," & nEstado & ",'" & sMovNro & "')"
coConex.ConexionActiva.Execute sSQL
End Sub

Public Sub ActualizaEstadoDocRecEst(ByVal nTipoDoc As TpoDoc, ByVal sNroDoc As String, _
        ByVal sCodIF As String, ByVal sMovNro As String, ByVal nEstado As ChequeEstado)
'Actualiza el ultimo estado a la tabla de Documento Recibidos
Dim sSQL As String
sSQL = "Update DocRec Set nEstado = " & nEstado & " WHERE " _
    & "nTpoDoc = '" & nTipoDoc & "' AND cNroDoc = '" & sNroDoc & "' AND cPersCod = '" & sCodIF & "'"
coConex.ConexionActiva.Execute sSQL
'Agrega un registro mas de estado a la historia de estados del documento recibido
sSQL = "INSERT DocRecEst (nTpoDoc,cNroDoc,cPersCod,cMovNro,nEstado) " _
    & "VALUES ('" & nTipoDoc & "','" & sNroDoc & "','" & sCodIF & "','" & sMovNro & "'," & nEstado & ")"
coConex.ConexionActiva.Execute sSQL
End Sub

Public Function GetMovExtorno(ByVal sDatoBus As String, ByVal pdFecSis As Date, Optional nTipoBus As Integer = 0) As Recordset
Dim rsMov As Recordset
Dim sWhere As String
Dim sSQL As String
If nTipoBus = 0 Then
    sWhere = "M.cMovNro LIKE '%" & sDatoBus & "%'"
ElseIf nTipoBus = 1 Then
    sWhere = "M.cMovNro LIKE '" & Format$(pdFecSis, "yyyymmdd") & "%' And C.cCtaCod = '" & sDatoBus & "'"
End If
sSQL = "Select M.cMovNro, O.cOpeDesc, C.cCtaCod, ISNULL(MD.nDocTpo,'') nDocTpo, ISNULL(MD.cDocNro,'') cDocNro, " _
    & "M.cMovDesc, C.cOpeCod, C.nMonto, M.nMovNro FROM MovDoc MD RIGHT JOIN Mov M INNER JOIN MovCap C " _
    & "INNER JOIN " & csCentralCom & "OpeTpo O ON C.cOpeCod = O.cOpeCod ON M.nMovNro = C.nMovNro " _
    & "ON MD.nMovNro = M.nMovNro WHERE " & sWhere & " AND C.cOpeCod NOT IN ('" & gAhoEstInacAct & "') " _
    & "AND M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & "," & gMovFlagDeExtorno & ")" _
    & "AND M.nMovEstado IN (" & gMovEstContabMovContable & ") " _
    & "ORDER BY M.nMovNro"
Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSQL, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetMovExtorno = rsMov
End Function

Public Sub ActualizaEstadoMov(ByVal nMovNro As Long, ByVal nMovFlag As MovFlag)
Dim sSQL As String
sSQL = "Update Mov Set nMovFlag = " & nMovFlag & " Where nMovNro = " & nMovNro
coConex.ConexionActiva.Execute sSQL
End Sub

Public Function EsOrdenPagoEmitida(ByVal sCuenta As String, ByVal nNumOP As Long) As Boolean
Dim rsOP As Recordset
Set rsOP = New Recordset
Dim sSQL As String
rsOP.CursorLocation = adUseClient
sSQL = "Select OP.cCtaCod, OP.nInicio, OP.nFin FROM Mov M INNER JOIN MovDocEmitidoRango OP ON " _
    & "M.nMovNro = OP.nMovNro WHERE M.nMovFlag <> " & gMovFlagExtornado & " AND OP.cCtaCod = '" _
    & sCuenta & "' And " & nNumOP & " Between OP.nInicio And OP.nFin"
rsOP.Open sSQL, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsOP.ActiveConnection = Nothing
If rsOP.EOF And rsOP.BOF Then
    EsOrdenPagoEmitida = False
Else
    EsOrdenPagoEmitida = True
End If
rsOP.Close
Set rsOP = Nothing
End Function

Public Function ImprimeBoleta(ByVal psTit As String, ByVal psText As String, ByVal psCodOpe As String, ByVal pnMonto As String, _
            ByVal psCliente As String, ByVal psCodCta As String, ByVal psNumDoc As String, ByVal pnSaldo As Double, _
            ByVal pnInteresA As String, NomDoc As String, ByVal pnNumExt As Long, ByVal pnSaldoC As Double, _
            Optional pbOpSaldoC As Boolean = True, Optional pbSaldoInt As Boolean = True, Optional psNumDias As String = "----", _
            Optional psNomAgeRem As String = "", Optional psCodUsuRem As String = "", Optional pbCuenta As Boolean = False, _
            Optional psComCmac As String = "XXX", Optional psLin3 As String = "XXX", Optional psTexto As String = "XXX", _
            Optional pdFecSis As Date = CDate("01/01/1950"), Optional psNomAge As String = "") As String

Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMonto As String
Dim sSDisp As String
Dim sIntAcum As String
Dim sMax As Integer
Dim saux As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsNroExt As String
Dim lnTope As Integer
Dim lsSaldoC As String
Dim lsCadArg As String
Dim lsInteres As String
Dim sCad As String
Dim lnCliAux As Long
Dim lsCliAux1 As String
Dim lsCliAux2 As String

Dim lnChq As Long
Dim lsChqAux1 As String
Dim lsChqAux2 As String
Dim lsNomAge As String

Dim lnNumLinCmac As Integer

Dim lsMensaje As String * 39

'ETIQ:
'
'On Error GoTo ERROR
'
'lnTope = 0 '6 'Tope de lineas en Boleta
'
'lsNegritaOn = Chr$(27) + Chr$(71)
'lsNegritaOff = Chr$(27) + Chr$(72)
'
'lsNroExt = Str(pnNumExt)
'
'nFicSal = FreeFile
'Open gsLpt For Output As nFicSal
'
'Print #nFicSal, Chr$(27) & Chr$(64);
'
'Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
'Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
'Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
'Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
'Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
'Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita
'
'sSep = 15
'sIni = 1
'sMax = 33
'sAux = 5
'
'
'sFecha = Format$(pdFecSis, "dd/mm/yyyy")
'sHora = Format$(Time, "hh:mm:ss")
'sMonto = Format$(pnMonto, "#,##0.00")
'sSDisp = Format$(pnSaldo, "#,##0.00")
'lsSaldoC = Format$(pnSaldoC, "#,##0.00")
'
'lsNomAge = psNomAge
'
''Print #nFicSal, Chr$(10);
'Print #nFicSal, lsNegritaOn; 'Activa Negrita
'Print #nFicSal, Tab(sIni); "CMACT - AHORRO"; Space(19 + sSep + sAux); "CMACT - AHORRO"
'
'If Mid(psCodCta, 9, 1) = 1 Then
'    Print #nFicSal, Tab(sIni); Trim(psNomAge) & "-SOLES"; Space(sAux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(psNomAge) & "-SOLES"; Space(sAux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt;
'Else
'    Print #nFicSal, Tab(sIni); Trim(psNomAge) & "-DOLARES"; Space(sAux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt; Space(sSep); Trim(psNomAge) & "-DOLARES"; Space(sAux + sMax - Len(Trim(psNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt;
'End If
'
'If psNomAgeRem = "" Then
'    Print #nFicSal, ""
'Else
'    Print #nFicSal, Tab(sIni); "Ag.Rem: " & Trim(psNomAgeRem); Space(sAux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))); "Ag.Rem: " & Trim(psNomAgeRem)
'End If
'
'If psComCmac = "XXX" Then
'    If psLin3 = "XXX" Then
'        Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
'    Else
'        Print #nFicSal, Tab(sIni); psLin3 & Space(sAux + sSep + sMax - Len(psLin3)) & psLin3 & lsNegritaOff;  'Desactiva Negrita
'        lnNumLinCmac = 1
'    End If
'    lnNumLinCmac = 0
'Else
'    Print #nFicSal, Tab(sIni); "NroDocCmac:" & psComCmac & Space(sAux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & "NroDocCmac:" & psComCmac & lsNegritaOff;  'Desactiva Negrita
'    lnNumLinCmac = 1
'End If
'
'Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(sAux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora
'
''psCliente = PstaNombre(psCliente)
'
'lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
'
'If lnCliAux = 0 Then
'    If sAux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + sAux)
'    Print #nFicSal, Tab(sIni); ImpreCarEsp(psCliente); Space(sAux + sMax + sSep - Len(psCliente)); ImpreCarEsp(psCliente)
'Else
'    lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
'    lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
'
'    If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + sAux)
'    If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + sAux)
'
'    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux1); Space(sAux + sMax + sSep - Len(lsCliAux1)); ImpreCarEsp(lsCliAux1)
'    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux2); Space(sAux + sMax + sSep - Len(lsCliAux2)); ImpreCarEsp(lsCliAux2)
'
'    lnCliAux = 1
'End If
'
'If pbSaldoInt Or pbCuenta Then
'    Print #nFicSal, Tab(sIni); "Cuenta:" & psCodCta; Space(8 + sSep + sAux); "Cuenta:" & psCodCta
'Else
'    Print #nFicSal, ""
'End If
'
'psTit = Trim(psTit)
'psTit = CentrarCadena(psTit, 28)
'Print #nFicSal, lsNegritaOn; 'Activa Negrita
'Print #nFicSal, Tab(sIni + 1); "-----" & psTit & "-----"; Space(-1 + sSep); "-----" & psTit & "-----"
'
'
'lnChq = InStr(1, psText, "*", vbTextCompare)
'
'If psTexto = "XXX" Then
'    If lnChq = 0 Then
'        Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(Mid(psText, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)); sMonto; Space(-1 + sSep); ImpreCarEsp(Trim(Mid(psText, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)); sMonto
'        Print #nFicSal, ""
'    Else
'        lsChqAux1 = (Mid(psText, 1, lnChq - 1))
'        lsChqAux2 = (Mid(psText, lnChq + 1))
'        Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(lsChqAux1)); Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)); sMonto; Space(-1 + sSep); ImpreCarEsp(Trim(lsChqAux1)); Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)); sMonto
'        Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(lsChqAux2)); Space(sMax + 6 - Len(Trim(lsChqAux2))); Space(-1 + sSep); ImpreCarEsp(Trim(lsChqAux2)); Space(sMax + 6 - Len(Trim(lsChqAux2)))
'    End If
'Else
'    Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(psTexto)); Space(sAux + sSep + sMax - Len(Trim(psTexto))); ImpreCarEsp(Trim(psTexto))
'    Print #nFicSal, ""
'End If
'
'
'Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
'
'If pbSaldoInt Then
'    If MsgBox("Desea Imprimir el Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
'        Print #nFicSal, Tab(sIni); "Saldo Disponible"; Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)); sSDisp; Space(-1 + sSep); "Saldo Disponible"; Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)); sSDisp
'        If pbOpSaldoC Then
'            Print #nFicSal, Tab(sIni); "Saldo Contable"; Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)); lsSaldoC; Space(-1 + sSep); "Saldo Contable"; Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)); lsSaldoC
'        Else
'            Print #nFicSal, ""
'        End If
'    Else
'        Print #nFicSal, ""
'        Print #nFicSal, ""
'        pbSaldoInt = False
'    End If
'Else
'    Print #nFicSal, ""
'End If
'
'lsInteres = pnInteresA
'
'If pbSaldoInt Then
'    If lsInteres <> "No Valido" Then
'        lsInteres = Format(lsInteres, "#,##0.00")
'        Print #nFicSal, Tab(sIni); "Interes del Mes"; Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)); lsInteres; Space(-1 + sSep); "Interes del Mes"; Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)); lsInteres
'    End If
'Else
'    Print #nFicSal, ""
'End If
'
'If Not psNumDoc = "" Then
'    Print #nFicSal, Tab(sIni); NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc; Space(-1 + sSep); NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc
'Else
'    Print #nFicSal, ""
'End If
'
'If Not psNumDias = "----" Then
'    Print #nFicSal, Tab(sIni); "Nro Dias Interes"; Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)); psNumDias; Space(-1 + sSep); "Nro Dias Interes"; Space(sMax + 6 - Len(psNumDias) - Len("Nro Dias Interes")); psNumDias
'    lnTope = 4 - lnCliAux
'Else
'    lnTope = 3 - lnCliAux
'End If
'
'Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
'If psCodUsuRem = "" Then
'    Print #nFicSal, Tab(sIni); ImpreCarEsp(gsCodUser); Space(29 + sSep + sAux); ImpreCarEsp(gsCodUser)
'Else
'    Print #nFicSal, Tab(sIni); "Loc/Rem"; Space(sMax + sAux - Len("Loc/Rem") - 1 - 8); ImpreCarEsp(gsCodUser) & "/"; ImpreCarEsp(psCodUsuRem); Space(sSep); "Loc/Rem"; Space(sMax + sAux - Len("Loc/Rem") - 1 - 8); ImpreCarEsp(gsCodUser) & "/"; ImpreCarEsp(psCodUsuRem)
'End If
'Dim clsGen As DGeneral
'Set clsGen = New DGeneral
'lsMensaje = clsGen.GetMensajeBoletas(psCodCta)
'Set clsGen = Nothing
'Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsMensaje); Space(-1 + sSep); ImpreCarEsp(lsMensaje); lsNegritaOff
'
'lnNumLinCmac = lnNumLinCmac + 1
'
''Print #nFicSal, Chr$(12)
'For sAux = 1 To (lnTope - lnNumLinCmac)
'    Print #nFicSal, ""
'Next sAux
'Close nFicSal
'Exit Function
'ERROR:
'    Close nFicSal
'    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
'        GoTo ETIQ
'    End If
End Function

Public Function GetCapParametro(ByVal nParametro As CaptacParametro) As Double
Dim rsVar As Recordset
Dim sSQL As String
sSQL = "SELECT nParValor FROM Parametro WHERE nParCod = " & nParametro & " " _
    & "And nParProd = " & gPrdParamCaptac
Set rsVar = New Recordset
rsVar.CursorLocation = adUseClient
rsVar.Open sSQL, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsVar.ActiveConnection = Nothing
If rsVar.EOF And rsVar.BOF Then
    GetCapParametro = 0
Else
    GetCapParametro = rsVar("nParValor")
End If
rsVar.Close
Set rsVar = Nothing
End Function


Public Function ValidaSaldoCuenta(ByVal sCuenta As String, ByVal nMonto As Double) As Boolean
Dim nMoneda As Moneda
Dim clsMant As NCapMantenimiento
Dim rsCta As Recordset
Dim nSaldo As Double, nMontoMinimo As Double
Set clsMant = New NCapMantenimiento
Set rsCta = clsMant.GetDatosCuenta(sCuenta)
nSaldo = rsCta("nSaldoDisp")
rsCta.Close
Set rsCta = Nothing
Set clsMant = Nothing
nMoneda = CLng(Mid(sCuenta, 9, 1))
If nMoneda = gMonedaNacional Then
    nMontoMinimo = GetCapParametro(gSaldMinAhoMN)
Else
    nMontoMinimo = GetCapParametro(gSaldMinAhoME)
End If
If nSaldo - nMontoMinimo - nMonto >= 0 Then
    ValidaSaldoCuenta = True
Else
    ValidaSaldoCuenta = False
End If
End Function

Private Function EmiteBoleta(ByVal sMsgProd As String, ByVal sMsgOpe As String, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal noperacion As CaptacOperacion, ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double, _
            ByVal nIntMes As Double, ByVal nExtracto As Long, Optional bDocumento As Boolean = False, Optional nDocumento As TpoDoc = TpoDocCheque, _
            Optional sNroDoc As String = "", Optional dFechaValor As Date, Optional bImpSaldos As Boolean = True, _
            Optional pdFecSis As Date = CDate("01/01/1950"), Optional psNomAge As String = "") As String

Dim bReImp As Boolean
Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(noperacion)
sModDep = sMsgOpe
sTipApe = sMsgProd
Dim clsMant As NCapMantenimiento
Set clsMant = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsMant.GetNombreTitulares(sCuenta))
Set clsMant = Nothing
bReImp = False
Do
    If bDocumento Then
        Select Case nDocumento
            Case TpoDocCheque
                ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntMes, "Fecha Valor", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecSis, psNomAge
            Case TpoDocNotaAbono, TpoDocNotaCargo
                ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntMes, "", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecSis, psNomAge
            Case TpoDocOrdenPago
                ImprimeBoleta sTipApe, ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntMes, "", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecSis, psNomAge
        End Select
    End If
    If MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        bReImp = True
    Else
        bReImp = False
    End If
Loop Until Not bReImp
End Function


Public Function GetInteres(ByVal nCapital As Double, ByVal nTasa As Double, _
            ByVal nPlazo As Long, Optional nTipoInteres As TipoCalculoInteres = TpoCalcIntSimple) As Double
If nTipoInteres = TpoCalcIntSimple Then
    GetInteres = Round((nTasa / 36000) * nPlazo * nCapital, 2)
ElseIf nTipoInteres = TpoCalcIntCompuesto Then
    GetInteres = Round((((nTasa / 36000) + 1) ^ nPlazo - 1) * nCapital, 2)
End If
End Function


Public Function CapAbonoCuentaAho(ByRef pMatAho As Variant, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal noperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", Optional sCodIF As String = "", _
            Optional dFechaValor As Date, Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
            Optional dFecSis As Date = CDate("01/01/1950"), Optional sNomAge As String = "") As Double

Dim clsMant As DCapMantenimiento
'Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nSaldoInac As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, nExtracto As Long
Dim sMsgOpe As String
Dim nMovNro As Long
bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, True, False)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion
'Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
'clsCap.dbCmact.BeginTrans
'bTrans = True
If sGlosa = "" Then sGlosa = "Abono Cuenta = " & sCuenta
If bActivaCta Then
    If bInactiva Then
        nSaldoInac = nSaldoCnt
    End If
End If
'Calcula intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
pMatAho(2) = Format(nIntGanado, "#0.00")

dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt + nMonto
'Call AgregaMov(sMovNro, nOperacion, sGlosa)
nMovNro = GetnMovNro(sMovNro)
If sNroDoc = "" Then ' Si la operacion es en efectivo
    Call ActualizaSaldoAnteriorAho(sCuenta, nSaldoDisp)
    nSaldoDisp = nSaldoDisp + nMonto
    Call ActualizaAbonoCaptacion(sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True)
    sMsgOpe = "Depósito Efectivo"
Else
    If nTipoDoc = TpoDocCheque Then 'Si el abono es con cheque
        nSaldoDisp = nSaldoDisp + 0
        nIntGanado = 0
        Call ActualizaAbonoCaptacion(sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True)
        Call AgregaCuentaDocumento(sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMovNro)
        sMsgOpe = "Depósito Cheque"
    ElseIf nTipoDoc = TpoDocNotaAbono Then 'Si el abono es con nota de abono
        Call ActualizaSaldoAnteriorAho(sCuenta, nSaldoDisp)
        nSaldoDisp = nSaldoDisp + nMonto
        Call ActualizaAbonoCaptacion(sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True)
        sMsgOpe = "Depósito Nota Abono"
    End If
End If
    Call AgregaMovCap(nMovNro, noperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt)
If bInactiva Then
    Call AgregaMovCapDet(nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac)
    Call ActualizaEstadoCuenta(sCuenta, gCapEstActiva, dFecSis, sMovNro)
End If
    Call AgregaMovCapDet(nMovNro, noperacion, sCuenta, gConcCapital, nMonto)
    Call UltimaActualizacionCuenta(sCuenta, sMovNro)
CapAbonoCuentaAho = nSaldoCnt

pMatAho(8) = nSaldoCnt
pMatAho(9) = nSaldoCnt

'If sNroDoc = "" Then
'    EmiteBoleta "DEPOSITO AHORROS", sMsgOpe, sCuenta, nMonto, nOperacion, nSaldoDisp, nSaldoCnt, nIntGanado, nExtracto, , , , , , dFecSis, sNomAge
'Else
'    EmiteBoleta "DEPOSITO AHORROS", sMsgOpe, sCuenta, nMonto, nOperacion, nSaldoDisp, nSaldoCnt, nIntGanado, nExtracto, True, nTipoDoc, sNroDoc, dFechaValor, , dFecSis, sNomAge
'End If
Exit Function
ErrGraba:
    Err.Raise Err.Number, "", Err.Description
    CapAbonoCuentaAho = 0
End Function

Public Function CapCargoCuentaAho(ByRef pMatAho As Variant, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal noperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
            Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional bOPExiste As Boolean = False, _
            Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
            Optional sCodCMAC As String = "", Optional sNomCmac As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "", Optional bCancelaciones As Boolean = False) As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nSaldoInac As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, i As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date

Dim clsMov As NContFunciones

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
Set clsMant = Nothing

nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion
Set clsCap = New DCapMovimientos

bTrans = True
Randomize
For i = 0 To Rnd(2000) * 1000
Next i
If Not ValidaSaldoCuenta(sCuenta, nMonto) Then 'Valida el saldo de la cuenta nuevamente
    Err.Raise 1000, "CapCargoCuentaAho", "Cuenta NO Posee Saldo Suficiente"
    CapCargoCuentaAho = 0
    Exit Function
End If
If bActivaCta Then
    If bInactiva Then
        nSaldoInac = nSaldoCnt
    End If
End If
'Calcula los intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
If Not bCancelaciones Then
    pMatAho(3) = Format(nIntGanado, "#0.00")
Else
    pMatAho(4) = Format(nIntGanado, "#0.00")
End If

dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt - nMonto

ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp

nSaldoDisp = nSaldoDisp - nMonto
If Not bCancelaciones Then
    pMatAho(10) = nSaldoDisp
    pMatAho(11) = nSaldoCnt
Else
    pMatAho(12) = nSaldoDisp
    pMatAho(13) = nSaldoCnt
End If

ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True

If sNroDoc <> "" Then
    If nTipoDoc = TpoDocOrdenPago Then 'Si es con orden de pago
        If bOPExiste Then
            AgregaOrdenPagoEstado sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstCobrada
        Else
            AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMonto, gCapOPEstCobrada
        End If
        If noperacion = gAhoRetOPCanje Then
            sMsgOpe = "Retiro OP Canje"
        ElseIf noperacion = gAhoRetOP Then
            sMsgOpe = "Retiro OP"
        ElseIf noperacion = gAhoOPCertificacion Then
            sMsgOpe = "Certificacion OP"
        End If
    ElseIf nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = "Retiro NC"
    End If
Else
    sMsgOpe = "Retiro Efectivo"
End If

'AgregaMov sMovNro, nOperacion, sGlosa

nMovNro = GetnMovNro(sMovNro)
AgregaMovCap nMovNro, noperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
If bInactiva Then
    AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
    ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
End If
AgregaMovCapDet nMovNro, noperacion, sCuenta, gConcCapital, nMonto
'If sCodCMAC <> "" Then
'    AgregaMovCMAC nMovNro, sCodCMAC, CLng(Mid(sCuenta, 9, 1)), nOperacion, nMonto
'End If
UltimaActualizacionCuenta sCuenta, sMovNro
'If sPersLavDinero <> "" Then
'    clsCap.AgregaMovLavDinero nMovNro, sPersLavDinero
'End If

CapCargoCuentaAho = nSaldoCnt
bTrans = False

'Dim sTipDep As String, sCodOpe As String
'Dim sModDep As String, sTipApe As String
'Dim sNomTit As String
'sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
'sCodOpe = Trim(nOperacion)
'sModDep = sMsgOpe
'sTipApe = "RETIRO AHORROS"
'Dim clsTit As NCapMantenimiento
'Set clsTit = New NCapMantenimiento
'sNomTit = ImpreCarEsp(clsTit.GetNombreTitulares(sCuenta))
'Set clsTit = Nothing
'
'Do
'    If sNroDoc <> "" Then
'        Select Case nTipoDoc
'            Case TpoDocNotaCargo
'                ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
'            Case TpoDocOrdenPago
'                If sCodCMAC <> "" Then
'                    sTipApe = "RETIRO AHORROS CMAC"
'                    ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , sNomCMAC, , dFecSis, sNomAge, sCodUser, sLpt
'                Else
'                    ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
'                End If
'        End Select
'    Else
'        If sCodCMAC <> "" Then
'            sTipApe = "RETIRO AHORROS CMAC"
'            ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , sNomCMAC, , dFecSis, sNomAge, sCodUser, sLpt
'        Else
'            ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt
'        End If
'    End If
'Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
'Exit Function
'ErrGraba:
'    If bTrans Then clsCap.dbCmact.RollbackTrans
'    Set clsCap = Nothing
'    Err.Raise Err.Number, "", Err.Description
'    CapCargoCuentaAho = 0
End Function

Public Function CapAperturaCuenta(ByVal nProducto As Producto, ByVal nMoneda As Moneda, ByVal rsPers As Recordset, ByVal sMovNro As String, _
        ByVal sAgencia As String, ByVal noperacion As CaptacOperacion, ByVal nTasa As Double, ByVal nSaldo As Double, _
        ByVal dFecha As Date, ByVal nFirmas As Integer, ByVal nPersoneria As PersPersoneria, ByVal sGlosa As String, _
        ByVal nTipoCuenta As ProductoCuentaTipo, _
        ByVal nTipoTasa As CaptacTipoTasa, Optional bCheque As Boolean = False, _
        Optional sNroDoc As String, Optional sCodIF As String, _
        Optional ByVal bOrdPag As Boolean = False, Optional nPlazo As Long = 0, _
        Optional nFormaRetiro As CaptacPFFormaRetiro = gCapPFFormRetMensual, _
        Optional bCtaAboInt As Boolean = False, Optional sCtaCodAbono As String = "", _
        Optional nPorcRetCTS As Double = 0, Optional sInstitucion As String = "", _
        Optional sPersLavDinero As String = "") As String

Dim nErrNum As Integer
Dim nErrDesc As String, sCuenta As String
Dim bTrans As Boolean
Dim nMovNro As Long


bTrans = False


bTrans = True
sCuenta = AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo, dFecha, nFirmas, nPersoneria, _
    nTipoCuenta, sMovNro, nTipoTasa, bCheque, bOrdPag, nPlazo, nFormaRetiro, bCtaAboInt, sCtaCodAbono, nPorcRetCTS, sInstitucion)
'AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = GetnMovNro(sMovNro)
Select Case noperacion
    Case gAhoApeChq, gPFApeChq, gCTSApeChq
        AgregaMovCap nMovNro, noperacion, sCuenta, nSaldo, 0, nSaldo
        AgregaMovCapDet nMovNro, noperacion, sCuenta, gConcCapital, nSaldo
        AgregaCuentaDocumento sCuenta, TpoDocCheque, sNroDoc, sCodIF, sMovNro, nMovNro
    Case gAhoApeTransf, gPFApeTransf, gCTSApeTransf
        AgregaMovCap nMovNro, noperacion, sCuenta, nSaldo, nSaldo, nSaldo
        AgregaMovCapDet nMovNro, noperacion, sCuenta, gConcCapital, nSaldo
        'Aqui debe registrarse la referencia con la transferencia de caja general
    Case Else
        AgregaMovCap nMovNro, noperacion, sCuenta, nSaldo, nSaldo, nSaldo
        AgregaMovCapDet nMovNro, noperacion, sCuenta, gConcCapital, nSaldo
End Select
'If nOperacion = gCTSApeEfec Or nOperacion = gCTSApeChq Or nOperacion = gCTSApeTransf Then
'    AgregaMovParamCTS nMovNro, nPorcRetCTS
'End If
'If sPersLavDinero <> "" Then
'    clsMant.AgregaMovLavDinero nMovNro, sPersLavDinero
'End If
rsPers.MoveFirst
Do While Not rsPers.EOF
    AgregaNuevoProdPers sCuenta, rsPers("Codigo"), CLng(Trim(Right(rsPers("Relacion"), 4)))
    rsPers.MoveNext
Loop
CapAperturaCuenta = sCuenta

End Function

Public Function GetSaldoFecha(ByVal sCuenta As String, ByVal dFecha As Date) As Recordset
Dim sSQL As String, sNumMov As String
Dim rsCta As Recordset
sSQL = "Select TOP 1 CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha, " _
    & "C.nSaldoDisponible, C.nSaldoContable FROM Mov M INNER JOIN MovCap C " _
    & "INNER JOIN MovCapDet CD INNER JOIN OpeTpo O INNER JOIN CapMovTipo CT ON " _
    & "O.cOpeCod = CT.cOpeCod ON CD.cOpeCod = O.cOpeCod ON C.nMovNro = CD.nMovNro ON " _
    & "M.nMovNro = C.nMovNro WHERE C.cCtaCod = '" & sCuenta & "' AND CD.nConceptoCod " _
    & "IN (" & gConcCapital & ") AND M.cMovNro <= '" & Format$(dFecha, "yyyymmdd") & "235959' " _
    & "AND M.nMovFlag NOT IN (" & gMovFlagExtornado & ") ORDER BY M.nMovNro DESC"
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSQL, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetSaldoFecha = rsCta
Set rsCta = Nothing
End Function

'******************************************************************************************
'************************************** Final de Copiado de Clases de Ahorros **********************
'******************************************************************************************
Public Sub dUpdateGarantiaBloquea(ByVal pCnumGarant As String, ByVal pnEstado As Integer)
Dim sSQL As String

    sSQL = " UPDATE Garantias Set  nEstado = " & pnEstado & " Where cNumGarant = '" & pCnumGarant & "'"
    coConex.Ejecutar sSQL
    
End Sub
Public Sub dUpdateGarantiaLibera(ByVal pCnumGarant As String, ByVal pnEstado As Integer, ByVal pnGravado As Currency)
Dim sSQL As String

    sSQL = " UPDATE Garantias Set nGravament = nGravament - " & pnGravado & _
            ", nEstado = " & pnEstado & " Where cNumGarant = '" & pCnumGarant & "'"
    coConex.Ejecutar sSQL
   
End Sub


Public Function GeneraMovNro(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String
    On Error GoTo GeneraMovNroErr
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    Set rs = New ADODB.Recordset
    
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & Right(psCodAge, 2) & "','" & psUser & "'"
    Else
       sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    
    Set rs = coConex.Ejecutar(sql)
    If Not rs.EOF Then
        GeneraMovNro = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
GeneraMovNroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function


Public Sub ActualizaFechaCalendarioMatriz(ByVal psCtaCod As String, ByVal pMatCal As Variant)
Dim i As Integer
Dim oCred As DCredito
Dim R As ADODB.Recordset
Dim nNroCalen As Integer
    
        Set oCred = New DCredito
        Set R = oCred.RecuperaColocacCred(psCtaCod)
        Set oCred = Nothing
        nNroCalen = R!nNroCalen
        
        dBeginTrans
        For i = 0 To UBound(pMatCal) - 1
            Call dUpdateColocCalendario(psCtaCod, nNroCalen, CInt(pMatCal(i, 1)), gColocCalendAplCuota, pMatCal(i, 0))
        Next i
        dCommitTrans
End Sub

Public Sub dInsertColocCartaFianza(ByVal pcCtaCod As String, ByVal pnCondicion As ColocCredCondicion, _
    ByVal pnModalidad As ColCFModalidad, _
    ByVal pdFecAsig As Date, ByVal psFecVenc As Date, _
    Optional ByVal pcFinalidad As String = "", _
    Optional ByVal pdEjecutado As Date = CDate("01/01/1950"), Optional ByVal pbEjecBatch As Boolean = False)

Dim sSQL As String

        sSQL = "INSERT INTO ColocCartaFianza(cCtaCod, nCondicion, nModalidad, dEjecutado, cFinalidad, dAsignacion, dVencimiento) "
        sSQL = sSQL & " VALUES('" & pcCtaCod & "'," & pnCondicion & "," & pnModalidad & ",NULL,'" & pcFinalidad & "','" & Format(pdFecAsig, "mm/dd/yyyy") & "','" & Format(psFecVenc, "mm/dd/yyyy") & "')"
        If pbEjecBatch Then
            coConex.AdicionaCmdBatch sSQL
        Else
            coConex.Ejecutar sSQL
        End If
End Sub

Public Sub dUpdateColocCartaFianza(ByVal pcCtaCod As String, Optional ByVal pnCondicion As ColocCredCondicion = -1, _
    Optional ByVal pnModalidad As ColCFModalidad = -1, _
    Optional ByVal pdFecAsig As Date = CDate("01/01/1950"), Optional ByVal pdFecVenc As Date = CDate("01/01/1950"), _
    Optional ByVal psFinalidad As String = "", _
    Optional ByVal pdEjecutado As Date = CDate("01/01/1950"), Optional ByVal pbEjecBatch As Boolean = False)

Dim sSQL As String
        
        sSQL = "UPDATE ColocCartaFianza SET "
        If pnCondicion <> -1 Then
           sSQL = sSQL & " nCondicion = " & pnCondicion & ","
        End If
        If pnModalidad <> -1 Then
            sSQL = sSQL & " nModalidad = " & pnModalidad & ","
        End If
        If pdFecAsig <> CDate("01/01/1950") Then
            sSQL = sSQL & " dAsignacion = '" & Format(pdFecAsig, "mm/dd/yyyy") & "',"
        End If
        If pdFecVenc <> CDate("01/01/1950") Then
            sSQL = sSQL & " dVencimiento = '" & Format(pdFecVenc, "mm/dd/yyyy") & "',"
        End If
        If psFinalidad <> "" Then
            sSQL = sSQL & " cFinalidad = '" & psFinalidad & "',"
        End If
        If pdEjecutado <> CDate("01/01/1950") Then
            sSQL = sSQL & " dEjecutado = '" & Format(pdEjecutado, "mm/dd/yyyy") & "',"
        End If
        sSQL = Mid(sSQL, 1, Len(sSQL) - 1)
        sSQL = sSQL & " WHERE cCtaCod = '" & pcCtaCod & "' "
        If pbEjecBatch Then
            coConex.AdicionaCmdBatch sSQL
        Else
            coConex.Ejecutar sSQL
        End If
End Sub

Public Sub dInsertColocCFEstado(ByVal psCtaCod As String, ByVal pdFechaPrdEstado As String, _
        ByVal pnPrdEstado As Integer, ByVal pnMonto As Currency, ByVal pdVenc As String, _
        ByVal psDescripcion As String, pnMotivoRechazo As Integer, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim sSQL As String

        sSQL = "INSERT INTO ColocCFEstado(cCtaCod, dPrdEstado, nPrdEstado, nMonto, dVenc, cDescripcion, nMotivoRechazo ) "
        sSQL = sSQL & " VALUES('" & psCtaCod & "','" & Format(CDate(Format(pdFechaPrdEstado, "dd/mm/yyyy") & Space(1) & Format(dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "'," & pnPrdEstado & ","
        sSQL = sSQL & pnMonto & ",'" & Format(CDate(Format(pdVenc, "dd/mm/yyyy")), "mm/dd/yyyy") & "','" & psDescripcion & "',"
        sSQL = sSQL & pnMotivoRechazo & " )"
        If pbEjecBatch Then
            coConex.AdicionaCmdBatch sSQL
        Else
            coConex.Ejecutar sSQL
        End If
End Sub

Public Sub dDeleteColocCFEstado(ByVal psCtaCod As String, ByVal pnPrdEstado As Integer, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    lsSQL = "DELETE ColocCFEstado WHERE cCtaCod = '" & psCtaCod & "' AND nPrdEstado = " & pnPrdEstado
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dInsertColocCFCredito(ByVal psCtaCod As String, ByVal psCredito As String, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim sSQL As String

        sSQL = "INSERT INTO ColocCFCredito (cCtaCod, cCtaCodCre  ) "
        sSQL = sSQL & " VALUES('" & psCtaCod & "','" & psCredito & "')"
        If pbEjecBatch Then
            coConex.AdicionaCmdBatch sSQL
        Else
            coConex.Ejecutar sSQL
        End If
End Sub

Public Sub dDeleteColocCFCredito(ByVal psCtaCod As String, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    lsSQL = "DELETE ColocCFCredito WHERE cCtaCod = '" & psCtaCod & "' "
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Function RecuperaLineasCreditoSaldo(ByVal cLineaCred As String, Optional pbBloqueo As Boolean = False) As ADODB.Recordset
Dim sSQL As String
Dim R As ADODB.Recordset

    On Error GoTo ErrorRecuperaLineasCreditoSaldo
    sSQL = "Select nMontoTotal, nSaldoCap,nMontoColocado,nMoneda, nMontoReservado from ColocLineaCreditosaldo "
    If pbBloqueo Then
        sSQL = sSQL & " with (tablockx) "
    End If
    sSQL = sSQL & " Where cLineaCred = '" & cLineaCred & "'"
    Set R = New ADODB.Recordset
    R.CursorLocation = adUseClient
    R.Open sSQL, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    R.ActiveConnection = Nothing
    Set RecuperaLineasCreditoSaldo = R
    Exit Function
ErrorRecuperaLineasCreditoSaldo:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Sub dBeginTrans()
    'If pbConcurrencia Then
    '    coConex.ConexionActiva.Execute "SET TRANSACTION ISOLATION LEVEL SERIALIZABLE"
    'End If
    coConex.BeginTrans
End Sub

Public Sub dRollbackTrans()
    coConex.RollbackTrans
    coConex.Ejecutar "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
End Sub
Public Sub dCommitTrans()
    coConex.CommitTrans
    'coConex.Ejecutar "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
End Sub
Private Sub Class_Initialize()
    Dim loIni As ClasIni
    
    Set loIni = New ClasIni
        csConexion = loIni.CadenaConexion
        csNegocio = loIni.BaseNegocio
        csCentralPer = loIni.BasePersonas
        csCentralCom = loIni.BaseComunes
        csCentralImg = loIni.BaseImagenes
        csAdminist = loIni.BaseAdministracion
    Set loIni = Nothing
    
Set coConex = New DConecta
If coConex.AbreConexion(csConexion) = False Then
    Call RaiseError(MyUnhandledError, "DColPCalculos:Initialize. Error en Conexion a Base de datos")
End If

End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub

'Ejecuta procesos Batch
Public Function dEjecutaBatch() As Integer
On Error GoTo EjecutaBatchErr
    dEjecutaBatch = coConex.EjecutarBatch
    Exit Function
EjecutaBatchErr:
    Call RaiseError(MyUnhandledError, "DCredActualizaBD - Ejecutar Batch")
End Function

Public Sub dEliminaCuotasPendientesCalifMiViv(ByVal psCtaCod As String)
Dim sSQL As String

    sSQL = " DELETE ColocCalifMiViv Where cCtaCod = '" & psCtaCod & "' AND nColocCalendEstado = " & gColocCalendEstadoPendiente
    coConex.Ejecutar sSQL
    
End Sub

Public Sub dUpdateColocCalifMiViv(ByVal psCodCta As String, ByVal pnNroCalen As Integer, _
    ByVal pnCuota As Integer, _
    Optional ByVal pdPago As Date = CDate("01/01/1900"), _
    Optional ByVal psColocMiVivEval As String = "", Optional ByVal pnColocCalendEstado As Integer = -1, _
    Optional ByVal pdCondonacion As Date = CDate("01/01/1900"))

Dim sSQL As String

    sSQL = "UPDATE ColocCalifMiViv SET "
    
    If pdPago <> CDate("01/01/1900") Then
        sSQL = sSQL & " dPago = '" & Format(pdPago, "mm/dd/yyyy hh:mm:ss") & "', "
    End If
    
    If pdCondonacion <> CDate("01/01/1900") Then
        sSQL = sSQL & " dFecCond = '" & Format(pdCondonacion, "mm/dd/yyyy hh:mm:ss") & "', "
    End If
    
    If psColocMiVivEval <> "" Then
        sSQL = sSQL & " cColocMiVivEval = '" & psColocMiVivEval & "', "
    End If
    
    If pnColocCalendEstado <> -1 Then
        sSQL = sSQL & " nColocCalendEstado = " & pnColocCalendEstado & ", "
    End If
    
    sSQL = Mid(sSQL, 1, Len(sSQL) - 2)
    If psColocMiVivEval = "" Then
        sSQL = sSQL & " WHERE  cCtaCod = '" & psCodCta & "' AND nNroCalen = " & pnNroCalen & " AND nCuota = " & pnCuota
    Else
        sSQL = sSQL & " WHERE  cCtaCod = '" & psCodCta & "' AND nCuotaOrig = " & pnCuota
    End If
    coConex.Ejecutar sSQL
    
End Sub

Public Sub dInsertColocCalifMiViv(ByVal psCodCta As String, ByVal pnCuotaOrig As Integer, ByVal pnNroCalen As Integer, _
    ByVal pnCuota As Integer, ByVal pdVenc As Date, _
    ByVal pnColocCalendEstado As Integer)

Dim sSQL As String

    sSQL = " INSERT INTO ColocCalifMiViv(cCtaCod, nCuotaOrig, nNroCalen, nCuota, dVenc, "
    sSQL = sSQL & " nColocCalendEstado) "
    sSQL = sSQL & " VALUES('" & psCodCta & "'," & pnCuotaOrig & "," & pnNroCalen & "," & pnCuota & ","
    sSQL = sSQL & "'" & Format(pdVenc, "mm/dd/yyyy") & "',"
    sSQL = sSQL & pnColocCalendEstado & ")"
    
    coConex.Ejecutar sSQL
    
End Sub

Public Sub dUpdateLineaCreditoSaldo(ByVal psLineaCred As String, Optional pnMontoTotal As Double = -1, Optional pnSaldoCap As Double = -1, Optional pnMontoColocado As Double = -1, Optional pnMoneda As Integer = -1, Optional ByVal pnMontoReservado As Double = -1, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    On Error GoTo ErrordUpdateLineaCreditoSaldo
    lsSQL = "UPDATE ColocLineaCreditoSaldo SET "
    
    If pnMontoReservado <> -1 Then
        lsSQL = lsSQL & " nMontoReservado = " & Format(pnMontoReservado, "#0.00") & ","
    End If
    If pnMontoTotal <> -1 Then
        lsSQL = lsSQL & " nMontoTotal = " & Format(pnMontoTotal, "#0.00") & ","
    End If
    If pnSaldoCap <> -1 Then
        lsSQL = lsSQL & " nSaldoCap = " & Format(pnSaldoCap, "#0.00") & ","
    End If
    If pnMontoColocado <> -1 Then
        lsSQL = lsSQL & " nMontoColocado = " & Format(pnMontoColocado, "#0.00") & ","
    End If
    If pnMoneda <> -1 Then
        lsSQL = lsSQL & " nMoneda = " & Format(pnMoneda, "#0.00") & ","
    End If
    lsSQL = Mid(lsSQL, 1, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cLineaCred = '" & psLineaCred & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub
ErrordUpdateLineaCreditoSaldo:
    Err.Raise Err.Number, "Error En Proceso dUpdateLineaCreditoSaldo", Err.Description

End Sub


Public Sub dDeleteProductoTasaInteres(ByVal psCtaCod As String, ByVal pnPrdTasaInteres As Integer, Optional ByVal pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    On Error GoTo ErrordDeleteProductoTasaInteres
    
    lsSQL = " DELETE ProductoTasaInteres WHERE cCtaCod = '" & psCtaCod & "' AND nPrdTasaInteres = " & pnPrdTasaInteres
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
    Exit Sub

ErrordDeleteProductoTasaInteres:
    Err.Raise Err.Number, "Error En Proceso dDeleteProductoTasaInteres", Err.Description
End Sub

Public Sub dDeleteProductoPersona(ByVal psCtaCod As String, Optional ByVal pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    On Error GoTo ErrordDeleteProductoTasaInteres
    
    lsSQL = " DELETE ProductoPersona WHERE cCtaCod = '" & psCtaCod & "' "
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
    Exit Sub

ErrordDeleteProductoTasaInteres:
    Err.Raise Err.Number, "Error En Proceso dDeleteProductoTasaInteres", Err.Description
End Sub


Public Sub dUpdateProductoTasaInteres(ByVal psCtaCod As String, ByVal pnPrdTasaInteres As Integer, Optional ByVal pnTasaInteres As Double = -1, _
    Optional ByVal pbEjecBatch As Boolean = False)
    
Dim lsSQL As String

    On Error GoTo ErrordInsertProductoTasaInteres
    lsSQL = " UPDATE ProductoTasaInteres SET "
    'If pnPrdTasaInteres <> -1 Then
    '    lsSQL = lsSQL & " nPrdTasaInteres = " & Format(pnPrdTasaInteres, "#0.0000") & ","
    'End If
    If pnTasaInteres <> -1 Then
        lsSQL = lsSQL & " nTasaInteres = " & Format(pnTasaInteres, "#0.0000") & ","
    End If
    lsSQL = Mid(lsSQL, 1, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "' AND nPrdTasaInteres = " & pnPrdTasaInteres
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

ErrordInsertProductoTasaInteres:
                  Err.Raise Err.Number, "Error En Proceso dInsertProductoTasaInteres", Err.Description
End Sub



Public Sub dInsertProductoTasaInteres(ByVal psCtaCod As String, ByVal pnPrdTasaInteres As Integer, ByVal pnTasaInteres As Double, _
    Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    On Error GoTo ErrordInsertProductoTasaInteres
    
    lsSQL = "INSERT ProductoTasaInteres (cCtaCod,nPrdTasaInteres,nTasaInteres) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "','" & pnPrdTasaInteres & "'," & Format(pnTasaInteres, "#0.0000") & ")"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

    Exit Sub

ErrordInsertProductoTasaInteres:
    Err.Raise Err.Number, "Error En Proceso dInsertProductoTasaInteres", Err.Description
    
End Sub

Public Sub dUpdateColocacConvenio(ByVal psCtaCod As String, Optional ByVal psPersCod As String = "", Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal psCodMod As String = "")
Dim lsSQL As String
    
    On Error GoTo ErrordUpdateColocacConvenio
    lsSQL = "UPDATE ColocacConvenio SET "
    If psPersCod <> "" Then
        lsSQL = lsSQL & " cPersCod = '" & psPersCod & "' "
    End If
    If psCodMod <> "" Then
        If psPersCod <> "" Then
            lsSQL = lsSQL & " , "
        End If
            lsSQL = lsSQL & " cCodModular = '" & psCodMod & "' "
    End If
    lsSQL = lsSQL & " Where cCtaCod = '" & psCtaCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

ErrordUpdateColocacConvenio:
    Err.Raise Err.Number, "Error En Proceso dUpdateColocacConvenio", Err.Description
End Sub

Public Sub EliminaGastosExonerados(ByVal psCtaCod As String)
Dim lsSQL As String
    
    lsSQL = "DELETE ColocCredGastosExon Where cCtaCod = '" & psCtaCod & "'"
    coConex.Ejecutar lsSQL
    
End Sub

Public Sub dInsertColocCredGastosExon(ByVal psCtaCod As String, ByVal pnPrdConceptoCod As Long)
Dim lsSQL As String

    lsSQL = "INSERT ColocCredGastosExon(cCtaCod, nPrdConceptoCod) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "'," & pnPrdConceptoCod & ")"
    
    coConex.Ejecutar lsSQL
    
End Sub

Public Sub dUpdateColocacCred(ByVal psCtaCod As String, Optional ByVal pnDiasAtraso As Integer = -1, Optional ByVal pnColocCondicion As Integer = -1, _
    Optional ByVal pnColocDestino As Integer = -1, Optional ByVal psProtesto As String = "", Optional ByVal pbCargoAuto As Integer = -1, _
    Optional ByVal psMetLiquidacion As String = "", Optional ByVal pbRefCapInt As Integer = -1, Optional ByVal pnNroProxCuota As Integer = -1, _
    Optional ByVal pnIntPend As Double = -1, Optional ByVal pnExoPenalidad As Integer = -1, Optional ByVal pnTpoCalend As Integer = -1, _
    Optional ByVal pnCalendDinamico As Integer = -1, Optional ByVal pnTipoDesembolso As Integer = -1, Optional ByVal pnNroCalen As Integer = -1, Optional ByVal pnNroProxDesemb As Integer = -1, _
    Optional psPersFondo As String = "", Optional pnMoneda As Integer = -1, Optional ByVal pbEjecBatch As Boolean = False, _
    Optional ByVal pnCuotaComodin As Integer = -1, Optional ByVal pnMiViv As Integer = -1, _
    Optional ByVal pnNroCalParalelo As Integer = -1, Optional ByVal pnCalifPagos As Integer = -1, _
    Optional ByVal pbPrepago As Integer = -1, Optional ByVal pnCalendDinamTipo As Integer = -1)
    
Dim lsSQL As String
    On Error GoTo ErrordUpdateColocacCred
  
    lsSQL = "UPDATE ColocacCred SET "
    If pnMoneda <> -1 Then
        lsSQL = lsSQL & " nFondoMoneda = " & pnMoneda & ","
    End If
    If pnCalifPagos <> -1 Then
        lsSQL = lsSQL & " nCalPago = " & pnCalifPagos & ","
    End If
    If pnCuotaComodin <> -1 Then
        lsSQL = lsSQL & " bCuotaCom = " & pnCuotaComodin & ","
    End If
    If pnMiViv <> -1 Then
        lsSQL = lsSQL & " bMiVivienda = " & pnMiViv & ","
    End If
    If pnNroCalParalelo <> -1 Then
        lsSQL = lsSQL & " nNroCalPar = " & pnNroCalParalelo & ","
    End If
    If psPersFondo <> "" Then
        lsSQL = lsSQL & " cPersCod = '" & psPersFondo & "',"
    End If
    If pnNroProxDesemb <> -1 Then
        lsSQL = lsSQL & " nNroProxDesemb = " & pnNroProxDesemb & ","
    End If
    If pnDiasAtraso <> -1 Then
        lsSQL = lsSQL & " nDiasAtraso = " & pnDiasAtraso & ","
    End If
    If pnColocCondicion <> -1 Then
        lsSQL = lsSQL & " nColocCondicion = " & pnColocCondicion & ","
    End If
    If pnColocDestino <> -1 Then
        lsSQL = lsSQL & " nColocDestino = " & pnColocDestino & ","
    End If
    If psProtesto <> "" Then
        lsSQL = lsSQL & " cProtesto = '" & psProtesto & "',"
    End If
    If pbCargoAuto <> -1 Then
        lsSQL = lsSQL & " bCargoAuto = " & pbCargoAuto & ","
    End If
    If psMetLiquidacion <> "" Then
        lsSQL = lsSQL & " cMetLiquidacion = '" & psMetLiquidacion & "',"
    End If
    If pbRefCapInt <> -1 Then
        lsSQL = lsSQL & " bRefCapInt = " & pbRefCapInt & ","
    End If
    If pnNroProxCuota <> -1 Then
        lsSQL = lsSQL & " nNroProxCuota = " & pnNroProxCuota & ","
    End If
    If pnIntPend <> -1 Then
        lsSQL = lsSQL & " nIntPend = " & Format(pnIntPend, "#0.00") & ","
    End If
    If pnExoPenalidad <> -1 Then
        lsSQL = lsSQL & " nExoPenalidad = " & pnExoPenalidad & ","
    End If
    If pnTpoCalend <> -1 Then
        lsSQL = lsSQL & " nColocCalendCod = " & pnTpoCalend & ","
    End If
    If pnNroCalen <> -1 Then
        lsSQL = lsSQL & " nNroCalen = " & pnNroCalen & ","
    End If
    If pnCalendDinamico <> -1 Then
        lsSQL = lsSQL & " nCalendDinamico = " & pnCalendDinamico & ","
    End If
    If pnTipoDesembolso <> -1 Then
        lsSQL = lsSQL & " nTipoDesembolso = " & pnTipoDesembolso & ","
    End If
    If pbPrepago <> -1 Then
        lsSQL = lsSQL & " bPrepago = " & pbPrepago & ","
    End If
    If pnCalendDinamTipo <> -1 Then
        lsSQL = lsSQL & " nCalendDinamTipo = " & pnCalendDinamTipo & ","
    End If
    
    lsSQL = Mid(lsSQL, 1, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "'"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

ErrordUpdateColocacCred:
    Err.Raise Err.Number, "Error En Proceso dUpdateColocacCred", Err.Description

End Sub

Public Sub dInsertColocRecupExpediente(ByVal psCodCta As String, ByVal psNumExp As String, _
    ByVal pnMonPetit As Double, ByVal pnMoneda As Integer, ByVal pnViaProce As Integer, _
    ByVal pnEstadoProceso As Integer, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    On Error GoTo ErrordInsertExpedJud
    
    lsSQL = "INSERT INTO ColocRecupExpediente(cCtaCod, cNumExp, nMonPetit, nMoneda,nViaProce, nEstadoProceso ) "
    lsSQL = lsSQL & " VALUES('" & psCodCta & "','" & psNumExp & "'," & Format(pnMonPetit, "#0.00") & "," & pnMoneda & ","
    lsSQL = lsSQL & pnViaProce & "," & pnEstadoProceso & ") "
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

ErrordInsertExpedJud:
    Err.Raise Err.Number, "Error En Proceso dInsertColocRecupExpediente", Err.Description
End Sub

Public Sub dInsertColocRecupExpedInf(ByVal psCodCta As String, ByVal pnTipoInf As Integer, _
    ByVal psInforme As String, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    On Error GoTo ErrordInsertColocRecupExpedInf
    
    lsSQL = "INSERT INTO ColocRecupExpedInf(cCtaCod, nTipoInf, psInforme) "
    lsSQL = lsSQL & " VALUES('" & psCodCta & "'," & pnTipoInf & ",'" & psInforme & "' )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

ErrordInsertColocRecupExpedInf:
    Err.Raise Err.Number, "Error En Proceso ErrordInsertColocRecupExpedInf", Err.Description

End Sub
Public Sub dInsertColocRecup(ByVal psCtaCod As String, ByVal psFecIngRecup As String, ByVal pnCodComi As Integer, ByVal pnSaldoIntComp As Double, _
    ByVal pnSaldoIntMor As Double, ByVal pnSaldoGasto As Double, ByVal pnIntCompGen As Double, _
    ByVal psMetLiquid As String, ByVal pnTipCj As Integer, ByVal pnDemanda As Integer, _
    ByVal pnNroCalen As Integer, Optional ByVal pbEjecBatch As Boolean = False)
    
Dim lsSQL As String
    On Error GoTo ErrordInsertColocCredRecup
    lsSQL = "INSERT INTO ColocRecup(cCtaCod,dIngRecup,nComisionCod,nSaldoIntComp,nSaldoIntMor,nSaldoGasto, "
    lsSQL = lsSQL & "nIntCompGen,cMetLiquid, nTipCj,nDemanda,nNroCalen )  "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "','" & Format(psFecIngRecup, "mm/dd/yyyy") & "'," & pnCodComi & "," & Format(pnSaldoIntComp, "#0.00") & ","
    lsSQL = lsSQL & Format(pnSaldoIntMor, "#0.00") & "," & Format(pnSaldoGasto, "#0.00") & ","
    lsSQL = lsSQL & Format(pnIntCompGen, "#0.00") & ",'" & psMetLiquid & "'," & pnTipCj & ","
    lsSQL = lsSQL & pnDemanda & "," & pnNroCalen & " ) "
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

ErrordInsertColocCredRecup:
    Err.Raise Err.Number, "Error En Proceso dInsertColocCredRecup", Err.Description

End Sub

Public Sub dInsertColocRecupGastos(ByVal psCtaCod As String, ByVal pnNroGastoCta As Integer, ByVal pnPrdConceptoCod As Integer, _
    ByVal psFecAsigna As String, ByVal pnMonto As Double, ByVal pnMontoPagado As Double, _
    ByVal pnColocRecGastoEstado As Integer, psMotivoGasto As String, _
    Optional ByVal pbEjecBatch As Boolean = False)
    
Dim lsSQL As String
    On Error GoTo ErrordInsertColocRecupGastos
    lsSQL = "INSERT INTO ColocRecupGastos(cCtaCod,nNroGastoCta, nPrdConceptoCod, dAsigna,nMonto," _
            & " nMontoPagado, nColocRecGastoEstado, cMotivoGasto ) " _
            & " VALUES('" & psCtaCod & "'," & pnNroGastoCta & "," & pnPrdConceptoCod & ",'" _
            & Format(psFecAsigna, "mm/dd/yyyy") & "'," & Format(pnMonto, "#0.00") & "," _
            & Format(pnMontoPagado, "#0.00") & "," & pnColocRecGastoEstado & ",'" & psMotivoGasto & "' )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

ErrordInsertColocRecupGastos:
    Err.Raise Err.Number, "Error En Proceso dInsertColocCredRecup", Err.Description

End Sub

Public Sub dInsertColocacCred(ByVal psCtaCod As String, ByVal pnDiasAtraso As Integer, ByVal pnColocCondicion As ColocCredCondicion, _
    ByVal pnColocDestino As ColocDestino, ByVal psProtesto As String, ByVal pbCargoAuto As Integer, _
    ByVal psMetLiquidacion As String, ByVal pbRefCapInt As Integer, ByVal pnNroProxCuota As Integer, _
    ByVal pnIntPend As Double, ByVal psExoPenalidad As String, ByVal pnTpoCalend As ColocTipoCalend, _
    ByVal pnCalendDinamico As Integer, ByVal pnTipoDesembolso As Integer, ByVal pnNroProxDesemb As Integer, _
    Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pbMiVivienda As Boolean = False)

Dim lsSQL As String

    On Error GoTo ErrordInsertColocacCred
    lsSQL = "INSERT INTO ColocacCred(cCtaCod,nDiasAtraso,nColocCondicion,nColocDestino,"
    lsSQL = lsSQL & "cProtesto,bCargoAuto,cMetLiquidacion,bRefCapInt,nNroProxCuota,nIntPend,"
    lsSQL = lsSQL & "nExoPenalidad,nColocCalendCod,nCalendDinamico,nTipoDesembolso, nNroProxDesemb,bMiVivienda)"
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "',"
    lsSQL = lsSQL & pnDiasAtraso & ","
    lsSQL = lsSQL & pnColocCondicion & ","
    lsSQL = lsSQL & pnColocDestino & ","
    lsSQL = lsSQL & "'" & psProtesto & "',"
    lsSQL = lsSQL & pbCargoAuto & ","
    lsSQL = lsSQL & "'" & psMetLiquidacion & "',"
    lsSQL = lsSQL & pbRefCapInt & ","
    lsSQL = lsSQL & pnNroProxCuota & ","
    lsSQL = lsSQL & Format(pnIntPend, "#0.00")
    lsSQL = lsSQL & "," & psExoPenalidad & ","
    lsSQL = lsSQL & pnTpoCalend & ","
    lsSQL = lsSQL & pnCalendDinamico & ","
    lsSQL = lsSQL & pnTipoDesembolso & ","
    lsSQL = lsSQL & pnNroProxDesemb
    If pbMiVivienda Then
        lsSQL = lsSQL & ",1)"
    Else
        lsSQL = lsSQL & ",0)"
    End If
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
    Exit Sub

ErrordInsertColocacCred:
                  Err.Raise Err.Number, "Error En Proceso dInsertColocacCred", Err.Description

End Sub



Public Sub dInsertColocaciones(ByVal psCtaCod As String, ByVal pnPlazo As Integer, ByVal pdVenc As Date, ByVal pnMontoCol As Double, ByVal pnColocCalendCod As Integer, psLineaCred As String, ByVal pdVigencia As Date, _
    ByVal psMovAct As String, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    
    
    lsSQL = "INSERT Colocaciones (cCtaCod,nPlazo,dVenc,nMontoCol,nColocCalendCod,cUltimaActualizacion,cLineaCred,dVigencia) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "',"
    lsSQL = lsSQL & Trim(Str(pnPlazo)) & ","
    lsSQL = lsSQL & "'" & Format(pdVenc, "mm/dd/yyyy") & "',"
    lsSQL = lsSQL & Trim(Str(pnMontoCol)) & ","
    lsSQL = lsSQL & "'" & Trim(Str(pnColocCalendCod)) & "',"
    lsSQL = lsSQL & "'" & psMovAct & "',"
    lsSQL = lsSQL & "'" & psLineaCred & "',"
    lsSQL = lsSQL & "'" & Format(pdVigencia, "mm/dd/yyyy") & "')"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
End Sub

Public Sub dInsertProducto(ByVal psCtaCod As String, ByVal pnTasaInteres As Double, ByVal pnSaldo As Double, ByVal pnPrdEstado As Integer, ByVal pdPrdEstado As Date, ByVal pnTransacc As Long, _
    Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    On Error GoTo ErrordInsertProducto
    
    lsSQL = "INSERT Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "',"
    lsSQL = lsSQL & Format(pnTasaInteres, "#0.00") & ","
    lsSQL = lsSQL & Format(pnSaldo, "#0.00") & ","
    lsSQL = lsSQL & Trim(Str(pnPrdEstado)) & ","
    lsSQL = lsSQL & Format(pdPrdEstado, "mm/dd/yyyy") & ","
    lsSQL = lsSQL & Trim(Str(pnTransacc)) & ")"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

ErrordInsertProducto:
    Err.Raise Err.Number, "Error En Proceso dInsertProducto", Err.Description
        
    
End Sub

Public Sub dInsertColocCredCredVig(ByVal psCtaCod As String, ByVal psCtaCodRef As String, Optional ByVal pbEjecBatch As Boolean = False)
Dim lsSQL As String
        
    lsSQL = "INSERT INTO ColocCredCredVig(cCtaCod, cCtaCodRef)"
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "','" & psCtaCodRef & "')"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
        
    
End Sub


'Colocaciones
Public Sub dUpdateMov(ByVal pnMovNro As Long, Optional ByVal psOpeCod As String = "@", _
    Optional ByVal psMovDesc As String = "@", Optional ByVal pnMovEstado As MovEstado = -1, _
    Optional ByVal pnMovFlag As MovFlag = -1, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE Mov SET "

    If psOpeCod <> "@" Then
        lsSQL = lsSQL & " cOpeCod = '" & psOpeCod & "',"
    End If
    If psMovDesc <> "@" Then
        lsSQL = lsSQL & " cMovDesc = '" & psMovDesc & "',"
    End If
    If pnMovEstado <> -1 Then
        lsSQL = lsSQL & " nMovEstado = " & pnMovEstado & ","
    End If
    If pnMovFlag <> -1 Then
        lsSQL = lsSQL & " nMovFlag = " & pnMovFlag & ","
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE nMovNro =" & pnMovNro & " "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub


Public Sub dUpdateProducto(ByVal psCtaCod As String, Optional ByVal pnTasaInteres As Double = -1, Optional ByVal pnSaldo As Double = -1, Optional ByVal pnPrdEstado As Integer = -1, Optional ByVal pdPrdEstado As Date = CDate("01/01/1950"), Optional ByVal pnTransacc As Long = -1, _
    Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pnExtorno As Integer = 0, Optional ByVal bIncrementarSaldo As Boolean = False)

Dim lsSQL As String
    
    On Error GoTo ErrordInsertProducto
    
    lsSQL = " UPDATE Producto Set "
    If pnTasaInteres <> -1 Then
        lsSQL = lsSQL & " nTasaInteres = " & Format(pnTasaInteres, "#0.0000") & ","
    End If
    If pnExtorno = 0 Then
        If pnSaldo <> -1 And Not bIncrementarSaldo Then
            lsSQL = lsSQL & " nSaldo = " & pnSaldo & ","
        End If
    Else
        If pnSaldo <> -1 And Not bIncrementarSaldo Then
            lsSQL = lsSQL & " nSaldo = nSaldo - " & pnSaldo & ","
        End If
    End If
    
    If bIncrementarSaldo Then
        If pnSaldo <> -1 Then
            lsSQL = lsSQL & " nSaldo = nSaldo + " & pnSaldo & ","
        End If
    End If
    If pnPrdEstado <> -1 Then
        lsSQL = lsSQL & " nPrdEstado = " & pnPrdEstado & ","
    End If
    If pdPrdEstado <> CDate("01/01/1950") Then
        lsSQL = lsSQL & " dPrdEstado = '" & Format(pdPrdEstado, "mm/dd/yyyy") & "',"
    End If
    If pnExtorno = 0 Then
        If pnTransacc = -2 Then ' Si se le envia (-2) aumenta el nro de transaccion en uno (LAYG)
            lsSQL = lsSQL & " nTransacc = nTransacc + 1,"
        ElseIf pnTransacc <> -1 Then
            lsSQL = lsSQL & " nTransacc = " & pnTransacc & ","
        End If
    Else
        If pnTransacc = -2 Then ' Si se le envia (-2) aumenta el nro de transaccion en uno (LAYG)
            lsSQL = lsSQL & " nTransacc = nTransacc + 1,"
        ElseIf pnTransacc <> -1 Then
            lsSQL = lsSQL & " nTransacc = nTransacc - 1, "
        End If
    End If
    lsSQL = Mid(lsSQL, 1, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "'"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
    Exit Sub

ErrordInsertProducto:
                  Err.Raise Err.Number, "Error En Proceso dInsertProducto", Err.Description

End Sub

Public Sub dInsertProductoPersona(ByVal psCuenta As String, ByVal psPersona As String, ByVal pnRelacion As ColocRelacPers, _
    Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
        & "VALUES ('" & psCuenta & "','" & psPersona & "'," & pnRelacion & ")"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
End Sub

Public Sub dInsertaMovDoc(ByVal pnMovNro As Long, ByVal pnDocTpo As Integer, _
    ByVal psDocNro As String, ByVal pdDocFecha As Date, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
        
    On Error GoTo ErrordInsertaMovDoc
    lsSQL = "INSERT INTO MovDoc(nMovNro,nDocTpo,cDocNro,dDocFecha) "
    lsSQL = lsSQL & " VALUES(" & pnMovNro & "," & pnDocTpo & ",'" & psDocNro & "','" & Format(pdDocFecha, "mm/dd/yyyy hh:mm:ss") & "')"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

ErrordInsertaMovDoc:
    Err.Raise Err.Number, "Error En Proceso dInsertaMovDoc", Err.Description
    
End Sub

'Inserta MovGarantDet
Public Sub dInsertMovGarantDet(ByVal pnMovNro As Long, ByVal psNumgarant As String, _
    ByVal pscOpeCod As String, ByVal psPrdConcepto As String, ByVal pdmonto As Double)
    
Dim lsSQL As String
On Error GoTo InsertMovErr

lsSQL = "Insert Into MovGarantDet(nMovNro,cNumgarant, cOpeCod, nPrdConceptoCod,nMonto)"
lsSQL = lsSQL & " Values(" & pnMovNro & ",'" & psNumgarant & "','" & pscOpeCod & "'," & psPrdConcepto & "," & pdmonto & ")"
coConex.Ejecutar lsSQL
Exit Sub
InsertMovErr:
   Call RaiseError(MyUnhandledError, "DCredActualizaBD Insert Mov")

End Sub

'Inserta MovGarantia
Public Sub dInsertMovGarantia(ByVal pnMovNro As Long, ByVal psNumgarant As String, _
    ByVal pscOpeCod As String, ByVal pdmonto As Double, ByVal psDescripcion As String)
    
Dim lsSQL As String
On Error GoTo InsertMovG_Err
'lssql = " Select Max(nMovNro) from MovGarantia"
lsSQL = " Insert MovGarantia(nMovNro,cNumgarant, cOpeCod, nMonto, cDescripcion)"
lsSQL = lsSQL & " Values(" & pnMovNro & ",'" & psNumgarant & "','" & pscOpeCod & "' ," & pdmonto & ",'" & psDescripcion & "')"
coConex.Ejecutar lsSQL
Exit Sub
InsertMovG_Err:
    Call RaiseError(MyUnhandledError, "DCredActualizaBD Insert Mov")

End Sub

'Inserta Mov
Public Sub dInsertMov(ByVal psMovNro As String, ByVal psOpeCod As String, _
    ByVal psMovDesc As String, ByVal pnMovEstado As MovEstado, _
    Optional pnMovFlag As MovFlag = gMovFlagVigente, _
    Optional pbEjecBatch As Boolean = False)
    
Dim lsSQL As String
On Error GoTo InsertMovErr
    
    lsSQL = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado, nMovFlag) " _
          & "VALUES ('" & psMovNro & "','" & psOpeCod & "','" & Replace(psMovDesc, "'", "''") & "'," _
          & pnMovEstado & "," & pnMovFlag & ")"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

InsertMovErr:
    Call RaiseError(MyUnhandledError, "DCredActualizaBD Insert Mov")
End Sub

'Inserta MovRef
Public Sub dInsertMovRef(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long, _
    Optional pbEjecBatch As Boolean = False)
    
Dim lsSQL As String
On Error GoTo InsertMovRefErr
    
    lsSQL = "INSERT MovRef (nMovNro,nMovNroRef) " _
          & "VALUES (" & pnMovNro & "," & pnMovNroRef & ")"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

InsertMovRefErr:
    Call RaiseError(MyUnhandledError, "DCredActualizaBD Insert Mov Ref")
End Sub


Public Sub dInsertMovCol(ByVal pnMovNro As Long, ByVal psOperacion As String, _
        ByVal psCuenta As String, ByVal pnNroCalend As Long, ByVal pnMonto As Currency, _
        ByVal pnDiasMora As Integer, ByVal psMetLiq As String, ByVal pnPlazo As Integer, ByVal pnSaldoCap As Double, ByVal pnEstado As Integer, _
        Optional pbEjecBatch As Boolean = False, Optional ByVal pnFlag As Long = -1, Optional ByVal pnPrepago As Integer = 0)

Dim lsSQL As String

    lsSQL = "INSERT MovCol (nMovNro,cOpeCod,cCtaCod,nNroCalen,nMonto,nDiasMora,cMetLiq,nPlazo, nSaldoCap, nCredEstado,nPrepago" & IIf(pnFlag > 0, ",nFlag)", ")") _
        & "VALUES (" & pnMovNro & ",'" & psOperacion & "','" & psCuenta & "'," & pnNroCalend & "," _
        & pnMonto & "," & pnDiasMora & ",'" & psMetLiq & "'," & pnPlazo & "," & Format(pnSaldoCap, "#0.00") & "," & pnEstado & "," & pnPrepago & IIf(pnFlag > 0, "," & pnFlag & ")", ")")

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

''''''''''''''
'
Public Sub dInsertaMovLavDinero(ByVal nMovNro As Long, ByVal sPersCod As String)
Dim lsSQL As String

On Error GoTo InsertMovLav
lsSQL = "INSERT MovLavDinero (nMovNro,cPersCod) " _
    & "VALUES (" & nMovNro & ",'" & sPersCod & "')"
coConex.Ejecutar lsSQL
Exit Sub

InsertMovLav:
    Call RaiseError(MyUnhandledError, "DCredActualizaBD Insert Mov Lab")

End Sub
'
''''''''''''''


'Obtiene el nMovNro  apartir del cMovNro
Public Function dGetnMovNro(ByVal psMovNro As String) As Long
Dim lsSQL As String
Dim lrs As ADODB.Recordset
Set lrs = New ADODB.Recordset

lsSQL = "Select nMovNro From Mov where cMovNro ='" & psMovNro & "'"
Set lrs = coConex.CargaRecordSet(lsSQL, adLockReadOnly)
'Set lrs.ActiveConnection = coConex.ConexionActiva
If Not lrs.EOF And Not lrs.BOF Then
    dGetnMovNro = lrs!nMovNro
End If
lrs.Close
Set lrs = Nothing
End Function

Public Sub dInsertMovColDet(ByVal pnMovNro As Long, ByVal psOperacion As String, _
        ByVal psCtaCod As String, pnNroCalend As Long, ByVal pnConcepto As Long, _
        ByVal pnNroCuota As Integer, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT MovColDet (nMovNro,cOpeCod,cCtaCod,nNroCalen,nPrdConceptoCod,nNroCuota,nMonto) " _
        & "VALUES (" & pnMovNro & ",'" & psOperacion & "','" & psCtaCod & "'," & pnNroCalend & "," _
        & pnConcepto & "," & pnNroCuota & "," & pnMonto & " ) "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dInsertColocCargoAutoma(ByVal psCtaCod As String, ByVal pnOrden As Integer, ByVal psCodCtaAho As String, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String

    lsSQL = "INSERT INTO ColocCargoAutoma(cCtaCod,nOrden,cCodCtaAho,cEstado)"
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "'," & pnOrden & ",'" & psCodCtaAho & "',1)"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub
Public Function dFechaHora(Optional psFecha As String = "") As Date
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConecta As DConecta

    On Error GoTo ErrordFechaHora
    sSQL = "Select GETDATE() as FechaHora"
    Set R = New ADODB.Recordset
    R.Open sSQL, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    If psFecha = "" Then
        dFechaHora = R!FechaHora
    Else
        dFechaHora = CDate(Format(Format(psFecha, "dd/mm/yyyy") & " " & Format(R!FechaHora, "hh:mm:ss"), "dd/mm/yyyy"))
    End If
    R.Close
    Set R = Nothing
    Exit Function

ErrordFechaHora:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
Public Sub dInsertColocacEstado(ByVal psCtaCod As String, ByVal pdFechaPrdEstado As String, _
        ByVal pnPrdEstado As Integer, ByVal pnCuotas As Integer, ByVal pnMonto As Currency, _
        ByVal psDescripcion As String, ByVal pnTipoCalendario As ColocTipoCalend, _
        ByVal pnPeriodoFechaFija As Integer, ByVal pnPeriodoGracia As Integer, _
        ByVal pnPlazo As Integer, ByVal pnTipoGracia As Integer, ByVal pnTipoDesembolso As ColocTiposDesembolso, _
        ByVal pnProxMes As Integer, ByVal pnCalendDinamico As Integer, Optional pbEjecBatch As Boolean = False, _
        Optional ByVal pnMotivRech As Integer = -1, Optional ByVal psTipoGasto As String = "F")

Dim lsSQL As String
    
    If pnMotivRech <> -1 Then
        lsSQL = "INSERT ColocacEstado (cCtaCod, dPrdEstado, nPrdEstado, nCuotas, nMonto, cDescripcion, " _
            & "nColocCalendCod, nPeriodoFechaFija, nPeriodoGracia, nPlazo, nTipoGracia,nTipoDesembolso,nProxMes,nCalendDinamico,nMotivoRechazo, cTipoGasto) " _
            & "VALUES ('" & psCtaCod & "','" & Format(CDate(Format(pdFechaPrdEstado, "dd/mm/yyyy") & Space(1) & Format(dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "'," & pnPrdEstado & "," & pnCuotas & "," _
            & pnMonto & ",'" & psDescripcion & "'," & pnTipoCalendario & "," & pnPeriodoFechaFija & "," _
            & pnPeriodoGracia & "," & pnPlazo & "," & pnTipoGracia & "," & pnTipoDesembolso & "," & pnProxMes & "," & pnCalendDinamico & "," & pnMotivRech & ",'" & psTipoGasto & "')"
    Else
        lsSQL = "INSERT ColocacEstado (cCtaCod, dPrdEstado, nPrdEstado, nCuotas, nMonto, cDescripcion, " _
            & "nColocCalendCod, nPeriodoFechaFija, nPeriodoGracia, nPlazo, nTipoGracia,nTipoDesembolso,nProxMes,nCalendDinamico,cTipoGasto) " _
            & "VALUES ('" & psCtaCod & "','" & Format(CDate(Format(pdFechaPrdEstado, "dd/mm/yyyy") & Space(1) & Format(dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "'," & pnPrdEstado & "," & pnCuotas & "," _
            & pnMonto & ",'" & psDescripcion & "'," & pnTipoCalendario & "," & pnPeriodoFechaFija & "," _
            & pnPeriodoGracia & "," & pnPlazo & "," & pnTipoGracia & "," & pnTipoDesembolso & "," & pnProxMes & "," & pnCalendDinamico & ",'" & psTipoGasto & "')"
    End If
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dUpdateColocacEstado(ByVal psCtaCod As String, ByVal pdFechaPrdEstado As String, _
        ByVal pnPrdEstado As Integer, ByVal pnCuotas As Integer, ByVal pnMonto As Currency, _
        ByVal psDescripcion As String, ByVal pnTipoCalendario As ColocTipoCalend, _
        ByVal pnPeriodoFechaFija As Integer, ByVal pnPeriodoGracia As Integer, _
        ByVal pnPlazo As Integer, ByVal pnTipoGracia As Integer, ByVal pnTipoDesembolso As ColocTiposDesembolso, _
        ByVal pnProxMes As Integer, ByVal pnCalendDinamico As Integer, Optional pbEjecBatch As Boolean = False, _
        Optional ByVal psTipoGastos As String = "F")

Dim lsSQL As String

    
    On Error GoTo ErrordUpdateColocacEstado
    
    lsSQL = " UPDATE ColocacEstado SET "
    lsSQL = lsSQL & " dPrdEstado = '" & Format(CDate(Format(pdFechaPrdEstado, "dd/mm/yyyy") & Space(1) & Format(dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
    lsSQL = lsSQL & " nCuotas = " & pnCuotas & ","
    lsSQL = lsSQL & " NMonto = " & Format(pnMonto, "#0.00") & ","
    lsSQL = lsSQL & " cDescripcion = '" & Replace(psDescripcion, "'", "''") & "',"
    lsSQL = lsSQL & " nColocCalendCod = " & pnTipoCalendario & ","
    lsSQL = lsSQL & " nPeriodoFechaFija  = " & pnPeriodoFechaFija & ","
    lsSQL = lsSQL & " nPeriodoGracia = " & pnPeriodoGracia & ","
    lsSQL = lsSQL & " nPlazo = " & pnPlazo & ","
    lsSQL = lsSQL & " nTipoGracia = " & pnTipoGracia & ","
    lsSQL = lsSQL & " nTipoDesembolso = " & pnTipoDesembolso & ","
    lsSQL = lsSQL & " nProxMes = " & pnProxMes & ","
    lsSQL = lsSQL & " nCalendDinamico = " & pnCalendDinamico & ", "
    lsSQL = lsSQL & " cTipoGasto = '" & psTipoGastos & "' "
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "' AND nPrdEstado = " & pnPrdEstado
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

    Exit Sub

ErrordUpdateColocacEstado:
    Err.Raise Err.Number, "Error En Proceso dUpdateColocacEstado", Err.Description

End Sub

Public Sub dUpdateColocCalendario(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
        ByVal pnCuota As Integer, ByVal pnColocCalendApl As ColocCalendApl, Optional ByVal pdVenc As Date = CDate("01/01/1950"), _
        Optional ByVal pnColocCalendEstado As Integer = -1, Optional ByVal psDescripcion As String = "", Optional ByVal pnCalendProc As ColocCalendConceptoProc = -1, _
        Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pcColocMiVivEval As String = "N", _
        Optional ByVal pdFecPago As Date = CDate("01/01/1900"))
        
Dim lsSQL As String

    On Error GoTo ErrordInsertColocCalendario
    'cCtaCod            nNroCalen   cColocCalendApl nCuota      dVenc                       nColocCalendEstado cDescripcion                                                                                                                                                                                                                                                    cColocCalenFlag nCalendProc
    
    lsSQL = "UPDATE ColocCalendario SET "
    If pdVenc <> CDate("01/01/1950") Then
        lsSQL = lsSQL & " dVenc = '" & Format(pdVenc, "mm/dd/yyyy") & "',"
    End If
    If pnColocCalendEstado <> -1 Then
        lsSQL = lsSQL & " nColocCalendEstado = " & pnColocCalendEstado & ","
    End If
    If psDescripcion <> "" Then
        lsSQL = lsSQL & " cDescripcion = '" & psDescripcion & "',"
    End If
    If pnCalendProc <> -1 Then
        lsSQL = lsSQL & " nCalendProc = " & pnCalendProc & ","
    End If
    If pcColocMiVivEval <> "N" Then
        lsSQL = lsSQL & " cColocMiVivEval = '" & pcColocMiVivEval & "',"
    End If
    If pdFecPago <> CDate("01/01/1900") Then
        lsSQL = lsSQL & " dPago = '" & Format(pdFecPago, "mm/dd/yyyy hh:mm:ss") & "',"
    End If
    
    lsSQL = Mid(lsSQL, 1, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & pnNroCalen & " AND nColocCalendApl = " & pnColocCalendApl & " AND nCuota = " & pnCuota
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
    Exit Sub

ErrordInsertColocCalendario:
    Err.Raise Err.Number, "Error En Proceso dInsertColocCalendario", Err.Description

End Sub

Public Sub dDeleteColocCalendario(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
        ByVal pnColocCalendApl As ColocCalendApl, Optional ByVal pnCuota As Integer = -1, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    On Error GoTo ErrordDeleteColocCalendario
    
    lsSQL = " DELETE ColocCalendario WHERE "
    lsSQL = lsSQL & " cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & pnNroCalen & " AND nColocCalendApl = " & pnColocCalendApl
    If pnCuota <> -1 Then
        lsSQL = lsSQL & " AND nCuota = " & pnCuota
    End If
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

ErrordDeleteColocCalendario:
    Err.Raise Err.Number, "Error En Proceso dDeleteColocCalendario", Err.Description

End Sub

Public Sub dInsertColocCalendario(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
        ByVal pnColocCalendApl As ColocCalendApl, ByVal pnCuota As Integer, ByVal pdVenc As Date, _
        ByVal pnColocCalendEstado As Integer, ByVal psDescripcion As String, ByVal pnCalendProc As ColocCalendConceptoProc, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ColocCalendario (cCtaCod,nNroCalen,nColocCalendApl,nCuota,dVenc,nColocCalendEstado,cDescripcion,cColocCalenFlag,nCalendProc) " _
        & "VALUES ('" & psCtaCod & "'," & pnNroCalen & "," & pnColocCalendApl & "," & pnCuota & ",'" _
        & Format(pdVenc, "mm/dd/yyyy") & "'," & pnColocCalendEstado & ",'" & psDescripcion & "',NULL," & pnCalendProc & ")"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If


End Sub

Public Sub dInsertMovCMAC(ByVal pnMovNro As Long, ByVal psPersCod As String, _
        ByVal psIFTipo As String, ByVal pnMoneda As Moneda, ByVal psCtaIf As String, _
        ByVal psDocumento As String, ByVal psOpeCod As String, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT MovCMAC (nMovNro, cPersCod, cIFTpo, nMoneda, cCtaIF, cDocumento, cOpeCod, nMonto  ) " _
        & "VALUES (" & pnMovNro & ",'" & psPersCod & "','" & psIFTipo & "'," & pnMoneda & ",'" _
        & psCtaIf & "','" & psDocumento & "','" & psOpeCod & "'," & pnMonto & ") "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
Public Sub dActualizarMovCmac(ByVal pnMovNro As Integer, ByVal pnMonto As Double, Optional pbEjecBatch As Boolean = False)
    Dim sSQL As String
    
    sSQL = "Update MovCMAC"
    sSQL = sSQL & " Set nMonto=nMonto+" & pnMonto
    sSQL = sSQL & " Where nMovNro=" & pnMovNro
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSQL
    Else
        coConex.Ejecutar sSQL
    End If
End Sub


Public Sub dInsertColocCalendDet(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
        ByVal pnColocCalendApl As ColocCalendApl, ByVal pnCuota As Integer, ByVal pnColocConceptoCod As Long, _
        ByVal pnMonto As Currency, ByVal pnMontoPagado As Currency, _
        ByVal psFlag As String, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ColocCalendDet (cCtaCod,nNroCalen,nColocCalendApl,nCuota,nPrdConceptoCod,nMonto,nMontoPagado,cFlag) " _
        & "VALUES ('" & psCtaCod & "'," & pnNroCalen & "," & pnColocCalendApl & "," & pnCuota & "," _
        & pnColocConceptoCod & "," & Format(pnMonto, "#0.00") & "," & Format(pnMontoPagado, "#0.00") & ",'" & psFlag & "') "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dUpdateColocCalendDet(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
        ByVal pnColocCalendApl As String, ByVal pnCuota As Integer, ByVal pnColocConceptoCod As Long, _
        Optional ByVal pnMonto As Double = -1, Optional ByVal pnMontoPagado As Double = -1, _
        Optional ByVal psFlag As String = "", Optional pbEjecBatch As Boolean = False, Optional pbAmortizar As Boolean = False _
        , Optional pbMontoIncrem As Boolean = False)

Dim lsSQL As String

    On Error GoTo ErrordUpdateColocCalendDet
    lsSQL = "UPDATE ColocCalendDet SET "
    If Not pbMontoIncrem Then
        If pnMonto <> -1 Then
            lsSQL = lsSQL & " nMonto = " & pnMonto & ","
        End If
    Else
        If pnMonto <> -1 Then
            lsSQL = lsSQL & " nMonto = nMonto + " & pnMonto & ","
        End If
    End If
    If pnMontoPagado <> -1 Then
        If pbAmortizar Then
            lsSQL = lsSQL & " nMontoPagado = nMontoPagado + " & pnMontoPagado & ","
        Else
            lsSQL = lsSQL & " nMontoPagado = " & pnMontoPagado & ","
        End If
    End If
    If psFlag <> "" Then
        lsSQL = lsSQL & " cFlag = '" & psFlag & "',"
    End If
    lsSQL = Mid(lsSQL, 1, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & pnNroCalen & " AND nColocCalendApl = " & pnColocCalendApl & " AND nCuota = " & pnCuota & " AND nPrdConceptoCod = " & pnColocConceptoCod
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

ErrordUpdateColocCalendDet:
    Err.Raise Err.Number, "Error En Proceso dUpdateColocCalendDet", Err.Description

End Sub

Public Sub dDeleteColocCalendDet(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
        ByVal pnColocCalendApl As ColocCalendApl, Optional ByVal pnCuota As Integer = -1, Optional ByVal pnColocConceptoCod As Long = -1, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    On Error GoTo ErrordDeleteColocCalendDet
    lsSQL = "DELETE ColocCalendDet WHERE "
    lsSQL = lsSQL & " cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & pnNroCalen & " AND nColocCalendApl = " & pnColocCalendApl
    If pnCuota <> -1 Then
        lsSQL = lsSQL & " AND nCuota = " & pnCuota
    End If
    If pnColocConceptoCod <> -1 Then
        lsSQL = lsSQL & " AND nPrdConceptoCod = " & pnColocConceptoCod
    End If
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

ErrordDeleteColocCalendDet:
    Err.Raise Err.Number, "Error En Proceso dDeleteColocCalendDet", Err.Description
    
End Sub

Public Sub dUpdatecorresponsalia(ByVal psCodCta As String, Optional ByVal pbEjecBatch As Boolean = False)
Dim lsSQL As String
    On Error GoTo ErrorddUpdatecorresponsalia
    
    lsSQL = "update corresponsalia set nflag=1 where cctacod= '" & psCodCta & "'"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

ErrorddUpdatecorresponsalia:
    Err.Raise Err.Number, "Error En Proceso dUpdatecorresponsalia", Err.Description
End Sub

Public Sub dUpdateDescripcionLote(ByVal psCtaCod As String, ByVal psDescLote As String, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "UPDATE ColocPignoraticio SET cLote ='" & psDescLote & "' " _
        & "WHERE cCtaCod ='" & psCtaCod & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dUpdateDocRec(ByVal psNumDoc As String, ByVal psPersCod As String, Optional pbEjecBatch As Boolean = False, Optional pnEstCheque As ChequeEstado = gChqEstValorizado)
Dim sSQL As String
    
    sSQL = "Update DocRec Set nEstado = " & pnEstCheque & " Where nTpoDoc = 47 AND cNroDoc = '" & psNumDoc & "' "
    sSQL = sSQL & " AND cPersCod = '" & psPersCod & "' "
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSQL
    Else
        coConex.Ejecutar sSQL
    End If
End Sub

Public Sub dUpdateProducto_Estado(ByVal psCtaCod As String, ByVal pnEstado As ColocEstado, _
        ByVal psFechaHora As String, Optional ByVal pnTransac As Integer = -1, Optional pbEjecBatch As Boolean = False, Optional ByVal pbIncremNumTran As Boolean = False)

Dim lsSQL As String

    If pbIncremNumTran Then
        lsSQL = "UPDATE Producto SET nPrdEstado = " & pnEstado & ", dPrdEstado = '" & Format(psFechaHora, "mm/dd/yyyy") & "', " _
            & "nTransacc = nTransacc + " & pnTransac & " WHERE  cCtaCod ='" & psCtaCod & "' "
    Else
        '**Update Producto
        lsSQL = "UPDATE Producto SET nPrdEstado = " & pnEstado & ", dPrdEstado = '" & Format(psFechaHora, "mm/dd/yyyy") & "', " _
            & "nTransacc = " & pnTransac & " WHERE  cCtaCod ='" & psCtaCod & "' "
    End If
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dUpdateColocaciones(ByVal psCtaCod As String, Optional ByVal pnPlazo As Integer = -1, _
    Optional ByVal pdVenc As Date = CDate("01/01/1950"), Optional ByVal pnMontoCol As Double = -1, Optional ByVal pnColocCalendCod As Integer = -1, _
    Optional ByVal psLineaCred As String = "", Optional ByVal pdVigencia As Date = CDate("01/01/1950"), Optional ByVal pbEjecBatch As Boolean = False, _
    Optional ByVal bIncrementarMontoCol As Boolean = False, Optional ByVal psUltimaActualizacion As String = "")

Dim lsSQL As String

    On Error GoTo ErrordUpdateColocaciones
    lsSQL = "UPDATE Colocaciones SET "
    If pnPlazo <> -1 Then
        lsSQL = lsSQL & " nPlazo = " & pnPlazo & ","
    End If
    If pdVenc <> CDate("01/01/1950") Then
        lsSQL = lsSQL & " dVenc = '" & Format(pdVenc, "mm/dd/yyyy") & "',"
    End If
    If pnMontoCol <> -1 Then
        If Not bIncrementarMontoCol Then
            lsSQL = lsSQL & " nMontoCol = " & Format(pnMontoCol, "#0.00") & ","
        Else
            lsSQL = lsSQL & " nMontoCol = nMontoCol + " & Format(pnMontoCol, "#0.00") & ","
        End If
    End If
    If pnColocCalendCod <> -1 Then
        lsSQL = lsSQL & " nColocCalendCod = " & pnColocCalendCod & ","
    End If
    If psLineaCred <> "" Then
        lsSQL = lsSQL & " cLineaCred = '" & psLineaCred & "', "
    End If
    If psUltimaActualizacion <> "" Then
        lsSQL = lsSQL & " cUltimaActualizacion = '" & psUltimaActualizacion & "',"
    End If
    If pdVigencia <> CDate("01/01/1950") Then
        lsSQL = lsSQL & " dVigencia = '" & Format(pdVigencia, "mm/dd/yyyy") & "' "
    End If
    
    lsSQL = Mid(lsSQL, 1, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
    Exit Sub

ErrordUpdateColocaciones:
    Err.Raise Err.Number, "Error En Proceso dUpdateColocaciones", Err.Description
End Sub

Public Sub dUpdateColocaciones_UltimaActualizacion(ByVal psCtaCod As String, ByVal psFechaHora As String, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    '**Update Producto
    lsSQL = "UPDATE Colocaciones SET cUltimaActualizacion = '" & psFechaHora & "' " _
        & "WHERE  cCtaCod ='" & psCtaCod & "' "
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dInsertColocacRefinanc(ByVal psCtaCod As String, ByVal psCtaCodRef As String, _
    ByVal pnMontoRef As Double, ByVal pnEstado As ColocEstadoRefinanc, ByVal pdFecha As Date, Optional pbEjecBatch As Boolean = False, Optional pbSustiDeudor As Boolean = False)
Dim lsSQL As String
    lsSQL = "INSERT INTO ColocacRefinanc(cCtaCod,cCtaCodRef,nEstado,dEstado,nMontoRef, bSustiDeudor)"
    lsSQL = lsSQL & "VALUES('" & psCtaCod & "','" & psCtaCodRef & "'," & pnEstado & ",'" & Format(pdFecha, "mm/dd/yyyy") & "'," & Format(pnMontoRef, "#0.00") & ", " & IIf(pbSustiDeudor, 1, 0) & ")"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dUpdateColocacRefinanc(ByVal psCtaCod As String, ByVal psCtaCodRef As String, _
      ByVal pnEstado As ColocEstadoRefinanc, ByVal pdFecha As Date, ByVal pnMontoRef As Double, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
           
       lsSQL = "UPDATE ColocacRefinanc Set "
       lsSQL = lsSQL & " nMontoRef = " & Format(pnMontoRef, "#0.00")
       lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "' AND cCtaCodRef = '" & psCtaCodRef & "' AND nEstado = " & pnEstado & "' AND dEstado = '" & pdFecha & "'"
        If pbEjecBatch Then
            coConex.AdicionaCmdBatch lsSQL
        Else
            coConex.Ejecutar lsSQL
        End If
End Sub

Public Sub dInsertColocacRefinancDet(ByVal psCtaCod As String, ByVal psCtaCodRef As String, _
    ByVal pnMonto As Double, ByVal pnMontoPagado As Double, ByVal pnEstado As ColocEstadoRefinanc, _
    ByVal pnColocConceptoCod As ColocConcepto, ByVal pdFecha As Date, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    
    lsSQL = "INSERT INTO ColocacRefinancDet(cCtaCod,cCtaCodRef,nEstado,dEstado,nPrdConceptoCod,nMonto,nMontoPagado)"
    lsSQL = lsSQL & "VALUES('" & psCtaCod & "','" & psCtaCodRef & "'," & pnEstado & ",'" & Format(pdFecha, "mm/dd/yyyy") & "'," & pnColocConceptoCod & "," & Format(pnMonto, "#0.00") & "," & Format(pnMontoPagado, "#0.00") & ")"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dUpdateColocacRefinancDet(ByVal psCtaCod As String, ByVal psCtaCodRef As String, _
     ByVal pnColocConcepto As ColocConcepto, ByVal pnEstado As ColocEstadoRefinanc, ByVal pdFecha As Date, Optional ByVal pnMonto As Double = -1, _
     Optional ByVal pnMontoPagado As Double = -1, Optional pbAmortizac As Boolean = False, _
     Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    lsSQL = "UPDATE ColocacRefinancDet SET "
    If pnMonto <> -1 Then
        lsSQL = lsSQL & "nMonto = " & Format(pnMonto, "#0.00") & ","
    End If
    If pnMontoPagado <> -1 Then
        If pbAmortizac Then
            lsSQL = lsSQL & "nMontoPagado = nMontoPagado + " & Format(pnMontoPagado, "#0.00") & ","
        Else
            lsSQL = lsSQL & "nMontoPagado = " & Format(pnMontoPagado, "#0.00") & ","
        End If
    End If
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "' AND cCtaCodRef = '" & psCtaCodRef & "' AND nEstado = " & pnEstado & "' AND dEstado = '" & pdFecha & "' AND nPrdConceptoCod = " & pnColocConcepto
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub
    
Public Sub dInsertColocCalificacionAnalista(ByVal psCtaCod As String, ByVal pdColocNotaFecha As Date, ByVal pnColocNota As Integer, ByVal psComentario As String, ByVal psTipo As String, _
        Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    lsSQL = "INSERT INTO ColocCalificacionAnalista(cCtaCod,dColocNotaFecha,nColocNota,cComentario,cTipo) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "','" & Format(pdColocNotaFecha, "mm/dd/yyyy hh:mm:ss") & "'," & pnColocNota & ",'" & psComentario & "','" & psTipo & "')"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
End Sub

Public Sub dInsertColocGarantRec(ByVal nNroGarantRec As Long, ByVal nTipo As Integer, ByVal psCtaCod As String, _
        ByVal dRegistro As Date, ByVal nValorTotal As Double, ByVal nEstado As Integer, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    lsSQL = "INSERT INTO ColocGarantRec(nNroGarantRec, nTipo, cCtaCod, dRegistro, nValorTotal, nEstado) "
    lsSQL = lsSQL & " VALUES(" & nNroGarantRec & "," & nTipo & ",'" & psCtaCod & "','" & Format(dRegistro, "mm/dd/yyyy") & "'," & Format(nValorTotal, "#0.00") & "," & nEstado & ")"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
End Sub

Public Sub dInsertColocFteIngreso(ByVal psNumFuente As String, ByVal psCtaCod As String, ByVal pdPersEval As Date, _
         Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    lsSQL = "INSERT INTO ColocFteIngreso(cNumFuente, cCtaCod, dPersEval) "
    lsSQL = lsSQL & " VALUES('" & psNumFuente & "','" & psCtaCod & "','" & Format(pdPersEval, "mm/dd/yyyy") & "')"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
End Sub


Public Function dRecuperaMovRegistroDacion(ByVal pnNroDacion As Long, ByVal psCtaCod As String) As Long
Dim sSQL As String
Dim R As ADODB.Recordset

    sSQL = "Select nMovNro From MovCol Where nNroCalen = " & pnNroDacion & " And cCtaCod = '" & psCtaCod & "'"
    Set R = New ADODB.Recordset
    R.Open sSQL, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    dRecuperaMovRegistroDacion = R!nMovNro
    R.Close
    Set R = Nothing
    
End Function

Public Sub dAnularColocGarantRec(ByVal nNroGarantRec As Long, ByVal nEstado As ColocGarantRecEstado, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "UPDATE ColocGarantRec SET nEstado = " & nEstado
    lsSQL = lsSQL & " Where nNroGarantRec = " & nNroGarantRec

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dUpdateColocGarantRec(ByVal nNroGarantRec As Long, ByVal psCtaCod As String, _
    ByVal nValorTotal As Double, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String

    lsSQL = "UPDATE ColocGarantRec SET cCtaCod = '" & psCtaCod & "', "
    lsSQL = lsSQL & " nValorTotal = " & Format(nValorTotal, "#0.00")
    lsSQL = lsSQL & " Where nNroGarantRec = " & nNroGarantRec
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Function dNuevoCodigoColocGarantRec() As Long
Dim lsSQL As String
Dim R As ADODB.Recordset

    lsSQL = " Select MAX(nNroGarantRec) as nTotal from ColocGarantRec with (tablockx) "
    Set R = New ADODB.Recordset
    R.Open lsSQL, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
        If IsNull(R!nTotal) Then
            dNuevoCodigoColocGarantRec = 1
        Else
            dNuevoCodigoColocGarantRec = R!nTotal + 1
        End If
    R.Close
    Set R = Nothing
End Function

Public Sub dInsertColocGarantRecDet(ByVal nNroGarantRec As Long, ByVal nCantidad As Integer, ByVal nPreTasacion As Double, _
        ByVal nPreRealizacion As Double, ByVal cComentario As String, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    lsSQL = "INSERT INTO ColocGarantRecDet(nNroGarantRec, nCantidad, nPreTasacion, nPreRealizacion, cComentario) "
    lsSQL = lsSQL & " VALUES(" & nNroGarantRec & "," & nCantidad & "," & Format(nPreTasacion, "#0.00") & "," & Format(nPreRealizacion, "#0.00") & ",'" & cComentario & "')"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
End Sub

Public Sub dDeleteColocGarantRecDet(ByVal pnNroGarantRec As Long, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
        
    lsSQL = "DELETE ColocGarantRecDet Where nNroGarantRec = " & pnNroGarantRec
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
End Sub

Public Sub dInsertColocMetasAnalista(ByVal psPersCod As String, ByVal pnTipoMeta As Integer, _
    ByVal pnTipoCred As Integer, ByVal pdInicial As Date, ByVal pdFinal As Date, _
    ByVal pnMonto As Double, ByVal pnMoneda As Integer, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT INTO ColocMetasAnalista(cPersCod,nTipoMeta,nTipoCred,dInicial,dFinal,nMonto,nMoneda)"
    lsSQL = lsSQL & " VALUES('" & psPersCod & "'," & pnTipoMeta & "," & pnTipoCred & ",'" & Format(pdInicial, "mm/dd/yyyy") & "','"
    lsSQL = lsSQL & Format(pdFinal, "mm/dd/yyyy") & "'," & Format(pnMonto, "0.00") & "," & pnMoneda & ")"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dUpdateColocMetasAnalista(ByVal psPersCod As String, ByVal pnTipoMeta As Integer, _
    ByVal pnTipoCred As Integer, ByVal pdInicial As Date, ByVal pdFinal As Date, _
    ByVal pnMonto As Double, ByVal pnMoneda As Integer, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE ColocMetasAnalista SET "
    lsSQL = lsSQL & " nMonto = " & Format(pnMonto, "#0.00")
    lsSQL = lsSQL & " WHERE cPersCod = '" & psPersCod & "' AND nTipoMeta = " & pnTipoMeta & " AND nTipoCred = " & pnTipoCred
    lsSQL = lsSQL & " AND dInicial = '" & Format(pdInicial, "mm/dd/yyyy") & "' AND dFinal = '" & Format(pdFinal, "mm/dd/yyyy") & "' AND nMoneda = " & pnMoneda
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dDeleteColocMetasAnalista(ByVal psPersCod As String, ByVal pnTipoMeta As Integer, _
     ByVal pdInicial As Date, ByVal pdFinal As Date, ByVal pnMoneda As Integer, _
    Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "DELETE ColocMetasAnalista "
    lsSQL = lsSQL & " WHERE cPersCod = '" & psPersCod & "' AND nTipoMeta = " & pnTipoMeta
    lsSQL = lsSQL & " dInicial = '" & Format(pdInicial, "mm/dd/yyyy") & "' AND dFinal = '" & Format(pdFinal, "mm/dd/yyyy") & "' AND nMoneda = " & pnMoneda
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dDeleteColocGarantia(ByVal psNumgarant As String, _
        ByVal psCtaCod As String, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    lsSQL = "DELETE ColocGarantia WHERE cNumGarant = '" & psNumgarant & "' "
    lsSQL = lsSQL & " AND cCtaCod = '" & psCtaCod & "' "
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dDeleteColocacEstado(ByVal psCtaCod As String, ByVal pnPrdEstado As Integer, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    lsSQL = "DELETE ColocacEstado WHERE cCtaCod = '" & psCtaCod & "' AND nPrdEstado = " & pnPrdEstado
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dDeleteProductoConceptoFitro(ByVal pnPrdConceptoCod As Integer, _
    ByVal psProdCod As String, ByVal pnMoneda As Integer, ByVal psAgeCod As String, _
    Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String

    lsSQL = "DELETE ProductoConceptoFitro "
    lsSQL = lsSQL & " WHERE nPrdConceptoCod = " & pnPrdConceptoCod & ","
    lsSQL = lsSQL & " cProdCod = '" & psProdCod & "', "
    lsSQL = lsSQL & " nMoneda = " & pnMoneda & ", "
    lsSQL = lsSQL & " cAgeCod = '" & psAgeCod & "' "
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dDeleteProductoConceptoFitroTotal(ByVal pnPrdConceptoCod As Long, _
    Optional psFiltroGaran As String = "P", Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String

    lsSQL = "DELETE ProductoConceptoFiltro "
    lsSQL = lsSQL & " WHERE nPrdConceptoCod = " & pnPrdConceptoCod & " AND  cTpoFiltro = '" & psFiltroGaran & "'"
 
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub


Public Sub dInsertProductoConceptoFitro(ByVal pnPrdConceptoCod As Long, _
    ByVal pnProdCod As Integer, ByVal psAgeCod As String, ByVal psTipoFiltro As String, _
    Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String

    lsSQL = "INSERT INTO ProductoConceptoFiltro (nPrdConceptoCod, nProdCod, cAgeCod,cTpoFiltro)"
    lsSQL = lsSQL & " VALUES(" & pnPrdConceptoCod & "," & pnProdCod & ",'"
    lsSQL = lsSQL & psAgeCod & "','" & psTipoFiltro & "')"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dDeleteColocFteIngreso(ByVal psCtaCod As String, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    lsSQL = "DELETE ColocFteIngreso WHERE cCtaCod = '" & psCtaCod & "' "
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dUpdateGarantias(ByVal psTpoDoc As String, ByVal psNroDoc As String, _
        ByVal pnGravament As Double, Optional pnTipoAcc As Integer = 1, _
        Optional pbEjecBatch As Boolean = False, _
        Optional psEstado As Double, Optional psNroGaran As String)
Dim lsSQL As String
    '1:Actualiza
    '2:Incrementa
    '3:Decrementa
    '4:Actualiza Estado Garantia Real
    Select Case pnTipoAcc
        Case 1
            lsSQL = "UPDATE Garantias SET nGravament = " & Format(pnGravament, "#0.00")
            lsSQL = lsSQL & " WHERE cTpoDoc = '" & psTpoDoc & "' AND cNroDoc = '" & psNroDoc & "'"
        Case 2
            lsSQL = "UPDATE Garantias SET nGravament = nGravament +" & Format(pnGravament, "#0.00")
            lsSQL = lsSQL & " WHERE cTpoDoc = '" & psTpoDoc & "' AND cNroDoc = '" & psNroDoc & "'"
        Case 3
            lsSQL = "UPDATE Garantias SET nGravament = nGravament - " & Format(pnGravament, "#0.00")
            lsSQL = lsSQL & " WHERE cTpoDoc = '" & psTpoDoc & "' AND cNroDoc = '" & psNroDoc & "'"
        Case 4
            lsSQL = " Update Garantias"
            lsSQL = lsSQL & " SET nEstado=" & psEstado
            lsSQL = lsSQL & " Where cNumGarant=" & psNroGaran

    End Select

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

'Copia Fiel de la Clase NCredito para poder realizar el pago en Lote en Una Sola Transaccion

Public Function AmortizarCreditoLote(ByVal psCtaCod, ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant, _
            ByVal pnMonto As Double, ByVal pdHoy As Date, ByVal psMetLiquid As String, _
            ByVal pnTipoPago As ColocTipoPago, ByVal psCodAge As String, ByVal psCodUser As String, _
            ByVal nMovNro As Long, ByVal pnPrdEstado As Integer, ByVal pnTransacc As Long, Optional psNroDoc As String = "", Optional pnITF As Double = 0#) As String
            
Dim nEstadoCred As Integer
Dim R As ADODB.Recordset
Dim oCred As DCredito
Dim oCalend As Dcalendario
Dim nTransacc As Long
Dim sLineaCred As String
Dim nMontoColocado As Double
Dim dFecPend As Date
Dim nDiasAtraso As Integer
Dim i As Integer
Dim nNroCalen As Integer
Dim nMontoGasto As Double
Dim nConsCred As String
Dim pnPlazo As Integer
Dim bTran As Boolean
Dim dFechaTran As Date
Dim oNegCred As NCredito

    On Error GoTo ErrorAmortizarCreditoLote
            
    nEstadoCred = pnPrdEstado
    nTransacc = pnTransacc
    
    'Definir Codigo de Operacion de Pago
    Select Case nEstadoCred
            'Si es Credito refinanciado
            Case gColocEstRefMor
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefMorEfec, gCredPagRefMorCh)
            Case gColocEstRefNorm
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefNorEfec, gCredPagRefNorCh)
            Case gColocEstRefVenc
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefVenEfec, gCredPagRefVenCh)
            'si es Credito Normal
            Case gColocEstVigMor
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorNorEfec, gCredPagNorNorCh)
            Case gColocEstVigNorm
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorMorEfec, gCredPagNorMorCh)
            Case gColocEstVigVenc
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorVenEfec, gCredPagNorVenCh)
    End Select
    
    
    Set oNegCred = New NCredito
    
    Set oCred = New DCredito
    Set R = oCred.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
    Set oCred = Nothing
    pnPlazo = IIf(IsNull(R!nPlazo), 0, R!nPlazo)
    R.Close
    Set R = Nothing
    
    Set oCred = New DCredito
    Set R = oCred.RecuperaColocacCred(psCtaCod)
    Set oCred = Nothing
    nNroCalen = R!nNroCalen
    R.Close
    Set R = Nothing
          
    dFechaTran = CDate(Format(pdHoy, "dd/mm/yyyy hh:mm:ss"))
    'Actualiza Producto
    If oNegCred.MatrizEstadoCalendario(MatCalendDistrib) = gColocCalendEstadoPagado Then
        nEstadoCred = gColocEstCancelado
    End If
    Call dUpdateProducto(psCtaCod, , oNegCred.MatrizSaldoCapital(MatCalend, MatCalendDistrib), nEstadoCred, pdHoy, nTransacc + 1, False)
    
    'Actualiza ColocLineaCreditoSaldo
    Set oCred = New DCredito
    Set R = oCred.RecuperaColocaciones(psCtaCod)
    Set oCred = Nothing
    sLineaCred = R!cLineaCred
    R.Close
    Set R = Nothing
    Set R = RecuperaLineasCreditoSaldo(sLineaCred, True)
    nMontoColocado = R!nMontoColocado
    R.Close
    Set R = Nothing
    nMontoColocado = nMontoColocado - pnMonto
    Call dUpdateLineaCreditoSaldo(sLineaCred, , , nMontoColocado, , , False)
    
    'Actualiza ColocacCred
    dFecPend = oNegCred.MatrizFechaCuotaPendiente(MatCalend, MatCalendDistrib)
    If oNegCred.MatrizCuotaPendiente(MatCalend, MatCalendDistrib) = 0 Then
        nDiasAtraso = 0
    Else
        nDiasAtraso = pdHoy - dFecPend
    End If
    
    Call dUpdateColocacCred(psCtaCod, nDiasAtraso, , , , , , , oNegCred.MatrizCuotaPendiente(MatCalend, MatCalendDistrib), , , , , , , , , , False)
        
    
    Call dInsertMovCol(nMovNro, nConsCred, psCtaCod, nNroCalen, pnMonto, 0, psMetLiquid, pnPlazo, False, pnPrdEstado)
    
    'ITF
    Call dInsertMovCol(nMovNro, gCredITF, psCtaCod, CLng(nNroCalen), pnITF, 0, psMetLiquid, pnPlazo, False, pnPrdEstado)
    Call dInsertMovColDet(nMovNro, gCredITF, psCtaCod, CLng(nNroCalen), gConcITFCliente, 0, pnITF, False)
    
    'Actualiza calendario (ColocCalendario y ColocCalendDet)
    For i = 0 To UBound(MatCalendDistrib) - 1
        Call dUpdateColocCalendario(psCtaCod, nNroCalen, CInt(MatCalendDistrib(i, 1)), gColocCalendAplCuota, , CInt(MatCalendDistrib(i, 2)), "Pago de Cuota", gColocCalendConceptoProcAprobado, False, , Format(dFechaTran, "mm/dd/yyyy"))
        
        'Amortizando Capital
        If CDbl(MatCalendDistrib(i, 3)) > 0 Then
            Call dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocConceptoCodCapital, , CDbl(MatCalendDistrib(i, 3)), , False, True)
            'Inserta Detalle Movimiento Capital
            Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodCapital, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 3)), False)
        End If
        'Amortizando Interes Compensatorio
        If CDbl(MatCalendDistrib(i, 4)) > 0 Then
            Call dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocConceptoCodInteresCompensatorio, , CDbl(MatCalendDistrib(i, 4)), , False, True)
            'Inserta Detalle Movimiento Interes Compensatorio
            Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresCompensatorio, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 4)), False)
        End If
        'Amortizando Interes Gracia
        If CDbl(MatCalendDistrib(i, 5)) > 0 Then
            Call dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocConceptoCodInteresGracia, , CDbl(MatCalendDistrib(i, 5)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresGracia, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 5)), False)
        End If
        'Amortizando Interes Moratorio
        If CDbl(MatCalendDistrib(i, 6)) > 0 Then
            Call dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocConceptoCodInteresMoratorio, , CDbl(MatCalendDistrib(i, 6)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresMoratorio, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 6)), False)
        End If
        'Amortizando Interes Reprog
        If CDbl(MatCalendDistrib(i, 7)) > 0 Then
            Call dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocConceptoCodInteresReprogramado, , CDbl(MatCalendDistrib(i, 7)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresReprogramado, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 7)), False)
        End If
        'Amortizando Interes Suspenso
        If CDbl(MatCalendDistrib(i, 8)) > 0 Then
            Call dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocConceptoCodInteresSuspenso, , CDbl(MatCalendDistrib(i, 8)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresSuspenso, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 7)), False)
        End If
        
        'Amortizando Interes Compensatorio Vencido
        If CDbl(MatCalendDistrib(i, 11)) > 0 Then
            Call dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocConceptoCodInteresCompVencido, , CDbl(MatCalendDistrib(i, 11)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresCompVencido, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 11)), False)
        End If
        
        'Amortizando Gastos
        If CDbl(MatCalendDistrib(i, 9)) > 0 Then
            nMontoGasto = CDbl(MatCalendDistrib(i, 9))
            Set oCalend = New Dcalendario
            Set R = oCalend.RecuperaCalendarioGastos(psCtaCod, nNroCalen, CInt(MatCalendDistrib(i, 1)), gColocCalendAplCuota)
            Set oCalend = Nothing
            Do While Not R.EOF
                If nMontoGasto >= R!nMonto Then
                    Call dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), R!nPrdConceptoCod, , R!nMonto, , False, True)
                    'Inserta Detalle Movimiento Gastos
                    Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), R!nPrdConceptoCod, CInt(MatCalendDistrib(i, 1)), R!nMonto, False)
                    nMontoGasto = nMontoGasto - R!nMonto
                Else
                    Call dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), R!nPrdConceptoCod, , nMontoGasto, , False, True)
                    'Inserta Detalle Movimiento Gastos
                    Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), R!nPrdConceptoCod, CInt(MatCalendDistrib(i, 1)), CDbl(Format(nMontoGasto, "#0.00")), False)
                    nMontoGasto = 0
                End If
                
                nMontoGasto = CDbl(Format(nMontoGasto, "#0.00"))
                If nMontoGasto = 0 Then
                    Exit Do
                End If
                R.MoveNext
             Loop
             R.Close
             Set R = Nothing
        End If
    Next i
    Exit Function

ErrorAmortizarCreditoLote:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Sub dExtornaSaldosCalendario(ByVal pnMovNro As Long, _
    ByVal nAplicado As ColocCalendApl, ByVal psCtaCod As String, Optional pbEjecBatch As Boolean = False)
Dim sSQL As String
    
    sSQL = "UPDATE ColocCalendario SET nColocCalendEstado = " & gColocCalendEstadoPendiente & ", dPago = NULL "
    sSQL = sSQL & " From ColocCalendario CC Inner Join MovColDet MC ON CC.cCtaCod = MC.cCtacod "
    sSQL = sSQL & "     And CC.nNroCalen = MC.nNroCalen "
    sSQL = sSQL & "     And CC.nCuota = MC.nNroCuota "
    sSQL = sSQL & "     Where MC.nMovNro = " & pnMovNro & " And CC.nColocCalendApl = " & nAplicado
    sSQL = sSQL & "     AND CC.cCtaCod = '" & psCtaCod & "'"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSQL
    Else
        coConex.Ejecutar sSQL
    End If
    
    sSQL = "UPDATE ColocCalendDet SET nMontoPagado = nMontoPagado - ABS(MC.nMonto) "
    sSQL = sSQL & " From ColocCalendDet CC Inner Join MovColDet MC ON CC.cCtaCod = MC.cCtacod "
    sSQL = sSQL & "     And CC.nNroCalen = MC.nNroCalen  And CC.nPrdConceptoCod = MC.nPrdConceptoCod"
    sSQL = sSQL & "     And CC.nCuota = MC.nNroCuota "
    sSQL = sSQL & "     Where MC.nMovNro = " & pnMovNro & " And CC.nColocCalendApl = " & nAplicado
    sSQL = sSQL & "     AND CC.cCtaCod = '" & psCtaCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSQL
    Else
        coConex.Ejecutar sSQL
    End If
    'Modificacion:LMMD por RFA extornar el interes diferido hacia el capital de la cuota
    
    sSQL = "Update ColocCalendDet"
    sSQL = sSQL & " Set nMontoPagado = nMontoPagado - Abs(MC.nMonto)"
    sSQL = sSQL & " From ColocCalendDet CC"
    sSQL = sSQL & " Inner Join MovColDet MC ON CC.cCtaCod = MC.cCtacod  And CC.nNroCalen = MC.nNroCalen  And CC.nPrdConceptoCod = 1000"
    sSQL = sSQL & " And CC.nCuota = MC.nNroCuota"
    sSQL = sSQL & " Where MC.nMovNro =" & pnMovNro & " And CC.nColocCalendApl =" & nAplicado & "  AND CC.cCtaCod = '" & psCtaCod & "' and MC.nPrdConceptoCod='124353'"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSQL
    Else
        coConex.Ejecutar sSQL
    End If
    
    
    'Extorno de Gastos que se cargaron al Momento de la Operacion
    sSQL = "UPDATE ColocCalendDet SET nMonto = CC.nMonto - ABS(MC.nMonto) "
    sSQL = sSQL & " From ColocCalendDet CC Inner Join MovColDet MC ON CC.cCtaCod = MC.cCtacod "
    sSQL = sSQL & "     And CC.nNroCalen = MC.nNroCalen  And CC.nPrdConceptoCod = MC.nPrdConceptoCod"
    sSQL = sSQL & "     And CC.nCuota = MC.nNroCuota "
    sSQL = sSQL & "     Where MC.nMovNro = " & pnMovNro & " And CC.nColocCalendApl = " & nAplicado
    sSQL = sSQL & "     AND CC.cCtaCod = '" & psCtaCod & "' AND "
    sSQL = sSQL & " MC.nPrdConceptoCod in ( Select nPrdConceptoCod from ProductoConcepto where nPrdConceptoCod like '12%' "
    sSQL = sSQL & " AND cAplicaProceso in ('CA','PA','PP')) "
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSQL
    Else
        coConex.Ejecutar sSQL
    End If
    

    
    
    If nAplicado = gColocCalendAplDesembolso Then
        sSQL = "Delete ColocCargoAutoma Where cCtaCod = '" & psCtaCod & "'"
        coConex.Ejecutar sSQL
    End If
End Sub

Public Sub dInsertZonas(ByVal psUbiGeoCod As String, ByVal psUbiGeoDescripcion As String, Optional pbEjecBatch As Boolean = False)
Dim sSQL As String
    
    sSQL = "INSERT INTO UbicacionGeografica(cUbiGeoCod, cUbiGeoDescripcion) "
    sSQL = sSQL & " VALUES('" & psUbiGeoCod & "','" & psUbiGeoDescripcion & "')"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSQL
    Else
        coConex.Ejecutar sSQL
    End If
End Sub

Public Sub dUpdateZonas(ByVal psUbiGeoCod As String, ByVal psUbiGeoDescripcion As String, Optional pbEjecBatch As Boolean = False)
Dim sSQL As String

    sSQL = "UPDATE UbicacionGeografica SET cUbiGeoDescripcion = '" & psUbiGeoDescripcion & "' "
    sSQL = sSQL & " Where cUbiGeoCod = '" & psUbiGeoCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSQL
    Else
        coConex.Ejecutar sSQL
    End If
End Sub

Public Sub dDeleteZonas(ByVal psUbiGeoCod As String, Optional pbEjecBatch As Boolean = False)
Dim sSQL As String

    sSQL = "DELETE UbicacionGeografica  Where cUbiGeoCod = '" & psUbiGeoCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSQL
    Else
        coConex.Ejecutar sSQL
    End If
End Sub

Public Sub dInsertFeriado(ByVal dFeriado As Date, ByVal cDescrip As String, ByVal cUltimaActualizacion As String, Optional pbEjecBatch As Boolean = False)
Dim sSQL As String
    
    sSQL = "INSERT INTO Feriado(dFeriado, cDescrip, cUltimaActualizacion) "
    sSQL = sSQL & " VALUES('" & Format(dFeriado, "mm/dd/yyyy") & "','" & cDescrip & "','" & cUltimaActualizacion & "')"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSQL
    Else
        coConex.Ejecutar sSQL
    End If
End Sub

Public Sub dDeleteFeriado(ByVal dFeriado As Date, Optional pbEjecBatch As Boolean = False)
Dim sSQL As String
    
    sSQL = "DELETE Feriado Where dFeriado = '" & Format(dFeriado, "mm/dd/yyyy") & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSQL
    Else
        coConex.Ejecutar sSQL
    End If

End Sub


Public Sub dInsertAgencia(ByVal cAgeCod As String, ByVal cAgeDescripcion As String, ByVal cAgeDireccion As String, _
    ByVal cAgeTelefono As String, ByVal cSubCtaCod As String, ByVal nAgeEspecial As Integer, _
    ByVal cUbiGeoCod As String, ByVal cUltimaActualizacion As String, Optional pbEjecBatch As Boolean = False)
Dim sSQL As String

    sSQL = "INSERT INTO Agencias(cAgeCod, cAgeDescripcion, cAgeDireccion, cAgeTelefono, cSubCtaCod, nAgeEspecial, cUbiGeoCod, cUltimaActualizacion) "
    sSQL = sSQL & " VALUES('" & cAgeCod & "','" & cAgeDescripcion & "','" & cAgeDireccion & "','" & cAgeTelefono & "','" & cSubCtaCod & "'," & nAgeEspecial & ",'" & Trim(Right(cUbiGeoCod, 15)) & "','" & cUltimaActualizacion & "')"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSQL
    Else
        coConex.Ejecutar sSQL
    End If

End Sub

Public Sub dUpdateAgencia(ByVal cAgeCod As String, ByVal cAgeDescripcion As String, ByVal cAgeDireccion As String, _
    ByVal cAgeTelefono As String, ByVal cSubCtaCod As String, ByVal nAgeEspecial As Integer, _
    ByVal cUbiGeoCod As String, ByVal cUltimaActualizacion As String, Optional pbEjecBatch As Boolean = False)

Dim sSQL As String

    sSQL = "Update Agencias SET "
    sSQL = sSQL & " cAgeDescripcion = '" & cAgeDescripcion & "',"
    sSQL = sSQL & " cAgeDireccion = '" & cAgeDireccion & "',"
    sSQL = sSQL & " cAgeTelefono = '" & cAgeTelefono & "',"
    sSQL = sSQL & " cSubCtaCod = '" & cSubCtaCod & "',"
    sSQL = sSQL & " nAgeEspecial = " & nAgeEspecial & ","
    sSQL = sSQL & " cUbiGeoCod = '" & Trim(Right(cUbiGeoCod, 20)) & "',"
    sSQL = sSQL & " cUltimaActualizacion = '" & cUltimaActualizacion & "' "
    sSQL = sSQL & " Where cAgeCod = '" & cAgeCod & "'"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSQL
    Else
        coConex.Ejecutar sSQL
    End If

End Sub

Public Sub dDeleteAgencia(ByVal cAgeCod As String, Optional pbEjecBatch As Boolean = False)
Dim sSQL As String
    sSQL = "DELETE Agencias "
    sSQL = sSQL & " Where cAgeCod = '" & cAgeCod & "'"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSQL
    Else
        coConex.Ejecutar sSQL
    End If

End Sub
'Para los extornos de Mivivienda - CAFF
Public Sub dDeleteColocCalifMiViv(ByVal psCodCta As String)
Dim sSQL  As String

    sSQL = "DELETE ColocCalifMiViv WHERE cCtaCod = '" & psCodCta & "'"
    coConex.Ejecutar sSQL

End Sub

'Por Error en la amortizacion se cuelga al abrir otra conexion - CAFF
Public Function RecuperaColocacCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaColocacCred
    sSQL = "Select * from ColocacCred where cCtacod = '" & psCtaCod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocacCred = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocacCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'Por Error en la amortizacion se cuelga al abrir otra conexion - CAFF
Public Function RecuperaColocaciones(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaColocaciones
    sSQL = "Select * from Colocaciones where cCtaCod = '" & psCtaCod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocaciones = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaColocaciones:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

'Por Error en la amortizacion se cuelga al abrir otra conexion - CAFF
Public Function RecuperaCalendarioGastos(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
                    ByVal pnNroCuota As Integer, ByVal pnAplicado As ColocCalendApl, _
                    Optional ByVal pbTodos As Boolean = False) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta


    On Error GoTo ErrorRecuperaCalendarioGastos
    sSQL = "Select * from ColocCalendDet where cCtaCod = '" & psCtaCod & _
            "' AND nNroCalen = " & pnNroCalen
    If Not pbTodos Then
        sSQL = sSQL & " AND nCuota = " & pnNroCuota
    End If
    sSQL = sSQL & " AND nColocCalendApl = " & pnAplicado
    sSQL = sSQL & " AND nPrdConceptoCod like '12%'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioGastos = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaCalendarioGastos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

