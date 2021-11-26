VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMCredActBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'* Modulo de Colocaciones Pignoraticio*
'* Clase con todas las Actualizaciones a BD
Option Explicit
Dim csConexion As String
Dim csNegocio As String
Dim csCentralPer As String
Dim csCentralCom As String
Dim csCentralImg As String
Dim csAdminist As String

Dim oError As COMConecta.COMErrorHandling
Public coConex As COMConecta.DCOMConecta

'******************************************************************************************
'************************************** Copiado de Clases de Ahorros **********************
'******************************************************************************************

Public Sub dUpdateColocGarantia(ByVal psNumGarant As String, psCodCta As String, pnEstado As Integer)
Dim sSql As String

    sSql = "UPDATE ColocGarantia Set nEstado = " & pnEstado & " WHERE cNumGarant = '" & psNumGarant & "' AND cCtaCod = '" & psCodCta & "'"
    coConex.Ejecutar sSql

End Sub

Public Sub dEliminaCargoResponsable(ByVal psCodCargo As String, ByVal psCodusu As String)
Dim sSql As String

    sSql = "DELETE credcargosusu Where cCodCargo = '" & psCodCargo & "' and ccodusu = '" & psCodusu & "'"
    
    coConex.Ejecutar sSql

End Sub

Public Sub dEliminarNivel(ByVal pnItem As Integer, ByVal pcCodRango As String, ByVal pcCodCargo As String)
Dim sSql As String

    sSql = "DELETE CredRangoCargo Where nItem = " & pnItem & " and cCodRango = '" & pcCodRango & "' AND cCodCargo = '" & pcCodCargo & "'"
    coConex.Ejecutar sSql


End Sub

Public Sub dEliminaCargo(ByVal psCodCargo As String)
Dim sSql As String

    sSql = "DELETE credcargos Where cCodCargo = '" & psCodCargo & "'"
    
    coConex.Ejecutar sSql
    
End Sub

Public Sub dEliminaRango(ByVal psCodRango As String)
Dim sSql As String

    sSql = "DELETE CredRangos Where cCodRango = '" & psCodRango & "'"
    
    coConex.Ejecutar sSql
    
End Sub


Public Sub dInsertaNiveles(ByVal nItem As Integer, ByVal cCodRango As String, _
    ByVal cCodCargo As String, ByVal pbComen As Boolean)
Dim sSql As String

        sSql = "INSERT INTO CredRangoCargo(nItem ,cCodRango,cCodCargo,bComentario) "
        sSql = sSql & "  VALUES(" & nItem & ",'" & cCodRango & "','" & cCodCargo & "',"
        sSql = sSql & IIf(pbComen, "1", "0") & ")"
            
    coConex.Ejecutar sSql
    
End Sub

Public Sub dInsertaCredCargosUsu(ByVal psCodCargo As String, ByVal psCodusu As String, ByVal psNomUsu As String, ByVal cAgeCod As String)
Dim sSql As String

    sSql = "INSERT INTO CredCargosUsu(cCodCargo, cCodUsu, cNomUsu,cAgeCod) "
    sSql = sSql & "  VALUES('" & psCodCargo & "','" & psCodusu & "','" & psNomUsu & "','" & cAgeCod & "')"
    
    coConex.Ejecutar sSql
    
End Sub

Public Sub dInsertaCredAprobacion(ByVal psCtaCod As String, ByVal psCodCargo As String, ByVal psCodusu As String, ByVal psNomUsu As String, ByVal pdFecApr As Date, ByVal psComen As String)
Dim sSql As String
    
    sSql = "UPDATE CredAprobacion SET cCodUsu ='" & psCodusu & "', cNomUsu = '" & psNomUsu & "', dFecApr = '" & Format(pdFecApr, "mm/dd/yyyy") & "', cComen = '" & psComen & "'"
    sSql = sSql & " Where  cCtaCod = '" & psCtaCod & "' AND cCodCargo = '" & psCodCargo & "'"
    
    coConex.Ejecutar sSql
    
End Sub

Public Sub dInsertaCredRangos(ByVal psCodRango As String, ByVal psCodProd As String, ByVal pnMoneda As Integer, ByVal pnTipoCred As Integer, ByVal pnMontoMax As Double, ByVal pnMontoMin As Double)
Dim sSql As String

    sSql = "INSERT INTO CredRangos(cCodRango, cCodProd, nMoneda, nTipoCred, nMontoMax, nMontoMin) "
    sSql = sSql & "  VALUES('" & psCodRango & "','" & psCodProd & "'," & pnMoneda & "," & pnTipoCred & "," & pnMontoMax & "," & pnMontoMin & ")"
    
    coConex.Ejecutar sSql
    
End Sub

Public Sub dInsertaCredCargos(ByVal psCodCargo As String, ByVal psNomCargo As String, _
                            ByVal pbTodaAgencia As Integer)
Dim sSql As String

    sSql = "INSERT INTO credcargos(cCodCargo, cNomCargo,bTodaAgencia) "
    sSql = sSql & "  VALUES('" & psCodCargo & "','" & psNomCargo & "'," & pbTodaAgencia & ")"
    
    coConex.Ejecutar sSql

End Sub

Public Sub dInsertaDesembolsosReprogramado(ByVal psCtaCod As String, pnNroCalen As Integer)
    Dim sSql As String
    
    sSql = "insert into ColocCalendario"
    sSql = sSql & " Select cCtaCod," & pnNroCalen & " as nNroCalen,nColocCalendApl,nCuota,dVenc,dPago,nColocCalendEstado,cDescripcion,cColocCalenFlag,nCalendProc,cColocMiVivEval"
    sSql = sSql & " From ColocCalendario"
    sSql = sSql & " Where cCtaCod='" & psCtaCod & "' and nColocCalendApl=0 "
    sSql = sSql & " and nNroCalen=(select nnrocalen from colocaccred where cctacod='" & psCtaCod & "')"
    
    coConex.Ejecutar sSql
    
    sSql = "insert into ColocCalendDet"
    sSql = sSql & " Select cCtaCod," & pnNroCalen & " as nNroCalen, nColocCalendApl,nCuota,nPrdConceptoCod,nMonto,nMontoPagado,cFlag"
    sSql = sSql & " From ColocCalendDet"
    sSql = sSql & " Where cCtaCod='" & psCtaCod & "' and nColocCalendApl=0"
    sSql = sSql & " and nNroCalen=(select nnrocalen from colocaccred where cctacod='" & psCtaCod & "')"
    
    coConex.Ejecutar sSql

    
End Sub


Public Function RecuperaMovimientoCapataciones(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String

    sSql = "Select M.cMovNro, MC.cOpeCod, MC.cCtaCod, MC.nMonto from MovCap MC "
    sSql = sSql & " Inner Join Mov M ON M.nMovNro = MC.nMovNro "
    sSql = sSql & " Where MC.nMovNro = " & pnMovNro
    
    Set RecuperaMovimientoCapataciones = New ADODB.Recordset
    
    RecuperaMovimientoCapataciones.CursorLocation = adUseClient
    RecuperaMovimientoCapataciones.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    RecuperaMovimientoCapataciones.ActiveConnection = Nothing

End Function


'Modificado por GIPO 19-01-2017  *** Vuelto a subir 14-03-2017
Public Function GetDatosCuentaAho(ByVal sCuenta As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

sSql = "Exec sp_sel_DatosCuentaAhorro '" & sCuenta & "' , " & gCaptacEstado & " , " & gProductoCuentaTipo & " , " & gCaptacTipoTasa

Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaAho = rsCta
Set rsCta = Nothing
End Function



Public Function GetDatosMovExtorno(ByVal sDatoBus As String, ByVal dFecha As Date, _
        ByVal nOperacion As CaptacOperacion, Optional nTipoBus As Integer = 0) As ADODB.Recordset

Set GetDatosMovExtorno = GetMovExtorno(sDatoBus, dFecha, nTipoBus)

End Function


Public Function CapExtornoApertura(ByVal sMovNroCap As String, ByVal pnMovNro As Long, ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional bDocumento As Boolean = False, Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional psMotExtorno As Variant)
'Optional psMotExtorno As Variant  *CTI3
Dim sMsgProd As String, sCodUser As String
Dim nProd As Producto
Dim nMovNro As Long
Dim dFecSis As Date
Dim oFun As COMNContabilidad.NCOMContFunciones

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
AgregaMov sMovNroCap, nOperacion, sGlosa, , gMovFlagDeExtorno
nMovNro = GetnMovNro(sMovNroCap)

'***CTI3 (ferimoro) 17102018
dInsertDetExtMovEcoTaxi nMovNro, nOperacion, psMotExtorno(0), psMotExtorno(1)

ActualizaEstadoCuenta sCuenta, gCapEstAnulada, dFecSis, sMovNro
AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, 0, 0
AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
dInsertMovRef nMovNro, pnMovNro, False
UltimaActualizacionCuenta sCuenta, sMovNro
If bDocumento Then
    If nTipoDoc = TpoDocCheque Then
        ActualizaEstadoDocRecEst TpoDocCheque, sNroDoc, sCodIF, sMovNro, gsChqEstExtornado
    End If
End If


End Function


Public Function CapExtornoAbonoAho(ByVal sMovNroCap As String, ByVal pnMovNro As Long, ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String, Optional sLpt As String = "", Optional ByVal psMotExtorno As Variant) As Double

Dim rsCta As ADODB.Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nDiasTranscurridos As Long, i As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
Dim oFun As COMNContabilidad.NCOMContFunciones

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
AgregaMov sMovNroCap, nOperacion, sGlosa, gMovEstContabMovContable, gMovFlagDeExtorno
nMovNro = GetnMovNro(sMovNroCap)
AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto

'***CTI3 (ferimoro)  17102018
'dInsertDetExtMovEcoTaxi nMovNro, nOperacion, psMotExtorno(0), psMotExtorno(1)

'ALPA20130827**************************************************************
'dInsertMovRef nMovNro, pnMovNro, False
dInsertMovRef nMovNro, IIf(pnMovNro = 0, nMovNroBus, pnMovNro), False
'**************************************************************************
UltimaActualizacionCuenta sCuenta, sMovNro

CapExtornoAbonoAho = nSaldoCnt

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "EXTORNO AHORROS"

End Function

Public Function CapExtornoCargoAho(ByVal pnMovNro As Long, ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String = "", Optional sLpt As String = "", Optional ByVal psMotExtorno As Variant) As Double
'Optional ByVal psMotExtorno As Variant --Agregado CTI3
Dim rsCta As ADODB.Recordset
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
        If nOperacion = gAhoExtRetOPCanje Then
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
'FRHU 20140228 RQ14006
'AgregaMov sMovNro, nOperacion, sGlosa, , gMovFlagDeExtorno
If nOperacion <> 990201 Then
    AgregaMov sMovNro, nOperacion, sGlosa, , gMovFlagDeExtorno
    '***CTI3 (ferimoro) 17102018
    'dInsertDetExtMovEcoTaxi sMovNro, nOperacion, psMotExtorno(0), psMotExtorno(1)
End If
'FIN FRHU 20140228
nMovNro = GetnMovNro(sMovNro)
'nMovNro = pnMovNro
AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
UltimaActualizacionCuenta sCuenta, sMovNro

CapExtornoCargoAho = nSaldoCnt

End Function


Public Function GetnMovNro(ByVal sMovNro As String) As Long
Dim rsMov As ADODB.Recordset
Dim sSql As String

sSql = "Select nMovNro From Mov Where cMovNro = '" & sMovNro & "'"
Set rsMov = New ADODB.Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
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
        Optional nPorcRetCTS As Double = 0, Optional sInstitucion As String = "", _
        Optional ByVal pnPrograma As Integer = 0, Optional ByVal pnMontoAbonar As Double, Optional ByVal pnPlazoAbono As Integer, Optional ByVal psPromotor As String, _
        Optional ByVal psRegla As String = "A") As String 'FRHU 20140228 RQ14006 - Se agrego psRegla

'ARCV 13-02-2007
'Se Agrego para los cambios en Captaciones pnPrograma,pnMontoAbonar,pnPlazoAbono
'--------
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim sCuenta As String, sFecha As String
Dim sFecUltCierre As String, sOrdPag As String, sCtaAbonoIntPF As String
Dim nSaldoDisp As Double, nSaldRetCTS As Double
Dim sSql As String
Set clsGen = New COMDConstSistema.DCOMGeneral
sCuenta = clsGen.GeneraNuevaCuenta(sAgencia, nProducto, nMoneda)
Set clsGen = Nothing
sFecha = Format$(dFecha, "mm/dd/yyyy") & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
sSql = "Insert Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) " _
    & "Values ('" & sCuenta & "'," & nTasa & "," & nSaldo & "," & gCapEstActiva & ",'" & sFecha & "',0)"
coConex.ConexionActiva.Execute sSql
nSaldoDisp = IIf(bCheque, 0, nSaldo)
sFecUltCierre = Format$(DateAdd("d", -1, dFecha), "mm/dd/yyyy")

'ARCV 13-02-2007
If nProducto <> gCapAhorros Then
    sSql = "Insert Captaciones (cCtaCod,nSaldoDisp,nPersoneria,nFirmas,nIntAcum,dUltCierre,dApertura,nPrdCtaTpo,nPrdTasaInteres,cUltimaActualizacion) " _
    & "Values ('" & sCuenta & "'," & nSaldoDisp & "," & nPersoneria & "," & nFirmas & ",0,'" & sFecUltCierre & "','" & sFecha & "'," & nTipoCuenta & "," _
    & nTipoTasa & ",'" & sMovNro & "')"
    
    coConex.ConexionActiva.Execute sSql
Else
    'FRHU 20140228 RQ14006 - Se agrego cReglas
    sSql = "Insert Captaciones (cCtaCod,nSaldoDisp,nPersoneria,nFirmas,nIntAcum,dUltCierre,dApertura,nPrdCtaTpo,nPrdTasaInteres,cUltimaActualizacion,nTpoPrograma,cReglas) " _
        & "Values ('" & sCuenta & "'," & nSaldoDisp & "," & nPersoneria & "," & nFirmas & ",0,'" & sFecUltCierre & "','" & sFecha & "'," & nTipoCuenta & "," _
        & nTipoTasa & ",'" & sMovNro & "'," & pnPrograma & ",'" & psRegla & "')"
        
    coConex.ConexionActiva.Execute sSql
End If
'-------

Select Case nProducto
    Case gCapAhorros
        sOrdPag = IIf(bOrdPag, "1", "0")
        'ARCV 13-02-2007
        sSql = "Insert CaptacAhorros (cCtaCod,nSaldoAnterior,bOrdPag,nSobregiro,dUltContacto,nPlazo,nMontoAbono) " _
            & "Values ('" & sCuenta & "',0," & sOrdPag & ",0,'" & sFecha & "'," & pnPlazoAbono & "," & pnMontoAbonar & ")"
        '------
    Case gCapPlazoFijo
        sCtaAbonoIntPF = IIf(bCtaAboInt, "1", "0")
        sSql = "Insert CaptacPlazoFijo (cCtaCod,nPlazo,nIntPag,dRenovacion,nApertura,dAuxiliar,nFormaRetiro,nDuplicado,bAbonoIntCtaAho) " _
            & "Values ('" & sCuenta & "'," & nPlazo & ",0,'" & sFecha & "'," & nSaldo & ",'" & sFecha & "'," & nFormaRetiro & ",0," & sCtaAbonoIntPF & ")"
    Case gCapCTS
        nSaldRetCTS = Round(nSaldo * nPorcRetCTS / 100, 2)
        sSql = "Insert CaptacCTS (cCtaCod,nSaldRetiro,nIntSaldo,cCodInst) " _
            & "Values ('" & sCuenta & "'," & nSaldRetCTS & ",0,'" & sInstitucion & "')"
End Select
coConex.ConexionActiva.Execute sSql
If nProducto = gCapPlazoFijo And bCtaAboInt Then
    sSql = "Insert CaptacCtaAboIntPF (cCtaCod,cCtaCodAbono) " _
        & "Values ('" & sCuenta & "','" & sCtaCodAbono & "')"
    coConex.ConexionActiva.Execute sSql
End If
AgregaNuevaCaptacion = sCuenta
End Function

Public Sub AgregaNuevoProdPers(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona, _
                                Optional ByVal psRegla As String = "A")
'FRHU 20140228 RQ14006 Se agrego psRegla
Dim sSql As String
sSql = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac,cGrupo) " _
    & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ",'" & psRegla & "')"
coConex.ConexionActiva.Execute sSql
End Sub

Public Function GetDatosOrdenPago(ByVal sCuenta As String, ByVal sNroDoc As String) As ADODB.Recordset
Dim rsOP As ADODB.Recordset
Dim sSql As String
sSql = "SELECT E.cMovNro, D.nTpoDoc, D.cNroDoc, D.cCtaCod, D.nMonto, P.cPersNombre, E.cEstado " _
    & "D.cIFCodPers FROM DocRecOPEst E INNER JOIN DocRecOP D INNER JOIN " & csCentralPer & "Persona P " _
    & "ON D.cIFCodPers = P.cPersCod ON E.nTpoDoc = D.nTpoDoc AND E.cNroDoc = D.cNroDoc AND E.cCtaCod " _
    & "= D.cCtaCod WHERE E.cMovNro IN (SELECT MAX(E1.cMovNro) FROM DocRecOPEst E WHERE E1.nTpoDoc = " _
    & "E.nTpoDoc AND E1.cNroDoc = E.cNroDoc AND E1.cCtaCod = E.cCtaCod)"
Set rsOP = New ADODB.Recordset
rsOP.CursorLocation = adUseClient
rsOP.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsOP.ActiveConnection = Nothing
Set GetDatosOrdenPago = rsOP
Set rsOP = Nothing
End Function

Public Sub AgregaOrdenPagoEstado(ByVal sCuenta As String, ByVal sNroDoc As String, _
        ByVal sMovNro As String, ByVal nMonto As Double, ByVal nEstadoOP As CaptacOrdPagoEstado)
Dim sSql As String
sSql = "UPDATE DocRecOP Set nMonto = " & nMonto & " WHERE nTpoDoc = " & TpoDocOrdenPago & " " _
    & "AND cNroDoc = '" & sNroDoc & "' And cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSql
sSql = "INSERT DocRecOPEst (nTpoDoc,cNroDoc,cCtaCod,cMovNro,nMonto,nEstado) " _
    & "VALUES (" & TpoDocOrdenPago & ",'" & sNroDoc & "','" & sCuenta & "','" & sMovNro & "'," & nMonto & "," & nEstadoOP & ")"
coConex.ConexionActiva.Execute sSql
End Sub

Public Sub AgregaCuentaDocumento(ByVal sCuenta As String, ByVal nTpoDoc As TpoDoc, _
        ByVal sNroDoc As String, ByVal sCodIF As String, ByVal sMovNro As String, _
        ByVal nMovNro As Long, Optional nMonto As Double = 0, _
        Optional nEstadoOP As CaptacOrdPagoEstado = gCapOPEstCobrada)
Dim sFecha As String
Dim sSql As String
sFecha = Mid(sMovNro, 1, 4) & "/" & Mid(sMovNro, 5, 2) & "/" & Mid(sMovNro, 7, 2) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
Select Case nTpoDoc
    Case TpoDocCheque
        sSql = "INSERT DocRecCapta (nTpoDoc,cNroDoc,cPersCod,cIFTpo,nMovNro,cCtaCod,nMonto) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "'," & nMovNro & ",'" & sCuenta & "'," & nMonto & ")"
    Case TpoDocOrdenPago
        sSql = "INSERT DocRecOP (nTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers,nEstado,cMovNro) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCuenta & "'," & nMonto & ",'" & sCodIF & "'," & nEstadoOP & ",'" & sMovNro & "')"
    Case TpoDocNotaAbono, TpoDocNotaCargo
End Select
coConex.ConexionActiva.Execute sSql
sSql = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
    & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
coConex.ConexionActiva.Execute sSql
End Sub

Public Sub AgregaMov(ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
        ByVal sGlosa As String, Optional nMovEstado As MovEstado = gMovEstContabMovContable, _
        Optional nMovFlag As MovFlag = gMovFlagVigente)
Dim sSql As String
sSql = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sGlosa & "'," & nMovEstado & "," & nMovFlag & ")"
coConex.ConexionActiva.Execute sSql
End Sub

Public Sub AgregaMovCap(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double)
Dim sSql As String

sSql = "INSERT MovCap (nMovNro,cOpeCod,cCtaCod,nMonto,nSaldoDisponible,nSaldoContable) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nMonto & "," & nSaldoDisp & "," & nSaldoCnt & ")"
coConex.ConexionActiva.Execute sSql
End Sub

Public Sub AgregaMovCapDet(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nConcepto As CaptacConcepto, ByVal nMonto As Double)
Dim sSql As String

sSql = "INSERT MovCapDet (nMovNro,cOpeCod,cCtaCod,nConceptoCod,nMonto) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nConcepto & "," & nMonto & ")"
coConex.ConexionActiva.Execute sSql
End Sub
Public Sub AgregaMovTipoCambio(ByVal nMovNro As Long, ByVal pnTipoCambio As Currency)
Dim sSql As String

sSql = " INSERT MovTpoCambio (nMovNro, nMovTpoCambio) " _
     & " VALUES (" & nMovNro & "," & pnTipoCambio & ")"

coConex.ConexionActiva.Execute sSql
End Sub

Public Sub ActualizaSaldoAnteriorAho(ByVal sCuenta As String, ByVal nSaldoAnt As Double)
Dim sSql As String

sSql = "Update CaptacAhorros Set nSaldoAnterior = " & nSaldoAnt & "  WHERE cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSql
End Sub
'By Capi 20012008
Public Sub ExoneraInactivas(ByVal sCuenta As String, ByVal sMovNro As String)
Dim sSql As String

sSql = "Update CaptacAhorros Set bNoCobroInactivas =1" & ",cNoCobroInactivas= '" & sMovNro & "' WHERE cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSql
End Sub


Public Sub ActualizaAbonoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)
Dim sSql As String

If bActExtracto Then
    sSql = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSql = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
coConex.ConexionActiva.Execute sSql
sSql = "Update Captaciones Set nSaldoDisp = nSaldoDisp + " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSql

End Sub

Public Sub ActualizaCargoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)

Dim sSql As String

If bActExtracto Then
    sSql = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSql = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
coConex.ConexionActiva.Execute sSql
sSql = "Update Captaciones Set nSaldoDisp = nSaldoDisp - " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSql
End Sub

Public Sub UltimaActualizacionCuenta(ByVal sCuenta As String, ByVal sMovNro As String)
'Actualiza la fecha de ultima actualizacion de la cuenta de captaciones
Dim sSql As String
sSql = "Update Captaciones Set cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSql
End Sub

Public Sub ActualizaEstadoCuenta(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, _
        ByVal dFecha As Date, ByVal sMovNro As String)
'Actualiza el ultimo estado a la tabla de producto
Dim sSql As String
sSql = "Update Producto Set nPrdEstado = " & nEstado & ", " _
    & "dPrdEstado = '" & Format$(dFecha, "mm/dd/yyyy") & "' WHERE " _
    & "cCtaCod = '" & sCuenta & "'"
coConex.ConexionActiva.Execute sSql
'Agrega un registro mas de estado a la historia de estados de la cuenta
sSql = "INSERT CaptacEstado (cCtaCod,nPrdEstado,cMovNro) " _
    & "VALUES ('" & sCuenta & "'," & nEstado & ",'" & sMovNro & "')"
coConex.ConexionActiva.Execute sSql
End Sub

Public Sub ActualizaEstadoDocRecEst(ByVal nTipoDoc As TpoDoc, ByVal sNroDoc As String, _
        ByVal sCodIF As String, ByVal sMovNro As String, ByVal nEstado As ChequeEstado)
'Actualiza el ultimo estado a la tabla de Documento Recibidos
Dim sSql As String
sSql = "Update DocRec Set nEstado = " & nEstado & " WHERE " _
    & "nTpoDoc = '" & nTipoDoc & "' AND cNroDoc = '" & sNroDoc & "' AND cPersCod = '" & sCodIF & "'"
coConex.ConexionActiva.Execute sSql
'Agrega un registro mas de estado a la historia de estados del documento recibido
sSql = "INSERT DocRecEst (nTpoDoc,cNroDoc,cPersCod,cMovNro,nEstado) " _
    & "VALUES ('" & nTipoDoc & "','" & sNroDoc & "','" & sCodIF & "','" & sMovNro & "'," & nEstado & ")"
coConex.ConexionActiva.Execute sSql
End Sub

'JUEZ 20150127 *******************************************************
Public Sub ActualizaInactivaAct(ByVal sCuenta As String)
Dim sSql As String
    sSql = "Update CaptacAhorros Set bInactiva = 0 WHERE cCtaCod = '" & sCuenta & "'"
    coConex.ConexionActiva.Execute sSql
End Sub
'END JUEZ ************************************************************

Public Function GetMovExtorno(ByVal sDatoBus As String, ByVal pdFecSis As Date, Optional nTipoBus As Integer = 0) As ADODB.Recordset
Dim rsMov As ADODB.Recordset
Dim sWhere As String
Dim sSql As String
If nTipoBus = 0 Then
    sWhere = "M.cMovNro LIKE '%" & sDatoBus & "%'"
ElseIf nTipoBus = 1 Then
    sWhere = "M.cMovNro LIKE '" & Format$(pdFecSis, "yyyymmdd") & "%' And C.cCtaCod = '" & sDatoBus & "'"
End If
sSql = "Select M.cMovNro, O.cOpeDesc, C.cCtaCod, ISNULL(MD.nDocTpo,'') nDocTpo, ISNULL(MD.cDocNro,'') cDocNro, " _
    & "M.cMovDesc, C.cOpeCod, C.nMonto, M.nMovNro FROM MovDoc MD RIGHT JOIN Mov M INNER JOIN MovCap C " _
    & "INNER JOIN " & csCentralCom & "OpeTpo O ON C.cOpeCod = O.cOpeCod ON M.nMovNro = C.nMovNro " _
    & "ON MD.nMovNro = M.nMovNro WHERE " & sWhere & " AND C.cOpeCod NOT IN ('" & gAhoEstInacAct & "') " _
    & "AND M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & "," & gMovFlagDeExtorno & ")" _
    & "AND M.nMovEstado IN (" & gMovEstContabMovContable & ") " _
    & "ORDER BY M.nMovNro"
Set rsMov = New ADODB.Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetMovExtorno = rsMov
End Function

Public Sub ActualizaEstadoMov(ByVal nMovNro As Long, ByVal nMovFlag As MovFlag)
Dim sSql As String
sSql = "Update Mov Set nMovFlag = " & nMovFlag & " Where nMovNro = " & nMovNro
coConex.ConexionActiva.Execute sSql
End Sub

Public Function EsOrdenPagoEmitida(ByVal sCuenta As String, ByVal nNumOP As Long) As Boolean
Dim rsOP As ADODB.Recordset
Set rsOP = New ADODB.Recordset
Dim sSql As String
rsOP.CursorLocation = adUseClient
sSql = "Select OP.cCtaCod, OP.nInicio, OP.nFin FROM Mov M INNER JOIN MovDocEmitidoRango OP ON " _
    & "M.nMovNro = OP.nMovNro WHERE M.nMovFlag <> " & gMovFlagExtornado & " AND OP.cCtaCod = '" _
    & sCuenta & "' And " & nNumOP & " Between OP.nInicio And OP.nFin"
rsOP.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
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
'Dim clsGen As DCOMGeneral
'Set clsGen = New DCOMGeneral
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
Dim rsVar As ADODB.Recordset
Dim sSql As String
sSql = "SELECT nParValor FROM Parametro WHERE nParCod = " & nParametro & " " _
    & "And nParProd = " & gPrdParamCaptac
Set rsVar = New ADODB.Recordset
rsVar.CursorLocation = adUseClient
rsVar.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsVar.ActiveConnection = Nothing
If rsVar.EOF And rsVar.BOF Then
    GetCapParametro = 0
Else
    GetCapParametro = rsVar("nParValor")
End If
rsVar.Close
Set rsVar = Nothing
End Function

Public Function GetDatosCuentaCTS(ByVal sCuenta As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas,c.nfirmasmin ,C.nIntAcum, dUltCierre=( case when dbo.ultimodiames(dUltCierre)=1 then dUltCierre else dateadd(dd,1, dUltCierre) end )  , C.dApertura, AG.cAgeDescripcion, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.nSaldRetiro, A.nIntSaldo, PE.cPersNombre cInstitucion, " _
    & "T.cConsDescripcion cEstado, T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.cCodInst,C.cAlias " _
    & "FROM Producto P INNER JOIN Captaciones C INNER JOIN CaptacCTS A INNER JOIN " _
    & "Persona PE ON A.cCodInst = PE.cPersCod ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod " _
    & "WHERE P.cCtaCod = '" & sCuenta & "' AND T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    & "AND T2.nConsCod = " & gCaptacTipoTasa

Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaCTS = rsCta
Set rsCta = Nothing

End Function

Public Function GetDatosCuentaPF(ByVal sCuenta As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, AG.cAgeDescripcion, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.nPlazo, A.nIntPag, A.dRenovacion, A.nApertura, " _
    & "A.nFormaRetiro, A.dAuxiliar, A.nDuplicado, T.cConsDescripcion cEstado, A.bBusCredPend, " _
    & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, T3.cConsDescripcion cRetiro, " _
    & "A.dUltCierreAnt, A.dAuxiliarAnt, ISNULL(CI.cCtaCodAbono,'') cCtaCodAbono,a.nformaretiro,c.calias,c.nfirmasmin FROM Producto P " _
    & "INNER JOIN Captaciones C INNER JOIN CaptacPlazoFijo A LEFT JOIN CaptacCtaAboIntPF CI " _
    & "ON A.cCtaCod = CI.cCtaCod ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Constante T3 ON A.nFormaRetiro = T3.nConsValor " _
    & "INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod WHERE P.cCtaCod = '" & sCuenta & "' AND " _
    & "T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    & "AND T2.nConsCod = " & gCaptacTipoTasa & " AND T3.nConsCod = " & gCaptacPFFormaRetiro


Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set GetDatosCuentaPF = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing

End Function


Public Function GetDatosCuenta(ByVal sCuenta As String) As ADODB.Recordset
Dim rsCta As ADODB.Recordset
Dim nProd As Producto
nProd = CInt(Mid(sCuenta, 6, 3))

Select Case nProd
    Case gCapAhorros
        Set rsCta = GetDatosCuentaAho(sCuenta)
        
    Case gCapPlazoFijo
        Set rsCta = GetDatosCuentaPF(sCuenta)
        
    Case gCapCTS
        Set rsCta = GetDatosCuentaCTS(sCuenta)
        
End Select
Set GetDatosCuenta = rsCta
Set rsCta = Nothing

End Function


Public Function ValidaSaldoCuenta(ByVal sCuenta As String, ByVal nMonto As Double) As Boolean
Dim nMoneda As Moneda
Dim rsCta As ADODB.Recordset
Dim nSaldo As Double, nMontoMinimo As Double

Set rsCta = GetDatosCuenta(sCuenta)
nSaldo = rsCta("nSaldoDisp")
rsCta.Close
Set rsCta = Nothing

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

Public Function GetNombreTitulares(ByVal sCuenta As String, Optional pbNombreCompleto As Boolean = False, Optional pnTitulares As Integer = 0, Optional pnMargenIzquierdo As Integer = 0) As String
GetNombreTitulares = GetNombreTitulares(sCuenta, pbNombreCompleto, pnTitulares, pnMargenIzquierdo)

End Function

Private Function EmiteBoleta(ByVal sMsgProd As String, ByVal sMsgOpe As String, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double, _
            ByVal nIntMes As Double, ByVal nExtracto As Long, Optional bDocumento As Boolean = False, Optional nDocumento As TpoDoc = TpoDocCheque, _
            Optional sNroDoc As String = "", Optional dFechaValor As Date, Optional bImpSaldos As Boolean = True, _
            Optional pdFecSis As Date = CDate("01/01/1950"), Optional psNomAge As String = "") As String

Dim bReImp As Boolean
Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
Dim oImpre As COMFunciones.FCOMImpresion

sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = sMsgProd

Set oImpre = New COMFunciones.FCOMImpresion

sNomTit = oImpre.ImpreCarEsp(GetNombreTitulares(sCuenta))

bReImp = False
'Do
    If bDocumento Then
        Select Case nDocumento
            Case TpoDocCheque
                ImprimeBoleta sTipApe, oImpre.ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntMes, "Fecha Valor", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecSis, psNomAge
            Case TpoDocNotaAbono, TpoDocNotaCargo
                ImprimeBoleta sTipApe, oImpre.ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntMes, "", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecSis, psNomAge
            Case TpoDocOrdenPago
                ImprimeBoleta sTipApe, oImpre.ImpreCarEsp(sModDep) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntMes, "", nExtracto, nSaldoCnt, bImpSaldos, , , , , , , , , pdFecSis, psNomAge
        End Select
    End If
    'If MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    '    bReImp = True
    'Else
    '    bReImp = False
    'End If
'Loop Until Not bReImp
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
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", Optional sCodIF As String = "", _
            Optional dFechaValor As Date, Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
            Optional dFecSis As Date = CDate("01/01/1950"), Optional sNomAge As String = "", _
            Optional pbITFAplica As Boolean = True, Optional pnITFValor As Double = 0, Optional pbITFAsumido As Boolean = False, _
            Optional psItfOperacion As String = "", Optional ByVal pnComAsumidaEcotaxi As Double = 0) As Double


Dim rsCta As ADODB.Recordset
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
pMatAho(2) = CDbl(Format(nIntGanado, "#0.00"))

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
    Call AgregaMovCap(nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt)

If bInactiva Then
    Call AgregaMovCap(nMovNro, gAhoEstInacAct, sCuenta, nSaldoInac, nSaldoDisp, nSaldoCnt)
    Call AgregaMovCapDet(nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac)
    Call ActualizaEstadoCuenta(sCuenta, gCapEstActiva, dFecSis, sMovNro)
    Call ActualizaInactivaAct(sCuenta) 'JUEZ 20150127
End If

'EJVG20140717 ***
'Call AgregaMovCapDet(nMovNro, nOperacion, sCuenta, gConcCapital, nMonto)
Call AgregaMovCapDet(nMovNro, nOperacion, sCuenta, gConcCapital, nMonto - pnComAsumidaEcotaxi)
If pnComAsumidaEcotaxi > 0 Then
    Call AgregaMovCapDet(nMovNro, nOperacion, sCuenta, gConcComAsumidoEcotaxi, pnComAsumidaEcotaxi)
End If
'END EJVG *******
Call UltimaActualizacionCuenta(sCuenta, sMovNro)
Call ActualizadUltContactoAho(sCuenta, dFecSis) 'EJVG20130920 Actualizacion Ultimo Contacto x Inactivas
CapAbonoCuentaAho = nSaldoCnt

pMatAho(8) = nSaldoCnt
pMatAho(9) = nSaldoCnt

If pbITFAplica And pnITFValor > 0 Then
    If pbITFAsumido Then
        AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFAsumido, pnITFValor
    Else
        AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp - pnITFValor, nSaldoCnt - pnITFValor
        AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFCliente, pnITFValor
    End If
End If

'If sNroDoc = "" Then ' Si la operacion es en efectivo
If (pbITFAplica And pnITFValor > 0) And (Not pbITFAsumido) And (val(psItfOperacion) = gITFCobroCargo Or val(psItfOperacion) = gITFCobroCMACCargo) Then
    ActualizaCargoCaptacion sCuenta, pnITFValor, pnITFValor, nIntGanado, dUltMov, sMovNro, bNumExtracto
    CapAbonoCuentaAho = nSaldoCnt - pnITFValor
    pMatAho(8) = nSaldoCnt - pnITFValor
    pMatAho(9) = nSaldoCnt - pnITFValor
End If


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
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, ByVal pdFecSis As Date, _
            Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional bOPExiste As Boolean = False, _
            Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
            Optional sCodCMAC As String = "", Optional sNomCmac As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "", Optional bCancelaciones As Boolean = False, _
            Optional pnMonedaCred As Moneda = 0, _
            Optional pbITFAplica As Boolean = True, Optional pnITFValor As Double = 0, Optional pbITFAsumido As Boolean = False, _
            Optional psItfOperacion As String = "", Optional pbRetiroEfectivo As Boolean = False) As Double
            'FRHU RQ14008 Se agrego: pbRetiroEfectivo

Dim rsCta As ADODB.Recordset
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
Dim lnMontoOri As Double
Dim oGen As COMDConstSistema.DCOMGeneral
Dim lnTipCambioV As Currency
Dim lnTipCambioC As Currency
Dim lnTipCambio As Currency

Dim clsMov As COMNContabilidad.NCOMContFunciones

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

Set oGen = New COMDConstSistema.DCOMGeneral

'GetTipCambio pdFecSis
lnTipCambioC = oGen.EmiteTipoCambio(pdFecSis, TCCompra)
lnTipCambioV = oGen.EmiteTipoCambio(pdFecSis, TCVenta)
lnTipCambio = oGen.EmiteTipoCambio(pdFecSis, TCFijoDia)

'RIRO 20200706 ****************
Dim sMensajeReactiva As String
sMensajeReactiva = oGen.LeeConstSistema("725")
If Not Len(Trim(sMensajeReactiva)) > 0 Then
    sMensajeReactiva = "Retiro Comisión Reactiva"
End If
'END RIRO *********************

If pnMonedaCred = 0 Then pnMonedaCred = Mid(sCuenta, 9, 1)

If pnMonedaCred <> Mid(sCuenta, 9, 1) Then
    lnMontoOri = nMonto
    If pnMonedaCred = gMonedaExtranjera And Mid(sCuenta, 9, 1) = gMonedaNacional Then
        nMonto = Round(nMonto * lnTipCambioV, 2)
    Else
        nMonto = Round(nMonto / lnTipCambioC, 2)
    End If
Else
    lnMontoOri = 0
End If

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

'FRHU20140228 RQ14006
If pbRetiroEfectivo Then
    pMatAho(15) = nSaldoDisp
    pMatAho(16) = nSaldoCnt
    pMatAho(17) = Format(nIntGanado, "#0.00")
End If
'FIN FRHU20140228 RQ14006
ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True

If sNroDoc <> "" Then
    If nTipoDoc = TpoDocOrdenPago Then 'Si es con orden de pago
        If bOPExiste Then
            AgregaOrdenPagoEstado sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstCobrada
        Else
            AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMonto, gCapOPEstCobrada
        End If
        If nOperacion = gAhoRetOPCanje Then
            sMsgOpe = "Retiro OP Canje"
        ElseIf nOperacion = gAhoRetOP Then
            sMsgOpe = "Retiro OP"
        ElseIf nOperacion = gAhoOPCertificacion Then
            sMsgOpe = "Certificacion OP"
        
        'RIRO 20200630 ******
        ElseIf nOperacion = "100949" Then
            sMsgOpe = sMensajeReactiva
        'END RIRO ***********
            
        End If
    ElseIf nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = "Retiro NC"
    End If
Else
        'RIRO 20200630 ******
        If nOperacion = "100949" Then
            sMsgOpe = sMensajeReactiva
        'END RIRO ***********
        Else
            sMsgOpe = "Retiro Efectivo"
        End If
End If

'AgregaMov sMovNro, nOperacion, sGlosa

nMovNro = GetnMovNro(sMovNro)
AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
If bInactiva Then
    AgregaMovCap nMovNro, gAhoEstInacAct, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
    AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
    ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
    ActualizaInactivaAct (sCuenta) 'JUEZ 20150127
End If

'RIRO 20200611 Comisión Reactiva ****
If nOperacion = "100949" Then
    AgregaMovCapDet nMovNro, nOperacion, sCuenta, 1251, nMonto
Else
    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
End If
'End RIRO 20200611 ******************

'AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto RIRO20200611 Comentado
'si el pago de credito es en diferente moneda a la cuenta de ahorros
Dim lnMontoCV As Double
If pnMonedaCred <> Mid(sCuenta, 9, 1) Then
    If Mid(sCuenta, 9, 1) = gMonedaExtranjera Then
        lnMontoCV = Round((lnMontoOri / lnTipCambioC) * lnTipCambio, 2)
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeCompra, lnMontoCV
        AgregaMovTipoCambio nMovNro, lnTipCambioC
        
        If lnTipCambioC < lnTipCambio Then
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - nMonto)
        Else
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - nMonto)
        End If
    Else
        lnMontoCV = Round(lnMontoOri * lnTipCambio, 2)
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeVenta, lnMontoCV
        AgregaMovTipoCambio nMovNro, lnTipCambioV
        
        If lnTipCambioV > lnTipCambio Then
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - Round(lnMontoOri * lnTipCambioV, 2))
        Else
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - Round(lnMontoOri * lnTipCambioV, 2))
        End If
    End If
End If
CapCargoCuentaAho = nSaldoCnt

If pbITFAplica And pnITFValor > 0 Then
    If pbITFAsumido Then
        AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFAsumido, pnITFValor
    Else
        ActualizaCargoCaptacion sCuenta, pnITFValor, pnITFValor, nIntGanado, dUltMov, sMovNro, bNumExtracto
        AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp - pnITFValor, nSaldoCnt - pnITFValor
        AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFCliente, pnITFValor
    End If
    
    nSaldoDisp = nSaldoDisp - pnITFValor
    If Not bCancelaciones Then
        pMatAho(10) = nSaldoDisp
        pMatAho(11) = nSaldoCnt - pnITFValor
    Else
        pMatAho(12) = nSaldoDisp
        pMatAho(13) = nSaldoCnt - pnITFValor
    End If
    'FRHU20140429 INCIDENTE: Se agrego para que en el vouche muestre el saldo correcto. cuando hace el retiro
    If pbRetiroEfectivo Then
        pMatAho(15) = nSaldoDisp
        pMatAho(16) = nSaldoCnt - pnITFValor
        pMatAho(17) = Format(nIntGanado, "#0.00")
    End If
    'FIN FRHU20140429
    CapCargoCuentaAho = nSaldoCnt - pnITFValor
End If
UltimaActualizacionCuenta sCuenta, sMovNro

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

Public Function CapAperturaCuenta(ByVal nProducto As Producto, ByVal nMoneda As Moneda, ByVal rsPers As ADODB.Recordset, ByVal sMovNro As String, _
        ByVal sAgencia As String, ByVal nOperacion As CaptacOperacion, ByVal nTasa As Double, ByVal nSaldo As Double, _
        ByVal dFecha As Date, ByVal nFirmas As Integer, ByVal nPersoneria As PersPersoneria, ByVal sGlosa As String, _
        ByVal nTipoCuenta As ProductoCuentaTipo, _
        ByVal nTipoTasa As CaptacTipoTasa, Optional bCheque As Boolean = False, _
        Optional sNroDoc As String, Optional sCodIF As String, _
        Optional ByVal bOrdPag As Boolean = False, Optional nPlazo As Long = 0, _
        Optional nFormaRetiro As CaptacPFFormaRetiro = gCapPFFormRetMensual, _
        Optional bCtaAboInt As Boolean = False, Optional sCtaCodAbono As String = "", _
        Optional nPorcRetCTS As Double = 0, Optional sInstitucion As String = "", _
        Optional sPersLavDinero As String = "", _
        Optional ByVal pnPrograma As Integer = 0, Optional ByVal pnMontoAbonar As Double, Optional ByVal pnPlazoAbono As Integer, Optional ByVal psPromotor As String) As String

'ARCV 13-02-2007 (NUEVOS DATOS DE AHORROS)
'pnPrograma , pnMontoAbonar ,pnPlazoAbono ,psPromotor

Dim nErrNum As Integer
Dim nErrDesc As String, sCuenta As String
Dim bTrans As Boolean
Dim nMovNro As Long


bTrans = False

bTrans = True
'FRHU 20140228 RQ14006 - Se quito psPromotor
'sCuenta = AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo, dFecha, nFirmas, nPersoneria, _
'    nTipoCuenta, sMovNro, nTipoTasa, bCheque, bOrdPag, nPlazo, nFormaRetiro, bCtaAboInt, sCtaCodAbono, nPorcRetCTS, sInstitucion, _
'    pnPrograma, pnMontoAbonar, pnPlazoAbono, psPromotor)
sCuenta = AgregaNuevaCaptacion(nProducto, nMoneda, Mid(sMovNro, 15, 3) & sAgencia, nTasa, nSaldo, dFecha, nFirmas, nPersoneria, _
    nTipoCuenta, sMovNro, nTipoTasa, bCheque, bOrdPag, nPlazo, nFormaRetiro, bCtaAboInt, sCtaCodAbono, nPorcRetCTS, sInstitucion, _
    pnPrograma, pnMontoAbonar, pnPlazoAbono)
'FIN FRHU 20140228
    
'AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = GetnMovNro(sMovNro)
Select Case nOperacion
    Case gAhoApeChq, gPFApeChq, gCTSApeChq
        AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, 0, nSaldo
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo
        AgregaCuentaDocumento sCuenta, TpoDocCheque, sNroDoc, sCodIF, sMovNro, nMovNro
    Case gAhoApeTransf, gPFApeTransf, gCTSApeTransf
        AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, nSaldo, nSaldo
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo
        'Aqui debe registrarse la referencia con la transferencia de caja general
    Case Else
        AgregaMovCap nMovNro, nOperacion, sCuenta, nSaldo, nSaldo, nSaldo
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nSaldo
End Select
rsPers.MoveFirst
Do While Not rsPers.EOF
    AgregaNuevoProdPers sCuenta, rsPers("Codigo"), CLng(Trim(Right(rsPers("Relacion"), 4)))
    rsPers.MoveNext
Loop
'ARCV 14-02-2007 (Agrega promotor)
    'AgregaNuevoProdPers sCuenta, psPromotor, gCapRelPersPromotor 'FRHU 20140228 RQ14006 Se quito AgregaNuevoProdPers
'--------
CapAperturaCuenta = sCuenta


End Function

'Public Function GetSaldoFecha(ByVal sCuenta As String, ByVal dFecha As Date) As Recordset
'Dim ssql As String, sNumMov As String
'Dim rsCta As Recordset
'ssql = "Select TOP 1 CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha, " _
'    & "C.nSaldoDisponible, C.nSaldoContable FROM Mov M INNER JOIN MovCap C " _
'    & "INNER JOIN MovCapDet CD INNER JOIN OpeTpo O INNER JOIN CapMovTipo CT ON " _
'    & "O.cOpeCod = CT.cOpeCod ON CD.cOpeCod = O.cOpeCod ON C.nMovNro = CD.nMovNro ON " _
'    & "M.nMovNro = C.nMovNro WHERE C.cCtaCod = '" & sCuenta & "' AND CD.nConceptoCod " _
'    & "IN (" & gConcCapital & ") AND M.cMovNro <= '" & Format$(dFecha, "yyyymmdd") & "235959' " _
'    & "AND M.nMovFlag NOT IN (" & gMovFlagExtornado & ") ORDER BY M.nMovNro DESC"
'Set rsCta = New Recordset
'rsCta.CursorLocation = adUseClient
'rsCta.Open ssql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
'Set rsCta.ActiveConnection = Nothing
'Set GetSaldoFecha = rsCta
'Set rsCta = Nothing
'End Function
Public Function GetSaldoFecha(ByVal sCuenta As String, ByVal dFecha As Date) As ADODB.Recordset
Dim sSql As String, sNumMov As String
Dim rsCta As ADODB.Recordset

sSql = "Select TOP 1 CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha, " _
    & "C.nSaldoDisponible, C.nSaldoContable FROM Mov M INNER JOIN MovCap C ON C.NMOVNRO = M.NMOVNRO  " _
    & "INNER JOIN MovCapDet CD ON CD.NMOVNRO = C.NMOVNRO AND CD.COPECOD = C.COPECOD AND CD.CCTACOD = C.CCTACOD INNER JOIN OpeTpo O ON O.COPECOD = CD.COPECOD INNER JOIN CapMovTipo CT ON " _
    & " O.cOpeCod = CT.cOpeCod " _
    & " WHERE C.cCtaCod = '" & Trim(sCuenta) & "' AND CD.nConceptoCod " _
     & "= " & gConcCapital & " AND LEFT(M.cMovNro,8) <= '" & Format$(dFecha, "yyyymmdd") & "' " _
    & "AND M.nMovFlag =" & gMovFlagVigente & " ORDER BY M.nMovNro DESC"

Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
coConex.ConexionActiva.CommandTimeout = 1200
rsCta.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetSaldoFecha = rsCta
Set rsCta = Nothing

End Function


'******************************************************************************************
'************************************** Final de Copiado de Clases de Ahorros **********************
'******************************************************************************************
Public Sub dUpdateGarantiaBloquea(ByVal pcNumGarant As String, ByVal pnEstado As Integer)
Dim sSql As String

    sSql = " UPDATE Garantias Set  nEstado = " & pnEstado & " Where cNumGarant = '" & pcNumGarant & "'"
    coConex.Ejecutar sSql
    
End Sub
Public Sub dUpdateGarantiaLibera(ByVal pcNumGarant As String, ByVal pnEstado As Integer, ByVal pnGravado As Currency)
Dim sSql As String

    sSql = " UPDATE Garantias Set nGravament = nGravament - " & pnGravado & _
            ", nEstado = " & pnEstado & " Where cNumGarant = '" & pcNumGarant & "'"
    coConex.Ejecutar sSql
   
End Sub


Public Function GeneraMovNro(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String
    On Error GoTo GeneraMovNroErr
    Dim rs As ADODB.Recordset
    Dim SQL As String
    
    Set rs = New ADODB.Recordset
    
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       SQL = "sp_GeneraMovNro '" & Format(pdFecha & " " & coConex.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & Right(psCodAge, 2) & "','" & psUser & "'"
    Else
       SQL = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    
    Set rs = coConex.Ejecutar(SQL)
    If Not rs.EOF Then
        GeneraMovNro = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
GeneraMovNroErr:
    
End Function


Public Sub ActualizaFechaCalendarioMatriz(ByVal psCtaCod As String, ByVal pMatCal As Variant)
Dim i As Integer
Dim oCred As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim nNroCalen As Integer
    
        Set oCred = New COMDCredito.DCOMCredito
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
    Optional ByVal pdEjecutado As Date = CDate("01/01/1950"), Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pcOtrsModalidades As String = "")
'joep20181221 CP Optional ByVal pcOtrsModalidades As String = ""
Dim sSql As String

'        sSQL = "INSERT INTO ColocCartaFianza(cCtaCod, nCondicion, nModalidad, dEjecutado, cFinalidad, dAsignacion, dVencimiento) "
'        sSQL = sSQL & " VALUES('" & pcCtaCod & "'," & pnCondicion & "," & pnModalidad & ",NULL,'" & pcFinalidad & "','" & Format(pdFecAsig, "mm/dd/yyyy") & "','" & Format(psFecVenc, "mm/dd/yyyy") & "')"
'ARCV 27-10-2006
        'sSql = "INSERT INTO ColocCartaFianza(cCtaCod, nCondicion, nModalidad, cFinalidad, dAsignacion, dVencimiento) "'Comento JOEP20190208 CP
        sSql = "INSERT INTO ColocCartaFianza(cCtaCod, nCondicion, nModalidad, cFinalidad, dAsignacion, dVencimiento,cOtrsModalidades) "
        'sSql = sSql & " VALUES('" & pcCtaCod & "'," & pnCondicion & "," & pnModalidad & ",'" & pcFinalidad & "','" & Format(pdFecAsig, "mm/dd/yyyy") & "','" & Format(psFecVenc, "mm/dd/yyyy") & "')"'Comento JOEP20190208 CP
        sSql = sSql & " VALUES('" & pcCtaCod & "'," & pnCondicion & "," & pnModalidad & ",'" & pcFinalidad & "','" & Format(pdFecAsig, "mm/dd/yyyy") & "','" & Format(psFecVenc, "mm/dd/yyyy") & "','" & IIf(pcOtrsModalidades = "", Null, pcOtrsModalidades) & "')" 'JOEP20181219 CP IIf(pcOtrsModalidades = "", Null, pcOtrsModalidades)

        If pbEjecBatch Then
            coConex.AdicionaCmdBatch sSql
        Else
            coConex.Ejecutar sSql
        End If
End Sub

Public Sub dUpdateColocCartaFianza(ByVal pcCtaCod As String, Optional ByVal pnCondicion As ColocCredCondicion = -1, _
    Optional ByVal pnModalidad As ColCFModalidad = -1, _
    Optional ByVal pdFecAsig As Date = CDate("01/01/1950"), Optional ByVal pdFecVenc As Date = CDate("01/01/1950"), _
    Optional ByVal psFinalidad As String = "", _
    Optional ByVal pdEjecutado As Date = CDate("01/01/1950"), Optional ByVal pbEjecBatch As Boolean = False, Optional pnRenovacion As Integer = 0, Optional pdFecRenovacion As Date, _
    Optional ByVal pbExtorno As Boolean = False, Optional ByVal psOtrsModalidades As String = "") 'WIOR 20130319 AGREGO pbExtorno
'JOEP20181219 CP psOtrsModalidades
Dim sSql As String
        
        sSql = "UPDATE ColocCartaFianza SET "
        If pnCondicion <> -1 Then
           sSql = sSql & " nCondicion = " & pnCondicion & ","
        End If
        If pnModalidad <> -1 Then
            sSql = sSql & " nModalidad = " & pnModalidad & ","
        End If
        If pdFecAsig <> CDate("01/01/1950") Then
            sSql = sSql & " dAsignacion = '" & Format(pdFecAsig, "mm/dd/yyyy") & "',"
        End If
        If pdFecVenc <> CDate("01/01/1950") Then
            sSql = sSql & " dVencimiento = '" & Format(pdFecVenc, "mm/dd/yyyy") & "',"
        End If
        If psFinalidad <> "" Then
            sSql = sSql & " cFinalidad = '" & psFinalidad & "',"
        End If
        If pdEjecutado <> CDate("01/01/1950") Then
            sSql = sSql & " dEjecutado = '" & Format(pdEjecutado, "mm/dd/yyyy") & "',"
        End If
        If pnRenovacion <> 0 Then
            sSql = sSql & " nRenovacion = '" & pnRenovacion & "',"
        End If
        'JOEP20181219 CP
        If psOtrsModalidades <> "" Then
            sSql = sSql & "  cOtrsModalidades= '" & psOtrsModalidades & "',"
        Else
            sSql = sSql & "  cOtrsModalidades= " & "Null" & ","
        End If
        'JOEP20181219 CP
        
        'WIOR 20130319 ************************
        If pbExtorno And pnRenovacion = 0 Then
            sSql = sSql & " nRenovacion = '" & pnRenovacion & "',"
        End If
        'WIOR FIN *****************************
        If pdFecRenovacion <> CDate("01/01/1950") Then
            sSql = sSql & " dFecRenova = '" & Format(pdFecRenovacion, "mm/dd/yyyy") & "',"
        End If
        sSql = Mid(sSql, 1, Len(sSql) - 1)
        sSql = sSql & " WHERE cCtaCod = '" & pcCtaCod & "' "
        If pbEjecBatch Then
            coConex.AdicionaCmdBatch sSql
        Else
            coConex.Ejecutar sSql
        End If
End Sub

Public Sub dInsertColocCFEstado(ByVal psCtaCod As String, ByVal pdFechaPrdEstado As String, _
        ByVal pnPrdEstado As Integer, ByVal pnMonto As Currency, ByVal pdVenc As String, _
        ByVal psDescripcion As String, pnMotivoRechazo As Integer, _
        Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pbExtorno As Boolean = False, _
        Optional ByVal pdAsig As String = "01/01/1900", _
        Optional ByVal pnTipoRiesgo As Integer = 0) 'WIOR 20130319 AGREGO pbExtorno
        'FRHU20131126 AGREGO: Optional ByVal pdAsig As String = "01/01/1900"
        'Agrego Optional ByVal pnTipoRiesgo As Integer = 0 JOEP20180622 Acta 122-2018 pnTipoRiesgo
        
Dim sSql As String
        
        If Not pbExtorno Then 'WIOR 20130319
            sSql = "INSERT INTO ColocCFEstado(cCtaCod, dPrdEstado, nPrdEstado, nMonto, dVenc, cDescripcion, nMotivoRechazo, dAsig,nTpRiesgo) " 'FRHU20131126 AGREGO dAsig 'Agrego JOEP20180622 Acta 122-2018 nTpRiesgo
            sSql = sSql & " VALUES('" & psCtaCod & "','" & Format(CDate(Format(pdFechaPrdEstado, "dd/mm/yyyy") & Space(1) & Format(dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "'," & pnPrdEstado & ","
            sSql = sSql & pnMonto & ",'" & Format(CDate(Format(pdVenc, "dd/mm/yyyy")), "mm/dd/yyyy") & "','" & psDescripcion & "',"
            'sSql = sSql & pnMotivoRechazo & " )" 'FRHU20131126 - antes
            sSql = sSql & pnMotivoRechazo & ",'" 'FRHU20131126
            sSql = sSql & Format(CDate(Format(pdAsig, "dd/mm/yyyy")), "mm/dd/yyyy") & "'," 'FRHU20131126
            sSql = sSql & pnTipoRiesgo & ")" 'Agrego JOEP20180622 Acta 122-2018 pnTipoRiesgo
         'WIOR 20130319*************************************
        Else
            sSql = "INSERT INTO ColocCFEstado(cCtaCod, dPrdEstado, nPrdEstado, nMonto, dVenc, cDescripcion, nMotivoRechazo, dAsig,nTpRiesgo) " 'FRHU20131126 AGREGO dAsig 'Agrego JOEP20180622 Acta 122-2018 nTpRiesgo
            sSql = sSql & " VALUES('" & psCtaCod & "','" & Format(CDate(pdFechaPrdEstado), "mm/dd/yyyy hh:mm:ss") & "'," & pnPrdEstado & ","
            sSql = sSql & pnMonto & ",'" & Format(CDate(Format(pdVenc, "dd/mm/yyyy")), "mm/dd/yyyy") & "','" & psDescripcion & "',"
            'sSql = sSql & pnMotivoRechazo & " )"       'FRHU20131126 - antes
             sSql = sSql & pnMotivoRechazo & ",'"       'FRHU20131126
            sSql = sSql & Format(CDate(Format(pdAsig, "dd/mm/yyyy")), "mm/dd/yyyy") & "',"  'FRHU20131126
            sSql = sSql & pnTipoRiesgo & ")" 'Agrego JOEP20180622 Acta 122-2018 pnTipoRiesgo
        End If
         'WIOR FIN **************************************
         
        If pbEjecBatch Then
            coConex.AdicionaCmdBatch sSql
        Else
            coConex.Ejecutar sSql
        End If
End Sub

Public Sub dInsertColocCFRenovacion(ByVal psCtaCodA As String, ByVal psCtaCodN As String)

    Dim sSql As String
   
    sSql = "INSERT INTO ColocCFRenovacion(cCtaCodAnterior, cCtaCodNueva) " & _
           " VALUES ('" & psCtaCodA & "','" & psCtaCodN & "')"
    coConex.Ejecutar sSql

    
End Sub

Public Sub dUpdateColocCFRenovacion(ByVal psCtaCodA As String, ByVal psCtaCodN As String)

    Dim sSql As String
        
    sSql = " DELETE ColocCFRenovacion  WHERE  cCtaCodAnterior ='" & psCtaCodA & "'"
    coConex.Ejecutar sSql
     
    sSql = " INSERT INTO ColocCFRenovacion(cCtaCodAnterior, cCtaCodNueva) " & _
           " VALUES ('" & psCtaCodA & "','" & psCtaCodN & "')"
    coConex.Ejecutar sSql

    
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

Dim sSql As String

        sSql = "INSERT INTO ColocCFCredito (cCtaCod, cCtaCodCre  ) "
        sSql = sSql & " VALUES('" & psCtaCod & "','" & psCredito & "')"
        If pbEjecBatch Then
            coConex.AdicionaCmdBatch sSql
        Else
            coConex.Ejecutar sSql
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
Dim sSql As String
Dim R As ADODB.Recordset

    On Error GoTo ErrorRecuperaLineasCreditoSaldo
    sSql = "Select nMontoTotal, nSaldoCap,nMontoColocado,nMoneda, nMontoReservado from ColocLineaCreditosaldo "
    If pbBloqueo Then
        sSql = sSql & " with (tablockx) "
    End If
    sSql = sSql & " Where cLineaCred = '" & cLineaCred & "'"
    Set R = New ADODB.Recordset
    R.CursorLocation = adUseClient
    R.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
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
    Dim loIni As COMConecta.DCOMClasIni
    
    Set loIni = New COMConecta.DCOMClasIni
        csConexion = loIni.CadenaConexion
        csNegocio = loIni.BaseNegocio
        csCentralPer = loIni.BasePersonas
        csCentralCom = loIni.BaseComunes
        csCentralImg = loIni.BaseImagenes
        csAdminist = loIni.BaseAdministracion
    Set loIni = Nothing
    
Set coConex = New COMConecta.DCOMConecta
Set oError = New COMConecta.COMErrorHandling

If coConex.AbreConexion(csConexion) = False Then
    Call oError.RaiseError(oError.MyUnhandledError, "DCOMCredActBD:Initialize. Error en Conexion a Base de datos")
End If

End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub

'Ejecuta procesos Batch
Public Function dEjecutaBatch() As Integer

    dEjecutaBatch = coConex.EjecutarBatch
    Exit Function

End Function

Public Sub dEliminaCuotasPendientesCalifMiViv(ByVal psCtaCod As String)
Dim sSql As String

    sSql = " DELETE ColocCalifMiViv Where cCtaCod = '" & psCtaCod & "' AND nColocCalendEstado = " & gColocCalendEstadoPendiente
    coConex.Ejecutar sSql
    
End Sub

Public Sub dUpdateColocCalifMiViv(ByVal psCodCta As String, ByVal pnNroCalen As Integer, _
    ByVal pnCuota As Integer, _
    Optional ByVal pdPago As Date = CDate("01/01/1900"), _
    Optional ByVal psColocMiVivEval As String = "", Optional ByVal pnColocCalendEstado As Integer = -1, _
    Optional ByVal pdCondonacion As Date = CDate("01/01/1900"))

Dim sSql As String

'    sSql = "UPDATE ColocCalifMiViv SET "
'
'    If pdPago <> CDate("01/01/1900") Then
'        sSql = sSql & " dPago = '" & Format(pdPago, "mm/dd/yyyy hh:mm:ss") & "', "
'    End If
'
'    If pdCondonacion <> CDate("01/01/1900") Then
'        sSql = sSql & " dFecCond = '" & Format(pdCondonacion, "mm/dd/yyyy hh:mm:ss") & "', "
'    End If
'
'    If psColocMiVivEval <> "" Then
'        sSql = sSql & " cColocMiVivEval = '" & psColocMiVivEval & "', "
'    End If
'
'    If pnColocCalendEstado <> -1 Then
'        sSql = sSql & " nColocCalendEstado = " & pnColocCalendEstado & ", "
'    End If
'
'    sSql = Mid(sSql, 1, Len(sSql) - 2)
'    If psColocMiVivEval = "" Then
'        sSql = sSql & " WHERE  cCtaCod = '" & psCodCta & "' AND nNroCalen = " & pnNroCalen & " AND nCuota = " & pnCuota
'    Else
'        sSql = sSql & " WHERE  cCtaCod = '" & psCodCta & "' AND nCuotaOrig = " & pnCuota
'    End If
    sSql = " exec stp_upd_ColocCalifMiViv '" & psCodCta & "'," & pnNroCalen & "," & pnCuota & ","
    sSql = sSql & pnColocCalendEstado & ",'" & psColocMiVivEval & "','" & Format(pdCondonacion, "mm/dd/yyyy") & "'"
    coConex.Ejecutar sSql
    
End Sub

Public Sub dInsertColocCalifMiViv(ByVal psCodCta As String, ByVal pnCuotaOrig As Integer, ByVal pnNroCalen As Integer, _
    ByVal pnCuota As Integer, ByVal pdVenc As Date, _
    ByVal pnColocCalendEstado As Integer)

Dim sSql As String
'
'    sSql = " INSERT INTO ColocCalifMiViv(cCtaCod, nCuotaOrig, nNroCalen, nCuota, dVenc, "
'    sSql = sSql & " nColocCalendEstado) "
'    sSql = sSql & " VALUES('" & psCodCta & "'," & pnCuotaOrig & "," & pnNroCalen & "," & pnCuota & ","
'    sSql = sSql & "'" & Format(pdVenc, "mm/dd/yyyy") & "',"
'    sSql = sSql & pnColocCalendEstado & ")"
    sSql = " exec stp_ins_ColocCalifMiViv '" & psCodCta & "'," & pnCuotaOrig & "," & pnNroCalen & "," & pnCuota & ","
    sSql = sSql & "'" & Format(pdVenc, "mm/dd/yyyy") & "'," & pnColocCalendEstado
    coConex.Ejecutar sSql
    
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
        lsSQL = lsSQL & " nTasaInteres = " & Format(pnTasaInteres, "#0.00000") & ","
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
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "','" & pnPrdTasaInteres & "'," & Format(pnTasaInteres, "#0.00000") & ")"
    
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

Public Sub dUpdateColocacCred(ByVal psCtaCod As String, Optional ByVal pnDiasAtraso As Integer = -1000, Optional ByVal pnColocCondicion As Integer = -1, _
    Optional ByVal pnColocDestino As Integer = -1, Optional ByVal psProtesto As String = "", Optional ByVal pbCargoAuto As Integer = -1, _
    Optional ByVal psMetLiquidacion As String = "", Optional ByVal pbRefCapInt As Integer = -1, Optional ByVal pnNroProxCuota As Integer = -1, _
    Optional ByVal pnIntPend As Double = -1, Optional ByVal pnExoPenalidad As Integer = -1, Optional ByVal pnTpoCalend As Integer = -1, _
    Optional ByVal pnCalendDinamico As Integer = -1, Optional ByVal pnTipoDesembolso As Integer = -1, Optional ByVal pnNroCalen As Integer = -1, Optional ByVal pnNroProxDesemb As Integer = -1, _
    Optional psPersFondo As String = "", Optional pnMoneda As Integer = -1, Optional ByVal pbEjecBatch As Boolean = False, _
    Optional ByVal pnCuotaComodin As Integer = -1, Optional ByVal pnMiViv As Integer = -1, _
    Optional ByVal pnNroCalParalelo As Integer = -1, Optional ByVal pnCalifPagos As Integer = -1, _
    Optional ByVal pbPrepago As Integer = -1, Optional ByVal pnCalendDinamTipo As Integer = -1, Optional ByVal pbBloqueo As Integer = -1, _
    Optional ByVal pnVAC As Integer = -1, Optional ByVal pnAproReglamento As Integer = 0, Optional ByVal pnExoSeguroDes As Integer = 0, _
    Optional ByVal pnNumConCer As Integer = -1, Optional ByVal pnTasCosEfeAnu As Double = 0, Optional ByVal psCtaCodAho As String = "", Optional ByVal psMovNro As String = "", Optional ByVal pnIdCamp As Integer = -1, _
    Optional ByVal pnNumConMic As Integer = -1, Optional ByVal pnDiasVoucher As Integer = -1, Optional ByVal pnDatoVivienda As Integer = -1, Optional ByVal pnCuotaBalon As Integer = -1, Optional pnPreParaAmpliacion As Integer = -1, Optional lnInteresDiferido As Currency = -1, _
    Optional ByVal pnExoSegMYPE As Integer = -1) 'DAOR 20061216: Se aumentó el parametro Numero de Consultas a Certicom
    'DAOR 20070419, Se agregó parametro pnTasCosEfeAnu:Tasa Costo Efectiva Anual as
    'By Capi 10042008 se insertó parametro psCtaCodAho:Cuenta Ahorro Desembolso
    'By Capi 14042008 se insertó parametro psMovNro: Ultima fecha y usuario que hizo la reprogramacion
    'Add Gitu 06042009 se agrego parametro para la actualizacion de la campaña
    'Add Gitu 20-05-2009 se agrego el parametro de consulta Score Microfinanzas
    'AMDO 20130712, TI-ERS077-2013 Se agregó parametro pnDiasVoucher para compra de deuda
    'JUEZ 20130913 Se agregó parametro pnDatoVivienda
    'WIOR 20131111 AGRGEO pnCuotaBalon
    'ALPA 20150804 Se agregó el campo lnInteresDiferido
    'APRI20180821 ERS061-2018 agregó pnExoSegMYPE
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
    If pnDiasAtraso <> -1000 Then
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
    If pbBloqueo <> -1 Then
        lsSQL = lsSQL & " bBloqueo = " & pbBloqueo & ","
    End If
    'Add By Gitu 06-04-2009 solo para aprobacion
    If pnIdCamp <> -1 Then
        lsSQL = lsSQL & " IdCampana = '" & pnIdCamp & "',"
    End If
    'Operaciones VAC
    If pnVAC <> -1 Then
        lsSQL = lsSQL & " bVAC=" & pnVAC & ","
    End If
    
    'If pnAproReglamento <> 0 Then
        lsSQL = lsSQL & " nAproReglamento =" & pnAproReglamento & ","
    'End If
 
    'If pnExoSeguroDes <> 0 Then
        lsSQL = lsSQL & " nExoSeguroDes =" & pnExoSeguroDes & ","
    'End If
    'APRI20180821 ERS061-2018
    If pnExoSegMYPE <> -1 Then
        lsSQL = lsSQL & " nExoSegMYPE =" & pnExoSegMYPE & ","
    End If
    'END APRI
    '**DAOR 20061216, Actualizar nNumConCer: Numero de Consultas a la central de Riesgos
    If pnNumConCer <> -1 Then
        lsSQL = lsSQL & " nNumConCer = " & pnNumConCer & ","
    End If
    '************************************************************************************
    'Add By Gitu 20-05-2009 Actualiza numero de consultas Score Microfinanzas
    If pnNumConMic <> -1 Then
        lsSQL = lsSQL & " nNumConMic = " & pnNumConMic & ","
    End If
    '-------------------------------------------------
    
    '**DAOR 20070419, Actualizar nTasCosEfeAnu: Tasa Costo Efectivo Anual
    If pnTasCosEfeAnu <> 0 Then
        lsSQL = lsSQL & " nTasCosEfeAnu = " & pnTasCosEfeAnu & ","
    End If
    '************************************************************************************
    'By Capi 10042008 para que actualice la Cuenta Ahorro Desembolso cuando corresponda
    If psMovNro <> "" Then
        lsSQL = lsSQL & " cMovNroRprg = '" & psMovNro & "',"
    End If
    If psCtaCodAho <> "" Then
        lsSQL = lsSQL & " cCtaCodAho = '" & psCtaCodAho & "',"
    End If
    '************************************************************************************

    '**AMDO 20130712 Dias de voucher de Compra de deuda
    If pnColocDestino = 14 And pnDiasVoucher <> -1 Then
        lsSQL = lsSQL & " nDiasVoucher = " & pnDiasVoucher & ","
    End If
    'END AMDO ******************************************
    
    'WIOR 20131111 ******************************
    If pnCuotaBalon <> -1 Then
        lsSQL = lsSQL & " nCuotaBalon = " & pnCuotaBalon & ","
    End If
    'ALPA 20150324******************************************************
    If pnPreParaAmpliacion <> -1 Then
        lsSQL = lsSQL & " bPreparaCobroAmpliacion = " & pnPreParaAmpliacion & ","
    End If
    '*******************************************************************
    'WIOR FIN **********************************************************
    If pnDatoVivienda <> -1 Then lsSQL = lsSQL & " nDatoVivienda = " & pnDatoVivienda & "," 'JUEZ 20130913
    'ALPA 20150804************************
    'If pnInteresDiferido <> -1 Then lsSql = lsSql & " nInteresDiferido = " & pnInteresDiferido & ","
    '*************************************
    
    lsSQL = lsSQL & " cMetodoCalend = '1' ," 'LUCV20180601, Agregó según ERS022-2018
    
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

'MADM 20111018----------------------------------------------------------------
Public Sub dUpdateColocacCredMora(ByVal psCtaCod As String, Optional ByVal pnDiasAtraso As Integer = -1, Optional pbEjecBatch As Boolean = False)
    
    Dim lsSQL As String
    On Error GoTo ErrordUpdateColocacCredMora
  
    lsSQL = "UPDATE ColocacCred SET "
    
    If pnDiasAtraso <> -1 Then
        lsSQL = lsSQL & " nDiasAtraso = " & pnDiasAtraso & ","
    End If
    
    lsSQL = Mid(lsSQL, 1, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "'"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

ErrordUpdateColocacCredMora:
    Err.Raise Err.Number, "Error En Proceso dUpdateColocacCred", Err.Description

End Sub

Public Sub dUpdateColocalendDetMora(ByVal psCtaCod As String, ByVal psNroCalend As Integer, ByVal psNroCuota As Integer, Optional ByVal pnMonto As Currency = 0, Optional pnPrdConceptoCod As String = 1101, Optional pbEjecBatch As Boolean = False)
    
    Dim lsSQL As String
    On Error GoTo ErrordUpdateColocalendDetMora
  
    lsSQL = "UPDATE ColocCalendDet SET "
    
    If pnMonto <> -1 Then
        lsSQL = lsSQL & " nMonto = " & pnMonto & ","
    End If
    
    lsSQL = Mid(lsSQL, 1, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "' and nNroCalen = " & psNroCalend & " and nPrdConceptoCod = '" & pnPrdConceptoCod & "' and nCuota = " & psNroCuota & " and nColocCalendApl = 1 "
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

ErrordUpdateColocalendDetMora:
    Err.Raise Err.Number, "Error En Proceso dUpdateColocacCred", Err.Description

End Sub

Public Sub dUpdatProductoPersona(ByVal psCtaCod As String, Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pnPrdpersRelac As String = "", Optional ByVal pcPersCod As String = "")
Dim lsSQL As String
    
    On Error GoTo ErrordUpdatProductoPersona
    
    If pnPrdpersRelac <> "" Then
        lsSQL = " Update ProductoPersona set cPersCod = '" & pcPersCod & "' WHERE cCtaCod = '" & psCtaCod & "' and nPrdPersRelac =  '" & pnPrdpersRelac & "'"
    End If
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
    Exit Sub

ErrordUpdatProductoPersona:
    Err.Raise Err.Number, "Error En Proceso dDeleteProductoTasaInteres", Err.Description
End Sub
'END MADM---------------------------------------------------------------------

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
    ByVal psMovAct As String, Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal psAgeCodAct As String = "", Optional ByVal psTpoProdCod As String = "", Optional ByVal psTpoCredCod As String = "")
    'MAVM 20100604 Se agrego los par: psAgeCodAct, psTpoProdCod, psTpoCredCod BAS II

Dim lsSQL As String
    
    
    
    lsSQL = "INSERT Colocaciones (cCtaCod,nPlazo,dVenc,nMontoCol,nColocCalendCod,cUltimaActualizacion,cLineaCred,dVigencia, cAgeCodAct, cTpoProdCod, cTpoCredCod) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "',"
    lsSQL = lsSQL & Trim(Str(pnPlazo)) & ","
    lsSQL = lsSQL & "'" & Format(pdVenc, "mm/dd/yyyy") & "',"
    lsSQL = lsSQL & Trim(Str(pnMontoCol)) & ","
    lsSQL = lsSQL & "'" & Trim(Str(pnColocCalendCod)) & "',"
    lsSQL = lsSQL & "'" & psMovAct & "',"
    lsSQL = lsSQL & "'" & psLineaCred & "',"
    lsSQL = lsSQL & "'" & Format(pdVigencia, "mm/dd/yyyy") & "', "
    'MAVM 20100604 BAS II ***
    lsSQL = lsSQL & "'" & psAgeCodAct & "', "
    lsSQL = lsSQL & "'" & psTpoProdCod & "', "
    lsSQL = lsSQL & "'" & psTpoCredCod & "')"
    '***
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
            lsSQL = lsSQL & " nTransacc = nTransacc - 1,"
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
    'RECO20092016 COMENTADO
    'lsSQL = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
        & "VALUES ('" & psCuenta & "','" & psPersona & "'," & pnRelacion & ")"
    
    lsSQL = "stp_sel_InsertProductoPersonaValida '" & psCuenta & "','" & psPersona & "'," & pnRelacion
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
'EJVG20140408 ***
'Inserta MovDocRec
Public Sub dInsertaMovDocRec(ByVal pnMovNro As Long, ByVal pnTpoDoc As Integer, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psIFCta As String, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    On Error GoTo ErrordInsertaMovDocRec
    lsSQL = "INSERT INTO MovDocRec(nMovNro,nTpoDoc,cNroDoc,cPersCod,cIFTpo,cIFCta) "
    lsSQL = lsSQL & " VALUES(" & pnMovNro & "," & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psIFCta & "')"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub
ErrordInsertaMovDocRec:
    Err.Raise Err.Number, "Error En Proceso dInsertaMovDocRec", Err.Description
End Sub
'Inserta DocRecColoc
Public Sub dInsertaCuentaDocumento(ByVal pnMovNro As Long, ByVal pnTpoDoc As Integer, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As String, _
                                    ByVal psIFCta As String, ByVal psCtaCod As String, ByVal pnMonto As Currency)
    Dim lsSQL As String
    On Error GoTo ErrordInsertaCuentaDocumento
    lsSQL = "INSERT DocRecColoc (nTpoDoc,cNroDoc,cPersCod,cIFTpo,cIFCta,nMovNro,cCtaCod,nMonto) " _
            & " VALUES (" & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psIFCta & "'," & pnMovNro & ",'" & psCtaCod & "'," & pnMonto & ")"
    coConex.Ejecutar lsSQL
    Exit Sub
ErrordInsertaCuentaDocumento:
    Err.Raise Err.Number, "Error En Proceso dInsertaCuentaDocumento", Err.Description
End Sub
'END EJVG *******
'Inserta MovGarantDet
Public Sub dInsertMovGarantDet(ByVal pnMovNro As Long, ByVal psNumGarant As String, _
    ByVal pscOpeCod As String, ByVal psPrdConcepto As String, ByVal pdmonto As Double)
    
Dim lsSQL As String


lsSQL = "Insert Into MovGarantDet(nMovNro,cNumgarant, cOpeCod, nPrdConceptoCod,nMonto)"
lsSQL = lsSQL & " Values(" & pnMovNro & ",'" & psNumGarant & "','" & pscOpeCod & "'," & psPrdConcepto & "," & pdmonto & ")"
coConex.Ejecutar lsSQL
Exit Sub


End Sub

'Inserta MovGarantia
Public Sub dInsertMovGarantia(ByVal pnMovNro As Long, ByVal psNumGarant As String, _
    ByVal pscOpeCod As String, ByVal pdmonto As Double, ByVal psDescripcion As String)
    
Dim lsSQL As String

'lssql = " Select Max(nMovNro) from MovGarantia"
lsSQL = " Insert MovGarantia(nMovNro,cNumgarant, cOpeCod, nMonto, cDescripcion)"
lsSQL = lsSQL & " Values(" & pnMovNro & ",'" & psNumGarant & "','" & pscOpeCod & "' ," & pdmonto & ",'" & psDescripcion & "')"
coConex.Ejecutar lsSQL
Exit Sub


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
    
End Sub

'Inserta MovRef
Public Sub dInsertMovRef(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long, _
    Optional pbEjecBatch As Boolean = False)
    
Dim lsSQL As String
'ALPA20130827*********
Dim oRs As ADODB.Recordset
Set oRs = New ADODB.Recordset
lsSQL = "Select * From MovRef where nMovNro=" & pnMovNro & " and nMovNroRef=" & pnMovNroRef & ""
Set oRs = coConex.CargaRecordSet(lsSQL)
If Not (oRs.BOF Or oRs.EOF) Then
    Exit Sub
End If
'*********************
    
    lsSQL = "INSERT MovRef (nMovNro,nMovNroRef) " _
          & "VALUES (" & pnMovNro & "," & pnMovNroRef & ")"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

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


Public Sub dInsertMovColDesembChq(ByVal pnMovNro As Long, ByVal psBancoCod As String, _
        ByVal psNumCheque As String, ByVal psCtaBco As String)

Dim lsSQL As String

    lsSQL = "INSERT MovColDesembChq (nMovNro, cBancoCod, cNumCheque, cCta)"
    lsSQL = lsSQL & "  VALUES(" & pnMovNro & ",'" & psBancoCod & "','" & psNumCheque & "','" & psCtaBco & "')"
    
    coConex.Ejecutar lsSQL
    

End Sub

''''''''''''''
'
Public Sub dInsertaMovLavDinero(ByVal nMovNro As Long, ByVal sPersCod As String, _
                                    Optional psPersCodOrd As String = "", _
                                    Optional psPersCodRea As String = "", _
                                    Optional psPersCodBen As String = "", _
                                    Optional psPersCodVisto As String = "", Optional nTipoREU As Integer = 1, Optional nMontoAcu As Double = 0, Optional sOrigen As String = "")  'DAOR 20070511, psPersCodVisto
'By Capi 20012008 se agrego los 3 ultimos
'ALPA 20081009 se agrego los 3 ultimos
Dim lsSQL As String

'Comentado por DAOR 20070511
'lsSQL = "INSERT MovLavDinero (nMovNro,cPersCod) " _
'    & "VALUES (" & nMovNro & ",'" & sPersCod & "')"

'**DAOR 20070511
lsSQL = "INSERT MovLavDinero (nMovNro,cPersCod,cPersCodVisto,cPersCodOrd,cPersCodRea,cPersCodBen,nTipoReu,nMontoAcum,cOrigenEfectivo) " _
    & "VALUES (" & nMovNro & ",'" & sPersCod & "','" & psPersCodVisto & "','" & psPersCodOrd & "','" & psPersCodRea & "','" & psPersCodBen & "'," & nTipoREU & ", " & nMontoAcu & ",'" & sOrigen & "')"

coConex.Ejecutar lsSQL
Exit Sub



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
        ByVal psCtaCod As String, ByVal pnNroCalend As Long, ByVal pnConcepto As Long, _
        ByVal pnNroCuota As Integer, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False) 'NAGL 20190716 Agregó ByVal en pnNroCalend As Long

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
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrordFechaHora
    sSql = "Select GETDATE() as FechaHora"
    Set R = New ADODB.Recordset
    R.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
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
        Optional ByVal pnMotivRech As Integer = -1, Optional ByVal psTipoGasto As String = "F", _
        Optional ByVal pnPeriodoFechaFija2 As Integer = 0, Optional ByVal pnIncremCapital As Integer = -1, _
        Optional ByVal pbGraciaEnCuotas As Integer = -1, Optional ByVal pnTasaGracia As Double = 0, Optional ByVal pnCuotaBalon As Integer = 0)
        'WIOR 20131111 AGREGO pnCuotaBalon
Dim lsSQL As String
    
    If pnMotivRech <> -1 Then
        lsSQL = "INSERT ColocacEstado (cCtaCod, dPrdEstado, nPrdEstado, nCuotas, nMonto, cDescripcion, " _
            & "nColocCalendCod, nPeriodoFechaFija, nPeriodoGracia, nPlazo, nTipoGracia,nTipoDesembolso,nProxMes,nCalendDinamico,nMotivoRechazo, cTipoGasto,nCuotaBalon) " _
            & "VALUES ('" & psCtaCod & "','" & Format(CDate(Format(pdFechaPrdEstado, "dd/mm/yyyy") & Space(1) & Format(dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "'," & pnPrdEstado & "," & pnCuotas & "," _
            & pnMonto & ",'" & psDescripcion & "'," & pnTipoCalendario & "," & pnPeriodoFechaFija & "," _
            & pnPeriodoGracia & "," & pnPlazo & "," & pnTipoGracia & "," & pnTipoDesembolso & "," & pnProxMes & "," & pnCalendDinamico & "," & pnMotivRech & ",'" & psTipoGasto & "'," & pnCuotaBalon & ")"
            'WIOR 20131111 AGREGO pnCuotaBalon
    Else
        lsSQL = "INSERT ColocacEstado (cCtaCod, dPrdEstado, nPrdEstado, nCuotas, nMonto, cDescripcion, " _
            & "nColocCalendCod, nPeriodoFechaFija, nPeriodoGracia, nPlazo, nTipoGracia,nTipoDesembolso,nProxMes,nCalendDinamico,cTipoGasto,nPeriodoFechaFija2,bIncremGraciaCap,bGraciaEnCuotas,nTasaGracia,nCuotaBalon) " _
            & "VALUES ('" & psCtaCod & "','" & Format(CDate(Format(pdFechaPrdEstado, "dd/mm/yyyy") & Space(1) & Format(dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "'," & pnPrdEstado & "," & pnCuotas & "," _
            & pnMonto & ",'" & psDescripcion & "'," & pnTipoCalendario & "," & pnPeriodoFechaFija & "," _
            & pnPeriodoGracia & "," & pnPlazo & "," & pnTipoGracia & "," & pnTipoDesembolso & "," & pnProxMes & "," & pnCalendDinamico & ",'" & psTipoGasto & _
            "'," & IIf(pnPeriodoFechaFija2 = 0, "Null", pnPeriodoFechaFija2) & _
            "," & IIf(pnIncremCapital = -1, "Null", pnIncremCapital) & _
            "," & IIf(pbGraciaEnCuotas = -1, "Null", pbGraciaEnCuotas) & _
            "," & IIf(pnTasaGracia = 0, "Null", pnTasaGracia) & "," & pnCuotaBalon & ")" 'ARCV 29-07-2006
            'WIOR 20131111 AGREGO pnCuotaBalon
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
        Optional ByVal psTipoGastos As String = "F", Optional ByVal pnCuotaBalon As Integer = 0)
        'WIOR 20131111 AGREGO pnCuotaBalon

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
    lsSQL = lsSQL & " nCuotaBalon = " & pnCuotaBalon 'WIOR 20131111 AGREGO pnCuotaBalon
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

Public Sub dInsertColocCalendarioReprog(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
        ByVal pnColocCalendApl As ColocCalendApl, ByVal pnCuota As Integer, ByVal pdVenc As Date, _
        ByVal pdFechaPago As Date, _
        ByVal pnColocCalendEstado As Integer, ByVal psDescripcion As String, ByVal pnCalendProc As ColocCalendConceptoProc, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT ColocCalendario (cCtaCod,nNroCalen,nColocCalendApl,nCuota,dVenc,dPago,nColocCalendEstado,cDescripcion,cColocCalenFlag,nCalendProc) " _
        & "VALUES ('" & psCtaCod & "'," & pnNroCalen & "," & pnColocCalendApl & "," & pnCuota & ",'" _
        & Format(pdVenc, "mm/dd/yyyy") & "','" & Format(pdFechaPago, "mm/dd/yyyy") & "'," & pnColocCalendEstado & ",'" & psDescripcion & "',NULL," & pnCalendProc & ")"

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
    Dim sSql As String
    
    sSql = "Update MovCMAC"
    sSql = sSql & " Set nMonto=nMonto+" & pnMonto
    sSql = sSql & " Where nMovNro=" & pnMovNro
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
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
        
    'Comentado por DAOR 20080308
    ''By Capi 28122007 se incluyo el un if para que no grabe gastos con valor cero
    'If pnMonto > 0 Then
        If pbEjecBatch Then
            coConex.AdicionaCmdBatch lsSQL
        Else
            coConex.Ejecutar lsSQL
        End If
    'End If

End Sub
'ALPA 20111031*******************************************************
Public Sub dInsertaConceptosInteresCorridos(ByVal psCtaCod As String, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String

    lsSQL = "stp_sel_insertaConceptosInteresCorridos '" & psCtaCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub
'********************************************************************
'
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
Dim sSql As String
    
    sSql = "Update DocRec Set nEstado = " & pnEstCheque & " Where nTpoDoc = 47 AND cNroDoc = '" & psNumDoc & "' "
    sSql = sSql & " AND cPersCod = '" & psPersCod & "' "
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If
End Sub

Public Sub dUpdateDocRecEst(ByVal psNumDoc As String, ByVal psPersCod As String, Optional pbEjecBatch As Boolean = False, Optional pnEstCheque As ChequeEstado = gChqEstValorizado)
Dim sSql As String
    
    sSql = "Update DocRecEst Set nEstado = " & pnEstCheque & " Where nTpoDoc = 47 AND cNroDoc = '" & psNumDoc & "' "
    sSql = sSql & " AND cPersCod = '" & psPersCod & "' "
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
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
    Optional ByVal bIncrementarMontoCol As Boolean = False, Optional ByVal psUltimaActualizacion As String = "", _
    Optional ByVal psAgeCodAct As String = "", Optional ByVal psTpoProdCod As String = "", Optional ByVal psTpoCredCod As String = "", Optional ByVal psTipoInstCod As String = "", Optional ByVal pnMontoMivivienda As Currency = -1, Optional ByVal pnCuotaMivienda As Integer = -1, _
    Optional pnTasaExononeracion As Double = -1, Optional pbPreferencial As Integer = -1, Optional ByVal pnRatioCobertura As Double = -1)
    'MAVM 20100604 Se agrego los par: psAgeCodAct, psTpoProdCod, psTpoCredCod BAS II

Dim lsSQL As String

    On Error GoTo ErrordUpdateColocaciones
    lsSQL = "UPDATE Colocaciones SET "
    If pnPlazo <> -1 Then
        lsSQL = lsSQL & " nPlazo = " & pnPlazo & ","
    End If
    If pdVenc <> CDate("01/01/1950") Then
        lsSQL = lsSQL & " dVenc = '" & Format(pdVenc, "mm/dd/yyyy") & "',"
    End If
    'ALPA 20140611************************
    If pnMontoMivivienda <> -1 Then
        lsSQL = lsSQL & " nMontoMiVivienda = " & pnMontoMivivienda & ","
    End If
    '************************************************************************
    'ALPA 20141030***********************************************************
    If pnTasaExononeracion <> -1 Then
        lsSQL = lsSQL & " nTasaExononeracion = " & pnTasaExononeracion & ","
        If pnTasaExononeracion = 0 Then
            lsSQL = lsSQL & " bExononeracionTasa = 0,"
        Else
            lsSQL = lsSQL & " bExononeracionTasa = 1,"
        End If
    End If
    If pbPreferencial <> -1 Then
        lsSQL = lsSQL & " bPreferencial = " & pbPreferencial & ","
    End If
    '*************************************************************************
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
    'ALPA 20141127*******************
    If pnCuotaMivienda <> -1 Then
        lsSQL = lsSQL & " nCuotaPolizaMivivienda = " & pnCuotaMivienda & ","
    End If
    '********************************
    If psLineaCred <> "" Then
        lsSQL = lsSQL & " cLineaCred = '" & psLineaCred & "', "
    End If
    If psUltimaActualizacion <> "" Then
        lsSQL = lsSQL & " cUltimaActualizacion = '" & psUltimaActualizacion & "',"
    End If
    If pdVigencia <> CDate("01/01/1950") Then
        lsSQL = lsSQL & " dVigencia = '" & Format(pdVigencia, "mm/dd/yyyy") & "',"
    End If
    
    'EJVG20150819 ***
    If pnRatioCobertura <> -1 Then
        lsSQL = lsSQL & " nRatioCobertura = " & pnRatioCobertura & ","
    End If
    'END EJVG *******
    
    'MAVM 20100604 BAS II
    If psAgeCodAct <> "" Then
        lsSQL = lsSQL & " cAgeCodAct = '" & psAgeCodAct & "',"
    End If
    If psTpoProdCod <> "" Then
        lsSQL = lsSQL & " cTpoProdCod = '" & psTpoProdCod & "',"
    End If
    If psTpoCredCod <> "" Then
        lsSQL = lsSQL & " cTpoCredCod = '" & psTpoCredCod & "',"
        'CTI3 ERS0032020
            lsSQL = lsSQL & " cTpoCredCodPadre = '" & Left(psTpoCredCod, 2) & "0" & "',"

        '******************************************
    End If
    If psTipoInstCod <> "" Then
         If Mid(psTpoCredCod, 1, 1) = Mid(gColCredCorpo, 1, 1) Then
            lsSQL = lsSQL & " nTpoInstCorp = '" & psTipoInstCod & "' "
         Else
            lsSQL = lsSQL & " nTpoInstCorp = null "
         End If
    End If
    '***
    
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
Dim sSql As String
Dim R As ADODB.Recordset

    sSql = "Select nMovNro From MovCol Where nNroCalen = " & pnNroDacion & " And cCtaCod = '" & psCtaCod & "'"
    Set R = New ADODB.Recordset
    R.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
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
'EJVG20150712 Comentar este método
'Public Sub dDeleteColocGarantia(ByVal psNumGarant As String, _
'        ByVal psCtaCod As String, Optional pbEjecBatch As Boolean = False)
'Dim lsSQL As String
'
'    lsSQL = "DELETE ColocGarantia WHERE cNumGarant = '" & psNumGarant & "' "
'    lsSQL = lsSQL & " AND cCtaCod = '" & psCtaCod & "' "
'    If pbEjecBatch Then
'        coConex.AdicionaCmdBatch lsSQL
'    Else
'        coConex.Ejecutar lsSQL
'    End If
'End Sub

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
    Optional pbEjecBatch As Boolean = False, Optional ByVal psCodInstitucion As String = "", _
    Optional ByVal pnIdCampana As Integer = 0)
Dim lsSQL As String
    On Error GoTo errorInsert
    lsSQL = "INSERT INTO ProductoConceptoFiltro (nPrdConceptoCod, nProdCod, cAgeCod,cTpoFiltro,cIntitucion, IDCampana)"
    lsSQL = lsSQL & " VALUES(" & pnPrdConceptoCod & "," & pnProdCod & ",'"
    'lsSQL = lsSQL & psAgeCod & "','" & psTipoFiltro & "'," & IIf(psCodInstitucion = "", Null, "'" & psCodInstitucion & "'") & ")"
    lsSQL = lsSQL & psAgeCod & "','" & psTipoFiltro & "','" & psCodInstitucion & "'," & pnIdCampana & ")"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
Exit Sub
errorInsert:
    MsgBox Err.Description, vbInformation, Err.Source
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

'Copia Fiel de la Clase NCOMCredOpe para poder realizar el pago en Lote en Una Sola Transaccion

Public Function AmortizarCreditoLote(ByVal psCtaCod, ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant, _
            ByVal pnMonto As Double, ByVal pdHoy As Date, ByVal psMetLiquid As String, _
            ByVal pnTipoPago As ColocTipoPago, ByVal psCodAge As String, ByVal psCodUser As String, _
            ByVal nMovNro As Long, ByVal pnPrdEstado As Integer, ByVal pnTransacc As Long, Optional psNroDoc As String = "", Optional pnITF As Double = 0#, _
            Optional ByVal psProyectoActual As String = "", _
            Optional ByVal nNroCalen As Integer, Optional ByVal pnPlazo As Integer, Optional ByVal pnCapitalPagado As Double = 0, _
            Optional ByVal psPersCod As String = "", Optional ByVal psIFTpo As String = "01", Optional ByVal psIFCta As String = "", Optional ByVal pnPorcentaje As Currency, Optional ByVal pbCobrarMora As Boolean = True) As String
            
'ARCV 22-06-2007 :Optional ByVal nNroCalen As Integer, Optional ByVal pnPlazo As Integer
'JUEZ 20131022 Se agregó parámetro pnCapitalPagado
'EJVG20140226 Se agregó psPersCod, psIFTpo, psIFCta
'NAGL 20190719 Según ERS036-2018 Agregó pbCobrarMora

Dim nEstadoCred As Integer
Dim R As ADODB.Recordset
Dim oCred As COMDCredito.DCOMCredito
Dim oCalend As COMDCredito.DCOMCalendario
Dim nTransacc As Long
Dim sLineaCred As String
Dim nMontoColocado As Double
Dim dFecPend As Date
Dim nDiasAtraso As Integer
Dim i As Integer
Dim k As Integer 'ADD CTI2 20190423
Dim nMontoGasto As Double
Dim nConsCred As String
'ARCV 22-06-2007
'Dim nNroCalen As Integer
'Dim pnPlazo As Integer
'-------
Dim bTran As Boolean
Dim dFechaTran As Date
'Dim oNegCred As NCOMCredOpe
Dim nPrestamo As Double 'JUEZ 20130429
Dim bReprogRoyaAmarilla As Boolean 'JUEZ 20131022
Dim nMovNroReversionReprog As Long 'JUEZ 20131022
Dim bReprogCovid19 As Boolean 'APRI20200430 POR COVID-19
Dim nMovNroReversionCovid19 As Long 'APRI20200430 POR COVID-19
'WIOR 20150421 ***
Dim nIntDiferido As Currency
Dim nRevIntDif As Currency
Dim nRevIntDifAcum As Currency
Dim sOpeRevIntDif As String
Dim nIntDiferidoPend As Currency
nRevIntDifAcum = 0
sOpeRevIntDif = ""
nIntDiferidoPend = 0
'WIOR FIN ********
'ALPA 20150708********************************
Dim nDatoVivienda As Integer
Dim sTpoCredCod As String
Dim dFecVig As Date
Dim CapitalPagado As Double
'*********************************************

'CTI2 20190423 INI
Dim MatGastosCred As Variant
Dim MatGastosCuota As Variant
Dim NumregGastosCred As Integer
Dim NumRegGastosCuota As Integer
'CTI2 END

'*********NAGL 20190719************
Dim lrsOperCond As ADODB.Recordset
Dim psOpeCodCond As String
'***********END NAGL***************

Dim lbRefVigente As Integer 'ALPA 20150803
    On Error GoTo ErrorAmortizarCreditoLote
            
    nEstadoCred = pnPrdEstado
    nTransacc = pnTransacc
    
    'Definir Codigo de Operacion de Pago
    Select Case nEstadoCred
            'Si es Credito refinanciado
            Case gColocEstRefMor
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefMorEfec, gCredPagRefMorCh)
                sOpeRevIntDif = gCredRevIntDifRef 'WIOR 20150421
            Case gColocEstRefNorm
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefNorEfec, gCredPagRefNorCh)
                sOpeRevIntDif = gCredRevIntDifRef 'WIOR 20150421
            Case gColocEstRefVenc
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefVenEfec, gCredPagRefVenCh)
                sOpeRevIntDif = gCredRevIntDifRef 'WIOR 20150421
            'si es Credito Normal
            Case gColocEstVigMor
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorNorEfec, gCredPagNorNorCh)
                sOpeRevIntDif = gCredRevIntDifVenc 'WIOR 20150421
            Case gColocEstVigNorm
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorMorEfec, gCredPagNorMorCh)
                sOpeRevIntDif = gCredRevIntDifVig 'WIOR 20150421
            Case gColocEstVigVenc
                nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorVenEfec, gCredPagNorVenCh)
                sOpeRevIntDif = gCredRevIntDifVenc 'WIOR 20150421
    End Select
    
    
    'Set oNegCred = New NCOMCredOpe
    
'ARCV 22-06-2007 : Ya no hacer busqueda...solo leerlos de las variables
'    Set oCred = New COMDCredito.DCOMCredito
'    Set R = oCred.RecuperaColocacEstado(psCtaCod, gColocEstAprob)
'    Set oCred = Nothing
'    pnPlazo = IIf(IsNull(R!nPlazo), 0, R!nPlazo)
'    R.Close
'    Set R = Nothing
'
    Set oCred = New COMDCredito.DCOMCredito
    Set R = oCred.RecuperaColocacCred(psCtaCod)
    Set oCred = Nothing
'    nNroCalen = R!nNroCalen
    nDatoVivienda = IIf(IsNull(R!nDatoVivienda), 0, R!nDatoVivienda) 'ALPA 20150529
    lbRefVigente = IIf(IsNull(R!bRefVigente), 0, R!bRefVigente) 'ALPA 20150803*****
    R.Close
    Set R = Nothing
'---------------------
    'JUEZ 20130429 ****************************************
    Set oCred = New COMDCredito.DCOMCredito
    Set R = oCred.RecuperaColocaciones(psCtaCod)
    nPrestamo = R!nMontoCol
    sTpoCredCod = R!cTpoCredCod 'ALPA 20150529
    dFecVig = R!dVigencia 'ALPA 20150529
    Set oCred = Nothing
    'END JUEZ *********************************************
    'JUEZ 20131022 ****************************************
    Set oCred = New COMDCredito.DCOMCredito
    bReprogRoyaAmarilla = False
    Set R = oCred.ExisteReprogramacionxDesastreNat(psCtaCod, gCredReprogRoyaAmarilla)
    If R.RecordCount > 0 Then
        bReprogRoyaAmarilla = True
        nMovNroReversionReprog = R!nMovNro
    End If
    Set oCred = Nothing
    'END JUEZ *********************************************
    'APRI20200430 POR COVID-19
    Set oCred = New COMDCredito.DCOMCredito
    Set R = oCred.ExisteReprogramacionxDesastreNat(psCtaCod, "100929")
    If R.RecordCount > 0 Then
        bReprogCovid19 = True
        nMovNroReversionCovid19 = R!nMovNro
    End If
    Set oCred = Nothing
    'END APRI
    'WIOR 20150421 ***
    Set oCred = New COMDCredito.DCOMCredito
    Set R = oCred.ObtenerInteresDiferidoCredito(psCtaCod)
    If R.RecordCount > 0 Then
        nIntDiferido = IIf(IsNull(R!nInteresDif), 0, R!nInteresDif)
        
        If nIntDiferido > 0 Then
            Set R = oCred.ObtenerInteresDiferidoPendCredito(psCtaCod)
            If R.RecordCount > 0 Then
                nIntDiferidoPend = IIf(IsNull(R!nIntDifPend), 0, R!nIntDifPend)
            End If
        End If
    End If
    Set oCred = Nothing
    'WIOR FIN ********
    
    '***********NAGL Según ERS036-2018***********
    Set lrsOperCond = New ADODB.Recordset
    Set oCred = New COMDCredito.DCOMCredito
    Set lrsOperCond = oCred.GetOpeCondonacionxEstado(psCtaCod, "cOpeCond")
    psOpeCodCond = lrsOperCond("cOpeCond")
    Set oCred = Nothing
    Set lrsOperCond = Nothing
    '**************END NAGL 20190722*************
    
    'dFechaTran = CDate(Format(pdHoy, "dd/mm/yyyy hh:mm:ss"))
    dFechaTran = CDate(Format(Format(pdHoy, "dd/mm/yyyy") & " " & Format(dFechaHora, "hh:mm:ss"), "dd/mm/yyyy hh:mm:ss")) 'JUEZ 20150410
    
    'Actualiza Producto
    'If oNegCred.MatrizEstadoCalendario(MatCalendDistrib) = gColocCalendEstadoPagado Then
    If MatrizEstadoCalendario(MatCalendDistrib) = gColocCalendEstadoPagado Then
        nEstadoCred = gColocEstCancelado
    End If
    '**Modificado por DAOR 20080115
    'Call dUpdateProducto(psCtaCod, , MatrizSaldoCapital(MatCalend, MatCalendDistrib), nEstadoCred, pdHoy, nTransacc + 1, False)
    'Call dUpdateProducto(psCtaCod, , MatrizSaldoCapital(MatCalend, MatCalendDistrib), nEstadoCred, Format(dFechaTran, "mm/dd/yyyy"), nTransacc + 1, False)
    Call dUpdateProducto(psCtaCod, , MatrizSaldoCapital(MatCalend, MatCalendDistrib), nEstadoCred, pdHoy, nTransacc + 1, False) 'JUEZ 20150410
    
    
    '05-05-2005
    'Actualiza ColocLineaCreditoSaldo
    'Set oCred = New COMDCredito.DCOMCredito
    'Set R = oCred.RecuperaColocaciones(psCtaCod)
    'Set oCred = Nothing
    'sLineaCred = R!cLineaCred
    'R.Close
    'Set R = Nothing
    'Set R = RecuperaLineasCreditoSaldo(sLineaCred, True)
    'nMontoColocado = R!nMontoColocado
    'R.Close
    'Set R = Nothing
    'nMontoColocado = nMontoColocado - pnMonto
    'Call dUpdateLineaCreditoSaldo(sLineaCred, , , nMontoColocado, , , False)
    '********************************************
    
    'Actualiza ColocacCred
    dFecPend = MatrizFechaCuotaPendiente(MatCalend, MatCalendDistrib)
    If MatrizCuotaPendiente(MatCalend, MatCalendDistrib) = 0 Then
        nDiasAtraso = 0
    Else
        nDiasAtraso = pdHoy - dFecPend
    End If
    
    'Diferencia proyectos - Para Huancayo
    If nDiasAtraso > 0 And psProyectoActual = "H" Then
        nDiasAtraso = 0
    End If
   
    
    Call dUpdateColocacCred(psCtaCod, nDiasAtraso, , , , , , , MatrizCuotaPendiente(MatCalend, MatCalendDistrib), , , , , , , , , , False)
        
    'EJVG20140408 ***
    If pnTipoPago = gColocTipoPagoCheque Then
        dInsertaCuentaDocumento nMovNro, TpoDocCheque, psNroDoc, psPersCod, psIFTpo, psIFCta, psCtaCod, pnMonto + pnITF
    End If
    'END EJVG *******
    'Call dInsertMovCol(nMovNro, nConsCred, psCtaCod, nNroCalen, pnMonto, 0, psMetLiquid, pnPlazo, False, pnPrdEstado)
    Call dInsertMovCol(nMovNro, nConsCred, psCtaCod, nNroCalen, pnMonto, 0, psMetLiquid, pnPlazo, MatrizSaldoCapital(MatCalend, MatCalendDistrib), pnPrdEstado) 'EJVG20130121
    
    If pbCobrarMora = False And (MatrizSaldoIntMoratCondonado(MatCalend, MatCalendDistrib) > 0) Then
       Call dInsertMovCol(nMovNro, psOpeCodCond, psCtaCod, nNroCalen, MatrizSaldoIntMoratCondonado(MatCalend, MatCalendDistrib), 0, psMetLiquid, pnPlazo, 0, pnPrdEstado)
    End If
    
    'ALPA 20150730******************
    If pnPorcentaje > 0# And lbRefVigente = 1 Then
        Call dInsertMovCol(nMovNro, "101109", psCtaCod, nNroCalen, Round(MatrizCapitalPagado(MatCalendDistrib) * pnPorcentaje, 2), 0, psMetLiquid, pnPlazo, MatrizSaldoCapital(MatCalend, MatCalendDistrib), pnPrdEstado, False)
        Call dInsertMovCol(nMovNro, "101110", psCtaCod, nNroCalen, Round(MatrizCapitalPagado(MatCalendDistrib), 2), 0, psMetLiquid, pnPlazo, MatrizSaldoCapital(MatCalend, MatCalendDistrib), pnPrdEstado, False)
    End If
    '*******************************
    'ITF
    Call dInsertMovCol(nMovNro, IIf(pnTipoPago = gColocTipoPagoCheque, "990107", gCredITF), psCtaCod, CLng(nNroCalen), pnITF, 0, psMetLiquid, pnPlazo, False, pnPrdEstado)
    Call dInsertMovColDet(nMovNro, IIf(pnTipoPago = gColocTipoPagoCheque, "990107", gCredITF), psCtaCod, CLng(nNroCalen), gConcITFCliente, 0, pnITF, False)
    
    'CTI2 20190423 - Carga Gastos en Memoria para Evitar Bloqueo
    MatGastosCred = DevuelveMatrizGastosCredito(NumregGastosCred, psCtaCod, nNroCalen)
    
    'Actualiza calendario (ColocCalendario y ColocCalendDet)
    For i = 0 To UBound(MatCalendDistrib) - 1
    
        If CDbl(MatCalendDistrib(i, 3)) > 0 Or CDbl(MatCalendDistrib(i, 4)) > 0 Or CDbl(MatCalendDistrib(i, 5)) > 0 Or _
          CDbl(MatCalendDistrib(i, 6)) > 0 Or CDbl(MatCalendDistrib(i, 7)) > 0 Or CDbl(MatCalendDistrib(i, 8)) > 0 Or _
          CDbl(MatCalendDistrib(i, 11)) > 0 Or CDbl(MatCalendDistrib(i, 9)) > 0 Then
              'Call dUpdateColocCalendario(psCtaCod, nNroCalen, CInt(MatCalendDistrib(i, 1)), gColocCalendAplCuota, , CInt(MatCalendDistrib(i, 2)), "Pago de Cuota", gColocCalendConceptoProcAprobado, False, , Format(dFechaTran, "mm/dd/yyyy"))
                  Call dUpdateColocCalendario(psCtaCod, nNroCalen, CInt(MatCalendDistrib(i, 1)), gColocCalendAplCuota, , CInt(MatCalendDistrib(i, 2)), "Pago de Cuota", gColocCalendConceptoProcAprobado, False, , dFechaTran) 'JUEZ 20150410
        Else
              Call dUpdateColocCalendario(psCtaCod, nNroCalen, CInt(MatCalendDistrib(i, 1)), gColocCalendAplCuota, , CInt(MatCalendDistrib(i, 2)), "Pago de Cuota", gColocCalendConceptoProcAprobado, False)
        End If
        
        'Amortizando Capital
        If CDbl(MatCalendDistrib(i, 3)) > 0 Then
            Call dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), gColocConceptoCodCapital, , CDbl(MatCalendDistrib(i, 3)), , False, True)
            'Inserta Detalle Movimiento Capital
            Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodCapital, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 3)), False)
            'ALPA 20150730************************
            If pnPorcentaje > 0# And lbRefVigente = 1 Then
                Call dInsertMovColDet(nMovNro, "101109", psCtaCod, CLng(nNroCalen), gColocRFAInteresDiferido, CInt(MatCalendDistrib(i, 1)), Round(CDbl(MatCalendDistrib(i, 3)) * pnPorcentaje, 2), False)
                Call dInsertMovColDet(nMovNro, "101110", psCtaCod, CLng(nNroCalen), 1000, CInt(MatCalendDistrib(i, 1)), Round(CDbl(MatCalendDistrib(i, 3)), 2), False)
            End If
            '*************************************
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
            
            '************NAGL 20190717 ERS036-2019***********
            If pbCobrarMora = True Then
            'Inserta Detalle Movimiento Interes Moratorio
                Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresMoratorio, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 6)), False)
            Else
                'Call dInsertMovCol(nMovNro, psOpeCodCond, psCtaCod, nNroCalen, CDbl(MatCalendDistrib(i, 6)), 0, psMetLiquid, pnPlazo, 0, pnPrdEstado)
                'Comentado by NAGL 20191231 y trasladado a la parte Superior
                Call dInsertMovColDet(nMovNro, psOpeCodCond, psCtaCod, CLng(nNroCalen), gColRecConceptoCodInteresMoratorio, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 6)), False)
            End If
            '**************************************************
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
            Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresSuspenso, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 8)), False) 'Juez 20120716
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
            
            'CTI2 20190423 Comentado para alinear distribución de gastos según la opción de pago nomrmal.
            'Set oCalend = New COMDCredito.DCOMCalendario
            'Set R = oCalend.RecuperaCalendarioGastos(psCtaCod, nNroCalen, CInt(MatCalendDistrib(i, 1)), gColocCalendAplCuota)
            'Set oCalend = Nothing
            
            'CTI2 20190423 COMENTADO
            'Do While Not R.EOF
            '    If nMontoGasto >= R!nMonto Then
            '        Call dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), R!nPrdConceptoCod, , R!nMonto, , False, True)
            '        'Inserta Detalle Movimiento Gastos
            '        Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), R!nPrdConceptoCod, CInt(MatCalendDistrib(i, 1)), R!nMonto, False)
            '        nMontoGasto = nMontoGasto - R!nMonto
            '    Else
            '        Call dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), R!nPrdConceptoCod, , nMontoGasto, , False, True)
            '        'Inserta Detalle Movimiento Gastos
            '        Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), R!nPrdConceptoCod, CInt(MatCalendDistrib(i, 1)), CDbl(Format(nMontoGasto, "#0.00")), False)
            '        nMontoGasto = 0
            '    End If
            '
            '    nMontoGasto = CDbl(Format(nMontoGasto, "#0.00"))
            '    If nMontoGasto = 0 Then
            '        Exit Do
            '    End If
            '    R.MoveNext
            'Loop
            'R.Close
            'END CTI2
            
            MatGastosCuota = DevuelveMatrizGastosCreditoCuota(NumRegGastosCuota, CInt(MatCalendDistrib(i, 1)), MatGastosCred, NumregGastosCred)
            
            For k = 0 To NumRegGastosCuota - 1
                If nMontoGasto >= CDbl(MatGastosCuota(k, 2)) Then
                    Call dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), CLng(MatGastosCuota(k, 1)), , CDbl(MatGastosCuota(k, 2)), , False, True)
                    'Inserta Detalle Movimiento Gastos
                    Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), CLng(MatGastosCuota(k, 1)), CInt(MatCalendDistrib(i, 1)), CDbl(MatGastosCuota(k, 2)), False)
                    nMontoGasto = nMontoGasto - CDbl(MatGastosCuota(k, 2))
                Else
                    Call dUpdateColocCalendDet(psCtaCod, nNroCalen, gColocCalendAplCuota, CInt(MatCalendDistrib(i, 1)), CLng(MatGastosCuota(k, 1)), , nMontoGasto, , False, True)
                    'Inserta Detalle Movimiento Gastos
                    Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), CLng(MatGastosCuota(k, 1)), CInt(MatCalendDistrib(i, 1)), CDbl(Format(nMontoGasto, "#0.00")), False)
                    nMontoGasto = 0
                End If

                nMontoGasto = CDbl(Format(nMontoGasto, "#0.00"))
                If nMontoGasto = 0 Then
                    Exit For
                End If
            Next k

            Set R = Nothing
        End If
        
        'WIOR 20150421 ***
        If nIntDiferido > 0 And nIntDiferidoPend > 0 Then
            If CDbl(MatCalendDistrib(i, 3)) > 0 Then
                nRevIntDif = CCur(Round((CDbl(MatCalendDistrib(i, 3)) / nPrestamo) * nIntDiferido, 2))
                nIntDiferidoPend = nIntDiferidoPend - nRevIntDif
                
                If nIntDiferidoPend < 0 Then
                    nRevIntDif = nRevIntDif + nIntDiferidoPend
                End If
                
                Call dInsertMovColDet(nMovNro, sOpeRevIntDif, psCtaCod, CLng(nNroCalen), gColocConceptoCodInteresDiferidoAmp, CInt(MatCalendDistrib(i, 1)), nRevIntDif, False)
                nRevIntDifAcum = nRevIntDifAcum + nRevIntDif
                nRevIntDif = 0
            End If
        End If
        'WIOR FIN ********
    Next i
    
    'WIOR 20150421 ***
    If nIntDiferido > 0 Then
        If MatrizCapitalPagado(MatCalendDistrib) > 0 Then
            If nRevIntDifAcum > 0 Then
                Call dInsertMovCol(nMovNro, sOpeRevIntDif, psCtaCod, CLng(nNroCalen), nRevIntDifAcum, 0, psMetLiquid, pnPlazo, False, pnPrdEstado)
            End If
        End If
        nRevIntDifAcum = 0
    End If
    'WIOR FIN ********
    
    'JUEZ 20130429 ****************************************
    If nEstadoCred = gColocEstCancelado Then
        Set oCred = New COMDCredito.DCOMCredito
        Set R = oCred.RecuperaGarantes(psCtaCod)
        If Not R.EOF Then
            Call dInsertMovCol(nMovNro, gReversionGarantAvales, psCtaCod, 0, nPrestamo, 0, "", 0, 0, 0, False)
            Call dInsertMovColDet(nMovNro, gReversionGarantAvales, psCtaCod, 0, 1000, 0, nPrestamo, False)
        End If
        Set R = Nothing
    End If
    'END JUEZ ********************************************
    'JUEZ 20131022 ***************************************
    If bReprogRoyaAmarilla Then
        Call dInsertMovCol(nMovNro, gReversionReprogRoyaAmarilla, psCtaCod, 0, pnCapitalPagado, 0, "", 0, 0, 0, False)
        Call dInsertMovColDet(nMovNro, gReversionReprogRoyaAmarilla, psCtaCod, 0, gColocConceptoCodCapitalReprogRoyaAmarilla, 0, pnCapitalPagado, False)
        If nEstadoCred = gColocEstCancelado Then
            Call dUpdateCredReprogNatEspeciales(nMovNroReversionReprog, 1)
        End If
    End If
    'END JUEZ ********************************************
    'APRI20200415 POR COVID-19 *********************************
    If bReprogCovid19 Then
        Call dInsertMovCol(nMovNro, "107337", psCtaCod, 0, pnCapitalPagado, 0, "", 0, 0, 0, False)
        Call dInsertMovColDet(nMovNro, "107337", psCtaCod, 0, 1080, 0, pnCapitalPagado, False)
        If nEstadoCred = gColocEstCancelado Then
            Call dUpdateCredReprogNatEspeciales(nMovNroReversionCovid19, 1)
        End If
    End If
    'APRI END***************************************
    'ALPA 20150529 ******************************************************************
        CapitalPagado = 0
        CapitalPagado = CapitalPagado + MatrizCapitalPagado(MatCalendDistrib)
        If Left(sTpoCredCod, 1) = "8" And DateDiff("D", dFecVig, "01/01/2013") <= 0 Then
            If nDatoVivienda = 1 Then
                Call dInsertMovCol(nMovNro, gReversionCapHipoPrimeraVivienda, psCtaCod, 0, CapitalPagado, 0, "", 0, 0, 0)
                Call dInsertMovColDet(nMovNro, gReversionCapHipoPrimeraVivienda, psCtaCod, 0, gColocConceptoCodCapital, 0, CapitalPagado, False)
            ElseIf nDatoVivienda = 2 Then
                Call dInsertMovCol(nMovNro, gReversionCapHipoNoPrimeraVivienda, psCtaCod, 0, CapitalPagado, 0, "", 0, 0, 0)
                Call dInsertMovColDet(nMovNro, gReversionCapHipoNoPrimeraVivienda, psCtaCod, 0, gColocConceptoCodCapital, 0, CapitalPagado, False)
            End If
        End If
    'END ALPA ***********************************************************************
    
    Call dLiberarGarantia(nMovNro, nConsCred, psCtaCod, CapitalPagado, nEstadoCred)  'EJVG20150820
    
    'WIOR 20150914 ***
    Call RegularPreparacionParaAmpliado(psCtaCod, pdHoy)
    'WIOR FIN ********
    
    Exit Function

ErrorAmortizarCreditoLote:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

'**********Para evitar el Cruce de Componentes N con Componentes D
Public Function MatrizEstadoCalendario(ByVal MatCalendDistrib As Variant) As Integer
Dim i As Integer
    MatrizEstadoCalendario = gColocCalendEstadoPagado
    For i = 0 To UBound(MatCalendDistrib) - 1
        If CInt(MatCalendDistrib(i, 2)) = gColocCalendEstadoPendiente Then
            MatrizEstadoCalendario = gColocCalendEstadoPendiente
            Exit For
        End If
    Next i
End Function

Public Function MatrizSaldoCapital(ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant) As Double
Dim i As Integer
    MatrizSaldoCapital = 0
    For i = 0 To UBound(MatCalend) - 1
        MatrizSaldoCapital = MatrizSaldoCapital + (CDbl(MatCalend(i, 3)) - CDbl(MatCalendDistrib(i, 3)))
        MatrizSaldoCapital = CDbl(Format(MatrizSaldoCapital, "#0.00"))
    Next i
End Function

Public Function MatrizSaldoIntMoratCondonado(ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant) As Double
Dim i As Integer
    MatrizSaldoIntMoratCondonado = 0
    For i = 0 To UBound(MatCalend) - 1
        MatrizSaldoIntMoratCondonado = MatrizSaldoIntMoratCondonado + CDbl(MatCalendDistrib(i, 6))
        MatrizSaldoIntMoratCondonado = CDbl(Format(MatrizSaldoIntMoratCondonado, "#0.00"))
    Next i
End Function 'Creado by NAGL 20191231

Public Function MatrizFechaCuotaPendiente(ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant) As Date
Dim i As Integer
    For i = 0 To UBound(MatCalend) - 1
        If CInt(MatCalendDistrib(i, 2)) = gColocCalendEstadoPendiente Then
            MatrizFechaCuotaPendiente = CDate(MatCalendDistrib(i, 0))
            Exit For
        End If
    Next i
End Function

Public Function MatrizCuotaPendiente(ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant) As Integer
Dim i As Integer
    MatrizCuotaPendiente = 0
    For i = 0 To UBound(MatCalend) - 1
        If CInt(MatCalendDistrib(i, 2)) = gColocCalendEstadoPendiente Then
            MatrizCuotaPendiente = CInt(MatCalendDistrib(i, 1))
            Exit For
        End If
    Next i
End Function
'*****************************************************************

Public Sub dExtornaSaldosCalendario(ByVal pnMovNro As Long, _
    ByVal nAplicado As ColocCalendApl, ByVal psCtaCod As String, Optional pbEjecBatch As Boolean = False, Optional psOpeCodCond As String = "")
Dim sSql As String
    
    'JIPR20201124 Mejora Extorno
'    sSql = "UPDATE ColocCalendario SET nColocCalendEstado = " & gColocCalendEstadoPendiente & ", dPago = NULL "
'    sSql = sSql & " From ColocCalendario CC Inner Join MovColDet MC ON CC.cCtaCod = MC.cCtacod "
'    sSql = sSql & "     And CC.nNroCalen = MC.nNroCalen "
'    sSql = sSql & "     And CC.nCuota = MC.nNroCuota "
'    sSql = sSql & "     Where MC.nMovNro = " & pnMovNro & " And CC.nColocCalendApl = " & nAplicado
'    sSql = sSql & "     AND CC.cCtaCod = '" & psCtaCod & "'"
    
    sSql = "EXEC stp_upd_ExtornarCuotaCalend " & pnMovNro & ",'" & psCtaCod & "'," & nAplicado & ""
      If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If
       
    'ALPA20130713***********
    
    If Mid(psCtaCod, 6, 3) = "801" Then
        sSql = " Update CC"
        sSql = sSql & "    set dPago=null,"
        sSql = sSql & "       nColocCalendEstado = 0"
        sSql = sSql & " from MovMiVivienda MV"
        sSql = sSql & "    inner join ColocCalendario CC on MV.cCtaCod=CC.cCtaCod"
        sSql = sSql & "      and MV.nNroCalen=CC.nNroCalen"
        sSql = sSql & "      and MV.nCuota=CC.nCuota"
        sSql = sSql & " Where MV.nMovNro = " & pnMovNro
        
        If pbEjecBatch Then
            coConex.AdicionaCmdBatch sSql
        Else
            coConex.Ejecutar sSql
        End If
    End If
    '***************************
    sSql = "UPDATE ColocCalendDet SET nMontoPagado = nMontoPagado - ABS(MC.nMonto) "
    sSql = sSql & " From ColocCalendDet CC Inner Join MovColDet MC ON CC.cCtaCod = MC.cCtacod "
    sSql = sSql & "     And CC.nNroCalen = MC.nNroCalen  And CC.nPrdConceptoCod = MC.nPrdConceptoCod"
    sSql = sSql & "     And CC.nCuota = MC.nNroCuota "
    sSql = sSql & "     Where MC.nMovNro = " & pnMovNro & " And CC.nColocCalendApl = " & nAplicado
    sSql = sSql & "     AND CC.cCtaCod = '" & psCtaCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If
    
    '*************NAGL ERS036-2018************'
    ' Para extornar el monto condonado
    If psOpeCodCond <> "" Then
        sSql = "UPDATE ColocCalendDet SET nMontoPagado = nMontoPagado - ABS(MC.nMonto) "
        sSql = sSql & " From ColocCalendDet CC Inner Join MovColDet MC ON CC.cCtaCod = MC.cCtacod "
        sSql = sSql & "     And CC.nNroCalen = MC.nNroCalen"
        sSql = sSql & "     And CC.nCuota = MC.nNroCuota "
        sSql = sSql & "     Where MC.nMovNro = " & pnMovNro & " And CC.nColocCalendApl = " & nAplicado
        sSql = sSql & "     AND CC.cCtaCod = '" & psCtaCod & "'" & " And MC.cOpeCod = '" & psOpeCodCond & "'"
        sSql = sSql & "     And CC.nPrdConceptoCod = IIF(MC.cOpeCod = '" & psOpeCodCond & "'" & " And MC.nPrdConceptoCod = " & gColRecConceptoCodInteresMoratorio & "," & gColocConceptoCodInteresMoratorio & ", MC.nPrdConceptoCod)"

        If pbEjecBatch Then
            coConex.AdicionaCmdBatch sSql
        Else
            coConex.Ejecutar sSql
        End If
    End If
    '************END NAGL 20190724************'
    
    If Mid(psCtaCod, 6, 3) = "801" Then
    
        sSql = " Update CC"
        sSql = sSql & "    set nMontoPagado = 0 "
        sSql = sSql & " from MovMiVivienda MV"
        sSql = sSql & "    inner join ColocCalendDet CC on MV.cCtaCod=CC.cCtaCod"
        sSql = sSql & "      and MV.nNroCalen=CC.nNroCalen"
        sSql = sSql & "      and MV.nCuota=CC.nCuota"
        sSql = sSql & " Where MV.nMovNro = " & pnMovNro
        
        If pbEjecBatch Then
            coConex.AdicionaCmdBatch sSql
        Else
            coConex.Ejecutar sSql
        End If
    End If
    '***************************
    'Modificacion:LMMD por RFA extornar el interes diferido hacia el capital de la cuota
    
'    sSQL = "Update ColocCalendDet"
'    sSQL = sSQL & " Set nMontoPagado = nMontoPagado - Abs(MC.nMonto)"
'    sSQL = sSQL & " From ColocCalendDet CC"
'    sSQL = sSQL & " Inner Join MovColDet MC ON CC.cCtaCod = MC.cCtacod  And CC.nNroCalen = MC.nNroCalen  And CC.nPrdConceptoCod =" & gColocRFACapital
'    sSQL = sSQL & " And CC.nCuota = MC.nNroCuota"
'    sSQL = sSQL & " Where MC.nMovNro =" & pnMovNro & " And CC.nColocCalendApl =" & nAplicado & "  AND CC.cCtaCod = '" & psCtaCod & "' and MC.nPrdConceptoCod =" & gColocRFACapital
    
'    sSQL = "Update ColocCalendDet"
'    sSQL = sSQL & " Set nMontoPagado = nMontoPagado - Abs((Select Sum(MC1.nMonto)"
'    sSQL = sSQL & " From ColocCalendDet CC1"
'    sSQL = sSQL & " Inner Join MovColDet MC1 ON CC1.cCtaCod = MC1.cCtacod  And CC1.nNroCalen = MC1.nNroCalen  And CC1.nPrdConceptoCod =1010"
'    sSQL = sSQL & " And CC1.nCuota = MC1.nNroCuota"
'    sSQL = sSQL & " Where MC1.nMovNro=" & pnMovNro & " And CC1.nColocCalendApl =" & nAplicado & "  AND CC1.cCtaCod = '" & psCtaCod & "' and MC1.nPrdConceptoCod  in(1110)))"
'    sSQL = sSQL & " From ColocCalendDet CC"
'    sSQL = sSQL & " Inner Join MovColDet MC ON CC.cCtaCod = MC.cCtacod  And CC.nNroCalen = MC.nNroCalen  And CC.nPrdConceptoCod =" & gColocRFACapital
'    sSQL = sSQL & " And CC.nCuota = MC.nNroCuota"
'    sSQL = sSQL & " Where MC.nMovNro =" & pnMovNro & " And CC.nColocCalendApl =" & nAplicado & "  AND CC.cCtaCod = '" & psCtaCod & "' and MC.nPrdConceptoCod =" & gColocRFACapital
'
    
    sSql = "Update ColocCalendDet"
    sSql = sSql & " Set nMontoPagado = nMontoPagado - Abs(MD.nMonto)"
    sSql = sSql & " From ColocCalendDet CD"
    sSql = sSql & " Inner Join ColocacCred C on C.cCtaCod=CD.cCtaCod and C.nNroCalen=CD.nNroCalen"
    sSql = sSql & " Inner Join MovColDet MD on MD.cCtaCod=CD.cCtaCod and MD.nNroCalen=CD.nNroCalen"
    sSql = sSql & " Where CD.nCuota=MD.nNroCuota and CD.nColocCalendApl=1 and MD.nNroCalen=CD.nNroCalen and"
    sSql = sSql & " MD.nMovNro=" & pnMovNro & "and MD.cCtaCod='" & psCtaCod & "' and MD.nPrdConceptoCod=1110 and"
    sSql = sSql & " CD.nPrdConceptoCod = 1010 "

    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If
       
'    sSQL = "Update ColocCalendDet"
'    sSQL = sSQL & " Set nMontoPagado = nMontoPagado - Abs(isnull((Select Sum(MC1.nMonto)"
'    sSQL = sSQL & " From ColocCalendDet CC1"
'    sSQL = sSQL & " Inner Join MovColDet MC1 ON CC1.cCtaCod = MC1.cCtacod  And CC1.nNroCalen = MC1.nNroCalen  And CC1.nPrdConceptoCod =1000"
'    sSQL = sSQL & " And CC1.nCuota = MC1.nNroCuota"
'    sSQL = sSQL & " Where MC1.nMovNro=" & pnMovNro & " And CC1.nColocCalendApl =" & nAplicado & "  AND CC1.cCtaCod = '" & psCtaCod & "' and MC1.nPrdConceptoCod  in(1109)),0))"
'    sSQL = sSQL & " From ColocCalendDet CC"
'    sSQL = sSQL & " Inner Join MovColDet MC ON CC.cCtaCod = MC.cCtacod  And CC.nNroCalen = MC.nNroCalen  And CC.nPrdConceptoCod =  1000"
'    sSQL = sSQL & " And CC.nCuota = MC.nNroCuota"
'    sSQL = sSQL & " Where MC.nMovNro =" & pnMovNro & " And CC.nColocCalendApl =" & nAplicado & "  AND CC.cCtaCod = '" & psCtaCod & "' and MC.nPrdConceptoCod = 1109"
'
'
    sSql = "Update ColocCalendDet"
    sSql = sSql & " Set nMontoPagado = nMontoPagado - Abs(MD.nMonto)"
    sSql = sSql & " From ColocCalendDet CD"
    sSql = sSql & " Inner Join ColocacCred C on C.cCtaCod=CD.cCtaCod and C.nNroCalen=CD.nNroCalen"
    sSql = sSql & " Inner Join MovColDet MD on MD.cCtaCod=CD.cCtaCod and MD.nNroCalen=CD.nNroCalen"
    sSql = sSql & " Where CD.nCuota=MD.nNroCuota and CD.nColocCalendApl=1 and MD.nNroCalen=CD.nNroCalen and "
    sSql = sSql & " MD.nMovNro=" & pnMovNro & "and MD.cCtaCod='" & psCtaCod & "' and MD.nPrdConceptoCod=1109 and"
    sSql = sSql & " CD.nPrdConceptoCod = 1000"

    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If
        
    'Extorno de Gastos que se cargaron al Momento de la Operacion
    sSql = "UPDATE ColocCalendDet SET nMonto = CC.nMonto - ABS(MC.nMonto) "
    sSql = sSql & " From ColocCalendDet CC Inner Join MovColDet MC ON CC.cCtaCod = MC.cCtacod "
    sSql = sSql & "     And CC.nNroCalen = MC.nNroCalen  And CC.nPrdConceptoCod = MC.nPrdConceptoCod"
    sSql = sSql & "     And CC.nCuota = MC.nNroCuota "
    sSql = sSql & "     Where MC.nMovNro = " & pnMovNro & " And CC.nColocCalendApl = " & nAplicado
    sSql = sSql & "     AND CC.cCtaCod = '" & psCtaCod & "' AND "
    sSql = sSql & " MC.nPrdConceptoCod in ( Select nPrdConceptoCod from ProductoConcepto where nPrdConceptoCod like '12%' "
    sSql = sSql & " AND cAplicaProceso in ('CA','PA','PP')) "
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If
    
    'EJVG20150822 ***
    sSql = "EXEC spt_ins_ERS0632014_ExtornoLiberaGarantia " & pnMovNro & ",'" & psCtaCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If
    'END EJVG *******
    
    If nAplicado = gColocCalendAplDesembolso Then
        sSql = "Delete ColocCargoAutoma Where cCtaCod = '" & psCtaCod & "'"
        coConex.Ejecutar sSql
    End If
End Sub

Public Sub dInsertZonas(ByVal psUbiGeoCod As String, ByVal psUbiGeoDescripcion As String, Optional pbEjecBatch As Boolean = False)
Dim sSql As String
    
    sSql = "INSERT INTO UbicacionGeografica(cUbiGeoCod, cUbiGeoDescripcion) "
    sSql = sSql & " VALUES('" & psUbiGeoCod & "','" & psUbiGeoDescripcion & "')"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If
End Sub

Public Sub dUpdateZonas(ByVal psUbiGeoCod As String, ByVal psUbiGeoDescripcion As String, Optional pbEjecBatch As Boolean = False)
Dim sSql As String

    sSql = "UPDATE UbicacionGeografica SET cUbiGeoDescripcion = '" & psUbiGeoDescripcion & "' "
    sSql = sSql & " Where cUbiGeoCod = '" & psUbiGeoCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If
End Sub

Public Sub dDeleteZonas(ByVal psUbiGeoCod As String, Optional pbEjecBatch As Boolean = False)
Dim sSql As String

    sSql = "DELETE UbicacionGeografica  Where cUbiGeoCod = '" & psUbiGeoCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If
End Sub

Public Sub dInsertFeriado(ByVal dFeriado As Date, ByVal cDescrip As String, ByVal cUltimaActualizacion As String, Optional pbEjecBatch As Boolean = False)
Dim sSql As String
    
    sSql = "INSERT INTO Feriado(dFeriado, cDescrip, cUltimaActualizacion) "
    sSql = sSql & " VALUES('" & Format(dFeriado, "mm/dd/yyyy") & "','" & cDescrip & "','" & cUltimaActualizacion & "')"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If
End Sub

Public Sub dDeleteFeriado(ByVal dFeriado As Date, Optional pbEjecBatch As Boolean = False)
Dim sSql As String
    
    sSql = "DELETE Feriado Where dFeriado = '" & Format(dFeriado, "mm/dd/yyyy") & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If

End Sub


Public Sub dInsertAgencia(ByVal cAgeCod As String, ByVal cAgeDescripcion As String, ByVal cAgeDireccion As String, _
    ByVal cAgeTelefono As String, ByVal cSubCtaCod As String, ByVal nAgeEspecial As Integer, _
    ByVal cUbiGeoCod As String, ByVal cUltimaActualizacion As String, Optional pbEjecBatch As Boolean = False)
Dim sSql As String

    sSql = "INSERT INTO Agencias(cAgeCod, cAgeDescripcion, cAgeDireccion, cAgeTelefono, cSubCtaCod, nAgeEspecial, cUbiGeoCod, cUltimaActualizacion) "
    sSql = sSql & " VALUES('" & cAgeCod & "','" & cAgeDescripcion & "','" & cAgeDireccion & "','" & cAgeTelefono & "','" & cSubCtaCod & "'," & nAgeEspecial & ",'" & Trim(Right(cUbiGeoCod, 15)) & "','" & cUltimaActualizacion & "')"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If

End Sub

Public Sub dUpdateAgencia(ByVal cAgeCod As String, ByVal cAgeDescripcion As String, ByVal cAgeDireccion As String, _
    ByVal cAgeTelefono As String, ByVal cSubCtaCod As String, ByVal nAgeEspecial As Integer, _
    ByVal cUbiGeoCod As String, ByVal cUltimaActualizacion As String, Optional pbEjecBatch As Boolean = False)

Dim sSql As String

    sSql = "Update Agencias SET "
    sSql = sSql & " cAgeDescripcion = '" & cAgeDescripcion & "',"
    sSql = sSql & " cAgeDireccion = '" & cAgeDireccion & "',"
    sSql = sSql & " cAgeTelefono = '" & cAgeTelefono & "',"
    sSql = sSql & " cSubCtaCod = '" & cSubCtaCod & "',"
    sSql = sSql & " nAgeEspecial = " & nAgeEspecial & ","
    sSql = sSql & " cUbiGeoCod = '" & Trim(Right(cUbiGeoCod, 20)) & "',"
    sSql = sSql & " cUltimaActualizacion = '" & cUltimaActualizacion & "' "
    sSql = sSql & " Where cAgeCod = '" & cAgeCod & "'"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If

End Sub

Public Sub dDeleteAgencia(ByVal cAgeCod As String, Optional pbEjecBatch As Boolean = False)
Dim sSql As String
    sSql = "DELETE Agencias "
    sSql = sSql & " Where cAgeCod = '" & cAgeCod & "'"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If

End Sub
'Para los extornos de Mivivienda - CAFF
Public Sub dDeleteColocCalifMiViv(ByVal psCodCta As String)
Dim sSql  As String

    sSql = "DELETE ColocCalifMiViv WHERE cCtaCod = '" & psCodCta & "'"
    coConex.Ejecutar sSql

End Sub

'Por Error en la amortizacion se cuelga al abrir otra conexion - CAFF
Public Function RecuperaColocacCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaColocacCred
    sSql = "Select * from ColocacCred where cCtacod = '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaColocacCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocacCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'Por Error en la amortizacion se cuelga al abrir otra conexion - CAFF
Public Function RecuperaColocaciones(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaColocaciones
    sSql = "Select * from Colocaciones where cCtaCod = '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaColocaciones = oConecta.CargaRecordSet(sSql)
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
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta


    On Error GoTo ErrorRecuperaCalendarioGastos
    sSql = "Select * from ColocCalendDet where cCtaCod = '" & psCtaCod & _
            "' AND nNroCalen = " & pnNroCalen
    If Not pbTodos Then
        sSql = sSql & " AND nCuota = " & pnNroCuota
    End If
    sSql = sSql & " AND nColocCalendApl = " & pnAplicado
    sSql = sSql & " AND nPrdConceptoCod like '12%'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioGastos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaCalendarioGastos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Sub dInsertaPenalidad(ByVal pnValorInicial As Double, ByVal pnValorFinal As Double, ByVal pnPenalidad As Double, ByVal psTipo As String)
    Dim sSql As String
    
    sSql = "Insert Into Penalidad(nPorcenIni,nPorcenFin,nPenalidad,TipoPenalidad)"
    sSql = sSql & "Values(" & pnValorInicial & "," & pnValorFinal & "," & pnPenalidad & ",'" & psTipo & "')"
    
   coConex.Ejecutar sSql
End Sub

Public Sub dActualizaPenalidad(ByVal pnCodigo As Integer, ByVal pnValorInicial As Double, ByVal pnValorFinal As Double, ByVal pnPenalida As Double, ByVal psTipo As String)
    Dim sSql As String
    
    sSql = "Update Penalidad"
    sSql = sSql & " set nPorcenIni=" & pnValorInicial & ","
    sSql = sSql & " nPorcenFin=" & pnValorFinal & ","
    sSql = sSql & " nPenalidad=" & pnPenalida
    sSql = sSql & " Where nCodigo=" & pnCodigo & ""
    
    coConex.Ejecutar sSql
End Sub

Public Sub dEliminacionPenalidad(ByVal pnCodigo As Integer)
    Dim sSql As String
    
    sSql = "Delete From Penalidad Where nCodigo=" & pnCodigo
    
    coConex.Ejecutar sSql
End Sub

Public Function BuscarMov(ByVal psMovNro As String, Optional psFiltro As String = "") As Boolean
    Dim SQL As String
    Dim rs As ADODB.Recordset

    BuscarMov = False
    
    SQL = "SELECT cMovnro from Mov WHere cMovNro LIKE '" & psMovNro & "%'" & IIf(psFiltro <> "", " AND ", "") & psFiltro
    Set rs = coConex.CargaRecordSet(SQL)
    BuscarMov = Not rs.EOF
    rs.Close
    Set rs = Nothing

End Function

Public Function InsertaMovCont(ByVal pnMovNro As Long, ByVal pnMovMonto As Double, ByVal pnMovCorrela As Integer, ByVal psMovParalelo As String) As Integer
    Dim SQL As String
    InsertaMovCont = 1
    
    SQL = "INSERT INTO MovCont (nMovNro, nMovMonto, nMovCorrela, cMovParalelo ) " _
         & "VALUES (" & pnMovNro & "," & pnMovMonto & "," & pnMovCorrela & ",'" & psMovParalelo & "')"
     
    coConex.Ejecutar SQL
    InsertaMovCont = 0
    Exit Function

End Function

Public Function InsertaMovPendientesRend(ByVal pnMovNro As Long, ByVal psCtaContCod As String, ByVal pnSaldo As Currency) As Integer
    Dim SQL As String
    Dim rs As ADODB.Recordset
    
    InsertaMovPendientesRend = 1
    SQL = "SELECT nMovNro FROM MovPendientesRend Where nMovNro = " & pnMovNro & " and cCtaContCod = '" & psCtaContCod & "' "
    Set rs = coConex.CargaRecordSet(SQL)
    If rs.EOF Then
        SQL = "INSERT INTO MovPendientesRend (nMovNro, nMovMonto, nMovCorrela, cMovParalelo ) " _
             & "VALUES (" & pnMovNro & ", '" & psCtaContCod & "'," & pnSaldo & ") "
    Else
        SQL = "UPDATE MovPendientesRend SET nSaldo = " & pnSaldo & " WHERE nMovNro = " & pnMovNro & " "
    End If
    coConex.Ejecutar SQL
    InsertaMovPendientesRend = 0

End Function


Public Function InsertaMovARendir(ByVal pnMovNro As Long, ByVal psTpoArendir As ArendirTipo, ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal psPersCod As String, ByVal pnMovSaldo As Currency, ByVal psCategCod As String) As Integer

    Dim SQL As String
    InsertaMovARendir = 1
    
    SQL = "INSERT INTO MovARendir(nMovNro,cTpoArendir, cAgeCod, cAreaCod, cPersCod, nMovSaldo , cCategCod ) " _
     & "VALUES (" & pnMovNro & ", '" & psTpoArendir & "','" & psAgeCod & "','" & psAreaCod & "','" & psPersCod & "'," & pnMovSaldo & ",'" & psCategCod & "')"
     
    coConex.Ejecutar SQL
    InsertaMovARendir = 0

End Function
Public Sub ActualizaMov(ByVal pnMovNro As Long, Optional ByVal psMovDesc As String = "", Optional ByVal pnMovEstado As MovEstado = -1, Optional ByVal psMovFlag As MovFlag = gMovFlagVigente)

    Dim SQL As String
    Dim lsActualiza As String
    lsActualiza = ""
    If psMovDesc <> "" Then
        lsActualiza = lsActualiza & " cMovDesc='" & Replace(psMovDesc, "'", "''") & "',"
    End If
    If pnMovEstado > -1 Then
      lsActualiza = lsActualiza & " nMovEstado=" & pnMovEstado & ","
    End If
    lsActualiza = lsActualiza & " nMovFlag=" & psMovFlag & ","
    lsActualiza = Mid(lsActualiza, 1, Len(lsActualiza) - 1)
    If lsActualiza = "" Then Exit Sub
    
    SQL = " Update Mov set " & lsActualiza & " Where nMovNro =" & pnMovNro & ""
    coConex.Ejecutar SQL

End Sub
Public Sub EliminaMov(ByVal psMovNro As String, Optional psMovNroActualiza As String)
Dim SQL As String

    SQL = "UPDATE Mov SET nMovFlag = '" & gMovFlagEliminado & "' WHERE cMovNro = '" & psMovNro & "'"
    coConex.Ejecutar SQL
    If Not psMovNroActualiza = "" Then
        InsertaMovModifica psMovNroActualiza, psMovNro
    End If
    ActualizaSaldoMovimiento psMovNro, "-"

End Sub
Public Sub InsertaMovModifica(ByVal psMovNro As String, ByVal psMovNroAnt As String, Optional psMovNroNew As String)
Dim SQL As String

    SQL = "INSERT MovModifica (cMovNro, cMovNroAnt,cMovNroNew) VALUES ('" & psMovNro & "','" & psMovNroAnt & "', '" & psMovNroNew & "')"
    coConex.Ejecutar SQL

End Sub
'##ModelId=3A848F08038A
Public Function InsertaMovCta(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psCtaContCod As String, ByVal pnMovImporte As Currency) As Integer

    Dim SQL As String
    InsertaMovCta = 1
    
    SQL = " INSERT INTO MOVCTA(nMovNro, nMovItem, cCtaContCod, nMovImporte )" _
        & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & ") "
      
    coConex.Ejecutar SQL
    InsertaMovCta = 0

End Function

Public Function InsertaMovCartaFianza(ByVal pnMovNro As Long, psPersCod As String, psIFPersCod As String, psCartaVenc As String) As Integer

    Dim SQL As String
    InsertaMovCartaFianza = 1
    
    SQL = " INSERT INTO MOVCartaFianza (nMovNro, cPersCod, cIFPersCod, dCartaVenc )" _
        & " VALUES(" & pnMovNro & ",'" & psPersCod & "','" & psIFPersCod & "','" & psCartaVenc & "') "
    coConex.Ejecutar SQL
    
    InsertaMovCartaFianza = 0

End Function

Public Sub ActualizaMovCta(ByVal pnMovNro As Long, ByVal pnMovItem As Long, Optional psCtaCod As String = "", Optional pnMovImporte As Currency = 0)

    Dim SQL As String
    Dim lsActualiza As String
    lsActualiza = ""
    If psCtaCod <> "" Then
        lsActualiza = lsActualiza & " cCtaContCod ='" & psCtaCod & "',"
    End If
    If pnMovImporte <> 0 Then
       lsActualiza = lsActualiza & " nMovImporte = " & pnMovImporte & ","
    End If
    If lsActualiza <> "" Then
      lsActualiza = Mid(lsActualiza, 1, Len(lsActualiza) - 1)
    End If
    If lsActualiza = "" Then Exit Sub
    
    SQL = " UPDATE MovCta SET " & lsActualiza & " WHERE nMovNro =" & pnMovNro & " and nMovItem = " & pnMovItem & ""
    coConex.Ejecutar SQL

End Sub
Public Function InsertaMovMe(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovMeImporte As Currency) As Integer

    Dim SQL As String
    InsertaMovMe = 1
    
    SQL = " INSERT INTO MOVME (nMovNro, nMovItem, nMovMEImporte )" _
        & " VALUES(" & pnMovNro & "," & pnMovItem & ", " & pnMovMeImporte & ") "
      
    coConex.Ejecutar SQL
    InsertaMovMe = 0

End Function

'##ModelId=3A848F3301A5
Public Function InsertaMovObj(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As Long, ByVal psObjetoCod As String) As Integer

    Dim SQL As String
    InsertaMovObj = 1
     
    SQL = "INSERT INTO MOVOBJ (nMovNro, nMovItem, nMovObjOrden, cObjetoCod) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & Format(psObjetoCod, "00") & "')"

    coConex.Ejecutar SQL
    InsertaMovObj = 0

End Function
Public Function InsertaMovObjViaticos(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psObjetoCod As String, ByVal pnImporte As Currency) As Integer

    Dim SQL As String
    InsertaMovObjViaticos = 1
     
    SQL = "INSERT INTO MovObjViaticos (nMovNro, nMovItem, cObjetoCod, nMovImporte) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psObjetoCod & "'," & pnImporte & ")"

    coConex.Ejecutar SQL
    InsertaMovObjViaticos = 0

End Function

Public Function InsertaMovObjIF(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As String, ByVal psPersCod As String, ByVal psTipoIF As String, ByVal psCtaIfCod As String) As Integer

    Dim SQL As String
    InsertaMovObjIF = 1
     
    SQL = "INSERT INTO MOVOBJIF(nMovNro, nMovItem, nMovObjOrden, cPersCod, cIFTpo, cCtaIfCod ) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & psPersCod & "','" & Format(psTipoIF, "00") & "','" & psCtaIfCod & "')"
    
    coConex.Ejecutar SQL
    InsertaMovObjIF = 0

End Function

'##ModelId=3A848F4F002E
Public Function InsertaMovDoc(ByVal pnMovNro As Long, ByVal pnDocTpo As TpoDoc, ByVal psDocNro As String, ByVal psDocFecha As String) As Integer

    Dim SQL As String
    InsertaMovDoc = 1
    
    SQL = "INSERT INTO MOVDOC (nMovNro, nDocTpo, cDocNro, dDocFecha ) " _
        & "VALUES (" & pnMovNro & "," & pnDocTpo & ",'" & psDocNro & "','" & psDocFecha & "')"
        
    coConex.Ejecutar SQL
    InsertaMovDoc = 0

End Function
'##ModelId=3A84A3FE038A
Public Function InsertaMovCajaChica(ByVal pnMovNro As Long, ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal pnProcNro As Integer, ByVal pnProcTpo As CHTipoProc) As Integer
 
    Dim SQL As String
    InsertaMovCajaChica = 1
    SQL = " INSERT INTO MOVCAJACHICA(nMovNro, cAreaCod, cAgeCod, nProcNro, cProcTpo ) " _
        & " VALUES(" & pnMovNro & ",'" & psAreaCod & "','" & psAgeCod & "'," & pnProcNro & ",'" & pnProcTpo & "')"
    
    coConex.Ejecutar SQL
    
    InsertaMovCajaChica = 0

End Function
Public Function ActualizaMovCajaChicaNroProc(ByVal pnMovNro As Long, Optional ByVal pnProcNro As Integer = 0) As Integer

    Dim SQL As String
    ActualizaMovCajaChicaNroProc = 1
    SQL = " UPDATE MOVCAJACHICA SET nProcNro = " & pnProcNro & " WHERE nMovNro = " & pnMovNro & " "
    coConex.Ejecutar SQL
    
    ActualizaMovCajaChicaNroProc = 0

End Function
Public Function InsertaCajaChica(ByVal psAreaCod As String, ByVal psAgeCod As String, _
                                ByVal pnProcNro As Integer, ByVal psPersCod As String, _
                                ByVal pnMontoAsig As Currency, ByVal pnTopeEgresos As Currency, _
                                ByVal pnMontoDesem As Currency, ByVal pnSaldoAnt As Currency, _
                                ByVal pnSaldo As Currency) As Integer

    Dim SQL As String
    InsertaCajaChica = 1
    SQL = " INSERT INTO CAJACHICA(cAreaCod ,cAgeCod,nProcNro,cPersCod,nMontoAsig," _
        & "                         nTopeEgresos,nMontoDesem,nSaldoAnt,nSaldo) " _
        & " VALUES('" & psAreaCod & "','" & psAgeCod & "'," & pnProcNro & ",'" & psPersCod & "'," _
        & pnMontoAsig & "," & pnTopeEgresos & "," & pnMontoDesem & "," & pnSaldoAnt & "," _
        & pnSaldo & ")"
    
    coConex.Ejecutar SQL
    InsertaCajaChica = 0

End Function

Public Function InsertaMovBS(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovBSOrden As Long, ByVal psBSCod As String) As Integer

    Dim SQL As String
    InsertaMovBS = 1
     
    SQL = "INSERT INTO MOVBS (nMovNro, nMovItem, nMovBSOrden, cBSCod) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovBSOrden & ",'" & psBSCod & "')"
    coConex.Ejecutar SQL
    InsertaMovBS = 0

End Function
'
'Private Sub Class_Initialize()
'    Dim oImp As DImpresoras
'    Set oImp = New DImpresoras
'
'    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
'
'    Set oImp = Nothing
'Dim oIni As ClasIni
'
'Set oIni = New ClasIni
'
'vsServerAdmin = oIni.BaseAdministracion
'vsServerComunes = oIni.BaseComunes
'vsServerImagenes = oIni.BaseImagenes
'vsServerNegocio = oIni.BaseNegocio
'vsServerPersonas = oIni.BasePersonas
'
'Set coConex = New COMConecta.DCOMConecta
'If coConex.AbreConexion = False Then
'    Call RaiseError(MyUnhandledError, "DMov:Initialize Method. Error en Conexion a Base de datos")
'End If
'Set oIni = Nothing
'End Sub
'Private Sub Class_Terminate()
'coConex.CierraConexion
'Set coConex = Nothing
'End Sub

Public Function GeneraMovME(ByVal pnMovNro As Long, ByVal psMovNro As String, Optional nTCMerc As Currency = 0) As Boolean
Dim nTCFijo As Currency
Dim nTCVent As Currency
Dim nTCComp As Currency
Dim sSql As String
Dim rs   As New ADODB.Recordset
Dim dFecha As Date
Dim nImporteD As Currency
Dim nImporteS As Currency
Dim oGen As COMDConstSistema.DCOMGeneral

Set oGen = New COMDConstSistema.DCOMGeneral
   GeneraMovME = False
   dFecha = CDate(oGen.GetFechaMov(psMovNro, True))
   
   Set oGen = Nothing
   
   Dim oTC As New NCOMTipoCambio
   nTCFijo = oTC.EmiteTipoCambio(dFecha, TCFijoDia)
   If nTCMerc = 0 Then
        nTCVent = oTC.EmiteTipoCambio(dFecha, TCVenta)
        nTCComp = oTC.EmiteTipoCambio(dFecha, TCCompra)
   Else
        nTCVent = nTCMerc
        nTCComp = nTCMerc
   End If
   Set oTC = Nothing
   coConex.Ejecutar " sp_GeneraMovME " & pnMovNro & ", " & nTCFijo & ", " & nTCComp & " , " & nTCVent
   
   GeneraMovME = True

End Function

Private Sub GetClaseCtaCont(ByRef pnBalaMoneda As Boolean, ByRef psCtaClase, ByVal psCtaCod As String)
Dim oCta As New COMDContabilidad.DCOMCtaCont
Dim prs  As ADODB.Recordset
Set prs = oCta.CargaCtaContClase(psCtaCod)
If Not prs.EOF Then
   pnBalaMoneda = prs!nCtaMoneda
   psCtaClase = prs!cCtaCaracter
Else
   pnBalaMoneda = 0
   psCtaClase = ""
End If
prs.Close
Set oCta = Nothing
End Sub

Public Function InsertaMovRef(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long) As Integer

    Dim SQL As String
    InsertaMovRef = 1
    
    SQL = "INSERT INTO MovRef(nMovNro, nMovNroRef) " _
         & "VALUES (" & pnMovNro & "," & pnMovNroRef & ")"
     
    coConex.Ejecutar SQL
    InsertaMovRef = 0

End Function

Public Function InsertaMovCajaChicaRend(ByVal pnMovNroRend As Long, ByVal pnMovNroAnt As Long, ByVal pnMovNroNew As Long) As Integer

    Dim SQL As String
    InsertaMovCajaChicaRend = 1
    
    SQL = "INSERT INTO MovCajaChicaRend(nMovNroRend, nMovNroAnt, nMovNroNew) " _
         & "VALUES (" & pnMovNroRend & "," & pnMovNroAnt & "," & pnMovNroNew & ")"
     
    coConex.Ejecutar SQL
    InsertaMovCajaChicaRend = 0

End Function

Public Function ActualizaMovArendir(ByVal pnMovNro As Long, Optional ByVal psTpoArendir As ArendirTipo = gArendirTipoCajaGeneral, Optional ByVal psAreaCod As String, Optional ByVal psAgeCod As String, Optional ByVal psPersCod As String, Optional ByVal pnMovSaldo As Currency = -9999) As Integer

    Dim SQL As String
    Dim lsActualiza As String
    lsActualiza = ""
    ActualizaMovArendir = 1
    If psAgeCod <> "" Then
        lsActualiza = lsActualiza & " cAgeCod='" & psAgeCod & "',"
    End If
    If psAreaCod <> "" Then
        lsActualiza = lsActualiza & " cAreaCod='" & psAreaCod & "',"
    End If
    If psPersCod <> "" Then
        lsActualiza = lsActualiza & " cPersCod='" & psPersCod & "',"
    End If
    If pnMovSaldo <> -9999 Then
        lsActualiza = lsActualiza & " nMovSaldo=" & pnMovSaldo & ","
    End If
    'nMovSaldo
    lsActualiza = lsActualiza & " cTpoArendir='" & psTpoArendir & "',"
    lsActualiza = Mid(lsActualiza, 1, Len(lsActualiza) - 1)
    
    If lsActualiza = "" Then Exit Function
    SQL = " Update MovARendir set " & lsActualiza & " Where nMovNro =" & pnMovNro & ""
    
    coConex.Ejecutar SQL
    ActualizaMovArendir = 0

End Function
Public Function InsertaMovObjPers(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As Long, ByVal psPersCod As String) As Integer

    Dim SQL As String
    InsertaMovObjPers = 1
    
    SQL = "INSERT INTO MovObjPers(nMovNro, nMovItem, nMovObjOrden, cPersCod ) " _
         & "VALUES (" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & psPersCod & "')"
     
    coConex.Ejecutar SQL
    InsertaMovObjPers = 0

End Function
Public Function InsertaMovObjAgenciaArea(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As Long, ByVal psAgeCod As String, ByVal psAreaCod As String) As Integer

    Dim SQL As String
    InsertaMovObjAgenciaArea = 1
    
    SQL = "INSERT INTO MovObjAreaAgencia(nMovNro, nMovItem, nMovObjOrden, cAreaCod, cAgeCod) " _
         & "VALUES (" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & psAreaCod & "','" & psAgeCod & "')"
     
    coConex.Ejecutar SQL
    InsertaMovObjAgenciaArea = 0

End Function
Public Function InsertaMovObjEfectivo(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As Long, ByVal psEfectivoCod As String) As Integer

    Dim SQL As String
    InsertaMovObjEfectivo = 1
     
    SQL = "INSERT INTO MovObjEfectivo(nMovNro, nMovItem, nMovObjOrden, cEfectivoCod ) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & psEfectivoCod & "')"

    coConex.Ejecutar SQL
    InsertaMovObjEfectivo = 0

End Function
Public Function InsertProv(ByVal psPersCod As String, psProvEstado As String, ByVal psUltimaActualizacion As String) As Integer

    Dim SQL As String
    InsertProv = 1
     
    SQL = "INSERT INTO Proveedor( cPersCod, nProvEstado, cUltimaActualizacion  ) " _
        & " VALUES('" & psPersCod & "','" & psProvEstado & "','" & psUltimaActualizacion & "')"

    coConex.Ejecutar SQL
    InsertProv = 0

End Function

Public Function InsertaMovGasto(ByVal pnMovNro As Long, ByVal psPersCod As String, ByVal psPersCodDest As String) As Integer

    Dim SQL As String
    InsertaMovGasto = 1
    
    SQL = "INSERT INTO MovGasto (nMovNro, cPersCod, cPersCodDest) " _
        & "VALUES (" & pnMovNro & ", '" & psPersCod & "','" & psPersCodDest & "')"
    coConex.Ejecutar SQL
    InsertaMovGasto = 0

End Function

Public Function InsertaMovOtrosItem(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psMovOtrasVar As String, ByVal pnMovOtroImporte As Currency, ByVal psMovOtroTexto As String) As Integer

    Dim SQL As String
    InsertaMovOtrosItem = 1
    
    SQL = "INSERT INTO MovOtrosItem (nMovNro,nMovItem,cMovOtroVariable,nMovOtroImporte,cMovOtroTexto) " _
     & "VALUES (" & pnMovNro & "," & pnMovItem & ",'" & psMovOtrasVar & "'," & pnMovOtroImporte & ",'" & psMovOtroTexto & "')"
     
    coConex.Ejecutar SQL
    InsertaMovOtrosItem = 0

End Function
Public Function InsertaMovArendirSust(ByVal pnMovNroSol As Long, ByVal pnMovNroSust As Long, Optional ByVal pnMovNroRend As Long = -1) As Integer

    Dim SQL As String
    InsertaMovArendirSust = 1
    
    SQL = "INSERT INTO MovArendirSust (nMovNro, nMovNroSust, nMovNroRend) " _
     & "VALUES (" & pnMovNroSol & "," & pnMovNroSust & "," & IIf(pnMovNroRend = -1, "Null", pnMovNroRend) & ")"
     
    coConex.Ejecutar SQL
    InsertaMovArendirSust = 0

End Function
'Public Function InsertaCheque(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psTipoIF As String, _
'            ByVal pnPlaza As ChequePlaza, ByVal psIFCta As String, ByVal pnMonto As Currency, _
'            ByVal pdValorizaRef As Date, ByVal pdValorizacion As Date, ByVal psDepIF As CGEstadosChq, _
'            ByVal pnConfCaja As CGEstadoConfCheque, ByVal pnMonedaCheque As Moneda, _
'            Optional ByVal pnEstado As ChequeEstado = gChqEstEnValorizacion, Optional psAreaCod As String, _
'            Optional psAgeCod As String, Optional pnProducto As Producto = 0)
'    On Error GoTo InsertaChequeErr
'    Dim sql As String
'    InsertaCheque = 1
'
'    sql = " INSERT INTO DOCREC(nTpoDoc, cNroDoc, cPersCod , cIFTpo, bPlaza, cIFCta, nMonto, dValorizaRef, dValorizacion, cDepIF, nConfCaja ,  nMoneda , nEstado, cAreaCod, cAgeCod, nProducto ) " _
'     & "    VALUES (" & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCod & "','" & Format(psTipoIF, "00") & "'," & pnPlaza & ",'" & psIFCta & "'," & pnMonto & ",'" _
'     & Format(pdValorizaRef, gsFormatoFecha) & "','" & Format(pdValorizacion, gsFormatoFecha) & "','" & psDepIF & "'," & pnConfCaja & "," & pnMonedaCheque & "," & pnEstado & ",'" & psAreaCod & "','" & psAgeCod & "'," & pnProducto & ")"
'
'    coConex.Ejecutar sql
'
'    InsertaCheque = 0
'    Exit Function
'InsertaChequeErr:
'    Call RaiseError(MyUnhandledError, "DMov:InsertaCheque Method")
'End Function

Public Function InsertaRegDocCuenta(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal cCtaCod As String) As Integer

    Dim SQL As String
    InsertaRegDocCuenta = 1
    
    SQL = " INSERT INTO RegDocCuenta(nDocTpo, cDocNro, cCtaCod)" _
        & " VALUES(" & pnTpoDoc & ",'" & psNroDoc & "','" & cCtaCod & "') "
    
    coConex.Ejecutar SQL
    InsertaRegDocCuenta = 0

End Function

Public Function InsertaDocRecCapta(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psTipoIF As String, _
                              ByVal psCodCta As String, ByVal pnMonto As Currency)

    Dim SQL As String
    InsertaDocRecCapta = 1
    
    SQL = "INSERT INTO DocRecCapta (nTpoDoc, cNroDoc, cPersCod, cIFTpo, cCtaCod, nMonto) VALUES " _
        & "(" & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCod & "','" & Format(psTipoIF, "00") & "', '" & psCodCta & "', " & pnMonto & ") "
    coConex.Ejecutar SQL
    
    InsertaDocRecCapta = 0

End Function
'
'Public Function ActualizaCheque(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psTipoIF As String, _
'                                Optional ByVal pdValorizaRef As String = "", Optional ByVal pdValorizacion As String = "", _
'                                Optional ByVal pnPlaza As ChequePlaza = -1, Optional ByVal psIFCta As String = "", _
'                                Optional ByVal pnMonto As Currency = -999999, _
'                                Optional ByVal pnDepIF As CGEstadosChq = -1, _
'                                Optional ByVal pnConfCaja As CGEstadoConfCheque = -1, _
'                                Optional ByVal pnMonedaCheque As Moneda = -1, _
'                                Optional ByVal pnEstado As ChequeEstado = -1)
'    On Error GoTo ActualizaChequeErr
'    Dim sql As String
'    Dim lsFiltro As String
'
'    ActualizaCheque = 1
'    lsFiltro = ""
'    If pdValorizaRef <> "" Then
'        lsFiltro = lsFiltro + " dValorizaRef  ='" & Format(pdValorizaRef, gsFormatoFecha) & "', "
'    End If
'    If pdValorizacion <> "" Then
'        lsFiltro = lsFiltro + " dValorizacion ='" & Format(pdValorizacion, gsFormatoFecha) & "', "
'    End If
'    If pnPlaza <> -1 Then
'        lsFiltro = lsFiltro + " bPlaza = " & pnPlaza & ","
'    End If
'    If psIFCta <> "" Then
'        lsFiltro = lsFiltro + " cIFCta = '" & psIFCta & "',"
'    End If
'    If pnMonto <> -999999 Then
'        lsFiltro = lsFiltro + " nMonto =" & pnMonto & ","
'    End If
'    If pnDepIF <> -1 Then
'        lsFiltro = lsFiltro + " cDepIF='" & pnDepIF & "',"
'    End If
'    If pnConfCaja <> -1 Then
'        lsFiltro = lsFiltro + " nConfCaja=" & pnConfCaja & ","
'    End If
'    If pnMonedaCheque <> -1 Then
'        lsFiltro = lsFiltro + " nMoneda=" & pnMonedaCheque & ","
'    End If
'    If pnEstado <> -1 Then
'        lsFiltro = lsFiltro + " nEstado=" & pnEstado & ","
'    End If
'    lsFiltro = Mid(lsFiltro, 1, Len(lsFiltro) - 1)
'
'    sql = " UPDATE  DOCREC SET " & lsFiltro & " " _
'     & "    WHERE   nTpoDoc =" & pnTpoDoc & " AND cNroDoc='" & psNroDoc & "' AND " _
'     & "            cPersCod='" & psPersCod & "' AND cIFTpo ='" & psTipoIF & "'"
'
'    coConex.Ejecutar sql
'    ActualizaCheque = 0
'    Exit Function
'ActualizaChequeErr:
'    Call RaiseError(MyUnhandledError, "DMov:ActualizaCheque Method")
'End Function

Public Function InsertaChequeEstado(ByVal psTpoDoc As TpoDoc, ByVal psNroDoc As String, psPersCod As String, ByVal psTipoIF As String, ByVal pdFechaEstado As Date, ByVal psEstado As ChequeEstado, ByVal psMovNro As String, Optional ByVal psCtaCheque As String) As Integer

    Dim SQL As String
    InsertaChequeEstado = 1
    
    SQL = " INSERT INTO DocRecEst(nTpoDoc, cNroDoc, cPersCod, cIFTpo , nEstado , cMovNro,CIFCTA ) " _
     & "    VALUES (" & psTpoDoc & ", '" & psNroDoc & "','" & psPersCod & "','" & Format(psTipoIF, "00") & "'," _
     & psEstado & ",'" & psMovNro & "','" & psCtaCheque & "')"
     
    coConex.Ejecutar SQL
    InsertaChequeEstado = 0

End Function
Public Function ActualizaMovimiento(ByVal psMovNroNew As String, ByVal pnMovNroAnt As Long, Optional pnMovEstado As MovEstado = gMovEstContabMovContable, _
                                Optional plActualizaAsiento As Boolean = True, Optional pbInsertaMov As Boolean = True, Optional pbDeleteMovAnt As Boolean = True, _
                                Optional ByVal psOpeCod As String = "", Optional ByVal psMovDesc As String = "", Optional ByVal lbUpdaMovRef As Boolean = True) As Integer

   Dim SQL As String
   Dim lnMovNronew As Long
   Dim lsMovNroAnt As String
   ActualizaMovimiento = 1
   
   lsMovNroAnt = GetcMovNro(pnMovNroAnt)
   If lsMovNroAnt = psMovNroNew Then
   End If
   If plActualizaAsiento Then
      If pbInsertaMov Then
        SQL = " INSERT INTO MOV (cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag )" _
              & " SELECT '" & psMovNroNew & "'," & IIf(psOpeCod = "", "cOpeCod", "'" & psOpeCod & "'") & "," _
            & IIf(psMovDesc = "", "cMovDesc", "'" & psMovDesc & "'") & ",'" & pnMovEstado & "', nMovFlag  " _
            & " FROM Mov WHERE nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        lnMovNronew = GetnMovNro(psMovNroNew)
      Else
        lnMovNronew = GetnMovNro(psMovNroNew)
      End If
      SQL = " UPDATE MOVCONT SET nMovNro=" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
      coConex.Ejecutar SQL
   
      If pbDeleteMovAnt Then
          SQL = "UPDATE MOV set nMovFlag=" & gMovFlagEliminado & " where nMovNro =" & pnMovNroAnt & " "
          coConex.Ejecutar SQL
          ActualizaSaldoMovimiento lsMovNroAnt, "-"
      End If
   Else
      lnMovNronew = GetnMovNro(psMovNroNew)
   End If
   
   'Cambio de Mov de Arqueos
   SQL = "INSERT INTO MOVARQUEO(nMovNro, cAreaCod, cPersCod, cHoraIni, cHoraFin, nEfectivo, nOPImporte, nTotDoc, nTotDocRend, nArendirPend, nDiferencia )      " _
    & " SELECT " & lnMovNronew & ",cAreaCod, cPersCod, cHoraIni, cHoraFin, nEfectivo, nOPImporte, nTotDoc, nTotDocRend, nArendirPend, nDiferencia FROM MOVARQUEO Where nMovNro =" & pnMovNroAnt & " "
   coConex.Ejecutar SQL
   
   SQL = "UPDATE MOVARQUEOOP SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & "; "
'   coConex.Ejecutar sql
   
   SQL = "UPDATE MOVARQUEOEFECTIVO      SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
   coConex.Ejecutar SQL
   
   SQL = "DELETE MOVARQUEO Where nMovNro =" & pnMovNroAnt & "; "
   coConex.Ejecutar SQL
   ' Fin de Cambio de Mov de Arqueos
   
   If lbUpdaMovRef Then
        SQL = "UPDATE MOVREF         SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & "; "
'        coConex.Ejecutar sql
        SQL = SQL & "UPDATE MOVREF         SET nMovNroRef =" & lnMovNronew & " Where nMovNroRef =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
   End If
   
   SQL = "UPDATE MOVARENDIR     SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & "; "
   'coConex.Ejecutar sql
   SQL = SQL & "UPDATE MOVARENDIRSUST SET nMovNroRend=" & lnMovNronew & " where nMovNroSust =" & pnMovNroAnt & "; "
   'coConex.Ejecutar sql
   SQL = SQL & "UPDATE MOVOTROSITEM   SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & "; "
   'coConex.Ejecutar sql
   SQL = SQL & "UPDATE MOVCAJACHICA   SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & "; "
   'coConex.Ejecutar sql
   SQL = SQL & "UPDATE MOVCARTAFIANZA SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & "; "
   'coConex.Ejecutar sql
   SQL = SQL & "UPDATE MOVTPOCAMBIO   SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & "; "
   'coConex.Ejecutar sql
   SQL = SQL & "UPDATE MOVVIATICOS    SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   coConex.Ejecutar SQL
    
   If plActualizaAsiento Then
        SQL = " INSERT MOVCTA (nMovNro, nMovItem, cCtaContCod, nMovImporte ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, cCtaContCod, nMovImporte FROM MOVCTA Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
        SQL = "INSERT MOVOBJ (nMovNro, nMovItem, nMovObjOrden, cObjetoCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cObjetoCod FROM MOVOBJ Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
        SQL = "UPDATE MOVOBJIF          SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & "; "
        'coConex.Ejecutar sql
        SQL = SQL & "UPDATE MOVOBJEfectivo    SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & "; "
        'coConex.Ejecutar sql
        
        SQL = SQL & "UPDATE MOVOBJAreaAgencia SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & "; "
        'coConex.Ejecutar sql
        
        SQL = SQL & "UPDATE MOVME             SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL

        SQL = "DELETE MOVOBJ            Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
        SQL = "DELETE MOVCTA            Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
        SQL = "UPDATE MOVDOC          SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
        SQL = " UPDATE MOVGASTO       SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
        ActualizaSaldoMovimiento psMovNroNew, "+"
   End If
   ActualizaMovimiento = 0

End Function
'**************************
'** ULTIMAS FUNCIONES
'**************************

Public Function CargaMovDocAsiento(ByVal pnMovNro As Long) As ADODB.Recordset
Dim psSql As String

   psSql = "SELECT md.nDocTpo, d.cDocDesc, cDocNro, Convert(varchar(10),md.dDocFecha,103) dDocFecha, m.nMovNro, m.cMovNro " _
         & "FROM Mov m JOIN MovDoc md ON md.nMovNro = m.nMovNro " _
         & "           JOIN Documento d ON d.nDocTpo = md.nDocTpo " _
         & "WHERE m.nMovNro = " & pnMovNro
   Set CargaMovDocAsiento = coConex.CargaRecordSet(psSql)

End Function
'
'Public Function CargaMovOpeAsiento(ByVal pnMovNro As Long, Optional ByVal psMovNro As String = "") As ADODB.Recordset
'Dim psSql As String
'   On Error GoTo CargaOpeMovAsientoErr
'      psSql = "SELECT m.cMovNro, m.cOpeCod, o.cOpeDesc, m.cMovDesc, m.nMovEstado, m.nMovFlag, ISNULL(mg.cPersCod,'') cPersCod, ISNULL(p.cPersNombre,'') cPersNombre, ISNULL(pid.cPersIDnro, '00000000') cRuc  " _
'            & "FROM Mov M JOIN " & vsServerComunes & "OpeTpo O ON O.cOpeCod = M.cOpeCod " _
'            & "      LEFT JOIN MovGasto mg ON mg.nMovNro = m.nMovNro " _
'            & "      LEFT JOIN " & vsServerPersonas & "Persona p ON p.cPersCod = mg.cPersCod LEFT JOIN  PersID pid ON pid.cPersCod = p.cPersCod and pid.cPersIDTpo = " & gPersIdRUC _
'            & "WHERE " & IIf(psMovNro <> "", " M.cMovNro = '" & psMovNro & "' ", " M.nMovNro = '" & pnMovNro & "'")
'      Set CargaMovOpeAsiento = coConex.CargaRecordSet(psSql)
'   Exit Function
'CargaOpeMovAsientoErr:
'   Call RaiseError(MyUnhandledError, "DMov:CargaMovOpeAsiento Method")
'End Function
'
'Public Function CargaMovCtaAsiento(ByVal pnMovNro As Long) As ADODB.Recordset
'Dim psSql As String
'   On Error GoTo CargaMovCtaAsientoErr
'      psSql = "SELECT cMovNro, a.nMovItem, a.cCtaContCod, isnull(b.cCtaContDesc,'') AS cCtaContDesc, " _
'         & "       CONVERT(varchar(20),ISNULL(SUM(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END),0),105) as nDebe, " _
'         & "       CONVERT(varchar(20),ISNULL(SUM(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END),0),105) as nHaber, " _
'         & "       CONVERT(varchar(20),ISNULL(SUM(CASE WHEN me.nMovMeImporte > 0 THEN me.nMovMeImporte END),0),105) as nDebeME, " _
'         & "       CONVERT(varchar(20),ISNULL(SUM(CASE WHEN me.nMovMeImporte < 0 THEN me.nMovMeImporte * -1 END),0),105) nHaberME " _
'         & " FROM  MovCta a      JOIN Mov M ON M.nMovNro = a.nMovNro " _
'         & "                LEFT JOIN MovMe Me ON Me.nMovNro = a.nMovNro and Me.nMovItem = a.nMovItem " _
'         & "                LEFT JOIN " & vsServerComunes & "CtaCont b ON a.cCtaContCod = b.cCtaContCod " _
'         & "WHERE  M.nMovNro = '" & pnMovNro & "'" _
'         & "GROUP BY cMovNro, a.nMovItem, a.cCtaContCod, b.cCtaContDesc ORDER BY a.nMovItem "
'      Set CargaMovCtaAsiento = coConex.CargaRecordSet(psSql)
'   Exit Function
'CargaMovCtaAsientoErr:
'   Call RaiseError(MyUnhandledError, "DMov:CargaMovCtaAsiento Method")
'End Function
'
'Public Function CargaMovObjAsiento(ByVal pnMovNro As Long) As ADODB.Recordset
'Dim psSql As String
'Dim lnMovNro As Long
'Dim prs      As ADODB.Recordset
'Dim prsObj   As ADODB.Recordset
'Dim prsRes   As ADODB.Recordset
'   On Error GoTo CargaMovObjAsientoErr
'    psSql = "SELECT cMovNro, m.nMovNro, mo.nMovItem, mo.nMovObjOrden, mo.cObjetoCod " _
'          & "FROM Mov m JOIN MovObj mo ON m.nMovNro = mo.nMovNro " _
'          & "WHERE M.nMovNro = '" & pnMovNro & "' "
'    Set prs = coConex.CargaRecordSet(psSql)
'    Do While Not prs.EOF
'        Select Case prs!cObjetoCod
'           Case ObjCMACAgencias, ObjCMACAgenciaArea, ObjCMACArea
'              psSql = "SELECT mo.nMovNro, mo.nMovItem, nMovObjOrden, " & prs!cObjetoCod & " ObjPadre, mo.cAreacod+mo.cAgeCod cObjetoCod, ISNULL(ag.cAgeDescripcion,'')+' '+ISNULL(a.cAreaDescripcion,'') cObjetoDesc " _
'                    & "FROM MovObjAreaAgencia mo JOIN AreaAgencia aa ON aa.cAreaCod = mo.cAreaCod and aa.cAgeCod = mo.cAgeCod " _
'                    & "     LEFT JOIN Areas a ON a.cAreaCod = aa.cAreaCod " _
'                    & "     LEFT JOIN Agencias ag ON ag.cAgeCod = aa.cAgeCod " _
'                    & "WHERE mo.nMovNro = '" & prs!nMovNro & "' and mo.nMovItem = '" & prs!nMovItem & "' and mo.nMovObjOrden = " & prs!nMovObjOrden & ""
'
'           Case ObjEntidadesFinancieras
'              psSql = "SELECT mif.nMovNro, mif.nMovItem, mif.nMovObjOrden, " & prs!cObjetoCod & " ObjPadre, cif.cPersCod + cif.cCtaIfCod cObjetoCod, RTRIM(p.cPersNombre) + ' ' + cCtaIFDesc cObjetoDesc " _
'                    & "FROM MovObjIF mif " _
'                    & "     JOIN Persona p ON p.cPersCod = mif.cPersCod " _
'                    & "     JOIN CtaIF cif ON cif.cPersCod = mif.cPersCod and cif.cCtaIFCod = mif.cCtaIFCod " _
'                    & "WHERE mif.nMovNro = '" & prs!nMovNro & "' and mif.nMovItem = " & prs!nMovItem & " and mif.nMovObjOrden = " & prs!nMovObjOrden
'
'           Case ObjDescomEfectivo
'              psSql = "SELECT mo.nMovNro, mo.nMovItem, nMovObjOrden, " & prs!cObjetoCod & " ObjPadre, aa.cEfectivoCod cObjetoCod, aa.nEfectivoValor cObjetoDesc " _
'                   & "FROM MovObjEfectivo mo JOIN Efectivo aa ON mo.cEfectivoCod = aa.cEfectivoCod " _
'                   & "    LEFT JOIN Areas a ON a.cAreaCod = aa.cAreaCod LEFT JOIN Agencias ag ON ag.cAgeCod = aa.cAgeCod " _
'                   & "WHERE  mo.nMovNro = " & prs!nMovNro & " and mo.nMovItem = " & prs!nMovItem & " and mo.nMovObjOrden = " & prs!nMovObjOrden
'
'           Case ObjBienesServicios
'              psSql = "SELECT bs.nMovNro, bs.nMovItem, bs.nMovBsOrden nMovObjOrden, " & prs!cObjetoCod & " ObjPadre, bs.cBSCod cObjetoCod, b.cBSDescripcion cObjetoDesc " _
'                   & "FROM MovBS bs JOIN BienesServicios b ON b.cBSCod = bs.cBsCod " _
'                   & "WHERE  bs.nMovNro = " & prs!nMovNro & " and bs.nMovItem = " & prs!nMovItem
'           Case Else
'              psSql = "SELECT mo.nMovNro, mo.nMovItem, nMovObjOrden, '" & prs!cObjetoCod & "' ObjPadre, mo.cObjetoCod, o.cObjetoDesc " _
'                   & "FROM MovObj mo JOIN  Objeto o ON o.cObjetoCod = mo.cObjetoCod " _
'                   & "WHERE  mo.nMovNro = " & prs!nMovNro & " and mo.nMovItem = " & prs!nMovItem & " and not mo.cObjetoCod IN ('" & ObjPersona & "', '" & ObjEntidadesFinancieras & "', '" & ObjCMACAgenciaArea & "', '" & ObjCMACArea & "','" & ObjCMACAgencias & "','" & ObjBienesServicios & "') "
'        End Select
'        Set prsObj = coConex.CargaRecordSet(psSql)
'        RecordSetAdiciona prsRes, prsObj
'        prs.MoveNext
'    Loop
'    Set CargaMovObjAsiento = prsRes
'Exit Function
'
'CargaMovObjAsientoErr:
'   Call RaiseError(MyUnhandledError, "DMov:CargaMovObjAsiento Method")
'End Function

Public Function CargaSumaMovAsiento(psMovNro As String) As ADODB.Recordset
Dim psSql As String

      psSql = "SELECT SUM(CASE WHEN mc.nMovImporte > 0 THEN mc.nMovImporte END) as nDebe, " _
            & "       SUM(CASE WHEN mc.nMovImporte < 0 THEN mc.nMovImporte * -1 END) as nHaber, " _
            & "       ISNULL(SUM(CASE WHEN me.nMovMeImporte > 0 THEN me.nMovMeImporte END),0) as nDebeME, " _
            & "       ISNULL(SUM(CASE WHEN me.nMovMeImporte < 0 THEN me.nMovMeImporte * -1 END),0) as nHaberME " _
            & "FROM   Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
            & "WHERE  m.cMovNro = '" & psMovNro & "'"
      Set CargaSumaMovAsiento = coConex.CargaRecordSet(psSql)

End Function
Public Function ActualizaArendirSust(ByVal pnMovNroSol As Long) As Integer

    Dim SQL As String
    ActualizaArendirSust = 1
    
    SQL = "UPDATE MovArendirSust SET nMovNroRend=NULL WHERE nMovNro =" & pnMovNroSol & ""
    coConex.Ejecutar SQL
    ActualizaArendirSust = 0

End Function
Public Function ExtornaMovimiento(ByVal psMovNroNew As String, ByVal pnMovNroAnt As Long, _
            Optional ByVal psOpeCod As String = "", Optional ByVal psMovDesc As String = "", _
            Optional ByVal pbEliminaMov As Boolean = False, Optional psMovActualiza As String = "") As Integer

    Dim SQL As String
    Dim lnMovNronew As Long
    Dim lsMovNroAnt As String
    Dim ldFechaAnt As Date
    Dim ldFechaExt As Date
    Dim oGen As COMDConstSistema.DCOMGeneral
    
    Set oGen = New COMDConstSistema.DCOMGeneral
    
    lsMovNroAnt = GetcMovNro(pnMovNroAnt)
    ldFechaAnt = CDate(oGen.GetFechaMov(lsMovNroAnt, True))
    ldFechaExt = CDate(oGen.GetFechaMov(psMovNroNew, True))
    
    Set oGen = Nothing
    
    ExtornaMovimiento = 1
    If ldFechaAnt = ldFechaExt Or pbEliminaMov = True Then
        EliminaMov lsMovNroAnt
        InsertaMovModifica psMovNroNew, lsMovNroAnt
        
    ElseIf ldFechaAnt <> ldFechaExt Or pbEliminaMov = False Then
        
        SQL = " INSERT INTO MOV (cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag )" _
            & " SELECT '" & psMovNroNew & "'," & IIf(psOpeCod = "", "cOpeCod", "'" & psOpeCod & "'") & ", " _
            & IIf(psMovDesc = "", "cMovDesc", "'" & Replace(psMovDesc, "''", "'") & "'") & ",nMovEstado, '" & gMovFlagDeExtorno & "'  " _
            & " FROM Mov WHERE nMovNro ='" & pnMovNroAnt & "' "
        coConex.Ejecutar SQL
        lnMovNronew = GetnMovNro(psMovNroNew)
        
        InsertaMovRef lnMovNronew, pnMovNroAnt
        ActualizaSaldoMovimiento lsMovNroAnt, "-"
        
        SQL = "INSERT MOVCONT (nMovNro, nMovMonto, nMovCorrela, cMovParalelo) " _
            & " SELECT nMovNro = " & lnMovNronew & ", nMovMonto, nMovCorrela, cMovParalelo FROM MovCont Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
       
        SQL = "INSERT MOVDOC (nMovNro, nDocTpo, cDocNro, dDocFecha) " _
            & " SELECT nMovNro = " & lnMovNronew & ", nDocTpo, cDocNro, dDocFecha FROM MovDoc Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
        SQL = "INSERT MOVGASTO (nMovNro, cPersCod, cPersCodDest) " _
            & " SELECT nMovNro = " & lnMovNronew & ",cPersCod, cPersCodDest FROM MovGasto Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
        SQL = "INSERT MOVCTA (nMovNro, nMovItem, cCtaContCod, nMovImporte ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, cCtaContCod, nMovImporte * -1  FROM MOVCTA Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
        SQL = "INSERT MOVME (nMovNro, nMovItem, nMovMEImporte ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovMEImporte*-1 FROM MOVME Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
        SQL = "INSERT MOVOBJ (nMovNro, nMovItem, nMovObjOrden, cObjetoCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cObjetoCod FROM MOVOBJ Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
        SQL = "INSERT MOVOBJIF (nMovNro, nMovItem, nMovObjOrden, cPersCod, cIFTpo, cCtaIfCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cPersCod, cIFTpo, cCtaIfCod FROM MOVOBJIF Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
'        sql = "INSERT MOVCTAIF (nMovNro, cPersCod, cIFTpo, cCtaIFCod, nConceptoCod, nMovImporte) " _
'            & " SELECT nMovNro =" & lnMovNronew & ", cPersCod, cIFTpo, cCtaIFCod, nConceptoCod, nMovImporte * -1 FROM MOVCTAIF Where nMovNro =" & pnMovNroAnt & " "
'        coConex.Ejecutar sql
                
        SQL = "INSERT MOVOBJEfectivo (nMovNro, nMovItem, nMovObjOrden, cEfectivoCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cEfectivoCod FROM MOVOBJEfectivo Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
        SQL = "INSERT MOVOBJAreaAgencia (nMovNro, nMovItem, nMovObjOrden, cAreaCod, cAgeCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cAreaCod, cAgeCod FROM MOVOBJAreaAgencia Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
        SQL = "INSERT INTO MOVHABILITACION(nMovNro, cTranspCod, nMontoTrans) " _
            & " SELECT nMovNro=" & lnMovNronew & ", cTranspCod, nMontoTrans * -1 FROM MOVHABILITACION Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
        SQL = "INSERT INTO MOVHABDEVCAJERO(nMovNro, cUser, nMovImporte)  " _
            & " SELECT nMovNro=" & lnMovNronew & ",cUser, nMovImporte *-1 FROM MOVHABDEVCAJERO Where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
        SQL = "UPDATE MOV SET nMovFlag = '" & gMovFlagExtornado & "' where nMovNro =" & pnMovNroAnt & " "
        coConex.Ejecutar SQL
        
        'ExtornaMovCtaIF pnMovNroAnt
    End If
    ActualizaSaldoMovimiento psMovNroNew, "+"
    ExtornaMovimiento = 0

End Function
'
'Public Function InsertaMovViaticos(ByVal pnViaticoMovNro As Long, ByVal pnMovNro As Long, _
'                                   ByVal psDestinoCod As ViaticosDestino, _
'                                   ByVal psTranspIda As ViaticosTransporte, _
'                                   ByVal psTranspVuelta As ViaticosTransporte, ByVal pdPartida As Date, _
'                                   ByVal pnMovViaticosDias As String, ByVal psDestinoDesc As String) As Integer
'
'    On Error GoTo InsertaMovViaticosErr
'    Dim sql As String
'    InsertaMovViaticos = 1
'
'    sql = "INSERT INTO MOVVIATICOS(nViaticoMovNro, nMovNro, " _
'        & "cDestinoCod, cTranspIda, cTranspVuelta,dPartida, nMovViaticosDias, cDestinoDesc  ) " _
'        & " VALUES (" & pnViaticoMovNro & "," & pnMovNro & ",'" _
'        & psDestinoCod & "','" & psTranspIda & "','" & psTranspVuelta & "','" _
'        & Format(pdPartida, gsFormatoFecha) & "'," & pnMovViaticosDias & ",'" & psDestinoDesc & "')"
'
'    coConex.Ejecutar sql
'    InsertaMovViaticos = 0
'    Exit Function
'InsertaMovViaticosErr:
'    Call RaiseError(MyUnhandledError, "DMov:InsertaMovViaticos Method")
'End Function

Public Sub InsertaMovLavDinero(ByVal nMovNro As Long, ByVal sPersCod As String, Optional SPersBen As String = "")
Dim sSql As String
If SPersBen = "" Then SPersBen = sPersCod

sSql = "INSERT MovLavDinero (nMovNro,cPersCod,cPersCodBen) " _
    & "VALUES (" & nMovNro & ",'" & sPersCod & "','" & SPersBen & "')"
coConex.Ejecutar sSql

End Sub

Public Function ActualizaCajaChica(ByVal psAreaCod As String, ByVal psAgeCod As String, _
                                    ByVal pnProcNro As Integer, _
                                    Optional ByVal psPersCod As String = "", _
                                    Optional ByVal pnMontoAsig As Currency = -99999, _
                                    Optional ByVal pnMontoDesem As Currency = -99999, _
                                    Optional ByVal pnMontoSaldAnt As Currency = -99999, _
                                    Optional ByVal pnMontoSaldo As Currency = -99999) As Integer
                                    

    Dim SQL As String
    Dim lsFiltro As String
    lsFiltro = ""
    If psPersCod <> "" Then
        lsFiltro = " cPersCod= '" & psPersCod & "',"
    End If
    If pnMontoAsig <> -99999 Then
        lsFiltro = lsFiltro & " nMontoAsig = " & pnMontoAsig & ","
    End If
    If pnMontoDesem <> -99999 Then
        lsFiltro = lsFiltro & " nMontoDesem = " & pnMontoDesem & ","
    End If
    If pnMontoSaldAnt <> -99999 Then
        lsFiltro = lsFiltro & " nMontoSaldAnt  = " & pnMontoSaldAnt & ","
    End If
    If pnMontoSaldo <> -99999 Then
        lsFiltro = lsFiltro & " nSaldo = " & pnMontoSaldo & ","
    End If
    lsFiltro = Mid(lsFiltro, 1, Len(lsFiltro) - 1)
    
    
    SQL = "UPDATE CAJACHICA SET " & lsFiltro _
    & "     WHERE  cAreaCod ='" & psAreaCod & "' AND  cAgeCod='" & psAgeCod & "' AND nProcNro=" & pnProcNro
    
    coConex.Ejecutar SQL
    

End Function
'
'Public Function CargaBilletaje(lsFecha As String, lsMoneda As String, lsCodAge As String, lnMoneda As String, Optional lsEfectivoCod As String = "") As ADODB.Recordset
'Dim psSql As String
'
'psSql = "SELECT m.cMovNro, mue.* " _
'    & "FROM   MovUserEfectivo MUE JOIN Mov M ON MUE.nMovNro = M.nMovNro " _
'    & "WHERE  M.nMovEstado = " & gMovEstContabNoContable & " and M.nMovFlag = " & gMovFlagVigente & " and " _
'    & "       m.cMovNro LIKE '" & lsFecha & "%' and " _
'    & "       m.cOpeCod = '" & gOpeHabCajRegEfect & "' and " _
'    & "       SubString(m.cMovNro,18,2) = '" & lsCodAge & "' " & IIf(lsEfectivoCod = "", "", " and cEfectivoCod = '" & lsEfectivoCod & "'")
'
'Set CargaBilletaje = coConex.CargaRecordSet(psSql)
'End Function
'
'Public Function CargaMovCta(Optional psMovNro As String = "", Optional psFiltro As String = "") As ADODB.Recordset
'Dim psSql As String
'   On Error GoTo CargaMovCtaErr
'      psSql = "SELECT cMovNro, nMovItem, mc.cCtaContCod, isnull(b.cCtaContDesc,'') AS cCtaContDesc, " _
'         & "       CONVERT(varchar(20),ISNULL(CASE WHEN nMovImporte > 0 THEN nMovImporte END,0),105) as nDebe, " _
'         & "       CONVERT(varchar(20),ISNULL(CASE WHEN nMovImporte < 0 THEN nMovImporte * -1 END,0),105) as nHaber " _
'         & " FROM  Mov M JOIN MovCta mc ON mc.nMovNro = m.nMovNro LEFT JOIN " & vsServerComunes & "CtaCont b ON mc.cCtaContCod = b.cCtaContCod " _
'         & IIf(psMovNro = "", "", "WHERE cMovNro = '" & psMovNro & "' ") _
'         & IIf(psFiltro = "", "", IIf(psMovNro = "", "WHERE ", "AND ")) & psFiltro
'      Set CargaMovCta = coConex.CargaRecordSet(psSql)
'   Exit Function
'CargaMovCtaErr:
'   Call RaiseError(MyUnhandledError, "DMov:CargaMovCta Method")
'End Function
Public Sub ActualizaMovCta_Cuenta(psCtaCodant As String, psCtaCodNew As String, psFecha1 As String, psFecha2 As String)
    Dim SQL As String

    SQL = "UPDATE mc SET cCtaContCod = '" & psCtaCodNew & "' FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro WHERE mc.cCtaContCod = '" & psCtaCodant & "' and LEFT(m.cMovNro,8) BETWEEN '" & psFecha1 & "' and '" & psFecha2 & "' "
    coConex.Ejecutar SQL


End Sub
Public Sub InsertaMovTpoCambio(ByVal pnMovNro As Long, ByVal pnMovTpoCambio As Currency)

    Dim SQL As String
    
    SQL = "INSERT INTO MovTpoCambio (nMovNro, nMovTpoCambio) " _
     & "VALUES (" & pnMovNro & ", " & pnMovTpoCambio & ")"
     
    coConex.Ejecutar SQL

End Sub
Public Function InsertaMov(ByVal psMovNro As String, _
                            ByVal psOpeCod As String, ByVal psMovDesc As String, _
                            Optional ByVal pnMovEstado As MovEstado = gMovEstContabMovContable, _
                            Optional ByVal pnMovFlag As MovFlag = gMovFlagVigente) As Integer

    Dim SQL As String
    
    SQL = "INSERT INTO MOV (cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag)" _
        & " VALUES('" & psMovNro & "','" & psOpeCod & "','" & Replace(psMovDesc, "'", "''") & "'," _
        & pnMovEstado & "," & pnMovFlag & ")"
     
    coConex.Ejecutar SQL
    

End Function
'Public Function GetnMovNro(ByVal psMovNro As String) As Long
'Dim sql As String
'Dim rs As ADODB.Recordset
'Set rs = New ADODB.Recordset
'
'sql = "Select Max(nMovNro) nMovNro From Mov where cMovNro ='" & psMovNro & "'"
'Set rs = coConex.CargaRecordSet(sql)
'If Not rs.EOF And Not rs.BOF Then
'    GetnMovNro = rs!nMovNro
'End If
'rs.Close
'Set rs = Nothing
'End Function
Public Function GetcMovNro(ByVal pnMovNro As Long) As String
Dim SQL As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

SQL = "Select cMovNro From Mov where nMovNro ='" & pnMovNro & "'"
Set rs = coConex.CargaRecordSet(SQL)
If Not rs.EOF And Not rs.BOF Then
    GetcMovNro = rs!cMovnro
End If
rs.Close
Set rs = Nothing
End Function
Public Sub EliminaMovRef(ByVal pnMovNro As Long)

    Dim SQL As String
    SQL = "DELETE MOVREF WHERE nMOVNRO =" & pnMovNro & ""
        
    coConex.Ejecutar SQL

End Sub
Public Function InsertaMovHabilitaciones(ByVal pnMovNro As Long, ByVal psCodTransp As String, _
    ByVal psAreaOrig As String, ByVal psAgeOrig As String, ByVal psAreaDest As String, _
    ByVal psAgeDest As String, ByVal pnMontoTrans As Double, ByVal pnMovImporte As Double, _
    ByVal nMoneda As Moneda) As Integer
    
    Dim SQL As String
    InsertaMovHabilitaciones = 1
    SQL = "INSERT INTO MovHabilitacion (nMovNro, cTranspCod, cAreaOrig, cAgeOrig, cAreaDest, cAgeDest, nMontoTrans, nMovImporte, nMoneda)" _
        & " VALUES(" & pnMovNro & ",'" & psCodTransp & "','" & psAreaOrig & "','" & psAgeOrig & "','" _
        & psAreaDest & "','" & psAgeDest & "'," & pnMontoTrans & "," & pnMovImporte & "," & nMoneda & ")"
     
    coConex.Ejecutar SQL
    InsertaMovHabilitaciones = 0

End Function

Public Function InsertaMovHabDev(ByVal pnMovNro As Long, ByVal pnImporte As Double, _
        ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal nMoneda As Moneda, _
        ByVal sUsuOrig As String, ByVal sUsuDest As String) As Integer

    Dim SQL As String
    InsertaMovHabDev = 1
    SQL = "Insert MovHabDevCajero(nMovNro,nMovImporte,cAgeCod,cAreaCod,nMoneda,cUsuOrig,cUsuDest) " _
        & "Values(" & pnMovNro & "," & pnImporte & ",'" & psAgeCod & "','" & psAreaCod & "'," & nMoneda & ",'" & sUsuOrig & "','" & sUsuDest & "')"
    coConex.Ejecutar SQL
    InsertaMovHabDev = 0
 
End Function

Public Function InsertaMovUserEfectivo(ByVal pnMovNro As Long, ByVal psCodUser As String, ByVal psEfectivoCod As String, _
                                        ByVal pnMonto As Currency) As Integer

    Dim SQL As String
    InsertaMovUserEfectivo = 1
     
    SQL = "INSERT INTO MovUserEfectivo(nMovNro, cUser,cEfectivoCod, nMonto) " _
        & " VALUES(" & pnMovNro & ",'" & psCodUser & "','" & psEfectivoCod & "'," & pnMonto & ")"

    coConex.Ejecutar SQL
    InsertaMovUserEfectivo = 0

End Function
Public Function ActualizaMovUserEfectivo(ByVal pnMovNro As Long, ByVal psCodUser As String, ByVal psEfectivoCod As String, _
                                         ByVal pnMonto As Currency) As Integer

    Dim SQL As String
    ActualizaMovUserEfectivo = 1
     
    SQL = "UPDATE MovUserEfectivo SET nMonto =" & pnMonto & " " _
        & " WHERE nMovNro =" & pnMovNro & " AND cUser='" & psCodUser & "' AND cEfectivoCod='" & psEfectivoCod & "'"

    coConex.Ejecutar SQL
    ActualizaMovUserEfectivo = 0

End Function
Public Sub ActualizaMovCont(ByVal pnMovNro As Long, Optional ByVal pnMovMonto As Currency, Optional ByVal pnMovCorrela As Integer, Optional ByVal psMovParalelo As String = "")
Dim sModif As String
Dim SQL    As String

    If pnMovMonto <> 0 Then
      sModif = "nMovMonto = " & pnMovMonto
    End If
    If pnMovCorrela <> 0 Then
      sModif = sModif & IIf(sModif = "", "", ",") & "nMovCorrela = " & pnMovCorrela
    End If
    If Not psMovParalelo = "" Then
      sModif = sModif & IIf(sModif = "", "", ",") & "cMovParalelo = '" & psMovParalelo
    End If
    If sModif = "" Then
       Err.Raise "50001", "DMov:ActualizaMovCont", "Indicar campos de Mov a Actualizar"
    Else
        SQL = "UPDATE MovCont SET " & sModif & " WHERE nMovNro = " & pnMovNro & ""
          
        coConex.Ejecutar SQL
    End If

End Sub
'Public Function GeneraMovNro(Optional ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional ByVal psMovNro As String = "", Optional pnCont As Integer = 0) As String
'    On Error GoTo GeneraMovNroErr
'    Dim rs As ADODB.Recordset
'    Dim sql As String
'    Set rs = New ADODB.Recordset
'    If pnCont = 0 And psMovNro <> "" Then
'        psMovNro = Left(psMovNro, 19) & Format(Val(Mid(psMovNro, 20, 2)) + 1, "00") & Right(psMovNro, 4)
'    End If
'    If psMovNro = "" Or Len(psMovNro) <> 25 Then
'       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & coConex.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & psCodAge & "','" & psUser & "'"
'    Else
'       sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
'    End If
'    Set rs = coConex.Ejecutar(sql)
'    If Not rs.EOF Then
'        GeneraMovNro = rs.Fields(0)
'    End If
'    rs.Close
'    Set rs = Nothing
'    Exit Function
'GeneraMovNroErr:
'    Call RaiseError(MyUnhandledError, "NContFunciones:GeneraMovNro Method")
'End Function
Public Function CargaRecordSet(ByVal SQL As String) As ADODB.Recordset
Set CargaRecordSet = coConex.CargaRecordSet(SQL)
End Function
Public Function InsertaMovCompraVenta(ByVal pnMovNro As Long, ByVal psPersCod As String, ByVal pnImporte As Currency) As Integer
    
    Dim SQL As String
    InsertaMovCompraVenta = 1
    
    SQL = " INSERT INTO MovCompraVenta(nMovNro, cPersCod, nMovImporte )" _
        & " VALUES(" & pnMovNro & ",'" & psPersCod & "'," & pnImporte & ") "
      
    coConex.Ejecutar SQL
    InsertaMovCompraVenta = 0

End Function

Public Function InsertaMovOpeVarias(ByVal pnMovNro As Long, ByVal psNroDoc As String, _
            ByVal psReferencia As String, ByVal pnImporte As Double, _
            Optional nMoneda As Moneda = gMonedaNacional) As Integer

    Dim SQL As String
    InsertaMovOpeVarias = 1
    
    SQL = " INSERT INTO MovOpeVarias(nMovNro, cNroDoc, cReferencia, nMovImporte,nMoneda)" _
        & " VALUES(" & pnMovNro & ",'" & psNroDoc & "','" & psReferencia & "'," & pnImporte & "," & nMoneda & ") "
      
    coConex.Ejecutar SQL
    InsertaMovOpeVarias = 0

End Function

Public Function InsertaNotaAbonoCargo(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pnEstado As NotaCargoAbonoEstado, _
                                    ByVal pnMotivo As Integer, ByVal pnMonto As Currency, Optional psObjetoCod As String = "", Optional psObjeto As String = "", Optional pnMoneda As Moneda = gMonedaNacional) As Integer

    Dim SQL As String
    InsertaNotaAbonoCargo = 1
    
    SQL = " INSERT INTO NotaAbonoCargo(nDocTpo, cDocNro, nEstado, nMotivoCod, nMonto, cObjetoCodPadre, cObjetoCod, nMoneda)" _
        & " VALUES(" & pnTpoDoc & ",'" & psNroDoc & "'," & pnEstado & "," & pnMotivo & "," & pnMonto & ",'" & psObjetoCod & "','" & psObjeto & "'," & pnMoneda & ") "
    
    coConex.Ejecutar SQL
    InsertaNotaAbonoCargo = 0

End Function
Public Function InsertaNotaAbonoCargoEst(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pnEstado As NotaCargoAbonoEstado, _
                                        ByVal psMovNro As String) As Integer

    Dim SQL As String
    InsertaNotaAbonoCargoEst = 1
    
    SQL = " INSERT INTO NotaAbonoCargoEst(nDocTpo, cDocNro, nEstado, cMovNro)" _
        & " VALUES(" & pnTpoDoc & ",'" & psNroDoc & "'," & pnEstado & ",'" & psMovNro & "') "
    
    coConex.Ejecutar SQL
    InsertaNotaAbonoCargoEst = 0
 
End Function

Public Function InsertaMovArqueo(ByVal pnMovNro As Long, ByVal psAreaCod As String, ByVal psPersCod As String, ByVal psHoraIni As String, _
                                ByVal psHoraFin As String, ByVal pnEfectivo As Currency, ByVal pnOPImporte As Currency, _
                                ByVal pnTotDoc As Currency, ByVal pnTotDocRend As Currency, ByVal pnArendirPend As Currency, _
                                ByVal pnDiferencia As Currency) As Integer
                                

    Dim SQL As String
    InsertaMovArqueo = 1
    
    SQL = " INSERT INTO MOVARQUEO(nMovNro, cAreaCod, cPersCod, cHoraIni, cHoraFin, " _
        & " nEfectivo, nOPImporte, nTotDoc, nTotDocRend, nArendirPend, nDiferencia)" _
        & " VALUES(" & pnMovNro & ",'" & psAreaCod & "','" & psPersCod & "','" & psHoraIni & "','" & psHoraFin & "'," _
        & pnEfectivo & "," & pnOPImporte & "," & pnTotDoc & "," & pnTotDocRend & "," & pnArendirPend & "," & pnDiferencia & ") "

    
    coConex.Ejecutar SQL
    InsertaMovArqueo = 0


End Function
Public Function InsertaMovArqueoOP(ByVal pnMovNro As Long, ByVal psNroOrden As String, ByVal pnImporte As String) As Integer

    Dim SQL As String
    InsertaMovArqueoOP = 1
    
    SQL = " INSERT INTO MovArqueoOp(nMovNro, cNroOrden, nImporte) " _
        & " VALUES(" & pnMovNro & ",'" & psNroOrden & "'," & pnImporte & ")"

    
    coConex.Ejecutar SQL
    InsertaMovArqueoOP = 0

End Function
Public Function InsertaMovArqueoEfectivo(ByVal pnMovNro As Long, ByVal psEfectivoCod As String, ByVal pnImporte As Currency) As Integer

    Dim SQL As String
    InsertaMovArqueoEfectivo = 1
     
    SQL = "INSERT INTO MovArqueoEfectivo(nMovNro, cEfectivoCod, nImporte ) " _
        & " VALUES(" & pnMovNro & ",'" & psEfectivoCod & "'," & pnImporte & ")"

    coConex.Ejecutar SQL
    InsertaMovArqueoEfectivo = 0

End Function
'
'Public Function AgregaNuevaCaptacion(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
'        ByVal sAgencia As String, ByVal nTasa As Double, ByVal nSaldo As Double, _
'        ByVal dFecha As Date, ByVal nFirmas As Integer, ByVal nPersoneria As PersPersoneria, _
'        ByVal nTipoCuenta As ProductoCuentaTipo, ByVal sMovNro As String, _
'        ByVal nTipoTasa As CaptacTipoTasa, Optional bCheque As Boolean = False, _
'        Optional ByVal bOrdPag As Boolean = False, Optional nPlazo As Long = 0, _
'        Optional nFormaRetiro As CaptacPFFormaRetiro = gCapPFFormRetMensual, _
'        Optional bCtaAboInt As Boolean = False, Optional sCtaCodAbono As String = "", _
'        Optional nPorcRetCTS As Double = 0, Optional sInstitucion As String = "") As String
'Dim clsGen As DCOMGeneral
'Dim sCuenta As String, sFecha As String
'Dim sFecUltCierre As String, sOrdPag As String, sCtaAbonoIntPF As String
'Dim nSaldoDisp As Double, nSaldRetCTS As Double
'Set clsGen = New DCOMGeneral
'sCuenta = clsGen.GeneraNuevaCuenta(sAgencia, nProducto, nMoneda)
'Set clsGen = Nothing
'sFecha = Format$(dFecha, gsFormatoFecha) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
'sSQL = "Insert Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) " _
'    & "Values ('" & sCuenta & "'," & nTasa & "," & nSaldo & "," & gCapEstActiva & ",'" & sFecha & "',0)"
'coConex.Ejecutar sSQL
'
'
'nSaldoDisp = IIf(bCheque, 0, nSaldo)
'sFecUltCierre = Format$(DateAdd("d", -1, dFecha), gsFormatoFecha)
'sSQL = "Insert Captaciones (cCtaCod,nSaldoDisp,nPersoneria,nFirmas,nIntAcum,dUltCierre,dApertura,nPrdCtaTpo,nPrdTasaInteres,cUltimaActualizacion) " _
'    & "Values ('" & sCuenta & "'," & nSaldoDisp & "," & nPersoneria & "," & nFirmas & ",0,'" & sFecUltCierre & "','" & sFecha & "'," & nTipoCuenta & "," _
'    & nTipoTasa & ",'" & sMovNro & "')"
'coConex.Ejecutar sSQL
'Select Case nProducto
'    Case gCapAhorros
'        sOrdPag = IIf(bOrdPag, "1", "0")
'        sSQL = "Insert CaptacAhorros (cCtaCod,nSaldoAnterior,bOrdPag,nSobregiro,dUltContacto) " _
'            & "Values ('" & sCuenta & "',0," & sOrdPag & ",0,'" & sFecha & "')"
'    Case gCapPlazoFijo
'        sCtaAbonoIntPF = IIf(bCtaAboInt, "1", "0")
'        sSQL = "Insert CaptacPlazoFijo (cCtaCod,nPlazo,nIntPag,dRenovacion,nApertura,dAuxiliar,nFormaRetiro,nDuplicado,bAbonoIntCtaAho) " _
'            & "Values ('" & sCuenta & "'," & nPlazo & ",0,'" & sFecha & "'," & nSaldo & ",'" & sFecha & "'," & nFormaRetiro & ",0," & sCtaAbonoIntPF & ")"
'    Case gCapCTS
'        nSaldRetCTS = Round(nSaldo * nPorcRetCTS / 100, 2)
'        sSQL = "Insert CaptacCTS (cCtaCod,nSaldRetiro,nIntSaldo,cCodInst) " _
'            & "Values ('" & sCuenta & "'," & nSaldRetCTS & ",0,'" & sInstitucion & "')"
'End Select
'coConex.Ejecutar sSQL
'If nProducto = gCapPlazoFijo And bCtaAboInt Then
'    sSQL = "Insert CaptacCtaAboIntPF (cCtaCod,cCtaCodAbono) " _
'        & "Values ('" & sCuenta & "','" & sCtaCodAbono & "')"
'    coConex.Ejecutar sSQL
'End If
'AgregaNuevaCaptacion = sCuenta
'End Function
'Public Sub AgregaNuevoProdPers(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona)
'sSQL = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
'    & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ")"
'coConex.Ejecutar sSQL
'End Sub
'
'Public Function GetDatosOrdenPago(ByVal sCuenta As String, ByVal sNroDoc As String) As Recordset
'Dim rsOP As Recordset
'sSQL = "SELECT E.cMovNro, D.nTpoDoc, D.cNroDoc, D.cCtaCod, D.nMonto, P.cPersNombre, E.cEstado " _
'    & "D.cIFCodPers FROM DocRecOPEst E INNER JOIN DocRecOP D INNER JOIN " & vsServerPersonas & "Persona P " _
'    & "ON D.cIFCodPers = P.cPersCod ON E.nTpoDoc = D.nTpoDoc AND E.cNroDoc = D.cNroDoc AND E.cCtaCod " _
'    & "= D.cCtaCod WHERE E.cMovNro IN (SELECT MAX(E1.cMovNro) FROM DocRecOPEst E WHERE E1.nTpoDoc = " _
'    & "E.nTpoDoc AND E1.cNroDoc = E.cNroDoc AND E1.cCtaCod = E.cCtaCod)"
'
'Set rsOP = New Recordset
'Set rsOP = coConex.CargaRecordSet(sSQL)
'Set GetDatosOrdenPago = rsOP
'End Function
'Public Sub AgregaOrdenPagoEstado(ByVal sCuenta As String, ByVal sNroDoc As String, _
'        ByVal sMovNro As String, ByVal nMonto As Double, ByVal nEstadoOP As CaptacOrdPagoEstado)
'sSQL = "UPDATE DocRecOP Set nMonto = " & nMonto & " WHERE nTpoDoc = " & TpoDocOrdenPago & " " _
'    & "AND cNroDoc = '" & sNroDoc & "' And cCtaCod = '" & sCuenta & "'"
'coConex.Ejecutar sSQL
'sSQL = "INSERT DocRecOPEst (nTpoDoc,cNroDoc,cCtaCod,cMovNro,nMonto,nEstado) " _
'    & "VALUES (" & TpoDocOrdenPago & ",'" & sNroDoc & "','" & sCuenta & "','" & sMovNro & "'," & nMonto & "," & nEstadoOP & ")"
'coConex.Ejecutar sSQL
'End Sub
'
'Public Sub AgregaCuentaDocumento(ByVal sCuenta As String, ByVal nTpoDoc As TpoDoc, _
'        ByVal sNroDoc As String, ByVal sCodIF As String, ByVal sMovNro As String, _
'        ByVal nMovNro As Long, Optional nMonto As Double = 0, _
'        Optional nEstadoOP As CaptacOrdPagoEstado = gCapOPEstCobrada)
'Dim sFecha As String
'sFecha = Mid(sMovNro, 1, 4) & "/" & Mid(sMovNro, 5, 2) & "/" & Mid(sMovNro, 7, 2) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
'Select Case nTpoDoc
'    Case TpoDocCheque
'        sSQL = "INSERT DocRecCapta (nTpoDoc,cNroDoc,cPersCod,cIFTpo,cCtaCod,nMonto) " _
'            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "','" & sCuenta & "'," & nMonto & ")"
'    Case TpoDocOrdenPago
'        sSQL = "INSERT DocRecOP (nTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers,nEstado,cMovNro) " _
'            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCuenta & "'," & nMonto & ",'" & sCodIF & "'," & nEstadoOP & ",'" & sMovNro & "')"
'    Case TpoDocNotaAbono, TpoDocNotaCargo
'End Select
'coConex.Ejecutar sSQL
'sSQL = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
'    & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
'coConex.Ejecutar sSQL
'End Sub
'Public Sub AgregaMovCap(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
'        ByVal sCuenta As String, ByVal nMonto As Double, _
'        ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double)
'sSQL = "INSERT MovCap (nMovNro,cOpeCod,cCtaCod,nMonto,nSaldoDisponible,nSaldoContable) " _
'    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nMonto & "," & nSaldoDisp & "," & nSaldoCnt & ")"
'coConex.Ejecutar sSQL
'End Sub
'Public Sub AgregaMovCapDet(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
'        ByVal sCuenta As String, ByVal nConcepto As CaptacConcepto, ByVal nMonto As Double)
'
'sSQL = "INSERT MovCapDet (nMovNro,cOpeCod,cCtaCod,nConceptoCod,nMonto) " _
'    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nConcepto & "," & nMonto & ")"
'coConex.Ejecutar sSQL
'End Sub
'Public Sub ActualizaSaldoAnteriorAho(ByVal sCuenta As String, ByVal nSaldoAnt As Double)
'sSQL = "Update CaptacAhorros Set nSaldoAnterior = " & nSaldoAnt & "  WHERE cCtaCod = '" & sCuenta & "'"
'coConex.Ejecutar sSQL
'End Sub
'Public Sub ActualizaAbonoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
'            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
'            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)
'
'If bActExtracto Then
'    sSQL = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
'        & "WHERE cCtaCod = '" & sCuenta & "'"
'Else
'    sSQL = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & " " _
'        & "WHERE cCtaCod = '" & sCuenta & "'"
'End If
'coConex.Ejecutar sSQL
'sSQL = "Update Captaciones Set nSaldoDisp = nSaldoDisp + " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
'    & ", dUltCierre = '" & Format$(dUltMov, gsFormatoFecha) & "', cUltimaActualizacion = '" & sMovNro & "' " _
'    & "WHERE cCtaCod = '" & sCuenta & "'"
'coConex.Ejecutar sSQL
'
'End Sub

'Public Sub ActualizaCargoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
'            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
'            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)
'
'If bActExtracto Then
'    sSQL = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
'        & "WHERE cCtaCod = '" & sCuenta & "'"
'Else
'    sSQL = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & " " _
'        & "WHERE cCtaCod = '" & sCuenta & "'"
'End If
'coConex.Ejecutar sSQL
'sSQL = "Update Captaciones Set nSaldoDisp = nSaldoDisp - " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
'    & ", dUltCierre = '" & Format$(dUltMov, gsFormatoFecha) & "', cUltimaActualizacion = '" & sMovNro & "' " _
'    & "WHERE cCtaCod = '" & sCuenta & "'"
'coConex.Ejecutar sSQL
'End Sub

'Public Sub UltimaActualizacionCuenta(ByVal sCuenta As String, ByVal sMovNro As String)
''Actualiza la fecha de ultima actualizacion de la cuenta de captaciones
'sSQL = "Update Captaciones Set cUltimaActualizacion = '" & sMovNro & "' " _
'    & "WHERE cCtaCod = '" & sCuenta & "'"
'coConex.Ejecutar sSQL
'End Sub
'Public Sub ActualizaEstadoCuenta(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, _
'        ByVal dFecha As Date, ByVal sMovNro As String)
''Actualiza el ultimo estado a la tabla de producto
'sSQL = "Update Producto Set nPrdEstado = " & nEstado & ", " _
'    & "dPrdEstado = '" & Format$(dFecha, gsFormatoFecha) & "' WHERE " _
'    & "cCtaCod = '" & sCuenta & "'"
'coConex.Ejecutar sSQL
''Agrega un registro mas de estado a la historia de estados de la cuenta
'sSQL = "INSERT CaptacEstado (cCtaCod,nPrdEstado,cMovNro) " _
'    & "VALUES ('" & sCuenta & "'," & nEstado & ",'" & sMovNro & "')"
'coConex.Ejecutar sSQL
'End Sub
'Public Sub ActualizaEstadoDocRecEst(ByVal nTipoDoc As TpoDoc, ByVal sNroDoc As String, _
'        ByVal sCodIF As String, ByVal sMovNro As String, ByVal nEstado As ChequeEstado)
''Actualiza el ultimo estado a la tabla de Documento Recibidos
'sSQL = "Update DocRec Set nEstado = " & nEstado & " WHERE " _
'    & "nTpoDoc = '" & nTipoDoc & "' AND cNroDoc = '" & sNroDoc & "' AND cPersCod = '" & sCodIF & "'"
'coConex.Ejecutar sSQL
''Agrega un registro mas de estado a la historia de estados del documento recibido
'sSQL = "INSERT DocRecEst (nTpoDoc,cNroDoc,cPersCod,cMovNro,nEstado) " _
'    & "VALUES ('" & nTipoDoc & "','" & sNroDoc & "','" & sCodIF & "','" & sMovNro & "'," & nEstado & ")"
'coConex.Ejecutar sSQL
'End Sub
'Public Function GetMovExtorno(ByVal sDatoBus As String, Optional nTipoBus As Integer = 0, Optional pdFecSis As Date) As Recordset
'Dim rsMov As Recordset
'Dim sWhere As String
'If nTipoBus = 0 Then
'    sWhere = "M.cMovNro LIKE '%" & sDatoBus & "%'"
'ElseIf nTipoBus = 1 Then
'    sWhere = "M.cMovNro LIKE '" & Format$(pdFecSis, "yyyymmdd") & "%' And C.cCtaCod = '" & sDatoBus & "'"
'End If
'sSQL = "Select M.cMovNro, O.cOpeDesc, C.cCtaCod, ISNULL(MD.nDocTpo,'') nDocTpo, ISNULL(MD.cDocNro,'') cDocNro, " _
'    & "M.cMovDesc, C.cOpeCod, C.nMonto, M.nMovNro FROM MovDoc MD RIGHT JOIN Mov M INNER JOIN MovCap C " _
'    & "INNER JOIN " & vsServerComunes & "OpeTpo O ON C.cOpeCod = O.cOpeCod ON M.nMovNro = C.nMovNro " _
'    & "ON MD.nMovNro = M.nMovNro WHERE " & sWhere & " AND C.cOpeCod NOT IN ('" & gAhoEstInacAct & "') " _
'    & "AND M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & "," & gMovFlagDeExtorno & ")" _
'    & "AND M.nMovEstado IN (" & gMovEstContabMovContable & ") " _
'    & "ORDER BY M.nMovNro"
'Set rsMov = New Recordset
'Set rsMov = coConex.CargaRecordSet(sSQL)
'Set GetMovExtorno = rsMov
'End Function
'
'Public Sub ActualizaMovDatos(ByVal pnMovNro As Long, Optional psMovNro As String = "", Optional pnMovEstado As MovEstado, Optional pnMovFlag As MovFlag)
'Dim lsFiltro As String
'If pnMovEstado <> -1 Then
'    lsFiltro = " nMovEstado  = " & pnMovEstado & ","
'End If
'If psMovNro <> "" Then
'    lsFiltro = lsFiltro & " cMovNro = '" & psMovNro & "',"
'End If
'If pnMovFlag <> -1 Then
'    lsFiltro = lsFiltro & " nMovFlag = " & pnMovFlag & ","
'End If
'If lsFiltro <> "" Then
'    lsFiltro = Left(lsFiltro, Len(lsFiltro) - 1)
'    sSQL = "UPDATE Mov SET " & lsFiltro & " WHERE nMovNro = " & pnMovNro
'    coConex.Ejecutar sSQL
'End If
'If pnMovEstado = gMovEstContabMovContable Then
'    ActualizaSaldoMovimiento psMovNro, "+"
'End If
'End Sub

'Public Sub ActualizaEstadoMov(ByVal nMovNro As Long, ByVal nMovFlag As MovFlag)
'sSQL = "Update Mov Set nMovFlag = " & nMovFlag & " Where nMovNro = " & nMovNro
'coConex.Ejecutar sSQL
'End Sub
'Public Function EsOrdenPagoEmitida(ByVal sCuenta As String, ByVal nNumOP As Long) As Boolean
'Dim rsOP As Recordset
'Set rsOP = New Recordset
'rsOP.CursorLocation = adUseClient
'sSQL = "Select OP.cCtaCod, OP.nInicio, OP.nFin FROM Mov M INNER JOIN MovDocEmitidoRango OP ON " _
'    & "M.nMovNro = OP.nMovNro WHERE M.nMovFlag <> " & gMovFlagExtornado & " AND OP.cCtaCod = '" _
'    & sCuenta & "' And " & nNumOP & " Between OP.nInicio And OP.nFin"
'
'Set rsOP = coConex.CargaRecordSet(sSQL)
'If rsOP.EOF And rsOP.BOF Then
'    EsOrdenPagoEmitida = False
'Else
'    EsOrdenPagoEmitida = True
'End If
'rsOP.Close
'Set rsOP = Nothing
'End Function
Public Function ActualizaNotaAbonoCargo(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, Optional ByVal pnEstado As NotaCargoAbonoEstado = -1, _
                                    Optional ByVal pnMotivo As Integer = -1, Optional ByVal pnMonto As Currency = -999999, _
                                    Optional psObjetoCod As String = "", Optional psObjeto As String = "") As Integer

    Dim SQL As String
    Dim lsActualiza As String
    ActualizaNotaAbonoCargo = 1
    lsActualiza = ""
    
    If pnMotivo <> -1 Then
        lsActualiza = lsActualiza + " nMotivoCod=" & pnMotivo & ", "
    End If
    If pnEstado <> -1 Then
        lsActualiza = lsActualiza + " nEstado=" & pnEstado & ","
    End If
    If pnMonto <> -999999 Then
        lsActualiza = lsActualiza + " nMonto=" & pnMonto & ","
    End If
    If psObjetoCod <> "" Then
        lsActualiza = lsActualiza + " cObjetoCodPadre= '" & psObjetoCod & "',"
    End If
    If psObjeto <> "" Then
        lsActualiza = lsActualiza + " cObjetoCod='" & psObjeto & "',"
    End If
    If lsActualiza <> "" Then
        lsActualiza = Mid(lsActualiza, 1, Len(lsActualiza) - 1)
        SQL = " UPDATE NotaAbonoCargo SET " & lsActualiza & " " _
            & " WHERE nDocTpo=" & pnTpoDoc & " AND cDocNro='" & psNroDoc & "'"
    End If
        
    coConex.Ejecutar SQL
    ActualizaNotaAbonoCargo = 0

End Function
'Public Function InsertaCuentaIF(ByVal psPersCod As String, ByVal psIFTpo As String, _
'                                ByVal psCtaIfCod As String, ByVal psCtaIFDesc As String, _
'                                ByVal pdCtaIFAper As Date, ByVal psCtaIFCap As String, ByVal psCtaIFInt As String, _
'                                ByVal pnCtaIFPlazo As Integer, ByVal pnCtaIFEstado As CGEstadoCtaIF, _
'                                ByVal psUltAct As String) As Integer
'    On Error GoTo InsertaCuentaIFErr
'    Dim sql As String
'    InsertaCuentaIF = 1
'
'    sql = "INSERT INTO CTAIF( cPersCod, cIFTpo, cCtaIFCod, cCtaIFDesc, " _
'        & "                dCtaIFAper, dCtaIFCap, nCtaIFPlazo, cCtaIFEstado, cUltimaActualizacion , dCtaIfInt ) " _
'        & " VALUES('" & psPersCod & "','" & psIFTpo & "','" & psCtaIfCod & "','" & psCtaIFDesc & "','" _
'        & Format(pdCtaIFAper, gsFormatoFecha) & "'," & IIf(psCtaIFCap = "", "Null", "'" & psCtaIFCap & "'") & "," _
'        & pnCtaIFPlazo & "," & pnCtaIFEstado & ",'" & psUltAct & "'," & IIf(psCtaIFInt = "", "Null", "'" & psCtaIFInt & "'") & ")"
'
'    coConex.Ejecutar sql
'    InsertaCuentaIF = 0
'    Exit Function
'InsertaCuentaIFErr:
'    Call RaiseError(MyUnhandledError, "DMov:InsertaCuentaIF Method")
'End Function
'Public Function InsertaCuentaIFInteres(ByVal psPersCod As String, ByVal psIFTpo As String, _
'                                ByVal psCtaIfCod As String, ByVal pdFechaInt As Date, ByVal pnPeriodo As Integer, _
'                                ByVal pnInteres As Currency, ByVal psMovNro As String) As Integer
'    On Error GoTo InsertaCuentaIFInteresErr
'    Dim sql As String
'    InsertaCuentaIFInteres = 1
'
'    sql = "INSERT INTO CtaIFInteres( cPersCod, cIFTpo, cCtaIFCod, dCtaIFIntRegistro , nCtaIFIntPeriodo, nCtaIFIntValor, cUltimaActualizacion  ) " _
'        & " VALUES('" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psCtaIfCod & "','" _
'        & Format(pdFechaInt, gsFormatoFecha) & "'," & pnPeriodo & "," & pnInteres & ",'" & psMovNro & "')"
'
'    coConex.Ejecutar sql
'    InsertaCuentaIFInteres = 0
'    Exit Function
'InsertaCuentaIFInteresErr:
'    Call RaiseError(MyUnhandledError, "DMov:InsertaCuentaIFInteres Method")
'End Function
'
'Public Function InsertaInteresCtaIF(ByVal psPersCod As String, ByVal psIFTpo As String, _
'                                    ByVal psCtaIfCod As String, _
'                                    ByVal pdCtaIFIntRegistro As Date, ByVal pnCtaIFIntPeriodo As Long, _
'                                    ByVal pnCtaIFIntValor As Currency, ByVal psUltAct As String) As Integer
'    On Error GoTo InsertaInteresCtaIFErr
'    Dim sql As String
'    InsertaInteresCtaIF = 1
'
'    sql = " INSERT INTO CtaIfInteres  " _
'        & " (cPersCod, cIFTpo, cCtaIFCod, dCtaIFIntRegistro, nCtaIFIntPeriodo, nCtaIFIntValor, cUltimaActualizacion) " _
'        & " VALUES('" & psPersCod & "','" & psIFTpo & "','" & psCtaIfCod & "','" _
'        & Format(pdCtaIFIntRegistro, gsFormatoFecha) & "'," & pnCtaIFIntPeriodo & "," _
'        & pnCtaIFIntValor & ",'" & psUltAct & "')"
'
'    coConex.Ejecutar sql
'    InsertaInteresCtaIF = 0
'    Exit Function
'InsertaInteresCtaIFErr:
'    Call RaiseError(MyUnhandledError, "DMov:InsertaInteresCtaIF Method")
'End Function

Public Function InsertaCuentaIFFiltro(ByVal psPersCod As String, ByVal psIFTpo As String, _
                                    ByVal psCtaIfCod As String, ByVal psCtaContCod As String, _
                                    ByVal psCtaIfSubCta As String) As Integer
                                    

    Dim SQL As String
    InsertaCuentaIFFiltro = 1
    
    SQL = "INSERT INTO CtaIFFiltro (cPersCod, cIFTpo, cCtaIfCod, cCtaContCod, cCtaIFSubCta)" _
        & " VALUES('" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psCtaIfCod & "','" _
        & psCtaContCod & "','" & psCtaIfSubCta & "')"
    
    coConex.Ejecutar SQL
    InsertaCuentaIFFiltro = 0

End Function

'Public Function InsertaMovCtaIF(ByVal pnMovNro As Long, ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, _
'                                 ByVal psCtaIfCod As String, ByVal pnConCod As CGCtaIFConceptos, ByVal pnImporte As Currency, Optional pnDiasTrans As Integer = 0) As Integer
'
'    On Error GoTo InsertaMovCtaIFErr
'    Dim sql As String
'    InsertaMovCtaIF = 1
'
'    sql = "INSERT INTO MovCtaIF(nMovNro, cPersCod, cIFTpo, cCtaIFCod, nConceptoCod, nMovImporte,nDiasTrans) " _
'        & "     VALUES(" & pnMovNro & ",'" & psPersCod & "','" & Format(pnIFTpo, "00") & "','" _
'        & psCtaIfCod & "'," & pnConCod & "," & pnImporte & "," & pnDiasTrans & ")"
'
'    coConex.Ejecutar sql
'    InsertaMovCtaIF = 0
'    Exit Function
'InsertaMovCtaIFErr:
'    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCtaIF Method")
'End Function

'Public Function ExtornaMovCtaIF(ByVal pnMovNro As Long) As Integer
'Dim sql As String
'Dim rs As adodb.Recordset
'Set rs = New adodb.Recordset
'
'sql = " SELECT  M.nMovNro,CONVERT(DATETIME , SUBSTRING(M.CMOVNRO,1,8)) AS FECHA, " _
'    & "         MC.cPersCod , MC.cIFTpo, MC.cCtaIFCod " _
'    & " FROM MOV M JOIN MOVCTAIF  MC ON MC.NMOVNRO =M.NMOVNRO  " _
'    & " WHERE M.NMOVNRO = " & pnMovNro & ""
'
'Set rs = coConex.CargaRecordSet(sql)
'If Not rs.EOF And Not rs.BOF Then
'    Do While Not rs.EOF
'        ActualizaSaldoCtaIF rs!Fecha, rs!cPersCod, rs!cIFTpo, rs!cCtaIFCod
'        rs.MoveNext
'    Loop
'End If
'rs.Close
'Set rs = Nothing
'End Function
'
'Public Function ActualizaSaldoCtaIF(ByVal pdFecha As Date, _
'                                    ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, _
'                                    ByVal psCtaIfCod As String) As Integer
'
'    On Error GoTo ActualizaSaldoCtaIFErr
'    Dim sql As String
'    ActualizaSaldoCtaIF = 1
'
'    sql = "sp_ActualizaSaldoCtaIF '" & Format(pdFecha, gsFormatoFecha) & "','" _
'                                    & psPersCod & "','" _
'                                    & Format(pnIFTpo, "00") & "', '" & psCtaIfCod & "'"
'
'    coConex.Ejecutar sql
'    ActualizaSaldoCtaIF = 0
'    Exit Function
'ActualizaSaldoCtaIFErr:
'    Call RaiseError(MyUnhandledError, "DMov:ActualizaSaldoCtaIF Method")
'End Function
Public Function ActualizaCtaIF(ByVal psPersCod As String, ByVal psIFTpo As String, _
                                  ByVal psCtaIfCod As String, _
                                  Optional ByVal psCtaIFDesc As String = "", _
                                  Optional ByVal pdCtaIFAper As String = "", _
                                  Optional ByVal pdCtaIFCap As String = "", _
                                  Optional ByVal pdCtaIFInt As String = "", _
                                  Optional ByVal pnCtaIFPlazo As Integer = -1, _
                                  Optional ByVal pnCtaIFEstado As CGEstadoCtaIF = -1, _
                                  Optional ByVal pnInteres As Currency = -999999, _
                                  Optional ByVal psUltAct As String = "") As Integer
                                  

    Dim SQL As String
    Dim lsFiltro As String
    ActualizaCtaIF = 1
    lsFiltro = ""
    If psCtaIFDesc <> "" Then
        lsFiltro = lsFiltro + " cCtaIFDesc ='" & psCtaIFDesc & "',"
    End If
    If pdCtaIFAper <> "" Then
        lsFiltro = lsFiltro + " dCtaIFAper= '" & pdCtaIFAper & "',"
    End If
    If pdCtaIFCap <> "" Then
        lsFiltro = lsFiltro + " dCtaIFCap= '" & pdCtaIFCap & "',"
    End If
    If pdCtaIFInt <> "" Then
        lsFiltro = lsFiltro + " dCtaIFInt= '" & pdCtaIFInt & "',"
    End If
    If pnCtaIFPlazo <> -1 Then
        lsFiltro = lsFiltro + " nCtaIFPlazo =" & pnCtaIFPlazo & ","
    End If
    If pnCtaIFEstado <> -1 Then
        lsFiltro = lsFiltro + " cCtaIFEstado =" & pnCtaIFEstado & ","
    End If
    If pnInteres <> -999999 Then
        lsFiltro = lsFiltro + " nInteres =" & pnInteres & ", "
    End If
    If psUltAct <> "" Then
        lsFiltro = lsFiltro + " cUltimaActualizacion ='" & psUltAct & "',"
    End If
    If lsFiltro <> "" Then
        lsFiltro = Mid(lsFiltro, 1, Len(lsFiltro) - 1)
        SQL = "UPDATE CTAIF SET " & lsFiltro _
            & " WHERE cPersCod = '" & psPersCod & "' AND cIFTpo='" & Format(psIFTpo, "00") & "' AND cCtaIFCod='" & psCtaIfCod & "'"
    End If
    coConex.Ejecutar SQL
    ActualizaCtaIF = 0

End Function

'
'Public Sub InsertaEfectivoAutomatico(ByVal pdFecha As Date, ByVal pcCodAge As String, ByVal pcUser As String)
'
'On Error GoTo InsertaEfectivoAutomaticoErr
'    Dim sql As String
'
'    sql = "INSERT INTO EfectivoAutomatico (dFecha, cCodAge, cUser, nFlag )" _
'        & " VALUES('" & Format(pdFecha, gsFormatoFecha) & "','" & Trim(pcCodAge) & "','" & Trim(pcUser) & "', 0)"
'
'    coConex.Ejecutar sql
'
'    Exit Sub
'InsertaEfectivoAutomaticoErr:
'    Call RaiseError(MyUnhandledError, "DMov:InsertaEfectivoAutomatico Method")
'
'End Sub


Public Sub ActualizaSaldoMovimiento(sMovNro As String, sSimbolo As String)
 Dim SQL As String
' Dim prs As adodb.Recordset
 Dim ldFecha As Date
' Dim ldFecSis As Date
 Dim oGen As COMDConstSistema.DCOMGeneral
 
 Set oGen = New COMDConstSistema.DCOMGeneral
 
 ldFecha = oGen.GetFechaMov(Left(sMovNro, 8), True)
 
 Set oGen = Nothing
 
 SQL = "spActualizaMovSaldo '" & sMovNro & "', '" & Format(ldFecha, "mm/dd/yyyy") & "', '" & sSimbolo & "' "
 coConex.Ejecutar SQL
 
 'CTASALDO
' sql = "SELECT left(m.cMovNro,8) as fecha, mc.cCtaContcod, cls.cCtaCaracter, SUM(mc.nMovImporte) as nImporte, ISNULL(SUM(me.nMovMeImporte),0) as nImporteME, sdo.dCtaSaldoFecha, ISNULL(sdo.nCtaSaldoImporte,0) nCtaSaldoImporte, ISNULL(sdo.nCtaSaldoImporteME,0) nCtaSaldoImporteME " _
'     & "FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro JOIN CtaContClase cls ON mc.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' " _
'     & "     LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'     & "     LEFT JOIN (SELECT cCtaContCod, dCtaSaldoFecha, nCtaSaldoImporte, nCtaSaldoImporteME FROM CtaSaldo cs WHERE cs.dCtaSaldoFecha = (SELECT Max(dCtaSaldoFecha) FROM CtaSaldo where cCtaContCod = cs.cCtaContCod and dCtaSaldoFecha <= '" & Format(ldFecha, gsFormatoFecha) & "') " _
'     & "               ) sdo ON sdo.cCtaContCod = mc.cCtaContCod " _
'     & "WHERE m.nMovEstado = " & gMovEstContabMovContable & " and m.cMovNro like '" & sMovNro & "%' and mc.cCtaContCod <> '' " _
'     & "GROUP BY left(m.cmovnro,8), mc.cCtaContCod, cls.cCtaCaracter, sdo.dCtaSaldoFecha, sdo.nCtaSaldoImporte, sdo.nCtaSaldoImporteME " _
'     & "ORDER BY left(m.cmovnro,8) "
' Set prs = coConex.CargaRecordSet(sql)
' If Not prs.EOF Then
'   Do While Not prs.EOF
'     If sSimbolo = "+" Then
'        If ldFecha <> prs!dCtaSaldofecha Then
'           sql = "INSERT CtaSaldo (cCtaContCod, dCtaSaldoFecha, nCtaSaldoImporte, nCtaSaldoImporteME) " _
'               & "VALUES ('" & prs!cCtaContCod & "', '" & Format(ldFecha, gsFormatoFecha) & "', " & prs!nCtaSaldoImporte & ", " & prs!nCtaSaldoImporteME & ")"
'           coConex.Ejecutar sql
'         End If
'     End If
'    sql = "UPDATE CtaSaldo SET nCtaSaldoImporte = nCtaSaldoImporte " & sSimbolo & "(" & prs!nImporte * IIf(prs!cCtaCaracter = "A", -1, 1) & "), " _
'        & "nCtaSaldoImporteME = nCtaSaldoImporteME " & sSimbolo & "(" & prs!nImporteME * IIf(prs!cCtaCaracter = "A", -1, 1) & ") " _
'        & "WHERE  cCtaContCod = '" & prs!cCtaContCod & "' and dCtaSaldoFecha >= '" & Format(ldFecha, gsFormatoFecha) & "'"
'    coConex.Ejecutar sql
'     prs.MoveNext
'   Loop
' End If
'
' 'CTAOBJSALDO
' sql = "SELECT left(m.cmovnro,8) as fecha, mc.cCtaContcod, mo.nMovObjOrden, mo.cObjetoCod, cls.cCtaCaracter, SUM(mc.nMovImporte) as nImporte, ISNULL(SUM(me.nMovMeImporte),0) as nImporteME, sdo.dCtaObjSaldoFecha, ISNULL(sdo.nCtaObjSaldoImporte,0) nCtaObjSaldoImporte, ISNULL(sdo.nCtaObjSaldoImporteME,0) nCtaObjSaldoImporteMe " _
'     & "FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro JOIN CtaContClase cls ON mc.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%'" _
'     & "     LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'     & "     JOIN MovObj mo ON mo.nMovNro = mc.nMovNro and mo.nMovItem = mc.nMovItem " _
'     & "     LEFT JOIN (SELECT cCtaContCod, nCtaObjOrden, cObjetoCod, dCtaObjSaldoFecha, nCtaObjSaldoImporte, nCtaObjSaldoImporteME " _
'     & "                FROM CtaObjSaldo cs WHERE cs.dCtaObjSaldoFecha = (SELECT Max(dCtaObjSaldoFecha) FROM CtaObjSaldo WHERE cCtaContCod = cs.cCtaContCod and nCtaObjOrden = cs.nCtaObjOrden and cObjetoCod = cs.cObjetoCod and dCtaObjSaldoFecha <= '" & Format(ldFecha, gsFormatoFecha) & "') " _
'     & "               ) sdo ON sdo.cCtaContCod = mc.cCtaContCod and sdo.nCtaObjOrden = mo.nMovObjOrden and sdo.cObjetoCod = mo.cObjetoCod " _
'     & "WHERE m.nMovEstado = " & gMovEstContabMovContable & " and m.cMovNro like '" & sMovNro & "%' and mc.cCtaContCod <> '' " _
'     & "GROUP BY left(m.cMovNro,8), mc.cCtaContCod, mo.nMovObjOrden, mo.cObjetoCod, cls.cCtaCaracter, sdo.dCtaObjSaldoFecha, ISNULL(sdo.nCtaObjSaldoImporte,0), ISNULL(sdo.nCtaObjSaldoImporteME,0) " _
'     & "ORDER BY left(m.cMovNro,8) "
' Set prs = coConex.CargaRecordSet(sql)
'
' If Not prs.EOF Then
'   Do While Not prs.EOF
'     ldFecha = GetFechaMov(prs!Fecha, True)
'     If sSimbolo = "+" Then
'        If prs!dCtaObjSaldoFecha <> ldFecha Then
'            sql = "INSERT CtaObjSaldo (cCtaContCod, nCtaObjOrden, cObjetoCod, dCtaObjSaldoFecha, nCtaObjSaldoImporte, nCtaObjSaldoImporteME) " _
'                & " VALUES ('" & prs!cCtaContCod & "','" & prs!nMovObjOrden & "', '" & prs!cObjetoCod & "', '" & Format(ldFecha, gsFormatoFecha) & "', " & prs!nCtaObjSaldoImporte & "," & prs!nCtaObjSaldoImporteME & " )"
'            coConex.Ejecutar sql
'        End If
'     End If
'     sql = "UPDATE CtaObjSaldo SET nCtaObjSaldoImporte = nCtaObjSaldoImporte " & sSimbolo & " (" & prs!nImporte * IIf(prs!cCtaCaracter = "A", -1, 1) & "), " _
'         & "nCtaObjSaldoImporteME = nCtaObjSaldoImporteME " & sSimbolo & "(" & prs!nImporteME * IIf(prs!cCtaCaracter = "A", -1, 1) & ") " _
'         & "WHERE  cCtaContCod = '" & prs!cCtaContCod & "' and nCtaObjOrden = " & prs!nMovObjOrden & " and cObjetoCod = '" & prs!cObjetoCod & "' and dCtaObjSaldoFecha >= '" & Format(ldFecha, gsFormatoFecha) & "' "
'     coConex.Ejecutar sql
'     prs.MoveNext
'   Loop
' End If
'
' 'CtaObjEfectivoSaldo
' sql = "SELECT LEFT(m.cMovNro,8) as fecha, mc.cCtaContcod, moef.nMovObjOrden, moef.cEfectivoCod, cls.cCtaCaracter, SUM(mc.nMovImporte) as nImporte, ISNULL(SUM(me.nMovMeImporte),0) as nImporteME, sdo.dEfectivoSaldoFecha, ISNULL(sdo.nEfectivoSaldoImporte,0) nEfectivoSaldoImporte, ISNULL(sdo.nEfectivoSaldoImporteME,0) nEfectivoSaldoImporteMe " _
'     & "FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro JOIN CtaContClase cls ON mc.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%'" _
'     & "     LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'     & "          JOIN MovObjEfectivo moef ON moef.nMovNro = mc.nMovNro and moef.nMovItem = mc.nMovItem " _
'     & "     LEFT JOIN (SELECT cCtaContCod, nCtaObjOrden, cEfectivoCod, dEfectivoSaldoFecha, nEfectivoSaldoImporte, nEfectivoSaldoImporteME " _
'     & "                FROM CtaObjEfectivoSaldo cs WHERE cs.dEfectivoSaldoFecha = (SELECT Max(dEfectivoSaldoFecha) FROM CtaObjEfectivoSaldo WHERE cCtaContCod = cs.cCtaContCod and nCtaObjOrden = cs.nCtaObjOrden and cEfectivoCod = cs.cEfectivoCod and dEfectivoSaldoFecha <= '" & Format(ldFecha, gsFormatoFecha) & "') " _
'     & "               ) sdo ON sdo.cCtaContCod = mc.cCtaContCod and sdo.nCtaObjOrden = moef.nMovObjOrden and sdo.cEfectivoCod = moef.cEfectivoCod " _
'     & "WHERE m.nMovEstado = " & gMovEstContabMovContable & " and m.cMovNro like '" & sMovNro & "%' and mc.cCtaContCod <> '' " _
'     & "GROUP BY left(m.cMovNro,8), mc.cCtaContCod, moef.nMovObjOrden, moef.cEfectivoCod, cls.cCtaCaracter, sdo.dEfectivoSaldoFecha, ISNULL(sdo.nEfectivoSaldoImporte,0), ISNULL(sdo.nEfectivoSaldoImporteME,0) " _
'     & "ORDER BY left(m.cMovNro,8) "
' Set prs = coConex.CargaRecordSet(sql)
'
' If Not prs.EOF Then
'   Do While Not prs.EOF
'     ldFecha = GetFechaMov(prs!Fecha, True)
'     If sSimbolo = "+" Then
'        If prs!dEfectivoSaldoFecha <> ldFecha Then
'            sql = "INSERT CtaObjEfectivoSaldo (cCtaContCod, nCtaObjOrden, cEfectivoCod, dEfectivoSaldoFecha, nEfectivoSaldoImporte, nEfectivoSaldoImporteME) " _
'                & " VALUES ('" & prs!cCtaContCod & "','" & prs!nMovObjOrden & "', '" & prs!cEfectivoCod & "', '" & Format(ldFecha, gsFormatoFecha) & "', " & prs!nEfectivoSaldoImporte & "," & prs!nEfectivoSaldoImporteME & " )"
'            coConex.Ejecutar sql
'        End If
'     End If
'     sql = "UPDATE CtaObjEfectivoSaldo SET nEfectivoSaldoImporte = nEfectivoSaldoImporte " & sSimbolo & " (" & prs!nImporte * IIf(prs!cCtaCaracter = "A", -1, 1) & "), " _
'         & "nEfectivoSaldoImporteME = nEfectivoSaldoImporteME " & sSimbolo & "(" & prs!nImporteME * IIf(prs!cCtaCaracter = "A", -1, 1) & ") " _
'         & "WHERE  cCtaContCod = '" & prs!cCtaContCod & "' and nCtaObjOrden = " & prs!nMovObjOrden & " and cEfectivoCod = '" & prs!cEfectivoCod & "' and dEfectivoSaldoFecha >= '" & Format(ldFecha, gsFormatoFecha) & "' "
'     coConex.Ejecutar sql
'     prs.MoveNext
'   Loop
' End If
'
' 'CtaObjAreaAgenciaSaldo
' sql = "SELECT LEFT(m.cMovNro,8) as fecha, mc.cCtaContcod, moef.nMovObjOrden, moef.cAreaCod, moef.cAgeCod, cls.cCtaCaracter, SUM(mc.nMovImporte) as nImporte, ISNULL(SUM(me.nMovMeImporte),0) as nImporteME, sdo.dAreaSaldoFecha, ISNULL(sdo.nAreaSaldoImporte,0) nAreaSaldoImporte, ISNULL(sdo.nAreaSaldoImporteME,0) nAreaSaldoImporteMe " _
'     & "FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro JOIN CtaContClase cls ON mc.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%'" _
'     & "     LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'     & "          JOIN MovObjAreaAgencia moef ON moef.nMovNro = mc.nMovNro and moef.nMovItem = mc.nMovItem " _
'     & "     LEFT JOIN (SELECT cCtaContCod, nCtaObjOrden, cAreaCod, cAgeCod, dAreaSaldoFecha, nAreaSaldoImporte, nAreaSaldoImporteME " _
'     & "                FROM CtaObjAreaAgenciaSaldo cs WHERE cs.dAreaSaldoFecha = (SELECT Max(dAreaSaldoFecha) FROM CtaObjAreaAgenciaSaldo WHERE cCtaContCod = cs.cCtaContCod and nCtaObjOrden = cs.nCtaObjOrden and cAreaCod = cs.cAreaCod and cAgeCod = cs.cAgeCod and dAreaSaldoFecha <= '" & Format(ldFecha, gsFormatoFecha) & "') " _
'     & "               ) sdo ON sdo.cCtaContCod = mc.cCtaContCod and sdo.nCtaObjOrden = moef.nMovObjOrden and sdo.cAreaCod = moef.cAreaCod and sdo.cAgeCod = moef.cAgeCod " _
'     & "WHERE m.nMovEstado = " & gMovEstContabMovContable & " and m.cMovNro like '" & sMovNro & "%' and mc.cCtaContCod <> '' " _
'     & "GROUP BY left(m.cMovNro,8), mc.cCtaContCod, moef.nMovObjOrden, moef.cAreaCod, moef.cAgeCod, cls.cCtaCaracter, sdo.dAreaSaldoFecha, ISNULL(sdo.nAreaSaldoImporte,0), ISNULL(sdo.nAreaSaldoImporteME,0) " _
'     & "ORDER BY left(m.cMovNro,8) "
' Set prs = coConex.CargaRecordSet(sql)
'
' If Not prs.EOF Then
'   Do While Not prs.EOF
'     ldFecha = GetFechaMov(prs!Fecha, True)
'     If sSimbolo = "+" Then
'        If prs!dAreaSaldoFecha <> ldFecha Then
'            sql = "INSERT CtaObjAreaAgenciaSaldo (cCtaContCod, nCtaObjOrden, cAreaCod, cAgeCod, dAreaSaldoFecha, nAreaSaldoImporte, nAreaSaldoImporteME) " _
'                & " VALUES ('" & prs!cCtaContCod & "','" & prs!nMovObjOrden & "', '" & prs!cAreaCod & "','" & prs!CAGECOD & "', '" & Format(ldFecha, gsFormatoFecha) & "', " & prs!nAreaSaldoImporte & "," & prs!nAreaSaldoImporteME & " )"
'            coConex.Ejecutar sql
'        End If
'     End If
'     sql = "UPDATE CtaObjAreaAgenciaSaldo SET nAreaSaldoImporte = nAreaSaldoImporte " & sSimbolo & " (" & prs!nImporte * IIf(prs!cCtaCaracter = "A", -1, 1) & "), " _
'         & "nAreaSaldoImporteME = nAreaSaldoImporteME " & sSimbolo & "(" & prs!nImporteME * IIf(prs!cCtaCaracter = "A", -1, 1) & ") " _
'         & "WHERE  cCtaContCod = '" & prs!cCtaContCod & "' and nCtaObjOrden = " & prs!nMovObjOrden & " and cAreaCod = '" & prs!cAreaCod & "' and cAgeCod = '" & prs!CAGECOD & "' and dAreaSaldoFecha >= '" & Format(ldFecha, gsFormatoFecha) & "' "
'     coConex.Ejecutar sql
'     prs.MoveNext
'   Loop
' End If
'
' 'CtaIFSaldo
' sql = "SELECT LEFT(m.cMovNro,8) as fecha, mc.cCtaContcod, moif.cPersCod, moif.cIFTpo, moif.cCtaIFCod, cls.cCtaCaracter, SUM(mc.nMovImporte) as nImporte, ISNULL(SUM(me.nMovMeImporte),0) as nImporteME, sdo.dCtaIFSaldo, ISNULL(sdo.nSaldo,0) nSaldo, ISNULL(sdo.nSaldoME,0) nSaldoME " _
'     & "FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro JOIN CtaContClase cls ON mc.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%'" _
'     & "     LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'     & "          JOIN MovObjIF moif ON moif.nMovNro = mc.nMovNro and moif.nMovItem = mc.nMovItem " _
'     & "     LEFT JOIN (SELECT cCtaContCod, cPersCod, cIFTpo, cCtaIFCod, dCtaIFSaldo, nSaldo, nSaldoME " _
'     & "                FROM CtaIFSaldo cs WHERE cs.dCtaIFSaldo = (SELECT Max(dCtaIFSaldo) FROM CtaIFSaldo WHERE cCtaContCod = cs.cCtaContCod and cPersCod = cs.cPersCod and cIFTpo = cs.cIFTpo and cCtaIFCod = cs.cCtaIFCod and dCtaIFSaldo <= '" & Format(ldFecha, gsFormatoFecha) & "') " _
'     & "               ) sdo ON sdo.cCtaContCod = mc.cCtaContCod and sdo.cPersCod = moif.cPersCod and sdo.cIFTpo = moif.cIFTpo and sdo.cCtaIFCod = moif.cCtaIFCod " _
'     & "WHERE m.nMovEstado = " & gMovEstContabMovContable & " and m.cMovNro LIKE '" & sMovNro & "%' and mc.cCtaContCod <> '' " _
'     & "GROUP BY left(m.cMovNro,8), mc.cCtaContCod, moif.cPersCod, moif.cIFTpo, moif.cCtaIFCod, cls.cCtaCaracter, sdo.dCtaIFSaldo, ISNULL(sdo.nSaldo,0), ISNULL(sdo.nSaldoME,0) " _
'     & "ORDER BY left(m.cMovNro,8) "
' Set prs = coConex.CargaRecordSet(sql)
' If Not prs.EOF Then
'   Do While Not prs.EOF
'     ldFecha = GetFechaMov(prs!Fecha, True)
'     If sSimbolo = "+" Then
'        If prs!dCtaIFSaldo <> ldFecha Then
'            sql = "INSERT CtaIFSaldo (cCtaContCod, cPersCod, cIFTpo, cCtaIFCod, dCtaIFSaldo, nSaldo, nSaldoME) " _
'                & " VALUES ('" & prs!cCtaContCod & "','" & prs!cPersCod & "','" & prs!cIFTpo & "', '" & prs!cCtaIFCod & "', '" & Format(ldFecha, gsFormatoFecha) & "', " & prs!nSaldo & "," & prs!nSaldoME & " )"
'            coConex.Ejecutar sql
'        End If
'     End If
'     sql = "UPDATE CtaIFSaldo SET nSaldo = nSaldo " & sSimbolo & " (" & prs!nImporte * IIf(prs!cCtaCaracter = "A", -1, 1) & "), " _
'         & "nSaldoME = nSaldoME " & sSimbolo & "(" & prs!nImporteME * IIf(prs!cCtaCaracter = "A", -1, 1) & ") " _
'         & "WHERE  cCtaContCod = '" & prs!cCtaContCod & "' and cPersCod = '" & prs!cPersCod & "' and cIFTpo = '" & prs!cIFTpo & "' and cCtaIFCod = '" & prs!cCtaIFCod & "' and dCtaIFSaldo >= '" & Format(ldFecha, gsFormatoFecha) & "' "
'     coConex.Ejecutar sql
'     prs.MoveNext
'   Loop
' End If
' RSClose prs
'
End Sub

'Public Sub InsertaDatosAdeudados(prsAdeud As ADODB.Recordset, psPersCod As String, psIFTpo As String, psCodAdeudado As String, pnCapital As Currency, pnNroCuotas As Integer, pnGracia As Currency, pnComisionIni As Currency, psInterno As String, pnCuotaPagoK As Currency, _
'                psFecha As String, pnMoneda As Integer)
'    Dim lbExiste As Boolean
'    Dim lsTipoCuota As String
'    Dim oDAdeudCal As New DAdeudCal
'    oDAdeudCal.Inicio coConex.ConexionActiva
'    lbExiste = oDAdeudCal.VerificaCtaIFAdeudado(psPersCod, psIFTpo, psCodAdeudado)
'    If lbExiste Then
'        oDAdeudCal.EliminarCtaIFAdeudado psPersCod, psIFTpo, psCodAdeudado
'    End If
'
'    oDAdeudCal.InsertaCtaIFAdeudado psPersCod, psIFTpo, psCodAdeudado, pnCapital, pnCapital, pnNroCuotas, _
'                pnGracia, pnComisionIni, psInterno, pnCuotaPagoK, _
'                psFecha, "1", pnMoneda, psFecha
'
'    'GRABAMOS EL CALENDARIO DE LOS ADEUDADOS
'    If Not prsAdeud Is Nothing Then
'        If Not prsAdeud.EOF Then
'            'GRABAMOS LOS POSIBLES PAGOS A REALIZARSE O PAGOS PROYECTADOS
'            Do While Not prsAdeud.EOF
'                'TipoCuota  : ---> A  Cuota Adelantada  ----> C Cuota Normal
'                lsTipoCuota = IIf(prsAdeud.Fields(1) = 0, "A", "C")
'                oDAdeudCal.InsertCalendario psPersCod, psIFTpo, psCodAdeudado, "A", prsAdeud.Fields(1), prsAdeud.Fields(0), "P", 0, prsAdeud.Fields(2), prsAdeud.Fields(3), 0
'                If prsAdeud.Fields(1) = 1 Then 'Si es primera Cuota
'                   oDAdeudCal.UpdateCtaIFAdeudado psPersCod, psIFTpo, psCodAdeudado, pnCapital, pnCapital, _
'                                                  pnNroCuotas, pnGracia, pnComisionIni, psInterno, pnCuotaPagoK, _
'                                                    Format(prsAdeud.Fields(0), gsFormatoFecha), "1", pnMoneda
'                End If
'                prsAdeud.MoveNext
'            Loop
'        End If
'    End If
'    Set oDAdeudCal = Nothing
'End Sub

'Public Sub ActualizaAdeudadosProvision(psPersCod As String, psIFTpo As String, psCtaIfCod As String, psFecha As String, pnInteresPag As Currency, pnNroCuota As Integer, psMovNro As String)
'Dim oDAdeudCal As New DAdeudCal
'    oDAdeudCal.Inicio coConex.ConexionActiva
'    oDAdeudCal.UpdateCtaIFAdeudado psPersCod, psIFTpo, psCtaIfCod, , , , , , , , , , , psFecha
'    oDAdeudCal.UpdateCalendario psPersCod, psIFTpo, psCtaIfCod, pnNroCuota, pnInteresPag, psMovNro
'Set oDAdeudCal = Nothing
'End Sub
            
'Public Sub ActualizaAdeudadosPagoCuota(psPersCod As String, psIFTpo As String, psCtaIfCod As String, psFecha As String, pnCapital As Currency, pnInteresPag As Currency, pnNroCuota As Integer, psMovNro As String, Optional pbCancela As Boolean = False)
'Dim oDAdeudCal As New DAdeudCal
'    oDAdeudCal.Inicio coConex.ConexionActiva
'    oDAdeudCal.UpdateCtaIFAdeudado psPersCod, psIFTpo, psCtaIfCod, , pnCapital, , , , , , , , , psFecha
'    oDAdeudCal.UpdateCalendario psPersCod, psIFTpo, psCtaIfCod, pnNroCuota, pnInteresPag, psMovNro, gTpoEstCuotaAdeudCanc
'Set oDAdeudCal = Nothing
'End Sub
      

Public Sub GrabaMovimientoEfectivo(ByVal pnMovNro As Long, ByRef pnMovItem As Long, rsEfectivo As ADODB.Recordset, psCuenta As String, psDH As String)
On Error GoTo GrabaMovimientoEfectivoErr
Dim lnMovOrden As Integer
'Guardamos el Billetaje
If Not rsEfectivo Is Nothing Then
    If rsEfectivo.State = adStateOpen Then
        If Not rsEfectivo.EOF And Not rsEfectivo.BOF Then
            Do While Not rsEfectivo.EOF
                If rsEfectivo!Monto <> 0 Then
                    pnMovItem = pnMovItem + 1: lnMovOrden = 0
                    InsertaMovCta pnMovNro, Format(pnMovItem, "#0"), psCuenta, rsEfectivo!Monto * IIf(psDH = "D", 1, -1)
                    lnMovOrden = lnMovOrden + 1
                    InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, ObjDescomEfectivo
                    InsertaMovObjEfectivo pnMovNro, pnMovItem, lnMovOrden, rsEfectivo!cEfectivoCod
                End If
                rsEfectivo.MoveNext
            Loop
        End If
        rsEfectivo.Close: Set rsEfectivo = Nothing
    End If
End If
Exit Sub
GrabaMovimientoEfectivoErr:
    Err.Raise Err.Number, "DMov: GrabaMovimientoEfectivo", Err.Description
End Sub

Public Sub GrabaMovOtrosCtaObj(pnMovNro As Long, pnMovItem As Long, prsOtro As ADODB.Recordset, prsObj As ADODB.Recordset, psDH As String)
Dim lnMovOrden As Long
Dim lsFiltro   As String

If Not prsOtro Is Nothing Then
    If prsOtro.State = adStateOpen Then
        Do While Not prsOtro.EOF
            pnMovItem = pnMovItem + 1: lnMovOrden = 0
            lsFiltro = ""
            If Not prsObj Is Nothing Then
                If prsObj.State = adStateOpen Then
                    Do While Not prsObj.EOF
                        If prsObj![#] = prsOtro![#] Then
                            lsFiltro = lsFiltro & prsObj!SubCta
                        End If
                        prsObj.MoveNext
                    Loop
                End If
            End If
            InsertaMovCta pnMovNro, pnMovItem, prsOtro!Cuenta & lsFiltro, prsOtro!Importe * IIf(psDH = "D", 1, -1)
            lsFiltro = ""
            If Not prsObj Is Nothing Then
                prsObj.MoveFirst
                lnMovOrden = 0
                Do While Not prsObj.EOF
                    If prsObj.Fields(0) = prsOtro.Fields(0) Then
                        Select Case prsObj!Objpadre
                            Case ObjCMACAgencias
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!Objpadre
                                InsertaMovObjAgenciaArea pnMovNro, pnMovItem, lnMovOrden, prsObj!Código, ""
                            Case ObjCMACArea
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!Objpadre
                                InsertaMovObjAgenciaArea pnMovNro, pnMovItem, lnMovOrden, "", prsObj!Código
                            Case ObjCMACAgenciaArea
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!Objpadre
                                InsertaMovObjAgenciaArea pnMovNro, pnMovItem, lnMovOrden, Mid(prsObj!Código, 4, 2), Mid(prsObj!Código, 1, 3)
                            Case ObjDescomEfectivo
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!Objpadre
                                InsertaMovObjEfectivo pnMovNro, pnMovItem, lnMovOrden, prsObj!Código
                            Case ObjEntidadesFinancieras
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!Objpadre
                                InsertaMovObjIF pnMovNro, pnMovItem, lnMovOrden, Mid(prsObj!Código, 4, 13), Mid(prsObj!Código, 1, 2), Mid(prsObj!Código, 18, 10)
                            Case Else
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!Código
                        End Select
                    End If
                    prsObj.MoveNext
                Loop
            End If
            prsOtro.MoveNext
        Loop
    End If
End If

End Sub


Public Function ObtieneMovNroRef(ByVal pnMovNro As Long) As Long
Dim coConex As COMConecta.DCOMConecta
Dim sSql As String

    Set coConex = New COMConecta.DCOMConecta
    Dim lrst As ADODB.Recordset
    
    coConex.AbreConexion
    sSql = "Select nMovNro from MovRef where nMovNroRef =" & pnMovNro
    Set lrst = coConex.CargaRecordSet(sSql)
    coConex.CierraConexion
    
    If Not lrst.EOF Then
        ObtieneMovNroRef = lrst("nMovNro")
    End If
        
End Function

Public Function GrabarDatosCComodin(ByVal psCtaCod As String, _
                                    ByVal pMatCal As Variant)
                                    
On Error GoTo ErrorGrabarDatosCComodin

    Call ActualizaFechaCalendarioMatriz(psCtaCod, pMatCal)
    Call dUpdateColocacCred(psCtaCod, , , , , , , , , , , , , , , , , , False, 2)

    Exit Function
    
ErrorGrabarDatosCComodin:
    Err.Raise Err.Number, "Grabar Datos Cuota Comodin", Err.Description
End Function

Public Function GrabarGastosExonerados(ByVal psCtaCod As String, _
                                    ByVal pMatGastos As Variant)
                                    
Dim i As Integer

On Error GoTo ErrorGrabarGastosExonerados

    dBeginTrans
    Call EliminaGastosExonerados(psCtaCod)
    For i = 0 To UBound(pMatGastos) - 1
        Call dInsertColocCredGastosExon(psCtaCod, CLng(pMatGastos(i)))
    Next i
    dCommitTrans

    Exit Function
    
ErrorGrabarGastosExonerados:
    Err.Raise Err.Number, "Grabar Datos Gastos Exonerados", Err.Description
End Function

Public Function ActualizaCodigoModular(ByVal pMatDatos As Variant)
                                    
Dim i As Integer
Dim sSql As String
On Error GoTo ErrorActualizaCodigoModular

    dBeginTrans
    For i = 0 To UBound(pMatDatos) - 1
        Call dUpdateColocacConvenio(pMatDatos(i, 0), , False, pMatDatos(i, 1))
        sSql = "UPDATE ColocacConvenio Set cCargo = '" & pMatDatos(i, 2) & "' Where cCtaCod = '" & pMatDatos(i, 0) & "'"
        coConex.Ejecutar sSql
        sSql = "UPDATE ColocacConvenio Set cCARBEN = '" & pMatDatos(i, 3) & "' Where cCtaCod = '" & pMatDatos(i, 0) & "'"
        coConex.Ejecutar sSql
        sSql = "UPDATE ColocacConvenio Set cT_Plani = '" & pMatDatos(i, 4) & "' Where cCtaCod = '" & pMatDatos(i, 0) & "'"
        coConex.Ejecutar sSql
    Next i
    dCommitTrans

    Exit Function
    
ErrorActualizaCodigoModular:
    Err.Raise Err.Number, "Actualizar Codigo Modular", Err.Description
End Function

Public Function ValorizarCheques(ByVal pMatCheques As Variant, _
                                ByVal pnEstChequeNew As ChequeEstado)
                                    
Dim i As Integer

On Error GoTo ErrorValorizarCheques

    dBeginTrans
    For i = 0 To UBound(pMatCheques) - 1
        Call dUpdateDocRec(pMatCheques(i, 0), pMatCheques(i, 1), False, pnEstChequeNew)
        Call dUpdateDocRecEst(pMatCheques(i, 0), pMatCheques(i, 1), False, pnEstChequeNew)
    Next i
    dCommitTrans

    Exit Function
    
ErrorValorizarCheques:
    Err.Raise Err.Number, "Valorizar Cheques", Err.Description
End Function

Public Function RegistrarDacionEnPago(ByVal pnValorTotal As Double, _
                                    ByVal pnAccion As Integer, _
                                    ByVal pdFecSis As Date, _
                                    ByVal psCtaCod As String, _
                                    ByVal pMatDatos As Variant, _
                                    ByVal pnNroGarantRec As Long, _
                                    ByVal psCodAge As String, _
                                    ByVal psCodUser As String)

Dim nCodNew As Long
Dim i As Integer
Dim oFun As COMNContabilidad.NCOMContFunciones
Dim sMovNro As String
Dim nMovNro As Long

On Error GoTo ErrorRegistrarDacionEnPago
    
    'Si es Nuevo
    If pnAccion = 1 Then
        dBeginTrans
        nCodNew = dNuevoCodigoColocGarantRec
        Call dInsertColocGarantRec(nCodNew, gColocGarantRecDacion, psCtaCod, pdFecSis, pnValorTotal, gColocGarantRecEstadoRegistrado, False)
        'Ingresa el Detalle
        For i = 0 To UBound(pMatDatos) - 1
            Call dInsertColocGarantRecDet(nCodNew, CInt(pMatDatos(i, 1)), CDbl(pMatDatos(i, 2)), CDbl(pMatDatos(i, 3)), pMatDatos(i, 0), False)
        Next i
        Set oFun = New COMNContabilidad.NCOMContFunciones
        sMovNro = oFun.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
        Set oFun = Nothing
        Call dInsertMov(sMovNro, gCredRegisDacion, "Registro de dacion de Pago", gMovEstContabMovContable, gMovFlagVigente, False)
        nMovNro = dGetnMovNro(sMovNro)
        Call dInsertMovCol(nMovNro, gCredRegisDacion, psCtaCod, nCodNew, pnValorTotal, 0, "", 0, 0, 0, False)
        dCommitTrans
    Else
        'Si Es Modificacion
        dBeginTrans
        Call dUpdateColocGarantRec(pnNroGarantRec, psCtaCod, pnValorTotal, False)
        Call dDeleteColocGarantRecDet(pnNroGarantRec, False)
        For i = 0 To UBound(pMatDatos) - 1
            Call dInsertColocGarantRecDet(pnNroGarantRec, CInt(pMatDatos(i, 1)), CDbl(pMatDatos(i, 2)), CDbl(pMatDatos(i, 3)), pMatDatos(i, 0), False)
        Next i
        dCommitTrans
    End If
    
    Exit Function
ErrorRegistrarDacionEnPago:
    Err.Raise Err.Number, "Registrar Dacion en Pago", Err.Description
End Function

'QUITADO RECO
'Public Function InsercionGastosxCuotaLote(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
'                                    ByVal pnColocCalendApl As ColocCalendApl, ByVal pMatCuotas As Variant, ByVal pnColocConceptoCod As Long, _
'                                    ByVal pnMonto As Currency) As String
'Dim i As Integer
'Dim oGasto As COMDCredito.DCOMGasto
'
'On Error GoTo ErrorInsercionGastosxCuotaLote
'Set oGasto = New COMDCredito.DCOMGasto
'
'For i = 0 To UBound(pMatCuotas) - 1
'    If oGasto.ExisteGasto(psCtaCod, pnNroCalen, pnColocCalendApl, pMatCuotas(i), pnColocConceptoCod) Then
'        InsercionGastosxCuotaLote = "Gasto ya Existe para la Cuota " & pMatCuotas(i)
'        Exit Function
'    End If
'    Call dInsertColocCalendDet(psCtaCod, pnNroCalen, pnColocCalendApl, pMatCuotas(i), pnColocConceptoCod, pnMonto, 0#, "", False)
'Next
'Set oGasto = Nothing
'
''No hay ningun Error
'InsercionGastosxCuotaLote = ""
'Exit Function
'ErrorInsercionGastosxCuotaLote:
'    InsercionGastosxCuotaLote = "Error en la Asignacion de los Gastos en Lote"
'    'Err.Raise Err.Number, "Error en la Asignacion de los Gastos en Lote", Err.Description
'End Function
Public Function InsercionGastosxCuotaLote(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
                                    ByVal pnColocCalendApl As ColocCalendApl, ByVal pMatCuotas As Variant, ByVal pnColocConceptoCod As Long, _
                                    ByVal pnMonto As Currency, Optional pbMantPolizaCI As Boolean = False) As String
Dim i As Integer
Dim oGasto As COMDCredito.DCOMGasto

On Error GoTo ErrorInsercionGastosxCuotaLote
Set oGasto = New COMDCredito.DCOMGasto

For i = 0 To UBound(pMatCuotas) - 1
    If oGasto.ExisteGasto(psCtaCod, pnNroCalen, pnColocCalendApl, pMatCuotas(i), pnColocConceptoCod) Then
        '*********RECO 20131122 ERS055************************************
        If pbMantPolizaCI Then
            Call dUpdateColocCalendDet(psCtaCod, pnNroCalen, pnColocCalendApl, pMatCuotas(i), pnColocConceptoCod, pnMonto, 0#, "", False)
        Else
            InsercionGastosxCuotaLote = "Gasto ya Existe para la Cuota " & pMatCuotas(i)
            Exit Function
        End If
        '***************END RECO******************************************
        'InsercionGastosxCuotaLote = "Gasto ya Existe para la Cuota " & pMatCuotas(i)
        'Exit Function
    Else
            Call dInsertColocCalendDet(psCtaCod, pnNroCalen, pnColocCalendApl, pMatCuotas(i), pnColocConceptoCod, pnMonto, 0#, "", False)
    End If
Next
Set oGasto = Nothing

'No hay ningun Error
InsercionGastosxCuotaLote = ""
Exit Function
ErrorInsercionGastosxCuotaLote:
    InsercionGastosxCuotaLote = "Error en la Asignacion de los Gastos en Lote"
    'Err.Raise Err.Number, "Error en la Asignacion de los Gastos en Lote", Err.Description
End Function



'**DAOR 20070131
'**Creditos que serán desembolsados en el banco de la nación
Public Sub dInsertColocacConvBcoNac(ByVal psCtaCod As String, ByVal psClaveCli As String, ByVal psCodAgeBcoNac As String, pnEstado As Integer)
Dim lsSQL As String
On Error GoTo ErrordInsertColocacConvBcoNac
    
    lsSQL = "INSERT INTO ColocConvBcoNac(cCtaCod,cClave,cCodAgeBcoNac,nEstado)  "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "','" & psClaveCli & "','" & psCodAgeBcoNac & "'," & pnEstado & ")"
    
    coConex.Ejecutar lsSQL

    Exit Sub
ErrordInsertColocacConvBcoNac:
    Err.Raise Err.Number, "Error en registrar datos para el desembolso en el banco de la nación (dInsertColocacConvBcoNac)", Err.Description
End Sub

'**DAOR 20070131
'**Creditos que serán desembolsados en el banco de la nación
Public Sub dDeleteColocConvBcoNac(ByVal psCtaCod As String)
Dim lsSQL As String
On Error GoTo ErrordDeleteColocConvBcoNac
    
    lsSQL = "DELETE ColocConvBcoNac WHERE cCtaCod = '" & psCtaCod & "' "
    coConex.Ejecutar lsSQL

    Exit Sub
ErrordDeleteColocConvBcoNac:
    Err.Raise Err.Number, "Error en Proceso (dDeleteColocConvBcoNac)", Err.Description
End Sub

'AVMM -- GESTION DE COBRANZAS -- 22-03-2007
'Buscar Gestiones de Cobranza para un Credito
Public Function ObtenerGestionCobranza(ByVal psCtaCod As String) As Recordset


    Dim lsSQL As String
    Dim rs As New ADODB.Recordset

    lsSQL = " SELECT A.cCtaCod, A.cMovNro, A.cComenta, A.dFechaAviso,A.dFechaVencimiento  " _
          & " FROM ColocacCredsGestionCobranza A " _
          & " WHERE A.cCtaCod ='" & psCtaCod & "'"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open lsSQL, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set ObtenerGestionCobranza = rs
    Set rs = Nothing
End Function


'Insert Gestión de Cobranza para un Credito
Public Sub dInsertGestionCobranza(ByVal psCtaCod As String, ByVal psMovNro As String, ByVal psComenta As String, _
                                          Optional ByVal pdFechaA As Date, Optional ByVal pdFechaV As Date)

Dim lsSQL As String

lsSQL = "INSERT ColocacCredsGestionCobranza (cCtaCod, cMovNro, cComenta,dFechaAviso,dFechaVencimiento) " _
    & " VALUES ('" & psCtaCod & "','" & psMovNro & "','" & psComenta & "','" & Format(pdFechaA, "yyyymmdd") & "','" & Format(pdFechaV, "yyyymmdd") & "') "
    
coConex.Ejecutar lsSQL
  
End Sub

'**DAOR 20070413, Funcion que recupera las cuentas de plazos fijos
'**asociados como garantías del crédito
Public Function RecuperaPlazosFijosAsocComoGarantia(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
On Error GoTo ErrorRecuperaPlazosFijosAsocComoGarantia
    '**Modificado por DAOR 20080108, Sólo considerar cuentas activas
    'sSql = "Select  P.cCtaCod,P.nSaldo,C.nSaldoDisp " & _
    '        "From ColocGarantia CG inner join Garantias G on CG.cNumGarant=G.cNumGarant " & _
    '        "inner join Producto P on rtrim(G.cNroDoc)=P.cCtaCod " & _
    '        "inner join Captaciones C on P.cCtaCod=C.cCtaCod " & _
    '        "where G.nTpoGarantia=6 and G.cTpoDoc=17 and CG.cCtaCod='" & psCtaCod & "'"

    sSql = " exec stp_sel_PlazosFijosAsocComoGarantia '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPlazosFijosAsocComoGarantia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaPlazosFijosAsocComoGarantia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function


'**DAOR 20070413, Funcion copiado de las clase DCOMCaptaGenerales
Public Function NuevoBloqueoRetiro(ByVal sCuenta As String, ByVal nMotivo As CaptacMotBloqueoRet, _
        cComentario As String, ByVal cMovnro As String)
Dim sSql As String
sSql = "INSERT ProductoBloqueos (cCtaCod,nBlqTpo,nBlqMotivo,cComentario,cMovNro,cMovNroDbl) " _
    & "VALUES ('" & sCuenta & "'," & gCapTpoBlqRetiro & "," & nMotivo & ",'" & cComentario & "','" & cMovnro & "',NULL)"
coConex.ConexionActiva.Execute sSql
End Function

'madm 20091216********************************
Public Sub dInsertCalendarioReprogradoLote(ByVal nNroCalen As Integer, ByVal psCtaCod As String, Optional pbEjecBatch As Boolean = False)

    Dim lsSQL As String
    
    lsSQL = "Insert ColocCalendario " _
        & "Select cCtaCod, " & nNroCalen & ",nColocCalendApl,nCuota,dVenc,dPago,nColocCalendEstado,cDescripcion,cColocCalenFlag,nCalendProc,cColocMiVivEval " _
        & "from ColocCalendario Where nColocCalendApl=0 and cCtaCod = '" & psCtaCod & "' And nNroCalen = " & nNroCalen - 1
    
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
Public Sub dInsertCalendarioDetReprogradoLote(ByVal nNroCalen As Integer, ByVal psCtaCod As String, Optional pbEjecBatch As Boolean = False)

    Dim lsSQL As String
    
   lsSQL = "Insert ColocCalenddet " _
        & "Select cCtaCod, " & nNroCalen & ", nColocCalendApl, nCuota, nPrdConceptoCod, nMonto, nMontoPagado, cFlag " _
        & "from ColocCalenddet Where nColocCalendApl=0 and cCtaCod = '" & psCtaCod & "' And nNroCalen = " & nNroCalen - 1
          
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
'*********************************************

'MAVM 20100823
Public Sub dInsertCredMantPrepago(ByVal psCtaCod As String, Optional ByVal psMovNro As String, Optional ByVal pnCalendDinamico As Integer, Optional ByVal pnMantNroCuota As Integer, Optional ByVal pnOpcion As Integer, Optional ByVal pnCuotasPend As Integer)
Dim lsSQL As String
    
    lsSQL = "UPDATE CredMantPrepago Set nEstado = 0 Where cCtaCod = '" & psCtaCod & "' and Substring(cMovNro, 1, 8) = '" & Mid(psMovNro, 1, 8) & "'"
    coConex.Ejecutar lsSQL
    
    If pnOpcion = 0 Then
        lsSQL = "INSERT CredMantPrepago (cCtaCod, cMovNro, nCalendDinamico, nMantNroCuota, nCuotasPend, nEstado) "
        lsSQL = lsSQL & " Values('" & psCtaCod & "','" & psMovNro & "', '" & pnCalendDinamico & "', '" & pnMantNroCuota & "', '" & pnCuotasPend & "', 1)"
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dUpdateColocacEstadoCD(ByVal psCtaCod As String, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    On Error GoTo ErrordUpdateColocacEstadoCD
    lsSQL = " UPDATE ColocacEstado SET "
    lsSQL = lsSQL & " nPlazo = 0"
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "' AND nPrdEstado = " & "2002"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub
ErrordUpdateColocacEstadoCD:
    Err.Raise Err.Number, "Error En Proceso dUpdateColocacEstadoCD", Err.Description
End Sub
'***

'MAVM 20101218
Public Sub dInsertColocCFHistorial(ByVal psCtaCod As String, ByVal pdFecVencAnt As Date, ByVal pnPeriodo As Integer, _
                                   ByVal pdFecVenc As Date, ByVal pnMonto As Currency, _
                                   Optional ByVal pbEjecBatch As Boolean = False)

Dim sSql As String
        sSql = "INSERT INTO ColocCFHistorial(cCtaCod, dVencAnt, nPeriodo, dVenc, nMonto) "
        sSql = sSql & " VALUES('" & psCtaCod & "','" & Format(CDate(Format(pdFecVencAnt, "dd/mm/yyyy")), "mm/dd/yyyy") & "', " & pnPeriodo & ",'" & Format(CDate(Format(pdFecVenc, "dd/mm/yyyy")), "mm/dd/yyyy") & "', " & pnMonto & ")"
        If pbEjecBatch Then
            coConex.AdicionaCmdBatch sSql
        Else
            coConex.Ejecutar sSql
        End If
End Sub
'ALPA 20111022
Public Sub dAgregaMovDoc(ByVal nMovNro As Long, ByVal nTipoDoc As TpoDoc, _
        ByVal sDocumento As String, ByVal dFecDocumento As Date, Optional ByVal pbEjecBatch As Boolean = False)
Dim sFecha As String
Dim sSql As String
sFecha = Format$(dFecDocumento, "mm/dd/yyyy")
sSql = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
    & "VALUES (" & nMovNro & "," & nTipoDoc & ",'" & sDocumento & "','" & sFecha & "')"
If pbEjecBatch Then
    coConex.AdicionaCmdBatch sSql
Else
    coConex.Ejecutar sSql
End If
End Sub
Public Sub dAgregaMovOpeVarias(ByVal pnMovNro As Long, ByVal pnMovImporte As Double, _
        ByVal psNroDoc As String, ByVal psReferencia As String, ByVal pnMoneda As Moneda, _
        Optional nOperacion As CaptacOperacion, Optional ByVal pbEjecBatch As Boolean = False)
    Dim SQL As String
    
    SQL = " Insert MovOpeVarias (nMovNro,nMovImporte,cNroDoc,cReferencia,nMoneda,cOpeCod)" _
        & " Values(" & pnMovNro & "," & pnMovImporte & ",'" & psNroDoc & "','" & psReferencia & "'," & pnMoneda & ",'" & Trim(nOperacion) & "')"
    
If pbEjecBatch Then
    coConex.AdicionaCmdBatch SQL
Else
    coConex.Ejecutar SQL
End If
End Sub

'MADM 20111116
'Inserta Mov
Public Sub dInsertMovModifica(ByVal psMovNro As String, ByVal psMovNroAnt As String, Optional ByVal pbEjecBatch As Boolean = False, Optional psMovNroNew As String = "")
Dim lsSQL As String
On Error GoTo dInsertMovModificaErr
      
    lsSQL = "INSERT MovModifica (cMovNro, cMovNroAnt,cMovNroNew)" _
          & "VALUES ('" & psMovNro & "','" & psMovNroAnt & "','" & psMovNroNew & "')"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

dInsertMovModificaErr:
     Err.Raise Err.Number, "Error En Proceso dUpdateColocacEstadoCD", Err.Description
End Sub

'madm 20111119
Public Sub Actualizar_EstxMov_MovCorresponsalia(ByVal nMovNro As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorActualizar_EstxMov_MovCorresponsalia
 
   sSql = "Exec stp_upd_LogCorresponsaliaBcoNacDet_MovEst '" & nMovNro & "'"
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
ErrorActualizar_EstxMov_MovCorresponsalia:
    Err.Raise Err.Number, "Actualizar Datos Detalle", Err.Description
End Sub
Public Sub Actualizar_EstxMov_MovConvenio(ByVal nMovNro As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorActualizar_EstxMov_MovConvenio
 
       sSql = "Exec stp_upd_LogConvenioBcoNacDet_MovEst '" & nMovNro & "'"
       Set oConecta = New COMConecta.DCOMConecta
       oConecta.AbreConexion
       oConecta.ConexionActiva.Execute sSql
       oConecta.CierraConexion
       Set oConecta = Nothing
    Exit Sub
ErrorActualizar_EstxMov_MovConvenio:
    Err.Raise Err.Number, "Actualizar Datos Detalle", Err.Description
End Sub
'END MADM
'ALPA 20111212******************************************************
Public Function dGetnMovNroLeasing(ByVal psMovNro As String, Optional ByVal pbEjecBatch As Boolean = False) As Long


Dim lsSQL As String
Dim lrs As ADODB.Recordset
Set lrs = New ADODB.Recordset

lsSQL = "Select nMovNro From Mov where cMovNro ='" & psMovNro & "'"
Set lrs = coConex.CargaRecordSet(lsSQL, adLockReadOnly)
'Set lrs.ActiveConnection = coConex.ConexionActiva
Set lrs = coConex.Ejecutar(lsSQL)
If Not lrs.EOF And Not lrs.BOF Then
    dGetnMovNroLeasing = lrs!nMovNro
End If
lrs.Close
Set lrs = Nothing
End Function
'ALPA 20111101
Public Sub dActualizaPagoSaf(Optional ByVal pbEjecBatch As Boolean = False)
    Dim lsSQL As String

    lsSQL = "exec stp_sel_actualiza_pagoSAF "
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

'MADM 20110105
Public Sub dDeleteProductoPersonaRelac(ByVal psCtaCod As String, ByVal nPrdPersRelac As Integer, Optional ByVal pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    On Error GoTo ErrordDeleteProductoPersonaRelac
    
    lsSQL = "Exec EliminarProductoPersRelac '" & psCtaCod & "'," & nPrdPersRelac & ""
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
    Exit Sub
ErrordDeleteProductoPersonaRelac:
    Err.Raise Err.Number, "Error En Proceso dDeleteProductoTasaInteres", Err.Description
End Sub
'END MADM

'*********************************************************************
'*** BRGO 20111103 *****************************************
Public Sub AgregaNuevoProdPersCofide(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As Integer, ByVal sCuentaAbono As String, ByVal nMontoAbono As Double)
    Dim sSql As String
    sSql = "INSERT ProductoPersonaInfogas " _
        & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ",'" & sCuentaAbono & "'," & Format(nMontoAbono, "0.00") & ")"
    coConex.ConexionActiva.Execute sSql
End Sub

Public Sub EliminaRelacionesInfogas(ByVal sCuenta As String)
    Dim sSql As String
    sSql = "Delete From ProductoPersonaInfogas Where cCtaCod = '" & sCuenta & "'"
    coConex.ConexionActiva.Execute sSql
End Sub
'*** END BRGO ****************************************************

'*** BRGO 20111227 *****************************************
Public Sub AgregaNuevoPersCofide(ByVal sPersona As String, ByVal nTipoPers As Integer, ByVal nCodCofide As Integer)
    Dim sSql As String
    sSql = "Exec stp_ins_PersonaCofide '" & sPersona & "', " & nTipoPers & "," & nCodCofide
    coConex.ConexionActiva.Execute sSql
End Sub
Public Sub ModificaPersCofide(ByVal sPersona As String, ByVal nTipoPers As Integer, ByVal nCodCofide As Integer)
    Dim sSql As String
    sSql = "Exec stp_upd_PersonaCofide '" & sPersona & "', " & nTipoPers & "," & nCodCofide
    coConex.ConexionActiva.Execute sSql
End Sub
Public Function ValidaExistePersCofide(ByVal sPersona As String, ByVal nTipoPers As Integer) As Boolean
    Dim lrs As ADODB.Recordset
    Dim sSql As String
    Dim bExiste As Boolean
    Set lrs = New ADODB.Recordset
    
    sSql = "Exec stp_sel_ValidaExistePersonaCofide '" & sPersona & "', " & nTipoPers
    Set lrs = coConex.CargaRecordSet(sSql, adLockReadOnly)
    bExiste = False
    If Not lrs.EOF And Not lrs.BOF Then
        bExiste = True
    End If
    ValidaExistePersCofide = bExiste
End Function
Public Function ListaPersonasCofide() As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim sSql As String
  
    sSql = "Exec stp_sel_ListaPersonasCofide"
    Set lrs = coConex.CargaRecordSet(sSql, adLockReadOnly)
    Set ListaPersonasCofide = lrs
End Function
Public Function ListaPersonasCofideTipo(ByVal nTipoPers As Integer) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim sSql As String
  
    sSql = "Exec stp_sel_ListaPersonasCofideTipo " & nTipoPers
    Set lrs = coConex.CargaRecordSet(sSql, adLockReadOnly)
    Set ListaPersonasCofideTipo = lrs
End Function

Public Sub AgregaNuevoVehiculoInfoGas(ByVal sCuenta As String, ByVal sCodConcesionario As String, ByVal sCodTaller As String, _
    ByVal nTipoVehiculo As Integer, ByVal sPlaca As String, ByVal sPlacaNueva As String, ByVal nMontoAprobado As Double, _
    ByVal nCuotaInicial As Double, ByVal nRecaudo As Double, ByVal nTipoCredito As Integer, ByVal dFecEmision As Date, _
    ByVal sUltActualizacion As String, ByVal sUser As String)
    Dim sSql As String
    sSql = "Exec stp_ins_VehiculoInfoGas '" & sCuenta & "','" & sCodConcesionario & "','" & sCodTaller & "'," & nTipoVehiculo & ",'" & sPlaca & "','" & sPlacaNueva & "'," & Format(nMontoAprobado, "0.00") & "," & Format(nCuotaInicial, "0.00") & "," & Format(nRecaudo, "0.00") & "," & nTipoCredito & ",'" & Format(dFecEmision, "MM/dd/yyyy") & "','" & sUltActualizacion & "','" & sUser & "'"
    coConex.ConexionActiva.Execute sSql
End Sub
Public Sub ModificaVehiculoInfoGas(ByVal sCuenta As String, ByVal sCodConcesionario As String, ByVal sCodTaller As String, _
    ByVal nTipoVehiculo As Integer, ByVal sPlaca As String, ByVal sPlacaNueva As String, ByVal nMontoAprobado As Double, _
    ByVal nCuotaInicial As Double, ByVal nRecaudo As Double, ByVal sUltActualizacion As String)
    Dim sSql As String
    sSql = "Exec stp_upd_VehiculoInfoGas '" & sCuenta & "','" & sCodConcesionario & "','" & sCodTaller & "'," & nTipoVehiculo & ",'" & sPlaca & "','" & sPlacaNueva & "'," & Format(nMontoAprobado, "0.00") & "," & Format(nCuotaInicial, "0.00") & "," & Format(nRecaudo, "0.00") & ",'" & sUltActualizacion & "'"
    coConex.ConexionActiva.Execute sSql
End Sub
Public Function ObtieneDatosVehiculoInfoGas(ByVal sCuenta As String) As ADODB.Recordset
    Dim lrs As ADODB.Recordset
    Dim sSql As String
  
    sSql = "Exec stp_sel_ObtieneDatosVehiculoInfoGas '" & sCuenta & "'"
    Set lrs = coConex.CargaRecordSet(sSql, adLockReadOnly)
    Set ObtieneDatosVehiculoInfoGas = lrs
End Function
'*** END BRGO ****************************************************

'** Juez 20120323 ************************************
Public Sub dInsertaPersonaRealizaPagoCuotas(ByVal nNroMov As Long, ByVal sPersCod As String, ByVal nPersDNI As String, ByVal sPerNombre As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_CredPagoCuotasPersRealiza " & nNroMov & ",'" & sPersCod & "','" & nPersDNI & "','" & sPerNombre & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub
'** End Juez *****************************************
'***Agregado por ELRO el 20120329, según RFC023-2012
Public Sub registrarColocacionPorAfectacion(ByVal pnMovNro As Long, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    lsSQL = "exec stp_ins_RegistrarAfectacion " & pnMovNro
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    lsSQL = ""
End Sub
Public Sub actualizarColocacionPorAfectacion(ByVal pnMovNro As Long, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    lsSQL = "exec stp_upd_ActualizarAfectacion " & pnMovNro
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    lsSQL = ""
End Sub
'***Fin Agregado por ELRO***************************
'WIOR 20120328-INICIO
Public Sub dInsertProductoPersona2(ByVal psCuenta As String, ByVal psPersona As String, ByVal pnRelacion As ColocRelacPers, ByVal psConsorcio As String, _
    Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    'Insertar ProductoPersona con el codigo mismo del Titular por ser un consorcio
    lsSQL = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
        & "VALUES ('" & psCuenta & "','" & psPersona & "'," & pnRelacion & ")"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
    'Actualiza la tabla ColocCartaFianza con el nombre del consorcio
    lsSQL = "UPDATE ColocCartaFianza SET cConsorcio= upper('" & psConsorcio & "') WHERE cCtaCod='" & psCuenta & "'"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub
'WIOR - FIN

'WIOR 20120330
Public Sub dDeleteConsorcio(ByVal psCtaCod As String, Optional ByVal pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    On Error GoTo ErrordDeleteConsorcio
 
    'Actualiza la tabla ColocCartaFianza quitando el nombre consorcio
    lsSQL = "UPDATE ColocCartaFianza SET cConsorcio= null WHERE cCtaCod='" & psCtaCod & "'"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
 Exit Sub
ErrordDeleteConsorcio:
    Err.Raise Err.Number, "Error En Proceso dDeleteConsorcio", Err.Description
End Sub
'WIOR FIN
'ALPA 20120421************************************************************
Public Sub dActualizarPagoLeasig(Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String

    lsSQL = "exec stp_upd_actualizarpagos "
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub
'**************************************************************************
'EJVG20120514 *************************************************************
Public Sub dLimpiaRecaudoDetTemp()
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    lsSQL = "Delete from RecaudoEcoTaxiDetTemp"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub
Public Sub dMigraRecaudosDet(ByVal pdFechaCarga As Date)
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    sSql = "Exec stp_ins_MigrarRecaudosDet '" & Format(pdFechaCarga, "yyyymmdd hh:mm:ss") & "'"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
End Sub
Public Sub dRegistrarRecaudoDetTemp(ByVal psPlaca As String, ByVal pdEESSFechaHora As Date, ByVal psEESSNombre As String, _
                        ByVal psEESSNroTicket As String, ByVal pnEESSRecaudoBruto As Currency, _
                        ByVal pnEESSITFEntrada As Currency, ByVal pnEESSITFSalida As Currency, ByVal pnEESSRecaudoReal As Currency, _
                        ByVal pnCOFIDEPorcentajeComision As Currency, ByVal pnIFIRecaudoNeto As Currency, _
                        ByVal psCtaCodCredito As String, ByVal psCtaCodAbono As String, ByVal psPersCod As String, ByVal psIFICod As String, _
                        ByVal psEESSId As String, ByVal bAptoAbono As Boolean)
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    sSql = "Exec stp_ins_RecaudoEcoTaxiDetTemp '" & psPlaca & "','" & Format(pdEESSFechaHora, "yyyymmdd hh:mm:ss") & "','" & psEESSNombre _
        & "','" & psEESSNroTicket & "'," & pnEESSRecaudoBruto _
        & "," & pnEESSITFEntrada & "," & pnEESSITFSalida & "," & pnEESSRecaudoReal _
        & "," & pnCOFIDEPorcentajeComision & "," & pnIFIRecaudoNeto _
        & ",'" & psCtaCodCredito & "','" & psCtaCodAbono & "','" & psPersCod & "','" & psIFICod & "','" & psEESSId & "'," & IIf(bAptoAbono = True, 1, 0)
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.Ejecutar sSql
    oConec.CierraConexion
    Set oConec = Nothing
End Sub
Public Sub dAsignaNroMovDepEcoTaxi(ByVal pnId As Long, ByVal pnMovNro As Long)
    Dim sSql As String
    sSql = "Exec stp_upd_AsignaNroMovDepEcoTaxi " & pnId & "," & pnMovNro
    coConex.Ejecutar sSql
End Sub
Public Function DameRecaudo(ByVal psCtaCod As String, ByVal psEESSId As String, ByVal psEESSNroTicket As String, ByVal pdFechaRecaudo As Date) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    sSql = "Exec stp_sel_ListaRecaudoPaVerificar '" & psCtaCod & "','" & psEESSId & "','" & psEESSNroTicket & "','" & Format(pdFechaRecaudo, "yyyymmdd hh:mm:ss") & "'"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set DameRecaudo = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Sub InsertaMovRedondeoITF(ByVal psMovNro As String, ByVal pnMovItem As Integer, ByVal pnITFCalculado As Currency, ByVal pnITFRedondeado As Currency, Optional ByVal pnMovNro As Long = 0)
Dim SQL As String
Dim lnMovNro As Long
    If pnMovNro > 0 Then
        lnMovNro = pnMovNro
    Else
        lnMovNro = GetnMovNro(psMovNro)
    End If

    SQL = " exec stp_ins_MovRedondeoITF " & lnMovNro & "," & pnMovItem & "," & pnITFCalculado & "," & pnITFRedondeado
    coConex.Ejecutar SQL
    Exit Sub
End Sub
Public Sub ActualizaOperacionITFPagoCuotaEcoTaxi(ByVal pnMovNro As Long, ByVal lsOpeITF As String)
    Dim lsSQL As String
    lsSQL = "Update MovCol Set cOpeCod = '" & lsOpeITF & "' Where nMovNro = " & pnMovNro & " And cOpeCod like '99%'"
    coConex.Ejecutar lsSQL
    lsSQL = "Update MovColDet Set cOpeCod = '" & lsOpeITF & "' Where nMovNro = " & pnMovNro & " And cOpeCod like '99%'"
    coConex.Ejecutar lsSQL
End Sub
Public Function DameOpeCodPagoCuotaEcoTaxi(ByVal pnMovNro As Long) As String
    Dim lsSQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    lsSQL = "Select cOpeCod from MovCol where nMovNro = " & pnMovNro & " And cOpeCod like '100[234567]02'"
    Set rs = coConex.CargaRecordSet(lsSQL)
    DameOpeCodPagoCuotaEcoTaxi = rs!cOpeCod
    Set rs = Nothing
End Function
'END EJVG *****************************************************************
'*** BRGO 20120106 *********************************************************
Public Sub NuevaAprobacionCreditoINFOGAS(ByVal psCuenta As String, ByVal psPlaca As String, _
                        ByVal psPersId As String, ByVal pdFechaGeneracion As Date, _
                        ByVal psIFICod As String, ByVal psTallerId As String)
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    sSql = "Exec stp_ins_AprobacionINFOGAS '" & psCuenta & "','" & psPlaca & "','" & psPersId & "','" & psIFICod & "','" & psTallerId & "','" & Format(pdFechaGeneracion, "MM/dd/yyyy hh:mm:ss") & "'"
    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing

End Sub
Public Sub ActivaCreditoINFOGAS(ByVal psCuenta As String, ByVal psPlaca As String, ByVal pdFecha As Date)
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta

    Set oCon = New COMConecta.DCOMConecta
    sSql = "Exec stp_upd_ActivaCreditoINFOGAS '" & psCuenta & "','" & psPlaca & "','" & Format(pdFecha, "MM/dd/yyyy hh:mm:ss") & "'"

    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing

End Sub
Public Function ObtieneEstadoVehiculoInfoGas(ByVal psCuenta As String) As Boolean

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim oRs As ADODB.Recordset

    sSql = "Exec stp_sel_ObtieneEstadoVehiculoINFOGAS '" & psCuenta & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set oRs = oCon.CargaRecordSet(sSql)
    If oRs!cActivo = 1 Then
        ObtieneEstadoVehiculoInfoGas = True
    Else
        ObtieneEstadoVehiculoInfoGas = False
    End If
    oCon.CierraConexion
    Set oRs = Nothing
    Set oCon = Nothing
End Function
'*** END BRGO **************************************************************

'** Juez 20120514 *********************************************************
Public Function obtieneListaCredPagoCuotasAgeParam() As ADODB.Recordset
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "Select A.cAgeCod,A.cAgeDescripcion,isnull(CP.nMontoMin,0)nMontoMin,isnull(CP.nMontoMax,0)nMontoMax"
    lsSQL = lsSQL & " From Agencias A"
    lsSQL = lsSQL & " left join CredPagoCuotasAgeParam CP on CP.cAgeCod = A.cAgeCod AND CP.bActivo = 1"
    lsSQL = lsSQL & " where A.nEstado = 1"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set obtieneListaCredPagoCuotasAgeParam = lrDatos
    Set lrDatos = Nothing
End Function
Public Function obtieneCredPagoCuotasAgeParam(ByVal sCodAge As String) As ADODB.Recordset
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset

    lsSQL = "Select isnull(nMontoMin,0)nMontoMin,isnull(nMontoMax,0)nMontoMax"
    lsSQL = lsSQL & " From CredPagoCuotasAgeParam"
    lsSQL = lsSQL & " where bActivo =1 and cAgeCod ='" & sCodAge & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set obtieneCredPagoCuotasAgeParam = lrDatos
    Set lrDatos = Nothing
End Function
Public Sub insertarCredPagoCuotasAgeParam(ByVal sAgeCod As String, ByVal nMontoMin As Double, ByVal nMontoMax As Double, ByVal sUltimaActualizacion As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset

    lsSQL = "insert into CredPagoCuotasAgeParam (cAgeCod,nMontoMin,nMontoMax,cUltimaActualizacion,bActivo)"
    lsSQL = lsSQL & " Values ('" & sAgeCod & "'," & nMontoMin & "," & nMontoMax & ",'" & sUltimaActualizacion & "',1)"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub
Public Sub actualizaCredPagoCuotasAgeParam()
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "update CredPagoCuotasAgeParam set bActivo=0 where bActivo=1"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub
'** End Juez **************************************************************
'**Juez 20120601 *********************************************************
Public Function obtieneNroCuotasPorVencUltimaCuotaPagada(ByVal psCtaCod As String, ByVal pdFecSis As String) As ADODB.Recordset
Dim lsSQL As String
Dim RFecha As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta

    lsSQL = "exec stp_sel_obtieneNroCuotasPorVencUltimaCuotaPagada '" & psCtaCod & "', '" & pdFecSis & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion

    Set obtieneNroCuotasPorVencUltimaCuotaPagada = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
End Function
'** End Juez *************************************************************
'WIOR 20120517 ************************************************************
Public Sub RegistrarColocacionMicroseguro(ByVal pcCtaCod As String, ByVal pnTipo As Integer, ByVal pdFecha As Date, ByVal pnEstado As Integer, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    lsSQL = "exec stp_ins_ColocacMicroseguro '" & pcCtaCod & "'," & pnTipo & ",'" & Format(pdFecha, "yyyymmdd") & "'," & pnEstado
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    lsSQL = ""
End Sub
Public Sub RegistrarColocacionMultiriesgo(ByVal pcCtaCod As String, ByVal pdFecha As Date, ByVal pnEstado As Integer, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    lsSQL = "exec stp_ins_ColocacMultiriesgo '" & pcCtaCod & "','" & Format(pdFecha, "yyyymmdd") & "'," & pnEstado
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    lsSQL = ""
End Sub
Public Sub QuitarColocacionMicroseguro(ByVal pcCtaCod As String, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    lsSQL = "delete from ColocacMicroseguro where cCtaCod='" & pcCtaCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    lsSQL = ""
End Sub
Public Sub QuitarColocacionMultiriesgo(ByVal pcCtaCod As String, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    lsSQL = "delete from ColocacMultiriesgo where cCtaCod='" & pcCtaCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    lsSQL = ""
End Sub

Public Sub ActualizarValorMultiriesgo(ByVal pcCtaCod As String, ByVal pnValorTotal As Double, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    lsSQL = "update ColocacMultiriesgo set nValorTotal=" & pnValorTotal & " where cCtaCod='" & pcCtaCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    lsSQL = ""
End Sub
Public Sub ActualizarNumCertificado(ByVal pcCtaCod As String, ByVal pcNumCertf As String, ByVal pcTabla As String, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    lsSQL = "update Colocac" & pcTabla & " set cNumCert='" & pcNumCertf & "' where cCtaCod='" & pcCtaCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    lsSQL = ""
End Sub
Public Sub RegistrarBeneficiariosMicroseguro(ByVal pcCtaCod As String, ByVal psPersCod As String, ByVal pdFecha As Date, ByVal pnEstado As Integer, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    
    lsSQL = " Insert into ColocacMicroseguroBeneficiarios(cCtaCod,cPersCod,dEstado,nEstado)"
    lsSQL = lsSQL & " values ('" & pcCtaCod & "','" & psPersCod & "','" & Format(pdFecha, "yyyymmdd") & "'," & pnEstado & ")"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    lsSQL = ""
End Sub
Public Sub RegistrarMueblesMultiriesgo(ByVal pcCtaCod As String, ByVal psDescripcion As String, ByVal pnAno As Integer, ByVal pnValor As Double, ByVal pdFecha As Date, ByVal pnEstado As Integer, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    
    lsSQL = " Insert into ColocacMultiriesgoMuebles(cCtaCod,cDescripcion,nAno,nValor,dEstado,nEstado)"
    lsSQL = lsSQL & " VALUES ('" & pcCtaCod & "','" & psDescripcion & "'," & pnAno & "," & pnValor & ",'" & Format(pdFecha, "yyyymmdd") & "'," & pnEstado & ")"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    lsSQL = ""
End Sub

Public Function ObtenerBeneficiariosMicroseguro(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "select cCtaCod,cPersCod,dEstado,nEstado from  ColocacMicroseguroBeneficiarios WHERE cCtaCod='" & psCtaCod & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set ObtenerBeneficiariosMicroseguro = rs
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ObtenerMueblesMultiriesgo(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "select cCtaCod,cDescripcion,nAno,nValor,dEstado,nEstado  from  ColocacMultiriesgoMuebles WHERE cCtaCod='" & psCtaCod & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set ObtenerMueblesMultiriesgo = rs
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Sub QuitarBeneficiariosMicroseguro(ByVal pcCtaCod As String, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    lsSQL = "delete from ColocacMicroseguroBeneficiarios where cCtaCod=  '" & pcCtaCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    lsSQL = ""
End Sub
Public Sub QuitarMueblesMultiriesgo(ByVal pcCtaCod As String, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    lsSQL = "delete from ColocacMultiriesgoMuebles where cCtaCod= '" & pcCtaCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    lsSQL = ""
End Sub
'WIOR FIN *****************************************************************
'EJVG20120622 *****************
Public Sub dRegistrarPenalidad(ByVal nMovNro As Long, ByVal nMotivoPenalidad As Double, ByVal cCtaCodCreTit As String, ByVal cCtaCodAhoTit As String, ByVal cPersCodTit As String, ByVal cCtaCodAhoConcesionario As String, ByVal cPersCodConcesionario As String, ByVal nDiasRetraso As Integer, ByVal nMontoxDiaRetraso As Double, ByVal cComentario As String, ByVal nMontoPenalidad As Double, ByVal nMontoFavorCliente As Double, ByVal nMontoFavorInstitucion As Double)
    Dim psSql As String
    psSql = "Exec stp_ins_PenalidadEcoTaxi " & nMovNro & "," & nMotivoPenalidad & ",'" & cCtaCodCreTit & "','" & cCtaCodAhoTit & "','" & cPersCodTit & "','" & cCtaCodAhoConcesionario & "','" & cPersCodConcesionario & "'," & nDiasRetraso & "," & nMontoxDiaRetraso & ",'" & cComentario & "'," & nMontoPenalidad & "," & nMontoFavorCliente & "," & nMontoFavorInstitucion
    coConex.Ejecutar psSql
End Sub
Public Function ListaPenalidadEcoTaxiAExtornarxUsuario(ByVal psUsuario As String, ByVal pdFecha As Date) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "Exec stp_sel_ListaExtornoPenalidadEcoTaxixUsuario '" & psUsuario & "','" & Format(pdFecha, "yyyymmdd") & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set ListaPenalidadEcoTaxiAExtornarxUsuario = rs
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function ListaPenalidadEcoTaxiAExtornarxCtaCod(ByVal psCtaCod As String, ByVal pdFecha As Date) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    Set rs = New Recordset
    On Error GoTo ErrorMsg
    
    If oConec.AbreConexion() Then
        sSql = "Exec stp_sel_ListaExtornoPenalidadEcoTaxixCtaCod '" & psCtaCod & "','" & Format(pdFecha, "yyyymmdd") & "'"
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
    End If
    
    Set ListaPenalidadEcoTaxiAExtornarxCtaCod = rs
    Exit Function
ErrorMsg:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'EJVG END *********************
'**Juez 20120803 **************
Public Sub InsertaDatosImprVoucher(ByVal nMovNro As Long, ByVal psGlosa As String, ByVal psMovNro As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_InsertaDatosImprVoucher " & nMovNro & ",'" & psGlosa & "','" & psMovNro & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub
'**End Juez *******************

'** JUEZ 20120910 *******************************************************
Public Sub GrabarCredEvaluacionFormatoIyII(ByVal psCtaCod As String, ByVal pnInsertActualiza As Integer, ByVal psGiroNeg As String, ByVal pnExpEmpAnio As Integer, _
                                        ByVal pnExpEmpMes As Integer, ByVal pnTiempoLocalAnio As Integer, ByVal pnTiempoLocalMes As Integer, ByVal pnCondLocal As Integer, _
                                        ByVal psCondLocalOtros As String, ByVal pnCuotaPagar As Double, ByVal pnCuotas As Integer, _
                                        ByVal pnUltEndeuda As Double, ByVal psFecUltEndeuda As String, ByVal pnMoneda As Integer, ByVal pnMontoSol As Double, _
                                        ByVal pnVentaProm As Double, ByVal pnCostoVenta As Double, ByVal pnUtilNeta As Double, ByVal pnExcedenteFam As Double, _
                                        ByVal pnCalcMonto As Double, ByVal pnCalcTEM As Double, ByVal pnCalcCuotas As Double, ByVal pnMontoMax As Double, _
                                        ByVal pnCuotaEstima As Double, ByVal pnCuotaUNM As Double, ByVal pnCuotaExcedeFam As Double, ByVal psComent As String, _
                                        Optional ByVal pnFormato As Integer)
                                
Dim lsSQL As String

    If pnInsertActualiza = 1 Then
        lsSQL = "exec stp_ins_GrabarCredEvaluacionFormatoIyII '"
    Else
        lsSQL = "exec stp_upd_GrabarCredEvaluacionFormatoIyII '"
    End If
    lsSQL = lsSQL & psCtaCod & "','" & psGiroNeg & "'," & pnExpEmpAnio & "," & pnExpEmpMes & "," & pnTiempoLocalAnio & "," & _
                    pnTiempoLocalMes & "," & pnCondLocal & ",'" & psCondLocalOtros & "'," & pnCuotaPagar & "," & pnCuotas & "," & _
                    pnUltEndeuda & ",'" & psFecUltEndeuda & "'," & pnMoneda & "," & pnMontoSol & "," & pnVentaProm & "," & pnCostoVenta & "," & _
                    pnUtilNeta & "," & pnExcedenteFam & "," & pnCalcMonto & "," & pnCalcTEM & "," & pnCalcCuotas & "," & pnMontoMax & "," & _
                    pnCuotaEstima & "," & pnCuotaUNM & "," & pnCuotaExcedeFam & ",'" & psComent & "'," & pnFormato
                    
    coConex.Ejecutar lsSQL
End Sub

Public Sub dInsertaCredEvalDatosGrillas(ByVal psCtaCod As String, ByVal psConcepto As String, _
                                        ByVal pnMonto As Double, ByVal pnTipoGrilla As Integer, _
                                        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT INTO ColocacCredEvalGrillas(cCtaCod, cConcepto, nMonto, nTipoGrilla) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "', '" & psConcepto & "'," & pnMonto & "," & pnTipoGrilla & ")"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dEliminaCredEvalDatosGrillas(ByVal psCtaCod As String, ByVal pnTipoGrilla As Integer)
Dim sSql As String

    sSql = " DELETE FROM ColocacCredEvalGrillas Where cCtaCod = '" & psCtaCod & "' AND nTipoGrilla = " & pnTipoGrilla
    coConex.Ejecutar sSql
    
End Sub

Public Sub dInsertaCredEvalDatosVentasSemana(ByVal psCtaCod As String, ByVal psProducto As String, ByVal pnVentaAlta As Double, _
                                                            ByVal pnDiasAlta As Integer, ByVal pnVentaBaja As Double, ByVal pnDiasBaja As Integer, _
                                                            ByVal pnTotalMes As Double, ByVal pnCosto As Double, ByVal pnParticip As Double, ByVal pnReal As Double)

Dim lsSQL As String

    lsSQL = "INSERT INTO ColocacCredEvalVentasSemana(cCtaCod, cProducto, nVentaAlta, nDiasAlta, nVentaBaja, nDiasBaja, nTotalMes, nCosto, nParticip, nReal) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "', '" & psProducto & "'," & pnVentaAlta & "," & pnDiasAlta & "," & pnVentaBaja & "," & _
                                pnDiasBaja & "," & pnTotalMes & "," & pnCosto & "," & pnParticip & "," & pnReal & ")"

    coConex.Ejecutar lsSQL


End Sub

Public Sub dEliminaCredEvalDatosVentasSemana(ByVal psCtaCod As String)
Dim sSql As String

    sSql = " DELETE FROM ColocacCredEvalVentasSemana Where cCtaCod = '" & psCtaCod & "'"
    coConex.Ejecutar sSql
    
End Sub

Public Sub dInsertaCredEvalDatosReferencia(ByVal psCtaCod As String, ByVal psNombre As String, ByVal psDNI As String, _
                                            ByVal psTelef As String, ByVal psReferido As String, ByVal psDNI2 As String, _
                                            Optional ByVal pbEjecBatch As Boolean = False)

    Dim lsSQL As String

    lsSQL = "INSERT INTO ColocacCredEvalReferencia(cCtaCod,cNombre,cDNI,cTelef,cReferido,cDNIRef) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "','" & psNombre & "','" & psDNI & "','" & psTelef & "','" & psReferido & "','" & psDNI2 & "')"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dEliminaCredEvalDatosReferencia(ByVal psCtaCod As String)
Dim sSql As String

    sSql = " DELETE FROM ColocacCredEvalReferencia Where cCtaCod = '" & psCtaCod & "'"
    coConex.Ejecutar sSql
    
End Sub

Public Sub GrabarCredEvaluacionVerif(ByVal psCtaCod As String, ByVal psVerif As String, ByVal psMovNro As String, Optional ByVal pbEjecBatch As Boolean = False)
    Dim lsSQL As String

    lsSQL = "exec stp_upd_GrabarCredEvaluacionVerif '" & psCtaCod & "','" & psVerif & "','" & psMovNro & "'"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dEliminaCredEvalParametrosTipo()
Dim sSql As String

    sSql = " DELETE FROM ColocacCredEvalParamTipo"
    coConex.Ejecutar sSql
    
End Sub

Public Sub dInsertaCredEvalParametrosTipo(ByVal psColocEvalId As String, ByVal psTpoProdCod As String, ByVal psTpoSubProdCod As String, _
                                            ByVal pnMin As Double, ByVal pnMax As Double, ByVal pnTpoEval As Integer, _
                                            Optional ByVal pbEjecBatch As Boolean = False)

    Dim lsSQL As String

    lsSQL = "INSERT INTO ColocacCredEvalParamTipo(cColocEvalId,cTpoProdCod,cTpoSubProdCod,nMin,nMax,nTpoEval) "
    lsSQL = lsSQL & " VALUES('" & psColocEvalId & "','" & psTpoProdCod & "','" & psTpoSubProdCod & "'," & pnMin & "," & pnMax & "," & pnTpoEval & ")"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dEliminaCredEvalParametrosIndicador()
Dim sSql As String

    sSql = " DELETE FROM ColocacCredEvalParamIndicador"
    coConex.Ejecutar sSql
    
End Sub

Public Sub dInsertaCredEvalParametrosIndicador(ByVal psIndID As String, ByVal psIndDesc As String, ByVal pnIndFormato As Integer, _
                                            ByVal pnIndCondicion As Integer, ByVal pnIndAcepMin As Double, ByVal pnIndAcepMax As Double, _
                                            ByVal pnIndCritMin As Double, ByVal pnIndCritMax As Double, _
                                            Optional ByVal pbEjecBatch As Boolean = False)

    Dim lsSQL As String

    lsSQL = "INSERT INTO ColocacCredEvalParamIndicador(cIndicadorID, cIndicadorDesc, cIndicadorCondicion, cIndicadorFormato, cIndicadorAcepMin, cIndicadorAcepMax, cIndicadorCritMin, cIndicadorCritMax) "
    lsSQL = lsSQL & " VALUES('" & psIndID & "','" & psIndDesc & "'," & pnIndCondicion & "," & pnIndFormato & "," & pnIndAcepMin & "," & pnIndAcepMax & "," & pnIndCritMin & "," & pnIndCritMax & ")"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
'** END JUEZ ************************************************************
'WIOR 20120918 **********************************************************
Public Sub dInsertaCredEvalGastoNegocio(ByVal psCtaCod As String, ByVal psConcepto As String, _
                                        ByVal pnTipoGasto As Double, ByVal pnMonto As Integer, _
                                        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "INSERT INTO ColocacCredEvalGastosNeg(cCtaCod, cConcepto, nTpoGasto, nMonto) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "', '" & psConcepto & "'," & pnTipoGasto & "," & pnMonto & ")"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
Public Sub dEliminaCredEvalGastoNegocio(ByVal psCtaCod As String)
Dim sSql As String

    sSql = " DELETE FROM ColocacCredEvalGastosNeg Where cCtaCod = '" & psCtaCod & "' "
    coConex.Ejecutar sSql
    
End Sub
Public Sub dEliminaCredEvalActivosPasivos(ByVal psCtaCod As String, ByVal pnTipo As Integer)
Dim sSql As String

    sSql = " DELETE FROM ColocacCredEvalActivosPasivos Where cCtaCod = '" & psCtaCod & "' AND nTpoActPas = " & pnTipo
    coConex.Ejecutar sSql
    
End Sub
Public Sub dInsertaCredEvalActivosPasivos(ByVal psCtaCod As String, ByVal nTipo As Integer, _
                                        ByVal psConcepto As String, ByVal pnPE As Double, ByVal pnPP As Double, _
                                        ByVal pnTotal As Double, Optional ByVal pcDesc As String = "", _
                                        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

    
    lsSQL = "INSERT INTO ColocacCredEvalActivosPasivos(cCtaCod,nTpoActPas, cConcepto, nPE,nPP, nTOTAL,cDesc) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "'," & nTipo & ", '" & psConcepto & "'," & pnPE & "," & pnPP & "," & pnTotal & ",'" & pcDesc & "')"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dEliminaCredEvalPDT(ByVal psCtaCod As String)
Dim sSql As String

    sSql = " DELETE FROM ColocacCredEvalPDT Where cCtaCod = '" & psCtaCod & "'"
    coConex.Ejecutar sSql
    
    sSql = " DELETE FROM ColocacCredEvalPDTDet Where cCtaCod = '" & psCtaCod & "'"
    coConex.Ejecutar sSql
End Sub
Public Sub dInsertaCredEvalPDT(ByVal psCtaCod As String, _
                                        ByVal pnMes1 As Integer, ByVal pnMes2 As Integer, ByVal pnMes3 As Integer, _
                                        ByVal pnAnio1 As Integer, ByVal pnAnio2 As Integer, ByVal pnAnio3 As Integer, _
                                        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

    
    lsSQL = "INSERT INTO ColocacCredEvalPDT(cCtaCod,nMes1, nAnio1, nMes2,nAnio2, nMes3,nAnio3) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "'," & pnMes1 & ", " & pnAnio1 & "," & pnMes2 & "," & pnAnio2 & "," & pnMes3 & "," & pnAnio3 & ")"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dInsertaCredEvalPDTDet(ByVal psCtaCod As String, ByVal nTipo As Integer, _
                                        ByVal pnMontoMes1 As Double, ByVal pnMontoMes2 As Double, ByVal pnMontoMes3 As Double, _
                                        ByVal pnPromedio As Double, _
                                        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

    
    lsSQL = "INSERT INTO ColocacCredEvalPDTDet(cCtaCod,nTipo, nMontoMes1, nMontoMes2,nMontoMes3, nPromedio) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "'," & nTipo & ", " & pnMontoMes1 & "," & pnMontoMes2 & "," & pnMontoMes3 & "," & pnPromedio & ")"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
Public Sub dEliminaCredEvalEstadosGP(ByVal psCtaCod As String)
Dim sSql As String

    sSql = " DELETE FROM ColocacCredEvalEstadosGP Where cCtaCod = '" & psCtaCod & "'"
    coConex.Ejecutar sSql
    
End Sub
Public Sub dInsertaCredEvalEstadosGP(ByVal psCtaCod As String, _
                                        ByVal pnIngresos As Double, ByVal pnEgresos As Double, ByVal pnMargen As Double, _
                                        ByVal pnGastos As Double, ByVal pnIngresoNeto As Double, _
                                        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

    
    lsSQL = "INSERT INTO ColocacCredEvalEstadosGP(cCtaCod, nIngresos, nEgresos,nMargen, nGastos,nIngresoNeto) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "'," & pnIngresos & ", " & pnEgresos & "," & pnMargen & "," & pnGastos & "," & pnIngresoNeto & ")"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dEliminaCredEvalRatios(ByVal psCtaCod As String)
Dim sSql As String

    sSql = " DELETE FROM ColocacCredEvalRatios Where cCtaCod = '" & psCtaCod & "'"
    coConex.Ejecutar sSql
    
End Sub
Public Sub dInsertaCredEvalRatios(ByVal psCtaCod As String, ByVal pnTipo As Integer, _
                                        ByVal pnEndCortoPlazo As Double, ByVal pnEndLargoPlazo As Double, ByVal pnLiqCorriente As Double, _
                                        ByVal pnCapTrabajo As Double, ByVal pnRotInventario As Double, _
                                        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

    
    lsSQL = "INSERT INTO ColocacCredEvalRatios(cCtaCod,nTipo, nEndCortoPlazo, nEndLargoPlazo,nLiqCorriente, nCapTrabajo,nRotInventario) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "'," & pnTipo & ", " & pnEndCortoPlazo & "," & pnEndLargoPlazo & "," & pnLiqCorriente & "," & pnCapTrabajo & "," & pnRotInventario & ")"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
Public Sub dEliminaCredEvalIngNeg(ByVal psCtaCod As String)
Dim sSql As String

    sSql = " DELETE FROM ColocacCredEvalIngNeg Where cCtaCod = '" & psCtaCod & "'"
    coConex.Ejecutar sSql
    
End Sub
Public Sub dInsertaCredEvalIngNeg(ByVal psCtaCod As String, ByVal pnTipo As Integer, _
                                        ByVal pnProd1 As Double, ByVal pnProd2 As Double, ByVal pnProd3 As Double, _
                                        ByVal pnResultado As Double, _
                                        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

    
    lsSQL = "INSERT INTO ColocacCredEvalIngNeg(cCtaCod,nTipo, nProd1, nProd2,nProd3, nResultado) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "'," & pnTipo & ", " & pnProd1 & "," & pnProd2 & "," & pnProd3 & "," & pnResultado & ")"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dEliminaCredEvalTasas(ByVal psCtaCod As String)
Dim sSql As String

    sSql = " DELETE FROM ColocacCredEvalTasasIngEg Where cCtaCod = '" & psCtaCod & "'"
    coConex.Ejecutar sSql
    
End Sub
Public Sub dInsertaCredEvalTasas(ByVal psCtaCod As String, ByVal pnTipo As Integer, ByVal pnTasa As Double, _
                                        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

    
    lsSQL = "INSERT INTO ColocacCredEvalTasasIngEg(cCtaCod,nTpoTasa, nTasa) "
    lsSQL = lsSQL & " VALUES('" & psCtaCod & "'," & pnTipo & ", " & pnTasa & ")"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
Public Sub GrabarCredEvaluacionFormato3y4(ByVal psCtaCod As String, ByVal pnInsertActualiza As Integer, ByVal psGiroNeg As String, _
                                        ByVal pnCuotaPagar As Double, ByVal pnCuotas As Integer, ByVal pnMoneda As Integer, ByVal pnMontoSol As Double, _
                                        ByVal pnCalcMonto As Double, ByVal pnCalcTEM As Double, ByVal pnCalcCuotas As Double, ByVal pnMontoMax As Double, _
                                        ByVal pnCuotaEstima As Double, ByVal pnEndeudTotal As Double, ByVal pnCapPagoEmp As Double, ByVal pnCapPagoTotal As Double, _
                                        ByVal pnFormato As Integer, ByVal psComent As String)
                                
Dim lsSQL As String

 
    If pnInsertActualiza = 1 Then
        lsSQL = "exec stp_ins_GrabarCredEvaluacionFormato3Y4 '"
    Else
        lsSQL = "exec stp_upd_GrabarCredEvaluacionFormato3y4 '"
    End If
    
    lsSQL = lsSQL & psCtaCod & "','" & psGiroNeg & "'," & pnCuotaPagar & "," & pnCuotas & "," & pnMoneda & "," & _
            pnMontoSol & "," & pnCalcMonto & "," & pnCalcTEM & "," & pnCalcCuotas & "," & pnMontoMax & "," & pnCuotaEstima & "," & _
            pnEndeudTotal & "," & pnCapPagoEmp & "," & pnCapPagoTotal & "," & pnFormato & ",'" & psComent & "'"
                    
    coConex.Ejecutar lsSQL
End Sub

Public Sub dEliminaCredEvalParametrosEspecializacion()
Dim sSql As String

    sSql = " DELETE FROM ColocacCredEvalEspecializacion"
    coConex.Ejecutar sSql
    
End Sub


Public Sub dInsertaCredEvalParametrosEspecializacion(ByVal psColocEspId As String, ByVal pnEsp As Integer, ByVal psTpoProdCod As String, ByVal psTpoSubProdCod As String, _
                                            ByVal psTpoCredCod As String, ByVal psTpoSubCredCod As String, ByVal pnUltEnd As Integer, ByVal pnEndMin As Double, ByVal pnEndMax As Double, _
                                            Optional ByVal pbEjecBatch As Boolean = False)

    Dim lsSQL As String

    lsSQL = "INSERT INTO ColocacCredEvalEspecializacion(cColocEspId,nTpoEsp,cTpoProdCod,cTpoSubProdCod,cTpoCredCod,cTpoSubCredCod,nUltEndeud,nEndeudCredMin,nEndeudCredMax) "
    lsSQL = lsSQL & " VALUES('" & psColocEspId & "'," & pnEsp & ",'" & psTpoProdCod & "','" & psTpoSubProdCod & "','" & psTpoCredCod & "','" & psTpoSubCredCod & "'," & pnUltEnd & "," & pnEndMin & "," & pnEndMax & ")"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

'WIOR FIN*****************************************************************************

'**JUEZ 20120921 ********************************************************
Public Sub dInsertaMotivoCancAnticipada(ByVal psCtaCod As String, ByVal pnMotivo As Integer, ByVal psOtros As String)

    Dim lsSQL As String

    lsSQL = "exec stp_ins_GrabarMotivoCancAnticipada '" & psCtaCod & "'," & pnMotivo & ",'" & psOtros & "'"

    coConex.Ejecutar lsSQL

End Sub
'**END JUEZ *************************************************************

'JUEZ 20121213 ***********************************************************************
Public Sub dInsertaCredNivAprResultado(ByVal psCtaCod As String, ByVal psNivAprCod As String, ByVal pnCuotas As Integer, _
                                       ByVal pnTasa As Double, ByVal pnMonto As Double, ByVal psComent As String, _
                                       ByVal psMovApr As String, Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

On Error GoTo ErrordInsertaCredNivAprResultado
    lsSQL = "exec stp_ins_dInsertaCredNivAprResultado '" & psCtaCod & "','" & psNivAprCod & "'," & pnCuotas & "," & pnTasa _
                                                         & "," & pnMonto & ",'" & psComent & "','" & psMovApr & "'"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
    Exit Sub

ErrordInsertaCredNivAprResultado:
    Err.Raise Err.Number, "Error En Proceso dInsertaCredNivAprResultado", Err.Description

End Sub

Public Sub dUpdateColocacEstadoUltimoRegNivApr(ByVal psCtaCod As String, ByVal pdFechaPrdEstado As String, _
        ByVal pnPrdEstado As Integer, ByVal pnCuotas As Integer, ByVal pnMonto As Currency, _
        ByVal psDescripcion As String, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
Dim rs As ADODB.Recordset
    
    On Error GoTo ErrordUpdateColocacEstado
    
    Set rs = coConex.CargaRecordSet("SELECT max(dPrdEstado) dPrdEstado FROM ColocacEstado WHERE cCtaCod = '" & psCtaCod & "' AND nPrdEstado = " & pnPrdEstado & " AND cDescripcion like '%por Niv. Aprob.'")
    
    lsSQL = " UPDATE ColocacEstado SET "
    lsSQL = lsSQL & " dPrdEstado = '" & Format(CDate(Format(pdFechaPrdEstado, "dd/mm/yyyy") & Space(1) & Format(dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "',"
    lsSQL = lsSQL & " nCuotas = " & pnCuotas & ","
    lsSQL = lsSQL & " NMonto = " & Format(pnMonto, "#0.00") & ","
    lsSQL = lsSQL & " cDescripcion = '" & Replace(psDescripcion, "'", "''") & "'"
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "' AND nPrdEstado = " & pnPrdEstado
    lsSQL = lsSQL & " AND dPrdEstado = '" & Format(CDate(Format(rs!dPrdEstado, "dd/mm/yyyy") & Space(1) & Format(rs!dPrdEstado, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "'"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

    Exit Sub

ErrordUpdateColocacEstado:
    Err.Raise Err.Number, "Error En Proceso dUpdateColocacEstado", Err.Description

End Sub

Public Sub dInsertColocacEstadoNivApr(ByVal psCtaCod As String, ByVal pdFechaPrdEstado As String, _
        ByVal pnPrdEstado As Integer, ByVal pnCuotas As Integer, ByVal pnMonto As Currency, _
        ByVal psDescripcion As String, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
Dim rs As ADODB.Recordset
    
    On Error GoTo ErrorInsertColocacEstado

    lsSQL = "INSERT INTO ColocacEstado (cCtaCod,dPrdEstado,nPrdEstado,nCuotas,nMonto,cDescripcion,nColocCalendCod,nPeriodoFechaFija,nPeriodoGracia,nPlazo,nTipoGracia,nTipoDesembolso,nProxMes,nCalendDinamico,nMotivoRechazo,cTipoGasto,nPeriodoFechaFija2,bIncremGraciaCap,nMontoCol,bGraciaEnCuotas,nTasaGracia) " 'JUEZ 20131223
    lsSQL = lsSQL & "SELECT '" & psCtaCod & "','" & Format(CDate(Format(pdFechaPrdEstado, "dd/mm/yyyy") & Space(1) & Format(dFechaHora, "hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "'," & pnPrdEstado & "," & pnCuotas & "," & pnMonto & ",'" & psDescripcion & "', "
    lsSQL = lsSQL & "nColocCalendCod,nPeriodoFechaFija,nPeriodoGracia,nPlazo,nTipoGracia,nTipoDesembolso,nProxMes, "
    lsSQL = lsSQL & "nCalendDinamico,nMotivoRechazo,cTipoGasto,nPeriodoFechaFija2,bIncremGraciaCap,nMontoCol,bGraciaEnCuotas,nTasaGracia "
    lsSQL = lsSQL & "FROM ColocacEstado WHERE cCtaCod='" & psCtaCod & "' and nprdestado=2001 "
    lsSQL = lsSQL & "and dprdestado=(select MAX(dprdestado) from ColocacEstado "
    lsSQL = lsSQL & "                        WHERE cCtaCod='" & psCtaCod & "' and nprdestado=2001 ) "
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

    Exit Sub

ErrorInsertColocacEstado:
    Err.Raise Err.Number, "Error En Proceso dInsertColocacEstadoNivApr", Err.Description

End Sub

Public Function ExisteColocacEstadoNivApr(ByVal psCtaCod As String) As Boolean
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset

    ExisteColocacEstadoNivApr = False
    
    sSql = "stp_sel_ExisteColocacEstadoNivApr '" & psCtaCod & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    If Not (rs.EOF And rs.BOF) Then
        ExisteColocacEstadoNivApr = True
    End If
    
    Set rs = Nothing
End Function

Public Sub dDeleteCredSugExonera(ByVal psCtaCod As String)
Dim lsSQL As String

    lsSQL = "DELETE FROM CredSugExonera WHERE cCtaCod='" & psCtaCod & "'"
    coConex.Ejecutar lsSQL
    
End Sub

Public Sub dInsertaCredSugExonera(ByVal psCtaCod As String, ByVal psExoneraCod As String)
Dim lsSQL As String
Dim rs As ADODB.Recordset

    Set rs = coConex.CargaRecordSet("exec stp_sel_RecuperaNivelesAprobDelCredito '" & psCtaCod & "','" & psExoneraCod & "'")
    Do While Not rs.EOF
        lsSQL = "INSERT INTO CredSugExonera (cCtaCod,cExoneraCod,cNivAprCod,nEstado) "
        lsSQL = lsSQL & "VALUES ('" & psCtaCod & "','" & psExoneraCod & "','" & rs!cNivAprCod & "',0)"
        coConex.Ejecutar lsSQL
        rs.MoveNext
    Loop
End Sub

Public Sub dInsertProductoPersonaNivAprApoderado(ByVal psCuenta As String, ByVal psUser As String)

Dim lsSQL As String
Dim psPersCod As String
Dim rs As ADODB.Recordset
    Set rs = coConex.CargaRecordSet("Select cPersCod From RRHH Where cUser='" & psUser & "'")
    psPersCod = rs!cPersCod
    Set rs = coConex.CargaRecordSet("Select cPersCod From ProductoPersona where cCtaCod= '" & psCuenta & "' and nPrdPersRelac=" & gColRelPersApoderado)

    If Not rs.EOF Then
        lsSQL = "UPDATE ProductoPersona set cPersCod='" & psPersCod & _
                "' WHERE cCtaCod='" & psCuenta & "' and nPrdPersRelac=" & gColRelPersApoderado
    Else
        lsSQL = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
            & "VALUES ('" & psCuenta & "','" & psPersCod & "'," & gColRelPersApoderado & ")"
    End If

    coConex.Ejecutar lsSQL
End Sub

'END JUEZ ****************************************************************************

'AMDO 20130319 CREDIPREMIAZO ***************************************************************************************
Public Sub InsertaPuntosCrediPremiazoPagoNormal(ByVal pnMovNro As Long, Optional ByVal pnDiasAtras As Integer = 0)
    Dim sSql As String
    sSql = "exec stp_ins_PuntosCrediPremiazoPagoLote " & pnMovNro & "," & pnDiasAtras
    coConex.Ejecutar sSql
End Sub

Public Sub ExtornaPuntosCrediPremiazo(ByVal pnMovNro As Long, ByVal psCtaCod As String)
    Dim sSql As String
    sSql = "exec stp_upd_ExtornosPuntosCrediPremiazo " & pnMovNro & ",'" & psCtaCod & "'"
    coConex.Ejecutar sSql
End Sub

Public Function DevuelveCantPuntosMovCredito(ByVal pnMovNro As Long, ByVal psCtaCod As String) As Integer
Dim R As ADODB.Recordset
Dim sSql As String

    sSql = "exec stp_sel_PuntosTotalMovCredito " & pnMovNro & ",'" & psCtaCod & "'"
    Set R = coConex.CargaRecordSet(sSql)
    If Not R.EOF Then
        DevuelveCantPuntosMovCredito = R!TotalPuntos
    Else
        DevuelveCantPuntosMovCredito = 0
    End If
End Function

Public Function DevuelveCantPuntosTotalCredito(ByVal psCtaCod As String) As Integer
    Dim R As ADODB.Recordset
    Dim sSql As String
    sSql = "exec stp_sel_PuntosTotalAcumuladosCrediPremiazo '" & psCtaCod & "'"
    Set R = coConex.CargaRecordSet(sSql)

    If Not R.EOF Then
        DevuelveCantPuntosTotalCredito = R!TotalPuntos
    Else
        DevuelveCantPuntosTotalCredito = 0
    End If
End Function
'END AMDO **********************************************************************************************************
'JUEZ 20130411 **********************************************************************************
Public Sub dInsertComision(ByVal pnMovNro As Long, ByVal psNumDoc As String, ByVal pnComision As Double, Optional ByVal pnCuota As Integer = 0, _
                           Optional ByVal psCodCom As String = "")
'JUEZ 20130529 Se agregó pnCuota
'JUEZ 20150930 se agregó psCodCom
Dim lsSQL As String
    'lsSQL = "exec stp_sel_InsertaMovComision " & pnMovNro & ",'" & psNumDoc & "'," & pnComision & "," & pnCuota
    lsSQL = "exec stp_sel_InsertaMovComision " & pnMovNro & ",'" & psNumDoc & "'," & pnComision & "," & pnCuota & ",'" & psCodCom & "'" 'JUEZ 20150930

    coConex.Ejecutar lsSQL
End Sub
Public Sub dUpdateComision(ByVal psNumDoc As String, ByVal psOpeCod As CaptacOperacion)
Dim lsSQL As String
    lsSQL = "exec stp_sel_ActualizaMovComision '" & psNumDoc & "','" & psOpeCod & "'"
    
    coConex.Ejecutar lsSQL
End Sub
'END JUEZ ***************************************************************************************
'ALPA20130801************************************************************************************
Public Sub dInsertaMovMiVivienda(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, ByVal pnCuota As Integer, _
                                       ByVal pnPrdConceptoCod As Integer, ByVal pnMovNro As Long, ByVal pnMontoPagado As Currency, _
                                       Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String

On Error GoTo ErrordInsertaMovMiVivienda
    lsSQL = "exec stp_ins_MovMiVivienda '" & psCtaCod & "'," & pnNroCalen & "," & pnCuota & "," & pnPrdConceptoCod _
                                                         & "," & pnMovNro & "," & pnMontoPagado & ""

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
    Exit Sub

ErrordInsertaMovMiVivienda:
    Err.Raise Err.Number, "Error En Proceso dInsertaMovMiVivienda", Err.Description

End Sub
'************************************************************************************************
'EJVG20130920 ***
Public Sub ActualizadUltContactoAho(ByVal sCuenta As String, ByVal dFecha As Date)
    Dim sSql As String
    sSql = "Update CaptacAhorros Set dUltContacto = '" & Format(dFecha, "yyyymmdd") & "'  WHERE cCtaCod = '" & sCuenta & "'"
    coConex.ConexionActiva.Execute sSql
End Sub
'END EJVG *******
'************************************************************************************************
'ALPA20130918*****************************************
Public Sub dExtornoLeasing(ByVal psCtaCod As String, ByVal psCuota As String, ByVal pdFecha As Date, Optional pbEjecBatch As Boolean = False)
Dim sSql As String

    sSql = "exec SAF_stp_ins_ExtornarCuotasLeasing '" & Format(pdFecha, "YYYY/MM/DD") & "','" & psCtaCod & "', '" & psCuota & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If
End Sub
'**********************************
'JUEZ 20130925 *********************************************************
Public Sub dInsertExonerarPreCanc(ByVal psCtaCod As String, ByVal pdFecha As Date, ByVal pnSaldoCanc As Double, _
                                  ByVal pnMontoCom As Double, ByVal psGlosa As String, ByVal psMovNro As String)
Dim lsSQL As String
    lsSQL = "exec stp_ins_ExonerarPreCancelacion '" & psCtaCod & "','" & Format(pdFecha, "yyyymmdd") & "'," & pnSaldoCanc _
                                                & "," & pnMontoCom & ",'" & psGlosa & "','" & psMovNro & "'"

    coConex.Ejecutar lsSQL
End Sub
'END JUEZ **************************************************************
'JUEZ 20131022 *********************************************************
Public Sub dInsertaCredReprogNatEspeciales(ByVal pnMovNro As Long, ByVal psCtaCod As String)
Dim lsSQL As String
    lsSQL = "exec stp_ins_CredReprogNatEspeciales " & pnMovNro & ",'" & psCtaCod & "'"

    coConex.Ejecutar lsSQL
End Sub
Public Sub dUpdateCredReprogNatEspeciales(ByVal pnMovNro As Long, ByVal pnReversion As Integer)
Dim lsSQL As String
    lsSQL = "exec stp_upd_CredReprogNatEspeciales " & pnMovNro & "," & pnReversion

    coConex.Ejecutar lsSQL
End Sub
'END JUEZ **************************************************************

'*********************RECO 20131122 ERS055************************************************
Public Sub MantenimientoGastoPolizaContraincendio(ByVal sCtaCod As String, ByVal sUser As String, _
                                                  ByVal sAgeCod As String, ByVal sFechaReg As String, _
                                                  ByVal nCuotaDesde As Integer, ByVal nCuotaHasta As Integer, _
                                                  ByVal nPrimaNeta As Double, ByVal nPrimaTotal As Double, _
                                                  Optional ByVal nPrimaTotalCT As Double = 0, Optional ByVal nMontoPrimaCuota As Double = 0, _
                                                  Optional ByVal sNumCertificado As String = "", Optional ByVal sPersCodAseguradora As String = "")
Dim lsSQL As String
    lsSQL = "exec stp_ins_MantenimientoGastoPolizaContraincendio '" & sCtaCod & "','" & sUser & "','" & sAgeCod & "','" _
    & sFechaReg & "'," & nCuotaDesde & "," & nCuotaHasta & "," & nPrimaNeta & "," & nPrimaTotal & "," _
    & nPrimaTotalCT & "," & nMontoPrimaCuota & ",'" & sNumCertificado & "','" & sPersCodAseguradora & "'"

    coConex.Ejecutar lsSQL
End Sub
Public Function ValidaNumeroCertificado(ByVal sPersCodAseg As String, ByVal sNumCertificado As String) As Boolean
    Dim lsSQL As String
    Dim R As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ObtieneNumeroCertificado '" & sPersCodAseg & "','" & sNumCertificado & "'"
    Set R = coConex.CargaRecordSet(lsSQL)
    If Not R.EOF Then
        ValidaNumeroCertificado = True
    Else
        ValidaNumeroCertificado = False
    End If
End Function
'*********************END RECO************************************************************

'JUEZ 20131116 *******************************************************************
Public Sub dExtornaCredReprogNatEspeciales(ByVal psCtaCod As String, ByVal psOpeCodReprog, ByVal pnReversion As Integer)
    Dim sSql As String
    sSql = "exec stp_upd_ExtornaCredReprogNatEspeciales '" & psCtaCod & "','" & psOpeCodReprog & "'," & pnReversion
    coConex.Ejecutar sSql
End Sub
'END JUEZ ************************************************************************
'WIOR 20140128 **********************************************************
Public Sub dDeleteCredVinculados(ByVal psCtaCod As String, ByVal pnEstado As Integer)
Dim lsSQL As String
On Error GoTo Error
    
    lsSQL = "EXEC stp_del_CredAsignacionSaldo '" & psCtaCod & "'," & pnEstado
    
    coConex.Ejecutar lsSQL

    Exit Sub
Error:
    Err.Raise Err.Number, "Error en Proceso (dDeleteColocConvBcoNac)", Err.Description
End Sub
Public Sub dInsertCredVinculados(ByVal psCtaCod As String, ByVal pnMonto As Double, ByVal psMovNroSol As String, ByVal pnTipoVinc As Integer, ByVal pnEstado As Integer)
Dim lsSQL As String
On Error GoTo Error
    
    lsSQL = "EXEC stp_ins_CredAsignacionSaldo '" & psCtaCod & "'," & pnMonto & ",'" & psMovNroSol & "'," & pnTipoVinc & "," & pnEstado
    
    coConex.Ejecutar lsSQL

    Exit Sub
Error:
    Err.Raise Err.Number, "Error en proceso (dInsertCredVinculados)", Err.Description
End Sub
Public Sub ActualizaCredVinculados(ByVal psCtaCod As String, ByVal psMovAprueba As String, ByVal pnEstadoA As Integer, ByVal pnEstadoB As Integer, Optional ByVal pnNoTransac As Boolean = False)
Dim lsSQL As String
On Error GoTo Error

    lsSQL = "EXEC stp_upd_CredAsignacionSaldo '" & psCtaCod & "','" & psMovAprueba & "'," & pnEstadoA & "," & pnEstadoB
    
    coConex.Ejecutar lsSQL

    Exit Sub
Error:
    Err.Raise Err.Number, "Error en Proceso(ActualizaCredVinculados)", Err.Description
End Sub
'WIOR FIN ******************************************************************
'***** FRHU 20140228 RQ14006
Public Sub dDeleteProductoBloqueo(ByVal psMovNro As String)
Dim sSql  As String

    sSql = "DELETE ProductoBloqueos WHERE cMovNro = '" & psMovNro & "'"
    coConex.Ejecutar sSql

End Sub
Public Function dGetcMovNro(ByVal pnMovNro As Long) As String
Dim lsSQL As String
Dim lrs As ADODB.Recordset
Set lrs = New ADODB.Recordset

lsSQL = "Select cMovNro From Mov where nMovNro = " & pnMovNro
Set lrs = coConex.CargaRecordSet(lsSQL, adLockReadOnly)
'Set lrs.ActiveConnection = coConex.ConexionActiva
If Not lrs.EOF And Not lrs.BOF Then
    dGetcMovNro = lrs!cMovnro
End If
lrs.Close
Set lrs = Nothing
End Function
'FRHU 20140228 RQ14006
'FRHU 20140301 RQ14006
Public Sub AgregarRelaAhoCred(ByVal psCtaAhorro As String, ByVal psCtaCred As String, ByVal psMovNroBlo As String, ByVal pbEstado As Boolean)
Dim sSql As String
sSql = "INSERT RelaAhoCred (cCtaCodAho,cCtaCodCred,cMovNroBlo,bEstado) " _
    & "VALUES ('" & psCtaAhorro & "','" & psCtaCred & "','" & psMovNroBlo & "'," & IIf(pbEstado = True, 1, 0) & ")"
coConex.ConexionActiva.Execute sSql
End Sub
'FIN FRHU 20140301 RQ14006
'EJVG20140325 ***

Public Sub dInsertColocCalendarioxMigracion(ByVal psCtaCod As String, ByVal pnNroCalenNew As Integer, ByVal pnNroCalenAnt As Integer, Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pnTipoReprogCovid As Integer = 0, Optional ByVal pnCantCuotasNew As Integer = 0) 'Add joep20200526 MAntener cuota Covid
'Public Sub dInsertColocCalendarioxMigracion(ByVal psCtaCod As String, ByVal pnNroCalenNew As Integer, ByVal pnNroCalenAnt As Integer, Optional ByVal pbEjecBatch As Boolean = False) 'Comento joep20200526 MAntener cuota Covid
    Dim lsSQL As String
    
    'JOEP20200327 Covid-19 comento
    'lsSQL = "Insert ColocCalendario " _
     '       & "Select cCtaCod, " & pnNroCalenNew & ",nColocCalendApl,nCuota,dVenc,dPago,nColocCalendEstado,cDescripcion,cColocCalenFlag,nCalendProc,cColocMiVivEval " _
      '      & "from ColocCalendario Where cCtaCod = '" & psCtaCod & "' And nNroCalen = " & pnNroCalenAnt
    'JOEP20200327 Covid-19 comento
    
    'lsSQL = "exec dInsertColocCalendarioxMigracion '" & psCtaCod & "'," & pnNroCalenAnt & "," & pnNroCalenNew & "" 'JOEP20200327 Covid-19 add
    lsSQL = "exec dInsertColocCalendarioxMigracion '" & psCtaCod & "'," & pnNroCalenAnt & "," & pnNroCalenNew & "," & pnTipoReprogCovid & "," & pnCantCuotasNew & "" 'JOEP20200327 Covid-19 add
        
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

'Public Sub dInsertColocCalendDetxMigracion(ByVal psCtaCod As String, ByVal pnNroCalenNew As Integer, ByVal pnNroCalenAnt As Integer, Optional ByVal pbEjecBatch As Boolean = False) 'Comento joep20200526 MAntener cuota Covid
Public Sub dInsertColocCalendDetxMigracion(ByVal psCtaCod As String, ByVal pnNroCalenNew As Integer, ByVal pnNroCalenAnt As Integer, Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pnTipoReprogCovid As Integer = 0, Optional ByVal pnCantCuotasNew As Integer = 0) 'Add joep20200526 MAntener cuota Covid
    Dim lsSQL As String
    'JOEP20200327 Covid-19 comento
    'lsSQL = "Insert ColocCalendDet " _
     '        & "Select cCtaCod, " & pnNroCalenNew & ", nColocCalendApl, nCuota, nPrdConceptoCod, nMonto, nMontoPagado, cFlag " _
      '       & "from ColocCalendDet Where cCtaCod = '" & psCtaCod & "' And nNroCalen = " & pnNroCalenAnt
    'JOEP20200327 Covid-19 comento
    'lsSQL = "exec dInsertColocCalendDetxMigracion '" & psCtaCod & "'," & pnNroCalenAnt & "," & pnNroCalenNew & "" 'JOEP20200327 Covid-19 add
    lsSQL = "exec dInsertColocCalendDetxMigracion '" & psCtaCod & "'," & pnNroCalenAnt & "," & pnNroCalenNew & "," & pnTipoReprogCovid & "," & pnCantCuotasNew & "" 'JOEP20200327 Covid-19 add
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub


Public Function dExisteConceptoColocCalendDet(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, ByVal pnColocCalendApl As ColocCalendApl, ByVal pnNroCuota As Long, ByVal pnPrdConceptoCod As Long) As Boolean
    Dim oRs As New ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "Select COUNT(cCtaCod) nCuenta " _
             & "From ColocCalenddet Where cCtaCod = '" & psCtaCod & "' And nNroCalen = " & pnNroCalen & " And nCuota=" & pnNroCuota & " And nPrdConceptoCod=" & pnPrdConceptoCod & " And nColocCalendApl=" & pnColocCalendApl
    Set oRs = coConex.CargaRecordSet(lsSQL)
    If Not oRs.EOF Then
        dExisteConceptoColocCalendDet = IIf(oRs!nCuenta > 0, True, False)
    End If
    oRs.Close
    Set oRs = Nothing
End Function
Public Sub dExtornaCredReprogNatEspeciales2(ByVal pnMovNro As Long)
    Dim sSql As String
    sSql = "exec stp_upd_ExtornaCredReprogNatEspeciales2 " & pnMovNro
    coConex.Ejecutar sSql
End Sub
Public Function dRecuperaMovimientoAprobacionxReprog(ByVal psCtaCod As String) As Long
    Dim oRs As New ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_ERS0582013_RecuperaUltMovAprobacionxRefinanc '" & psCtaCod & "'"
    Set oRs = coConex.CargaRecordSet(lsSQL)
    If Not oRs.EOF Then
        dRecuperaMovimientoAprobacionxReprog = oRs!nMovNro
    End If
    oRs.Close
    Set oRs = Nothing
End Function
'END EJVG *******
'WIOR 20140604 **************************************************
Public Sub ActualizaMontoPolizaCalend(ByVal psCtaCod As String, ByVal pnDesde As Integer, ByVal pnMonto As Double)
Dim lsSQL As String
On Error GoTo Error
    
    lsSQL = "EXEC stp_upd_ActualizarMontoPolizaRenova '" & psCtaCod & "'," & pnDesde & "," & pnMonto
    
    coConex.Ejecutar lsSQL

    Exit Sub
Error:
    Err.Raise Err.Number, "Error en proceso (ActualizaMontoPolizaCalend)", Err.Description
End Sub
Public Sub RegistrarRenovacionPoliza(ByVal psCtaCod As String, ByVal psNumGarant As String, ByVal psNumPoliza As String, ByVal pnTipoRenovacion As Integer, ByVal psPersCodAseguradora As String, _
                        ByVal psNumCertificado As String, ByVal pnMoneda As Integer, ByVal pdTasacion As Date, ByVal pnValorEdificacion As Double, ByVal pnInmueble As Double, _
                        ByVal pnPrimaNeta As Double, ByVal pnDerechoEm As Double, ByVal pnIGV As Double, ByVal pnPrimaTotal As Double, ByVal pnPrimaTotalTC As Double, _
                        ByVal pnDesde As Integer, ByVal pnPrimaCuota As Double, ByVal pdVigencia As Date, ByVal pdVencimiento As Date, ByVal pdVigenciaNueva As Date, _
                        ByVal pdVencimientoNueva As Date, ByVal psMovRegistro As String)
Dim lsSQL As String
On Error GoTo Error
    
    lsSQL = "EXEC stp_ins_CredPolizasRenova '" & psCtaCod & "','" & psNumGarant & "','" & psNumPoliza & "'," & pnTipoRenovacion & ",'" & psPersCodAseguradora & "','" & psNumCertificado & "'," & _
            pnMoneda & ",'" & Format(pdTasacion, "yyyymmdd") & "'," & pnValorEdificacion & "," & pnInmueble & "," & pnPrimaNeta & "," & pnDerechoEm & "," & pnIGV & "," & pnPrimaTotal & "," & _
            pnPrimaTotalTC & "," & pnDesde & "," & pnPrimaCuota & ",'" & Format(pdVigencia, "yyyymmdd") & "','" & Format(pdVencimiento, "yyyymmdd") & "','" & Format(pdVigenciaNueva, "yyyymmdd") & "','" & _
            Format(pdVencimientoNueva, "yyyymmdd") & "','" & psMovRegistro & "'"
    
    coConex.Ejecutar lsSQL


    Exit Sub
Error:
    Err.Raise Err.Number, "Error en proceso (ActualizaMontoPolizaCalend)", Err.Description
End Sub
Public Sub RegistarPendientPagoAutomatico(ByVal pnMovNro As Long, ByVal pnMonto As Double)
    Dim SQL As String
    
    SQL = "EXEC stp_ins_CredPendientePagoAuto " & pnMovNro & "," & pnMonto
    coConex.Ejecutar SQL
End Sub
'WIOR FIN *******************************************************
'ALPA 20120421************************************************************
Public Sub dActualizarPagoLeasigPreCancelados(Optional pbEjecBatch As Boolean = False, Optional ByVal psCtaCod As String = "", Optional ByVal pdFecha As Date = "1990/01/01")
Dim lsSQL As String

    lsSQL = "exec stp_upd_actualizarpagos_PreCancelados '" & psCtaCod & "','" & Format(pdFecha, "YYYY/MM/DD") & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub
'*************************************************************************
'RIRO20140610 ERS017 *****************************************************
Public Sub actualizarRelacionExtornoVoucherCaptacion(ByVal pnMovNroOpe As Long)
    Dim lsSQL As String
    lsSQL = "exec stp_upd_ActualizarCapVoucherDeposito_3  " & pnMovNroOpe & ""
    coConex.ConexionActiva.Execute lsSQL
End Sub
Public Sub ActualizaMovPendientesOpe(ByVal nMovNroOpe As Long, ByVal pnMonto As Currency)
    Dim lsSQL As String
    lsSQL = "STP_UPD_ACTUALIZAPENDOPE " & nMovNroOpe & ", " & pnMonto
    coConex.ConexionActiva.Execute lsSQL
End Sub
'END RIRO ****************************************************************
'WIOR 20140825 ***********************************************************
Public Sub dInsertColocCalendDetSegDes(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, ByVal pnCuota As Integer, ByVal pnAsegurados As Integer, _
                                        ByVal pnSaldoCalc As Double, ByVal pnTasa As Double, Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String

    lsSQL = "stp_ins_ColocCalendDetSegDes '" & psCtaCod & "'," & pnNroCalen & "," & pnCuota & "," & pnAsegurados & "," & pnSaldoCalc & "," & pnTasa
        
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dDeleteColocCalendDetSegDes(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, Optional pbEjecBatch As Boolean = False, Optional pOpCovid As Integer = 0, Optional ByVal pnCantCuotasNew As Integer = 0) 'Add JOEP20201012 Tasa Especial reduccion de monto covid
'Public Sub dDeleteColocCalendDetSegDes(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, Optional pbEjecBatch As Boolean = False)'Comento JOEP20201012 Tasa Especial reduccion de monto covid
Dim lsSQL As String
On Error GoTo ErrordDeleteColocCalendDetSegDes

    'lsSQL = "EXEC stp_del_ColocCalendDetSegDes '" & psCtaCod & "'," & pnNroCalen
    lsSQL = "EXEC stp_del_ColocCalendDetSegDes '" & psCtaCod & "'," & pnNroCalen & "," & pOpCovid & "," & pnCantCuotasNew & ""
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
Exit Sub
ErrordDeleteColocCalendDetSegDes:
    Err.Raise Err.Number, "Error En Proceso dDeleteColocCalendDetSegDes", Err.Description
End Sub

Public Sub dMigrarColocCalendDetSegDes(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, Optional pbEjecBatch As Boolean = False, Optional pOpCovid As Integer = 0, Optional ByVal pnCantCuotasNew As Integer = 0) 'Add JOEP20201012 Tasa Especial reduccion de monto covid
'Public Sub dMigrarColocCalendDetSegDes(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, Optional pbEjecBatch As Boolean = False)'Comento JOEP20201012 Tasa Especial reduccion de monto covid
Dim lsSQL As String
On Error GoTo ErrordMigrarColocCalendDetSegDes

    'lsSQL = "EXEC stp_ins_MigrarDatosSegDes '" & psCtaCod & "'," & pnNroCalen
    lsSQL = "EXEC stp_ins_MigrarDatosSegDes '" & psCtaCod & "'," & pnNroCalen & "," & pOpCovid & "," & pnCantCuotasNew & ""
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    
Exit Sub
ErrordMigrarColocCalendDetSegDes:
    Err.Raise Err.Number, "Error En Proceso dMigrarColocCalendDetSegDes", Err.Description
End Sub
'WIOR FIN ****************************************************************
'ALPA 20141223************************************************************
Public Sub GrabarVerEntidades(ByVal lsTipoEntidad As String, ByVal lsCtaCod As String, ByVal lsPersCod As String, ByVal lnCodEntidad As Integer, ByVal lbAnulado As Integer, ByVal psCodEdu As String, ByVal psMovNro As String, ByVal lsDocumento As String, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    lsSQL = "exec stp_ins_EndeudamientoEntidades '" & lsTipoEntidad & "','" & lsCtaCod & "','" & lsPersCod & "','" & lnCodEntidad & "','" & lbAnulado & "','" & psMovNro & "','" & psCodEdu & "','" & lsDocumento & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub
Public Sub dDeleteEndeudamientoEntidades(ByVal psCtaCod As String, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String

    lsSQL = "DELETE FROM EndeudamientoEntidades WHERE cCtaCod='" & psCtaCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
'**************************************************************************
'EJVG20141230 ***
Public Function RecuperaRelacPers(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrorRecuperaRelacPers
    lsSQL = "EXEC stp_sel_RecuperaRelacPers '" & psCtaCod & "'"
    Set RecuperaRelacPers = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrorRecuperaRelacPers:
    Err.Raise Err.Number, "Error En Proceso de Recuperar Relaciones de Personas con el Credito", Err.Description
End Function
'END EJVG *******

'JUEZ 20150218 ******************************************************************
Public Sub InsertaMovCapServDebitoAuto(ByVal pnMovNro As Long, ByVal psCtaCod As String, ByVal pnMonto As Double, ByVal pnComision As Double, _
                                        ByVal pnTipoDebito As CaptacServDebitoAutomaticoTipo, Optional ByVal psCtaCodCred As String = "", _
                                        Optional psCodConvenio As String = "", Optional psCodCliente As String = "", Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String

    lsSQL = "stp_ins_MovCapServDebitoAuto " & pnMovNro & ",'" & psCtaCod & "','" & psCtaCodCred & "','" & psCodConvenio & "','" _
                                            & psCodCliente & "'," & pnMonto & "," & pnComision & "," & pnTipoDebito
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.ConexionActiva.Execute lsSQL
    End If
End Sub
Public Function CapCargoCuentaCTS(ByRef pMatAho As Variant, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, ByVal pdFecSis As Date, _
            Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional bOPExiste As Boolean = False, _
            Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
            Optional sCodCMAC As String = "", Optional sNomCmac As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "", Optional bCancelaciones As Boolean = False, _
            Optional pnMonedaCred As Moneda = 0, _
            Optional pbITFAplica As Boolean = True, Optional pnITFValor As Double = 0, Optional pbITFAsumido As Boolean = False, _
            Optional psItfOperacion As String = "", Optional pbRetiroEfectivo As Boolean = False) As Double

Dim rsCta As ADODB.Recordset
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
Dim lnMontoOri As Double
Dim oGen As COMDConstSistema.DCOMGeneral
Dim lnTipCambioV As Currency
Dim lnTipCambioC As Currency
Dim lnTipCambio As Currency

Dim clsMov As COMNContabilidad.NCOMContFunciones

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False

Set rsCta = GetDatosCuentaCTS(sCuenta)


nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
'bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

Set oGen = New COMDConstSistema.DCOMGeneral

'GetTipCambio pdFecSis
lnTipCambioC = oGen.EmiteTipoCambio(pdFecSis, TCCompra)
lnTipCambioV = oGen.EmiteTipoCambio(pdFecSis, TCVenta)
lnTipCambio = oGen.EmiteTipoCambio(pdFecSis, TCFijoDia)

If pnMonedaCred = 0 Then pnMonedaCred = Mid(sCuenta, 9, 1)

If pnMonedaCred <> Mid(sCuenta, 9, 1) Then
    lnMontoOri = nMonto
    If pnMonedaCred = gMonedaExtranjera And Mid(sCuenta, 9, 1) = gMonedaNacional Then
        nMonto = Round(nMonto * lnTipCambioV, 2)
    Else
        nMonto = Round(nMonto / lnTipCambioC, 2)
    End If
Else
    lnMontoOri = 0
End If

bTrans = True
Randomize
For i = 0 To Rnd(2000) * 1000
Next i

If Not ValidaSaldoCuenta(sCuenta, nMonto) Then 'Valida el saldo de la cuenta nuevamente
    Err.Raise 1000, "CapCargoCuentaCTS", "Cuenta NO Posee Saldo Suficiente"
    CapCargoCuentaCTS = 0
    Exit Function
End If
'Calcula los intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntCompuesto)
If Not bCancelaciones Then
    pMatAho(3) = Format(nIntGanado, "#0.00")
Else
    pMatAho(4) = Format(nIntGanado, "#0.00")
End If

dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt - nMonto

nSaldoDisp = nSaldoDisp - nMonto
If Not bCancelaciones Then
    pMatAho(10) = nSaldoDisp
    pMatAho(11) = nSaldoCnt
Else
    pMatAho(12) = nSaldoDisp
    pMatAho(13) = nSaldoCnt
End If

If pbRetiroEfectivo Then
    pMatAho(15) = nSaldoDisp
    pMatAho(16) = nSaldoCnt
    pMatAho(17) = Format(nIntGanado, "#0.00")
End If

ActualizaSaldoRetiroCTS sCuenta, nSaldoDisp, 0
ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True

If sNroDoc <> "" Then
    If nTipoDoc = TpoDocOrdenPago Then 'Si es con orden de pago
        If bOPExiste Then
            AgregaOrdenPagoEstado sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstCobrada
        Else
            AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMonto, gCapOPEstCobrada
        End If
        If nOperacion = gAhoRetOPCanje Then
            sMsgOpe = "Retiro OP Canje"
        ElseIf nOperacion = gAhoRetOP Then
            sMsgOpe = "Retiro OP"
        ElseIf nOperacion = gAhoOPCertificacion Then
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
AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
'si el pago de credito es en diferente moneda a la cuenta de CTS
Dim lnMontoCV As Double
If pnMonedaCred <> Mid(sCuenta, 9, 1) Then
    If Mid(sCuenta, 9, 1) = gMonedaExtranjera Then
        lnMontoCV = Round((lnMontoOri / lnTipCambioC) * lnTipCambio, 2)
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeCompra, lnMontoCV
        AgregaMovTipoCambio nMovNro, lnTipCambioC
        
        If lnTipCambioC < lnTipCambio Then
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - nMonto)
        Else
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - nMonto)
        End If
    Else
        lnMontoCV = Round(lnMontoOri * lnTipCambio, 2)
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeVenta, lnMontoCV
        AgregaMovTipoCambio nMovNro, lnTipCambioV
        
        If lnTipCambioV > lnTipCambio Then
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - Round(lnMontoOri * lnTipCambioV, 2))
        Else
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - Round(lnMontoOri * lnTipCambioV, 2))
        End If
    End If
End If
CapCargoCuentaCTS = nSaldoCnt

UltimaActualizacionCuenta sCuenta, sMovNro

bTrans = False
End Function
Public Sub ActualizaSaldoRetiroCTS(ByVal sCuenta As String, ByVal nSaldoRet As Double, ByVal nIntSaldo As Double)
Dim sSql As String
    sSql = "exec stp_upd_ActualizaSaldoRetiroCTS '" & sCuenta & "'," & nSaldoRet & "," & nIntSaldo
    coConex.Ejecutar sSql
End Sub
 'END JUEZ ***********************************************************************
'WIOR 20150404 *************************
Public Sub dExtornoPerdonMora(ByVal pnMovNro As Double, Optional pbEjecBatch As Boolean = False)
Dim sSql As String

    sSql = "EXEC stp_upd_ExtornoPerdonMoraCampRecup " & pnMovNro
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If
End Sub
'WIOR FIN  *****************************
'WIOR 20150602 *************************
Public Sub dCampRecupExtornoPagoConPerdonMora(ByVal pnMovNro As Double, Optional pbEjecBatch As Boolean = False)
Dim sSql As String

    sSql = "EXEC stp_upd_CampanasRecupExtornoPagoConPerdon " & pnMovNro
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch sSql
    Else
        coConex.Ejecutar sSql
    End If
End Sub
'WIOR FIN  *****************************
'WIOR 20150421 ***
Public Function MatrizCapitalPagado(ByVal MatCalendDistrib As Variant) As Double
Dim i As Integer
    MatrizCapitalPagado = 0
    For i = 0 To UBound(MatCalendDistrib) - 1
        MatrizCapitalPagado = MatrizCapitalPagado + CDbl(MatCalendDistrib(i, 3))
    Next i

End Function
'WIOR FIN ********
'FRHU 20150725 INCIDENTE: Este metodo tambien esta en DCOMCredito, por tema de que todo este en la misma transaccion se paso para DCOMCredActBD
Public Function ExisteReprogramacionxDesastreNat(ByVal psCtaCod As String, ByVal psOpeCod As String, Optional ByVal nReversion As Integer = 0, Optional pbEjecBatch As Boolean = False) As ADODB.Recordset
    Dim sSql As String
    sSql = "Exec stp_sel_ExisteReprogramacionxDesastreNat '" & psCtaCod & "','" & psOpeCod & "'," & nReversion

    If pbEjecBatch Then
        Set ExisteReprogramacionxDesastreNat = coConex.AdicionaCmdBatch(sSql)
    Else
        Set ExisteReprogramacionxDesastreNat = coConex.Ejecutar(sSql)
    End If
End Function
'FIN FRHU
'ALPA 20150804**************************************
Public Sub dUpdateColocacCred_InteresDiferido(ByVal psCtaCod As String, ByVal lnInteresDiferido As Currency, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    On Error GoTo ErrordUpdateColocacCred_InteresDiferido
  
    lsSQL = "UPDATE ColocacCred SET "
    lsSQL = lsSQL & " nInteresDiferido = " & lnInteresDiferido & ""
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psCtaCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub

ErrordUpdateColocacCred_InteresDiferido:
    Err.Raise Err.Number, "Error En Proceso dUpdateColocacCred_InteresDiferido", Err.Description

End Sub
'***************************************************
'ALPA 20150804**************************************
Public Function dObtenerDiferidoCreditoRefinanciadoDeVigenteCancelado(ByVal psCtaCod As String, ByVal pdFecha As Date, ByVal pnMoneda As Integer, Optional pbEjecBatch As Boolean = False) As Currency
    Dim sSql As String
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    sSql = "Exec stp_sel_obtenerDiferidoCreditoRefinanciadoDeVigenteCancelado '" & psCtaCod & "','" & Format(pdFecha, "YYYY/MM/DD") & "'," & pnMoneda
    If pbEjecBatch Then
        Set oRs = coConex.AdicionaCmdBatch(sSql)
    Else
        Set oRs = coConex.Ejecutar(sSql)
    End If
    dObtenerDiferidoCreditoRefinanciadoDeVigenteCancelado = 0
    If Not (oRs.BOF Or oRs.EOF) Then
        dObtenerDiferidoCreditoRefinanciadoDeVigenteCancelado = oRs!nMontoDiferidoCredCancelado
    End If
End Function
Public Function dObtenerPorcentajeDiferidos(ByVal psCtaCod As String, _
    Optional pbEjecBatch As Boolean = False) As Currency
    
Dim lsSQL As String
Dim lrs As ADODB.Recordset
Set lrs = New ADODB.Recordset
On Error GoTo ObtenerPorcentajeDiferidosErr
    dObtenerPorcentajeDiferidos = 0
    lsSQL = "exec stp_sel_ObtenerPorcentajeDiferidos '" & psCtaCod & "'"
    If pbEjecBatch Then
        Set lrs = coConex.AdicionaCmdBatch(lsSQL)
    Else
        Set lrs = coConex.Ejecutar(lsSQL)
    End If
    If Not (lrs.EOF Or lrs.BOF) Then
        dObtenerPorcentajeDiferidos = lrs!nMontoPorcentaje
    End If
    Exit Function
ObtenerPorcentajeDiferidosErr:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'***************************************************

'WIOR 20150914 ***
Public Sub RegularPreparacionParaAmpliado(ByVal psCtaCod As String, ByVal pdFecha As Date, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String

    lsSQL = "EXEC stp_upd_ColocacCredPreAmpliadosCred '" & psCtaCod & "','" & Format(pdFecha, "yyyymmdd") & "'"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
'WIOR FIN ********
'EJVG20150513 ***
Public Sub RegistrarColocGravamen(ByVal psCtaCod As String, ByVal psTpoProdCod As String, ByVal pnMontoColocado As Currency, Optional ByVal pnCliPreferencial As Integer = -1)
    Dim lsSQL As String
    On Error GoTo ErrRegistrarColocGravamen
    'lsSQL = "EXEC stp_ins_ERS0632014_ColocGravamen '" & psCtaCod & "','" & psTpoProdCod & "'," & pnMontoColocado
    lsSQL = "EXEC stp_ins_ERS0632014_ColocGravamen '" & psCtaCod & "','" & psTpoProdCod & "'," & pnMontoColocado & "," & IIf(pnCliPreferencial = -1, "NULL", IIf(pnCliPreferencial = 0, 0, 1)) 'EJVG20160426

    coConex.Ejecutar lsSQL
    Exit Sub
ErrRegistrarColocGravamen:
    Err.Raise Err.Number, "Error al registrar datos del gravamen para el producto con código " & psCtaCod, Err.Description
End Sub
Public Sub RegistrarCoberturaGarantia(ByVal psNumGarant As String, ByVal pnMoneda As Integer, ByVal psCtaCod As String, ByVal pdValorizacion As Date, _
            ByVal pnSaldoCobertura As Currency, ByVal pnRatio As Double, ByVal pnCoberturaMax As Currency, ByVal pnCobertura As Currency, ByVal pnOrden As Integer, _
            Optional ByVal pdTramiteLegal As Date = CDate("1900-01-01"))
    Dim lsSQL As String
    On Error GoTo ErrRegistrarCoberturaGarantia
    lsSQL = "EXEC stp_ins_ERS0632014_ColocGarantia '" & psNumGarant & "'," & pnMoneda & ",'" & psCtaCod & "','" & Format(pdValorizacion, "yyyymmdd hh:mm:ss") & "'," _
            & IIf(DateDiff("d", pdTramiteLegal, "1900-01-01") = 0, "NULL", "'" & Format(pdTramiteLegal, "yyyymmdd hh:mm:ss") & "'") & "," & pnSaldoCobertura & "," _
            & pnRatio & "," & pnCoberturaMax & "," & pnCobertura & "," & pnOrden
    coConex.Ejecutar lsSQL
    Exit Sub
ErrRegistrarCoberturaGarantia:
    Err.Raise Err.Number, "Error al registrar la cobertura para el producto con código " & psCtaCod, Err.Description
End Sub
Public Sub EliminarCoberturaGarantia(ByVal psCtaCod As String)
    Dim lsSQL As String
    On Error GoTo ErrEliminarCoberturaGarantia
    lsSQL = "EXEC stp_del_ERS0632014_ColocGarantia '" & psCtaCod & "'"
    coConex.Ejecutar (lsSQL)
    Exit Sub
ErrEliminarCoberturaGarantia:
    Err.Raise Err.Number, "Error al eliminar las coberturas relaciondas al producto con código " & psCtaCod, Err.Description
End Sub
Public Sub InsertarSolicitudPoliza(ByVal psCtaCod As String, ByVal pdSolicitud As Date)
    Dim lsSQL As String
    On Error GoTo ErrInsertarSolicitudPoliza
    lsSQL = "EXEC stp_ins_ERS0632014_SolicitudPoliza '" & psCtaCod & "','" & Format(pdSolicitud, "yyyymmdd hh:mm:ss") & "'"
    coConex.Ejecutar (lsSQL)
    Exit Sub
ErrInsertarSolicitudPoliza:
    Err.Raise Err.Number, "Error al insertar solicitud de Poliza para el crédito con código " & psCtaCod, Err.Description
End Sub
Public Sub CancelarSolicitudPoliza(ByVal psCtaCod As String, ByVal psMovNroCancela As String, ByVal psGlosa As String, ByVal pbCancelarTodo As Boolean)
    Dim lsSQL As String
    On Error GoTo ErrEliminarSolicitudPoliza
    lsSQL = "EXEC stp_upd_ERS0632014_CancelarSolicitudPoliza '" & psCtaCod & "','" & psMovNroCancela & "','" & psGlosa & "'," & IIf(pbCancelarTodo, 1, 0)
    coConex.Ejecutar lsSQL
    Exit Sub
ErrEliminarSolicitudPoliza:
    Err.Raise Err.Number, "Error al eliminar la solicitud de Poliza", Err.Description
End Sub
'CTI3 ERS0032020************************************************************************************
Public Sub RegistrarCuotasRSE_Admision(ByVal psCtaCod As String)
    Dim lsSQL As String
    On Error GoTo ErrRegistrarCuotasSSF
    lsSQL = "EXEC stp_ins_RSE_Admision '" & psCtaCod & "'"
    coConex.Ejecutar lsSQL
    Exit Sub
ErrRegistrarCuotasSSF:
    Err.Raise Err.Number, "Error al registrar cuotas del Sistema Financiero", Err.Description
End Sub
'****************************************************************************************************
Public Function RequierePoliza(ByVal psCtaCod As String, ByVal psTpoCredCod As String, ByVal pnMonto As Currency, ByVal pdFecVencimiento As Date, Optional ByVal psNumGarant As String = "") As Boolean
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRequierePoliza
    lsSQL = "SELECT bRequierePoliza = dbo.fnc_ERS0632014_RequierePoliza( '" & psCtaCod & "','" & psTpoCredCod & "'," & pnMonto & ",'" & Format(pdFecVencimiento, "yyyymmdd") & "','" & psNumGarant & "')"
    Set rs = coConex.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        RequierePoliza = rs!bRequierePoliza
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrRequierePoliza:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
Public Function InsertarPoliza(ByVal psNumPoliza As String, ByVal psPersCodContr As String, ByVal psPersCodAseg As String, ByVal pnMoneda As Moneda, ByVal psNumCertif As String, _
                                ByVal pnTipoPoliza As Integer, ByVal pdRegistro As Date, ByVal pbPolizaExterna As Boolean, ByVal pdVigencia As Date, ByVal pdVencimiento As Date, _
                                ByVal pnSumaAsegurada As Currency, ByVal pnTasacionTC As Currency, ByVal pnPrimaNeta As Currency, ByVal pnDerechoEmision As Currency, ByVal pnIGV As Currency, pnPrimaTotal As Currency, ByVal pnPrimaCuota As Currency, _
                                ByVal pnNroCuotaIni As Integer, ByVal pnNroCuotaFin As Integer, ByVal pnNumAnio As Integer, ByVal pnTipoPago As Integer, ByVal pnTpoProceso As Integer)
    Dim lsSQL As String
    On Error GoTo ErrInsertarPoliza
    lsSQL = "EXEC stp_ins_ERS0632014_Poliza '" & psNumPoliza & "','" & psPersCodContr & "','" & psPersCodAseg & "'," & pnMoneda & ",'" & psNumCertif & "'," & _
                                pnTipoPoliza & ",'" & Format(pdRegistro, "yyyymmdd") & "'," & IIf(pbPolizaExterna, 1, 0) & ",'" & Format(pdVigencia, "yyyymmdd") & "','" & Format(pdVencimiento, "yyyymmdd") & "'," & _
                                pnSumaAsegurada & "," & pnTasacionTC & "," & pnPrimaNeta & "," & pnDerechoEmision & "," & pnIGV & "," & pnPrimaTotal & "," & pnPrimaCuota & "," & _
                                pnNroCuotaIni & "," & pnNroCuotaFin & "," & pnNumAnio & "," & pnTipoPago & "," & pnTpoProceso
    coConex.Ejecutar lsSQL
    Exit Function
ErrInsertarPoliza:
    Err.Raise Err.Number, "Error al Insertar una Poliza", Err.Description
End Function
Public Function RecuperaNumeroPoliza() As String
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrRecuperaNumeroPoliza
    
    Set rs = coConex.CargaRecordSet("SELECT cNumPoliza=ISNULL(MAX(convert(int,cNumPoliza)),0) FROM Poliza")
    
    If rs.EOF Then
        RecuperaNumeroPoliza = "00000001"
    Else
        RecuperaNumeroPoliza = Format(rs!cNumPoliza + 1, "00000000")
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrRecuperaNumeroPoliza:
    Err.Raise Err.Number, "Error al recuperar correlativo número de Poliza", Err.Description
End Function
'Public Function InsertarPolizaGarantia(ByVal psNumPoliza As String, ByVal psNumGarant As String, ByVal pdTasacion As Date, ByVal pdFecha As Date)
'    Dim lsSql As String
'    On Error GoTo ErrInsertarPolizaGarantia
'    lsSql = "EXEC stp_ins_ERS0632014_PolizaGarantia '" & psNumPoliza & "','" & psNumGarant & "','" & Format(pdTasacion, "yyyymmdd hh:mm:ss") & "','" & Format(pdFecha, "yyyymmdd") & "'"
'    coConex.Ejecutar lsSql
'    Exit Function
'ErrInsertarPolizaGarantia:
'    Err.Raise Err.Number, "Error al Insertar relación Poliza-Garantía", Err.Description
'End Function
'Public Function InsertarPolizaCredito(ByVal psNumPoliza As String, ByVal psCtaCod As String, ByVal pdFecha As Date)
'    Dim lsSql As String
'    On Error GoTo ErrInsertarPolizaCredito
'    lsSql = "EXEC stp_ins_ERS0632014_PolizaCredito '" & psNumPoliza & "','" & psCtaCod & "','" & Format(pdFecha, "yyyymmdd") & "'"
'    coConex.Ejecutar lsSql
'    Exit Function
'ErrInsertarPolizaCredito:
'    Err.Raise Err.Number, "Error al Insertar relación Poliza-Crédito", Err.Description
'End Function
'Public Function InsertarCreditoGarantiaPoliza(ByVal psCtaCod As String, ByVal psNumGarant As String, ByVal pdValorizacion As Date, ByVal psNumPoliza As String, ByVal pnMonedaGarantia As Integer, ByVal pnValorEdifComGarantia As Currency)
'    Dim lsSql As String
'    On Error GoTo ErrInsertarCreditoGarantiaPoliza
'    lsSql = "EXEC stp_ins_ERS0632014_ColocGarantiaPoliza '" & psCtaCod & "','" & psNumGarant & "','" & Format(pdValorizacion, "yyyymmdd hh:mm:ss") & "','" & psNumPoliza & "'," & pnMonedaGarantia & "," & pnValorEdifComGarantia
'    coConex.Ejecutar lsSql
'    Exit Function
'ErrInsertarCreditoGarantiaPoliza:
'    Err.Raise Err.Number, "Error al Insertar relación Crédito-Garantía-Póliza", Err.Description
'End Function
Public Function InsertarCreditoGarantiaPoliza(ByVal psCtaCod As String, ByVal psNumGarant As String, ByVal pdValorizacion As Date, ByVal psNumPoliza As String, ByVal pdRegistro As Date, ByVal pnMonedaGarantia As Integer, ByVal pnValorEdifComGarantia As Currency)
    Dim lsSQL As String
    On Error GoTo ErrInsertarCreditoGarantiaPoliza
    lsSQL = "EXEC stp_ins_ERS0632014_ColocGarantiaPoliza '" & psCtaCod & "','" & psNumGarant & "','" & Format(pdValorizacion, "yyyymmdd hh:mm:ss") & "','" & psNumPoliza & "','" & Format(pdRegistro, "yyyymmdd") & "'," & pnMonedaGarantia & "," & pnValorEdifComGarantia
    coConex.Ejecutar lsSQL
    Exit Function
ErrInsertarCreditoGarantiaPoliza:
    Err.Raise Err.Number, "Error al Insertar relación Crédito-Garantía-Póliza", Err.Description
End Function
Public Function EliminarCreditoGarantiaPoliza(ByVal psCtaCod As String, ByVal psNumGarant As String, ByVal psNumPoliza As String) As Boolean
    Dim lsSQL As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrInsertarCreditoGarantiaPoliza
    lsSQL = "EXEC stp_del_ERS0632014_ColocGarantiaPoliza '" & psCtaCod & "','" & psNumGarant & "','" & psNumPoliza & "'"
    Set rs = coConex.CargaRecordSet(lsSQL)
    
    If Not rs.EOF Then
        EliminarCreditoGarantiaPoliza = IIf(rs!nNroRegistroAfecta > 0, True, False)
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrInsertarCreditoGarantiaPoliza:
    Err.Raise Err.Number, "Error al Eliminar lógicamente relación Crédito-Garantía-Póliza", Err.Description
End Function
Public Function InsertarGastosxCuotaLote(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
                                    ByVal pnColocCalendApl As ColocCalendApl, ByVal pnColocConceptoCod As ColocConcepto, ByVal pnCuotaIni As Integer, ByVal pnCuotaFin As Integer, _
                                    ByVal pnMonto As Currency, ByVal pbIncrementa As Boolean, Optional ByVal pbBackup As Boolean = False, Optional ByVal psNumPoliza As String = "") As Boolean
    Dim nCuota As Integer
    On Error GoTo ErrInsertarGastosxCuotaLote
    
    If pbBackup Then
        Call dInsertColocCalendDetBackup(psCtaCod, psNumPoliza, pnNroCalen, pnColocConceptoCod, pnCuotaIni, pnCuotaFin)
    End If
    
    For nCuota = pnCuotaIni To pnCuotaFin
        If ExisteGasto(psCtaCod, pnNroCalen, pnColocCalendApl, nCuota, pnColocConceptoCod) Then
            pbIncrementa = True 'APRI 20170227 'FRHU20170323 Descomentado ok
            Call dUpdateColocCalendDet(psCtaCod, pnNroCalen, pnColocCalendApl, nCuota, pnColocConceptoCod, pnMonto, 0#, "", False, , IIf(pbIncrementa, True, False))
        Else
            Call dInsertColocCalendDet(psCtaCod, pnNroCalen, pnColocCalendApl, nCuota, pnColocConceptoCod, pnMonto, 0#, "", False)
        End If
    Next

    InsertarGastosxCuotaLote = True
    Exit Function
ErrInsertarGastosxCuotaLote:
    InsertarGastosxCuotaLote = False
    Err.Raise Err.Number, "Error al insertar gasto", Err.Description
End Function
Public Function ExisteGasto(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, ByVal pnColocCalendApl As ColocCalendApl, ByVal pnCuota As Integer, ByVal pnConcepto As Long) As Boolean
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    
    sSql = "Select nMonto,nMontoPagado From ColocCalendDet WHERE cCtaCod = '" & psCtaCod & "' "
    sSql = sSql & " AND nNroCalen = " & pnNroCalen
    sSql = sSql & " AND nColocCalendApl = " & pnColocCalendApl
    sSql = sSql & " AND nCuota = " & pnCuota
    sSql = sSql & " AND nPrdConceptoCod = " & pnConcepto
    
    Set rs = coConex.CargaRecordSet(sSql)
    If rs.RecordCount > 0 Then
        ExisteGasto = True
    Else
        ExisteGasto = False
    End If
    rs.Close
    Set rs = Nothing
End Function
'Public Sub InsertarRenovacionPoliza(ByVal psCtaCod As String, ByVal psNumGarant As String, ByVal psNumPoliza As String, ByVal pnTipoRenovacion As Integer, ByVal psPersCodAseguradora As String, _
'                        ByVal psNumCertificado As String, ByVal pnMoneda As Integer, ByVal pdTasacion As Date, ByVal pnValorEdificacion As Double, ByVal pnInmueble As Integer, _
'                        ByVal pnPrimaNeta As Currency, ByVal pnDerechoEm As Currency, ByVal pnIGV As Currency, ByVal pnPrimaTotal As Currency, ByVal pnPrimaTotalTC As Currency, _
'                        ByVal pnDesde As Integer, ByVal pnPrimaCuota As Currency, ByVal pdVigencia As Date, ByVal pdVencimiento As Date, ByVal psMovRegistro As String, ByVal pnInmueble As Integer)
'    Dim sSql As String
'On Error GoTo ErrInsertarRenovacionPoliza
'    sSql = "EXEC stp_ins_ERS0632014_RenovarPoliza '" & psCtaCod & "','" & psNumGarant & "','" & psNumPoliza & "'," & pnTipoRenovacion & ",'" & psPersCodAseguradora & "','" & psNumCertificado & "'," & _
'            pnMoneda & ",'" & Format(pdTasacion, "yyyymmdd") & "'," & pnValorEdificacion & "," & pnInmueble & "," & pnPrimaNeta & "," & pnDerechoEm & "," & pnIGV & "," & pnPrimaTotal & "," & _
'            pnPrimaTotalTC & "," & pnDesde & "," & pnPrimaCuota & ",'" & Format(pdVigencia, "yyyymmdd") & "','" & Format(pdVencimiento, "yyyymmdd") & "','" & psMovRegistro & "'," & pnInmueble
'    coConex.Ejecutar sSql
'    Exit Sub
'ErrInsertarRenovacionPoliza:
'    Err.Raise Err.Number, "Error en proceso (InsertaRenovacionPoliza)", Err.Description
'End Sub
Public Function ModificarPoliza(ByVal pcNumPoliza As String, _
                                Optional ByVal pcPersCodContr As String = "", _
                                Optional ByVal pcPersCodAseg As String = "", _
                                Optional ByVal pcNumCertif As String = "", _
                                Optional ByVal pnMoneda As Integer = -1, _
                                Optional ByVal pnSumaAsegurada As Currency = -1, _
                                Optional ByVal pnTasacionTC As Currency = -1, _
                                Optional ByVal pnPrimaNeta As Currency = -1, _
                                Optional ByVal pnDerechoEmi As Currency = -1, _
                                Optional ByVal pnIGV As Currency = -1, _
                                Optional ByVal pnPrimaTotal As Currency = -1, _
                                Optional ByVal pnTipoPoliza As Integer = -1, _
                                Optional ByVal pnPrimaCuota As Currency = -1, _
                                Optional ByVal pnNroCuotaDesde As Integer = -1, _
                                Optional ByVal pnNroCuotaHasta As Integer = -1, _
                                Optional ByVal pdVigencia As Date = CDate("1900-01-01"), _
                                Optional ByVal pdVencimiento As Date = CDate("1900-01-01"))
    Dim sSql As String
    
    On Error GoTo ErrModificarPoliza
    sSql = " UPDATE Poliza SET "
    If pcPersCodContr <> "" Then
        sSql = sSql & " cPersCodContr='" & pcPersCodContr & "',"
    End If
    If pcPersCodAseg <> "" Then
        sSql = sSql & " cPersCodAseg='" & pcPersCodAseg & "',"
    End If
    If pcNumCertif <> "" Then
        sSql = sSql & " cCodPolizaAseg='" & pcNumCertif & "',"
    End If
    If pnMoneda <> -1 Then
        sSql = sSql & " nMoneda=" & pnMoneda & ","
    End If
    If pnTasacionTC <> -1 Then
        sSql = sSql & " nTasacionTC=" & pnTasacionTC & ","
    End If
    If pnSumaAsegurada <> -1 Then
        sSql = sSql & " nSumaAsegurada=" & pnSumaAsegurada & ","
    End If
    If pnPrimaNeta <> -1 Then
        sSql = sSql & " nPrimaNeta=" & pnPrimaNeta & ","
    End If
    If pnDerechoEmi <> -1 Then
        sSql = sSql & " nDrchoEmi=" & pnDerechoEmi & ","
    End If
    If pnIGV <> -1 Then
        sSql = sSql & " nIGV=" & pnIGV & ","
    End If
    If pnPrimaTotal <> -1 Then
        sSql = sSql & " nMontoPrimaTotal=" & pnPrimaTotal & ","
    End If
    If pnTipoPoliza <> -1 Then
        sSql = sSql & " nTipoPoliza=" & pnTipoPoliza & ","
    End If
    If pnPrimaCuota <> -1 Then
        sSql = sSql & " nTipoPoliza=" & pnPrimaCuota & ","
    End If
    If pnNroCuotaDesde <> -1 Then
        sSql = sSql & " nNroCuotaIni=" & pnNroCuotaDesde & ","
    End If
    If pnNroCuotaHasta <> -1 Then
        sSql = sSql & " nNroCuotaFin=" & pnNroCuotaHasta & ","
    End If
    If DateDiff("D", pdVigencia, CDate("1900-01-01")) <> 0 Then
        sSql = sSql & " dVigenciaAseg='" & Format(pdVigencia, "yyyymmdd") & "',"
    End If
    If DateDiff("D", pdVencimiento, CDate("1900-01-01")) <> 0 Then
        sSql = sSql & " dVencAseg='" & Format(pdVencimiento, "yyyymmdd") & "',"
    End If
    
    sSql = Left(sSql, Len(sSql) - 1)
    sSql = sSql & " WHERE cNumPoliza='" & pcNumPoliza & "'"
    coConex.Ejecutar sSql
    Exit Function
ErrModificarPoliza:
    Err.Raise Err.Number, "Error al actualizar la Poliza (ModificarPoliza)", Err.Description
End Function
Public Sub dInsertColocCalendDetBackup(ByVal psCtaCod As String, ByVal psNumPoliza As String, ByVal pnNroCalen As Integer, ByVal pnConcepto As Integer, ByVal pnCuotaIni As Integer, ByVal pnCuotaFin As Integer)
    Dim sSql As String
On Error GoTo ErrdInsertColocCalendDetBackup
    sSql = "EXEC stp_ins_ERS0632014_ColocCalendDetBackup '" & psCtaCod & "','" & psNumPoliza & "'," & pnNroCalen & "," & pnConcepto & "," & pnCuotaIni & "," & pnCuotaFin
    coConex.Ejecutar sSql
    Exit Sub
ErrdInsertColocCalendDetBackup:
    Err.Raise Err.Number, "Error en proceso (dInsertColocCalendDetBackup)", Err.Description
End Sub
Public Sub ConsolidarGarantia(Optional ByVal psNumGarant As String = "")
    Dim lsSQL As String
    On Error GoTo ErrConsolidaGarantia
    lsSQL = "EXEC stp_upd_ERS0632014_ConsolidaGarantia '" & psNumGarant & "'"
    coConex.Ejecutar lsSQL
    Exit Sub
ErrConsolidaGarantia:
    Err.Raise Err.Number, "Error al consolidar datos de la Garantía", Err.Description
End Sub
Public Sub dAsientoGarantiaDesembolso(ByVal pnMovNro As Long, ByVal psCtaCod As String, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    On Error GoTo ErrAsientoGarantiaDesembolso
    lsSQL = "EXEC stp_ins_ERS0632014_AsientoGarantiaDesembolso " & pnMovNro & ",'" & psCtaCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub
ErrAsientoGarantiaDesembolso:
    Err.Raise Err.Number, "Error al registrar las operaciones de asiento de garantía", Err.Description
End Sub
Public Sub dLiberarGarantia(pnMovNro As Long, ByVal psOpeTpo As String, ByVal psCtaCod As String, ByVal pnMontoCapAmortiza As Currency, ByVal pnEstadoCred As ColocEstado, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    On Error GoTo ErrLiberarGarantia
    lsSQL = "EXEC spt_ins_ERS0632014_LiberaGarantia " & pnMovNro & ",'" & psOpeTpo & "','" & psCtaCod & "'," & pnMontoCapAmortiza & "," & pnEstadoCred
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub
ErrLiberarGarantia:
    Err.Raise Err.Number, "Error al liberar las garantías en el proceso de Pago de Créditos", Err.Description
End Sub
Public Function dInsertSolicitudExoneraCobertura(ByVal psCtaCod As String, ByVal psComentario As String, ByVal psMovNro As String, ByVal pnRatio As Double) As Long
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
    
    On Error GoTo ErrInsertSolicitudExoneracionTasa
    
    lsSQL = "EXEC stp_ins_ERS0632014_ExoneraCobertura '" & psCtaCod & "','" & psComentario & "','" & psMovNro & "'," & pnRatio
    Set rs = coConex.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        dInsertSolicitudExoneraCobertura = rs!nId
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrInsertSolicitudExoneracionTasa:
    Err.Raise Err.Number, "Error al Solicitar Exoneración de Tasas", Err.Description
End Function
Public Sub dUpdateSolicitudExoneraCobertura(ByVal pnId As Long, ByVal pbAutoriza As Boolean, ByVal psMovNro As String, ByVal pnTasa As Currency, ByVal psComentario)
    Dim lsSQL As String
    
    On Error GoTo ErrInsertSolicitudExoneracionTasa
    
    lsSQL = "EXEC stp_upd_ERS0632014_ExoneraCobertura " & pnId & "," & IIf(pbAutoriza, 1, 0) & ",'" & psMovNro & "'," & pnTasa & ",'" & psComentario & "'"
    coConex.Ejecutar lsSQL
    Exit Sub
ErrInsertSolicitudExoneracionTasa:
    Err.Raise Err.Number, "Error al Solicitar Exoneración de Tasas", Err.Description
End Sub
Public Sub dUpdateCoberturaGarantia(ByVal psNumGarant As String)
    Dim lsSQL As String
    On Error GoTo ErrdUpdateGravament
    lsSQL = "EXEC stp_upd_ERS0632014_CoberturaGarantia '" & psNumGarant & "'"
    coConex.Ejecutar lsSQL
    Exit Sub
ErrdUpdateGravament:
    Err.Raise Err.Number, "Error al actualizar gravament de la Garantía", Err.Description
End Sub
Public Function RequiereReLiberaGarantiasCreditoDiferenteMoneda(ByVal pdFecha As Date) As Boolean
    Dim lsSQL As String
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrRequiere
    lsSQL = "SELECT bRealizarReLiberacion = DBO.fnc_ERS0632014_RequiereReLiberaGarantiaCreditoDiferenteMoneda('" & Format(pdFecha, "yyyymmdd") & "')"
    Set rs = coConex.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        RequiereReLiberaGarantiasCreditoDiferenteMoneda = rs!bRealizarReLiberacion
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrRequiere:
    Err.Raise Err.Number, "Error al consultar si se va a reliberar garantías", Err.Description
End Function
Public Sub ReLiberaGarantiasCreditoDiferenteMoneda(ByVal pdFecha As Date)
    Dim lsSQL As String
    On Error GoTo ErrReLiberaGarantiasCreditoDiferenteMoneda
    lsSQL = "EXEC stp_upd_ERS0632014_ReLiberaGarantiaCreditoDiferenteMoneda '" & Format(pdFecha, "yyyymmdd") & "'"
    coConex.Ejecutar lsSQL
    Exit Sub
ErrReLiberaGarantiasCreditoDiferenteMoneda:
    Err.Raise Err.Number, "Error al ReLiberar las garantías de los créditos de difernte monedas", Err.Description
End Sub
Public Sub dUpdateCoberturaGarantiaxCredito(ByVal psCtaCod As String)
    Dim lsSQL As String
    On Error GoTo ErrUpdateCoberturaGarantiaxCredito
    lsSQL = "EXEC stp_upd_ERS0632014_CoberturaGarantiaxCredito '" & psCtaCod & "'"
    coConex.Ejecutar lsSQL
    Exit Sub
ErrUpdateCoberturaGarantiaxCredito:
    Err.Raise Err.Number, "Error al actualizar las coberturas de las garantías del producto " & psCtaCod, Err.Description
End Sub
Public Sub dAsientoGarantiaRefinancia(ByVal pnMovNro As Long, ByVal psCtaCod As String, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    On Error GoTo ErrAsientoGarantiaRefinancia
    lsSQL = "EXEC stp_ins_ERS0632014_AsientoGarantiaRefinancia " & pnMovNro & ",'" & psCtaCod & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub
ErrAsientoGarantiaRefinancia:
    Err.Raise Err.Number, "Error al registrar las operaciones de asiento de garantía por refinanciación", Err.Description
End Sub
'END EJVG *******
'JUEZ 20160216 ********************************************************************
Public Sub dInsertaColocacReprogramadoEstado(ByVal psCtaCod As String, ByVal psFechaEstado As String, ByVal pnPrdEstado As ColocacReprogEstado, ByVal pnSaldoReprog As Double, _
                                            ByVal psMovNro As String, ByVal pnDiasAtraso As Integer, ByVal pnCuotasReprog As Integer, ByVal pdFecCuotaVenc As Date, _
                                            ByVal pdFecNuevaCuotaVenc As Date, ByVal psMotivo As String, ByVal pbSolicAutorizacion As Boolean, _
                                            Optional pnMenuOpcion As Integer = -1, Optional pOpCovid As Integer = -1, Optional pArrayDatos As Variant = Nothing) 'Add JOEP20210306 garantia covid
Dim lsSQL As String
'Add JOEP20210306 garantia covid
    'lsSQL = "Exec stp_ins_upd_ColocacReprogramado '" & psCtaCod & "'," & pnSaldoReprog & "," & pnPrdEstado & ",'" & psFechaEstado & "','" & psMovNro & "'," & pOpCovid & "," & pnMontoCuota & "" 'JOEP Tasa Especial Reprogramacion Propuesta
    If IsArray(pArrayDatos) = False Then
        lsSQL = "Exec stp_ins_ColocacReprogramadoEstado '" & psCtaCod & "','" & psFechaEstado & "'," & pnPrdEstado & "," & pnSaldoReprog & ",'" & psMovNro & "'," & _
                                                     pnDiasAtraso & "," & pnCuotasReprog & ",'" & Format(pdFecCuotaVenc, "yyyyMMdd") & "','" & _
                                                     Format(pdFecNuevaCuotaVenc, "yyyyMMdd") & "','" & psMotivo & "'," & IIf(pbSolicAutorizacion, 1, 0) & "," _
                                                     & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & pnMenuOpcion & ""
    Else
        lsSQL = "Exec stp_ins_ColocacReprogramadoEstado '" & psCtaCod & "','" & psFechaEstado & "'," & pnPrdEstado & "," & pnSaldoReprog & ",'" & psMovNro & "'," & _
                                                         pnDiasAtraso & "," & pnCuotasReprog & ",'" & Format(pdFecCuotaVenc, "yyyyMMdd") & "','" & _
                                                        Format(pdFecNuevaCuotaVenc, "yyyyMMdd") & "','" & psMotivo & "'," & IIf(pbSolicAutorizacion, 1, 0) & "," _
                                                        & pArrayDatos(1) & "," & pArrayDatos(2) & "," & pArrayDatos(3) & "," & pArrayDatos(4) & "," & pArrayDatos(5) & "," & pArrayDatos(6) & "," & pArrayDatos(7) & "," & pnMenuOpcion & ""
    End If
'Add JOEP20210306 garantia covid
    coConex.Ejecutar lsSQL
End Sub

Public Sub dInsertaActualizaColocacReprogramado(ByVal psCtaCod As String, ByVal pnSaldoReprog As Double, ByVal pnPrdEstado As ColocacReprogEstado, _
                                               ByVal psFechaEstado As String, ByVal psMovNro As String, Optional pOpCovid As Integer = -1, Optional pnMontoCuota As Currency = -1, _
                                               Optional pnMenuOpcion As Integer = -1) 'JOEP20210211 Garantia covid
Dim lsSQL As String

    'lsSQL = "Exec stp_ins_upd_ColocacReprogramado '" & psCtaCod & "'," & pnSaldoReprog & "," & pnPrdEstado & ",'" & psFechaEstado & "','" & psMovNro & "'"'Comento JOEP Tasa Especial Reprogramacion Propuesta
    'lsSQL = "Exec stp_ins_upd_ColocacReprogramado '" & psCtaCod & "'," & pnSaldoReprog & "," & pnPrdEstado & ",'" & psFechaEstado & "','" & psMovNro & "'," & pOpCovid & "," & pnMontoCuota & "" 'JOEP Tasa Especial Reprogramacion Propuesta
    lsSQL = "Exec stp_ins_upd_ColocacReprogramado '" & psCtaCod & "'," & pnSaldoReprog & "," & pnPrdEstado & ",'" & psFechaEstado & "','" & psMovNro & "'," & pOpCovid & "," & pnMontoCuota & "," _
        & pnMenuOpcion & "" 'Add JOEP20210306 garantia covid
    coConex.Ejecutar lsSQL
End Sub
Public Sub dEliminaColocacReprogramado(ByVal psCtaCod As String, ByVal pnPrdEstado As ColocacReprogEstado, ByVal psMovNro As String)
Dim lsSQL As String

    lsSQL = "Exec stp_del_ColocacReprogramado '" & psCtaCod & "'," & pnPrdEstado & ",'" & psMovNro & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Sub dInsertaColocacReprogExoneraSolicitud(ByVal psCtaCod As String, ByVal nExoneraCod As Integer, ByVal psFechaSolicitud As String)
Dim lsSQL As String

    lsSQL = "Exec stp_ins_ColocacReprogExoneraSolicitud '" & psCtaCod & "'," & nExoneraCod & ",'" & psFechaSolicitud & "'"
    coConex.Ejecutar lsSQL
End Sub
'END JUEZ *************************************************************************
'JUEZ 20160407 **************************************************************
Public Sub dInsertaCredMetasAnalistas(ByVal psPersCod As String, ByVal psUser As String, ByVal psAgeCod As String, ByVal pnAnio As Integer, ByVal pnMes As Integer, _
                                      ByVal pnMetaCrecCartera As Double, ByVal pnMetaCrecClientes As Integer, ByVal pnMetaNumDesemb As Integer, ByVal pnMetaPorcMoraMayor30 As Double, _
                                      ByVal pnSaldoBajarMoraMayor30 As Double, ByVal pnMetaPorcCAR As Double, ByVal pnSaldoBajarCAR As Double, ByVal psUltimaActualizacion As String)
Dim lsSQL As String

    lsSQL = "Exec stp_ins_upd_CredMetasAnalistas '" & psPersCod & "','" & psUser & "','" & psAgeCod & "'," & pnAnio & "," & pnMes & "," & pnMetaCrecCartera & "," & _
                                                      pnMetaCrecClientes & "," & pnMetaNumDesemb & "," & pnMetaPorcMoraMayor30 & "," & pnSaldoBajarMoraMayor30 & "," & _
                                                      pnMetaPorcCAR & "," & pnSaldoBajarCAR & ",'" & psUltimaActualizacion & "'"
    coConex.Ejecutar lsSQL
End Sub
'END JUEZ *******************************************************************

'WIOR 20160419 ***
Public Sub GestionSiniestrorLiquidacionReg(ByVal pnIdSiniestro As Long, ByVal pnIdRegAtencion As Long, ByVal pcCtaCod As String, ByVal pnMovNro As Long, _
                                        ByVal pcNroDoc As String, ByVal pnMontoPago As Currency, ByVal pnMontoDevolucion As Currency, ByVal pdFechaReg As Date, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String



    lsSQL = "EXEC stp_ins_GestionSiniestroLiquidacion " & pnIdSiniestro & "," & pnIdRegAtencion & ",'" & pcCtaCod & "'," & pnMovNro & ",'" & _
                                        pcNroDoc & "'," & pnMontoPago & "," & pnMontoDevolucion & ",'" & Format(pdFechaReg + coConex.GetHoraServer, "yyyymmdd hh:mm:ss") & "'"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub
Public Sub GestionSiniestrorLiquidacionEliminar(ByVal pnMovNro As Long, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    lsSQL = "EXEC stp_del_GestionSiniestroLiquidacion " & pnMovNro

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub
'WIOR FIN ********
'EJVG20160616 ***
Public Sub dAjusteResgistroGarantia(ByVal psCtaCod As String, ByVal pdFecha As Date, ByVal pnMontoCol As Double, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    On Error GoTo ErrdAjusteResgistroGarantia
    lsSQL = "EXEC stp_upd_AjusteRegistroCobertura '" & psCtaCod & "','" & Format(pdFecha, "yyyymmdd") & "'," & pnMontoCol
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub
ErrdAjusteResgistroGarantia:
    Err.Raise Err.Number, "Error en el proceso<dAjusteResgistroGarantia>", Err.Description
End Sub
'END EJVG *******


'WIOR 20160613 ***
Public Sub SobreEndEliminarSolicitudBloq(ByVal pcCtaCod As String, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    lsSQL = "EXEC stp_del_SobreEndCodigos '" & pcCtaCod & "'"

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub SobreEndRegistrarSolicitudBloq(ByVal pcCtaCod As String, ByVal prsDatos As ADODB.Recordset, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    
    lsSQL = "EXEC stp_ins_SobreEndCodigos '" & pcCtaCod & "','" & Format(CDate(prsDatos!dCierre), "yyyymmdd") & "','" & Format(CDate(prsDatos!dRCCAct), "yyyymmdd") & "','" & _
    Format(CDate(prsDatos!dRCC11), "yyyymmdd") & "'," & _
    CCur(prsDatos!nDeudaTotalAct) & "," & _
    CCur(prsDatos!nDeudaTotal11) & "," & _
    CInt(prsDatos!nNroEnt) & "," & _
    CCur(prsDatos!nPatrimonio) & ",'" & _
    (prsDatos!cCalifInterna) & "','" & _
    (prsDatos!cCalifInterna) & "'," & _
    CCur(prsDatos!nCuota) & "," & _
    CCur(prsDatos!nUtilidad) & "," & _
    CInt(prsDatos!nMaxDiasAtraso) & "," & _
    CInt(prsDatos!nCodigo1) & "," & CInt(prsDatos!nCodigo2) & "," & CInt(prsDatos!nCodigo3) & "," & _
    CInt(prsDatos!nCodigo4) & "," & CInt(prsDatos!nCodigo5) & "," & CInt(prsDatos!cCodigo1Det) & "," & CInt(prsDatos!cCodigo2Det) & "," & CInt(prsDatos!cCodigo3Det) & "," & _
    CInt(prsDatos!cCodigo4Det) & "," & CInt(prsDatos!cCodigo5Det)

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub
'WIOR FIN ********
'EJVG20160713 ***
Public Sub dUpdateMontoExposicionEsteCredito(ByVal pcCtaCod As String, Optional ByVal pnMontoExpEsteCred As Double = 0#, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    
    lsSQL = "EXEC stp_upd_MontoExposicionEsteCredito '" & pcCtaCod & "'," & pnMontoExpEsteCred

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub
'Public Sub dDeleteEvaluacionCredito(ByVal pcCtaCod As String, Optional pbEjecBatch As Boolean = False)
'    Dim lsSQL As String
'
'    lsSQL = "EXEC spt_del_CredFormEval '" & pcCtaCod & "'"
'
'    If pbEjecBatch Then
'        coConex.AdicionaCmdBatch lsSQL
'    Else
'        coConex.Ejecutar lsSQL
'    End If
'End Sub
'Public Sub dRecalculaIndicadoresyRatiosEvaluacion(ByVal pcCtaCod As String, Optional pbEjecBatch As Boolean = False)
'    Dim lsSQL As String
'
'    lsSQL = "EXEC stp_upd_CredformEvalratiosIndicadores '" & pcCtaCod & "'"
'
'    If pbEjecBatch Then
'        coConex.AdicionaCmdBatch lsSQL
'    Else
'        coConex.Ejecutar lsSQL
'    End If
'End Sub
'END EJVG *******

'JOEP 20170311 REPROGRAMACION
Public Sub dInsertaConceptosIntGracia(ByVal psCtaCod As String, ByVal psNroCalen As Integer, ByVal psnColocCalenApl As Integer, ByVal psnCuota As Integer, ByVal psnConceptoCod As Integer, ByVal psnMonto As Double, ByVal psnMontoPagado As String, ByVal pscFlag As String, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String

  On Error GoTo ErrordInsertaConceptosIntGracia
    
    lsSQL = "sp_inst_InsertarConceptoIntGracia '" & psCtaCod & "'," & psNroCalen & "," & psnColocCalenApl & "," & psnCuota & "," & psnConceptoCod & "," & psnMonto & "," & psnMontoPagado & ",'" & pscFlag & "'"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub
ErrordInsertaConceptosIntGracia:
    Err.Raise Err.Number, "Error En Proceso dInsertaConceptosIntGracia", Err.Description
End Sub

Public Sub ActualizarTCEA(ByVal psCtaCod As String, ByVal pnMontoTCEA As Double, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String

  On Error GoTo ErrorActualizarTCEA
    
    lsSQL = "sp_sel_ActualizarTCEA '" & psCtaCod & "'," & pnMontoTCEA & ""
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
    Exit Sub
ErrorActualizarTCEA:
    Err.Raise Err.Number, "Error En Proceso ActualizarTCEA", Err.Description
End Sub

'JOEP 20170311 REPROGRAMACION

'PASI20170622 ***********************************************
Public Function DevuelveOpeTpoMov(ByVal pnMovNro As Long) As String
    Dim rs As ADODB.Recordset
    Dim sSql As String
    
    sSql = "select cOpeCod from Mov where nMovNro = " & pnMovNro
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    If Not (rs.EOF And rs.BOF) Then
         DevuelveOpeTpoMov = rs!cOpeCod
    Else
         DevuelveOpeTpoMov = ""
    End If
    rs.Close
    Set rs = Nothing
End Function
'PASI END****************************************************
'VAPA20171019 DIFERIDOS FROM 60'********Comentado by NAGL 20180609, para ser transferido al cierre Diario**********
'Public Sub dInsertaInteresDiferido(ByVal psCtaCodAmp As String, ByVal pnMonto As Double, ByVal pnMovNro As Long, Optional pnMovNroDesem As Long = 0)
'    Dim lsSQL As String
'
'    lsSQL = "stp_ins_InteresDiferido '" & psCtaCodAmp & "'," & pnMonto & "," & pnMovNro & "," & pnMovNroDesem & ""
'    coConex.Ejecutar lsSQL
'    Exit Sub
'End Sub
'Public Sub dActualizaInteresDiferido(ByVal psCtaCodAmp As String, ByVal pnMonto As Double, ByVal pnMovNro As Long, Optional pnMovNroDesem As Long = 0)
'    Dim lsSQL As String
'
'    lsSQL = "stp_upd_InteresDiferido '" & psCtaCodAmp & "'," & pnMonto & "," & pnMovNro & "," & pnMovNroDesem & ""
'    coConex.Ejecutar lsSQL
'    Exit Sub
'End Sub
'Public Sub dActualizaIntDifDesembolso(ByVal psCtaCodAmp As String, ByVal pnMonto As Double, Optional pnMovNroDesem As Long = 0, Optional pbEjecBatch As Boolean = False)
'    Dim lsSQL As String
'
'    lsSQL = "stp_upd_InteresDiferidoDesembolso '" & psCtaCodAmp & "'," & pnMonto & "," & pnMovNroDesem & ""
'
'    If pbEjecBatch Then
'        coConex.AdicionaCmdBatch lsSQL
'    Else
'        coConex.Ejecutar lsSQL
'    End If
'
'    Exit Sub
'End Sub
'
'Public Function ValidaPreAmpliacion(ByVal pnMonNro As Integer) As Boolean
'Dim sSql As String
'Dim oConecta As COMConecta.DCOMConecta
'Dim R As ADODB.Recordset
'
'    On Error GoTo ErrorPreAmpliacion
'    sSql = "exec stp_sel_ValidaPreAmpliacion " & pnMonNro & ""
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set R = oConecta.CargaRecordSet(sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    If Not R.BOF And Not R.EOF Then
'        ValidaPreAmpliacion = True
'    Else
'        ValidaPreAmpliacion = False
'    End If
'    R.Close
'    Set R = Nothing
'    Exit Function
'
'ErrorPreAmpliacion:
'                  Err.Raise Err.Number, "Error En Proceso", Err.Description
'End Function
'VAPA20171019 END DIFERIDOS '******************END NAGL 20180609********************


'CTI2 20181011
Public Sub OtrasOpeDepCtaBanco(ByVal sMovNro As String, ByVal sPersCod As String, _
        ByVal sDocumento As String, ByVal nMonto As Double, ByVal sGlosa As String, _
        ByVal sNomAge As String, ByVal nMoneda As Moneda, _
        ByVal psOpeCod As COMDConstantes.CaptacOperacion, _
        Optional ByRef pnMovNro As Long = 0)

'Dim clsMov As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim nMovNro As Long
Dim dFecSis As Date
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
'Set clsMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
On Error GoTo ErrGraba
    
    AgregaMov sMovNro, psOpeCod, sGlosa, gMovEstContabMovContable, gMovFlagVigente
    nMovNro = GetnMovNro(sMovNro)
    AgregaMovDoc nMovNro, TpoDocBolDep, sDocumento, dFecSis
    AgregaMovOpeVarias nMovNro, nMonto, sDocumento, sPersCod, nMoneda, psOpeCod
    pnMovNro = nMovNro
    
    Exit Sub
ErrGraba:
    Err.Raise Err.Number, "Error", Err.Description
End Sub
Public Function GrabaDeposiBancoNegocioFinan(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                            ByVal psCtaContBanco As String, ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                            ByVal pnDocTpo As TpoDoc, ByVal psDocNro As String, ByVal psFechaDoc As String, _
                                            ByVal psDocNroVoucher As String, _
                                            ByVal lnTpoObjIf As TpoObjetos, ByVal psCtaIf As String, ByVal pbDeposito As Boolean, _
                                            Optional pnMotivoNANC As Long = -1, Optional psMotObjPadre As String = "", Optional psMotObjeto As String = "", Optional ByVal psctaAho As String = "", _
                                            Optional ByVal pnDocTpoNANCAux As TpoDoc = -1, Optional ByVal psNroDocNANCAux As String = "", Optional ByVal pnImporteAux As Currency, _
                                            Optional pnMotivoNANCAux As Long = -1, Optional psMotObjPadreAux As String = "", Optional psMotObjetoAux As String = "", Optional ByVal psCtaAhoAux As String = "", _
                                            Optional ByVal pbGrabaNegocio As Boolean = False, Optional ByVal pbBitCentral As Boolean = False, _
                                            Optional psOpeCodNegocio As String = "", _
                                            Optional psAjusteCta As String = "", Optional pnAjusteMonto As Currency = 0, Optional pnTCBanco As Currency = 0, Optional pnMovNroRef As Long) As Integer

Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim lsMovNro As String
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim lbTransGeneral As Boolean
Dim lsSubCta As String
Dim lsMovNroNegocio As String
Dim lsMsgErr As String

On Error GoTo ErrorGrabaDepRetBancosVarios

GrabaDeposiBancoNegocioFinan = 1
lbTransGeneral = False
lbTrans = False

InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = GetnMovNro(psMovNro)
InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1
lsSubCta = GetFiltroObjetos(ObjEntidadesFinancieras, psCtaContBanco, psCtaIf, False)

InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaContBanco + lsSubCta, pnImporte * IIf(pbDeposito, 1, -1)

lnMovOrden = lnMovOrden + 1
InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, Mid(psCtaIf, 4, 13), Mid(psCtaIf, 1, 2), Mid(psCtaIf, 18, 10)

lnMovItem = lnMovItem + 1: lnMovOrden = 0

InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber + lsSubCta, (pnImporte - pnAjusteMonto) * IIf(pbDeposito, -1, 1)

InsertaMovRef lnMovNro, pnMovNroRef

GrabaDeposiBancoNegocioFinan = 0

If Mid(psOpeCod, 3, 1) = 2 Then
    GeneraMovME lnMovNro, psMovNro, pnTCBanco
End If

ActualizaSaldoMovimiento psMovNro, "+"

Exit Function
ErrorGrabaDepRetBancosVarios:
    lsMsgErr = Err.Description
    Err.Raise vbObjectError + 100, "GrabaRendicionGiroDocumento", lsMsgErr
End Function
Public Sub AgregaMovDoc(ByVal nMovNro As Long, ByVal nTipoDoc As TpoDoc, _
        ByVal sDocumento As String, ByVal dFecDocumento As Date)
Dim sFecha As String
Dim sSql As String
sFecha = Format$(dFecDocumento, "mm/dd/yyyy")
sSql = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
    & "VALUES (" & nMovNro & "," & nTipoDoc & ",'" & sDocumento & "','" & sFecha & "')"
coConex.Ejecutar sSql
End Sub
Public Sub AgregaMovOpeVarias(ByVal pnMovNro As Long, ByVal pnMovImporte As Double, _
        ByVal psNroDoc As String, ByVal psReferencia As String, ByVal pnMoneda As Moneda, _
        Optional nOperacion As CaptacOperacion, _
        Optional ByVal pnConcepto As Integer = 0)
        
    Dim SQL As String
    
    SQL = " Insert MovOpeVarias (nMovNro,nMovImporte,cNroDoc,cReferencia,nMoneda,cOpeCod,nConcepto)" _
        & " Values(" & pnMovNro & "," & pnMovImporte & ",'" & psNroDoc & "','" & psReferencia & "'," & pnMoneda & ",'" & Trim(nOperacion) & "'," & pnConcepto & ")"
        
    coConex.Ejecutar SQL
End Sub
Public Function GetFiltroObjetos(ByVal pnTipoObj As TpoObjetos, ByVal psCtaContCod As String, ByVal psObjetoCod As String, _
                                 Optional lbMuestraCta As Boolean = True, Optional lbExisteFiltro As Boolean = False) As String
    On Error GoTo GetCtaObjFiltroErr
    Dim SQL As String
    Dim rs   As New ADODB.Recordset
'    Dim oConect As COMConecta.DCOMConecta
    Dim lsPersCtaIF As String
    Dim lsCtaIf As String
    Dim lsTpoIf As String
    
'    Set oConect = New COMConecta.DCOMConecta
'    If oConect.AbreConexion = False Then Exit Function
    Dim vsServerCom As String

    GetFiltroObjetos = ""
    Select Case pnTipoObj
        Case ObjCMACAgencias
            SQL = "SELECT cAgeCod,  cSubCtaCod as SubCta  FROM " & vsServerCom & "Agencias where cAgeCod='" & psObjetoCod & "'"
        Case ObjCMACArea
            SQL = "SELECT cAreaCod, cSubCtaCod  as SubCta  FROM " & vsServerCom & "AREAS where cAreaCod='" & psObjetoCod & "'"
        Case ObjCMACAgenciaArea
           SQL = "     SELECT   AA.cAreaCod, AA.cAgeCod, " _
                & "             CASE " _
                & "                  WHEN CF.cSubctaCod IS NULL THEN AA.cSubCtaCod " _
                & "                  Else CF.cSubCtaCod END As SubCta " _
                & "    FROM     " & vsServerCom & "AREAAGENCIA AA " _
                & "             LEFT JOIN ( Select cAreaCod , cAgeCod , cSubCtaCod " _
                & "                          From " & vsServerCom & "CtaAreaAgeFiltro " _
                & "                          WHERE  cCtaContcod in ('" & psCtaContCod & "')) AS CF " _
                & "             ON CF.cAreacod = AA.cAreaCod and CF.cAgeCod = AA.cAgeCod " _
                & "    WHERE    AA.cAgeCod='" & Mid(psObjetoCod, 4, 2) & "' AND  AA.cAreaCod='" & Mid(psObjetoCod, 1, 3) & "'"
                
        Case ObjEntidadesFinancieras
            'If Len(psObjetoCod) > 15 Then
                lsTpoIf = Mid(psObjetoCod, 1, 2)
                lsPersCtaIF = Mid(psObjetoCod, 4, 13)
                lsCtaIf = Mid(psObjetoCod, 18, 10)
            'Else
            '    If psPersCodIf <> "" Then Exit Function
            '    lsPersCtaIF = psPersCodIf
            '    lsCtaIf = psObjetoCod
            'End If
            If Len(psObjetoCod) = 13 Then
                SQL = "select cPersCod, cSubCtaContcod SubCta FROM InstitucionFinanc where cPersCod = '" & psObjetoCod & "' "
            Else
                SQL = "Select cPersCod, cCtaIfCod, cCtaContCod, cCtaIFSubCta as SubCta From " & vsServerCom & "CtaIFFiltro WHERE  cPersCod = '" & lsPersCtaIF & "' and cIFTpo='" & lsTpoIf & "' AND cCtaIfCod='" & lsCtaIf & "' AND cCtaContCod in ('" & psCtaContCod & "')"
            End If
        Case Else
            SQL = "SELECT cCtaObjSubCta as SubCta FROM " & vsServerCom & "CtaObjFiltro WHERE cCtaContCod in ('" & psCtaContCod & "') and cObjetoCod = '" & psObjetoCod & "'"
    End Select
    Set rs = coConex.CargaRecordSet(SQL)
    If Not rs.EOF Then
        lbExisteFiltro = True
        GetFiltroObjetos = IIf(lbMuestraCta, psCtaContCod, "") & rs!SubCta
    End If
    If GetFiltroObjetos = "" Then
        GetFiltroObjetos = IIf(lbMuestraCta, psCtaContCod, "")
    End If
    Exit Function
GetCtaObjFiltroErr:
    'Call RaiseError(MyUnhandledError, "NContFunciones:GetCtaObjFiltro Method")
End Function
Public Sub Actualizar_DetallexEst_MovCorresponsalia(ByVal id As Double, ByVal nMovNro As String, ByVal nMovNroR As String, _
                                                    ByVal nMovNroC As String, Optional ByVal nComision As Double)
    
    Dim sSql As String
    
    On Error GoTo ErrorActualizar_DetallexEst_Mov
 
    sSql = "Exec stp_upd_LogCorresponsaliaBcoNacDet_Mov " & id & ",'" & nMovNro & "','" & nMovNroR & "','" & nMovNroC & "'," & nComision & ""
      
    coConex.Ejecutar sSql
        
    Exit Sub
ErrorActualizar_DetallexEst_Mov:
    Err.Raise Err.Number, "Actualizar Datos Detalle", Err.Description
End Sub
Public Sub Insertar_Datos_Detalle_BN(ByVal id As Double, ByVal prs As ADODB.Recordset, _
                                     Optional x As Boolean = False)
    
    Dim sSql As String
    Dim aux As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorInsertar_Datos_Detalle_BN
   
   If x Then
    aux = 1
   Else
    aux = 0
   End If
   
   sSql = "Exec stp_ins_LogCobrosBcoNacDet " & id & ",'" & prs("cCodCta") & "','" & prs("cCodCtaF") & "','" & prs("nNroCta") & "','" & prs("cSituacion") & "','" & prs("nmoneda") & "','" & prs("cNomCliente") & "', '" & prs("nImporteCuota") & "', '"
   sSql = sSql & prs("dCaducidad") & "','" & prs("nIndTasa") & "', '"
   sSql = sSql & prs("nfactorMora") & "','" & prs("nFactorCompensacion") & "','" & prs("nImporteGastos") & "' ,'"
   sSql = sSql & prs("cCuentaCliente") & "','" & prs("cOrdenCobro") & "','" & prs("nMora") & "','" & prs("nCompensacion") & "', '"
   sSql = sSql & prs("nImporteCobrado") & "','" & prs("cAgencia") & "','" & prs("dFechaCobro") & "' ,'"
   sSql = sSql & prs("dHoraCobro") & "' , '00',"
   sSql = sSql & aux
       
    coConex.Ejecutar sSql

    Exit Sub
ErrorInsertar_Datos_Detalle_BN:
    Err.Raise Err.Number, "Error Insertar Datos Detalle BN", Err.Description
End Sub
Public Function DevolverIDDatos_Detalle(ByVal id_cab As Double, ByVal nCuota As Integer, ByVal cCtaCod As String, _
                                        ByVal dfpago As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
       
    sSql = " exec stp_sel_LogCobrosBcoNacDetalleID " & id_cab & " , " & nCuota & " , '" & cCtaCod & "' , '" & dfpago & "'"
    Set DevolverIDDatos_Detalle = coConex.CargaRecordSet(sSql)
    
End Function
Public Sub Actualizar_DetallexEst_Mov(ByVal id As Double, ByVal nMovNro As String, ByVal nMovNroR As String, _
                                      ByVal nMovNroC As String)
    
    Dim sSql As String
    On Error GoTo ErrorActualizar_DetallexEst_Mov
    
    sSql = "Exec stp_upd_LogCobrosBcoNacDet_Mov " & id & ",'" & nMovNro & "','" & nMovNroR & "','" & nMovNroC & "'"

    coConex.Ejecutar sSql

    Exit Sub
ErrorActualizar_DetallexEst_Mov:
    Err.Raise Err.Number, "Actualizar Datos Detalle", Err.Description
End Sub
Public Sub Insertar_Datos_SobrantesCorresponsalia(cCtaCod As String, nCuota As Integer, nMontoDif As Double, _
                                                  dfpago As Date, cNomCliente As String, nMovNro As Long)
    Dim sSql As String
    On Error GoTo ErrorInsertar_Datos_Sobrantes
   
    sSql = "Exec stp_ins_MovCorresponsaliaBcoNacSobrante '" & cCtaCod & "'," & nCuota & "," & nMontoDif & ", '" & Format(dfpago, "YYYYMMDD") & "','" & cNomCliente & "'," & nMovNro & ""
   
    coConex.Ejecutar sSql
       
    Exit Sub
ErrorInsertar_Datos_Sobrantes:
    Err.Raise Err.Number, "Cargar Datos Generales", Err.Description
End Sub
Public Sub Insertar_Datos_Sobrantes(ByVal cCtaCod As String, ByVal nCuota As Integer, ByVal nMontoDif As Double, _
                                    ByVal dfpago As Date, ByVal cNomCliente As String, ByVal nMovNro As Long)
    Dim sSql As String
    On Error GoTo ErrorInsertar_Datos_Sobrantes
   
    sSql = "Exec stp_ins_MovCobrosBcoNacSobrante '" & cCtaCod & "'," & nCuota & "," & nMontoDif & ", '" & Format(dfpago, "YYYYMMDD") & "','" & cNomCliente & "'," & nMovNro & ""
       
    coConex.Ejecutar sSql
      
    Exit Sub
ErrorInsertar_Datos_Sobrantes:
    Err.Raise Err.Number, "Cargar Datos Generales", Err.Description
End Sub
Public Sub Insertar_Datos_Detalle_corresponsaliaBN(ByVal id As Double, ByVal prs As ADODB.Recordset, _
                                                   Optional x As Boolean = False)
    
    Dim sSql As String
    Dim aux As String
    
    On Error GoTo ErrorInsertar_Datos_Detalle_corresponsaliaBN
    If x Then
     aux = 1
    Else
     aux = 0
    End If
    
    sSql = "Exec stp_ins_LogCorresponsaliaBcoNacDet " & id & ",'" & prs("cCodCta") & "','" & prs("cCodCtaF") & "','" & prs("nNroCta") & "','" & prs("nmoneda") & "','" & prs("cNomCliente") & "', '" & prs("nImporteCobrado") & "', '"
    sSql = sSql & prs("dFechaCobro") & "',"
    sSql = sSql & aux
     
    coConex.Ejecutar sSql

    Exit Sub
ErrorInsertar_Datos_Detalle_corresponsaliaBN:
    Err.Raise Err.Number, "Error Insertar Datos Detalle corresponsalia BN", Err.Description
End Sub
Public Function DevolverIDDatos_DetalleCorresponsalia(ByVal id_cab As Double, ByVal nCuota As Integer, _
                                                      ByVal cCtaCod As String, ByVal dfpago As String) As ADODB.Recordset

    Dim sSql As String
       
    sSql = " exec stp_sel_LogCorresponsaliaBcoNacDetalleID " & id_cab & " , " & nCuota & " , '" & cCtaCod & "' , '" & dfpago & "'"
    Set DevolverIDDatos_DetalleCorresponsalia = coConex.CargaRecordSet(sSql)

End Function
'END CTI2
'APRI20181026 ERS071-2018 - MEJORA
Public Sub InsertarObservacionesCredSegDesg(ByVal psCtaCod As String, ByVal pnMonto As Currency)
    Dim lsSQL As String
    On Error GoTo ErrInsertarObservacionesCredSegDesg
    lsSQL = "EXEC STP_INS_SEGDESGCREDITOOBS '" & psCtaCod & "'," & pnMonto
    coConex.Ejecutar lsSQL
    Exit Sub
ErrInsertarObservacionesCredSegDesg:
    Err.Raise Err.Number, "Error al insertar obs. Seg. Desg.", Err.Description
End Sub
Public Sub InsertaSegMultiriesgoDesgravamenTrama(ByVal psCtaCod As String)
    Dim sSql As String
    On Error GoTo ErrInsertaSegMultiriesgoDesgravamenTrama
    sSql = "exec stp_ins_SegMultiriesgoDesgravamenTrama '" & psCtaCod & "'"
    coConex.Ejecutar sSql
    Exit Sub
ErrInsertaSegMultiriesgoDesgravamenTrama:
    Err.Raise Err.Number, "Error al insertar Seg. Desg.y Seg. Mult.", Err.Description
End Sub
Public Sub ExonerarSeguroTrama(ByVal psCtaCod As String)
    Dim sSql As String
    On Error GoTo ErrExonerarSegDesgravamen
    sSql = "EXEC stp_upd_ExonerarSeguroTrama '" & psCtaCod & "'"
    coConex.Ejecutar sSql
    Exit Sub
ErrExonerarSegDesgravamen:
    Err.Raise Err.Number, "Error al exonerar Seg. Desg.", Err.Description
End Sub
'END APRI

'CTI2 20181205 ******************
Public Sub ActualizarDetallexEstMovBCP(ByVal id As Double, ByVal nMovNro As String, ByVal nMovNroR As String, ByVal nMovNroC As String)
    Dim sSql As String
    
    On Error GoTo ErrorActualizarDetallexEstMovBCP
 
    sSql = "Exec stp_upd_LogCobrosBcoBCPDet_Mov " & id & ",'" & nMovNro & "','" & nMovNroR & "','" & nMovNroC & "'"
       
    coConex.Ejecutar sSql
    
    Exit Sub
ErrorActualizarDetallexEstMovBCP:
    Err.Raise Err.Number, "Actualizar Datos Detalle", Err.Description
End Sub
Public Sub InsertarDatosBCPSobrantes(ByVal cCtaCod As String, ByVal nCuota As Integer, ByVal nMontoDif As Double, ByVal dfpago As Date, ByVal cNomCliente As String, ByVal nMovNro As Long)
    Dim sSql As String
    
    On Error GoTo ErrorInsertarDatosBCPSobrantes
   
    sSql = "Exec stp_ins_MovCobrosBcoBCPSobrante '" & cCtaCod & "'," & nCuota & "," & nMontoDif & ", '" & Format(dfpago, "YYYYMMDD") & "','" & cNomCliente & "'," & nMovNro & ""
    coConex.Ejecutar sSql
       
    Exit Sub
ErrorInsertarDatosBCPSobrantes:
    Err.Raise Err.Number, "Cargar Datos Generales", Err.Description
End Sub
Public Sub InsertarDatosDetalleBCP(ByVal id As Double, ByVal prs As ADODB.Recordset)
    Dim sSql As String
    On Error GoTo ErrorInsertarDatosDetalleBCP
    
    sSql = "Exec stp_ins_LogCobrosBcoBCPDet " & id & ",'" & prs("cCodCtaF") & "'," & CInt(prs("Cuota")) & ",'" & prs("Titular") & "','" & prs("c1") & "','" & prs("c2") & "','" & prs("c3") & "', '" & prs("c4") & "', '"
    sSql = sSql & prs("c5") & "','" & prs("c6") & "', '" & prs("c7") & "','" & prs("c8") & "'," & prs("c9") & " ,"
    sSql = sSql & prs("c10") & "," & prs("c11") & ",'" & prs("c12") & "','" & prs("c13") & "', '"
    sSql = sSql & prs("c14") & "','" & prs("c15") & "','" & prs("c16") & "' ,'" & prs("c17") & "','" & prs("c18") & "','" & prs("c19") & "' ,'"
    sSql = sSql & prs("c20") & "','" & prs("c21") & "' "
           
    coConex.Ejecutar sSql
    
    Exit Sub
ErrorInsertarDatosDetalleBCP:
    Err.Raise Err.Number, "Error Insertar Datos BCP Detalle", Err.Description
End Sub
Public Function getDatosDetalleBcoBCPxId(ByVal id_cab As Double, ByVal cCtaCod As String, _
                                         Optional ByVal nCuota As Integer = -1) As ADODB.Recordset
    'nCuota ADD CTI2 20190802
    Dim sSql As String
    sSql = " exec stp_sel_LogCobrosBcoBCPDetalleID " & id_cab & " ,'" & cCtaCod & "', " & nCuota
    Set getDatosDetalleBcoBCPxId = coConex.CargaRecordSet(sSql)
    
    
End Function
'CTI2 END ***********************
'CTI2 20181207 ACTA 182-2018 Mejoras Voucher *******
Public Sub actualizarRelacionVoucherCaptacion(ByVal pnMovNroOpe As Long, ByVal pnMovNroRVD As Long)
    Dim lsSQL As String
    lsSQL = "exec stp_upd_ActualizarCapVoucherDeposito_2  " & pnMovNroOpe & ", " & pnMovNroRVD & ""
    coConex.Ejecutar lsSQL
End Sub
Public Sub ActualizaMovPendientesRend(ByVal nMovNro As Long, ByVal pnMonto As Currency)
    Dim sSql As String
    sSql = " UPDATE MovPendientesRend " _
         & " Set nSaldo = nSaldo - " & pnMonto & "" _
         & " Where nMovNro = " & nMovNro & ""
    coConex.Ejecutar sSql
End Sub
'CTI2 END ******************************************
'*** CTI3 (ferimoro) 16102018
Public Sub dInsertDetExtMovEcoTaxi(ByVal pnMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        Optional ByVal psMotExtorno As String, _
        Optional ByVal psDetMotExtorno As String)
Dim sSql As String
Dim psOperacion, psCuenta As String
 psCuenta = ""
psOperacion = nOperacion
    sSql = "INSERT MovDetExtornos(nMovNro,cOpedCod,cCtaCod,cMotExtorno,cDetMotExto) " _
          & "VALUES (" & pnMovNro & ",'" & psOperacion & "','" & psCuenta & "','" & psMotExtorno & "','" & UCase(psDetMotExtorno) & "')"

coConex.ConexionActiva.Execute sSql

End Sub
'*** CTI3 (ferimoro) 16102018
Public Sub dInsertDetExtMov(ByVal pnMovNro As Long, ByVal psOperacion As String, _
        ByVal pnEstado As Integer, Optional psCuenta As String, _
        Optional pbEjecBatch As Boolean = False, _
        Optional ByVal psMotExt As String, _
        Optional ByVal psDetMotExt As String)

Dim lsSQL As String
Dim lrs As ADODB.Recordset
Set lrs = New ADODB.Recordset
 
    lsSQL = "INSERT MovDetExtornos (nMovNro,cOpedCod,cCtaCod,cMotExtorno,cDetMotExto) " _
          & "VALUES (" & pnMovNro & ",'" & psOperacion & "','" & psCuenta & "','" & psMotExt & "','" & UCase(psDetMotExt) & "')"

    Set lrs = coConex.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrs = Nothing

End Sub
'CTI2 20181219 ERS075-2018 INI *************************
Public Sub RegistraMovPerdon(ByVal pnIdSolicitudPerdon As Long, ByVal pnMovNro As Long)
    Dim sSql As String
    sSql = " stp_upd_AutorizacionActualizacion " & pnIdSolicitudPerdon & ", " & pnMovNro
    coConex.Ejecutar sSql
End Sub
'CTI2 20181219 ERS075-2018 FIN *************************

'CTI2 20190423 Corrección de distribución de gastos en pago de creditos convenio en lote
Private Function DevuelveMatrizGastosCredito(ByRef pnNumreg As Integer, _
                                             ByVal psCtaCod As String, _
                                             ByVal nNroCalen As Integer) As Variant
Dim R As ADODB.Recordset
Dim oCalend As COMDCredito.DCOMCalendario
Dim MatGastosCred() As String
        Set oCalend = New COMDCredito.DCOMCalendario
        Set R = oCalend.RecuperaCalendarioGastos(psCtaCod, nNroCalen, 0, gColocCalendAplCuota, True) 'CTI2 ADD pbExcluyeGastos 20190101
        Set oCalend = Nothing
        pnNumreg = 0
        pnNumreg = R.RecordCount
        If R.RecordCount > 0 Then
            ReDim MatGastosCred(R.RecordCount, 3)
            Do While Not R.EOF
                MatGastosCred(R.Bookmark - 1, 0) = Trim(Str(R!nCuota))
                MatGastosCred(R.Bookmark - 1, 1) = Trim(Str(R!nPrdConceptoCod))
                MatGastosCred(R.Bookmark - 1, 2) = Format(R!nMonto - R!nMontoPagado, "#0.00")
                R.MoveNext
            Loop
        End If
        R.Close

        DevuelveMatrizGastosCredito = MatGastosCred

End Function
Private Function DevuelveMatrizGastosCreditoCuota(pnNumreg As Integer, _
                                                  ByVal pnCuota As Integer, _
                                                  MatGastosCred As Variant, _
                                                  ByVal pnNumRegGasCred As Integer) As Variant
Dim i As Integer
Dim nCont As Integer
Dim MatGastosCredCuota() As String
    nCont = 0
    For i = 0 To pnNumRegGasCred - 1
        If pnCuota = CInt(MatGastosCred(i, 0)) Then
           nCont = nCont + 1
        End If
    Next i

    ReDim MatGastosCredCuota(nCont, 3)
    nCont = 0
    For i = 0 To pnNumRegGasCred - 1
        If pnCuota = CInt(MatGastosCred(i, 0)) Then
           nCont = nCont + 1
           MatGastosCredCuota(nCont - 1, 0) = MatGastosCred(i, 0) 'nCuota
           MatGastosCredCuota(nCont - 1, 1) = MatGastosCred(i, 1) 'nPrdConceptoCod
           MatGastosCredCuota(nCont - 1, 2) = MatGastosCred(i, 2) 'nMontoPendiente
        End If
    Next i
    pnNumreg = nCont
    DevuelveMatrizGastosCreditoCuota = MatGastosCredCuota
End Function
'END CTI2 END


'ANGC/RIRO 20200520 *****

Public Sub CofideReactivaValidaComision(ByRef psCtaCod As String, _
                                         ByRef pbAplicaComision As Boolean, _
                                         ByRef pnMontoComision As Double)
Dim rsComision As ADODB.Recordset

Dim sSql As String
sSql = " sp_sel_Calcula_Comision '" & psCtaCod & "'"
Set rsComision = coConex.CargaRecordSet(sSql)

pbAplicaComision = False
pnMontoComision = -1
    
If Not rsComision Is Nothing Then
    If rsComision.RecordCount > 0 Then
          pbAplicaComision = rsComision!bAplicaComision
          pnMontoComision = rsComision!nComision
    End If
    rsComision.Close
End If

Set rsComision = Nothing

End Sub
'END

'CTI5 INICIO 20200803
Public Function CalcularDiasAtrasoDespuesPago(ByVal sCtaCod As String, ByVal nDiasAtrasoCalc As Integer, ByVal sFecha As String) As Integer
    Dim lrs As ADODB.Recordset
    Dim sSql As String
    Dim nDiasAtrasoFinal As Integer
    
    sSql = "Exec stp_sel_CalcularDiasAtrasoDespuesPago '" & sCtaCod & "'," & nDiasAtrasoCalc & ",'" & sFecha & "'"
    
    Set lrs = coConex.CargaRecordSet(sSql, adLockReadOnly)
    nDiasAtrasoFinal = nDiasAtrasoCalc
    If Not lrs.EOF And Not lrs.BOF Then
        nDiasAtrasoFinal = CInt(lrs!nDiasAtrasoFinal)
    End If
    CalcularDiasAtrasoDespuesPago = nDiasAtrasoFinal
End Function
'CTI5 FIN 20200803

'RIRO 20200901 **********************************************
Public Sub RegistraLiquidacion(ByVal psCtaCod As String, _
                               ByVal psUser As String, _
                               ByVal pnMovNro As Long, _
                               ByRef pnIdLiquidacion As Long, _
                               ByVal pnCuotasTotal As Integer, _
                               ByVal pnTipoLiquidacion As Integer)
                               
Dim rsLiquidacion As ADODB.Recordset
Dim sSql As String

sSql = " stp_ins_Liquidacion '" & psCtaCod & "', '" & psUser & "', " & _
                                  pnMovNro & ", " & pnIdLiquidacion & ", " & _
                                  pnCuotasTotal & ", " & pnTipoLiquidacion

Set rsLiquidacion = coConex.CargaRecordSet(sSql)
 
If Not (rsLiquidacion Is Nothing) Then
    If Not rsLiquidacion.EOF And Not rsLiquidacion.BOF Then
        If rsLiquidacion.RecordCount > 0 Then
              pnIdLiquidacion = rsLiquidacion!nIdLiquidacion
        End If
    End If
End If
Set rsLiquidacion = Nothing
End Sub
Public Sub RegistraLiquidacionDet(ByVal pnIdLiquidacion As Long, _
                                  ByVal pnCuota As Integer, _
                                  ByVal pnPrdConceptoCod As Integer, _
                                  ByVal pnMonto As Double, _
                                  ByVal pnTipo As Integer)
                                  
    'RIRO 20200904 / pnTipo : 1-> Calculado, 2 -> asignado
                               
    Dim rsLiquidacion As ADODB.Recordset
    Dim sSql As String
    
    sSql = "stp_ins_InsertColocCalendLiquidacionDet " & _
            pnIdLiquidacion & ", " & _
            pnCuota & ", " & _
            pnPrdConceptoCod & ", " & _
            pnMonto & "," & pnTipo
          
    Set rsLiquidacion = coConex.CargaRecordSet(sSql)
    
End Sub
Public Sub ExtornaLiquidacion(ByVal psCtaCod As String)

    Dim rsLiquidacion As ADODB.Recordset
    Dim sSql As String
    
    sSql = "stp_upd_ActualizaLiquidacionCred '" & psCtaCod & "'"
    coConex.CargaRecordSet (sSql)
    
End Sub
Public Sub ActualizaCalendario(ByVal psCtaCod As String, ByVal pnMovNro As Long, _
                               ByVal pbVigeref As Boolean, ByVal psCodOpe As String)

    Dim rsLiquidacion As ADODB.Recordset
    Dim sSql As String
    
    sSql = "stp_upd_ActualizaCalendario '" & psCtaCod & "', " & pnMovNro & ", " & IIf(pbVigeref, 1, 0) & ", '" & psCodOpe & "'"
    coConex.CargaRecordSet (sSql)
    
End Sub
'END RIRO ****************************************************


'ACTA Nº 061 - 2020 JOEP-LIPA Mejora en la Propuesta de Reprogramacion
Public Sub dInsertaColocacReprogActiEco(ByVal psCtaCod As String, ByVal ActividadFormEval As String, ByVal NegVinFunc As Integer, ByVal CliMatAct As Integer, NewActividad, ByVal IngActuales As Currency, ByVal DimEs As Integer, ByVal DimPorc As Integer, ByVal IngActPermt As Currency, ByVal IngVentasNet As Currency, ByVal GastTotal As Currency, ByVal Excedente As Currency, ByVal nPrdEstado As Integer, ByVal dPrdEstado As String, ByVal cMovnro As String)
Dim lsSQL As String

    lsSQL = "Exec stp_ins_ReprogPropuestaActEco '" & psCtaCod & "','" & ActividadFormEval & "'," & NegVinFunc & "," & CliMatAct & ",'" & NewActividad & "'," & IngActuales & "," & DimEs & "," & DimPorc & "," & IngActPermt & "," & IngVentasNet & "," & GastTotal & "," & Excedente & "," & nPrdEstado & ",'" & dPrdEstado & "','" & cMovnro & "'"
    coConex.Ejecutar lsSQL
End Sub

Public Sub dInsertaColocacReprogIfis(ByVal psCtaCod As String, ByVal pdFechaRcc As String, ByVal pnMonedad As Integer, ByVal pnIfis As String, ByVal pnMonto As Currency, ByVal pnMontoCuota As Currency, ByVal nPrdEstado As Integer, ByVal dPrdEstado As String, ByVal cMovnro As String)
Dim lsSQL As String

    lsSQL = "Exec stp_ins_ReprogPropuestaIfis '" & psCtaCod & "','" & Format(pdFechaRcc, "yyyymmdd") & "'," & pnMonedad & ",'" & pnIfis & "'," & pnMonto & "," & pnMontoCuota & "," & nPrdEstado & ",'" & dPrdEstado & "','" & cMovnro & "'"
    coConex.Ejecutar lsSQL
End Sub

Public Sub dInsertaColocacReprogModalidades(ByVal psCtaCod As String, ByVal pnModalidad As Integer, ByVal pcCalif As String, ByVal pnCuota As Currency, ByVal pdPrdEstado As String, ByVal nPrdEstado As Integer, ByVal cMovnro As String)
Dim lsSQL As String

    lsSQL = "Exec stp_ins_ReprogPropuestaModalidades '" & psCtaCod & "'," & pnModalidad & ",'" & pcCalif & "'," & pnCuota & ",'" & pdPrdEstado & "'," & nPrdEstado & ",'" & cMovnro & "'"
    coConex.Ejecutar lsSQL
End Sub
'ACTA Nº 061 - 2020 JOEP-LIPA Mejora en la Propuesta de Reprogramacion

'reversion
Public Sub dInsertaColocacReprogRechazo(ByVal psCtaCod As String, ByVal pdPrdEstado As String, ByVal pnPrdEstado As Integer, ByVal pcMovNro As String, ByVal pcMotivo As String, ByVal pnNewDias As Integer, ByVal pnOpcion As Integer)
Dim lsSQL As String

    lsSQL = "Exec stp_inst_ReprogramacionRechazo_Datos '" & psCtaCod & "','" & pdPrdEstado & "'," & pnPrdEstado & ",'" & pcMovNro & "','" & pcMotivo & "'," & pnNewDias & "," & pnOpcion & ""
    coConex.Ejecutar lsSQL
End Sub
'reversion
'CTI5 ERS0012021**********************************************************************************************************
Public Sub dUpdateGarantiaFechaConstatacion(ByVal psCtaCod As String, ByVal pdFecha As Date)

Dim lsSQL As String

    On Error GoTo ErrordUpdateGarantiaFechaConstatacion
    lsSQL = "UPDATE Gar SET "
    lsSQL = lsSQL & " dConstatacion = '" & Format(pdFecha, "yyyymmdd") & "' "
    lsSQL = lsSQL & " From Garantias Gar inner join ColocGarantia CG on Gar.cNumGarant=CG.cNumGarant "
    lsSQL = lsSQL & " WHERE CG.cCtaCod = '" & psCtaCod & "' and Gar.nEstado=1 "
    coConex.Ejecutar lsSQL
    Exit Sub

ErrordUpdateGarantiaFechaConstatacion:
    Err.Raise Err.Number, "Error En Proceso ErrordUpdateGarantiaFechaConstatacion", Err.Description

End Sub
'*************************************************************************************************************************
Public Sub ActualizaCuentaAbonoGastoFMV(ByVal psCtaCod As String, ByVal psCtaCodAhoGasto As String)
    Dim sSql As String
    
    sSql = "UPDATE CredMiVivienda SET cCtaCodGasto = '" & psCtaCodAhoGasto & "' WHERE cCtaCod = '" & psCtaCod & "' AND nPrdEstado = 2002"
    coConex.ConexionActiva.Execute sSql
End Sub
'JOEP20210211 Garantia covid
Public Sub dInsertaColocacReprogCalen(ByVal psCtaCod As String, ByVal pnPrdEstado As Integer, ByVal pdPrdEstado As String, ByVal pnNroCalen As Integer, ByVal pnColocCalenApl As Integer, ByVal pnCuota As Integer, ByVal pdVenc As String, ByVal pdPago As String, ByVal pnColcClenEstado As Integer, ByVal pnOpCovid As Integer)
Dim lsSQL As String

    lsSQL = "Exec stp_ins_ReprogramadoGarantiaCalendario '" & psCtaCod & "'," & pnPrdEstado & ",'" & pdPrdEstado & "'," & pnNroCalen & "," & pnColocCalenApl & "," & pnCuota & ",'" & Format(pdVenc, "yyyymmdd") & "','" & Format(pdPago, "yyyymmdd hh:mm:ss") & "'," & pnColcClenEstado & "," & pnOpCovid & ""
    
    coConex.Ejecutar lsSQL
End Sub

Public Sub dInsertaColocacReprogCalenDet(ByVal psCtaCod As String, ByVal pnPrdEstado As Integer, ByVal pdPrdEstado As String, ByVal pnNroCalen As Integer, ByVal pnColocCalenApl As Integer, ByVal pnCuota As Integer, ByVal pnPrdConceptoCod As Integer, ByVal pnMonto As Currency, ByVal pnMontoPagado As Currency, ByVal pnOpCovid As Integer)
Dim lsSQL As String

    lsSQL = "Exec stp_ins_ReprogramadoGarantiaCalendarioDet '" & psCtaCod & "'," & pnPrdEstado & ",'" & pdPrdEstado & "'," & pnNroCalen & "," & pnColocCalenApl & "," & pnCuota & "," & pnPrdConceptoCod & "," & pnMonto & "," & pnMontoPagado & "," & pnOpCovid & ""
    coConex.Ejecutar lsSQL
End Sub
'JOEP20210211 Garantia covid
'RIRO 20210512 ***************
Public Sub PerdonarMoraICV(ByVal psCtaCod As String, ByVal pdFecha As Date, ByVal psUser As String)
    
    Dim lsSQL As String
    On Error GoTo ErrorPerdonarMoraICV
  
    lsSQL = "exec stp_upd_PerdonMoraICV '" & Format(pdFecha, "yyyyMMdd") & "', '" & psUser & "', '" & psCtaCod & "'"
    coConex.Ejecutar lsSQL
        
    Exit Sub

ErrorPerdonarMoraICV:
    Err.Raise Err.Number, "Error En Proceso de Perdón de Mora e ICV", Err.Description

End Sub
'END RIRO ********************
