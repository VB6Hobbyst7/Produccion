VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMCaptaMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public dbCmact As ADODB.Connection
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String
Dim sSql As String
Public bTransaccion As Boolean

Dim oConstante As New COMFunciones.FCOMVarPublicas
Public gcPEN As New COMDConstantes.DCOMConstantes 'CTI7 OPEv2
Public oImp As COMFunciones.FCOMVarImpresion 'CTI7 OPEv2
'ALPA20130930**********************************************************************
Public Sub dBeginTrans()
    dbCmact.BeginTrans
End Sub

Public Sub dRollbackTrans()
    dbCmact.RollbackTrans
End Sub
Public Sub dCommitTrans()
    dbCmact.CommitTrans
End Sub
'**********************************************************************************
Public Function GetDataInteresEspecial(ByVal cPersCod As String, ByVal nEstado As Integer, ByVal nProd As Integer, ByVal nMoneda As Integer, ByVal dFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim rstemp As ADODB.Recordset, cFecha1 As String, sqlaux As String
Dim dFecha2 As Date, cFecha2 As String
cFecha1 = Year(dFecha) & String(2 - Len(Month(dFecha)), "0") & CStr(Month(dFecha)) & String(2 - Len(Day(dFecha)), "0") & CStr(Day(dFecha))
dFecha2 = DateAdd("d", -2, dFecha)
cFecha2 = Year(dFecha2) & String(2 - Len(Month(dFecha2)), "0") & CStr(Month(dFecha2)) & String(2 - Len(Day(dFecha2)), "0") & CStr(Day(dFecha2))
If nEstado = 99 Then
    sqlaux = " And nestado in (0,1,3)"
Else
    sqlaux = " And nestado=" & nEstado
End If

Set rstemp = New ADODB.Recordset
    sSql = "Select ntasa,cestado=(case when nestado=0 then 'SOLICITADA' "
    sSql = sSql & "                    when nestado=1 then 'APROBADA' "
    sSql = sSql & "                    when nestado=2 then 'APERTURADA'   "
    sSql = sSql & "                    when nestado=3 then 'RECHAZADA'"
    sSql = sSql & "                    when nestado=4 then 'EXTORNADA'  end ),"
    sSql = sSql & " nestado,nmonto from captasaespecial where cPersCod='" & cPersCod & "' and nProducto= " & nProd & " And nMoneda = " & nMoneda & sqlaux
    sSql = sSql & " and cast(left(cmovnro,8) as bigint)<= cast('" & cFecha1 & "' as bigint) and cast(left(cmovnro,8) as bigint)>=cast('" & cFecha2 & "' as bigint) "
    sSql = sSql & " ORDER BY CAST(LEFT(CMOVNRO,14) AS BIGINT) DESC "
    rstemp.CursorLocation = adUseClient
    rstemp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
   ' Set rstemp.ActiveConnection = Nothing
    Set GetDataInteresEspecial = rstemp
    Set rstemp.ActiveConnection = Nothing
    If rstemp.State = 1 Then rstemp.Close
    
    Set rstemp = Nothing

End Function

Public Function GetnMovNro(ByVal sMovNro As String) As Long
Dim rsMov As ADODB.Recordset
sSql = "Select nMovNro From Mov Where cMovNro = '" & sMovNro & "'"
Set rsMov = New ADODB.Recordset
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
        Optional ByVal bOrdPag As Boolean = False, Optional ByVal nPlazo As Long = 0, _
        Optional nFormaRetiro As CaptacPFFormaRetiro = gCapPFFormRetMensual, _
        Optional bCtaAboInt As Boolean = False, Optional sCtaCodAbono As String = "", _
        Optional nPorcRetCTS As Double = 0, Optional sInstitucion As String = "", _
        Optional ByVal cAlias As String = "", Optional ByVal nFirmasMin As Integer = 1, _
        Optional cPersTasaEspecial As String, Optional bEsTasaPermanente As Boolean, _
        Optional ByVal nMonITF As Currency = 0, Optional ByVal nTpoPrograma As Integer = -1, _
        Optional ByVal nMontoAbono As Double = 0, Optional ByVal nPlazoAho As Long = 0, _
        Optional ByVal nGarantDepCF As Integer = -1, Optional ByVal psCodInstConv As String, _
        Optional ByVal nSubasta As Integer = -1, Optional ByVal sReglas As String, _
        Optional ByVal psCtaCodCargo As String = "", Optional ByVal pnCampanaCod As Long = 0, _
        Optional ByVal pnEsITFCargo As Integer = 0) As String
        'SE AGEREGÓ PARÁMETRO "sReglas", RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
        'JUEZ 20131212 Se agregó psCtaCodCargo
        'JUEZ 20160420 Se agregó pnCampanaCod
        
    Dim clsGen As COMDConstSistema.DCOMGeneral
 
    Dim sCuenta As String, sFecha As String
    Dim sFecUltCierre As String, sOrdPag As String, sCtaAbonoIntPF As String
    Dim nSaldoDisp As Double, nSaldRetCTS As Double
    
    Set clsGen = New COMDConstSistema.DCOMGeneral
    sCuenta = clsGen.GeneraNuevaCuenta(sAgencia, nProducto, nMoneda)
    Set clsGen = Nothing
    sFecha = Format$(dFecha, "mm/dd/yyyy") & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
    sSql = "Insert Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) " _
        & "Values ('" & sCuenta & "'," & nTasa & "," & nSaldo & "," & gCapEstActiva & ",'" & sFecha & "',1)"
    dbCmact.Execute sSql
    'nSaldoDisp = IIf(bCheque, 0, nSaldo)
    nSaldoDisp = IIf(bCheque, nSaldo, nSaldo) 'EJVG20140408 xq cheque ya estará valorizado
    sFecUltCierre = Format$(DateAdd("d", -1, dFecha), "mm/dd/yyyy")
    If nProducto = gCapAhorros Or nProducto = gCapCTS Then 'BRGO 20111021 Se Agregó registro de subproducto de CTS
        'RIRO 20131102 SEGUN "CAMBIOS EN PODERES", SE AGREGO PARAMETRO "sReglas"
        sSql = "Insert Captaciones (cCtaCod,nSaldoDisp,nPersoneria,nFirmas,nIntAcum,dUltCierre,dApertura,nPrdCtaTpo,nPrdTasaInteres,cUltimaActualizacion,cAlias,nFirmasMin,nTpoPrograma,cReglas,cCtaCodCargo,nCampanaCod) " _
            & " Values ('" & sCuenta & "'," & nSaldoDisp & "," & nPersoneria & "," & nFirmas & ",0,'" & sFecUltCierre & "','" & sFecha & "'," & nTipoCuenta & "," _
            & nTipoTasa & ",'" & sMovNro & "','" & cAlias & "'," & nFirmasMin & "," & nTpoPrograma & ",'" & sReglas & "','" & psCtaCodCargo & "'," & pnCampanaCod & ")"
        'JUEZ 20131212 Se agregó psCtaCodCargo
        dbCmact.Execute sSql
    Else
        'RIRO 20131102 SEGUN "CAMBIOS EN PODERES", SE AGREGO PARAMETRO "sReglas"
        sSql = "Insert Captaciones (cCtaCod,nSaldoDisp,nPersoneria,nFirmas,nIntAcum,dUltCierre,dApertura,nPrdCtaTpo,nPrdTasaInteres,cUltimaActualizacion,cAlias,nFirmasMin,nTpoPrograma,cReglas,cCtaCodCargo,nCampanaCod) " _
            & "Values ('" & sCuenta & "'," & nSaldoDisp & "," & nPersoneria & "," & nFirmas & ",0,'" & sFecUltCierre & "','" & sFecha & "'," & nTipoCuenta & "," _
            & nTipoTasa & ",'" & sMovNro & "','" & cAlias & "'," & nFirmasMin & "," & nTpoPrograma & ",'" & sReglas & "','" & psCtaCodCargo & "'," & pnCampanaCod & ")"
        dbCmact.Execute sSql
    End If
    Select Case nProducto
        Case gCapAhorros
            sOrdPag = IIf(bOrdPag, "1", "0")
            If nTpoPrograma = 2 Or nTpoPrograma = 3 Then
                sSql = "Insert CaptacAhorros (cCtaCod,nSaldoAnterior,bOrdPag,nSobregiro,dUltContacto,nPlazo,nMontoAbono,cCodInst,dRenoPanderito) " _
                    & "Values ('" & sCuenta & "',0," & sOrdPag & ",0,'" & sFecha & "'," & nPlazoAho & "," & nMontoAbono & ",'" & sInstitucion & "','" & Format$(dFecha, "mm/dd/yyyy") & "')"
            'By Capi 19082008 para exonerar de cobro de inactivas a la cuenta soñada
            ElseIf nTpoPrograma = 5 Then
                sSql = "Insert CaptacAhorros (cCtaCod,nSaldoAnterior,bOrdPag,nSobregiro,dUltContacto,nPlazo,nMontoAbono,cCodInst,bNoCobroInactivas,cNoCobroInactivas,cPersCodConv) " _
                & "Values ('" & sCuenta & "',0," & sOrdPag & ",0,'" & sFecha & "'," & nPlazoAho & "," & nMontoAbono & ",'" & sInstitucion & "',1, '" & sMovNro & "','" & psCodInstConv & "')"
            '
            Else
                sSql = "Insert CaptacAhorros (cCtaCod,nSaldoAnterior,bOrdPag,nSobregiro,dUltContacto,nPlazo,nMontoAbono,cCodInst,cPersCodConv) " _
                    & "Values ('" & sCuenta & "',0," & sOrdPag & ",0,'" & sFecha & "'," & nPlazoAho & "," & nMontoAbono & ",'" & sInstitucion & "','" & psCodInstConv & "')"
            End If
        Case gCapPlazoFijo
            sCtaAbonoIntPF = IIf(bCtaAboInt, "1", "0")
            '*** BRGO 20111219 Se agregó el campo nMontoAbono al registro
            sSql = "Insert CaptacPlazoFijo (cCtaCod,nPlazo,nIntPag,dRenovacion,nApertura,dAuxiliar,nFormaRetiro,nDuplicado,bAbonoIntCtaAho,bBusCredPend,bPermanente,GarantDepCF,nMontoAbono,Subasta) " _
                & "Values ('" & sCuenta & "'," & nPlazo & ",0,'" & sFecha & "'," & nSaldo & ",'" & sFecha & "'," & nFormaRetiro & ",0," & sCtaAbonoIntPF & ",0," & IIf(bEsTasaPermanente, 1, 0) & "," & IIf(nGarantDepCF = -1, 0, 1) & "," & nMontoAbono & "," & IIf(nSubasta = -1, 0, 1) & ") " 'MIOL 20121110, SEGUN RFC098; Agrego: nSubasta
            '*** END BRGO ************************************
        Case gCapCTS
            nSaldRetCTS = Round(nSaldoDisp * nPorcRetCTS / 100, 2)
            sSql = "Insert CaptacCTS (cCtaCod,nSaldRetiro,nIntSaldo,cCodInst,dUltCierreMes) " _
                 & "Values ('" & sCuenta & "'," & nSaldRetCTS & ",0,'" & sInstitucion & "','" & sFecha & "')"
            'JUEZ 20140304 ********************************
            dbCmact.Execute sSql
            sSql = "Insert Into CapCambioTasaCTS (cCtaCod, nTasaAnt, nTasaCambio, nTipoMov,cMovNroReg,nEstado)" _
                 & "Values ('" & sCuenta & "',0," & nTasa & ",1,'" & sMovNro & "',1)"
            'END JUEZ *************************************
    End Select
    dbCmact.Execute sSql
    If nProducto = gCapPlazoFijo And bCtaAboInt Then
        sSql = "Insert CaptacCtaAboIntPF (cCtaCod,cCtaCodAbono) " _
            & "Values ('" & sCuenta & "','" & sCtaCodAbono & "')"
        dbCmact.Execute sSql
    End If

    AgregaNuevaCaptacion = sCuenta
End Function

' *** RIRO20131102 SEGUN "CAMBIOS EN PODERES"
Public Sub AgregaNuevoProdPers(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona, Optional ByVal COBLIGATORIO As String, Optional ByVal sGrupos As String)
    If Mid(sCuenta, 6, 3) = "232" Or Mid(sCuenta, 6, 3) = "233" Then
        
        ' COMENTADO RIRO20131102
        'sSql = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac,cObligatorio) " _
        '& "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ",'" & COBLIGATORIO & "')"
        
        ' AGREGADO RIRO20131102
        sSql = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac,cObligatorio, cGrupo) " _
             & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ",'N', '" & sGrupos & "')"
        
    Else
        ' COMENTADO RIRO20131102
        'sSql = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac,cObligatorio) " _
        '    & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ",'S')"
        
        ' AGREGADO RIRO20131102
        sSql = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac,cObligatorio, cGrupo) " _
             & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ",'N', '" & sGrupos & "')"
        
    End If
    dbCmact.Execute sSql
End Sub

Public Function GetDatosCheque(ByVal sNroDoc As String, ByVal scperscod As String, ByVal sTipoIF As String) As ADODB.Recordset
Dim rsChe As ADODB.Recordset
sSql = "  select cNroDoc,cPersCod,cIFTpo, bPlaza,cIFCta=isnull(cIFCta,''),nMonto, " _
       & " dValorizacion,cDepIF,nConfCaja,nEstado,nMoneda,cAreaCod,cAgeCod, nProducto  from  docrec " _
       & " where cnrodoc='" & sNroDoc & "' and cperscod='" & scperscod & "' and ciftpo='" & sTipoIF & "'"

Set rsChe = New ADODB.Recordset
rsChe.CursorLocation = adUseClient
rsChe.Open sSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
Set rsChe.ActiveConnection = Nothing
Set GetDatosCheque = rsChe
Set rsChe = Nothing

End Function

Public Function GetVerificaRetiro(ByVal sCuenta As String) As Boolean
Dim rs As ADODB.Recordset

GetVerificaRetiro = False
sSql = "  select mc.nmonto from mov m join movcap mc on mc.nmovnro=m.nmovnro "
sSql = sSql & " where mc.cctacod = '" & sCuenta & "' and mc.copecod='220303' and m.nmovflag=0 "
       
Set rs = New ADODB.Recordset
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


Public Function GetDatosOrdenPago(ByVal sCuenta As String, ByVal sNroDoc As String) As ADODB.Recordset
Dim rsOP As ADODB.Recordset
sSql = "SELECT E.cMovNro, D.nTpoDoc, D.cNroDoc, D.cCtaCod, D.nMonto, P.cPersNombre, E.cEstado " _
    & "D.cIFCodPers FROM DocRecOPEst E INNER JOIN DocRecOP D INNER JOIN " & sDBPersona & "Persona P " _
    & "ON D.cIFCodPers = P.cPersCod ON E.nTpoDoc = D.nTpoDoc AND E.cNroDoc = D.cNroDoc AND E.cCtaCod " _
    & "= D.cCtaCod WHERE E.cMovNro IN (SELECT MAX(E1.cMovNro) FROM DocRecOPEst E WHERE E1.nTpoDoc = " _
    & "E.nTpoDoc AND E1.cNroDoc = E.cNroDoc AND E1.cCtaCod = E.cCtaCod)"
Set rsOP = New ADODB.Recordset
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
        ByVal nMonto As Double, Optional ByVal psIFTpo As String = "01", Optional ByVal psIFCta As String = "")
sSql = "INSERT DocRecCapta (nTpoDoc,cNroDoc,cPersCod,cIFTpo,nMovNro,cCtaCod,nMonto,cIFCta) " _
    & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format(psIFTpo, "00") & "'," & nMovNro & ",'" & sCuenta & "'," & nMonto & ",'" & psIFCta & "')"
dbCmact.Execute sSql
End Sub

Public Sub AgregaCuentaDocumento(ByVal sCuenta As String, ByVal nTpoDoc As TpoDoc, _
        ByVal sNroDoc As String, ByVal sCodIF As String, ByVal sMovNro As String, _
        ByVal nMovNro As Long, Optional nMonto As Double = 0, _
        Optional nEstadoOP As CaptacOrdPagoEstado = gCapOPEstCobrada, _
        Optional psCtaIf As String = "", Optional ByVal psIFTpo As String = "01" _
        , Optional ByVal pbvalor As Boolean = True) 'pbvalor agregado PASI20140617
Dim sFecha As String
sFecha = Mid(sMovNro, 1, 4) & "/" & Mid(sMovNro, 5, 2) & "/" & Mid(sMovNro, 7, 2) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)


Select Case nTpoDoc
    Case TpoDocCheque
        '***Modificado por ELRO el 20121025, según OYP-RFC101-2012
        'sSql = "INSERT DocRecCapta (nTpoDoc,cNroDoc,cPersCod,cIFTpo,nMovNro,cCtaCod,nMonto,cIFCta) " _
        '    & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "'," & nMovNro & ",'" & sCuenta & "'," & nMonto & ", '" & psCtaIf & "')"
        'dbCmact.Execute sSql
        'sSql = "Update DocRec Set nEstado = " & gChqEstEnValorizacion & " Where cNroDoc = '" & sNroDoc & "' And " _
        '    & "cIFTPo = '" & Format$(gTpoIFBanco, "00") & "' And cPersCod = '" & sCodIF & "' and CIFCTA='" & psCtaIf & "'"
        'dbCmact.Execute sSql
        'sSql = "INSERT DocRecEst (nTpoDoc,cNroDoc,cPersCod,cIFTpo,cMovNro,nEstado,CIFCTA ) " _
        '    & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "','" & sMovNro & "'," & gChqEstEnValorizacion & ",'" & psCtaIf & "')"
        'dbCmact.Execute sSql
        'sSql = "exec stp_ins_RFC1012012_AgregaCuentaDocumento " & nTpoDoc & ",'" & sNroDoc & "', '" _
                                                                & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "', '" _
                                                                & sMovNro & "','" & psCtaIf & "'," _
                                                                & gChqEstEnValorizacion & "," & nMonto & ", '" _
                                                                & sCuenta & "'," & nMovNro & ",'" & sFecha & "'"
        'EJVG20140408 ***
        sSql = "INSERT DocRecCapta (nTpoDoc,cNroDoc,cPersCod,cIFTpo,nMovNro,cCtaCod,nMonto,cIFCta) " _
            & " VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format(psIFTpo, "00") & "'," & nMovNro & ",'" & sCuenta & "'," & nMonto & ", '" & psCtaIf & "')"
        'END EJVG *******
         dbCmact.Execute sSql
         '***Fin Modificado por ELRO el 20121025*******************
    Case TpoDocOrdenPago
        sSql = "INSERT DocRecOP (nTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers,nEstado,cMovNro) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCuenta & "'," & nMonto & ",'" & sCodIF & "'," & nEstadoOP & ",'" & sMovNro & "')"
        dbCmact.Execute sSql
End Select
'Modificado PASI20140529
''***Modificado por ELRO el 20121025, según OYP-RFC101-2012
''sSql = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
''    & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
''dbCmact.Execute sSql
'If nTpoDoc <> TpoDocCheque Then
'    sSql = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
'    & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
'    dbCmact.Execute sSql
'End If
''***Fin Modificado por ELRO el 20121025*******************
If pbvalor = True Then
    sSql = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
        & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
    dbCmact.Execute sSql
End If
'end PASI
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

'ALPA 20081003
'Public Sub AgregaMovLavDinero(ByVal nMovNro As Long, ByVal sPersCod As String, Optional sPersCodTit As String = "", Optional sPersCodOrd As String = "", Optional sPersCodRea As String = "", Optional sPersCodBen As String = "", Optional sPersCodVisto As String = "")
'EJVG20120327
'JACA 20110224
'Public Sub AgregaMovLavDinero(ByVal nMovNro As Long, ByVal sPersCod As String, Optional sPersCodTit As String = "", Optional sPersCodOrd As String = "", Optional sPersCodRea As String = "", Optional sPersCodBen As String = "", Optional sPersCodVisto As String = "", Optional nTipoREU As Integer = 1, Optional nMontoAcu As Double = 0, Optional sOrigen As String = "", Optional sPersCodBen2 As String = "", Optional sPersCodBen3 As String = "", Optional sPersCodBen4 As String = "")
'Public Sub AgregaMovLavDinero(ByVal nMovNro As Long, ByVal sPersCod As String, Optional sPersCodTit As String = "", Optional sPersCodOrd As String = "", Optional sPersCodRea As String = "", Optional sPersCodBen As String = "", Optional sPersCodVisto As String = "", Optional nTipoREU As Integer = 1, Optional nMontoAcu As Double = 0, Optional sOrigen As String = "")
'JACA END
Public Sub AgregaMovLavDinero(ByVal nMovNro As Long, Optional sPersCodTit As String = "", Optional sPersCodVisto As String = "", Optional nTipoREU As Integer = 1, Optional nMontoAcu As Double = 0, Optional sOrigen As String = "")
'By Capi 20012008
'sSql = "INSERT MovLavDinero (nMovNro,cPersCod,cPersCodBen,cPersCodOrd,cPersCodRea,cPersCodVisto) " _
'    & "VALUES (" & nMovNro & ",'" & sPersCod & "','" & IIf(sPersCodBen = "", sPersCod, sPersCodBen) & "','" & sPersCodOrd & "','" & sPersCodRea & "','" & sPersCodVisto & "')"
    
'JACA 20110224
'sSql = "INSERT MovLavDinero (nMovNro,cPersCod,cPersCodOrd,cPersCodRea,cPersCodBen,cPersCodVisto,nTipoReu,nMontoAcum,cOrigenEfectivo) " _
'    & "VALUES (" & nMovNro & ",'" & sPersCodTit & "','" & sPersCodOrd & "','" & sPersCodRea & "','" & sPersCodBen & "','" & sPersCodVisto & "'," & nTipoREU & ", " & nMontoAcu & ",'" & sOrigen & "')"
'sSql = "INSERT MovLavDinero (nMovNro,cPersCod,cPersCodOrd,cPersCodRea,cPersCodBen,cPersCodVisto,nTipoReu,nMontoAcum,cOrigenEfectivo,cPersCodBen2,cPersCodBen3,cPersCodBen4) " _
'    & "VALUES (" & nMovNro & ",'" & sPersCodTit & "','" & sPersCodOrd & "','" & sPersCodRea & "','" & sPersCodBen & "','" & sPersCodVisto & "'," & nTipoREU & ", " & nMontoAcu & ",'" & sOrigen & "','" & sPersCodBen2 & "','" & sPersCodBen3 & "','" & sPersCodBen4 & "')"
'JACA END

'dbCmact.Execute sSql
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    sSql = "Exec stp_ins_MovLavDinero " & nMovNro & ",'" & sPersCodTit & "','" & sPersCodVisto & "'," & nTipoREU & "," & nMontoAcu & ",'" & sOrigen & "'"
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
'END EJVG
End Sub


Public Sub AgregaMovTipoCambio(ByVal nMovNro As Long, ByVal pnTipoCambio As Currency)
'***Modificado por ELRO el 20121119, según OYP-RFC101-2012
'sSql = " INSERT MovTpoCambio (nMovNro, nMovTpoCambio) " _
'     & " VALUES (" & nMovNro & "," & pnTipoCambio & ")"
sSql = "exec stp_ins_RFC1012012_MovTpoCambio " & nMovNro & ", " & pnTipoCambio & ""
'***Modificado por ELRO el 20121119***********************
dbCmact.Execute sSql
End Sub

'RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
Public Sub AgregaMovCap(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double, _
        Optional ByVal pnOpeTarj As Integer, Optional ByVal psNumTarj As String, Optional ByVal psRegla As String)
Dim oCon As COMDConstSistema.FCOMITF
Set oCon = New COMDConstSistema.FCOMITF
    If Mid(nOperacion, 1, 2) = "99" Then
        nMonto = oCon.fgTruncar(nMonto, 2)
    End If
Set oCon = Nothing

sSql = "INSERT MovCap (nMovNro,cOpeCod,cCtaCod,nMonto,nSaldoDisponible,nSaldoContable,nOpeTarj,cNumTarjeta, cRegla) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nMonto & "," & nSaldoDisp & "," & nSaldoCnt & "," & pnOpeTarj & ",'" & psNumTarj & "', '" & psRegla & "')"
dbCmact.Execute sSql
End Sub

Public Sub AgregaMovCapDet(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nConcepto As CaptacConcepto, ByVal nMonto As Double)
Dim oCon As COMDConstSistema.FCOMITF
Set oCon = New COMDConstSistema.FCOMITF
    If Mid(nOperacion, 1, 2) = "99" Then
        nMonto = oCon.fgTruncar(nMonto, 2)
    End If
Set oCon = Nothing
sSql = "INSERT MovCapDet (nMovNro,cOpeCod,cCtaCod,nConceptoCod,nMonto) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nConcepto & "," & nMonto & ")"
dbCmact.Execute sSql
End Sub

Public Sub AgregaCtaExoneradaITF(ByVal sCuenta As String, ByVal TpoExoneracionITF As Integer, ByVal cMovnro As String)
        
sSql = "Insert ITFExoneracionCta(cCtaCod, nExoTpo, cRegistro, cExtorno) " _
    & "VALUES ('" & sCuenta & "'," & TpoExoneracionITF & ",'" & cMovnro & "',Null)"
dbCmact.Execute sSql
        
sSql = "Insert ITFExoneracionCtaDet(cMovNro, cCtaCod, nExoTpo) " _
    & "VALUES ('" & cMovnro & "','" & sCuenta & "'," & TpoExoneracionITF & ")"
dbCmact.Execute sSql

End Sub

Public Sub ActualizaMovPendientesRend(ByVal nMovNro As Long, ByVal pnMonto As Currency)
    sSql = " UPDATE MovPendientesRend " _
         & " Set nSaldo = nSaldo - " & pnMonto & "" _
         & " Where nMovNro = " & nMovNro & ""
    dbCmact.Execute sSql
End Sub

'RIRO20140610 ERS017 ************
Public Sub ActualizaMovPendientesOpe(ByVal nMovNroOpe As Long, ByVal pnMonto As Currency)
    sSql = "STP_UPD_ACTUALIZAPENDOPE " & nMovNroOpe & ", " & pnMonto
    dbCmact.Execute sSql
End Sub
'END RIRO ***********************

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
'***Modificado por ELRO el 20121029, según OYP-RFC101-2012
'sSql = "INSERT MovRef (nMovNro,nMovNroRef) " _
'    & "VALUES (" & nMovNro & "," & nMovRef & ")"
sSql = "exec stp_ins_RegistrarMovRef " & nMovNro & "," & nMovRef & ""
'***Fin Modificado por ELRO el 20121029
dbCmact.Execute sSql
End Sub

Public Function GetEstadoCuenta(ByVal sCuenta As String) As CaptacEstado
    Dim rsEst As ADODB.Recordset
    Set rsEst = New ADODB.Recordset
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

Public Function GetCuentaBloqueos(ByVal sCuenta As String, ByVal nTipoBloqueo As CaptacTipoBloqueo) As ADODB.Recordset
    Dim rsEst As ADODB.Recordset
    Set rsEst = New ADODB.Recordset
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

Public Sub ActualizaSaldoRetiroCTS(ByVal sCuenta As String, ByVal nSaldoRet As Double, ByVal nIntSaldo As Double, Optional ByVal nDU01 As Double)
'sSql = "Update CaptacCTS Set nSaldRetiro = " & nSaldoRet & ", nIntSaldo = " & nIntSaldo & " " _
'    & "WHERE cCtaCod = '" & sCuenta & "'"
'sSql = "exec stp_upd_ActualizaSaldoRetiroCTS '" & sCuenta & "'," & nSaldoRet & "," & nIntSaldo   'JUEZ 20130724
sSql = "exec stp_upd_ActualizaSaldoRetiroCTS '" & sCuenta & "'," & nSaldoRet & "," & nIntSaldo & "," & nDU01 'APRI20200415 POR COVID-19
dbCmact.Execute sSql
End Sub

Public Sub ActualizaSaldoAnteriorAho(ByVal sCuenta As String, ByVal nSaldoAnt As Double)
    sSql = "Update CaptacAhorros Set nSaldoAnterior = " & nSaldoAnt & "  WHERE cCtaCod = '" & sCuenta & "'"
    dbCmact.Execute sSql
End Sub

Public Sub ActualizadUltContactoAho(ByVal sCuenta As String, ByVal dFecha As Date)
    sSql = "Update CaptacAhorros Set dUltContacto = '" & Format(dFecha, oConstante.gsFormatoFecha) & "'  WHERE cCtaCod = '" & sCuenta & "'"
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

Public Sub ActualizaAumentoCapitalPF(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, ByVal dRenovacion As Date, ByVal nPlazo As Long, _
            ByVal nTasaNueva As Double, ByVal dAuxiliar As Date, _
            Optional ByVal bActExtracto As Boolean = True, Optional ByVal nTpoPrograma As Integer = -1, Optional ByVal pbExtorno As Boolean = False)
            'APRI20171227 TIC1712200001 -> ADD pbExtorno
            
'*** Modificado por BRGO 20111228 El interés calculado se guarda como Interés Acumulado
If bActExtracto Then
    sSql = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & ", nTransacc = nTransacc + 1, " _
        & "nTasaInteres = " & nTasaNueva & " WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSql = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & ", " _
        & "nTasaInteres = " & nTasaNueva & " WHERE cCtaCod = '" & sCuenta & "'"
End If
dbCmact.Execute sSql

'APRI20171227 TIC1712200001
If pbExtorno Then
    sSql = "Update Captaciones Set nSaldoDisp = nSaldoDisp + " & nMontoDisp & ", nIntAcum = " & nIntGanado & ", " _
    & "dUltCierre = '" & Format$(dUltMov, "yyyy/mm/dd") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSql = "Update Captaciones Set nSaldoDisp = nSaldoDisp + " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado & ", " _
    & "dUltCierre = '" & Format$(dUltMov, "yyyy/mm/dd") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
End If
'END APRI

'sSql = "Update Captaciones Set nSaldoDisp = nSaldoDisp + " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado & ", " _
'    & "dUltCierre = '" & Format$(dUltMov, "yyyy/mm/dd") & "', cUltimaActualizacion = '" & sMovNro & "' " _
'    & "WHERE cCtaCod = '" & sCuenta & "'"

dbCmact.Execute sSql

sSql = "Update CaptacPlazoFijo Set nPlazo = " & nPlazo & ", nIntPag = 0, nIntDevChq = 0 "
If nTpoPrograma = 0 Then
    sSql = sSql & ", dRenovacion = '" & Format$(dRenovacion, "yyyy/mm/dd") & "',  "
    sSql = sSql & " dAuxiliar = '" & Format$(dAuxiliar, "yyyy/mm/dd") & "' "
End If
sSql = sSql & " WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaDisminucionCapitalPF(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, ByVal dRenovacion As Date, ByVal nPlazo As Long, _
            ByVal nTasaNueva As Double, ByVal dAuxiliar As Date, _
            Optional ByVal bActExtracto As Boolean = True)
If bActExtracto Then
    sSql = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt + nIntGanado & ", nTransacc = nTransacc + 1, " _
        & "nTasaInteres = " & nTasaNueva & " WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSql = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt + nIntGanado & ", " _
        & "nTasaInteres = " & nTasaNueva & " WHERE cCtaCod = '" & sCuenta & "'"
End If
dbCmact.Execute sSql
sSql = "Update Captaciones Set nSaldoDisp = nSaldoDisp - " & nMontoDisp & ", nIntAcum = 0, " _
    & "dUltCierre = '" & Format$(dUltMov, "yyyy/mm/dd") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql

sSql = "Update CaptacPlazoFijo Set nPlazo = " & nPlazo & ", nIntPag = 0, " _
    & "dRenovacion = '" & Format$(dRenovacion, "yyyy/mm/dd") & "', " _
    & "dAuxiliar = '" & Format$(dAuxiliar, "yyyy/mm/dd") & "' WHERE cCtaCod = '" & sCuenta & "'"
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
    '***Agregado por ELRO el 20130118, según TI-ERS031-2013****
    'sSql = "Update CaptacCTS Set nSaldRetiro = nSaldRetiro - " & nMonto & ", " _
    '    & "nIntSaldo = nIntSaldo + " & nIntSaldo & " WHERE cCtaCod = '" & sCuenta & "'"
    sSql = "exec stp_sel_ERS0312013_ActualizaCargoCTS '" & sCuenta & "', " & nMonto & ", " & nIntSaldo & ""
    '***Fin Agregado por ELRO el 20130118, según TI-ERS031-2013
    dbCmact.Execute sSql
End Sub



Public Sub ActualizaRetiroInteresPF(ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nIntGanado As Double, ByVal dUltMov As Date, ByVal sMovNro As String, _
            ByVal dAuxiliar As Date, Optional bActExtracto As Boolean = True, _
            Optional ByVal bAdelantado As Boolean = False)
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
    Dim rsChq As ADODB.Recordset
    sSql = "Select C.cCtaCod, D.nTpoDoc, D.cNroDoc, D.cPersCod, P.cPersNombre " _
        & "FROM Captaciones C INNER JOIN DocRecCapta D INNER JOIN DocRecEst E " _
        & "INNER JOIN " & sDBPersona & " Persona P ON E.cPersCod = P.cPersCod ON " _
        & "D.nTpoDoc = E.nTpoDoc AND D.cNroDoc = E.cNroDoc AND D.cPersCod = " _
        & "E.cPersCod AND D.cIFTpo = E.cIFTpo ON C.cCtaCod = D.cCtaCod WHERE C.cCtaCod = '" & sCuenta & "' " _
        & "AND E.cMovNro = (Select MAX(cMovNro) FROM DocRecEst E1 WHERE E1.nTpoDoc " _
        & "= E.nTpoDoc AND E1.cNroDoc = E.cNroDoc AND E1.cPersCod = E.cPersCod) AND " _
        & "E.nEstado = " & gChqEstEnValorizacion
    Set rsChq = New ADODB.Recordset
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
Dim rs As New ADODB.Recordset
Dim lscIFcta As String
'elimina docreccapta
sSql = " delete docreccapta  WHERE " _
    & "nTpoDoc = '" & nTipoDoc & "' AND cNroDoc = '" & sNroDoc & "' AND cPersCod = '" & sCodIF & "' and cctacod='" & sCuenta & "'"
dbCmact.Execute sSql
'Actualiza el ultimo estado a la tabla de Documento Recibidos
If nEstado <> 5 Then 'PASI20140611
    sSql = "Update DocRec Set nEstado = " & nEstado & " WHERE " _
        & "nTpoDoc = '" & nTipoDoc & "' AND cNroDoc = '" & sNroDoc & "' AND cPersCod = '" & sCodIF & "'"
    dbCmact.Execute sSql
End If
'Obtener Numero de CIFCTA

sSql = " select cIFcta From DocRec WHERE nTpoDoc = '" & nTipoDoc & "' AND cNroDoc = '" & sNroDoc & "' AND cPersCod = '" & sCodIF & "' "
rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
If Not (rs.EOF And rs.BOF) Then
    lscIFcta = rs!cIFcta
End If
rs.Close
'Agrega un registro mas de estado a la historia de estados del documento recibido
If nEstado <> 5 Then 'PASI20140611 then
    sSql = " INSERT DocRecEst (nTpoDoc,cNroDoc,cPersCod,cMovNro,nEstado,cIFTpo,cIFCta) " _
        & " VALUES ('" & nTipoDoc & "','" & sNroDoc & "','" & sCodIF & "','" & sMovNro & "'," & nEstado & ",'" & Format$(nIFTpo, "00") & "','" & lscIFcta & "')"
    dbCmact.Execute sSql
End If
End Sub

Public Function GetMovExtorno(ByVal sDatoBus As String, ByVal dFecha As Date, _
        ByVal nOperacion As CaptacOperacion, psAgeCod As String, Optional nTipoBus As Integer = 0) As ADODB.Recordset
Dim rsMov As ADODB.Recordset
Dim sWhere As String
'***Modificado por ELRO el 20121106, según OYP-RFC101-2012
'If nTipoBus = 0 Then
'    sWhere = "M.cMovNro LIKE '%" & sDatoBus & "%'"
'ElseIf nTipoBus = 1 Then
'    sWhere = "M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And C.cCtaCod = '" & sDatoBus & "'"
'ElseIf nTipoBus = 2 Then
'    sWhere = "M.nMovNro LIKE '%" & sDatoBus & "%'"
'End If
''If nOperacion = 200307 Then
'' sSql = "Select Distinct M.cMovNro, O.cOpeDesc, ('') Cuenta, ISNULL(MD.nDocTpo,'') nDocTpo, " _
''        & " ISNULL(MD.cDocNro,'') cDocNro,  M.cMovDesc, M.cOpeCod, nMovImporte, M.nMovNro," _
''        & " ('') cPersCod,(0) ITFCargo,(0) ITFOperacion,('')ITFConcepto,(0) RetOtraAge,(0) RetxMaxOpe" _
''        & "  FROM (Select MD.nMovNro, MD.nDocTpo, MD.cDocNro" _
''        & " From DocRecop D" _
''        & " RIGHT JOIN MovDoc MD ON d.nTpoDoc = MD.nDocTpo And D.cNroDoc = MD.cDocNro) MD" _
''        & " JOIN Mov M  ON MD.nMovNro = M.nMovNro " _
''        & " JOIN MovOpeVarias V  ON M.nMovNro = V.nMovNro" _
''        & " INNER JOIN OpeTpo O ON O.cOpeCod = V.cOpeCod" _
''        & " WHERE M.cMovNro LIKE '20060913%' AND M.cOpeCod  IN ('200307')" _
''        & " AND M.nMovFlag NOT IN (2,1,3)AND M.nMovEstado IN (10)" _
''        & " And SubString(M.cMovNro,18,2) = '01' ORDER BY M.nMovNro"
''Else
'sSql = "Select Distinct M.cMovNro, O.cOpeDesc, C.cCtaCod, ISNULL(MD.nDocTpo,'') nDocTpo, ISNULL(MD.cDocNro,'') cDocNro, " _
'    & " M.cMovDesc, C.cOpeCod, C.nMonto, M.nMovNro, cPersCod=isnull((select cperscod from docreccapta where cctacod=c.cctacod and nmovnro=c.nmovnro ),0), " _
'    & " IsNull(CD.nMonto,0) ITFCargo, ISNULL(CD.cOpeCod,'') ITFOperacion, ISNULL(CD.nConceptoCod,'') ITFConcepto, IsNull(CD1.nMonto,0) RetOtraAge,IsNull(CD2.nMonto,0) RetxMaxOpe, IsNull(CD3.nMonto,0) DepOtraAge FROM " _
'    & "(Select MD.nMovNro, MD.nDocTpo, MD.cDocNro From DocRecop D " _
'    & "        RIGHT JOIN MovDoc MD ON d.nTpoDoc = MD.nDocTpo And D.cNroDoc = MD.cDocNro) MD RIGHT JOIN Mov M JOIN MovCap C INNER JOIN OpeTpo O ON C.cOpeCod = O.cOpeCod ON M.nMovNro = C.nMovNro " _
'    & "ON MD.nMovNro = M.nMovNro Left Join MovCapDet CD On C.nMovNro = CD.nMovNro And C.cCtaCod = CD.cCtaCod And " _
'    & "CD.cOpeCod IN ('" & gITFCobroCargo & "','" & gITFCobroEfectivo & "','" & gITFCobroCMAC & "','" & gITFCobroCMACCargo & "','" & gITFCobroCargoPF & "') And CD.nConceptoCod IN " _
'    & "(" & gConcITFCliente & "," & gConcITFAsumido & ") " _
'    & " left Join MovCapDet CD1 On C.nMovNro = CD1.nMovNro And C.cCtaCod = CD1.cCtaCod And CD1.cOpeCod = " & gAhoRetComision & "" _
'    & " And CD1.nConceptoCod = " & gConcOpeRetOtraAge & "" _
'    & " Left Join MovCapDet CD2 On C.nMovNro = CD2.nMovNro And C.cCtaCod = CD2.cCtaCod And CD2.cOpeCod = " & gAhoRetComision & "" _
'    & " And CD2.nConceptoCod = " & gConcOpeRetxMaxOpe & "" _
'    & " Left Join MovCapDet CD3 On C.nMovNro = CD3.nMovNro And C.cCtaCod = CD3.cCtaCod And CD3.cOpeCod = " & gAhoDepComision & "" _
'    & " And CD3.nConceptoCod = " & gConcOpeDepOtraAge & "" _
'    & " WHERE " & sWhere
'
'    If nOperacion = gAhoTransAbonoL Then
'        sSql = sSql & "AND C.cOpeCod  IN ('" & nOperacion & "','" & gAhoTransCargoL & "') "
'    Else
'        sSql = sSql & "AND C.cOpeCod  IN ('" & nOperacion & "') "
'    End If
'
'    sSql = sSql & "AND M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & "," & gMovFlagDeExtorno & ")" _
'    & "AND M.nMovEstado IN (" & gMovEstContabMovContable & ") And SubString(M.cMovNro,18,2) = '" & psAgeCod & "'" _
'    & "ORDER BY M.nMovNro"
'
'    '***Modificado por ELRO el 20120312, según Acta N° 252-2011/TI-D para agregar la variable gITFCobroCargoPF de la Apertura en Efectivo de un Plazo Fijo
''End If
sSql = "exec stp_sel_RFC1012012_DevolverMovimientoExtornar '" & sDatoBus & "', '" & Format$(dFecha, "yyyymmdd") & "', '" & CStr(nOperacion) & "', '" & psAgeCod & "', " & nTipoBus & " "
'***Fin Modificado por ELRO el 20121106*******************
Set rsMov = New ADODB.Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetMovExtorno = rsMov
End Function

Public Function GetMovExtornoGiro(ByVal sDatoBus As String, ByVal dFecha As Date, _
        ByVal nOperacion As CaptacOperacion, psAgeCod As String, Optional nTipoBus As Integer = 0) As ADODB.Recordset
Dim rsMov As ADODB.Recordset
Dim sWhere As String
'If nTipoBus = 0 Then
'    '***Modificado por ELRO el 20110920, segun Acta 261-2011/TI-D
'    'sWhere = "M.cMovNro LIKE '%" & sDatoBus & "%'" 'comentado por ELRO el 20110920
'    sWhere = "M.cMovNro LIKE '" & sDatoBus & "%'"
'    '***Fin Modificado por ELRO
'ElseIf nTipoBus = 1 Then
'    sWhere = "M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' And C.cNumRecibo = '" & sDatoBus & "'"
'End If
'sSql = "Select distinct M.cMovNro, O.cOpeDesc, C.cNumRecibo cCtaCod, nDocTpo = '', cDocNro = '', M.cMovDesc, M.cOpeCod, "
'sSql = sSql & " C.nMonto, M.nMovNro, PP.cPersCod cPersCod "
'
''Agregado P
'sSql = sSql & ", IsNull(CD.nMonto,0) ITFCargo, ISNULL(CD.cOpeCod,'') ITFOperacion, ISNULL(CD.nPrdConcepto,'') ITFConcepto "
'' Fin Agregado
'
'sSql = sSql & " FROM OpeTpo O INNER JOIN Mov M JOIN MovServicios C "
'sSql = sSql & " JOIN ProductoPersona PP ON C.cNumRecibo = PP.cCtaCod ON M.nMovNro = C.nMovNro ON O.cOpeCod = M.cOpeCod "
'
''Agregado P
'sSql = sSql & " Left Join MovServiciosDet CD On   CD.nMovNro=C.nMovNro  And  CD.cNumRecibo=C.cNumRecibo  "
'sSql = sSql & " And CD.cOpeCod IN ('990101','990102','990302') "
'sSql = sSql & " And CD.nPrdConcepto IN (20,21) "
'' Fin Agregado
''***Modificado por ELRO el 20110919, segun Acta 261-2011/TI-D
'If nOperacion = gServGiroApertEfec Then
'sSql = sSql & " Left Join Producto P On PP.cCtaCod =P.cCtaCod "
'End If
''***Fin Modificado por ELRO
'sSql = sSql & " WHERE " & sWhere & " AND M.cOpeCod ='" & nOperacion & "' and C.copecod not like '99%'  AND PP.nPrdPersRelac = " & gGiroRelPersRemitente & " "
'sSql = sSql & " AND M.nMovFlag=" & gMovFlagVigente
'sSql = sSql & " AND M.nMovEstado=" & gMovEstContabMovContable & "  And SubString(M.cMovNro,18,2) = '" & psAgeCod & "' "
''***Modificado por ELRO el 20110919, segun Acta 261-2011/TI-D
'If nOperacion = gServGiroApertEfec Then
'sSql = sSql & " And P.nPrdEstado =1000"
'End If
'sSql = sSql & " AND O.cOpeCod ='" & nOperacion & "'"
''***Fin Modificado por ELRO
'sSql = sSql & " ORDER BY M.nMovNro "
'Comentado by NAGL 20181030
sSql = "Exec stp_sel_GetMovExtornoGiro '" & sDatoBus & "', '" & Format(dFecha, "yyyymmdd") & "', '" & nOperacion & "','" & psAgeCod & "'," & nTipoBus & "" 'Agregado by NAGL Según RFC1807260001
Set rsMov = New ADODB.Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetMovExtornoGiro = rsMov
End Function

Public Function GetITFMOV(ByVal nMovNroBus As Long) As Double
Dim rstemp As ADODB.Recordset
    
    sSql = "Select nmonto from movcapdet where nconceptocod=20 and nmovnro=" & nMovNroBus
    Set rstemp = New ADODB.Recordset
    rstemp.CursorLocation = adUseClient
    rstemp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rstemp.ActiveConnection = Nothing
   GetITFMOV = IIf(Not rstemp.EOF, rstemp!nMonto, 0)
End Function

Public Function GetIntaRetMov(ByVal nMovNroBus As Long) As Double
Dim rstemp As ADODB.Recordset
    
    sSql = "Select nmonto from movcapdet where nconceptocod=30 and nmovnro=" & nMovNroBus
    Set rstemp = New ADODB.Recordset
    rstemp.CursorLocation = adUseClient
    rstemp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rstemp.ActiveConnection = Nothing
   GetIntaRetMov = IIf(Not rstemp.EOF, rstemp!nMonto, 0)

End Function

Public Function GetMovExtornoCMAC(ByVal sDatoBus As String, ByVal dFecha As Date, _
        ByVal nOperacion As CaptacOperacion, Optional nTipoBus As Integer = 0) As ADODB.Recordset
Dim rsMov As ADODB.Recordset
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
Set rsMov = New ADODB.Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetMovExtornoCMAC = rsMov
End Function
Public Sub ActualizaEstadoMov(ByVal nMovNro As Long, ByVal nMovFlag As MovFlag)
sSql = "Update Mov Set nMovFlag = " & nMovFlag & " Where nMovNro = " & nMovNro
dbCmact.Execute sSql
End Sub
Public Function MovAsociados(ByVal nMovNro As Long) As ADODB.Recordset
Dim rstemp As ADODB.Recordset
sSql = "SELECT NMOVNROREF FROM MOVREF WHERE NMOVNRO=" & nMovNro
Set rstemp = New ADODB.Recordset
rstemp.CursorLocation = adUseClient
rstemp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rstemp.ActiveConnection = Nothing
Set MovAsociados = rstemp
Set rstemp = Nothing
End Function

Public Sub ActualizaDuplicadoCertPF(ByVal sCuenta As String)
sSql = "Update CaptacPlazoFijo Set nDuplicado = nDuplicado + 1 Where cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Function GetMovTransferencia(ByVal nMovNro As Long, Optional ByVal psCtaCod As String) As ADODB.Recordset
Dim rsTrans As ADODB.Recordset
sSql = "SELECT M.nMovNro, C.cOpeCod, C.cCtaCod, C.nMonto From Mov M INNER JOIN " _
    & "MovCap C ON M.nMovNro = C.nMovNro Where M.nMovNro = " & nMovNro & " and c.copecod not like '99%' "
If psCtaCod <> "" Then
    'sSql = sSql & "and Substring(C.cCtaCod,6,3) ='" & psCtaCod & "'" 'CTI5 ERS0112020 COMENTÓ
    sSql = sSql & "and C.cCtaCod = '" & psCtaCod & "'" 'CTI5 ERS0112020
End If
sSql = sSql & "ORDER BY M.cOpeCod DESC"
Set rsTrans = New ADODB.Recordset
rsTrans.CursorLocation = adUseClient
rsTrans.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTrans.ActiveConnection = Nothing
Set GetMovTransferencia = rsTrans
Set rsTrans = Nothing

End Function

Public Function GetMovTransferenciaITF(ByVal nMovNro As Long, ByVal sCuenta As String) As ADODB.Recordset
Dim rsTrans As ADODB.Recordset
sSql = " SELECT M.nMovNro, C.cOpeCod, C.cCtaCod, C.nMonto,c.nconceptocod From Mov M INNER JOIN " _
    & " MovCapdet C ON M.nMovNro = C.nMovNro Where M.nMovNro = " & nMovNro & " and c.CCTACOD='" & sCuenta & "' and c.copecod like '99%'"
    
Set rsTrans = New ADODB.Recordset
rsTrans.CursorLocation = adUseClient
rsTrans.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTrans.ActiveConnection = Nothing
Set GetMovTransferenciaITF = rsTrans
Set rsTrans = Nothing

End Function

Public Function GetFechaUltimoRetiroIntPF(ByVal sCuenta As String) As Date
Dim rsMov As ADODB.Recordset
sSql = "Select TOP 1 PF.dRenovacion, M.cMovNro FROM CaptacPlazoFijo PF LEFT JOIN " _
    & "MovCap C INNER JOIN Mov M ON C.nMovNro = M.nMovNro ON PF.cCtaCod = C.cCtaCod " _
    & "WHERE PF.cCtaCod = '" & sCuenta & "' OR (C.cOpeCod IN ('" & gPFRetInt & "','" & gPFRetIntAboAho & "') " _
    & "AND M.nMovFlag NOT IN ('" & gMovFlagExtornado & "') AND PF.cCtaCod = '" & sCuenta & "') " _
    & "ORDER BY M.nMovNro "
Set rsMov = New ADODB.Recordset
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
Dim rsMov As ADODB.Recordset
sSql = "Select TOP 2 PF.dRenovacion, M.cMovNro FROM CaptacPlazoFijo PF LEFT JOIN " _
    & "MovCap C INNER JOIN Mov M ON C.nMovNro = M.nMovNro ON PF.cCtaCod = C.cCtaCod " _
    & "WHERE PF.cCtaCod = '" & sCuenta & "' OR (C.cOpeCod IN ('" & gPFRetInt & "','" & gPFRetIntAboAho & "') " _
    & "AND M.nMovFlag NOT IN ('" & gMovFlagExtornado & "') AND PF.cCtaCod = '" & sCuenta & "') " _
    & "ORDER BY M.nMovNro "
Set rsMov = New ADODB.Recordset
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
Dim rsChq As ADODB.Recordset
sSql = "SELECT C.dValorizacion FROM DocRecCapta A INNER JOIN DocRec C ON " _
    & "A.nTpoDoc = C.nTpoDoc AND A.cNroDoc = C.cNroDoc AND A.cPersCod = C.cPersCod WHERE " _
    & "C.nEstado = " & gChqEstValorizado & " AND A.cCtaCod = '" & sCuenta & "'"
Set rsChq = New ADODB.Recordset
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
Dim rsCta As ADODB.Recordset
sSql = "SELECT cCtaCodAbono FROM CaptacAboIntPF WHERE cCtaCod = '" & sCuenta & "'"
Set rsCta = New ADODB.Recordset
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
Dim rsOP As ADODB.Recordset
Set rsOP = New ADODB.Recordset
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
    ByVal nProducto As Producto) As ADODB.Recordset
Dim rsMov As ADODB.Recordset
sSql = "Select M.cMovNro, MC.nMovNro, MC.cOpeCod, MC.cCtaCod, MC.nMonto, MC.nSaldoDisponible, " _
    & "MC.nSaldoContable, MD.nMonto nMontoConcepto FROM Mov M INNER JOIN MovCap MC INNER JOIN " _
    & "MovCapDet MD ON MC.nMovNro = MD.nMovNro AND MC.cCtaCod = MD.cCtaCod ON M.nMovNro = MC.nMovNro " _
    & "WHERE M.nMovNro = " & nMovNro & " AND MD.nConceptoCod = " & nConcepto & " AND " _
    & "SUBSTRING(MC.cCtaCod,6,3) = '" & nProducto & "'"
Set rsMov = New ADODB.Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetCapMov = rsMov
Set rsMov = Nothing
End Function

Public Function GetCapMovCMAC(ByVal nMovNro As Long) As ADODB.Recordset
Dim rsMov As ADODB.Recordset
sSql = "Select M.cMovNro, M.nMovNro, MC.cOpeCod, MC.cCtaIF cCtaCod, MC.nMonto, " _
    & "MC.nMoneda, MC.cDocumento FROM Mov M INNER JOIN MovCMAC MC ON " _
    & "M.nMovNro = MC.nMovNro WHERE M.nMovNro = " & nMovNro
Set rsMov = New ADODB.Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetCapMovCMAC = rsMov
Set rsMov = Nothing
End Function

Public Function GetCapMovParamCTS(ByVal nMovNro As Long) As Double
Dim rsCTS As ADODB.Recordset
sSql = "Select nPorcentajeAbono From MovParamCTS Where nMovNro = " & nMovNro
Set rsCTS = New ADODB.Recordset
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

Public Sub AgregaMovTranferBanco(ByVal pnMovNro As Long, ByVal psPersCodIF As String, psCtaBco As String, ByVal pnMoneda As Moneda, Optional ByVal sTitular As String = "")
    'sSql = " Insert MovTransferBco (nMovNro, cPersCodIF, cCtaBco, nMoneda) Values (" & pnMovNro & ",'" & psPersCodIF & "','" & psCtaBco & "'," & pnMoneda & ")" RIRO20131212 ERS137 Comentado
    sSql = " Insert MovTransferBco (nMovNro, cPersCodIF, cCtaBco, nMoneda, cTitular) Values (" & pnMovNro & ",'" & psPersCodIF & "','" & psCtaBco & "'," & pnMoneda & ", '" & sTitular & "')" ' RIRO20131212 ERS137 Agregado
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
IniciaImpresora
End Sub
Public Sub IniciaImpresora(Optional ByVal nImpresora As COMFunciones.Impresoras = gEPSON)
Set oImp = New COMFunciones.FCOMVarImpresion
oImp.Inicia nImpresora
End Sub
Private Sub Class_Terminate()
If bTransaccion = True Then
Else
    dbCmact.Close
    Set dbCmact = Nothing
End If

Set oConstante = Nothing
Set oImp = Nothing
End Sub

Public Function CtaConFirmas(psCtaCod As String) As Boolean
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim clsCon As COMConecta.DCOMConecta
    Set clsCon = New COMConecta.DCOMConecta

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

Public Sub AgregaMovOpeVarias(ByVal pnMovNro As Long, ByVal pnMovImporte As Double, _
        ByVal psNroDoc As String, ByVal psReferencia As String, ByVal pnMoneda As Moneda, _
        Optional nOperacion As CaptacOperacion, _
        Optional ByVal pnConcepto As Integer = 0)
        'JUEZ 201541019 Se agregó pnConcepto
    Dim sql As String
    
    'JUEZ 20151019 *******************
    'sql = " Insert MovOpeVarias (nMovNro,nMovImporte,cNroDoc,cReferencia,nMoneda,cOpeCod)" _
    '    & " Values(" & pnMovNro & "," & pnMovImporte & ",'" & psNroDoc & "','" & psReferencia & "'," & pnMoneda & ",'" & Trim(nOperacion) & "')"
    sql = " Insert MovOpeVarias (nMovNro,nMovImporte,cNroDoc,cReferencia,nMoneda,cOpeCod,nConcepto)" _
        & " Values(" & pnMovNro & "," & pnMovImporte & ",'" & psNroDoc & "','" & psReferencia & "'," & pnMoneda & ",'" & Trim(nOperacion) & "'," & pnConcepto & ")"
    'END JUEZ ************************
    
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

Public Function GetOtrasOperaciones(pdFecha As Date, Optional cUser As String, Optional ByVal psOpeCod As String = "") As ADODB.Recordset
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim clsCon As COMConecta.DCOMConecta
    Set clsCon = New COMConecta.DCOMConecta
    
    clsCon.AbreConexion
    
    'RIRO20140530 ERS017 - Comentado
    'sql = " SELECT M.nMovNro, Convert(DateTime,Left(M.cMovNro,8)) dFecTran, M.cOpeCod, Right(M.cMovNro,4) cCodUsu, OP.cOpeDesc, MOV.nMovImporte, MOV.nMoneda " _
    '& " FROM Mov M JOIN MovOpeVarias MOV On M.nMovNro = MOV.nMovNro JOIN OpeTpo OP ON M.cOpeCod = OP.cOpeCod " _
    '& " LEFT JOIN CredComisiones CC ON M.nMovNro=CC.nMovNro " _
    '& " WHERE M.nMovFlag = 0 And (Left(M.cOpeCod,4) IN ('" & Left(gOtrOpeAhoOtrosIngresos, 4) & "', " _
    '& " '" & Left(gOtrOpeAhoOtrosEgresos, 4) & " ','" & Left(gOtrOpeDuplicadoTarjeta, 4) & "','" & Left(gOtrOpeDepCtaBcoEfec, 4) & "','" _
    '& Left(gComisionReprogCredito, 4) & "') or  M.cOpeCod='" & gAhoRetFondoFijo & "') " _
    '& " And M.cMovNro LIKE '" & Format(pdFecha, oConstante.gsFormatoMovFecha) & "%' and right(M.cMovNro,4)='" & cUser & "' And IsNull(CC.nUtilizado,0) = 0 "
    ''JUEZ 20130315 Se agregó tabla CredComisiones y el Cod. Operacion gComisionReprogCredito
    
    'RIRO20140530 ERS017 - Agregado
    'sql = "STP_SEL_OTRASOPERACIONES '" & Format(pdFecha, oConstante.gsFormatoMovFecha) & "', '" & cUser & "'" 'RECO20160408
    sql = "STP_SEL_OTRASOPERACIONES '" & Format(pdFecha, oConstante.gsFormatoMovFecha) & "', '" & cUser & "','" & psOpeCod & "'" 'RECO20160408
    
    Set rs = clsCon.CargaRecordSet(sql)
    
    Set GetOtrasOperaciones = rs
    
End Function

Public Function GetOtrasOperacionesDet(psMovNro As Long) As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim clsCon As COMConecta.DCOMConecta
    Set clsCon = New COMConecta.DCOMConecta
    
    clsCon.AbreConexion
    
    sSql = "exec STP_SEL_GetOtrasOperacionesDet " & psMovNro 'ORCR20140714
'    sSql = "SELECT M.nMovNro, Convert(DateTime,Left(M.cMovNro,8)) dFecTran, M.cOpeCod, Right(M.cMovNro,4) cCodUsu, OP.cOpeDesc, MOV.nMovImporte, PE.cPersCod , PE.cPersNombre, M.cMovDesc, MOV.cNroDoc " _
'        & "FROM Mov M JOIN MovOpeVarias MOV On M.nMovNro = MOV.nMovNro " _
'        & "JOIN MovGasto MG On M.nMovNro = MG.nMovNro JOIN Persona PE On MG.cPersCod = PE.cPersCod " _
'        & "JOIN OpeTpo OP ON M.cOpeCod = OP.cOpeCod WHERE M.nMovNro = " & psMovNro & " " _
'        & "UNION " _
'        & "SELECT M.nMovNro, Convert(DateTime,Left(M.cMovNro,8)) dFecTran, M.cOpeCod, Right(M.cMovNro,4) cCodUsu, OP.cOpeDesc, MOV.nMovImporte, PE.cPersCod , PE.cPersNombre, M.cMovDesc, MOV.cNroDoc " _
'        & "FROM Mov M JOIN MovOpeVarias MOV On M.nMovNro = MOV.nMovNro " _
'        & "JOIN Persona PE On MOV.cReferencia = PE.cPersCod OR SUBSTRING(MOV.cReferencia,4,13) = PE.cPersCod " _
'        & "JOIN OpeTpo OP ON M.cOpeCod = OP.cOpeCod WHERE M.nMovNro = " & psMovNro
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
    Dim clsCon As COMConecta.DCOMConecta
    Set clsCon = New COMConecta.DCOMConecta
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
    Dim clsCon As COMConecta.DCOMConecta
    Set clsCon = New COMConecta.DCOMConecta
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

Public Function BuscaGarantiaCreditosPlazoFijo(ByVal psCuenta As String) As Boolean
    Dim rsCta As ADODB.Recordset
    Set rsCta = New ADODB.Recordset
    Dim clsCon As COMConecta.DCOMConecta
    Set clsCon = New COMConecta.DCOMConecta
    Dim sql As String
    
    clsCon.AbreConexion
    
    sql = " select count(*) nGarantia from Garantias G " & _
          " inner join ColocGarantia CG on G.cNumGarant=CG.cNumGarant" & _
          " inner join Producto P on P.cCtacod=CG.cCtaCod and nPrdEstado in('2020','2021','2022','2030','2031','2032')" & _
          " where cNroDoc = '" & psCuenta & "'"
    Set rsCta = clsCon.CargaRecordSet(sql)
    
    If rsCta!nGarantia <> 0 Then
        BuscaGarantiaCreditosPlazoFijo = True
    Else
        BuscaGarantiaCreditosPlazoFijo = False
    End If
    'rsCta.Close
    Set rsCta = Nothing
    Set clsCon = Nothing
End Function



Public Function GetFirma(psPersCod As String, Optional ByVal psCodAge As String = "") As ADODB.Recordset
    Dim sSql As String
    Dim R As ADODB.Recordset
    Set R = New ADODB.Recordset
    
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    sSql = "Select cPersCod,iPersFirma,cUltimaActualizacion from " & sDBImagenes & "PersImagen PI Where cPersCod = '" & psPersCod & "'"
    'ande 20170914
    'Set R = New ADODB.Recordset
    'R.Open sSql, dbCmact, adOpenKeyset, adLockOptimistic
    'end ande
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set GetFirma = R
End Function

Public Function GetDatosCrePendiente(psTipo As String, pbVencidos As Boolean, psCuentas As String, pnDiasAtraso As Integer, pgdFecSis As Date) As ADODB.Recordset
    Dim sql As String
    Dim clsCon As COMConecta.DCOMConecta
    Dim sVencidos As String
    Set clsCon = New COMConecta.DCOMConecta
    
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
            & "     Where P1.cCtaCod = P.cCtaCod And P.nNroCalen = P1.nNroCalen And P1.nColocCalendEstado = " & ColocCalendEstado.gColocCalendEstadoPendiente & " And P1.nColocCalendApl = " & ColocCalendApl.gColocCalendAplCuota & " And DateDiff(dd,P1.dVenc,'" & Format$(pgdFecSis, oConstante.gsFormatoFecha) & "') > 0)" _
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
            & "   Where  " & sVencidos & " Left(PRD.nPrdEstado,3) = " & Left(gColocEstVigNorm, 3) & " And P.nColocCalendEstado = " & ColocCalendEstado.gColocCalendEstadoPendiente & " And P.nColocCalendApl = " & ColocCalendApl.gColocCalendAplCuota & " And C.cCtaCod In " & psCuentas & " And DateDiff(dd,P.dVenc,'" & Format$(pgdFecSis, oConstante.gsFormatoFecha) & "') > 0 ) T " _
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
Dim rsMov As ADODB.Recordset
sSql = "Select ISNULL(Count(*),0) nNumMov From Mov M JOIN MovCap C ON M.nMovNro = C.nMovNro " _
    & "Where M.nMovFlag = " & gMovFlagVigente & " And C.cCtaCod = '" & sCuenta & "' And C.cOpeCod Not In ('" & gITFCobroCargo & "','" & gITFCobroEfectivo & "','" & gITFCobroEfectivoAsumidoPF & "','" & gITFCobroCargoPF & "') " _
    & "And M.nMovEstado = " & gMovEstContabMovContable
'***Agregado por ELRO gITFCobroCargoPF
Set rsMov = New ADODB.Recordset
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
    Dim rsVar As ADODB.Recordset
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    ListaCtasCont = " "
    
    If oCon.AbreConexion = False Then Exit Function
    sSql = "Select nConsSisDesc, nConsSisValor From ConstSistema where nConsSisCod =" & psConstSistCod & ""
    Set rsVar = New ADODB.Recordset
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
    
    Dim oCon As COMDConstSistema.NCOMConstSistema
    Set oCon = New COMDConstSistema.NCOMConstSistema
    
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
     
    sql = " Select MOV.cCtaCod, MOV.nMonto, ISNull(ITF.nMonto,0) ITF,cOpeCodITF=ITF.cOpeCod From " _
        & " (Select cCtaCod, nMonto From movcap Where nMovNro = " & pnMovNro & " And cOpeCod Not Like '" & Left(gITFCobroCargo, 2) & "%') MOV" _
        & " Left Join (Select cCtaCod, nMonto,cOpeCod From movcap Where nMovNro = " & pnMovNro & " And cOpeCod Like '" & Left(gITFCobroCargo, 2) & "%') ITF On MOV.cCtaCod = ITF.cCtaCod"
    
    rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.EOF And rs.BOF Then
        Set GetApeLoteCuentas = rs
    Else
        Set GetApeLoteCuentas = rs
    End If
End Function

Public Function GetEstadoAhoPanderito(ByVal sCuenta As String, ByVal dFecha As Date) As String
    Dim rsEst As ADODB.Recordset
    Set rsEst = New ADODB.Recordset
    
    sSql = sSql & " select count(*)"
    sSql = sSql & " from Producto P"
    sSql = sSql & " inner join Captaciones C on P.cCtaCod=C.cCtaCod"
    sSql = sSql & " inner join CaptacAhorros CA on CA.cCtaCod=C.cCtacod"
    sSql = sSql & " inner join (select MC.cCtacod,Monto=sum(nMonto) from Mov M"
    sSql = sSql & "        inner join MovCap MC on M.nMovNro=MC.nMovNro"
    sSql = sSql & "        inner join CaptacAhorros C1 on C1.cCtaCod = MC.cCtacod"
    sSql = sSql & "        where nMovFlag=0 and MC.cOpeCod ='200201'"
    sSql = sSql & "        and datediff(d,dbo.FechaHoraMovDate(cMovNro),C1.dRenoPanderito)<=0 and datediff(d, dbo.FechaHoraMovDate(cMovNro),'" & Format(dFecha, "yyyy/mm/dd") & "')>=0"
    sSql = sSql & "        group by MC.cCtacod) MO ON MO.cCtacod=CA.cCtaCod"
    sSql = sSql & " where nprdestado not in(1400,1300) and C.nTpoPrograma=2 and substring(P.cCtaCod,6,3)='232'"
    sSql = sSql & " and (MO.Monto < (CA.nMontoAbono*CA.nPlazo)or MO.Monto >= (CA.nMontoAbono*CA.nPlazo) ) and datediff(d,dRenoPanderito,'" & Format(dFecha, "yyyy/mm/dd") & "')<>0"
    sSql = sSql & " and datediff(d,dRenoPanderito,'" & Format(dFecha, "yyyy/mm/dd") & "')-(select nFeriado=count(*)"
    sSql = sSql & "                                from Feriado where datediff(d,dFeriado,dRenoPanderito)<=0"
    sSql = sSql & "                                and datediff(d,dFeriado,'" & Format(dFecha, "yyyy/mm/dd") & "')>=0)=CA.nPlazo"
    sSql = sSql & " and CA.cCtaCod ='" & sCuenta & "'"

    rsEst.CursorLocation = adUseClient
    rsEst.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsEst.ActiveConnection = Nothing
    If Not (rsEst.EOF And rsEst.BOF) Then
        If rsEst(0) = 0 Then
            GetEstadoAhoPanderito = "No se puede Renovar, aún no vence el Plazo"
        Else
            GetEstadoAhoPanderito = ""
        End If
    End If
    rsEst.Close
    Set rsEst = Nothing
End Function

Public Sub ActualizaCaptaAhoPanderito(ByVal sCuenta As String, ByVal nPlazo As Integer, ByVal dFecha As Date)

sSql = "Update CaptacAhorros Set nPlazo =  " & nPlazo & ", dRenoPanderito = '" & Format(dFecha, "yyyy/mm/dd") & "' " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql

End Sub

Public Sub SetConexion(pCon As ADODB.Connection)
    Set dbCmact = pCon
End Sub

Public Function RecuperaFirma(ByVal psCodPers As String, ByVal psCodAge As String) As ADODB.Recordset

Dim sSql As String
Dim R As ADODB.Recordset
Dim sCadCon As String

'
'    ssql = "Select * from Servidor Where RIGHT(cCodAge,2) = '" & psCodAge & "' AND cNroSer = '05'"
'    Set R = New ADODB.Recordset
'    R.Open ssql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'    If R.RecordCount > 0 Then
'        sCadCon = "Provider=" & Trim(R!cProvider) & ";User ID=" & Trim(R!cLogin) & ";Initial Catalog=" & Trim(R!cDataBase) & ";Data Source=" & Trim(R!cNomSer) & ";PWD=" & Trim(R!cPass)
'    End If
'    R.Close
'
'    If sCadCon <> "" Then
'        Set oConecta = New COMConecta.DCOMConecta
'        Call oConecta.AbreConexion(sCadCon)
'        Set ConexionActiva = oConecta.ConexionActiva
'    End If
    sSql = "Select cPersCod,iPersFirma,cUltimaActualizacion from " & sDBImagenes & "PersImagen PI Where cPersCod = '" & psCodPers & "'"
    Set R = New ADODB.Recordset
    R.Open sSql, dbCmact, adOpenKeyset, adLockOptimistic
'    If R.BOF And R.EOF Then
'        R.Close
'        Set RecuperaFirma = Nothing
'        Exit Function
'    End If
    Set RecuperaFirma = R
End Function
'ALPA 20081007*****************************************************************
Public Sub AgregaMovWesterUnion(ByVal nMovNro As Long, ByVal sPersCod As String)
    
sSql = "INSERT MovWesterUnion (nMovNro,cPersWesterUnion) " _
    & "VALUES (" & nMovNro & ",'" & sPersCod & "')"
    
dbCmact.Execute sSql

End Sub
'****************************************************************************
'** BRGO 20110309 ***********************************************************
Public Sub AgregaMovWesterUnionRemitente(nMovNro As Long, sPersNombre As String, sPersCiudadPais As String, dFecha As Date)
    sSql = " INSERT MovWUnionDatRemite (nMovNro, cNombres, cCiudadPais, dFecha) " _
    & " VALUES (" & nMovNro & ",'" & sPersNombre & "','" & sPersCiudadPais & "','" & dFecha & "')"
    dbCmact.Execute sSql
End Sub
'****************************************************************************

'JACA 20110316==============================================================
Public Sub registrarCapAbonoAgeParam(ByVal sAgecod As String, ByVal nMontoMin As Double, ByVal nMontoMax As Double, ByVal sUltimaActualizacion As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_CapAbonoAgeParam '" & sAgecod & "'," & nMontoMin & "," & nMontoMax & ",'" & sUltimaActualizacion & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub
Public Sub registrarCapAbonoPersRea(ByVal nNroMov As Long, ByVal sPersCod As String, ByVal nPersDNI As String, ByVal sPerNombre As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_CapAbonoPersRea " & nNroMov & ",'" & sPersCod & "','" & nPersDNI & "','" & sPerNombre & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub
Public Sub modificaCapAbonoAgeParam()
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_upd_CapAbonoAgeParam"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub

Public Function obtenerCapAbonoAgeParamListar() As ADODB.Recordset
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "stp_sel_CapAbonoAgeParam_Lista"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set obtenerCapAbonoAgeParamListar = lrDatos
    Set lrDatos = Nothing
End Function
Public Function obtenerCapAbonoAgeParam(ByVal sCodage As String) As ADODB.Recordset
Dim lsSQL As String
Dim oConec As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "stp_sel_CapAbonoAgeParam '" & sCodage & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set obtenerCapAbonoAgeParam = lrDatos
    Set lrDatos = Nothing
End Function
'END DJACA==============================================================

'*******BRGO 20110405 **************************************************
Public Sub AgregaDatosSueldosClientesCTS(ByVal sMovNro As String, ByVal sCuenta As String, ByVal nMonedaSueldo As Integer, ByVal nSueldoTotal As Currency)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_CapSueldosClientesCTS '" & sMovNro & "','" & sCuenta & "'," & nMonedaSueldo & "," & nSueldoTotal
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub
Public Function ObtenerCapCuentasCTSInstitucion(ByVal sRUC As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
        lsSQL = "stp_sel_ObtieneCuentasCTSInstitucion '" & sRUC & "'"
        
        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set ObtenerCapCuentasCTSInstitucion = lrDatos
        Set lrDatos = Nothing
End Function
Public Function ObtenerCapSaldosCuentasCTS(ByVal sCuenta As String, nTipoCambio As Double) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
        lsSQL = "stp_sel_ObtieneSaldosYSueldosCTS '" & sCuenta & "'," & nTipoCambio
        
        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set ObtenerCapSaldosCuentasCTS = lrDatos
        Set lrDatos = Nothing
End Function

Public Function ObtenerFecUltimaActSueldosCTS(ByVal sCuenta As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
        lsSQL = "stp_sel_ObtieneFecUltimaActSueldosCTS '" & sCuenta & "'"
        
        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set ObtenerFecUltimaActSueldosCTS = lrDatos
        Set lrDatos = Nothing
End Function

'MADM 20111207
Public Sub ActualizarEstMovPagoServicios(ByVal pnDeuNro As Long)
    sSql = "Exec upd_PagoClientesVariosPagoMov " & pnDeuNro & ""
    dbCmact.Execute sSql
End Sub
'END MADM

'***Agregado por ELRO el 20120307, según Acta N° 039-2012/TI-D
Public Function GetTipoProgramaCaptacion(ByVal pcCtaCod As String) As Integer
    Dim sSql As String
    Dim oDCOMConecta  As COMConecta.DCOMConecta
    Set oDCOMConecta = New COMConecta.DCOMConecta
    Dim rsTipoPrograma As ADODB.Recordset
    Set rsTipoPrograma = New ADODB.Recordset
        
    oDCOMConecta.AbreConexion
    'sSql = "exec stp_sel_DevolverTipoProgramaCaptacion " & pcCtaCod
    sSql = "exec stp_sel_DevolverTipoProgramaCaptacion '" & pcCtaCod & "'" 'JUEZ 20151229
    Set rsTipoPrograma = oDCOMConecta.CargaRecordSet(sSql)
    If Not rsTipoPrograma.BOF And Not rsTipoPrograma.EOF Then
        GetTipoProgramaCaptacion = rsTipoPrograma!nTpoPrograma
    Else
        GetTipoProgramaCaptacion = -1
    End If

    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
    Set rsTipoPrograma = Nothing
End Function
'***Fin Agregado por ELRO*************************************
'***Agregado por ELRO el 20120327, según OYP-RFC023-2012
Public Sub registrarAfectacion(ByVal pnMovNro As Long, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
  
    lsSQL = "exec stp_ins_RegistrarAfectacion " & pnMovNro
    
    If pbEjecBatch = False Then
        dbCmact.Execute lsSQL
    End If
    
    lsSQL = ""
End Sub
Public Sub actualizarAfectacion(ByVal pnMovNro As Long, Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String

    lsSQL = "exec stp_upd_ActualizarAfectacion " & pnMovNro

    If pbEjecBatch = False Then
        dbCmact.Execute lsSQL
    End If
    lsSQL = ""
End Sub
'***Fin Agregado por ELRO*******************************
'***Agregado por ELRO el 20120417, según OYP-RFC005-2012
Public Function ObtenerDesembolsoAprobacionARendirViaticos(ByVal pnMovNro As Long, _
                                                           ByVal psUsuario As String, _
                                                           ByVal psFecha As String, _
                                                           ByVal psAgeCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim rsDesembolso As ADODB.Recordset
        
    lsSQL = "exec stp_sel_ObtenerDesembolsoAprobacionARendirViaticos " & pnMovNro & ", '" & psUsuario & "', '" & psFecha & "', '" & psAgeCod & "'"
    
    Set oDCOMConecta = New COMConecta.DCOMConecta
    If oDCOMConecta.AbreConexion = False Then Exit Function
    Set rsDesembolso = oDCOMConecta.CargaRecordSet(lsSQL)
    Set ObtenerDesembolsoAprobacionARendirViaticos = rsDesembolso
    Set rsDesembolso = Nothing
    Set oDCOMConecta = Nothing
End Function

Public Function ObtenerDesembolsoAprobacionARendirCuentas(ByVal pnMovNro As Long, _
                                                          ByVal psUsuario As String, _
                                                          ByVal psFecha As String, _
                                                          ByVal psAgeCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim rsDesembolso As ADODB.Recordset
        
    lsSQL = "exec stp_sel_ObtenerDesembolsoAprobacionARendirCuentas " & pnMovNro & ", '" & psUsuario & "', '" & psFecha & "', '" & psAgeCod & "'"
    
    Set oDCOMConecta = New COMConecta.DCOMConecta
    If oDCOMConecta.AbreConexion = False Then Exit Function
    Set rsDesembolso = oDCOMConecta.CargaRecordSet(lsSQL)
    Set ObtenerDesembolsoAprobacionARendirCuentas = rsDesembolso
    Set rsDesembolso = Nothing
    Set oDCOMConecta = Nothing
End Function
'***Fin Agregado por ELRO*******************************
'***Agregado por ELRO el 20120608, según OYP-RFC047-2012
Public Function ObtenerDesembolsoAprobacionCH(ByVal pnMovNro As Long, _
                                              ByVal psUsuario As String, _
                                              ByVal psFecha As String, _
                                              ByVal psAgeCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim rsDesembolso As ADODB.Recordset

    lsSQL = "exec stp_sel_ObtenerDesembolsoAprobacionCH " & pnMovNro & ", '" & psUsuario & "', '" & psFecha & "', '" & psAgeCod & "'"

    Set oDCOMConecta = New COMConecta.DCOMConecta
    If oDCOMConecta.AbreConexion = False Then Exit Function
    Set rsDesembolso = oDCOMConecta.CargaRecordSet(lsSQL)
    Set ObtenerDesembolsoAprobacionCH = rsDesembolso
    Set rsDesembolso = Nothing
    Set oDCOMConecta = Nothing
End Function
'***Fin Agregado por ELRO*******************************
'***Agregado por ELRO el 20120717, según OYP-RFC024-2012
Public Sub actualizarRelacionVoucherCaptacion(ByVal pnMovNroOpe As Long, _
                                              ByVal pnMovNroRVD As Long)
    Dim lsSQL As String
    lsSQL = "exec stp_upd_ActualizarCapVoucherDeposito_2  " & pnMovNroOpe & ", " & pnMovNroRVD & ""
    dbCmact.Execute lsSQL
End Sub

Public Sub actualizarRelacionExtornoVoucherCaptacion(ByVal pnMovNroOpe As Long)
    Dim lsSQL As String
    lsSQL = "exec stp_upd_ActualizarCapVoucherDeposito_3  " & pnMovNroOpe & ""
    dbCmact.Execute lsSQL
End Sub

'***Fin Agregado por ELRO*******************************
'***Agregado por ELRO el 20120911, según OYP-RFC087-2012
Public Sub devolverIndicadorCliente(ByVal psPersCod As String, _
                                    ByVal pnTipoCambio As Currency, _
                                    ByVal psFecha As String, _
                                    ByVal pnMontoApertura As Currency, _
                                    ByVal psAgeCod As String, _
                                    ByVal psTpoProducto As String, _
                                    ByVal pnTasaInteres As Currency, _
                                    ByVal psMov As String, _
                                    ByRef pnIndicadorInterno As Currency, _
                                    ByRef pnIndicadorPrimerCliente As Currency, _
                                    ByRef pnIndicadorDiezCliente As Currency, _
                                    ByRef pnIndicadorVeinteCliente As Currency, _
                                    ByRef pnSaldoTotalDepositos As Currency, _
                                    ByRef pnIdVoBoConCli As Long)
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim rsIndice As ADODB.Recordset

    lsSQL = "exec stp_sel_RFC0872012_DevolverIndicadorCliente '" & psPersCod & "', " & pnTipoCambio & ", '" & psFecha & " ', " & pnMontoApertura & ", '" & psAgeCod & "', '" & psTpoProducto & "', " & pnTasaInteres & ", '" & psMov & "'"

    Set oDCOMConecta = New COMConecta.DCOMConecta
    If oDCOMConecta.AbreConexion = False Then Exit Sub
    Set rsIndice = oDCOMConecta.CargaRecordSet(lsSQL)
    
    If Not rsIndice.BOF And Not rsIndice.EOF Then
        pnIndicadorInterno = rsIndice!nIndicadorInterno
        pnIndicadorPrimerCliente = rsIndice!nIndicadorPrimerCliente
        pnIndicadorDiezCliente = rsIndice!nIndicadorDiezCliente
        pnIndicadorVeinteCliente = rsIndice!nIndicadorVeinteCliente
        pnSaldoTotalDepositos = rsIndice!nSaldoTotalDepositos
        pnIdVoBoConCli = rsIndice!nIdVoBoConCli
    
    Else
        pnIndicadorInterno = 0#
        pnIndicadorPrimerCliente = 0#
        pnIndicadorDiezCliente = 0#
        pnIndicadorVeinteCliente = 0#
        pnSaldoTotalDepositos = 0#
        pnIdVoBoConCli = 0#
    End If
    
    
    Set rsIndice = Nothing
    Set oDCOMConecta = Nothing
End Sub

Public Function DevolverVoBoConcentracionCliente(ByVal pnId_VoBoConCli As Long) As Integer

    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim rsIndice As ADODB.Recordset

    lsSQL = "exec stp_sel_RFC0872012_DevolverVoBoConcentracionCliente " & pnId_VoBoConCli & ""

    Set oDCOMConecta = New COMConecta.DCOMConecta
    If oDCOMConecta.AbreConexion = False Then Exit Function
    Set rsIndice = oDCOMConecta.CargaRecordSet(lsSQL)
    
    If Not rsIndice.BOF And Not rsIndice.EOF Then
        DevolverVoBoConcentracionCliente = rsIndice!nEstado
    Else
        DevolverVoBoConcentracionCliente = -1
    End If
    
    
    Set rsIndice = Nothing
    Set oDCOMConecta = Nothing
    
    
End Function
'***Fin Agregado por ELRO 20120911**********************

'Add By Gitu 22-10-2012
Public Sub registraOpeFrecuentes(ByVal psDescri As String, ByVal pdFecha As String, ByVal psCodUser As String, ByVal pnEstado As Integer, _
                                ByRef pnCodOpeFre As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_ins_InsertaOperacionesFrecuente '" & psDescri & "','" & pdFecha & "','" & psCodUser & "'," & pnEstado
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    'oConec.CargaRecordSet (lsSQL)
    Set lrDatos = oConec.CargaRecordSet(lsSQL)
    pnCodOpeFre = lrDatos("nCodOpeFre")
    oConec.CierraConexion
End Sub

Public Sub registraDetalleOpeFrecuentes(ByVal pnCodOpeFre As Integer, ByVal psCtaCod As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_InsertaDetalleOperacionesFrecuente " & pnCodOpeFre & ",'" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub

Public Function DevolverCodInstConv(ByVal psCtaCod As String) As String

    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim rsCod As ADODB.Recordset

    lsSQL = "exec stp_sel_DevolverCodInstConv '" & psCtaCod & "'"

    Set oDCOMConecta = New COMConecta.DCOMConecta
    If oDCOMConecta.AbreConexion = False Then Exit Function
    Set rsCod = oDCOMConecta.CargaRecordSet(lsSQL)
    
    If Not rsCod.BOF And Not rsCod.EOF Then
        DevolverCodInstConv = rsCod!cPersCodConv
    Else
        DevolverCodInstConv = ""
    End If
    
    Set rsCod = Nothing
    Set oDCOMConecta = Nothing
    
End Function
Public Sub ActualizaCodPersConv(ByVal psCtaCod As String, ByVal psCodPersConv As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_upd_ActualizarCodPersConv '" & psCtaCod & "','" & psCodPersConv & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub
'End GITU
'***Agregado por ELRO el 20121120, según OYP-RFC101-2012
Public Function devolverCuentasPorExtornarLote(ByVal pnMovNro As Long, ByVal psOpeCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim rsCtas As New ADODB.Recordset
    lsSQL = "exec stp_sel_RFC1012012_DevolverCuentasPorExtornarLote  " & pnMovNro & ", '" & psOpeCod & "'"
    rsCtas.Open lsSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set devolverCuentasPorExtornarLote = rsCtas
End Function

Public Function actualizarCTSNoActivo(ByVal psCtaCod As String, ByVal pnTasaInteres As Currency, _
                                      ByVal pnTpoPrograma As Integer) As Boolean
    On Error GoTo ErractualizarCTSNoActivo
    sSql = "exec stp_sel_RFC1012012_ActualizarCTSNoActivo '" & psCtaCod & "', " & pnTasaInteres & ", " & pnTpoPrograma & ""
    dbCmact.Execute sSql
    actualizarCTSNoActivo = True
    Exit Function
ErractualizarCTSNoActivo:
    Err.Raise Err.Number, "NCapDefinicion:actualizarCTSNoActivo", Err.Description
    actualizarCTSNoActivo = False
End Function
Public Function devolverUltimoMovimientoDeposito(ByVal psCtaCod As String, ByVal psFecha As String) As Long
    Dim lsSQL As String
    Dim rsCtas As New ADODB.Recordset
    lsSQL = "exec stp_sel_DevolverUltimoMovimientoDeposito  '" & psCtaCod & "',  '" & psFecha & "'"
    rsCtas.Open lsSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rsCtas.BOF And rsCtas.EOF) Then
        devolverUltimoMovimientoDeposito = rsCtas!nMovNro
    Else
        devolverUltimoMovimientoDeposito = 0
    End If
End Function
'***Fin Agregado por ELRO 20121120**********************
'ALPA 20130110******************************************
Public Sub RegistrarClienteEcotaxi(ByVal pnMovNro As Long, ByVal psPersCod As String)
sSql = "exec stp_ins_MovEcotaxi " & pnMovNro & ", '" & psPersCod & "'"
dbCmact.Execute sSql
End Sub
'*******************************************************
'***Agregado por ELRO el 20130325, según OYP-ERS011-2013****
Public Sub registrarHistorialCuentasMigradas(ByVal lnMovNro As Long, ByVal psCtaCod As String)

    sSql = "exec stp_ins_ERS0112012_RegistrarHistorialCuentasMigradas  " & lnMovNro & ", '" & psCtaCod & "'"
        
    dbCmact.Execute sSql
    
End Sub

Public Sub migrarCuenta(ByVal psCtaCod As String, ByVal pnSaldo As Currency, ByVal pdUltCierre As Date, ByVal pnTpoPrograma As Integer, ByVal pnTasaInteres As Double, ByVal psPersCod As String)

    sSql = "exec stp_upd_ERS0112013_MigrarCuenta '" & psCtaCod & "', " & pnSaldo & ",  '" & Format$(pdUltCierre, "mm/dd/yyyy") & "', " & pnTpoPrograma & ", " & pnTasaInteres & ", '" & psPersCod & "'"
        
    dbCmact.Execute sSql
    
End Sub

Public Sub extornarMigracionCuentas(ByVal lnMovNro As Long)

    sSql = "exec stp_upd_ERS0112013_ExtornarMigracionCuentas  " & lnMovNro & ""
        
    dbCmact.Execute sSql
    
End Sub

'***Fin Agregado por ELRO el 20130325, según OYP-ERS011-2013
'WIOR 20130419 ****************************************************
Public Function EsDepositoGarantia(ByVal psCtaCod As String) As Boolean
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    
    Dim lrDatos As ADODB.Recordset

    lsSQL = "select ISNULL(GarantDepCF,0) GarantDepCF from CaptacPlazoFijo where cCtaCod='" & psCtaCod & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    
    Set lrDatos = oConec.CargaRecordSet(lsSQL)
    
    If Not (lrDatos.EOF And lrDatos.BOF) Then
        If CInt(lrDatos!GarantDepCF) = 1 Then
            EsDepositoGarantia = True
        Else
            EsDepositoGarantia = False
        End If
    Else
        EsDepositoGarantia = False
    End If
    
    oConec.CierraConexion
End Function
'WIOR FIN ********************************************************

'Add By GITU 2013-05-28
Public Function ValidaEsOficialoAuxiliardeCumplimiento(ByVal pCargo As String) As Boolean
    Dim oGen As COMDConstSistema.DCOMGeneral
    Dim sPosicion As String
    Dim sCod() As String
    Dim i As Integer
    
    Set oGen = New COMDConstSistema.DCOMGeneral
    sPosicion = oGen.LeeConstSistema(437)
    
    sCod = Split(sPosicion, ",")
    ValidaEsOficialoAuxiliardeCumplimiento = False
    For i = 0 To UBound(sCod)
        If pCargo = sCod(i) Then
            ValidaEsOficialoAuxiliardeCumplimiento = True
        End If
    Next i
End Function
'End GITU
'***Agregado por ELRO el 20130711, según RFC1306270002****

Public Function devolverDatosCuentaConvenioServicioPago(ByVal psIdSerPag As Long) As ADODB.Recordset

    sSql = "exec stp_sel_DevolverDatosCuentaConvenioServicioPago  " & psIdSerPag & ""
        
    Set devolverDatosCuentaConvenioServicioPago = dbCmact.Execute(sSql)
    
End Function

Public Function actualizarCuentaConvenioServicioPago(ByVal pnIdPagSeg As Long, ByVal pnNroOpe As Long, _
                                                     ByVal psCtaCod As String, ByVal pnDebito As Currency, _
                                                     ByVal pnComision As Currency, ByVal pnPenalidad As Currency, _
                                                     ByVal pnITF As Currency, ByVal pnIntAcum As Currency, _
                                                     ByVal psUltCierre As String, ByVal pnSaldoAnterior As Currency) As Long
    Dim rsDebito As ADODB.Recordset
    Set rsDebito = New ADODB.Recordset
    sSql = "exec stp_upd_ActualizarCuentaConvenioServicioPago  " & pnIdPagSeg & ", " & pnNroOpe & ", '" & psCtaCod & "', " & pnDebito & ", " & pnComision & ", " & pnPenalidad & ", " & pnITF & ", " & pnIntAcum & ", '" & psUltCierre & "', " & pnSaldoAnterior & ""
        
    Set rsDebito = dbCmact.Execute(sSql)
    
    If Not (rsDebito.BOF And rsDebito.EOF) Then
        actualizarCuentaConvenioServicioPago = rsDebito!IdSerPag
    Else
        actualizarCuentaConvenioServicioPago = 0
    End If
     Set rsDebito = Nothing
End Function

Public Function actualizarServicioPagoDebito(ByVal pnIdPagSegDeb As Long, ByVal pnMovNro As Long) As Long
    Dim rsDebito As ADODB.Recordset
    Set rsDebito = New ADODB.Recordset
    
    sSql = "exec stp_upd_ActualizarServicioPagoDebito  " & pnIdPagSegDeb & ", " & pnMovNro & ""
        
    Set rsDebito = dbCmact.Execute(sSql)
    
    If Not (rsDebito.BOF And rsDebito.EOF) Then
        actualizarServicioPagoDebito = rsDebito!IdPagSegDeb
    Else
        actualizarServicioPagoDebito = 0
    End If
    Set rsDebito = Nothing
End Function

Public Function devolverConvenioServicioPagoDebitoParaExtornar(ByVal psDatoBusqueda As String, ByVal psFecha As String, _
                                                          ByVal psOpeCod As String, ByVal psAgeCod As String, _
                                                          ByVal pnTipoBusqueda As Integer, ByVal pnMovNroBus As Long) As ADODB.Recordset
    Dim lsSQL As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Dim rsDebitos As ADODB.Recordset

    lsSQL = "exec stp_sel_DevolverConvenioServicioPagoDebitoParaExtornar '" & psDatoBusqueda & "', '" & psFecha & "', '" & psOpeCod & "', '" & psAgeCod & "', " & pnTipoBusqueda & ", " & pnMovNroBus & " "

    Set oDCOMConecta = New COMConecta.DCOMConecta
    If oDCOMConecta.AbreConexion = False Then Exit Function
    Set rsDebitos = oDCOMConecta.CargaRecordSet(lsSQL)
    Set devolverConvenioServicioPagoDebitoParaExtornar = rsDebitos
    oDCOMConecta.CierraConexion
    Set rsDebitos = Nothing
    Set oDCOMConecta = Nothing
End Function

Public Function extornarCuentaConvenioServicioPagoDebito(ByVal pnIdPagSeg As Long, ByVal pnNroOpe As Long, _
                                                         ByVal psCtaCod As String, ByVal pnDebito As Currency, _
                                                         ByVal pnComision As Currency, ByVal pnPenalidad As Currency, _
                                                         ByVal pnITF As Currency, ByVal pnIntAcum As Currency, _
                                                         ByVal psUltCierre As String, ByVal pnSaldoAnterior As Currency, _
                                                         ByVal pnMovNro As Long) As Long
    Dim rsDebito As ADODB.Recordset
    Set rsDebito = New ADODB.Recordset
    
    sSql = "exec stp_upd_ExtornarCuentaConvenioServicioPagoDebito  " & pnIdPagSeg & ", " & pnNroOpe & ", '" & psCtaCod & "', " & pnDebito & ", " & pnComision & ", " & pnPenalidad & ", " & pnITF & ", " & pnIntAcum & ", '" & psUltCierre & "', " & pnSaldoAnterior & ", " & pnMovNro & ""
        
    Set rsDebito = dbCmact.Execute(sSql)
    
    If Not (rsDebito.BOF And rsDebito.EOF) Then
        extornarCuentaConvenioServicioPagoDebito = rsDebito!nMovNro
    Else
        extornarCuentaConvenioServicioPagoDebito = 0
    End If
    Set rsDebito = Nothing
End Function
'***Fin Agregado por ELRO el 20130711, según RFC1306270002
'***Agregado por ELRO el 20130723, según TI-ERS019-2013****
Public Sub registrarMedioRetiro(ByVal pnMedio As Long, ByVal pnMovNro As Long)
    sSql = "exec stp_ins_ERS0792013_RegistrarMedioRetiro " & pnMedio & ", " & pnMovNro & ""
    dbCmact.Execute (sSql)
End Sub
'***Fin Agregado por ELRO el 20130723, según TI-ERS019-2013

'JUEZ 20130805 *************************************************
Public Function ObtenerDatosSueldosClienteInstitucionCTS(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    
    lsSQL = "exec stp_sel_ObtenerDatosSueldosClienteInstitucionCTS '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    If oConecta.AbreConexion = False Then Exit Function
    Set ObtenerDatosSueldosClienteInstitucionCTS = oConecta.CargaRecordSet(lsSQL)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'END JUEZ ******************************************************
'JUEZ 20130902 ******************************************************
Public Sub InsertaCapComisiones(ByVal pnMovNro As Long, ByVal psCodCom As String, ByVal psPersCod As String, ByVal pnTipoPago As Integer, ByVal psCtaCod As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_CapComisiones " & pnMovNro & ",'" & psCodCom & "','" & psPersCod & "'," & pnTipoPago & ",'" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub
Public Function GetMovxMovRef(ByVal pnMovNroRef As Long, ByVal pnMovNroExt As Long) As ADODB.Recordset
Dim rstemp As ADODB.Recordset
sSql = "SELECT NMOVNRO FROM MOVREF WHERE NMOVNROREF=" & pnMovNroRef & " AND NMOVNRO <> " & pnMovNroExt
Set rstemp = New ADODB.Recordset
rstemp.CursorLocation = adUseClient
rstemp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rstemp.ActiveConnection = Nothing
Set GetMovxMovRef = rstemp
Set rstemp = Nothing
End Function
'END JUEZ ***********************************************************
'******RECO 20131030 ERS141********************************************
Public Sub RegistraDevFonfoGarantEcotaxi(ByVal nMovNro As Long, ByVal sCtaCod As String, ByVal sdFechaDev As String, ByVal nMontoDev As Double _
        , ByVal cOpeCod As String, ByVal sCtaCodAbono As String, ByVal nFormaDev As Integer, ByVal sGlosa As String)
        
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_ins_RegistraDevFonfoGarantEcotaxi " & nMovNro & ",'" & sCtaCod & "','" & Format(sdFechaDev, "yyyy/MM/dd") & "'," & _
            nMontoDev & ",'" & cOpeCod & "','" & sCtaCodAbono & "'," & nFormaDev & ",'" & sGlosa & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub
Public Sub RegistraDevolucionRecaudoEcotaxi(ByVal nMovNro As Long, ByVal sCtaCod As String, ByVal sdFechaDev As String, ByVal nMontoDev As Double _
        , ByVal cOpeCod As String, ByVal sGlosa As String)
        
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_ins_RegistraDevolucionRecaudoEcotaxi " & nMovNro & ",'" & sCtaCod & "','" & Format(sdFechaDev, "yyyy/MM/dd") & "'," & _
            nMontoDev & ",'" & cOpeCod & "','" & sGlosa & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
End Sub
'***************************************END RECO************************

' RIRO20131212 ERS137 - Registro de Comision en Efectivo ****
Public Sub RegistrarTransferencia(ByVal psCuenta As String, _
                                  ByVal pnMoneda As Moneda, _
                                  ByVal psOperacion As String, _
                                  ByVal psMovNroRet As String, _
                                  ByVal psMovNroCom As String, _
                                  ByVal pnComisionTranf As Double, _
                                  ByVal pnSaldoDisp As Double, _
                                  ByVal pnSaldoContable As Double, _
                                  ByVal pdUltMov As Date, _
                                  ByVal pnComRetTransVB As Integer, _
                                  Optional ByVal dAuxiliar As Date)
    Dim nMovNro As Long
    Dim nMovNroRet As Long
    Dim sOperacionComision As CaptacOperacion
    
    Select Case psOperacion
    
        Case "200303" 'Retiro Ahorro
            If pnComRetTransVB = 1 Then
                sOperacionComision = 200370
            Else
                sOperacionComision = 200371
            End If
   
        Case "200411" 'Cancelacion Ahorro
            If pnComRetTransVB = 1 Then
                sOperacionComision = 200372
            Else
                sOperacionComision = 200373
            End If
        
        Case "210209" 'Retiro P. Fijo
            If pnComRetTransVB = 1 Then
                sOperacionComision = 210211
            Else
                sOperacionComision = 210212
            End If
        
        Case "210302" 'Cancelacion P. Fijo
            If pnComRetTransVB = 1 Then
                sOperacionComision = 210213
            Else
                sOperacionComision = 210214
            End If
        
        Case "220302" 'Retiro CTS
            If pnComRetTransVB = 1 Then
                sOperacionComision = 220306
            Else
                sOperacionComision = 220307
            End If
     
        Case "220403" 'Cancelacion CTS
            If pnComRetTransVB = 1 Then
                sOperacionComision = 220308
            Else
                sOperacionComision = 220309
            End If
        
    End Select
    
    If pnComRetTransVB = 1 Then
        AgregaMov psMovNroCom, sOperacionComision, "COMISION TRANSFERENCIA BANCO", gMovEstContabNoContable, gMovFlagVigente
        nMovNro = GetnMovNro(psMovNroCom)
        AgregaMovOpeVarias nMovNro, pnComisionTranf, "", "", pnMoneda, sOperacionComision
        AgregaMovGasto nMovNro, "", ""
        nMovNroRet = GetnMovNro(psMovNroRet)
        AgregaMovRef nMovNroRet, nMovNro
        
    Else
        nMovNro = GetnMovNro(psMovNroRet)
        Select Case psOperacion
            Case "200303" 'Retiro Ahorro
                
                ActualizaCargoCaptacion psCuenta, pnComisionTranf, pnComisionTranf, 0, pdUltMov, psMovNroRet
                AgregaMovCap nMovNro, sOperacionComision, psCuenta, pnComisionTranf, pnSaldoDisp, pnSaldoContable
                AgregaMovCapDet nMovNro, sOperacionComision, psCuenta, 27, pnComisionTranf
                                
            Case "200411" 'Cancelacion Ahorro
                
                AgregaMovCap nMovNro, sOperacionComision, psCuenta, pnComisionTranf, 0, 0
                AgregaMovCapDet nMovNro, sOperacionComision, psCuenta, 27, pnComisionTranf
                     
            Case "210209" 'Retiro P. Fijo
                'ActualizaRetiroInteresPF psCuenta, pnComisionTranf, 0, pdUltMov, psMovNroRet, dAuxiliar, False
                AgregaMovCap nMovNro, sOperacionComision, psCuenta, pnComisionTranf, pnSaldoDisp, pnSaldoContable
                AgregaMovCapDet nMovNro, sOperacionComision, psCuenta, 27, pnComisionTranf
            
            Case "210302" 'Cancelacion P. Fijo

                AgregaMovCap nMovNro, sOperacionComision, psCuenta, pnComisionTranf, 0, 0
                AgregaMovCapDet nMovNro, sOperacionComision, psCuenta, 27, pnComisionTranf
            
            Case "220302" 'Retiro CTS

                ActualizaCargoCaptacion psCuenta, pnComisionTranf, pnComisionTranf, 0, pdUltMov, psMovNroRet, False
                AgregaMovCap nMovNro, sOperacionComision, psCuenta, pnComisionTranf, pnSaldoDisp, pnSaldoContable
                AgregaMovCapDet nMovNro, sOperacionComision, psCuenta, 27, pnComisionTranf

            Case "220403" 'Cancelacion CTS
                
                AgregaMovCap nMovNro, sOperacionComision, psCuenta, pnComisionTranf, 0, 0
                AgregaMovCapDet nMovNro, sOperacionComision, psCuenta, 27, pnComisionTranf
        
        End Select
    
    End If
        
End Sub
' END RIRO ****
'JUEZ 20140305 ************************************************************************
Public Sub RevertirCTSActivo(ByVal psPersCod As String, ByVal psCodInst As String, ByVal psCodUser As String, ByVal psCodAge As String, ByRef pbExito As Boolean)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    lsSQL = "exec stp_upd_RevertirCTSActivo '" & psPersCod & "','" & psCodInst & "','" & psCodUser & "','" & psCodAge & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(lsSQL)
    If Not (rs.EOF And rs.BOF) Then
        pbExito = IIf(rs!bExito, True, False)
    Else
        pbExito = False
    End If
    oConec.CierraConexion
End Sub
Public Sub CTSCeseLaboral(ByVal psPersCod As String, ByVal psCodInst As String, ByVal psCodUser As String, ByVal psCodAge As String, ByRef pbExito As Boolean)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    lsSQL = "exec stp_upd_CTSCeseLaboral '" & psPersCod & "','" & psCodInst & "','" & psCodUser & "','" & psCodAge & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(lsSQL)
    If Not (rs.EOF And rs.BOF) Then
        pbExito = IIf(rs!bExito, True, False)
    Else
        pbExito = False
    End If
    oConec.CierraConexion
End Sub
'END JUEZ *****************************************************************************
'EJVG20140205 ***
Public Sub AgregaMovDocRec(ByVal pnMovNro As Long, ByVal pnDocTpo As TpoDoc, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psIFCta As String)
    Dim lsSQL As String
    lsSQL = "Exec stp_ins_ERS1262013_MovDocRec " & pnMovNro & "," & pnDocTpo & ",'" & psNroDoc & "','" & psPersCod & "','" & psIFTpo & "','" & psIFCta & "'"
    dbCmact.Execute lsSQL
End Sub
'END EJVG *******
'RIRO 20140430 ERS017 ******
Public Sub RegistrarSobrante(ByVal pnId As Integer, ByVal pnMovNroSob As Long)
    Dim lsSQL As String
    lsSQL = "Exec STP_UPD_REGISTRARSOBRANTEVOUCHER " & pnId & "," & pnMovNroSob
    dbCmact.Execute (lsSQL)
    Exit Sub
End Sub
Public Sub ExtornarSobrante(ByVal pnMovNro As Long)
    Dim lsSQL As String
    lsSQL = "Exec STP_UPD_EXTORNARSOBRANTEVOUCHER " & pnMovNro
    dbCmact.Execute (lsSQL)
    Exit Sub
End Sub
Public Function registrarVoucherErr(ByVal sMovNro As String, _
                                    ByVal nTipMot As Integer, _
                                    ByVal nMovNroErr As Long, _
                                    ByVal nMovNroPend As Long, _
                                    ByVal nOpeRestantes As Integer, _
                                    ByVal nMonto As Currency) As Long
    Dim lsSQL As String
    Dim rsCtas As New ADODB.Recordset
    Dim nMovNro As Long
    
    On Error GoTo error
    
    lsSQL = "exec stp_ins_RegistrarCapRegVouDep_Error '" & sMovNro & "', " _
                                                         & nTipMot & ", " _
                                                         & nMovNroErr & ", " _
                                                         & nMovNroPend & ", " _
                                                         & nOpeRestantes & ", " _
                                                         & nMonto
    rsCtas.Open lsSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    If Not rsCtas Is Nothing Then
        If Not (rsCtas.BOF And rsCtas.EOF) Then
            nMovNro = rsCtas!nMovNro
        Else
            nMovNro = 0
        End If
    End If
    registrarVoucherErr = nMovNro
Exit Function
error:
    Err.Raise Err.Number, "DCOMCaptaMovimiento:registrarVoucherErr", Err.Description
    registrarVoucherErr = 0
End Function
Public Sub ActualizarVoucherErr(ByVal nMovNroErr As Long, _
                                ByVal nMovNro As Long, _
                                ByVal nMonto As Currency)
    
    Dim lsSQL As String
    Dim rs As ADODB.Recordset
    
    On Error GoTo error
    lsSQL = "exec stp_ins_UpdateCapRegVouDep_Error " & nMovNroErr & "," _
                                                     & nMovNro & "," _
                                                     & nMonto & ""

    Set rs = dbCmact.Execute(lsSQL)
Exit Sub
error:
    Err.Raise Err.Number, "DCOMCaptaMovimiento:ActualizarVoucherErr", Err.Description
End Sub
Public Sub EliminarVoucherErr(ByVal nMovNroErr As Long, _
                              ByVal nMovNro As Long, _
                              ByVal nMovNroOpe As Long, _
                              ByVal nDetalle As Long, _
                              ByVal nMonto As Currency)
        
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    
    On Error GoTo error
    
    lsSQL = "exec stp_del_DelCapRegVouDep_Error " & nMovNroErr & "," _
                                                  & nMovNro & "," _
                                                  & nMovNroOpe & "," _
                                                  & nDetalle & "," _
                                                  & nMonto & ""

    Set rs = dbCmact.Execute(lsSQL)
Exit Sub
error:
    Err.Raise Err.Number, "DCOMCaptaMovimiento:EliminarVoucherErr", Err.Description
End Sub
'END RIRO ********
'JUEZ 20140710 **********************************************************
Public Sub AgregaSegTarjetaAfiliacion(ByVal psNumCert As String, ByVal psNumTarj As String, ByVal psCtaCod As String, _
                                       ByVal pdFecSis As String, ByVal psMovNro As String, ByVal pnMovNro As Long, ByVal psPersCod As String, ByVal pbPrimaAnual As Integer, _
                                       Optional ByVal pcUser As String = "", Optional ByVal pcUserVenta As String = "")
'APRI20181023 ERS071-2018 ADD pbPrimaAnual
'CTI2 FERIMORO ERS034-2019   31082020  : Optional ByVal pcUser As String = "", Optional ByVal pcUserVenta As String = "")
Dim oConec As COMConecta.DCOMConecta
Dim sHora As String
Dim lsSQL As String
    sHora = Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2)
'CTI2 FERIMORO ERS034-2019  31082020
'    lsSQL = "Exec stp_ins_SegTarjetaAfiliacion '" & psNumCert & "','" & psNumTarj & "','" & psCtaCod & "','" & _
'            Format(pdFecSis & " " & sHora, "mm/dd/yyyy hh:mm:ss") & "','" & psMovNro & "'," & pnMovNro & ",'" & psPersCod & "'," & pbPrimaAnual

    lsSQL = "Exec stp_ins_SegTarjetaAfiliacion '" & psNumCert & "','" & psNumTarj & "','" & psCtaCod & "','" & _
            Format(pdFecSis & " " & sHora, "mm/dd/yyyy hh:mm:ss") & "','" & psMovNro & "'," & pnMovNro & ",'" & psPersCod & "'," & pbPrimaAnual & ",'" & pcUser & "','" & pcUserVenta & "'"  'CTI2 FERIMORO ERS034-2019 31082020 '" & pcUser & "','" & pcUserVenta & "'"

    dbCmact.Execute (lsSQL)
End Sub
'END JUEZ ***************************************************************
'JUEZ 20140915 ****************************************************************
Public Sub ActualizaSueldosCTS(ByVal pnTpoPrograma As Integer, ByVal pnUltRemSueldosAnt As Integer, ByVal pnUltRemSueldos As Integer, ByVal psMovNro As String)
Dim oConec As COMConecta.DCOMConecta
Dim sHora As String
Dim lsSQL As String
    lsSQL = "Exec stp_upd_ActualizaSueldosCTS " & pnTpoPrograma & "," & pnUltRemSueldosAnt & "," & pnUltRemSueldos & ",'" & psMovNro & "'"

    dbCmact.Execute (lsSQL)
End Sub
'END JUEZ *********************************************************************

'RIRO20141202 Correccion SR ********************
Public Function RegistroServicioRecaudo(ByVal arrDetalleCobro As Variant, _
                                    ByVal psCodConvenio As String, _
                                    ByVal psCodCliente As String, _
                                    ByVal pnTipoDoi As Integer, _
                                    ByVal psDoi As String, _
                                    ByVal psNombreCliente As String, _
                                    ByVal pdFechaCobro As Date, _
                                    ByVal psMovNro As String, _
                                    ByVal psMovNro2 As String, _
                                    ByVal pnComiEmp As Double, _
                                    ByVal pnComiCli As Double, _
                                    ByVal pnITF As Double, _
                                    ByVal pnIntGanado As Double, _
                                    ByVal psCuenta As String, _
                                    Optional ByVal pnConnBase As ADODB.Connection = Nothing, _
                                    Optional ByVal pnTipoPagoRecaudo As TipoPagoRecaudo = gPagoNormal, _
                                    Optional ByVal pbTieneWS As Boolean = False, _
                                    Optional ByVal cUrlSimaynas As String = "", Optional ByVal pnTipoPago As Integer = 1, Optional ByVal psCtaCodCargo As String = "", _
                                    Optional ByRef pMatAho As Variant = "", _
                                    Optional ByVal pbITFAplica As Boolean, _
                                    Optional ByVal pnMontoITF As Currency, _
                                    Optional ByVal pnMontoCargoCta As Currency = 0#, _
                                    Optional ByVal pnMovNroTransfer As Long = -1, Optional pnMonedaTrans As Moneda = gMonedaNacional, _
                                    Optional ByVal pnMovNroRVD As Long = 0, _
                                    Optional ByVal pnMontoRVD As Currency = 0) As Boolean 'CTI1 ERS027-2019 add pbTieneWS, cUrlSimaynas
                                    'JUEZ 20150209 Se agregó pnConnBase y pnTipoPagoRecaudo

Dim v As Variant
Dim arrCadena() As String
Dim sSql As String
Dim nMovNro As Long
Dim nMonto As Double
Dim nMontoTotalCargoCta As Double 'CTI6ERS0112020

Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nTasa As Double
Dim dUltMov As Date
Dim bInactiva As Boolean
Dim bInamovilizada As Boolean
Dim nExtracto As Long
Dim dFecSis As Date
Dim rsCta As ADODB.Recordset
Dim nOpeDepositoRecaudo As OperacionesRecaudo 'JUEZ 20150209
'CTI1 ERS027-2019
Dim ClsServicioRecaudoWS As COMDCaptaServicios.DCOMSrvRecaudoWS
Dim nCodConvenioRecaudoWS As Integer
Dim sCodigoCobranza As String
Dim sCodigoComprobante As String
Dim sCodigoCaja As String
Dim sCodigoContratista As String
Dim sCodigoContratistaContrato As String
Dim sAperturaCaja As String
Dim sMontoCobrado As String
Dim sMontoRecibido As String
Dim sFechaCobranza As String
Dim sCodigoCobranzaRes As String
'CTI1 ERS027-2019

On Error GoTo error
    
    If pnTipoPago = gCVTipoPagoEfectivo Then
    'JUEZ 20150209 *******************
    If pnTipoPagoRecaudo = gPagoDebitoAutomatico Then
        nOpeDepositoRecaudo = gDepositoRecaudoCargoCta
    Else
        nOpeDepositoRecaudo = gDepositoRecaudo
    End If
    
    If Not pnConnBase Is Nothing Then
        SetConexion pnConnBase
        bTransaccion = True
    End If
    'END JUEZ ************************
    ElseIf pnTipoPago = gCVTipoPagoCargoCta Then
        nOpeDepositoRecaudo = gRecuadoCargoCtaAhorro
    ElseIf pnTipoPago = gCVTipoPagoVoucher Then
         nOpeDepositoRecaudo = gServicioRecaudoVoucher
    End If

    'dbCmact.BeginTrans
    If Not bTransaccion Then dbCmact.BeginTrans 'JUEZ 20150209
    
    'Obteniendo Mov
    AgregaMov psMovNro, nOpeDepositoRecaudo, "Abono Cuenta = " & psCuenta
    nMovNro = GetnMovNro(psMovNro)
    
    'Insertando Valores a tabla IngresosCobroRecaudo
    If pbTieneWS = False Then 'CTI1 ERS027-2019
        For Each v In arrDetalleCobro
            arrCadena = Split(v, "|")
            If v = "" Then Exit For
            sSql = "exec stp_ins_upd_RegistrarCobroServicioRecaudo "
    
            'Codigo de convenio
            sSql = sSql & "'" & psCodConvenio & "', "
            'Id de la trama
            sSql = sSql & "'" & arrCadena(0) & "',"
            'Codigo de cliente
            sSql = sSql & "'" & psCodCliente & "',"
            'Tipo de Doi
            sSql = sSql & pnTipoDoi & ","
            'Doi
            sSql = sSql & "'" & psDoi & "',"
            'Nombre de Cliente
            sSql = sSql & "'" & psNombreCliente & "',"
            'Servicio
            sSql = sSql & "'" & arrCadena(1) & "',"
            'Concepto
            sSql = sSql & "'" & arrCadena(2) & "',"
            'Importe
            sSql = sSql & arrCadena(3) & ","
            'Deuda Actual
            sSql = sSql & arrCadena(4) & ","
            'Monto Cobro
            sSql = sSql & arrCadena(5) & ","
            'Estado
            sSql = sSql & "'" & arrCadena(6) & "',"
            'FechaCobro
            sSql = sSql & "'" & Format$(pdFechaCobro, "yyyymmdd") & "',"
            'cMovNro
            sSql = sSql & "'" & psMovNro & "',"
            'Comision Empresa
            sSql = sSql & CDbl(pnComiEmp) & ","
            'Comision Cliente
            sSql = sSql & CDbl(pnComiCli) & ","
            'Mora
            sSql = sSql & CDbl(arrCadena(7)) & ","
            'Fecha Vencimiento
            sSql = sSql & "'" & arrCadena(8) & "'"
    
            nMonto = nMonto + Val(arrCadena(5)) + CDbl(arrCadena(7)) ' Acumulamos el monto Sub total
            dbCmact.Execute sSql
            sSql = ""
        Next
    Else 'CTI1 ERS027-2019
        For Each v In arrDetalleCobro
            arrCadena = Split(v, "|")
            If v = "" Then Exit For
            nMonto = nMonto + Val(arrCadena(5)) + CDbl(arrCadena(7)) ' Acumulamos el monto Sub total
        Next
    End If 'CTI1 ERS027-2019

    dFecSis = CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4))
    
    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsCta = clsMant.GetDatosCuentaAho(psCuenta)
    Set clsMant = Nothing
    
    nEstado = rsCta("nPrdEstado")
    nSaldoDisp = rsCta("nSaldoDisp")
    nSaldoCnt = rsCta("nSaldo")
    nTasa = rsCta("nTasaInteres")
    nIntAcum = rsCta("nIntAcum")
    bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
    bInamovilizada = IIf(rsCta("bInmovilizada") = 0, False, True)
    nExtracto = rsCta("nTransacc") + 1
    dUltMov = DateAdd("d", -1, dFecSis)
    'Deposito
    If nMonto > 0 Then
        nSaldoCnt = nSaldoCnt + nMonto
        ActualizaSaldoAnteriorAho psCuenta, nSaldoDisp
        nSaldoDisp = nSaldoDisp + nMonto
        ActualizadUltContactoAho psCuenta, dFecSis
        AgregaMovCap nMovNro, nOpeDepositoRecaudo, psCuenta, nMonto, nSaldoDisp, nSaldoCnt 'JUEZ 20150209
        AgregaMovCapDet nMovNro, nOpeDepositoRecaudo, psCuenta, gConcCapital, nMonto 'JUEZ 20150209
        ActualizaAbonoCaptacion psCuenta, nMonto, nMonto, pnIntGanado, dUltMov, psMovNro, True
        UltimaActualizacionCuenta psCuenta, psMovNro 'nMovNro CTI2 20190410 Corrección de actualizació de campo cUltimaActualizacion
    End If
    'comision empresa
    If pnComiEmp > 0 Then
        nSaldoCnt = nSaldoCnt - pnComiEmp
        nSaldoDisp = nSaldoDisp - pnComiEmp
        AgregaMovCap nMovNro, gCargoEmpresaRecaudo, psCuenta, pnComiEmp, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro, gCargoEmpresaRecaudo, psCuenta, cComiServRecaudoEmpresa, pnComiEmp
        ActualizaAbonoCaptacion psCuenta, -pnComiEmp, -pnComiEmp, 0, dUltMov, psMovNro, True
        UltimaActualizacionCuenta psCuenta, psMovNro 'nMovNro CTI2 20190410 Corrección de actualizació de campo cUltimaActualizacion
    End If
    'ITF
    If pnITF > 0 Then
        nSaldoCnt = nSaldoCnt - pnITF
        nSaldoDisp = nSaldoDisp - pnITF
        AgregaMovCap nMovNro, gIFTRecaudo, psCuenta, pnITF, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro, gIFTRecaudo, psCuenta, gConcITFCliente, pnITF
        ActualizaAbonoCaptacion psCuenta, -pnITF, -pnITF, 0, dUltMov, psMovNro, True
        UltimaActualizacionCuenta psCuenta, psMovNro 'nMovNro CTI2 20190410 Corrección de actualizació de campo cUltimaActualizacion
    End If
    'comision cliente
    If pnComiCli > 0 And Not pnTipoPago = gCVTipoPagoCargoCta Then
        'JUEZ 20150209 *********************************************
        'RegistraComisionCliente nMovNro, pnComiCli, CInt(Mid(psCuenta, 9, 1)), psMovNro2
        If pnTipoPagoRecaudo <> gPagoDebitoAutomatico Then
            RegistraComisionCliente nMovNro, pnComiCli, CInt(Mid(psCuenta, 9, 1)), psMovNro2
        End If
        'END JUEZ **************************************************
    End If
    'CTI6 ERS0112020 gCVTipoPagoEfectivo
    If pnTipoPago = gCVTipoPagoCargoCta Then


        Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsCta = clsMant.GetDatosCuentaAho(psCtaCodCargo)
        Set clsMant = Nothing
        
        nEstado = rsCta("nPrdEstado")
        nSaldoDisp = rsCta("nSaldoDisp")
        nSaldoCnt = rsCta("nSaldo")



        Dim lsFechaHoraGrab As String
        lsFechaHoraGrab = SacarFechaHoraGrab(psMovNro)
        Dim lsOpeCodCargoCta As String
        Dim lsOpeCodCargoCtaDescripcion As String
        lsOpeCodCargoCtaDescripcion = "Cargo a Cuenta Recaudo: " & lsOpeCodCargoCta
        lsOpeCodCargoCta = gAhoCargoServicioRecaudo
        CapCargoCuentaAho pMatAho, psCtaCodCargo, pnMontoCargoCta + pnMontoITF - pnComiCli, lsOpeCodCargoCta, psMovNro, lsOpeCodCargoCtaDescripcion, lsFechaHoraGrab, , , , , , , , , , , , , , , , , , , pnTipoPago
        If pnMontoCargoCta - pnComiCli > 0 And pnComiCli > 0 Then
            nSaldoCnt = nSaldoCnt - (pnMontoCargoCta + pnMontoITF - pnComiCli) - pnComiCli
            nSaldoDisp = nSaldoDisp - (pnMontoCargoCta + pnMontoITF - pnComiCli) - pnComiCli
            AgregaMovCap nMovNro, gCargoClienteRecaudoCargoCta, psCtaCodCargo, pnComiCli, nSaldoDisp, nSaldoCnt
            AgregaMovCapDet nMovNro, gCargoClienteRecaudoCargoCta, psCtaCodCargo, cComiServRecaudoCliente, pnComiCli
            ActualizaAbonoCaptacion psCtaCodCargo, -pnComiCli, -pnComiCli, 0, dUltMov, psMovNro, True
            UltimaActualizacionCuenta psCtaCodCargo, psMovNro
        End If
        
'        Dim lnMovNro_Cargo As Integer
'        lnMovNro_Cargo = GetnMovNro(sMovNro)
    End If
    'END CTI6
    If pnTipoPago = gCVTipoPagoVoucher Then
        Dim nMoneda As Integer
        nMoneda = pnMonedaTrans
        Dim oGen As New COMDConstSistema.DCOMGeneral
        If pnMovNroTransfer <> -1 Then AgregaMovRef nMovNro, pnMovNroTransfer
        If pnMovNroTransfer <> -1 Then ActualizaMovPendientesRend pnMovNroTransfer, pnMontoRVD
        actualizarRelacionVoucherCaptacion nMovNro, pnMovNroRVD
    End If
    RegistraOperacion psCodConvenio
    'CTI1 ERS027-2019
    If pbTieneWS = True Then
        Set ClsServicioRecaudoWS = New COMDCaptaServicios.DCOMSrvRecaudoWS
        Dim rRes As ADODB.Recordset
        
        nCodConvenioRecaudoWS = ClsServicioRecaudoWS.ObtenerCodConvenioRecaudoWebService(psCodConvenio)
        
        Select Case nCodConvenioRecaudoWS
            Case 1 'ELUC
                Set rRes = ClsServicioRecaudoWS.ObtenerValoresRecaudoWebService(1) 'ELUC
                sCodigoCobranza = "0" '0 para realizar la solicitud
                sCodigoCaja = rRes!nCodigoCaja
                sCodigoContratista = rRes!nCodigoContratista
                sCodigoContratistaContrato = rRes!nCodigoContratistaContrato
                sAperturaCaja = Format$(rRes!dAperturaDia, "dd/mm/yyyy")
                sFechaCobranza = Format$(pdFechaCobro, "dd/mm/yyyy")
                
                For Each v In arrDetalleCobro
                    arrCadena = Split(v, "|")
                    If v = "" Then Exit For
                    sCodigoComprobante = arrCadena(0)
                    sMontoCobrado = arrCadena(5)
                    sMontoRecibido = arrCadena(5)
                    
                    'If sAperturaCaja = sFechaCobranza Then 'CTI7 COMENTO ESTO
                    'sCodigoCobranzaRes = ClsServicioRecaudoWS.GenerarComprobanteELUC(sCodigoCobranza, sCodigoComprobante, sCodigoCaja, _
                    '                                        sCodigoContratista, sCodigoContratistaContrato, sMontoCobrado, _
                    '                                        sMontoRecibido, sFechaCobranza, cUrlSimaynas, psCodConvenio, psCodCliente)
                    sCodigoCobranzaRes = ClsServicioRecaudoWS.GenerarComprobanteELUC(psCodCliente, sCodigoComprobante, _
                                                            sMontoRecibido, psMovNro, cUrlSimaynas)
                   'CTI1 ERS027-2019 ELOR
                    sSql = StringcobroServRecaudo(arrCadena, psCodConvenio, psCodCliente, pnTipoDoi, _
                                                  psDoi, psNombreCliente, pdFechaCobro, psMovNro, _
                                                  pnComiEmp, pnComiCli, sCodigoCobranzaRes)
                                                  'CTI1 ERS027-2019 ELOR
                    dbCmact.Execute sSql
                    sSql = ""
                   
                    'CTI7 COMENTO ESTO
                    'Else
                    '    MsgBox "La fecha del sistema no coincide" & Chr(13) & _
                    '    "con la fecha de Apertura Caja ELUC." & Chr(13) & _
                    '    "Comuniquese con el area de TI.", vbInformation, "Aviso"
                    '    sCodigoCobranzaRes = ""
                    '    Exit For
                    'End If
                Next
            Case 2 'ELOR CTI1 ERS027-2019 ELOR
                 Set rRes = ClsServicioRecaudoWS.ObtenerValoresRecaudoWebService(2) 'ELOR
                sAperturaCaja = Format$(rRes!dAperturaDia, "dd/mm/yyyy")
                sFechaCobranza = Format$(pdFechaCobro, "dd/mm/yyyy")
                
                For Each v In arrDetalleCobro
                    arrCadena = Split(v, "|")
                    If v = "" Then Exit For
                    sCodigoComprobante = arrCadena(0)
                    sMontoCobrado = arrCadena(5)
                    
                    'If sAperturaCaja = sFechaCobranza Then CTI7 COMENTO ESTO
                    sCodigoCobranzaRes = ClsServicioRecaudoWS.RealizarPagoELOR(sCodigoComprobante, sMontoCobrado, psMovNro, cUrlSimaynas)
                    
                    sSql = StringcobroServRecaudo(arrCadena, psCodConvenio, psCodCliente, pnTipoDoi, _
                                                  psDoi, psNombreCliente, pdFechaCobro, psMovNro, _
                                                  pnComiEmp, pnComiCli)
                    dbCmact.Execute sSql
                    sSql = ""
                   
                    'CTI7 COMENTO ESTO
                    'Else
                    '    MsgBox "La fecha del sistema no coincide" & Chr(13) & _
                    '    "con la fecha de Apertura ELOR." & Chr(13) & _
                    '    "Comuniquese con el area de TI.", vbInformation, "Aviso"
                    '    sCodigoCobranzaRes = ""
                    '    Exit For
                    'End If
                Next
                 'CTI1 ERS027-2019 ELOR
            Case Else
                MsgBox "Codigo de Servicio Web Incorrecto", vbInformation, "Aviso"
        End Select
        If sCodigoCobranzaRes <> "" Then
            If Not bTransaccion Then dbCmact.CommitTrans
            RegistroServicioRecaudo = True
            Exit Function
        Else
            If Not bTransaccion Then dbCmact.RollbackTrans
            RegistroServicioRecaudo = False
            Exit Function
        End If
    End If
    'CTI1 ERS027-2019
    'dbCmact.CommitTrans
    If Not bTransaccion Then dbCmact.CommitTrans 'JUEZ 20150209
    RegistroServicioRecaudo = True
Exit Function
error:
    'dbCmact.RollbackTrans
    If Not bTransaccion Then dbCmact.RollbackTrans 'JUEZ 20150209
    RegistroServicioRecaudo = False
End Function
'CTI6 ERS0112020
Public Function SacarFechaHoraGrab(ByVal psMovNro As String) As String
    SacarFechaHoraGrab = Mid(psMovNro, 1, 4) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 7, 2) & " " & Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2)
End Function
'End CTI6
Public Function ExtornoRegistroServicioRecaudo(ByVal nComiEmp As Double, _
                                               ByVal nMoneda As Integer, _
                                               ByVal psCuenta As String, _
                                               ByVal nITF As Double, _
                                               ByVal nMovNro As Double, _
                                               ByVal strMovNro As String, _
                                               ByVal strMovNro2 As String, _
                                               ByVal nMonto As Double, _
                                               ByVal sGlosa As String, _
                                               Optional ByVal sCodConvenio As String, _
                                               Optional ByVal sCodCliente As String, _
                                               Optional ByVal bTieneWS As Boolean = False, _
                                               Optional ByVal cUrlSimaynas As String, Optional ByVal cOpeCod As String, Optional ByVal psMovNroCap As String = "", _
                                               Optional ByRef psOpeCodExtornoCap As String = "", Optional ByRef psCtaAhoExt As String = "", _
                                               Optional ByRef pnMontoAhorroExt As Double = 0#, Optional ByRef psOperacionDescExt As String, _
                                               Optional ByRef pnITFAhoExt As Double, Optional ByRef psMovNroAExt As String = "", Optional ByRef psClienteExt As String = "") As Boolean
                                               'CTI1 ERS027-2019 add sCodConvenio, bTieneWS, cUrlSimaynas

    Dim sSql As String
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim reRef As ADODB.Recordset
    Dim nMovNro2 As Long
    
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim nEstado As CaptacEstado
    Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
    Dim nTasa As Double
    Dim dUltMov As Date
    Dim bInactiva As Boolean
    Dim bInamovilizada As Boolean
    Dim nExtracto As Long
    Dim dFecSis As Date
    Dim rsCta As ADODB.Recordset
    'CTI1 ERS027-2019
    Dim ClsServicioRecaudoWS As COMDCaptaServicios.DCOMSrvRecaudoWS
    Dim nCodConvenioRecaudoWS As Integer
    Dim sCodigoCobranza As String
    Dim sjustificacion As String
    Dim sloginUsuario As String
    Dim bRespuestaWS As Boolean
    'CTI1 ERS027-2019
        
    On Error GoTo errorRegistrarCobroServicioConvenio
    
    dbCmact.BeginTrans

    If bTieneWS = False Then 'CTI1 ERS027-2019
        sSql = "exec stp_upd_ExtornarCobroServRecaudo " & nMovNro & ", " & nMonto
        dbCmact.Execute sSql
    End If 'CTI1 ERS027-2019
    ActualizaEstadoMov nMovNro, gMovFlagExtornado
    AgregaMov strMovNro, gExtornoDepositoRecaudo, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
    nMovNro2 = GetnMovNro(strMovNro)
    
    dFecSis = CDate(Mid(strMovNro, 7, 2) & "/" & Mid(strMovNro, 5, 2) & "/" & Left(strMovNro, 4)) 'RIRO20150303 Correccion para el calculo de interés.
    
    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsCta = clsMant.GetDatosCuentaAho(psCuenta)
    Set clsMant = Nothing
    nEstado = rsCta("nPrdEstado")
    nSaldoDisp = rsCta("nSaldoDisp")
    nSaldoCnt = rsCta("nSaldo")
    nTasa = rsCta("nTasaInteres")
    nIntAcum = rsCta("nIntAcum")
    bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
    bInamovilizada = IIf(rsCta("bInmovilizada") = 0, False, True)
    nExtracto = rsCta("nTransacc") + 1
    dUltMov = DateAdd("d", -1, dFecSis)
    'Extorno Deposito
    If nMonto > 0 Then
        nSaldoCnt = nSaldoCnt - nMonto
        ActualizaSaldoAnteriorAho psCuenta, nSaldoDisp
        nSaldoDisp = nSaldoDisp - nMonto
        ActualizadUltContactoAho psCuenta, dFecSis
        AgregaMovCap nMovNro2, gExtornoDepositoRecaudo, psCuenta, nMonto, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro2, gExtornoDepositoRecaudo, psCuenta, gConcCapital, nMonto
        ActualizaAbonoCaptacion psCuenta, -nMonto, -nMonto, 0, dUltMov, strMovNro, True
        UltimaActualizacionCuenta psCuenta, nMovNro
    End If
    'Extorno Comision Empresa
    If nComiEmp > 0 Then
        nSaldoCnt = nSaldoCnt + nComiEmp
        nSaldoDisp = nSaldoDisp + nComiEmp
        AgregaMovCap nMovNro2, gExtornoComisionEmpresa, psCuenta, nComiEmp, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro2, gExtornoComisionEmpresa, psCuenta, cComiServRecaudoEmpresa, nComiEmp
        ActualizaAbonoCaptacion psCuenta, nComiEmp, nComiEmp, 0, dUltMov, strMovNro, True
    End If
    'Extorno ITF
    If nITF > 0 Then
        nSaldoCnt = nSaldoCnt + nITF
        nSaldoDisp = nSaldoDisp + nITF
        AgregaMovCap nMovNro2, gExtornoITF, psCuenta, nITF, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro2, gExtornoITF, psCuenta, gConcITFCliente, nITF
        ActualizaAbonoCaptacion psCuenta, nITF, nITF, 0, dUltMov, strMovNro, True
    End If
    'Extorno comision de cliente
    Set reRef = MovAsociados(nMovNro)
    If Not reRef Is Nothing Then
        If Not reRef.EOF And Not reRef.BOF Then
            ActualizaEstadoMov reRef!nMovNroRef, gMovFlagExtornado
            AgregaMov strMovNro2, gExtornoComisionCliente, sGlosa, gMovEstContabNoContable, gMovFlagDeExtorno
        End If
    End If
    'CTI1 ERS027-2019
    If bTieneWS = True Then
        Set ClsServicioRecaudoWS = New COMDCaptaServicios.DCOMSrvRecaudoWS
        Dim rRes As ADODB.Recordset
        
        nCodConvenioRecaudoWS = ClsServicioRecaudoWS.ObtenerCodConvenioRecaudoWebService(sCodConvenio)
        
        Select Case nCodConvenioRecaudoWS
            Case 1 'ELUC
                Set rRes = ClsServicioRecaudoWS.ObtenerValoresRecaudoWebService(1) 'ELUC
                sjustificacion = sGlosa
                sloginUsuario = rRes!cLoginUsuario
                sCodigoCobranza = ClsServicioRecaudoWS.ObtenerCodigoCobranzaWS(sCodConvenio, sCodCliente) '
                'bRespuestaWS = ClsServicioRecaudoWS.AnularComprobanteELUC(sCodigoCobranza, sjustificacion, sloginUsuario, cUrlSimaynas)'CTI7
                bRespuestaWS = ClsServicioRecaudoWS.AnularComprobanteELUC(sCodigoCobranza, nMovNro, cUrlSimaynas)
            Case 2 'ELOR
                bRespuestaWS = True 'CTI1 ERS027-2019 ELOR
            Case Else
                MsgBox "Codigo de Servicio Web Incorrecto", vbInformation, "Aviso"
        End Select
        If bRespuestaWS = True Then
            sSql = "exec stp_upd_ExtornarCobroServRecaudo " & nMovNro & ", " & nMonto
            dbCmact.Execute sSql
            dbCmact.CommitTrans
            ExtornoRegistroServicioRecaudo = True
            Exit Function
        Else
            dbCmact.RollbackTrans
            ExtornoRegistroServicioRecaudo = False
            Exit Function
        End If
    End If
    'CTI6 ERS0112020
    If cOpeCod = gExtornoDepositoRecaudo Then
        Dim RCapMov As ADODB.Recordset
        Dim cCtaCodCargo As String
        Set RCapMov = RecuperaMovimientoCapataciones(nMovNro)
         Do While Not RCapMov.EOF
                 If RCapMov!nMonto > 0 Then
                         If RCapMov!cOpeCod = gAhoCargoServicioRecaudo Then
                            cCtaCodCargo = RCapMov!cCtaCod
                         End If
                 End If
                 RCapMov.MoveNext
             Loop
       
        RCapMov.MoveFirst
        pnITFAhoExt = 0#
        psMovNroAExt = strMovNro
        Do While Not RCapMov.EOF
            If RCapMov!nMonto > 0 And cCtaCodCargo = RCapMov!cCtaCod Then
                If Mid(RCapMov!cOpeCod, 1, 2) <> "99" Then
                    If RCapMov!cOpeCod = gAhoCargoServicioRecaudo Then
                        psOpeCodExtornoCap = RCapMov!cOpeCod
                        psCtaAhoExt = RCapMov!cCtaCod
                        pnMontoAhorroExt = RCapMov!nMonto
                        psOperacionDescExt = "Extorno Cargo a Cuenta por Servicio de Recaudo"
                        Call CapExtornoCargoAho(nMovNro, 0, gAhoExtCargoServicioRecaudo, RCapMov!cCtaCod, psMovNroCap, psOperacionDescExt, RCapMov!nMonto)
                    End If
                Else
                    Call CapExtornoCargoAho(nMovNro, 0, gITFCobroCargoExt, RCapMov!cCtaCod, psMovNroCap, "Extorno Cargo a Cuenta por Servicio de Recaudo", RCapMov!nMonto)
                    pnITFAhoExt = RCapMov!nMonto
                End If
            End If
            RCapMov.MoveNext
        Loop
      
    End If
    If cOpeCod = gExtornoDepositoRecaudo Then
     'CTI7 ERSOpeV2
            actualizarRelacionExtornoVoucherCaptacion (nMovNro)
        '************************
    End If
    'CTI6 ERS0112020
    'CTI1 ERS027-2019
    dbCmact.CommitTrans
    ExtornoRegistroServicioRecaudo = True
    Exit Function
    
errorRegistrarCobroServicioConvenio:
    dbCmact.RollbackTrans
    ExtornoRegistroServicioRecaudo = False
End Function
Public Sub RegistraComisionCliente(ByVal nMovNro As Long, _
                                        ByVal nMonto As Double, _
                                        ByVal nMonedaComision As Integer, _
                                        ByVal sMovNro As String)

    Dim nMovNro2 As Long
    AgregaMov sMovNro, gCargoClienteRecaudo, "COMISIÓN SERVICIO RECAUDO - CLIENTE", gMovEstContabMovContable, gMovFlagVigente
    nMovNro2 = GetnMovNro(sMovNro)
    AgregaMovOpeVarias nMovNro2, nMonto, "", "COMISIÓN SERVICIO RECAUDO - CLIENTE", nMonedaComision, 0
    AgregaMovRef nMovNro, nMovNro2
End Sub
Public Function RegistraOperacion(ByVal cCodConvenio As String) As Boolean
    Dim sSql As String
    sSql = "exec stp_upd_IngresarOperacionConvenio '"
    sSql = sSql & cCodConvenio & "'"
    dbCmact.Execute (sSql)
End Function
'END RIRO **************************************
'JUEZ 20150209 ***************************************
Public Function RegistrarServCobDebitoAuto(ByVal psCtaCod As String, ByVal pnMontoCom As Double, ByVal pnMovNro As Long, ByVal psMovNro As String) As Long
Dim rs As New ADODB.Recordset
Dim lsSQL As String
    
    RegistrarServCobDebitoAuto = 0
    lsSQL = "Exec stp_ins_RegistrarServCobDebitoAuto '" & psCtaCod & "', " & pnMontoCom & "," & pnMovNro & ", '" & psMovNro & "'"
    Set rs = dbCmact.Execute(lsSQL)
    If Not (rs.BOF And rs.EOF) Then
        RegistrarServCobDebitoAuto = rs!nCodId
    End If
End Function
Public Sub RegistrarServCobDebitoAutoDetServ(ByVal pnCodId As Long, ByVal psCodConvenio As String, ByVal psCodCliente As String, ByVal pnDia1 As Integer, ByVal pnDia2 As Integer, ByVal pnMontoMax As Double, ByVal psMovNro As String)
    Dim lsSQL As String
    lsSQL = "Exec stp_ins_RegistrarServCobDebitoAutoDetServ " & pnCodId & ",'" & psCodConvenio & "','" & psCodCliente & "'," & pnDia1 & "," & pnDia2 & "," & pnMontoMax & ",'" & psMovNro & "'"
    dbCmact.Execute (lsSQL)
End Sub
Public Sub RegistrarServCobDebitoAutoDetCred(ByVal pnCodId As Long, ByVal psCtaCodCred As String, ByVal pnMontoMax As Double, ByVal psMovNro As String)
    Dim lsSQL As String
    lsSQL = "Exec stp_ins_RegistrarServCobDebitoAutoDetCred " & pnCodId & ",'" & psCtaCodCred & "'," & pnMontoMax & ",'" & psMovNro & "'"
    dbCmact.Execute (lsSQL)
End Sub
Public Sub EliminarServCobDebitoAuto(ByVal psCtaCod As String)
    Dim lsSQL As String
    lsSQL = "Exec stp_upd_EliminarServCobDebitoAuto '" & psCtaCod & "'"
    dbCmact.Execute lsSQL
End Sub
Public Sub EliminarServCobDebitoAutoDet(ByVal pnCodId As Long)
    Dim lsSQL As String
    lsSQL = "Exec stp_upd_EliminarServCobDebitoAutoDet " & pnCodId
    dbCmact.Execute (lsSQL)
End Sub
Public Sub RegistrarCaptacServDebitoAutoLog(ByVal psCtaCod As String, ByVal psCodServPago As String, ByVal nMonto As Double, ByVal pnMontoMax As Double, _
                                                   ByVal pnComision As Double, ByVal pbPago As Boolean, ByVal psPagoDesc As String, ByVal psFecSis As Date)
Dim rs As New ADODB.Recordset
Dim lsSQL As String
    
    lsSQL = "Exec stp_ins_RegistrarCaptacServDebitoAutoLog '" & psCtaCod & "', '" & psCodServPago & "'," & nMonto & "," & pnMontoMax & "," _
                                                              & pnComision & "," & IIf(pbPago, 1, 0) & ",'" & psPagoDesc & "','" & _
                                                              Format(psFecSis, "yyyyMMdd") & "'"
    Set rs = dbCmact.Execute(lsSQL)
End Sub
Public Function ValidaComisionExtractoServCobDebitoAuto(ByVal psCtaCod As String, ByVal pdFecSis As Date) As Boolean
Dim lsSQL As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
    ValidaComisionExtractoServCobDebitoAuto = False
    
    lsSQL = "exec stp_sel_ValidaComisionExtractoServCobDebitoAuto '" & psCtaCod & "','" & Format(pdFecSis, "yyyyMM") & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    If oConecta.AbreConexion = False Then Exit Function
    Set rs = oConecta.CargaRecordSet(lsSQL)
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not rs.EOF And Not rs.BOF Then
        ValidaComisionExtractoServCobDebitoAuto = True
    End If
End Function
Public Sub InsertaMovCapServDebitoAuto(ByVal pnMovNro As Long, ByVal psCtaCod As String, ByVal pnMonto As Double, ByVal pnComision As Double, _
                                        ByVal pnTipoDebito As CaptacServDebitoAutomaticoTipo, Optional ByVal psCtaCodCred As String = "", _
                                        Optional psCodConvenio As String = "", Optional psCodCliente As String = "", Optional pbEjecBatch As Boolean = False)
    Dim lsSQL As String
    lsSQL = "stp_ins_MovCapServDebitoAuto " & pnMovNro & ",'" & psCtaCod & "','" & psCtaCodCred & "','" & psCodConvenio & "','" _
                                            & psCodCliente & "'," & pnMonto & "," & pnComision & "," & pnTipoDebito
    dbCmact.Execute (lsSQL)
End Sub
'END JUEZ ********************************************


'AMDO20150331 TI-ERS018-2015 ************************************************
Public Sub ExtornoTarjetasCuentaCancelada(ByVal pnMovNro As Long) 'AMDO20150331
    sSql = "exec DBTarjeta..stp_upd_ExtornoTajetaCuentaCancelada " & pnMovNro
    dbCmact.Execute sSql
End Sub
'***************************************************************************
'EJVG20151016 ***
Public Function RealizoInteresCash(ByVal psCtaCod As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
    
    lsSQL = "SELECT bRealizo = DBO.fnc_ERS0632014_RealizoInteresCash('" & psCtaCod & "')"
    Set rs = dbCmact.Execute(lsSQL)
    If Not (rs.BOF And rs.EOF) Then
        RealizoInteresCash = rs!bRealizo
    End If
    rs.Close
    Set rs = Nothing
End Function
'END EJVG *******
'JUEZ 20150922 ***************************************************************************
Public Function GeneraMovNro(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "01", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String
    Dim rs As New ADODB.Recordset
    Dim sql As String

    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & Format(Time, "hh:mm:ss"), "mm/dd/yyyy hh:mm:ss") & "','" & Right(psCodAge, 2) & "','" & psUser & "'"
    Else
       sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    Set rs = dbCmact.Execute(sql)
    If Not rs.EOF Then
        GeneraMovNro = rs.Fields(0)
    End If
    Set rs = Nothing
    Exit Function
End Function
Public Sub AgregaMovCompraVenta(ByVal pnMovNro As Long, ByVal psPersCod As String, ByVal pnImporte As Currency)
    sSql = " INSERT INTO MovCompraVenta(nMovNro, cPersCod, nMovImporte )" _
        & " VALUES(" & pnMovNro & ",'" & psPersCod & "'," & pnImporte & ") "
      
    dbCmact.Execute sSql
End Sub
Public Sub AgregaMovTpoCambio(ByVal pnMovNro As Long, ByVal pnMovTpoCambio As Currency, _
Optional ByVal pnValVent As Currency = 0, _
Optional ByVal pnValComp As Currency = 0, _
Optional ByVal pnValFijo As Currency = 0, _
Optional ByVal pnValFijoDia As Currency = 0, _
Optional ByVal pnValVentEsp As Currency = 0, _
Optional ByVal pnValCompEsp As Currency = 0, _
Optional ByVal pnValPond As Currency = 0, _
Optional ByVal pnValPondVenta As Currency = 0, _
Optional ByVal pnValPondREU As Currency = 0, _
Optional ByVal pnValSBSDia As Currency = 0, _
Optional ByVal pnValCompTr As Currency = 0, _
Optional ByVal pnValVentTr As Currency = 0)
    
    sSql = "INSERT INTO MovTpoCambio (nMovNro, nMovTpoCambio,nValVent,nValComp,nValFijo,nValFijoDia,nValVentEsp,nValCompEsp,nValPond,nValPondVenta,nValPondREU,nValSBSDia,nValCompTr,nValVentTr) " _
     & "VALUES (" & pnMovNro & ", " & pnMovTpoCambio & ", " & pnValVent & ", " & pnValComp & "," & pnValFijo & ", " & pnValFijoDia & ", " & pnValVentEsp & ", " & pnValCompEsp & ", " & pnValPond & ", " & pnValPondVenta & ", " & pnValPondREU & ", " & pnValSBSDia & "," & pnValCompTr & ", " & pnValVentTr & ")"
     
    dbCmact.Execute sSql
End Sub
'END JUEZ ******************************************************************************
'RECO20151123 ERS073-2015 ***************************************************
Public Sub AgregaSegSepelioAfiliacion(ByVal psNumCert As String, ByVal psCtaCod As String, ByVal pdFecSis As String, ByVal psMovNro As String, _
                                    ByVal pnMovNro As Long, ByVal psPersCod As String, ByVal psAgeCod As String, ByVal pbPrimaAnual As Integer, _
                                    Optional ByVal cUser As String = "", Optional ByVal cUserVenta As String = "") 'CTI2 FERIMORO ERS034-2019
'APRI20181023 ERS071-2018 ADD pbPrimaAnual

    Dim oConec As COMConecta.DCOMConecta
    Dim sHora As String
    Dim lsSQL As String
    Dim dFecSis As String   'ADDED PAVA20210907
        dFecSis = CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4))  'ADDED PAVA20210907
        sHora = Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2)
        'CTI2 FERIMORO ERS034-2019 :   31082020
        'lsSQL = "Exec stp_ins_SegSepelioAfiliacion '" & psNumCert & "','" & psCtaCod & "','" & _
        '        Format(pdFecSis & " " & sHora, "mm/dd/yyyy hh:mm:ss") & "'," & pnMovNro & ",'" & psPersCod & "','" & psAgeCod & "'," & pbPrimaAnual
    
         lsSQL = "Exec stp_ins_SegSepelioAfiliacion '" & psNumCert & "','" & psCtaCod & "','" & _
                Format(dFecSis & " " & sHora, "mm/dd/yyyy hh:mm:ss") & "'," & pnMovNro & ",'" & psPersCod & "','" & psAgeCod & "'," & pbPrimaAnual & ",'" & _
                cUser & "','" & cUserVenta & "'"
       
        dbCmact.Execute (lsSQL)
End Sub
Public Sub AgregaSegSepelioAfiliacionHis(ByVal psNumCert As String, ByVal psCtaCod As String, ByVal pdFecSis As String, ByVal psMovNro As String, ByVal pnMovNro As Long, ByVal psPersCod As String, psAgeCod As String, ByVal pnSegEstado As Integer)
    Dim oConec As COMConecta.DCOMConecta
    Dim sHora As String
    Dim lsSQL As String
    Dim dFecSis As String   'ADDED PAVA20210907
        dFecSis = CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4))  'ADDED PAVA20210907
        sHora = Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2)
        lsSQL = "Exec stp_ins_SegSepelioAfiliacionHis '" & psNumCert & "','" & psCtaCod & "','" & Format(dFecSis & " " & sHora, "mm/dd/yyyy hh:mm:ss") & "'," & pnMovNro & "," & pnSegEstado
        dbCmact.Execute (lsSQL)
End Sub
Public Sub SepelioActualizaEstadoHis(ByVal pnMovNro As Long, ByVal pnEstado As Integer)
    Dim oConec As COMConecta.DCOMConecta
    Dim sHora As String
    Dim lsSQL As String
        lsSQL = "Exec stp_upd_SepelioActualizaEstadoHis " & pnMovNro & "," & pnEstado
        dbCmact.Execute (lsSQL)
End Sub
'RECO FIN *******************************************************************

'RIRO20160109 **********************************
Public Function ExtornoPagoCC(ByVal pnMovNro As Long) As Boolean
    
    Dim sSql As String
    Dim rsCCorresponsal As ADODB.Recordset
    Dim bResult As Boolean
    Dim nResp As Integer
        
    Set rsCCorresponsal = New ADODB.Recordset
    bResult = False
    nResp = 0
    sSql = "exec stp_upd_ExtornoPagoCC " & pnMovNro
    
    rsCCorresponsal.CursorLocation = adUseClient
    rsCCorresponsal.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsCCorresponsal.ActiveConnection = Nothing
    
    If Not rsCCorresponsal Is Nothing Then
        If Not rsCCorresponsal.BOF And Not rsCCorresponsal.EOF Then
            If rsCCorresponsal.RecordCount > 0 Then
                If rsCCorresponsal!nResp = 1 Then
                    bResult = True
                ElseIf rsCCorresponsal!nResp = 0 Then
                    bResult = False
                End If
            End If
        End If
    End If
    ExtornoPagoCC = bResult
    
End Function
Public Sub ObtieneDatosVoucherCajeroCorresponsal(ByVal pnMovNro As Long, _
                                                 ByRef psTituloOpe As String, _
                                                 ByRef psOperador As String, _
                                                 ByRef psCodOperador As String, _
                                                 ByRef pnMonto As Currency, _
                                                 ByRef pnMoneda As Integer)
    Dim sSql As String
    Dim rsCCorresponsal As ADODB.Recordset
    Set rsCCorresponsal = New ADODB.Recordset
    
    sSql = "exec stp_sel_obtenerDatosVoucher " & pnMovNro
    
    rsCCorresponsal.CursorLocation = adUseClient
    rsCCorresponsal.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not rsCCorresponsal Is Nothing Then
        If Not rsCCorresponsal.BOF And Not rsCCorresponsal.EOF Then
            If rsCCorresponsal.RecordCount > 0 Then
                psTituloOpe = rsCCorresponsal!cTituloOpe
                psOperador = rsCCorresponsal!cOperador
                psCodOperador = rsCCorresponsal!cCodOperador
                pnMonto = rsCCorresponsal!nMonto
                pnMoneda = rsCCorresponsal!nMoneda
            End If
        End If
    End If
    Set rsCCorresponsal = Nothing
    
End Sub
Public Function ObtieneDatosCajeroCorresponsal(ByVal sCodigoCC As String) As ADODB.Recordset
    Dim sSql As String
    Dim rsCCorresponsal As New ADODB.Recordset
    
    sSql = "exec stp_sel_ObtenerDatosCC '" & sCodigoCC & "'"
    
    rsCCorresponsal.CursorLocation = adUseClient
    rsCCorresponsal.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsCCorresponsal.ActiveConnection = Nothing
    Set ObtieneDatosCajeroCorresponsal = rsCCorresponsal
    Set rsCCorresponsal = Nothing
    
End Function
Public Function PagoRecaudoCajeroCorresponsal(ByVal pnIdCC As String, ByVal pnMonto As Double) As ADODB.Recordset
    Dim sSql As String
    Dim rsCCorresponsal As New ADODB.Recordset
    
    sSql = "exec stp_upd_PagoRecaudoCajeroCorresponsal " & pnIdCC & ", " & pnMonto
    
    rsCCorresponsal.CursorLocation = adUseClient
    rsCCorresponsal.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsCCorresponsal.ActiveConnection = Nothing
    Set PagoRecaudoCajeroCorresponsal = rsCCorresponsal
    Set rsCCorresponsal = Nothing
End Function
'END RIRO **************************************
'JUEZ 20151229 ************************************************************************
Public Sub GrabarOpeComisionConteoMoneda(ByVal pnMovNro As Long, ByVal psEfectivoCod As String, ByVal pnMonto As Double)

    sSql = "Exec stp_ins_OpeComisionConteoMoneda " & pnMovNro & ",'" & psEfectivoCod & "'," & pnMonto
     
    dbCmact.Execute sSql
End Sub
'END JUEZ *****************************************************************************
'PASI20160610 CCE**********************************************************************
Public Function CCE_InsertaCheque(ByVal pnId As Integer, ByVal psNombreBcoDest As String, nMisMoTit As Integer, ByVal pnDocTpoBenef As Integer, ByVal psDocNroBenef As String) As Long
Dim lsSQL As String
Dim rs As ADODB.Recordset
On Error GoTo ErrInsertaChequeCCE
    lsSQL = "exec stp_ins_CCE_DocRec " & pnId & ",'" & psNombreBcoDest & "'," & nMisMoTit & "," & pnDocTpoBenef & ",'" & psDocNroBenef & "'"
    rs.CursorLocation = adUseClient
    rs.Open lsSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        CCE_InsertaCheque = rs!nIdDRC
    End If
    Set rs = Nothing
Exit Function
ErrInsertaChequeCCE:
    Err.Raise Err.Number, "DMov:InsertaChequeCCE Method", Err.Description
    'Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaChequeCCE Method")
End Function
Public Function CCE_RegistraTransferencia(ByVal psOpeTransf As String, ByVal psCodBCRBcoDest As String, ByVal psPersCodOrd As String _
                                            , ByVal psPersNombreOrd As String, ByVal pnTpoDocOrd As Integer _
                                            , ByVal psNroDocOrd As String, ByVal psDireccionOrd As String, ByVal psTelefOrd As String, ByVal pnMoneda As Integer _
                                            , ByVal pnMontoTotal As Currency) As Long
Dim rs As ADODB.Recordset
    sSql = "stp_ins_CCE_RegistrarTransferencia '" & psOpeTransf & "','" & psCodBCRBcoDest & "','" & psPersCodOrd & "','" & psPersNombreOrd & "'," _
        & pnTpoDocOrd & ",'" & psNroDocOrd & "','" & psDireccionOrd & "','" & psTelefOrd & "'," & pnMoneda & "," & pnMontoTotal
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        CCE_RegistraTransferencia = rs!nTranNro
    End If
    Set rs = Nothing
End Function
Public Function CCE_RegistraTransfOrdinaria(ByVal nTranNro As Long, ByVal psCCIDestCod As String _
                                            , ByVal pnCargoCta As Integer, ByVal psCtaCod As String, ByVal psCCIOrigCod As String _
                                            , ByVal pnMismoTit As Integer, ByVal pnMontoIncComis As Integer, ByVal pnMonto As Currency _
                                            , ByVal psTariCod As String, ByVal psTariCriApl As String, ByVal psComisSigno As String, ByVal pnComisMonto As Currency)

    sSql = "stp_ins_CCE_RegistrarTransfOrdinaria " & nTranNro & ",'" & psCCIDestCod & "'," _
            & pnCargoCta & ",'" & psCtaCod & "','" & psCCIOrigCod & "'," & pnMismoTit & "," & pnMontoIncComis & "," & pnMonto & "," _
            & "'" & psTariCod & "','" & psTariCriApl & "','" & psComisSigno & "'," & pnComisMonto
    dbCmact.Execute sSql
End Function
'VAPA20170628
Public Function CCE_RegistraTransfOrdinariaSabado(ByVal nTranNro As Long, ByVal psCCIDestCod As String _
                                            , ByVal pnCargoCta As Integer, ByVal psCtaCod As String, ByVal psCCIOrigCod As String _
                                            , ByVal pnMismoTit As Integer, ByVal pnMontoIncComis As Integer, ByVal pnMonto As Currency _
                                            , ByVal psTariCod As String, ByVal psTariCriApl As String, ByVal psComisSigno As String, ByVal pnComisMonto As Currency, Optional ByVal psSesion As String, Optional ByVal pdFecha As Date)

    sSql = "stp_ins_CCE_RegistrarTransfOrdinaria " & nTranNro & ",'" & psCCIDestCod & "'," _
            & pnCargoCta & ",'" & psCtaCod & "','" & psCCIOrigCod & "'," & pnMismoTit & "," & pnMontoIncComis & "," & pnMonto & "," _
            & "'" & psTariCod & "','" & psTariCriApl & "','" & psComisSigno & "'," & pnComisMonto & ",'" & psSesion & "','" & Format(pdFecha, "YYYYmmdd") & "'"
    dbCmact.Execute sSql
End Function
'VAPA
Public Function CCE_RegistraTransfPagoHaberes(ByVal nTranNro As Integer, ByVal psCCIDestCod As String, ByVal psPersCodBenef As String _
                                            , ByVal psPersNombreBenef As String, ByVal pnTpoDocBenef As Integer, ByVal psNroDocBenef As String, ByVal psDireccionBenef As String _
                                            , ByVal psTelefBenef As String, ByVal pnTpoHaberCod As Integer _
                                            , ByVal pnMonto As Currency _
                                            , ByVal psTariCod As String, ByVal psTariCriApl As String, ByVal psComisSigno As String, ByVal pnComisMonto As Currency)
                                            
    sSql = "stp_ins_CCE_RegistrarTransfPagoHaberes " & nTranNro & ",'" & psCCIDestCod & "','" & psPersCodBenef & "','" & psPersNombreBenef & "'," _
                & pnTpoDocBenef & ",'" & psNroDocBenef & "','" & psDireccionBenef & "','" & psTelefBenef & "'," & pnTpoHaberCod & "," & pnMonto & "," _
                & "'" & psTariCod & "','" & psTariCriApl & "','" & psComisSigno & "'," & pnComisMonto
    dbCmact.Execute sSql
End Function
Public Function CCE_RegistraTransfPagoCTS(ByVal nTranNro As Long, ByVal psCCIDestCod As String, ByVal psPersCodBenef As String _
                                        , ByVal psPersNombreBenef As String, ByVal pnTpoDocBenef As Integer, ByVal psNroDocBenef As String, ByVal psDireccionBenef As String _
                                        , ByVal psTelefBenef As String, ByVal psMesPago As String _
                                        , ByVal psAnio As String, ByVal pnImporteNSueldosBrutos As Currency, ByVal pnMonto As Currency _
                                        , ByVal psTariCod As String, ByVal psTariCriApl As String, ByVal psComisSigno As String, ByVal pnComisMonto As Currency)
    
    sSql = "stp_ins_CCE_RegistrarTransfPagoCTS " & nTranNro & ",'" & psCCIDestCod & "','" & psPersCodBenef & "','" & psPersNombreBenef & "'," _
                & pnTpoDocBenef & ",'" & psNroDocBenef & "','" & psDireccionBenef & "','" & psTelefBenef & "','" & psMesPago & "','" & psAnio & "'," & pnImporteNSueldosBrutos & "," & pnMonto & "," _
                & "'" & psTariCod & "','" & psTariCriApl & "','" & psComisSigno & "'," & pnComisMonto
    dbCmact.Execute sSql
End Function
Public Function CCE_RegistraTransfPagTarjCred(ByVal nTranNro As Long, ByVal psNumTarj As String, ByVal psPersCodBenef As String _
                                             , ByVal psPersNombreBenef As String, ByVal pnTpoDocBenef As Integer, ByVal psNroDocBenef As String, ByVal psDireccionBenef As String _
                                            , ByVal psTelefBenef As String, ByVal pnMontoIncComis, ByVal pnMonto As Currency _
                                            , ByVal psTariCod As String, ByVal psTariCriApl As String, ByVal psComisSigno As String, ByVal pnComisMonto As Currency)
    
    sSql = "stp_ins_CCE_RegistrarTransfPagTarjCred " & nTranNro & ",'" & psNumTarj & "','" & psPersCodBenef & "','" & psPersNombreBenef & "'," _
                & pnTpoDocBenef & ",'" & psNroDocBenef & "','" & psDireccionBenef & "','" & psTelefBenef & "'," & pnMontoIncComis & "," & pnMonto & "," _
                & "'" & psTariCod & "','" & psTariCriApl & "','" & psComisSigno & "'," & pnComisMonto
    dbCmact.Execute sSql
End Function
Public Function CCE_RegistraTransfPagoProveedor(ByVal nTranNro As Long, ByVal psCCIDestCod As String, ByVal psPersCodBenef As String _
                                                , ByVal psPersNombreBenef As String, ByVal pnTpoDocBenef As Integer, ByVal psNroDocBenef As String, ByVal psDireccionBenef As String _
                                                , ByVal psTelefBenef As String, ByVal pnTpoDocComp As Integer _
                                                , ByVal psNroDocComp As String, ByVal pdFecEmisionComp As Date, ByVal pnMonto As Currency _
                                                , ByVal psTariCod As String, ByVal psTariCriApl As String, ByVal psComisSigno As String, ByVal pnComisMonto As Currency)
    
    sSql = "stp_ins_CCE_RegistrarTransfPagoProveedor " & nTranNro & ",'" & psCCIDestCod & "','" & psPersCodBenef & "','" & psPersNombreBenef & "'," _
                                                    & pnTpoDocBenef & ",'" & psNroDocBenef & "','" & psDireccionBenef & "','" & psTelefBenef & "'," & pnTpoDocComp & "," _
                                                    & "'" & psNroDocComp & "','" & Format(pdFecEmisionComp, "YYYYmmdd") & "', " & pnMonto & "," _
                                                    & "'" & psTariCod & "','" & psTariCriApl & "','" & psComisSigno & "'," & pnComisMonto
    dbCmact.Execute sSql
End Function
Public Function CCE_RegistraMovOpeVarias(ByVal pnMovNro As Long, ByVal pnMovImporte As Currency, ByVal pnMoneda As Integer, ByVal psOpeCod As String, ByVal pnConcepto As Integer)
    
    sSql = "stp_ins_CCE_RegistrarMovOpeVarias " & pnMovNro & "," & pnMovImporte & "," & pnMoneda & ",'" & psOpeCod & "'," & pnConcepto
    dbCmact.Execute sSql
End Function
Public Function CCE_CargoCta_OffHost(ByVal pdFecha As Date, ByVal psAgeCod As String, ByVal psCtaCod As String, ByVal pnMonto As Currency, ByVal pnComisMonto As Currency, ByVal pnMontoITF As Currency, ByVal pnMontoComision As Currency, _
                                    ByVal pnMoneda As Integer, ByVal psOpeCod As String, ByVal psOpeCodComCCE As String, ByVal psOpeCodComision As String, ByVal psOpeCodITF As String, _
                                    ByVal psUser As String) As Long
                                    'VAPA20170119 CCE se agrego psOpeCodComCCE

Dim rs As New ADODB.Recordset
    sSql = "stp_ins_CCE_CargoCta_OffHost '" & Format(pdFecha, "YYYYmmdd") & "','" & psAgeCod & "','" & psCtaCod & "'," & pnMonto & "," & pnComisMonto & "," & pnMontoITF & "," & pnMontoComision & "," _
                                            & pnMoneda & ",'" & psOpeCod & "','" & psOpeCodComCCE & "','" & psOpeCodComision & "','" & psOpeCodITF & "'," _
                                            & "'" & psUser & "'"
                                     'VAPA20170119 CCE se agrego psOpeCodComCCE
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    
    CCE_CargoCta_OffHost = IIf(Not (rs.EOF And rs.BOF), rs!pnResultado, 0)
End Function
'VAPA20170612
Public Function CCE_CargoCta(ByVal pdFecha As Date, ByVal psAgeCod As String, ByVal psCtaCod As String, ByVal pnMonto As Currency, ByVal pnComisMonto As Currency, ByVal pnMontoITF As Currency, ByVal pnMontoComision As Currency, _
                                    ByVal pnMoneda As Integer, ByVal psOpeCod As String, ByVal psOpeCodComCCE As String, ByVal psOpeCodComision As String, ByVal psOpeCodITF As String, _
                                    ByVal psUser As String, ByVal pnMovNro As Long, ByVal psMovNro As String) As Long
                                    

Dim rs As New ADODB.Recordset
    sSql = "stp_ins_CCE_CargoCta '" & Format(pdFecha, "YYYYmmdd") & "','" & psAgeCod & "','" & psCtaCod & "'," & pnMonto & "," & pnComisMonto & "," & pnMontoITF & "," & pnMontoComision & "," _
                                            & pnMoneda & ",'" & psOpeCod & "','" & psOpeCodComCCE & "','" & psOpeCodComision & "','" & psOpeCodITF & "'," _
                                            & "'" & psUser & "'," & pnMovNro & ",'" & psMovNro & "'"
                                    
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    
    CCE_CargoCta = IIf(Not (rs.EOF And rs.BOF), rs!pnResultado, 0)
End Function
Public Function CCE_ExtornaTramaPesentados(ByVal pnMovNroDeExtorno As Long)
    Dim sql As String
    sSql = "stp_upd_CCE_ExtornaTramaPre " & pnMovNroDeExtorno
    dbCmact.Execute sSql
End Function
'END VAPA20170612
Public Function CCE_ExtornaTransf(ByVal pnMovNroDeExtorno As Long, ByVal pnMovNroRef As Long)
    Dim sql As String
    sSql = "stp_upd_CCE_ExtornaTransferencia " & pnMovNroDeExtorno & "," & pnMovNroRef
    dbCmact.Execute sSql
End Function
Public Function CCE_RegistrarMovRef(ByVal nMovNro As Long, ByVal pnTranNroRef As Long)
    sSql = "stp_ins_CCE_RegistrarMovRef " & nMovNro & "," & pnTranNroRef
    dbCmact.Execute sSql
End Function
Public Function CCE_RegistrarMovOffHostRef(ByVal nMovNro As Long, ByVal pnTranNroRef As Long)
    sSql = "stp_ins_CCE_RegistrarMovOffHostRef " & nMovNro & "," & pnTranNroRef
    dbCmact.Execute sSql
End Function
Public Function CCE_RegistrarRecRef(ByVal nMovNro As Long, ByVal pnMovNroRef As Long, ByVal psEfinCodBCR As String)
    sSql = "stp_ins_CCE_RegistrarRecRef " & nMovNro & "," & pnMovNroRef & ",'" & psEfinCodBCR & "'"
    dbCmact.Execute sSql
End Function
'end PASI
'*** CTI3 (ferimoro) 16102018
Public Sub dInsertDetExtMov(ByVal pnMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        Optional ByVal psMotExtorno As String, _
        Optional ByVal psDetMotExtorno As String)
Dim psCuenta As String
psCuenta = ""
    sSql = "INSERT MovDetExtornos (nMovNro,cOpedCod,cCtaCod,cMotExtorno,cDetMotExto) " _
          & "VALUES (" & pnMovNro & ",'" & nOperacion & "','" & psCuenta & "','" & psMotExtorno & "','" & UCase(psDetMotExtorno) & "')"

    dbCmact.Execute sSql
End Sub
Public Function ExisteNroMov(ByVal pnMovNro As Long) As Boolean
Dim rs As ADODB.Recordset

ExisteNroMov = False
sSql = "  select * from MovDetExtornos where nMovNro=" & pnMovNro & ""
       
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open sSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
If rs.RecordCount > 0 Then ExisteNroMov = True
If rs.RecordCount = 0 Then ExisteNroMov = False

Set rs.ActiveConnection = Nothing
Set rs = Nothing

End Function
'***
'*********************add pti1 23-01-19 ERS081-2018 ACTA 06-2019 **********************
Public Function GetEstadoCuentaCCI(ByVal sCuenta As String) As Boolean
    Dim sSql As String
    Dim rsEst As ADODB.Recordset

    sSql = "EXEC SP_SEL_ESTADOCONFIGURACCI_PRO '" & sCuenta & "'"
    
       Set rsEst = New ADODB.Recordset
    rsEst.CursorLocation = adUseClient
    rsEst.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsEst.ActiveConnection = Nothing
    If Not (rsEst.EOF And rsEst.BOF) Then
        
        If rsEst!bEstado = 1 Then
            GetEstadoCuentaCCI = True
        Else
            GetEstadoCuentaCCI = False
        End If
        
    Else
        GetEstadoCuentaCCI = False
    End If
    
    rsEst.Close
    Set rsEst = Nothing
End Function
'*********************END pti1 23-01-19 ERS081-2018 ACTA 06-2019 **********************

'CTI2 20190408 INI *****
Public Function ValidaExtornoLote(ByVal nMovNro As Long, _
                                  Optional ByVal sCtaCod As String = "") As Boolean

Dim sSql As String
Dim rs As New ADODB.Recordset
Dim brespuesta As Boolean
brespuesta = False

sSql = "exec Stp_Sel_ValidaExtorno " & nMovNro & ", '" & sCtaCod & "'"
Set rs = dbCmact.Execute(sSql)

If Not rs Is Nothing Then
    If Not rs.EOF Then
        If rs.RecordCount > 0 Then
            brespuesta = rs!bResultado
        End If
    End If
End If
Set rs = Nothing
ValidaExtornoLote = brespuesta

End Function
'CTI2 FIN ****

'CTI1 ERS027-2019 ELOR begin
Private Function StringcobroServRecaudo(ByVal arrCadena As Variant, ByVal psCodConvenio As String, _
                                        ByVal psCodCliente As String, ByVal pnTipoDoi As String, _
                                        ByVal psDoi As String, ByVal psNombreCliente As String, _
                                        ByVal pdFechaCobro As String, ByVal psMovNro As String, _
                                        ByVal pnComiEmp As String, ByVal pnComiCli As String, _
                                        Optional ByVal sCodigoCobranzaRes As String = "0") As String
Dim sSql As String
sSql = "exec stp_ins_upd_RegistrarCobroServicioRecaudo "
    'Codigo de convenio
    sSql = sSql & "'" & psCodConvenio & "', "
    'Id de la trama
    sSql = sSql & "'" & arrCadena(0) & "',"
    'Codigo de cliente
    sSql = sSql & "'" & psCodCliente & "',"
    'Tipo de Doi
    sSql = sSql & pnTipoDoi & ","
    'Doi
    sSql = sSql & "'" & psDoi & "',"
    'Nombre de Cliente
    sSql = sSql & "'" & psNombreCliente & "',"
    'Servicio
    sSql = sSql & "'" & arrCadena(1) & "',"
    'Concepto
    sSql = sSql & "'" & arrCadena(2) & "',"
    'Importe
    sSql = sSql & arrCadena(3) & ","
    'Deuda Actual
    sSql = sSql & arrCadena(4) & ","
    'Monto Cobro
    sSql = sSql & arrCadena(5) & ","
    'Estado
    sSql = sSql & "'" & arrCadena(6) & "',"
    'FechaCobro
    sSql = sSql & "'" & Format$(pdFechaCobro, "yyyymmdd") & "',"
    'cMovNro
    sSql = sSql & "'" & psMovNro & "',"
    'Comision Empresa
    sSql = sSql & CDbl(pnComiEmp) & ","
    'Comision Cliente
    sSql = sSql & CDbl(pnComiCli) & ","
    'Mora
    sSql = sSql & CDbl(arrCadena(7)) & ","
    'Fecha Vencimiento
    sSql = sSql & "'" & arrCadena(8) & "',"
    'Codigo Cobranza
    sSql = sSql & "'" & sCodigoCobranzaRes & "'"
    StringcobroServRecaudo = sSql
End Function
'CTI1 ERS027-2019 ELOR end
'CTI4 ERS0112020
Public Function CapCargoCuentaAho(ByRef pMatAho As Variant, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, ByVal pdFecSis As Date, _
            Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional bOPExiste As Boolean = False, _
            Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
            Optional sCodCmac As String = "", Optional sNomCmac As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "", Optional bCancelaciones As Boolean = False, _
            Optional pnMonedaCred As Moneda = 0, _
            Optional pbITFAplica As Boolean = True, Optional pnITFValor As Double = 0, Optional pbITFAsumido As Boolean = False, _
            Optional psITFOperacion As String = "", Optional pbRetiroEfectivo As Boolean = False, _
            Optional ByVal pnTipoPago As Integer = 1) As Double
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

'Dim clsMov As COMNContabilidad.NCOMContFunciones

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
        'CTI6 ERS0112020
        Select Case pnTipoPago
         Case gCVTipoPagoEfectivo   ' Efectivo
                If nOperacion = "100949" Then
                    sMsgOpe = sMensajeReactiva
                Else
                    sMsgOpe = "Retiro Efectivo"
                End If
           Case gCVTipoPagoCargoCta
            If nOperacion = gAhoCargoServicioRecaudo Then
                sMsgOpe = "Retiro OP Recaudo Cargo a Cuenta"
            Else
                sMsgOpe = "Retiro Cargo a Cuenta"
            End If
            

        End Select
        'END CTI6
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
        AgregaMovCap nMovNro, psITFOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro, psITFOperacion, sCuenta, gConcITFAsumido, pnITFValor
    Else
        ActualizaCargoCaptacion sCuenta, pnITFValor, pnITFValor, nIntGanado, dUltMov, sMovNro, bNumExtracto
        AgregaMovCap nMovNro, psITFOperacion, sCuenta, pnITFValor, nSaldoDisp - pnITFValor, nSaldoCnt - pnITFValor
        AgregaMovCapDet nMovNro, psITFOperacion, sCuenta, gConcITFCliente, pnITFValor
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
End Function
Public Function GetDatosCuentaAho(ByVal sCuenta As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

sSql = "Exec sp_sel_DatosCuentaAhorro '" & sCuenta & "' , " & gCaptacEstado & " , " & gProductoCuentaTipo & " , " & gCaptacTipoTasa

Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaAho = rsCta
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
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaCTS = rsCta
Set rsCta = Nothing

End Function
Public Function GetInteres(ByVal nCapital As Double, ByVal nTasa As Double, _
            ByVal nPlazo As Long, Optional nTipoInteres As TipoCalculoInteres = TpoCalcIntSimple) As Double
If nTipoInteres = TpoCalcIntSimple Then
    GetInteres = Round((nTasa / 36000) * nPlazo * nCapital, 2)
ElseIf nTipoInteres = TpoCalcIntCompuesto Then
    GetInteres = Round((((nTasa / 36000) + 1) ^ nPlazo - 1) * nCapital, 2)
End If
End Function
Public Sub AgregaOrdenPagoEstado(ByVal sCuenta As String, ByVal sNroDoc As String, _
        ByVal sMovNro As String, ByVal nMonto As Double, ByVal nEstadoOP As CaptacOrdPagoEstado)
Dim sSql As String
sSql = "UPDATE DocRecOP Set nMonto = " & nMonto & " WHERE nTpoDoc = " & TpoDocOrdenPago & " " _
    & "AND cNroDoc = '" & sNroDoc & "' And cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
sSql = "INSERT DocRecOPEst (nTpoDoc,cNroDoc,cCtaCod,cMovNro,nMonto,nEstado) " _
    & "VALUES (" & TpoDocOrdenPago & ",'" & sNroDoc & "','" & sCuenta & "','" & sMovNro & "'," & nMonto & "," & nEstadoOP & ")"
dbCmact.Execute sSql
End Sub
Public Function GetCapParametro(ByVal nParametro As CaptacParametro) As Double
Dim rsVar As ADODB.Recordset
Dim sSql As String
sSql = "SELECT nParValor FROM Parametro WHERE nParCod = " & nParametro & " " _
    & "And nParProd = " & gPrdParamCaptac
Set rsVar = New ADODB.Recordset
rsVar.CursorLocation = adUseClient
rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsVar.ActiveConnection = Nothing
If rsVar.EOF And rsVar.BOF Then
    GetCapParametro = 0
Else
    GetCapParametro = rsVar("nParValor")
End If
rsVar.Close
Set rsVar = Nothing
End Function

'CTI2 FERIMORO : ERS030-2019 05082020
Public Function ObtieneCoberturaGarantiaDPF(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim rsCobertura As New ADODB.Recordset
    
    sSql = "exec stp_sel_ERS030_2019_ObtenerCoberturaGarantiaDPF '" & psCtaCod & "'"
    
    rsCobertura.CursorLocation = adUseClient
    rsCobertura.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsCobertura.ActiveConnection = Nothing
    Set ObtieneCoberturaGarantiaDPF = rsCobertura
    Set rsCobertura = Nothing
    
End Function
'CTI2 FERIMORO : ERS030-2019 10082020
Public Function insCalendarioIntDPFMensual(ByVal cCtaCod As String, ByVal nInte As Double, ByVal fIntePago As Date, ByVal esta As Integer)
Dim sSql As String
sSql = "exec stp_ins_ERS030_2019_CalendarioRetiroIntPF '" & cCtaCod & "'," & nInte & ",'" & Format(fIntePago, "YYYYmmdd") & "','" & esta & "'"
dbCmact.Execute sSql

End Function

'CTI2 FERIMORO : ERS034-2019 25082020
Public Function SeguroCronogramaAfiliacion(ByVal tpoAfiliacion As Integer, ByVal nFecha As Date, _
                                           ByVal nTpoSeguro As Integer, ByVal cNumCertificado As String, ByVal nMovNro As Long)
Dim sSql As String
sSql = "exec stp_ins_cns_ERS034_2019_SegurosCronograma " & tpoAfiliacion & ",'" & Format(nFecha, "YYYYmmdd") & "'," & nTpoSeguro & "," & nMovNro & ",'" & cNumCertificado & "'"
dbCmact.Execute sSql

End Function

'CTI2 FERIMORO: ERS034-2019 25082020
Public Function cnsFechaSeguroCronograma(ByVal cDocumento As String) As String

Dim sSql As String
Dim rs As New ADODB.Recordset
Dim brespuesta As String

sSql = "exec stp_cns_ERS034_2019_SeguroCronogramaVigente '" & cDocumento & "'"
Set rs = dbCmact.Execute(sSql)

If Not rs Is Nothing Then
    If Not rs.EOF Then
        If rs.RecordCount > 0 Then
            brespuesta = rs!fecha
        End If
    End If
End If
Set rs = Nothing
cnsFechaSeguroCronograma = brespuesta

End Function

'CTI2 FERIMORO: ERS034-2019 25082020
Public Function cnsVeriPagoAnualSeg(ByVal cDocumento As String, ByVal sFecha As Date) As String

Dim sSql As String
Dim rs As New ADODB.Recordset
Dim brespuesta As String

sSql = "exec stp_cns_ERS034_2019_VerificaSegCronoAnual '" & cDocumento & "','" & Format(sFecha, "YYYYmmdd") & "'"
Set rs = dbCmact.Execute(sSql)

If Not rs Is Nothing Then
    If Not rs.EOF Then
        If rs.RecordCount > 0 Then
            brespuesta = rs!bMensaje
        End If
    End If
End If
Set rs = Nothing
cnsVeriPagoAnualSeg = brespuesta

End Function

'CTI2 FERIMORO : ERS034-2019 25082020
Public Function veriCroSegSepAnual(ByVal cDocumento As String, ByVal sFecha As Date)
Dim sSql As String
sSql = "exec stp_ins_ERS034_2019_SegSepAnualCronoVerifica '" & cDocumento & "','" & Format(sFecha, "YYYYmmdd") & "'"
dbCmact.Execute sSql

End Function

'CTI2 FERIMORO : ERS034-2019 25082020
'Public Function cnsVerificaTiempoGraciaD(ByVal cDocumento As String, ByVal sFecha As Date) As Integer
'
'Dim sSql As String
'Dim rs As New ADODB.Recordset
'Dim brespuesta As String
'
'sSql = "exec stp_cns_ERS034_2019_VerificaTiempoGraciaSegSepelio '" & cDocumento & "','" & Format(sFecha, "YYYYmmdd") & "'"
'Set rs = dbCmact.Execute(sSql)
'
'If Not rs Is Nothing Then
'    If Not rs.EOF Then
'        If rs.RecordCount > 0 Then
'            brespuesta = rs!nApagar
'        End If
'    End If
'End If
'Set rs = Nothing
'cnsVerificaTiempoGraciaD = brespuesta
'End Function

'CTI2 FERIMORO : ERS034-2019 25082020
Public Sub AgregaSegSoat(ByVal psNumCert As String, ByVal psPersNombre As String, psPersCod As String, _
                         ByVal psUso As String, ByVal psUsuario As String, ByVal psPrima As Double, _
                         ByVal pnMovNro As Long, psAgeCod As String, ByVal pnSegEstado As Integer, ByVal pdFecSis As String, ByVal psDoi As String)
    Dim oConec As COMConecta.DCOMConecta
    Dim sHora As String
    Dim lsSQL As String
      '  sHora = Mid(psMovNro, 9, 2) & ":" & Mid(psMovNro, 11, 2) & ":" & Mid(psMovNro, 13, 2)
        lsSQL = "Exec stp_ins_ERS034_2019_AfiliacionSegSoat '" & psNumCert & "','" & psPersNombre & "','" & psDoi & "','" & psPersCod & "','" & psUso & "','" & psUsuario & "'," & psPrima & "," & pnMovNro & "," & pnSegEstado & ",'" & psAgeCod & "','" & Format(pdFecSis & " " & sHora, "mm/dd/yyyy hh:mm:ss") & "'"
        dbCmact.Execute (lsSQL)
End Sub
'CTI2 FERIMORO : ERS030-2019 10082020
Public Function insDocVersionadosEmitidosD(ByVal cCtaCod As String, ByVal sDocEmitidos As String, _
                                           ByVal sDocEmitidosHistorial As String, ByVal nFecha As Date)
Dim sSql As String
sSql = "exec stp_ins_ERS039_2019_DocumentosVersionadosEmitidos '" & cCtaCod & "','" & sDocEmitidos & "','" & sDocEmitidosHistorial & "','" & Format(nFecha, "YYYYmmdd") & "'"
dbCmact.Execute sSql

End Function
'CTI2 FERIMORO : ERS030-2019 10082020
Public Function updCalendariointPFMensualD(ByVal sCuenta As String, ByVal nFecha As Date, _
                                           ByVal nIntNuevo As Double)
Dim sSql As String
sSql = "exec updCalendarioIntDPFMensual '" & sCuenta & "'," & nIntNuevo & ",'" & Format(nFecha, "YYYYmmdd") & "'"
dbCmact.Execute sSql

End Function
'CTI2 FERIMORO : ERS030-2019 25082020
Public Function cnsRecuperaFechas(psCuenta As String, ByVal nFecha As Date, ByVal nCtaAbono As Integer) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

sSql = "exec stp_ERS030_2019_ConsFechaCambioTasa '" & Format(nFecha, "YYYYmmdd") & "','" & psCuenta & "'," & nCtaAbono & ""
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsCta.ActiveConnection = Nothing
Set cnsRecuperaFechas = rsCta
Set rsCta = Nothing
End Function

''CTI7 OpeV2********************************************************************
Public Function CapExtornoCargoAho(ByVal pnMovNro As Long, ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String = "", Optional sLpt As String = "") As Double

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
'CTI6 ERS0112020
If nOperacion = gAhoExtCargoServicioRecaudo Then
    sMsgOpe = sGlosa
Else
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
End If
'end CTI6
ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
AgregaMov sMovNro, nOperacion, sGlosa, , gMovFlagDeExtorno
nMovNro = GetnMovNro(sMovNro)
'nMovNro = pnMovNro
AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
UltimaActualizacionCuenta sCuenta, sMovNro

CapExtornoCargoAho = nSaldoCnt

End Function
Public Function RecuperaMovimientoCapataciones(ByVal pnMovNro As Long, Optional ByVal psOpeCod As String = "") As ADODB.Recordset
Dim sSql As String

    'Comentado CTI4 ERS0112020
    'sSql = "Select M.cMovNro, MC.cOpeCod, MC.cCtaCod, MC.nMonto from MovCap MC "
    'sSql = sSql & " Inner Join Mov M ON M.nMovNro = MC.nMovNro "
    'sSql = sSql & " Where MC.nMovNro = " & pnMovNro
    'If Not psOpeCod = "" Then
    '    sSql = sSql & " and Mc.cOpeCod = '" & psOpeCod & "'"
    'End If
    
    sSql = "exec stp_sel_RecuperaMovimientoCapataciones " & pnMovNro 'CTI4 ERS0112020
    
    Set RecuperaMovimientoCapataciones = New ADODB.Recordset
    
    RecuperaMovimientoCapataciones.CursorLocation = adUseClient
    RecuperaMovimientoCapataciones.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    RecuperaMovimientoCapataciones.ActiveConnection = Nothing

End Function
'END CTI4

'PAVITO 20210625
Public Function ActualizaEstadoSeguroSepelio(ByVal psNumCertificado As String, ByVal psFecAfiliacion As String, ByVal psMovNro As String, ByVal pnMovNroReg As Long, ByVal pnEstado As Integer) As String
                                                    
    Dim oDCOMSeg As New COMDCaptaGenerales.DCOMSeguros
    Dim rs As New ADODB.Recordset
    Dim dFecSis As String   'ADDED PAVA20210907
        dFecSis = CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4))  'ADDED PAVA20210907
   
    Set rs = oDCOMSeg.ActualizaEstadoSeguroSepelio(psNumCertificado, dFecSis, pnMovNroReg, pnEstado)
    If Not (rs.BOF And rs.EOF) Then
        ActualizaEstadoSeguroSepelio = rs!cNumCertificado
    End If
    Set oDCOMSeg = Nothing
End Function
'*************************
Public Function getVerificaCC(ByVal sCuenta As String, ByVal psOpeCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim rsCta As ADODB.Recordset
    sSql = " exec STP_SEL_Valida_Excluir_Comision_CC '" & sCuenta & "','" & psOpeCod & "'"
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsCta.ActiveConnection = Nothing
    Set getVerificaCC = rsCta
    Set rsCta = Nothing
End Function
Public Sub setExclusionComisionCC(ByVal sCuenta As String, Optional sCodOpe As String, Optional pnMonto As Double, Optional pnMovNro As String, Optional pnMontoComision As Double, Optional pnTipo As Integer)

    Dim sSql As String
    Dim rs As New ADODB.Recordset
    sSql = " exec STP_INSERT_EXC_COMISION '" & sCuenta & "','" & sCodOpe & "','" & pnMovNro & "'," & pnMonto & "," & pnMontoComision & "," & pnTipo
    Set rs = dbCmact.Execute(sSql)
    Set rs = Nothing
End Sub
    
'CTI7 ERS32-2021***********************************************************************************


Public Function CapAbonoCuentaAho(ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
        Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", Optional sCodIF As String = "", _
        Optional dFechaValor As Date, Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
        Optional sCodCmac As String = "", Optional sNomCmac As String = "", Optional sNomAge As String = "", _
        Optional sLpt As String = "", Optional sPersLavDinero As String = "", Optional bTransaccion As Boolean = False, _
        Optional bImprimeBoleta As Boolean = True, Optional pnMovNroTransfer As Long = -1, _
        Optional pnMonedaTrans As Moneda = gMonedaNacional, Optional ByVal psCodCMAC As String, Optional pCon As ADODB.Connection = Nothing, _
        Optional pbITFAplica As Boolean = True, Optional pnITFValor As Double = 0, Optional pbITFAsumido As Boolean = False, Optional psITFOperacion As String = "", Optional sOPedesc As String = "", Optional ByRef pnExtracto As Long, _
        Optional sBenPersLavDinero As String, Optional ByRef psMensajeValidacion As String = "", Optional ByRef psImpBoleta As String, Optional ByRef psITFImprimeBoleta As String, Optional pbImpTMU As Boolean, _
        Optional psTitPersLavDinero As String = "", Optional psOrdPersLavDinero As String = "", Optional psReaPersLavDinero As String = "", _
        Optional psBenPersLavDinero As String = "", Optional psVisPersLavDinero As String = "", Optional pnMovNro As Long = 0, _
        Optional pnComiDepOtraAge As Variant = 0, Optional pbPagoServicio As Boolean = False, Optional pcPersNomInst As String = "", _
        Optional nComisionPServicio As Double = 0, Optional pcImrimePer As String = "", Optional ByVal pnConceptoSer As String = "", _
        Optional ByVal pnNomCliServ As String = "", Optional ByVal pnMontoSer As String = "", Optional ByVal psNumTarj As String, _
        Optional ByVal pnMovNroRVD As Long = 0, Optional ByVal pnMontoRVD As Long = 0, Optional sPerEcotaxi As Variant = "", Optional pnLogEcotaxi As Integer = 0, _
        Optional ByVal psCtaCodAbono As String = "", Optional ByVal pbDevolRecaudoEcotaxi As Boolean = False, Optional ByVal sDocReferencia As String = "", _
        Optional ByVal psIFTpo As String = "", Optional ByVal psIFCta As String = "", _
        Optional ByVal pnConcepto As Long = 0, Optional ByVal pnMovNroVistoElec As Long = 0) As Double
        '*** RIRO20131102 ERS145 Se agregó parámetro sDocReferencia para los depósitos al club de trabajadores.
        '****RECO20131112 ERS141 Se agrego el parametro psCtaCodAbono con la finalidad de poder registrar la cuenta que realiza el deposito fr Fonfo de garantia ecotaxi
        '****RECO20131112 ERS141 Se agrego el parametro pbDevolRecaudoEcotaxi con la finalidad de poder registrar la devolucion de cuenta recaudo ecotaxi
        'By Capi 12022008 se adiciono Optional sOrdPersLavDinero As String, Optional sReaPersLavDinero As String, Optional VisPersLavDinero As String
        'By 20111213 MADM pnNomCliServ 20110613 MADM para pago de servicios pcPersNomInst - pbPagoServicio - pcImrimePer
        '*** Parametros: pnMovNroRVD, pnMontoRVD agregado por ELRO el 20120717, según OYP-RFC024-2012
        'sPerEcotaxi(1)=psPersCodEcotaxi
        'sPerEcotaxi(2)=psPersNomEcotaxi
        'sPerEcotaxi(3)=psPersDNIEcotaxi
        'FRHU 20141203 ERS048-2014 Se agrego psConcepto y pnMovVistoElec
        'APRI20190109 ERS077-2018 -> Change pnComiDepOtraAge As Double To As Variant
        
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim rsCta As New ADODB.Recordset
    Dim nEstado As CaptacEstado
    Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
    Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
    Dim nSaldoInac As Double
    Dim dUltMov As Date
    Dim bInactiva As Boolean, bTrans As Boolean, bInamovilizada As Boolean
    Dim nNumExtracto As Long, nDiasTranscurridos As Long, nExtracto As Long
    Dim sMsgOpe As String, sCodUser As String, sAgencia As String
    Dim nMovNro As Long
    Dim dFecSis As Date
    Dim sConceptoDesc As String 'FRHU 20141203 ERS048-2014
    Dim oCon As COMDConstSistema.NCOMConstSistema
    Set oCon = New COMDConstSistema.NCOMConstSistema
    Dim oGen As New COMDConstSistema.DCOMGeneral
    Dim oImpre As New COMFunciones.FCOMImpresion
    Dim oConcepto As New COMDConstantes.DCOMConstantes 'FRHU 20141203 ERS048-2014

    Dim lnSaldoDisp As Double
    Dim lnSaldoCont As Double
    Dim bImpreSaldos As Boolean

    Dim lnMontoCV As Currency

    'Para el manejo del Tipo de Cambio
    Dim lnTipCambio As Currency
    Dim lnTipCambioV As Currency
    Dim lnTipCambioC As Currency
    
    Dim lnOpeTarj As Integer 'Indica si la operacion se realizo con tarjeta
    If psNumTarj <> "" Then
        lnOpeTarj = 1
    End If
    
    dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
    sCodUser = Right(sMovNro, 4)
    sAgencia = Mid(sMovNro, 18, 2)

    Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento

    Dim MontoItf As Double
    nMovNro = GetnMovNro(sMovNro)
    
    'APRI20190109 ERS077-2018
    Dim nComiDepOtraAge As Double
    Dim nComiDepxMaxOpe As Double
    Dim nComisionesDeposito As Double
    If IsArray(pnComiDepOtraAge) Then
        nComiDepOtraAge = CDbl(pnComiDepOtraAge(0))
        nComiDepxMaxOpe = CDbl(pnComiDepOtraAge(1))
    Else
        nComiDepOtraAge = 0
        nComiDepxMaxOpe = 0
    End If
    nComisionesDeposito = nComiDepOtraAge + nComiDepxMaxOpe
    'END APRI

     'ande ERS021-2018 20180430 campaña mundial
    Dim cDatosCamp As String
    If Left(psMensajeValidacion, 1) = "#" And Right(psMensajeValidacion, 1) = "&" Then
        cDatosCamp = psMensajeValidacion
        psMensajeValidacion = ""
    End If

    bTrans = False
    'Obtiene los datos para el calculo
    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
    Set clsMant = Nothing
    nEstado = rsCta("nPrdEstado")
    nSaldoDisp = rsCta("nSaldoDisp")
    nSaldoCnt = rsCta("nSaldo")
    dUltMov = rsCta("dUltCierre")
    nTasa = rsCta("nTasaInteres")
    nIntAcum = rsCta("nIntAcum")
    bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
    bInamovilizada = IIf(rsCta("bInmovilizada") = 0, False, True)
    nExtracto = rsCta("nTransacc") + 1
    pnExtracto = nExtracto
    
    rsCta.Close
    Set rsCta = Nothing
    
    'Inicia la transaccion
    Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
    
    On Error GoTo ErrGraba
    
    If Not pCon Is Nothing Then
        clsCap.SetConexion pCon
        clsCap.bTransaccion = True
        bTransaccion = True
    End If
    
    If Not bTransaccion Then clsCap.dbCmact.BeginTrans
    bTrans = True


    If Not ValidaEstadoCuenta(sCuenta, False) Then   'Valida el estado de la cuenta nuevamente
        clsCap.dbCmact.RollbackTrans
        Set clsCap = Nothing
        psMensajeValidacion = "Cuenta NO Tiene un estado valido para la Operacion, consulte con el Asistente de Agencia."
        CapAbonoCuentaAho = 0
        Exit Function
    End If
     'APRI20190109 ERS077-2018 - MEJORA
    If Not ValidaSaldoCuenta_abono(sCuenta, 0, , nComiDepOtraAge, nComiDepxMaxOpe, , , , nMonto) Then 'Valida el saldo de la cuenta nuevamente
        clsCap.dbCmact.RollbackTrans
        Set clsCap = Nothing
        psMensajeValidacion = "Cuenta NO Posee Saldo Suficiente para la operación."
        CapAbonoCuentaAho = 0
        Exit Function
    End If
    'END APRI
    If sGlosa = "" Then sGlosa = "Abono Cuenta = " & sCuenta
    If bActivaCta Then
        If bInactiva Then
            nSaldoInac = nSaldoCnt
        End If
    End If

    'Calcula intereses
    nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
    If nDiasTranscurridos < 0 Then
        nDiasTranscurridos = 0
    End If
    nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
    dUltMov = DateAdd("d", -1, dFecSis)
    nSaldoCnt = nSaldoCnt + nMonto
    If pnMovNro = 0 Then
        AgregaMov sMovNro, nOperacion, sGlosa
    End If
    nMovNro = GetnMovNro(sMovNro)
    
    If sNroDoc = "" Then ' Si la operacion es en efectivo
        ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
        nSaldoDisp = nSaldoDisp + nMonto
        ActualizadUltContactoAho sCuenta, dFecSis
    Else
        If nTipoDoc = TpoDocCheque And nOperacion <> "200252" Then 'Si el abono es con cheque
            'EJVG20140208 *** xq los cheques ya estarán valorizados
            nSaldoDisp = nSaldoDisp + 0
            clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
            nSaldoDisp = nSaldoDisp + nMonto
            'nIntGanado = 0
            'END EJVG *******
            ActualizadUltContactoAho sCuenta, dFecSis
        ElseIf nTipoDoc = TpoDocCheque And nOperacion = "200252" Then
            ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
            nSaldoDisp = nSaldoDisp + nMonto
            ActualizadUltContactoAho sCuenta, dFecSis
        ElseIf nTipoDoc = TpoDocNotaAbono Then 'Si el abono es con nota de abono
            ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
            nSaldoDisp = nSaldoDisp + nMonto
        End If
    End If

    '******** para Boleta ITF
    
     lnSaldoDisp = nSaldoDisp
     lnSaldoCont = nSaldoCnt
    '-----
    AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt, lnOpeTarj, psNumTarj
    
    '***********RECO20131106 ERS-141**************************
    If nOperacion = "200201" And psCtaCodAbono <> "" Then
        RegistraDevFonfoGarantEcotaxi nMovNro, sCuenta, dFecSis, nMonto, nOperacion, psCtaCodAbono, 1, sGlosa
    ElseIf nOperacion = "200203" And psCtaCodAbono <> "" Then
        RegistraDevFonfoGarantEcotaxi nMovNro, sCuenta, dFecSis, nMonto, nOperacion, psCtaCodAbono, 2, sGlosa
    End If
    
    If nOperacion = "200201" And pbDevolRecaudoEcotaxi = True Then
        RegistraDevolucionRecaudoEcotaxi nMovNro, sCuenta, dFecSis, nMonto, nOperacion, sGlosa
    ElseIf nOperacion = "200203" And pbDevolRecaudoEcotaxi = True Then
        RegistraDevolucionRecaudoEcotaxi nMovNro, sCuenta, dFecSis, nMonto, nOperacion, sGlosa
    End If
    '*****************END RECO********************************
    
    '***Modificado por ELRO el 20120828, según OYP-RFC024-2012
    'clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
    'If nOperacion = gAhoDepTransf Then
    'If nOperacion = gAhoDepTransf Or nOperacion = "200245" Then 'EJVG20130923
    If nOperacion = gAhoDepTransf Or nOperacion = "200245" Or nOperacion = gAhoDepActivacionSegTarjeta Then 'FRHU 20140903: gAhoDepActivacionSegTarjeta = 200269
        If pnMovNroTransfer <> -1 Then
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcPenTran, nMonto
        ElseIf pnMovNroTransfer = -1 Then
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
        End If
    Else
        'FRHU 20141203 ERS048-2014
        'clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
        If nOperacion = gCapNotaDeAbono Then
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, CStr(pnConcepto), nMonto
        Else
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
        End If
        'FIN FRHU 20141203
    End If
    '***Fin Modificado por ELRO el 20120828*******************
    If pbITFAplica And pnITFValor > 0 Then
        If pbITFAsumido Then
            AgregaMovCap nMovNro, psITFOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
            AgregaMovCapDet nMovNro, psITFOperacion, sCuenta, gConcITFAsumido, pnITFValor
        Else
            If psITFOperacion = gITFCobroEfectivo Then
                AgregaMovCap nMovNro, psITFOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
                AgregaMovCapDet nMovNro, psITFOperacion, sCuenta, gConcITFCliente, pnITFValor
            Else
                AgregaMovCap nMovNro, psITFOperacion, sCuenta, pnITFValor, nSaldoDisp - pnITFValor, nSaldoCnt - pnITFValor
                AgregaMovCapDet nMovNro, psITFOperacion, sCuenta, gConcITFCliente, pnITFValor
            End If
        End If
    End If
    'FRHU 20141203 ERS048-2014
    sConceptoDesc = oConcepto.GetProductoConcepto(pnConcepto)
    If pnMovNroVistoElec <> 0 Then
        AgregaMovRef nMovNro, pnMovNroVistoElec
    End If
    'FIN FRHU 20141203
     'APRI20190109 ERS077-2018
    If nComisionesDeposito > 0 Then
     ActualizaCargoCaptacion sCuenta, nComisionesDeposito, nComisionesDeposito, 0, dUltMov, sMovNro, bNumExtracto
     AgregaMovCap nMovNro, gAhoDepComision, sCuenta, nComisionesDeposito, nSaldoDisp - pnITFValor - nComisionesDeposito, nSaldoCnt - pnITFValor - nComisionesDeposito
        '*** BRGO 20110127 ***** COBRO DE COMISION X DEPOSITOS OTRA AGENCIA *********
        If nComiDepOtraAge <> 0 Then
             AgregaMovCapDet nMovNro, gAhoDepComision, sCuenta, gConcOpeDepOtraAge, nComiDepOtraAge
        End If
        '*** END BRGO ***************************************************************
        If nComiDepxMaxOpe <> 0 Then
             AgregaMovCapDet nMovNro, gAhoDepComision, sCuenta, gConcOpeRetxMaxOpe, nComiDepxMaxOpe
        End If
    End If
    'END APRI


    If sNroDoc = "" Then ' Si la operacion es en efectivo
        If (pbITFAplica And pnITFValor > 0) And (Not pbITFAsumido) And (Val(psITFOperacion) = gITFCobroCargo Or Val(psITFOperacion) = gITFCobroCMACCargo) Then
            ActualizaAbonoCaptacion sCuenta, nMonto - pnITFValor, nMonto - pnITFValor, nIntGanado, dUltMov, sMovNro, bNumExtracto
        Else
            ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, bNumExtracto
        End If
        
        If nOperacion = gAhoDepEfec Then
            sMsgOpe = "Depósito Efectivo"
        ElseIf nOperacion = gAhoDepTransf Then
            sMsgOpe = "Depósito Transferencia"
        ElseIf nOperacion = gAhoDepDevFondoGar Then
            sMsgOpe = "Depósito Dev Fondo Gar"
        ElseIf nOperacion = gAhoDepPagServEdelnor Then
            sMsgOpe = "Depósito Serv Edelnor"
        ElseIf nOperacion = gAhoDepPagServHidrandina Then
            sMsgOpe = "Depósito Serv Hidrandina"
        ElseIf nOperacion = gAhoDepPagServSedalib Then
            sMsgOpe = "Depósito Serv Sedalib"
        ElseIf nOperacion = gAhoDepSobCaja Then
            sMsgOpe = "Depósito Sobrante"
        ElseIf nOperacion = gAhoDepAboOtrosConceptos Then
            sMsgOpe = "Depósito Otros Conceptos"
        ElseIf nOperacion = gAhoDepDevCredPersonales Then
            sMsgOpe = "Depósito Dev Cred Pers"
        ElseIf nOperacion = "200242" Then
            sMsgOpe = "Abono por otros Servicios"
        ElseIf nOperacion = "200209" Then
            sMsgOpe = "Otros Conceptos RRHH"
        ElseIf nOperacion >= 200247 And nOperacion <= 200251 Then
            sMsgOpe = Trim(sOPedesc)
        'FRHU 20140603 ERS068-2014
        ElseIf nOperacion = 200269 Then
            sMsgOpe = "Abono por Seguro Tarjeta"
        'FRHU 20140603
        'YIHU 20150108 ERS 160
        ElseIf nOperacion = "200245" Then
            sMsgOpe = "Abono Cta. Haberes Transf."
        ElseIf nOperacion = gOpeTransferenciaDeposito Then 'CTI7 OPEv2
            sMsgOpe = "Abono por Conversión de dinero"
        ElseIf nOperacion = gCapNotaDeAbono Then 'FRHU 20141203 ERS048-2014
            sMsgOpe = sConceptoDesc
        End If
        'END YIHU
    Else
        If nTipoDoc = TpoDocCheque Then  'Si el abono es con cheque
        
           If nOperacion <> "200252" Then
              If (pbITFAplica And pnITFValor > 0) And (Not pbITFAsumido) And (Val(psITFOperacion) = gITFCobroCargo Or Val(psITFOperacion) = gITFCobroCMACCargo) Then
                   'clsCap.ActualizaAbonoCaptacion sCuenta, nMonto - pnITFValor, (pnITFValor * -1), nIntGanado, dUltMov, sMovNro, True
                   ActualizaAbonoCaptacion sCuenta, nMonto - pnITFValor, nMonto - pnITFValor, nIntGanado, dUltMov, sMovNro, True 'EJVG20140408
              Else
                   'clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
                   ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True 'EJVG20140408
              End If
           Else
              If (pbITFAplica And pnITFValor > 0) And (Not pbITFAsumido) And (Val(psITFOperacion) = gITFCobroCargo Or Val(psITFOperacion) = gITFCobroCMACCargo) Then
                    ActualizaAbonoCaptacion sCuenta, nMonto - pnITFValor, nMonto - pnITFValor, nIntGanado, dUltMov, sMovNro, bNumExtracto
              Else
                    'clsCap.ActualizaAbonoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
                    ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True 'EJVG20140408
              End If
           End If
             
           If nOperacion <> "200252" Then
                 AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMovNro, nMonto, , psIFCta, psIFTpo
                 AgregaMovDocRec nMovNro, nTipoDoc, sNroDoc, sCodIF, psIFTpo, psIFCta
                 'END EJVG *******
           End If
            
            sMsgOpe = "Dep.Chq"
        End If
    End If
    
    If sMsgOpe = "" Then sMsgOpe = Left(GetCaptaOperacion(Trim(Str(nOperacion))), 28) 'YIHU 20150101 ERS160

    If bInamovilizada Then
        If oCon.LeeConstSistema(gConstSistBitCtrCtasInmovil) = 1 Then
            AgregaMovCap nMovNro, gAhoEstInmovAct, sCuenta, nSaldoInac, nSaldoDisp, nSaldoCnt
            AgregaMovCapDet nMovNro, gAhoEstInmovAct, sCuenta, gConcCapital, nSaldoInac
            ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
            ActualizaInmovilizadaAct sCuenta
            ActualizaInactivaAct sCuenta
        End If
    End If
    
    If bInactiva And Not bInamovilizada Then
        AgregaMovCap nMovNro, gAhoEstInacAct, sCuenta, nSaldoInac, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
        ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
        ActualizaInactivaAct sCuenta
    End If
    'ALPA 20130108*****************************************************
    If pnLogEcotaxi <> 0 Then
        If Trim(sPerEcotaxi(1)) <> "" Then
            Call RegistrarClienteEcotaxi(nMovNro, sPerEcotaxi(1))
        End If
    End If
    '******************************************************************
    If pnMovNroTransfer <> -1 And (nOperacion = gAhoDepTransf Or nOperacion = "200245") Then 'EJVG20130923
    '***Fin Modificado por ELRO el 20120829******************
        AgregaMovRef nMovNro, pnMovNroTransfer
        '***Agregado por ELRO el 20120716, según OYP-RFC024-2012
        ActualizaMovPendientesRend pnMovNroTransfer, pnMontoRVD
        '***Fin Agregado por ELRO el 20120716********************
        If pnMonedaTrans <> Mid(sCuenta, 9, 1) Then
            lnTipCambio = oGen.GetTipCambio(dFecSis, TCFijoMes)
            lnTipCambioC = oGen.GetTipCambio(dFecSis, TCCompra)
            lnTipCambioV = oGen.GetTipCambio(dFecSis, TCVenta)
            
            If Mid(sCuenta, 9, 1) = gMonedaNacional Then
                '***Comentado por ELRO el 201208028, según OYP-RFC024-2012
                'clsCap.ActualizaMovPendientesRend pnMovNroTransfer, Round(nMonto / lnTipCambioC, 2)
                '***Fin Comentado por ELRO el 201208028*******************
                lnMontoCV = Round((nMonto / lnTipCambioC) * lnTipCambio, 2)
                '***Comentado por ELRO el 201208028, según OYP-RFC024-2012
                'clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeCompra, lnMontoCV
                '***Fin Comentado por ELRO el 201208028*******************
                AgregaMovTipoCambio nMovNro, lnTipCambioC
                If lnTipCambioC < lnTipCambio Then
                    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - nMonto)
                Else
                    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - nMonto)
                End If
            Else
                lnMontoCV = Round(nMonto * lnTipCambio, 2)
                AgregaMovTipoCambio nMovNro, lnTipCambioV
                If lnTipCambioV > lnTipCambio Then
                    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - Round(nMonto * lnTipCambioV, 2))
                Else
                    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - Round(nMonto * lnTipCambioV, 2))
                End If
            End If
        End If
    ElseIf pnMovNroTransfer = -1 And (nOperacion = gAhoDepTransf Or nOperacion = "200245") Then 'EJVG20130923
        If pnMonedaTrans <> Mid(sCuenta, 9, 1) Then
            lnTipCambio = oGen.GetTipCambio(dFecSis, TCFijoMes)
            lnTipCambioC = oGen.GetTipCambio(dFecSis, TCCompra)
            lnTipCambioV = oGen.GetTipCambio(dFecSis, TCVenta)
            
            If Mid(sCuenta, 9, 1) = gMonedaNacional Then
                lnMontoCV = Round((nMonto / lnTipCambioC) * lnTipCambio, 2)
                AgregaMovTipoCambio nMovNro, lnTipCambioC
                If lnTipCambioC < lnTipCambio Then
                    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - nMonto)
                Else
                    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - nMonto)
                End If
            Else
                lnMontoCV = Round(nMonto * lnTipCambio, 2)
                AgregaMovTipoCambio nMovNro, lnTipCambioV
                If lnTipCambioV > lnTipCambio Then
                    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - Round(nMonto * lnTipCambioV, 2))
                Else
                    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - Round(nMonto * lnTipCambioV, 2))
                End If
            End If
        End If

    ElseIf pnMovNroTransfer <> -1 And nOperacion = gAhoDepActivacionSegTarjeta Then
        AgregaMovRef nMovNro, pnMovNroTransfer
        ActualizaMovPendientesRend pnMovNroTransfer, pnMontoRVD
    End If
    
    '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
    If (nOperacion = gAhoDepTransf Or nOperacion = "200245") And pnMovNroRVD > 0 Then 'EJVG20130923
        actualizarRelacionVoucherCaptacion nMovNro, pnMovNroRVD
    End If
    '***Fin Agregado por ELRO el 20120717********************

    If sCodCmac <> "" Then
        AgregaMovCMAC nMovNro, sCodCmac, CLng(Mid(sCuenta, 9, 1)), nOperacion, nMonto
    End If
    
    clsCap.UltimaActualizacionCuenta sCuenta, sMovNro

    ' RIRO20131102 ERS145
    If Len(sDocReferencia) > 0 Then
        RegistrarMovDepositoClub nMovNro, nOperacion, sDocReferencia
    End If
    ' FIN RIRO
    
    
    CapAbonoCuentaAho = nSaldoCnt
    
    pnMovNro = nMovNro
    '******************************************************************
    If (Not bImprimeBoleta) Or nOperacion = gAhoDepEntConv Then
            Exit Function
    End If

    '****************************************
    '*************************************************
    If Not bTransaccion Then clsCap.dbCmact.CommitTrans
    bTrans = False
    
    
    Dim sTipDep As String, sCodOpe As String
    Dim sModDep As String, sTipApe As String
    Dim sNomTit As String
    '''sTipDep = IIf(Mid(sCuenta, 9, 1) = Moneda.gMonedaNacional, "SOLES", "DOLARES") 'marg ers044-2016
    sTipDep = IIf(Mid(sCuenta, 9, 1) = Moneda.gMonedaNacional, StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), "DOLARES") 'marg ers044-2016
    sCodOpe = Trim(nOperacion)
    sModDep = sMsgOpe
    sTipApe = "DEPOSITO AHORROS"
    If nOperacion = gCapNotaDeAbono Then sTipApe = "NOTA DE ABONO" 'FRHU 20140630 ERS048-2014
    If nOperacion = gotropeDepUtilidadesTrans Then sTipApe = "PAGO UTILIDADES ABONO EN CTA" 'RIRO 20150601 ERS162-2014
    
    'MADM 20110318
    If pbPagoServicio Then
        Dim pbPagoServiciox As Boolean
        Dim pcPersNomInstx As String
        Dim nComisionPServiciox As Double
        Dim pcImrimePerx As String
        pcImrimePerx = pcImrimePer
        pcPersNomInstx = pcPersNomInst
        pbPagoServiciox = pbPagoServicio
        nComisionPServiciox = nComisionPServicio
        sTipApe = "PAGO SERVICIOS"
    End If
    'END MADM
    
    
    sNomTit = oImpre.ImpreCarEsp(GetNombreTitulares(sCuenta))
    
    If pbITFAplica And pnITFValor > 0 And pbITFAsumido = False Then
        
            If (Val(psITFOperacion) = gITFCobroCargo Or Val(psITFOperacion) = gITFCobroCMACCargo) Then
                nSaldoDisp = nSaldoDisp - pnITFValor
                nSaldoCnt = nSaldoCnt - pnITFValor
                
            End If
    End If
    
    '**** BRGO 20110127
    If nComiDepOtraAge <> 0 Then 'APRI20190109 ERS077-2018 Change pnComiDepOtraAge To nComiDepOtraAge
        nSaldoDisp = nSaldoDisp - nComiDepOtraAge
        nSaldoCnt = nSaldoCnt - nComiDepOtraAge
    End If
    '******************
    If nComiDepxMaxOpe <> 0 Then
        nSaldoDisp = nSaldoDisp - nComiDepxMaxOpe
        nSaldoCnt = nSaldoCnt - nComiDepxMaxOpe
    End If
    'END APRI
    
    Dim nTipoPag As Integer


    If pnITFValor > 0 Then
        If psITFOperacion = gITFCobroEfectivo Then
            nTipoPag = 1
        Else
            nTipoPag = 2
        End If
    End If
    'Do
    If sNroDoc <> "" Then
        Select Case nTipoDoc
            Case TpoDocCheque
                If sCodCmac <> "" Then
                    psImpBoleta = psImpBoleta & ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntGanado, "Fecha Valor", nExtracto, nSaldoCnt, True, , , , , , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, sAgencia, , bImpreSaldos, CStr(nOperacion), , , , , , , , False, nTipoPag, CDbl(pnITFValor), True, , , pbImpTMU, , , , , , , , , , , , , sPerEcotaxi, pnLogEcotaxi)
                Else
                    psImpBoleta = psImpBoleta & ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntGanado, "Fecha Valor", nExtracto, nSaldoCnt, True, , , , , , , , cDatosCamp, dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, sAgencia, , bImpreSaldos, CStr(nOperacion), , , , , , , , False, nTipoPag, CDbl(pnITFValor), True, , , pbImpTMU, , , , , , , , , , , , , sPerEcotaxi, pnLogEcotaxi)
                End If
            Case TpoDocNotaAbono
                psImpBoleta = psImpBoleta & ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, sAgencia, , bImpreSaldos, CStr(nOperacion), , , , , , , , False, nTipoPag, CDbl(pnITFValor), True, , , pbImpTMU, , , , , , , , , , , , , sPerEcotaxi, pnLogEcotaxi)
        End Select
    Else
        If sCodCmac <> "" Then
            sTipApe = "DEPOSITO CMAC AHORROS"
            psImpBoleta = psImpBoleta & ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, sAgencia, , bImpreSaldos, CStr(nOperacion), , , , , , , , False, nTipoPag, CDbl(pnITFValor), True, nComiDepOtraAge, nComiDepxMaxOpe, pbImpTMU, , , , , , , , , , , , , sPerEcotaxi, pnLogEcotaxi)  'APRI20190109 ERS077-2018
        Else
            psImpBoleta = psImpBoleta & ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , cDatosCamp, dFecSis, sNomAge, sCodUser, sLpt, , psCodCMAC, sAgencia, , bImpreSaldos, CStr(nOperacion), , , , , , , , False, nTipoPag, CDbl(pnITFValor), True, nComiDepOtraAge, nComiDepxMaxOpe, pbImpTMU, psNumTarj, pbPagoServiciox, pcPersNomInstx, nComisionPServiciox, pcImrimePerx, , , pnConceptoSer, pnNomCliServ, pnMontoSer, , , sPerEcotaxi, pnLogEcotaxi) 'APRI20190109 ERS077-2018 JIPR20191018
        End If
    End If
    

Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    'MsgBox "OCURRIO UN ERROR, VERIFIQUE ESTA OPERACION ERROR:" & Err.Description, vbOKOnly + vbInformation, "AVISO"
    Err.Raise Err.Number, "OCURRIO UN ERROR, VERIFIQUE ESTA OPERACION ", Err.Description
    Set clsCap = Nothing
    'Err.Raise Err.Number, "", Err.Description
    CapAbonoCuentaAho = 0
End Function

Public Function ValidaEstadoCuenta(ByVal sCuenta As String, Optional pbParaRetiro As Boolean = True) As Boolean
Dim rsCta As ADODB.Recordset
Set rsCta = New ADODB.Recordset
Dim nEstado As CaptacEstado

Set rsCta = GetDatosCuenta(sCuenta)

nEstado = rsCta("nPrdEstado")

rsCta.Close
Set rsCta = Nothing


If pbParaRetiro Then
      If nEstado = gCapEstActiva Or nEstado = gCapEstBloqRetiro Then 'GIPO 04-03-2017
        ValidaEstadoCuenta = True
    Else
        ValidaEstadoCuenta = False
    End If
Else
    If nEstado = gCapEstActiva Or nEstado = gCapEstBloqRetiro Or nEstado = gCapEstBloqTotal Then
        ValidaEstadoCuenta = True
    Else
        ValidaEstadoCuenta = False
    End If
End If
End Function

Public Sub RegistrarMovDepositoClub(ByVal nMovNro As Long, ByVal sOpeCod As CaptacOperacion, ByVal sDocReferencia As String)

    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo error
    
        sSql = "exec stp_ins_MovDepositoClub " & _
                nMovNro & ", '" & _
                sOpeCod & "', '" & _
                sDocReferencia & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.CargaRecordSet (sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Sub
error:
    
    Set oConecta = Nothing
    Err.Raise Err.Number, "Buscar Usuario ", Err.Description

End Sub

Public Function ImprimeBoleta(ByVal psTit As String, ByVal pstext As String, ByVal psCodOpe As String, ByVal pnMonto As String, _
        ByVal psCliente As String, ByVal psCodCta As String, ByVal psNumDoc As String, ByVal pnSaldo As Double, _
        ByVal pnInteresA As String, NomDoc As String, ByVal pnNumExt As Long, ByVal pnSaldoC As Double, _
        Optional pbOpSaldoC As Boolean = True, Optional pbSaldoInt As Boolean = True, Optional psNumDias As String = "----", _
        Optional psNomAgeRem As String = "", Optional psCodUsuRem As String = "", Optional pbCuenta As Boolean = False, _
        Optional psComCmac As String = "XXX", Optional psLin3 As String = "XXX", Optional psTexto As String = "XXX", _
        Optional pdFecSis As Date, Optional sNomAge As String = "", Optional sCodUser As String = "", _
        Optional sLpt As String = "", Optional bPregImpSaldos As Boolean = True, Optional ByVal psCodCMAC As String, _
        Optional sCodage As String = "", Optional pnMonedaCMAC As Integer = 1, Optional ByRef bImpreSaldos As Boolean = False, Optional psOpeCod As String = "", _
        Optional sUserBus As String = "", Optional bExtorno As Boolean = False, Optional ByVal sProxReno As String = "", Optional psDisponible As String = "", Optional psIntangible As String = "", _
        Optional psDispCta As String = "", Optional psIntaCta As String = "", Optional ByVal pbImprimeSaldos As Boolean = True, Optional pnTipoPago As Integer, Optional pnMontoIFT As Double, _
        Optional pbImpItf As Boolean = False, Optional pnComixRetOtraAge As Double = 0, Optional pnComixMaxRet As Double = 0, Optional pbImpTMU As Boolean = False, Optional psNumTarj As String = "", _
        Optional pbPagoServicio As Boolean = False, Optional pcPersNomInst As String = "", Optional pnComisionPServiciox As Double = 0, Optional pcImrimePerx As String = "", _
        Optional ByVal pnComiVB As Double = 0, Optional ByVal psCodOpeVB As String = "", Optional ByVal pnConceptoSer As String = "", Optional ByVal psNomCliServ As String = "", Optional ByVal pnMontoSer As String = "0", _
        Optional ByVal psProxReno2 As String = "", Optional ByVal psTipoCambio As String = "XXX", Optional sPerEcotaxi As Variant = "", Optional pnLogEcotaxi As Integer = 0) As String

        
    Dim sFecha As String
    Dim sHora As String
    Dim sSep As Integer
    Dim sIni As Integer
    Dim sMonto As String
    Dim sSDisp As String
    Dim sIntAcum As String
    Dim sMax As Integer
    Dim saux As Integer
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

    Dim oVarImpre As New COMFunciones.FCOMVarImpresion

    '** Para la Impresion del Codigo Antiguo en las Boletas
    Dim lsCodAntiguo As String
    Dim lscad As String

    Dim nComiTotal As Double
    Dim nComiTotalServicio As Double
    Dim nTotalServicio As Double '20110701
    Dim pcImrimePer As String

    Dim loConstSistema As COMDConstSistema.NCOMConstSistema
    Set loConstSistema = New COMDConstSistema.NCOMConstSistema
    If loConstSistema.LeeConstSistema(122) = "1" Then
        Dim loGeneral As COMDConstSistema.DCOMGeneral
        Set loGeneral = New COMDConstSistema.DCOMGeneral
        lsCodAntiguo = loGeneral.GetCuentaAntigua(psCodCta)
    Else
        lsCodAntiguo = ""
    End If
    Set loConstSistema = Nothing
    Dim nComisionTrans As String 'RIRO20131212 ERS137
    
    'ande ERS021-2018
    Dim bPartipaCamp As Boolean, nPunto As Integer, nCondicion As Integer, nPTotalAcumulado As Integer, cNombreCliente As String, cTexto As String, cadenita As String
    If Left(psTexto, 1) = "#" And Right(psTexto, 1) = "&" Then
        cTexto = Trim(Replace(Replace(psTexto, "#", ""), "&", ""))
        bPartipaCamp = IIf(Mid(cTexto, 1, InStr(cTexto, ".") - 1) = "1", True, False)
        nPunto = CInt(Mid(cTexto, InStr(cTexto, ".") + 1, InStr(cTexto, "$") - InStr(cTexto, ".") - 1))
        nCondicion = CInt(Mid(cTexto, InStr(cTexto, "$") + 1, InStr(cTexto, "_") - InStr(cTexto, "$") - 1))
        nPTotalAcumulado = CInt(Right(cTexto, Len(cTexto) - InStr(cTexto, "_")))
        psTexto = "XXX"
    ElseIf psTexto = "" Then
        psTexto = "XXX"
    End If
    
ETIQ:

    On Error GoTo error

    lnTope = 0 '6 'Tope de lineas en Boleta

    lsNroExt = Str(pnNumExt)
    

    lscad = lscad & oImp.gPrnInicializa
    lscad = lscad & oImp.gPrnEspaLineaN   'espaciamiento lineas 1/6 pulg.
    lscad = lscad & oImp.gPrnTamPagina22  'Longitud de página a 22 líneas'
    lscad = lscad & oImp.gPrnTamLetra10CPI   'Tamaño 10 cpi
    lscad = lscad & oImp.gPrnTpoLetraRoman     'Tipo de Letra Sans Serif
    lscad = lscad & oImp.gPrnCondensadaOFF ' cancela condensada
    lscad = lscad & oImp.gPrnNegritaOFF ' desactiva negrita

    sSep = 15
    sIni = 0
    sMax = 33
    saux = 5

    sFecha = Format$(pdFecSis, "dd/mm/yyyy")
    sHora = Format$(Time, "hh:mm:ss")
    sMonto = Format$(pnMonto, "#,##0.00")
    sSDisp = Format$(pnSaldo, "#,##0.00")
    lsSaldoC = Format$(pnSaldoC, "#,##0.00")

    lsNomAge = sNomAge

    '' Declaracion para los mensajes
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Dim oImpre As New COMFunciones.FCOMImpresion
    
       'CTI3 04122018 ERS082-2018 ********************************
        Dim TpoCta As String
        'TpoCta = Space(4)
     '       If psCodOpe = gAhoApeEfec Or psCodOpe = gAhoApeLoteEfec Or psCodOpe = gAhoApeTransf Or psCodOpe = gAhoApeCargoCta Or psCodOpe = gAhoApeLoteTransfBanco _
            Or psCodOpe = gAhoRetEfec Or psCodOpe = gAhoRetOP Or psCodOpe = gAhoRetTransf Or psCodOpe = gAhoRetOPCanje Or psCodOpe = gPFApeEfec Then

 
                
                  TpoCta = "(" & GetTpoCta(psCodCta) & ")"
                
               
     '       End If
        '*********************************************************
    
    
    'JUEZ 20131021 *******************************************************************************
    Dim oGen As New COMDConstSistema.DCOMGeneral, nEstadoCampDatos As Integer
    nEstadoCampDatos = oGen.LeeConstSistema(gConstSistEstadoCampActualizaDatos)
    If nEstadoCampDatos = 1 Then
        If pbImpTMU = False Then
            lscad = lscad & oImp.gPrnSaltoLinea & oImp.gPrnNegritaON
            lscad = lscad & "                            * *         " & Space(sSep - 2) & "                            * *         " & oImp.gPrnSaltoLinea
            lscad = lscad & "                         *      *       " & Space(sSep - 2) & "                         *      *       " & oImp.gPrnSaltoLinea
            lscad = lscad & "                        *               " & Space(sSep - 2) & "                        *               " & oImp.gPrnSaltoLinea
            lscad = lscad & "      * *              *          *     " & Space(sSep - 2) & "      * *              *          *     " & oImp.gPrnSaltoLinea
            lscad = lscad & "    *     *      *   *             *    " & Space(sSep - 2) & "    *     *      *   *             *    " & oImp.gPrnSaltoLinea
            lscad = lscad & "   *        *  *   *                    " & Space(sSep - 2) & "   *        *  *   *                    " & oImp.gPrnSaltoLinea
            lscad = lscad & "  *           *                     *   " & Space(sSep - 2) & "  *           *                     *   " & oImp.gPrnSaltoLinea
            lscad = lscad & " *                                   *  " & Space(sSep - 2) & " *                                   *  " & oImp.gPrnSaltoLinea
            lscad = lscad & "*                                     * " & Space(sSep - 2) & "*                                     * " & oImp.gPrnSaltoLinea
            lscad = lscad & "ACTUALIZA Y GANA UN PASAJE DOBLE A CUSCO" & Space(sSep - 2) & "ACTUALIZA Y GANA UN PASAJE DOBLE A CUSCO" & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
            lscad = lscad & oImp.gPrnSaltoLinea
        Else
            lscad = lscad & oImp.gPrnSaltoLinea & oImp.gPrnNegritaON & "C"
            lscad = lscad & "                            * *" & oImp.gPrnSaltoLinea
            lscad = lscad & "                         *      *" & oImp.gPrnSaltoLinea
            lscad = lscad & "                        *         " & oImp.gPrnSaltoLinea
            lscad = lscad & "      * *              *          *" & oImp.gPrnSaltoLinea
            lscad = lscad & "    *     *      *   *             *" & oImp.gPrnSaltoLinea
            lscad = lscad & "   *        *  *   *                " & oImp.gPrnSaltoLinea
            lscad = lscad & "  *           *                     *" & oImp.gPrnSaltoLinea
            lscad = lscad & " *                                   *" & oImp.gPrnSaltoLinea
            lscad = lscad & "*                                     *" & oImp.gPrnSaltoLinea
            lscad = lscad & "ACTUALIZA Y GANA UN PASAJE DOBLE A CUSCO" & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
            lscad = lscad & oImp.gPrnSaltoLinea
        End If
    End If
    'END JUEZ ************************************************************************************
    
    'RIRO20131212 ERS137 *******
    If InStr(1, "200303,200411,230303,230411,210209,210302,240204,240304,220302,220403,250302,250403", psCodOpe) > 0 And Not pbPagoServicio Then
        nComisionTrans = Trim(pnMontoSer)
    End If
    'END RIRO            *******
        
    '' liquidacion Trujillo
    If pbImpTMU = False Then
        lscad = lscad & oImp.gPrnNegritaON  'Activa Negrita
        lscad = lscad & oImp.gPrnSaltoLinea
        nComiTotalServicio = 0
        nTotalServicio = 0
        'CTI5 ERS0112020
        If psCodOpe = "" Then
            psCodOpe = 0
        End If
        'CTI5 ERS0112020 FIN
        'RIRO20160111 ADD gOtrOpePagoRecaudoCajeroCorresponsal
        If pbPagoServicio Or psCodOpe = gOtrOpePagoRecaudoCajeroCorresponsal Then
            lscad = lscad & Space(sIni) & "CMAC-MAYNAS - SERVICIOS" & Space(16 + sSep + saux) & "CMAC-MAYNAS - SERVICIOS" & oImp.gPrnSaltoLinea
            psCliente = pcPersNomInst
            nComiTotalServicio = pnComisionPServiciox
            pcImrimePer = Trim(pcImrimePerx)
        Else
            lscad = lscad & Space(sIni) & "CMAC-MAYNAS - AHORRO" & Space(16 + sSep + saux) & "CMAC-MAYNAS - AHORRO" & oImp.gPrnSaltoLinea
        End If

        If Len(Trim(psCodCta)) < 18 Then
            If pnMonedaCMAC = 1 Then
                '''lscad = lscad & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & oImp.gPrnSaltoLinea'marg ers044-2016
                lscad = lscad & Space(sIni) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
            Else
                lscad = lscad & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & oImp.gPrnSaltoLinea
            End If
        Else
            If Mid(psCodCta, 9, 1) = 1 Then
                '''lscad = lscad & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
                lscad = lscad & Space(sIni) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
            Else
                lscad = lscad & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & oImp.gPrnSaltoLinea
            End If
        End If

        If psNomAgeRem = "" Then
            lscad = lscad & ""
        Else
            lscad = lscad & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(saux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & "Ag.Rem: " & Trim(psNomAgeRem) & oImp.gPrnSaltoLinea
        End If
        
        If psComCmac = "XXX" Then
            If psLin3 = "XXX" Then
                lscad = lscad & oImp.gPrnNegritaOFF 'Desactiva Negrita
            Else
                lscad = lscad & Space(sIni) & psLin3 & Space(saux + sSep + sMax - Len(psLin3)) & psLin3 & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea  'Desactiva Negrita
                lnNumLinCmac = 1
            End If
            lnNumLinCmac = 0
        Else
            lscad = lscad & Space(sIni) & "NroDocCmac:" & psComCmac & Space(saux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & "NroDocCmac:" & psComCmac & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea 'Desactiva Negrita
            lnNumLinCmac = 1
        End If
        
        lscad = lscad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(saux + sSep - 6) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & oImp.gPrnSaltoLinea
        
        lnCliAux = InStr(1, pcPersNomInst, "*", vbTextCompare)
        
        If lnCliAux = 0 Then
            If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(psCliente) & Space(saux + sMax + sSep - Len(psCliente)) & oImpre.ImpreCarEsp(psCliente) & oImp.gPrnSaltoLinea
        Else
            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
    
            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
            
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux1) & Space(saux + sMax + sSep - Len(lsCliAux1)) & oImpre.ImpreCarEsp(lsCliAux1) & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux2) & Space(saux + sMax + sSep - Len(lsCliAux2)) & oImpre.ImpreCarEsp(lsCliAux2) & oImp.gPrnSaltoLinea
            
            lnCliAux = 1
        End If
        
        If pbSaldoInt Or pbCuenta Then
            If lsCodAntiguo <> "" Then  ' Para Codigo Antiguo - LAYG
              'CTI3 - FERIMORO : ERS082-2018
                lscad = lscad & Space(sIni) & "Cuenta:" & lsCodAntiguo & Space(2) & TpoCta & Space(2 + sSep + saux) & "Cuenta:" & lsCodAntiguo & Space(2) & TpoCta & oImp.gPrnSaltoLinea
                'lscad = lscad & Space(sIni) & "Cuenta:" & lsCodAntiguo & Space(8 + sSep + saux) & "Cuenta:" & lsCodAntiguo & oImp.gPrnSaltoLinea
                'lsCadImp = lsCadImp & " Contrato  : " & lsCodAntiguo & Space(10) & Space(1 + vEspacio) _
                                            & " Contrato  : " & lsCodAntiguo & Space(10) & Chr(10)
            Else
               'CTI3 FERIMORO : ERS082-2018
                If Left(psCodOpe, 4) = "2001" Or Left(psCodOpe, 4) = "2003" Or Left(psCodOpe, 4) = "2004" _
                Or Left(psCodOpe, 4) = "2101" Or Left(psCodOpe, 4) = "2102" Or Left(psCodOpe, 4) = "2103" Then
                  lscad = lscad & Space(sIni) & "Cuenta:" & psCodCta & Space(2) & TpoCta & Space(2 + sSep + saux) & "Cuenta:" & psCodCta & Space(2) & TpoCta & oImp.gPrnSaltoLinea
                Else ' EN CASO SEAN EXTORNOS
                  lscad = lscad & Space(sIni) & "Cuenta:" & psCodCta & Space(8 + sSep + saux) & "Cuenta:" & psCodCta & oImp.gPrnSaltoLinea
                End If
            End If
        Else
            lscad = lscad & ""
        End If
        
        psTit = Trim(psTit)
        psTit = oImpre.CentrarCadena(psTit, 28)
        lscad = lscad & oImp.gPrnNegritaON 'Activa Negrita
        lscad = lscad & Space(sIni + 1) & "-----" & psTit & "-----" & Space(-1 + sSep) & "-----" & psTit & "-----" & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnNegritaOFF
        lnChq = InStr(1, pstext, "*", vbTextCompare)
        
        If psTexto = "XXX" Then
            If lnChq = 0 Then
                lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(pstext, 1, 28))) - Len(sMonto)) & sMonto & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
                'MADM 20111209
                If pbPagoServicio Then
                     Dim xx1 As Integer
                     Dim psConceptoSerTemp As String
                     Dim psMontoSerTemp As String
                     psConceptoSerTemp = pnConceptoSer
                     psMontoSerTemp = pnMontoSer
                     pnConceptoSer = Trim(pnConceptoSer)
                     pnMontoSer = Trim(pnMontoSer)
                     pnConceptoSer = Mid(pnConceptoSer, 1, Len(pnConceptoSer) - 5)
                     If Int(Right(psConceptoSerTemp, 5)) = 1 Then
                            pnConceptoSer = pnConceptoSer & " XXX"
                            pnMontoSer = pnMontoSer & " XXX"
                     End If
                     lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(Trim(Mid(psNomCliServ, 1, 30))) & Space(sMax + 6 - Len(Trim(Mid(psNomCliServ, 1, 30)))) & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(Mid(psNomCliServ, 1, 30))) & Space(sMax + 6 - Len(Trim(Mid(psNomCliServ, 1, 30)))) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
                     For xx1 = 1 To Int(Right(psConceptoSerTemp, 5))
                        lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(Trim(Mid(Mid(pnConceptoSer, 4, InStr(4, pnConceptoSer, "XXX", vbTextCompare) - 4), 1, 31))) & Space(sMax + 6 - (Len(Trim(Mid(Mid(pnConceptoSer, 4, InStr(4, pnConceptoSer, "XXX", vbTextCompare) - 4), 1, 31))) + Len(Trim(Mid(Mid(pnMontoSer, 4, InStr(4, pnMontoSer, "XXX", vbTextCompare) - 4), 1, 8))))) & oImpre.ImpreCarEsp(Trim(Mid(Mid(pnMontoSer, 4, InStr(4, pnMontoSer, "XXX", vbTextCompare) - 4), 1, 8))) & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(Mid(Mid(pnConceptoSer, 4, InStr(4, pnConceptoSer, "XXX", vbTextCompare) - 4), 1, 31))) & Space(sMax + 6 - (Len(Trim(Mid(Mid(pnConceptoSer, 4, InStr(4, pnConceptoSer, "XXX", vbTextCompare) - 4), 1, 31))) + Len(Trim(Mid(Mid(pnMontoSer, 4, InStr(4, pnMontoSer, "XXX", vbTextCompare) - 4), 1, 8))))) & oImpre.ImpreCarEsp(Trim(Mid(Mid(pnMontoSer, 4, InStr(4, pnMontoSer, "XXX", vbTextCompare) - 4), 1, 8))) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
                        pnConceptoSer = Trim(Mid(pnConceptoSer, InStr(5, pnConceptoSer, "XXX", vbTextCompare), Len(pnConceptoSer) - 6))
                        If Right(Trim(pnConceptoSer), 3) <> "XXX" Then
                            pnConceptoSer = pnConceptoSer & " XXX"
                        End If
                        If Right(Trim(pnMontoSer), 3) <> "XXX" Then
                            pnMontoSer = pnMontoSer & " XXX"
                        End If
                     Next xx1
                End If
                lscad = lscad & ""
            Else
                lsChqAux1 = (Mid(pstext, 1, lnChq - 1))
                lsChqAux2 = (Mid(pstext, lnChq + 1))
                lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Space(-1 + sSep) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
            End If
        Else
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & Space(saux + sSep + sMax - Len(Trim(psTexto))) & oImpre.ImpreCarEsp(Trim(psTexto)) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
            If psDisponible = "" And psIntangible = "" Then
                lscad = lscad & ""
            End If
        End If
        
        If psDisponible <> "" And psIntangible <> "" Then
           'psDispCta psIntaCta
           'gPrnCondensadaOFF
           'gPrnCondensadaON
           lscad = lscad & oVarImpre.gPrnCondensadaON & Space(sIni) & oImpre.ImpreCarEsp(Trim(psDispCta)) & Space(saux + sSep + sMax - Len(Trim(psDispCta))) & oImpre.ImpreCarEsp(Trim(psDispCta)) & oImp.gPrnSaltoLinea
           lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psIntaCta)) & Space(saux + sSep + sMax - Len(Trim(psIntaCta))) & oImpre.ImpreCarEsp(Trim(psIntaCta)) & oImp.gPrnSaltoLinea
           lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psDisponible)) & Space(saux + sSep + sMax - Len(Trim(psDisponible))) & oImpre.ImpreCarEsp(Trim(psDisponible)) & oImp.gPrnSaltoLinea
           lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psIntangible)) & Space(saux + sSep + sMax - Len(Trim(psIntangible))) & oImpre.ImpreCarEsp(Trim(psIntangible)) & oVarImpre.gPrnCondensadaOFF & oImp.gPrnSaltoLinea

        End If
        'Insertar ITF
        If pbImpItf = True Then
            If pnMontoIFT <> 0 Then
                If pnTipoPago = 1 Then
                    psTexto = "ITF Efectivo"
                 
                ElseIf pnTipoPago = 2 Then
                    psTexto = "ITF Cargo a Cuenta"
                'ALPA 20091120************************************************
                ElseIf pnTipoPago = 3 Then
                    psTexto = "ITF Abono a Cuenta"
                End If
                '*************************************************************
                lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(pnMontoIFT)) & Format(pnMontoIFT, "0.00") & Space(-1 + 15) & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(pnMontoIFT)) & Format(pnMontoIFT, "0.00") & oImp.gPrnSaltoLinea
            End If
        End If
        '***    IMPRIMIR COMISON DE RETIRO -- AVMM 10-08-2006  ***
        'nComiTotal = pnComixRetOtraAge + pnComixMaxRet'Comentado by JACA 20111025
        'If nComiTotal <> 0 Then'Comentado by JACA 20111025
            If pnComixRetOtraAge <> 0 Then
               'psTexto = "Comisión Operación Otra Age."
               psTexto = "Comisión " & IIf(Left(psCodOpe, 4) = "2002", "Dep.", IIf(Left(psCodOpe, 4) = "2003", "Ret.", "Ope.")) & " Otra Loc." 'JUEZ 20150930
              'lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(pnComixRetOtraAge)) & Format(pnComixRetOtraAge, "0.00") & Space(-1 + 15) & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(pnComixRetOtraAge)) & Format(pnComixRetOtraAge, "0.00") & oImp.gPrnSaltoLinea
               lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 28))) & Space(33 + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(Format(pnComixRetOtraAge, "#,##0.00"))) & Format(pnComixRetOtraAge, "#,##0.00") & Space(-1 + 15) & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 28))) & Space(33 + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(Format(pnComixRetOtraAge, "#,##0.00"))) & Format(pnComixRetOtraAge, "#,##0.00") & oImp.gPrnSaltoLinea 'APRI20190109 MEJORA SUGERENCIA CALIDAD
            End If
            If pnComixMaxRet <> 0 Then
               psTexto = "Comisión x Max Nro Ope."
               lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(pnComixMaxRet)) & Format(pnComixMaxRet, "0.00") & Space(-1 + 15) & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(pnComixMaxRet)) & Format(pnComixMaxRet, "0.00") & oImp.gPrnSaltoLinea
            End If
        'End If
        '*********************************************************
        'MADM 20110815 - 20110318
        If nComiTotalServicio <> 0 Then
           nTotalServicio = nComiTotalServicio + pnMontoIFT + sMonto 'Format(nComiTotalServicio + pnMontoIFT + sMonto, "#,###,##0.00")
           psTexto = "Comision Pago de Servicios:"
           lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 27))) & Space(39 - Len(psTexto) - Len(Format(nComiTotalServicio, "0.00"))) & Format(nComiTotalServicio, "0.00") & Space(-1 + 15) & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 27))) & Space(39 - Len(psTexto) - Len(Format(nComiTotalServicio, "0.00"))) & Format(nComiTotalServicio, "0.00") & oImp.gPrnSaltoLinea
           psTexto = "Total Pagado              :"
           lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 27))) & Space(39 - Len(psTexto) - Len(Format(nTotalServicio, "#,###,##0.00"))) & Format(nTotalServicio, "#,###,##0.00") & Space(-1 + 15) & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 27))) & Space(39 - Len(psTexto) - Len(Format(nTotalServicio, "#,###,##0.00"))) & Format(nTotalServicio, "#,###,##0.00") & oImp.gPrnSaltoLinea
        End If
       '*********************************************************
       
       'Add by GITU 30-09-2011
       If pnComiVB <> 0 Then
           'psTexto = "Comision Ope Sin Tarjeta" 'Comentado CTI4 ERS0112020
           'CTI4 ERS0112020
            Select Case psCodOpeVB
                Case "300627"
                    psTexto = "Comisión Ope Sin Tarjeta"
                Case gComisionEmisionChequeCargoCta
                    psTexto = "Comisión Cargo a Cuenta"
                Case gComisionEmisionChequeEfectivo
                    psTexto = "Comisión Efectivo"
            End Select
           'CTI4 End
           lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 26))) & Space(39 - Len(psTexto) - Len(Format(pnComiVB, "0.00"))) & Format(pnComiVB, "0.00") & Space(-1 + 15) & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 26))) & Space(39 - Len(psTexto) - Len(Format(pnComiVB, "0.00"))) & Format(pnComiVB, "0.00") & oImp.gPrnSaltoLinea
       End If
       'END GITU
       
       'RIRO20131212 ERS137
       If Len(Trim(nComisionTrans)) <> 0 Then
           psTexto = "Comision Transf:"
           lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 26))) & Space(39 - Len(psTexto) - Len(Format(nComisionTrans, "0.00"))) & Format(nComisionTrans, "0.00") & Space(-1 + 15) & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 26))) & Space(39 - Len(psTexto) - Len(Format(nComisionTrans, "0.00"))) & Format(nComisionTrans, "0.00") & oImp.gPrnSaltoLinea
       End If
       'END RIRO
        
        'If pbSaldoInt Then
        If bPregImpSaldos Then
            'If MsgBox("Desea Imprimir los Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            If pbImprimeSaldos Then
                If psCodOpe <> "200399" Then 'MADM 20101115
                    bImpreSaldos = True
                    lscad = lscad & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Space(-1 + sSep) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & oImp.gPrnSaltoLinea
                    If pbOpSaldoC Then
                        lscad = lscad & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Space(-1 + sSep) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & oImp.gPrnSaltoLinea
                    Else
                        lscad = lscad & ""
                    End If
                End If
            Else
                bImpreSaldos = False
                lscad = lscad & ""
                lscad = lscad & ""
                pbSaldoInt = False
            End If
        Else
            lscad = lscad & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Space(-1 + sSep) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & oImp.gPrnSaltoLinea
            If pbOpSaldoC Then
               lscad = lscad & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Space(-1 + sSep) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & oImp.gPrnSaltoLinea
            Else
                lscad = lscad & ""
            End If
        End If
        'Else
        '   lscad = lscad & ""
        'End If
 
        lsInteres = pnInteresA
        
        If pbSaldoInt Then
            If lsInteres <> "No Valido" Then
                lsInteres = Format(lsInteres, "#,##0.00")
                lscad = lscad & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Space(-1 + sSep) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & oImp.gPrnSaltoLinea
            End If
        Else
            lscad = lscad & ""
        End If
        
        '***Modificado por ELRO el 20121109, según OYP-RFC098-2012
        If Trim(psProxReno2) <> "" Then
            lscad = lscad & Space(sIni) & psProxReno2 & Space(sMax + 20 - Len(psProxReno2)) & psProxReno2 & oImp.gPrnSaltoLinea
        End If
        '***Fin Modificado por ELRO el 20121109*******************

        If Trim(sProxReno) <> "" Then
              lscad = lscad & Space(sIni) & sProxReno & Space(sMax + 20 - Len(sProxReno)) & sProxReno & oImp.gPrnSaltoLinea
        End If
        
        
        If Not psNumDoc = "" Then
            lscad = lscad & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Space(-1 + sSep) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & oImp.gPrnSaltoLinea
        Else
            If Trim(psDisponible) = "" And Trim(psIntangible) = "" Then
                lscad = lscad & ""
            End If
        End If
        
        If Not psNumDias = "----" Then
            lscad = lscad & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & Space(-1 + sSep) & "Nro Dias Interes" & Space(sMax + 6 - Len(psNumDias) - Len("Nro Dias Interes")) & psNumDias & oImp.gPrnSaltoLinea
            lnTope = 4 - lnCliAux
        Else
            lnTope = 3 - lnCliAux
        End If
        
        lscad = lscad & oImp.gPrnNegritaON & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & "---------------------------------------" & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        If psCodUsuRem = "" Then
            If bExtorno Then
                lscad = lscad & Space(sIni) & "USUARIO EXTORNA : " & oImpre.ImpreCarEsp(sCodUser) & Space(11 + sSep + saux) & "USUARIO EXTORNA : " & oImpre.ImpreCarEsp(sCodUser) & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & "USUARIO OPERADOR: " & oImpre.ImpreCarEsp(sUserBus) & Space(11 + sSep + saux) & "USUARIO OPERADOR: " & oImpre.ImpreCarEsp(sUserBus) & oImp.gPrnSaltoLinea
            Else
                'ALPA 20130110****************************************************************************************************************************************
                If pnLogEcotaxi <> 0 Then
                    If Len(Trim(sPerEcotaxi(1))) <> 0 Then
                        lscad = lscad & Space(sIni) & Left(oImpre.ImpreCarEsp("ABONADO POR: " & sPerEcotaxi(2)) & Space(40), 40) & Space(13) & Left(oImpre.ImpreCarEsp("ABONADO POR: " & sPerEcotaxi(2)) & Space(40), 40) & oImp.gPrnSaltoLinea
                        lscad = lscad & Space(sIni) & Left(oImpre.ImpreCarEsp("DNI: " & sPerEcotaxi(3)) & Space(40), 40) & Space(13) & Left(oImpre.ImpreCarEsp("DNI: " & sPerEcotaxi(3)) & Space(40), 40) & oImp.gPrnSaltoLinea
                    End If
                End If
                '*****************************************************************************************************************************************************
                'WIOR 20130419 *******************************************
                If psCodOpe = gPFApeEfec Or psCodOpe = gPFApeChq Or psCodOpe = gPFApeTransf Then
                   If EsDepositoGarantia(psCodCta) Then
                        lscad = lscad & oImp.gPrnNegritaON & Space(sIni) & oImpre.ImpreCarEsp("Depósito en Garantía ------------------") & Space(-1 + sSep) & oImpre.ImpreCarEsp("Depósito en Garantía ------------------") & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
                   End If
                End If
                'WIOR FIN ************************************************
                '***Modificado por ELRO el 20121029, según OYP-RFC101-2012
                'lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & Space(29 + sSep + saux) & oImpre.ImpreCarEsp(sCodUser) & oImp.gPrnSaltoLinea
                If psTipoCambio = "XXX" Then
                    lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & Space(29 + sSep + saux) & oImpre.ImpreCarEsp(sCodUser) & oImp.gPrnSaltoLinea
                Else
                    lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & Space(5) & psTipoCambio & Space(sSep + saux - 2) & oImpre.ImpreCarEsp(sCodUser) & Space(5) & psTipoCambio & oImp.gPrnSaltoLinea
                    
                End If
                '***Modificado por ELRO el 20121029**********************
            End If
        Else
            lscad = lscad & Space(sIni) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & Space(sSep) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & oImp.gPrnSaltoLinea
        End If

        lsMensaje = clsGen.GetMensajeBoletas(psCodCta, sCodage)
        Set clsGen = Nothing
        lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(lsMensaje) & Space(-1 + sSep) & oImpre.ImpreCarEsp(lsMensaje) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        lnNumLinCmac = lnNumLinCmac + 1
        
        If Left(psOpeCod, 4) = "2605" Or Left(psOpeCod, 4) = "1270" Or Left(psOpeCod, 4) = "1070" Then
            lscad = lscad & ""
        ElseIf bExtorno = True Then
            For saux = 1 To (lnTope - lnNumLinCmac - 1)
                lscad = lscad & ""
            Next saux
        ElseIf sProxReno <> "" Then
            For saux = 1 To (lnTope - lnNumLinCmac - 1)
                lscad = lscad & ""
            Next saux
        ElseIf Trim(psDisponible) <> "" And Trim(psIntangible) <> "" Then
           lscad = lscad & ""
        Else
            For saux = 1 To (lnTope - lnNumLinCmac)
                lscad = lscad & ""
            Next saux
        End If
        If psNumTarj <> "" Then
            lscad = lscad & "Operacion realizada con Tarjeta Electronica" & Space(14) & "Operacion realizada con Tarjeta Electronica" & oImp.gPrnSaltoLinea
        End If
        
        'ande ers021-2018
        If bPartipaCamp Then
            'cadenita = IIf(nCondicion = 0, "¡Gol!", "¡Autogol!") & " Con " & Chr(130) & "sta operaci" & Chr(162) & "n" & IIf(nCondicion = 0, " ganaste", " perdiste")
            'lscad = lscad & cadenita & IIf(nCondicion = 0, Space(21), Space(17)) & cadenita & oImp.gPrnSaltoLinea
            'cadenita = IIf(Len(CStr(nPunto)) < 5, CStr(nPunto) & String(5 - Len(CStr(nPunto)), " "), CStr(nPunto)) & " punto(s). Sigue participando en"
            'lscad = lscad & cadenita & Space(17) & cadenita & oImp.gPrnSaltoLinea
            'cadenita = "nuestra Campa" & Chr(164) & "a Mundial del Ahorro 2018"
            'APRI20191128
            cadenita = IIf(nCondicion = 0, "Felicitaciones sumas ", "Lo sentimos, perdiste ") & IIf(Len(CStr(nPunto)) < 5, CStr(nPunto) & String(5 - Len(CStr(nPunto)), " "), CStr(nPunto)) & " punto(s) para"
            lscad = lscad & cadenita & IIf(nCondicion = 0, Space(13), Space(12)) & cadenita & oImp.gPrnSaltoLinea
            cadenita = "el sorteo de practicos electrodomesticos"
            lscad = lscad & cadenita & Space(13) & cadenita & oImp.gPrnSaltoLinea
            cadenita = "y televisores para tu hogar!"
            'END APRI
            lscad = lscad & cadenita & Space(25) & cadenita & oImp.gPrnSaltoLinea
            cadenita = "Pto(s) acumulado(s): " & IIf(Len(CStr(nPTotalAcumulado)) < 6, CStr(nPTotalAcumulado) & String(6 - Len(CStr(nPTotalAcumulado)), " "), CStr(nPTotalAcumulado))
            lscad = lscad & cadenita & Space(26) & cadenita & oImp.gPrnSaltoLinea
        End If
        
        ImprimeBoleta = lscad & oImp.gPrnSaltoLinea
    Else 'TMU
        lscad = lscad & oImp.gPrnNegritaON  'Activa Negrita
        '***Modificado por ELRO el 20120524, según OYP-RFC005-2012 y OYP-RFC016-2012
        'lscad = lscad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea '& oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        If psCodOpe <> gOtrOpeDesParGas And psCodOpe <> gOtrOpeDesParVia And psCodOpe <> gOtrOpeDesParCaj Then
            lscad = lscad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea '& oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
        End If
        '***Fin Modificado por ELRO
        nComiTotalServicio = 0 'MADM 20111209
        nTotalServicio = 0 'MADM 20111209
        'RIRO20160111 ADD gOtrOpePagoRecaudoCajeroCorresponsal
        If pbPagoServicio Or psCodOpe = gOtrOpePagoRecaudoCajeroCorresponsal Then
            lscad = lscad & Space(sIni) & "CMAC-MAYNAS - SERVICIOS" & Space(16 + sSep + saux) & oImp.gPrnSaltoLinea
            psCliente = pcPersNomInst
            nComiTotalServicio = pnComisionPServiciox
            pcImrimePer = Trim(pcImrimePerx)
        Else
            lscad = lscad & Space(sIni) & "CMAC-MAYNAS - AHORRO" & Space(16 + sSep + saux) & oImp.gPrnSaltoLinea
        End If
        
        If Len(Trim(psCodCta)) < 18 Then
            If pnMonedaCMAC = 1 Then
                '''lscad = lscad & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
                lscad = lscad & Space(sIni) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
            Else
                lscad = lscad & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & oImp.gPrnSaltoLinea
            End If
        Else
            If Mid(psCodCta, 9, 1) = 1 Then
                '''lscad = lscad & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
                lscad = lscad & Space(sIni) & Trim(sNomAge) & "-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase) & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-" & StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase))) + lsNroExt & oImp.gPrnSaltoLinea 'marg ers044-2016
            Else
                lscad = lscad & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & oImp.gPrnSaltoLinea
            End If
        End If

        If psNomAgeRem = "" Then
            lscad = lscad & ""
        Else
            lscad = lscad & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(saux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & oImp.gPrnSaltoLinea
        End If
        
        If psComCmac = "XXX" Then
            If psLin3 = "XXX" Then
                lscad = lscad & oImp.gPrnNegritaOFF 'Desactiva Negrita
            Else
                lscad = lscad & Space(sIni) & psLin3 & Space(saux + sSep + sMax - Len(psLin3)) & psLin3 & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea  'Desactiva Negrita
                lnNumLinCmac = 1
            End If
            lnNumLinCmac = 0
        Else
            lscad = lscad & Space(sIni) & "NroDocCmac:" & psComCmac & Space(saux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea 'Desactiva Negrita
            lnNumLinCmac = 1
        End If
        
        lscad = lscad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(saux + sSep - 6) & oImp.gPrnSaltoLinea

        lnCliAux = InStr(1, psCliente, "*", vbTextCompare)

        If lnCliAux = 0 Then
            If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(psCliente) & Space(saux + sMax + sSep - Len(psCliente)) & oImp.gPrnSaltoLinea
        Else
            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
    
            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
            
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux1) & oImpre.ImpreCarEsp(lsCliAux1) & oImp.gPrnSaltoLinea
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux2) & oImpre.ImpreCarEsp(lsCliAux2) & oImp.gPrnSaltoLinea
            'Space(saux + sMax + sSep - Len(lsCliAux2))
            lnCliAux = 1
        End If
        
        If pbSaldoInt Or pbCuenta Then
            If lsCodAntiguo <> "" Then  ' Para Codigo Antiguo - LAYG
                lscad = lscad & Space(sIni) & "Cuenta:" & lsCodAntiguo & oImp.gPrnSaltoLinea
                'lsCadImp = lsCadImp & " Contrato  : " & lsCodAntiguo & Space(10) & Space(1 + vEspacio) _
                                            & " Contrato  : " & lsCodAntiguo & Space(10) & Chr(10)
            Else
                lscad = lscad & Space(sIni) & "Cuenta:" & psCodCta & oImp.gPrnSaltoLinea
            End If
        Else
            lscad = lscad & ""
        End If
        
        psTit = Trim(psTit)
        psTit = oImpre.CentrarCadena(psTit, 28)
        lscad = lscad & oImp.gPrnNegritaON 'Activa Negrita
        lscad = lscad & Space(sIni + 1) & "-----" & psTit & "-----" & oImp.gPrnSaltoLinea
        lscad = lscad & oImp.gPrnNegritaOFF
        lnChq = InStr(1, pstext, "*", vbTextCompare)
        
        If psTexto = "XXX" Then
            If lnChq = 0 Then
                'lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(Trim(Mid(psText, 1, 28))) & Space(sMax - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
                lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(pstext, 1, 28))) & Space((sMax + 6 - (Len(pstext) + Len(Format(sMonto, "0.00"))))) & Format(sMonto, "0.00") & oImp.gPrnSaltoLinea
                 'MADM 20111209
                If pbPagoServicio Then
                     Dim xxx1 As Integer
                     Dim psConceptoSerTemp1 As String
                     psConceptoSerTemp1 = pnConceptoSer
                     pnConceptoSer = Trim(pnConceptoSer)
                     pnMontoSer = Trim(pnMontoSer)
                     pnConceptoSer = Mid(pnConceptoSer, 1, Len(pnConceptoSer) - 5)
                     If Int(Right(psConceptoSerTemp1, 5)) = 1 Then
                            pnConceptoSer = pnConceptoSer & " XXX"
                            pnMontoSer = pnMontoSer & " XXX"
                     End If
                     lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(psNomCliServ, 1, 30))) & oImp.gPrnSaltoLinea
                     For xxx1 = 1 To Int(Right(psConceptoSerTemp1, 5))
                        lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(Mid(pnConceptoSer, 4, InStr(4, pnConceptoSer, "XXX", vbTextCompare) - 4), 1, 31))) & Space(sMax + 6 - (Len(Trim(Mid(Mid(pnConceptoSer, 4, InStr(4, pnConceptoSer, "XXX", vbTextCompare) - 4), 1, 31))) + Len(Trim(Mid(Mid(pnMontoSer, 4, InStr(4, pnMontoSer, "XXX", vbTextCompare) - 4), 1, 8))))) & oImpre.ImpreCarEsp(Trim(Mid(Mid(pnMontoSer, 4, InStr(4, pnMontoSer, "XXX", vbTextCompare) - 4), 1, 8))) & oImp.gPrnSaltoLinea
                        pnConceptoSer = Trim(Mid(pnConceptoSer, InStr(5, pnConceptoSer, "XXX", vbTextCompare), Len(pnConceptoSer) - 6))
                        If Right(Trim(pnConceptoSer), 3) <> "XXX" Then
                            pnConceptoSer = pnConceptoSer & " XXX"
                        End If
                         If Right(Trim(pnMontoSer), 3) <> "XXX" Then
                            pnMontoSer = pnMontoSer & " XXX"
                        End If
                     Next xxx1
                End If
                lscad = lscad & ""
            Else
                lsChqAux1 = (Mid(pstext, 1, lnChq - 1))
                lsChqAux2 = (Mid(pstext, lnChq + 1))
                lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Space(-1 + sSep) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
            End If
        Else
            lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
            If psDisponible = "" And psIntangible = "" Then
                lscad = lscad & ""
            End If
        End If
        
        If psDisponible <> "" And psIntangible <> "" Then
           'psDispCta psIntaCta
           'gPrnCondensadaOFF
           'gPrnCondensadaON
           lscad = lscad & oVarImpre.gPrnCondensadaON & Space(sIni) & oImpre.ImpreCarEsp(Trim(psDispCta)) & oImp.gPrnSaltoLinea
           lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psIntaCta)) & oImp.gPrnSaltoLinea
           lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psDisponible)) & oImp.gPrnSaltoLinea
           lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psIntangible)) & Space(saux + sSep + sMax - Len(Trim(psIntangible))) & oImpre.ImpreCarEsp(Trim(psIntangible)) & oVarImpre.gPrnCondensadaOFF & oImp.gPrnSaltoLinea

        End If
        'Insertar ITF
        If pbImpItf = True Then
            If pnMontoIFT <> 0 Then
                If pnTipoPago = 1 Then
                    psTexto = "ITF Efectivo"
                 
                ElseIf pnTipoPago = 2 Then
                    psTexto = "ITF Cargo a Cuenta"
                End If
                lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(pnMontoIFT)) & Format(pnMontoIFT, "0.00") & oImp.gPrnSaltoLinea
            End If
        End If
        '***    IMPRIMIR COMISON DE RETIRO -- AVMM 10-08-2006  ***
        'nComiTotal = pnComixRetOtraAge + pnComixMaxRet 'Comentado by JACA 20111129
        'If nComiTotal <> 0 Then'Comentado by JACA 20111129
            If pnComixRetOtraAge <> 0 Then
               'psTexto = "Comisión Operación Otra Age."
               psTexto = "Comisión " & IIf(Left(psCodOpe, 4) = "2002", "Dep.", IIf(Left(psCodOpe, 4) = "2003", "Ret.", "Ope.")) & " Otra Loc." 'JUEZ 20150930
               lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(pnComixRetOtraAge)) & Format(pnComixRetOtraAge, "0.00") & oImp.gPrnSaltoLinea
            End If
            If pnComixMaxRet <> 0 Then
               psTexto = "Comisión x Max Nro Ope."
               lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 28))) & Space(37 + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(pnComixMaxRet)) & Format(pnComixMaxRet, "0.00") & oImp.gPrnSaltoLinea
            End If
        'End If
        '*********************************************************
        
        'Add by GITU 30-09-2011
        'If pnComiVB <> 0 And psCodOpeVB = "300627" Then 'Comentado CTI4 ERS0112020
        If pnComiVB <> 0 Then
           'psTexto = "Comision Ope Sin Tarjeta" 'Comentado CTI4 ERS0112020
           'CTI4 ERS0112020
            Select Case psCodOpeVB
                Case "300627"
                    psTexto = "Comisión Ope Sin Tarjeta"
                Case gComisionEmisionChequeCargoCta
                    psTexto = "Comisión Cargo a Cuenta"
                Case gComisionEmisionChequeEfectivo
                    psTexto = "Comisión Efectivo"
            End Select
           'CTI4 End
            lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 26))) & Space(37 + 6 - Len(Trim(Mid(psTexto, 1, 28))) - Len(pnComiVB)) & Format(pnComiVB, "0.00") & oImp.gPrnSaltoLinea
        End If
       'END GITU
       
       'RIRO20131212 ERS137
       If Len(Trim(nComisionTrans)) <> 0 Then
           psTexto = "Comision Transf:"
           lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 26))) & Space(39 - Len(psTexto) - Len(Format(nComisionTrans, "0.00"))) & Format(nComisionTrans, "0.00") & oImp.gPrnSaltoLinea
       End If
       'END RIRO
       
       'MADM 20111209
        If nComiTotalServicio <> 0 Then
           nTotalServicio = nComiTotalServicio + pnMontoIFT + sMonto 'Format(nComiTotalServicio + pnMontoIFT + sMonto, "#,###,##0.00")
           psTexto = "Comision Pago Servicios"
           'lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 27))) & Space(sMax + 6 - (Len(Trim(Mid(psTexto, 1, 27)))) + Len(Format(nComiTotalServicio, "#,###,##0.00"))) & oImpre.ImpreCarEsp(Trim(Format(nComiTotalServicio, "#,###,##0.00"))) & oImp.gPrnSaltoLinea
           lscad = lscad & oImpre.ImpreCarEsp(Trim(psTexto)) & Space(sMax + 6 - (Len(psTexto) + Len(Format(nComiTotalServicio, "0.00")))) & Format(nComiTotalServicio, "0.00") & oImp.gPrnSaltoLinea
           psTexto = "Total Pagado"
           lscad = lscad & oImpre.ImpreCarEsp(Trim(psTexto)) & Space(sMax + 6 - (Len(psTexto) + Len(Format(nTotalServicio, "#,###,##0.00")))) & Format(nTotalServicio, "#,###,##0.00") & oImp.gPrnSaltoLinea
           'lscad = lscad & oImpre.ImpreCarEsp(Trim(Mid(psTexto, 1, 27))) & Space(sMax + 6 - (Len(Mid(psTexto, 1, 27))) + Len(Format(nComiTotalServicio, "#,###,##0.00"))) & oImpre.ImpreCarEsp(Trim(Format(nComiTotalServicio, "#,###,##0.00"))) & oImp.gPrnSaltoLinea
        End If
       'END MADM
        'If pbSaldoInt Then
        If bPregImpSaldos Then
            'If MsgBox("Desea Imprimir los Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            If pbImprimeSaldos Then
                bImpreSaldos = True
                lscad = lscad & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & oImp.gPrnSaltoLinea
                If pbOpSaldoC Then
                    lscad = lscad & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & oImp.gPrnSaltoLinea
                Else
                    lscad = lscad & ""
                End If
            Else
                bImpreSaldos = False
                lscad = lscad & ""
                lscad = lscad & ""
                pbSaldoInt = False
            End If
        Else
            lscad = lscad & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & oImp.gPrnSaltoLinea
            If pbOpSaldoC Then
               lscad = lscad & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & oImp.gPrnSaltoLinea
            Else
                lscad = lscad & ""
            End If
        End If
        'Else
        '   lscad = lscad & ""
        'End If

        lsInteres = pnInteresA
        
        If pbSaldoInt Then
            If lsInteres <> "No Valido" Then
                lsInteres = Format(lsInteres, "#,##0.00")
                lscad = lscad & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & oImp.gPrnSaltoLinea
            End If
        Else
            lscad = lscad & ""
        End If

        If Trim(sProxReno) <> "" Then
              lscad = lscad & Space(sIni) & sProxReno & Space(sMax + 20 - Len(sProxReno)) & sProxReno & oImp.gPrnSaltoLinea
        End If
        
        
        If Not psNumDoc = "" Then
            lscad = lscad & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & oImp.gPrnSaltoLinea
        Else
            If Trim(psDisponible) = "" And Trim(psIntangible) = "" Then
                lscad = lscad & ""
            End If
        End If
        
        If Not psNumDias = "----" Then
            lscad = lscad & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & oImp.gPrnSaltoLinea
            lnTope = 4 - lnCliAux
        Else
            lnTope = 3 - lnCliAux
        End If


        lscad = lscad & oImp.gPrnNegritaON & Space(sIni) & "---------------------------------------" & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        If psCodUsuRem = "" Then
            If bExtorno Then
                lscad = lscad & Space(sIni) & "USUARIO EXTORNA : " & oImpre.ImpreCarEsp(sCodUser) & oImp.gPrnSaltoLinea
                lscad = lscad & Space(sIni) & "USUARIO OPERADOR: " & oImpre.ImpreCarEsp(sUserBus) & oImp.gPrnSaltoLinea
            Else
                'WIOR 20130419 *******************************************
                If psCodOpe = gPFApeEfec Or psCodOpe = gPFApeChq Or psCodOpe = gPFApeTransf Then
                   If EsDepositoGarantia(psCodCta) Then
                        lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp("Depósito en Garantía ------------------") & oImp.gPrnSaltoLinea
                   End If
                End If
                'WIOR FIN ************************************************
                lscad = lscad & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & oImp.gPrnSaltoLinea
            End If
        Else
            lscad = lscad & Space(sIni) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & oImpre.ImpreCarEsp(sCodUser) & "/" & oImpre.ImpreCarEsp(psCodUsuRem) & oImp.gPrnSaltoLinea
        End If

        lsMensaje = clsGen.GetMensajeBoletas(psCodCta, sCodage)
        Set clsGen = Nothing
        lscad = lscad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(lsMensaje) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
        lnNumLinCmac = lnNumLinCmac + 1
        
        If Left(psOpeCod, 4) = "2605" Or Left(psOpeCod, 4) = "1270" Or Left(psOpeCod, 4) = "1070" Then
            lscad = lscad & ""
        ElseIf bExtorno = True Then
            For saux = 1 To (lnTope - lnNumLinCmac - 1)
                lscad = lscad & ""
            Next saux
        ElseIf sProxReno <> "" Then
            For saux = 1 To (lnTope - lnNumLinCmac - 1)
                lscad = lscad & ""
            Next saux
        ElseIf Trim(psDisponible) <> "" And Trim(psIntangible) <> "" Then
           lscad = lscad & ""
        Else
            For saux = 1 To (lnTope - lnNumLinCmac)
                lscad = lscad & ""
            Next saux
        End If
        If psNumTarj <> "" Then
            lscad = lscad & "Operacion realizada con Tarjeta Electronica"
        End If
        If bPartipaCamp Then
            cadenita = IIf(nCondicion = 0, "Felicitaciones sumas ", "Lo sentimos, perdiste ") & IIf(Len(CStr(nPunto)) < 5, CStr(nPunto) & String(5 - Len(CStr(nPunto)), " "), CStr(nPunto)) & " punto(s) para"
            lscad = lscad & cadenita & oImp.gPrnSaltoLinea
            cadenita = "el sorteo de practicos electrodomesticos"
            lscad = lscad & cadenita & oImp.gPrnSaltoLinea
            cadenita = "y televisores para tu hogar!"
            'END APRI
            lscad = lscad & cadenita & oImp.gPrnSaltoLinea
            cadenita = "Pto(s) acumulado(s): " & CStr(nPTotalAcumulado)
            lscad = lscad & cadenita & oImp.gPrnSaltoLinea
        End If
 
        If psCodOpe <> gOtrOpeDesParGas And psCodOpe <> gOtrOpeDesParVia And psCodOpe <> gOtrOpeDesParCaj Then
            'Linea que corta el ticket en TMU
            lscad = lscad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
        ElseIf psCodOpe = gOtrOpeDesParCaj Then
            lscad = lscad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
        End If
        '***Fin Modificado por ELRO*******************************
        
        ImprimeBoleta = lscad
    End If
Exit Function
error:
    Err.Raise Err.Number, "Comprueba la conexion de su impresora... ", Err.Description

End Function

Public Function ValidaSaldoCuenta_abono(ByVal sCuenta As String, ByVal nMonto As Double, Optional nOperacion As CaptacOperacion, _
                                  Optional pnComiRetOtraAge As Double, Optional pnComix31Ope As Double, _
                                  Optional pnComRetSintarj As Double, Optional ByVal pnComTransf As Double = 0, _
                                  Optional ByVal pbExtornoAbono As Boolean = False, Optional nMontoDep As Double = 0) As Boolean
                                  'RIRO20131212 ERS137 - Se agregó pnComTransf
                                  'FRHU CORREO 20150414 Se agrego pbExtornoAbono
                                  'APRI20190109 ERS077-2018 Add nMontoDep
Dim nMoneda As Moneda
Dim nProducto As Producto
'Dim clsMant As NCOMCaptaGenerales
'Dim clsPar As NCOMCaptaDefinicion
Dim rsCta As New ADODB.Recordset
Dim nSaldo As Double, nMontoMinimo As Double
Dim bOrdPag As Boolean
Dim nPersoneria As PersPersoneria
Dim nEstado As CaptacEstado
Dim nBloqueoParcial As Double
'Set clsMant = New NCOMCaptaGenerales
Set rsCta = GetDatosCuenta(sCuenta)
nProducto = CLng(Mid(sCuenta, 6, 3))
bOrdPag = False
If nProducto = gCapAhorros Then
    nSaldo = rsCta("nSaldoDisp")
    bOrdPag = rsCta("bOrdPag")
ElseIf nProducto = gCapCTS Then
    If nOperacion = gCTSDepChq Or nOperacion = gCTSDepEfec Or nOperacion = gCTSDepPlanRRHH Or nOperacion = gCTSDepTransf Or nOperacion = "220303" Then
        nSaldo = rsCta("nSaldo")
    Else
        nSaldo = rsCta("nSaldRetiro")
    End If
End If

nPersoneria = rsCta("nPersoneria")
nEstado = rsCta("nPrdEstado")
'nBloqueoParcial = rsCta("nBloqueoParcial")
nBloqueoParcial = IIf(nMontoDep > 0, 0, rsCta("nBloqueoParcial")) 'APRI20200113 INC2001130013

rsCta.Close
Set rsCta = Nothing
'Set clsMant = Nothing
nMoneda = CLng(Mid(sCuenta, 9, 1))

'If nEstado <> gCapEstActiva then
    'ValidaSaldoCuenta = False
'Else
'If nEstado = gCapEstActiva Or (nEstado = gCapEstBloqRetiro And pbExtornoAbono) Then     'FRHU CORREO 20150414
If nEstado = gCapEstActiva Or (nEstado = gCapEstBloqRetiro And pbExtornoAbono) Or nMontoDep > 0 Then  'APRI20191107 MEJORA
    If CLng(Mid(sCuenta, 6, 3)) = gCapAhorros Then
        'Set clsPar = New NCOMCaptaDefinicion
        nMontoMinimo = GetSaldoMinimoPersoneria(nProducto, nMoneda, nPersoneria, bOrdPag)
       ' Set clsPar = Nothing
    Else
        nMontoMinimo = 0
    End If
    ' Agrego montos de Comision AVMM -- 03-06-2006
    ' RIRO20131212 ERS13 - Se agregó pnComTransf
    'If nSaldo - nBloqueoParcial - nMontoMinimo - nMonto - pnComiRetOtraAge - pnComix31Ope - pnComRetSintarj - pnComTransf >= 0 Then
    'If Round(nSaldo - nBloqueoParcial - nMontoMinimo - nMonto - pnComiRetOtraAge - pnComix31Ope - pnComRetSintarj - pnComTransf, 4) >= 0 Then 'JUEZ 20141021 Para corregir casos de decimales
    If Round((nSaldo + nMontoDep) - nBloqueoParcial - nMontoMinimo - nMonto - pnComiRetOtraAge - pnComix31Ope - pnComRetSintarj - pnComTransf, 4) >= 0 Then 'APRI20190109 ERS077-2018
        ValidaSaldoCuenta_abono = True
    Else
        ValidaSaldoCuenta_abono = False
    End If
'End If
Else
    ValidaSaldoCuenta_abono = False
End If

End Function
Public Function GetSaldoMinimoPersoneria(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal nPersoneria As PersPersoneria, ByVal bOrdPag As Boolean, _
        Optional ByVal nTpoPrograma As Integer = 0, Optional ByVal psCtaCod As String = "") As Double
'APRI20190109 ERS077-2018 add psCtaCod
    Dim rsPar As ADODB.Recordset
    Dim sOrdPag As String
    sOrdPag = IIf(bOrdPag, "1", "0")
    Set rsPar = New ADODB.Recordset
    rsPar.CursorLocation = adUseClient
'    sSql = "Select nSaldoMin From CapPersParam Where nPersoneria = " & nPersoneria & " And " _
'        & "nProducto = " & nProducto & " And nMoneda = " & nMoneda & " And cOrdPag = '" & sOrdPag & "' " _
'        & "And nTpoPrograma = " & nTpoPrograma
    sSql = "EXEC stp_sel_GetSaldoMinimoPersoneria " & nPersoneria & "," & nProducto & "," & nMoneda & ",'" & sOrdPag & "'," & nTpoPrograma & ",'" & psCtaCod & "'" 'APRI20190109 ERS077-2018
    rsPar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsPar.ActiveConnection = Nothing
    If rsPar.EOF And rsPar.BOF Then
        GetSaldoMinimoPersoneria = 0
    Else
        GetSaldoMinimoPersoneria = rsPar("nSaldoMin")
    End If
    rsPar.Close
    Set rsPar = Nothing
End Function
Public Function GetNombreTitulares(ByVal sCuenta As String, Optional pbNombreCompleto As Boolean = False, Optional pnTitulares As Integer = 0, Optional pnMargenIzquierdo As Integer = 0) As String
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
GetNombreTitulares = clsMant.GetNombreTitulares(sCuenta, pbNombreCompleto, pnTitulares, pnMargenIzquierdo)
Set clsMant = Nothing
End Function

Public Function GetTpoCta(ByVal sCuenta As String) As String
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim R As ADODB.Recordset
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales

Set R = clsMant.ConsultaTpoCta(sCuenta)

If Not R Is Nothing Then
    If Not R.EOF Then
        If R.RecordCount > 0 Then
           GetTpoCta = R!codLeyenda
        End If
    End If
End If
Set clsMant = Nothing
End Function

Public Function RecuperaMovimientoCapatacionesOperaciones(ByVal pnMovNro As Long, Optional ByVal psOpeCod As String = "") As ADODB.Recordset
Dim sSql As String

    
    sSql = "exec stp_sel_RecuperaMovimientoCapatacionesOperaciones " & pnMovNro & ",'" & psOpeCod & "'"
    
    Set RecuperaMovimientoCapatacionesOperaciones = New ADODB.Recordset
    
    RecuperaMovimientoCapatacionesOperaciones.CursorLocation = adUseClient
    RecuperaMovimientoCapatacionesOperaciones.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    RecuperaMovimientoCapatacionesOperaciones.ActiveConnection = Nothing

End Function
'******************************************************************************

