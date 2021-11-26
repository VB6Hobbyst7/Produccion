VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dFuncionesNeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public oImp As COMFunciones.FCOMVarImpresion
'Public oImp As NImpresion

Public Enum Moneda
    gMonedaNacional = 1
    gMonedaExtranjera = 2
End Enum

Dim dbCmact As COMConecta.DCOMConecta

Dim sSQL As String
Public Enum TpoDoc
    TpoDocCheque = 47
    TpoDocOrdenPago = 48
    TpoDocNotaCargo = 57
    TpoDocNotaAbono = 58
End Enum

Public Function GeneraMovNro(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String
    On Error GoTo GeneraMovNroErr
    Dim rs As ADODB.Recordset
    Dim Sql As String
    Set rs = New ADODB.Recordset
    Set dbCmact = New COMConecta.DCOMConecta

    If dbCmact.AbreConexion = False Then Exit Function
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       Sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & dbCmact.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & Right(psCodAge, 2) & "','" & psUser & "'"
    Else
       Sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    Set rs = dbCmact.Ejecutar(Sql)
    If Not rs.EOF Then
        GeneraMovNro = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
    dbCmact.CierraConexion
    'Set dbCmact = Nothing
    Exit Function
GeneraMovNroErr:
    Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function

Public Function GetCapParametro(ByVal nParametro As CaptacParametro) As Double
    Dim rsVar As ADODB.Recordset
    Set dbCmact = New COMConecta.DCOMConecta
    dbCmact.AbreConexion
    
    sSQL = "SELECT nParValor FROM Parametro WHERE nParCod = " & nParametro & " " _
        & "And nParProd = " & gPrdParamCaptac
    Set rsVar = New ADODB.Recordset
    rsVar.CursorLocation = adUseClient
    
    'rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsVar = dbCmact.CargaRecordSet(sSQL)
    
    Set rsVar.ActiveConnection = Nothing
    If rsVar.EOF And rsVar.BOF Then
        GetCapParametro = 0
    Else
        GetCapParametro = rsVar("nParValor")
    End If
    rsVar.Close
    dbCmact.CierraConexion
    Set rsVar = Nothing
End Function




Public Function IncrementaMovNro(ByVal cMovnro As String) As String
Dim cTexto1 As String
Dim cTexto2 As String
Dim cTexto3 As String
Dim cNumero As Integer
Dim nNumero As Integer
Dim sNumero As String
    cTexto1 = Mid(cMovnro, 1, 19)
    cTexto2 = Mid(cMovnro, 20, 2)
    cTexto3 = Mid(cMovnro, 22, 4)
    nNumero = Val(cTexto2) + 1
    sNumero = Right("00" & Trim(Str(nNumero)), 2)
    IncrementaMovNro = cTexto1 + sNumero + cTexto3
End Function


Public Function PstaNombre(psNombre As String, Optional pbNombApell As Boolean = False) As String
Dim Total As Long
Dim Pos As Long
Dim CadAux As String
Dim lsApellido As String
Dim lsNombre As String
Dim lsMaterno As String
Dim lsConyugue As String
Dim CadAux2 As String
Dim posAux As Integer
Dim lbVda As Boolean
lbVda = False
Total = Len(Trim(psNombre))
Pos = InStr(psNombre, "/")
If Pos <> 0 Then
    lsApellido = Left(psNombre, Pos - 1)
    CadAux = Mid(psNombre, Pos + 1, Total)
    Pos = InStr(CadAux, "\")
    If Pos <> 0 Then
        lsMaterno = Left(CadAux, Pos - 1)
        CadAux = Mid(CadAux, Pos + 1, Total)
        Pos = InStr(CadAux, ",")
        If Pos > 0 Then
            CadAux2 = Left(CadAux, Pos - 1)
            posAux = InStr(CadAux, "VDA")
            If posAux = 0 Then
                lsConyugue = CadAux2
            Else
                lbVda = True
                lsConyugue = CadAux2
            End If
        Else
            lsMaterno = CadAux
        End If
    Else
        CadAux = Mid(CadAux, Pos + 1, Total)
        Pos = InStr(CadAux, ",")
        If Pos <> 0 Then
            lsMaterno = Left(CadAux, Pos - 1)
            lsConyugue = ""
        Else
            lsMaterno = CadAux
        End If
    End If
    lsNombre = Mid(CadAux, Pos + 1, Total)
    If pbNombApell = True Then
        If Len(Trim(lsConyugue)) > 0 Then
            PstaNombre = Trim(lsNombre) & " " & Trim(lsApellido) & " " & Trim(lsMaterno) & IIf(lbVda = False, " DE ", " ") & Trim(lsConyugue)
        Else
            PstaNombre = Trim(lsNombre) & " " & Trim(lsApellido) & " " & Trim(lsMaterno)
        End If
    Else
        If Len(Trim(lsConyugue)) > 0 Then
            PstaNombre = Trim(lsApellido) & " " & Trim(lsMaterno) & IIf(lbVda = False, " DE ", " ") & Trim(lsConyugue) & " " & Trim(lsNombre)
        Else
            PstaNombre = Trim(lsApellido) & " " & Trim(lsMaterno) & " " & Trim(lsNombre)
        End If
    End If
Else
    PstaNombre = Trim(psNombre)
End If
End Function


Public Function ImpreCarEsp(ByVal vCadena As String) As String
    vCadena = Replace(vCadena, "á", Chr(160), , , vbTextCompare)
    vCadena = Replace(vCadena, "é", Chr(130), , , vbTextCompare)
    vCadena = Replace(vCadena, "í", Chr(161), , , vbTextCompare)
    vCadena = Replace(vCadena, "ó", Chr(162), , , vbTextCompare)
    vCadena = Replace(vCadena, "ú", Chr(163), , , vbTextCompare)
    vCadena = Replace(vCadena, "ñ", Chr(164), , , vbTextCompare)
    vCadena = Replace(vCadena, "Ñ", Chr(165), , , vbTextCompare)
    vCadena = Replace(vCadena, "°", Chr(248), , , vbTextCompare)
    vCadena = Replace(vCadena, "¦", Chr(179), , , vbTextCompare)
    ImpreCarEsp = vCadena
End Function

Public Function GetCaptaOperacion(psOpeCod As String) As String
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Set dbCmact = New COMConecta.DCOMConecta
    Set rs = New ADODB.Recordset
    
    If dbCmact.AbreConexion = False Then Exit Function
    
    Sql = "Select IsNull(cOpeDesc,'') cOpeDesc From OpeTpo where copecod = '" & Trim(psOpeCod) & "'"
    'rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rs = dbCmact.CargaRecordSet(Sql)
    If rs.EOF And rs.BOF Then
        GetCaptaOperacion = ""
    Else
        GetCaptaOperacion = rs!cOpeDesc
    End If
    dbCmact.CierraConexion
End Function

Public Function LeeConstSistema(ByVal psConstSistCod As ConstSistemas) As String
    Dim rsVar As Recordset
    Dim sSQL As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    sSQL = "Select nConsSisDesc, nConsSisValor From ConstSistema where nConsSisCod =" & psConstSistCod & ""
    Set rsVar = New Recordset
    Set rsVar = oCon.CargaRecordSet(sSQL)
    LeeConstSistema = ""
    If Not rsVar.EOF And Not rsVar.BOF Then
        LeeConstSistema = rsVar("nConsSisValor")
    End If
    rsVar.Close
    Set rsVar = Nothing
    Set oCon = Nothing
End Function

'Public Function GetDatosCuenta(ByVal sCuenta As String) As ADODB.Recordset
'Dim rsCta As ADODB.Recordset
'Dim nProd As Producto
''Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
'nProd = CInt(Mid(sCuenta, 6, 3))
''Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
'Select Case nProd
'    Case gCapAhorros
'        Set rsCta = GetDatosCuentaAho(sCuenta)
'
'    Case gCapPlazoFijo
'        Set rsCta = GetDatosCuentaPF(sCuenta)
'
'    Case gCapCTS
'        Set rsCta = GetDatosCuentaCTS(sCuenta)
'
'End Select
'Set GetDatosCuenta = rsCta
'Set rsCta = Nothing
'Set clsMant = Nothing
'End Function


Public Sub ActualizaSaldoAnteriorAho(ByVal sCuenta As String, ByVal nSaldoAnt As Double)
    Set dbCmact = New COMConecta.DCOMConecta
    dbCmact.AbreConexion
    
    sSQL = "Update CaptacAhorros Set nSaldoAnterior = " & nSaldoAnt & "  WHERE cCtaCod = '" & sCuenta & "'"
    dbCmact.Ejecutar (sSQL)
    dbCmact.CierraConexion
End Sub


Public Property Get gsFormatoFecha() As Variant
    gsFormatoFecha = "mm/dd/yyyy"
End Property

Public Function ValidaEstadoCuenta(ByVal sCuenta As String, Optional pbParaRetiro As Boolean = True) As Boolean
'Dim clsMant As NCOMCaptaGenerales
'Set clsMant = New NCOMCaptaGenerales
Dim rsCta As ADODB.Recordset
Set rsCta = New ADODB.Recordset
Dim nEstado As CaptacEstado

'Set rsCta = GetDatosCuenta(sCuenta)

'nEstado = rsCta("nPrdEstado")

rsCta.Close
Set rsCta = Nothing
'Set clsMant = Nothing

If pbParaRetiro Then
    If nEstado = gCapEstActiva Then
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


Public Function GetTarifaParametro(ByVal nOperacion As CaptacOperacion, _
            ByVal nMoneda As Moneda, ByVal nParametro As CaptacParametro) As Recordset
    Dim rsPar As ADODB.Recordset
    Set dbCmact = New COMConecta.DCOMConecta
    dbCmact.AbreConexion
    
    Set rsPar = New ADODB.Recordset
    rsPar.CursorLocation = adUseClient
    sSQL = "Select P.nParValor From Parametro P JOIN Tarifas T ON P.nParProd = T.nParProd And " _
        & "P.nParCod = T.nParCod Where T.cOpeCod = '" & nOperacion & "' And " _
        & "T.nMoneda = " & nMoneda & " And T.nParCod = " & nParametro
    
    Set rsPar = dbCmact.CargaRecordSet(sSQL)
    'rsPar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    Set rsPar.ActiveConnection = Nothing
    Set GetTarifaParametro = rsPar
    Set rsPar = Nothing
    dbCmact.CierraConexion

End Function

Public Function obtenerTipoOperacionesInterCMAC() As ADODB.Recordset
Dim lCn As COMConecta.DCOMConecta
Dim lsSql As String
    
    lsSql = " exec PIT_stp_sel_OperacionesInterCMACSEnvio "
    
    Set lCn = New COMConecta.DCOMConecta
    
    lCn.AbreConexion
    
    Set obtenerTipoOperacionesInterCMAC = lCn.CargaRecordSet(lsSql)
    
    lCn.CierraConexion
    
    Set lCn = Nothing
End Function

Public Function obtenerCMACS() As ADODB.Recordset
Dim lCn As COMConecta.DCOMConecta
Dim lsSql As String
    
    lsSql = " exec PIT_stp_sel_CMACS "
    
    Set lCn = New COMConecta.DCOMConecta
    lCn.AbreConexion
    Set obtenerCMACS = lCn.CargaRecordSet(lsSql)
    lCn.CierraConexion
    Set lCn = Nothing
End Function

Public Function obtenerDatosAgencia(psAgeCod As String) As ADODB.Recordset
Dim lCn As COMConecta.DCOMConecta
Dim lsSql As String
    
    lsSql = " exec PIT_stp_sel_Agencia '" & psAgeCod & "'"
    
    Set lCn = New COMConecta.DCOMConecta
    lCn.AbreConexion
    Set obtenerDatosAgencia = lCn.CargaRecordSet(lsSql)
    lCn.CierraConexion
    Set lCn = Nothing
End Function



'Public Sub GetOperaciones(rsOpeLLama As ADODB.Recordset, ByRef rsCMACs As ADODB.Recordset)
'    Set dbCmact = New COMConecta.DCOMConecta
'    dbCmact.AbreConexion
'
'    sSQL = " exec PIT_stp_sel_OperacionesInterCMACSEnvio"
'
'    Set rsOpeLLama = New ADODB.Recordset
'    Set rsOpeLLama = dbCmact.CargaRecordSet(sSQL)
'
'
'    sSQL = "exec PIT_stp_sel_CMACS"
'    Set rsCMACs = New ADODB.Recordset
'    Set rsCMACs = dbCmact.CargaRecordSet(sSQL)
'
'    dbCmact.CierraConexion
'    Set dbCmact = Nothing
'
'End Sub

Public Function GetVarSistema() As ADODB.Recordset
    Dim rsPar As ADODB.Recordset
    Set dbCmact = New COMConecta.DCOMConecta
    dbCmact.AbreConexion
    
    Set rsPar = New ADODB.Recordset
    sSQL = " Select nConsSisCod,nConsSisValor,nConsSisDesc From ConstSistema " _
            & " Where nConsSisCod in (" & gConstSistFechaSistema & "," & gConstSistNombreAbrevCMAC & "," _
            & gConstSistRutaBackup & "," & gConstSistCodCMAC & "," & gConstSistMargenSupCartas & "," & gConstSistMagenIzqCartas & "," _
            & gConstSistMargenDerCartas & "," & gConstSistNroLineasPagina & "," & gConstSistNroLineasOrdenPago & "," _
            & gConstSistCtaConversionMEDol & "," & gConstSistCtaConversiónMESoles & "," & gConstSistTipoConverión & "," _
            & gConstSistNombreModulo & "," & gConstSistFechaInicioDia & "," & gConstSistDominio & "," & gConstSistPDC _
            & ",40,50," & gConstPersCodCMACT & "," & gConstSistAgenciaEspecial & "," & gConstSistCierreMesNegocio & ", " _
            & gConstSistRutaIcono & ",84,85," & gConstSistVerificaRegistroEfectivo & ",301,302,303,304,305) " _
            & "ORDER BY nConsSisCod"

    Set rsPar = dbCmact.CargaRecordSet(sSQL)
    
    Set rsPar.ActiveConnection = Nothing
    Set GetVarSistema = rsPar
    Set rsPar = Nothing
    dbCmact.CierraConexion
End Function
Public Function DevuelveHoraServer() As String
    Dim rsH As New ADODB.Recordset
    Set dbCmact = New COMConecta.DCOMConecta

    If dbCmact.AbreConexion = False Then Exit Function
    sSQL = "Select Convert(varchar(10),getdate(),108) as sHora"
    Set rsH = dbCmact.CargaRecordSet(sSQL)
    If Not rsH.EOF Then
        GetHoraServer = rsH!sHora
    Else
        GetHoraServer = Format(Time, "hh:mm:ss")
    End If
    rsH.Close

    dbCmact.CierraConexion
    Set dbCmact = Nothing
End Function
Public Function GetMensajeBoletas(ByVal sCuenta As String, Optional sAgencia As String) As String
Dim sSQL As String
Dim rsMsg As New ADODB.Recordset
Set dbCmact = New COMConecta.DCOMConecta

If dbCmact.AbreConexion = False Then Exit Function
rsMsg.CursorLocation = adUseClient
sSQL = "Select cMensaje from Mensajes where cCodPro = '" & Mid(sCuenta, 6, 3) & "' And bMenEst = 1"
If sAgencia <> "" Then
    sSQL = sSQL & " And cCodAge = '" & sAgencia & "'"
End If

'rsMsg.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMsg = dbCmact.CargaRecordSet(sSQL)

Set rsMsg.ActiveConnection = Nothing
If rsMsg.EOF And rsMsg.BOF Then
    GetMensajeBoletas = ""
Else
    GetMensajeBoletas = Trim(rsMsg("cMensaje"))
End If
rsMsg.Close
Set rsMsg = Nothing
End Function

Public Function DevuelveCodAutorizaCMAC(ByVal psPersCod As String, ByRef psPersDesc As String) As String
    Dim rsA As New ADODB.Recordset
    Set dbCmact = New COMConecta.DCOMConecta

    If dbCmact.AbreConexion = False Then Exit Function
    sSQL = "Exec PIT_stp_sel_ObtieneCodigoAutCMACS '" & psPersCod & "'"
    Set rsA = dbCmact.CargaRecordSet(sSQL)
    
    If Not rsA.EOF Then
        DevuelveCodAutorizaCMAC = rsA!cCodAutOIT
        psPersDesc = rsA!cCMACDesc
    Else
        DevuelveCodAutorizaCMAC = ""
        psPersDesc = ""
    End If
    rsA.Close

    dbCmact.CierraConexion
    Set dbCmact = Nothing
End Function



Public Function nRegistraTramaEnvio(psTramaMovNro As String, psOpeCod As String, psTramaEnvio As String, _
        psCodTX As String, psCuentaTX As String, psPANTX As String, psDNITX As String, psMonedaTX As String, _
        pnMontoTX As Currency, psCodInstDest As String)
Dim lCn As New COMConecta.DCOMConecta
Dim lrs As New ADODB.Recordset
Dim lsSql As String


    lsSql = "exec PIT_stp_ins_TramaEnvioModoCli '" & psTramaMovNro & "','" & psOpeCod & "','" & psTramaEnvio & "','" & _
                psCodTX & "','" & psCuentaTX & "','" & psPANTX & "','" & psDNITX & "','" & psMonedaTX & "'," & pnMontoTX & ",'" & psCodInstDest & "'"
    

    lCn.AbreConexion

    Set lrs = lCn.ConexionActiva.Execute(lsSql)
    
    nRegistraTramaEnvio = lrs("nTramaId")
    
    lCn.CierraConexion
    
    Set lCn = Nothing
    Set lrs = Nothing
End Function

Public Sub RegistraTramaRecepcion(pnTramaId As Long, psTramaRecep As String, pnCodRespSwitch As Integer, psCodErrorTX As String, psCodRespTX As String, pnDenegada As Integer)
Dim lCn As New COMConecta.DCOMConecta
Dim lsSql As String


    lsSql = "exec PIT_stp_upd_RegistraTramaRecepcionModoCli " & pnTramaId & ",'" & psTramaRecep & "'," & pnCodRespSwitch & ",'" & psCodErrorTX & "','" & psCodRespTX & "'," & pnDenegada

    lCn.AbreConexion

    lCn.ConexionActiva.Execute (lsSql)
        
    lCn.CierraConexion
    
    Set lCn = Nothing
End Sub


Public Function nRegistraOperacionInterCMACM_Envio(psMovNro As String, psOpeCod As String, psCuenta As String, _
    psDescOperacion As String, psCMACPersCod As String, pnMoneda As Integer, psDocumento As String, _
    pnMonto As Currency, Optional psCliente As String = "", Optional psNomCmac As String = "", _
    Optional psNomAge As String = "", Optional pnComision As Currency = 0, Optional psGlosa As String = "")
Dim sMsgOpe As String
Dim bTrans As Boolean
Dim lnMovNro As Long
Dim nTempo As Currency
Dim nBan As Boolean
Dim lsCtaIF As String
Dim lsSql As String
Dim lCn As New COMConecta.DCOMConecta
Dim lrs As New ADODB.Recordset
    
    If psGlosa = "" Then
        psGlosa = psDescOperacion & ", " & psNomCmac & ", Cuenta = " & psCuenta
    Else
        psGlosa = sGlosa & Chr$(13) & psDescOperacion & ", " & psNomCmac & ", Cuenta = " & psCuenta
    End If
    
    
    lsSql = "exec PIT_stp_ins_OperacionInterCMACM_Envio '" & psMovNro & "','" & psOpeCod & "','" & _
            psGlosa & "','" & psCMACPersCod & "'," & pnMoneda & ",'" & psCuenta & "','" & _
            psDocumento & "'," & pnMonto & ",'" & psCliente & "'," & pnComision
            
    bTrans = False
    
    'Inicia la transaccion
    Set lCn = New COMConecta.DCOMConecta
On Error GoTo ErrGraba
    lCn.AbreConexion
    lCn.BeginTrans
    bTrans = True
    
    Set lrs = lCn.ConexionActiva.Execute(lsSql)
    
    If Not lrs.EOF And Not lrs.BOF Then
        lnMovNro = lrs("nMovNro")
    End If
    
    'Fin transaccion
    lCn.CommitTrans
    bTrans = False
    Set lCn = Nothing
    
    nRegistraOperacionInterCMACM_Envio = lnMovNro
    
    Exit Function
ErrGraba:
    If bTrans Then lCn.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Function



Public Function nRegistraExtornoOperacionInterCMACM_Envio(pnMovNroAExtonar As Long, psMovNro As String, psOpeCod As String, psCuenta As String, _
    psDescOperacion As String, psCMACPersCod As String, pnMoneda As Integer, psDocumento As String, _
    pnMonto As Currency, Optional psCliente As String = "", Optional psNomCmac As String = "", _
    Optional psNomAge As String = "", Optional pnComision As Currency = 0, Optional psGlosa As String = "") As Long
Dim sMsgOpe As String
Dim bTrans As Boolean
Dim lnMovNro As Long
Dim nTempo As Currency
Dim nBan As Boolean
Dim lsCtaIF As String
Dim lsSql As String
Dim lCn As New COMConecta.DCOMConecta
Dim lrs As New ADODB.Recordset
    
    If psGlosa = "" Then
        psGlosa = psDescOperacion & ", " & psNomCmac & ", Cuenta = " & psCuenta
    Else
        psGlosa = sGlosa & Chr$(13) & psDescOperacion & ", " & psNomCmac & ", Cuenta = " & psCuenta
    End If
    
    
    lsSql = "exec PIT_stp_ins_ExtornoOperacionInterCMACM_Envio '" & psMovNro & "','" & psOpeCod & "','" & _
            psGlosa & "','" & psCMACPersCod & "'," & pnMoneda & ",'" & psCuenta & "','" & _
            psDocumento & "'," & pnMonto & ",'" & psCliente & "'," & pnComision & "," & pnMovNroAExtonar
            

    bTrans = False
    
    'Inicia la transaccion
    Set lCn = New COMConecta.DCOMConecta
On Error GoTo ErrGraba
    lCn.AbreConexion
    lCn.BeginTrans
    bTrans = True
    
    Set lrs = lCn.ConexionActiva.Execute(lsSql)
    
    If Not lrs.EOF And Not lrs.BOF Then
        lnMovNro = lrs("nMovNro")
    End If
    
    'Fin transaccion
    lCn.CommitTrans
    bTrans = False
    Set lCn = Nothing
    
    nRegistraExtornoOperacionInterCMACM_Envio = lnMovNro
    
    Exit Function
ErrGraba:
    If bTrans Then lCn.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Function

Public Function ImprimeBoleta(ByVal psTit As String, ByVal psText As String, ByVal psMovNro As String, ByVal psMovNroCom As String, ByVal psCtaCod As String, _
                              ByVal pdFecSis As Date, ByVal sNomAge As String, ByVal psNomCMACDest As String, ByVal psNumTarj As String, ByVal psCodAutDest As String, _
                              ByVal psUser As String, ByVal psOpeCod As String, ByVal psDNI As String, ByVal pnMoneda As Integer, Optional pbImpTMU As Boolean = False) As String

Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMax As Integer
Dim sAux As Integer
Dim sCad As String
Dim lsNomAge As String
Dim lsMoneda As String
Dim lsCad As String
Dim lsSalLin As String
Dim oImpre As COMFunciones.FCOMImpresion
    
    On Error GoTo ERROR
    
    Set oImp = New NImpresion
    lsSalLin = Chr$(10)

    lnTope = 0
    sSep = 15
    sIni = 0
    sMax = 33
    sAux = 5
    

    lsNroExt = Str(pnNumExt)
    lsMoneda = IIf(pnMoneda = 1, "SOLES", "DOLARES")
    sFecha = Format$(pdFecSis, "dd/mm/yyyy")
    sHora = Format$(Time, "hh:mm:ss")
    lsNomAge = psNomAge
    
    
    lsCad = lsCad & Chr$(27) & Chr$(64)
    lsCad = lsCad & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
    lsCad = lsCad & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
    lsCad = lsCad & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
    lsCad = lsCad & Chr$(27) & Chr$(107) & Chr$(0)     'Tipo de Letra Sans Serif
    lsCad = lsCad & Chr$(27) & Chr$(18) ' cancela condensada
    lsCad = lsCad & Chr$(27) & Chr$(72) ' desactiva negrita


    Set oImpre = New COMFunciones.FCOMImpresion

    
    If Not pbImpTMU Then
        'Cabecera Operaciones InterCajas
        lsCad = lsCad & Chr$(27) & Chr$(71)  'Activa Negrita
        lsCad = lsCad & lsSalLin
        lsCad = lsCad & Space(sIni) & "CMAC MAYNAS" & Space(16 + sSep + sAux) & "CMAC MAYNAS" & lsSalLin 'oImp.gPrnSaltoLinea
        lsCad = lsCad & lsNomAge & Space(5) & psMovNro & Space(sAux + sMax + sSep - Len(lsNomAge) - Len(psMovNro)) & lsNomAge & Space(5) & psMovNro & lsSalLin & lsSalLin
        lsCad = lsCad & sFecha & Space(2) & sHora & Space(sAux + sMax + sSep - (Len(sFecha) + 2) - Len(sHora)) & sFecha & Space(2) & sHora & lsSalLin & lsSalLin
        lsCad = lsCad & Space(15) & psTit & Space(sAux + sMax + sSep - Len(psTit)) & Space(15) & psTit & lsSalLin & lsSalLin
        lsCad = lsCad & psNomCMACDest & Space(10) & psCodAutDest & Space(sAux + sMax + sSep - Len(psNomCMACDest) - Len(psCodAutDest)) & psNomCMACDest & Space(10) & psCodAutDest & lsSalLin
        
        Select Case psOpeCod
            Case "260503"  'Retiro
                lsCad = lsCad & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & Space(5) & lsMoneda & Space(16 + sSep + sAux) & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & Space(5) & lsMoneda & lsSalLin
                lsCad = lsCad & "Nº CTA :" & psCtaCod & Space(16 + sSep + sAux) & "Nº CTA :" & psCtaCod & lsSalLin
                lsCad = lsCad & psText & Space(16 + sSep + sAux) & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "260501" 'Deposito
                lsCad = lsCad & Right(psNumTarj, 6) & Space(5) & lsMoneda & Space(16 + sSep + sAux) & Right(psNumTarj, 6) & Space(5) & lsMoneda & lsSalLin
                lsCad = lsCad & "Nº CTA :" & psCtaCod & Space(16 + sSep + sAux) & "Nº CTA :" & psCtaCod & lsSalLin
                lsCad = lsCad & psText & Space(16 + sSep + sAux) & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "107001" 'Pago de Creditos
                lsCad = lsCad & "Nº CTA :" & psCtaCod & Space(16 + sSep + sAux) & "Nº CTA :" & psCtaCod & lsSalLin
                lsCad = lsCad & psText & Space(16 + sSep + sAux) & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "260505" 'Consulta de Cuentas de Ahorro
                lsCad = lsCad & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & Space(16 + sSep + sAux) & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & lsSalLin & lsSalLin
                lsCad = lsCad & psText & Space(16 + sSep + sAux) & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "260506" 'Consulta de Movimientos de Cuentas de Ahorro
                lsCad = lsCad & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & Space(5) & lsMoneda & Space(16 + sSep + sAux) & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & Space(5) & lsMoneda & lsSalLin
                lsCad = lsCad & "Nº CTA :" & psCtaCod & Space(16 + sSep + sAux) & "Nº CTA :" & psCtaCod & lsSalLin
                lsCad = lsCad & psText & Space(16 + sSep + sAux) & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "107004" 'Consulta de Cuentas de Credito
                lsCad = lsCad & psDNI & " / " & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & Space(16 + sSep + sAux) & psDNI & " / " & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & lsSalLin
                lsCad = lsCad & psText & Space(16 + sSep + sAux) & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
        End Select
        
        'Pie de Boleta
        lsCad = lsCad & "COMISION TRANSACCION" & Space(5) & "S/. 1.00" & Space(5) & psnMovNroCom & Space(16 + sSep + sAux) & "COMISION TRANSACCION" & Space(5) & "S/. 1.00" & Space(5) & psnMovNroCom & lsSalLin & lsSalLin & lsSalLin
        lsCad = lsCad & psUser & " / " & sFecha & " : " & sHora & Space(16 + sSep + sAux) & psUser & " / " & sFecha & " : " & sHora & lsSalLin
        lsCad = lsCad & "CARACTERISTICAS CMAC ORIGEN" & Space(16 + sSep + sAux) & "CARACTERISTICAS CMAC ORIGEN"
            
        ImprimeBoleta = lsCad & lsSalLin
        
    Else 'Imprmir en TMU
        lsCad = lsCad & Chr$(27) & Chr$(71)  'Activa Negrita
        lsCad = lsCad & lsSalLin
        lsCad = lsCad & Space(sIni) & "CMAC-MAYNAS" & lsSalLin 'oImp.gPrnSaltoLinea
        lsCad = lsCad & lsNomAge & Space(10) & psMovNro & lsSalLin & lsSalLin
        lsCad = lsCad & sFecha & Space(2) & sHora & lsSalLin & lsSalLin
        lsCad = lsCad & Space(15) & psTit & lsSalLin & lsSalLin
        lsCad = lsCad & psNomCMACDest & Space(10) & psCodAutDest & lsSalLin
        
        Select Case psOpeCod
            Case "260503"  'Retiro
                lsCad = lsCad & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & Space(5) & lsMoneda & lsSalLin
                lsCad = "Nº CTA :" & psCtaCod & lsSalLin
                lsCad = lsCad & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "260501" 'Deposito
                lsCad = lsCad & Right(psNumTarj, 6) & Space(5) & lsMoneda & lsSalLin
                lsCad = "Nº CTA :" & psCtaCod & lsSalLin
                lsCad = lsCad & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "107001" 'Pago de Creditos
                lsCad = "Nº CTA :" & psCtaCod & lsSalLin
                lsCad = lsCad & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "260505" 'Consulta de Cuentas de Ahorro
                lsCad = lsCad & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & lsSalLin & lsSalLin
                lsCad = lsCad & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "260506" 'Consulta de Movimientos de Cuentas de Ahorro
                lsCad = lsCad & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & Space(5) & lsMoneda & lsSalLin
                lsCad = "Nº CTA :" & psCtaCod & lsSalLin
                lsCad = lsCad & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "107004" 'Consulta de Cuentas de Credito
                lsCad = lsCad & psDNI & " / " & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & lsSalLin
                lsCad = lsCad & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
        End Select
                
        lsCad = lsCad & "COMISION TRANSACCION" & Space(5) & "S/. 1.00" & Space(5) & psNroRecIngXCom & lsSalLin & lsSalLin & lsSalLin
        lsCad = lsCad & psUser & " / " & sFecha & " : " & sHora & lsSalLin
        lsCad = lsCad & "CARACTERISTICAS CMAC ORIGEN"

        lsCad = lsCad & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & Chr(27) & "m" & Chr(10)
        ImprimeBoleta = lsCad
    End If
Exit Function
ERROR:
    Err.Raise Err.Number, "Comprueba la conexion de su impresora... ", Err.Description
'    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
'        GoTo ETIQ
'    End If
End Function


Public Function ImprimeBoletaOpeInterCMAC(ByVal psTit As String, ByVal psText As String, ByVal psMovNro As String, ByVal psMovNroCom As String, ByVal psCtaCod As String, _
                              ByVal pdFecSis As Date, ByVal psNomAge As String, ByVal psNomCMACDest As String, ByVal psNumTarj As String, ByVal psCodAutDest As String, _
                              ByVal psUser As String, ByVal psOpeCod As String, ByVal psDNI As String, ByVal pnMoneda As Integer, pnMonto As Currency, Optional pbImpTMU As Boolean = False) As String
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMonto As String
Dim lsNroExt As String
Dim lsNomAge As String
Dim lsMensaje As String * 39
Dim lsCad As String
Dim lnMaxCar As String
Dim lsTarjeta As String
Dim lsTexto As String
Dim lsSaldoC As String
Dim lsSaldoD As String
Dim lnLenTexto As Integer, lnPosi As Integer, lnLineas As Integer, i As Integer, lnPosiT As Integer
Dim oImpre As COMFunciones.FCOMImpresion
Dim clsGen As COMDConstSistema.DCOMGeneral



ETIQ:
    
On Error GoTo ERROR
    
    sSep = 14
    sIni = 0
    lnMaxCar = 39
    lnLineas = 0
       
    lsNroExt = Str(pnNumExt)
    sFecha = Format$(pdFecSis, "DD/MM/YYYY")
    sHora = Format$(Time, "HH:MM:SS")
    sMonto = Format$(pnMonto, "#,##0.00")

    lsNomAge = psNomAge
    lsTarjeta = Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4)
    lsMoneda = Left(IIf(pnMoneda = 1, "SOLES", "DOLARES") & Space(7), 7)
    
    If psOpeCod = "261003" Or psOpeCod = "261004" Or psOpeCod = "104001" Or psOpeCod = "104002" Then
        lnPosiT = 0
        lnPosiT = InStr(1, psText, "1P")
        If lnPosiT = 1 Then
            psText = "000" & psText
        End If
    End If
    
    lsTexto = psText
    lnLenTexto = Len(psText)
    
    lsCad = lsCad & oImp.gPrnInicializa
    lsCad = lsCad & oImp.gPrnEspaLineaN   'espaciamiento lineas 1/6 pulg.
    lsCad = lsCad & oImp.gPrnTamPagina22  'Longitud de página a 22 líneas'
    lsCad = lsCad & oImp.gPrnTamLetra10CPI   'Tamaño 10 cpi
    lsCad = lsCad & oImp.gPrnTpoLetraRoman     'Tipo de Letra Sans Serif
    lsCad = lsCad & oImp.gPrnCondensadaOFF ' cancela condensada
    lsCad = lsCad & oImp.gPrnNegritaOFF ' desactiva negrita
        
    Set oImpre = New COMFunciones.FCOMImpresion
    Set clsGen = New COMDConstSistema.DCOMGeneral
    
    If Not pbImpTMU Then
        'Cabecera
        lsCad = lsCad & oImp.gPrnNegritaON 'Activa Negrita
        
        lsCad = lsCad & Space(sIni) & "CMAC-MAYNAS         " & Space(19) & Space(sSep) & "CMAC-MAYNAS - AHORRO" & oImp.gPrnSaltoLinea
        lsCad = lsCad & Space(sIni) & Trim(lsNomAge) & Space(lnMaxCar - Len(Trim(lsNomAge))) & Space(sSep) & Trim(lsNomAge) & Space(lnMaxCar - Len(Trim(lsNomAge))) & oImp.gPrnSaltoLinea
        lsCad = lsCad & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(14) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & oImp.gPrnSaltoLinea

        psTit = Trim(psTit)
        psTit = oImpre.CentrarCadena(psTit, 34)
        lsCad = lsCad & Space(sIni) & "---" & psTit & "--" & Space(sSep) & "------" & psTit & "-----" & oImp.gPrnSaltoLinea
        lsCad = lsCad & Space(sIni) & Left(psNomCMACDest & Space(30), 30) & Space(2) & Right(Space(7) & psCodAutDest, 7) & Space(sSep) & Left(psNomCMACDest & Space(30), 30) & Space(2) & Right(Space(7) & psCodAutDest, 7) & oImp.gPrnSaltoLinea
        lsCad = lsCad & oImp.gPrnNegritaOFF

                
        Select Case psOpeCod
            Case "261001"  'Retiro
                lsCad = lsCad & Space(sIni) & "NRO TAR: " & lsTarjeta & Space(16) & Space(sSep) & "NRO TAR: " & lsTarjeta & oImp.gPrnSaltoLinea
                lsCad = lsCad & Space(sIni) & "NRO CTA: " & Left(psCtaCod & Space(21), 21) & Space(2) & lsMoneda & Space(sSep) & "NRO CTA: " & Left(psCtaCod & Space(21), 21) & Space(2) & lsMoneda & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                lsCad = lsCad & Space(sIni) & Space(5) & "MONTO     ******: " & Left(sMonto & Space(16), 16) & Space(sSep) & Space(5) & "MONTO     ******: " & Left(sMonto & Space(16), 16) & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                lsSaldoC = Format$(Mid(lsTexto, 3, 9) & "." & Mid(lsTexto, 12, 2), "#,##0.00")
                lsSaldoD = Format$(Mid(lsTexto, 16, 8) & "." & Mid(lsTexto, 24, 2), "#,##0.00")
                
                lsCad = lsCad & Space(sIni) & Space(5) & "SALDO DISPONIBLE: " & Left(lsSaldoD & Space(16), 16) & Space(sSep) & Space(5) & "SALDO DISPONIBLE: " & Left(lsSaldoD & Space(16), 16) & oImp.gPrnSaltoLinea
                lsCad = lsCad & Space(sIni) & Space(5) & "SALDO CONTABLE**: " & Left(lsSaldoC & Space(16), 16) & Space(sSep) & Space(5) & "SALDO CONTABLE**: " & Left(lsSaldoC & Space(16), 16) & oImp.gPrnSaltoLinea
            Case "261002" 'Deposito
                lsCad = lsCad & Space(sIni) & "NRO TAR: " & lsTarjeta & Space(16) & Space(sSep) & "NRO TAR: " & lsTarjeta & oImp.gPrnSaltoLinea
                lsCad = lsCad & Space(sIni) & "NRO CTA: " & Left(psCtaCod & Space(21), 21) & Space(2) & lsMoneda & Space(sSep) & "NRO CTA: " & Left(psCtaCod & Space(21), 21) & Space(2) & lsMoneda & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                lsCad = lsCad & Space(sIni) & Space(5) & "MONTO     ******: " & Left(sMonto & Space(16), 16) & Space(sSep) & Space(5) & "MONTO     ******: " & Left(sMonto & Space(16), 16) & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
            Case "261003" 'Consulta de Cuentas de Ahorro
                lsCad = lsCad & Space(sIni) & "NRO TAR: " & lsTarjeta & Space(16) & Space(sSep) & "NRO TAR: " & lsTarjeta & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                If (lnLenTexto > 15) Then
                    lnPosi = 16
                    While lnPosi < lnLenTexto
                        lsCad = lsCad & Space(sIni) & Left(Mid(lsTexto, lnPosi, 40) & Space(40), 40) & Space(sSep - 1) & Left(Mid(lsTexto, lnPosi, 40) & Space(40), 40) & oImp.gPrnSaltoLinea
                        lnPosi = lnPosi + 40
                        lnLineas = lnLineas + 1
                    Wend
                    
                    For i = 1 To (8 - lnLineas)
                        lsCad = lsCad & oImp.gPrnSaltoLinea
                    Next
                End If

            Case "261004" 'Consulta de Movimientos de Cuentas de Ahorro
                lsCad = lsCad & Space(sIni) & "NRO TAR: " & lsTarjeta & Space(16) & Space(sSep) & "NRO TAR: " & lsTarjeta & oImp.gPrnSaltoLinea
                lsCad = lsCad & Space(sIni) & "NRO CTA: " & Left(psCtaCod & Space(21), 21) & Space(2) & lsMoneda & Space(sSep) & "NRO CTA: " & Left(psCtaCod & Space(21), 21) & Space(2) & lsMoneda & oImp.gPrnSaltoLinea
                
                If (lnLenTexto > 15) Then
                    lnPosi = 16
                    While lnPosi < lnLenTexto
                        lsCad = lsCad & Space(sIni) & Left(Mid(lsTexto, lnPosi, 40) & Space(40), 40) & Space(sSep - 1) & Left(Mid(lsTexto, lnPosi, 40) & Space(40), 40) & oImp.gPrnSaltoLinea
                        lnPosi = lnPosi + 40
                        lnLineas = lnLineas + 1
                    Wend
                    
                    For i = 1 To (8 - lnLineas)
                        lsCad = lsCad & oImp.gPrnSaltoLinea
                    Next
                End If
                
            Case "104001" 'Pago de Creditos
                lsCad = lsCad & Space(sIni) & "NRO CTA: " & Left(psCtaCod & Space(21), 21) & Space(2) & lsMoneda & Space(sSep) & "NRO CTA: " & Left(psCtaCod & Space(21), 21) & Space(2) & lsMoneda & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                
                If (lnLenTexto > 15) Then
                    lnPosi = 16
                    While lnPosi < lnLenTexto
                        lsCad = lsCad & Space(sIni) & Left(Mid(lsTexto, lnPosi, 40) & Space(40), 40) & Space(sSep - 1) & Left(Mid(lsTexto, lnPosi, 40) & Space(40), 40) & oImp.gPrnSaltoLinea
                        lnPosi = lnPosi + 40
                        lnLineas = lnLineas + 1
                    Wend
                    
                    For i = 1 To (8 - lnLineas)
                        lsCad = lsCad & oImp.gPrnSaltoLinea
                    Next
                    
                End If
                
            Case "104002" 'Consulta de Cuentas de Credito
                lsCad = lsCad & Space(sIni) & "NRO DNI: " & psDNI & Space(22) & Space(sSep) & "NRO TAR: " & lsTarjeta & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                
                If (lnLenTexto > 15) Then
                    lnPosi = 16
                    While lnPosi < lnLenTexto
                        lsCad = lsCad & Space(sIni) & Left(Mid(lsTexto, lnPosi, 40) & Space(40), 40) & Space(sSep - 1) & Left(Mid(lsTexto, lnPosi, 40) & Space(40), 40) & oImp.gPrnSaltoLinea
                        lnPosi = lnPosi + 40
                        lnLineas = lnLineas + 1
                    Wend
                    
                    For i = 1 To (8 - lnLineas)
                        lsCad = lsCad & oImp.gPrnSaltoLinea
                    Next
                    
                End If
            Case "159101", "279101", "279102" 'Extornos
                lsCad = lsCad & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                lsCad = lsCad & Space(sIni) & oImpre.CentrarCadena("EXTORNO DE OPERACIÓN REALIZADA", 40) & Space(sSep - 1) & oImpre.CentrarCadena("EXTORNO DE OPERACIÓN REALIZADA", 40) & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
                lsCad = lsCad & oImp.gPrnSaltoLinea
        End Select
        'PIE DE PÁGINA
        lsCad = lsCad & Space(sIni) & "---------------------------------------" & Space(sSep) & "---------------------------------------" & oImp.gPrnSaltoLinea
        lsCad = lsCad & Space(sIni) & oImpre.ImpreCarEsp(psUser) & Space(35) & Space(sSep) & oImpre.ImpreCarEsp(psUser) & oImp.gPrnSaltoLinea
        lsCad = lsCad & Space(sIni) & "---------------------------------------" & Space(sSep) & "---------------------------------------" & oImp.gPrnSaltoLinea
        If psOpeCod = "261001" Or psOpeCod = "261003" Or psOpeCod = "261004" Then 'Para retiros
            lsCad = lsCad & Space(sIni) & Left("Esta operación fue autorizada" & Space(39), 39) & Space(sSep) & Left("Esta operación fue autorizada" & Space(39), 39) & oImp.gPrnSaltoLinea
            lsCad = lsCad & Space(sIni) & Left("con el uso de mi clave secreta" & Space(39), 39) & Space(sSep) & Left("con el uso de mi clave secreta" & Space(39), 39) & oImp.gPrnSaltoLinea
        Else
            lsCad = lsCad & Space(sIni) & Left("CARACTERISTICAS CMAC ORIGEN" & Space(39), 39) & Space(sSep) & Left("CARACTERISTICAS CMAC ORIGEN" & Space(39), 39) & oImp.gPrnSaltoLinea
            lsCad = lsCad & Space(sIni) & Left("" & Space(39), 39) & Space(sSep) & Left("" & Space(39), 39) & oImp.gPrnSaltoLinea
        End If
        
        Set clsGen = Nothing
        
    Else
'        lsCad = lsCad & oImp.gPrnNegritaON 'Activa Negrita
'        lsCad = lsCad & Space(sIni) & "CMAC-MAYNAS - AHORRO" & oImp.gPrnSaltoLinea
'        If nMoneda = gMonedaNacional Then
'            lsCad = lsCad & Space(sIni) & Trim(sNomAge) & "-SOLES " & lsNroExt & oImp.gPrnSaltoLinea
'        Else
'            lsCad = lsCad & Space(sIni) & Trim(sNomAge) & "-DOLARES " & lsNroExt & Space(sSep) & Trim(sNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & oImp.gPrnSaltoLinea
'        End If
'
'        lsCad = lsCad & ""
'
'        If psComCmac = "XXX" Then
'            If psLin3 = "XXX" Then
'                lsCad = lsCad & oImp.gPrnNegritaOFF 'Desactiva Negrita
'            Else
'                lsCad = lsCad & Space(sIni) & psLin3 & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea   'Desactiva Negrita
'                lnNumLinCmac = 1
'            End If
'            lnNumLinCmac = 0
'        Else
'            lsCad = lsCad & Space(sIni) & "NroDocCmac:" & psComCmac & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea 'Desactiva Negrita
'            lnNumLinCmac = 1
'        End If
'
'        lsCad = lsCad & Space(sIni) & "Fecha:" & sFecha & Space(5) & "Hora:" & sHora & oImp.gPrnSaltoLinea
'
'        lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
'
'        If lnCliAux = 0 Then
'            If sAux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + sAux)
'            lsCad = lsCad & Space(sIni) & oImpre.ImpreCarEsp(psCliente) & oImp.gPrnSaltoLinea
'            lsCad = lsCad & Space(sIni) & "Cuenta:" & psCodCta & oImp.gPrnSaltoLinea
'        Else
'            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
'            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
'
'            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + sAux)
'            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + sAux)
'
'            lsCad = lsCad & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux1) & oImp.gPrnSaltoLinea
'            lsCad = lsCad & Space(sIni) & oImpre.ImpreCarEsp(lsCliAux2) & oImp.gPrnSaltoLinea
'            lnCliAux = 1
'            lsCad = lsCad & Space(sIni) & "Cuenta:" & psCodCta & oImp.gPrnSaltoLinea
'        End If
'
'        psTit = Trim(psTit)
'        psTit = oImpre.CentrarCadena(psTit, 28)
'        lsCad = lsCad & oImp.gPrnNegritaON 'Activa Negrita
'        lsCad = lsCad & Space(sIni + 1) & "-----" & psTit & "-----" & oImp.gPrnSaltoLinea
'
'
'        lnChq = InStr(1, psText, "*", vbTextCompare)
'
'        If psTexto = "XXX" Then
'            If lnChq = 0 Then
'                lsCad = lsCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(Mid(psText, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)) & sMonto & oImp.gPrnSaltoLinea
'                lsCad = lsCad & ""
'            Else
'                lsChqAux1 = (Mid(psText, 1, lnChq - 1))
'                lsChqAux2 = (Mid(psText, lnChq + 1))
'                lsCad = lsCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & oImp.gPrnSaltoLinea
'                lsCad = lsCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(lsChqAux2)) & oImp.gPrnSaltoLinea
'            End If
'        Else
'            lsCad = lsCad & Space(sIni) & oImpre.ImpreCarEsp(Trim(psTexto)) & oImp.gPrnSaltoLinea
'            lsCad = lsCad & ""
'        End If
'
'        lsCad = lsCad & oImp.gPrnNegritaOFF 'Desactiva Negrita
'        If pbSaldoInt Then
'            If bPregImpSaldos Then
'                If pbImprimeSaldos Then
'                     lsCad = lsCad & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & oImp.gPrnSaltoLinea
'                    If pbOpSaldoC Then
'                       lsCad = lsCad & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & oImp.gPrnSaltoLinea
'                    Else
'                       lsCad = lsCad & ""
'                    End If
'                Else
'                    lsCad = lsCad & ""
'                    lsCad = lsCad & ""
'                    pbSaldoInt = False
'                End If
'            Else
'                lsCad = lsCad & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & oImp.gPrnSaltoLinea
'                If pbOpSaldoC Then
'                    lsCad = lsCad & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & oImp.gPrnSaltoLinea
'                Else
'                    lsCad = lsCad & ""
'                End If
'            End If
'        Else
'            lsCad = lsCad & ""
'            lsCad = lsCad & ""
'            lsCad = lsCad & ""
'        End If
'        lsCad = lsCad & ""
'        If Not psNumDoc = "" Then
'            lsCad = lsCad & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & oImp.gPrnSaltoLinea
'        Else
'            lsCad = lsCad & ""
'        End If
'
'        lnTope = 3 - lnCliAux
'
'        lsCad = lsCad & Space(sIni) & "---------------------------------------" & oImp.gPrnSaltoLinea
'        lsCad = lsCad & Space(sIni) & oImpre.ImpreCarEsp(sCodUser) & oImp.gPrnSaltoLinea
'
'        'lsMensaje = clsGen.GetMensajeBoletas("XXXXX" & gCapAhorros, sCodAge)
'        Set clsGen = Nothing
'        lsCad = lsCad & Space(sIni) & oImp.gPrnNegritaON & oImpre.ImpreCarEsp(lsMensaje) & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
'        lnNumLinCmac = lnNumLinCmac + 1
'        lsCad = lsCad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & Chr(27) & "m" & Chr(10)
    End If
    ImprimeBoletaOpeInterCMAC = lsCad
Exit Function
ERROR:
    Err.Raise Err.Number, "Comprueba la conexion de su impresora...", Err.Description
    
End Function
Public Sub IniciaImpresora(Optional ByVal nImpresora As COMFunciones.Impresoras = gEPSON)
    Set oImp = New COMFunciones.FCOMVarImpresion
        oImp.Inicia nImpresora
End Sub
Private Sub Class_Initialize()
    IniciaImpresora
End Sub

Private Sub Class_Terminate()
    Set oImp = Nothing
    Set dbCmact = Nothing
End Sub

Public Function obtenerMovimientosInterCajasParaExtorno(pnTipoBusq As Integer, psDatoBusq As String, psOpeCod As String, psFecha As String) As ADODB.Recordset
Dim lCn As COMConecta.DCOMConecta
Dim lsSql As String
    
    lsSql = " exec PIT_stp_sel_MovimientosInterCajasParaextorno " & pnTipoBusq & ",'" & psDatoBusq & "','" & psOpeCod & "','" & psFecha & "'"
    
    Set lCn = New COMConecta.DCOMConecta
    lCn.AbreConexion
    Set obtenerMovimientosInterCajasParaExtorno = lCn.CargaRecordSet(lsSql)
    lCn.CierraConexion
    Set lCn = Nothing
End Function


Public Function obtenerMovimientoInterCajaParaExtorno(pnMovNro) As ADODB.Recordset
Dim lCn As COMConecta.DCOMConecta
Dim lsSql As String
    
    lsSql = " exec PIT_stp_sel_MovimientoInterCajaParaExtorno " & pnMovNro
    
    Set lCn = New COMConecta.DCOMConecta
    lCn.AbreConexion
    Set obtenerMovimientoInterCajaParaExtorno = lCn.CargaRecordSet(lsSql)
    lCn.CierraConexion
    Set lCn = Nothing
End Function


Public Function obtenerParametros() As ADODB.Recordset
Dim lCn As COMConecta.DCOMConecta
Dim lsSql As String
    
    lsSql = " exec PIT_stp_sel_Parametros "
    
    Set lCn = New COMConecta.DCOMConecta
    lCn.AbreConexion
    Set obtenerParametros = lCn.CargaRecordSet(lsSql)
    lCn.CierraConexion
    Set lCn = Nothing
End Function


Public Function obtenerCodigosISO() As ADODB.Recordset
Dim lCn As COMConecta.DCOMConecta
Dim lsSql As String
    
    lsSql = " exec PIT_stp_sel_CodigosISO "
    
    Set lCn = New COMConecta.DCOMConecta
    lCn.AbreConexion
    Set obtenerCodigosISO = lCn.CargaRecordSet(lsSql)
    lCn.CierraConexion
    Set lCn = Nothing
End Function

Public Function obtenerDescripcionPorCodigoISO(psCodigoISO As String) As String
Dim lrs As ADODB.Recordset
Dim lCn As COMConecta.DCOMConecta
Dim lsSql As String
    
    lsSql = " exec PIT_stp_sel_CodigoISO '" & psCodigoISO & "'"
    
    Set lCn = New COMConecta.DCOMConecta
    lCn.AbreConexion
    Set lrs = lCn.CargaRecordSet(lsSql)
    lCn.CierraConexion
    
    If Not (lrs.EOF And lrs.BOF) Then
        obtenerDescripcionPorCodigoISO = lrs("cDescripcion")
    Else
        obtenerDescripcionPorCodigoISO = "Error desconocido"
    End If
    
    Set lCn = Nothing
    Set lrs = Nothing
    
End Function

Public Function ObtenerMontoOpeXDia(ByVal psCuenta As String, ByVal pnMoneda As Integer, ByVal pdFecha As String, _
                                    ByRef pnNumOpe As Integer) As Double
                                    
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsSql As String
    
    lsSql = " exec PIT_stp_sel_MontoOpeXdia '" & psCuenta & "', " & pnMoneda & ",'" & Format(pdFecha, "YYYY/MM/DD") & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(lsSql)
    oCon.CierraConexion
    
    If Not (rs.EOF And rs.BOF) Then
        ObtenerMontoOpeXDia = rs("nMonto")
        pnNumOpe = rs("nNumOpe")
    Else
        ObtenerMontoOpeXDia = 0
        pnNumOpe = 0
    End If
    
    Set oCon = Nothing
    Set lrs = Nothing
    
'Dim oConect As DConecta
'Dim Cmd As New Command
'Dim Prm As New ADODB.Parameter
'
'    Set oConect = New DConecta
'    oConect.AbreConexion
'
'    Set Cmd = New ADODB.Command
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@sCuenta", adVarChar, adParamInput, 18, psCuenta)
'    Cmd.Parameters.Append Prm
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@nMoneda", adInteger, adParamInput, , pnMoneda)
'    Cmd.Parameters.Append Prm
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@pdFecha", adDBDate, adParamInput, , Format(pdFecha, "YYYY/MM/DD"))
'    Cmd.Parameters.Append Prm
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@nMonto", adDouble, adParamOutput)
'    Cmd.Parameters.Append Prm
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@nNumOpe", adInteger, adParamOutput)
'    Cmd.Parameters.Append Prm
'
'    oConect.AbreConexion
'    Cmd.ActiveConnection = oConect.ConexionActiva ' AbrirConexion
'    Cmd.CommandType = adCmdStoredProc
'    Cmd.CommandText = "PIT_stp_sel_MontoOpeXdia"
'    Cmd.Execute
'
'    ObtenerMontoOpeXDia = Cmd.Parameters(3).Value
'    pnNumOpe = Cmd.Parameters(4).Value
'
'    Set Prm = Nothing
'    Set Cmd = Nothing
'
'    oConect.CierraConexion
'    Set oConect = Nothing
   
End Function

Public Function ObtenerMontoOpeXMes(ByVal psCuenta As String, ByVal pnMoneda As Integer, ByVal pdFecha As String, _
                                    ByRef pnNumOpe As Integer) As Double
                                    
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsSql As String
    
    lsSql = " exec PIT_stp_sel_MontoOpeXMes '" & psCuenta & "', " & pnMoneda & ",'" & Format(pdFecha, "YYYY/MM/DD") & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(lsSql)
    oCon.CierraConexion
    
    If Not (rs.EOF And rs.BOF) Then
        ObtenerMontoOpeXMes = rs("nMonto")
        pnNumOpe = rs("nNumOpe")
    Else
        ObtenerMontoOpeXMes = 0
        pnNumOpe = 0
    End If
    
    Set oCon = Nothing
    Set lrs = Nothing
    
'Dim oConect As DConecta
'Dim Cmd As New Command
'Dim Prm As New ADODB.Parameter
'
'    Set oConect = New DConecta
'    oConect.AbreConexion
'
'    Set Cmd = New ADODB.Command
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@sCuenta", adVarChar, adParamInput, 18, psCuenta)
'    Cmd.Parameters.Append Prm
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@nMoneda", adInteger, adParamInput, , pnMoneda)
'    Cmd.Parameters.Append Prm
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@pdFecha", adDBDate, adParamInput, , Format(pdFecha, "YYYY/MM/DD"))
'    Cmd.Parameters.Append Prm
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@nMonto", adDouble, adParamOutput)
'    Cmd.Parameters.Append Prm
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@nNumOpe", adInteger, adParamOutput)
'    Cmd.Parameters.Append Prm
'
'    oConect.AbreConexion
'    Cmd.ActiveConnection = oConect.ConexionActiva ' AbrirConexion
'    Cmd.CommandType = adCmdStoredProc
'    Cmd.CommandText = "PIT_stp_sel_MontoOpeXMes"
'    Cmd.Execute
'
'    ObtenerMontoOpeXMes = Cmd.Parameters(3).Value
'    pnNumOpe = Cmd.Parameters(4).Value
'
'    Set Prm = Nothing
'    Set Cmd = Nothing
'
'    oConect.CierraConexion
'    Set oConect = Nothing
   
End Function



