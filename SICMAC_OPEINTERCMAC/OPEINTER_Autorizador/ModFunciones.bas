Attribute VB_Name = "ModFunciones"
Option Explicit
Public C As ADODB.Connection
Global Const gsCodCMAC As String = "231"
Global Const gsCanal As String = "CMAC"

Global gnMontoMinRetMN As Double
Global gnMontoMaxRetMN As Double
Global gnMontoMinRetME As Double
Global gnMontoMaxRetME As Double
Global gnMontoMinRetMNReqDNI As Double
Global gnMontoMinRetMEReqDNI As Double
Global gnMontoMaxOpeMNxDia As Double
Global gnMontoMaxOpeMExDia As Double
Global gnMontoMaxOpeMNxMes As Double
Global gnMontoMaxOpeMExMes As Double
Global gnNumeroMaxOpeXDia As Double
Global gnNumeroMaxOpeXMes As Double

Global gnMontoMinDepMN As Double
Global gnMontoMaxDepMN As Double
Global gnMontoMinDepME As Double
Global gnMontoMaxDepME As Double

Public gsCadenaConexion As String
Dim loConec As New DConecta

Public Enum TipoCambio
    TCFijoMes = 0
    TCFijoDia = 1
    TCVenta = 2
    TCCompra = 3
    TCVentaEsp = 4
    TCCompraEsp = 5
    TCPonderado = 6
    TCPondVenta = 7
    TCPondREU = 8
End Enum

Enum ColocTipoPago
    gColocTipoPagoEfectivo = 1
    gColocTipoPagoCheque = 2
    gColocTipoPagoDacionPago = 3
    gColocTipoPagoCargoCta = 4
End Enum

Public Enum ColocEstado
    gColocEstSolic = 2000     'Colocaciones Solicitado
    gColocEstSug = 2001       'Colocaciones Sugerido
    gColocEstAprob = 2002     'Colocaciones Aprobado
    gColocEstRech = 2003      'Colocaciones Rechazado
    gColocEstDesemb = 2004    'ColoCaCiones Desembolso
    gColocEstVigNorm = 2020   'ColoCaCiones Vigente Normal
    gColocEstVigVenc = 2021   'ColoCaCiones Vigente VenCido
    gColocEstVigMor = 2022    'ColoCaCiones Vigente Moroso
    gColocEstRefNorm = 2030   'ColoCaCiones RefinanCiado Normal
    gColocEstRefVenc = 2031   'ColoCaCiones RefinanCiado VenCido
    gColocEstRefMor = 2032    'ColoCaCiones RefinanCiado Moroso
    gColocEstEmbarg = 2040    'ColoCaCiones Embargado
    gColocEstCancelado = 2050 'Colocaciones Cancelado
    gColocEstRefinanc = 2060  'Colocaciones Refinanciado
    gColocEstJudicial = 2070  'Colocaciones Paso a Judicial
    gColocEstRetirado = 2080  'Colocaciones Retirado
    gColocEstHonrada = 2090   'Colocaciones Honrada
    gColocEstDevuelta = 2091  'Colocaciones Devuelta
    gColocEstRenovada = 2092  'Colocaciones Renovada
End Enum

Public Enum CodigoRespuesta
    'Captaciones y Colocaciones
    gPITCuentaNoExiste = 30
    gPITCuentaNoPerteneceTitular = 31
    gPITCuentaNoVigente = 32
    gPITCuentaCancelada = 33

    'Captaciones
    gPITCuentaBloqueada = 40
    gPITFondosInsuficienteparaExt = 50
    gPITTarjetaNoValida = 60
    gPITTarjetaCaducada = 61
    gPITCuentaNoAdmiteRet = 62
    gPITFondosInsuficienteparaRet = 63
    gPITRetencionJudicial = 64
    
    'Colocaciones
    gPITPagoExcedeDeuda = 70
    gPITCuentaCanceldaDebeEntregar = 71
    gPITVuelto = 72
    gPITClientedeberecogerExcedente = 73
End Enum

Public Enum OperacionInterCMAC
    gPITColocPagoCredito = 105001
    gPITColocConsCtasCredito = 105002
    gPITColocConsMovCtasCred = 105003
    
    gPITCaptacRetiro = 261501
    gPITCaptaDeposito = 261502
    gPITCaptaConsCtasAhorro = 261503
    gPITCaptaConsMovCtasAho = 261504
    gPITCaptaConsSaldo = 261505
    
    gPITColocExtPagoCredito = 159201
    gPITCaptacExtRetiro = 279201
    gPITCaptacExtDeposito = 279202
    
End Enum


Public Enum CredOperacion
        
    gCredPagNorNorCC = 100202
    'Dacion en Pago
    gCredPagNorNorDacion = 100207
    
    'Pagos RFA
    gCredPagNorMorCC = 100302
    gCredPagNorMorDacion = 100307
    
    gCredPagNorVenCC = 100402
    'Dacion en Pago
    gCredPagNorVenDacion = 100407

    gCredPagRefNorEfec = 100501
    gCredPagRefNorCC = 100502
    gCredPagRefNorCh = 100503

    'Dacion en Pago
    gCredPagRefNorDacion = 100507
    
    gCredPagNorMorEfec = 100301
    gCredPagNorNorEfec = 100201
    gCredPagNorVenEfec = 100401
    
    gCredPagNorVenCh = 100403
    gCredPagNorNorCh = 100203
    gCredPagNorMorCh = 100303
    
    gCredPagRefMorEfec = 100601
    gCredPagRefMorCC = 100602
    gCredPagRefMorCh = 100603
    'Dacion en Pago
    gCredPagRefMorDacion = 100607

    gCredPagRefVenEfec = 100701
    gCredPagRefVenCC = 100702
    gCredPagRefVenCh = 100703
    'Dacion en Pago
    gCredPagRefVenDacion = 100707
End Enum

Public gnITFPorcent As Double
Public gbITFAplica As Boolean
Public gnITFMontoMin As Double


Private pMatCalend As Variant
Private pMatCalend_2 As Variant
Private pMatCalendTmp As Variant
Private pMatCalendDistribuido As Variant

Public sUserATM As String
Public sHora As String
Public sMesDia As String


Public gMESSAGE_TYPE As String, gTRACE As String, gPRCODE As String
Public gsPAN As String, gTIME_LOCAL As String, gDATE_LOCAL As String
Public gTERMINAL_ID As String, gACCT_1 As String, gCARD_ACCEPTOR As String
Public gACQ_INST As String, gPOS_COND_CODE As String, gTXN_AMOUNT As String
Public gCUR_CODE As String, gACCT_2 As String, gDATE_EXP As String
Public gCARD_LOCATION As String, gDATA_ATM_ADD As String

Public gsMonCta As String, gsCtaCod As String, gsDNI As String, gsMovNro As String
Public gnTramaIDExt As Long, gnTramaId As Long

'Dim gnMonMinMN As Double, gnMonMinME As Double, gnMonMaxMN As Double, gnMonMaxME As Double

Dim nTipoCambioVenta As Double, nTipoCambioCompra As Double, nTipoCambio As Double

Dim dFecSis As Date
Dim nOFFHost As Integer, nTarjCondicion As Integer, nRetenerTarjeta As Integer
Dim nNOOperMonExt As Integer, nSuspOper As Integer
Dim dFecVenc As Date
Dim sCtaCod As String
Dim nCtaSaldo As Double
Dim pINXml As String, pOUTXml As String
Dim nTipoEquipo As Integer, nTipoServicio As Integer
Dim nProced As Integer, nTipoOperac As Integer, nMoneda As Integer
Dim bOpcConsu As Boolean


Dim sOpeCod As String, sOpeCodComision As String, sOpeCodITF As String, sOpeCodComisionITF As String
Dim sIDTrama As String

Dim sOpeExtorno As String, sOpeExtornoComision As String, sOpeCodTransferencia As String, sOpeCodExtornoTransfer As String

Dim nMontoTran As Double, nMontoITF As Double, nMontoComision As Double, nMontoComisionITF As Double
Dim lnMovNro As Long
Dim sCadResp As String
Dim sCadAmount As String
Dim nMontoEquiv As Double
Dim nResultado As Long

Public Function RecuperaValorXML(ByVal pINXml As String, ByVal psEtiqueta As String)
Dim i As Integer
Dim sCadTempo As String
Dim sCadValor As String
Dim bIniCad As Boolean

    i = 1
    Do While i <= Len(pINXml)
        
        Do While Mid(pINXml, i, 1) <> "<" And i <= Len(pINXml)
            i = i + 1
        Loop
        i = i + 1
        sCadTempo = ""
        Do While Mid(pINXml, i, 1) <> " " And Mid(pINXml, i, 1) <> ">" And i <= Len(pINXml)
            sCadTempo = sCadTempo & Mid(pINXml, i, 1)
            i = i + 1
        Loop
        
        If UCase(Trim(sCadTempo)) = UCase(Trim(psEtiqueta)) Then
                Do While Mid(pINXml, i, 1) <> "=" And Mid(pINXml, i, 1) <> ">" And i <= Len(pINXml)
                    i = i + 1
                Loop
                i = i + 1
                sCadValor = ""
                Do While Mid(pINXml, i, 1) <> "/" And Mid(pINXml, i, 1) <> ">" And i <= Len(pINXml) And Mid(pINXml, i, 1) <> "<"
                    sCadValor = sCadValor & Mid(pINXml, i, 1)
                    i = i + 1
                Loop
                sCadValor = Trim(Replace(sCadValor, """", ""))
                Exit Do
        End If
        
    Loop
    
    RecuperaValorXML = sCadValor

End Function

Public Function GeneraXMLSalida(ByVal RESP_CODE As String, Optional ByVal CARDISS_AMOUNT As String, _
    Optional ByVal AUTH_CODE As String, Optional ByVal ADD_RESP_DATA As String, Optional ByVal CUR_CODE_CARDISS As String, _
    Optional ByVal AMOUNTS_ADD As String, Optional ByVal PRIV_USE As String, Optional ByVal MESSAGE_TYPE As String = "0210") As String

Dim sXML As String


    sXML = "<MESSAGE_TYPE = " & MESSAGE_TYPE & " />"
    sXML = sXML & "<CARDISS_AMOUNT  value=""" & CARDISS_AMOUNT & """ />"
    sXML = sXML & "<AUTH_CODE      value=""" & AUTH_CODE & """ />"  '// Código de autorización (si la transacción es aprobada).
    sXML = sXML & "<RESP_CODE    value=""" & RESP_CODE & """ />" '// Código de Respuesta (Valores Válidos indicados en el Anexo)
    sXML = sXML & "<ADD_RESP_DATA      value=""" & ADD_RESP_DATA & """ />"
    sXML = sXML & "<CUR_CODE_CARDISS   value=""" & CUR_CODE_CARDISS & """ />"
    sXML = sXML & "<AMOUNTS_ADD   value=""" & AMOUNTS_ADD & """ />" '// Datos de la Consulta
    sXML = sXML & "<PRIV_USE   value=""" & PRIV_USE & """ />" '// Glosa a imprimir en el recibo de ATM (aplica en ciertas transacciones)
    
    GeneraXMLSalida = sXML

End Function

Public Function DE_TRAMA_ConvierteAMontoReal(ByVal psMontoTxN As String) As Double

    DE_TRAMA_ConvierteAMontoReal = CDbl(Mid(psMontoTxN, 1, Len(psMontoTxN) - 2) & "." & Right(psMontoTxN, 2))
    
End Function

Public Function ObtieneITF() As Double
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter

    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nITF", adDouble, adParamOutput, , ObtieneITF)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaValorITF"
    Cmd.Execute
    
    
    ObtieneITF = Cmd.Parameters(0).Value
    
    loConec.CierraConexion
    Set loConec = Nothing
    
    Set Cmd = Nothing
    Set Prm = Nothing
    
    
End Function

Public Function RecuperaSaldoDisp(ByVal psCtaCod As String) As Double
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim loConec As New DConecta

    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCodCta", adVarChar, adParamInput, 50, psCtaCod)
    Cmd.Parameters.Append Prm
        
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSaldoDisp", adDouble, adParamOutput, , psCtaCod)
    Cmd.Parameters.Append Prm
        
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaSaldoCuenta"
    Cmd.Execute
    
    RecuperaSaldoDisp = Cmd.Parameters(1).Value
    
    loConec.CierraConexion
    Set loConec = Nothing

    Set Cmd = Nothing
    Set Prm = Nothing

End Function

'****************************************
'RECUPERA COMISION DE TARIFARIO
'****************************************
Public Function ObtieneComisionPorOperacion(ByVal pnTipoEquipo As Integer, ByVal pnTiposerv As Integer, _
    ByVal pnProced As Integer, ByVal pnMoneda As Integer, ByVal pnTipoOper As Integer, _
    ByVal pnMontoOper As Double, ByVal pPAN As String) As Double
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim nComision As Double
Dim loConec As New DConecta

    Set Cmd = New ADODB.Command
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnTipoEquipo", adInteger, adParamInput, , pnTipoEquipo)
    Cmd.Parameters.Append Prm
        
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnTipoServ", adDouble, adParamInput, , pnTiposerv)
    Cmd.Parameters.Append Prm
        
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnProced", adDouble, adParamInput, , pnProced)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMoneda", adDouble, adParamInput, , pnMoneda)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnTipoOper", adDouble, adParamInput, , pnTipoOper)
    Cmd.Parameters.Append Prm
        
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMonto", adDouble, adParamInput, , pnMontoOper)
    Cmd.Parameters.Append Prm
            
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnValor", adDouble, adParamOutput)
    Cmd.Parameters.Append Prm

    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecupComision"
    Cmd.Execute
    
    nComision = IIf(IsNull(Cmd.Parameters(6).Value), 0, Cmd.Parameters(6).Value)
          
    loConec.CierraConexion
    Set loConec = Nothing
    
    Set Cmd = Nothing
    Set Prm = Nothing
    
    ObtieneComisionPorOperacion = nComision
    
End Function

Public Sub RegistraSucesos(ByVal pdFecha As Date, ByVal psProceso As String, ByVal psDescrip As String)
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim loConec As New DConecta
   
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFecha", adDBDate, adParamInput, , pdFecha)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cProceso", adVarChar, adParamInput, 150, psProceso)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cDescripcion", adVarChar, adParamInput, 5000, psDescrip)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RegistraSucesos"
    
    Cmd.Execute
    
    loConec.CierraConexion
    Set loConec = Nothing

    Set Cmd = Nothing
    Set Prm = Nothing
    
End Sub

Public Sub RegistraBitacora(ByVal psNumTarjeta As String, ByVal psCanal As String, ByVal pdFecha As Date, ByVal psIDTrama As String, ByVal psProceso As String, ByVal psDescrip As String, Optional ByVal pnTramaId As Long = 0)
    Dim sSQL As String
    Dim loConec As DConecta
    Set loConec = New DConecta
    sSQL = "exec PIT_stp_ins_RegistraBitacora '" & psNumTarjeta & "','" & psCanal & "','" & Format(pdFecha, "YYYYMMDD HH:MM:SS") & "','" & psIDTrama & "','" & psProceso & "','" & psDescrip & "'," & CStr(pnTramaId)
    loConec.AbreConexion
    loConec.ConexionActiva.Execute sSQL
    loConec.CierraConexion
    Set loConec = Nothing
 End Sub

Public Function ValidaCampos(ByVal psCtaCod As String, ByVal pnMontoTran As Double, _
    ByVal pnMontoITF As Double, ByVal pnMontoComision As Double, ByVal pnMontoComisionITF As Double, _
    ByVal pnMoneda As Integer, ByVal psOpeCod As String, ByVal psOpeCodComision As String, _
    ByVal psOpeCodITF As String, ByVal psOpeCodComisionITF As String, ByVal psOpeExtorno As String, _
    ByVal psOpeExtornoComision As String, ByVal psOpeCodTransferencia As String, ByVal psOpeCodExtornoTransfer As String, _
    ByVal pnTipoCambioCompra As Double, ByVal pnTipoCambioVenta As Double, ByVal psIDTrama As String, ByVal psCtaDeposito As String, _
    ByVal psProceso As String, ByVal pbOpcConsul As Boolean) As Boolean
Dim dFecha As Date
    dFecha = Now()
    
    ValidaCampos = True
    If pbOpcConsul = False Then 'Add By gitu 30-07-2009
        If Len(psCtaCod) <> 18 Then
            ValidaCampos = False
            Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama :" & psIDTrama, "Cuenta con longitud menor a 15 digitos")
        End If
        If (pnMontoTran <= 0 Or pnMontoITF < 0 Or pnMontoComision < 0 Or pnMontoComisionITF < 0) And psProceso = "31" And psProceso = "39" And psProceso = "93" And psProceso = "91" And psProceso = "98" Then
            ValidaCampos = False
            Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama :" & psIDTrama, "Algunos de los Montos son menores que cero")
        End If
        If pnMoneda <> 1 And pnMoneda <> 2 Then
            ValidaCampos = False
            Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama :" & psIDTrama, "Moneda es diferente de 1 o 2")
        End If
    End If 'End Gitu
    
    If pnTipoCambioCompra <= 0 Or pnTipoCambioVenta <= 0 Then
        ValidaCampos = False
        Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama :" & psIDTrama, "Monto de Compra o venta menor o igual a Cero")
    End If
    
    If Len(Trim(psIDTrama)) = 0 Then
        ValidaCampos = False
        Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama :" & psIDTrama, "ID de la Trama en blanco")
    End If
    
    If Len(psCtaDeposito) <> 18 And psProceso = "40" Then
        ValidaCampos = False
        Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama :" & psIDTrama, "Cuenta de Deposito Invalida")
    End If
    
End Function

Public Function ValidaEstadoCuenta(ByVal psCtaCod As String, Optional pnTipo As Integer = 0) As Integer
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim loConec As New DConecta
   
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 18, psCtaCod)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnTipProd", adInteger, adParamInput, , pnTipo)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnResultado", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "PIT_stp_sel_EstadoCuenta"
    
    Cmd.Execute
    ValidaEstadoCuenta = Cmd.Parameters(2).Value
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
            
 
    Set Cmd = Nothing
    Set Prm = Nothing

End Function

Public Function ValidaLimitesOperacion(ByVal psNumTarj As String, ByVal psCodOpeCaj As String, ByVal pnMonto As Double) As Integer

    ValidaLimitesOperacion = 0
    'Valida Retiro de Cajero - Numero Maximo de Retiro
    
            
End Function

'NSSE 07/06/2008
Public Function RecuperaCuentaCascada(ByVal psPAN As String, ByVal psMoneda As String, ByVal pnMonto As Double, _
    ByVal pnTipoCambioCompra As Double, ByVal pnTipoCambioVenta As Double, ByVal psProd As String) As String
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim loConec As New DConecta
   
    Set Cmd = New ADODB.Command
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psPAN", adVarChar, adParamInput, 20, psPAN)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCompra", adDouble, adParamInput, , pnMonto)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@sMoneda", adChar, adParamInput, 1, psMoneda)
    Cmd.Parameters.Append Prm
                                   
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnTipoCambioCompra", adDouble, adParamInput, , pnTipoCambioCompra)
    Cmd.Parameters.Append Prm
                                   
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnTipoCambioVenta", adDouble, adParamInput, , pnTipoCambioVenta)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@sCtaCascada", adVarChar, adParamOutput, 18)
    Cmd.Parameters.Append Prm
    
            
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psProd", adVarChar, adParamInput, 10, psProd)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaCuentaCascada "
    
    Cmd.Execute
    RecuperaCuentaCascada = Cmd.Parameters(5).Value
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
            
 
    Set Cmd = Nothing
    Set Prm = Nothing
    
End Function

'NSSE 07/06/2008

Public Function ValidaOperacion(ByRef psCtaCascada As String, ByVal psCtaCod As String, ByVal pnMontoTran As Double, _
    ByVal pnMontoITF As Double, ByVal pnMontoComision As Double, ByVal pnMontoComisionITF As Double, _
    ByVal pnMoneda As Integer, ByVal psOpeCod As String, ByVal psOpeCodComision As String, _
    ByVal psOpeCodITF As String, ByVal psOpeCodComisionITF As String, ByVal psOpeExtorno As String, _
    ByVal psOpeExtornoComision As String, ByVal psOpeCodTransferencia As String, ByVal psOpeCodExtornoTransfer As String, _
    ByVal pnTipoCambioCompra As Double, ByVal pnTipoCambioVenta As Double, ByVal psIDTrama As String, ByVal psCtaDeposito As String, _
    ByVal pPAN As String, ByVal dFecha As Date, ByVal dFecVenc As Date, ByVal pnCondicion As Integer, _
    ByVal pnRetenerTarjeta As Integer, ByVal pnCtaSaldo As Double, ByVal psProceso As String, _
    ByVal pnNOOperMonExt As Integer, ByVal pnSuspOper As Integer, ByVal pbOpcConsul As Boolean, ByVal psDNI As String, _
    ByRef psCadResp As String) As String


Dim sRespuetas As String
Dim lnValResp As Integer

'00 Acepta la Transaccion
ValidaOperacion = "00"

'14 Si recibe un numero de tarjeta Invalido
If Len(pPAN) < 16 And (gPRCODE = "351100" Or Left(gPRCODE, 2) = "94" Or Left(gPRCODE, 2) = "01") Then
    ValidaOperacion = "14" 'Tarjeta Invalida
    psCadResp = "Numero Tarjeta Invalida"
    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Numero de Tarjeta Invalido", gnTramaId)
    Exit Function
End If


If pnCondicion <> 1 And (psOpeCod = CStr(gPITCaptacRetiro) Or psOpeCod = CStr(gPITCaptaConsCtasAhorro) Or psOpeCod = CStr(gPITCaptaConsMovCtasAho) _
                        Or psOpeCod = CStr(gPITCaptaConsSaldo)) Then
    ValidaOperacion = "14" 'Tarjeta Invalida, Cancelada, etc.
    psCadResp = "Tarjeta Invalida o Cancelada"
    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Tarjeta Invalida o Cancelada", gnTramaId)
    Exit Function
End If

'43 Si la tarjeta esta en la condicion de robada o HOT(Fuerza al ATM a Capturar la Tarjeta)
' Se usa la Tabla Tarjeta campo nCondicion=10 y nRetenerTarjeta=1
If pnCondicion = 10 Or pnRetenerTarjeta = 1 Then
    ValidaOperacion = "41"
    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Tarjeta Retenida", gnTramaId)
    Exit Function
End If

'81 Si la tarjeta esta cancelada
'Se usa la Tabla Tarjeta campo nCondicion=50
If pnCondicion = 50 Then
    ValidaOperacion = "90"
    Exit Function
End If

'33 Si la tarjeta esta Vencida
If dFecha > dFecVenc And (psOpeCod = CStr(gPITCaptacRetiro) Or psOpeCod = CStr(gPITCaptaConsCtasAhorro) Or psOpeCod = CStr(gPITCaptaConsMovCtasAho) _
                          Or psOpeCod = CStr(gPITCaptaConsSaldo)) Then
    ValidaOperacion = "54" 'Tarjeta Vencida
    psCadResp = "Tarjeta Vencida"
    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Tarjeta Vencida", gnTramaId)
    Exit Function
End If


If pnMoneda = 1 Then
    If Mid(gPRCODE, 1, 2) = "01" Then
        If pnMontoTran < gnMontoMinRetMN Then
            ValidaOperacion = "65" 'MONTO MENOR AL MINIMO DE RETIRO
            psCadResp = "Monto menor al minimo de Retiro"
            Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", psCadResp, gnTramaId)
            Exit Function
        End If
        If pnMontoTran > gnMontoMaxRetMN Then
            ValidaOperacion = "65" 'MONTO MENOR AL MINIMO DE RETIRO
            psCadResp = "Monto mayor al minimo de Retiro"
            Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", psCadResp, gnTramaId)
            Exit Function
        End If
        
    End If
Else
    If (pnMontoTran < gnMontoMinDepMN Or pnMontoTran > gnMontoMaxDepMN) And Mid(gPRCODE, 1, 2) = "01" Then
        ValidaOperacion = "65" 'MONTO MAYOR AL MAXIMO DE RETIRO
        psCadResp = "Monto mayor al maximo de retiro"
        Exit Function
    End If
End If

If pnMoneda = 1 Then
    If (pnMontoTran < gnMontoMinRetMN Or pnMontoTran > gnMontoMaxRetMN) And Mid(gPRCODE, 1, 2) = "20" Then
        ValidaOperacion = "65" 'MONTO MENOR AL MINIMO DE DEPOSITO
        psCadResp = "Monto mayor al minimo de deposito"
        Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Algunos de los Montos son menores que cero", gnTramaId)
        Exit Function
    End If
Else
    If (pnMontoTran < gnMontoMinDepME Or pnMontoTran > gnMontoMaxDepME) And Mid(gPRCODE, 1, 2) = "20" Then
        ValidaOperacion = "65" 'MONTO MAYOR AL MAXIMO DE DEPOSITO
        psCadResp = "Monto mayor al maximo de deposito"
        Exit Function
    End If
End If

lnValResp = 0
If Left(gPRCODE, 2) = "01" Or Left(gPRCODE, 2) = "20" Or Left(gPRCODE, 4) = "9411" Then
    lnValResp = ValidaEstadoCuenta(psCtaCod, 1) 'Ahorros
ElseIf Left(gPRCODE, 2) = "50" Or Left(gPRCODE, 4) = "9435" Then
    lnValResp = ValidaEstadoCuenta(psCtaCod, 2) ' créditos
End If

If pbOpcConsul = False Then 'Add By gitu 30-07-2009
    If Len(psCtaCod) <> 18 Then
        ValidaOperacion = "78"
        psCadResp = "Cuenta con longitud menor a los 18 digitos"
        Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Cuenta con longitud menor a 18 digitos", gnTramaId)
        Exit Function
    End If

    If (pnMontoTran <= 0 Or pnMontoITF < 0 Or pnMontoComision < 0 Or pnMontoComisionITF < 0) And psProceso = "31" And psProceso = "39" And psProceso = "93" And psProceso = "91" And psProceso = "98" Then
        ValidaOperacion = "65"
        psCadResp = "Monto de la operacion menores que cero"
        Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Algunos de los Montos son menores que cero", gnTramaId)
        Exit Function
    End If
    
    Select Case lnValResp
        Case 1
            ValidaOperacion = "53"
            psCadResp = "Cuenta con estado nulo"
            Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Cuenta con estado Nulo", gnTramaId)
            Exit Function
        Case 2
            ValidaOperacion = "80"
            psCadResp = "Cuenta bloqueada para retiros"
            Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Cuenta Bloqueada para Retiros", gnTramaId)
            Exit Function
        Case 3
            ValidaOperacion = "81"
            psCadResp = "Cuenta bloqueada totalmente"
            Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Cuenta Bloqueada Totalmente", gnTramaId)
            Exit Function
        Case 4
            ValidaOperacion = "82"
            psCadResp = "Cuenta cancelada"
            Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Cuenta Cancelada", gnTramaId)
            Exit Function
        Case 5
            ValidaOperacion = "78"
            psCadResp = "Cuenta no existe"
            Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Cuenta no Existe", gnTramaId)
            Exit Function
    End Select
    
    If lnValResp = 6 And psOpeCod = gPITCaptaDeposito Then
        ValidaOperacion = "12"
        psCadResp = "Cuenta de heberes No puede realizar deposito"
        Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Cuenta de Haberes", gnTramaId)
        Exit Function
    End If
    
    If pnCtaSaldo < pnMontoTran And psOpeCod = gPITCaptacRetiro Then
        ValidaOperacion = "51"
        psCadResp = "Cuenta no posee saldo sufieciente"
        Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Cuenta no posee saldo suficiente", gnTramaId)
        Exit Function
    End If
    
End If 'End Gitu


If Left(gPRCODE, 2) = "01" Or Left(gPRCODE, 2) = "20" Or Left(gPRCODE, 2) = "50" Then 'Verificar moneda
    If pnMoneda <> CInt(Mid(psCtaCod, 9, 1)) Then
        ValidaOperacion = "12"
        psCadResp = "Moneda de la operacion diferente al de la cuenta"
        Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Moneda de transacción difiere a la moneda de la cuenta", gnTramaId)
        Exit Function
    End If
End If

If Left(gPRCODE, 2) = "01" Then
    If (pnMoneda = 1 And pnMontoTran > gnMontoMinRetMNReqDNI) Or (pnMoneda = 2 And pnMontoTran > gnMontoMinRetMEReqDNI) Then
        If ValidaDNI(psCtaCod, psDNI) = 1 Then
            ValidaOperacion = "34"
            psCadResp = "DNI no pertenece al titular de la cuenta"
            Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "DNI no pertenece al titular de la cuenta", gnTramaId)
            Exit Function
        End If
    End If
End If

If pnTipoCambioCompra <= 0 Or pnTipoCambioVenta <= 0 Then
    ValidaOperacion = 65
    psCadResp = "Monto de Compra o venta menor o igual a Cero"
    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Monto de Compra o venta menor o igual a Cero", gnTramaId)
    Exit Function
End If
'
'If Len(Trim(psIDTrama)) = 0 Then
'    ValidaCampos = False
'    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama :" & psIDTrama, "ID de la Trama en blanco")
'End If

If Len(psCtaDeposito) <> 18 And psProceso = "40" Then
    ValidaOperacion = 14
    psCadResp = "Cuenta de Deposito Invalida"
    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Validacion de Campos - Trama ", "Cuenta de Deposito Invalida", gnTramaId)
    Exit Function
End If

'51 Si la cuenta no tiene el saldo suficiente para atender el requerimiento del Cliente
If pbOpcConsul = False Then
     psCtaCascada = psCtaCod
    
    If pnNOOperMonExt = 1 And (Mid(psCtaCod, 9, 1) = "2" Or _
            (Mid(psProceso, 1, 2) = "40" And Mid(psCtaCod, 9, 1) = "2") _
                Or pnMoneda = 2) Then
        ValidaOperacion = "90" 'Servidor en cierre o regreso de offhost
        Exit Function
    End If
    
'    If VerificaEstadoCuenta(psCtaCod) <> 0 Then
'        ValidaOperacion = "53" 'Cuenta invalida
'        Exit Function
'    End If
'
'    If Mid(psProceso, 1, 2) = "40" Then
'        If VerificaEstadoCuenta(psCtaCod) <> 0 Then
'            ValidaOperacion = "53" 'Cuenta Invalida
'            Exit Function
'        End If
'    End If
End If


'80 Cuando por algun motivo interno del banco no puede procesar la transaccion.
If pnSuspOper <> 0 Then
    ValidaOperacion = "90"
    Exit Function
End If

End Function

Public Function RecuperaSaldoDeCuenta(ByVal psCtaCod As String, ByRef pnSaldoDisp As Double, _
    ByRef pnSaldoTot As Double) As Double
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim loConec As New DConecta

    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 18, psCtaCod)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnSaldoDisponilble", adDouble, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnSaldoContable", adDouble, adParamOutput)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ConsultaSaldo"
    
    Cmd.Execute
      
    pnSaldoDisp = Cmd.Parameters(1).Value
    pnSaldoTot = Cmd.Parameters(2).Value
    
    loConec.CierraConexion
    Set loConec = Nothing

    Set Cmd = Nothing
    Set Prm = Nothing
End Function

Public Function RecuperaTipoCambio(ByVal pdFecha As Date, ByRef pnCompra As Double, _
    ByRef pnVenta As Double) As Double
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim loConec As New DConecta

    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pdFecha", adDate, adParamInput, 18, pdFecha)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnTCCompra", adDouble, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnTCVenta", adDouble, adParamOutput)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ConsultaTC"
    
    Cmd.Execute
      
    pnCompra = Cmd.Parameters(1).Value
    pnVenta = Cmd.Parameters(2).Value
    
    loConec.CierraConexion
    Set loConec = Nothing
    
    Set Cmd = Nothing
    Set Prm = Nothing
            
End Function

Public Function ValidaLimitesOperacionATMPOS(ByVal psNumTarj As String, ByVal psCodTran As String, ByVal pnMonto As Double, ByVal pnTiposerv As Integer, _
    ByVal pnMoneda As Integer) As Integer
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim loConec As New DConecta

    Set Cmd = New ADODB.Command
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 50, psNumTarj)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pcCodTran", adVarChar, adParamInput, 50, psCodTran)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMonto", adDouble, adParamInput, , pnMonto)
    Cmd.Parameters.Append Prm
                
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnTiposerv", adInteger, adParamInput, , pnTiposerv)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nMoneda", adInteger, adParamInput, , pnMoneda)
    Cmd.Parameters.Append Prm
                
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnResul", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ValidaLimitesOperativos"
    
    Cmd.Execute
      
    ValidaLimitesOperacionATMPOS = Cmd.Parameters(5).Value
    
    loConec.CierraConexion
    Set loConec = Nothing
    
    Set Cmd = Nothing
    Set Prm = Nothing
    
End Function




Public Sub RegistrarTrama(ByVal psIDTrama As String, ByVal psTramas As String, ByVal pnDenegada As Integer)
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim loConec As New DConecta
    
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pcIDTrama", adVarChar, adParamInput, 50, psIDTrama)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cTrama", adVarChar, adParamInput, 5000, psTramas)
    Cmd.Parameters.Append Prm
               
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nDenegada", adInteger, adParamInput, , pnDenegada)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RegistrarTrama"
    
    Cmd.Execute
    
    loConec.CierraConexion
    Set loConec = Nothing
        
    Set Cmd = Nothing
    Set Prm = Nothing

End Sub


Public Function nRegistrarTramaRecepcion(psMovNro As String, psOpeCod As String, psTramaRecep As String, psCodTX As String, psCuentaTX As String, _
                                         psPANTX As String, psDNITX As String, psMonedaTX As String, pnMontoTX As Double, psCodInstTX As String) As Long
Dim lCn As New DConecta
Dim lRs As New ADODB.Recordset
Dim lsSql As String

    lsSql = "exec PIT_stp_ins_RegistraTramaRecepcionModoAut '" & psMovNro & "','" & psOpeCod & "','" & psTramaRecep & "','" & psCodTX & "','" & psCuentaTX & "','" & _
                                         psPANTX & "','" & psDNITX & "','" & psMonedaTX & "'," & pnMontoTX & ",'" & psCodInstTX & "'"
    

    lCn.AbreConexion

    Set lRs = lCn.ConexionActiva.Execute(lsSql)
    
    nRegistrarTramaRecepcion = lRs("nTramaId")
    gnTramaId = nRegistrarTramaRecepcion
    lCn.CierraConexion
    
    Set lCn = Nothing
    Set lRs = Nothing

End Function

Public Sub RegistrarTramaEnvio(pnTramaId As Integer, psTramaEnvio As String, psCodRespTX As String, pnDenegada As Integer)
Dim lCn As New DConecta
Dim lsSql As String
    
    
    lsSql = "exec PIT_stp_upd_TramaEnvioModoAut " & pnTramaId & ",'" & psTramaEnvio & "','" & psCodRespTX & "'," & pnDenegada
    
    lCn.AbreConexion

    lCn.ConexionActiva.Execute (lsSql)
    
    lCn.CierraConexion
    
    Set lCn = Nothing

End Sub

Public Function RecuperaCtaDisponible(ByVal psNumTarjeta As String, ByVal pnMonCta As Integer, ByVal psProd As String) As String
Dim sCta As String
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim loConec As New DConecta
    
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psTarjeta", adVarChar, adParamInput, 20, psNumTarjeta)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psProd", adVarChar, adParamInput, 20, psProd)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamOutput, 18)
    Cmd.Parameters.Append Prm
                
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMonCta", adInteger, adParamInput, , pnMonCta)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaCuetaDisponible"
    
    Cmd.Execute
    
    sCta = IIf(IsNull(Cmd.Parameters(2).Value), "", Cmd.Parameters(2).Value)
    
    loConec.CierraConexion
    Set loConec = Nothing
        
    Set Cmd = Nothing
    Set Prm = Nothing
    
    RecuperaCtaDisponible = sCta

End Function

Public Function RecuperaMovimDeCuenta(ByVal psCtaCod As String, pnMoneda As Integer) As String
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim R As ADODB.Recordset
Dim nSaldoDisp As Double
Dim nSaldoTot  As Double
Dim sCabe As String 'DAOR 20081112
Dim sFecha As String
Dim lsMes As String
Dim lsSimMon As String
Dim loConec As New DConecta

    sCabe = "1P" & Format(Now, "YYMMDD") & "0040"
    RecuperaMovimDeCuenta = sCabe
    RecuperaMovimDeCuenta = RecuperaMovimDeCuenta & " FECHA  MOVIMIENTO    MONTO       ITF   "
    
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 18, psCtaCod)
    Cmd.Parameters.Append Prm
                
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "PIT_stp_sel_ConsultaMovCuenta"
    
    Set R = New ADODB.Recordset
    R.Open Cmd
    
    sFecha = IIf(IsNull(R!Fecha), "", R!Fecha)
    
    lsMes = DevuelveMes(Mid(sFecha, 3, 2))
    
    '**DAOR 20081112 ***************************************************************
    If Not (R.EOF Or R.BOF) Then
        'R.Sort = "Fecha"
        Do While Not R.EOF
            RecuperaMovimDeCuenta = RecuperaMovimDeCuenta & " " & UCase(Format(R!Fecha, "DDMMM"))
            RecuperaMovimDeCuenta = RecuperaMovimDeCuenta & "  " & UCase(Left(R!Operacion & Space(12), 12))
            RecuperaMovimDeCuenta = RecuperaMovimDeCuenta & "  " & UCase(Right(Space(10) & Format(R!Monto, "#,0.00"), 10))
            RecuperaMovimDeCuenta = RecuperaMovimDeCuenta & "  " & UCase(Right(Space(6) & Format(R!MontoITF, "#,0.00"), 6))
            R.MoveNext
        Loop
    Else
        RecuperaMovimDeCuenta = RecuperaMovimDeCuenta & "    *** CUENTA SIN MOVIMIENTOS ***      "
    End If
    R.Close
    
    'Call CerrarConexion
    loConec.CierraConexion

    
    Call RecuperaSaldoDeCuenta(psCtaCod, nSaldoDisp, nSaldoTot)
    If pnMoneda = 1 Then
        lsSimMon = "S/."
    Else
        lsSimMon = "$. "
    End If
    
    RecuperaMovimDeCuenta = RecuperaMovimDeCuenta & " SALDO CONTABLE    " & lsSimMon & "      " & Right(Space(12) & Format(nSaldoTot, "#,0.00"), 12)
    RecuperaMovimDeCuenta = RecuperaMovimDeCuenta & " SALDO DISPONIBLE  " & lsSimMon & "      " & Right(Space(12) & Format(nSaldoDisp, "#,0.00"), 12)
    
    RecuperaMovimDeCuenta = RecuperaMovimDeCuenta & psCtaCod
    '*******************************************************************************
             
    'RecuperaMovimDeCuenta = CStr(Len(RecuperaMovimDeCuenta) + 16) & RecuperaMovimDeCuenta
    Set Cmd = Nothing
    Set Prm = Nothing
    Set R = Nothing
    
    Set loConec = Nothing
            
End Function
'Create by GITU 22-07-2009 Recupera todas las cuentas asociadas a una tarjeta
Public Function RecuperaCuentasAho(ByVal psNumTarjeta As String) As String
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim R As ADODB.Recordset
Dim sCabe As String 'DAOR 20081112
Dim loConec As New DConecta

    sCabe = "1P" & Format(Now, "YYMMDD") & "0040"
    RecuperaCuentasAho = sCabe
    RecuperaCuentasAho = RecuperaCuentasAho & "CUENTA            MONEDA         SALDO  "
    
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psTarjeta", adVarChar, adParamInput, 18, psNumTarjeta)
    Cmd.Parameters.Append Prm
 
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMoneda", adInteger, adParamInput, , nMoneda)
    Cmd.Parameters.Append Prm
                
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ConsultaCuentasAho"
    
    Set R = New ADODB.Recordset
    R.Open Cmd
      
    If Not (R.EOF Or R.BOF) Then
        Do While Not R.EOF
            RecuperaCuentasAho = RecuperaCuentasAho & UCase(Trim(R!cCtaCod))
            RecuperaCuentasAho = RecuperaCuentasAho & " " & UCase(Trim(R!MONEDA)) & Space(5)
            RecuperaCuentasAho = RecuperaCuentasAho & " " & UCase(Right(Space(12) & Format(R!SaldoDisponible, "#,0.00"), 12))
            R.MoveNext
        Loop
    Else
        RecuperaCuentasAho = RecuperaCuentasAho & "    *** TARJETA SIN CUENTAS ***        "
    End If
    R.Close
    
    loConec.CierraConexion
    'RecuperaCuentasAho = CStr(Len(RecuperaCuentasAho) + 16) & RecuperaCuentasAho
    '*******************************************************************************
    
    Set Cmd = Nothing
    Set Prm = Nothing
    Set R = Nothing
    
    Set loConec = Nothing
            
End Function
'Create by GITU 22-07-2009 Recupera las cuentas de credito
Public Function RecuperaCuentasCred(ByVal psDNI As String, ByVal pdFecSis As String) As String
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim R As ADODB.Recordset
Dim RGas As ADODB.Recordset
Dim sCabe As String 'DAOR 20081112
Dim loConec As New DConecta
Dim MatCalend As Variant
Dim nMonPago As Double
Dim nMora As Double
Dim nCuotasMora As Double
Dim sFecha As String
Dim sCtaCod As String
Dim nGastos As Double
Dim dFecIni As String
Dim nCuotaAFecha As Double
Dim lsMes As String
Dim i As Integer
Dim nNroCuotas As Integer

    sCabe = "1P" & Format(Now, "YYMMDD") & "0040"
    RecuperaCuentasCred = sCabe
    RecuperaCuentasCred = RecuperaCuentasCred & "NRO CUENTA         MONEDA MONTO FEC VENC"
    
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psDNI", adVarChar, adParamInput, 13, psDNI)
    Cmd.Parameters.Append Prm
                
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMoneda", adInteger, adParamInput, , nMoneda)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "PIT_ConsultaCuentasCred"
    
    Set R = New ADODB.Recordset
    R.Open Cmd

    pdFecSis = Format(pdFecSis, "dd/mm/yyyy")
    
    If Not (R.EOF Or R.BOF) Then
        
        Do While Not R.EOF
        
            sCtaCod = R!cCtaCod
            MatCalend = RecuperaMatrizCalendarioPendiente(sCtaCod, , , nNroCuotas)
            
            Set RGas = loConec.CargaRecordSet("SELECT nGasto=DBCMAC.dbo.ColocCred_ObtieneGastoFechaCredito('" & sCtaCod & "','" & Format(pdFecSis, "mm/dd/yyyy") & "')")
            nGastos = RGas!nGasto
        
            nMonPago = MatrizMontoAPagar(MatCalend, pdFecSis)
            
            Dim nInteresFecha As Currency
            Dim nInterFechaGra As Currency
            Dim nMontoFecha As Currency
            
            
            nMora = Format(MatrizMoraTotal(MatCalend, pdFecSis), "#0.00")
            
'
            nCuotasMora = MatrizCuotasEnMora(MatCalend, pdFecSis)
            
            nInteresFecha = DevuelveInteresAFecha(sCtaCod, MatCalend, pdFecSis)
            nInterFechaGra = MatrizInteresGraAFecha(sCtaCod, MatCalend, pdFecSis)
            'nMontoFecha = MatrizCapitalAFecha(sCtaCod, MatCalend, pdFecSis)

            'mody by gitu 20-06-2009
            nCuotaAFecha = Format(nInteresFecha + nInterFechaGra + nMonPago, "#0.00")

            'ARCV 28-02-2007
            nCuotaAFecha = Format(nCuotaAFecha + RGas!nGasto, "#0.00")
            

            If nNroCuotas = Val(MatCalend(0, 1)) And nCuotaAFecha < nMonPago Then
                nMonPago = nCuotaAFecha
            End If
            
            sFecha = Format(MatCalend(0, 0), "DDMMYY")
            
            lsMes = DevuelveMes(Mid(sFecha, 3, 2))
            
            RecuperaCuentasCred = RecuperaCuentasCred & UCase(Trim(R!cCtaCod))
            RecuperaCuentasCred = RecuperaCuentasCred & " " & UCase(Trim(R!sMoneda))
            RecuperaCuentasCred = RecuperaCuentasCred & " " & UCase(Right(Space(10) & Format(nMonPago, "#,0.00"), 10))
            RecuperaCuentasCred = RecuperaCuentasCred & " " & Left(sFecha, 2) & lsMes & Right(sFecha, 2)
            R.MoveNext
        Loop
    Else
        RecuperaCuentasCred = RecuperaCuentasCred & "    *** CLIENTE SIN CUENTAS ***      "
    End If
    R.Close
    
    'RecuperaCuentasCred = CStr(Len(RecuperaCuentasCred) + 16) & RecuperaCuentasCred
    
    loConec.CierraConexion

    '*******************************************************************************
    
    Set Cmd = Nothing
    Set Prm = Nothing
    Set R = Nothing
    
    Set loConec = Nothing
            
End Function


'NSSE 07/06/2008
Public Function RecuperaUserATM(ByVal psTerminalID As String) As String
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim loConec As New DConecta
   
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psTerminalID", adVarChar, adParamInput, 50, psTerminalID)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psUser", adChar, adParamOutput, 4)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaUserATM"
    
    Cmd.Execute
    
    RecuperaUserATM = Cmd.Parameters(1).Value
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    
    Set Cmd = Nothing
    Set Prm = Nothing
    
End Function

Public Function RecuperaConsultaIntegrada(ByVal psNumTarjeta As String) As String
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim R As ADODB.Recordset
Dim loConec As New DConecta

    RecuperaConsultaIntegrada = ""
    
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarjeta", adVarChar, adParamInput, 18, psNumTarjeta)
    Cmd.Parameters.Append Prm
                
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ConsultaIntegral"
    
    RecuperaConsultaIntegrada = "CONSULTA INTEGRADA;"
    Set R = New ADODB.Recordset
    R.Open Cmd
    Do While Not R.EOF
        'Recorrer y armar la cadena
        '3 bytes Agencia
        '1 byte espacio en blanco
        '15 bytes para identificar la cuenta
        '1 byte espacio en blanco
        '1 byte para indicar la moneda de la cuenta
        '16 bytes para indicar el saldo disponible de la cuenta
        '1 byte para poner "-" en caso de saldo negativo
        '1 byte para cambio de linea emplear ";"
        RecuperaConsultaIntegrada = RecuperaConsultaIntegrada & "0" & Mid(R!cCtaCod, 4, 2) & " " & Right("000000000000000" & Mid(R!cCtaCod, 6, 13), 15) & " "
        RecuperaConsultaIntegrada = RecuperaConsultaIntegrada & IIf(Mid(R!cCtaCod, 9, 1) = "1", "gPRCODE/.", "US$")
        'LSDO 2008/01/07
        'RecuperaConsultaIntegrada = RecuperaConsultaIntegrada & Right("                " & Replace(Format(R!SaldoDisponible, "#0.00"), ".", ""), 16)
        RecuperaConsultaIntegrada = RecuperaConsultaIntegrada & Right("                " & Format(R!SaldoDisponible, "#,0.00"), 16)
        RecuperaConsultaIntegrada = RecuperaConsultaIntegrada & " ;"
    
        R.MoveNext
    Loop
    R.Close
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    
    Set Cmd = Nothing
    Set Prm = Nothing
    Set R = Nothing
            
End Function
Public Function Transaccion(ByVal psEntrada As String) As String

End Function
 
 'NSSE 05/12/2008
Public Function RecuperaCondicionDeTarjeta(ByVal psPAN As String) As Integer
Dim cmdNeg As New Command
Dim Prm As New ADODB.Parameter
Dim prmNegFecha As New ADODB.Parameter
Dim loConec As New DConecta

    '***************************************************
    'RECUPERA DATOS DE TARJETA
    '***************************************************
    Set cmdNeg = New ADODB.Command
    Set prmNegFecha = New ADODB.Parameter
    Set prmNegFecha = cmdNeg.CreateParameter("@PAN", adVarChar, adParamInput, 20, psPAN)
    cmdNeg.Parameters.Append prmNegFecha
    
    Set prmNegFecha = New ADODB.Parameter
    Set prmNegFecha = cmdNeg.CreateParameter("@nCondicion", adInteger, adParamOutput)
    cmdNeg.Parameters.Append prmNegFecha
    
    Set prmNegFecha = New ADODB.Parameter
    Set prmNegFecha = cmdNeg.CreateParameter("@nRetenerTarjeta", adInteger, adParamOutput)
    cmdNeg.Parameters.Append prmNegFecha
    
    Set prmNegFecha = New ADODB.Parameter
    Set prmNegFecha = cmdNeg.CreateParameter("@nNOOperMonExt", adInteger, adParamOutput)
    cmdNeg.Parameters.Append prmNegFecha
    
    Set prmNegFecha = New ADODB.Parameter
    Set prmNegFecha = cmdNeg.CreateParameter("@nSuspOper", adInteger, adParamOutput)
    cmdNeg.Parameters.Append prmNegFecha
    
    Set prmNegFecha = New ADODB.Parameter
    Set prmNegFecha = cmdNeg.CreateParameter("@dfecVenc", adDate, adParamOutput)
    cmdNeg.Parameters.Append prmNegFecha
    
    Set prmNegFecha = New ADODB.Parameter
    Set prmNegFecha = cmdNeg.CreateParameter("@psDescEstado", adVarChar, adParamOutput, 100)
    cmdNeg.Parameters.Append prmNegFecha
    
    loConec.AbreConexion
    cmdNeg.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    cmdNeg.CommandType = adCmdStoredProc
    cmdNeg.CommandText = "ATM_RecuperaDatosTarjeta"
    
    cmdNeg.Execute
      
    RecuperaCondicionDeTarjeta = cmdNeg.Parameters(1).Value
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    
    Set cmdNeg = Nothing
    Set prmNegFecha = Nothing
    
End Function


Public Function TransaccionGlobalNet(ByVal psMESSAGE_TYPE As String, ByVal psTRACE As String, ByVal psPRCODE As String, _
    ByVal psPAN As String, ByVal psTIME_LOCAL As String, ByVal psDATE_LOCAL As String, ByVal psTERMINAL_ID As String, _
    ByVal psACCT_1 As String, ByVal psCARD_ACCEPTOR As String, ByVal psACQ_INST As String, ByVal psPOS_COND_CODE As String, _
    ByVal psTXN_AMOUNT As String, ByVal psCUR_CODE As String, ByVal psACCT_2 As String, ByVal psDATE_EXP As String, ByVal psCARD_LOCATION As String, _
    ByVal psMonCta As String, ByVal psMovNro As String, ByVal pnTramaId As Long, Optional ByVal psCtaCod As String, Optional psDNI As String) As String

Dim sCtaDeposito As String
Dim XmlExt As String
Dim nResValLimOper As Integer
Dim sXMLTrama As String
Dim sCtaCascada As String
Dim sValida As String
Dim nMontoTotal As Double
 
    gMESSAGE_TYPE = psMESSAGE_TYPE
    gTRACE = psTRACE
    gPRCODE = psPRCODE
    gsPAN = psPAN
    gTIME_LOCAL = psTIME_LOCAL
    gDATE_LOCAL = psDATE_LOCAL
    gTERMINAL_ID = psTERMINAL_ID
    gACCT_1 = psACCT_1
    gCARD_ACCEPTOR = psCARD_ACCEPTOR
    gACQ_INST = psACQ_INST
    gPOS_COND_CODE = psPOS_COND_CODE
    gTXN_AMOUNT = psTXN_AMOUNT
    gCUR_CODE = psCUR_CODE
    gACCT_2 = psACCT_2
    gDATE_EXP = psDATE_EXP
    gCARD_LOCATION = psCARD_LOCATION
    gsMonCta = psMonCta
    gnTramaIDExt = pnTramaId
    gsMovNro = psMovNro
    gsCtaCod = psCtaCod
    gsDNI = psDNI

    sIDTrama = psTRACE
        
    gPRCODE = IIf(Len(gPRCODE) <> 6, Right("000000" & gPRCODE, 6), gPRCODE)
    
    sOpeCod = getOpeCod(gPRCODE, gMESSAGE_TYPE)
    
    sXMLTrama = GeneraTramaEnXML()
    pINXml = sXMLTrama
    
 
    Call RecuperaDatosNegocio
    
    Call InicializaParametros
        
        
    If gMESSAGE_TYPE = "0800" Then ' Tokens, LogOn/LogOff
        TransaccionGlobalNet = LogONOFF()
        Exit Function
    End If
    
    
    If gMESSAGE_TYPE = "0420" Then ' Solicitud de reverso de transacción (extorno)
        TransaccionGlobalNet = TXExtornoTransaccion()
        Exit Function
    End If
    
    
    If gMESSAGE_TYPE = "0200" Then ' Solicitud de transacción financiera
        Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Inicio de Transaccion " & gMESSAGE_TYPE, "", gnTramaId)
        sCtaCod = psCtaCod
        If Mid(gPRCODE, 1, 2) = "35" Then
            bOpcConsu = True
        End If
        
        
        If Mid(gPRCODE, 1, 2) <> "35" Then
            
            gTXN_AMOUNT = IIf(gTXN_AMOUNT = "[.....]", "000000000000", gTXN_AMOUNT)
            
            nMontoTran = DE_TRAMA_ConvierteAMontoReal(gTXN_AMOUNT)
            
            nMontoITF = CalculaITF(nMontoTran, sCtaCod)
            
            'nMontoComision = ObtieneComisionPorOperacion(nTipoEquipo, nTipoServicio, nProced, _
                IIf(gCUR_CODE = "604", 1, 2), nTipoOperac, DE_TRAMA_ConvierteAMontoReal(gTXN_AMOUNT), gsPAN)
            nMontoComision = 0
            
            'nMontoComisionITF = CalculaITF(nMontoComision, sCtaCod) 'CDbl(Format(ObtieneITF * nMontoComision, "#0.00"))
            nMontoComisionITF = 0
            
        End If
        
        Call RecuperaDatosTarjeta
        
        If Mid(gPRCODE, 1, 2) <> "35" Then
            Call RecuperaDatosCuenta
        End If
        
        'Validación de operación
        sCtaCascada = ""
        sCtaCod = gsCtaCod
        
        sValida = ValidaOperacion(sCtaCascada, sCtaCod, nMontoTran, nMontoITF, nMontoComision, nMontoComisionITF, _
                                    nMoneda, sOpeCod, sOpeCodComision, sOpeCodITF, sOpeCodComisionITF, sOpeExtorno, _
                                    sOpeExtornoComision, sOpeCodTransferencia, sOpeCodExtornoTransfer, _
                                    nTipoCambioCompra, nTipoCambioVenta, sIDTrama, sCtaDeposito, _
                                    gsPAN, dFecSis, dFecVenc, nTarjCondicion, nRetenerTarjeta, nCtaSaldo, gPRCODE, nNOOperMonExt, nSuspOper, bOpcConsu, gsDNI, sCadResp)
        
        
        If sValida <> "00" Then
            pOUTXml = GeneraXMLSalida(sValida, , sHora, sCtaCod, , sCadAmount, sCadResp)
            TransaccionGlobalNet = pOUTXml
            Exit Function
        End If
        
        Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Valido la Operacion ", "", gnTramaId)
    
        If nResValLimOper = 0 Then
            nMontoComision = 0
            nMontoComisionITF = 0
        End If
        Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Valido Limites de Operacion ", "", gnTramaId)
        
        Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Inicio de Operacion ", "", gnTramaId)
        Select Case Mid(gPRCODE, 1, 2) 'Por cada tipo de operación
            Case "01" 'Retiros
                TransaccionGlobalNet = TXRetiro()
                Exit Function
            Case "20" 'Depósitos
                TransaccionGlobalNet = TXDeposito()
                Exit Function
            Case "50" 'Pago de crédito
                TransaccionGlobalNet = TXPagoCredito()
                Exit Function
            Case "31" 'Consulta saldo cta. ahorro
                TransaccionGlobalNet = TXConsultaCuentaAhorro()
                Exit Function
            Case "35" 'Consulta de cuentas de ahorro, Consulta de cuentas de crédito
                If Mid(gPRCODE, 3, 2) = "11" Then 'Consulta ahorro
                    TransaccionGlobalNet = TXConsultaCuentasAhorro
                    Exit Function
                ElseIf Mid(gPRCODE, 3, 2) = "35" Then 'Consulta crédito
                    TransaccionGlobalNet = TXConsultaCuentasCredito
                    Exit Function
                End If
            Case "94" 'Consulta de movimientos de cuentas de ahorro, Consulta de movimientos de cuentas de crédito
                If Mid(gPRCODE, 3, 2) = "11" Then 'Consulta ahorro
                    TransaccionGlobalNet = TXConsultaMovimientosAhorro
                    Exit Function
                ElseIf Mid(gPRCODE, 3, 2) = "35" Then 'Consulta crédito

                End If
        End Select
    
        Exit Function
    End If
  
    
    Set loConec = Nothing
    TransaccionGlobalNet = pOUTXml
End Function

Sub InicializaParametros()
    Dim rsParametros As ADODB.Recordset
    'Tipo de equipo de cajero
    nTipoEquipo = 1
    If gCARD_ACCEPTOR <> "000000000000000" Then 'Procedencia Perú (PE) o internacional
        nProced = IIf(Right(gCARD_LOCATION, 2) = "PE", 1, 2)
    Else
        nProced = 1
    End If
    
    nTipoServicio = 2
    
    If nTipoEquipo = 2 Then  'Si tipo equipo es POS, entonces servicio=Compras
        nTipoServicio = 4
    End If

    If Mid(gPRCODE, 1, 2) = "31" Or Mid(gPRCODE, 1, 2) = "96" Or Mid(gPRCODE, 1, 2) = "94" Then 'Consultas y cambio de clave
        nTipoOperac = 2
    Else
        nTipoOperac = 1
    End If
    
    sHora = Trim(gTIME_LOCAL)
    sMesDia = Trim(gDATE_LOCAL)
    sUserATM = Trim(RecuperaUserATM(Trim(gTERMINAL_ID)))
    If Len(Trim(sUserATM)) = 0 Then sUserATM = "AT00"

    Select Case gCUR_CODE
        Case "604"
            nMoneda = 1
        Case "840"
            nMoneda = 2
        Case "000"
            nMoneda = 0
    End Select
    'Códigos de operación para InterCMACs
    Select Case Mid(gPRCODE, 1, 2)
        Case "01" 'Retiro
            'sOpeCod = gPITCaptacRetiro
            sOpeCodComision = "208022" 'comision
            sOpeCodITF = "261506" 'ITF Comision
            sOpeCodComisionITF = "208024"
            sOpeExtorno = "279101"
            sOpeExtornoComision = "279101"
        Case "20" 'Depósito
            'sOpeCod = gPITCaptaDeposito
            sOpeCodComision = "208022" 'comision
            sOpeCodITF = "261506" 'ITF Comision
            sOpeCodComisionITF = "208024"
            sOpeExtorno = "279102"
            sOpeExtornoComision = "279102"
        Case "50"
            'sOpeCod = gPITColocPagoCredito
            sOpeCodComision = "208022" 'comision
            sOpeCodITF = "105004" 'ITF Comision
            sOpeCodComisionITF = "208024"
    End Select
    
    Set rsParametros = obtenerParametros()
    
    While Not (rsParametros.EOF Or rsParametros.BOF)
        Select Case rsParametros!nParametroId
            Case 1000
                gnMontoMinRetMN = rsParametros!nValor
            Case 1001
                gnMontoMaxRetMN = rsParametros!nValor
            Case 1002
                gnMontoMinRetME = rsParametros!nValor
            Case 1003
                gnMontoMaxRetME = rsParametros!nValor
            Case 1004
                gnMontoMinRetMNReqDNI = rsParametros!nValor
            Case 1005
                gnMontoMinRetMEReqDNI = rsParametros!nValor
            Case 1006
                gnMontoMaxOpeMNxDia = rsParametros!nValor
            Case 1007
                gnMontoMaxOpeMExDia = rsParametros!nValor
            Case 1008
                gnMontoMaxOpeMNxMes = rsParametros!nValor
            Case 1009
                gnMontoMaxOpeMExMes = rsParametros!nValor
            Case 1010
                gnNumeroMaxOpeXDia = rsParametros!nValor
            Case 1011
                gnNumeroMaxOpeXMes = rsParametros!nValor
            Case 1012
                gnMontoMinDepMN = rsParametros!nValor
            Case 1013
                gnMontoMaxDepMN = rsParametros!nValor
            Case 1014
                gnMontoMinDepME = rsParametros!nValor
            Case 1015
                gnMontoMaxDepME = rsParametros!nValor
        End Select
        rsParametros.MoveNext
    Wend

End Sub
Public Function LogONOFF() As String

    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Transaccion ECO: " & sIDTrama, "")
    Call RegistrarTrama(gTRACE, pINXml, 1)
    pOUTXml = GeneraXMLSalida("00", , sHora)
    LogONOFF = pOUTXml
End Function


Public Function TXRetiro() As String
Dim lsCadResp As String
Dim lsCadAmount As String
        
    'If nOFFHost = 0 Then
    '    nResultado = PIT_RetiroOFFHost(dFecSis, sCtaCod, nMontoTran, nMontoITF, nMontoComision, nMontoComisionITF, nMoneda, _
            sOpeCod, sOpeCodComision, sOpeCodITF, sOpeCodComisionITF, nTipoCambio, pINXml, 0, gsPAN, sHora, sMesDia)
    'Else
    
    'MsgBox "Monto" & CStr(nMontoTran)
    
        nResultado = PITRetiro(dFecSis, sCtaCod, nMontoTran, nMontoITF, nMontoComision, nMontoComisionITF, nMoneda, _
            sOpeCod, sOpeCodComision, sOpeCodITF, sOpeCodComisionITF, nTipoCambio, pINXml, 0, gsPAN, sHora, sMesDia, nMontoEquiv, gsMovNro)
    'End If
    
    'MsgBox "Paso retiro"

    lnMovNro = nResultado
        
    lsCadAmount = "1001" & gCUR_CODE & "C" & Right("000000000000" & Replace(Trim(Format(RecuperaSaldoDisp(sCtaCod), "###0.00")), ".", ""), 12)
    lsCadAmount = lsCadAmount & "1002" & gCUR_CODE & "C" & Right("000000000000" & Replace(Trim(Format(RecuperaSaldoDisp(sCtaCod), "###0.00")), ".", ""), 12)
    lsCadResp = ""
    nMontoEquiv = CDbl(Format(nTipoCambio, "###0.000"))
    lsCadResp = Right("000000000000" & Replace(Format(nMontoEquiv, "#000.00"), ".", ""), 12)
    pOUTXml = GeneraXMLSalida(IIf(lnMovNro > 0, "00", "89"), , gnTramaId, sCtaCod, , lsCadAmount, lsCadResp)
    
    TXRetiro = pOUTXml
    
    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Registro Retiro", "", gnTramaId)
End Function

Public Function TXDeposito() As String
Dim lsCadResp As String
Dim lsCadAmount As String

    'If nOFFHost = 0 Then
        nResultado = PITDeposito(dFecSis, sCtaCod, nMontoTran, nMontoITF, nMontoComision, nMontoComisionITF, nMoneda, _
            sOpeCod, sOpeCodComision, sOpeCodITF, sOpeCodComisionITF, nTipoCambio, gnTramaId, gsPAN, sHora, sMesDia, sUserATM, gsMovNro)
    'Else
    '    nResultado = DepositoOffHost(dFecSis, sCtaCod, nMontoTran, nMontoITF, nMontoComision, nMontoComisionITF, nMoneda, _
    '        sOpeCod, sOpeCodComision, sOpeCodITF, sOpeCodComisionITF, nTipoCambio, pINXml, 0, gsPAN, sHora, sMesDia, sUserATM)
    'End If
        
    lnMovNro = nResultado
    
    lsCadAmount = "1001" & gCUR_CODE & "C" & Right("000000000000" & Replace(Trim(Format(RecuperaSaldoDisp(sCtaCod), "###0.00")), ".", ""), 12)
    lsCadAmount = lsCadAmount & "1002" & gCUR_CODE & "C" & Right("000000000000" & Replace(Trim(Format(RecuperaSaldoDisp(sCtaCod), "###0.00")), ".", ""), 12)
    nMontoEquiv = CDbl(Format(nTipoCambio, "#,000.00"))
    lsCadResp = Right("000000000000" & Replace(Format(nMontoEquiv, "###0.00"), ".", ""), 12)
    
    pOUTXml = GeneraXMLSalida(IIf(nResultado > 0, "00", "89"), , gnTramaId, sCtaCod, , lsCadAmount, lsCadResp)

    TXDeposito = pOUTXml
    
    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Registro Deposito", "", gnTramaId)
End Function

Public Function TXPagoCredito() As String
        
    TXPagoCredito = PagoCredito(sCtaCod, nMontoTran, nMoneda, dFecSis, lnMovNro)
      
    pOUTXml = GeneraXMLSalida(IIf(lnMovNro > 0, "00", "89"), , gnTramaId, sCtaCod, , nMontoTran, TXPagoCredito, "0215")
    
    TXPagoCredito = pOUTXml
    
    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Registro Pago de Credito", "", gnTramaId)
End Function

Public Function TXConsultaCuentaAhorro() As String
Dim nSaldoCtaDisp As Double
Dim nSaldoCtaTot As Double
Dim sCadRespSalCta As String

    nResultado = 1
    lnMovNro = nResultado

    Call RecuperaSaldoDeCuenta(sCtaCod, nSaldoCtaDisp, nSaldoCtaTot)

    sCadRespSalCta = "1001"
    If Mid(sCtaCod, 9, 1) = "1" Then
        sCadRespSalCta = sCadRespSalCta & "604"
    Else
        sCadRespSalCta = sCadRespSalCta & "840"
    End If
    
    sCadRespSalCta = sCadRespSalCta & "C"
    sCadRespSalCta = sCadRespSalCta & Right("000000000000" & Replace(Format(nSaldoCtaTot, "#0.00"), ".", ""), 12)
    
    sCadRespSalCta = sCadRespSalCta & "1002"
    If Mid(sCtaCod, 9, 1) = "1" Then
        sCadRespSalCta = sCadRespSalCta & "604"
    Else
        sCadRespSalCta = sCadRespSalCta & "840"
    End If
    sCadRespSalCta = sCadRespSalCta & "C"
    sCadRespSalCta = sCadRespSalCta & Right("000000000000" & Replace(Format(nSaldoCtaDisp, "#0.00"), ".", ""), 12)
                
    pOUTXml = GeneraXMLSalida("00", , sHora, sCtaCod, , sCadRespSalCta)
    TXConsultaCuentaAhorro = pOUTXml
    
    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Realizo Consulta Cuenta Ahorro", "", gnTramaId)
End Function

Public Function TXConsultaCuentasAhorro() As String


    'If nOFFHost = 0 Then
        nResultado = Consulta(dFecSis, nMoneda, sOpeCod, 0, gsPAN, sHora, sMesDia)
    'Else
    '    nResultado = ConsultaOFFHost(dFecSis, nMoneda, sOpeCod, 0, gsPAN, sHora, sMesDia)
    'End If

    lnMovNro = nResultado
              
    pOUTXml = GeneraXMLSalida(IIf(lnMovNro > 0, "00", "89"), , gnTramaId, sCtaCod, , 0, RecuperaCuentasAho(gsPAN), "0215")
    
    TXConsultaCuentasAhorro = pOUTXml
    
    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Realizo Consulta Cuentas Ahorro", "", gnTramaId)
End Function

Public Function TXConsultaCuentasCredito() As String
    
    'If nOFFHost = 0 Then
        nResultado = Consulta(dFecSis, nMoneda, sOpeCod, 0, gsPAN, sHora, sMesDia)
    'Else
    '    nResultado = ConsultaOFFHost(dFecSis, nMoneda, sOpeCod, 0, gsPAN, sHora, sMesDia)
    'End If
    
    'Call RegistraPITMov(nResultado, 0, -1, gsPAN, gsDNI, 0, sHora, sMesDia, nMoneda, gnTramaId)
    
    lnMovNro = nResultado
       
    pOUTXml = GeneraXMLSalida(IIf(lnMovNro > 0, "00", "89"), , gnTramaId, sCtaCod, , 0, RecuperaCuentasCred(gsDNI, dFecSis), "0215")
    
    TXConsultaCuentasCredito = pOUTXml
    
    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Realizo Consulta Cuentas Credito", "", gnTramaId)
End Function

Public Function TXConsultaMovimientosAhorro() As String
    
    'If nOFFHost = 0 Then
        nResultado = Consulta(dFecSis, nMoneda, sOpeCod, 0, gsPAN, sHora, sMesDia)
    'Else
    '    nResultado = ConsultaOFFHost(dFecSis, nMoneda, sOpeCod, 0, gsPAN, sHora, sMesDia)
    'End If

    'Call RegistraPITMov(nResultado, 0, -1, gsPAN, gsDNI, 0, sHora, sMesDia, nMoneda, gnTramaId)
    
    lnMovNro = nResultado
    
    pOUTXml = GeneraXMLSalida(IIf(lnMovNro > 0, "00", "89"), , gnTramaId, sCtaCod, , 0, RecuperaMovimDeCuenta(sCtaCod, nMoneda), "0215")

    TXConsultaMovimientosAhorro = pOUTXml
    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Realizo Consulta Movimientos Ahorro", "", gnTramaId)
End Function

Public Function TXExtornoTransaccion() As String
Dim lnMovNro As Long
Dim lsTramaXmlEnvio As String, lsTramaXmlRecep As String
Dim lRsMovExt As ADODB.Recordset
Dim lsOpecod As String
Dim loConec As New DConecta, lsSql As String
Dim lRsMov As ADODB.Recordset

    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Inicio Operacion de Extorno", "", gnTramaId)

    Set lRsMovExt = RecuperaMovimientoInterCajaParaExtorno(gnTramaIDExt, "")

    
    If Not lRsMovExt.EOF And Not lRsMovExt.BOF Then
        lsTramaXmlEnvio = lRsMovExt("cTramaEnvio")
        lsTramaXmlRecep = lRsMovExt("cTramaRecep")
        lnMovNro = lRsMovExt("nMovNro")
        lsOpecod = lRsMovExt("cOpeCod")
    Else 'No existe la operación a extornar o ya se encuentra extornado
        Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Recupera Movimiento InterCaja Para Extorno", "Error - no existe movimiento a extornar u operacio se encuentra extornada", gnTramaId)
        pOUTXml = GeneraXMLSalida("12", gnTramaId)
        Exit Function
    End If

    gMESSAGE_TYPE = "0420"
    gTRACE = RecuperaValorXML(lsTramaXmlRecep, "TRACE")
    gPRCODE = RecuperaValorXML(lsTramaXmlRecep, "PRCODE")
    gsPAN = RecuperaValorXML(lsTramaXmlRecep, "PAN")
    gTIME_LOCAL = RecuperaValorXML(lsTramaXmlRecep, "TIME_LOCAL")
    gDATE_LOCAL = RecuperaValorXML(lsTramaXmlRecep, "DATE_LOCAL")
    gTERMINAL_ID = RecuperaValorXML(lsTramaXmlRecep, "TERMINAL_ID")
    gCARD_ACCEPTOR = RecuperaValorXML(lsTramaXmlRecep, "CARD_ACCEPTOR")
    gACQ_INST = RecuperaValorXML(lsTramaXmlRecep, "ACQ_INST")
    gPOS_COND_CODE = RecuperaValorXML(lsTramaXmlRecep, "POS_COND_CODE")
    gTXN_AMOUNT = RecuperaValorXML(lsTramaXmlRecep, "TXN_AMOUNT")
    gCUR_CODE = RecuperaValorXML(lsTramaXmlRecep, "CUR_CODE")
    gDATE_EXP = RecuperaValorXML(lsTramaXmlRecep, "DATE_EXP")
    gCARD_LOCATION = RecuperaValorXML(lsTramaXmlRecep, "CARD_LOCATION")
    gACCT_1 = RecuperaValorXML(lsTramaXmlRecep, "ACCT_1")
    gACCT_2 = RecuperaValorXML(lsTramaXmlRecep, "ACCT_2")
    
    gsCtaCod = Right(gACCT_1, 18)
    
    pINXml = Replace(pINXml, "<MESSAGE_TYPE = 0420 />", "<MESSAGE_TYPE = 0200 />")
    
    If Left(gPRCODE, 2) = "01" Or Left(gPRCODE, 2) = "20" Then
        If Not ValidaExtorno(gsCtaCod, lnMovNro, 0) Then
            pOUTXml = GeneraXMLSalida("89", gnTramaId)
            Exit Function
        End If
    Else
        If Not ValidaExtorno(gsCtaCod, lnMovNro, 1) Then
            pOUTXml = GeneraXMLSalida("89", gnTramaId)
            Exit Function
        End If
    End If
   
    loConec.AbreConexion
    
    Select Case Left(gPRCODE, 2)
        Case "01" 'Extorno de retiro de cuenta de ahorro
            
            lsSql = "Exec PIT_stp_ins_ExtornoRetiro " & lnMovNro & ",'" & lsOpecod & "','" _
                    & sOpeExtorno & "','" & sOpeCodITF & "','" & sOpeCodComision & "','" _
                    & sOpeExtornoComision & "','" & sOpeCodComisionITF & "','" & gsMovNro & "'"
            
            Set lRsMov = loConec.ConexionActiva.Execute(lsSql)
    
            If Not lRsMov.BOF And Not lRsMov.EOF Then
                nResultado = lRsMov("nResultado")
            End If
            
        Case "20" 'Extorno de depósito de cuenta de ahorro
            
            lsSql = "Exec PIT_stp_ins_ExtornoDeposito '" & gsCtaCod & "'," & lnMovNro & ",'" _
                    & lsOpecod & "','" & sOpeExtorno & "','" & sOpeCodITF & "','" & sOpeCodComision & "','" _
                    & sOpeExtornoComision & "','" & sOpeCodComisionITF & "','" & gsMovNro & "'"
            
            Set lRsMov = loConec.ConexionActiva.Execute(lsSql)
            
            If Not lRsMov.BOF And Not lRsMov.EOF Then
                nResultado = lRsMov("nResultado")
            End If
            
        Case "50" 'Extorno de pago de cuenta de crédito
        
            Call ExtornarPagoCredito(lnMovNro, gsCtaCod, Val(gTXN_AMOUNT), dFecSis, "CMAC", "01", nResultado)
            
    End Select

    If nResultado > 0 Then
        Call RegistraPITMov(nResultado, 0, -1, gsPAN, gsDNI, nMontoTran, sHora, sMesDia, nMoneda, gnTramaId, 4) 'pnEstado=4: Registro de extorno
    End If
    
    loConec.CierraConexion
    
    Set loConec = Nothing

    pOUTXml = GeneraXMLSalida(IIf(nResultado > 0, "00", "89"), , gnTramaId)
            
    TXExtornoTransaccion = pOUTXml
End Function

Public Function RetenerTarjetaPorPosibleFraude(ByVal psNumTarj As String) As Integer
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim loConec As New DConecta

    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pPAN", adVarChar, adParamInput, 20, psNumTarj)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnRes", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RetenerTarjetaPorPosibleFraude"
    
    Cmd.Execute
    
    RetenerTarjetaPorPosibleFraude = Cmd.Parameters(1).Value
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    
    Set Cmd = Nothing
    Set Prm = Nothing
            
End Function
 
Public Function RecuperaPVV(ByVal pPAN As String) As String
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim loConec As New DConecta

    Set Cmd = New ADODB.Command
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cNumTarj", adVarChar, adParamInput, 50, pPAN)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psPVV", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaPVV"
    
    Cmd.Execute
        
    RecuperaPVV = Cmd.Parameters(1).Value
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    
    Set Cmd = Nothing
    Set Prm = Nothing
    
End Function
 
Public Sub RegistraOperacionLimitesCajeroPOS(ByVal pdFecha As Date, ByVal psCodTranCaj As String, _
        ByVal psNumTarj As String, ByVal pnMonto As Double, ByVal pnMoneda As Integer)
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim loConec As New DConecta
 
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFecha", adDBDate, adParamInput, , pdFecha)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cCodTranCaj", adVarChar, adParamInput, 50, psCodTranCaj)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cNumTarj", adVarChar, adParamInput, 50, psNumTarj)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nMonto", adDouble, adParamInput, , pnMonto)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMoneda", adInteger, adParamInput, , pnMoneda)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "RegistraOperacionLimitesCajeroPOS"
    
    Cmd.Execute
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
            
    Set Cmd = Nothing
    Set Prm = Nothing
        
End Sub
 
Public Sub ActualizaPVV(ByVal pssPVV As String, ByVal psPAN As String)
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim loConec As New DConecta

    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psPan", adVarChar, adParamInput, 16, psPAN)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psPVV", adVarChar, adParamInput, 10, pssPVV)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_ActualizaPVV"
    
    Cmd.Execute
    
    'Call CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    
    Set Cmd = Nothing
    Set Prm = Nothing
    
End Sub
 
 '***********************************************
 'Modificado por NSSE 07/06/2008
 '***********************************************
 
Public Function Retiro(ByVal pdFecSis As Date, ByVal psCtaCod As String, ByVal pnMontoTran As Double, _
    ByVal pnMontoITF As Double, ByVal pnMontoComision As Double, ByVal pnMontoComisionITF As Double, _
    ByVal pnMoneda As Integer, ByVal psOpeCod As String, ByVal psOpeCodComision As String, ByVal psOpeCodITF As String, _
    ByVal psOpeCodComisionITF As String, ByVal pnTipoCambio As Double, ByVal psIDTrama As String, _
    ByVal pnoffHost As Integer, ByVal pPAN As String, sHora As String, ByVal sMesDia As String, ByRef nMontoEquiv As Double) As Integer
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim nResultado As Integer
Dim loConec As New DConecta
        
    Set Cmd = New ADODB.Command
    '0
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pdFecha", adDBDate, adParamInput, , pdFecSis)
    Cmd.Parameters.Append Prm
    '1
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 18, psCtaCod)
    Cmd.Parameters.Append Prm
    '2
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMonto", adDouble, adParamInput, , pnMontoTran)
    Cmd.Parameters.Append Prm
    
    '3
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMontoITF", adDouble, adParamInput, , pnMontoITF)
    Cmd.Parameters.Append Prm
    
    '4
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMontoComision", adDouble, adParamInput, , pnMontoComision)
    Cmd.Parameters.Append Prm
    '5
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMontoComisionITF", adDouble, adParamInput, , pnMontoComisionITF)
    Cmd.Parameters.Append Prm
    '6
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMoneda", adSmallInt, adParamInput, , pnMoneda)
    Cmd.Parameters.Append Prm
    
    '7
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCod", adVarChar, adParamInput, 6, psOpeCod)
    Cmd.Parameters.Append Prm
    '8
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCodComision", adVarChar, adParamInput, 6, psOpeCodComision)
    Cmd.Parameters.Append Prm
    '9
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCodITF", adVarChar, adParamInput, 6, psOpeCodITF)
    Cmd.Parameters.Append Prm
    
    '10
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("sOpeCodComisionITF", adVarChar, adParamInput, 6, psOpeCodComisionITF)
    Cmd.Parameters.Append Prm
    
    '11
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnTipoCambio", adDouble, adParamInput, 6, pnTipoCambio)
    Cmd.Parameters.Append Prm
    '12
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psIDTrama", adVarChar, adParamInput, 5000, psIDTrama)
    Cmd.Parameters.Append Prm
    '13
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pPAN", adVarChar, adParamInput, 16, pPAN)
    Cmd.Parameters.Append Prm
    '14
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psHora", adVarChar, adParamInput, 20, sHora)
    Cmd.Parameters.Append Prm
    '15
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psMesDia", adVarChar, adParamInput, 20, sMesDia)
    Cmd.Parameters.Append Prm
     
    '16
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnResultado", adSmallInt, adParamOutput, , nResultado)
    Cmd.Parameters.Append Prm
                       
    '17
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMovNroOffHost", adSmallInt, adParamInput, , pnoffHost)
    Cmd.Parameters.Append Prm
                            
    '18
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psUser", adChar, adParamInput, 4, sUserATM)
    Cmd.Parameters.Append Prm
    
     '19
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMontoEquiv", adDouble, adParamOutput)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_Retiro"
    Cmd.Execute
    'Call CerrarConexion
    
    Retiro = Cmd.Parameters(16).Value
    nMontoEquiv = Cmd.Parameters(19).Value
    
    loConec.CierraConexion
    Set loConec = Nothing

End Function

 '***********************************************
 'Modificado por NSSE 15/05/2008 OFFHOST
 '***********************************************
 
Public Function RetiroOFFHost(ByVal pdFecSis As Date, ByVal psCtaCod As String, ByVal pnMontoTran As Double, _
    ByVal pnMontoITF As Double, ByVal pnMontoComision As Double, ByVal pnMontoComisionITF As Double, _
    ByVal pnMoneda As Integer, ByVal psOpeCod As String, ByVal psOpeCodComision As String, ByVal psOpeCodITF As String, _
    ByVal psOpeCodComisionITF As String, ByVal pnTipoCambio As Double, ByVal psIDTrama As String, _
    ByVal pnoffHost As Integer, ByVal pPAN As String, ByVal sHora As String, ByVal sMesDia As String) As Integer
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim nResultado As Integer
Dim loConec As New DConecta
        
    Set Cmd = New ADODB.Command
    '0
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pdFecha", adDBDate, adParamInput, , pdFecSis)
    Cmd.Parameters.Append Prm
    '1
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 18, psCtaCod)
    Cmd.Parameters.Append Prm
    '2
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMonto", adDouble, adParamInput, , pnMontoTran)
    Cmd.Parameters.Append Prm
    
    '3
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMontoITF", adDouble, adParamInput, , pnMontoITF)
    Cmd.Parameters.Append Prm
    
    '4
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMontoComision", adDouble, adParamInput, , pnMontoComision)
    Cmd.Parameters.Append Prm
    '5
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMontoComisionITF", adDouble, adParamInput, , pnMontoComisionITF)
    Cmd.Parameters.Append Prm
    '6
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMoneda", adSmallInt, adParamInput, , pnMoneda)
    Cmd.Parameters.Append Prm
    
    '7
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCod", adVarChar, adParamInput, 6, psOpeCod)
    Cmd.Parameters.Append Prm
    '8
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCodComision", adVarChar, adParamInput, 6, psOpeCodComision)
    Cmd.Parameters.Append Prm
    '9
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCodITF", adVarChar, adParamInput, 6, psOpeCodITF)
    Cmd.Parameters.Append Prm
    
    '10
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("sOpeCodComisionITF", adVarChar, adParamInput, 6, psOpeCodComisionITF)
    Cmd.Parameters.Append Prm
    
    '11
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnTipoCambio", adDouble, adParamInput, 6, pnTipoCambio)
    Cmd.Parameters.Append Prm
    '12
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psIDTrama", adVarChar, adParamInput, 5000, psIDTrama)
    Cmd.Parameters.Append Prm
    
    '13
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pPAN", adVarChar, adParamInput, 16, pPAN)
    Cmd.Parameters.Append Prm
    '14
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psHora", adVarChar, adParamInput, 6, sHora)
    Cmd.Parameters.Append Prm
    '15
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psMesDia", adVarChar, adParamInput, 4, sMesDia)
    Cmd.Parameters.Append Prm
     
    '16
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnResultado", adSmallInt, adParamOutput, , nResultado)
    Cmd.Parameters.Append Prm
    
    '17
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psUser", adChar, adParamInput, 4, sUserATM)
    Cmd.Parameters.Append Prm
                            
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_Retiro_OffHost"
    Cmd.Execute
    'Call CerrarConexion
    
    RetiroOFFHost = Cmd.Parameters(16).Value

    loConec.CierraConexion
    Set loConec = Nothing
    
End Function

Public Function PIT_RetiroOFFHost(ByVal pdFecSis As Date, ByVal psCtaCod As String, ByVal pnMontoTran As Double, _
    ByVal pnMontoITF As Double, ByVal pnMontoComision As Double, ByVal pnMontoComisionITF As Double, _
    ByVal pnMoneda As Integer, ByVal psOpeCod As String, ByVal psOpeCodComision As String, ByVal psOpeCodITF As String, _
    ByVal psOpeCodComisionITF As String, ByVal pnTipoCambio As Double, ByVal psIDTrama As String, _
    ByVal pnoffHost As Integer, ByVal pPAN As String, ByVal sHora As String, ByVal sMesDia As String) As Long
'Dim nResultado As Integer
Dim loConec As New DConecta, lsSql As String
Dim lRsMov As ADODB.Recordset
        

    loConec.AbreConexion
    
    lsSql = " exec PIT_stp_sel_Retiro_OffHost '" & Format(pdFecSis, "YYYYMMDD") & "','" & psCtaCod & "'," & _
                pnMontoTran & "," & pnMontoITF & "," & pnMontoComision & "," & pnMontoComisionITF & "," & _
                pnMoneda & ",'" & psOpeCod & "','" & psOpeCodComision & "','" & psOpeCodITF & "','" & _
                psOpeCodComisionITF & "'," & pnTipoCambio & ",'" & pPAN & "','" & sHora & "','" & sMesDia & "','" & _
                "CMAC" & "'," & gnTramaId
    
    Set lRsMov = loConec.ConexionActiva.Execute(lsSql)
    
    If Not lRsMov.BOF And Not lRsMov.EOF Then
        nResultado = lRsMov("nResultado")
    End If

    PIT_RetiroOFFHost = nResultado

    loConec.CierraConexion
    
    Set loConec = Nothing
    
End Function

 
 '***********************************************
 'Modificado por NSSE 07/06/2008
 '***********************************************
 
Public Function PITRetiro(ByVal pdFecSis As Date, ByVal psCtaCod As String, ByVal pnMontoTran As Double, _
    ByVal pnMontoITF As Double, ByVal pnMontoComision As Double, ByVal pnMontoComisionITF As Double, _
    ByVal pnMoneda As Integer, ByVal psOpeCod As String, ByVal psOpeCodComision As String, ByVal psOpeCodITF As String, _
    ByVal psOpeCodComisionITF As String, ByVal pnTipoCambio As Double, ByVal psIDTrama As String, _
    ByVal pnoffHost As Integer, ByVal psPAN As String, psHora As String, ByVal psMesDia As String, ByRef pnMontoEquiv As Double, _
    ByVal psMovNro As String) As Long
Dim nResultado As Long
Dim loConec As New DConecta, lsSql As String
Dim lRsMov As ADODB.Recordset
            
    lsSql = "Exec PIT_stp_ins_Retiro '" & Format(pdFecSis, "yyyy-mm-dd hh:mm:ss") & "','" & psCtaCod & "'," & pnMontoTran & "," & pnMontoITF & "," & _
            pnMontoComision & "," & pnMontoComisionITF & "," & pnMoneda & ",'" & psOpeCod & "','" & psOpeCodComision & "','" & _
            psOpeCodITF & "','" & psOpeCodComisionITF & "'," & pnTipoCambio & ",'" & psPAN & "','" & psHora & "','" & psMesDia & "'," & _
            pnoffHost & ",'" & "CMAC" & "','" & psMovNro & "'"
    
    loConec.AbreConexion
    
    Set lRsMov = loConec.ConexionActiva.Execute(lsSql)

    If Not lRsMov.BOF And Not lRsMov.EOF Then
        nResultado = lRsMov("nResultado")
    End If
    
    If nResultado > 0 Then
        Call RegistraPITMov(nResultado, 0, -1, psPAN, gsDNI, pnMontoTran, psHora, psMesDia, pnMoneda, gnTramaId, 0)
    End If

    PITRetiro = nResultado

    loConec.CierraConexion
    
    Set loConec = Nothing

End Function

Public Function PITDeposito(ByVal pdFecSis As Date, ByVal psCtaCod As String, ByVal pnMontoTran As Double, _
    ByVal pnMontoITF As Double, ByVal pnMontoComision As Double, ByVal pnMontoComisionITF As Double, _
    ByVal pnMoneda As Integer, ByVal psOpeCod As String, ByVal psOpeCodComision As String, ByVal psOpeCodITF As String, _
    ByVal psOpeCodComisionITF As String, ByVal pnTipoCambio As Double, ByVal pnTramaId As Long, _
    ByVal psPAN As String, psHora As String, ByVal psMesDia As String, psUserATM As String, psMovNro As String) As Long
'Dim nResultado As Integer
Dim loConec As New DConecta, lsSql As String
Dim lRsMov As ADODB.Recordset
                
    
    lsSql = "Exec PIT_stp_ins_Deposito '" & Format(pdFecSis, "yyyy-mm-dd hh:mm:ss") & "','" & psCtaCod & "'," & pnMontoTran & "," & pnMontoITF & "," & _
                pnMontoComision & "," & pnMontoComisionITF & "," & pnMoneda & ",'" & psOpeCod & "','" & _
                psOpeCodComision & "','" & psOpeCodITF & "','" & psOpeCodComisionITF & "'," & pnTipoCambio & "," & 0 & ",'" & "CMAC" & "','" & psMovNro & "'"
    
    loConec.AbreConexion
    
    'MsgBox lsSql
    
    Set lRsMov = loConec.ConexionActiva.Execute(lsSql)
    
    If Not lRsMov.BOF And Not lRsMov.EOF Then
        nResultado = lRsMov("nResultado")
    End If
        

    If nResultado > 0 Then
        Call RegistraPITMov(nResultado, 0, -1, psPAN, gsDNI, pnMontoTran, psHora, psMesDia, pnMoneda, pnTramaId, 0)
    End If

    PITDeposito = nResultado

    loConec.CierraConexion
    
    Set loConec = Nothing

End Function
Public Function DepositoOffHost(ByVal pdFecSis As Date, ByVal psCtaCod As String, ByVal pnMontoTran As Double, _
    ByVal pnMontoITF As Double, ByVal pnMontoComision As Double, ByVal pnMontoComisionITF As Double, _
    ByVal pnMoneda As Integer, ByVal psOpeCod As String, ByVal psOpeCodComision As String, ByVal psOpeCodITF As String, _
    ByVal psOpeCodComisionITF As String, ByVal pnTipoCambio As Double, ByVal psIDTrama As String, _
    ByVal pnResultado As Integer, ByVal pPAN As String, sHora As String, ByVal sMesDia As String, psUserATM As String) As Integer

    Dim Cmd As ADODB.Command
    Dim Prm As ADODB.Parameter
    Dim nResultado As Integer
    Dim loConec As New DConecta

    Set Cmd = New ADODB.Command
    '0
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pdFecha", adDBDate, adParamInput, , pdFecSis)
    Cmd.Parameters.Append Prm
    '1
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 18, psCtaCod)
    Cmd.Parameters.Append Prm
    '2
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMonto", adDouble, adParamInput, , pnMontoTran)
    Cmd.Parameters.Append Prm
    
    '*******************************************************************
    'MODIFICADO NSSE 16/06/2008
    '*******************************************************************
    '3
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMontoITF", adDouble, adParamInput, , pnMontoITF)
    'Set Prm = Cmd.CreateParameter("@pnMontoITF", adDouble, adParamInput, , 0)
    Cmd.Parameters.Append Prm
    
    '4
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMontoComision", adDouble, adParamInput, , pnMontoComision)
    Cmd.Parameters.Append Prm
    '*******************************************************************
    'MODIFICADO NSSE 16/06/2008
    '*******************************************************************
    '5
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMontoComisionITF", adDouble, adParamInput, , pnMontoComisionITF)
    'Set Prm = Cmd.CreateParameter("@pnMontoComisionITF", adDouble, adParamInput, , 0)
    Cmd.Parameters.Append Prm
    '6
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMoneda", adSmallInt, adParamInput, , pnMoneda)
    Cmd.Parameters.Append Prm
    
    '7
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCod", adVarChar, adParamInput, 6, psOpeCod)
    Cmd.Parameters.Append Prm
    '8
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCodComision", adVarChar, adParamInput, 6, psOpeCodComision)
    Cmd.Parameters.Append Prm
    '9
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCodITF", adVarChar, adParamInput, 6, psOpeCodITF)
    Cmd.Parameters.Append Prm
    
    '10
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("sOpeCodComisionITF", adVarChar, adParamInput, 6, psOpeCodComisionITF)
    Cmd.Parameters.Append Prm
    
    '11
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnTipoCambio", adDouble, adParamInput, 6, pnTipoCambio)
    Cmd.Parameters.Append Prm
    '12
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psIDTrama", adVarChar, adParamInput, 5000, psIDTrama)
    Cmd.Parameters.Append Prm
    '13
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnResultado", adSmallInt, adParamOutput, , pnResultado)
    Cmd.Parameters.Append Prm
                       
    '14
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psUser", adChar, adParamInput, 4, psUserATM)
    Cmd.Parameters.Append Prm
                             
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_Desposito_OffHost"
    Cmd.Execute
    'Call CerrarConexion
    loConec.CierraConexion
    DepositoOffHost = Cmd.Parameters(13).Value
    Set Cmd = Nothing
End Function
 '***********************************************
 'Modificado por NSSE 07/06/2008
 '***********************************************
 
Public Function Consulta(ByVal pdFecSis As Date, ByVal pnMoneda As Integer, ByVal psOpeCod As String, _
                         ByVal pnoffHost As Integer, ByVal pPAN As String, sHora As String, _
                         ByVal sMesDia As String) As Long
    
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim loConec As New DConecta, lsSql As String
Dim lRsMov As ADODB.Recordset
    
    loConec.AbreConexion

    lsSql = "exec PIT_stp_ins_CONSULTA '" & Format(pdFecSis, "YYYYMMDD") & "'," & pnMoneda & ",'" & psOpeCod & "','" & pPAN & "','" & _
                sHora & "','" & sMesDia & "'," & pnoffHost & ",'" & "CMAC" & "','" & gsMovNro & "'"
    
    Set lRsMov = loConec.ConexionActiva.Execute(lsSql)

    If Not lRsMov.BOF And Not lRsMov.EOF Then
        nResultado = lRsMov("nResultado")
    End If

    Call RegistraPITMov(nResultado, 0, -1, gsPAN, gsDNI, 0, sHora, sMesDia, pnMoneda, gnTramaId, 0)

    Consulta = nResultado
    
    loConec.CierraConexion
    Set loConec = Nothing

End Function
Public Function ConsultaOFFHost(ByVal pdFecSis As Date, ByVal pnMoneda As Integer, _
                                ByVal psOpeCod As String, ByVal pnoffHost As Integer, _
                                ByVal pPAN As String, sHora As String, ByVal sMesDia As String) As Integer

Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim nResultado As Integer
Dim loConec As New DConecta
        
    Set Cmd = New ADODB.Command
    '0
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pdFecha", adDBDate, adParamInput, , pdFecSis)
    Cmd.Parameters.Append Prm
    '1
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMoneda", adSmallInt, adParamInput, , pnMoneda)
    Cmd.Parameters.Append Prm
    '2
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCod", adVarChar, adParamInput, 6, psOpeCod)
    Cmd.Parameters.Append Prm
    '4
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pPAN", adVarChar, adParamInput, 16, pPAN)
    Cmd.Parameters.Append Prm
    '5
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psHora", adVarChar, adParamInput, 6, sHora)
    Cmd.Parameters.Append Prm
    '6
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psMesDia", adVarChar, adParamInput, 4, sMesDia)
    Cmd.Parameters.Append Prm
    '7
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnResultado", adSmallInt, adParamOutput, , nResultado)
    Cmd.Parameters.Append Prm
    '8
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psUser", adChar, adParamInput, 4, sUserATM)
    Cmd.Parameters.Append Prm
                            
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "PIT_Consulta_OffHost"
    Cmd.Execute
    'Call CerrarConexion
    
    ConsultaOFFHost = Cmd.Parameters(7).Value

    loConec.CierraConexion
    Set loConec = Nothing
    
End Function
Public Function CuentaExoneradaITF(ByVal psCtaCod As String) As Boolean
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim loConec As New DConecta
    
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCodCta", adVarChar, adParamInput, 50, psCtaCod)
    Cmd.Parameters.Append Prm
        
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnValor", adDouble, adParamOutput, , psCtaCod)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaCtaExoITF"
    Cmd.Execute
   
    CuentaExoneradaITF = IIf(Cmd.Parameters(1).Value = 1, True, False)
    
    Set Cmd = Nothing
    Set Prm = Nothing
    
    loConec.CierraConexion
    Set loConec = Nothing

End Function

Public Function CalculaITF(ByVal pnMonto As Double, ByVal psCtaCod As String) As Double
Dim lnValor As Double
lnValor = pnMonto
    
    If Not CuentaExoneradaITF(psCtaCod) And Mid(psCtaCod, 6, 3) <> "234" Then
        
        lnValor = pnMonto * ObtieneITF
        
        Dim aux As Double
        If InStr(1, CStr(lnValor), ".", vbTextCompare) > 0 Then
            aux = CDbl(CStr(Int(lnValor)) & "." & Mid(CStr(lnValor), InStr(1, CStr(lnValor), ".", vbTextCompare) + 1, 2))
        Else
            aux = CDbl(CStr(Int(lnValor)))
        End If
        lnValor = aux
   
        lnValor = fgTruncar(lnValor, 2)
        CalculaITF = lnValor
    Else
        CalculaITF = 0
    End If

End Function

Public Function fgTruncar(pnNumero As Double, pnDecimales As Integer) As String

    Dim i As Integer
    Dim sEnt As String
    Dim sDec As String
    Dim sNum As String
    Dim sPunto As String
    Dim sResul As String
    
    sNum = Trim(Str(pnNumero))
    sDec = ""
    sPunto = ""
    sEnt = ""
    For i = 1 To Len(Trim(sNum))
        If Mid(sNum, i, 1) = "." Then
            sPunto = "."
        Else
            If sPunto = "" Then
                sEnt = sEnt & Mid(sNum, i, 1)
            Else
                sDec = sDec & Mid(sNum, i, 1)
            End If
        End If
    Next i
    If sDec = "" Then
        sDec = "00"
    End If
    sResul = sEnt & "." & Left(sDec, 2)
    fgTruncar = sResul
    
End Function

Private Function cCargoPorOpe(pnCargo As Double, Optional pnCantCaracteres As Integer = 8) As String
    cCargoPorOpe = ";CARGO POR OPE.: " & Right(Space(pnCantCaracteres) & Format(pnCargo, "#,0.00"), pnCantCaracteres) & ";"
End Function

'*************************************
' Retiro para operaciones Intercajas *
'*************************************

Public Function RetiroInterCMAC(ByVal pdFecSis As Date, ByVal psCtaCod As String, ByVal pnMontoTran As Double, _
    ByVal pnMontoComision As Double, ByVal pnMoneda As Integer, ByVal psOpeCod As String, ByVal psOpeCodComision As String, _
    ByVal pnTipoCambio As Double, ByVal psIDTrama As String, ByVal pnoffHost As Integer, ByVal pPAN As String, _
    sHora As String, ByVal sMesDia As String, ByRef nMontoEquiv As Double) As Integer
    
    Dim Cmd As ADODB.Command
    Dim Prm As ADODB.Parameter
    Dim nResultado As Integer
    Dim loConec As New DConecta
        
    Set Cmd = New ADODB.Command
    '0
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pdFecha", adDBDate, adParamInput, , pdFecSis)
    Cmd.Parameters.Append Prm
    '1
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 18, psCtaCod)
    Cmd.Parameters.Append Prm
    '2
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMonto", adDouble, adParamInput, , pnMontoTran)
    Cmd.Parameters.Append Prm
    
    '3
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@pnMontoITF", adDouble, adParamInput, , pnMontoITF)
'    Cmd.Parameters.Append Prm
    
    '4
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMontoComision", adDouble, adParamInput, , pnMontoComision)
    Cmd.Parameters.Append Prm
    '5
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@pnMontoComisionITF", adDouble, adParamInput, , pnMontoComisionITF)
'    Cmd.Parameters.Append Prm
    '6
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMoneda", adSmallInt, adParamInput, , pnMoneda)
    Cmd.Parameters.Append Prm
    
    '7
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCod", adVarChar, adParamInput, 6, psOpeCod)
    Cmd.Parameters.Append Prm
    '8
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCodComision", adVarChar, adParamInput, 6, psOpeCodComision)
    Cmd.Parameters.Append Prm
    '9
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@psOpeCodITF", adVarChar, adParamInput, 6, psOpeCodITF)
'    Cmd.Parameters.Append Prm
    
    '10
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("sOpeCodComisionITF", adVarChar, adParamInput, 6, psOpeCodComisionITF)
'    Cmd.Parameters.Append Prm

    '11
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnTipoCambio", adDouble, adParamInput, 6, pnTipoCambio)
    Cmd.Parameters.Append Prm
    '12
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psIDTrama", adVarChar, adParamInput, 5000, psIDTrama)
    Cmd.Parameters.Append Prm
    '13
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pPAN", adVarChar, adParamInput, 16, pPAN)
    Cmd.Parameters.Append Prm
    '14
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psHora", adVarChar, adParamInput, 20, sHora)
    Cmd.Parameters.Append Prm
    '15
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psMesDia", adVarChar, adParamInput, 20, sMesDia)
    Cmd.Parameters.Append Prm
     
    '16
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnResultado", adSmallInt, adParamOutput, , nResultado)
    Cmd.Parameters.Append Prm
                       
    '17
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMovNroOffHost", adSmallInt, adParamInput, , pnoffHost)
    Cmd.Parameters.Append Prm
                            
    '18
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psUser", adChar, adParamInput, 4, sUserATM)
    Cmd.Parameters.Append Prm
    
     '19
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnMontoEquiv", adDouble, adParamOutput)
    Cmd.Parameters.Append Prm
    
    loConec.AbreConexion
    Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RetiroInterCMAC"
    Cmd.Execute
    'Call CerrarConexion
    
    RetiroInterCMAC = Cmd.Parameters(16).Value
    nMontoEquiv = Cmd.Parameters(19).Value
    
    loConec.CierraConexion
    Set loConec = Nothing

End Function

'Add by Gitu 30-07-2009
Public Function MatrizMontoAPagar(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim i As Integer
Dim J As Integer
    MatrizMontoAPagar = 0
    For i = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(i, 0)) Then
            For J = 3 To 9
                MatrizMontoAPagar = MatrizMontoAPagar + CDbl(MatCalend(i, J))
            Next J
            MatrizMontoAPagar = MatrizMontoAPagar + CDbl(MatCalend(i, 11))
        End If
    Next i
    If MatrizMontoAPagar = 0 Then
        For J = 3 To 9
            MatrizMontoAPagar = MatrizMontoAPagar + CDbl(MatCalend(0, J))
        Next J
        MatrizMontoAPagar = MatrizMontoAPagar + CDbl(MatCalend(0, 11))
    End If
End Function

'Add by gitu 30-07-2009
Function RecuperaMatrizCalendarioPendiente(ByVal psCtaCod As String, Optional ByVal pbParalelo As Boolean = False, _
                                           Optional ByVal pBrfa As Boolean, Optional ByRef pnNroCuotas As Integer) As Variant
Dim R As ADODB.Recordset
Dim RTmp As ADODB.Recordset
Dim RCred As ADODB.Recordset
Dim MatCalend() As String
Dim nMontoSaldo As Double
Dim nNumDiasfer As Integer
Dim dFecSisTmp As Date
Dim bFeriado As Boolean
Dim i As Integer
Dim nIntPend, nColocCalendCod, nIntFecha  As Double
Dim dVigencia As Date
Dim nSaldo As Double
Dim nDiasTranscurridos As Integer
Dim nTasaInt As Double
Dim nDiasAtraso As Integer

    'On Error GoTo ErrorRecuperaMatrizCalendarioPendiente

    dFecSisTmp = CDate(LeeConstSistema(16))

    Set R = RecuperaColocacCred(psCtaCod)
    If R.RecordCount > 0 Then
        nIntPend = R!nIntPend
        nColocCalendCod = R!nColocCalendCod
        nDiasAtraso = R!nDiasAtraso
        R.Close
    Else
        Set R = Nothing
        Exit Function
    End If
    
    'Para el caso de los Creditos Solicitados
    Set R = RecuperaColocaciones(psCtaCod)
    If R.RecordCount > 0 Then
        dVigencia = IIf(IsNull(R!dVigencia), 0, R!dVigencia)
        R.Close
    Else
        Set R = Nothing
        Exit Function
    End If
    
    Set R = RecuperaProducto(psCtaCod)
    If R.RecordCount > 0 Then
        nSaldo = R!nSaldo
        nTasaInt = R!nTasaInteres
        R.Close
    Else
        Set R = Nothing
        Exit Function
    End If
                  
    Set R = RecuperaNroCuotas(psCtaCod)
    If R.RecordCount > 0 Then
        pnNroCuotas = R!nCuotas
        R.Close
    Else
        Set R = Nothing
        Exit Function
    End If
    
    Set R = RecuperaCalendarioPagosPendiente(psCtaCod, pbParalelo, pBrfa)
    
    ReDim MatCalend(R.RecordCount, 13)
    If R.RecordCount > 0 Then
        nMontoSaldo = CDbl(Format(IIf(IsNull(R!nMontoPrestamo), 0, R!nMontoPrestamo), "#0.00"))
        
        Do While Not R.EOF

            MatCalend(R.Bookmark - 1, 0) = Format(R!dVenc, "dd/mm/yyyy")
            MatCalend(R.Bookmark - 1, 1) = Trim(Str(R!nCuota))
            MatCalend(R.Bookmark - 1, 2) = Trim(Str(R!ncoloccalendestado))
            MatCalend(R.Bookmark - 1, 3) = Format(IIf(IsNull(R!nCapital), 0, R!nCapital), "#0.00")
            MatCalend(R.Bookmark - 1, 4) = Format(IIf(IsNull(R!nIntComp), 0, R!nIntComp), "#0.00")
            MatCalend(R.Bookmark - 1, 5) = Format(IIf(IsNull(R!nIntGracia), 0, R!nIntGracia), "#0.00")
            If dFecSisTmp > R!dVenc Then
               bFeriado = False
                'Si se le cobra Mora
                If R!dVenc <> 0 Then
                    For i = 0 To DateDiff("d", R!dVenc, dFecSisTmp) - 1

                        Set RTmp = DetallaFeriado(R!dVenc + i, Mid(psCtaCod, 4, 2))
                        If RTmp.RecordCount > 0 Then
                            bFeriado = True
                        Else
                            bFeriado = False
                            Exit For
                        End If
                        RTmp.Close

                        If i = 2 Then
                            Exit For
                        End If
                    Next i
                End If
                If bFeriado Then
                    MatCalend(R.Bookmark - 1, 6) = "0.00"
                Else
                    MatCalend(R.Bookmark - 1, 6) = Format(IIf(IsNull(R!nIntMor), 0, R!nIntMor), "#0.00")
                End If
            Else
                MatCalend(R.Bookmark - 1, 6) = Format(IIf(IsNull(R!nIntMor), 0, R!nIntMor), "#0.00")
            End If
            MatCalend(R.Bookmark - 1, 7) = Format(IIf(IsNull(R!nIntReprog), 0, R!nIntReprog), "#0.00")
            MatCalend(R.Bookmark - 1, 8) = Format(IIf(IsNull(R!nIntSuspenso), 0, R!nIntSuspenso), "#0.00")
            MatCalend(R.Bookmark - 1, 9) = Format(IIf(IsNull(R!nGasto), 0, R!nGasto), "#0.00")
            MatCalend(R.Bookmark - 1, 12) = Format(nMontoSaldo, "#0.00")
            nMontoSaldo = nMontoSaldo - IIf(IsNull(R!nSaldoCap), 0, R!nSaldoCap)
            nMontoSaldo = CDbl(Format(nMontoSaldo, "#0.00"))
            MatCalend(R.Bookmark - 1, 10) = Format(nMontoSaldo, "#0.00")
            MatCalend(R.Bookmark - 1, 11) = Format(IIf(IsNull(R!nIntCompVenc), 0, R!nIntCompVenc), "#0.00")
            MatCalend(R.Bookmark - 1, 12) = Format(IIf(IsNull(R!nitf), 0, R!nitf), "#0.00")
            R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing

    'Si es cuota libre hallar Interes a la Fecha
    If nColocCalendCod = 70 And nDiasAtraso <= 0 Then
        Set R = RecuperaUltimoMovimiento(psCtaCod)
        If R.RecordCount > 0 Then
            dVigencia = R!dFecPago
        End If
        R.Close
        nDiasTranscurridos = DateDiff("d", dVigencia, dFecSisTmp)
        If nDiasTranscurridos > 0 Then
            nIntFecha = (TasaIntPerDias(nTasaInt, nDiasTranscurridos) * nSaldo) + nIntPend
        Else
            nIntFecha = nIntPend
        End If

        For i = 0 To UBound(MatCalend) - 1
            If i = 0 Then
                MatCalend(i, 4) = Format(nIntFecha, "#0.00")
            Else
                MatCalend(i, 4) = "0.00"
            End If
        Next i

    End If

    RecuperaMatrizCalendarioPendiente = MatCalend
    Exit Function

ErrorRecuperaMatrizCalendarioPendiente:
    ReDim MatCalend(0, 0)
    RecuperaMatrizCalendarioPendiente = MatCalend
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'Add GITU 30-07-2009
Public Function RecuperaColocacCred(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaColocacCred
    sSQL = "Select * from DBCMAC..ColocacCred where cCtacod = '" & psCtaCod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocacCred = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocacCred:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaColocaciones(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaColocaciones
    
    '*** peac 20071201
    'sSql = "Select * from Colocaciones where cCtaCod = '" & psCtaCod & "'"
    
    '*** peac 20071204
    sSQL = " Select a.*, isnull(w.nmonto,0) nMontoColUltCal, isnull(w.dvenc,0) dMontoColUltCal"
    sSQL = sSQL & " from DBCMAC..Colocaciones a"
    sSQL = sSQL & " Left Join"
    sSQL = sSQL & " (SELECT cc.cctacod,cc.nmonto, ccc.dvenc"
    sSQL = sSQL & " from DBCMAC..coloccalenddet cc"
    sSQL = sSQL & " inner join DBCMAC..colocaccred c on cc.cctacod=c.cctacod and c.nnrocalen=cc.nnrocalen"
    sSQL = sSQL & " inner join DBCMAC..coloccalendario ccc on ccc.cctacod=cc.cctacod and ccc.nnrocalen=cc.nnrocalen and ccc.ncoloccalendapl=0"
    sSQL = sSQL & " where cc.ncoloccalendapl=0 and nprdconceptocod=1000) w on  a.cctacod=w.cctacod"
    sSQL = sSQL & " where a.cCtaCod = '" & psCtaCod & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocaciones = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaColocaciones:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaProducto(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaProducto
    sSQL = "Select * from DBCMAC..Producto where cCtacod = '" & psCtaCod & "' "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaProducto = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaProducto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
Public Function RecuperaNroCuotas(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaProducto
    sSQL = "Select  top 1 nCuotas From DBCMAC..ColocacEstado Where cCtaCod = '" & psCtaCod & "' and nPrdEstado in (2020,2030) order by dPrdEstado"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaNroCuotas = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaProducto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaCalendarioPagosPendiente(ByVal psCtaCod As String, Optional ByVal pbCalParalelo As Boolean = False, _
Optional ByVal pBrfa As Boolean) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

'ARCV 27-02-2007
Dim dFecSisTmp As Date

dFecSisTmp = CDate(LeeConstSistema(16))

'--------

    On Error GoTo ErrorRecuperaCalendarioDesemb
    
    If pBrfa = False Then
            sSQL = "select C.dVenc,C.nCuota, C.nColocCalendEstado,"
            sSQL = sSQL & " nMontoPrestamo = (select SUM(nMonto) from DBCMAC..ColocCalendDet CD2 Inner Join DBCMAC..ColocCalendario C2 ON CD2.cCtaCod=C2.cCtaCod AND CD2.nNroCalen=C2.nNroCalen AND CD2.nColocCalendApl = C2.nColocCalendApl AND CD2.nCuota = C2.nCuota where CD2.cCtaCod = C.cCtaCod And CD2.nNroCalen = C.nNroCalen And CD2.nColocCalendApl=C.nColocCalendApl and CD2.nPrdConceptoCod in(1000,1010) AND C2.nColocCalendEstado = 0 ), "
            sSQL = sSQL & " nSaldoCap=(select nMonto from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(1000,1010)),"
            sSQL = sSQL & " nCapital=(select nMonto - nMontoPagado from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(1000,1010)),"
            sSQL = sSQL & " nIntComp=(select nMonto - nMontoPagado from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(1100,1107)),"
            sSQL = sSQL & " nIntCompVenc=(select nMonto - nMontoPagado from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1105),"
            sSQL = sSQL & " nIntGracia=(select nMonto - nMontoPagado from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1102),"
            
            sSQL = sSQL & " nIntMor=(select ROUND(nMonto,2) - ROUND(nMontoPagado,2) from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in ( 1101,1108)),"
            '----------
            sSQL = sSQL & " nIntReprog=(select nMonto - nMontoPagado from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1103),"
            sSQL = sSQL & " nIntSuspenso=(select nMonto - nMontoPagado from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1104),"
            sSQL = sSQL & " nGasto= DBCMAC.dbo.ColocCred_ObtieneGastoFechaCuota('" & psCtaCod & "',C.nCuota,'" & Format(dFecSisTmp, "mm/dd/yyyy") & "'), "
            '****************
            sSQL = sSQL & " nITF=(select nMonto - nMontoPagado from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(20,21)) "
            
            sSQL = sSQL & " from DBCMAC..ColocCalendario C "
            sSQL = sSQL & " Where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= 1 And nNroCalen = (select " & IIf(pbCalParalelo, "nNroCalPar", "nNroCalen") & " from DBCMAC..ColocacCred where cCtaCod = C.cCtaCod) "
            
            sSQL = sSQL & " AND C.nColocCalendEstado = 0"
            sSQL = sSQL & " order by C.nCuota"

    Else
        
            sSQL = "select C.dVenc,C.nCuota, C.nColocCalendEstado,"
            sSQL = sSQL & " nMontoPrestamo = (select SUM(nMonto) from DBCMAC..ColocCalendDet CD2 Inner Join DBCMAC..ColocCalendario C2 ON CD2.cCtaCod=C2.cCtaCod AND CD2.nNroCalen=C2.nNroCalen AND CD2.nColocCalendApl = C2.nColocCalendApl AND CD2.nCuota = C2.nCuota where CD2.cCtaCod = C.cCtaCod And CD2.nNroCalen = C.nNroCalen And CD2.nColocCalendApl=C.nColocCalendApl and CD2.nPrdConceptoCod in(1000,1010) AND C2.nColocCalendEstado = 0), "
            sSQL = sSQL & " nSaldoCap=(select nMonto from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (1000,1010)),"
            sSQL = sSQL & " nCapital=(select nMonto - nMontoPagado from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in (1000,1010)),"
            sSQL = sSQL & " nIntComp=(select nMonto - nMontoPagado from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in ( 1100,1107)),"
            sSQL = sSQL & " nIntCompVenc=(select nMonto - nMontoPagado from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1105),"
            sSQL = sSQL & " nIntGracia=(select nMonto - nMontoPagado from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1102),"
            sSQL = sSQL & " nIntMor=(select ROUND(nMonto,2) - ROUND(nMontoPagado,2) from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in ( 1101,1108)),"
            '----------
            sSQL = sSQL & " nIntReprog=(select nMonto - nMontoPagado from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1103),"
            sSQL = sSQL & " nIntSuspenso=(select nMonto - nMontoPagado from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod = 1104),"
            sSQL = sSQL & " nGasto= DBCMAC.dbo.ColocCred_ObtieneGastoFechaCuota('" & psCtaCod & "',C.nCuota,'" & Format(dFecSisTmp, "mm/dd/yyyy") & "'), "
            '****************

            sSQL = sSQL & " nITF=(select nMonto - nMontoPagado from DBCMAC..ColocCalendDet where cCtaCod = C.cCtaCod And nNroCalen = C.nNroCalen And nColocCalendApl=C.nColocCalendApl and nCuota = C.nCuota and nPrdConceptoCod in(20,21)) "
            
            sSQL = sSQL & " from DBCMAC..ColocCalendario C "
            sSQL = sSQL & " Where C.cCtaCod = '" & psCtaCod & "' And C.nColocCalendApl= 1 And nNroCalen = (select " & IIf(pbCalParalelo, "nNroCalPar", "nNroCalen") & " from DBCMAC..ColocacCred where cCtaCod = C.cCtaCod) "
            
            sSQL = sSQL & " AND C.nColocCalendEstado = 0"
            sSQL = sSQL & " order by C.nCuota"

    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaCalendarioPagosPendiente = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCalendarioDesemb:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaUltimoMovimiento(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    sSQL = " Select MC.nMovNro, MC.cOpeCod, dbo.FechaMov(M.cMovNro) as dFecPago,"
    sSQL = sSQL & " nNroCuota = (Select MAX(nNroCuota) from DBCMAC..MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro),"
    sSQL = sSQL & " nMontoPagado = (Select ISNULL(SUM(nMonto),0) from DBCMAC..MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro),"
    sSQL = sSQL & " nCapital = (Select ISNULL(SUM(nMonto),0) from DBCMAC..MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod = 1000 ),"
    sSQL = sSQL & " nInteres = (Select ISNULL(SUM(nMonto),0) from DBCMAC..MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod in (1100,1102,1103,1104,1105)),"
    sSQL = sSQL & " nMora = (Select ISNULL(SUM(nMonto),0) from DBCMAC..MovColDet Where cCTaCod = MC.cCtaCod AND  nMovNro = MC.nMovNro AND nPrdConceptoCod = 1101),"
    sSQL = sSQL & " nGastos = (Select ISNULL(SUM(nMonto),0) from DBCMAC..MovColDet Where cCTaCod = MC.cCtaCod AND nMovNro = MC.nMovNro AND nPrdConceptoCod like  '12%'),"
    sSQL = sSQL & " MC.nDiasMora , MC.nSaldoCap"
    sSQL = sSQL & " from DBCMAC..MovCol MC Inner Join DBCMAC..Mov M ON M.nMovNro = MC.nMovNro"
    sSQL = sSQL & " where MC.cOpeCod like '100%' AND MC.cCtaCod = '" & psCtaCod & "' "
    sSQL = sSQL & " AND (SUBSTRING(MC.cOpeCod,1,4) in ('1002','1003','1004','1005','1006','1007','1070'))"
    sSQL = sSQL & " AND MC.nMovNro = (Select MAX(nMovNro) From DBCMAC..MovCol Where cCtaCod = '" & psCtaCod & "' AND (SUBSTRING(cOpeCod,1,4) in ('1002','1003','1004','1005','1006','1007','1070')) ) "

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaUltimoMovimiento = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function DetallaFeriado(ByVal FecVer As Date, ByVal psCodAge As String) As ADODB.Recordset
Dim VSQL As String
Dim Co As DConecta
Dim rs As New ADODB.Recordset
Set Co = New DConecta
' Verifica si la Fecha seleccionada ya existe en la tabla feriados
'VSQL = "select dFeriado, cDescrip from Feriado where dFeriado = '" & Format(FecVer, "MM/DD/YYYY") & "' "

VSQL = " Select F.dferiado, FA.cCodAge"
VSQL = VSQL & " from DBCMAC..Feriado F"
VSQL = VSQL & " Inner Join DBCMAC..FeriadoAge FA ON F.dFeriado=FA.dFeriado"
'By Capi 06122007 por formato fecha
'VSQL = VSQL & " Where F.dFeriado='" & Format(FecVer, "MM/DD/YYYY") & "' AND FA.cCodAge = '" & psCodAge & "'"
VSQL = VSQL & " Where F.dFeriado='" & Format(FecVer, "YYYYMMDD") & "' AND FA.cCodAge = '" & psCodAge & "'"

Co.AbreConexion
Set rs = Co.CargaRecordSet(VSQL)
Co.CierraConexion
Set DetallaFeriado = rs
End Function
Public Function LeeConstSistema(ByVal psConstSistCod As Integer) As String
    Dim rsVar As Recordset
    Dim sSQL As String
    Dim oCon  As DConecta
    Set oCon = New DConecta
    
    If oCon.AbreConexion = False Then Exit Function
    sSQL = "Select nConsSisDesc, nConsSisValor From DBCMAC..ConstSistema where nConsSisCod =" & psConstSistCod & ""
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
'***********************************************
'*  FUNCION QUE HALLA EL INTERES DE UN PERIODO DE DIAS TRANACURRIDOS
'***********************************************
Public Function TasaIntPerDias(ByVal pnTasaInter As Double, ByVal pnDiasTrans As Integer) As Double
    TasaIntPerDias = ((1 + pnTasaInter / 100) ^ (pnDiasTrans / 30)) - 1
End Function

Public Function PagoCredito(ByVal psCtaCod As String, ByVal pnMonto As Double, ByVal pnMoneda As Integer, ByVal pdFecSis As String, ByRef pnMovNro As Long) As String
'Dim MatCalend As Variant
'Dim MatCalendTmp As Variant
'Dim MatCalendDistribuido As Variant
Dim nMonPago As Double
'Dim lnMontoPago As Double
Dim prsCredVig As ADODB.Recordset
Dim RGas As ADODB.Recordset
Dim oBase As New DConecta
Dim lnGastos  As Double
Dim lnMora As Double
Dim lnMovNro As Long
Dim lnCuotasMora As Double
Dim nInteresFecha  As Double
Dim nInterFechaGra As Double
Dim lnTotalDeuda As Double
Dim nMontoFecha As Double
Dim lnTC As Double
Dim lnITF As Double
Dim lnMonCuotaPend As Double
Dim lnNewSalCap As Double
Dim lnNewCPend As Integer
Dim psAgencia As String
Dim lsCodAge As String
Dim lsCodUser As String
Dim psPersCod As String
Dim ldProxfec As Date
Dim lsEstado As String
Dim lsMetLiquid As String
Dim sCabe As String
Dim lsCadena As String
Dim lnDifPag As Double

    Set prsCredVig = RecuperaDatosCreditoVigente(psCtaCod, pdFecSis)

    If Not prsCredVig.BOF And Not prsCredVig.EOF Then
        psAgencia = ""
        pMatCalend = RecuperaMatrizCalendarioPendiente(psCtaCod)
        pMatCalendTmp = pMatCalend
        pMatCalendDistribuido = CrearMatrizparaAmortizacion(pMatCalend)
        
        psPersCod = prsCredVig!cPersCod

        oBase.AbreConexion
        
        Set RGas = oBase.CargaRecordSet("SELECT nGasto=DBCMAC.dbo.ColocCred_ObtieneGastoFechaCredito('" & psCtaCod & "','" & Format(pdFecSis, "mm/dd/yyyy") & "')")
        lnGastos = RGas!nGasto
        
        nMonPago = MatrizMontoAPagar(pMatCalend, pdFecSis)
        
        lnMora = Format(MatrizMoraTotal(pMatCalend, pdFecSis), "#0.00")
        
        lnCuotasMora = MatrizCuotasEnMora(pMatCalend, pdFecSis)
        
        nInteresFecha = MatrizInteresGastosAFecha(psCtaCod, pMatCalend, pdFecSis, True, False)
        nInterFechaGra = MatrizInteresGraAFecha(psCtaCod, pMatCalend, pdFecSis)
        
        nMontoFecha = MatrizCapitalAFecha(psCtaCod, pMatCalend, pdFecSis)
        
        lnTotalDeuda = Format(nInteresFecha + nInterFechaGra + nMontoFecha, "#0.00")
        
        lnTotalDeuda = lnTotalDeuda - CDbl(MatrizGastosFecha(psCtaCod, pMatCalend))
        
        lnTotalDeuda = Format(lnTotalDeuda + RGas!nGasto, "#0.00")
        
        If lnTotalDeuda < pnMonto Then
            pnMonto = lnTotalDeuda
            lnDifPag = lnTotalDeuda - pnMonto
        End If

        If Mid(psCtaCod, 9, 1) = 2 Then
            lnTC = EmiteTipoCambio(pdFecSis, TCFijoDia)
        Else
            lnTC = 1
        End If
        oBase.CierraConexion
        Set oBase = Nothing

        Call ActualizaMontoPago(pnMonto, lnTotalDeuda, psCtaCod, pdFecSis, Trim(prsCredVig!cMetLiquidacion), IIf(IsNull(prsCredVig!nIntPend), 0, prsCredVig!nIntPend), 0, False, False, False, nMonPago, _
                0, "", lnITF, 0, lnMonCuotaPend, lnNewSalCap, lnNewCPend, ldProxfec, lsEstado)


        Call AmortizarCredito(psCtaCod, pMatCalend, pMatCalendDistribuido, _
                  pnMonto, pdFecSis, lsMetLiquid, 1, _
                  lsCodAge, lsCodUser, , , , , , _
                  0, , , , , , _
                  , , , , , , , , , , , , , , pnMovNro)
                  
        If pnMovNro > 0 Then
            Call RegistraPITMov(pnMovNro, 0, -1, gsPAN, gsDNI, pnMonto, sHora, sMesDia, nMoneda, gnTramaId, 0)
        End If
        

    sCabe = "1P" & Format(Now, "YYMMDD") & "0040"
    lsCadena = sCabe
    lsCadena = lsCadena & "MONTO DE PAGO     :        " & Right(Space(12) & Format(pnMonto - lnITF, "#,000.00"), 12)
    lsCadena = lsCadena & "I.T.F.            :        " & Right(Space(12) & Format(lnITF, "#,#00.00"), 12)
    lsCadena = lsCadena & "TOTAL             :        " & Right(Space(12) & Format(pnMonto, "#,#00.00"), 12)
    lsCadena = lsCadena & "FECHA VENCIMIENTO :        " & Right(Space(12) & Format(ldProxfec, "dd/mm/yyyy"), 12)
    lsCadena = lsCadena & "MONTO PROX CUOTA  :        " & Right(Space(12) & Format(lnMonCuotaPend, "#,#00.00"), 12)
    If lnDifPag > 0 Then
        lsCadena = lsCadena & "DEVOLUCION POR EXCESO CUOTA :  " & Right(Space(8) & Format(lnDifPag, "#,#00.00"), 8)
    End If
    
    'lsCadena = CStr(Len(lsCadena) + 16) & lsCadena
    
    PagoCredito = lsCadena
    
    loConec.CierraConexion

    '*******************************************************************************
    

    
    Set loConec = Nothing
                  
    
    End If
End Function

Public Function ActualizaMontoPago(ByVal pnMontoAPagar As Double, ByVal pnTotalDeuda As Double, _
                                    ByVal psCtaCod As String, ByVal pdFecSis As Date, _
                                    ByVal psMetLiquidacion As String, _
                                    ByVal pvnIntPendiente As Double, ByVal pvnIntPendientePagado As Double, _
                                    ByVal pbCalenCuotaLibre As Boolean, ByVal pbCalenDinamic As Boolean, ByVal pbPrepago As Integer, _
                                    ByRef pnMontoPago As Double, ByVal pnMonCalDin As Double, _
                                    ByRef psMensaje As String, ByRef pnITF As Double, _
                                    ByRef pnInteresDesagio As Double, ByRef pnMonCuotaPend As Double, _
                                    ByRef pnNewSalCap As Double, ByRef pnNewCPend As Integer, _
                                    ByRef pdProxfec As Date, ByRef psEstado As String, Optional ByVal pnMonIntGra As Double) As Boolean


Dim nMontoGastoGen As Double

Dim nInteresFecha As Double
Dim AcumMontoAPagar As Double
Dim i As Integer
Dim nInterFechaGra As Double 'Add by GITU 07-05-2009
Dim nMonIntGra As Double



On Error GoTo ErrorActualizaMontoPago

ActualizaMontoPago = True

    If pnMontoAPagar = 0 Then
        psMensaje = "Monto de Pago Debe ser mayor que Cero"
        ActualizaMontoPago = False
        Exit Function
    End If

    fgITFParametros

    If Mid(psCtaCod, 6, 3) = "423" Then
        pnITF = 0
    Else
        Dim lnValor As Double
        lnValor = pnMontoAPagar * gnITFPorcent
        lnValor = CortaDosITF(lnValor)
        pnITF = lnValor 'CalculoSinRedondeo(pnMontoAPagar)
    End If
    
    pnMontoPago = fgITFCalculaImpuestoNOIncluido(pnMontoAPagar)

    pnMontoPago = pnMontoAPagar

    nInteresFecha = MatrizInteresGastosAFecha(psCtaCod, pMatCalend, pdFecSis, True)
    nInterFechaGra = MatrizInteresGraAFecha(psCtaCod, pMatCalend, pdFecSis)
    pnInteresDesagio = 0
    
    'Se modifico 23-03 (Gastos en COM)
    Dim nNumGastosFinal As Integer
    Dim MatGastosFinal As Variant


        'Si es Pago Normal
        If pnMontoPago <> pnTotalDeuda Then
           pMatCalend = pMatCalendTmp

            'Distribuye Monto

            If Mid(psMetLiquidacion, 3, 1) = "i" Or Mid(psMetLiquidacion, 3, 1) = "Y" Then
                If pnMontoAPagar >= pnTotalDeuda Then
                    pMatCalendDistribuido = MatrizDistribuirCancelacion(psCtaCod, pMatCalend, pnMontoPago, psMetLiquidacion, pdFecSis, True)
                Else
                    pMatCalendDistribuido = MatrizDistribuirCancelacion(psCtaCod, pMatCalend, pnMontoPago, psMetLiquidacion, pdFecSis, False, , False)
                End If
            Else
                pMatCalendDistribuido = MatrizDistribuirMonto(pMatCalend, pnMontoPago, psMetLiquidacion, , pnMonIntGra)
            End If
        Else 'Si es una Cancelacion del Credito
            'Distribuye Monto
            pMatCalendDistribuido = MatrizDistribuirCancelacion(psCtaCod, pMatCalend, pnMontoPago, psMetLiquidacion, pdFecSis, True, pbCalenDinamic)
        End If
'    End If

    pnNewSalCap = MatrizSaldoCapital(pMatCalend, pMatCalendDistribuido)
    pnNewCPend = MatrizCuotaPendiente(pMatCalend, pMatCalendDistribuido)
    pnMonCuotaPend = MatrizMontoCuotaPendiente(pMatCalend, pMatCalendDistribuido)
    pdProxfec = Format(MatrizFechaCuotaPendiente(pMatCalend, pMatCalendDistribuido), "dd/mm/yyyy")
    psEstado = IIf(MatrizEstadoCalendario(pMatCalendDistribuido) = 1, "CANCELADO", "VIGENTE")
    If psEstado = "CANCELADO" Then
        pdProxfec = 0   ' Se pone 0 en label=""
    End If

    Exit Function
ErrorActualizaMontoPago:
    Err.Raise Err.Number, "Error", Err.Description

End Function
Public Function MatrizMoraTotal(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim i As Integer
    MatrizMoraTotal = 0
    For i = 0 To UBound(MatCalend) - 1
        If pdHoy >= CDate(MatCalend(0, 0)) Then
            MatrizMoraTotal = MatrizMoraTotal + CDbl(MatCalend(i, 6))
        End If
    Next i
End Function
Public Function MatrizCuotasEnMora(ByVal MatCalend As Variant, ByVal pdHoy As Date) As Integer
Dim i As Integer
    MatrizCuotasEnMora = 0
    For i = 0 To UBound(MatCalend) - 1
        If pdHoy > CDate(MatCalend(i, 0)) Then
            MatrizCuotasEnMora = MatrizCuotasEnMora + 1
        End If
    Next i
End Function
Public Function SaldoK(ByVal pcCtaCod As String, ByVal dPago As Date) As Currency
    Dim oConec As DConecta
    Dim sSQL As String
    Dim rs As ADODB.Recordset

    sSQL = "select nMonto  as SK from DBCMAC..ColocCalendDet "
    sSQL = sSQL & " where cCtaCod='" & pcCtaCod & "' and nColocCalendApl=0 and nPrdConceptoCod=1000 "
    sSQL = sSQL & "     and nNroCalen=(select nNroCalen from DBCMAC..ColocacCred where cCtaCod='" & pcCtaCod & "') "
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)

    If Not rs.EOF And Not rs.BOF Then
        SaldoK = rs!SK
    End If
    'En el Caso pague cuotas atrasada y cancele el credito
    sSQL = "Select isnull(Sum(nMonto),0) as Capital"
    sSQL = sSQL & " From DBCMAC..ColocCalendDet a "
    sSQL = sSQL & " inner join DBCMAC..ColocCalendario b on a.cctacod=b.cctacod and a.nnrocalen=b.nnrocalen and "
    sSQL = sSQL & " a.nColocCalendApl=b.nColocCalendApl and a.ncuota=b.nCuota "
    sSQL = sSQL & " Where a.nPrdConceptoCod in(1000,1010) and b.dVenc<='" & Format(CStr(dPago), "MM/dd/yyyy") & "' and b.nColocCalendApl=1 "
    sSQL = sSQL & " and a.cctacod='" & pcCtaCod & "' and a.nNroCalen=(Select nNroCalen From DBCMAC..ColocacCred Where cCtaCod='" & pcCtaCod & "')"

    Set rs = New ADODB.Recordset
    Set rs = oConec.CargaRecordSet(sSQL)
    If Not rs.EOF And Not rs.BOF Then
        SaldoK = SaldoK - IIf(IsNull(rs!capital), 0, rs!capital)
    End If
    Set rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function RecuperaFechaInicioCuota(ByVal psCtaCod As String, ByVal pnCuota As Integer, ByVal pnAplicado As Integer, Optional ByVal pbCalendDin As Boolean = False) As Date
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaFechaInicioCuota
    If pnCuota > 1 Then
        sSQL = "Select CC.dVenc from DBCMAC..ColocCalendario CC "
        sSQL = sSQL & " Where CC.nCuota = " & pnCuota - 1 & " And nColocCalendApl = " & pnAplicado & " AND cCtaCod = '" & psCtaCod & "'"
        sSQL = sSQL & " And nNroCalen = (Select nNroCalen From DBCMAC..ColocacCred Where cCtaCod = CC.cCtaCod)"
    Else
        If pbCalendDin Then
            sSQL = "Select CC.dVenc DBCMAC..from ColocCalendario CC "
            sSQL = sSQL & " Where CC.nCuota = " & pnCuota & " And nColocCalendApl = 0 AND cCtaCod = '" & psCtaCod & "'"
            sSQL = sSQL & " And nNroCalen = (Select nNroCalen From DBCMAC..ColocacCred Where cCtaCod = CC.cCtaCod)"
        Else
            'sSql = "Select dVigencia as dVenc from Colocaciones Where cCtaCod = '" & psCtaCod & "'"
            'peac 20071219
            sSQL = " select dvenc from DBCMAC..coloccalendario where cctacod='" & psCtaCod & "'"
            sSQL = sSQL & " and ncoloccalendapl=0"
            sSQL = sSQL & " and nnrocalen=(select nnrocalen from DBCMAC..colocaccred where cctacod='" & psCtaCod & "')"
        End If
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    RecuperaFechaInicioCuota = Format(R!dVenc, "dd/mm/yyyy")
    
    Exit Function

ErrorRecuperaFechaInicioCuota:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function DevuelveInteresAFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
    Dim R As ADODB.Recordset
    Dim oConec As DConecta
    Dim nTasa As Double
    Dim dFecIni As Date
    Dim nMontoAFecha As Double
    Dim sSQL As String
    Dim i As Integer
    
    Set R = RecuperaProducto(psCtaCod)
    nTasa = R!nTasaInteres
    R.Close
    Set R = Nothing
    
    dFecIni = Format(RecuperaFechaInicioCuota(psCtaCod, CInt(MatCalend(0, 1)), 1, 0), "dd/mm/yyyy")
    
    nMontoAFecha = TasaIntPerDias(nTasa, pdHoy - dFecIni) * CDbl(SaldoK(psCtaCod, pdHoy))
    
    sSQL = "Select isnull(nMontoPagado,0)  as nMontoPagado From DBCMAC..ColocCalendDet Where cCtaCod = '" & psCtaCod & "' AND nColocCalendApl = 1 "
    sSQL = sSQL & " AND nNroCalen =  (Select nNroCalen from DBCMAC..ColocacCred Where cctaCod = '" & psCtaCod & "')"
    sSQL = sSQL & " AND nPrdConceptoCod = 1100 AND nCuota = " & MatCalend(i, 1)
                
    Set oConec = New DConecta
    oConec.AbreConexion
    Set R = oConec.CargaRecordSet(sSQL)
    Dim nMonto As Double
    If Not R.EOF And Not R.BOF Then
        nMonto = IIf(IsNull(R!nMontoPagado), 0, R!nMontoPagado)
    Else
        nMonto = 0
    End If

    nMontoAFecha = nMontoAFecha - nMonto
    
    'ARCV 14-11-2006 (YA NO COBRAMOS DESAGIO)
    If nMontoAFecha < 0 Then
        nMontoAFecha = 0
    End If
                
'    If Not pbDesagio And nMontoAFecha < 0 Then
'        nMontoAFecha = 0
'    End If
    
    DevuelveInteresAFecha = nMontoAFecha
    
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function MatrizInteresGraAFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date) As Double
Dim R As ADODB.Recordset
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados

    Set R = RecuperaColocaciones(psCtaCod)
     
    dFecIni = R!dMontoColUltCal
    nSaldoCal = R!nMontoColUltCal
    
    R.Close
    Set R = Nothing
    
    Set R = RecuperaProducto(psCtaCod)
    pnTasa = R!nTasaInteres
    R.Close
    Set R = Nothing

    'MatrizInteresGraAFecha = MatrizInteresGraciaFecha(psCtaCod, MatCalend, pdHoy, nSaldoCal)
                
End Function
Public Function MatrizCapitalAFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant, Optional ByVal pdFecSis As Date) As Double
Dim nMontoGanado As Double
Dim i As Integer
    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    For i = 0 To UBound(MatCalend) - 1
            nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 3)) 'Capital
    Next i

    'Total Calculado es
    MatrizCapitalAFecha = CDbl(Format(nMontoGanado, "#0.00"))

End Function
Public Function RecuperaDatosCreditoVigente(ByVal psCtaCod As String, Optional ByVal pdHoy As Date = CDate("01/01/1900"), _
    Optional ByVal pbIncluirbAprobados As Boolean = False) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaDatosCreditoVigente
    sSQL = "Select PP.cPersCod, CC.cMetLiquidacion, DBCMAC.dbo.CuotaAprobada(P.cCtaCod) as CuotaAprobada, CC.bCuotaCom, CC.bMiVivienda, CC.nCalendDinamTipo, CC.bPrepago, CC.nCalPago, Pers.cPersNombre, C.cLineaCred,C.nMontoCol, P.nSaldo, "
    sSQL = sSQL & " cMoneda = (Select cConsDescripcion from DBCMAC..Constante where nConsCod = 1011 AND nConsValor = " & CInt(Mid(psCtaCod, 9, 1)) & "),"
    sSQL = sSQL & " CE.nCuotas nCuotasApr, CE.nColocCalendCod " & ", "
    sSQL = sSQL & " CC.nDiasAtraso, CC.cMetLiquidacion, P.nTransacc, CC.nCalendDinamico, CC.nIntPend, CC.nNroCalen, CC.nNroProxCuota, CC.nNroProxDesemb "
    If pdHoy <> CDate("01/01/1900") Then
        sSQL = sSQL & ", DATEDIFF(year,C.dVigencia,'" & Format(pdHoy, "mm/dd/yyyy") & "') as nPlazoTranscurrido "
    End If
    sSQL = sSQL & " From DBCMAC..Producto P Inner join DBCMAC..ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = 20 "
    sSQL = sSQL & "                 Inner Join DBCMAC..Persona Pers ON PP.cPersCod = Pers.cPersCod "
    sSQL = sSQL & "                 Inner Join DBCMAC..Colocaciones C ON C.cCtaCod =  P.cCtaCod "
    sSQL = sSQL & "                 Inner Join DBCMAC..ColocacCred CC ON CC.cCtaCod =  P.cCtaCod "
    sSQL = sSQL & "                 Inner Join DBCMAC..ColocacEstado CE ON CE.cCtaCod = P.cCtaCod And CE.nPrdEstado = 2002 "
    sSQL = sSQL & " WHERE P.cCtaCod = '" & psCtaCod & "'"

    If pbIncluirbAprobados Then
        sSQL = sSQL & " AND P.nPrdEstado in (2032,2030,2031,2022,2020,2021,2002)"
    Else
        sSQL = sSQL & " AND P.nPrdEstado in (2032,2030,2031,2022,2020,2021)"
    End If

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaDatosCreditoVigente = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaDatosCreditoVigente:
    Err.Raise Err.Number, "Error En Proceso RecuperaDatosCreditoVigente", Err.Description

End Function

Public Function CrearMatrizparaAmortizacion(ByVal MatCalend As Variant) As Variant
Dim MatCalendAmortiz() As String
Dim i As Integer
    ReDim MatCalendAmortiz(UBound(MatCalend), 12)

    For i = 0 To UBound(MatCalend) - 1
        MatCalendAmortiz(i, 0) = MatCalend(i, 0)
        MatCalendAmortiz(i, 1) = MatCalend(i, 1)
        MatCalendAmortiz(i, 2) = MatCalend(i, 2)
        MatCalendAmortiz(i, 3) = "0.00"
        MatCalendAmortiz(i, 4) = "0.00"
        MatCalendAmortiz(i, 5) = "0.00"
        MatCalendAmortiz(i, 6) = "0.00"
        MatCalendAmortiz(i, 7) = "0.00"
        MatCalendAmortiz(i, 8) = "0.00"
        MatCalendAmortiz(i, 9) = "0.00"
        MatCalendAmortiz(i, 11) = "0.00"
        MatCalendAmortiz(i, 10) = MatCalend(i, 10)
    Next i
    CrearMatrizparaAmortizacion = MatCalendAmortiz
End Function
Public Function MatrizInteresGastosAFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date, Optional ByVal pbDesagio As Boolean = False, Optional ByVal pbCalenDin As Boolean = False) As Double
Dim R As ADODB.Recordset
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados

    'Set oCredito = New COMDCredito.DCOMCredito
    Set R = RecuperaColocaciones(psCtaCod)
    'Set oCredito = Nothing

    '*** PEAC 20071204
    'dFecIni = R!dVigencia
    'nSaldoCal = R!nMontoCol

    dFecIni = R!dMontoColUltCal
    nSaldoCal = R!nMontoColUltCal


    R.Close
    Set R = Nothing

    'Set oCredito = New COMDCredito.DCOMCredito
    Set R = RecuperaProducto(psCtaCod)
    'Set oCredito = Nothing
    pnTasa = R!nTasaInteres
    R.Close
    Set R = Nothing

    MatrizInteresGastosAFecha = CDbl(Format(MatrizGastosFecha(psCtaCod, MatCalend) + _
                MatrizInteresReprogramadoFecha(psCtaCod, MatCalend) + _
                MatrizInteresSuspensoFecha(psCtaCod, MatCalend) + _
                MatrizInteresMorFecha(psCtaCod, MatCalend) + _
                MatrizInteresCompAFecha(psCtaCod, MatCalend, pdHoy, nSaldoCal, pnTasa, pbDesagio, pbCalenDin) + MatrizInteresCompensatorioVencido(MatCalend), "#0.00"))


End Function

Public Function MatrizGastosFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant) As Double
Dim R As ADODB.Recordset
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim i As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 9)) 'Gastos
    Next i

    'Total Calculado es
    MatrizGastosFecha = CDbl(Format(nMontoGanado, "#0.00"))

End Function

Public Function MatrizInteresReprogramadoFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant) As Double
Dim R As ADODB.Recordset
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim i As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 7)) 'Interes Reprogramado
    Next i

    'Total Calculado es
    MatrizInteresReprogramadoFecha = CDbl(Format(nMontoGanado, "#0.00"))

End Function
Public Function MatrizInteresSuspensoFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant) As Double
Dim R As ADODB.Recordset
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim i As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 8)) 'Interes Suspenso
    Next i

    'Total Calculado es
    MatrizInteresSuspensoFecha = CDbl(Format(nMontoGanado, "#0.00"))

End Function
Public Function MatrizInteresMorFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant) As Double
Dim R As ADODB.Recordset
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim i As Integer
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    nSaldoCap = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 6)) 'Interes Moratorio
    Next i

    'Total Calculado es
    MatrizInteresMorFecha = CDbl(Format(nMontoGanado, "#0.00"))

End Function
Public Function MatrizInteresCompAFecha(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date, _
    Optional pnMontoCol As Double = -1, Optional pnTasaInteres As Double = -1, Optional ByVal pbDesagio As Boolean = False, Optional ByVal pbCalenDin As Boolean = False) As Double
Dim R As ADODB.Recordset
'Dim oCalend As COMDCredito.DCOMCalendario
Dim nMontoGanado As Double
Dim nMontoAFecha As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim nSaldoCap As Double
Dim i As Integer
Dim pnTasa As Double
Dim nColocCalendCod As Integer
Dim sSQL As String
Dim oConec As DConecta
Dim nDiasGracia As Long
Dim nPlazo As Long
Dim ldFecha As Date
Dim dFechaDesp As Date
Dim ldFecIniGracia As Date
Dim nTipoPeriodo As Integer

    'Calculo de Montos Ya Ganados


    'Set oCredito = New COMDCredito.DCOMCredito
    Set R = RecuperaColocacCred(psCtaCod)
    'Set oCredito = Nothing
    nColocCalendCod = R!nColocCalendCod
    R.Close



    nMontoGanado = 0
    nSaldoCap = 0
    nMontoAFecha = 0

    If nColocCalendCod <> 70 Then
        For i = 0 To UBound(MatCalend) - 1
            If pdHoy >= Format(CDate(MatCalend(i, 0)), "dd/mm/yyyy") Then
                nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 4))
            End If
        Next i
    Else
        nMontoGanado = nMontoGanado + CDbl(MatCalend(0, 4))
    End If


    'revisar nSaldoCal
    If nColocCalendCod <> 70 Then
        If pnMontoCol = -1 Then
            Set R = RecuperaColocaciones(psCtaCod)

            '*** PEAC 20071204 **********
            'nSaldoCal = R!nMontoCol
            nSaldoCal = R!nMontoColUltCal
            '****************************

            R.Close
            Set R = Nothing
        Else
            nSaldoCal = pnMontoCol
        End If

        dFecIni = Format(RecuperaFechaInicioCuota(psCtaCod, CInt(MatCalend(0, 1)), 1, pbCalenDin), "dd/mm/yyyy")

        If pnTasaInteres = -1 Then
            Set R = RecuperaProducto(psCtaCod)
            pnTasa = R!nTasaInteres
            R.Close
            Set R = Nothing
        Else
            pnTasa = pnTasaInteres
        End If

        Set R = RecuperaColocacEstado(psCtaCod, 2002)

        If R.EOF Then
            nDiasGracia = 0
        Else
            nDiasGracia = R!nPeriodoGracia
            'Add by Gitu 21-08-08
            nTipoPeriodo = R!nColocCalendCod

            If R!nColocCalendCod = 20 Or R!nColocCalendCod = 21 _
                Or R!nColocCalendCod = 30 Or R!nColocCalendCod = 31 _
                Or R!nColocCalendCod = 10 Or R!nColocCalendCod = 11 Then
                nPlazo = R!nPlazo 'GITU 10/04/2008
            End If

            If R!nColocCalendCod = 50 Or R!nColocCalendCod = 51 _
                Or R!nColocCalendCod = 60 Or R!nColocCalendCod = 61 _
                Or R!nColocCalendCod = 40 Or R!nColocCalendCod = 41 Then
                nPlazo = 30
            End If

            ldFecIniGracia = CDate(Format(dFecIni, "dd/mm/yyyy")) + nPlazo
            'End Gitu
        End If

        R.Close
        Set R = Nothing

        'Interes a la fecha
        nMontoAFecha = 0
        For i = 0 To UBound(MatCalend) - 1
            If pdHoy >= Format(CDate(MatCalend(i, 0)), "dd/mm/yyyy") Then
                dFecIni = Format(CDate(MatCalend(i, 0)), "dd/mm/yyyy")
            Else
                If (pdHoy - dFecIni) > 0 Then
                    If MatCalend(i, 1) = "1" Then
                       'If TasaIntPerDias(pnTasa, (pdHoy - dFecIni) - nDiasGracia) > 0 Then
                       '     nMontoAFecha = TasaIntPerDias(pnTasa, (pdHoy - dFecIni) - nDiasGracia) * nSaldoCal
                       'Else
                            'Modify Gitu 21-08-2008
                            '** Se Agrego la validacion porque si cancela antes del plazo calculaba con todo los dias de gracia
                            If CDate(Format(pdHoy, "dd/mm/yyyy")) < CDate(Format(ldFecIniGracia, "yyyy/mm/dd")) Then
                                nMontoAFecha = TasaIntPerDias(pnTasa, (pdHoy - dFecIni)) * nSaldoCal
                            Else
                                'nMontoAFecha = TasaIntPerDias(pnTasa, (pdHoy - ldFecIniGracia)) * nSaldoCal
                                nMontoAFecha = TasaIntPerDias(pnTasa, (ldFecIniGracia - dFecIni)) * nSaldoCal
                            End If
                       'End If
                        'nMontoAFecha = TasaIntPerDias(pnTasa, pdHoy -dFecIni) * (CDbl(MatCalend(I, 10)) + CDbl(MatCalend(I, 3)))
                    Else
                        nMontoAFecha = TasaIntPerDias(pnTasa, pdHoy - dFecIni) * CDbl(SaldoK(psCtaCod, pdHoy))
                        'nMontoAFecha = TasaIntPerDias(pnTasa, pdHoy - dFecIni) * CDbl(MatCalend(I, 10))
                    End If
                End If
                sSQL = "Select isnull(nMontoPagado,0)  as nMontoPagado From DBCMAC..ColocCalendDet Where cCtaCod = '" & psCtaCod & "' AND nColocCalendApl = 1 "
                sSQL = sSQL & " AND nNroCalen =  (Select nNroCalen from DBCMAC..ColocacCred Where cctaCod = '" & psCtaCod & "')"
                sSQL = sSQL & " AND nPrdConceptoCod = 1100 AND nCuota = " & MatCalend(i, 1)

                'Comentado por gitu 23-09-2008
'                sSql = "Select nMontoPagado From ColocCalendDet Where cCtaCod = '" & psCtaCod & "' AND nColocCalendApl = 1 "
'                sSql = sSql & " AND nNroCalen =  (Select nNroCalen from ColocacCred Where cctaCod = '" & psCtaCod & "')"
'                sSql = sSql & " AND nPrdConceptoCod = 1100 AND nCuota = " & MatCalend(i, 1)
                Set oConec = New DConecta
                oConec.AbreConexion
                Set R = oConec.CargaRecordSet(sSQL)
                Dim nMonto As Double
                If Not R.EOF And Not R.BOF Then
                    nMonto = IIf(IsNull(R!nMontoPagado), 0, R!nMontoPagado)
                Else
                    nMonto = 0
                End If


                nMontoAFecha = nMontoAFecha - nMonto

                'ARCV 14-11-2006 (YA NO COBRAMOS DESAGIO)
                If nMontoAFecha < 0 Then
                    nMontoAFecha = 0
                End If

                If Not pbDesagio And nMontoAFecha < 0 Then
                    nMontoAFecha = 0
                End If
                oConec.CierraConexion
                Set oConec = Nothing
                Exit For
            End If
        Next i

    End If
    'Total Calculado es
    MatrizInteresCompAFecha = CDbl(Format(nMontoGanado + nMontoAFecha, "#0.00"))

End Function
Public Function RecuperaColocacEstado(ByVal psCtaCod As String, ByVal pnEstado As Integer, Optional pbTodos As Boolean = False) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaColocacEstado
    If Not pbTodos Then
        sSQL = "Select * from DBCMAC..ColocacEstado Where cCtaCod = '" & psCtaCod & "' AND nPrdEstado =  " & pnEstado
    Else
        sSQL = "Select * from DBCMAC..ColocacEstado Where cCtaCod = '" & psCtaCod & "' Order By nPrdEstado"
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaColocacEstado = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocacEstado:
    Err.Raise Err.Number, "Error En Proceso", Err.Description


End Function
Public Function MatrizInteresCompensatorioVencido(ByVal MatCalend As Variant) As Double
Dim nMontoGanado As Double
Dim i As Integer

    'Calculo de Montos Ya Ganados
    nMontoGanado = 0
    For i = 0 To UBound(MatCalend) - 1
        nMontoGanado = nMontoGanado + CDbl(MatCalend(i, 11)) 'Interes Suspenso
    Next i

    'Total Calculado es
    MatrizInteresCompensatorioVencido = CDbl(Format(nMontoGanado, "#0.00"))

End Function

Public Function EmiteTipoCambio(ByVal dFecha As Date, ByVal nTpoTipoCambio As TipoCambio) As Double
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As New DConecta
    
    Set oCon = New DConecta
    
    EmiteTipoCambio = 0

    If oCon.AbreConexion = False Then Exit Function
    rs.CursorLocation = adUseClient
    sql = "Select isnull(nValPondREU,0) nValPondREU,nValFijo, nValFijoDia, nValVent, nValComp, nValVentEsp, nValCompEsp, nValPond,nValPondVenta From DBCMAC..TipoCambio " _
        & " WHERE dFecCamb = (   Select Max(dFecCamb)" _
        & "                      From DBCMAC..TipoCambio " _
        & "                      Where datediff(day,dFecCamb,'" & Format$(dFecha, "mm/dd/yyyy") & "')=0)"
  
    Set rs = oCon.CargaRecordSet(sql)
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF And Not rs.BOF Then
        Select Case nTpoTipoCambio
            Case TCFijoMes
                EmiteTipoCambio = rs("nValFijo")
            Case TCFijoDia
                EmiteTipoCambio = rs("nValFijoDia")
            Case TCVenta
                EmiteTipoCambio = rs("nValVent")
            Case TCCompra
                EmiteTipoCambio = rs("nValComp")
            Case TCVentaEsp
                EmiteTipoCambio = rs("nValVentEsp")
            Case TCCompraEsp
                EmiteTipoCambio = rs("nValCompEsp")
            Case TCPonderado
                EmiteTipoCambio = rs("nValPond")
            Case TCPondVenta
                EmiteTipoCambio = IIf(IsNull(rs("nValPondVenta")), 0, rs("nValPondVenta"))
            Case TCPondREU
                EmiteTipoCambio = rs("nValPondREU")
        End Select
    End If
    rs.Close
    Set rs = Nothing
End Function

'*** Obtiene los parametros de ITF
Public Function fgITFParametros()
Dim oCon As DConecta
Set oCon = New DConecta
Dim lsSql As String
Dim lr As ADODB.Recordset
Set lr = New ADODB.Recordset
    
    lsSql = "select nParCod, nParValor FROM DBCMAC..PARAMETRO WHERE nParProd = 1000 And nParCod In (1001,1002,1003)"
    oCon.AbreConexion
    Set lr = oCon.CargaRecordSet(lsSql)
    
    Do While Not lr.EOF
        Select Case lr!nParCod
            Case 1001
                gbITFAplica = IIf(lr!nParValor = 0, False, True)
            Case 1003
                gnITFPorcent = lr!nParValor
            Case 1002
                gnITFMontoMin = lr!nParValor
        End Select
        lr.MoveNext
    Loop
    lr.Close
    Set lr = Nothing

    oCon.CierraConexion
    Set oCon = Nothing

End Function
Public Function fgITFCalculaImpuestoNOIncluido(ByVal pnMonto As Double, Optional ByVal bCancelacion As Boolean) As Double
Dim lnValor As Double
lnValor = pnMonto
If gbITFAplica = True Then
        If bCancelacion = True Then
            lnValor = fgTruncar(pnMonto * (1 + gnITFPorcent), 2)
        Else
            lnValor = pnMonto * (1 + gnITFPorcent)
        End If

        Dim aux As Double
        If bCancelacion = True Then
            If InStr(1, CStr(lnValor), ".", vbTextCompare) <> 0 Then
                aux = CDbl(CStr(Int(lnValor)) & "." & Mid(CStr(lnValor), InStr(1, CStr(lnValor), ".", vbTextCompare) + 1, 2))
                'ARCV 08-06-2006
                'aux = CDbl(lnValor)
            Else
                aux = lnValor
            End If
        Else
            aux = CDbl(CStr(Int(lnValor)) & "." & Mid(CStr(lnValor), InStr(1, CStr(lnValor), ".", vbTextCompare) + 1, 2))
            'ARCV 08-06-2006
            'aux = CDbl(lnValor)
        End If
        
        lnValor = aux
        lnValor = fgTruncar(lnValor, 2)
End If
fgITFCalculaImpuestoNOIncluido = lnValor
End Function

Public Function CortaDosITF(ByVal lnITF As Double) As Double
Dim intpos  As Integer
Dim lnDecimal As Double
Dim lsDec As String
Dim lnEntero As Long
Dim lnPos As Long

lnEntero = Int(lnITF)
lnDecimal = Round(lnITF - Int(lnEntero), 6)
lnPos = InStr(1, Trim(Str(lnDecimal)), ".")
If lnPos > 0 Then
    'ARCV 26-10-2006
    lsDec = Mid(Trim(Str(lnDecimal)), lnPos + 1, 2)
    lsDec = IIf(Len(lsDec) = 1, lsDec * 10, lsDec)
    lnDecimal = Val(lsDec) / 100
    'ARCV 08-06-2006
    CortaDosITF = lnEntero + lnDecimal
Else
    lnDecimal = 0
    CortaDosITF = lnEntero
End If
End Function
Public Function MontoTotalGastosGenerado(ByVal MatGastos As Variant, ByVal pnNumGastosCancel As Integer, _
    Optional ByVal psTipoGastoProc As Variant = "") As Double
Dim i As Integer
    MontoTotalGastosGenerado = 0
    For i = 0 To pnNumGastosCancel - 1
        If MatGastos(i, 4) = psTipoGastoProc(0) Or MatGastos(i, 4) = psTipoGastoProc(1) Or MatGastos(i, 4) = psTipoGastoProc(2) Then
            MontoTotalGastosGenerado = MontoTotalGastosGenerado + CDbl(MatGastos(i, 3))
        End If
    Next i
End Function

'ARCV 12-07-2006
'pbVerificaSoloCapital para las Cancelaciones de Credito con desembolso de otros
Public Function MatrizDistribuirCancelacion(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pnMontoPago As Double, _
                 ByVal psMetLiquidacion As String, ByVal pdHoy As Date, Optional ByVal pbNoCancelar As Boolean = False, Optional ByVal pbCalenDin As Boolean = False, _
                 Optional ByVal pCancel As Boolean = True, _
                 Optional ByVal pbVerificaSoloCapital As Boolean = False) As Variant

Dim MatCalendDistrib As Variant
Dim nMontoGastos As Double
Dim nMontoMora As Double
Dim nMontoInteres As Double
Dim nMontoInterGra As Double 'Add by gitu 07-05-2009
Dim nMontoCapital As Double
Dim nMontoTotalTmp As Double
Dim nMontoPagoTmp As Double
Dim J As Integer

Dim nMontoInteresFecha As Double 'ARCV 30-07-2006

        MatCalendDistrib = CrearMatrizparaAmortizacion(MatCalend)
        nMontoTotalTmp = 0
        nMontoPagoTmp = pnMontoPago
        For J = 1 To 4
                Select Case Mid(psMetLiquidacion, J, 1)
                    Case "G"
                        nMontoGastos = MatrizGastosFecha(psCtaCod, MatCalend)
                        If nMontoPagoTmp > nMontoGastos Then
                            nMontoPagoTmp = nMontoPagoTmp - nMontoGastos
                        Else
                            nMontoGastos = nMontoPagoTmp
                            nMontoPagoTmp = 0#
                        End If
                        nMontoTotalTmp = nMontoTotalTmp + nMontoGastos
                        Call MatrizDistribuirGastos(MatCalend, MatCalendDistrib, nMontoGastos, True)
                    Case "M"
                        nMontoMora = MatrizInteresMorFecha(psCtaCod, MatCalend)
                        If nMontoPagoTmp > nMontoMora Then
                            nMontoPagoTmp = nMontoPagoTmp - nMontoMora
                        Else
                            nMontoMora = nMontoPagoTmp
                            nMontoPagoTmp = 0#
                        End If
                        nMontoTotalTmp = nMontoTotalTmp + nMontoMora
                        Call MatrizDistribuirMora(MatCalend, MatCalendDistrib, nMontoMora, True)
                    Case "I", "i", "Y"
                        nMontoInteres = CDbl(Format(MatrizInteresTotalesAFechaSinMora(psCtaCod, MatCalend, pdHoy, pbCalenDin), "#0.00"))
                        nMontoInterGra = CDbl(Format(MatrizInteresGraAFecha(psCtaCod, MatCalend, pdHoy), "#0.00"))
                                                
                        nMontoInteresFecha = nMontoInteres 'ARCV 30-07-2006
                        
                        If pCancel = True Then
                            If nMontoPagoTmp > nMontoInteres Then
                                nMontoPagoTmp = nMontoPagoTmp - (nMontoInteres + nMontoInterGra)
                            Else
                                nMontoInteres = nMontoPagoTmp
                                nMontoPagoTmp = 0#
                            End If
                            nMontoTotalTmp = nMontoTotalTmp + nMontoInteres + nMontoInterGra
                            Call MatrizDistribuirInteresI(MatCalend, MatCalendDistrib, nMontoInteres + nMontoInterGra, True, nMontoInterGra)
                       Else
                            If nMontoPagoTmp > IIf(nMontoInteres > 0, nMontoInteres, 0) Then
                                nMontoPagoTmp = nMontoPagoTmp - (IIf(nMontoInteres > 0, nMontoInteres, 0) + nMontoInterGra)
                            Else
                                nMontoInteres = nMontoPagoTmp
                                nMontoPagoTmp = 0#
                            End If
                            nMontoTotalTmp = nMontoTotalTmp + IIf(nMontoInteres > 0, nMontoInteres, 0) + nMontoInterGra 'Corregido Gitu 26-08-09
                            Call MatrizDistribuirInteresI(MatCalend, MatCalendDistrib, IIf(nMontoInteres > 0, nMontoInteres, 0) + nMontoInterGra, True) 'Corregido Gitu 26-08-09
                       End If
                    Case "C"
                        nMontoCapital = MatrizCapitalAFecha(psCtaCod, MatCalend)
                        If nMontoPagoTmp > nMontoCapital Then
                            nMontoPagoTmp = nMontoPagoTmp - nMontoCapital
                        Else
                            nMontoCapital = nMontoPagoTmp
                            nMontoPagoTmp = 0#
                        End If
                        nMontoTotalTmp = nMontoTotalTmp + nMontoCapital
                        Call MatrizDistribuirCapital(MatCalend, MatCalendDistrib, nMontoCapital, True)
                        'Call MatrizDistribuirCapital(MatCalend, MatCalendDistrib, nMontoTotalTmp, True) 'Coment gitu 13-06-2009
                End Select
        Next J
        If (Mid(psMetLiquidacion, 3, 1) = "i" Or Mid(psMetLiquidacion, 3, 1) = "Y") And pbNoCancelar Then
            If Mid(psMetLiquidacion, 3, 1) = "Y" Then
                Call MatrizActualizarEstadoCalendCancelado(MatCalendDistrib, True, pbNoCancelar, MatCalend)
            Else
                Call MatrizActualizarEstadoCalendCancelado(MatCalendDistrib, False, pbNoCancelar, MatCalend)
            End If
        Else
               If Mid(psMetLiquidacion, 3, 1) = "Y" And pCancel = False Then
                    Call MatrizActualizarEstadoCalendCancelado(MatCalendDistrib, True, False, MatCalend)
               Else
                    If Mid(psMetLiquidacion, 3, 1) = "i" Then
                        Call MatrizActualizarEstadoCalendCancelado(MatCalendDistrib, True, pbNoCancelar, MatCalend)
                    Else
                        'Call MatrizActualizarEstadoCalendCancelado(MatCalendDistrib, , pbNoCancelar, MatCalend)
                        Call MatrizActualizarEstadoCalendCancelado(MatCalendDistrib, pbVerificaSoloCapital, pbNoCancelar, MatCalend, nMontoInteresFecha) 'ARCV 12-07-2006 + 30-07-2006
                    End If
               End If
        End If
        MatrizDistribuirCancelacion = MatCalendDistrib
End Function
Public Function MatrizDistribuirMonto(ByVal MatCalend As Variant, ByVal pnMontoPago As Double, _
                    ByVal psMetLiquidacion As String, Optional ByVal pnMontoIntFecha As Double = 0#, Optional ByVal pnMontoIntFechaGra As Double = 0#) As Variant
Dim i As Integer
Dim J As Integer
Dim MatCalendDistrib As Variant
Dim nMonto As Double
Dim nMontoTemp As Double

        nMonto = pnMontoPago
        MatCalendDistrib = CrearMatrizparaAmortizacion(MatCalend)
        Do While nMonto > 0
            For J = 1 To 4
                Select Case Mid(psMetLiquidacion, J, 1)
                    Case "G"
                        Call MatrizDistribuirGastos(MatCalend, MatCalendDistrib, nMonto)
                    Case "M"
                        Call MatrizDistribuirMora(MatCalend, MatCalendDistrib, nMonto)
                    Case "I"
                        Call MatrizDistribuirInteresI(MatCalend, MatCalendDistrib, nMonto, , pnMontoIntFechaGra)
                    Case "C"
                        Call MatrizDistribuirCapital(MatCalend, MatCalendDistrib, nMonto)
                End Select
                Call MatrizActualizarEstadoCuota(MatCalend, MatCalendDistrib)
            Next J
        Loop
        MatrizDistribuirMonto = MatCalendDistrib
End Function

Public Function MatrizSaldoCapital(ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant) As Double
Dim i As Integer
    MatrizSaldoCapital = 0
    For i = 0 To UBound(MatCalend) - 1
        MatrizSaldoCapital = MatrizSaldoCapital + (CDbl(MatCalend(i, 3)) - CDbl(MatCalendDistrib(i, 3)))
        MatrizSaldoCapital = CDbl(Format(MatrizSaldoCapital, "#0.00"))
    Next i
End Function
Public Function MatrizCuotaPendiente(ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant) As Integer
Dim i As Integer
    MatrizCuotaPendiente = 0
    For i = 0 To UBound(MatCalend) - 1
        If CInt(MatCalendDistrib(i, 2)) = 0 Then
            MatrizCuotaPendiente = CInt(MatCalendDistrib(i, 1))
            Exit For
        End If
    Next i
End Function
Public Function MatrizMontoCuotaPendiente(ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant) As Double
Dim i As Integer
    MatrizMontoCuotaPendiente = 0
    For i = 0 To UBound(MatCalend) - 1
        If CInt(MatCalendDistrib(i, 2)) = 0 Then
            MatrizMontoCuotaPendiente = CInt(MatCalendDistrib(i, 3)) + CInt(MatCalendDistrib(i, 4)) + CInt(MatCalendDistrib(i, 5)) + CInt(MatCalendDistrib(i, 6))
            Exit For
        End If
    Next i
End Function
Public Function MatrizFechaCuotaPendiente(ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant) As Date
Dim i As Integer
    For i = 0 To UBound(MatCalend) - 1
        If CInt(MatCalendDistrib(i, 2)) = 0 Then
            MatrizFechaCuotaPendiente = CDate(MatCalendDistrib(i, 0))
            Exit For
        End If
    Next i
End Function
Public Function MatrizEstadoCalendario(ByVal MatCalendDistrib As Variant) As Integer
Dim i As Integer
    MatrizEstadoCalendario = 1
    For i = 0 To UBound(MatCalendDistrib) - 1
        If CInt(MatCalendDistrib(i, 2)) = 0 Then
            MatrizEstadoCalendario = 0
            Exit For
        End If
    Next i
End Function
Private Function MatrizDistribuirGastos(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant, ByRef pnMonto As Double, _
        Optional ByVal DistVert As Boolean = False)
Dim i As Integer

    For i = 0 To UBound(MatCalend) - 1
        If i > 0 And Not DistVert Then
            'Si Aun queda pendiente capital, interes, gastos, mora de la cuota anterior
            If (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 3)) - CDbl(MatCalendDistrib(i - 1, 3))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 8)) - CDbl(MatCalendDistrib(i - 1, 8))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 7)) - CDbl(MatCalendDistrib(i - 1, 7))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 5)) - CDbl(MatCalendDistrib(i - 1, 5))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 4)) - CDbl(MatCalendDistrib(i - 1, 4))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 6)) - CDbl(MatCalendDistrib(i - 1, 6))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 And _
                (CDbl(MatCalend(i - 1, 9)) - CDbl(MatCalendDistrib(i - 1, 9))) > 0) Then

                Exit Function
            End If
        End If

        If CInt(MatCalend(i, 2)) = 0 And pnMonto > 0 And (CDbl(MatCalend(i, 9)) - CDbl(MatCalendDistrib(i, 9))) > 0 Then
            If pnMonto > (CDbl(MatCalend(i, 9)) - CDbl(MatCalendDistrib(i, 9))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 9)) - CDbl(MatCalendDistrib(i, 9)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 9) = MatCalend(i, 9)
            Else
                MatCalendDistrib(i, 9) = Format(CDbl(MatCalendDistrib(i, 9)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
            If Not DistVert Then
                Exit For
            End If
        End If
    Next i

End Function

Private Function MatrizDistribuirMora(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant, _
ByRef pnMonto As Double, Optional ByVal DistVert As Boolean = False)
Dim i As Integer

    For i = 0 To UBound(MatCalend) - 1

        If i > 0 And Not DistVert Then
            'Si Aun queda pendiente capital, interes, gastos, mora de la cuota anterior
            If (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 3)) - CDbl(MatCalendDistrib(i - 1, 3))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 8)) - CDbl(MatCalendDistrib(i - 1, 8))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 7)) - CDbl(MatCalendDistrib(i - 1, 7))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 5)) - CDbl(MatCalendDistrib(i - 1, 5))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 4)) - CDbl(MatCalendDistrib(i - 1, 4))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 6)) - CDbl(MatCalendDistrib(i - 1, 6))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 And _
                (CDbl(MatCalend(i - 1, 9)) - CDbl(MatCalendDistrib(i - 1, 9))) > 0) Then

                Exit Function
            End If
        End If

        If CInt(MatCalend(i, 2)) = 0 And pnMonto > 0 _
            And (CDbl(MatCalend(i, 6)) - CDbl(MatCalendDistrib(i, 6))) > 0 Then
            If pnMonto > (CDbl(MatCalend(i, 6)) - CDbl(MatCalendDistrib(i, 6))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 6)) - CDbl(MatCalendDistrib(i, 6)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 6) = MatCalend(i, 6)
            Else
                MatCalendDistrib(i, 6) = Format(CDbl(MatCalendDistrib(i, 6)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
    Next i
End Function

Public Function MatrizInteresTotalesAFechaSinMora(ByVal psCtaCod As String, ByVal MatCalend As Variant, ByVal pdHoy As Date, Optional ByVal pbCalenDin As Boolean = False) As Double
Dim R As ADODB.Recordset
'Dim oCredito As COMDCredito.DCOMCredito
Dim nMontoGanado As Double
Dim dFecIni As Date
Dim nSaldoCal As Double
Dim pnTasa As Double

    'Calculo de Montos Ya Ganados
    Set R = RecuperaColocaciones(psCtaCod)

    'posicion cliente
    
    dFecIni = R!dMontoColUltCal
    nSaldoCal = R!nMontoColUltCal
    
    R.Close
    Set R = Nothing

    Set R = RecuperaProducto(psCtaCod)

    pnTasa = R!nTasaInteres
    R.Close
    Set R = Nothing

    'Total Calculado a la fecha
    
    MatrizInteresTotalesAFechaSinMora = CDbl(Format(MatrizInteresReprogramadoFecha(psCtaCod, MatCalend) + _
                MatrizInteresSuspensoFecha(psCtaCod, MatCalend) + _
                MatrizInteresCompAFecha(psCtaCod, MatCalend, pdHoy, nSaldoCal, pnTasa, True, pbCalenDin) + MatrizInteresCompensatorioVencido(MatCalend), "#0.00"))
    
End Function
Private Function MatrizDistribuirInteresI(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant, _
    ByRef pnMonto As Double, Optional ByVal DistVert As Boolean = False, Optional ByRef pnMontoGra As Double)
Dim i As Integer
Dim bSiCubrio As Boolean
    bSiCubrio = False
    For i = 0 To UBound(MatCalend) - 1
        If i > 0 And Not DistVert Then
            'Si Aun queda pendiente capital, interes, gastos, mora de la cuota anterior
            If (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 3)) - CDbl(MatCalendDistrib(i - 1, 3))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 8)) - CDbl(MatCalendDistrib(i - 1, 8))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 7)) - CDbl(MatCalendDistrib(i - 1, 7))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 5)) - CDbl(MatCalendDistrib(i - 1, 5))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 4)) - CDbl(MatCalendDistrib(i - 1, 4))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 6)) - CDbl(MatCalendDistrib(i - 1, 6))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 And _
                (CDbl(MatCalend(i - 1, 9)) - CDbl(MatCalendDistrib(i - 1, 9))) > 0) Then

                Exit Function
            End If
        End If

        'Cubre Interes en Suspenso
        If CInt(MatCalend(i, 2)) = 0 And pnMonto > 0 _
            And (CDbl(MatCalend(i, 8)) - CDbl(MatCalendDistrib(i, 8))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(i, 8)) - CDbl(MatCalendDistrib(i, 8))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 8)) - CDbl(MatCalendDistrib(i, 8)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 8) = MatCalend(i, 8)
            Else
                MatCalendDistrib(i, 8) = Format(CDbl(MatCalendDistrib(i, 8)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        'Cubre Intereses Reprogramados
        If CInt(MatCalend(i, 2)) = 0 And pnMonto > 0 _
            And (CDbl(MatCalend(i, 7)) - CDbl(MatCalendDistrib(i, 7))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(i, 7)) - CDbl(MatCalendDistrib(i, 7))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 7)) - CDbl(MatCalendDistrib(i, 7)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 7) = MatCalend(i, 7)
            Else
                MatCalendDistrib(i, 7) = Format(CDbl(MatCalendDistrib(i, 7)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        'Intereses de Gracia
        'If CInt(MatCalend(i, 2)) = gColocCalendEstadoPendiente And pnMontoGra > 0 _
            And (CDbl(MatCalend(i, 5)) - CDbl(MatCalendDistrib(i, 5))) > 0 Then
        If CInt(MatCalend(i, 2)) = 0 And pnMonto > 0 _
            And (CDbl(MatCalend(i, 5)) - CDbl(MatCalendDistrib(i, 5))) > 0 Then
            bSiCubrio = True
            'Mody By gitu 07-05-2009
'            If pnMontoGra > (CDbl(MatCalend(i, 5)) - CDbl(MatCalendDistrib(i, 5))) Then
'                pnMontoGra = pnMontoGra - (CDbl(MatCalend(i, 5)) - CDbl(MatCalendDistrib(i, 5)))
'                pnMontoGra = CDbl(Format(pnMontoGra, "#0.00"))
'                MatCalendDistrib(i, 5) = MatCalend(i, 5)
'            Else
'                If pnMontoGra > pnMonto Then
'                    pnMontoGra = pnMonto
'                    MatCalendDistrib(i, 5) = Format(CDbl(MatCalendDistrib(i, 5)) + pnMontoGra, "#0.00")
'                    pnMontoGra = 0
'                    pnMonto = 0
'                Else
'                    MatCalendDistrib(i, 5) = Format(CDbl(MatCalendDistrib(i, 5)) + pnMontoGra, "#0.00")
'                    pnMontoGra = 0
'                End If
'
'            End If
            If pnMonto > (CDbl(MatCalend(i, 5)) - CDbl(MatCalendDistrib(i, 5))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 5)) - CDbl(MatCalendDistrib(i, 5)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 5) = MatCalend(i, 5)
            Else
                MatCalendDistrib(i, 5) = Format(CDbl(MatCalendDistrib(i, 5)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        'Intereses Compensatorio Vencido
        If CInt(MatCalend(i, 2)) = 0 And pnMonto > 0 _
            And (CDbl(MatCalend(i, 11)) - CDbl(MatCalendDistrib(i, 11))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(i, 11)) - CDbl(MatCalendDistrib(i, 11))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 11)) - CDbl(MatCalendDistrib(i, 11)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 11) = MatCalend(i, 11)
            Else
                MatCalendDistrib(i, 11) = Format(CDbl(MatCalendDistrib(i, 11)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        'Intereses Compensatorios
        If CInt(MatCalend(i, 2)) = 0 And pnMonto > 0 _
            And (CDbl(MatCalend(i, 4)) - CDbl(MatCalendDistrib(i, 4))) > 0 Then
            bSiCubrio = True
            If pnMonto > (CDbl(MatCalend(i, 4)) - CDbl(MatCalendDistrib(i, 4))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 4)) - CDbl(MatCalendDistrib(i, 4)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 4) = MatCalend(i, 4)
            Else
                MatCalendDistrib(i, 4) = Format(CDbl(MatCalendDistrib(i, 4)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
        End If
        If bSiCubrio Then
            If Not DistVert Then
                Exit For
            End If
        End If
    Next i
End Function

Private Function MatrizDistribuirCapital(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant, _
    ByRef pnMonto As Double, Optional ByVal DistVert As Boolean = False)
Dim i As Integer

    For i = 0 To UBound(MatCalend) - 1

        If i > 0 And Not DistVert Then
            'Si Aun queda pendiente capital, interes, gastos, mora de la cuota anterior
            If (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 3)) - CDbl(MatCalendDistrib(i - 1, 3))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 8)) - CDbl(MatCalendDistrib(i - 1, 8))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 7)) - CDbl(MatCalendDistrib(i - 1, 7))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 5)) - CDbl(MatCalendDistrib(i - 1, 5))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 4)) - CDbl(MatCalendDistrib(i - 1, 4))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 _
                And (CDbl(MatCalend(i - 1, 6)) - CDbl(MatCalendDistrib(i - 1, 6))) > 0) Or _
                (CInt(MatCalend(i - 1, 2)) = 0 And pnMonto > 0 And _
                (CDbl(MatCalend(i - 1, 9)) - CDbl(MatCalendDistrib(i - 1, 9))) > 0) Then

                Exit Function
            End If
        End If

        If CInt(MatCalend(i, 2)) = 0 And pnMonto > 0 _
            And (CDbl(MatCalend(i, 3)) - CDbl(MatCalendDistrib(i, 3))) > 0 Then
            If pnMonto > (CDbl(MatCalend(i, 3)) - CDbl(MatCalendDistrib(i, 3))) Then
                pnMonto = pnMonto - (CDbl(MatCalend(i, 3)) - CDbl(MatCalendDistrib(i, 3)))
                pnMonto = CDbl(Format(pnMonto, "#0.00"))
                MatCalendDistrib(i, 3) = MatCalend(i, 3)
            Else
                MatCalendDistrib(i, 3) = Format(CDbl(MatCalendDistrib(i, 3)) + pnMonto, "#0.00")
                pnMonto = 0
            End If
            If Not DistVert Then
                Exit For
            End If
        End If
    Next i
End Function
Private Sub MatrizActualizarEstadoCalendCancelado(ByRef MatCalendDistrib As Variant, Optional ByVal pbVerficaSolosCapital As Boolean = False, _
    Optional ByVal pbNoCancelar As Boolean = True, Optional ByVal MatCalend As Variant = "", _
    Optional ByVal pnMontoInteresFecha As Double = -1)  'ARCV 30-07-2006 (pnInteresAFecha)

Dim i As Integer

    For i = 0 To UBound(MatCalendDistrib) - 1
        If pbNoCancelar Then

            If pbVerficaSolosCapital Then
                If CInt(MatCalendDistrib(i, 2)) = 0 Then
                    If CDbl(MatCalend(i, 3)) = CDbl(MatCalendDistrib(i, 3)) Then
                        MatCalendDistrib(i, 2) = Trim(Str(1))
                    Else
                        MatCalendDistrib(i, 2) = Trim(Str(0))
                    End If
                End If
            Else
                If pnMontoInteresFecha > 0 Then 'ARCV 30-07-2006
                    'CDbl(MatCalend(i, 5)) = CDbl(MatCalendDistrib(i, 5)) And _ 'ARCV 26-06-2007 (Interes a la Fecha ya incluye el Periodo de Gracia)

                    If CDbl(MatCalend(i, 8)) = CDbl(MatCalendDistrib(i, 8)) And _
                        CDbl(MatCalend(i, 7)) = CDbl(MatCalendDistrib(i, 7)) And _
                        CDbl(MatCalend(i, 6)) = CDbl(MatCalendDistrib(i, 6)) And _
                        CDbl(MatCalend(i, 3)) = CDbl(MatCalendDistrib(i, 3)) Then
                        MatCalendDistrib(i, 2) = Trim(Str(1))
                    End If
                Else
                    'ARCV 30-07-2006 (Solo Capital...los interes son 0)
                    If CDbl(MatCalend(i, 3)) = CDbl(MatCalendDistrib(i, 3)) Then
                        MatCalendDistrib(i, 2) = Trim(Str(1))
                    End If
                End If
            End If
        Else
            If Not pbVerficaSolosCapital Then
                If CInt(MatCalendDistrib(i, 2)) = 0 Then
                     '3 Capital 9 Gastos 6 Mora 4 int.Compensatorio
                    If CDbl(MatCalend(i, 3)) = CDbl(MatCalendDistrib(i, 3)) And _
                        CDbl(MatCalend(i, 9)) = CDbl(MatCalendDistrib(i, 9)) And _
                        CDbl(MatCalend(i, 6)) = CDbl(MatCalendDistrib(i, 6)) And _
                        CDbl(MatCalend(i, 4)) = CDbl(MatCalendDistrib(i, 4)) Then
                        MatCalendDistrib(i, 2) = Trim(Str(1))
                    Else
                        MatCalendDistrib(i, 2) = Trim(Str(0))
                    End If
                End If
             Else
             
                If CInt(MatCalendDistrib(i, 2)) = 0 Then
                    If CDbl(MatCalend(i, 3)) = CDbl(MatCalendDistrib(i, 3)) Then
                        MatCalendDistrib(i, 2) = Trim(Str(1))
                    Else
                        MatCalendDistrib(i, 2) = Trim(Str(0))
                    End If
                End If
                             
            End If
        End If
    Next i
End Sub

Private Sub MatrizActualizarEstadoCuota(ByVal MatCalend As Variant, ByRef MatCalendDistrib As Variant)
Dim i As Integer
    For i = 0 To UBound(MatCalend) - 1
        If CInt(MatCalendDistrib(i, 2)) = 0 Then
                If CDbl(MatCalend(i, 8)) = CDbl(MatCalendDistrib(i, 8)) And _
                    CDbl(MatCalend(i, 7)) = CDbl(MatCalendDistrib(i, 7)) And _
                    CDbl(MatCalend(i, 6)) = CDbl(MatCalendDistrib(i, 6)) And _
                    CDbl(MatCalend(i, 5)) = CDbl(MatCalendDistrib(i, 5)) And _
                    CDbl(MatCalend(i, 4)) = CDbl(MatCalendDistrib(i, 4)) And _
                    CDbl(MatCalend(i, 3)) = CDbl(MatCalendDistrib(i, 3)) And CDbl(MatCalend(i, 11)) = CDbl(MatCalendDistrib(i, 11)) _
                    And CDbl(MatCalend(i, 9)) = CDbl(MatCalendDistrib(i, 9)) Then   'Faltaba el Concepto de Gastos
                    MatCalendDistrib(i, 2) = Trim(Str(1))
                End If
        End If
    Next i
End Sub

Public Function AmortizarCredito(ByVal psCtaCod, ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant, _
            ByVal pnMonto As Double, ByVal pdHoy As Date, ByVal psMetLiquid As String, _
            ByVal pnTipoPago As Integer, ByVal psCodAge As String, ByVal psCodUser As String, _
            Optional psNroDoc As String = "", _
            Optional ByRef pnMovNro As Long = -1, Optional ByVal pnNroDacion As Integer = -1, Optional pbEnOtraCmac As Boolean = False, _
            Optional psPersCmac As String = "", Optional ByVal pnIntPend As Double = 0, Optional ByVal pnIntPendPagado As Double = 0, _
            Optional psMovnroTemp As String = "", Optional ByVal pMatGastosGen As Variant = "", Optional ByVal pnNumGastosGen As Integer = -1, _
            Optional ByVal MatCalendDistribParalelo As Variant = "", Optional ByVal pnCalPago As Integer = 1, Optional ByVal MatCalendParalelo As Variant = "", _
            Optional ByVal pnPrepago As Integer = 0, Optional psPersLavDinero As String = "", Optional pnITF As Double = 0#, Optional ByVal pnMontoDesagio As Double = 0, _
            Optional ByVal pbInicioTrans As Boolean = False, Optional ByRef psMensajeValidacion As String = "", _
            Optional ByVal psProyectoActual As String = "", _
            Optional ByVal psTitLavDinero As String = "", _
            Optional ByVal psOrdLavDinero As String = "", _
            Optional ByVal psReaLavDinero As String = "", _
            Optional ByVal psBenLavDinero As String = "", _
            Optional ByVal psVisLavDinero As String = "", _
            Optional pnMovNroTem As Long = 0) As String  'DAOR 20070511,psVisLavDinero
            
                                

'By Capi 20012008 se agrego los ultimos 5 parametros

'Dim oBase As COMDCredito.DCOMCredActBD
Dim nEstadoCred As Integer
Dim nEstadoCredTemp As Integer
Dim R As ADODB.Recordset
'Dim oCred As COMDCredito.DCOMCredito
'Dim oCalend As COMDCredito.DCOMCalendario
Dim nTransacc As Long
Dim sLineaCred As String
Dim nMontoColocado As Double
Dim dFecPend As Date
Dim nDiasAtraso As Integer
Dim i, k As Integer
Dim nNroCalen As Integer
Dim nMontoGasto As Double
Dim sMovNro As String
Dim nMovNro As Long
Dim nMovNroOffHost As Long
Dim nConsCred As String
Dim pnPlazo As Integer
Dim bTran As Boolean
'Dim oFunciones As New COMNContabilidad.NCOMContFunciones
Dim dFechaTran As Date
Dim nIntPend As Double
Dim nMontoPago As Double
Dim nMontoPago_2 As Double
Dim nNroCalPar As Integer
Dim nMivivienda As Integer
Dim MatGastosCred As Variant
Dim NumregGastosCred As Integer
Dim MatGastosCuota As Variant
Dim NumRegGastosCuota As Integer
Dim nBuenPagador As Integer
Dim nPrestamo As Double
Dim CapitalPagado As Double
Dim nDiasAtrasoMov As Integer
Dim opeITFChequeEfect As String
'Se Agrego para el Manejo de las Operaciones VAC
Dim nCapitalVAC As Double
nCapitalVAC = 0
'************************

    On Error GoTo ErrorAmortizarPago


    'opeITFChequeEfect = "990103"

    Set R = RecuperaProducto(psCtaCod)

    nEstadoCred = R!nPrdEstado
    nEstadoCredTemp = R!nPrdEstado
    nTransacc = R!nTransacc
    R.Close
    Set R = Nothing
    
    Set R = RecuperaColocaciones(psCtaCod)
    
    nPrestamo = R!nMontoCol
    R.Close
    Set R = Nothing

'    If pnTipoPago <> gColocTipoPagoDacionPago Then
'        If pnTipoPago <> gColocTipoPagoCargoCta Then
'            'Definir Codigo de Operacion de Pago
'            Select Case nEstadoCred
'                'Si es Credito refinanciado
'                Case gColocEstRefMor
'                    nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefMorEfec, gCredPagRefMorCh)
'                Case gColocEstRefNorm
'                    nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefNorEfec, gCredPagRefNorCh)
'                Case gColocEstRefVenc
'                    nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagRefVenEfec, gCredPagRefVenCh)
'                'si es Credito Normal
'                Case gColocEstVigMor
'                    nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorMorEfec, gCredPagNorMorCh)
'                Case gColocEstVigNorm
'                    nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorNorEfec, gCredPagNorNorCh)
'                Case gColocEstVigVenc
'                    nConsCred = IIf(pnTipoPago = gColocTipoPagoEfectivo, gCredPagNorVenEfec, gCredPagNorVenCh)
'            End Select
'        Else
'            Select Case nEstadoCred
'            'Si es Credito refinanciado
'            Case gColocEstRefMor
'                nConsCred = gCredPagRefMorCC
'            Case gColocEstRefNorm
'                nConsCred = gCredPagRefNorCC
'            Case gColocEstRefVenc
'                nConsCred = gCredPagRefVenCC
'            'si es Credito Normal
'            Case gColocEstVigMor
'                nConsCred = gCredPagNorMorCC
'            Case gColocEstVigNorm
'                nConsCred = gCredPagNorNorCC
'            Case gColocEstVigVenc
'                nConsCred = gCredPagNorVenCC
'            End Select
'        End If
'    Else
'        'Definir Codigo de Operacion de Pago
'        Select Case nEstadoCred
'            'Si es Credito refinanciado
'            Case gColocEstRefMor
'                nConsCred = gCredPagRefMorDacion
'            Case gColocEstRefNorm
'                nConsCred = gCredPagRefNorDacion
'            Case gColocEstRefVenc
'                nConsCred = gCredPagRefVenDacion
'            'si es Credito Normal
'            Case gColocEstVigMor
'                nConsCred = gCredPagNorMorDacion
'            Case gColocEstVigNorm
'                nConsCred = gCredPagNorNorDacion
'            Case gColocEstVigVenc
'                nConsCred = gCredPagNorVenDacion
'        End Select
'    End If
    
    nConsCred = gPITColocPagoCredito
    
    If nConsCred = "" Then
        psMensajeValidacion = "No se ha definido la operación correctamente. Consulte con la Oficina de T.I"
        Exit Function
    End If

    Set R = RecuperaColocacEstado(psCtaCod, gColocEstAprob)

    pnPlazo = IIf(IsNull(R!nPlazo), 0, R!nPlazo)

    R.Close
    Set R = Nothing

    Set R = RecuperaColocacCred(psCtaCod)

    'Manejo de Operaciones VAC
    Dim bVAC As Boolean
    bVAC = IIf(IsNull(R!bVAC), False, R!bVAC)
    '**********************
    nNroCalen = R!nNroCalen
    nNroCalPar = R!nNroCalPar
    nMivivienda = IIf(IsNull(R!bMiVivienda), 0, R!bMiVivienda)
    nBuenPagador = IIf(IsNull(R!nCalPago), 0, R!nCalPago)
    nDiasAtrasoMov = IIf(IsNull(R!nDiasAtraso), 0, R!nDiasAtraso)
    R.Close
    Set R = Nothing

    'Set oBase = New COMDCredito.DCOMCredActBD

    If psMovnroTemp <> "" Then
        sMovNro = psMovnroTemp
    Else
        'sMovNro = GeneraMovNro(pdHoy, psCodAge, psCodUser)
        sMovNro = gsMovNro
    End If

    dFechaTran = CDate(Format(Format(pdHoy, "dd/mm/yyyy") & " " & Format(dFechaHora, "hh:mm:ss"), "dd/mm/yyyy hh:mm:ss"))

    'Actualiza Producto
    If MatrizEstadoCalendario(MatCalendDistrib) = 1 Then
        nEstadoCred = 2050
    End If
    
    Set R = RecuperaColocaciones(psCtaCod)

    sLineaCred = R!cLineaCred
    R.Close
    Set R = Nothing

    'Actualiza ColocacCred
    dFecPend = MatrizFechaCuotaPendiente(MatCalend, MatCalendDistrib)
    
    If MatrizCuotaPendiente(MatCalend, MatCalendDistrib) = 0 Then
        nDiasAtraso = 0
    Else
        nDiasAtraso = pdHoy - dFecPend
    End If

    nIntPend = pnIntPend - pnIntPendPagado
    
    'Call dInsertMovOffHost(sMovNro, nConsCred, "", 10, 0, False)
    'nMovNroOffHost = dGetnMovNroOffHost(sMovNro)
    
    Call dInsertMov(sMovNro, nConsCred, "", 10, 0, False)
    nMovNro = dGetnMovNro(sMovNro)

    pnMovNroTem = nMovNro

    'Call dInsertMovCMAC(nMovNro, psPersCmac, Format$(gTpoIFCmac, "00"), CInt(Mid(psCtaCod, 9, 1)), "", psNroDoc, nConsCred, pnMonto, False)

    If pnTipoPago <> gColocTipoPagoDacionPago Then
        'Call dInsertMovColOffHost(nMovNroOffHost, nConsCred, psCtaCod, nNroCalen, pnMonto, nDiasAtrasoMov, psMetLiquid, pnPlazo, MatrizSaldoCapital(MatCalend, MatCalendDistrib), nEstadoCredTemp, False, , pnPrepago)
        Call dInsertMovCol(nMovNro, nConsCred, psCtaCod, nNroCalen, pnMonto, nDiasAtrasoMov, psMetLiquid, pnPlazo, MatrizSaldoCapital(MatCalend, MatCalendDistrib), nEstadoCredTemp, False, , pnPrepago)
    Else
        'Call dInsertMovColOffHost(nMovNroOffHost, nConsCred, psCtaCod, nNroCalen, pnMonto, nDiasAtrasoMov, psMetLiquid, pnPlazo, MatrizSaldoCapital(MatCalend, MatCalendDistrib), nEstadoCredTemp, False, pnNroDacion, pnPrepago)
        Call dInsertMovCol(nMovNro, nConsCred, psCtaCod, nNroCalen, pnMonto, nDiasAtrasoMov, psMetLiquid, pnPlazo, MatrizSaldoCapital(MatCalend, MatCalendDistrib), nEstadoCredTemp, False, pnNroDacion, pnPrepago)
    End If

    '*********  ITF  *****************

    'Call dInsertMovColOffHost(nMovNroOffHost, opeITFChequeEfect, psCtaCod, CLng(nNroCalen), pnITF, 0, "", 0, 0#, nEstadoCredTemp, False)
    Call dInsertMovCol(nMovNro, sOpeCodITF, psCtaCod, CLng(nNroCalen), pnITF, 0, "", 0, 0#, nEstadoCredTemp, False)
    'Call dInsertMovColDetOffHost(nMovNroOffHost, opeITFChequeEfect, psCtaCod, CLng(nNroCalen), 20, 0, pnITF, False)
    Call dInsertMovColDet(nMovNro, sOpeCodITF, psCtaCod, CLng(nNroCalen), 20, 0, pnITF, False)
    
    'Carga Gastos en Memoria para Evitar Bloqueo
    MatGastosCred = DevuelveMatrizGastosCredito(NumregGastosCred, psCtaCod, nNroCalen)

    'Actualiza calendario (ColocCalendario y ColocCalendDet)
    For i = 0 To UBound(MatCalendDistrib) - 1
        If MatrizMontoPagado(MatCalendDistrib, CInt(MatCalendDistrib(i, 1))) > 0 Then
            Call dUpdateColocCalendario(psCtaCod, nNroCalen, CInt(MatCalendDistrib(i, 1)), 1, , CInt(MatCalendDistrib(i, 2)), "Pago de Cuota", 2, False, , dFechaTran)
        Else
            Call dUpdateColocCalendario(psCtaCod, nNroCalen, CInt(MatCalendDistrib(i, 1)), 1, , CInt(MatCalendDistrib(i, 2)), "Pago de Cuota", 2, False)
        End If

        'Amortizando Capital
        If CDbl(MatCalendDistrib(i, 3)) > 0 Then
            Call dUpdateColocCalendDet(psCtaCod, nNroCalen, 1, CInt(MatCalendDistrib(i, 1)), 1000, , CDbl(MatCalendDistrib(i, 3)), , False, True)
            'Inserta Detalle Movimiento Capital
            'Call dInsertMovColDetOffHost(nMovNroOffHost, nConsCred, psCtaCod, CLng(nNroCalen), 1000, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 3)), False)
            Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), 1000, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 3)), False)
        End If
        'Amortizando Interes Compensatorio
        If CDbl(MatCalendDistrib(i, 4)) > 0 Then
            Call dUpdateColocCalendDet(psCtaCod, nNroCalen, 1, CInt(MatCalendDistrib(i, 1)), 1100, , CDbl(MatCalendDistrib(i, 4)), , False, True)
            'Inserta Detalle Movimiento Interes Compensatorio
            'Call dInsertMovColDetOffHost(nMovNroOffHost, nConsCred, psCtaCod, CLng(nNroCalen), 1100, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 4)), False)
            Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), 1100, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 4)), False)
        End If
        'Amortizando Interes Gracia
        If CDbl(MatCalendDistrib(i, 5)) > 0 Then
            Call dUpdateColocCalendDet(psCtaCod, nNroCalen, 1, CInt(MatCalendDistrib(i, 1)), 1102, , CDbl(MatCalendDistrib(i, 5)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            'Call dInsertMovColDetOffHost(nMovNroOffHost, nConsCred, psCtaCod, CLng(nNroCalen), 1102, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 5)), False)
            Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), 1102, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 5)), False)
        End If
        'Amortizando Interes Moratorio
        If CDbl(MatCalendDistrib(i, 6)) > 0 Then
            Call dUpdateColocCalendDet(psCtaCod, nNroCalen, 1, CInt(MatCalendDistrib(i, 1)), 1101, , CDbl(MatCalendDistrib(i, 6)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            'Call dInsertMovColDetOffHost(nMovNroOffHost, nConsCred, psCtaCod, CLng(nNroCalen), 1101, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 6)), False)
            Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), 1101, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 6)), False)
        End If
        'Amortizando Interes Reprog
        If CDbl(MatCalendDistrib(i, 7)) > 0 Then
            Call dUpdateColocCalendDet(psCtaCod, nNroCalen, 1, CInt(MatCalendDistrib(i, 1)), 1103, , CDbl(MatCalendDistrib(i, 7)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            Call dInsertMovColDetOffHost(nMovNroOffHost, nConsCred, psCtaCod, CLng(nNroCalen), 1103, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 7)), False)
            Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), 1103, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 7)), False)
        End If
        'Amortizando Interes Suspenso
        If CDbl(MatCalendDistrib(i, 8)) > 0 Then
            Call dUpdateColocCalendDet(psCtaCod, nNroCalen, 1, CInt(MatCalendDistrib(i, 1)), 1104, , CDbl(MatCalendDistrib(i, 8)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            'Call dInsertMovColDetOffHost(nMovNroOffHost, nConsCred, psCtaCod, CLng(nNroCalen), 1104, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 7)), False)
            Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), 1104, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 7)), False)
        End If

        'Amortizando Interes Compensatorio Vencido
        If CDbl(MatCalendDistrib(i, 11)) > 0 Then
            Call dUpdateColocCalendDet(psCtaCod, nNroCalen, 1, CInt(MatCalendDistrib(i, 1)), 1105, , CDbl(MatCalendDistrib(i, 11)), , False, True)
            'Inserta Detalle Movimiento Interes Gracia
            'Call dInsertMovColDetOffHost(nMovNroOffHost, nConsCred, psCtaCod, CLng(nNroCalen), 1105, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 11)), False)
            Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), 1105, CInt(MatCalendDistrib(i, 1)), CDbl(MatCalendDistrib(i, 11)), False)
        End If

        'Amortizando Gastos
        If CDbl(MatCalendDistrib(i, 9)) > 0 Then
            nMontoGasto = CDbl(MatCalendDistrib(i, 9))

            MatGastosCuota = DevuelveMatrizGastosCreditoCuota(NumRegGastosCuota, CInt(MatCalendDistrib(i, 1)), MatGastosCred, NumregGastosCred)

            For k = 0 To NumRegGastosCuota - 1
                If nMontoGasto >= CDbl(MatGastosCuota(k, 2)) Then
                    Call dUpdateColocCalendDet(psCtaCod, nNroCalen, 1, CInt(MatCalendDistrib(i, 1)), CLng(MatGastosCuota(k, 1)), , CDbl(MatGastosCuota(k, 2)), , False, True)
                    'Inserta Detalle Movimiento Gastos
                    'Call dInsertMovColDetOffHost(nMovNroOffHost, nConsCred, psCtaCod, CLng(nNroCalen), CLng(MatGastosCuota(k, 1)), CInt(MatCalendDistrib(i, 1)), CDbl(MatGastosCuota(k, 2)), False)
                    Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), CLng(MatGastosCuota(k, 1)), CInt(MatCalendDistrib(i, 1)), CDbl(MatGastosCuota(k, 2)), False)
                    nMontoGasto = nMontoGasto - CDbl(MatGastosCuota(k, 2))
                Else
                    Call dUpdateColocCalendDet(psCtaCod, nNroCalen, 1, CInt(MatCalendDistrib(i, 1)), CLng(MatGastosCuota(k, 1)), , nMontoGasto, , False, True)
                    'Inserta Detalle Movimiento Gastos
                    'Call dInsertMovColDetOffHost(nMovNroOffHost, nConsCred, psCtaCod, CLng(nNroCalen), CLng(MatGastosCuota(k, 1)), CInt(MatCalendDistrib(i, 1)), CDbl(Format(nMontoGasto, "#0.00")), False)
                    Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), CLng(MatGastosCuota(k, 1)), CInt(MatCalendDistrib(i, 1)), CDbl(Format(nMontoGasto, "#0.00")), False)
                    nMontoGasto = 0
                End If

                nMontoGasto = CDbl(Format(nMontoGasto, "#0.00"))
                If nMontoGasto = 0 Then
                    Exit For
                End If
            Next k
            ' R.Close
             Set R = Nothing
        End If
   
    Next i

    'Amortizando Desagio si Hubiere
    If pnMontoDesagio > 0 Then
        'Call dInsertMovColDetOffHost(nMovNroOffHost, nConsCred, psCtaCod, CLng(nNroCalen), 1106, CInt(MatCalendDistrib(0, 1)), pnMontoDesagio, False)
        Call dInsertMovColDet(nMovNro, nConsCred, psCtaCod, CLng(nNroCalen), 1106, CInt(MatCalendDistrib(0, 1)), pnMontoDesagio, False)
    End If

    If pnTipoPago = gColocTipoPagoDacionPago Then
        Call dAnularColocGarantRec(pnNroDacion, 2, False)
    End If

    CapitalPagado = 0
    CapitalPagado = CapitalPagado + MatrizCapitalPagado(MatCalendDistrib)

    ''''''''''''''''''''
    ' Reversion de Garantia
    ''''''''''''''''''''
    Call LiberaGarantiaPago(nMovNro, psCtaCod, psCodAge, psCodUser, pdHoy, nPrestamo, CapitalPagado, IIf(nEstadoCred = gColocEstCancelado, True, False))

    ''''''''''''''''''''

    Exit Function

ErrorAmortizarPago:
    If bTran Then
        'Call oBase.dRollbackTrans
        'Set oBase = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function GeneraMovNro(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String
    On Error GoTo GeneraMovNroErr
    Dim rs As ADODB.Recordset
    Dim coConex As New DConecta
    Dim sql As String
    
    Set rs = New ADODB.Recordset
    coConex.AbreConexion
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       sql = "DBCMAC..sp_GeneraMovNro '" & Format(pdFecha, "mm/dd/yyyy hh:mm:ss") & "','" & Right(psCodAge, 2) & "','" & psUser & "'"
    Else
       sql = "DBCMAC..sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    
    Set rs = coConex.Ejecutar(sql)
    If Not rs.EOF Then
        GeneraMovNro = rs.Fields(0)
    End If
    coConex.CierraConexion
    rs.Close
    Set rs = Nothing
    Exit Function
GeneraMovNroErr:
    
End Function
Public Function dFechaHora(Optional psFecha As String = "") As Date
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConecta As New DConecta

    On Error GoTo ErrordFechaHora
    oConecta.AbreConexion
    sSQL = "Select GETDATE() as FechaHora"
    Set R = New ADODB.Recordset
    R.Open sSQL, oConecta.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    If psFecha = "" Then
        dFechaHora = R!FechaHora
    Else
        dFechaHora = CDate(Format(Format(psFecha, "dd/mm/yyyy") & " " & Format(R!FechaHora, "hh:mm:ss"), "dd/mm/yyyy"))
    End If
    R.Close
    oConecta.CierraConexion
    Set R = Nothing
    Exit Function

ErrordFechaHora:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
Public Sub dUpdateProducto(ByVal psCtaCod As String, Optional ByVal pnTasaInteres As Double = -1, Optional ByVal pnSaldo As Double = -1, Optional ByVal pnPrdEstado As Integer = -1, Optional ByVal pdPrdEstado As Date = CDate("01/01/1950"), Optional ByVal pnTransacc As Long = -1, _
    Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pnExtorno As Integer = 0, Optional ByVal bIncrementarSaldo As Boolean = False)

Dim coConex As New DConecta
Dim lsSql As String


    On Error GoTo ErrordInsertProducto
    
    coConex.AbreConexion
    lsSql = " UPDATE DBCMAC..Producto Set "
    If pnTasaInteres <> -1 Then
        lsSql = lsSql & " nTasaInteres = " & Format(pnTasaInteres, "#0.0000") & ","
    End If
    If pnExtorno = 0 Then
        If pnSaldo <> -1 And Not bIncrementarSaldo Then
            lsSql = lsSql & " nSaldo = " & pnSaldo & ","
        End If
    Else
        If pnSaldo <> -1 And Not bIncrementarSaldo Then
            lsSql = lsSql & " nSaldo = nSaldo - " & pnSaldo & ","
        End If
    End If
    
    If bIncrementarSaldo Then
        If pnSaldo <> -1 Then
            lsSql = lsSql & " nSaldo = nSaldo + " & pnSaldo & ","
        End If
    End If
    If pnPrdEstado <> -1 Then
        lsSql = lsSql & " nPrdEstado = " & pnPrdEstado & ","
    End If
    If pdPrdEstado <> CDate("01/01/1950") Then
        lsSql = lsSql & " dPrdEstado = '" & Format(pdPrdEstado, "mm/dd/yyyy") & "',"
    End If
    If pnExtorno = 0 Then
        If pnTransacc = -2 Then ' Si se le envia (-2) aumenta el nro de transaccion en uno (LAYG)
            lsSql = lsSql & " nTransacc = nTransacc + 1,"
        ElseIf pnTransacc <> -1 Then
            lsSql = lsSql & " nTransacc = " & pnTransacc & ","
        End If
    Else
        If pnTransacc = -2 Then ' Si se le envia (-2) aumenta el nro de transaccion en uno (LAYG)
            lsSql = lsSql & " nTransacc = nTransacc + 1,"
        ElseIf pnTransacc <> -1 Then
            lsSql = lsSql & " nTransacc = nTransacc - 1, "
        End If
    End If
    lsSql = Mid(lsSql, 1, Len(lsSql) - 1)
    lsSql = lsSql & " WHERE cCtaCod = '" & psCtaCod & "'"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSql
    Else
        coConex.Ejecutar lsSql
    End If
    coConex.CierraConexion
    Exit Sub

ErrordInsertProducto:
                  Err.Raise Err.Number, "Error En Proceso dInsertProducto", Err.Description

End Sub

Public Sub dUpdateColocacCred(ByVal psCtaCod As String, Optional ByVal pnDiasAtraso As Integer = -1, Optional ByVal pnColocCondicion As Integer = -1, _
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
    Optional ByVal pnNumConMic As Integer = -1) 'DAOR 20061216: Se aumentó el parametro Numero de Consultas a Certicom
    'DAOR 20070419, Se agregó parametro pnTasCosEfeAnu:Tasa Costo Efectiva Anual as
    'By Capi 10042008 se insertó parametro psCtaCodAho:Cuenta Ahorro Desembolso
    'By Capi 14042008 se insertó parametro psMovNro: Ultima fecha y usuario que hizo la reprogramacion
    'Add Gitu 06042009 se agrego parametro para la actualizacion de la campaña
    'Add Gitu 20-05-2009 se agrego el parametro de consulta Score Microfinanzas
Dim coConex As New DConecta
Dim lsSql As String
    On Error GoTo ErrordUpdateColocacCred
    
    coConex.AbreConexion
    lsSql = "UPDATE DBCMAC..ColocacCred SET "
    
       
    If pnMoneda <> -1 Then
        lsSql = lsSql & " nFondoMoneda = " & pnMoneda & ","
    End If
    If pnCalifPagos <> -1 Then
        lsSql = lsSql & " nCalPago = " & pnCalifPagos & ","
    End If
    If pnCuotaComodin <> -1 Then
        lsSql = lsSql & " bCuotaCom = " & pnCuotaComodin & ","
    End If
    If pnMiViv <> -1 Then
        lsSql = lsSql & " bMiVivienda = " & pnMiViv & ","
    End If
    If pnNroCalParalelo <> -1 Then
        lsSql = lsSql & " nNroCalPar = " & pnNroCalParalelo & ","
    End If
    If psPersFondo <> "" Then
        lsSql = lsSql & " cPersCod = '" & psPersFondo & "',"
    End If
    If pnNroProxDesemb <> -1 Then
        lsSql = lsSql & " nNroProxDesemb = " & pnNroProxDesemb & ","
    End If
    If pnDiasAtraso <> -1 Then
        lsSql = lsSql & " nDiasAtraso = " & pnDiasAtraso & ","
    End If
    If pnColocCondicion <> -1 Then
        lsSql = lsSql & " nColocCondicion = " & pnColocCondicion & ","
    End If
    If pnColocDestino <> -1 Then
        lsSql = lsSql & " nColocDestino = " & pnColocDestino & ","
    End If
    If psProtesto <> "" Then
        lsSql = lsSql & " cProtesto = '" & psProtesto & "',"
    End If
    If pbCargoAuto <> -1 Then
        lsSql = lsSql & " bCargoAuto = " & pbCargoAuto & ","
    End If
    If psMetLiquidacion <> "" Then
        lsSql = lsSql & " cMetLiquidacion = '" & psMetLiquidacion & "',"
    End If
    If pbRefCapInt <> -1 Then
        lsSql = lsSql & " bRefCapInt = " & pbRefCapInt & ","
    End If
    If pnNroProxCuota <> -1 Then
        lsSql = lsSql & " nNroProxCuota = " & pnNroProxCuota & ","
    End If
    If pnIntPend <> -1 Then
        lsSql = lsSql & " nIntPend = " & Format(pnIntPend, "#0.00") & ","
    End If
    If pnExoPenalidad <> -1 Then
        lsSql = lsSql & " nExoPenalidad = " & pnExoPenalidad & ","
    End If
    If pnTpoCalend <> -1 Then
        lsSql = lsSql & " nColocCalendCod = " & pnTpoCalend & ","
    End If
    If pnNroCalen <> -1 Then
        lsSql = lsSql & " nNroCalen = " & pnNroCalen & ","
    End If
    If pnCalendDinamico <> -1 Then
        lsSql = lsSql & " nCalendDinamico = " & pnCalendDinamico & ","
    End If
    If pnTipoDesembolso <> -1 Then
        lsSql = lsSql & " nTipoDesembolso = " & pnTipoDesembolso & ","
    End If
    If pbPrepago <> -1 Then
        lsSql = lsSql & " bPrepago = " & pbPrepago & ","
    End If
    If pnCalendDinamTipo <> -1 Then
        lsSql = lsSql & " nCalendDinamTipo = " & pnCalendDinamTipo & ","
    End If
    If pbBloqueo <> -1 Then
        lsSql = lsSql & " bBloqueo = " & pbBloqueo & ","
    End If
    'Add By Gitu 06-04-2009 solo para aprobacion
    If pnIdCamp <> -1 Then
        lsSql = lsSql & " IdCampana = '" & pnIdCamp & "',"
    End If
    'Operaciones VAC
    If pnVAC <> -1 Then
        lsSql = lsSql & " bVAC=" & pnVAC & ","
    End If
    
    'If pnAproReglamento <> 0 Then
        lsSql = lsSql & " nAproReglamento =" & pnAproReglamento & ","
    'End If
 
    'If pnExoSeguroDes <> 0 Then
        lsSql = lsSql & " nExoSeguroDes =" & pnExoSeguroDes & ","
    'End If
    
    '**DAOR 20061216, Actualizar nNumConCer: Numero de Consultas a la central de Riesgos
    If pnNumConCer <> -1 Then
        lsSql = lsSql & " nNumConCer = " & pnNumConCer & ","
    End If
    '************************************************************************************
    'Add By Gitu 20-05-2009 Actualiza numero de consultas Score Microfinanzas
    If pnNumConMic <> -1 Then
        lsSql = lsSql & " nNumConMic = " & pnNumConMic & ","
    End If
    '-------------------------------------------------
    
    '**DAOR 20070419, Actualizar nTasCosEfeAnu: Tasa Costo Efectivo Anual
    If pnTasCosEfeAnu <> 0 Then
        lsSql = lsSql & " nTasCosEfeAnu = " & pnTasCosEfeAnu & ","
    End If
    '************************************************************************************
    'By Capi 10042008 para que actualice la Cuenta Ahorro Desembolso cuando corresponda
    If psMovNro <> "" Then
        lsSql = lsSql & " cMovNroRprg = '" & psMovNro & "',"
    End If
    If psCtaCodAho <> "" Then
        lsSql = lsSql & " cCtaCodAho = '" & psCtaCodAho & "',"
    End If
    '************************************************************************************

    lsSql = Mid(lsSql, 1, Len(lsSql) - 1)
    lsSql = lsSql & " WHERE cCtaCod = '" & psCtaCod & "'"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSql
    Else
        coConex.Ejecutar lsSql
    End If
    coConex.CierraConexion
    Exit Sub

ErrordUpdateColocacCred:
    Err.Raise Err.Number, "Error En Proceso dUpdateColocacCred", Err.Description

End Sub
Public Sub dInsertMovOffHost(ByVal psMovNro As String, ByVal psOpeCod As String, _
    ByVal psMovDesc As String, ByVal pnMovEstado As Integer, _
    Optional pnMovFlag As Integer = 0, _
    Optional pbEjecBatch As Boolean = False)
    
Dim coConex As New DConecta
Dim lsSql As String
On Error GoTo InsertMovErr
    coConex.AbreConexion
    
    lsSql = "INSERT MovOffHost (cMovNro,cOpeCod,cMovDesc,nMovEstado, nMovFlag) " _
          & "VALUES ('" & psMovNro & "','" & psOpeCod & "','" & Replace(psMovDesc, "'", "''") & "'," _
          & pnMovEstado & "," & pnMovFlag & ")"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSql
    Else
        coConex.Ejecutar lsSql
    End If
    coConex.CierraConexion
    
    Exit Sub

InsertMovErr:
    
End Sub
Public Sub dInsertMov(ByVal psMovNro As String, ByVal psOpeCod As String, _
    ByVal psMovDesc As String, ByVal pnMovEstado As Integer, _
    Optional pnMovFlag As Integer = 0, _
    Optional pbEjecBatch As Boolean = False)
    
Dim coConex As New DConecta
Dim lsSql As String
On Error GoTo InsertMovErr
    coConex.AbreConexion
    
    lsSql = "INSERT DBCMAC..Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado, nMovFlag) " _
          & "VALUES ('" & psMovNro & "','" & psOpeCod & "','" & Replace(psMovDesc, "'", "''") & "'," _
          & pnMovEstado & "," & pnMovFlag & ")"
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSql
    Else
        coConex.Ejecutar lsSql
    End If
    coConex.CierraConexion
    
    Exit Sub

InsertMovErr:
    
End Sub
Private Function DevuelveMatrizGastosCredito(ByRef pnNumreg As Integer, ByVal psCtaCod As String, ByVal nNroCalen As Integer) As Variant
Dim R As ADODB.Recordset
'Dim oCalend As COMDCredito.DCOMCalendario
Dim MatGastosCred() As String
        'Set oCalend = New COMDCredito.DCOMCalendario
        Set R = RecuperaCalendarioGastos(psCtaCod, nNroCalen, 0, 1, True)
        'Set oCalend = Nothing
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
Public Function MatrizMontoPagado(ByVal MatCalendDistrib As Variant, _
    Optional ByVal pnCuota As Integer = -1) As Double
Dim i, J As Integer
    MatrizMontoPagado = 0

    If pnCuota <> -1 Then
        For i = 0 To UBound(MatCalendDistrib) - 1
            If CInt(MatCalendDistrib(i, 1)) = pnCuota Then
                For J = 3 To 9
                    MatrizMontoPagado = MatrizMontoPagado + MatCalendDistrib(i, J)
                Next J
                MatrizMontoPagado = MatrizMontoPagado + MatCalendDistrib(i, 11)
                Exit For
            End If
        Next i
    Else
        For i = 0 To UBound(MatCalendDistrib) - 1
            For J = 3 To 9
                MatrizMontoPagado = MatrizMontoPagado + MatCalendDistrib(i, J)
            Next J
            MatrizMontoPagado = MatrizMontoPagado + MatCalendDistrib(i, 11)
        Next i
    End If
End Function
Public Sub dUpdateColocCalendario(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
        ByVal pnCuota As Integer, ByVal pnColocCalendApl As Integer, Optional ByVal pdVenc As Date = CDate("01/01/1950"), _
        Optional ByVal pnColocCalendEstado As Integer = -1, Optional ByVal psDescripcion As String = "", Optional ByVal pnCalendProc As Integer = -1, _
        Optional ByVal pbEjecBatch As Boolean = False, Optional ByVal pcColocMiVivEval As String = "N", _
        Optional ByVal pdFecPago As Date = CDate("01/01/1900"))
        
Dim lsSql As String
Dim coConex As New DConecta

    On Error GoTo ErrordInsertColocCalendario
    'cCtaCod            nNroCalen   cColocCalendApl nCuota      dVenc                       nColocCalendEstado cDescripcion                                                                                                                                                                                                                                                    cColocCalenFlag nCalendProc
    coConex.AbreConexion
    
    lsSql = "UPDATE DBCMAC..ColocCalendario SET "
    If pdVenc <> CDate("01/01/1950") Then
        lsSql = lsSql & " dVenc = '" & Format(pdVenc, "mm/dd/yyyy") & "',"
    End If
    If pnColocCalendEstado <> -1 Then
        lsSql = lsSql & " nColocCalendEstado = " & pnColocCalendEstado & ","
    End If
    If psDescripcion <> "" Then
        lsSql = lsSql & " cDescripcion = '" & psDescripcion & "',"
    End If
    If pnCalendProc <> -1 Then
        lsSql = lsSql & " nCalendProc = " & pnCalendProc & ","
    End If
    If pcColocMiVivEval <> "N" Then
        lsSql = lsSql & " cColocMiVivEval = '" & pcColocMiVivEval & "',"
    End If
    If pdFecPago <> CDate("01/01/1900") Then
        lsSql = lsSql & " dPago = '" & Format(pdFecPago, "mm/dd/yyyy hh:mm:ss") & "',"
    End If
    
    lsSql = Mid(lsSql, 1, Len(lsSql) - 1)
    lsSql = lsSql & " WHERE cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & pnNroCalen & " AND nColocCalendApl = " & pnColocCalendApl & " AND nCuota = " & pnCuota
    
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSql
    Else
        coConex.Ejecutar lsSql
    End If
    coConex.CierraConexion
    Exit Sub

ErrordInsertColocCalendario:
    Err.Raise Err.Number, "Error En Proceso dInsertColocCalendario", Err.Description

End Sub

Private Function DevuelveMatrizGastosCreditoCuota(pnNumreg As Integer, ByVal pnCuota As Integer, MatGastosCred As Variant, ByVal pnNumRegGasCred As Integer) As Variant
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
           MatGastosCredCuota(nCont - 1, 0) = MatGastosCred(i, 0)
           MatGastosCredCuota(nCont - 1, 1) = MatGastosCred(i, 1)
           MatGastosCredCuota(nCont - 1, 2) = MatGastosCred(i, 2)
        End If
    Next i
    pnNumreg = nCont
    DevuelveMatrizGastosCreditoCuota = MatGastosCredCuota
End Function

Public Sub dAnularColocGarantRec(ByVal nNroGarantRec As Long, ByVal nEstado As Integer, Optional pbEjecBatch As Boolean = False)
Dim lsSql As String
Dim coConex As DConecta

    lsSql = "UPDATE ColocGarantRec SET nEstado = " & nEstado
    lsSql = lsSql & " Where nNroGarantRec = " & nNroGarantRec

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSql
    Else
        coConex.Ejecutar lsSql
    End If
End Sub



Public Function MatrizCapitalPagado(ByVal MatCalendDistrib As Variant) As Double
Dim i As Integer
    MatrizCapitalPagado = 0
    For i = 0 To UBound(MatCalendDistrib) - 1
        MatrizCapitalPagado = MatrizCapitalPagado + CDbl(MatCalendDistrib(i, 3))
    Next i
End Function
Public Sub LiberaGarantiaPago(ByVal pnMovNro As Long, ByVal psCtaCod As String, ByVal psCodAge As String, ByVal psCodUser As String, _
    ByVal pdFecha As Date, ByVal pnPrestamo As Double, ByVal pnMontoCapital As Double, ByVal bCancelado As Boolean)

'Dim oCred As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim R2 As ADODB.Recordset
'Dim oBase As COMDCredito.DCOMCredActBD
Dim sSQL As String
Dim sMovNro As String
Dim nMovNro As Long
Dim sOpeCod As String
Dim nSumaTotalGarantia As Double
Dim nPago As Double
Dim nMontoLibera As Double
Dim nMontoLiberaTotal As Double ' Total a liberar
Dim nMonToGravament As Double
Dim coConex As New DConecta

    'Set oBase = poBase
    'Set oCred = New COMDCredito.DCOMCredito
    coConex.AbreConexion
    Set R = RecuperaColocGarantia(psCtaCod)
    nSumaTotalGarantia = 0
    Do While Not R.EOF
        nSumaTotalGarantia = nSumaTotalGarantia + R!nGravado
        R.MoveNext
    Loop
    nSumaTotalGarantia = Format(nSumaTotalGarantia, "#0.00")
    nPago = Format(pnMontoCapital * nSumaTotalGarantia / pnPrestamo, "#0.00")
    R.Close

    nMontoLiberaTotal = 0

    Set R = RecuperaColocGarantia(psCtaCod)
    Do While Not R.EOF
        sSQL = "Select cOpeCod from DBCMAC..GarantiaOpe where nTpoGarantia = " & R!nTpoGarantia & " and nMoneda = " & R!nMoneda & " AND bReversion = 1 "
        Set R2 = New ADODB.Recordset
        R2.Open sSQL, coConex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
        If R2.RecordCount > 0 Then
            sOpeCod = R2!COPECOD
        Else
            sOpeCod = ""
        End If
        R2.Close
        nMontoLibera = nPago * (R!nGravado / nSumaTotalGarantia)
        'calculamos el nuevo monto de gravament
        nMonToGravament = IIf(R!nGravament - nMontoLibera <= 0, 0, R!nGravament - nMontoLibera)
        sSQL = "UPDATE DBCMAC..GARANTIAS SET nGravament = " & Format(nMonToGravament, "#0.00")
        sSQL = sSQL & " WHERE cNumGarant = '" & R!cNumGarant & "'"
        '-----------------------------
        coConex.ConexionActiva.Execute sSQL
        If bCancelado Then
            sSQL = " UPDATE DBCMAC..ColocGarantia Set nEstado = 0 WHERE cNumGarant = '" & R!cNumGarant & "' "
            sSQL = sSQL & " AND  cCtaCod = '" & psCtaCod & "'"
            coConex.ConexionActiva.Execute sSQL
        End If
        nMovNro = pnMovNro

        'En el numero de Calendario en caso de las garantias se asigna el cnumgarant convertido a entero

        'Call dInsertMovCol(nMovNro, sOpeCod, psCtaCod, CLng(IIf(IsNull(R!cNumGarant), 0, R!cNumGarant)), nMontoLibera, 0, "", 0, 0, 0, False)
        'Call dInsertMovColDet(nMovNro, sOpeCod, psCtaCod, CLng(IIf(IsNull(R!cNumGarant), 0, R!cNumGarant)), 1000, 0, nMontoLibera, False)

        R.MoveNext
    Loop
    ' LAYG - 2004/09/22   luis
    'nMovNro = pnMovNro
    'Call oBase.dInsertMovCol(nMovNro, sOpeCod, psCtacod, 0, nMontoLiberaTotal, 0, "", 0, 0, 0, False)
    'Call oBase.dInsertMovColDet(nMovNro, sOpeCod, psCtacod, 0, 1000, 0, nMontoLiberaTotal, False)

    R.Close
    coConex.CierraConexion
    'Set oCred = Nothing

End Sub

Public Function RecuperaCalendarioGastos(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
                    ByVal pnNroCuota As Integer, ByVal pnAplicado As Integer, _
                    Optional ByVal pbTodos As Boolean = False) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As New DConecta

    On Error GoTo ErrorRecuperaCalendarioGastos
    sSQL = "Select * from DBCMAC..ColocCalendDet where cCtaCod = '" & psCtaCod & _
            "' AND nNroCalen = " & pnNroCalen
    If Not pbTodos Then
        sSQL = sSQL & " AND nCuota = " & pnNroCuota
    End If
    sSQL = sSQL & " AND nColocCalendApl = " & pnAplicado
    sSQL = sSQL & " AND nPrdConceptoCod like '12%' Order By nCuota "

    oConecta.AbreConexion
    Set RecuperaCalendarioGastos = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaCalendarioGastos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


'Obtiene el nMovNro  apartir del cMovNro
Public Function dGetnMovNroOffHost(ByVal psMovNro As String) As Long
Dim lsSql As String
Dim lRs As ADODB.Recordset
Set lRs = New ADODB.Recordset
Dim coConex As New DConecta

coConex.AbreConexion
lsSql = "Select nMovNro From MovOffHost where cMovNro ='" & psMovNro & "'"
Set lRs = coConex.CargaRecordSet(lsSql, adLockReadOnly)
'Set lrs.ActiveConnection = coConex.ConexionActiva
If Not lRs.EOF And Not lRs.BOF Then
    dGetnMovNroOffHost = lRs!nMovNro
End If
lRs.Close
coConex.CierraConexion
Set lRs = Nothing
End Function
'Obtiene el nMovNro  apartir del cMovNro
Public Function dGetnMovNro(ByVal psMovNro As String) As Long
Dim lsSql As String
Dim lRs As ADODB.Recordset
Set lRs = New ADODB.Recordset
Dim coConex As New DConecta

coConex.AbreConexion
lsSql = "Select nMovNro From DBCMAC..Mov where cMovNro ='" & psMovNro & "'"
Set lRs = coConex.CargaRecordSet(lsSql, adLockReadOnly)
'Set lrs.ActiveConnection = coConex.ConexionActiva
If Not lRs.EOF And Not lRs.BOF Then
    dGetnMovNro = lRs!nMovNro
End If
lRs.Close
coConex.CierraConexion
Set lRs = Nothing
End Function
Public Sub dInsertMovColOffHost(ByVal pnMovNro As Long, ByVal psOperacion As String, _
        ByVal psCuenta As String, ByVal pnNroCalend As Long, ByVal pnMonto As Currency, _
        ByVal pnDiasMora As Integer, ByVal psMetLiq As String, ByVal pnPlazo As Integer, ByVal pnSaldoCap As Double, ByVal pnEstado As Integer, _
        Optional pbEjecBatch As Boolean = False, Optional ByVal pnFlag As Long = -1, Optional ByVal pnPrepago As Integer = 0)

Dim lsSql As String
Dim coConex As New DConecta
    coConex.AbreConexion
    lsSql = "INSERT MovColOffHost (nMovNro,cOpeCod,cCtaCod,nNroCalen,nMonto,nDiasMora,cMetLiq,nPlazo, nSaldoCap, nCredEstado,nPrepago" & IIf(pnFlag > 0, ",nFlag)", ")") _
        & "VALUES (" & pnMovNro & ",'" & psOperacion & "','" & psCuenta & "'," & pnNroCalend & "," _
        & pnMonto & "," & pnDiasMora & ",'" & psMetLiq & "'," & pnPlazo & "," & Format(pnSaldoCap, "#0.00") & "," & pnEstado & "," & pnPrepago & IIf(pnFlag > 0, "," & pnFlag & ")", ")")

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSql
    Else
        coConex.Ejecutar lsSql
    End If
    coConex.CierraConexion
End Sub
Public Sub dInsertMovCol(ByVal pnMovNro As Long, ByVal psOperacion As String, _
        ByVal psCuenta As String, ByVal pnNroCalend As Long, ByVal pnMonto As Currency, _
        ByVal pnDiasMora As Integer, ByVal psMetLiq As String, ByVal pnPlazo As Integer, ByVal pnSaldoCap As Double, ByVal pnEstado As Integer, _
        Optional pbEjecBatch As Boolean = False, Optional ByVal pnFlag As Long = -1, Optional ByVal pnPrepago As Integer = 0)

Dim lsSql As String
Dim coConex As New DConecta
    coConex.AbreConexion
    lsSql = "INSERT DBCMAC..MovCol (nMovNro,cOpeCod,cCtaCod,nNroCalen,nMonto,nDiasMora,cMetLiq,nPlazo, nSaldoCap, nCredEstado,nPrepago" & IIf(pnFlag > 0, ",nFlag)", ")") _
        & "VALUES (" & pnMovNro & ",'" & psOperacion & "','" & psCuenta & "'," & pnNroCalend & "," _
        & pnMonto & "," & pnDiasMora & ",'" & psMetLiq & "'," & pnPlazo & "," & Format(pnSaldoCap, "#0.00") & "," & pnEstado & "," & pnPrepago & IIf(pnFlag > 0, "," & pnFlag & ")", ")")

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSql
    Else
        coConex.Ejecutar lsSql
    End If
    coConex.CierraConexion
End Sub
Public Sub dInsertMovColDetOffHost(ByVal pnMovNro As Long, ByVal psOperacion As String, _
        ByVal psCtaCod As String, pnNroCalend As Long, ByVal pnConcepto As Long, _
        ByVal pnNroCuota As Integer, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSql As String
Dim coConex As New DConecta
    coConex.AbreConexion
    lsSql = "INSERT MovColDetOffHost (nMovNro,cOpeCod,cCtaCod,nNroCalen,nPrdConceptoCod,nNroCuota,nMonto) " _
        & "VALUES (" & pnMovNro & ",'" & psOperacion & "','" & psCtaCod & "'," & pnNroCalend & "," _
        & pnConcepto & "," & pnNroCuota & "," & pnMonto & " ) "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSql
    Else
        coConex.Ejecutar lsSql
    End If
    coConex.CierraConexion
End Sub

Public Sub dInsertMovColDet(ByVal pnMovNro As Long, ByVal psOperacion As String, _
        ByVal psCtaCod As String, pnNroCalend As Long, ByVal pnConcepto As Long, _
        ByVal pnNroCuota As Integer, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSql As String
Dim coConex As New DConecta
    coConex.AbreConexion
    lsSql = "INSERT DBCMAC..MovColDet(nMovNro,cOpeCod,cCtaCod,nNroCalen,nPrdConceptoCod,nNroCuota,nMonto) " _
        & "VALUES (" & pnMovNro & ",'" & psOperacion & "','" & psCtaCod & "'," & pnNroCalend & "," _
        & pnConcepto & "," & pnNroCuota & "," & pnMonto & " ) "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSql
    Else
        coConex.Ejecutar lsSql
    End If
    coConex.CierraConexion
End Sub

Public Sub dUpdateColocCalendDet(ByVal psCtaCod As String, ByVal pnNroCalen As Integer, _
        ByVal pnColocCalendApl As String, ByVal pnCuota As Integer, ByVal pnColocConceptoCod As Long, _
        Optional ByVal pnMonto As Double = -1, Optional ByVal pnMontoPagado As Double = -1, _
        Optional ByVal psFlag As String = "", Optional pbEjecBatch As Boolean = False, Optional pbAmortizar As Boolean = False _
        , Optional pbMontoIncrem As Boolean = False)

Dim lsSql As String
Dim coConex As New DConecta

    On Error GoTo ErrordUpdateColocCalendDet
    coConex.AbreConexion
    lsSql = "UPDATE DBCMAC..ColocCalendDet SET "
    If Not pbMontoIncrem Then
        If pnMonto <> -1 Then
            lsSql = lsSql & " nMonto = " & pnMonto & ","
        End If
    Else
        If pnMonto <> -1 Then
            lsSql = lsSql & " nMonto = nMonto + " & pnMonto & ","
        End If
    End If
    If pnMontoPagado <> -1 Then
        If pbAmortizar Then
            lsSql = lsSql & " nMontoPagado = nMontoPagado + " & pnMontoPagado & ","
        Else
            lsSql = lsSql & " nMontoPagado = " & pnMontoPagado & ","
        End If
    End If
    If psFlag <> "" Then
        lsSql = lsSql & " cFlag = '" & psFlag & "',"
    End If
    lsSql = Mid(lsSql, 1, Len(lsSql) - 1)
    lsSql = lsSql & " WHERE cCtaCod = '" & psCtaCod & "' AND nNroCalen = " & pnNroCalen & " AND nColocCalendApl = " & pnColocCalendApl & " AND nCuota = " & pnCuota & " AND nPrdConceptoCod = " & pnColocConceptoCod
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSql
    Else
        coConex.Ejecutar lsSql
    End If
    
    coConex.CierraConexion
    Exit Sub

ErrordUpdateColocCalendDet:
    Err.Raise Err.Number, "Error En Proceso dUpdateColocCalendDet", Err.Description

End Sub

Public Function RecuperaColocGarantia(ByVal psCtaCod As String, Optional poConex As ADODB.Connection) As ADODB.Recordset
Dim oConecta As New DConecta
Dim sSQL As String

    On Error GoTo ErrorRecuperaColocGarantia
    sSQL = "Select G.cTpoDoc, G.cNroDoc, G.nTpoGarantia, G.IdSupGarant, CG.cNumGarant, CG.cCtaCod, CG.nMoneda, CG.nGravado, ISNULL(G.nGravament,0) AS nGravament "
    sSQL = sSQL & " from DBCMAC..ColocGarantia CG Inner Join DBCMAC..Garantias G ON CG.cNumGarant = G.cNumGarant "
    sSQL = sSQL & " Where cCtaCod = '" & psCtaCod & "'"

   ' Set oConecta = New COMConecta.DCOMConecta
    'oConecta.AbreConexion
   ' If Not poConex Is Nothing Then
  '      oConecta.ConexionActiva = poConex
  '  End If
    If poConex Is Nothing Then
        Set oConecta = New DConecta
        oConecta.AbreConexion
        Set RecuperaColocGarantia = oConecta.CargaRecordSet(sSQL)
        oConecta.CierraConexion
        Set oConecta = Nothing
    Else
        Set RecuperaColocGarantia = poConex.Execute(sSQL)
    End If

    'oConecta.CierraConexion
    'Set oConecta = Nothing
    Exit Function

ErrorRecuperaColocGarantia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Sub RecuperaDatosNegocio()
Dim cmdNeg As New Command
Dim prmNegFecha As New ADODB.Parameter

    Set cmdNeg = New ADODB.Command
    Set prmNegFecha = New ADODB.Parameter
    Set prmNegFecha = cmdNeg.CreateParameter("@dFecSis", adDBDate, adParamOutput)
    cmdNeg.Parameters.Append prmNegFecha
    
    Set prmNegFecha = New ADODB.Parameter
    Set prmNegFecha = cmdNeg.CreateParameter("@nTipoCambioVenta", adDouble, adParamOutput)
    cmdNeg.Parameters.Append prmNegFecha
    
    Set prmNegFecha = New ADODB.Parameter
    Set prmNegFecha = cmdNeg.CreateParameter("@nTipoCambioCompra", adDouble, adParamOutput)
    cmdNeg.Parameters.Append prmNegFecha
    
    Set prmNegFecha = New ADODB.Parameter
    Set prmNegFecha = cmdNeg.CreateParameter("@nOFFHost", adInteger, adParamOutput)
    cmdNeg.Parameters.Append prmNegFecha
    
    loConec.AbreConexion
    cmdNeg.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    cmdNeg.CommandType = adCmdStoredProc
    cmdNeg.CommandText = "ATM_RecuperaDatosNegocio"
    
    cmdNeg.Execute
    
    
    dFecSis = cmdNeg.Parameters(0).Value
    
    'Coordinar operacion de madrugada y fecha de sistema aun sigue con dia de ayer por falgta de cierre de dia
    dFecSis = CDate(Format(dFecSis, "dd/MM/yyyy") & " " & Mid(Format(Now(), "dd/MM/yyyy hh:mm:ss.000"), 12, 8))
    nTipoCambioVenta = cmdNeg.Parameters(1).Value
    nTipoCambioCompra = cmdNeg.Parameters(2).Value
    nOFFHost = cmdNeg.Parameters(3).Value

    'las operaciones deben grabarse con la fecha y hora real**********
    If nOFFHost = 1 Then
        dFecSis = CDate(Format(Now(), "dd/MM/yyyy hh:mm:ss"))
    End If
    
    loConec.CierraConexion
    
    'Call RegistraSucesos(Now, "Recupero datos del Negocio : " & gTRACE, "")
     Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Recupero datos del Negocio : " & gTRACE, "")
     
    Set cmdNeg = Nothing
    Set prmNegFecha = Nothing

End Sub

Sub RecuperaDatosTarjeta()
Dim lRs As ADODB.Recordset
Dim lsSql As String
    
    lsSql = "exec PIT_RecuperaDatosTarjeta '" & gsPAN & "'"
      
    loConec.AbreConexion
    
    Set lRs = loConec.ConexionActiva.Execute(lsSql)
    
    If (Not lRs.EOF And Not lRs.BOF) Then
        nTarjCondicion = lRs.Fields("nCondicion")
        nRetenerTarjeta = lRs.Fields("nRetenerTarjeta")
        nNOOperMonExt = lRs.Fields("nNOOperMonExt")
        nSuspOper = lRs.Fields("nSuspOper")
        dFecVenc = lRs.Fields("dFecVenc")
    Else
        gsPAN = ""
    End If

    loConec.CierraConexion

    Set lRs = Nothing
    
    If Mid(gDATE_EXP, 2, 1) <> "." Then
        dFecVenc = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Mid(gDATE_EXP, 3, 2) & "/20" & Mid(gDATE_EXP, 1, 2))))
    Else
        'dFecSis = dFecVenc
    End If

    Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Recupero Datos Tarjeta ", "", gnTramaId)
    
End Sub

Sub RecuperaDatosCuenta()
Dim cmdNeg As New Command
Dim prmNegFecha As New ADODB.Parameter
    
    If sCtaCod <> "" And Mid(gPRCODE, 1, 2) <> "50" Then
        Set cmdNeg = New ADODB.Command
        Set prmNegFecha = New ADODB.Parameter
        Set prmNegFecha = cmdNeg.CreateParameter("@psCtaCod", adVarChar, adParamInput, 20, sCtaCod)
        cmdNeg.Parameters.Append prmNegFecha
        
        Set prmNegFecha = New ADODB.Parameter
        Set prmNegFecha = cmdNeg.CreateParameter("@nSaldo", adDouble, adParamOutput)
        cmdNeg.Parameters.Append prmNegFecha
        
        loConec.AbreConexion
        cmdNeg.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
        cmdNeg.CommandType = adCmdStoredProc
        cmdNeg.CommandText = "ATM_RecuperaDatosCuenta"

        cmdNeg.Execute
    
        nCtaSaldo = IIf(IsNull(cmdNeg.Parameters(1).Value), 0, cmdNeg.Parameters(1).Value)
    
        loConec.CierraConexion

        Set cmdNeg = Nothing
        Set prmNegFecha = Nothing
        
        Call RegistraBitacora(gsPAN, gsCanal, Now, sIDTrama, "Recupero Datos Cuenta ", "", gnTramaId)

    End If
End Sub

Public Function GeneraTramaEnXML() As String
Dim sXMLTrama As String
    sXMLTrama = "<MESSAGE_TYPE = " & gMESSAGE_TYPE & " />"
    sXMLTrama = sXMLTrama & " <TRACE = " & gTRACE & " />"
    sXMLTrama = sXMLTrama & " <PRCODE = " & gPRCODE & " />"
    sXMLTrama = sXMLTrama & " <PAN = " & gsPAN & " />"
    sXMLTrama = sXMLTrama & " <TIME_LOCAL = " & gTIME_LOCAL & " />"
    sXMLTrama = sXMLTrama & " <DATE_LOCAL = " & gDATE_LOCAL & " />"
    sXMLTrama = sXMLTrama & " <TERMINAL_ID = " & gTERMINAL_ID & " />"
    sXMLTrama = sXMLTrama & " <CARD_ACCEPTOR = " & gCARD_ACCEPTOR & " />"
    sXMLTrama = sXMLTrama & " <ACQ_INST = " & gACQ_INST & " />"
    sXMLTrama = sXMLTrama & " <POS_COND_CODE = " & gPOS_COND_CODE & " />"
    sXMLTrama = sXMLTrama & " <TXN_AMOUNT = " & gTXN_AMOUNT & " />"
    sXMLTrama = sXMLTrama & " <CUR_CODE = " & gCUR_CODE & " />"
    sXMLTrama = sXMLTrama & " <DATE_EXP = " & gDATE_EXP & " />"
    sXMLTrama = sXMLTrama & " <CARD_LOCATION = " & gCARD_LOCATION & " />"
    sXMLTrama = sXMLTrama & " <CUENTA = " & gsCtaCod & " />"
    
    GeneraTramaEnXML = sXMLTrama
End Function


Public Function RecuperaMovimientoInterCajaParaExtorno(pnTramaId As Long, psCodCMACOrigen As String) As ADODB.Recordset
Dim lCn As DConecta
Dim lsSql As String
    
    lsSql = " exec PIT_stp_sel_MovimientoAutInterCajaParaExtorno " & pnTramaId & ",'" & psCodCMACOrigen & "'"
    
    Set lCn = New DConecta
    lCn.AbreConexion
    Set RecuperaMovimientoInterCajaParaExtorno = lCn.CargaRecordSet(lsSql)
    lCn.CierraConexion
    Set lCn = Nothing
End Function

Public Sub RegistraPITMov(pnMovNro As Long, pnMovNroOffHost As Long, pnSecuencia As Integer, psPAN As String, _
            psDNI As String, pnMonto As Double, psHora As String, psMesDia As String, _
            pnMoneda As Integer, pnTramaId As Long, pnEstado As Integer)
Dim lCn As New DConecta
Dim lsSql As String

    lsSql = " exec PIT_stp_ins_PITMov " & pnMovNro & "," & pnMovNroOffHost & "," & pnSecuencia & ",'" & psPAN & "','" & _
                        psDNI & "'," & pnMonto & ",'" & psHora & "','" & psMesDia & "'," & pnMoneda & "," & pnTramaId & "," & pnEstado
    
    lCn.AbreConexion
    lCn.ConexionActiva.Execute (lsSql)
    lCn.CierraConexion

    Set lCn = Nothing
End Sub
Public Function getOpeCod(psPRCODE As String, Optional psCodTxFinanc As String = "0200") As String
    If psCodTxFinanc = "0200" Then
        Select Case psPRCODE
            Case "500035"
                getOpeCod = "105001"
            Case "353500"
                getOpeCod = "105002"
            Case "943500"
                getOpeCod = "104003"
            Case "011100"
                getOpeCod = "261501"
            Case "200011"
                getOpeCod = "261502"
            Case "351100"
                getOpeCod = "261503"
            Case "941100"
                getOpeCod = "261504"
            Case "311100"
                getOpeCod = "261505"
         End Select
     Else
         Select Case psPRCODE
            Case "500035"
                getOpeCod = "159201"
            Case "011100"
                getOpeCod = "279201"
            Case "200011"
                getOpeCod = "279202"
        End Select
    End If
End Function

Public Function ExtornarPagoCredito(ByVal pnMovNroExt As Long, _
                                    ByVal psCtaCod As String, _
                                    ByVal pnMonto As Double, _
                                    ByVal pdFecSis As Date, _
                                    ByVal psCodUser As String, _
                                    ByVal psCodAge As String, _
                                    ByRef pnResultado As Long)
Dim R As New ADODB.Recordset
Dim sSQL As String
Dim sMovNro As String
Dim oBase As New DConecta

Set oBase = New DConecta

If psCtaCod = "" Then
    psCtaCod = RecuperaCtaExtorno(pnMovNroExt)
End If

Set R = RecuperaDatosExtorno(pnMovNroExt, psCtaCod)
If R.EOF Then
    'psMensaje = "No se puede obtener los Datos del Extorno"
    'Set oDCred = Nothing
    Exit Function
End If

'sMovNro = GeneraMovNro(pdFecSis, psCodAge, psCodUser)
sMovNro = gsMovNro

Call dUpdateMov(pnMovNroExt, , , , 2, False)
Call dInsertMov(sMovNro, gPITColocExtPagoCredito, "Extorno de Pago", 13, 3, False)
lnMovNro = dGetnMovNro(sMovNro)

Call dInsertMovCol(lnMovNro, gPITColocExtPagoCredito, psCtaCod, 0, pnMonto, 0, "", 0, 0, 0, False)

Call dExtornaSaldosCalendario(pnMovNroExt, 1, psCtaCod, False)

Call dUpdateColocacCred(psCtaCod, R!nDiasMora, , , , , , , R!nMinCuota, , , , , , , R!nMinCuota, , , False)

Call dUpdateProducto(psCtaCod, , R!nCapital, R!nCredEstado, , , False, 1, True)

oBase.AbreConexion

sSQL = " Update PITMov Set nEstado = 2 Where nMovNro = " & pnMovNroExt 'Extornado por autorizador

oBase.Ejecutar (sSQL)
            
If Mid(psCtaCod, 6, 3) <> 423 Then

    sSQL = "DELETE From DBCMAC..ColocCalendario where cctaCod = '" & psCtaCod & "' and nnrocalen > ( "
    sSQL = sSQL & " Select min(nNrocalen) From DBCMAC..MovCol where cctaCod = '" & psCtaCod & "' and nMovNro = " & pnMovNroExt & " and nNroCalen <> 0)"
    oBase.Ejecutar (sSQL)

    sSQL = "DELETE From DBCMAC..ColocCalendDet where cctaCod = '" & psCtaCod & "' and nnrocalen > ( "
    sSQL = sSQL & " Select min(nNrocalen) From DBCMAC..MovCol where cctaCod = '" & psCtaCod & "' and nMovNro = " & pnMovNroExt & " and nNroCalen <> 0)"
    oBase.Ejecutar (sSQL)

    sSQL = " Update DBCMAC..ColocacCred SET nNroCalen = (Select min(nNrocalen) From DBCMAC..MovCol where cctaCod = '" & psCtaCod & "' and nMovNro = " & pnMovNroExt & " and nNroCalen <> 0), "
    sSQL = sSQL & " nNroCalPar = (Select min(nNrocalen) From DBCMAC..MovCol where cctaCod = '" & psCtaCod & "' and nMovNro = " & pnMovNroExt & " and nNroCalen <> 0 ) + 1 "
    sSQL = sSQL & "  From DBCMAC..ColocacCred "
    sSQL = sSQL & "  where cctaCod = '" & psCtaCod & "'"
    oBase.Ejecutar (sSQL)

End If

pnResultado = lnMovNro

R.Close
Set R = Nothing
oBase.CierraConexion
    
End Function

Public Function RecuperaDatosExtorno(ByVal pnMovNro As Long, ByVal psCtaCod As String) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As New DConecta

Set oConecta = New DConecta

    
    sSQL = "Select M.nCredEstado, SUM(MD.nMonto) as nCapital," _
         & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag, " _
         & " nMovNroRef = IsNull(MR.nMovNro, 0)" _
         & " From DBCMAC..MovCol M " _
         & "    Inner Join DBCMAC..MovColDet MD ON M.nMovNro = MD.nMovNro " _
         & "        AND M.cOpeCod = MD.cOpeCod AND MD.nPrdConceptoCod = 1000" _
         & "    Left Join DBCMAC..MovRef MR ON M.nMovNro = MR.nMovNroRef" _
         & " Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psCtaCod & "'" _
         & "     AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12'" _
         & "     AND M.cCtaCod = '" & psCtaCod & "' And M.nCredEstado <> 0 " _
         & "     and MD.COPECOD = '105001' and M.cOpeCod not Like '107[123456789]%'" _
         & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag, MR.nMovNro"

    oConecta.AbreConexion
    Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSQL)


    If RecuperaDatosExtorno.RecordCount = 0 Then
        sSQL = "Select M.nCredEstado, 0 as nCapital, "
        sSQL = sSQL & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
        sSQL = sSQL & " ,nMovNroRef=ISNULL(MR.nMovNro,0)" 'ARCV 06-09
        sSQL = sSQL & " From DBCMAC..MovCol M Inner Join DBCMAC..MovColDet MD ON M.nMovNro = MD.nMovNro "
        sSQL = sSQL & " AND M.cOpeCod = MD.cOpeCod " 'ARCV 28-05-07
        sSQL = sSQL & " AND MD.nPrdConceptoCod = 1100 "
        sSQL = sSQL & "     Left Join DBCMAC..MovRef MR ON M.nMovNro = MR.nMovNroRef " 'ARCV 06-09
        sSQL = sSQL & " Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psCtaCod & "' AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12' AND M.cCtaCod = '" & psCtaCod & "' and M.cOpeCod not Like '107[123456789]%'"
        sSQL = sSQL & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
        sSQL = sSQL & " ,MR.nMovNro" 'ARCV 06-09
        Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSQL)
    End If

    If RecuperaDatosExtorno.RecordCount = 0 Then
        sSQL = "Select M.nCredEstado, 0 as nCapital, "
        sSQL = sSQL & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
        sSQL = sSQL & " ,nMovNroRef=ISNULL(MR.nMovNro,0)" 'ARCV 06-09
        sSQL = sSQL & " From DBCMAC..MovCol M Inner Join DBCMAC..MovColDet MD ON M.nMovNro = MD.nMovNro "
        sSQL = sSQL & " AND M.cOpeCod = MD.cOpeCod " 'ARCV 28-05-07
        sSQL = sSQL & " AND MD.nPrdConceptoCod = 1101 "
        sSQL = sSQL & "     Left Join DBCMAC..MovRef MR ON M.nMovNro = MR.nMovNroRef " 'ARCV 06-09
        sSQL = sSQL & " Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psCtaCod & "' AND SUBSTRING(CONVERT(varchar(4),MD.nPrdConceptoCod),1,2)<>'12' AND M.cCtaCod = '" & psCtaCod & "' and M.cOpeCod not Like '107[123456789]%'"
        sSQL = sSQL & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
        sSQL = sSQL & " ,MR.nMovNro" 'ARCV 06-09
        Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSQL)
    End If

   ' Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSQL)
    If RecuperaDatosExtorno.RecordCount = 0 Then
        sSQL = "Select M.nCredEstado, 0 as nCapital, "
        sSQL = sSQL & " MIN(MD.nNroCuota) nMinCuota, M.nDiasMora, M.nFlag "
        sSQL = sSQL & " ,nMovNroRef=ISNULL(MR.nMovNro,0)" 'ARCV 06-09
        sSQL = sSQL & " From DBCMAC..MovCol M Inner Join DBCMAC..MovColDet MD ON M.nMovNro = MD.nMovNro "
        sSQL = sSQL & " AND M.cOpeCod = MD.cOpeCod " 'ARCV 28-05-07
        sSQL = sSQL & " AND Left(Cast(nPrdConceptoCod as Varchar(10)),2)='12' "
        sSQL = sSQL & "     Left Join DBCMAC..MovRef MR ON M.nMovNro = MR.nMovNroRef " 'ARCV 06-09
        sSQL = sSQL & " Where M.nMovNro = " & pnMovNro & " And MD.cCtaCod = '" & psCtaCod & "' AND Left(Cast(nPrdConceptoCod as Varchar(10)),2)='12' AND M.cCtaCod = '" & psCtaCod & "' and M.cOpeCod not Like '107[123456789]%'"
        sSQL = sSQL & " Group By  M.nCredEstado,M.nDiasMora, M.nFlag "
        sSQL = sSQL & " ,MR.nMovNro" 'ARCV 06-09
        Set RecuperaDatosExtorno = oConecta.CargaRecordSet(sSQL)
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function
Public Sub dUpdateMov(ByVal pnMovNro As Long, Optional ByVal psOpeCod As String = "@", _
    Optional ByVal psMovDesc As String = "@", Optional ByVal pnMovEstado As Integer = -1, _
    Optional ByVal pnMovFlag As Integer = -1, Optional pbEjecBatch As Boolean = False)
Dim coConex As New DConecta
Dim lsSql As String

Set coConex = New DConecta

coConex.AbreConexion
    lsSql = "UPDATE DBCMAC..Mov SET "

    If psOpeCod <> "@" Then
        lsSql = lsSql & " cOpeCod = '" & psOpeCod & "',"
    End If
    If psMovDesc <> "@" Then
        lsSql = lsSql & " cMovDesc = '" & psMovDesc & "',"
    End If
    If pnMovEstado <> -1 Then
        lsSql = lsSql & " nMovEstado = " & pnMovEstado & ","
    End If
    If pnMovFlag <> -1 Then
        lsSql = lsSql & " nMovFlag = " & pnMovFlag & ","
    End If
    
    lsSql = Left(lsSql, Len(lsSql) - 1)
    lsSql = lsSql & " WHERE nMovNro =" & pnMovNro & " "

    'If pbEjecBatch Then
    '    coConex.AdicionaCmdBatch lsSQL
    'Else
        coConex.Ejecutar lsSql
    'End If
    
    coConex.CierraConexion

End Sub

Public Sub dExtornaSaldosCalendario(ByVal pnMovNro As Long, _
    ByVal nAplicado As Integer, ByVal psCtaCod As String, Optional pbEjecBatch As Boolean = False)
Dim sSQL As String
Dim coConex As New DConecta

Set coConex = New DConecta

coConex.AbreConexion

    sSQL = "UPDATE DBCMAC..ColocCalendario SET nColocCalendEstado = 0, dPago = NULL "
    sSQL = sSQL & " From DBCMAC..ColocCalendario CC Inner Join DBCMAC..MovColDet MC ON CC.cCtaCod = MC.cCtacod "
    sSQL = sSQL & "     And CC.nNroCalen = MC.nNroCalen "
    sSQL = sSQL & "     And CC.nCuota = MC.nNroCuota "
    sSQL = sSQL & "     Where MC.nMovNro = " & pnMovNro & " And CC.nColocCalendApl = " & nAplicado
    sSQL = sSQL & "     AND CC.cCtaCod = '" & psCtaCod & "'"
    
    
        coConex.Ejecutar sSQL

    
    sSQL = "UPDATE DBCMAC..ColocCalendDet SET nMontoPagado = nMontoPagado - ABS(MC.nMonto) "
    sSQL = sSQL & " From DBCMAC..ColocCalendDet CC Inner Join DBCMAC..MovColDet MC ON CC.cCtaCod = MC.cCtacod "
    sSQL = sSQL & "     And CC.nNroCalen = MC.nNroCalen  And CC.nPrdConceptoCod = MC.nPrdConceptoCod"
    sSQL = sSQL & "     And CC.nCuota = MC.nNroCuota "
    sSQL = sSQL & "     Where MC.nMovNro = " & pnMovNro & " And CC.nColocCalendApl = " & nAplicado
    sSQL = sSQL & "     AND CC.cCtaCod = '" & psCtaCod & "'"
        
        coConex.Ejecutar sSQL

    
    sSQL = "Update DBCMAC..ColocCalendDet"
    sSQL = sSQL & " Set nMontoPagado = nMontoPagado - Abs(MD.nMonto)"
    sSQL = sSQL & " From DBCMAC..ColocCalendDet CD"
    sSQL = sSQL & " Inner Join DBCMAC..ColocacCred C on C.cCtaCod=CD.cCtaCod and C.nNroCalen=CD.nNroCalen"
    sSQL = sSQL & " Inner Join DBCMAC..MovColDet MD on MD.cCtaCod=CD.cCtaCod and MD.nNroCalen=CD.nNroCalen"
    sSQL = sSQL & " Where CD.nCuota=MD.nNroCuota and CD.nColocCalendApl=1 and MD.nNroCalen=CD.nNroCalen and"
    sSQL = sSQL & " MD.nMovNro=" & pnMovNro & "and MD.cCtaCod='" & psCtaCod & "' and MD.nPrdConceptoCod=1110 and"
    sSQL = sSQL & " CD.nPrdConceptoCod = 1010"

    

        coConex.Ejecutar sSQL


    sSQL = "Update DBCMAC..ColocCalendDet"
    sSQL = sSQL & " Set nMontoPagado = nMontoPagado - Abs(MD.nMonto)"
    sSQL = sSQL & " From DBCMAC..ColocCalendDet CD"
    sSQL = sSQL & " Inner Join DBCMAC..ColocacCred C on C.cCtaCod=CD.cCtaCod and C.nNroCalen=CD.nNroCalen"
    sSQL = sSQL & " Inner Join DBCMAC..MovColDet MD on MD.cCtaCod=CD.cCtaCod and MD.nNroCalen=CD.nNroCalen"
    sSQL = sSQL & " Where CD.nCuota=MD.nNroCuota and CD.nColocCalendApl=1 and MD.nNroCalen=CD.nNroCalen and "
    sSQL = sSQL & " MD.nMovNro=" & pnMovNro & "and MD.cCtaCod='" & psCtaCod & "' and MD.nPrdConceptoCod=1109 and"
    sSQL = sSQL & " CD.nPrdConceptoCod = 1000"

        coConex.Ejecutar sSQL
    
    'Extorno de Gastos que se cargaron al Momento de la Operacion
    sSQL = "UPDATE DBCMAC..ColocCalendDet SET nMonto = CC.nMonto - ABS(MC.nMonto) "
    sSQL = sSQL & " From DBCMAC..ColocCalendDet CC Inner Join DBCMAC..MovColDet MC ON CC.cCtaCod = MC.cCtacod "
    sSQL = sSQL & "     And CC.nNroCalen = MC.nNroCalen  And CC.nPrdConceptoCod = MC.nPrdConceptoCod"
    sSQL = sSQL & "     And CC.nCuota = MC.nNroCuota "
    sSQL = sSQL & "     Where MC.nMovNro = " & pnMovNro & " And CC.nColocCalendApl = " & nAplicado
    sSQL = sSQL & "     AND CC.cCtaCod = '" & psCtaCod & "' AND "
    sSQL = sSQL & " MC.nPrdConceptoCod in ( Select nPrdConceptoCod from DBCMAC..ProductoConcepto where nPrdConceptoCod like '12%' "
    sSQL = sSQL & " AND cAplicaProceso in ('CA','PA','PP')) "

        coConex.Ejecutar sSQL

    
    If nAplicado = 0 Then
        sSQL = "Delete DBCMAC..ColocCargoAutoma Where cCtaCod = '" & psCtaCod & "'"
        coConex.Ejecutar sSQL
    End If
    coConex.CierraConexion
    
End Sub

Public Function ValidaDNI(ByVal psCtaCod As String, ByVal psDNI As String) As Integer
Dim lCn As DConecta
Dim lRs As ADODB.Recordset
Dim lsSql As String
    
    lsSql = " exec PIT_stp_sel_VerificaDNIyCuenta '" & psCtaCod & "','" & psDNI & "'"
    
    Set lCn = New DConecta
    lCn.AbreConexion
    Set lRs = lCn.CargaRecordSet(lsSql)
    
    If (Not lRs.EOF And Not lRs.BOF) Then
        ValidaDNI = 0 'DNI si pertenece a la cuenta
    Else
        ValidaDNI = 1 'DNI no pertenece a la cuenta
    End If

    
    lCn.CierraConexion
    
    Set lRs = Nothing
    Set lCn = Nothing
End Function

Public Function DevuelveMes(ByVal psMes As String) As String
Dim sMes As String
    Select Case psMes
        Case "01"
            sMes = "ENE"
        Case "02"
            sMes = "FEB"
        Case "03"
            sMes = "MAR"
        Case "04"
            sMes = "ABR"
        Case "05"
            sMes = "MAY"
        Case "06"
            sMes = "JUN"
        Case "07"
            sMes = "JUL"
        Case "08"
            sMes = "AGO"
        Case "09"
            sMes = "SET"
        Case "10"
            sMes = "OCT"
        Case "11"
            sMes = "NOV"
        Case "12"
            sMes = "DIC"
    End Select
    DevuelveMes = sMes
End Function
Public Function obtenerParametros() As ADODB.Recordset
Dim lCn As DConecta
Dim lsSql As String
    
    lsSql = " exec PIT_stp_sel_Parametros "
    
    Set lCn = New DConecta
    lCn.AbreConexion
    Set obtenerParametros = lCn.CargaRecordSet(lsSql)
    lCn.CierraConexion
    Set lCn = Nothing
End Function
Public Function ValidaExtorno(psCtaCod As String, pnMovNro As Long, pnTipProd As Integer) As Boolean
Dim lCn As DConecta
Dim rs As ADODB.Recordset
Dim lsSql As String
    
    lsSql = " exec PIT_stp_sel_ValidaExtorno '" & psCtaCod & "'," & pnMovNro & "," & pnTipProd
    
    Set lCn = New DConecta
    lCn.AbreConexion
    Set rs = lCn.CargaRecordSet(lsSql)
    
    If Not rs.EOF() And Not rs.BOF() Then
        ValidaExtorno = False
    Else
        ValidaExtorno = True
    End If
    
    lCn.CierraConexion
    Set lCn = Nothing
End Function

Public Function RecuperaCtaExtorno(ByVal pnMovNro As Long) As String
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConecta As New DConecta

    Set oConecta = New DConecta
  
    sSQL = "Select  B.cMovNro, A.cCtaCod " _
         & "From DBCMAC..MovColDet A" _
         & "   Inner Join DBCMAC..Mov B On A.nMovNro = B.nMovNro " _
         & "Where A.nMovNro = " & pnMovNro & " And B.nMovEstado = 10 And B.nMovFlag = 0 " _
         & "Group By A.cCtaCod, B.cMovNro"

    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    
    RecuperaCtaExtorno = R!cCtaCod

    
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

