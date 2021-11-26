VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAutorizador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private nTramaID As Long

Property Get TramaId() As String
    TramaId = nTramaID
End Property


Public Function EjecutorGlobalNet(ByVal MESSAGE_TYPE As String, ByVal TRACE As String, ByVal PRCODE As String, _
    ByVal PAN As String, ByVal TIME_LOCAL As String, ByVal DATE_LOCAL As String, ByVal TERMINAL_ID As String, _
    ByVal ACCT_1 As String, ByVal CARD_ACCEPTOR As String, ByVal ACQ_INST As String, ByVal POS_COND_CODE As String, _
    ByVal TXN_AMOUNT As String, ByVal CUR_CODE As String, ByVal ACCT_2 As String, ByVal DATE_EXP As String, ByVal CARD_LOCATION As String, _
    ByVal psMonCta As String, ByVal psMovNro As String, ByVal pnTranaID As Long, Optional ByVal psCtaCod As String, Optional ByVal psDNI As String) As String
    
    EjecutorGlobalNet = TransaccionGlobalNet(MESSAGE_TYPE, TRACE, PRCODE, PAN, TIME_LOCAL, DATE_LOCAL, TERMINAL_ID, _
        ACCT_1, CARD_ACCEPTOR, ACQ_INST, POS_COND_CODE, TXN_AMOUNT, CUR_CODE, ACCT_2, DATE_EXP, CARD_LOCATION, psMonCta, psMovNro, pnTranaID, psCtaCod, psDNI)
    
        
End Function
Public Sub RegistrarTrama_1(ByVal psIDTrama As String, ByVal psTramas As String, ByVal pnDenegada As Integer)
    Call RegistrarTrama(psIDTrama, psTramas, pnDenegada)
End Sub
  'NSSE 05/12/2008
Public Function RecuperaCondicionDeTarjeta_1(ByVal PAN As String) As Integer
    RecuperaCondicionDeTarjeta_1 = RecuperaCondicionDeTarjeta(PAN)
End Function

 Public Function RetenerTarjetaPorPosibleFraude_1(ByVal psNumTarj As String) As Integer
    RetenerTarjetaPorPosibleFraude_1 = RetenerTarjetaPorPosibleFraude(psNumTarj)
 
 End Function
 Public Sub RegistraOperacionLimitesCajeroPOS_1(ByVal pdFecha As Date, ByVal psCodTranCaj As String, _
        ByVal psNumTarj As String, ByVal pnMonto As Double)
        
        Call RegistraOperacionLimitesCajeroPOS(pdFecha, psCodTranCaj, psNumTarj, pnMonto, 1)
        
        End Sub
 Public Function RecuperaPVV_1(ByVal psPAN As String) As String
        RecuperaPVV_1 = RecuperaPVV(psPAN)
 End Function
  
Public Sub ActualizaPVV_1(ByVal pssPVV As String, ByVal psPAN As String)
   Call ActualizaPVV(pssPVV, psPAN)

End Sub
Public Function Ejecutor(ByVal psEntrada As String) As String
    Ejecutor = Transaccion(psEntrada)
End Function
Public Function Prueba(ByVal psPAN As String, ByVal pnMonto As String)
    Dim sCtaCod  As String
    Dim PAN As String
    Dim sHora As String
    Dim sMesDia As String
    Dim sUserATM As String
    
    Dim Cmd As New Command
    Dim cmdNeg As New Command
    Dim prmNegFecha As New ADODB.Parameter
    Dim Prm As New ADODB.Parameter
    
    Dim loConec As New DConecta
    Dim sCtaCascada As String
    Dim dFecSis As Date
    Dim nOFFHost As Integer
    Dim PRCODE As String
    Dim sMonCta As String
    Dim nResultado As Integer
    
    Dim nSaldoCtaDisp As Double
    Dim nSaldoCtaTot As Double
    Dim sCadRespSalCta As String
    
    sMonCta = ""
    PAN = Trim(psPAN)
    sHora = Time() 'Trim(TIME_LOCAL)
    sMesDia = CStr(Month(Date)) + CStr(Day(Date)) 'Trim(DATE_LOCAL)
    'sUserATM = Trim(RecuperaUserATM(Trim(TERMINAL_ID)))
    sUserATM = "AT00"
    'If Len(Trim(sUserATM)) = 0 Then
    '    sUserATM = "AT00"
    'End If
    
    '*****************************************
    'RECUPERA DATOS DEL NEGOCIO
    '*****************************************
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
    dFecSis = CDate(Format(dFecSis, "dd/MM/yyyy") & " " & Mid(Format(Now(), "dd/MM/yyyy hh:mm:ss"), 12, 8))
    'nTipoCambioVenta = cmdNeg.Parameters(1).Value
    'nTipoCambioCompra = cmdNeg.Parameters(2).Value
    nOFFHost = cmdNeg.Parameters(3).Value
    
    '**DAOR 20081218, las operaciones en ATM deben grabarse con la fecha y hora real**********
    If nOFFHost = 1 Then
        dFecSis = CDate(Format(Now(), "dd/MM/yyyy hh:mm:ss"))
    End If
    '*****************************************************************************************
    
    'Call CerrarConexion
    loConec.CierraConexion
    
    sCtaCod = RecuperaCtaDisponible(psPAN, IIf(sMonCta = "604", 1, 2), IIf(Mid(PRCODE, 3, 2) = "12", "234", "232"))
    
    Call RecuperaSaldoDeCuenta(sCtaCod, nSaldoCtaDisp, nSaldoCtaTot)
    
    sCadRespSalCta = "1001"
        If Mid(sCtaCod, 9, 1) = "1" Then
            sCadRespSalCta = sCadRespSalCta & "604"
        Else
            sCadRespSalCta = sCadRespSalCta & "840"
        End If
        sCadRespSalCta = sCadRespSalCta & "C"
        sCadRespSalCta = sCadRespSalCta & Right("000000000000" & Replace(Format(nSaldoCtaTot, "#0.00"), ".", ""), 12)
        
        'SALDO DISPONIBLE
        sCadRespSalCta = sCadRespSalCta & "1002"
        If Mid(sCtaCod, 9, 1) = "1" Then
            sCadRespSalCta = sCadRespSalCta & "604"
        Else
            sCadRespSalCta = sCadRespSalCta & "840"
        End If
        sCadRespSalCta = sCadRespSalCta & "C"
        sCadRespSalCta = sCadRespSalCta & Right("000000000000" & Replace(Format(nSaldoCtaDisp, "#0.00"), ".", ""), 12)
    
    nResultado = RetiroInterCMAC(dFecSis, sCtaCod, pnMonto, 0, 1, "", 0, 3.14, "", 0, PAN, sHora, sMesDia, 0)
    
    
End Function
Public Function PIT_GeneraMovNro(pdFecha As Date, psCodAge As String, psUser As String) As String
    PIT_GeneraMovNro = GeneraMovNro(pdFecha, psCodAge, psUser)
End Function
Public Function PIT_nRegistrarTramaRecepcion(psMovNro As String, psOpeCod As String, psTramaRecep As String, psCodTX As String, psCuentaTX As String, _
                                         psPANTX As String, psDNITX As String, psMonedaTX As String, pnMontoTX As Double, psCodInstTX As String) As Long
    PIT_nRegistrarTramaRecepcion = nRegistrarTramaRecepcion(psMovNro, psOpeCod, psTramaRecep, psCodTX, psCuentaTX, psPANTX, psDNITX, psMonedaTX, pnMontoTX, psCodInstTX)
    nTramaID = PIT_nRegistrarTramaRecepcion
End Function
Public Function PIT_RegistrarTramaEnvio(pnTramaId As Integer, psTramaEnvio As String, psCodRespTX As String, pnDenegada As Integer)
    Call RegistrarTramaEnvio(pnTramaId, psTramaEnvio, psCodRespTX, pnDenegada)
End Function

Public Sub PIT_RegistraBitacora(ByVal psNumTarjeta As String, ByVal psCanal As String, ByVal pdFecha As Date, ByVal psIDTrama As String, ByVal psProceso As String, ByVal psDescrip As String, Optional ByVal pnTramaId As Long = 0)
    Call RegistraBitacora(psNumTarjeta, psCanal, pdFecha, psIDTrama, psProceso, psDescrip, pnTramaId)
End Sub

'DAOR de prueba
Public Function PIT_ValidaExtorno(psCtaCod As String, pnMovNro As Long, pnTipProd As Integer) As Boolean
    PIT_ValidaExtorno = ValidaExtorno(psCtaCod, pnMovNro, pnTipProd)
End Function

