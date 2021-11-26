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

Public Function EjecutorGlobalNet(ByVal MESSAGE_TYPE As String, ByVal TRACE As String, ByVal PRCODE As String, _
    ByVal PAN As String, ByVal TIME_LOCAL As String, ByVal DATE_LOCAL As String, ByVal TERMINAL_ID As String, _
    ByVal ACCT_1 As String, ByVal CARD_ACCEPTOR As String, ByVal ACQ_INST As String, ByVal POS_COND_CODE As String, _
    ByVal TXN_AMOUNT As String, ByVal CUR_CODE As String, ByVal ACCT_2 As String, ByVal DATE_EXP As String, ByVal CARD_LOCATION As String, ByVal psMonCta As String, ByRef pnMovNro As Long) As String
    
    EjecutorGlobalNet = TransaccionGlobalNet(MESSAGE_TYPE, TRACE, PRCODE, PAN, TIME_LOCAL, DATE_LOCAL, TERMINAL_ID, _
        ACCT_1, CARD_ACCEPTOR, ACQ_INST, POS_COND_CODE, TXN_AMOUNT, CUR_CODE, ACCT_2, DATE_EXP, CARD_LOCATION, psMonCta, pnMovNro)
    
        
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
 Public Sub RegistraOperacionLimitesCajeroPOS_1(ByVal pdfecha As Date, ByVal psCodTranCaj As String, _
        ByVal psNumTarj As String, ByVal pnMonto As Double)
        
        Call RegistraOperacionLimitesCajeroPOS(pdfecha, psCodTranCaj, psNumTarj, pnMonto, 1)
        
        End Sub
 Public Function RecuperaPVV_1(ByVal psPan As String) As String
        RecuperaPVV_1 = RecuperaPVV(psPan)
 End Function
  
Public Sub ActualizaPVV_1(ByVal pssPVV As String, ByVal psPan As String)
   Call ActualizaPVV(pssPVV, psPan)

End Sub
Public Function Ejecutor(ByVal psEntrada As String) As String
    Ejecutor = Transaccion(psEntrada)
End Function


'Option Explicit
'Private C As ADODB.Connection
'
'Private Function AbrirConexion() As ADODB.Connection
'Dim sCadCadConex As String
'
'    'sCadCadConex = "Provider=SQLOLEDB.1;Password=DBNegocio;User ID=sa;Initial Catalog=DBTarjeta;Data Source=SVRCOM02"
'    sCadCadConex = "Provider=SQLOLEDB.1;Password=1234;User ID=sa;Initial Catalog=DBTarjeta;Data Source=10.0.0.251"
'    Set C = New ADODB.Connection
'    C.Open sCadCadConex
'
'    Set AbrirConexion = C
'
'End Function
'
'Private Sub CerrarConexion()
'    C.Close
'    Set C = Nothing
'
'End Sub
'
'Public Function RecuperaValorXML(ByVal pINXml As String, ByVal psEtiqueta As String)
'Dim i As Integer
'Dim sCadTempo As String
'Dim sCadValor As String
'Dim bIniCad As Boolean
'
'    i = 1
'    Do While i <= Len(pINXml)
'
'        Do While Mid(pINXml, i, 1) <> "<" And i <= Len(pINXml)
'            i = i + 1
'        Loop
'        i = i + 1
'        sCadTempo = ""
'        Do While Mid(pINXml, i, 1) <> " " And Mid(pINXml, i, 1) <> ">" And i <= Len(pINXml)
'            sCadTempo = sCadTempo & Mid(pINXml, i, 1)
'            i = i + 1
'        Loop
'
'        If UCase(Trim(sCadTempo)) = UCase(Trim(psEtiqueta)) Then
'                Do While Mid(pINXml, i, 1) <> "=" And Mid(pINXml, i, 1) <> ">" And i <= Len(pINXml)
'                    i = i + 1
'                Loop
'                i = i + 1
'                sCadValor = ""
'                Do While Mid(pINXml, i, 1) <> "/" And Mid(pINXml, i, 1) <> ">" And i <= Len(pINXml) And Mid(pINXml, i, 1) <> "<"
'                    sCadValor = sCadValor & Mid(pINXml, i, 1)
'                    i = i + 1
'                Loop
'                sCadValor = Trim(Replace(sCadValor, """", ""))
'                Exit Do
'        End If
'
'    Loop
'
'    RecuperaValorXML = sCadValor
'
'End Function
'
'Public Function GeneraXMLSalida(ByVal RESP_CODE As String, Optional ByVal CARDISS_AMOUNT As String, _
'    Optional ByVal AUTH_CODE As String, Optional ByVal ADD_RESP_DATA As String, Optional ByVal CUR_CODE_CARDISS As String, _
'    Optional ByVal AMOUNTS_ADD As String, Optional ByVal PRIV_USE As String) As String
'
'Dim sXML As String
'
'sXML = "<?xml version=""1.0""?>"
'sXML = sXML & "<Messages>"
'sXML = sXML & "<TXN_FIN_RES>"
'sXML = sXML & "<CARDISS_AMOUNT  value=""" & CARDISS_AMOUNT & """ />"
'sXML = sXML & "<AUTH_CODE      value=""" & AUTH_CODE & """ />"  '// Código de autorización (si la transacción es aprobada).
'sXML = sXML & "<RESP_CODE    value=""" & RESP_CODE & """ />" '// Código de Respuesta (Valores Válidos indicados en el Anexo)
'sXML = sXML & "<ADD_RESP_DATA      value=""" & ADD_RESP_DATA & """ />"
'sXML = sXML & "<CUR_CODE_CARDISS   value=""" & CUR_CODE_CARDISS & """ />"
'sXML = sXML & "<AMOUNTS_ADD   value=""" & AMOUNTS_ADD & """ />" '// Datos de la Consulta
'sXML = sXML & "<PRIV_USE   value=""" & PRIV_USE & """ />" '// Glosa a imprimir en el recibo de ATM (aplica en ciertas transacciones)
'sXML = sXML & "</TXN_FIN_RES>"
'sXML = sXML & "</Messages>"
'
'GeneraXMLSalida = sXML
'
'End Function
'
'Private Function DE_TRAMA_ConvierteAMontoReal(ByVal psMontoTxN As String) As Double
'
'    DE_TRAMA_ConvierteAMontoReal = CDbl(Mid(psMontoTxN, 1, Len(psMontoTxN) - 2) & "." & Right(psMontoTxN, 2))
'
'End Function
'
'Private Function ObtieneITF() As Double
'Dim Cmd As New Command
'Dim Prm As New ADODB.Parameter
'
'
'    Set Cmd = New ADODB.Command
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@nITF", adDouble, adParamOutput, , ObtieneITF)
'    Cmd.Parameters.Append Prm
'
'
'    Cmd.ActiveConnection = AbrirConexion
'    Cmd.CommandType = adCmdStoredProc
'    Cmd.CommandText = "ATM_RecuperaValorITF"
'    Cmd.Execute
'
'
'    ObtieneITF = Cmd.Parameters(0).Value
'
'    Set Cmd = Nothing
'    Set Prm = Nothing
'
'
'End Function
'
'Public Function RecuperaSaldoDisp(ByVal psCtaCod As String) As Double
'Dim Cmd As New Command
'Dim Prm As New ADODB.Parameter
'
'    Set Cmd = New ADODB.Command
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@psCodCta", adVarChar, adParamInput, 50, psCtaCod)
'    Cmd.Parameters.Append Prm
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@nSaldoDisp", adDouble, adParamOutput, , psCtaCod)
'    Cmd.Parameters.Append Prm
'
'    Cmd.ActiveConnection = AbrirConexion
'    Cmd.CommandType = adCmdStoredProc
'    Cmd.CommandText = "ATM_RecuperaSaldoCuenta"
'    Cmd.Execute
'
'    RecuperaSaldoDisp = Cmd.Parameters(1).Value
'
'    Set Cmd = Nothing
'    Set Prm = Nothing
'
'End Function
'
''****************************************
''RECUPERA COMISION DE TARIFARIO
''****************************************
'Public Function ObtieneComisionPorOperacion(ByVal pnTipoEquipo As Integer, ByVal pnTipoServ As Integer, _
'    ByVal pnProced As Integer, ByVal pnMoneda As Integer, ByVal pnTipoOper As Integer, _
'    ByVal pnMontoOper As Double) As Double
'
'
'Dim Cmd As New Command
'Dim Prm As New ADODB.Parameter
'
'
'    Set Cmd = New ADODB.Command
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@pnTipoEquipo", adInteger, adParamInput, , pnTipoEquipo)
'    Cmd.Parameters.Append Prm
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@pnTipoServ", adDouble, adParamInput, , pnTipoServ)
'    Cmd.Parameters.Append Prm
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@pnProced", adDouble, adParamInput, , pnProced)
'    Cmd.Parameters.Append Prm
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@pnMoneda", adDouble, adParamInput, , pnMoneda)
'    Cmd.Parameters.Append Prm
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@pnTipoOper", adDouble, adParamInput, , pnTipoOper)
'    Cmd.Parameters.Append Prm
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@pnMonto", adDouble, adParamInput, , pnMontoOper)
'    Cmd.Parameters.Append Prm
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@pnValor", adDouble, adParamOutput)
'    Cmd.Parameters.Append Prm
'
'
'    Cmd.ActiveConnection = AbrirConexion
'    Cmd.CommandType = adCmdStoredProc
'    Cmd.CommandText = "ATM_RecupComision"
'    Cmd.Execute
'
'    ObtieneComisionPorOperacion = Cmd.Parameters(6).Value
'
'    Set Cmd = Nothing
'    Set Prm = Nothing
'
'End Function
'
'Public Sub RegistraSucesos(ByVal pdFecha As Date, ByVal psProceso As String, ByVal psDescrip As String)
'    Dim Cmd As ADODB.Command
'    Dim Prm As ADODB.Parameter
'
'
'            Set Cmd = New ADODB.Command
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@dFecha", adDBDate, adParamInput, , pdFecha)
'            Cmd.Parameters.Append Prm
'
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@cProceso", adVarChar, adParamInput, 150, psProceso)
'            Cmd.Parameters.Append Prm
'
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@cDescripcion", adVarChar, adParamInput, 5000, psDescrip)
'            Cmd.Parameters.Append Prm
'
'            Cmd.ActiveConnection = AbrirConexion
'            Cmd.CommandType = adCmdStoredProc
'            Cmd.CommandText = "ATM_RegistraSucesos"
'
'            Cmd.Execute
'
'            Call CerrarConexion
'
'
'    Set Cmd = Nothing
'    Set Prm = Nothing
'
'End Sub
'
'Public Function ValidaCampos(ByVal psCtaCod As String, ByVal pnMontoTran As Double, _
'    ByVal pnMontoITF As Double, ByVal pnMontoComision As Double, ByVal pnMontoComisionITF As Double, _
'    ByVal pnMoneda As Integer, ByVal psOpeCod As String, ByVal psOpeCodComision As String, _
'    ByVal psOpeCodITF As String, ByVal psOpeCodComisionITF As String, ByVal psOpeExtorno As String, _
'    ByVal psOpeExtornoComision As String, ByVal psOpeCodTransferencia As String, ByVal psOpeCodExtornoTransfer As String, _
'    ByVal pnTipoCambioCompra As Double, ByVal pnTipoCambioVenta As Double, ByVal psIDTrama As String, ByVal psCtaDeposito As String, _
'    ByVal psProceso As String) As Boolean
'
'    Dim dfecha As Date
'    dfecha = Now()
'
'    ValidaCampos = True
'
'    If Len(psCtaCod) <> 18 Then
'        ValidaCampos = False
'        Call RegistraSucesos(dfecha, "Validacion de Campos - Trama :" & psIDTrama, "Cuenta con longitud menor a 15 digitos")
'    End If
'    If (pnMontoTran <= 0 Or pnMontoITF < 0 Or pnMontoComision < 0 Or pnMontoComisionITF < 0) And psProceso = "31" And psProceso = "39" And psProceso = "93" And psProceso = "91" And psProceso = "98" Then
'        ValidaCampos = False
'        Call RegistraSucesos(dfecha, "Validacion de Campos - Trama :" & psIDTrama, "alguno de los Montos son menores que cero")
'    End If
'    If pnMoneda <> 1 And pnMoneda <> 2 Then
'        ValidaCampos = False
'        Call RegistraSucesos(dfecha, "Validacion de Campos - Trama :" & psIDTrama, "Moneda es diferente de 1 o 2")
'    End If
'    'david se agregho And psProceso <> "31"
'    If (Len(psOpeCod) <> 6 Or Len(psOpeCodComision) <> 6 Or Len(psOpeCodITF) <> 6 _
'        Or Len(psOpeCodComisionITF) <> 6 Or Len(psOpeExtorno) <> 6 Or Len(psOpeExtornoComision) <> 6 _
'        Or Len(psOpeCodTransferencia) <> 6 Or Len(psOpeCodExtornoTransfer) <> 6) And psProceso <> "31" Then
'        ValidaCampos = False
'        Call RegistraSucesos(dfecha, "Validacion de Campos - Trama :" & psIDTrama, "Codigo de Operacion en diferente de 6 digitos")
'    End If
'
'    If pnTipoCambioCompra <= 0 Or pnTipoCambioVenta <= 0 Then
'        ValidaCampos = False
'        Call RegistraSucesos(dfecha, "Validacion de Campos - Trama :" & psIDTrama, "Monto de Compra o venta menor o igual a Cero")
'    End If
'
'    If Len(Trim(psIDTrama)) = 0 Then
'        ValidaCampos = False
'        Call RegistraSucesos(dfecha, "Validacion de Campos - Cuenta :" & psCtaCod, "ID de la Trama en blanco")
'    End If
'
'    If Len(psCtaDeposito) <> 18 And psProceso = "40" Then
'        ValidaCampos = False
'        Call RegistraSucesos(dfecha, "Validacion de Campos - Trama :" & psIDTrama, "Cuenta de Deposito Invalida")
'    End If
'
'    End Function
'
'Public Function VerificaEstadoCuenta(ByVal psCtaCod As String) As Integer
'    Dim Cmd As ADODB.Command
'    Dim Prm As ADODB.Parameter
'
'
'            Set Cmd = New ADODB.Command
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 18, psCtaCod)
'            Cmd.Parameters.Append Prm
'
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pnResultado", adInteger, adParamOutput)
'            Cmd.Parameters.Append Prm
'
'
'            Cmd.ActiveConnection = AbrirConexion
'            Cmd.CommandType = adCmdStoredProc
'            Cmd.CommandText = "ATM_EstadoCuenta"
'
'            Cmd.Execute
'            VerificaEstadoCuenta = Cmd.Parameters(1).Value
'
'            Call CerrarConexion
'
'
'    Set Cmd = Nothing
'    Set Prm = Nothing
'
'
'
'End Function
'
'Public Function ValidaOperacion(ByVal psCtaCod As String, ByVal pnMontoTran As Double, _
'    ByVal pnMontoITF As Double, ByVal pnMontoComision As Double, ByVal pnMontoComisionITF As Double, _
'    ByVal pnMoneda As Integer, ByVal psOpeCod As String, ByVal psOpeCodComision As String, _
'    ByVal psOpeCodITF As String, ByVal psOpeCodComisionITF As String, ByVal psOpeExtorno As String, _
'    ByVal psOpeExtornoComision As String, ByVal psOpeCodTransferencia As String, ByVal psOpeCodExtornoTransfer As String, _
'    ByVal pnTipoCambioCompra As Double, ByVal pnTipoCambioVenta As Double, ByVal psIDTrama As String, ByVal psCtaDeposito As String, _
'    ByVal pPAN As String, ByVal dfecha As Date, ByVal dFecVenc As Date, ByVal pnCondicion As Integer, _
'    ByVal pnRetenerTarjeta As Integer, ByVal pnCtaSaldo As Double, ByVal psProceso As String, _
'    ByVal pnNOOperMonExt As Integer, ByVal pnSuspOper As Integer) As String
'
'
'Dim sRespuetas As String
'
''00 Acepta la Transaccion
'    ValidaOperacion = "00"
'
''12 Si el Recibe un mensaje no soportado o encuentra un problema de formato en uno de los campos
''30 Cuando en el mensaje recibido se encuentra algun error de formato.
'If ValidaCampos(psCtaCod, pnMontoTran, pnMontoITF, pnMontoComision, pnMontoComisionITF, pnMoneda, psOpeCod, psOpeCodComision, _
'    psOpeCodITF, psOpeCodComisionITF, psOpeExtorno, psOpeExtornoComision, psOpeCodTransferencia, psOpeCodExtornoTransfer, _
'    pnTipoCambioCompra, pnTipoCambioVenta, psIDTrama, psCtaDeposito, Mid(psProceso, 1, 2)) = False Then
'
'    ValidaOperacion = "30"
'    Exit Function
'End If
'
''14 Si recibe un numero de tarjeta Invalido
''->Cusco Alcanzará la estructura de la tarjeta
'If Len(pPAN) = 0 Then
'    ValidaOperacion = "14"
'    Exit Function
'End If
'
'
''28 Si el archivo, el registro de la cuenta, o la tarjeta esta en Uso por otro aplicativo del
''    interno del banco
''-> Definir con CMAC Cusco
'
''56 Si la tarjeta no existe
'If pnCondicion = 0 Then
'        ValidaOperacion = "56"
'        Exit Function
'End If
'
''33 Si la tarjeta esta Vencida
'If dfecha > dFecVenc Then
'        ValidaOperacion = "33"
'        Exit Function
'End If
'
''41 Si la tarjeta esta en condicion de PERDIDA
'' Se usa la Tabla Tarjeta campo nCondicion
'If pnCondicion = 2 Then
'    ValidaOperacion = "41"
'    Exit Function
'End If
'
''43 Si la tarjeta esta en la condicion de robada o HOT(Fuerza al ATM a Capturar la Tarjeta)
'' Se usa la Tabla Tarjeta campo nCondicion=10 y nRetenerTarjeta=1
'If pnCondicion = 10 Or pnRetenerTarjeta = 1 Then
'    ValidaOperacion = "43"
'    Exit Function
'End If
'
''51 Si la cuenta no tiene el saldo suficiente para atender el requerimiento del Cliente
'
'Dim nMontoTotal As Double
'nMontoTotal = pnMontoTran + pnMontoITF + pnMontoComision + pnMontoComisionITF
'
'If pnMoneda <> CInt(Mid(psCtaCod, 9, 1)) Then
'    If pnMoneda = 1 Then 'Soles
'        nMontoTotal = nMontoTotal / pnTipoCambioCompra
'
'    Else  'Dolares
'        nMontoTotal = nMontoTotal * pnTipoCambioVenta
'    End If
'End If
'If pnCtaSaldo <= nMontoTotal Then
'        ValidaOperacion = "51"
'        Exit Function
'End If
'
''57 Si la transaccion no esta permitida para la tarjeta
'If Not (psProceso = "910000" Or (Mid(psProceso, 1, 2) = "01" And Mid(psProceso, 5, 2) = "00") _
'    Or (Mid(psProceso, 1, 2) = "00" And Mid(psProceso, 5, 2) = "00") _
'    Or (Mid(psProceso, 1, 2) = "97" And Mid(psProceso, 5, 2) = "00") _
'    Or (Mid(psProceso, 1, 2) = "31" And Mid(psProceso, 5, 2) = "00") _
'    Or (Mid(psProceso, 1, 2) = "39") Or (Mid(psProceso, 1, 2) = "99") _
'    Or (Mid(psProceso, 1, 2) = "90") Or (Mid(psProceso, 1, 2) = "40") Or (psProceso = "930099") Or (psProceso = "980000")) Then
'
'    ValidaOperacion = "57"
'    Exit Function
'End If
'
''61 Si el monto de la transaccion excede el limite de retiro por transaccion de la tarjeta y/o Cuenta
''   o si el monto acumulado de retiros excede al limite diario de retiro de la tarjeta
''Definir cuando alcance CMAC Cusco restricciones de Operaciones
'
'
''76 Si ha suspendido temporalmente sus operaciones de cambio de moneda extrangera
'
''// 910000 = Cambio de Clave
''// 01XX00 = Retiro
''// 00XX00 = Compra
''// 97XX00 = Compra Intermoneda
''// 31XX00 = Consulta de Cuenta
''// 39XX10 = Consulta de movimientos
''// 8900XX = Pago de Servicios
''// 40XXYY = Transferencia
''// 930099 = Consulta Integrada
''// 910000 = Cambio de Clave
''// 90XXYY = Transferencia Intermoneda
''// 48XXYY = Transferencia Interbancaria
''// 980000 = Consulta Tipo de Cambio
''Donde "XX" y "YY" son tipos de cuenta
'
'If pnNOOperMonExt = 1 And (Mid(psCtaCod, 9, 1) = "2" Or _
'        (Mid(psProceso, 1, 2) = "40" And Mid(psCtaDeposito, 9, 1) = "2") _
'            Or pnMoneda = 2) Then
'    ValidaOperacion = "76"
'    Exit Function
'End If
'
'
''78 Si la tarjeta no esta activa
'' Se usa la Tabla Tarjeta campo nCondicion<>1
'If pnCondicion <> 1 Then
'    ValidaOperacion = "78"
'    Exit Function
'End If
'
''79 Si algunas de las cuentas involucradas en la transaccion esta en algunas de las siguientes
''   Condiciones :
''                   Cuenta No Existe, Cuenta Bloqueada, Cuenta Cancelada, Cuenta Reasignada,
''                   Cuenta de Cobranza Judicial, Cuenta No Asociada a la Tarjeta.
'
''Jomark haga un sp que verifique estos datos
'If VerificaEstadoCuenta(psCtaCod) <> 0 Then
'    ValidaOperacion = "79"
'    Exit Function
'End If
'If Mid(psProceso, 1, 2) = "40" Then
'    If VerificaEstadoCuenta(psCtaDeposito) <> 0 Then
'        ValidaOperacion = "79"
'        Exit Function
'    End If
'End If
'
'
'
''80 Cuando por algun motivo interno del banco no puede procesar la transaccion.
'If pnSuspOper <> 0 Then
'    ValidaOperacion = "80"
'    Exit Function
'End If
'
''81 Si la tarjeta esta cancelada
''Se usa la Tabla Tarjeta campo nCondicion=50
'If pnCondicion = 50 Then
'    ValidaOperacion = "81"
'    Exit Function
'End If
'
'
'
'End Function
'
'Public Function RecuperaSaldoDeCuenta(ByVal psCtaCod As String, ByRef pnSaldoDisp As Double, _
'    ByRef pnSaldoTot As Double) As Double
'Dim Cmd As ADODB.Command
'Dim Prm As ADODB.Parameter
'
'            Set Cmd = New ADODB.Command
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 18, psCtaCod)
'            Cmd.Parameters.Append Prm
'
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pnSaldoDisponilble", adDouble, adParamOutput)
'            Cmd.Parameters.Append Prm
'
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pnSaldoContable", adDouble, adParamOutput)
'            Cmd.Parameters.Append Prm
'
'            Cmd.ActiveConnection = AbrirConexion
'            Cmd.CommandType = adCmdStoredProc
'            Cmd.CommandText = "ATM_ConsultaSaldo"
'
'            Cmd.Execute
'
'            pnSaldoDisp = Cmd.Parameters(1).Value
'            pnSaldoTot = Cmd.Parameters(2).Value
'
'            Call CerrarConexion
'            Set Cmd = Nothing
'            Set Prm = Nothing
'
'End Function
'
'Public Function RecuperaTipoCambio(ByVal pdFecha As Date, ByRef pnCompra As Double, _
'    ByRef pnVenta As Double) As Double
'
'Dim Cmd As ADODB.Command
'Dim Prm As ADODB.Parameter
'
'            Set Cmd = New ADODB.Command
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pdFecha", adDate, adParamInput, 18, pdFecha)
'            Cmd.Parameters.Append Prm
'
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pnTCCompra", adDouble, adParamOutput)
'            Cmd.Parameters.Append Prm
'
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pnTCVenta", adDouble, adParamOutput)
'            Cmd.Parameters.Append Prm
'
'            Cmd.ActiveConnection = AbrirConexion
'            Cmd.CommandType = adCmdStoredProc
'            Cmd.CommandText = "ATM_ConsultaTC"
'
'            Cmd.Execute
'
'            pnCompra = Cmd.Parameters(1).Value
'            pnVenta = Cmd.Parameters(2).Value
'
'            Call CerrarConexion
'            Set Cmd = Nothing
'            Set Prm = Nothing
'
'End Function
'
'Public Sub RegistrarCambioClave(ByVal pdFecha As Date, ByVal psNumTarjeta As String, ByVal psIDTrama As String)
'
'    Dim Cmd As ADODB.Command
'    Dim Prm As ADODB.Parameter
'
'
'            Set Cmd = New ADODB.Command
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pdFecha", adDBDate, adParamInput, , pdFecha)
'            Cmd.Parameters.Append Prm
'
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pcNumTarjeta", adVarChar, adParamInput, 50, psNumTarjeta)
'            Cmd.Parameters.Append Prm
'
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pcIDTrama", adVarChar, adParamInput, 50, psIDTrama)
'            Cmd.Parameters.Append Prm
'
'            Cmd.ActiveConnection = AbrirConexion
'            Cmd.CommandType = adCmdStoredProc
'            Cmd.CommandText = "ATM_RegistraCambioClave"
'
'            Cmd.Execute
'
'            Call CerrarConexion
'
'
'    Set Cmd = Nothing
'    Set Prm = Nothing
'
'
'End Sub
'
'
'Public Sub RegistrarTrama(ByVal psIDTrama As String, ByVal psTramas As String)
'
'    Dim Cmd As ADODB.Command
'    Dim Prm As ADODB.Parameter
'
'
'            Set Cmd = New ADODB.Command
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pcIDTrama", adVarChar, adParamInput, 50, psIDTrama)
'            Cmd.Parameters.Append Prm
'
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@cTrama", adVarChar, adParamInput, 5000, psTramas)
'            Cmd.Parameters.Append Prm
'
'
'            Cmd.ActiveConnection = AbrirConexion
'            Cmd.CommandType = adCmdStoredProc
'            Cmd.CommandText = "ATM_RegistrarTrama"
'
'            Cmd.Execute
'
'            Call CerrarConexion
'
'
'    Set Cmd = Nothing
'    Set Prm = Nothing
'
'
'End Sub
'
'Public Function RecuperaCtaDisponible(ByVal psNumTarjeta As String) As String
'Dim sCta As String
'
'
'    Dim Cmd As ADODB.Command
'    Dim Prm As ADODB.Parameter
'
'
'            Set Cmd = New ADODB.Command
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psTarjeta", adVarChar, adParamInput, 20, psNumTarjeta)
'            Cmd.Parameters.Append Prm
'
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamOutput, 18)
'            Cmd.Parameters.Append Prm
'
'
'            Cmd.ActiveConnection = AbrirConexion
'            Cmd.CommandType = adCmdStoredProc
'            Cmd.CommandText = "ATM_RecuperaCuetaDisponible"
'
'            Cmd.Execute
'
'            sCta = Cmd.Parameters(1).Value
'
'            Call CerrarConexion
'
'
'    Set Cmd = Nothing
'    Set Prm = Nothing
'
'    RecuperaCtaDisponible = sCta
'
'End Function
'
'Public Function RecuperaMovimDeCuenta(ByVal psCtaCod As String) As String
'Dim Cmd As ADODB.Command
'Dim Prm As ADODB.Parameter
'Dim R As ADODB.Recordset
'Dim nSaldoDisp As Double
'Dim nSaldoTot  As Double
'
'
'            RecuperaMovimDeCuenta = "MOVIMIENTOS;"
'            RecuperaMovimDeCuenta = "  FECHA    DESCRIPCION      MONTO  ;"
'
'            Set Cmd = New ADODB.Command
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 18, psCtaCod)
'            Cmd.Parameters.Append Prm
'
'
'            Cmd.ActiveConnection = AbrirConexion
'            Cmd.CommandType = adCmdStoredProc
'            Cmd.CommandText = "ATM_ConsultaMovCuenta"
'
'            Set R = New ADODB.Recordset
'            R.Open Cmd
'            Do While Not R.EOF
'                'Recorrer y armar la cadena
'                'fecha 5 bytes dd/mm, 8 bytes Operacion, 16 bytes Monto de Movim
'                '1 Byte para Cargo "-"
'                'Ultima linea el Saldo
'                'LSDO 2008/01/07
'                'RecuperaMovimDeCuenta = RecuperaMovimDeCuenta & Mid(Format(R!fecha, "dd/MM/yyyy"), 1, 5)
'                RecuperaMovimDeCuenta = RecuperaMovimDeCuenta & Format(R!fecha, "dd/MM/yy")
'                'LSDO 2008/01/07
'                'RecuperaMovimDeCuenta = RecuperaMovimDeCuenta & Right("                " & Replace(Format(R!Monto, "#0.00"), ".", ""), 16)
'                RecuperaMovimDeCuenta = RecuperaMovimDeCuenta & " " & Left(R!Operacion & "               ", 15) & " " & Right("          " & Format(R!Monto, "#,0.00"), 10)
'                RecuperaMovimDeCuenta = RecuperaMovimDeCuenta & ";"
'                R.MoveNext
'            Loop
'            R.Close
'
'            Call CerrarConexion
'
'            Call RecuperaSaldoDeCuenta(psCtaCod, nSaldoDisp, nSaldoTot)
'            RecuperaMovimDeCuenta = RecuperaMovimDeCuenta & ";SALDO TOTAL : " & IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & " " & Right("                 " & Format(nSaldoTot, "#,0.00"), 17)
'
'            Set Cmd = Nothing
'            Set Prm = Nothing
'            Set R = Nothing
'
'End Function
'
'Public Function RecuperaConsultaIntegrada(ByVal psNumTarjeta As String) As String
'Dim Cmd As ADODB.Command
'Dim Prm As ADODB.Parameter
'Dim R As ADODB.Recordset
'
'            RecuperaConsultaIntegrada = ""
'
'            Set Cmd = New ADODB.Command
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psNumTarjeta", adVarChar, adParamInput, 18, psNumTarjeta)
'            Cmd.Parameters.Append Prm
'
'
'            Cmd.ActiveConnection = AbrirConexion
'            Cmd.CommandType = adCmdStoredProc
'            Cmd.CommandText = "ATM_ConsultaIntegral"
'
'            RecuperaConsultaIntegrada = "CONSULTA INTEGRADA;"
'            Set R = New ADODB.Recordset
'            R.Open Cmd
'            Do While Not R.EOF
'                'Recorrer y armar la cadena
'                '3 bytes Agencia
'                '1 byte espacio en blanco
'                '15 bytes para identificar la cuenta
'                '1 byte espacio en blanco
'                '1 byte para indicar la moneda de la cuenta
'                '16 bytes para indicar el saldo disponible de la cuenta
'                '1 byte para poner "-" en caso de saldo negativo
'                '1 byte para cambio de linea emplear ";"
'                RecuperaConsultaIntegrada = RecuperaConsultaIntegrada & "0" & Mid(R!cctacod, 4, 2) & " " & Right("000000000000000" & Mid(R!cctacod, 6, 13), 15) & " "
'                RecuperaConsultaIntegrada = RecuperaConsultaIntegrada & IIf(Mid(R!cctacod, 9, 1) = "1", "S/.", "US$")
'                'LSDO 2008/01/07
'                'RecuperaConsultaIntegrada = RecuperaConsultaIntegrada & Right("                " & Replace(Format(R!SaldoDisponible, "#0.00"), ".", ""), 16)
'                RecuperaConsultaIntegrada = RecuperaConsultaIntegrada & Right("                " & Format(R!SaldoDisponible, "#,0.00"), 16)
'                RecuperaConsultaIntegrada = RecuperaConsultaIntegrada & " ;"
'
'                R.MoveNext
'            Loop
'            R.Close
'            Call CerrarConexion
'            Set Cmd = Nothing
'            Set Prm = Nothing
'            Set R = Nothing
'
'End Function
'
'
'Public Function Transaccion(ByVal psEntrada As String) As String
'Dim pINXml As String
'Dim pOUTXml As String
'Dim S As String
'Dim Cmd As New Command
'Dim cmdNeg As New Command
'Dim prmNegFecha As New ADODB.Parameter
'Dim Prm As New ADODB.Parameter
'
'Dim sCtaCod As String
'Dim nMontoTran As Double
'Dim nMontoITF As Double
'Dim nMontoComision As Double
'Dim nMontoComisionITF As Double
'Dim nMoneda As Integer
'Dim sOpeCod As String
'Dim sOpeCodComision As String
'Dim sOpeCodITF As String
'Dim nTipoCambioVenta As Double
'Dim nTipoCambioCompra As Double
'Dim nTipoCambio As Double
'Dim dFecSis As Date
'Dim sIDTrama As String
'Dim nResultado As Integer
'Dim sOpeCodComisionITF As String
'Dim sCadResp As String
'Dim sOpeExtorno As String
'Dim sOpeExtornoComision As String
'Dim sOpeCodTransferencia As String
'Dim sOpeCodExtornoTransfer As String
'Dim sCtaDeposito As String
'Dim dFecVenc As Date
'Dim PAN As String
'Dim nTarjCondicion As Integer
'Dim nRetenerTarjeta As Integer
'Dim nCtaSaldo As Double
'Dim nNOOperMonExt As Integer
'Dim nSuspOper As Integer
'Dim XmlExt As String
'
''LSDO INICIO prueba
'Dim x As Integer
'x = FreeFile
'Open "c:\test.txt" For Output As x
'Print #x, Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss AMPM") & " Inicio de operación Retiro Normal"
'Close x
'
''FIN
'
'
'    pINXml = psEntrada
'
''    RetiroNormal = "Prueba"
''    Exit Function
''
'    'Call RegistrarTrama(sIDTrama, pINXml)
'
'    '*****************************************
'    'RECUPERA TRAMA DE BASE
'    'SOLO PARA PRUEBAS
'    '*****************************************
''    Dim sTramaBase As String
''    Set cmdNeg = New ADODB.Command
''    Set prmNegFecha = New ADODB.Parameter
''    Set prmNegFecha = cmdNeg.CreateParameter("@XML", adVarChar, adParamOutput, 5000)
''    cmdNeg.Parameters.Append prmNegFecha
''
''    cmdNeg.ActiveConnection = AbrirConexion
''    cmdNeg.CommandType = adCmdStoredProc
''    cmdNeg.CommandText = "ATM_RecuepraTRamaBase"
''
''    cmdNeg.Execute
''
''    sTramaBase = cmdNeg.Parameters(0).Value
''    pINXml = sTramaBase
'    '**********************************************************************
'
'     If RecuperaValorXML(pINXml, "MESSAGE_TYPE") = "0800" Then
'        Call RegistraSucesos(Now, "Transaccion ECO: " & RecuperaValorXML(pINXml, "TRACE"), "")
'        Call RegistrarTrama(RecuperaValorXML(pINXml, "TRACE"), pINXml)
'
'        pOUTXml = GeneraXMLSalida("00", , "123456")
'        Transaccion = pOUTXml
'        Exit Function
'     End If
'
'
'    Call RegistrarTrama(RecuperaValorXML(pINXml, "TRACE"), pINXml)
'    Call RegistraSucesos(Now, "Registro Trama : " & RecuperaValorXML(pINXml, "TRACE"), "")
'
'    If RecuperaValorXML(pINXml, "MESSAGE_TYPE") = "0400" Then
'
'            '*****************************************
'            'RECUPERA TRAMA DE BASE
'            '*****************************************
'            Set prmNegFecha = Nothing
'            Set cmdNeg = Nothing
'            Set cmdNeg = New ADODB.Command
'
'            Set prmNegFecha = New ADODB.Parameter
'            Set prmNegFecha = cmdNeg.CreateParameter("@cID", adVarChar, adParamInput, 50)
'            prmNegFecha.Value = RecuperaValorXML(pINXml, "TRACE")
'            cmdNeg.Parameters.Append prmNegFecha
'
'            Set prmNegFecha = New ADODB.Parameter
'            Set prmNegFecha = cmdNeg.CreateParameter("@XML", adVarChar, adParamOutput, 5000)
'            cmdNeg.Parameters.Append prmNegFecha
'
'            cmdNeg.ActiveConnection = AbrirConexion
'            cmdNeg.CommandType = adCmdStoredProc
'            cmdNeg.CommandText = "ATM_RecuperaTramaBaseExt"
'
'            cmdNeg.Execute
'
'            XmlExt = cmdNeg.Parameters(1).Value
'            pINXml = XmlExt
'            pINXml = Replace(pINXml, "<MESSAGE_TYPE value=""200""/>", "<MESSAGE_TYPE value=""400""/>")
'
'    End If
'
'
'    Call RegistraSucesos(Now, "Inicio de Transaccion : " & RecuperaValorXML(pINXml, "TRACE"), "")
'
'
'    S = RecuperaValorXML(pINXml, "PRCODE")
'    If Len(S) <> 6 Then
'            S = Right("000000" & S, 6)
'    End If
'
'    PAN = Trim(RecuperaValorXML(pINXml, "PAN"))
'
'    '*****************************************
'    'RECUPERA DATOS DEL NEGOCIO
'    '*****************************************
'    Set cmdNeg = New ADODB.Command
'    Set prmNegFecha = New ADODB.Parameter
'    Set prmNegFecha = cmdNeg.CreateParameter("@dFecSis", adDBDate, adParamOutput)
'    cmdNeg.Parameters.Append prmNegFecha
'
'    Set prmNegFecha = New ADODB.Parameter
'    Set prmNegFecha = cmdNeg.CreateParameter("@nTipoCambioVenta", adDouble, adParamOutput)
'    cmdNeg.Parameters.Append prmNegFecha
'
'    Set prmNegFecha = New ADODB.Parameter
'    Set prmNegFecha = cmdNeg.CreateParameter("@nTipoCambioCompra", adDouble, adParamOutput)
'    cmdNeg.Parameters.Append prmNegFecha
'
'    cmdNeg.ActiveConnection = AbrirConexion
'    cmdNeg.CommandType = adCmdStoredProc
'    cmdNeg.CommandText = "ATM_RecuperaDatosNegocio"
'
'    cmdNeg.Execute
'
'    dFecSis = cmdNeg.Parameters(0).Value
'    'Coordinar operacion de madrugada y fecha de sistema aun sigue con dia de ayer por falgta de cierre de dia
'    dFecSis = CDate(Format(dFecSis, "dd/MM/yyyy") & " " & Mid(Format(Now(), "dd/MM/yyyy hh:mm:ss"), 12, 8))
'    nTipoCambioVenta = cmdNeg.Parameters(1).Value
'    nTipoCambioCompra = cmdNeg.Parameters(2).Value
'    Call CerrarConexion
'
'    Call RegistraSucesos(Now, "Recupero datos del Negocio : " & RecuperaValorXML(pINXml, "TRACE"), "")
'
'    Set cmdNeg = Nothing
'    Set prmNegFecha = Nothing
'
'    If S = "910000" Or S = "930099" Or S = "980000" Then
'        sCtaCod = RecuperaCtaDisponible(PAN)
'    Else
'        sCtaCod = Trim("106" & Right(RecuperaValorXML(pINXml, "ACCT_1"), 2) & Mid(Trim(RecuperaValorXML(pINXml, "ACCT_1")), 1, 13))
'    End If
'
'    Dim nTipoEquipo As Integer
'    Dim nTipoServicio As Integer
'    Dim nProced As Integer
'    Dim nTipoOperac As Integer
'    nTipoEquipo = CInt(IIf(Mid(S, 1, 2) <> "00" And Mid(S, 1, 2) <> "97", 1, 2))
'    'nTipoEquipo = CInt(IIf(Mid(S, 1, 2) = "01", 1, 2))
'    If RecuperaValorXML(pINXml, "CARD_ACCEPTOR") <> "000000000000000" Then
'        nProced = IIf(Right(RecuperaValorXML(pINXml, "CARD_LOCATION"), 2) = "PE", 1, 2)
'    Else
'        nProced = 1
'    End If
'
'    nTipoServicio = IIf(Trim(RecuperaValorXML(pINXml, "ACQ_INST")) = "426154", 1, IIf(CInt(RecuperaValorXML(pINXml, "POS_COND_CODE")) = 2, 2, 3))
'
'    If Mid(S, 1, 2) = "31" Or Mid(S, 1, 2) = "39" Or Mid(S, 1, 2) = "93" Or Mid(S, 1, 2) = "98" Then
'        nTipoOperac = 2
'    Else
'        nTipoOperac = 1
'    End If
'
'    nMontoTran = DE_TRAMA_ConvierteAMontoReal(RecuperaValorXML(pINXml, "TXN_AMOUNT"))
'    nMontoITF = CDbl(Format(ObtieneITF * DE_TRAMA_ConvierteAMontoReal(RecuperaValorXML(pINXml, "TXN_AMOUNT")), "#0.00"))
'
'    nMontoComision = ObtieneComisionPorOperacion(nTipoEquipo, nTipoServicio, nProced, _
'        IIf(RecuperaValorXML(pINXml, "CUR_CODE") = "604", 1, 2), nTipoOperac, DE_TRAMA_ConvierteAMontoReal(RecuperaValorXML(pINXml, "TXN_AMOUNT")))
'
'    nMontoComisionITF = CDbl(Format(ObtieneITF * nMontoComision, "#0.00"))
'    nMoneda = IIf(RecuperaValorXML(pINXml, "CUR_CODE") = "604", 1, 2)
'
'    'RETIRO POR CAJEROS
'    'david
'    'If Mid(S, 1, 2) = "01" Then
'    If Mid(S, 1, 2) <> "00" And Mid(S, 1, 2) <> "97" Then
'
'        If CInt(RecuperaValorXML(pINXml, "POS_COND_CODE")) = 2 Then
'            sOpeCod = "200361" 'Retiro por ATM Unibanca
'            sOpeCodComision = "200362" 'comision
'            sOpeCodITF = "200363" 'ITF Comision
'            sOpeCodComisionITF = "200364"
'            sOpeExtorno = "200373"
'            sOpeExtornoComision = "200376"
'            sOpeCodTransferencia = "200379"
'            sOpeCodExtornoTransfer = "200381"
'
'        End If
'        If RecuperaValorXML(pINXml, "POS_COND_CODE") = "51" Then
'            sOpeCod = "200365" 'Retiro por OTRAS REDES VISA
'            sOpeCodComision = "200366" 'comision
'            sOpeCodITF = "200367" 'ITF Comision
'            sOpeCodComisionITF = "200368"
'            sOpeExtorno = "200374"
'            sOpeExtornoComision = "200377"
'            sOpeCodTransferencia = "200380"
'            sOpeCodExtornoTransfer = "200382"
'        End If
'        If RecuperaValorXML(pINXml, "POS_COND_CODE") = "64" Then
'            sOpeCod = "200369" 'Retiro por Ventanilla
'            sOpeCodComision = "200370" 'comision
'            sOpeCodITF = "200371" 'Itf
'            sOpeCodComisionITF = "200372" 'ITF Comision
'            sOpeExtorno = "200375"
'            sOpeExtornoComision = "200378"
'            sOpeCodTransferencia = ""
'            sOpeCodExtornoTransfer = ""
'        End If
'
'    End If
'
'    'COMPRAS POR POS
'    'david
'    'If Mid(S, 1, 2) = "00" Then
'    If Mid(S, 1, 2) = "00" Or Mid(S, 1, 2) = "97" Then
'            sOpeCod = "200391" 'Retiro por ATM Unibanca
'            sOpeCodComision = "200392" 'comision
'            sOpeCodITF = "200393" 'ITF Comision
'            sOpeCodComisionITF = "200394"
'            sOpeExtorno = "200379"
'            sOpeExtornoComision = "200380"
'            sOpeCodTransferencia = ""
'            sOpeCodExtornoTransfer = ""
'    End If
'
'    'Si operacion es de diferente Moneda
'    If nMoneda <> CInt(Mid(sCtaCod, 9, 1)) Then
'        'Si es Soles
'        If nMoneda = 1 Then
'            nTipoCambio = nTipoCambioCompra
'        Else
'            nTipoCambio = nTipoCambioVenta
'        End If
'    End If
'
'    sIDTrama = RecuperaValorXML(pINXml, "TRACE")
'    'sCtaDeposito = RecuperaValorXML(pINXml, "ACCT_2")
'    sCtaDeposito = Trim("106" & Right(RecuperaValorXML(pINXml, "ACCT_2"), 2) & Mid(RecuperaValorXML(pINXml, "ACCT_2"), 1, 13))
'
'    dFecVenc = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Mid(RecuperaValorXML(pINXml, "DATE_EXP"), 3, 2) & "/20" & Mid(RecuperaValorXML(pINXml, "DATE_EXP"), 1, 2))))
'
'
'    '***************************************************
'    'RECUPERA DATOS DE TARJETA
'    '***************************************************
'    Set cmdNeg = New ADODB.Command
'    Set prmNegFecha = New ADODB.Parameter
'    Set prmNegFecha = cmdNeg.CreateParameter("@PAN", adVarChar, adParamInput, 20, PAN)
'    cmdNeg.Parameters.Append prmNegFecha
'
'    Set prmNegFecha = New ADODB.Parameter
'    Set prmNegFecha = cmdNeg.CreateParameter("@nCondicion", adInteger, adParamOutput)
'    cmdNeg.Parameters.Append prmNegFecha
'
'    Set prmNegFecha = New ADODB.Parameter
'    Set prmNegFecha = cmdNeg.CreateParameter("@nRetenerTarjeta", adInteger, adParamOutput)
'    cmdNeg.Parameters.Append prmNegFecha
'
'    Set prmNegFecha = New ADODB.Parameter
'    Set prmNegFecha = cmdNeg.CreateParameter("@nNOOperMonExt", adInteger, adParamOutput)
'    cmdNeg.Parameters.Append prmNegFecha
'
'    Set prmNegFecha = New ADODB.Parameter
'    Set prmNegFecha = cmdNeg.CreateParameter("@nSuspOper", adInteger, adParamOutput)
'    cmdNeg.Parameters.Append prmNegFecha
'
'    cmdNeg.ActiveConnection = AbrirConexion
'    cmdNeg.CommandType = adCmdStoredProc
'    cmdNeg.CommandText = "ATM_RecuperaDatosTarjeta"
'
'    cmdNeg.Execute
'
'    nTarjCondicion = cmdNeg.Parameters(1).Value
'    nRetenerTarjeta = cmdNeg.Parameters(2).Value
'    nNOOperMonExt = cmdNeg.Parameters(3).Value
'    nSuspOper = cmdNeg.Parameters(4).Value
'    Call CerrarConexion
'
'    Set cmdNeg = Nothing
'    Set prmNegFecha = Nothing
'
'    Call RegistraSucesos(Now, "Recupero datos de la Tarjeta : " & RecuperaValorXML(pINXml, "TRACE"), "")
'
'    '**********************************************************************
'    'RECUPERA DATOS DE CUENTA
'    '**********************************************************************
'    Set cmdNeg = New ADODB.Command
'    Set prmNegFecha = New ADODB.Parameter
'    Set prmNegFecha = cmdNeg.CreateParameter("@psCtaCod", adVarChar, adParamInput, 20, sCtaCod)
'    cmdNeg.Parameters.Append prmNegFecha
'
'    Set prmNegFecha = New ADODB.Parameter
'    Set prmNegFecha = cmdNeg.CreateParameter("@nSaldo", adInteger, adParamOutput)
'    cmdNeg.Parameters.Append prmNegFecha
'
'    cmdNeg.ActiveConnection = AbrirConexion
'    cmdNeg.CommandType = adCmdStoredProc
'    cmdNeg.CommandText = "ATM_RecuperaDatosCuenta"
'
'    cmdNeg.Execute
'
'    nCtaSaldo = cmdNeg.Parameters(1).Value
'
'    Call CerrarConexion
'
'    Set cmdNeg = Nothing
'    Set prmNegFecha = Nothing
'
'
'    '**********************************************************************
'    Call RegistraSucesos(Now, "Recupero datos de la cuenta : " & RecuperaValorXML(pINXml, "TRACE"), "")
'
'    '*****************************************
'    'VALIDACION DE OPERACION
'    '*****************************************
'    Dim sValida As String
'    sValida = ValidaOperacion(sCtaCod, nMontoTran, nMontoITF, nMontoComision, nMontoComisionITF, _
'    nMoneda, sOpeCod, sOpeCodComision, sOpeCodITF, sOpeCodComisionITF, sOpeExtorno, _
'    sOpeExtornoComision, sOpeCodTransferencia, sOpeCodExtornoTransfer, _
'    nTipoCambioCompra, nTipoCambioVenta, sIDTrama, sCtaDeposito, _
'    PAN, dFecSis, dFecVenc, nTarjCondicion, nRetenerTarjeta, nCtaSaldo, S, nNOOperMonExt, nSuspOper)
'
'    If sValida <> "00" Then
'        pOUTXml = GeneraXMLSalida(sValida)
'        Transaccion = pOUTXml
'        Exit Function
'    End If
'
'    '**************************************************************************************************
'    'REGISTRO DE TRAMA
'    '**************************************************************************************************
'    Call RegistraSucesos(Now, "Valido la Operacion : " & RecuperaValorXML(pINXml, "TRACE"), "")
'
'
'    '**************************************************************************************************
'    'TRANSFERENCIA, TRANSFERENCIA INTERMONEDA, RETIRO POR CAJERO, COMPRA Y COMPRA INTERMONEDA
'    '**************************************************************************************************
'    Call RegistraSucesos(Now, "Inicio de Operacion : " & RecuperaValorXML(pINXml, "TRACE"), "")
'
'    If ((Mid(S, 1, 2) = "01" And Mid(S, 5, 2) = "00") Or (Mid(S, 1, 2) = "00" And Mid(S, 5, 2) = "00") _
'        Or (Mid(S, 1, 2) = "97" And Mid(S, 5, 2) = "00") Or (Mid(S, 1, 2) = "40") Or (Mid(S, 1, 2) = "90") Or (Mid(S, 1, 2) = "99")) _
'        And RecuperaValorXML(pINXml, "MESSAGE_TYPE") <> "0400" Then
'
'        '********************************
'        'INICIA PROCESO DE RETIRO
'        '********************************
'        nResultado = Retiro(dFecSis, sCtaCod, nMontoTran, nMontoITF, nMontoComision, nMontoComisionITF, nMoneda, _
'            sOpeCod, sOpeCodComision, sOpeCodITF, sOpeCodComisionITF, nTipoCambio, sIDTrama, 0)
'
'
'        '******************************************************
'        'TRANFERENCIA
'        '******************************************************
'        If (Mid(S, 1, 2) = "40") Or (Mid(S, 1, 2) = "90") Then
'
'            'Si operacion es de diferente Moneda
'            If nMoneda <> CInt(Mid(sCtaDeposito, 9, 1)) Then
'                'Si es Soles
'                If nMoneda = 1 Then
'                    nTipoCambio = nTipoCambioCompra
'                Else
'                    nTipoCambio = nTipoCambioVenta
'                End If
'            End If
'
'            Set Cmd = New ADODB.Command
'            '0
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pdFecha", adDBDate, adParamInput, , dFecSis)
'            Cmd.Parameters.Append Prm
'            '1
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 18, sCtaDeposito)
'            Cmd.Parameters.Append Prm
'            '2
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pnMonto", adDouble, adParamInput, , nMontoTran)
'            Cmd.Parameters.Append Prm
'
'            '3
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pnMontoITF", adDouble, adParamInput, , nMontoITF)
'            Cmd.Parameters.Append Prm
'
'            '4
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pnMontoComision", adDouble, adParamInput, , nMontoComision)
'            Cmd.Parameters.Append Prm
'            '5
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pnMontoComisionITF", adDouble, adParamInput, , nMontoComisionITF)
'            Cmd.Parameters.Append Prm
'            '6
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pnMoneda", adSmallInt, adParamInput, , nMoneda)
'            Cmd.Parameters.Append Prm
'
'            '7
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psOpeCod", adVarChar, adParamInput, 6, sOpeCodTransferencia)
'            Cmd.Parameters.Append Prm
'            '8
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psOpeCodComision", adVarChar, adParamInput, 6, sOpeCodComision)
'            Cmd.Parameters.Append Prm
'            '9
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psOpeCodITF", adVarChar, adParamInput, 6, sOpeCodITF)
'            Cmd.Parameters.Append Prm
'
'            '10
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("sOpeCodComisionITF", adVarChar, adParamInput, 6, sOpeCodComisionITF)
'            Cmd.Parameters.Append Prm
'
'            '11
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pnTipoCambio", adDouble, adParamInput, 6, nTipoCambio)
'            Cmd.Parameters.Append Prm
'            '12
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psIDTrama", adVarChar, adParamInput, 5000, sIDTrama)
'            Cmd.Parameters.Append Prm
'            '13
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pnResultado", adSmallInt, adParamOutput, , nResultado)
'            Cmd.Parameters.Append Prm
'
'            '14
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pnMovNroOffHost", adSmallInt, adParamInput, , 0)
'            Cmd.Parameters.Append Prm
'
'            Cmd.ActiveConnection = AbrirConexion
'            Cmd.CommandType = adCmdStoredProc
'            Cmd.CommandText = "ATM_Deposito"
'            Cmd.Execute
'            Call CerrarConexion
'            nResultado = Cmd.Parameters(13).Value
'        End If
'
'        'GENERA RESPUESTA PARA ATM
'        If (Mid(S, 1, 2) = "01" And Mid(S, 5, 2) = "00") Or (Mid(S, 1, 2) = "99") Then
'            sCadResp = "1001" & RecuperaValorXML(pINXml, "CUR_CODE") & "C" & Right("000000000000" & Replace(Trim(Format(RecuperaSaldoDisp(sCtaCod), "0.00")), ".", ""), 12)
'            sCadResp = sCadResp & "1002" & RecuperaValorXML(pINXml, "CUR_CODE") & "C" & Right("000000000000" & Replace(Trim(Format(RecuperaSaldoDisp(sCtaCod), "0.00")), ".", ""), 12)
'            pOUTXml = GeneraXMLSalida("00", , "123456", , , sCadResp)
'        End If
'
'         'GENERA RESPUESTA PARA TRANSFERENCIA
'        If (Mid(S, 1, 2) = "40") Or (Mid(S, 1, 2) = "90") Then
'            sCadResp = "1001" & RecuperaValorXML(pINXml, "CUR_CODE") & "C" & Right("000000000000" & Replace(Trim(Str(RecuperaSaldoDisp(sCtaCod))), ".", ""), 12)
'            sCadResp = sCadResp & "1002" & RecuperaValorXML(pINXml, "CUR_CODE") & "C" & Right("000000000000" & Replace(Trim(Str(RecuperaSaldoDisp(sCtaCod))), ".", ""), 12)
'            pOUTXml = GeneraXMLSalida("00", "000000000000", "123456", , RecuperaValorXML(pINXml, "CUR_CODE"), sCadResp)
'        End If
'
'
'        'GENERA RESPUESTA PARA COMPRA y COMPRA INTERMONEDA
'        If (Mid(S, 1, 2) = "00" And Mid(S, 5, 2) = "00") Or (Mid(S, 1, 2) = "97" And Mid(S, 5, 2) = "00") Then
'            pOUTXml = GeneraXMLSalida("00", Right("000000000000" & Replace(Str(Trim(nMontoTran)), ".", ""), 12), "123456", , "000")
'        End If
'        Transaccion = pOUTXml
'        Exit Function
'    End If
'
'    '*****************************************************************************
'    'CONSULTAS : DE CUENTA
'    '*****************************************************************************
'    If RecuperaValorXML(pINXml, "MESSAGE_TYPE") = "0200" _
'        And (Mid(S, 1, 2) = "31" And Mid(S, 5, 2) = "00") Then
'        Dim nSaldoCtaDisp As Double
'        Dim nSaldoCtaTot As Double
'        Dim sCadRespSalCta As String
'
'        nResultado = Retiro(dFecSis, sCtaCod, nMontoTran, nMontoITF, nMontoComision, nMontoComisionITF, nMoneda, _
'            sOpeCod, sOpeCodComision, sOpeCodITF, sOpeCodComisionITF, nTipoCambio, sIDTrama, 0)
'
'        Call RecuperaSaldoDeCuenta(sCtaCod, nSaldoCtaDisp, nSaldoCtaTot)
'
'        'SALDO TOTAL
'        sCadRespSalCta = "1001"
'        If Mid(sCtaCod, 9, 1) = "1" Then
'            sCadRespSalCta = sCadRespSalCta & "604"
'        Else
'            sCadRespSalCta = sCadRespSalCta & "840"
'        End If
'        sCadRespSalCta = sCadRespSalCta & "C"
'        sCadRespSalCta = sCadRespSalCta & Right("000000000000" & Replace(Format(nSaldoCtaTot, "#0.00"), ".", ""), 12)
'
'        'SALDO DISPONIBLE
'        sCadRespSalCta = sCadRespSalCta & "1002"
'        If Mid(sCtaCod, 9, 1) = "1" Then
'            sCadRespSalCta = sCadRespSalCta & "604"
'        Else
'            sCadRespSalCta = sCadRespSalCta & "840"
'        End If
'        sCadRespSalCta = sCadRespSalCta & "C"
'        sCadRespSalCta = sCadRespSalCta & Right("000000000000" & Replace(Format(nSaldoCtaDisp, "#0.00"), ".", ""), 12)
'
'
'        pOUTXml = GeneraXMLSalida("00", , "123456", , , sCadRespSalCta)
'        Transaccion = pOUTXml
'        Exit Function
'
'    End If
'
'    '*****************************************************************************
'    'CONSULTAS : DE MOVIMIENTOS
'    '*****************************************************************************
'    If RecuperaValorXML(pINXml, "MESSAGE_TYPE") = "0200" _
'        And (Mid(S, 1, 2) = "39") Then
'
'        nResultado = Retiro(dFecSis, sCtaCod, nMontoTran, nMontoITF, nMontoComision, nMontoComisionITF, nMoneda, _
'            sOpeCod, sOpeCodComision, sOpeCodITF, sOpeCodComisionITF, nTipoCambio, sIDTrama, 0)
'
'        pOUTXml = GeneraXMLSalida("00", "000000000000", "123456", , "000", , RecuperaMovimDeCuenta(sCtaCod))
'        Transaccion = pOUTXml
'        Exit Function
'    End If
'
'    '*****************************************************************************
'    'CONSULTAS : INTEGRADA
'    '*****************************************************************************
'    If RecuperaValorXML(pINXml, "MESSAGE_TYPE") = "0200" _
'        And (Mid(S, 1, 2) = "93" And Mid(S, 5, 2) = "99") Then
'
'        nResultado = Retiro(dFecSis, sCtaCod, nMontoTran, nMontoITF, nMontoComision, nMontoComisionITF, nMoneda, _
'            sOpeCod, sOpeCodComision, sOpeCodITF, sOpeCodComisionITF, nTipoCambio, sIDTrama, 0)
'
'        pOUTXml = GeneraXMLSalida("00", "000000000000", "123456", , "000", , RecuperaConsultaIntegrada(PAN))
'        Transaccion = pOUTXml
'        Exit Function
'    End If
'
'    '*****************************************************************************
'    'CONSULTAS : TIPO DE CAMBIO
'    '*****************************************************************************
'    If RecuperaValorXML(pINXml, "MESSAGE_TYPE") = "0200" _
'        And (Mid(S, 1, 2) = "98" And Mid(S, 5, 2) = "00") Then
'
'        nResultado = Retiro(dFecSis, sCtaCod, nMontoTran, nMontoITF, nMontoComision, nMontoComisionITF, nMoneda, _
'            sOpeCod, sOpeCodComision, sOpeCodITF, sOpeCodComisionITF, nTipoCambio, sIDTrama, 0)
'
'        Dim nConsTipCamCompra As Double
'        Dim nConsTipCamVenta As Double
'        Dim sCadRespConsTipCam As String
'
'        Call RecuperaTipoCambio(dFecSis, nConsTipCamCompra, nConsTipCamVenta)
'
'        sCadRespConsTipCam = "CMAC CUSCO SA;"
'        sCadRespConsTipCam = sCadRespConsTipCam & " COMPRA DE US $: " & Format(nConsTipCamCompra, "#0.00") & ";"
'        sCadRespConsTipCam = sCadRespConsTipCam & " VENTA DE US $: " & Format(nConsTipCamVenta, "#0.00") & ";"
'        sCadRespConsTipCam = sCadRespConsTipCam & " TIPO DE CAMBIO SUJETO A VARIACION;"
'
'        pOUTXml = GeneraXMLSalida("00", , "123456", , , , sCadRespConsTipCam)
'        Transaccion = pOUTXml
'        Exit Function
'
'    End If
'
'
'    '*****************************************************************************
'    'CAMBIO DE CLAVE
'    '*****************************************************************************
'    If RecuperaValorXML(pINXml, "MESSAGE_TYPE") = "0200" _
'        And (Mid(S, 1, 2) = "91" And Mid(S, 5, 2) = "00") Then
'
'
'        nResultado = Retiro(dFecSis, sCtaCod, nMontoTran, nMontoITF, nMontoComision, nMontoComisionITF, nMoneda, _
'            sOpeCod, sOpeCodComision, sOpeCodITF, sOpeCodComisionITF, nTipoCambio, sIDTrama, 0)
'
'        Call RegistrarCambioClave(dFecSis, PAN, sIDTrama)
'
'        pOUTXml = GeneraXMLSalida("00", , "123456", , , "")
'        Transaccion = pOUTXml
'        Exit Function
'
'    End If
'    '*****************************************************************************
'
'
'
'    '*******************************************************
'    'EXTORNO TRANSFERENCIA, RETIRO DE CAJERO, COMPRA Y COMPRA INTERMONEDA
'    '*******************************************************
'    If ((Mid(S, 1, 2) = "01" And Mid(S, 5, 2) = "00") Or (Mid(S, 1, 2) = "00" And Mid(S, 5, 2) = "00") _
'        Or (Mid(S, 1, 2) = "97" And Mid(S, 5, 2) = "00") Or (Mid(S, 1, 2) = "40") Or (Mid(S, 1, 2) = "99")) _
'        And RecuperaValorXML(pINXml, "MESSAGE_TYPE") = "0400" Then
'
'        Dim nMovNro As Long
'
'        '************************************
'        'RECUPERA MOV DE LA TRAMA A EXTORNAR
'        '************************************
'        Set Cmd = New ADODB.Command
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@psNumTrama", adVarChar, adParamInput, 50, sIDTrama)
'        Cmd.Parameters.Append Prm
'
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@pcCtaCod", adVarChar, adParamInput, 18, sCtaCod)
'        Cmd.Parameters.Append Prm
'
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@pnMovnro", adInteger, adParamOutput, , nMovNro)
'        Cmd.Parameters.Append Prm
'
'
'        Cmd.ActiveConnection = AbrirConexion
'        Cmd.CommandType = adCmdStoredProc
'        Cmd.CommandText = "ATM_RecuperaMovDeTrama"
'        Cmd.Execute
'
'        nMovNro = Cmd.Parameters(2).Value
'
'        '********************************
'        'INICIA PROCESO DE RETIRO
'        '********************************
'
'        Set Cmd = New ADODB.Command
'        '0
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@pnMovNro", adInteger, adParamInput, , nMovNro)
'        Cmd.Parameters.Append Prm
'        '1
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@psOpeRetiro", adVarChar, adParamInput, 6, sOpeCod)
'        Cmd.Parameters.Append Prm
'
'        '2
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@psOpeRetiroExtorno", adVarChar, adParamInput, 6, sOpeExtorno)
'        Cmd.Parameters.Append Prm
'
'        '3
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@psOpeRetiroITF", adVarChar, adParamInput, 6, sOpeCodITF)
'        Cmd.Parameters.Append Prm
'
'        '4
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@psOpeRetiroComision", adVarChar, adParamInput, 6, sOpeCodComision)
'        Cmd.Parameters.Append Prm
'
'        '5
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@psOpeRetiroComisionExtorno", adVarChar, adParamInput, 6, sOpeExtornoComision)
'        Cmd.Parameters.Append Prm
'
'
'        '6
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@psOpeRetiroComisionITF", adVarChar, adParamInput, 6, sOpeCodComisionITF)
'        Cmd.Parameters.Append Prm
'
'        '7
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@pnResultado", adVarChar, adParamInput, 6, nResultado)
'        Cmd.Parameters.Append Prm
'
'        Cmd.ActiveConnection = C
'        Cmd.CommandType = adCmdStoredProc
'        Cmd.CommandText = "ATM_ExtornoRetiro"
'        Cmd.Execute
'
'        nResultado = Cmd.Parameters(7).Value
'
'        Call CerrarConexion
'
'        '******************************************************
'        'EXTORNO DE TRANFERENCIA
'        '******************************************************
'        If (Mid(S, 1, 2) = "40") Then
'            '************************************
'            'RECUPERA MOV DE LA TRAMA A EXTORNAR
'            '************************************
'
'            Set Cmd = Nothing
'            Set Prm = Nothing
'
'            Set Cmd = New ADODB.Command
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psNumTrama", adVarChar, adParamInput, 50, sIDTrama)
'            Cmd.Parameters.Append Prm
'
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pcCtaCod", adVarChar, adParamInput, 18, sCtaDeposito)
'            Cmd.Parameters.Append Prm
'
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pnMovnro", adInteger, adParamOutput, , nMovNro)
'            Cmd.Parameters.Append Prm
'
'            Cmd.ActiveConnection = AbrirConexion
'            Cmd.CommandType = adCmdStoredProc
'            Cmd.CommandText = "ATM_RecuperaMovDeTrama"
'            Cmd.Execute
'
'            nMovNro = Cmd.Parameters(2).Value
'
'            '********************************
'            'INICIA PROCESO DE RETIRO
'            '********************************
'
'            Set Cmd = New ADODB.Command
'            '0
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pnMovNro", adInteger, adParamInput, , nMovNro)
'            Cmd.Parameters.Append Prm
'            '1
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psOpeDeposito", adVarChar, adParamInput, 6, sOpeCodTransferencia)
'            Cmd.Parameters.Append Prm
'
'            '2
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psOpeDepositoExtorno", adVarChar, adParamInput, 6, sOpeCodExtornoTransfer)
'            Cmd.Parameters.Append Prm
'
'            '3
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psOpeDepositoITF", adVarChar, adParamInput, 6, sOpeCodITF)
'            Cmd.Parameters.Append Prm
'
'            '4
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psOpeDepositoComision", adVarChar, adParamInput, 6, sOpeCodComision)
'            Cmd.Parameters.Append Prm
'
'            '5
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psOpeDepositoComisionExtorno", adVarChar, adParamInput, 6, sOpeExtornoComision)
'            Cmd.Parameters.Append Prm
'
'
'            '6
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@psOpeDepositoComisionITF", adVarChar, adParamInput, 6, sOpeCodComisionITF)
'            Cmd.Parameters.Append Prm
'
'            '7
'            Set Prm = New ADODB.Parameter
'            Set Prm = Cmd.CreateParameter("@pnResultado", adVarChar, adParamInput, 6, nResultado)
'            Cmd.Parameters.Append Prm
'
'            Cmd.ActiveConnection = C
'            Cmd.CommandType = adCmdStoredProc
'            Cmd.CommandText = "ATMExtornoDeposito"
'            Cmd.Execute
'
'            nResultado = Cmd.Parameters(7).Value
'
'            Call CerrarConexion
'        End If
'
'        pOUTXml = GeneraXMLSalida("00", , "123456")
'        Transaccion = pOUTXml
'        Exit Function
'    End If
'    Transaccion = pOUTXml
'End Function

' Public Sub Instancia(ByVal pINXml As String, pOUTXml As String)
'
' End Sub
'
'Public Function Retiro(ByVal pdFecSis As Date, ByVal psCtaCod As String, ByVal pnMontoTran As Double, _
'    ByVal pnMontoITF As Double, ByVal pnMontoComision As Double, ByVal pnMontoComisionITF As Double, _
'    ByVal pnMoneda As Integer, ByVal psOpeCod As String, ByVal psOpeCodComision As String, ByVal psOpeCodITF As String, _
'    ByVal psOpeCodComisionITF As String, ByVal pnTipoCambio As Double, ByVal psIDTrama As String, _
'    ByVal pnoffHost As Integer) As Integer
'
'Dim Cmd As ADODB.Command
'Dim Prm As ADODB.Parameter
'
'        Dim nResultado As Integer
'
'        Set Cmd = New ADODB.Command
'        '0
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@pdFecha", adDBDate, adParamInput, , pdFecSis)
'        Cmd.Parameters.Append Prm
'        '1
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 18, psCtaCod)
'        Cmd.Parameters.Append Prm
'        '2
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@pnMonto", adDouble, adParamInput, , pnMontoTran)
'        Cmd.Parameters.Append Prm
'
'        '3
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@pnMontoITF", adDouble, adParamInput, , pnMontoITF)
'        Cmd.Parameters.Append Prm
'
'        '4
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@pnMontoComision", adDouble, adParamInput, , pnMontoComision)
'        Cmd.Parameters.Append Prm
'        '5
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@pnMontoComisionITF", adDouble, adParamInput, , pnMontoComisionITF)
'        Cmd.Parameters.Append Prm
'        '6
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@pnMoneda", adSmallInt, adParamInput, , pnMoneda)
'        Cmd.Parameters.Append Prm
'
'        '7
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@psOpeCod", adVarChar, adParamInput, 6, psOpeCod)
'        Cmd.Parameters.Append Prm
'        '8
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@psOpeCodComision", adVarChar, adParamInput, 6, psOpeCodComision)
'        Cmd.Parameters.Append Prm
'        '9
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@psOpeCodITF", adVarChar, adParamInput, 6, psOpeCodITF)
'        Cmd.Parameters.Append Prm
'
'        '10
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("sOpeCodComisionITF", adVarChar, adParamInput, 6, psOpeCodComisionITF)
'        Cmd.Parameters.Append Prm
'
'        '11
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@pnTipoCambio", adDouble, adParamInput, 6, pnTipoCambio)
'        Cmd.Parameters.Append Prm
'        '12
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@psIDTrama", adVarChar, adParamInput, 5000, psIDTrama)
'        Cmd.Parameters.Append Prm
'        '13
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@pnResultado", adSmallInt, adParamOutput, , nResultado)
'        Cmd.Parameters.Append Prm
'
'        '14
'        Set Prm = New ADODB.Parameter
'        Set Prm = Cmd.CreateParameter("@pnMovNroOffHost", adSmallInt, adParamInput, , pnoffHost)
'        Cmd.Parameters.Append Prm
'
'        Cmd.ActiveConnection = AbrirConexion
'        Cmd.CommandType = adCmdStoredProc
'        Cmd.CommandText = "ATM_Retiro"
'        Cmd.Execute
'        Call CerrarConexion
'        Retiro = Cmd.Parameters(13).Value
'
'End Function
'
'
'
'
