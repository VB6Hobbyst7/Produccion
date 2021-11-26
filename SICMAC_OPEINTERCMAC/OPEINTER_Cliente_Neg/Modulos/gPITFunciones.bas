Attribute VB_Name = "gPITFunciones"
Option Explicit

Global Const gsPATHTramas As String = "\Tramas\"

'Global Const gsIPRQ As String = "192.168.0.21:81" '"192.168.15.35:81"
'Global Const gsCodPITInstCMACM As String = "810900"
'Global Const gsCodSegRQClient As String = "05090B0A3F33353F3431"
'Global Const gnTiempoEspTX As Integer = 120 ' En segundos
'Global Const gsRQ_Redirect_IP As String = "192.168.0.21" '"192.168.15.35"
'Global Const gnRQ_Redirect_Port As Integer = 1000

Public gsIPRQ As String
Public gsCodPITInstCMACM As String
Public gsCodSegRQClient As String
Public gnTiempoEspTX As Long
Public gsRQ_Redirect_IP As String
Public gnRQ_Redirect_Port As Long
Public gsIPRQPinVerify As String

Public ArchivoEnvio As String, ArchivoRecepcion As String

Dim sProcessCode As String
Dim sTxnAmount As String
Dim sDateTimeTrans As String
Dim sTerminalId As String
Dim sCardAceptor As String
Dim sCurCode As String
Dim nTramaId As Long
Dim sCOD_ERR_TX As String

Global gsAgeCiudad As String

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


Dim gsMSG_TYPE As String, gsPR_CODE As String, gsTXN_AMOUNT As String, gsDATETIME_TRANS As String
Dim gsAUDIT_NUMBER_TRACE As String, gsTIME_TX_LOCAL As String, gsDATE_TX_LOCAL As String, gsDATE_CAPTURE As String
Dim gsCODE_INST_ACQ As String, gsCODE_INST_FWD As String, gsTRACK2 As String, gsPAN As String, gsNUM_REF_RETRIEVAL As String
Dim gsRESP_IDENT_AUTH As String, gsCODE_RESP As String, gsTERMINAL_ID As String, gsCARD_ACEPTOR As String
Dim gsDATA_RESP_ADD As String, gsCUR_CODE As String, gsDATA_TERMINAL As String, gsDATA_AUTHORIZER As String
Dim gsCODEID_INST_RECEIVING As String, gsIDENT1_ACCOUNT As String, gsIDENT2_ACCOUNT As String
Dim gsDATA_ATM_ADD As String
Dim gsDATA_ELEMENTS_ORIG As String
        
        
                  
Private Declare Function RQxDFClientSend _
Lib "RQxDFClientDLL.dll" _
             (ByVal Name As String, _
              ByVal SecCode As String, _
              ByVal oPer As String, _
              ByVal FileInput As String, _
              ByVal FileOutput As String, _
              ByVal Ip As String, _
              ByVal port As Integer, _
              ByVal timeout As Integer) As Integer

 Private Declare Function GetTokenInfo _
 Lib "RQxDFTk.dll" _
                 (ByVal file As String, _
                  ByVal info As String, _
                  ByVal subinfo As String, _
                  ByVal tokenitem As String _
                 ) As Long
    
    
    
Private Declare Function pinblock _
Lib "PINVerify.dll" _
             (ByVal IPServer As String, _
              ByVal Key As String, _
              ByVal PAN As String, _
              ByVal PVKI As String, _
              ByVal PIN As String, _
              ByVal KeyPINBlock As String _
             ) As Long
    
    
    
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'Para el poder cargar los Datos de la Maquina Cliente
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long


Private Function DevuelveParametro(ByVal x As Long) As String
Dim Buffer() As Byte
Dim nLen As Long
Dim Res As String
Dim i As Integer
    nLen = lstrlenW(x) * 2
    ReDim Buffer(0 To (nLen - 1)) As Byte
    CopyMemory Buffer(0), ByVal x, nLen
    Res = ""
    For i = 0 To nLen - 1
    If (Buffer(i) = 0) Then
    Exit For
    End If
    Res = Res + Chr(Buffer(i))
    Next
    DevuelveParametro = Res
End Function

Private Function getNombrePCUsuario() As String  'Para obtener la Maquina del Usuario
    Dim buffMaq As String
    Dim lSizeMaq As Long
    buffMaq = Space(255)
    lSizeMaq = Len(buffMaq)
    GetComputerName buffMaq, lSizeMaq
    getNombrePCUsuario = Trim(Left$(buffMaq, lSizeMaq))
End Function


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

'Convert a LPWSTR pointer to a VB string
Private Function PtrToString(lpwString As Long) As String
    Dim Buffer() As Byte
    Dim nLen As Long

    If lpwString Then
        nLen = lstrlenW(lpwString) * 2
        If nLen Then
            ReDim Buffer(0 To (nLen - 1)) As Byte
                CopyMemory Buffer(0), ByVal lpwString, nLen
            PtrToString = Buffer
        End If
    End If
End Function



Public Sub GeneraTrama(psPAN As String, psCuenta As String, psOpeCod As String, ByVal pnMonto As Currency, _
        psTRACK2 As String, psCodCMACAuto As String, pnMoneda As Integer, psDNI As String, _
        psCodCMACOrig As String, psPINBlock As String, psCodAge As String, psMovNro As String, _
        Optional psCodTXFinanciera As String = "0200", Optional pnMovNro As Long = 0, Optional ByRef psTramaExtAut As String)
Dim lsTramaXML As String, lsNewPINBlock As String
Dim lsTramaId As String
Dim lsCampo(129) As String '**Campo Primarios (1 al 64) y Secundarios (64 al 128)
Dim COMPITNeg As New COMOpeInterCMAC.dFuncionesNeg
Dim lRsMovAExtornar As ADODB.Recordset
Dim lsTramaXMLEnvio As String, lsTramaXMLRecep As String
Dim x As Long
    
    
    'X = pinblock("192.168.15.35:81", "123", psPAN, "1", "t" & psPINBlock & "_02", "_01")
    x = pinblock(gsIPRQPinVerify, "123", psPAN, "1", "t" & psPINBlock & "_02", "_01")
    lsNewPINBlock = DevuelveParametro(x)
   
    psCuenta = IIf(psCuenta = "", " ", psCuenta)
    
    
    
    ArchivoEnvio = App.path & gsPATHTramas & "_ISO" & Format(Date, "yyyymmdd") & Format(Time, "hhmmss") & "_" & psCodTXFinanciera & ".txt"
    ArchivoRecepcion = App.path & gsPATHTramas & "_ISO" & Format(Date, "yyyymmdd") & Format(Time, "hhmmss") & "_" & IIf(psCodTXFinanciera = "0200", "0210", "0430") & ".txt"


    gsMSG_TYPE = psCodTXFinanciera
    If pnMovNro > 0 Then 'Extorno a solicitud de cliente
        Set lRsMovAExtornar = COMPITNeg.obtenerMovimientoInterCajaParaExtorno(pnMovNro)
        If Not lRsMovAExtornar.EOF And Not lRsMovAExtornar.BOF Then
            lsTramaXMLEnvio = lRsMovAExtornar("cTramaEnvio")
            lsTramaXMLRecep = lRsMovAExtornar("cTramaRecep")
        End If
                
        gsPR_CODE = RecuperaValorXML(lsTramaXMLEnvio, "PR_CODE")
        gsTXN_AMOUNT = RecuperaValorXML(lsTramaXMLEnvio, "TXN_AMOUNT")
        gsDATETIME_TRANS = Format(Date, "MMDD") & Format(Time, "HHMMSS")
        gsAUDIT_NUMBER_TRACE = String(6, "0")
        gsTIME_TX_LOCAL = Format(Time, "HHMMSS")
        gsDATE_TX_LOCAL = Format(Date, "MMDD")
        gsDATE_CAPTURE = Format(Date, "MMDD")
        gsCODE_INST_ACQ = psCodCMACOrig
        gsCODE_INST_FWD = RecuperaValorXML(lsTramaXMLEnvio, "CODE_INST_FWD")
        gsTRACK2 = RecuperaValorXML(lsTramaXMLEnvio, "TRACK2")
        gsPAN = Left(gsTRACK2, 16)
        gsTERMINAL_ID = RecuperaValorXML(lsTramaXMLEnvio, "TERMINAL_ID")
        gsCARD_ACEPTOR = RecuperaValorXML(lsTramaXMLEnvio, "CARD_ACEPTOR")
        gsCUR_CODE = RecuperaValorXML(lsTramaXMLEnvio, "CUR_CODE")
        gsDATA_TERMINAL = RecuperaValorXML(lsTramaXMLRecep, "DATA_TERMINAL")
        gsDATA_AUTHORIZER = RecuperaValorXML(lsTramaXMLRecep, "DATA_AUTHORIZER")
        gsCODEID_INST_RECEIVING = RecuperaValorXML(lsTramaXMLEnvio, "CODEID_INST_RECEIVING")
        gsIDENT1_ACCOUNT = RecuperaValorXML(lsTramaXMLEnvio, "IDENT1_ACCOUNT_DNI")
        gsIDENT2_ACCOUNT = RecuperaValorXML(lsTramaXMLEnvio, "IDENT2_ACCOUNT")
        gsDATA_ATM_ADD = RecuperaValorXML(lsTramaXMLEnvio, "DATA_ATM_ADD")
        
        gsNUM_REF_RETRIEVAL = RecuperaValorXML(lsTramaXMLRecep, "NUM_REF_RETRIEVAL")
        gsRESP_IDENT_AUTH = RecuperaValorXML(lsTramaXMLRecep, "RESP_IDENT_AUTH")
        gsCODE_RESP = "22"
        gsDATA_RESP_ADD = "[.....]"
        
        gsDATA_ELEMENTS_ORIG = RecuperaValorXML(lsTramaXMLEnvio, "MSG_TYPE") & Left(RecuperaValorXML(lsTramaXMLRecep, "NUM_REF_RETRIEVAL") & Space(12), 12) & RecuperaValorXML(lsTramaXMLEnvio, "DATE_TX_LOCAL") & RecuperaValorXML(lsTramaXMLEnvio, "TIME_TX_LOCAL") & "00" & RecuperaValorXML(lsTramaXMLEnvio, "DATE_CAPTURE") & "0000000000"
        
    Else
        
        gsPR_CODE = getPROCESS_CODE(psOpeCod)
        gsTXN_AMOUNT = Right("000000000000" & Replace(Format(pnMonto, "#0.00"), ".", ""), 12)
        gsDATETIME_TRANS = Format(Date, "MMDD") & Format(Time, "HHMMSS")
        gsAUDIT_NUMBER_TRACE = String(6, "0")
        gsTIME_TX_LOCAL = Format(Time, "HHMMSS")
        gsDATE_TX_LOCAL = Format(Date, "MMDD")
        gsDATE_CAPTURE = Format(Date, "MMDD")
        gsCODE_INST_ACQ = psCodCMACOrig
        gsCODE_INST_FWD = psCodCMACAuto
        gsTRACK2 = psTRACK2
        gsPAN = Left(psTRACK2, 16)
        gsTERMINAL_ID = Left(getNombrePCUsuario & String(16, " "), 16)
        gsCARD_ACEPTOR = getCARD_ACEPTOR(psCodAge)
        gsCUR_CODE = IIf(pnMoneda = 0, "000", IIf(pnMoneda = gMonedaNacional, "604", "840"))
        gsDATA_TERMINAL = String(15, "0")
        gsDATA_AUTHORIZER = "MAYNCER1000000P"
        gsCODEID_INST_RECEIVING = Left(psCodCMACAuto & "00000000000", 11) 'String(4, "0")
        gsIDENT1_ACCOUNT = Right(String(28, "0") & IIf(gsPR_CODE = "353500", psDNI, psCuenta), 28) ' DNI o Cuenta
        gsIDENT2_ACCOUNT = Right(String(28, "0") & IIf(gsPR_CODE = "011100", psDNI, ""), 28) 'DNI para retiro que los requiera
        
        If gsCUR_CODE = "840" Then
            'If (gsPR_CODE = "500035" Or gsPR_CODE = "200011") Then
            '    gsDATA_ATM_ADD = "& 0000200080! Q300058 " & gsCUR_CODE & "000" & gsCUR_CODE & "000000000000 " & gsTXN_AMOUNT & "000000000000000000000000"
            'Else
                gsDATA_ATM_ADD = "& 0000200080! Q300058 " & gsCUR_CODE & gsCUR_CODE & "000000000000000 " & gsTXN_AMOUNT & "000000000000000000000000"
                gsTXN_AMOUNT = "000000000000"
            'End If
        Else
            'If (gsPR_CODE = "500035" Or gsPR_CODE = "200011") Then
            '    gsDATA_ATM_ADD = "& 0000200080! Q300058 " & gsCUR_CODE & "000" & gsCUR_CODE & "000000000000 000000000000000000000000000000000000"
                'gsDATA_ATM_ADD = "& 0000200080! Q300058 " & gsCUR_CODE & gsCUR_CODE & "000000000000000 000000000000000000000000000000000000"
            'Else
                gsDATA_ATM_ADD = "& 0000200080! Q300058 " & gsCUR_CODE & gsCUR_CODE & "000000000000000 000000000000000000000000000000000000"
            'End If
        End If
        
        gsNUM_REF_RETRIEVAL = "[.....]"
        gsRESP_IDENT_AUTH = "[.....]"
        gsCODE_RESP = "[.....]"
        gsDATA_RESP_ADD = "[.....]"
        gsDATA_ELEMENTS_ORIG = "[.....]"
        
    End If
    

    lsCampo(1) = "0000000016000004"
    lsCampo(3) = gsPR_CODE
    lsCampo(4) = gsTXN_AMOUNT
    lsCampo(7) = gsDATETIME_TRANS
    lsCampo(11) = gsAUDIT_NUMBER_TRACE
    lsCampo(12) = gsTIME_TX_LOCAL
    lsCampo(13) = gsDATE_TX_LOCAL
    lsCampo(17) = gsDATE_CAPTURE
    lsCampo(32) = gsCODE_INST_ACQ
    lsCampo(33) = gsCODE_INST_FWD
    lsCampo(35) = gsTRACK2
    lsCampo(41) = Left(gsTERMINAL_ID & Space(16), 16)
    lsCampo(43) = gsCARD_ACEPTOR
    lsCampo(49) = gsCUR_CODE
    lsCampo(60) = gsDATA_TERMINAL
    lsCampo(61) = gsDATA_AUTHORIZER
    lsCampo(100) = gsCODEID_INST_RECEIVING
    lsCampo(102) = gsIDENT1_ACCOUNT
    lsCampo(103) = gsIDENT2_ACCOUNT
    lsCampo(126) = gsDATA_ATM_ADD
             

    lsTramaXML = GeneraTramaEnXML(gsMSG_TYPE, gsPR_CODE, gsTXN_AMOUNT, gsDATETIME_TRANS, gsAUDIT_NUMBER_TRACE, gsTIME_TX_LOCAL, gsDATE_TX_LOCAL, gsDATE_CAPTURE, _
                            gsCODE_INST_ACQ, gsCODE_INST_FWD, gsTRACK2, gsNUM_REF_RETRIEVAL, gsRESP_IDENT_AUTH, gsCODE_RESP, gsTERMINAL_ID, gsCARD_ACEPTOR, gsDATA_RESP_ADD, _
                            gsCUR_CODE, gsDATA_TERMINAL, gsDATA_AUTHORIZER, gsDATA_ELEMENTS_ORIG, gsCODEID_INST_RECEIVING, gsIDENT1_ACCOUNT, gsIDENT2_ACCOUNT, gsDATA_ATM_ADD)
             
    'lsTramaXML = GeneraTramaEnXML("0200", sProcessCode, sTxnAmount, sDateTimeTrans, lsCampo(11), lsCampo(12), lsCampo(13), lsCampo(17), _
    '                        psCodCMACOrig, psCodCMACAuto, psPAN, "[.....]", "[.....]", "[.....]", sTerminalId, sCardAceptor, "[.....]", _
    '                        sCurCode, lsCampo(60), lsCampo(61), lsCampo(100), lsCampo(102), lsCampo(103), lsCampo(126))
    psTramaExtAut = lsTramaXML
    
    nTramaId = COMPITNeg.nRegistraTramaEnvio(psMovNro, psOpeCod, lsTramaXML, gsPR_CODE, psCuenta, psPAN, psDNI, gsCUR_CODE, pnMonto, gsCODE_INST_FWD)
    Set COMPITNeg = Nothing
    
    If gsNUM_REF_RETRIEVAL <> "[.....]" Then 'Obtenido para reverso
        lsCampo(37) = gsNUM_REF_RETRIEVAL
        lsCampo(38) = gsRESP_IDENT_AUTH
        lsCampo(39) = gsCODE_RESP
        lsCampo(90) = gsDATA_ELEMENTS_ORIG
    Else
        lsTramaId = Right(String(12, "0") & CStr(nTramaId), 12)
        lsCampo(37) = lsTramaId
        lsCampo(52) = psPINBlock
    End If
    
    
    Call crearArchivoTrama(ArchivoEnvio, lsCampo)
    
End Sub

Public Sub RegistraRespuestaDeTramaEnvio(pnTramaId As Long, pnCodRespSwitch As Integer)
Dim lsTramaXML As String, lsCOD_MSG As String
Dim lsCampo(129) As String
Dim x As Long, lnDenegada As Integer
Dim COMPITNeg As COMOpeInterCMAC.dFuncionesNeg


    x = GetTokenInfo(ArchivoRecepcion, "O", "*", "*")
    lsCOD_MSG = Trim(DevuelveParametro(x))
        

    x = GetTokenInfo(ArchivoRecepcion, "F1", "*", "*")
    lsCampo(1) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F3", "*", "*")
    lsCampo(3) = Trim(DevuelveParametro(x))
    
    sCOD_ERR_TX = "000000"
    If (Left(lsCampo(3), 4) = "9900") Then
        sCOD_ERR_TX = lsCampo(3)
    End If
    
    
    x = GetTokenInfo(ArchivoRecepcion, "F4", "*", "*")
    lsCampo(4) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F7", "*", "*")
    lsCampo(7) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F11", "*", "*")
    lsCampo(11) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F12", "*", "*")
    lsCampo(12) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F13", "*", "*")
    lsCampo(13) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F17", "*", "*")
    lsCampo(17) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F32", "*", "*")
    lsCampo(32) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F33", "*", "*")
    lsCampo(33) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F35", "*", "*")
    lsCampo(35) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F37", "*", "*")
    lsCampo(37) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F38", "*", "*")
    lsCampo(38) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F39", "*", "*")
    lsCampo(39) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F41", "*", "*")
    lsCampo(41) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F43", "*", "*")
    lsCampo(43) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F44", "*", "*")
    lsCampo(44) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F49", "*", "*")
    lsCampo(49) = Trim(DevuelveParametro(x))
   
    x = GetTokenInfo(ArchivoRecepcion, "F60", "*", "*")
    lsCampo(60) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F61", "*", "*")
    lsCampo(61) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F90", "*", "*")
    lsCampo(90) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F100", "*", "*")
    lsCampo(100) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F102", "*", "*")
    lsCampo(102) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F103", "*", "*")
    lsCampo(103) = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoRecepcion, "F126", "*", "*")
    lsCampo(126) = Trim(DevuelveParametro(x))
             
    lsTramaXML = GeneraTramaEnXML(lsCOD_MSG, lsCampo(3), lsCampo(4), lsCampo(7), lsCampo(11), lsCampo(12), lsCampo(13), lsCampo(17), _
                        lsCampo(32), lsCampo(33), lsCampo(35), lsCampo(37), lsCampo(38), lsCampo(39), lsCampo(41), lsCampo(43), _
                        lsCampo(44), lsCampo(49), lsCampo(60), lsCampo(61), lsCampo(90), lsCampo(100), lsCampo(102), lsCampo(103), lsCampo(126))
        
    lnDenegada = IIf(lsCampo(39) = "00", 0, 1)
    
    Set COMPITNeg = New COMOpeInterCMAC.dFuncionesNeg
    Call COMPITNeg.RegistraTramaRecepcion(pnTramaId, lsTramaXML, pnCodRespSwitch, sCOD_ERR_TX, lsCampo(39), lnDenegada)
    Set COMPITNeg = Nothing
       

End Sub


Private Sub crearArchivoTrama(ByVal psArchivo As String, ByVal pMatTrama As Variant)
Dim i As Integer, lsCampo As String
    
    Open psArchivo For Append As #1
  
    For i = 1 To 128 '128 Campos según standar ISO 8583
        If pMatTrama(i) = "" Then
            lsCampo = "[.....]"
        Else
            lsCampo = pMatTrama(i)
        End If
                
        Print #1, lsCampo
    Next
    Close #1
    
End Sub

Sub RegistrarOperacionInterCMAC(psPAN As String, ByVal psPINBlock As String, ByVal psCuenta As String, ByVal psOpeCod As String, psTRACK2 As String, _
            pnMoneda As Integer, psDNI As String, psPersCodCMAC As String, sLpt As String, psOpeDescripcion As String, _
            psNombreCMAC As String, pdFecSis As Date, psCodAge As String, psCodUser As String, _
            Optional ByVal pnMonto As Currency, Optional psGlosa As String, Optional psIFTipo As String, _
            Optional pbImpTMU As Boolean = False, Optional pnMovNroAExtornar As Long = 0, _
            Optional ByVal pnComision As Currency = 0)
Dim lnResp As Integer
Dim lnRespExt As Integer
Dim lsCampo38 As String, lsCampo39 As String, lsCampo44 As String, lsCampo125 As String, lsCampo126 As String
Dim x As Long
Dim lsImpBoleta As String
Dim lsNomTit As String, lsTexto As String, lsDescripcionISO As String
Dim nFicSal As Integer
Dim lsCodAutCMAC As String, lsNomCMACDest As String, lsCodAutDest As String
Dim lnMovNroCom As Long
Dim lsMovNro As String
Dim lsCodTXFinanciera As String, lsCabOpeTX As String, lsTramaExtAut As String

Dim loCOMPITNeg As COMOpeInterCMAC.dFuncionesNeg
    
On Error GoTo Manejador

    Set loCOMPITNeg = New COMOpeInterCMAC.dFuncionesNeg
        
    lsMovNro = loCOMPITNeg.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
    
    Call CargaParametrosRQRedirect
        
    If Left(psOpeCod, 2) = "15" Or Left(psOpeCod, 2) = "27" Then 'Extorno a solicitud
        lsCodTXFinanciera = "0420"
    Else
        lsCodTXFinanciera = "0200"
        lsCodAutCMAC = loCOMPITNeg.DevuelveCodAutorizaCMAC(psPersCodCMAC, lsNomCMACDest)
    End If
    
    Call GeneraTrama(psPAN, psCuenta, psOpeCod, pnMonto, psTRACK2, lsCodAutCMAC, pnMoneda, psDNI, gsCodPITInstCMACM, psPINBlock, psCodAge, lsMovNro, lsCodTXFinanciera, pnMovNroAExtornar, lsTramaExtAut)
    
    lsCabOpeTX = "030000000" & lsCodTXFinanciera
    
    'lnResp = RQxDFClientSend("CAJA MAYNAS", gsCodSegRQClient, lsCabOpeTX, ArchivoEnvio, ArchivoRecepcion, gsRQ_Redirect_IP, lnPortPr, lnTimeOutPr)
    lnResp = RQxDFClientSend("CAJA MAYNAS", gsCodSegRQClient, lsCabOpeTX, ArchivoEnvio, ArchivoRecepcion, gsRQ_Redirect_IP, gnRQ_Redirect_Port, gnTiempoEspTX)
    If lnResp <> 1 Then 'Errores que se detectan al momento de enviar trama
        'Extorno Automatico
        If (psOpeCod = "104001" Or psOpeCod = "261001" Or psOpeCod = "261002") Then
            
            Call GeneraTramaExtornoAut(psPAN, psCuenta, psOpeCod, pnMonto, psTRACK2, lsCodAutCMAC, pnMoneda, psDNI, gsCodPITInstCMACM, psPINBlock, psCodAge, lsMovNro, "0420", pnMovNroAExtornar, lsTramaExtAut, nTramaId)
            Sleep 1000
            lsCabOpeTX = "0300000000420"
    
            lnRespExt = RQxDFClientSend("CAJA MAYNAS", gsCodSegRQClient, lsCabOpeTX, ArchivoEnvio, ArchivoRecepcion, gsRQ_Redirect_IP, gnRQ_Redirect_Port, gnTiempoEspTX)
'
        End If
'
        If lnRespExt < 0 Then
            Call MsgBox(getMensajeErrorRQClient(lnRespExt) & " EA", vbCritical + vbExclamation)
            Exit Sub
        End If
        'Fin Extorno Automatico
        Call MsgBox(getMensajeErrorRQClient(lnResp), vbCritical + vbExclamation)
        Exit Sub
        
    End If
    
    If lnResp = 1 Then 'Hay respuesta
        
        Call RegistraRespuestaDeTramaEnvio(nTramaId, lnResp)
        
        If Left(sCOD_ERR_TX, 4) = "9900" Then 'Existe error en la transacción
            Call MsgBox(getMensajeErrorEnTransaccion(sCOD_ERR_TX), vbCritical + vbExclamation)
            Exit Sub
        End If
    
        'Comenzar a leer campos del archivo de salida
        x = GetTokenInfo(ArchivoRecepcion, "F38", "*", "*")
        lsCampo38 = Trim(DevuelveParametro(x))
        x = GetTokenInfo(ArchivoRecepcion, "F39", "*", "*")
        lsCampo39 = Trim(DevuelveParametro(x))
        x = GetTokenInfo(ArchivoRecepcion, "F44", "*", "*")
        lsCampo44 = Trim(DevuelveParametro(x))
        x = GetTokenInfo(ArchivoRecepcion, "F125", "*", "*")
        lsCampo125 = Trim(DevuelveParametro(x))
        x = GetTokenInfo(ArchivoRecepcion, "F126", "*", "*")
        lsCampo126 = Trim(DevuelveParametro(x))

               
        If lsCampo39 <> "00" Then
            lsDescripcionISO = loCOMPITNeg.obtenerDescripcionPorCodigoISO(lsCampo39)
            If lsCampo39 = "53" And Left(psOpeCod, 2) = "10" Then
                lsDescripcionISO = lsDescripcionISO & " Credito"
            Else
                lsDescripcionISO = lsDescripcionISO & " Ahorros"
            End If
            Call MsgBox("Transacción denegada, código de respuesta " & lsCampo39 & ": " & lsDescripcionISO & " Detalle : " & Trim(lsCampo125), vbInformation)
            Exit Sub
        End If
        
        lsCodAutDest = lsCampo38
        lsTexto = lsCampo125

        
        If lsCampo39 = "00" Then
            Select Case psOpeCod
                Case "104001" 'Pago Credito
                    lsNomTit = "(INTERCAJAS) PAGO DE CREDITOS"
                    Call loCOMPITNeg.nRegistraOperacionInterCMACM_Envio(lsMovNro, psOpeCod, psCuenta, psOpeDescripcion, psPersCodCMAC, pnMoneda, "", pnMonto, "", psNombreCMAC, gsNomAge, pnComision, psGlosa)
                    'MsgBox "La operación se registró sstisfactoriamente"
                Case "104002" 'Consulta de Cuentas de crédito
                    lsNomTit = "(INTERCAJAS) CONSULTA DE CREDITOS"
                    Call loCOMPITNeg.nRegistraOperacionInterCMACM_Envio(lsMovNro, psOpeCod, psCuenta, psOpeDescripcion, psPersCodCMAC, pnMoneda, "", pnMonto, "", psNombreCMAC, gsNomAge, pnComision, psGlosa)
                    'MsgBox "La operación se registró sstisfactoriamente"
                Case "104003" 'Consulta de movimientos de crédito
                    lsNomTit = "(INTERCAJAS) CONSULTA MOV. CRÉDITO"
                    Call loCOMPITNeg.nRegistraOperacionInterCMACM_Envio(lsMovNro, psOpeCod, psCuenta, psOpeDescripcion, psPersCodCMAC, pnMoneda, "", pnMonto, "", psNombreCMAC, gsNomAge, pnComision, psGlosa)
                    'MsgBox "La operación se registró sstisfactoriamente"
                Case "261001" 'Retiro
                    lsNomTit = "(INTERCAJAS) RETIRO EFECTIVO"
                    Call loCOMPITNeg.nRegistraOperacionInterCMACM_Envio(lsMovNro, psOpeCod, psCuenta, psOpeDescripcion, psPersCodCMAC, pnMoneda, "", pnMonto, "", psNombreCMAC, gsNomAge, pnComision, psGlosa)
                    If pnMoneda = 1 Then
                        lsTexto = lsCampo44
                    Else
                        lsTexto = "3+" & Mid(lsCampo126, 58, 11) & "+" & Mid(lsCampo126, 70, 11)
                    End If
                     
                    'MsgBox "La operación se registró sstisfactoriamente"
                Case "261002" 'Deposito
                    lsNomTit = "(INTERCAJAS) DEPOSITO EFECTIVO"
                    Call loCOMPITNeg.nRegistraOperacionInterCMACM_Envio(lsMovNro, psOpeCod, psCuenta, psOpeDescripcion, psPersCodCMAC, pnMoneda, "", pnMonto, "", psNombreCMAC, gsNomAge, pnComision, psGlosa)
                    'MsgBox "La operación se registró sstisfactoriamente"
                Case "261003" 'Consulta de Cuentas de Ahorro
                    lsNomTit = "(INTERCAJAS) CONSULTA CUENTAS AHORRO"
                    Call loCOMPITNeg.nRegistraOperacionInterCMACM_Envio(lsMovNro, psOpeCod, psCuenta, psOpeDescripcion, psPersCodCMAC, pnMoneda, "", pnMonto, "", psNombreCMAC, gsNomAge, pnComision, psGlosa)
                    'MsgBox "La operación se registró sstisfactoriamente"
                Case "261004" 'Consulta de movimientos de Cuentas de Ahorro
                    lsNomTit = "(INTERCAJAS) CONSULTA DE MOVIMIENTOS"
                    Call loCOMPITNeg.nRegistraOperacionInterCMACM_Envio(lsMovNro, psOpeCod, psCuenta, psOpeDescripcion, psPersCodCMAC, pnMoneda, "", pnMonto, "", psNombreCMAC, gsNomAge, pnComision, psGlosa)
                    'MsgBox "La operación se registró sstisfactoriamente"
                Case "159101" 'Extorno pago de crédito
                    lsNomTit = "(INTERCAJAS) EXTORNO DE PAGO DE CREDITO"
                    Call loCOMPITNeg.nRegistraExtornoOperacionInterCMACM_Envio(pnMovNroAExtornar, lsMovNro, psOpeCod, psCuenta, psOpeDescripcion, psPersCodCMAC, pnMoneda, "", pnMonto, "", psNombreCMAC, gsNomAge, 0, psGlosa)
                    'MsgBox "La operación se registró sstisfactoriamente"
                Case "279101" 'Extorno retiro
                    lsNomTit = "(INTERCAJAS) EXTORNO DE RETIRO"
                    Call loCOMPITNeg.nRegistraExtornoOperacionInterCMACM_Envio(pnMovNroAExtornar, lsMovNro, psOpeCod, psCuenta, psOpeDescripcion, psPersCodCMAC, pnMoneda, "", pnMonto, "", psNombreCMAC, gsNomAge, 0, psGlosa)
                    'MsgBox "La operación se registró sstisfactoriamente"
                Case "279102" 'Extorno deposito
                    lsNomTit = "(INTERCAJAS) EXTORNO DE DEPOSITO"
                    Call loCOMPITNeg.nRegistraExtornoOperacionInterCMACM_Envio(pnMovNroAExtornar, lsMovNro, psOpeCod, psCuenta, psOpeDescripcion, psPersCodCMAC, pnMoneda, "", pnMonto, "", psNombreCMAC, gsNomAge, 0, psGlosa)
                    'MsgBox "La operación se registró sstisfactoriamente"
            End Select
        End If
            
        lsImpBoleta = lsImpBoleta & loCOMPITNeg.ImprimeBoletaOpeInterCMAC(lsNomTit, lsTexto, CStr(gnMovNro), CStr(lnMovNroCom), psCuenta, gdFecSis, gsNomAge, lsNomCMACDest, psPAN, lsCodAutDest, psCodUser, psOpeCod, psDNI, pnMoneda, pnMonto, pbImpTMU)
           
        'lsImpBoleta = lsImpBoleta & Chr(10)

        Do
            If Trim(lsImpBoleta) <> "" Then
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, lsImpBoleta
                    Print #nFicSal, ""
                Close #nFicSal
            End If
        Loop Until MsgBox("¿Desea Re-Imprimir Boletas ?", vbQuestion + vbYesNo, "Aviso") = vbNo
    
    End If
    Set loCOMPITNeg = Nothing
    
    Exit Sub
Manejador:
    MsgBox "Ocurrió un error inesperado", vbCritical + vbExclamation
End Sub

Public Function GeneraTramaEnXML(psMSG_TYPE As String, psPR_CODE As String, psTXN_AMOUNT As String, psDATETIME_TRANS As String, _
        psAUDIT_NUMBER_TRACE As String, psTIME_TX_LOCAL As String, psDATE_TX_LOCAL As String, psDATE_CAPTURE As String, _
        psCODE_INST_ACQ As String, psCODE_INST_FWD As String, psTRACK2 As String, psNUM_REF_RETRIEVAL As String, _
        psRESP_IDENT_AUTH As String, psCODE_RESP As String, psTERMINAL_ID As String, psCARD_ACEPTOR As String, _
        psDATA_RESP_ADD As String, psCUR_CODE, psDATA_TERMINAL As String, psDATA_AUTHORIZER As String, _
        psDATA_ELEMENTS_ORIG As String, psCODEID_INST_RECEIVING As String, psIDENT1_ACCOUNT As String, _
        psIDENT2_ACCOUNT As String, psDATA_ATM_ADD As String) As String
Dim sXMLTrama As String
    sXMLTrama = "<MSG_TYPE = " & psMSG_TYPE & " />"
    sXMLTrama = sXMLTrama & " <PR_CODE = " & psPR_CODE & " />"
    sXMLTrama = sXMLTrama & " <TXN_AMOUNT = " & psTXN_AMOUNT & " />"
    sXMLTrama = sXMLTrama & " <DATETIME_TRANS = " & psDATETIME_TRANS & " />"
    sXMLTrama = sXMLTrama & " <AUDIT_NUMBER_TRACE = " & psAUDIT_NUMBER_TRACE & " />"
    sXMLTrama = sXMLTrama & " <TIME_TX_LOCAL = " & psTIME_TX_LOCAL & " />"
    sXMLTrama = sXMLTrama & " <DATE_TX_LOCAL = " & psDATE_TX_LOCAL & " />"
    sXMLTrama = sXMLTrama & " <DATE_CAPTURE = " & psDATE_CAPTURE & " />"
    sXMLTrama = sXMLTrama & " <CODE_INST_ACQ = " & psCODE_INST_ACQ & " />"
    sXMLTrama = sXMLTrama & " <CODE_INST_FWD = " & psCODE_INST_FWD & " />"
    sXMLTrama = sXMLTrama & " <TRACK2 = " & psTRACK2 & " />"
    sXMLTrama = sXMLTrama & " <NUM_REF_RETRIEVAL = " & psNUM_REF_RETRIEVAL & " />"
    sXMLTrama = sXMLTrama & " <RESP_IDENT_AUTH = " & psRESP_IDENT_AUTH & " />"
    sXMLTrama = sXMLTrama & " <CODE_RESP = " & psCODE_RESP & " />"
    sXMLTrama = sXMLTrama & " <TERMINAL_ID = " & psTERMINAL_ID & " />"
    sXMLTrama = sXMLTrama & " <CARD_ACEPTOR = " & psCARD_ACEPTOR & " />"
    sXMLTrama = sXMLTrama & " <DATA_RESP_ADD = " & psDATA_RESP_ADD & " />"
    sXMLTrama = sXMLTrama & " <CUR_CODE = " & psCUR_CODE & " />"
    sXMLTrama = sXMLTrama & " <DATA_TERMINAL = " & psDATA_TERMINAL & " />"
    sXMLTrama = sXMLTrama & " <DATA_AUTHORIZER = " & psDATA_AUTHORIZER & " />"
    sXMLTrama = sXMLTrama & " <DATA_ELEMENTS_ORIG = " & psDATA_ELEMENTS_ORIG & " />"
    sXMLTrama = sXMLTrama & " <CODEID_INST_RECEIVING = " & psCODEID_INST_RECEIVING & " />"
    sXMLTrama = sXMLTrama & " <IDENT1_ACCOUNT_DNI = " & psIDENT1_ACCOUNT & " />"
    sXMLTrama = sXMLTrama & " <IDENT2_ACCOUNT = " & psIDENT2_ACCOUNT & " />"
    sXMLTrama = sXMLTrama & " <DATA_ATM_ADD = " & psDATA_ATM_ADD & " />"
    
    GeneraTramaEnXML = sXMLTrama
End Function

Public Function getCARD_ACEPTOR(psCodAge As String) As String
    getCARD_ACEPTOR = Left("CMAC MAYNAS S.A.    " & Space(22), 22) & Left(gsAgeCiudad & Space(13), 13) & "MAYPE"
End Function

Public Function getPROCESS_CODE(psOpeCod As String) As String
    Select Case psOpeCod
        Case "104001"
            getPROCESS_CODE = "500035"
        Case "104002"
            getPROCESS_CODE = "353500"
        Case "104003"
            getPROCESS_CODE = "943500"
        Case "261001"
            getPROCESS_CODE = "011100"
        Case "261002"
            getPROCESS_CODE = "200011"
        Case "261003"
            getPROCESS_CODE = "351100"
        Case "261004"
            getPROCESS_CODE = "941100"
        Case "261005"
            getPROCESS_CODE = "311100"
            
        Case "159101"
            getPROCESS_CODE = "500035"
        Case "279101"
            getPROCESS_CODE = "011100"
        Case "279102"
            getPROCESS_CODE = "200011"
    End Select
End Function

Private Function getMensajeOpeInvalida(psCodRespuesta As String) As String

    Select Case psCodRespuesta
        Case "06"
            getMensajeOpeInvalida = " (Err 06) Error"
        Case "12"
            getMensajeOpeInvalida = " (Err 12) No existen datos para la tranacción"
        Case "14"
            getMensajeOpeInvalida = " (Err 14) Número de tarjeta no existe"
        Case "47"
            getMensajeOpeInvalida = " (Err 47) Cuenta para retención judicial"
        Case "50"
            getMensajeOpeInvalida = " (Err 50) Pago excede deuda"
        Case "51"
            getMensajeOpeInvalida = " (Err 51) Cuenta no posee fondos suficientes"
        Case "54"
            getMensajeOpeInvalida = " (Err 54) Tarjeta caducada"
        Case "55"
            getMensajeOpeInvalida = " (Err 55) PIN Inválido"
        Case "62"
            getMensajeOpeInvalida = " (Err 62) Tarjeta inválida"
        Case "65"
            getMensajeOpeInvalida = " (Err 65) Excede los límites mínimos o  máximos para realizar eta operación"
        Case "78"
            getMensajeOpeInvalida = " (Err 78) Cuenta no existe"
        Case "79"
            getMensajeOpeInvalida = " (Err 79) Cuenta no pertenece al titular"
        Case "80"
            getMensajeOpeInvalida = " (Err 80) Cuenta no admite retiros"
        Case "81"
            getMensajeOpeInvalida = " (Err 81) Cuenta bloqueada"
        Case "82"
            getMensajeOpeInvalida = " (Err 82) Cuenta se encuentra cancelado"
        Case Else
            getMensajeOpeInvalida = " (Err " & psCodRespuesta & ") Código de respuesta desconocido"
    End Select
End Function

Private Function getMensajeErrorEnTransaccion(psCodErrTX As String) As String
    Select Case Right(psCodErrTX, 2)
        Case "01"
            getMensajeErrorEnTransaccion = "Formato ISO incorrecto"
        Case "02"
            getMensajeErrorEnTransaccion = "Longitud ISO incorrecto"
        Case "03"
            getMensajeErrorEnTransaccion = "Targer Error"
        Case "04"
            getMensajeErrorEnTransaccion = "Tranacción denegada"
        Case "05"
            getMensajeErrorEnTransaccion = "No Target"
        Case "06"
            getMensajeErrorEnTransaccion = "Error: Bad Program Response"
        Case Else
            getMensajeErrorEnTransaccion = "Error de transacción desconocido"
    End Select
End Function

Private Function getMensajeErrorRQClient(pnRespRQClient As Integer) As String
    Select Case pnRespRQClient
        Case 1
            getMensajeErrorRQClient = "Congrats!!! recibio respuesta"
        Case -1
            getMensajeErrorRQClient = "No hay respuesta del receptor"
        Case -2
            getMensajeErrorRQClient = "No existe archivo de entrada"
        Case -3
            getMensajeErrorRQClient = "Error en codigo de seguridad"
    End Select
End Function

'Funcion que llena un Combo con un recordset
Sub PIT_Llenar_Combo_con_Recordset(pRs As ADODB.Recordset, pcboObjeto As ComboBox)

    pcboObjeto.Clear
    Do While Not pRs.EOF
        pcboObjeto.AddItem Trim(pRs!cCMACDesc) & Space(70) & Trim(pRs!cCMACPersCod)
        pRs.MoveNext
    Loop
    pRs.Close
        
End Sub

Public Function getTarjetaFormateado(psTarjeta As String) As String
    getTarjetaFormateado = Left(psTarjeta, 4) & "-" & Mid(psTarjeta, 5, 4) & "-" & Mid(psTarjeta, 9, 4) & "-" & Right(psTarjeta, 4)
End Function
Public Function CargaParametrosRQRedirect()
    gsIPRQPinVerify = Trim(LeeConstanteSist(403)) 'IP y PUERTO PinVerify
    gsRQ_Redirect_IP = Trim(LeeConstanteSist(404)) 'IP RQ Redirect
    gnRQ_Redirect_Port = CLng(Trim(LeeConstanteSist(405))) 'Puerto RQ Redirect
    gnTiempoEspTX = CLng(Trim(LeeConstanteSist(406))) 'Tiempo Espera Transaccion RQ Redirect
    gsCodSegRQClient = Trim(LeeConstanteSist(407)) 'Codigo Seguridad RQ Client
    gsCodPITInstCMACM = Trim(LeeConstanteSist(408)) 'Codigo de CMAC en el PIT
    gsIPRQ = Trim(LeeConstanteSist(409)) 'IP y Puerto RQ Intercajas
End Function

Public Function RecuperaIpPuertoPinVerifyPOS() As String
    RecuperaIpPuertoPinVerifyPOS = Trim(LeeConstanteSist(403))
End Function

Public Sub GeneraTramaExtornoAut(psPAN As String, psCuenta As String, psOpeCod As String, ByVal pnMonto As Currency, _
        psTRACK2 As String, psCodCMACAuto As String, pnMoneda As Integer, psDNI As String, _
        psCodCMACOrig As String, psPINBlock As String, psCodAge As String, psMovNro As String, _
        psCodTXFinanciera As String, pnMovNro As Long, psTramaExtAuto As String, pnTramaId As Long)
        
Dim lsTramaXML As String ' lsNewPINBlock As String
Dim lsTramaIdExt As String
Dim lsCampo(129) As String '**Campo Primarios (1 al 64) y Secundarios (64 al 128)
Dim COMPITNeg As New COMOpeInterCMAC.dFuncionesNeg
Dim lRsMovAExtornar As ADODB.Recordset
Dim lsTramaXMLEnvio As String, lsTramaXMLRecep As String
Dim x As Long
    
    
    'X = pinblock("192.168.15.35:81", "123", psPAN, "1", "t" & psPINBlock & "_02", "_01")
    'x = pinblock(gsIPRQPinVerify, "123", psPAN, "1", "t" & psPINBlock & "_02", "_01")
    'lsNewPINBlock = DevuelveParametro(x)
   
    psCuenta = IIf(psCuenta = "", " ", psCuenta)
    
    'Set lrsMov = loCOMPITNeg.obtenerMovimientosInterCajasParaExtorno(2, lsDatoBusq, CStr(nOperacion), lsFecha)
    
    ArchivoEnvio = App.path & gsPATHTramas & "_ISO" & Format(Date, "yyyymmdd") & Format(Time, "hhmmss") & "_" & psCodTXFinanciera & ".txt"
    ArchivoRecepcion = App.path & gsPATHTramas & "_ISO" & Format(Date, "yyyymmdd") & Format(Time, "hhmmss") & "_0430" & ".txt"


    gsMSG_TYPE = psCodTXFinanciera
    'If pnMovNro > 0 Then 'Extorno a solicitud de cliente
        'Set lRsMovAExtornar = COMPITNeg.obtenerMovimientoInterCajaParaExtorno(pnMovNro)
        'If Not lRsMovAExtornar.EOF And Not lRsMovAExtornar.BOF Then
        '    lsTramaXMLEnvio = lRsMovAExtornar("cTramaEnvio")
        '    lsTramaXMLRecep = lRsMovAExtornar("cTramaRecep")
        'End If
        
        lsTramaXMLEnvio = psTramaExtAuto
        'lsTramaXMLRecep = ArchivoRecepcion
        
        gsPR_CODE = RecuperaValorXML(lsTramaXMLEnvio, "PR_CODE")
        gsTXN_AMOUNT = RecuperaValorXML(lsTramaXMLEnvio, "TXN_AMOUNT")
        gsDATETIME_TRANS = Format(Date, "MMDD") & Format(Time, "HHMMSS")
        gsAUDIT_NUMBER_TRACE = String(6, "0")
        gsTIME_TX_LOCAL = Format(Time, "HHMMSS")
        gsDATE_TX_LOCAL = Format(Date, "MMDD")
        gsDATE_CAPTURE = Format(Date, "MMDD")
        gsCODE_INST_ACQ = psCodCMACOrig
        gsCODE_INST_FWD = RecuperaValorXML(lsTramaXMLEnvio, "CODE_INST_FWD")
        gsTRACK2 = RecuperaValorXML(lsTramaXMLEnvio, "TRACK2")
        gsPAN = Left(gsTRACK2, 16)
        gsTERMINAL_ID = RecuperaValorXML(lsTramaXMLEnvio, "TERMINAL_ID")
        gsCARD_ACEPTOR = RecuperaValorXML(lsTramaXMLEnvio, "CARD_ACEPTOR")
        gsCUR_CODE = RecuperaValorXML(lsTramaXMLEnvio, "CUR_CODE")
        gsDATA_TERMINAL = RecuperaValorXML(lsTramaXMLEnvio, "DATA_TERMINAL")
        gsDATA_AUTHORIZER = RecuperaValorXML(lsTramaXMLEnvio, "DATA_AUTHORIZER")
        gsCODEID_INST_RECEIVING = RecuperaValorXML(lsTramaXMLEnvio, "CODEID_INST_RECEIVING")
        gsIDENT1_ACCOUNT = RecuperaValorXML(lsTramaXMLEnvio, "IDENT1_ACCOUNT_DNI")
        gsIDENT2_ACCOUNT = RecuperaValorXML(lsTramaXMLEnvio, "IDENT2_ACCOUNT")
        gsDATA_ATM_ADD = RecuperaValorXML(lsTramaXMLEnvio, "DATA_ATM_ADD")
        
        lsTramaIdExt = Right(String(12, "0") & CStr(nTramaId), 12)
        gsNUM_REF_RETRIEVAL = lsTramaIdExt
        'gsNUM_REF_RETRIEVAL = RecuperaValorXML(lsTramaXMLEnvio, "NUM_REF_RETRIEVAL")
        
        gsRESP_IDENT_AUTH = RecuperaValorXML(lsTramaXMLEnvio, "RESP_IDENT_AUTH")
        gsCODE_RESP = "68"
        gsDATA_RESP_ADD = "[.....]"
        
        gsDATA_ELEMENTS_ORIG = RecuperaValorXML(lsTramaXMLEnvio, "MSG_TYPE") & Right(String(12, "0") & CStr(nTramaId), 12) & RecuperaValorXML(lsTramaXMLEnvio, "DATE_TX_LOCAL") & RecuperaValorXML(lsTramaXMLEnvio, "TIME_TX_LOCAL") & "00" & RecuperaValorXML(lsTramaXMLEnvio, "DATE_CAPTURE") & "0000000000"
        
    'End If
    

    lsCampo(1) = "0000000016000004"
    lsCampo(3) = gsPR_CODE
    lsCampo(4) = gsTXN_AMOUNT
    lsCampo(7) = gsDATETIME_TRANS
    lsCampo(11) = gsAUDIT_NUMBER_TRACE
    lsCampo(12) = gsTIME_TX_LOCAL
    lsCampo(13) = gsDATE_TX_LOCAL
    lsCampo(17) = gsDATE_CAPTURE
    lsCampo(32) = gsCODE_INST_ACQ
    lsCampo(33) = gsCODE_INST_FWD
    lsCampo(35) = gsTRACK2
    lsCampo(41) = Left(gsTERMINAL_ID & Space(16), 16)
    lsCampo(43) = gsCARD_ACEPTOR
    lsCampo(49) = gsCUR_CODE
    lsCampo(60) = gsDATA_TERMINAL
    lsCampo(61) = gsDATA_AUTHORIZER
    lsCampo(100) = gsCODEID_INST_RECEIVING
    lsCampo(102) = gsIDENT1_ACCOUNT
    lsCampo(103) = gsIDENT2_ACCOUNT
    lsCampo(126) = gsDATA_ATM_ADD
             

    lsTramaXML = GeneraTramaEnXML(gsMSG_TYPE, gsPR_CODE, gsTXN_AMOUNT, gsDATETIME_TRANS, gsAUDIT_NUMBER_TRACE, gsTIME_TX_LOCAL, gsDATE_TX_LOCAL, gsDATE_CAPTURE, _
                            gsCODE_INST_ACQ, gsCODE_INST_FWD, gsTRACK2, gsNUM_REF_RETRIEVAL, gsRESP_IDENT_AUTH, gsCODE_RESP, gsTERMINAL_ID, gsCARD_ACEPTOR, gsDATA_RESP_ADD, _
                            gsCUR_CODE, gsDATA_TERMINAL, gsDATA_AUTHORIZER, gsDATA_ELEMENTS_ORIG, gsCODEID_INST_RECEIVING, gsIDENT1_ACCOUNT, gsIDENT2_ACCOUNT, gsDATA_ATM_ADD)
             
    'lsTramaXML = GeneraTramaEnXML("0200", sProcessCode, sTxnAmount, sDateTimeTrans, lsCampo(11), lsCampo(12), lsCampo(13), lsCampo(17), _
    '                        psCodCMACOrig, psCodCMACAuto, psPAN, "[.....]", "[.....]", "[.....]", sTerminalId, sCardAceptor, "[.....]", _
    '                        sCurCode, lsCampo(60), lsCampo(61), lsCampo(100), lsCampo(102), lsCampo(103), lsCampo(126))
    
    
    nTramaId = COMPITNeg.nRegistraTramaEnvio(psMovNro, psOpeCod, lsTramaXML, gsPR_CODE, psCuenta, psPAN, psDNI, gsCUR_CODE, pnMonto, gsCODE_INST_FWD)
    
    Set COMPITNeg = Nothing
    
    'If gsNUM_REF_RETRIEVAL <> "[.....]" Then 'Obtenido para reverso
        lsCampo(37) = gsNUM_REF_RETRIEVAL
        lsCampo(38) = gsRESP_IDENT_AUTH
        lsCampo(39) = gsCODE_RESP
        lsCampo(90) = gsDATA_ELEMENTS_ORIG
    'Else
'        lsTramaIdExt = Right(String(12, "0") & CStr(nTramaId), 12)
'        lsCampo(37) = lsTramaIdExt
'        lsCampo(52) = psPINBlock
    'End If
    
    
    Call crearArchivoTrama(ArchivoEnvio, lsCampo)
    
End Sub

Public Function RecuperaMontoOpeXdia(ByVal psCuenta As String, ByVal pnMoneda As Integer, ByRef pnNumOpe As Integer) As Double
Dim COMPITNeg As New COMOpeInterCMAC.dFuncionesNeg
    RecuperaMontoOpeXdia = COMPITNeg.ObtenerMontoOpeXDia(psCuenta, pnMoneda, gdFecSis, pnNumOpe)
    Set COMPITNeg = Nothing
End Function

Public Function RecuperaMontoOpeXMes(ByVal psCuenta As String, ByVal pnMoneda As Integer, ByRef pnNumOpe As Integer) As Double
Dim COMPITNeg As New COMOpeInterCMAC.dFuncionesNeg
    RecuperaMontoOpeXMes = COMPITNeg.ObtenerMontoOpeXMes(psCuenta, pnMoneda, gdFecSis, pnNumOpe)
    Set COMPITNeg = Nothing
End Function

