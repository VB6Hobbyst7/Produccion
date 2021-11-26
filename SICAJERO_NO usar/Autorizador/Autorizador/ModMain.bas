Attribute VB_Name = "ModMain"
Option Explicit

Global Const gsIPRQ As String = "192.168.0.9:81"

Dim sTramaXML As String

Public MESSAGE_TYPE As String
Public TRACE As String
Public PRCODE As String
Public PAN As String
Public TIME_LOCAL As String
Public DATE_LOCAL As String
Public TERMINAL_ID As String
Public ACCT_1 As String
Public CARD_ACCEPTOR As String
Public ACQ_INST As String
Public POS_COND_CODE As String
Public TXN_AMOUNT As String
Public CUR_CODE As String
Public ACCT_2 As String
Public DATE_EXP As String
Public CARD_LOCATION As String
 
 
 Private Declare Function GetTokenInfo _
    Lib "RQxDFTk.dll" _
                 (ByVal file As String, _
                  ByVal info As String, _
                  ByVal subinfo As String, _
                  ByVal tokenitem As String _
                 ) As Long
                 
    Private Declare Function pinverify _
    Lib "PINVerify.dll" _
                 (ByVal ippuerto As String, _
                  ByVal key As String, _
                  ByVal PAN As String, _
                  ByVal pvki As String, _
                  ByVal pin As String, _
                  ByVal PVV As String _
                 ) As Integer
    
    Private Declare Function changepin _
    Lib "PINVerify.dll" _
                 (ByVal ippuerto As String, _
                  ByVal key As String, _
                  ByVal PAN As String, _
                  ByVal pvki As String, _
                  ByVal pin As String, _
                  ByVal PVV As String, _
                  ByVal npin As String _
                 ) As Long
    
    Private Declare Function genpckey Lib "PINVerify.dll" () As Long
    
    Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long
    
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)


Public Sub ArmaTramaXML(ByRef psTramaXML As String, ByVal pnCodCampo As Integer, ByVal psValor As String)

sTramaXML = " [<?xml version=""1.0""?> <Messages> <TXN_FIN_REQ> <MESSAGE_TYPE value=""0200""/>  <PAN "
sTramaXML = sTramaXML & " value=""4261540000000127""/>  <PRCODE value=""011000""/>  <TXN_AMOUNT "
sTramaXML = sTramaXML & " value=""000000010000""/>  <CARDISS_AMOUNT     value=""null""/>  <TXN_DATE_TIME "
sTramaXML = sTramaXML & " value=""0717104456""/>  <CONVRATE           value=""4261540000000159""/>  <TRACE "
sTramaXML = sTramaXML & " value=""370846""/>  <TIME_LOCAL value=""104456""/>  <DATE_LOCAL value=""0717""/> "
sTramaXML = sTramaXML & " <DATE_EXP           value=""1111""/>  <DATE_STTL          value=""0717""/> "
sTramaXML = sTramaXML & " <DATE_CAPTURE       value=""0717""/>  <MERCHANT           value=""0000""/> "
sTramaXML = sTramaXML & " <COUNTRY_CODE       value=""null""/>  <POS_ENTRY_MODE     value=""000""/> "
sTramaXML = sTramaXML & " <POS_COND_CODE      value=""02""/>  <ACQ_INST value=""426154""/>  <ISS_INST "
sTramaXML = sTramaXML & " value=""426154""/>  <PAN_EXT  value=""null""/>  <TRACK2 "
sTramaXML = sTramaXML & " value=""4261540000000159=11111201705242600000""/>  <REFNUM "
sTramaXML = sTramaXML & " value=""000000000000""/>  <AUTH_CODE          value=""null""/>  <RESP_CODE"
sTramaXML = sTramaXML & " value=""null""/>  <TERMINAL_ID value=""00000370""/>  <CARD_ACCEPTOR"
sTramaXML = sTramaXML & " value=""000000000000000""/>  <CARD_LOCATION      value=""  RED UNICARD       CAJERO"
sTramaXML = sTramaXML & " TEST CUZCO   ""/>  <ADD_RESP_DATA value=""""/>  <CUR_CODE           value=""604""/>"
sTramaXML = sTramaXML & " <CUR_CODE_CARDISS   value=""null""/>  <PIN_BLOCK          value=""E2ADA72FBC4EB9FD""/>"
sTramaXML = sTramaXML & " <ADD_POS_INFO       value=""null""/>  <NET_INF            value=""null""/>  <ORG_DATA "
sTramaXML = sTramaXML & "        value=""null""/>  <REP_AMOUNT         value=""null""/>  <REQ_INST"
sTramaXML = sTramaXML & "  value=""426154""/>  <ACCT_1             value=""2321000000019 26154    001""/>"
sTramaXML = sTramaXML & "  <ACCT_2             value=""""/>  <CUSTOMER_INF_RESP  value=""null""/>  <PRIV_USE"
sTramaXML = sTramaXML & "    value=""null""/>  </TXN_FIN_REQ> </Messages>]"


End Sub

Public Function CambioClave(ByVal psAntg As String, ByVal psNuevo As String) As String

    CambioClave = "1234"

End Function

Sub Main()
Dim i As Integer
Dim x As Long
Dim A As Autorizador.ClsAutorizador
Dim sTramaSalida As String
Dim sTramaXML As String
Dim bPinValido As Boolean
Dim psPVVNew As String

Dim lnMovNro As Long

'**DAOR 20081016 ***************************************************************
'**Campo Primarios
Dim sCampoP001, sCampoP003, sCampoP004, sCampoP006, sCampoP007 As String
Dim sCampoP011, sCampoP012, sCampoP013, sCampoP017, sCampoP032 As String
Dim sCampoP035, sCampoP037, sCampoP038, sCampoP039, sCampoP041, sCampoP042, sCampoP043 As String
Dim sCampoP044, sCampoP049, sCampoP060, sCampoP061, sCampoP052 As String

'**Campo secundarios
Dim sCampoS100, sCampoS102, sCampoS103, sCampoS125, sCampoS126 As String

'**
Dim sSaldoD, sSaldoC, sMonRetCta, psMonCta As String
Dim bRetenerTarjeta As Boolean
Dim nPos, nLen, campo As Integer
Dim sPinNew, sPVVNew, sCampo, PVV, fcampo As String
Dim Argumento, arrParametros, ArchivoEntrada, ArchivoSalida, ArchivoSalida2 As Variant

'*********************************************************************************
    Argumento = Command()
    arrParametros = Split(Argumento, " ")
    ArchivoEntrada = CStr(Trim(Replace(arrParametros(0), """", "")))
    ArchivoSalida = CStr(Trim(Replace(arrParametros(1), """", "")))
    ArchivoSalida2 = Replace(ArchivoSalida, ".txt", "_newcode.txt")

    Set A = New Autorizador.ClsAutorizador
        

    x = GetTokenInfo(ArchivoEntrada, "O", "*", "*")
    MESSAGE_TYPE = Trim(DevuelveParametro(x))

    x = GetTokenInfo(ArchivoEntrada, "F37", "*", "*")
    TRACE = Trim(DevuelveParametro(x))

    x = GetTokenInfo(ArchivoEntrada, "F3", "*", "*")
    PRCODE = Trim(DevuelveParametro(x))

    x = GetTokenInfo(ArchivoEntrada, "F35", "*", "*")
    sCampo = Trim(DevuelveParametro(x))
    PAN = Mid(sCampo, 1, 16)

    x = GetTokenInfo(ArchivoEntrada, "F12", "*", "*")
    TIME_LOCAL = Trim(DevuelveParametro(x))

    x = GetTokenInfo(ArchivoEntrada, "F13", "*", "*")
    DATE_LOCAL = Trim(DevuelveParametro(x))

    x = GetTokenInfo(ArchivoEntrada, "F41", "*", "*")
    TERMINAL_ID = Trim(DevuelveParametro(x))

    ACCT_1 = ""

    x = GetTokenInfo(ArchivoEntrada, "F43", "*", "*")
    CARD_ACCEPTOR = Trim(DevuelveParametro(x))

    x = GetTokenInfo(ArchivoEntrada, "F19", "*", "*")
    ACQ_INST = Trim(DevuelveParametro(x))

    x = GetTokenInfo(ArchivoEntrada, "F25", "*", "*")
    POS_COND_CODE = Trim(DevuelveParametro(x))

    x = GetTokenInfo(ArchivoEntrada, "F126", "*", "*")
    sCampo = Trim(DevuelveParametro(x))
    
    sMonRetCta = Mid(sCampo, 23, 6)
    CUR_CODE = Mid(sCampo, 23, 3)
    psMonCta = Mid(sCampo, 26, 3)
    
    x = GetTokenInfo(ArchivoEntrada, "F126", "*", "*")
    sCampoS126 = Trim(DevuelveParametro(x))
       
    x = GetTokenInfo(ArchivoEntrada, "F4", "*", "*")
    sCampo = Trim(DevuelveParametro(x))
    TXN_AMOUNT = sCampo
    sCampoP004 = sCampo
        
    If CUR_CODE = "840" Then
        TXN_AMOUNT = Mid(sCampoS126, 45, 12)
        sCampoP004 = Mid(sCampoS126, 45, 12)
    End If

    ACCT_2 = ""
        
    x = GetTokenInfo(ArchivoEntrada, "F14", "*", "*")
    DATE_EXP = Trim(DevuelveParametro(x))

    x = GetTokenInfo(ArchivoEntrada, "F43", "*", "*")
    CARD_LOCATION = Trim(DevuelveParametro(x))
                             
             
    x = GetTokenInfo(ArchivoEntrada, "F32", "*", "*")
    sCampoP032 = Trim(DevuelveParametro(x))
            
    x = GetTokenInfo(ArchivoEntrada, "F42", "*", "*")
    sCampoP042 = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoEntrada, "F43", "*", "*")
    sCampoP043 = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoEntrada, "F44", "*", "*")
    sCampoP044 = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoEntrada, "F52", "*", "*")
    sCampoP052 = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoEntrada, "F60", "*", "*")
    sCampoP060 = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoEntrada, "F61", "*", "*")
    sCampoP061 = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoEntrada, "F100", "*", "*")
    sCampoS100 = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoEntrada, "F102", "*", "*")
    sCampoS102 = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoEntrada, "F103", "*", "*")
    sCampoS103 = Trim(DevuelveParametro(x))
    
    x = GetTokenInfo(ArchivoEntrada, "F125", "*", "*")
    sCampoS125 = Trim(DevuelveParametro(x))
                                      

    If MESSAGE_TYPE = "0420" Then 'Si es extorno, no validar PIN
        bPinValido = True
    Else
        
        PVV = A.RecuperaPVV_1(PAN)
    
        i = pinverify(gsIPRQ, "444", PAN, "1", "t" & sCampoP052, PVV)
    
        bPinValido = False
        
        If (i = 1) Then ' VALIDO
           bPinValido = True
        Else
            PRCODE = "99"
           Call A.RegistraOperacionLimitesCajeroPOS_1(Now, "990000", PAN, 0)
        End If
    End If
    
    If bPinValido Then

        sTramaSalida = A.EjecutorGlobalNet(MESSAGE_TYPE, TRACE, PRCODE, PAN, TIME_LOCAL, DATE_LOCAL, TERMINAL_ID, _
                     ACCT_1, CARD_ACCEPTOR, ACQ_INST, POS_COND_CODE, TXN_AMOUNT, CUR_CODE, ACCT_2, _
                     DATE_EXP, CARD_LOCATION, psMonCta, lnMovNro)
                                                               
    End If


    '**Armar trama de salida ***************************************
    
    Select Case Left(PRCODE, 2)
        Case "01" 'retiro
            sCampoP039 = Trim(RecuperaValorXML(sTramaSalida, "RESP_CODE"))
            sCampoP042 = "[.....]"
            sCampoP043 = "[.....]"
            
            If MESSAGE_TYPE <> "0420" Then 'Si no es extorno
                sCampoP006 = "123456789012"
                sCampoP038 = Right("000000" & lnMovNro, 6) '"245621"
                
                        
                sSaldoC = Right("00000000000" & Mid(RecuperaValorXML(sTramaSalida, "AMOUNTS_ADD"), 10, 11), 11)
                sSaldoD = Right("00000000000" & Mid(RecuperaValorXML(sTramaSalida, "AMOUNTS_ADD"), 30, 11), 11)
                
                sCampoP044 = "3+" & sSaldoD & "+" & sSaldoC
                sCampoS102 = Left("0" & RecuperaValorXML(sTramaSalida, "ADD_RESP_DATA") & String(28, "0"), 28)
                sCampoS126 = Mid(sCampoS126, 1, 56) & "0" & sSaldoD & "0" & sSaldoC
    
                            
                'Retira soles de una cta dolares
                If Mid(sMonRetCta, 1, 6) = "604840" Or Mid(sMonRetCta, 1, 6) = "840840" Then
                    sCampoS126 = Left(Mid(sCampoS126, 1, 31) & Right("000000" & Mid(RecuperaValorXML(sTramaSalida, "PRIV_USE"), 1, 9), 6) & Left(Mid(RecuperaValorXML(sTramaSalida, "PRIV_USE"), 10, 3) & "000000", 6) & " " & TXN_AMOUNT & "0" & sSaldoD & "0000000000000", 80)
                End If
                
                'Retira dolares de una cta soles
                If Mid(sMonRetCta, 1, 6) = "840604" Then
                     sCampoS126 = Left(Mid(sCampoS126, 1, 31) & Right("000000" & Mid(RecuperaValorXML(sTramaSalida, "PRIV_USE"), 1, 9), 6) & Left(Mid(RecuperaValorXML(sTramaSalida, "PRIV_USE"), 10, 3) & "000000", 6) & " " & TXN_AMOUNT & "0" & sSaldoD & "0000000000000", 80)
                End If
            Else
                sCampoP006 = "[.....]"
                sCampoP038 = "[.....]"
                sCampoP044 = "[.....]"
            End If
                        
        Case "31" 'consulta saldo
            'NSSE 12/12/2008
            sCampoP006 = "[.....]"
            
            sCampoP038 = Right("000000" & lnMovNro, 6) '"245621"
            sCampoP039 = Trim(RecuperaValorXML(sTramaSalida, "RESP_CODE"))

            sCampoP042 = "[.....]"
            sCampoP043 = "[.....]"
            
            sSaldoC = Right("00000000000" & Mid(RecuperaValorXML(sTramaSalida, "AMOUNTS_ADD"), 10, 11), 11)
            sSaldoD = Right("00000000000" & Mid(RecuperaValorXML(sTramaSalida, "AMOUNTS_ADD"), 30, 11), 11)

            sCampoP044 = "3+" & sSaldoD & "+" & sSaldoC
            sCampoS102 = Left(RecuperaValorXML(sTramaSalida, "ADD_RESP_DATA") & String(28, " "), 28)

        
            If psMonCta <> "604" Then
                sCampoS126 = Left(Mid(sCampoS126, 1, 44) & "000000000000" & "0" & sSaldoD & "0000000000000", 80)
            End If

            
        Case "94" 'ultimos movimientos
            sCampoP004 = String(12, "0")
            sCampoP006 = "[.....]"
            'sCampoP011 = "000190"
            sCampoP039 = Mid(sTramaSalida, 1, 2)
            
            
            

            
            sSaldoC = Mid(RecuperaValorXML(sTramaSalida, "AMOUNTS_ADD"), 10, 11)
            sSaldoD = Mid(RecuperaValorXML(sTramaSalida, "AMOUNTS_ADD"), 30, 11)
            sCampoP038 = Right("000000" & lnMovNro, 6) '"000000"
            sCampoP042 = Left(sCampoP042 & Space(15), 15)


            sCampoP044 = "3+00000000000+00000000000" '"3+" & sSaldoD & "+" & sSaldoC

            sCampoP060 = String(15, "0")
            sCampoP061 = String(16, "0")
            sCampoS100 = sCampoP032
            
            sCampoS102 = Left(Right(sTramaSalida, 18) & String(28, " "), 28)
            
            sTramaSalida = UCase(sTramaSalida)
            sCampoS125 = Mid(sTramaSalida, 3, Len(sTramaSalida) - 20)
                       
            Open ArchivoSalida2 For Append As #2
                Print #2, "0215"
            Close #2
            
        Case "96" 'cambio de clave
            sCampoP006 = "[.....]"
            
            
             nPos = InStr(1, sCampoS126, "0600052")
             sPinNew = Mid(sCampoS126, nPos + 28, 16)
             
            '**DAOR 20081110 *********************************************************
            'FALTA CODIGO DE RESUESTA DE CLAVE INCORRECTA
            'TEMPORALMENTE SALDRA CUENTA INVALIDA
            'NSSE FECHA : 11/11/2008
             x = changepin(gsIPRQ, "444", PAN, "1", "t" & sCampoP052, PVV, "t" & sPinNew)
             psPVVNew = DevuelveParametro(x)
             If psPVVNew = "ERR1" Then
                'PIN Invalido
                sCampoP039 = "55" '"53"
             ElseIf psPVVNew = "ERR2" Then
                'No hay conexión con el servidor
                sCampoP039 = "53"
             Else
                Call A.ActualizaPVV_1(psPVVNew, PAN)
                sCampoP039 = "00"
             End If
            '*************************************************************************
             
            sCampoP038 = Right("000000" & lnMovNro, 6) '"245621"
            
            sCampoP039 = "00"
            sCampoP042 = "[.....]"
            sCampoP043 = "[.....]"
            sCampoS102 = String(28, "0")
        
            sCampoS126 = "& 0000200080! Q300058 0006040000000000000000000000000000000000000000000000000000! Q200014 & 0000300142! Q300"
           
        Case "99" 'Pin Invalido
            Dim sXMLTrama As String
            Dim nCondicionTarjeta As Integer
            
            'ARMA TRAMA
            
            sXMLTrama = "<MESSAGE_TYPE = " & MESSAGE_TYPE & " />"
            sXMLTrama = sXMLTrama & " <TRACE = " & TRACE & " />"
            sXMLTrama = sXMLTrama & " <PRCODE = " & PRCODE & " />"
            sXMLTrama = sXMLTrama & " <PAN = " & PAN & " />"
            sXMLTrama = sXMLTrama & " <TIME_LOCAL = " & TIME_LOCAL & " />"
            sXMLTrama = sXMLTrama & " <DATE_LOCAL = " & DATE_LOCAL & " />"
            sXMLTrama = sXMLTrama & " <TERMINAL_ID = " & TERMINAL_ID & " />"
            sXMLTrama = sXMLTrama & " <CARD_ACCEPTOR = " & CARD_ACCEPTOR & " />"
            sXMLTrama = sXMLTrama & " <ACQ_INST = " & ACQ_INST & " />"
            sXMLTrama = sXMLTrama & " <POS_COND_CODE = " & POS_COND_CODE & " />"
            sXMLTrama = sXMLTrama & " <TXN_AMOUNT = " & TXN_AMOUNT & " />"
            sXMLTrama = sXMLTrama & " <CUR_CODE = " & CUR_CODE & " />"
            sXMLTrama = sXMLTrama & " <DATE_EXP = " & DATE_EXP & " />"
            sXMLTrama = sXMLTrama & " <CARD_LOCATION = " & CARD_LOCATION & " />"
            sXMLTrama = sXMLTrama & " <CUENTA = " & "PIN INVALIDO" & " />"
            
            'Registra Trama
            Call A.RegistrarTrama_1("0", sXMLTrama, 1)
                    
            sCampoP006 = "000000000000"
            
            sCampoP038 = Right("000000" & lnMovNro, 6) '"245621"
            
            bRetenerTarjeta = IIf(A.RetenerTarjetaPorPosibleFraude_1(PAN) = 1, True, False)
            nCondicionTarjeta = A.RecuperaCondicionDeTarjeta_1(PAN)
                        
            If nCondicionTarjeta = 10 Or nCondicionTarjeta = 2 Or nCondicionTarjeta = 3 Or nCondicionTarjeta = 50 Then
                bRetenerTarjeta = True
            End If
            
            If Not bRetenerTarjeta Then
                sCampoP039 = "55" 'Falta definir el codigo correcto
            Else
                sCampoP039 = "34"
            End If
            
            sCampoP042 = "[.....]"
            sCampoP043 = "[.....]"
            
            sSaldoC = "00000000000"
            sSaldoD = "00000000000"
            
            sCampoP044 = "[.....]"
            
            sCampoS102 = String(28, "0")
                    

    End Select
    '******************************************************************

   
    
    Open ArchivoSalida For Append As #1
    
    For campo = 1 To 128
        fcampo = "F" + CStr(campo)
        x = GetTokenInfo(ArchivoEntrada, fcampo, "*", "*")
        sCampo = DevuelveParametro(x) 'Trim(DevuelveParametro(x))
        
        
        If (campo = 1) Or (campo = 3) Or (campo = 4) Or (campo = 6) Or (campo = 7) Or (campo = 11) Or (campo = 12) Or (campo = 13) _
            Or (campo = 17) Or (campo = 32) Or (campo = 35) Or (campo = 37) Or (campo = 38) Or (campo = 39) Or (campo = 41) _
            Or (campo = 49) Or (campo = 60) Or (campo = 61) Or (campo = 100) Or (campo = 15) Or (campo = 42) Or (campo = 43) Or (campo = 44) _
            Or (campo = 52) Or (campo = 54) Or (campo = 63) Or (campo = 64) Or (campo = 90) Or (campo = 102) Or (campo = 103) _
            Or (campo = 124) Or (campo = 125) Or (campo = 126) Or (campo = 128) Then
                                              
                                              
            If (campo = 4) Then
                sCampo = sCampoP004
            End If

            If (campo = 6) Then
                sCampo = sCampoP006
            End If

            If (campo = 38) Then
                sCampo = sCampoP038
            End If

            If (campo = 39) Then
                sCampo = sCampoP039
            End If

            If (campo = 42) Then
                sCampo = sCampoP042
            End If

            If (campo = 43) Then
                sCampo = sCampoP043
            End If

            If (campo = 44) Then
               sCampo = sCampoP044
            End If

            If (campo = 52) Then 'PINBlock
                sCampo = "[.....]"
            End If

            If (campo = 60) Then
               sCampo = sCampoP060
            End If
            
            If (campo = 61) Then
               sCampo = sCampoP061
            End If

            If (campo = 100) Then
                 sCampo = sCampoS100
            End If
            
            If (campo = 102) Then
               sCampo = sCampoS102
            End If
            
            If (campo = 103) Then
                 sCampo = sCampoS103
            End If
                        
            If (campo = 125) Then
                 sCampo = sCampoS125
            End If
            
            If (campo = 126) Then
                 sCampo = sCampoS126
            End If
            
          
        Else
            sCampo = "[.....]"
        End If
        
        
        Print #1, sCampo
    Next
    Close #1
    

Set A = Nothing

End Sub


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

Private Function DevuelveParametro(ByVal x As Long) As String
Dim Buffer() As Byte
Dim nLen As Long
Dim res As String
Dim i As Integer
    nLen = lstrlenW(x) * 2
    ReDim Buffer(0 To (nLen - 1)) As Byte
    CopyMemory Buffer(0), ByVal x, nLen
    res = ""
    For i = 0 To nLen - 1
        If (Buffer(i) = 0) Then
            Exit For
        End If
        res = res + Chr(Buffer(i))
    Next
    DevuelveParametro = res
End Function

Private Sub CrearArchivo(ByVal ArchivoSalida As String, ByVal Cadena As String)
Dim fso As Scripting.FileSystemObject
Dim lsMensaje As String
Dim ts As TextStream
Dim lsFile As String

'    lsFile = App.Path & "\" & Format(Date, "yymmdd") & ".RX"
    lsFile = ArchivoSalida
    Set fso = New Scripting.FileSystemObject
'    If fso.FileExists(lsFile) Then
'        If MsgBox("El archivo ya existe, desea reemplazarlo", vbYesNo + vbInformation, "Aviso") = vbNo Then
'            Set fso = Nothing
'            Exit Sub
'        End If
'    End If
    Set ts = fso.CreateTextFile(lsFile, True)
        ts.Write (Cadena)
        'MsgBox "El archivo se generó satisfactoriamente", vbInformation, "Aviso"
        ts.Close
    Set fso = Nothing
End Sub
