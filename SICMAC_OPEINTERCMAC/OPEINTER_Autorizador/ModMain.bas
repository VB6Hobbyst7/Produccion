Attribute VB_Name = "ModMain"
Option Explicit
Global Const gsIPRQ As String = "192.168.15.35:81"
'Global Const gsIPRQPVV As String = "192.168.0.9:81"

Global Const gsCanal As String = "CMAC"
Global Const gsCodInstAutorizadora = "810900"

Public gsMESSAGE_TYPE As String
Public gsTRACE As String
Public gsPRCODE As String
Public gsPAN As String
Public gsTIME_LOCAL As String
Public gsDATE_LOCAL As String
Public gsTERMINAL_ID As String
Public ACCT_1 As String
Public gsCARD_ACCEPTOR As String
Public gsACQ_INST As String
Public gsPOS_COND_CODE As String
Public gsTXN_AMOUNT As String
Public gsCUR_CODE As String
Public ACCT_2 As String
Public gsDATE_EXP As String
Public gsCARD_LOCATION As String
Public gsTRACK2 As String
Public gsDATA_ATM_ADD As String

Public gsMovNro As String
Public gdFecha As Date
Public gsCodAge As String
Public gsCodUser As String
Public gnTramaId As Integer
Public gsOpeCod As String

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
    
Dim Argumento, arrParametros, ArchivoEntrada, ArchivoSalida, ArchivoSalida2 As Variant
Dim sCampo(129) As String


Sub Main()
Dim i As Integer, x As Long, lnTramaID As Long
Dim loAut As AutorizadorIT.ClsAutorizador

Dim bPinValido As Boolean, bRetenerTarjeta As Boolean

Dim sTramaSalida As String, sSaldoD As String, sSaldoC As String
Dim sMonRetCta As String, psMonCta As String, sCodCta As String, sDNI As String

Dim lsCampo As String, lsPVV As String, fCampo As String
Dim lsTramaXML As String, lsCodRespTX As String, lsPRCODE As String
Dim nCondicionTarjeta As Integer, lnDenegada As Integer

    
    Argumento = Command()
    arrParametros = Split(Argumento, " ")
    ArchivoEntrada = CStr(Trim(Replace(arrParametros(0), """", "")))
    ArchivoSalida = CStr(Trim(Replace(arrParametros(1), """", "")))
    ArchivoSalida2 = Replace(ArchivoSalida, ".txt", "_newcode.txt")
    
    sCampo(1) = GetTokenParametro(ArchivoEntrada, "F1")
    sCampo(2) = GetTokenParametro(ArchivoEntrada, "F2")
    sCampo(3) = GetTokenParametro(ArchivoEntrada, "F3")
    gsMESSAGE_TYPE = GetTokenParametro(ArchivoEntrada, "O")
    
    gsPRCODE = GetTokenParametro(ArchivoEntrada, "F3")
    lsPRCODE = gsPRCODE
    
    sCampo(4) = GetTokenParametro(ArchivoEntrada, "F4")
    gsTXN_AMOUNT = sCampo(4)
    sCampo(6) = GetTokenParametro(ArchivoEntrada, "F6")
    sCampo(7) = GetTokenParametro(ArchivoEntrada, "F7")
    sCampo(11) = GetTokenParametro(ArchivoEntrada, "F11")
    gsTIME_LOCAL = GetTokenParametro(ArchivoEntrada, "F12")
    gsDATE_LOCAL = GetTokenParametro(ArchivoEntrada, "F13")
    gsDATE_EXP = GetTokenParametro(ArchivoEntrada, "F14")
    gsACQ_INST = GetTokenParametro(ArchivoEntrada, "F19")
    gsPOS_COND_CODE = GetTokenParametro(ArchivoEntrada, "F25")
    sCampo(32) = GetTokenParametro(ArchivoEntrada, "F32")
    sCampo(35) = GetTokenParametro(ArchivoEntrada, "F35")
    gsPAN = Mid(sCampo(35), 1, 16)
    gsTRACK2 = sCampo(35)
    sCampo(37) = GetTokenParametro(ArchivoEntrada, "F37")
    gsTRACE = sCampo(37)
    sCampo(38) = GetTokenParametro(ArchivoEntrada, "F38")
    gsTERMINAL_ID = GetTokenParametro(ArchivoEntrada, "F41")
    'sCampo(42) = GetTokenParametro(ArchivoEntrada, "F42")
    sCampo(43) = GetTokenParametro(ArchivoEntrada, "F43")
    gsCARD_ACCEPTOR = sCampo(43)
    gsCARD_LOCATION = sCampo(43)
    sCampo(44) = GetTokenParametro(ArchivoEntrada, "F44")
    sCampo(52) = GetTokenParametro(ArchivoEntrada, "F52")
    sCampo(60) = GetTokenParametro(ArchivoEntrada, "F60")
    sCampo(61) = GetTokenParametro(ArchivoEntrada, "F61")
    sCampo(100) = GetTokenParametro(ArchivoEntrada, "F100")
    sCampo(102) = GetTokenParametro(ArchivoEntrada, "F102")
    sCampo(103) = GetTokenParametro(ArchivoEntrada, "F103")
    ACCT_1 = GetTokenParametro(ArchivoEntrada, "F102")
    ACCT_2 = GetTokenParametro(ArchivoEntrada, "F103")
    
    sCampo(125) = GetTokenParametro(ArchivoEntrada, "F125")
    gsDATA_ATM_ADD = sCampo(125)
    sCampo(126) = GetTokenParametro(ArchivoEntrada, "F126")
    sMonRetCta = Mid(sCampo(126), 23, 6)
    gsCUR_CODE = Mid(sCampo(126), 23, 3)
    psMonCta = Mid(sCampo(126), 26, 3)
    
    If Left(gsPRCODE, 3) <> "353" Then
        sCodCta = Right(sCampo(102), 18)
        If Left(gsPRCODE, 2) = "01" Then
            If gsCUR_CODE = "604" And ConvierteAMontoReal(gsTXN_AMOUNT) > 1500 Then
                sDNI = Right(sCampo(103), 8)
            End If
            If gsCUR_CODE = "840" And ConvierteAMontoReal(gsTXN_AMOUNT) > 500 Then
                sDNI = Right(sCampo(103), 8)
            End If
        End If
    Else
        sDNI = Right(sCampo(102), 8)
    End If
   
    If gsCUR_CODE = "840" And Val(gsTXN_AMOUNT) = 0 Then
        gsTXN_AMOUNT = Mid(sCampo(126), 45, 12)
        sCampo(4) = Mid(sCampo(126), 45, 12)
    End If
    
    If gsMESSAGE_TYPE = "0420" Then lnTramaID = sCampo(38)

    Set loAut = New AutorizadorIT.ClsAutorizador
    
    bPinValido = True
    
    If Left(gsPRCODE, 2) = "01" Or Left(gsPRCODE, 3) = "351" Or Left(gsPRCODE, 2) = "94" Then
'        If sCampo(32) = "810900" Or sCampo(32) = "811100" Or sCampo(32) = "810800" Or sCampo(32) = "426154" Or sCampo(32) = "810700" _
'           Or sCampo(32) = "811400" Then
        'If sCampo(32) <> "810900" Then
            lsPVV = loAut.RecuperaPVV_1(gsPAN)
    
            bPinValido = False
            i = pinverify(gsIPRQ, "444", gsPAN, "1", "t" & sCampo(52), lsPVV)
    
            'VALIDO
            
            If (i = 1) Then
               bPinValido = True
            Else
                lsPRCODE = "99"
               'gsPRCODE = "99"
               'Call A.RegistraOperacionLimitesCajeroPOS_1(Now, "990000", PAN, 0)
            End If
        'End If
'        End If
    End If
    
    gdFecha = Now
    gsCodAge = "01"
    gsCodUser = "CMAC"
    gsOpeCod = getOpeCod(gsPRCODE, gsMESSAGE_TYPE)
    lsTramaXML = GeneraTramaEnXML()
    
    gsMovNro = loAut.PIT_GeneraMovNro(gdFecha, gsCodAge, gsCodUser)
    
    gnTramaId = loAut.PIT_nRegistrarTramaRecepcion(gsMovNro, gsOpeCod, lsTramaXML, gsPRCODE, sCodCta, gsPAN, sDNI, gsCUR_CODE, ConvierteAMontoReal(gsTXN_AMOUNT), sCampo(32))
    Call loAut.PIT_RegistraBitacora(gsPAN, gsCanal, Now, gsTRACE, "Registro Trama Recepcion", "", gnTramaId)
    
    If bPinValido Then
        sTramaSalida = loAut.EjecutorGlobalNet(gsMESSAGE_TYPE, gsTRACE, gsPRCODE, gsPAN, gsTIME_LOCAL, gsDATE_LOCAL, gsTERMINAL_ID, _
                     ACCT_1, gsCARD_ACCEPTOR, gsACQ_INST, gsPOS_COND_CODE, gsTXN_AMOUNT, gsCUR_CODE, ACCT_2, _
                     gsDATE_EXP, gsCARD_LOCATION, psMonCta, gsMovNro, lnTramaID, sCodCta, sDNI)
    Else
        sTramaSalida = GeneraXMLSalida("55", "", gnTramaId, "", gsCUR_CODE, "0", "Contraseña Invalida", gsMESSAGE_TYPE)

    End If
        
    lsCodRespTX = Trim(RecuperaValorXML(sTramaSalida, "RESP_CODE"))

    lnDenegada = IIf(lsCodRespTX = "00", 0, 1)
    
    Call loAut.PIT_RegistrarTramaEnvio(gnTramaId, sTramaSalida, lsCodRespTX, lnDenegada)
    Call loAut.PIT_RegistraBitacora(gsPAN, gsCanal, Now, gsTRACE, "Registro Trama Envio", "", gnTramaId)
    
    sCampo(38) = Right("000000" & gnTramaId, 6)
    sCampo(39) = Trim(RecuperaValorXML(sTramaSalida, "RESP_CODE"))
    sCampo(44) = "3+00000000000+00000000000"
    'sCampo(52) = "[.....]"
    
    sCampo(100) = Left(gsCodInstAutorizadora & "00000000000", 11)
    
    Select Case Left(lsPRCODE, 2)
        Case "01" 'retiro

            sSaldoC = Right("00000000000" & Mid(RecuperaValorXML(sTramaSalida, "AMOUNTS_ADD"), 10, 11), 11)
            sSaldoD = Right("00000000000" & Mid(RecuperaValorXML(sTramaSalida, "AMOUNTS_ADD"), 30, 11), 11)
            
            sCampo(44) = "3+" & sSaldoD & "+" & sSaldoC
            sCampo(125) = Trim(RecuperaValorXML(sTramaSalida, "PRIV_USE"))
            sCampo(126) = Mid(sCampo(126), 1, 56) & "0" & sSaldoD & "0" & sSaldoC

        Case "20" 'Deposito

            sSaldoC = Right("00000000000" & Mid(RecuperaValorXML(sTramaSalida, "AMOUNTS_ADD"), 10, 11), 11)
            sSaldoD = Right("00000000000" & Mid(RecuperaValorXML(sTramaSalida, "AMOUNTS_ADD"), 30, 11), 11)
            
            sCampo(44) = "3+" & sSaldoD & "+" & sSaldoC
            'sCampo(125) = "[.....]"
            sCampo(125) = Trim(RecuperaValorXML(sTramaSalida, "PRIV_USE"))
            sCampo(126) = Mid(sCampo(126), 1, 56) & "0" & sSaldoD & "0" & sSaldoC
            
        Case "31" 'consulta saldo

            sSaldoC = Right("00000000000" & Mid(RecuperaValorXML(sTramaSalida, "AMOUNTS_ADD"), 10, 11), 11)
            sSaldoD = Right("00000000000" & Mid(RecuperaValorXML(sTramaSalida, "AMOUNTS_ADD"), 30, 11), 11)

            sCampo(44) = "3+" & sSaldoD & "+" & sSaldoC

            If psMonCta <> "604" Then
                sCampo(126) = Left(Mid(sCampo(126), 1, 44) & "000000000000" & "0" & sSaldoD & "0000000000000", 80)
            End If
            
        Case "35" 'consulta de cuentas de ahorro y credito

            sTramaSalida = UCase(sTramaSalida)
            sCampo(125) = Trim(RecuperaValorXML(sTramaSalida, "PRIV_USE"))
                        
            Open ArchivoSalida2 For Append As #2
                Print #2, "0215"
            Close #2
        Case "50" 'Pago de credito
            
            sTramaSalida = UCase(sTramaSalida)
            sCampo(125) = Trim(RecuperaValorXML(sTramaSalida, "PRIV_USE"))

            Open ArchivoSalida2 For Append As #2
                Print #2, "0215"
            Close #2
        Case "94" 'ultimos movimientos

            sCampo(125) = Trim(RecuperaValorXML(sTramaSalida, "PRIV_USE"))
            
            Open ArchivoSalida2 For Append As #2
                Print #2, "0215"
            Close #2
            
        Case "99" 'Pin Invalido
            
            bRetenerTarjeta = IIf(loAut.RetenerTarjetaPorPosibleFraude_1(gsPAN) = 1, True, False)
            nCondicionTarjeta = loAut.RecuperaCondicionDeTarjeta_1(gsPAN)
            
            If nCondicionTarjeta = 10 Or nCondicionTarjeta = 2 Or nCondicionTarjeta = 3 Or nCondicionTarjeta = 50 Then
                bRetenerTarjeta = True
            End If
            
            If Not bRetenerTarjeta Then
                sCampo(39) = "55"
            Else
                sCampo(39) = "34"
            End If

            sCampo(44) = "[.....]"
            
    End Select
    
    Open ArchivoSalida For Append As #1
    For i = 1 To 128
        fCampo = "F" + CStr(i)
        x = GetTokenInfo(ArchivoEntrada, fCampo, "*", "*")
        lsCampo = DevuelveParametro(x)
        
        If sCampo(i) <> "" Then 'Personalizado
            lsCampo = sCampo(i)
        End If
        
        'If lsCampo = "" And sCampo(i) = "" Then
        '    lsCampo = "[.....]"
        'ElseIf lsCampo <> "" And sCampo(i) = "" Then
        '    lsCampo = "[.....]"
        'Else
        '    lsCampo = sCampo(i)
        'End If
        
        Print #1, lsCampo
    Next
    Close #1
    
    Set loAut = Nothing

End Sub

Public Function GetTokenParametro(ByVal ArchivoEntrada As String, ByVal Token As String)
    Dim x As Long
    Dim sCampo As String
    
    x = GetTokenInfo(ArchivoEntrada, Token, "*", "*")
    sCampo = Trim(DevuelveParametro(x))
    GetTokenParametro = sCampo
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

    lsFile = ArchivoSalida
    Set fso = New Scripting.FileSystemObject
    Set ts = fso.CreateTextFile(lsFile, True)
        ts.Write (Cadena)
        ts.Close
    Set fso = Nothing
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

Public Function GeneraTramaEnXML() As String
Dim sXMLTrama As String
    sXMLTrama = "<MESSAGE_TYPE = " & gsMESSAGE_TYPE & " />"
    sXMLTrama = sXMLTrama & " <TRACE = " & gsTRACE & " />"
    sXMLTrama = sXMLTrama & " <PRCODE = " & gsPRCODE & " />"
    sXMLTrama = sXMLTrama & " <PAN = " & gsPAN & " />"
    sXMLTrama = sXMLTrama & " <TIME_LOCAL = " & gsTIME_LOCAL & " />"
    sXMLTrama = sXMLTrama & " <DATE_LOCAL = " & gsDATE_LOCAL & " />"
    sXMLTrama = sXMLTrama & " <TERMINAL_ID = " & gsTERMINAL_ID & " />"
    sXMLTrama = sXMLTrama & " <CARD_ACCEPTOR = " & gsCARD_ACCEPTOR & " />"
    sXMLTrama = sXMLTrama & " <ACQ_INST = " & gsACQ_INST & " />"
    sXMLTrama = sXMLTrama & " <POS_COND_CODE = " & gsPOS_COND_CODE & " />"
    sXMLTrama = sXMLTrama & " <TXN_AMOUNT = " & gsTXN_AMOUNT & " />"
    sXMLTrama = sXMLTrama & " <CUR_CODE = " & gsCUR_CODE & " />"
    sXMLTrama = sXMLTrama & " <DATE_EXP = " & gsDATE_EXP & " />"
    sXMLTrama = sXMLTrama & " <CARD_LOCATION = " & gsCARD_LOCATION & " />"
    sXMLTrama = sXMLTrama & " <TRACK2 = " & gsTRACK2 & " />"
    sXMLTrama = sXMLTrama & " <DATA_ATM_ADD = " & gsDATA_ATM_ADD & " />"
    sXMLTrama = sXMLTrama & " <ACCT_1 = " & ACCT_1 & " />"
    sXMLTrama = sXMLTrama & " <ACCT_2 = " & ACCT_2 & " />"
    GeneraTramaEnXML = sXMLTrama
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
Public Function ConvierteAMontoReal(ByVal psMontoTxN As String) As Double

    ConvierteAMontoReal = CDbl(Mid(psMontoTxN, 1, Len(psMontoTxN) - 2) & "." & Right(psMontoTxN, 2))
    
End Function
