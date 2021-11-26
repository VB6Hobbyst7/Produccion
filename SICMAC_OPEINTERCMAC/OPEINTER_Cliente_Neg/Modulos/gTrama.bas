Attribute VB_Name = "gOpeIntercmac"
Option Explicit

Global Const gsIPRQ As String = "192.168.15.35:81"
Global Const gsPATHTramas As String = "C:\SICMACM\SICMAC_OPEINTERCMAC\Tramas\"

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
Public ArchivoEntrada As String, ArchivoSalida As String
 
'Private Declare Function pinverify _
'Lib "PINVerify.dll" _
'             (ByVal ippuerto As String, _
'              ByVal key As String, _
'              ByVal PAN As String, _
'              ByVal pvki As String, _
'              ByVal pin As String, _
'              ByVal pvv As String _
'             ) As Integer
                 
Private Declare Function RQxDFClientSend _
Lib "RQxDFClientDLL.dll" _
             (ByVal Name As String, _
              ByVal SecCode As String, _
              ByVal Oper As String, _
              ByVal FileInput As String, _
              ByVal FileOutput As String, _
              ByVal Ip As String, _
              ByVal Port As Integer, _
              ByVal TimeOut As Integer) As Integer

 Private Declare Function GetTokenInfo _
 Lib "RQxDFTk.dll" _
                 (ByVal file As String, _
                  ByVal info As String, _
                  ByVal subinfo As String, _
                  ByVal tokenitem As String _
                 ) As Long
    
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

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

Public Sub GeneraTrama(psNumTarj As String, psCodcta As String, psOpeCod As String, _
                        ByVal pnMonto As Double, psTrack2 As String, psCodAutCMAC As String, _
                        pnMoneda As Moneda, psDni As String, psNomCmac As String, psCodCMACOrig As String, psClave As String)
Dim sTramaSalida As String
Dim lsCampo(129) As String '**Campo Primarios (1 al 64) y Secundarios (64 al 128)
Dim lsTipoOperacionId As String
Dim lsMonto As String
Dim lsDecimal As String
Dim lsMoneda As String

    ArchivoEntrada = "_ISO" & Format(Date, "yyyymmdd") & Format(Time, "hhmmss") & "_" & "0200.txt"
    If psOpeCod = "260506" Then
        ArchivoSalida = "_ISO" & Format(Date, "yyyymmdd") & Format(Time, "hhmmss") & "_" & "0215Salida.txt"
    Else
        ArchivoSalida = "_ISO" & Format(Date, "yyyymmdd") & Format(Time, "hhmmss") & "_" & "0210Salida.txt"
    End If
    
    If pnMoneda = 1 Then
        lsMoneda = "604"
    Else
        lsMoneda = "840"
    End If
    
    If pnMonto > 0 Then
        lsMonto = CStr(pnMonto)
        lsDecimal = Right(lsMonto, 2)
        lsMonto = Format(lsMonto, "000000000000")
    Else
        lsMonto = "000000000000"
    End If
    
    lsCampo(1) = psNumTarj
    lsCampo(3) = lsTipoOperacionId
    lsCampo(4) = lsMonto
    lsCampo(7) = Format(Now, "MMDDHHMMSS")
    lsCampo(11) = "000001"
    lsCampo(12) = "113020"
    lsCampo(13) = Format(Date, "MMDD")
    lsCampo(17) = Format(Date, "MMDD")
    lsCampo(32) = psCodCMACOrig
    lsCampo(33) = psCodAutCMAC
    lsCampo(35) = psTrack2
    lsCampo(37) = "6043        "
    lsCampo(41) = "TIDE-01-109     "
    lsCampo(43) = psNomCmac & "            N 0PE"
    lsCampo(49) = lsMoneda
    lsCampo(52) = psClave
    lsCampo(60) = "VENTCER10000000"
    lsCampo(61) = "MAYNCER11100P"
    lsCampo(100) = "00000000000"
    lsCampo(102) = psCodcta
    lsCampo(103) = " "
    lsCampo(126) = "& 0000200080! Q300058 604604000000000000000 000000000000000000000000000000000000"

    Select Case gsOpeCod
        Case "260503" 'Retiro
            lsTipoOperacionId = "011100"
            lsCampo(3) = lsTipoOperacionId
            'lsCampo(4) = "000000010000"
        Case "260501" 'Deposito
            lsTipoOperacionId = "200011"
            lsCampo(3) = lsTipoOperacionId
        Case "107001" 'Pago de Credito
            lsTipoOperacionId = "500035"
            lsCampo(3) = lsTipoOperacionId
        Case "260505" 'Consulta de Cuentas de Ahorro
            lsTipoOperacionId = "351100"
            lsCampo(3) = lsTipoOperacionId
        Case "260506" 'Consulta de Movimientos de una cuenta Ahorro
            lsTipoOperacionId = "941100"
            lsCampo(3) = lsTipoOperacionId
        Case "107004" 'Consulta de Cuentas de Credito o Prestamo
            lsTipoOperacionId = "353500"
            lsCampo(3) = lsTipoOperacionId
            lsCampo(102) = psDni
    End Select
    
    
    Call crearArchivoTrama(ArchivoEntrada, lsCampo)
    
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

Sub RegistrarOperacionInterCMAC(psNumTarj As String, ByVal psClave As String, ByVal psOpeCod As String, psTrack2 As String, pnMoneda As Moneda, _
                                psDni As String, psPersCodCMAC As String, psMovNro As String, sLpt As String, psOpeDescripcion As String, _
                                psNombreCMAC As String, Optional psCuenta As String, Optional ByVal pnMonto As Double, Optional psGlosa As String, _
                                Optional psIFTipo As String, Optional pbImpTMU As Boolean = False)
    Dim lnResp As Integer
    Dim lsCampo38 As String
    Dim lsCampo39 As String
    Dim lsCampo125 As String
    Dim x As Long
    Dim lsFechaHoraGrab As String
    Dim lsBoleta As String
    Dim nFicSal As Integer
    Dim lsCodAutCMAC As String
    Dim lsImpBoleta As String
    Dim lsNomTit As String
    Dim lsTexto As String
    Dim lsNomCMACDest As String
    Dim lsCodAutDest As String
    Dim lsNumRecRegXCom As String
    Dim lnMovNroCom As Long
    Dim lsUser As String
    Dim clsFun As COMOpeInterCMAC.dFuncionesNeg
    
    Set clsFun = New COMOpeInterCMAC.dFuncionesNeg
    
    lsCodAutCMAC = clsFun.DevuelveCodAutorizaCMAC(psPersCodCMAC, lsNomCMACDest)
    
    Call GeneraTrama(psNumTarj, "", "", pnMonto, psTrack2, lsCodAutCMAC, pnMoneda, psDni, psNombreCMAC, "810900", psClave)
    
    lnResp = RQxDFClientSend("CAJA MAYNAS", "05090B0A3F33353F3431", "0160000100200", ArchivoEntrada, ArchivoSalida, "192.168.15.35", 1000, 120)
    MsgBox lnResp

    Select Case lnResp
        Case 1
            MsgBox "Congrats!!! recibio respuesta", vbCritical + vbExclamation, "Mensaje del Sistema"
        Case -1
            MsgBox "No Hay Respuesta del Receptor", vbCritical + vbExclamation, "Mensaje del Sistema"
        Case -2
            MsgBox "No existe archivo de Entrada", vbCritical + vbExclamation, "Mensaje del Sistema"
        Case -3
            MsgBox "Error en codigo de seguridad", vbCritical + vbExclamation, "Mensaje del Sistema"
    End Select
    
    If lnResp = 1 Then
        
        'Comenzar a leer campos del archivo de salida
        x = GetTokenInfo(ArchivoSalida, "F38", "*", "*")
        lsCampo38 = Trim(DevuelveParametro(x))
        x = GetTokenInfo(ArchivoSalida, "F39", "*", "*")
        lsCampo39 = Trim(DevuelveParametro(x))
        x = GetTokenInfo(ArchivoSalida, "F125", "*", "*")
        lsCampo125 = Trim(DevuelveParametro(x))
        lsTexto = lsCampo125
        
        MsgBox "Prueba " & lsCampo38 & " " & lsCampo39 & " " & lsCampo125
        
        If lsCampo39 = "00" Then
            Select Case psOpeCod
                Case "260501" 'Deposito
                    lsNomTit = "DEPOSITO"
                    clsFun.CapOpeAhoCMACPIT psMovNro, pnMoneda, psOpeCod, psOpeDescripcion, pnMonto, psCuenta, 0, psPersCodCMAC, psNombreCMAC, gsNomAge, 1, gnMovNro, lnMovNroCom, psGlosa
    '            Case "260502" 'Deposito Cheque
    '                clsFun.CapOpeAhoCMACLlamada sMovNro, nmoneda, sGlosa, "260502", nExtracto, sDescOperacion, nMonto, sCuenta, nSaldo, sPersCodCMAC, sNombreCMAC, sCliente, TpoDocCheque, sDocumento, , gsNomAge, sLpt, CDbl(Val(lblComision.Caption)), gMonedaNacional, CCur(Val(Me.lblITF.Caption)), , sBenPersLavDinero, lsBoleta, txtCuenta.Text, gbImpTMU, , , , , , gnMovNro
                Case "260503" 'Retiro
                    lsNomTit = "RETIRO"
                    clsFun.CapOpeAhoCMACPIT psMovNro, pnMoneda, psOpeCod, psOpeDescripcion, pnMonto, psCuenta, 0, psPersCodCMAC, psNombreCMAC, gsNomAge, 1, gnMovNro, lnMovNroCom, psGlosa
    '            Case "260504" 'Retiro OP
    '                clsFun.CapOpeAhoCMACLlamada sMovNro, nmoneda, sGlosa, "260504", nExtracto, sDescOperacion, nMonto, sCuenta, nSaldo, sPersCodCMAC, sNombreCMAC, sCliente, TpoDocOrdenPago, sDocumento, sCtaAbono, gsNomAge, sLpt, CDbl(Val(lblComision.Caption)), gMonedaNacional, CCur(Val(Me.lblITF.Caption)), , sBenPersLavDinero, lsBoleta, txtCuenta.Text, gbImpTMU, , , , , , gnMovNro
                Case "107001" 'Pago Credito
                    lsNomTit = "PAGO DE CREDITOS"
                    lsFechaHoraGrab = fgFechaHoraGrab(psMovNro)
                    'clsFun.nOpeCMACLlamadaCredPignoraticio 107001, "", "", "", "", "", gMonedaNacional, "", "", 100
                    clsFun.nOpeCMACLlamadaCredPignoraticio psOpeCod, lsFechaHoraGrab, psMovNro, psGlosa, psPersCodCMAC, psIFTipo, pnMoneda, psCuenta, "", pnMonto, False, 1, gMonedaNacional, "", gsNomAge, sLpt, "", , "", lsBoleta, , , , , , , , gnMovNro
                Case "260505" 'Consulta de Cuentas de Ahorro
                    lsNomTit = "CONSULTAS DE CUENTAS - AHORRO"
                    clsFun.nOpeCMACConsultaCTAAhoCreMov psMovNro, gnMovNro, pnMoneda, 1, psOpeCod, lnMovNroCom, psOpeDescripcion, gsNomAge, sLpt, lsBoleta, psCuenta, psPersCodCMAC, psNombreCMAC, "", psGlosa
                Case "260506" 'Consulta de Movimientos de Cuentas de Ahorro
                    lsNomTit = "CONSULTA DE MOVIMIENTOS"
                    clsFun.nOpeCMACConsultaCTAAhoCreMov psMovNro, gnMovNro, pnMoneda, 1, psOpeCod, lnMovNroCom, psOpeDescripcion, gsNomAge, sLpt, lsBoleta, psCuenta, psPersCodCMAC, psNombreCMAC, "", psGlosa
                Case "107004" 'Consulta de Cuentas de Credito
                    lsNomTit = "CONSULTA DE CREDITOS"
                    clsFun.nOpeCMACConsultaCTAAhoCreMov psMovNro, gnMovNro, pnMoneda, 1, psOpeCod, lnMovNroCom, psOpeDescripcion, gsNomAge, sLpt, lsBoleta, psCuenta, psPersCodCMAC, psNombreCMAC, "", psGlosa
            End Select
        Else
            MsgBox "Error " & lsCampo39
            Exit Sub
        End If
        
    lsUser = Right(psMovNro, 4)
    
    lsImpBoleta = lsImpBoleta & clsFun.ImprimeBoleta(lsNomTit, lsTexto, CStr(gnMovNro), CStr(lnMovNroCom), psCuenta, gdFecSis, psNombreCMAC, lsNomCMACDest, psNumTarj, lsCodAutDest, lsUser, psOpeCod, psDni, pnMoneda, pbImpTMU)
    
    Dim vVez As Integer
    
    vVez = 1
    'nBan = True
        
    lsImpBoleta = lsImpBoleta & Chr(10) & Chr(10) & Chr(10) & Chr(10)

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
    Set clsFun = Nothing
End Sub

'Funcion que llena un Combo con un recordset
Sub Llenar_Combo_con_RecordsetPIT(prs As ADODB.Recordset, pcboObjeto As ComboBox)

pcboObjeto.Clear
Do While Not prs.EOF
    pcboObjeto.AddItem Trim(prs!cPersNombre) & Space(70) & Trim(prs!cPersCod)
    prs.MoveNext
Loop
prs.Close
    
End Sub

