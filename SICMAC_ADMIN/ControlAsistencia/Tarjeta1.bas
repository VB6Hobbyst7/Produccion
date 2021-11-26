Attribute VB_Name = "Tarjeta1"
'Centralizacion
'Usuario :NSSE
'Fecha :05/12/2000

'*****************************************************************************
'*      List of  Functions implemented in the file:
'*      Function name    -   Description
'*      Main             -  Program entry point
'*      MakeWord         -  Converts an integer to a positive long
'*      McrCancelEvent   -  Executes Api function PpMcrCancelEvent
'*      McrRead          -  Executes Api function PpMcrRead
'*      SignalPINPad     -  Executes Api function PpDvcSignal
'*      WriteToLcd       -  Writes a string to the PINPad Lcd
'*
'*****************************************************************************

Option Explicit
' Copy constants and declarations required from apiuser.vbh
' file here or include the "apiuser.vbh" module into your project

'*************************
' DECLARACIONES   ********
'*************************

' Constantes para Lectora de Tarjeta

Global Const ID_SYS = &H0
Global Const ID_PARENT_KS = &H100
Global Const ID_PARENT_PP = &H200
Global Const MAX_PARENT = 2 + 1
Global Const ID_CHILD_DVC = &H1
Global Const ID_CHILD_LCD = &H2
Global Const ID_CHILD_MCRE = &H3
Global Const ID_CHILD_MCR = &H4
Global Const ID_CHILD_SCT = &H5
Global Const ID_CHILD_COM = &H6
Global Const ID_CHILD_CON = &H7
'Global Const MAX_CHILD = 7 + 1
Global Const MAX_CHILD_INDEX = 7
Global Const MAX_CHILD = 7 + 1 + MAX_CHILD_INDEX
Global Const CHILD_INDEX_BIT_SHIFT = 5
Global Const INDEX_OFFSET = 7
Global Const ID_SMARTKEY_INDEX = &HE0
Global Const MAX_SMARTKEY_ADDRESS = 64
Global Const MAX_PORT = 8 + 1

' Constantes de Error de Lectora de Tarjeta

Global Const EROF = 61000              'Error code offset, limit = 61255
Global Const API_CLEAR = 0             'No errors encountered
Global Const ERR_API_ARG = 61001       'At least one API argument is wrong
Global Const ERR_GEN_COM = 61002       'KeyStation is NAKing
Global Const ERR_GEN_TO = 61003        'General timeout
Global Const ERR_UNKNOWN_FUNC = 61004  'Function not implemented
Global Const ERR_KS_OPEN = 61005       'Could not open KeyStation
Global Const ERR_KS_CLOSE = 61006      'Could not close KeyStation
Global Const ERR_NO_DVC = 61007        'No LL_ACK received
Global Const ERR_BAD_PACKET = 61008    'Packet is corrupted
Global Const ERR_API_BUSY = 61009      'Async API process is busy
Global Const ERR_DVC_LOCKED = 61010    'Async API dvc is prohibited to process
Global Const ERR_UNDEF_HPROC = 61011   'Async Undefined process handle
Global Const ERR_EVENT_ABORTED = 61012 'Sync event aborted
Global Const ERR_DMON_OFF = 61013      'DMON application is not running now
Global Const ERR_MEMORY = 61014        'Cannot allocate memory for the API packet
Global Const ERR_ALR_OPEN = 61015      'KeyStation is already open, and ready to be used
Global Const PP_BUSY = 61060
Global Const ERR_GENERAL = 65535

' Funciones Prototipo Generales / Prototipos API para Lectora de Tarjetas

Declare Function PpConCancelEvent Lib "PPN.DLL" () As Integer
Declare Function PpLcdClear Lib "PPN.DLL" () As Integer
Declare Function PpLcdPutString Lib "PPN.DLL" (ByVal wLcdPosition%, ByVal pbMessString$, ByVal wMessLength%) As Integer
Declare Function PpMcrRead Lib "PPN.DLL" (ByVal pi%, ByVal Track1$, ByVal bufLen1%, actLen1%, ByVal Track2$, ByVal bufLen2%, actLen2%, ByVal Track3$, ByVal bufLen3%, actLen3%) As Integer
Declare Function PpDvcSignal Lib "PPN.DLL" () As Integer
Declare Function PpMcrCancelEvent Lib "PPN.DLL" () As Integer
Declare Function PpConGetData Lib "PPN.DLL" (ByVal pi%, ByVal dataBuf$, ByVal nk%, ByVal eo%, ByVal ep%, ByVal tk%, Al%) As Integer

Declare Function ApiOpen Lib "BIOSN.DLL" (ByVal pcComPort$, ByVal wLen%, ByVal pcNull$, ByVal wNull%) As Integer
Declare Function ApiClose Lib "BIOSN.DLL" () As Integer
Declare Function ApiReadError Lib "BIOSN.DLL" (ByVal dvc%) As Integer
Declare Function ApiCheckIn Lib "BIOSN.DLL" (ByVal dvc%) As Integer
Declare Function ApiCheckOut Lib "BIOSN.DLL" (ByVal dvc%) As Integer
Declare Function ApiIsCheckIn Lib "BIOSN.DLL" (ByVal dvc%) As Integer
Declare Function ApiIsDmonRunning Lib "BIOSN.DLL" () As Integer
Declare Function ApiIsOpen Lib "BIOSN.DLL" () As Integer
Declare Function ApiGetCommId Lib "BIOSN.DLL" (ByVal pcId$) As Integer
Declare Function ApiEncrypt Lib "BIOSN.DLL" (ByVal keyp$, ByVal wKeyLen%, ByVal vecp$, ByVal wVecLen%, ByVal datap$, ByVal wDataLen%) As Integer
Declare Function ApiDecrypt Lib "BIOSN.DLL" (ByVal keyp$, ByVal wKeyLen%, ByVal vecp$, ByVal wVecLen%, ByVal datap$, ByVal wDataLen%) As Integer



'------------------- pinpad verifone modelo sc552 - CMACICA --------------------------------------------

Public Type RegOut
   Trama As String * 16
End Type
 
Declare Function Connect Lib "HPSerialDLL.dll" (ByVal port As Long) As Integer
Declare Function Reset Lib "HPSerialDLL.dll" () As Integer
Declare Function ReadCardIni Lib "HPSerialDLL.dll" () As Integer
Declare Function ReadCard Lib "HPSerialDLL.dll" (ByRef Out As RegOut) As Integer
Declare Function ReadPin Lib "HPSerialDLL.dll" (ByRef Out As RegOut) As Integer
Declare Function ReadPinIni Lib "HPSerialDLL.dll" (MK As RegOut, WK As RegOut, Tarjeta As RegOut) As Integer
Declare Function Disconnect Lib "HPSerialDLL.dll" () As Integer
Declare Function TransferMK Lib "HPSerialDLL.dll" (sNMK As RegOut, sMK As RegOut) As Integer
Declare Function ShowMessage Lib "HPSerialDLL.dll" (sMessage As RegOut) As Integer
Declare Function ConnectionTest Lib "HPSerialDLL.dll" () As Integer
Declare Function UnencryptedPIN Lib "HPSerialDLL.dll" () As Integer



'--------------------------------------------------------------------------------------------------------



' Const definitions
Const MAX_CHAR_TRACK1 = 76 + 1 'Maximum number of characters on track1 is 76
Const MAX_CHAR_TRACK2 = 37 + 1 'Maximum number of characters on track1 is 37
Const MAX_CHAR_TRACK3 = 104 + 1 'Maximum number of characters on track1 is 104

' Local module Variables ( for 'ppmcrtst.bas' module)
Dim StopPpMcrRead% 'Rise this flag to stop execution of PpMcrRead API
Dim gsErrLec As String

Public GnTipoPinPad As Integer


'*****************************************************************************
'*
'* FUNCTION NAME:           MakeWord
'*
'* DESCRIPTION:
'*
'* This function takes an integer and returns corresponding
'* positive long. We need this function because API's returns
'* unsigned integer (word), but Visual basic has only signed
'* numeric types.
'*
'*****************************************************************************
'
Function MakeWord(ByVal i As Integer) As Long
If (i < 0) Then
    MakeWord = i And &HFFFF&
End If
End Function

'*****************************************************************************
'*
'* FUNCTION NAME:           McrCancelEvent
'*
'* DESCRIPTION:
'*
'* This function rises flag to stop exectution of PpMcrRead function,
'* checks into corresponding device, executes PpMcrCancelEvent function,
'* checks out of the device
'*
'*
'*****************************************************************************
'
Sub McrCancelEvent()
Dim wRetVal& ' Use long to represent word (unsigned integer)
Dim Msg

If ApiCheckIn(ID_PARENT_PP Or ID_CHILD_MCR) <> 0 Then
    
    MsgBox "CheckIn failed", , "PpMcrCancelEvent"
    Exit Sub
    
End If

wRetVal = ERR_API_BUSY
StopPpMcrRead = True

While wRetVal = ERR_API_BUSY
    wRetVal = MakeWord(PpMcrCancelEvent())
    DoEvents  ' Lets "Dmonpro" do communication
Wend

If ApiCheckOut(ID_PARENT_PP Or ID_CHILD_MCR) Then
End If

If wRetVal <> 0 Then

    gsErrLec = wRetVal
'    Msg = "Error code:" & Str$(wRetVal)
'    MsgBox Msg, , "PpMcrRead"

End If

End Sub

'*****************************************************************************
'*
'* FUNCTION NAME:           McrRead
'*
'* DESCRIPTION:
'*
'* This function checks into corresponding device,
'* executes PpMcrRead function with PromptIndex set to 255(ie no prompt),
'* checks out of the device.
'* The function returns False if any error occurs, otherwise it returns True
'*
'*****************************************************************************
'
Function McrRead(Track1$, ByVal len1%, RetLen1%, Track2$, ByVal len2%, RetLen2%, Track3$, ByVal Len3%, RetLen3%) As Integer
            
Dim wRetVal&
Dim szTrack1 As String * MAX_CHAR_TRACK1  'Always allocate space for the buffer,
                                          'Don't pass variable length strings
                                          'to the API functions
                                          
Dim szTrack2 As String * MAX_CHAR_TRACK2
Dim szTrack3 As String * MAX_CHAR_TRACK3
    

'To execute PpMcrRead you need to check into PINPad MCR device
'As an alternative you can try to check into all devices you will
'at the start of the program in the "Sub Main" and check out
'at the end of the program

If ApiCheckIn(ID_PARENT_PP Or ID_CHILD_MCR) <> 0 Then
    
    MsgBox "CheckIn failed", , "PpMcrRead"
    Exit Function
    
End If

StopPpMcrRead = False

wRetVal = ERR_API_BUSY

                      
'Call the function continuously while "ERR_API_BUSY" returned
'and the StoPpMcrRead flag is not raised


While wRetVal = ERR_API_BUSY And StopPpMcrRead <> True
    wRetVal = MakeWord(PpMcrRead(255, szTrack1, len1, RetLen1, szTrack2, len2, RetLen2, szTrack3, Len3, RetLen3))
    DoEvents  ' Lets "Dmonpro" do communication
Wend

If ApiCheckOut(ID_PARENT_PP Or ID_CHILD_MCR) <> 0 Then
End If

If StopPpMcrRead Then
    Exit Function
End If

If wRetVal = 0 Then

    Track1 = Left$(szTrack1, RetLen1) 'Api Don't return zero ended strings
                                      'use "Left" function to make string
                                      'zero ended
    Track2 = Left$(szTrack2, RetLen2)
    Track3 = Left$(szTrack3, RetLen3)
    McrRead = True
Else
    gsErrLec = wRetVal
'    MsgBox "Error code: " & Str$(wRetVal), , "PpMcrRead"
End If

End Function

'*****************************************************************************
'*
'* FUNCTION NAME:           SignalPINPad
'*
'* DESCRIPTION:
'*
'* This function checks into corresponding device,
'* executes PpDvcSignal function, checks out of the device
'* The function returns True if "ApiCheckIn" failed or
'* no error occured
'*
'*****************************************************************************
'
Function SignalPINPad() As Integer
Dim wRetVal&
Dim Msg

If ApiCheckIn(ID_PARENT_PP Or ID_CHILD_DVC) <> 0 Then
    
    ' Because some other application already using PINPad
    ' Lets assume that PINPad is present

    SignalPINPad = True
    Exit Function
    
End If

wRetVal = ERR_API_BUSY

While wRetVal = ERR_API_BUSY
    wRetVal = MakeWord(PpDvcSignal())
    DoEvents  ' Lets "Dmonpro" do communication
Wend

If ApiCheckOut(ID_PARENT_PP Or ID_CHILD_DVC) Then
End If

If wRetVal = 0 Then
    SignalPINPad = True
Else
    gsErrLec = wRetVal
End If

End Function

'*****************************************************************************
'*
'* FUNCTION NAME:           WriteToLcd
'*
'* DESCRIPTION:
'*
'* This subroutine checks into corresponding device,
'* executes PpLcdPutString function, checks out of the device
'* It always starts write at position 1 of the LCD
'*
'*****************************************************************************
'

Function WriteToLcd(szMessage$) As Boolean

    Select Case GnTipoPinPad
        Case gPerifPINPAdUnisys  'UNISYS
            WriteToLcd = WriteToLcd_UNISYS("                                       ")
            WriteToLcd = WriteToLcd_UNISYS(szMessage)
        Case gPerifPINPAdVrf5000  'VERIFONE 5000
            WriteToLcd = WriteToLcd_Vrf5000(szMessage)
        Case gPerifPINPAdVrfSC552  'VERIFONE SC 552
            WriteToLcd = True  '
            WriteToLcd_VrfSC552 (szMessage)
        Case gPerifPINPAdHiperCom  'HIPERCOM
            'WriteToLcd = WriteToLcd_HiperCom(szMessage)
            
    End Select
End Function
Function WriteToLcd_VrfSC552(szMessage$) As Boolean

    Dim lbResult As Boolean
    Dim lalgo As Long
    Dim Mensaje As RegOut
    Mensaje.Trama = ""
    Mensaje.Trama = Trim(Left(szMessage, 16))
    lbResult = Not ShowMessage(Mensaje)
    WriteToLcd_VrfSC552 = lbResult

End Function

Function WriteToLcd_Vrf5000(szMessage$) As Boolean

    Dim lbResult As Boolean
    Dim lalgo As Long
    
    'lalgo = GmyPSerial.ReadCardIniConf(szMessage)
    
    If GmyPSerial.ReadCardIniConf(szMessage) = 1 Then
        lbResult = True
    Else
        lbResult = False
    End If
    
    
    WriteToLcd_Vrf5000 = lbResult
    
End Function

Function WriteToLcd_UNISYS(szMessage$) As Boolean
    Dim wRetVal&
    Dim Msg
    
    If ApiCheckIn(ID_PARENT_PP Or ID_CHILD_LCD) <> 0 Then
        Exit Function
    End If
    
    wRetVal = ERR_API_BUSY
    
    While wRetVal = ERR_API_BUSY
        wRetVal = MakeWord(PpLcdPutString(1, szMessage, Len(szMessage)))
        DoEvents  ' Lets "Dmonpro" do communication
    Wend
    
    If ApiCheckOut(ID_PARENT_PP Or ID_CHILD_LCD) Then
    End If
    
    WriteToLcd_UNISYS = True

End Function




'*****************************************************************************
'*
'* FUNCTION NAME:           McrReadPin
'*
'* DESCRIPTION:
'*
'* This function checks into corresponding device,
'* executes PpMcrRead function with PromptIndex set to 255(ie no prompt),
'* checks out of the device.
'* The function returns False if any error occurs, otherwise it returns True
'*
'*****************************************************************************
'
Function McrReadPin(Track1$, ByVal len1%, RetLen1%, Track2$, ByVal len2%, RetLen2%, Track3$, ByVal Len3%, RetLen3%) As Integer
            
Dim wRetVal&
Dim szTrack1 As String * MAX_CHAR_TRACK1  'Always allocate space for the buffer,
                                          'Don't pass variable length strings
                                          'to the API functions
                                          
Dim szTrack2 As String * MAX_CHAR_TRACK2
Dim szTrack3 As String * MAX_CHAR_TRACK3
    

'To execute PpMcrRead you need to check into PINPad MCR device
'As an alternative you can try to check into all devices you will
'at the start of the program in the "Sub Main" and check out
'at the end of the program

If ApiCheckIn(ID_PARENT_PP Or ID_CHILD_CON) <> 0 Then
    
   If gsErrLec = 0 Then
      MsgBox "CheckIn failed", , "PpMcrRead"
      Exit Function
   End If
    
End If

StopPpMcrRead = False
len1 = 10
wRetVal = ERR_API_BUSY
                      
'Call the function continuously while "ERR_API_BUSY" returned
'and the StoPpMcrRead flag is not raised

While wRetVal = ERR_API_BUSY And StopPpMcrRead <> True
    wRetVal = MakeWord(PpConGetData(0, szTrack1, len1, 42, 28, 13, RetLen1))
    DoEvents  ' Lets "Dmonpro" do communication
Wend

If ApiCheckOut(ID_PARENT_PP Or ID_CHILD_CON) <> 0 Then
End If

If StopPpMcrRead Then
    Exit Function
End If

If wRetVal = 0 Then

    Track1 = Left$(szTrack1, RetLen1) 'Api Don't return zero ended strings
                                      'use "Left" function to make string
                                      'zero ended
    'Track2 = Left$(szTrack2, RetLen2)
    'Track3 = Left$(szTrack3, RetLen3)
    McrReadPin = True
Else
    gsErrLec = wRetVal
'    MsgBox "Error code: " & Str$(wRetVal), , "PpMcrReadPin"
End If

End Function
Public Function GetNumTarjeta() As String

    Select Case GnTipoPinPad
        Case gPerifPINPAdUnisys  'UNISYS
            GetNumTarjeta = GetNumTarjeta_UNISYS
        Case gPerifPINPAdVrf5000  'VERIFONE 5000
            GetNumTarjeta = GetNumTarjeta_Vrf5000
        Case gPerifPINPAdVrfSC552  'VERIFONE SC 552
            GetNumTarjeta = GetNumTarjeta_VrfSC552
        Case gPerifPINPAdHiperCom  'HIPERCOM
            'GetNumTarjeta = GetNumTarjeta_HiperCom
    End Select

End Function
Public Function GetNumTarjeta_VrfSC552() As String

    Dim lsNumTarTemp As String
    Dim lsNumTar As String
    Dim routTarjeta As RegOut
    Dim i As Long
    If ReadCardIni() = 1 Then
        i = 0
        Do While i = 0
            routTarjeta.Trama = ""
            i = ReadCard(routTarjeta)
            If i = 1 Then
                lsNumTarTemp = routTarjeta.Trama
                Exit Do
            End If
            DoEvents
        Loop
        lsNumTar = Trim(lsNumTarTemp)
        GetNumTarjeta_VrfSC552 = lsNumTar
    End If

End Function
Public Function GetNumTarjeta_Vrf5000() As String

    Dim lsNumTarTemp As String
    Dim lsNumTar As String
    
    
    While lsNumTar = ""
        lsNumTar = GmyPSerial.ReadCard
        DoEvents
    Wend
    Debug.Print lsNumTar
    lsNumTarTemp = lsNumTar
    lsNumTar = Trim(Mid(lsNumTarTemp, 2, 16))
    If Not IsNumeric(lsNumTar) Then
      lsNumTar = Trim(Mid(lsNumTarTemp, 3, 16))
    End If
    
    GetNumTarjeta_Vrf5000 = lsNumTar
    
End Function

Public Function GetNumTarjeta_UNISYS() As String

    Dim Result&
    Dim sTarjeta As String
    Dim sCadena As String
    
    Dim lsGetNumTarjeta As String
    
    
    Result = McrRead(sTarjeta, 76, 0, sCadena, 0, 0, "", 0, 0)
    If Result <= 0 Then
       lsGetNumTarjeta = Format$(Mid$(sTarjeta, 2, 16), "0000-0000-0000-0000")
       lsGetNumTarjeta = Trim(Replace(lsGetNumTarjeta, "-", "", 1, , vbTextCompare))    'ppoa formateo trasladado
    Else
        MsgBox "Error : " & Trim(Result) & ". Consulte con Servicio Tecnico.", vbInformation, "Aviso"
        lsGetNumTarjeta = ""
    End If
    
    GetNumTarjeta_UNISYS = lsGetNumTarjeta
    
End Function


Public Function GetClaveTarjeta(Optional psCadena As String = "") As String

    Select Case GnTipoPinPad
        Case gPerifPINPAdUnisys  'UNISYS
            GetClaveTarjeta = GetClaveTarjeta_UNISYS
        Case gPerifPINPAdVrf5000  'VERIFONE 5000
            GetClaveTarjeta = GetClaveTarjeta_Vrf5000(psCadena)
        Case gPerifPINPAdVrfSC552  'VERIFONE SC 552
            GetClaveTarjeta = GetClaveTarjeta_VrfSC552("1234123412341234") 'psNumTarj)
        Case gPerifPINPAdHiperCom  'HIPERCOM
            'GetClaveTarjeta = GetClaveTarjeta_HiperCom
    End Select

End Function
Public Function GetClaveTarjeta_VrfSC552(ByVal psNumTarj As String) As String

    Dim lnClaveTar As String
    Dim iPin As Integer
    Dim lnI As Long
    Dim i As Long
    Dim routPinDes As RegOut
    Dim routMK As RegOut
    Dim routWK As RegOut
    Dim routTarj As RegOut
'    On Error Resume Next
    
    routMK.Trama = ""
    routWK.Trama = ""
    routTarj.Trama = ""
    routMK.Trama = "1234567890123456"
    routWK.Trama = "1234567890123456"
    routTarj.Trama = Trim(psNumTarj)
    
    i = ReadPinIni(routMK, routWK, routTarj)
    
    Dim icry As Long
    
    icry = 0
    Do While icry = 0
        icry = UnencryptedPIN
    Loop
    
    
    If i = 1 Then
        i = 0
        Do While True
            routPinDes.Trama = ""
            i = ReadPin(routPinDes)
            If i = 1 Then 'Or Len(Trim(routPinDes.Trama)) > 0 Then
                GetClaveTarjeta_VrfSC552 = Mid(routPinDes.Trama, 1, 4)
                Exit Do
            End If
            DoEvents
        Loop
    End If
    
End Function

Public Function GetClaveTarjeta_Vrf5000(ByVal lsMensaje As String) As String


    Dim lnClaveTar As String
    Dim iPin As Integer
    Dim lnI As Long
'    On Error Resume Next
    
    
    'ppoa Modificacion
    If Not WriteToLcd("Ingrese Clave") Then
        FinalizaPinPad
        MsgBox "No se Realizó Envío", vbInformation, "Aviso"
        Exit Function
    End If
    
    
    
    For lnI = 0 To 7000000
    Next lnI
    
    GmyPSerial.ReadPinIniConf "0", "1111111111111111", "0000000000000000", "INGRESE CLAVE " & lsMensaje
    lnClaveTar = GmyPSerial.ReadPinDes(iPin)
    
    While (lnClaveTar = "" And iPin <> 0)
      lnClaveTar = GmyPSerial.ReadPinDes(iPin)
    Wend
    
    If (lnClaveTar <> "" And iPin <> 0) Then
      GetClaveTarjeta_Vrf5000 = Mid(lnClaveTar, 1, 4)
    ElseIf iPin = 0 Then
         'TimerPinDes.Enabled = False
    End If


End Function


Public Function GetClaveTarjeta_UNISYS() As String

    Dim Result&
    Dim sClave As String
    
    
    'ppoa Modificacion
    If Not WriteToLcd("Ingrese Clave") Then
        FinalizaPinPad
        MsgBox "No se Realizó Envío", vbInformation, "Aviso"
        Exit Function
    End If
    
    
    
    Result = McrReadPin(sClave, 76, 0, "", 0, 0, "", 0, 0)
    If Result <= 0 Then
        GetClaveTarjeta_UNISYS = sClave
    Else
        MsgBox "Error : " & Trim(Result) & ". Consulte con Servicio Tecnico.", vbInformation, "Aviso"
        GetClaveTarjeta_UNISYS = ""
    End If
    
End Function
Private Function ObtieneTipoPinPad() As TipoPinPad
    
    Dim lsNombrePc As String
    Dim lnPeriferico   As TipoPeriferico
    
    Dim oConec As DConecta
    Dim lrst As ADODB.Recordset
    Dim lsql As String
    
    
    lsNombrePc = GetComputerName
    lnPeriferico = gPerifPINPAD
        
    
    Set oConec = New DConecta
    oConec.AbreConexion
    lsql = "SELECT nMarca FROM PERIFERICO where cPCNombre = '" & lsNombrePc & "' and  nPeriferico = " & lnPeriferico
    Set lrst = oConec.CargaRecordSet(lsql)
    
    If Not lrst.EOF() Then ObtieneTipoPinPad = lrst("nMarca")
    
End Function

Public Function IniciaPinPad(ByVal nPuerto As TipoPuertoSerial) As Boolean

    GnTipoPinPad = ObtieneTipoPinPad()

    Select Case GnTipoPinPad
        Case gPerifPINPAdUnisys   'UNISYS
            IniciaPinPad = IniciaPinPad_UNISYS(nPuerto)
        Case gPerifPINPAdVrf5000   'VERIFONE 5000
            IniciaPinPad = IniciaPinPad_Vrf5000(nPuerto)
        Case gPerifPINPAdVrfSC552   'VERIFONE SC 552
            IniciaPinPad = IniciaPinPad_VrfSC552(nPuerto)
        Case GnTipoPinPad = gPerifPINPAdHiperCom  'HIPERCOM
            'IniciaPinPad = IniciaPinPad_HiperCom(nPuerto)
    End Select

End Function
Public Function IniciaPinPad_VrfSC552(ByVal nPuerto As TipoPuertoSerial) As Boolean

    Disconnect
    Reset
    IniciaPinPad_VrfSC552 = CBool(Connect(CInt(nPuerto)))

End Function
Public Function IniciaPinPad_Vrf5000(ByVal nPuerto As TipoPuertoSerial) As Boolean

    Set GmyPSerial = CreateObject("HComPinpad.Pinpad")
    GmyPSerial.Reset

    'GmyPSerial.Connect nPuerto, 9600
    IniciaPinPad_Vrf5000 = CBool(GmyPSerial.Connect(nPuerto, 9600))
    
End Function

Public Function IniciaPinPad_UNISYS(ByVal nPuerto As TipoPuertoSerial) As Boolean
Dim Result As Long
Dim bExito As Boolean
bExito = False
Do While Not bExito
    Result = ApiOpen(Trim(nPuerto), 1, "", 0)  'Com1
    If Result <> 0 Then
        Select Case Result
            Case -4523 'Dmont no está corriendo
                Dim nRetVal As Long
                Dim X As Double
                ChDrive App.Path
                ChDir App.Path
                nRetVal = MakeWord(PpDvcSignal())
                If nRetVal = ERR_DMON_OFF Then
                    X = Shell(App.Path & "\Dmonnt.exe", vbMinimizedNoFocus)
                End If
            Case 61010
                If ApiCheckIn(ID_PARENT_PP Or ID_CHILD_MCR) <> 0 Then
                    MsgBox "Fallo ApiCheckIn, verificar estado del PINPAD", vbExclamation, "Aviso"
                    Exit Do
                End If
            Case Else
                MsgBox "Error al conectarse al PINPAD. Consulte con Servicio Tecnico.", vbInformation, "Aviso"
                Exit Do
        End Select
    Else
        bExito = True
    End If
Loop
IniciaPinPad_UNISYS = bExito
End Function
    
Public Function FinalizaPinPad()
    Select Case GnTipoPinPad
        Case gPerifPINPAdUnisys  'UNISYS
            FinalizaPinPad_UNISYS
        Case gPerifPINPAdVrf5000  'VERIFONE 5000
            FinalizaPinPad_Vrf5000
        Case gPerifPINPAdVrfSC552  'VERIFONE SC 552
            FinalizaPinPad_VrfSC552
        Case gPerifPINPAdHiperCom  'HIPERCOM
            'FinalizaPinPad_HiperCom
    End Select
End Function
Private Sub FinalizaPinPad_VrfSC552()

        'SetIdlePrompt
        WriteToLcd ("Gracias por su  Preferencia...")
        Disconnect
        Reset

End Sub

Private Sub FinalizaPinPad_UNISYS()
    WriteToLcd ("Gracias por su  Preferencia...")
    PpLcdClear
    PpConCancelEvent
    ApiClose
End Sub

Private Sub FinalizaPinPad_Vrf5000()
        WriteToLcd ("Gracias por su  Preferencia...")
        GmyPSerial.ReturnIdleState
        GmyPSerial.Disconnect
End Sub
