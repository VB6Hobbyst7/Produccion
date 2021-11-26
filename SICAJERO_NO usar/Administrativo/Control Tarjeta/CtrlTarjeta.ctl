VERSION 5.00
Begin VB.UserControl CtrlTarjeta 
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2745
   LockControls    =   -1  'True
   ScaleHeight     =   1620
   ScaleWidth      =   2745
   Begin VB.Timer TimerCard 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimerPin 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   465
      Top             =   0
   End
   Begin VB.Timer TimerPinDes 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer TimerPinDes2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1410
      Top             =   0
   End
End
Attribute VB_Name = "CtrlTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public myPSerial As Object
Dim SCard As String
Dim sPin As String

'****************** PARA PINPAD ACS
'**********************************************************************

Dim x As Long
Private Declare Function pinverify _
Lib "PINVerify.dll" _
             (ByVal ippuerto As String, _
              ByVal Key As String, _
              ByVal PAN As String, _
              ByVal PVKI As String, _
              ByVal PIN As String, _
              ByVal pvv As String _
             ) As Integer


Private Declare Function genpckey Lib "PINVerify.dll" () As Long

Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Declare Function OpenReader Lib "PinPad.dll" () As Integer
Private Declare Sub CloseReader Lib "PinPad.dll" ()
Private Declare Function ReadMagneticCard Lib "PinPad.dll" (ByVal min As Integer) As Integer
Private Declare Sub ClearDisplay Lib "PinPad.dll" ()
Private Declare Sub LightOn Lib "PinPad.dll" (ByVal min As Integer)
Private Declare Sub BuzzerOn Lib "PinPad.dll" (ByVal min As Integer)
Private Declare Sub DisplayString Lib "PinPad.dll" (ByVal line As Integer, ByVal offset As Integer, ByVal tipo As Integer, ByVal msg As String)
Private Declare Function GetPAN Lib "PinPad.dll" (ByVal timeout As Integer) As Long
Private Declare Function CheckPIN _
Lib "PinPad.dll" _
             (ByVal timeout As Integer, _
              ByVal ippuerto As String, _
              ByVal PAN As String, _
              ByVal PVKI As String, _
              ByVal pvv As String _
             ) As Integer
             
Private Declare Function ChangePIN _
Lib "PinPad.dll" _
             (ByVal timeout As Integer, _
              ByVal ippuerto As String, _
              ByVal PAN As String, _
              ByVal PVKI As String, _
              ByVal confirm As Integer, _
              ByVal mensaje As String _
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
             
Private Function DevuelveParametro(ByVal x As Long) As String
Dim Buffer() As Byte
Dim i, nLen As Long
Dim res As String
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

'************************** FIN DE CODIGO PARA PINPAD ACS


Public Function LeerTarjeta(ByVal psMensaje As String, Optional ByVal pnTipoPinPad As Integer = 1, Optional ByVal pnPuerto As Integer = 1) As String
Dim M As HCOMPINPADLib.Pinpad
Dim sResp As String
  
On Error GoTo Error1
  
    Select Case pnTipoPinPad
        Case 1
  
            Set myPSerial = CreateObject("HComPinpad.Pinpad")
            sResp = conectar(pnPuerto)
            If sResp <> "" Then
                LeerTarjeta = sResp
                Exit Function
            End If
            
            If myPSerial.ReadCardIniConf(psMensaje) = 1 Then
              SCard = ""
              Do While Len(SCard) = 0
                  SCard = myPSerial.ReadCard
                  If (SCard <> "") Then
                   Exit Do
                  End If
                  DoEvents
              Loop
            End If
            Call Desconectar
                
        Case 2
  
            '************ CONECTAR **********************
            x = OpenReader
            If x > 0 Then
                LeerTarjeta = ""
                Else
                LeerTarjeta = "No hay conexión"
            End If
            
            '******* PARA PINPAD ACS ********************************************************
            LightOn (1)
            ClearDisplay
            DisplayString 1, 8, 1, psMensaje
            x = GetPAN(100)
            SCard = DevuelveParametro(x)
            SCard = " " & SCard
            ClearDisplay
            LightOn (0)
            '**********************************************************************
            
            CloseReader
  
  End Select
  
  LeerTarjeta = SCard
  
  Exit Function
Error1:
        MsgBox ("Error en el Pinpad, Avise al Area de Sistemas")
  
End Function

Public Function PedirPinDes(ByVal sNroTarjeta As String, ByVal NMK As String, ByVal WK As String, Optional ByVal pnTipoPinPad As Integer = 1, Optional ByVal pnPuerto As Integer = 1) As String
Dim sResp As String
  
  sResp = conectar(pnPuerto)
  If sResp <> "" Then
        PedirPinDes = sResp
        Exit Function
  End If
  
        If myPSerial.ReadPinIni(NMK, WK, sNroTarjeta) = 1 Then
        'If myPSerial.ReadPinIni("0", "1111111111111111", "0000000000000000") = 1 Then
              Dim iPin As Integer
              sPin = ""
              Do While Len(sPin) = 0
                sPin = myPSerial.ReadPinDes(iPin)
                If (sPin <> "" And iPin <> 0) Then
                    sPin = Mid(sPin, 1, 4)
                    Exit Do
                  ElseIf iPin = 0 Then
                    Exit Do
                End If
            Loop
        End If
      Call Desconectar
        PedirPinDes = sPin

End Function

Public Function PedirPinEnc(ByVal sNroTarjeta As String, ByVal NMK As String, ByVal WK As String, Optional ByVal pnTipoPinPad As Integer = 1, Optional ByVal pnPuerto As Integer = 1) As String
Dim sResp As String
Dim iPin As Integer
  
    Select Case pnTipoPinPad
        Case 1
            sResp = conectar(pnPuerto)
            If sResp <> "" Then
                  PedirPinEnc = sResp
                  Exit Function
            End If
            
            If myPSerial.ReadPinIni(NMK, WK, sNroTarjeta) = 1 Then
                sPin = ""
                Do While Len(sPin) = 0
                  sPin = myPSerial.ReadPin(iPin)
                  If (sPin <> "" And iPin <> 0) Then
                      Exit Do
                    ElseIf iPin = 0 Then
                      Exit Do
                  End If
                Loop
            End If
            Call Desconectar
        Case 2
    End Select
    
    PedirPinEnc = sPin
        
End Function

'**DAOR 20091216, Obtiene el PINBlock para operaciones intercajas
Public Function PITPedirPinEnc(ByVal psPAN As String, ByVal psNMK As String, ByVal psWK As String, ByVal psIpPuertoPinVerify As String, ByVal psCanalIdPOS As String, ByVal psCanalIdATM, Optional ByVal pnTipoPinPad As Integer = 1, Optional ByVal pnPuerto As Integer = 1) As String
Dim sResp As String, lsNewPINBlock As String
Dim iPin As Integer
Dim lnTimeOut As Integer, i As Integer
Dim lsPvki As String, lsKey As String

    lnTimeOut = 100
    lsPvki = "1"
    lsKey = "444"
  
    Select Case pnTipoPinPad
        Case 1
            sResp = conectar(pnPuerto)
            If sResp <> "" Then
                  PITPedirPinEnc = sResp
                  Exit Function
            End If
            
            If myPSerial.ReadPinIni(psNMK, psWK, psPAN) = 1 Then
                sPin = ""
                Do While Len(sPin) = 0
                  sPin = myPSerial.ReadPin(iPin)
                  If (sPin <> "" And iPin <> 0) Then
                      Exit Do
                    ElseIf iPin = 0 Then
                      Exit Do
                  End If
                Loop
            End If
            Call Desconectar
            
        Case 2
        
    End Select
    
   
    x = pinblock(psIpPuertoPinVerify, lsKey, psPAN, lsPvki, "t" & sPin & psCanalIdPOS, psCanalIdATM)
    lsNewPINBlock = DevuelveParametro(x)
    
    
    PITPedirPinEnc = lsNewPINBlock
        
End Function


Private Function conectar(Optional ByVal pnPuerto As Integer = 1) As String
 If myPSerial.Connect(pnPuerto, 19200) Then
        conectar = ""
    Else
        conectar = "No hay conexión"
    End If
End Function

Private Sub Desconectar()
    myPSerial.Disconnect
End Sub

Private Sub TimerCard_Timer()
  SCard = myPSerial.ReadCard
    If (SCard <> "") Then
     TimerCard.Enabled = False
    End If
End Sub

Private Sub TimerPinDes2_Timer()
'  Dim sPin As String
'  Dim iPin As Integer
'  sPin = myPSerial.ReadPinDes(iPin)
'  If (sPin <> "" And iPin <> 0) Then
'      txtPin.Text = sPin
'      TimerPinDes2.Enabled = False
'    ElseIf iPin = 0 Then
'      TimerPinDes2.Enabled = False
'  End If
End Sub

                

'**DAOR 20081126, ************************************************************
Public Function PedirPinYValida(ByVal sNroTarjeta As String, ByVal NMK As String, ByVal WK As String, ByVal psIpPuertoPinVerify As String, ByVal psPvv As String, ByVal psCanalIdPOS As String, Optional ByVal pnTipoPinPad As Integer = 1, Optional ByVal pnPuerto As Integer = 1, Optional ByRef psPINBlock As String = "") As Integer
Dim sResp As String
Dim lnTimeOut As Integer, i As Integer
Dim lsPvki As String, lsKey As String
    lnTimeOut = 100
    lsPvki = "1"
    lsKey = "444"
    PedirPinYValida = 0
    Select Case pnTipoPinPad
    Case 1
        sResp = conectar(pnPuerto)
        If sResp <> "" Then
              PedirPinYValida = -2 'No hay conexion
              Exit Function
        End If
    
        If myPSerial.ReadPinIni(NMK, WK, sNroTarjeta) = 1 Then
            Dim iPin As Integer
            sPin = ""
            Do While Len(sPin) = 0
                sPin = myPSerial.ReadPin(iPin)
                If (sPin <> "" And iPin <> 0) Then
                    Exit Do
                ElseIf iPin = 0 Then
                    Exit Do
                End If
            Loop
        End If
        psPINBlock = sPin
        Call Desconectar
                                   
        i = pinverify(psIpPuertoPinVerify, lsKey, sNroTarjeta, lsPvki, "t" & sPin & psCanalIdPOS, psPvv)
        PedirPinYValida = i
    
    Case 2
        x = OpenReader
        If x <= 0 Then
            PedirPinYValida = -2 'No hay conexion
            Exit Function
        End If
  
        LightOn (1)
        ClearDisplay
        DisplayString 1, 8, 1, "INGRESE SU CLAVE"
        i = CheckPIN(lnTimeOut, psIpPuertoPinVerify, sNroTarjeta, lsPvki, psPvv)
        ClearDisplay
        LightOn (0)
        PedirPinYValida = i
            
        Call CloseReader
    End Select
            
End Function

Public Function PedirPinDevPvvACS(ByVal sNroTarjeta As String, ByVal psIpPuertoPinVerify As String, Optional ByVal pnTipoPinPad As Integer = 1, Optional ByVal pnPuerto As Integer = 1) As String
Dim lnTimeOut As Integer, i As Integer
Dim lsPvki As String, lsKey As String
    lnTimeOut = 100
    lsPvki = "1"
    lsKey = "444"
  
    x = OpenReader
    If x <= 0 Then
        PedirPinDevPvvACS = "ERR2" 'No hay conexion
        Exit Function
    End If
  
    LightOn (1)
    ClearDisplay
    DisplayString 1, 8, 1, "INGRESE SU CLAVE"
    x = ChangePIN(lnTimeOut, psIpPuertoPinVerify, sNroTarjeta, lsPvki, 0, "INGRESE SU CLAVE")
    PedirPinDevPvvACS = DevuelveParametro(x)
    ClearDisplay
    LightOn (0)
        
    Call CloseReader
End Function

'*****************************************************************
