VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPinpad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CREADO POR JHCU ENCUESTAS PINPADS DEL PROYECTO OCXTARJETA
'20-11-2019 17:25 PM

Option Explicit
Public myPSerial As Object
Dim sPin As String 'added 20191202

Private Function conectar(Optional ByVal pnPuerto As Integer = 1) As String
 If myPSerial.Connect(pnPuerto, 19200) Then
        myPSerial.TO_READ = CInt(600) 'add jhcu 18-11-2019
        myPSerial.TO_TIMEOUT = CInt(6000) 'add jhcu 18-11-2019
        conectar = ""
    Else
        conectar = "No hay conexión"
    End If
End Function
'ADD JHCU 18-11-2019
Public Function PedirTecla(Optional ByRef psMensaje As String = "Presione una tecla", Optional ByRef psPuerto As Integer = 9) As String
Dim strReadKey As String
Dim rc As Integer
Dim sResp As String
Dim nTimeoutLeerTecla As Integer
Set myPSerial = CreateObject("HComPinpad.Pinpad")
sResp = conectar(psPuerto)
If sResp <> "" Then
    PedirTecla = -2 'NO HAY CONEXIÓN
    Exit Function
End If
myPSerial.SetIdlePrompt (psMensaje)
nTimeoutLeerTecla = CInt(60)
rc = myPSerial.Reset
strReadKey = myPSerial.ReadKey(nTimeoutLeerTecla)
myPSerial.ReturnIdleState
PedirTecla = strReadKey
Call Desconectar
End Function
Private Sub Desconectar()
    myPSerial.Disconnect
End Sub
Public Function PedirPinEncNDig(ByVal sNroTarjeta As String, ByVal NMK As String, ByVal WK As String, Optional ByVal pnTipoPinPad As Integer = 1, Optional ByVal pnPuerto As Integer = 1, Optional ByVal pnDigtos = 4) As String
Dim sResp As String
Dim iPin As Integer
    
    If myPSerial Is Nothing Then
        Set myPSerial = CreateObject("HComPinpad.Pinpad")
    End If
    
    Select Case pnTipoPinPad
        Case 1
            sResp = conectar(pnPuerto)
            If sResp <> "" Then
                  'PedirPinEncNDig = sResp
                  PedirPinEncNDig = ""
                  Exit Function
            End If
            
            If myPSerial.ReadPinVarIni(NMK, WK, sNroTarjeta, pnDigtos, pnDigtos) = 1 Then
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
    
    PedirPinEncNDig = sPin
        
End Function
