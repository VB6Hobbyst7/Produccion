Attribute VB_Name = "gFunciones"
Option Explicit
Global gsErrLec As Long
Global oCon As DConecta

Global Const gsCodAge = "11207"
Global Const gsCodUsu = "AUTO"
Global Const gnHoraFinal = 20
Global Const gnHoraMenor = 19
Global VSQL As String
Global gsPASS As String
Public gsPuerto As String

Public Sub IniLectora()
    Dim Result As Long
'    Result = ApiOpen(gsPuerto, 1, "", 0)  'Com1
'    If Result <> 0 Then
'        MsgBox GetErrorPINPAD(Result) & " Consulte con Servicio Tecnico.", vbInformation, "Aviso"
'    End If
End Sub

Public Function GetClaveTarjeta() As String
    Dim Result As Long
    Dim lsClave As String
    Result = McrReadPin(lsClave, 76, 0, "", 0, 0, "", 0, 0)
    If Result <= 0 Then
       GetClaveTarjeta = lsClave
    Else
        MsgBox GetErrorPINPAD(Result) & " Consulte con Servicio Tecnico.", vbInformation, "Aviso"
        GetClaveTarjeta = ""
    End If
End Function

Public Sub Main()
    Dim X As Double
    Dim lnRetVal As Long
    ChDrive App.Path
    ChDir App.Path
    'lnRetVal = MakeWord(PpDvcSignal())
    'If lnRetVal = ERR_DMON_OFF Then
    '    X = Shell(App.Path & "\Dmonnt.exe", vbMinimizedNoFocus)
    'End If
    frmAsistencia.Show 1
End Sub

Public Sub FinLectora()
Dim i As Integer
    i = Reset()
    i = Disconnect()
'    PpLcdClear
'    PpConCancelEvent
'    ApiClose
End Sub

Public Function IniAsistencia() As String
    Dim Result As Long
    Dim lsClave As String
    Result = McrReadPin(lsClave, 76, 0, "", 0, 0, "", 0, 0)
    If Result <= 0 Then
       IniAsistencia = lsClave
    Else
        MsgBox GetErrorPINPAD(Result) & " Consulte con Servicio Tecnico.", vbInformation, "Aviso"
        IniAsistencia = "X"
    End If
End Function

Public Function GetErrorPINPAD(pnNumber As Long) As String
    GetErrorPINPAD = "Error de PINPAN, Verifique si el Programa DMONNT.EXE esta en la Barra de Tareas. Verifique la Conexión del PINPAD."
End Function

Public Sub OpenDB()
End Sub
    
Public Sub CloseDB()
'    dbCmact.Close
'    Set dbCmact = Nothing
End Sub
