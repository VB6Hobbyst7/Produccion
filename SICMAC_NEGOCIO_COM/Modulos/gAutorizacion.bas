Attribute VB_Name = "gAutorizacion"
''Option Explicit
''
'''p.cperscod,p.cpersnombre,p.dpersnaccreac,c.cnivel,rhct.crhcargodescripcion,rhc.crhcargocod
''
''Public GAutcperscod  As String
''Public GAutcpersnombre As String
''Public GAutdpersnaccreac As Date
''Public GAutcnivel As String
''Public GAutcrhcargodescripcion As String
''Public GAutcrhcargocod As String
''
''Public GAutCodOP As String
''Public GAutDescOP As String
'
'
'Private Const OPEN_EXISTING = 3
'Private Const GENERIC_WRITE = &H40000000
'Private Const FILE_SHARE_READ = &H1
'Private Const FILE_ATTRIBUTE_NORMAL = &H80
'Private Const INVALID_HANDLE_VALUE = -1
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hHandle As Long) As Long
'Private Declare Function WriteFile Lib "kernel32" (ByVal hFileName As Long, ByVal lpBuff As Any, ByVal nNrBytesToWrite As Long, lpNrOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
'Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwAccess As Long, ByVal dwShare As Long, ByVal lpSecurityAttrib As Long, ByVal dwCreationDisp As Long, ByVal dwAttributes As Long, ByVal hTemplateFile As Long) As Long
'Public Function EnviarMensajeA(ByVal Emisor As String, ByVal Receptor As String, ByVal Mensaje As String)
'Dim lngH As Long, strTextoAEnviar As String, lngResult As Long
'   strTextoAEnviar = Emisor & Chr(0) & Receptor & Chr(0) & Mensaje & Chr(0)
'   lngH = CreateFile("\\" & Receptor & "\mailslot\messngr", GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
'   If lngH = INVALID_HANDLE_VALUE Then
'      EnviarMensaje = 0
'   Else
'      If WriteFile(lngH, strTextoAEnviar, Len(strTextoAEnviar), lngResult, 0) = 0 Then
'         EnviarMensaje = 0
'      Else
'         EnviarMensaje = lngResult
'      End If
'      CloseHandle lngH
'   End If
'End Function
'
