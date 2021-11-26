Attribute VB_Name = "gApisIni"
Option Explicit
'Declare Function SetWindowWord Lib "User" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal wNewWord As Integer) As Integer
Declare Function SetWindowWord Lib "User32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
'Crear una ventana flotante al estilo de los tool-bar
'Cuando se minimiza la ventana padre, también lo hace ésta.
Public Const SWW_hParent = -8
'APIS DE LECTURA Y ESCRITURA DE ARCHIVOS INI
Declare Function GetPrivateProfileStringByKeyName& Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function GetPrivateProfileStringKeys& Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function GetPrivateProfileStringSections& Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName&, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function WritePrivateProfileStringByKeyName& Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)
Declare Function WritePrivateProfileStringToDeleteKey& Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Long, ByVal lplFileName As String)

Public Const BUFF_SIZ As Long = 9160
Public Const READ_BUFF As Long = 255
Public Function LeerArchivoIni(Seccion As String, Dato As String, lsArchivo As String) As String
    Dim strValue As String
    Dim lngRetLen As Long
    
    strValue = String(READ_BUFF + 1, Space(1))
    lngRetLen = GetPrivateProfileStringByKeyName(Seccion, Dato, "", strValue, READ_BUFF, lsArchivo)
    If lngRetLen > 1 Then
        LeerArchivoIni = Left(strValue, lngRetLen)
    Else
        LeerArchivoIni = ""
    End If
End Function
Public Function EscribirIni(strFileSection As String, strKey As String, strValue As String, sFile As String) As Long

    If Len(strKey) > READ_BUFF Or Len(strValue) > READ_BUFF Then
        MsgBox "No puede escribir Cadena mas grande de " & READ_BUFF & " caracteres para el valor o la clave."
        EscribirIni = -1
        Exit Function
    End If
    EscribirIni = WritePrivateProfileStringByKeyName(strFileSection, strKey, strValue, sFile)
End Function

