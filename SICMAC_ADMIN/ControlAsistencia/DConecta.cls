VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DConecta"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim DBConec As ADODB.Connection
Dim sCentralCom As String
Dim sSql As String
Dim prs  As ADODB.Recordset

' --------------------------------------------------------------
' Registry Specific Access Rights
' --------------------------------------------------------------
  Private Const KEY_QUERY_VALUE As Long = &H1
  Private Const KEY_SET_VALUE As Long = &H2
  Private Const KEY_CREATE_SUB_KEY As Long = &H4
  Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
  Private Const KEY_NOTIFY As Long = &H10
  Private Const KEY_CREATE_LINK As Long = &H20
  Private Const KEY_ALL_ACCESS As Long = &H3F
  Private Const HKEY_CURRENT_USER As Long = &O0
  Public ConexionActiva As ADODB.Connection


Public Function GetStringConnection(Optional psApp As String = "SICMACT") As String
'Dim sConex As String
'Dim sPath  As String
'Dim lngRootKey As Long
'sPath = "Software\" & psApp & "\Init\"
'lngRootKey = HKEY_CURRENT_USER
'
'sConex = regQuery_A_Key(lngRootKey, sPath, "Provider")
'If sConex <> "" Then
'   GetStringConnection = "PROVIDER=" & Encripta(sConex, False) & ";"
'   sConex = regQuery_A_Key(lngRootKey, sPath, "Server")
'   If sConex <> "" Then
'      GetStringConnection = GetStringConnection & "SERVER=" & Encripta(sConex, False) & ";"
'      sConex = regQuery_A_Key(lngRootKey, sPath, "DataBase")
'      If sConex <> "" Then
'         GetStringConnection = GetStringConnection & "DATABASE=" & Encripta(sConex, False) & ";"
'         sConex = regQuery_A_Key(lngRootKey, sPath, "Uid")
'         If sConex <> "" Then
'            GetStringConnection = GetStringConnection & "UID=" & Encripta(sConex, False) & ";"
'            sConex = regQuery_A_Key(lngRootKey, sPath, "Pwd")
'            If sConex <> "" Then
'               GetStringConnection = GetStringConnection & "PWD=" & Encripta(sConex, False) & ";"
'            End If
'         End If
'      End If
'   End If
'End If
End Function

Public Sub CierraConexion()
On Error Resume Next
If DBConec.State = adStateOpen Then
   DBConec.Close
   Set DBConec = Nothing
End If
Err.Clear
End Sub

Public Function AbreConexion(Optional psStringConec As String = "") As Boolean
On Error GoTo AbreConexionErr
AbreConexion = True
If DBConec Is Nothing Then
   Set DBConec = New ADODB.Connection
End If
If Not DBConec.State = adStateOpen Then
   If psStringConec = "" Then
      Dim oConec As clsIni.ClasIni
      Set oConec = New clsIni.ClasIni
      psStringConec = oConec.CadenaConexion()
      MsgBox psStringConec, vbInformation, "Aviso"
   End If
   DBConec.CursorLocation = adUseClient
   DBConec.Open psStringConec
End If
Set ConexionActiva = DBConec
Exit Function
AbreConexionErr:
   AbreConexion = False
   MsgBox Err.Description, vbInformation, "Aviso"
End Function

Public Function CargaRecordSet(sql As String, Optional LockType As LockTypeEnum = adLockReadOnly, Optional CursorType As CursorTypeEnum = adOpenStatic, Optional pbCierraTranasaccionActiva As Boolean = True) As ADODB.Recordset
   Dim prs As ADODB.Recordset
   On Error GoTo Error
   Set prs = New ADODB.Recordset
   If prs.State = adStateOpen Then prs.Close
   prs.CursorLocation = adUseClient
   DBConec.CommandTimeout = 120
   prs.Open sql, DBConec, CursorType, LockType, adCmdText
   If pbCierraTranasaccionActiva Then prs.ActiveConnection = Nothing
   Set CargaRecordSet = prs
   Set prs = Nothing
   Exit Function
Error:
   Set CargaRecordSet = Nothing
End Function

Public Function Ejecutar(psql As String) As ADODB.Recordset
On Error GoTo ErrEjecutar
Set Ejecutar = DBConec.Execute(psql)
Exit Function
ErrEjecutar:
   
End Function

Public Function AdicionaCmdBatch(psql As String)
If sSql = "" Then
   sSql = psql
Else
   sSql = sSql & "; " & psql
End If
End Function

Public Function EjecutarBatch() As Integer
On Error GoTo Ejecutar
Dim lTransac As Boolean
If sSql <> "" Then
   lTransac = True
   DBConec.BeginTrans
   DBConec.Execute sSql
   DBConec.CommitTrans
   sSql = ""
   lTransac = False
End If
Exit Function
Ejecutar:
   If lTransac Then
      DBConec.RollbackTrans
      lTransac = False
      sSql = ""
   End If
   ''Call RaiseError(MyUnhandledError, "DConecta:EjecutarBatch Method")
End Function

Public Sub BeginTrans()
On Error GoTo Ejecutar
   DBConec.BeginTrans
Exit Sub
Ejecutar:
   'Call RaiseError(MyUnhandledError, "DConecta:BeginTrans Method")
End Sub

Public Sub CommitTrans()
On Error GoTo Ejecutar
   DBConec.CommitTrans
Exit Sub
Ejecutar:
   'Call RaiseError(MyUnhandledError, "DConecta:CommitTrans Method")
End Sub
Public Sub RollbackTrans()
    On Error GoTo Ejecutar
       DBConec.RollbackTrans
    Exit Sub
Ejecutar:
       'Call RaiseError(MyUnhandledError, "DConecta:RollbackTrans Method")
End Sub

Public Function DatabaseName() As String
    DatabaseName = DBConec.DefaultDatabase
End Function

Public Function CadenaConexion() As String
    CadenaConexion = DBConec.ConnectionString
End Function

Public Function ServerName() As String
    Dim rs As ADODB.Recordset
    Dim sql As String
    ServerName = ""
    sql = "Select cNomSer = @@ServerName"
    Set rs = CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        ServerName = IIf(IsNull(rs!cNomSer), "", Trim(rs!cNomSer))
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Function GetHoraServer() As String
    Dim sql As String
    Dim rsH As New ADODB.Recordset
    sql = "SELECT convert(varchar(10),getdate(),108) as sHora"
    Set rsH = CargaRecordSet(sql)
    If Not rsH.EOF Then
       GetHoraServer = rsH!sHora
    Else
       GetHoraServer = Format(Time, "hh:mm:ss")
    End If
    rsH.Close: Set rsH = Nothing
End Function


Public Function GetFechaHoraServer() As Date
    Dim sql As String
    Dim rsH As New ADODB.Recordset
    sql = "Select GetDate() as dFecha"
    
    Set rsH = CargaRecordSet(sql)
    If Not rsH.EOF Then
       GetFechaHoraServer = rsH!dFecha
    Else
       GetFechaHoraServer = Now()
    End If
    rsH.Close
    Set rsH = Nothing
End Function

