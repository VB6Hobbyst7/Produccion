VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMConecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Dim dbConec As ADODB.Connection
Dim sCentralCom As String
Dim sSQL As String
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
  
Public Sub CierraConexion()
'On Error Resume Next
    If Not (dbConec Is Nothing) Then
        If dbConec.State = adStateOpen Then
            dbConec.Close
            Set dbConec = Nothing
        End If
    End If
Err.Clear
End Sub

Public Function AbreConexion(Optional psStringConec As String = "") As Boolean
'On Error GoTo AbreConexionErr
AbreConexion = True
If dbConec Is Nothing Then
   Set dbConec = New ADODB.Connection
End If
If Not dbConec.State = adStateOpen Then
   If psStringConec = "" Then
      Dim oConec As DCOMClasIni
      Set oConec = New DCOMClasIni
      
      psStringConec = oConec.CadenaConexion()
   End If
   dbConec.CursorLocation = adUseClient
   dbConec.Open psStringConec
End If
dbConec.Execute "SET DATEFORMAT mdy"
'Set ConexionActiva = DBConec
'Exit Function
'AbreConexionErr:
'   AbreConexion = False
'   Call RaiseError(MyUnhandledError, "DConecta:AbreConexion Method")
End Function

Public Function CargaRecordSet(sql As String, Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
Dim prs As ADODB.Recordset
   Set prs = New ADODB.Recordset
   If prs.State = adStateOpen Then prs.Close
   prs.CursorLocation = adUseClient
   dbConec.CommandTimeout = 8000
   prs.Open sql, dbConec, adOpenStatic, LockType, adCmdText
   prs.ActiveConnection = Nothing
   Set CargaRecordSet = prs
   Set prs = Nothing
End Function

Public Function Ejecutar(psql As String) As ADODB.Recordset
On Error GoTo ErrEjecutar
Set Ejecutar = dbConec.Execute(psql)
Exit Function

ErrEjecutar:
   Call RaiseError(MyUnhandledError, "DConecta:Ejecutar Method")
End Function

Public Function AdicionaCmdBatch(psql As String)
If sSQL = "" Then
   sSQL = psql
Else
   sSQL = sSQL & "; " & psql
End If
End Function

Public Function EjecutarBatch() As Integer
On Error GoTo Ejecutar
Dim lTransac As Boolean
If sSQL <> "" Then
   lTransac = True
   dbConec.BeginTrans
   dbConec.Execute sSQL
   dbConec.CommitTrans
   sSQL = ""
   lTransac = False
End If
Exit Function
Ejecutar:
   If lTransac Then
      dbConec.RollbackTrans
      lTransac = False
      sSQL = ""
   End If
   Call RaiseError(MyUnhandledError, "DConecta:EjecutarBatch Method")
End Function

Public Sub BeginTrans()
On Error GoTo Ejecutar
   dbConec.BeginTrans
Exit Sub
Ejecutar:
   Call RaiseError(MyUnhandledError, "DConecta:BeginTrans Method")
End Sub

Public Sub CommitTrans()
On Error GoTo Ejecutar
   dbConec.CommitTrans
Exit Sub
Ejecutar:
   Call RaiseError(MyUnhandledError, "DConecta:CommitTrans Method")
End Sub
Public Sub RollbackTrans()
On Error GoTo Ejecutar
   dbConec.RollbackTrans
Exit Sub
Ejecutar:
   Call RaiseError(MyUnhandledError, "DConecta:RollbackTrans Method")
End Sub

Public Function DatabaseName() As String
    DatabaseName = dbConec.DefaultDatabase
End Function

Public Function CadenaConexion() As String
    CadenaConexion = dbConec.ConnectionString
End Function

Public Function servername() As String
Dim rs As ADODB.Recordset
Dim sql As String
servername = ""
sql = "Select cNomSer = @@ServerName"
Set rs = CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    servername = IIf(IsNull(rs!cNomSer), "", Trim(rs!cNomSer))
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
   RSClose rsH
End Function

Public Property Get ConexionActiva() As ADODB.Connection
    Set ConexionActiva = dbConec
End Property

Public Property Let ConexionActiva(ByVal vNewValue As ADODB.Connection)
    Set dbConec = vNewValue
End Property


