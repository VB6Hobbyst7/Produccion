VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DConecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit
Dim DBConec As ADODB.Connection
Dim sCentralCom As String
Dim sSql As String
Dim pRs  As ADODB.Recordset

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
      Dim oConec As COMConecta.DCOMClasIni
      Set oConec = New COMConecta.DCOMClasIni
      psStringConec = oConec.CadenaConexion()
   End If
   DBConec.CursorLocation = adUseClient
   DBConec.CommandTimeout = 1200 'ARCV 22-06-2007
   DBConec.Open psStringConec
End If
DBConec.Execute "SET DATEFORMAT mdy"
Set ConexionActiva = DBConec
Exit Function
AbreConexionErr:
   AbreConexion = False
   Call RaiseError(MyUnhandledError, "DConecta:AbreConexion Method")
End Function

Public Property Get CommadTimeOut() As Long
'    dbConec.CommandTimeout = pnTiempodeEspera
End Property

Public Function CargaRecordSet(sql As String, Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
Dim pRs As ADODB.Recordset
   Set pRs = New ADODB.Recordset
   If pRs.State = adStateOpen Then pRs.Close
   pRs.CursorLocation = adUseClient
   'By Capi 20071106 sugerido por Danny
   'DBConec.CommandTimeout = 120
   DBConec.CommandTimeout = 1200
   pRs.Open sql, DBConec, adOpenStatic, LockType, adCmdText
   pRs.ActiveConnection = Nothing
   Set CargaRecordSet = pRs
   Set pRs = Nothing
End Function

Public Function Ejecutar(pSQL As String) As ADODB.Recordset
On Error GoTo ErrEjecutar
Set Ejecutar = DBConec.Execute(pSQL)
Exit Function

ErrEjecutar:
   Call RaiseError(MyUnhandledError, "DConecta:Ejecutar Method")
End Function

Public Function AdicionaCmdBatch(pSQL As String)
If sSql = "" Then
   sSql = pSQL
Else
   sSql = sSql & "; " & pSQL
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
   Call RaiseError(MyUnhandledError, "DConecta:EjecutarBatch Method")
End Function

Public Sub BeginTrans()
On Error GoTo Ejecutar
   DBConec.BeginTrans
Exit Sub
Ejecutar:
   Call RaiseError(MyUnhandledError, "DConecta:BeginTrans Method")
End Sub

Public Sub CommitTrans()
On Error GoTo Ejecutar
   DBConec.CommitTrans
Exit Sub
Ejecutar:
   Call RaiseError(MyUnhandledError, "DConecta:CommitTrans Method")
End Sub
Public Sub RollbackTrans()
On Error GoTo Ejecutar
   DBConec.RollbackTrans
Exit Sub
Ejecutar:
   Call RaiseError(MyUnhandledError, "DConecta:RollbackTrans Method")
End Sub

Public Function DatabaseName() As String
    DatabaseName = DBConec.DefaultDatabase
End Function

Public Function CadenaConexion() As String
    CadenaConexion = DBConec.ConnectionString
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



'****************************************
'AbreConexionRemota: Función que devvuelve True si la se realiza la nueva conexión y false
'si la nueva conexion falla, se ingresa el codigo de agencia y False
'o los dos numeros iniciales del codigo de agencia y true
'****************************************
Public Function AbreConexionRemota(psAgecod As String, Optional pbMensaje As Boolean = True, Optional pbConIniCie As Boolean = False, Optional psNroServidor As String = "01") As Boolean
    Dim lsConexionCad As String
    Dim ldFecha  As Date
    Dim ldFecSis As Date
    Dim lsAgeNom As String
    
    On Error GoTo AbreConexionRemotaErr
    AbreConexionRemota = True
    lsConexionCad = GetCadenaConexion(psAgecod, psNroServidor)
    If lsConexionCad = "" Then
       AbreConexionRemota = False
       Exit Function
    End If
    
   If DBConec Is Nothing Then
      Set DBConec = New ADODB.Connection
   End If
   If DBConec.State = adStateOpen Then
      DBConec.Close
   End If
   DBConec.CommandTimeout = 3000
   DBConec.CursorLocation = adUseClient
   DBConec.Open lsConexionCad
   DBConec.Execute "Set Dateformat mdy"
   AbreConexionRemota = True
   
   If Not pbConIniCie Then
      Exit Function
   End If
    
   Dim oConst As New NConstSistemas
   Dim oArea  As New DActualizaDatosArea
   ldFecha = CDate(oConst.LeeConstSistema(gConstSistFechaSistema))
   lsAgeNom = oArea.GetNombreAgencia(psAgecod)
   If DateDiff("d", ldFecha, ldFecSis) = 0 Then
      If DateDiff("d", CDate(oConst.LeeConstSistema(gConstSistCierreDiaCapta)), ldFecSis) = 0 Or DateDiff("d", CDate(oConst.LeeConstSistema(gConstSistCierreDiaColoc)), ldFecSis) = 0 Or DateDiff("d", CDate(oConst.LeeConstSistema(gConstSistCierreDiaColocPrend)), ldFecSis) = 0 Then
         AbreConexionRemota = False
         CierraConexion
         Err.Raise 50001, "AbreConexionRemota", "La Agencia: " & lsAgeNom & " ya realizo el cierre de día"
      Else
         AbreConexionRemota = True
      End If
   Else
      AbreConexionRemota = False
      CierraConexion
      Err.Raise 50001, "AbreConexionRemota", "La Agecia: " & lsAgeNom & " todavia no ha hecho Inico de día, o tiene una Fecha diferente a la agencia  local. La Fecha es :  " & ldFecha & " "
   End If
   Set oArea = Nothing
   Set oConst = Nothing
   Exit Function
AbreConexionRemotaErr:
    If pbMensaje Then
       Err.Raise 50001, "AbreConexionRemota", "El Servidor De la Agencia Nº : " & Mid(psAgecod, 1, 2) & Space(2) & lsAgeNom & " no responde." & Chr(13) & "Consulte al Departamento de Sistemas"
    End If
    AbreConexionRemota = False
End Function

'****************************************************
'GetCadenaConexion:Funcion que devuelve la cadena de conección de la
'agencia segun el codigo que se ingresa
'****************************************************
Public Function GetCadenaConexion(psAgecod As String, psNroServidor As String) As String
   Dim pRs  As ADODB.Recordset
   Dim sSql As String
    
   On Error GoTo GetCadenaConexionErr
   GetCadenaConexion = ""
   AbreConexion
   sSql = "SELECT * FROM Servidor " _
        & "WHERE  substring(cCodAge,4,2) = '" + Mid(psAgecod, 1, 2) + "' AND cEstado = 'A' And cNroSer = '" & psNroServidor & "'"
   Set pRs = CargaRecordSet(sSql)
   If pRs.EOF And pRs.BOF Then
      Err.Raise 50001, "DConecta", "No se Configuró acceso con Agencia " & psAgecod & ". Consultar con Sistemas"
      Exit Function
   End If
   GetCadenaConexion = "PROVIDER=" + Trim(pRs!cProvider) + ";uid=" + Trim(pRs!cLogin) + ";pwd=" + Trim(pRs!cPass) + ";DATABASE=" + Trim(pRs!cDataBase) + ";SERVER=" + Trim(pRs!cNomSer)
   RSClose pRs
   CierraConexion
   Exit Function
GetCadenaConexionErr:
    Err.Raise 50001, "DConecta", "Error de Conexión " + Err.Description
End Function


