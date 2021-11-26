VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DConecta"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
If dbConec Is Nothing Then
    Exit Sub
End If
If dbConec.State = adStateOpen Then
   dbConec.Close
   Set dbConec = Nothing
End If
Err.Clear
End Sub

Public Function AbreConexion(Optional psStringConec As String = "") As Boolean
On Error GoTo AbreConexionErr
AbreConexion = True
If dbConec Is Nothing Then
   Set dbConec = New ADODB.Connection
End If
If dbConec.State = adStateOpen Then
    dbConec.Close
End If
   If psStringConec = "" Then
      Dim oConec As ClsIni.ClasIni
      Set oConec = New ClsIni.ClasIni
      psStringConec = oConec.CadenaConexion(App.Path & "\SicmactF.Ini")
   End If
   dbConec.CursorLocation = adUseClient
   dbConec.ConnectionTimeout = 3600
   dbConec.Open psStringConec
   dbConec.Execute "set dateformat mdy"
Exit Function
AbreConexionErr:
   AbreConexion = False
   Call RaiseError(MyUnhandledError, "DConecta:AbreConexion Method")
End Function


Public Function Ejecutar(pSql As String) As ADODB.Recordset
On Error GoTo ErrEjecutar
dbConec.CommandTimeOut = 2000
Set Ejecutar = dbConec.Execute(pSql)
Exit Function
ErrEjecutar:
   Call RaiseError(MyUnhandledError, "DConecta:Ejecutar Method")
End Function

Public Function AdicionaCmdBatch(pSql As String)
If sSQL = "" Then
   sSQL = pSql
Else
   sSQL = sSQL & "; " & pSql
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

'**********************
'** NUEVAS FUNCIONES
'**********************

Public Function CargaRecordSet(sql As String, Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
Dim prs As ADODB.Recordset
   Set prs = New ADODB.Recordset
   If prs.State = adStateOpen Then prs.Close
   prs.CursorLocation = adUseClient
   dbConec.CommandTimeOut = 3000
   prs.Open sql, dbConec, adOpenStatic, LockType, adCmdText
   prs.ActiveConnection = Nothing
   Set CargaRecordSet = prs
   Set prs = Nothing
End Function

Public Sub CommandTimeOut(pnValor)
   dbConec.CommandTimeOut = pnValor
End Sub

Public Property Get ConexionActiva() As ADODB.Connection
Set ConexionActiva = dbConec
End Property

Public Property Let ConexionActiva(ByVal vNewValue As ADODB.Connection)
Set dbConec = vNewValue
End Property

'****************************************
'AbreConexionRemota: Función que devvuelve True si la se realiza la nueva conexión y false
'si la nueva conexion falla, se ingresa el codigo de agencia y False
'o los dos numeros iniciales del codigo de agencia y true
'****************************************
Public Function AbreConexionRemota(psAgeCod As String, Optional pbMensaje As Boolean = True, Optional pbConIniCie As Boolean = False, Optional psNroServidor As String = "01") As Boolean
    Dim lsConexionCad As String
    Dim ldFecha  As Date
    Dim ldFecSis As Date
    Dim lsAgeNom As String
    
    On Error GoTo AbreConexionRemotaErr
    AbreConexionRemota = True
    lsConexionCad = GetCadenaConexion(psAgeCod, psNroServidor)
    If lsConexionCad = "" Then
       AbreConexionRemota = False
       Exit Function
    End If
    
   If dbConec Is Nothing Then
      Set dbConec = New ADODB.Connection
   End If
   If dbConec.State = adStateOpen Then
      dbConec.Close
   End If
   dbConec.CommandTimeOut = 3000
   dbConec.CursorLocation = adUseClient
   dbConec.Open lsConexionCad
   dbConec.Execute "Set Dateformat mdy"
   AbreConexionRemota = True
   
   If Not pbConIniCie Then
      Exit Function
   End If
    
   Dim oConst As New NConstSistemas
   Dim oArea  As New DActualizaDatosArea
   ldFecha = CDate(oConst.LeeConstSistema(gConstSistFechaSistema))
   lsAgeNom = oArea.GetNombreAgencia(psAgeCod)
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
      Err.Raise 50001, "AbreConexionRemota", "La Agencia: " & lsAgeNom & " todavia no ha hecho Inico de día, o tiene una Fecha diferente a la agencia  local. La Fecha es :  " & ldFecha & " "
   End If
   Set oArea = Nothing
   Set oConst = Nothing
   Exit Function
AbreConexionRemotaErr:
    If pbMensaje Then
       Err.Raise 50001, "AbreConexionRemota", "El Servidor De la Agencia Nº : " & Mid(psAgeCod, 1, 2) & Space(2) & lsAgeNom & " no responde." & Chr(13) & "Consulte al Departamento de Sistemas"
    End If
    AbreConexionRemota = False
End Function

'****************************************************
'GetCadenaConexion:Funcion que devuelve la cadena de conección de la
'agencia segun el codigo que se ingresa
'****************************************************
Public Function GetCadenaConexion(psAgeCod As String, psNroServidor As String) As String
   Dim prs  As ADODB.Recordset
   Dim sSQL As String
    
   On Error GoTo GetCadenaConexionErr
   GetCadenaConexion = ""
   AbreConexion
   sSQL = "SELECT * FROM Servidor " _
        & "WHERE  substring(cCodAge,4,2) = '" + Mid(psAgeCod, 1, 2) + "' AND cEstado = 'A' And cNroSer = '" & psNroServidor & "'"
   Set prs = CargaRecordSet(sSQL)
   If prs.EOF And prs.BOF Then
      Err.Raise 50001, "DConecta", "No se Configuró acceso con Agencia " & psAgeCod & ". Consultar con Sistemas"
      Exit Function
   End If
   GetCadenaConexion = "PROVIDER=" + Trim(prs!cProvider) + ";uid=" + Trim(prs!cLogin) + ";pwd=" + Trim(prs!cPass) + ";DATABASE=" + Trim(prs!cDataBase) + ";SERVER=" + Trim(prs!cNomSer)
   RSClose prs
     
   CierraConexion
   Exit Function
GetCadenaConexionErr:
    Err.Raise 50001, "DConecta", "Error de Conexión " + Err.Description
End Function

Public Function StringServidorRemoto(psAgeCod As String, psNroServidor As String, Optional GetServidor As Boolean = True) As String
   Dim prs  As ADODB.Recordset
   Dim sSQL As String
    
   On Error GoTo GetCadenaConexionErr
   StringServidorRemoto = ""
   sSQL = "SELECT cNomser, cDataBase FROM Servidor " _
        & "WHERE  substring(cCodAge,4,2) = '" + Mid(psAgeCod, 1, 2) + "' AND cEstado = 'A' And cNroSer = '" & psNroServidor & "'"
   Set prs = CargaRecordSet(sSQL)
   If prs.EOF And prs.BOF Then
      Err.Raise 50001, "DConecta", "No se Configuró acceso con Agencia " & psAgeCod & ". Consultar con Sistemas"
      Exit Function
   End If
   If GetServidor Then
      StringServidorRemoto = "[" & Trim(prs!cNomSer) & "].[" & Trim(prs!cDataBase) & "].DBO."
   Else
      StringServidorRemoto = "[" & Trim(prs!cDataBase) & "].DBO."
   End If
   RSClose prs
      
   Exit Function
GetCadenaConexionErr:
    Err.Raise 50001, "DConecta: StringServidorRemoto", "Error de Conexión " + Err.Description
End Function


'****************************************************
'GetCadenaConexionEnlazado:Funcion que devuelve la cadena de
'conexión a setrvidores enlazados
'****************************************************
Public Function GetCadenaConexionEnlazado(psCuenta As String, psNroServidor As String, Optional pbSoloServidor As Boolean = False) As String
    Dim RegSistema As New ADODB.Recordset
    Dim SQL1 As String
    
    On Error GoTo ERROR
    
    SQL1 = "Select cNomSer, cDataBase From Servidor Where substring(cCodAge,4,2) = '" + Mid(psCuenta, 1, 2) + "' AND cEstado = 'A' And cNroSer = '" & psNroServidor & "'"
    RegSistema.Open SQL1, dbConec, adOpenStatic, adLockOptimistic, adCmdText
    
    If RegSistema.EOF And RegSistema.BOF Then
        MsgBox "No hay una conexion Activa con ese numero de Agencia", vbInformation, "Aviso"
        GetCadenaConexionEnlazado = ""
        Exit Function
    End If
    
    If Not pbSoloServidor Then
        GetCadenaConexionEnlazado = "[" & Trim(RegSistema!cNomSer) & "].[" & Trim(RegSistema!cDataBase) & "].dbo."
    Else
        GetCadenaConexionEnlazado = "[" & Trim(RegSistema!cNomSer) & "]."
    End If
    
    RegSistema.Close
    Set RegSistema = Nothing

    Exit Function
ERROR:
    MsgBox "Error de Conexión " + Err.Description, vbCritical, "Servicio Tecnico"
    
End Function

Public Property Get CommadTimeOut() As Long
'    dbConec.CommandTimeout = pnTiempodeEspera
End Property

Public Property Let CommadTimeOut(ByVal pnTiempodeEspera As Long)
    dbConec.CommandTimeOut = pnTiempodeEspera
End Property




