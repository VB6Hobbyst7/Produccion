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
Public sCentralCom As String
Public sCentralPers As String
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

Public Sub CierraConexion()
On Error Resume Next

If dbConec Is Nothing Then Exit Sub

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
   psStringConec = oConec.CadenaConexion(App.Path & "\SicmactA.Ini")

   'psStringConec = CadenaConexionIniSegInt()
End If
dbConec.CursorLocation = adUseClient
dbConec.Open psStringConec
dbConec.Execute "Set DateFormat mdy"

Set ConexionActiva = dbConec
Exit Function

AbreConexionErr:
   AbreConexion = False
   Call RaiseError(MyUnhandledError, "DConecta:AbreConexion Method")
End Function

Public Function CargaRecordSet(sql As String, Optional LockType As LockTypeEnum = adLockReadOnly, Optional CursorType As CursorTypeEnum = adOpenStatic, Optional pbCierraTranasaccionActiva As Boolean = True) As ADODB.Recordset
   Dim prs As ADODB.Recordset
   On Error GoTo ERROR
   Set prs = New ADODB.Recordset
   If prs.State = adStateOpen Then prs.Close
   prs.CursorLocation = adUseClient
   dbConec.CommandTimeout = 20000
   prs.Open sql, dbConec, CursorType, LockType, adCmdText
   If pbCierraTranasaccionActiva Then prs.ActiveConnection = Nothing
   Set CargaRecordSet = prs
   Set prs = Nothing
   Exit Function
ERROR:
   Set CargaRecordSet = Nothing
   Call RaiseError(MyUnhandledError, "DConecta:CargaRecordSet Method")
End Function

Public Function Ejecutar(psql As String) As ADODB.Recordset
On Error GoTo ErrEjecutar
Set Ejecutar = dbConec.Execute(psql)
Exit Function
ErrEjecutar:
   Call RaiseError(MyUnhandledError, "DConecta:Ejecutar Method")
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
   dbConec.BeginTrans
   dbConec.Execute sSql
   dbConec.CommitTrans
   sSql = ""
   lTransac = False
End If
Exit Function
Ejecutar:
   If lTransac Then
      dbConec.RollBackTrans
      lTransac = False
      sSql = ""
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
Public Sub RollBackTrans()
On Error GoTo Ejecutar
   dbConec.RollBackTrans
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

Public Function ConexionActivaReal() As ADODB.Connection
    Set ConexionActivaReal = dbConec
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
Public Function AbreConexionRemota(psAgeCod As String, Optional pbMensaje As Boolean = True, Optional pbConIniCie As Boolean = False, Optional psNroServidor As String = "01", Optional pdFecSis As Date = CDate("01/01/1900")) As Boolean
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
   dbConec.CommandTimeout = 640
   dbConec.CursorLocation = adUseClient
   dbConec.Open lsConexionCad
   dbConec.Execute "Set Dateformat mdy"
   AbreConexionRemota = True
      
   sCentralCom = ReadVarSisCon("ADM", "cCentralCom")
   sCentralPers = ReadVarSisCon("ADM", "cCentralPers")
   
   If Not pbConIniCie Then
      Exit Function
   End If

   Dim oConst As New NConstSistemas
   Dim oArea  As New DActualizaDatosArea
   
   ldFecha = CDate(oConst.LeeConstSistema(gConstSistFechaSistema))
   ldFecha = CDate(ReadVarSisCon("ADM", "dFecSis"))
   ldFecSis = pdFecSis
   lsAgeNom = oArea.GetNomAge(psAgeCod)
   If DateDiff("d", ldFecha, ldFecSis) = 0 Then
      If DateDiff("d", CDate(ReadVarSisCon("ADM", "dCieAho")), ldFecSis) = 0 Or DateDiff("d", CDate(ReadVarSisCon("ADM", "dCieCre")), ldFecSis) = 0 Or DateDiff("d", CDate(ReadVarSisCon("ADM", "dCiePre")), ldFecSis) = 0 Then
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
   Dim sSql As String
    
   On Error GoTo GetCadenaConexionErr
   GetCadenaConexion = ""
   AbreConexion
   sSql = "SELECT * FROM Servidor " _
        & "WHERE  substring(cCodAge,4,2) = '" + Mid(psAgeCod, 1, 2) + "' AND cEstado = 'A' And cNroSer = '" & psNroServidor & "'"
   Set prs = CargaRecordSet(sSql)
   If prs.EOF And prs.BOF Then
      Err.Raise 50001, "DConecta", "No se Configuró acceso con Agencia " & psAgeCod & ". Consultar con Sistemas"
      Exit Function
   End If
   'GetCadenaConexion = "PROVIDER=" & Trim(prs!cProvider) & ";User ID=" & Trim(prs!cLogin) & ";Password=" & Trim(prs!cPass) & ";Initial Catalog=" & Trim(prs!cDataBase) & ";Data Source=" & Trim(prs!cNomSer) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False"
   GetCadenaConexion = "Provider=" & Trim(prs!cProvider) & ";Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & Trim(prs!cDataBase) & ";Data Source=" & Trim(prs!cNomSer)

   RSClose prs
   CierraConexion
   Exit Function
GetCadenaConexionErr:
    Err.Raise 50001, "DConecta", "Error de Conexión " + Err.Description
End Function


Public Function StringServidorRemoto(psAgeCod As String, psNroServidor As String, Optional GetServidor As Boolean = True) As String
   Dim prs  As ADODB.Recordset
   Dim sSql As String
    
   On Error GoTo GetCadenaConexionErr
   StringServidorRemoto = ""
   sSql = "SELECT cNomser, cDataBase FROM Servidor " _
        & "WHERE  substring(cCodAge,4,2) = '" + Mid(psAgeCod, 1, 2) + "' AND cEstado = 'A' And cNroSer = '" & psNroServidor & "'"
   Set prs = CargaRecordSet(sSql)
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

'********************************
'ReadVarSisCon: Funcion devuelve el valor de Varsistema pero de la conexion que se le envia
'********************************
Private Function ReadVarSisCon(psCod As String, psNomVar As String) As String
    Dim sqlF As String
    Dim rsF As New ADODB.Recordset
    
    sqlF = "Select cValorVar from VarSistema where cCodProd = '" & psCod & "' and cNomVar = '" & psNomVar & "'"
    rsF.Open sqlF, dbConec, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ReadVarSisCon = Trim(rsF!cValorVar)
    
    rsF.Close
    Set rsF = Nothing
End Function


'****************************************************
'GetCadenaConexionEnlazado:Funcion que devuelve la cadena de
'conexión a setrvidores enlazados
'****************************************************
Public Function GetCadenaConexionEnlazado(psCuenta As String, psNroServidor As String, Optional pbSoloServidor As Boolean = False) As String
    Dim RegSistema As New ADODB.Recordset
    Dim Sql1 As String
    
    On Error GoTo ERROR
    
    Sql1 = "Select cNomSer, cDataBase From Servidor Where substring(cCodAge,4,2) = '" + Mid(psCuenta, 1, 2) + "' AND cEstado = 'A' And cNroSer = '" & psNroServidor & "'"
    RegSistema.Open Sql1, dbConec, adOpenStatic, adLockOptimistic, adCmdText
    
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

Public Function CadenaConexionIni(Optional psArchivo As String = "", Optional psAplicacion As String = "SICMACT") As String
    Dim lsProvider As String
    Dim lsServer As String
    Dim lsUser As String
    Dim lsPassword As String
    Dim lsDatabase As String
    Dim lsCadena As String
    
    On Error GoTo ErrorCadCon
    If psArchivo = "" Then
        psArchivo = App.Path & "\SicmactA.Ini"
    End If
       
    lsProvider = LeerArchivoIni(Encripta(psAplicacion), Encripta("Provider"), psArchivo)
    lsServer = LeerArchivoIni(Encripta(psAplicacion), Encripta("Server"), psArchivo)
    lsUser = LeerArchivoIni(Encripta(psAplicacion), Encripta("User"), psArchivo)
    lsPassword = LeerArchivoIni(Encripta(psAplicacion), Encripta("Password"), psArchivo)
    lsDatabase = LeerArchivoIni(Encripta(psAplicacion), Encripta("DataBase"), psArchivo)
    
    lsCadena = "PROVIDER=" & Trim(Encripta(lsProvider, False)) & ";User ID=" & Trim(Encripta(lsUser, False)) & ";Password=" & Trim(Encripta(lsPassword, False)) & ";INITIAL CATALOG=" & Trim(Encripta(lsDatabase, False)) & ";DATA SOURCE=" & Trim(Encripta(lsServer, False)) & ""
    CadenaConexionIni = lsCadena
    
    Exit Function
ErrorCadCon:
    Err.Raise Err.Number, "Cadena de Conexion", "Error al Generar cadena de Conexion. Revise su Archivo de Inicio" & vbCrLf & "DESCRIPCION DEL ERROR:" & vbCrLf & Err.Description
End Function

Public Function CadenaConexionIniSegInt(Optional psArchivo As String = "", Optional psAplicacion As String = "SICMACT") As String
    Dim lsProvider As String
    Dim lsServer As String
    Dim lsUser As String
    Dim lsPassword As String
    Dim lsDatabase As String
    Dim lsCadena As String
    
    On Error GoTo ErrorCadCon
    If psArchivo = "" Then
        psArchivo = App.Path & "\SicmactA.Ini"
    End If
       
    lsProvider = LeerArchivoIni(Encripta(psAplicacion), Encripta("Provider"), psArchivo)
    lsServer = LeerArchivoIni(Encripta(psAplicacion), Encripta("Server"), psArchivo)
    lsUser = LeerArchivoIni(Encripta(psAplicacion), Encripta("User"), psArchivo)
    lsPassword = LeerArchivoIni(Encripta(psAplicacion), Encripta("Password"), psArchivo)
    lsDatabase = LeerArchivoIni(Encripta(psAplicacion), Encripta("DataBase"), psArchivo)
    
    lsCadena = "Provider=" & Trim(Encripta(lsProvider, False)) & ";Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & Trim(Encripta(lsDatabase, False)) & ";Data Source=" & Trim(Encripta(lsServer, False))
    CadenaConexionIniSegInt = lsCadena
    
    Exit Function
ErrorCadCon:
    Err.Raise Err.Number, "Cadena de Conexion", "Error al Generar cadena de Conexion. Revise su Archivo de Inicio" & vbCrLf & "DESCRIPCION DEL ERROR:" & vbCrLf & Err.Description
End Function

