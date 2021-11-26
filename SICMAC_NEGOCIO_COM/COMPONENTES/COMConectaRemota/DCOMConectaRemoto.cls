VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMConectaRemoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Dim dbConec As ADODB.Connection

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
    Dim ocon As New COMConecta.DCOMConecta
    
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
   dbConec.CommandTimeout = 3000
   dbConec.CursorLocation = adUseClient
   dbConec.Open lsConexionCad
   dbConec.Execute "Set Dateformat mdy"
   AbreConexionRemota = True
   
   If Not pbConIniCie Then
      Exit Function
   End If
    
   Dim oConst As New COMConstSistema.NCOMConstSistema
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
      ocon.CierraConexion
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
   Dim ocon As New COMConecta.DCOMConecta
   Dim sSQL As String
    
   On Error GoTo GetCadenaConexionErr
   GetCadenaConexion = ""
   ocon.AbreConexion
   sSQL = "SELECT * FROM Servidor " _
        & "WHERE  substring(cCodAge,4,2) = '" + Mid(psAgeCod, 1, 2) + "' AND cEstado = 'A' And cNroSer = '" & psNroServidor & "'"
   Set prs = ocon.CargaRecordSet(sSQL)
   If prs.EOF And prs.BOF Then
      Err.Raise 50001, "DConecta", "No se Configuró acceso con Agencia " & psAgeCod & ". Consultar con Sistemas"
      Exit Function
   End If
   GetCadenaConexion = "PROVIDER=" + Trim(prs!cProvider) + ";uid=" + Trim(prs!cLogin) + ";pwd=" + Trim(prs!cPass) + ";DATABASE=" + Trim(prs!cDataBase) + ";SERVER=" + Trim(prs!cNomSer)
   RSClose prs
   ocon.CierraConexion
   Exit Function
GetCadenaConexionErr:
    Err.Raise 50001, "DConecta", "Error de Conexión " + Err.Description
End Function
