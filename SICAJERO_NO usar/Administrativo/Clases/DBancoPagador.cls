VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBancoPagador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : DBancoPagador
'** Descripción : Para el manejo de Abonos del Banco Pagador
'** Creación : AMDO, 20131125 10:00:00 AM
'********************************************************************
Option Explicit
Dim dbConec As DConecta
Dim lsSQL As String

Private Sub Class_Initialize()
    Set dbConec = New DConecta
    If Not dbConec.AbreConexion Then
        Call RaiseError(MyUnhandledError, "DBien:Initialize Method. Error en Conexion a Base de datos")
    End If
End Sub
Private Sub Class_Terminate()
    If Not dbConec Is Nothing Then
        dbConec.CierraConexion
        Set dbConec = Nothing
    End If
End Sub
Public Sub dBeginTrans()
    dbConec.BeginTrans
End Sub
Public Sub dRollbackTrans()
    dbConec.RollbackTrans
    dbConec.Ejecutar "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
End Sub
Public Sub dCommitTrans()
    dbConec.CommitTrans
End Sub

Public Function InsertaAbonoBancoPagadorInicial(ByVal pnIdProc As String, ByVal psUserProcesa As String, ByVal psOpeCod As String, ByVal psAgeCod As String) As Boolean
Dim Cmd As New ADODB.Command
Dim Prm As New ADODB.Parameter

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nIdProc", adInteger, adParamInput, , pnIdProc)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cUserProcesa", adVarChar, adParamInput, 20, psUserProcesa)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCod", adVarChar, adParamInput, 6, psOpeCod)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psAgeCod", adVarChar, adParamInput, 3, gsCodAge)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnResult", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Cmd.ActiveConnection = dbConec.ConexionActiva
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "stp_ins_AbonoBancoPagadorInicial"
    Cmd.Execute
    
    If Cmd.Parameters(4).Value = 0 Then
        InsertaAbonoBancoPagadorInicial = True
    Else
        InsertaAbonoBancoPagadorInicial = False
    End If
    
    Set Prm = Nothing
    Set Cmd = Nothing
End Function

Public Function DevuelveEstadoProceso(ByVal psArchivoNombre As String, ByVal psFecha As String) As ADODB.Recordset
    Dim ssql As String
    Dim rs As New ADODB.Recordset
    Dim oconect As New DConecta

    ssql = " exec stp_sel_BP_DevuelveEstadoProceso '" & psArchivoNombre & "','" & psFecha & "'"
    oconect.AbreConexion
    Set rs = oconect.CargaRecordSet(ssql)
    
    Set DevuelveEstadoProceso = rs
    oconect.CierraConexion
    Set rs = Nothing
    Set oconect = Nothing
End Function

Public Function InsertaAbonoBancoPagadorConfirmacion(ByVal pnIdProc As String, ByVal psUserProcesa As String, ByVal psOpeCod As String, ByVal psAgeCod As String) As Boolean
Dim Cmd As New ADODB.Command
Dim Prm As New ADODB.Parameter

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nIdProc", adInteger, adParamInput, , pnIdProc)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cUserProcesa", adVarChar, adParamInput, 20, psUserProcesa)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCod", adVarChar, adParamInput, 6, psOpeCod)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psAgeCod", adVarChar, adParamInput, 3, gsCodAge)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnResult", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Cmd.ActiveConnection = dbConec.ConexionActiva
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "stp_ins_AbonoBancoPagadorConfirmacion"
    Cmd.Execute
    
    If Cmd.Parameters(4).Value = 0 Then
        InsertaAbonoBancoPagadorInicial = True
    Else
        InsertaAbonoBancoPagadorInicial = False
    End If
    
    Set Prm = Nothing
    Set Cmd = Nothing
End Function


