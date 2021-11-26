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
Attribute VB_Exposed = True
'********************************************************************
'** Nombre : DBancoPagador
'** Descripción : Para el manejo de Abonos del Banco Pagador
'** Creación : AMDO, 20131207 10:00:00 AM
'********************************************************************
Option Explicit
Dim dbConec As DConecta
Dim lsSql As String

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

Public Function DevuelveEstadoProceso(ByVal psArchivoNombre As String, ByVal psFecha As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConect As New DConecta

    sSql = " exec stp_sel_BP_DevuelveEstadoProceso '" & psArchivoNombre & "','" & psFecha & "'"
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(sSql)
    
    Set DevuelveEstadoProceso = rs
    oConect.CierraConexion
    Set rs = Nothing
    Set oConect = Nothing
End Function

Public Function InsertaAbonoBancoPagadorConfirmacion(ByVal pnIdProc As String, ByVal psUserProcesa As String, ByVal psOpeCod As String, ByVal psOpeCodITF As String, ByVal psAgeCod As String, ByVal psFecha As String) As Boolean
Dim cmd As New ADODB.Command
Dim prm As New ADODB.Parameter

    Set prm = New ADODB.Parameter
    Set prm = cmd.CreateParameter("@nIdProc", adInteger, adParamInput, , pnIdProc)
    cmd.Parameters.Append prm

    Set prm = New ADODB.Parameter
    Set prm = cmd.CreateParameter("@cUserConfirma", adVarChar, adParamInput, 20, psUserProcesa)
    cmd.Parameters.Append prm
    
    Set prm = New ADODB.Parameter
    Set prm = cmd.CreateParameter("@psOpeCod", adVarChar, adParamInput, 6, psOpeCod)
    cmd.Parameters.Append prm
    
    Set prm = New ADODB.Parameter
    Set prm = cmd.CreateParameter("@psOpeCodITF", adVarChar, adParamInput, 6, psOpeCodITF)
    cmd.Parameters.Append prm
    
    Set prm = New ADODB.Parameter
    Set prm = cmd.CreateParameter("@psAgeCod", adVarChar, adParamInput, 3, psAgeCod)
    cmd.Parameters.Append prm
    
    Set prm = New ADODB.Parameter
    Set prm = cmd.CreateParameter("@pdFecha", adDBDate, adParamInput, , psFecha)
    cmd.Parameters.Append prm
    
    Set prm = New ADODB.Parameter
    Set prm = cmd.CreateParameter("@pnResult", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    cmd.ActiveConnection = dbConec.ConexionActiva
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "stp_ins_AbonoBancoPagadorConfirmacion"
    cmd.Execute
    
    If cmd.Parameters(6).value = 0 Then
        InsertaAbonoBancoPagadorConfirmacion = True
    Else
        InsertaAbonoBancoPagadorConfirmacion = False
    End If
    
    Set prm = Nothing
    Set cmd = Nothing
End Function

Public Function DevuelveRegistrosxMoneda(ByVal pnIdProc As Long, ByVal pnEstadoReg As Integer, ByVal pnMoneda As Integer) As ADODB.Recordset	' ANGC20211020 INT A LONG
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConect As New DConecta

    sSql = "exec stp_sel_BP_DevuelveRegistrosxMoneda " & pnIdProc & "," & pnEstadoReg & "," & pnMoneda
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(sSql)
    
'    If Not rs.EOF Then
        Set DevuelveRegistrosxMoneda = rs
'    End If

    oConect.CierraConexion
    Set rs = Nothing
    Set oConect = Nothing
End Function

Public Function DevuelveValoresFinProcesoInicial(ByVal pnIdProc As Long) As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConect As New DConecta

    sSql = "exec stp_sel_BP_ValoresFinProcesoInicial " & pnIdProc
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(sSql)
    
    
    If Not rs.EOF Then
        Set DevuelveValoresFinProcesoInicial = rs
    End If
    
    oConect.CierraConexion
    Set rs = Nothing
    Set oConect = Nothing
End Function

Public Function DevuelveArchivosProcesados() As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConect As New DConecta

    sSql = "exec stp_sel_DevuelveArchivosProcesados"
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(sSql)
    
    Set DevuelveArchivosProcesados = rs

    oConect.CierraConexion
    Set rs = Nothing
    Set oConect = Nothing
End Function


