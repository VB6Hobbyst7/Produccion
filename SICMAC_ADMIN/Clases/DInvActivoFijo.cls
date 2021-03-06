VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DInvActivoFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ListarOrdenCompraXAF(ByVal sDesde As String, ByVal sHasta As String, ByVal sMoneda As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_InvListarOrdenCompraXAF '" & sDesde & "', '" & sHasta & "', '" & sMoneda & "'"
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set ListarOrdenCompraXAF = lrDatos
    Set lrDatos = Nothing
End Function

Public Function ListarBBSXOC(ByVal sMovNro As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_InvListarBBSXOC '" & sMovNro & "'"
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set ListarBBSXOC = lrDatos
    Set lrDatos = Nothing
End Function

Public Function DevolverDatosBBS(ByVal sMovNro As String, ByVal sBSCod As String, ByVal iMovItem As Integer) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_DarDatosBBS '" & sMovNro & "', '" & sBSCod & "', " & iMovItem & ""
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set DevolverDatosBBS = lrDatos
    Set lrDatos = Nothing
End Function

Public Function DevolverAgencia(ByVal sMovNro As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_DarAgencia '" & sMovNro & "'"
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set DevolverAgencia = lrDatos
    Set lrDatos = Nothing
End Function

Public Function DevolverCorrelativo(ByVal sBSCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_sel_DarMaximo '" & sBSCod & "'"
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set DevolverCorrelativo = lrDatos
    Set lrDatos = Nothing
End Function

'JIPR20200313 INICIO
Public Function DevolverCorrelativoAF(ByVal sBSCod As String, Optional sAgencia As String, Optional nTipo As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec sp_sel_SerieActivoFijo '" & sBSCod & "', '" & sAgencia & "', " & nTipo & ""
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set DevolverCorrelativoAF = lrDatos
    Set lrDatos = Nothing
End Function
'JIPR20200313 FIN


Public Function DevolverTipoBien(ByVal sBSCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_sel_DarTipoBien '" & sBSCod & "'"
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set DevolverTipoBien = lrDatos
    Set lrDatos = Nothing
End Function

Public Function DevolverActivadoXBBS(ByVal sMovNro As String, ByVal sBSCod As String, ByVal iMovItem As Integer) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_sel_ActivadoXBBS '" & sMovNro & "', '" & sBSCod & "', '" & iMovItem & "'"
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set DevolverActivadoXBBS = lrDatos
    Set lrDatos = Nothing
End Function

Public Sub RegistrarAF(ByVal sCodInventario As String, ByVal sNombre As String, ByVal sAgencia As String, ByVal sMarca As String, ByVal sModelo As String, ByVal sSerie As String, ByVal sFIngreso As String, ByVal sMovNro As String, ByVal sBSCod As String, ByVal iMovItem As Integer)
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_ActivoFijo '" & sCodInventario & "', '" & sNombre & "', '" & sAgencia & "', '" & sMarca & "', '" & sModelo & "', '" & sSerie & "', '" & sFIngreso & "', '" & sMovNro & "', '" & sBSCod & "', " & iMovItem & ""
    Set loReg = New DConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Function CargarTipoAF() As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_TipoAF"
        
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarTipoAF = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarBienXTipo(ByVal sBSCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_BienXTipo '" & sBSCod & "'"
        
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarBienXTipo = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarReporteAF(ByVal sAgeCod As String, ByVal sBSCod As String, ByVal sBSCodBien As String, ByVal sFecha As String, ByVal psCategoBien As String, ByVal psDel As String, ByVal psAl As String, Optional ByVal psSerie As String = "") As ADODB.Recordset
    '*** PEAC 20110705 - SE AGREGO PARAMETRO (psCategoBien)
    '*** PEAC 20120425 - SE AGREG� PARAMETRO psSerie
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_ReporteAF '" & sAgeCod & "', '" & sBSCod & "', '" & sBSCodBien & "', '" & Format(sFecha, "yyyymmdd") & "','" & psCategoBien & "','" & psDel & "','" & psAl & "','" & psSerie & "'"
        
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarReporteAF = lrDatos
    Set lrDatos = Nothing
End Function

Public Function DevolverBienes(ByVal sCodInv As String, ByVal sFI As String, ByVal sFF As String, ByVal sDescripcion As String, ByVal sFDepre As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_Bienes '" & sCodInv & "', '" & sFI & "', '" & sFF & "', '" & sDescripcion & "', '" & sFDepre & "'"
        
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set DevolverBienes = lrDatos
    Set lrDatos = Nothing
End Function


'*** PEAC 20200511
Public Function DevolverCorrelativoNoDepre(ByVal sBSCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_sel_DarMaximoNoDepre '" & sBSCod & "'"
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set DevolverCorrelativoNoDepre = lrDatos
    Set lrDatos = Nothing
End Function

