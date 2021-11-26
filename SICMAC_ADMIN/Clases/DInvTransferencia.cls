VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DInvTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub RegistrarTransferencia(ByVal iTipoTransferencia As Integer, ByVal sFecha As String, ByVal sOrigen As String, ByVal sDestino As String, ByVal sOrigenCod As String, ByVal sDestinoCod As String)
    Dim lsSQL As String
    Dim loReg As DConecta
    lsSQL = "exec stp_ins_Transferencia " & iTipoTransferencia & ", '" & sFecha & "', '" & sOrigen & "', '" & sDestino & "', '" & sOrigenCod & "', '" & sDestinoCod & "'"
    Set loReg = New DConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Function MaximoIdTransferencia() As Integer
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim lrDatos As ADODB.Recordset
    Dim lIdTrans As Integer
    lsSQL = "exec stp_sel_MaximoIdTransferencia"
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL)
    If lrDatos.RecordCount > 0 Then
        lIdTrans = lrDatos!Valor
    Else
        lIdTrans = lIdTrans
    End If
    MaximoIdTransferencia = lIdTrans
    Set loReg = Nothing
End Function

Public Sub RegistrarInventarioTransferencia(ByVal iTransferenciaId As Integer, ByVal cBSCod As String)
    Dim lsSQL As String
    Dim loReg As DConecta
    lsSQL = "exec stp_ins_InventarioTransferencia " & iTransferenciaId & ", '" & cBSCod & "'"
    Set loReg = New DConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Sub ModificarBS(ByVal sMovNro As String, ByVal cBSCod As String, ByVal sAreCod As String, ByVal sAgeCod As String, ByVal cCodBSNvo As String)
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_upd_ActualizarBS '" & sMovNro & "', '" & cBSCod & "', '" & sAreCod & "', '" & sAgeCod & "', '" & cCodBSNvo & "'"
    Set loReg = New DConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

'Reporte de Transferencias
Public Function CargarReporteTransferencia(ByVal sFI As String, ByVal sFF As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_ReporteTransferencia '" & sFI & "', '" & sFF & "'"
        
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarReporteTransferencia = lrDatos
    Set lrDatos = Nothing
End Function

'Asiento Transferencia
Public Sub RegistrarAsientoTransferencia(ByVal sCodInventario As String, ByVal sNvoCodInventario As String, ByVal cBSValor As Currency, ByVal dFechaRegistro As String, ByVal sTipo As String, ByVal cValorDepre As Currency)
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_AsientoTransferencia '" & sCodInventario & "', '" & sNvoCodInventario & "', " & cBSValor & ", '" & dFechaRegistro & "', '" & sTipo & "', " & cValorDepre & ""
    Set loReg = New DConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

'Reporte Asiento Tranferencia
Public Function CargarReporteAsientoTransferencia() As ADODB.Recordset 'ByVal sFI As String, ByVal sFF As String
    Dim lsSQL As String
    Dim loReg As DConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_ReporteAsientoTransferencia" ' '" & sFI & "', '" & sFF & "'"
        
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarReporteAsientoTransferencia = lrDatos
    Set lrDatos = Nothing
End Function
