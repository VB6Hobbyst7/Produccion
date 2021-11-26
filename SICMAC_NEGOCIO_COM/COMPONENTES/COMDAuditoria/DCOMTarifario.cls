VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMTarifario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim oCon As COMConecta.DCOMConecta

Public Function CargarTarifarioGastos(ByVal psProducto As String, ByRef psMensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_TarifarioGastos '" & psProducto & "'"
    
    Set loReg = New COMConecta.DCOMConecta
        loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    If lrDatos Is Nothing Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psMensaje = "No existen datos"
        Exit Function
    End If
    Set CargarTarifarioGastos = lrDatos
    Set lrDatos = Nothing
End Function

