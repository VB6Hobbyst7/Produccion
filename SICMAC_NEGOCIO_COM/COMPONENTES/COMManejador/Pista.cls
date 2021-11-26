VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim oCon As COMConecta.DCOMConecta

'Acciones: 1 Insertar; 2 Modificar; 3 Eliminar; 4 Alerta de Error; 5 Ingresar al Sistema; 6 Salir del Sistema
Public Sub InsertarPista(ByVal sOpeCod As String, ByVal sMovNro As String, ByVal sCodPersUser As String, ByVal sEquipo As String, ByVal sAccion As String, Optional ByVal sComentario As String = "", Optional ByVal sCodIdent As String = "", Optional ByVal nCodTipoIdent As Integer = 0)
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim pbTran As Boolean

    lsSQL = "exec stp_ins_Pistas '" & sOpeCod & "', '" & sMovNro & "', '" & sCodPersUser & "', '" & sEquipo & "', '" & sAccion & "', '" & sComentario & "', '" & sCodIdent & "', " & nCodTipoIdent & ""
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.Ejecutar (lsSQL)
    loReg.CierraConexion
End Sub

Public Sub InsertarPistaSesion(ByVal sOpeCod As String, ByVal sMovNro As String, ByVal sCodPersUser As String, ByVal sEquipo As String, ByVal sAccion As String, ByVal pnTipoApp As Integer) 'pnTipoApp 0: SICMAC Operaciones - 1: SICMAC Negocio
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim pbTran As Boolean

    lsSQL = "exec stp_ins_PistaSession '" & sOpeCod & "', '" & sMovNro & "', '" & sCodPersUser & "', '" & sEquipo & "', '" & sAccion & "'," & pnTipoApp
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.Ejecutar (lsSQL)
    loReg.CierraConexion
End Sub

Public Function ValidaAccesoPistaRF(ByVal pUser As String) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    sSql = "stp_sel_ValidaRF '" & pUser & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If R.BOF And R.EOF Then
        ValidaAccesoPistaRF = False
    Else
        ValidaAccesoPistaRF = True
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

End Function

Public Function ValidaIniciaSessionRF(ByVal pUser As String, fecha As Date, ByVal pMaquina As String, ByVal pnTipoApp As Integer, ByRef psMaquinaActiva As String) As Boolean 'pnTipoApp 0: SICMAC Operaciones - 1: SICMAC Negocio
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
       
   sSql = " exec stp_sel_validaSession '" & pUser & "','" & Format(fecha, "YYYYMMDD") & "','" & pMaquina & "'," & pnTipoApp
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If R.BOF And R.EOF Then
        ValidaIniciaSessionRF = False
    Else
        psMaquinaActiva = R!cMaquina 'JUEZ 20160405
        ValidaIniciaSessionRF = True
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
End Function
'JUEZ 20160125 ***********************************************
Public Sub ActualizarPistaSesion(ByVal sCodPersUser As String, ByVal sEquipo As String, ByVal pnTipoApp As Integer) 'pnTipoApp 0: SICMAC Operaciones - 1: SICMAC Negocio
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim pbTran As Boolean

    lsSQL = "exec stp_upd_PistaSession '" & sCodPersUser & "', '" & sEquipo & "'," & pnTipoApp
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.Ejecutar (lsSQL)
    loReg.CierraConexion
End Sub
Public Function ListarPistaSesiones(ByVal pbActivo As Boolean) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    sSql = "stp_sel_PistaSession " & IIf(pbActivo, 1, 0)
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ListarPistaSesiones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'END JUEZ ****************************************************
