VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMTCEspPermiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsConexion As String
Dim vsServerCom As String
Dim vsServerPers As String
Dim vsFormatoFecha As String
Dim oCon As COMConecta.DCOMConecta
Dim oerror As New COMConecta.COMErrorHandling

Private Sub Class_Initialize()

    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then
        oerror.RaiseError oerror.MyUnhandledError, "DTipoCambio:Initialize Method. Error en Conexion de Base de datos"
        Exit Sub
    End If
End Sub

Private Sub Class_Terminate()
oCon.CierraConexion
Set oCon = Nothing
End Sub
'Public Function InsertaTipoCambio(ByVal psMovNro As String, _
'ByVal psCodOpe As String, ByVal pnTCReg As Currency, ByVal pdFecSis As Date, _
'ByVal pnMontoReg As Currency, ByVal psPersCod As String) As Integer
'    Dim sql As String
'    'GIPO ERS0692016
'    Dim ObjHora As COMConecta.DCOMConecta
'    Dim dFechaReg As String
'    Set ObjHora = New COMConecta.DCOMConecta
'    ObjHora.AbreConexion
'    dFechaReg = Format(pdFecSis & " " & oCon.GetHoraServer, "MM/DD/YYYY hh:mm:ss AMPM")
'     'dFechaReg = Format(pdFecSis, "MM/DD/YYYY")
'
''    ObjHora.CierraConexion
''    Set ObjHora = Nothing
'    InsertaTipoCambio = 1
'    sql = "Insert Into CapAutorizacionTC"
'    sql = sql & " (cMovNro,cCodOpe,nTCReg,dFechaReg,nMontoReg,cPersCod,nEstado,cCodUserAprob,cMovNroAprob,nTCAprob) values "
'    sql = sql & " ('" & psMovNro & "','" & psCodOpe & "'," & pnTCReg & ",'" & dFechaReg & "'," & pnMontoReg & ",'" & psPersCod & "',0,null,null,null)"
'
'    'oCon.BeginTrans
'    oCon.Ejecutar sql
'    'oCon.CommitTrans
'
'    InsertaTipoCambio = 0
'End Function
'APRI20180201 MEJORA INC180105004

'Public Sub InsertaTipoCambio(ByVal psMovNro As String, _
'ByVal psCodOpe As String, ByVal pnTCReg As Currency, ByVal pdFecSis As Date, _ revertir
'ByVal pnMontoReg As Currency, ByVal psPersCod As String)
'    Dim Sql As String
'    Dim dFechaReg As String
'
'    dFechaReg = Format(pdFecSis & " " & oCon.GetHoraServer, "MM/DD/YYYY hh:mm:ss AMPM")
'
'    Dim conec As COMConecta.DCOMConecta
'    Set conec = New COMConecta.DCOMConecta
'    conec.AbreConexion
'
'    Sql = "EXEC stp_ins_CapAutorizacionTC '" & psMovNro & "','" & psCodOpe & "'," & pnTCReg & ",'" & dFechaReg & "'," & pnMontoReg & ",'" & psPersCod & "'," & 0 'APRI20180205 MEJORA
'
'    conec.Ejecutar Sql
'    Set conec = Nothing
'
'End Sub

'LUCV20180224: El proceso de este método se agregó en el mismo formulario: frmCompraVenta
Public Function InsertaTipoCambio(ByVal psMovNro As String, _
                                  ByVal psCodOpe As String, ByVal pnTCReg As Currency, ByVal pdFecSis As Date, _
                                  ByVal pnMontoReg As Currency, ByVal psPersCod As String) As Integer
    Dim Sql As String
    'GIPO ERS0692016
    Dim ObjHora As COMConecta.DCOMConecta
    Dim dFechaReg As String
    Set ObjHora = New COMConecta.DCOMConecta
    ObjHora.AbreConexion
    dFechaReg = Format(pdFecSis & " " & oCon.GetHoraServer, "MM/DD/YYYY hh:mm:ss AMPM")
    dFechaReg = Format(pdFecSis, "MM/DD/YYYY")

    ObjHora.CierraConexion
    Set ObjHora = Nothing
    InsertaTipoCambio = 1
    Sql = "Insert Into CapAutorizacionTC"
    Sql = Sql & " (cMovNro,cCodOpe,nTCReg,dFechaReg,nMontoReg,cPersCod,nEstado,cCodUserAprob,cMovNroAprob,nTCAprob) values "
    Sql = Sql & " ('" & psMovNro & "','" & psCodOpe & "'," & pnTCReg & ",'" & dFechaReg & "'," & pnMontoReg & ",'" & psPersCod & "',0,null,null,null)"

    'oCon.BeginTrans
    oCon.Ejecutar Sql
    'oCon.CommitTrans
    InsertaTipoCambio = 0
End Function

Public Sub CambioEstado_TC_Especial(ByVal psMovNro As String, ByVal pnEstado As Integer, _
ByVal pnTCReg As Currency, ByVal pdFecSis As Date, _
ByVal psCodUser As String, ByVal pnIdAprob As Long)

    Dim Sql As String
    Dim dFechaReg As String

    dFechaReg = Format(pdFecSis, "MM/DD/YYYY")

    Sql = "EXEC stp_upd_EstadoCapAutorizacionTC " & pnEstado & ",'" & psCodUser & "','" & dFechaReg & "','" & psMovNro & "'," & pnTCReg & "," & pnIdAprob

    oCon.Ejecutar Sql

End Sub
Public Sub OpeTipoCambioEspecialCliente(ByVal psNroMov As String, ByVal pnNroMov As Long)
    Dim Sql As String
    Sql = "Exec stp_upd_OpeCapAutorizacionTC '" & psNroMov & "'," & pnNroMov
    oCon.Ejecutar Sql
End Sub
Public Sub OpeExtTCEspecialCliente(ByVal pnNroMov As Long)
    Dim Sql As String
    Sql = "Exec stp_upd_ExtornoOpeCapAutorizacionTC " & pnNroMov
    oCon.Ejecutar Sql
End Sub
Public Function ObtieneOpeTipoCambioEspecialCliente(ByVal pnNroMov As Long) As Boolean
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Sql = "Exec stp_sel_ObtieneCapAutorizacionTC " & pnNroMov
    Set rs = oCon.CargaRecordSet(Sql)
    If Not (rs.BOF And rs.EOF) Then
        ObtieneOpeTipoCambioEspecialCliente = True
    Else
        ObtieneOpeTipoCambioEspecialCliente = False
    End If

End Function
'END APRI

Public Sub Aprobar_TC_Especial(ByVal psMovNro As String, _
ByVal pnTCReg As Currency, ByVal pdFecSis As Date, _
ByVal psCodUser As String, ByVal pnIdAprob As Long)
    
    Dim Sql As String
'    Dim ObjHora As COMConecta.DCOMConecta
    Dim dFechaReg As String
    
'    Set ObjHora = New COMConecta.DCOMConecta
'    ObjHora.AbreConexion
'    dFechaReg = Format(pdFecSis & " " & oCon.GetHoraServer, "MM/DD/YYYY hh:mm:ss AMPM")
    dFechaReg = Format(pdFecSis, "MM/DD/YYYY")

'    ObjHora.CierraConexion
'    Set ObjHora = Nothing
    
    Sql = "update CapAutorizacionTC set nEstado = 2, cCodUserAprob = '" & psCodUser & "',"
    Sql = Sql & " dFechaAProb = '" & dFechaReg & "', cMovNroAprob = '" & psMovNro & "', nTCAprob = " & pnTCReg
    Sql = Sql & " Where nCodAut = " & pnIdAprob

    'oCon.BeginTrans
    oCon.Ejecutar Sql
    'oCon.CommitTrans
    
End Sub
Public Sub Rechazar_TC_Especial(ByVal psMovNro As String, _
ByVal pdFecSis As Date, _
ByVal psCodUser As String, ByVal pnIdAprob As Long)
    
    Dim Sql As String
'    Dim ObjHora As COMConecta.DCOMConecta
    Dim dFechaReg As String
    
'    Set ObjHora = New COMConecta.DCOMConecta
'    ObjHora.AbreConexion
'    dFechaReg = Format(pdFecSis & " " & oCon.GetHoraServer, "MM/DD/YYYY hh:mm:ss AMPM")
     dFechaReg = Format(pdFecSis, "MM/DD/YYYY")

'    ObjHora.CierraConexion
'    Set ObjHora = Nothing
'
    Sql = "update CapAutorizacionTC set nEstado = 1, cCodUserAprob = '" & psCodUser & "',"
    Sql = Sql & " dFechaAProb = '" & dFechaReg & "', cMovNroAprob = '" & psMovNro & "', nTCAprob = 0"
    Sql = Sql & " Where nCodAut = " & pnIdAprob

    'oCon.BeginTrans
    oCon.Ejecutar Sql
    'oCon.CommitTrans
    
End Sub
'GIPO ERS069-2016  01-01-2017
Public Sub RechazarTipoCambioEspecialCliente(ByVal psNroMov As String)
    Dim Sql As String
    Sql = "Exec sp_upd_ERS0692016_RechazarAprobacionTCEspecial '" & psNroMov & "'"
    oCon.Ejecutar Sql
End Sub

Public Sub AceptarTipoCambioEspecialCliente(ByVal psNroMov As String)
    Dim Sql As String
    Sql = "Exec sp_upd_ERS0692016_AceptarAprobacionTCEspecial '" & psNroMov & "'"
    oCon.Ejecutar Sql
End Sub

Public Function GetTc_AUT(ByVal pdFecSis As Date, ByVal pnCodAut As Long, _
ByVal psUser As String, ByVal psOpeCod As String, ByVal psPersCod As String) As ADODB.Recordset
Dim Sql As String

Sql = " select"
Sql = Sql & " cMovNro,nCodAut,cCodOpe,nTCReg,dFechaReg,nMontoReg,P.cPersCod,"
Sql = Sql & " nEstado , cCodUserAprob, dFechaAprob, cMovNroAprob, nTCAprob,"
Sql = Sql & " cPersNombre,cPersDireccDomicilio,nPersPersoneria,"
Sql = Sql & " IsNull((Select cPersIdnro from PersID D where D.cPersCod = P.cPersCod and cPersIdTpo ='1'),'') DNI,"
Sql = Sql & " IsNull((Select cPersIdnro from PersID D where D.cPersCod = P.cPersCod and cPersIdTpo= '2'),'') RUC"
Sql = Sql & " from CapAutorizacionTC A "
Sql = Sql & " Inner Join Persona P on P.cPersCod = A.cPersCod"
Sql = Sql & " where A.nEstado = 2 and nCodAut =" & pnCodAut
Sql = Sql & " and datediff(day,dFechaReg,'" & Format(pdFecSis, "MM/DD/YYYY") & "') = 0"
Sql = Sql & " AND RIGHT(cMovNro,4) = '" & psUser & "' and cCodOpe = '" & psOpeCod & "'"

Set GetTc_AUT = oCon.CargaRecordSet(Sql)
End Function

'GIPO ERS069-2016
Public Function GetTc_Autorizacion(ByVal pdFecSis As Date, _
ByVal psUser As String, ByVal psOpeCod As String, ByVal psPersCod As String) As ADODB.Recordset
Dim Sql As String
Sql = "Exec sp_sel_ERS0692016_ObtenerAutorizacion '" & psPersCod & "','" & Format(pdFecSis, "MM/DD/YYYY") & "','" & psUser & "','" & psOpeCod & "'"
Set GetTc_Autorizacion = oCon.CargaRecordSet(Sql)
End Function

Public Function GetTc_AutorizacionByNroMov(ByVal cMovnro As String) As ADODB.Recordset
Dim Sql As String
Sql = "Exec sp_sel_ERS0692016_ObtenerAutorizacionByNroMov '" & cMovnro & "'"
Set GetTc_AutorizacionByNroMov = oCon.CargaRecordSet(Sql)
End Function

Public Function GetTc_AccesoFormulario(ByVal cUser As String) As ADODB.Recordset
Dim Sql As String
Sql = "Exec sp_sel_ERS0692016_permitirAccesoAprobacionTC '" & cUser & "'"
Set GetTc_AccesoFormulario = oCon.CargaRecordSet(Sql)
End Function
'END GIPO

Public Function Get_All(ByVal pdFecSis As Date, Optional psCodAge As String = "00") As ADODB.Recordset
Dim Sql As String

Sql = "Exec sp_sel_ERS0692016_ObtenerAutorizacionesTCEspecial '" & Format(pdFecSis, "YYYY/MM/DD") & "'"

Set Get_All = oCon.CargaRecordSet(Sql)
End Function


Public Function GetDoc_Persona(ByVal psPersCod As String) As ADODB.Recordset
Dim Sql As String
Sql = " Select cConsDescripcion ,cPersIDnro"
Sql = Sql & " From PersID P"
Sql = Sql & " Inner Join Constante C on C.nConsValor  = P.cPersIDTpo and nConsCod = 1003"
Sql = Sql & " and P.cPersCod = '" & psPersCod & "'"
Sql = Sql & " Order by C.nConsValor"
Set GetDoc_Persona = oCon.CargaRecordSet(Sql)
End Function


