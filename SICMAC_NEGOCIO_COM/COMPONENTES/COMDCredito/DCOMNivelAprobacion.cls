VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMNivelAprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'EJVG20160205 ERS002-2016***
Public coConex As COMConecta.DCOMConecta

Public Sub dBeginTrans()
    coConex.BeginTrans
End Sub

Public Sub dRollbackTrans()
    coConex.RollbackTrans
    coConex.Ejecutar "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
End Sub
Public Sub dCommitTrans()
    coConex.CommitTrans
End Sub
Private Sub Class_Initialize()
    Dim loIni As COMConecta.DCOMClasIni
    Dim csConexion As String
    
    Set loIni = New COMConecta.DCOMClasIni
    csConexion = loIni.CadenaConexion
    Set loIni = Nothing
    
    Set coConex = New COMConecta.DCOMConecta
    coConex.AbreConexion csConexion
End Sub
Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub
'END EJVG *******
'MAVM 20110429 ***
Public Sub InsertarNivApr(ByVal cNivAprCod As String, ByVal nNivelPor As Integer, ByVal cNivelCod As String, ByVal nMontoMin As Currency, ByVal nMontoMax As Currency, ByVal cRiesgoCod As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    
On Error GoTo ControlError
    lsSQL = "exec stp_ins_ColocCredNivApr " & "'" & cNivAprCod & "', " & nNivelPor & ", '" & cNivelCod & "', " & nMontoMin & ", " & nMontoMax & ", '" & cRiesgoCod & "'"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
    Exit Sub
ControlError:
    Set loReg = Nothing
End Sub

Public Sub InsertarNivAprCargo(ByVal cNivAprCod As String, ByVal cCargoCod As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    
On Error GoTo ControlError
    lsSQL = "exec stp_ins_ColocCredNivAprCargo " & "'" & cNivAprCod & "', '" & cCargoCod & "'"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
    Exit Sub
ControlError:
    Set loReg = Nothing
End Sub

Public Sub InsertarNivAprTipoCredito(ByVal cNivAprCod As String, ByVal cTipoCreditoCod As String, ByVal cAgeCod As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    
On Error GoTo ControlError
    lsSQL = "exec stp_ins_ColocCredNivAprTipoCredito " & "'" & cNivAprCod & "', '" & cTipoCreditoCod & "', '" & cAgeCod & "'"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
    Exit Sub
ControlError:
    Set loReg = Nothing
End Sub

Public Function ObtenerCorrelativoNivAprCod() As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_ObtenerCorrelativoNivAprCod"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerCorrelativoNivAprCod = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function ListarNivelesAprobacion(ByRef psMensaje As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ListarNivelesAprobacion"
    
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
    Set ListarNivelesAprobacion = lrDatos
    Set lrDatos = Nothing
End Function

Public Function ListarHistorialNivApr(ByVal cCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ListarHistorialNivApr " & "'" & cCtaCod & "'"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set ListarHistorialNivApr = lrDatos
    Set lrDatos = Nothing
End Function

Public Function CargarDatosPreApr(ByVal cCtaCod As String, ByVal cCargoCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_CargarDatosPreApr " & "'" & cCtaCod & "', '" & cCargoCod & "'"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarDatosPreApr = lrDatos
    Set lrDatos = Nothing
End Function

Public Sub ModificarEstadoPreApr(ByVal cCtaCod As String, ByVal sCargoCod As String, ByVal sComentario As String, ByVal nEstado As Integer, ByVal sMovNro As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_upd_ColocCredEstadoPreApr '" & cCtaCod & "', '" & sCargoCod & "', '" & sComentario & "', " & nEstado & ", '" & sMovNro & "'"
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Function CargarEncargaturaNivApr(ByVal cUserCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ColocCredEncargaturaNivApr '" & cUserCod & "'"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set CargarEncargaturaNivApr = lrDatos
    Set lrDatos = Nothing
End Function

'JUEZ 20121128 ************************************************************************
Public Function RecuperaGruposApr(Optional ByVal psIdGrupo As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaGruposAprobacion
    
    sSql = "exec stp_sel_RecuperaGruposAprobacion '" & psIdGrupo & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGruposApr = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaGruposAprobacion:
    Err.Raise Err.Number, "Recupera Grupos de Aprobacion ", Err.Description
End Function

Public Function RecuperaGruposAprDetalle(ByVal psIdGrupo As String, ByVal pnTipo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo RecuperaGruposAprDetalle
    
    sSql = "exec stp_sel_RecuperaGruposAprDetalle '" & psIdGrupo & "'," & pnTipo
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGruposAprDetalle = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

RecuperaGruposAprDetalle:
    Err.Raise Err.Number, "Recupera Grupos de Aprobacion Detalle", Err.Description
End Function

Public Sub dInsertaGruposApr(ByVal psGrupoDesc As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_GrabarGruposApr '" & psGrupoDesc & "'"
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Function RecuperaGrupoAprCod(ByVal psGrupoDesc As String) As String
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaGrupoAprCod
    
    sSql = "exec stp_sel_RecuperaGrupoAprCod '" & psGrupoDesc & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    If Not rs.EOF Then
        RecuperaGrupoAprCod = rs!cGrupoCod
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaGrupoAprCod:
    Err.Raise Err.Number, "Error en Recupera Cod Grupo", Err.Description
End Function

Public Sub dActualizaGruposApr(ByVal psGrupoCod As String, ByVal psGrupoDesc As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_upd_dActualizaGruposApr '" & psGrupoCod & "','" & psGrupoDesc & "'"
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dEliminaGruposAprDetalle(ByVal psGrupoCod As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_del_EliminaGruposAprDetalle '" & psGrupoCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dInsertaGruposAprDetalle(ByVal psGrupoCod As String, ByVal psAgeProdValor As String, ByVal psAgeProdCod As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_GrabarGruposAprDetalle '" & psGrupoCod & "','" & psAgeProdValor & "'," & psAgeProdCod
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dEliminaGrupoApr(ByVal psGrupoCod As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_upd_EliminaGrupoApr '" & psGrupoCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Function VerificaDatosGrupoApr(ByVal psGrupoCod As String, ByVal psGrupoDesc As String, ByVal psTpoProd As String, ByVal psAgencias As String) As Boolean
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    VerificaDatosGrupoApr = False
    lsSQL = "exec stp_sel_VerificaDatosGrupoApr '" & psGrupoCod & "','" & psGrupoDesc & "','" & psTpoProd & "','" & psAgencias & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    If Not lrDatos.EOF Then
        VerificaDatosGrupoApr = True
    End If
    Set lrDatos = Nothing
End Function

Public Function RecuperaNivApr(Optional ByVal psNivAprCod As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaNivApr
    
    sSql = "exec stp_sel_RecuperaNivelesAprobacion '" & psNivAprCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaNivApr = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaNivApr:
    Err.Raise Err.Number, "Recupera Niveles de Aprobacion ", Err.Description
End Function

Public Function VerificaDatosNivApr(ByVal psNivAprCod As String, ByVal psNivAprDesc As String) As Boolean
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    VerificaDatosNivApr = False
    lsSQL = "exec stp_sel_VerificaDatosNivApr '" & psNivAprCod & "','" & psNivAprDesc & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    If Not lrDatos.EOF Then
        VerificaDatosNivApr = True
    End If
    Set lrDatos = Nothing
End Function

Public Function dInsertaNivApr(ByVal psNivAprDesc As String, ByVal pnTipoReg As Integer, ByVal pnNumCantFirmas As Integer, _
                            ByVal pnCorrigeSug As Integer, ByVal pnValidaAg As Integer) As String
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_GrabarNivApr '" & psNivAprDesc & "'," & pnTipoReg & "," & pnNumCantFirmas & "," & pnCorrigeSug & "," & pnValidaAg
            
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.Ejecutar (lsSQL)
    Set lrDatos = oConec.CargaRecordSet("Select cNivAprCod from CredNivelesApr where cNivAprDesc= '" & psNivAprDesc & "' and nEstado=1")
    dInsertaNivApr = lrDatos.Fields(0)
    oConec.CierraConexion
    Set oConec = Nothing
    Set lrDatos = Nothing
End Function

Public Sub dActualizaNivApr(ByVal psNivAprCod As String, ByVal psNivAprDesc As String, ByVal pnTipoReg As Integer, _
                            ByVal pnNumCantFirmas As Integer, ByVal pnCorrigeSug As Integer, ByVal pnValidaAg As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_upd_dActualizaNivApr '" & psNivAprCod & "','" & psNivAprDesc & "'," & pnTipoReg & "," & _
                                                pnNumCantFirmas & "," & pnCorrigeSug & "," & pnValidaAg
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dEliminaNivAprValores(ByVal psNivAprCod As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_del_EliminaNivAprValores '" & psNivAprCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dInsertaNivAprValores(ByVal psNivAprCod As String, ByVal psValorCod As String, ByVal psTipoCod As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_GrabarNivAprValores '" & psNivAprCod & "','" & psValorCod & "','" & psTipoCod & "'"
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Function RecuperaNivAprValores(ByVal psNivAprCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaNivAprValores
    
    sSql = "exec stp_sel_RecuperaNivAprValores '" & psNivAprCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaNivAprValores = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaNivAprValores:
    Err.Raise Err.Number, "Recupera Niveles de Aprobacion Valores", Err.Description
End Function

Public Sub dEliminaNivApr(ByVal psNivAprCod As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_upd_EliminaNivApr '" & psNivAprCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Function ObtenerUltimoOrdenNivel() As Integer
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

    On Error GoTo ErrorObtenerUltimoOrdenNivel
    
    sSql = "exec stp_sel_ObtenerUltimoOrdenNivel"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    ObtenerUltimoOrdenNivel = rs!nOrden
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtenerUltimoOrdenNivel:
    Err.Raise Err.Number, "Obtiene Ultimo Orden Nivel ", Err.Description
End Function

Public Sub dActualizaOrdenNivel(ByVal psNivAprCodAct As String, ByVal psNivAprCodReemp As String, ByVal pnOrdenAct As Integer, ByVal pnOrdenReemp As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_upd_ActualizaOrdenNivel '" & psNivAprCodAct & "','" & psNivAprCodReemp & "'," & pnOrdenAct & "," & pnOrdenReemp
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Function RecuperaParametrosGrupoApr(ByVal psGrupoCod As String, ByVal pnTpoCliente As Integer) As ADODB.Recordset
                                    'RECO20150202 ERS173-2014 Se agregó el parametro: pnTpoCliente
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaParametrosGrupoApr
    
    'sSql = "exec stp_sel_RecuperaParametrosGrupoApr '" & psGrupoCod & "'"
    sSql = "exec stp_sel_RecuperaParametrosGrupoApr '" & psGrupoCod & "'," & pnTpoCliente 'RECO20150202 ERS173-2014
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaParametrosGrupoApr = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaParametrosGrupoApr:
    Err.Raise Err.Number, "Recupera Parámetros de Aprobacion ", Err.Description
End Function

Public Function RecuperaNivAprListaParam(ByVal pnTipo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaNivAprListaParam
    
    If pnTipo = 1 Then
        sSql = "exec stp_sel_RecuperaNivAprListaParam"
    ElseIf pnTipo = 2 Then
        sSql = "exec stp_sel_RecuperaNivAprListaCombo"
    End If

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaNivAprListaParam = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaNivAprListaParam:
    Err.Raise Err.Number, "Recupera Niveles de Aprobacion Lista", Err.Description
End Function

Public Function RecuperaParametrosGrupoNivelesApr(ByVal psParamCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaParametrosGrupoNivelesApr
    
    sSql = "exec stp_sel_RecuperaParametrosGrupoNivelesApr '" & psParamCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaParametrosGrupoNivelesApr = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaParametrosGrupoNivelesApr:
    Err.Raise Err.Number, "Recupera Parámetros de Aprobacion ", Err.Description
End Function

Public Sub dInsertaParamGruposNiveles(ByVal psGrupoCod As String, ByVal psNivelesDesc As String, ByVal pnDesdeR1 As Double, _
                                      ByVal pnHastaR1 As Double, ByVal pnDesdeR2 As Double, ByVal pnHastaR2 As Double, _
                                      Optional ByVal pnTpoCliente As Integer)
                                      'RECO20150129 ERS173-2014 Se agrego tipo de cliente
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_GrabarParamGruposNiveles '" & psGrupoCod & "'," & psNivelesDesc & "," & pnDesdeR1 & "," & pnHastaR1 & "," & _
                                                        pnDesdeR2 & "," & pnHastaR2 & "," & pnTpoCliente
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dEliminaParamGruposNiveles(ByVal psGrupoCod As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_upd_EliminaParamGruposNiveles '" & psGrupoCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dEliminaParamApr(ByVal psParamCod As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_upd_EliminaParamApr '" & psParamCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dActualizaParamGruposNiveles(ByVal psParamCod As String, ByVal psGrupoCod As String, ByVal psNivelesDesc As String, _
                                        ByVal pnDesdeR1 As Double, ByVal pnHastaR1 As Double, ByVal pnDesdeR2 As Double, ByVal pnHastaR2 As Double, _
                                        Optional ByVal pnTpoCliente As Integer)
                                        'RECO20150129 ERS173-2014 Se agrego tipo de cliente
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_upd_ActualizaParamGruposNiveles '" & psParamCod & "','" & psGrupoCod & "'," & psNivelesDesc & "," & _
                                                            pnDesdeR1 & "," & pnHastaR1 & "," & pnDesdeR2 & "," & pnHastaR2 & "," & pnTpoCliente
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Function VerificaUsuarioSiTieneNivel(ByVal psCodUser As String, Optional ByRef psNomUsu As String = "", _
                                            Optional ByRef psTipoNiv As String = "", Optional ByRef psNivAprCod As String = "") As Boolean
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    VerificaUsuarioSiTieneNivel = False
    lsSQL = "exec stp_sel_VerificaUsuarioSiTieneNivel '" & psCodUser & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    If Not lrDatos.EOF Then
        psNomUsu = lrDatos!cPersNombre
        psTipoNiv = lrDatos!cTipoReg
        psNivAprCod = lrDatos!cNivAprCod
        VerificaUsuarioSiTieneNivel = True
    End If
    Set lrDatos = Nothing
End Function

Public Function VerificaUsuarioSiTieneDelegacion(ByVal psCodUser As String) As Boolean
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    VerificaUsuarioSiTieneDelegacion = False
    
    lsSQL = "exec stp_sel_VerificaUsuarioSiTieneDelegacion '" & psCodUser & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    If Not lrDatos.EOF Then
        VerificaUsuarioSiTieneDelegacion = True
    End If
    Set lrDatos = Nothing
End Function

Public Function VerificaUsuarioSiTieneNivelDelegado(ByVal psCodUser As String) As Boolean
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    VerificaUsuarioSiTieneNivelDelegado = False
    
    lsSQL = "exec stp_sel_VerificaUsuarioSiTieneNivelDelegado '" & psCodUser & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    If Not lrDatos.EOF Then
        VerificaUsuarioSiTieneNivelDelegado = True
    End If
    Set lrDatos = Nothing
End Function

Public Sub dDelegarAprobacion(ByVal psUserDelega As String, ByVal psUserDelegado As String, ByVal pdFechaDesde As String, ByVal pdFechaHasta As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_DelegarAprobacion '" & psUserDelega & "','" & psUserDelegado & "','" & Format(CDate(Format(pdFechaDesde, "dd/mm/yyyy hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "','" & Format(CDate(Format(pdFechaHasta, "dd/mm/yyyy hh:mm:ss")), "mm/dd/yyyy hh:mm:ss") & "'"
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dRetirarDelegacionAprobacion(ByVal psUserDelega As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_RetirarDelegacionAprobacion '" & psUserDelega & "'"
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Function RecuperaDelegacion(ByVal psCodUser As String, ByVal pnTipoBusq As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDelegacion
    
    sSql = "exec stp_sel_RecuperaDelegacion '" & psCodUser & "'," & pnTipoBusq
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDelegacion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDelegacion:
    Err.Raise Err.Number, "Recupera Delegación ", Err.Description
End Function

'Public Function RecuperaTiposExoneraciones() As ADODB.Recordset
Public Function RecuperaTiposExoneraciones(Optional nTpOpcion As Integer = 0) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaTiposExoneraciones
    
    'sSql = "exec stp_sel_RecuperaTiposExoneraciones "
    sSql = "exec stp_sel_RecuperaTiposExoneraciones " & nTpOpcion & ""
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaTiposExoneraciones = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaTiposExoneraciones:
    Err.Raise Err.Number, "Recupera Tipos de Exoneraciones ", Err.Description
End Function

Public Sub dEliminaTipoExoneracion(Optional ByVal psExoneraCod As String = "")
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    If psExoneraCod <> "" Then
        lsSQL = "exec stp_upd_EliminaTipoExoneracion '" & psExoneraCod & "'"
    Else
        lsSQL = "exec stp_del_EliminaTipoExoneracion "
    End If
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dInsertaTiposExoneraciones(ByVal psExoneraDesc As String, ByVal pnTipoCod As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_GrabarTipoExoneracion '" & psExoneraDesc & "'," & pnTipoCod
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dActualizaTiposExoneraciones(ByVal psExoneraCod As String, ByVal psExoneraDesc As String, ByVal pnTipoCod As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_upd_ActualizaTipoExoneracion '" & psExoneraCod & "','" & psExoneraDesc & "'," & pnTipoCod
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Function RecuperaNivelesExoneracion(ByVal psNivExoneraCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaNivelesExoneracion
    
    sSql = "exec stp_sel_RecuperaNivelesExoneracion '" & psNivExoneraCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaNivelesExoneracion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaNivelesExoneracion:
    Err.Raise Err.Number, "Recupera Niveles de Exoneraciones ", Err.Description
End Function

Public Sub dEliminaNivelesExoneracion(ByVal psTipoExoneraCod As String, Optional ByVal psNivExoneraCod As String = "")
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    If psNivExoneraCod <> "" Then
        lsSQL = "exec stp_upd_EliminaNivelesExoneracion '" & psTipoExoneraCod & "','" & psNivExoneraCod & "'"
    Else
        lsSQL = "exec stp_del_EliminaNivelesExoneracion '" & psTipoExoneraCod & "'"
    End If
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Sub dInsertaNivelesExoneraciones(ByVal psTipoExoneraCod As String, ByVal psNivAprCod As String, ByVal pnDesde As Double, ByVal pnHasta As Double)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_GrabarNivelExoneracion '" & psTipoExoneraCod & "','" & psNivAprCod & "'," & pnDesde & "," & pnHasta
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Function ObtieneNivelAprobacionDelegacion(ByVal psCodUser As String) As String
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    ObtieneNivelAprobacionDelegacion = ""
    lsSQL = "exec stp_sel_ObtieneNroNivelAprobacionDelegacion '" & psCodUser & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    If Not lrDatos.EOF Then
        ObtieneNivelAprobacionDelegacion = lrDatos!cNivAprCod
    End If
    Set lrDatos = Nothing
End Function

Public Function RecuperaCreditosPorAprobar(ByVal psCodUser As String, ByVal pnTipo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaCreditosPorAprobar
    
    sSql = "exec stp_sel_RecuperaCreditosPorAprobar '" & psCodUser & "'," & pnTipo
   
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosPorAprobar = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCreditosPorAprobar:
    Err.Raise Err.Number, "Recupera Creditos Por Aprobar ", Err.Description
End Function

Public Function RecuperaDatosCredResultado(ByVal psCtaCod As String, Optional ByVal psNivAprCod As String = "", Optional ByVal psCodUser As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDatosCredResultado
    
    sSql = "exec stp_sel_RecuperaDatosCredResultado '" & psCtaCod & "','" & psNivAprCod & "','" & psCodUser & "'"
   
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosCredResultado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosCredResultado:
    Err.Raise Err.Number, "Recupera Datos Resultado Creditos", Err.Description
End Function

Public Function RecuperaDatosCredExoneracion(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaDatosCredExoneracion
    
    sSql = "exec stp_sel_RecuperaDatosCredExoneracion '" & psCtaCod & "'"
   
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosCredExoneracion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaDatosCredExoneracion:
    Err.Raise Err.Number, "Recupera Datos Exoneracion Creditos", Err.Description
End Function

Public Sub dSolicitaModifAprobacionResultado(ByVal psCtaCod As String, ByVal psComent As String, _
                Optional ByVal pnTipoReg As TipoNivelReg = TipoNivelReg.gTipoNivelAprob, Optional ByVal psMovNro As String = "") 'RECO20160628 ERS002-2016 SE AGREGO "pnTipoReg","psMovNro"
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    'lsSql = "exec stp_ins_SolicitaModifAprobacionResultado '" & psCtaCod & "','" & psComent & "'"
    lsSQL = "exec stp_ins_SolicitaModifAprobacionResultado '" & psCtaCod & "','" & psComent & "'," & pnTipoReg & ",'" & psMovNro & "'" 'RECO20160628 ERS002-2016
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub
Public Sub dDesactivaAprobacionResultado(ByVal psCtaCod As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_upd_DesactivaAprobacionResultado '" & psCtaCod & "'"
            
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
End Sub

Public Function RecuperaHistorialCredAprobados(ByVal psCtaCod As String, _
                Optional ByVal pnTpoClient As Integer) As ADODB.Recordset
                'RECO20150217 ERS173-2014 Agrego pnTpoClient
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaHistorialCredAprobados
    
    'sSql = "exec stp_sel_RecuperaHistorialCredAprobados '" & psCtaCod & "'"
    sSql = "exec stp_sel_RecuperaHistorialCredAprobados '" & psCtaCod & "'," & pnTpoClient
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaHistorialCredAprobados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaHistorialCredAprobados:
    Err.Raise Err.Number, "Recupera Grupos de Aprobacion ", Err.Description
End Function

Public Function ExisteAprobacionCredNivelesPendientes(ByVal psCtaCod As String, ByVal pnTipoCliente As Integer) As Boolean
    'pnTipoCliente : Se creó el parametro RECO20160419
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    ExisteAprobacionCredNivelesPendientes = False
    lsSQL = "exec stp_sel_VerificaAprobacionCredNivelesPendientes '" & psCtaCod & "'," & pnTipoCliente
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    If Not lrDatos.EOF Then
        ExisteAprobacionCredNivelesPendientes = True
    End If
    Set lrDatos = Nothing
End Function

Public Function ExisteRegistroResultadoUserNivApr(ByVal psCtaCod As String, ByVal psNivAprCod As String, ByVal psCodUser As String) As Boolean
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    ExisteRegistroResultadoUserNivApr = False
    lsSQL = "exec stp_sel_VerificaRegistroResultadoUserNivApr '" & psCtaCod & "','" & psNivAprCod & "','" & psCodUser & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    If Not lrDatos.EOF Then
        ExisteRegistroResultadoUserNivApr = True
    End If
    Set lrDatos = Nothing
End Function
'END JUEZ *****************************************************************************
'RECO20140225 ERS174-2013**************************************************************
Public Sub RegistraExoneraAutoriza(ByVal psCtaCod As String, ByVal psExoneraCod As String, ByVal pnEstado As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    On Error GoTo ErrorRegistraExoneraAutoriza
    
    lsSQL = "exec stp_ins_RegistraExoAut '" & psCtaCod & "','" & psExoneraCod & "'," & pnEstado
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
    Exit Sub
ErrorRegistraExoneraAutoriza:
    Err.Raise Err.Number, "Registra Exoneraciones/Autorizaciones", Err.Description
End Sub

Public Sub EliminaRegExoneraAutoriza(ByVal psCtaCod As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    On Error GoTo ErrorEliminaRegExoneraAutoriza
    
    lsSQL = "exec stp_del_EliminaRegExoneraAutorizaCred '" & psCtaCod & "'"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
    Exit Sub
ErrorEliminaRegExoneraAutoriza:
    Err.Raise Err.Number, "Elimana el registro Exoneraciones/Autorizaciones", Err.Description
End Sub

Public Function ObtieneDatosExoAutCred(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
        
    lsSQL = "exec stp_sel_ObtieneDatosExoAutCred '" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set ObtieneDatosExoAutCred = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
End Function

Public Function ObtieneExoneraAutoriCred(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
        
    lsSQL = "exec stp_sel_ObtieneExoneraAutoriCred '" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set ObtieneExoneraAutoriCred = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
End Function
Public Function ObtineRiesgoUnicoCred(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
        
    lsSQL = "exec stp_sel_ObtineRiesgoUnicoCred '" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set ObtineRiesgoUnicoCred = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
End Function

Public Function ObtieneNivelAprResultado(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
        
    lsSQL = "exec stp_sel_ObtieneNivelAprResultado '" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set ObtieneNivelAprResultado = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
End Function

Public Function RecuperaRespExoAut(ByVal psNivExo As String, ByVal pnMonto As Double) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
        
    lsSQL = "exec stp_sel_RecuperaRespExoAut '" & psNivExo & "'," & pnMonto
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set RecuperaRespExoAut = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
End Function

Public Function RecuperaTpoExoneraciones() As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
        
    lsSQL = "exec stp_sel_RecuperaTpoExoneraciones "
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set RecuperaTpoExoneraciones = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
End Function

Public Function RecuperaTpoAutorizaciones() As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
        
    lsSQL = "exec stp_sel_RecuperaTpoAutorizaciones "
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set RecuperaTpoAutorizaciones = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
End Function

Public Function RecuperaNivelesAprValores(ByVal psNiv As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
        
    lsSQL = "exec stp_sel_RecuperaNivelesAprValores '" & psNiv & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set RecuperaNivelesAprValores = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
End Function

Public Function RecuperaCargosNivelesExoneracion(ByVal psNivExo As String, psNivAprCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
        
    lsSQL = "exec stp_sel_RecuperaCargosNivelesExoneracion '" & psNivExo & "','" & psNivAprCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
    Set RecuperaCargosNivelesExoneracion = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
End Function

Public Sub RegistraExoneraAdm(ByVal psCtaCod As String, ByVal psExoneraCod As String, ByVal psNivExoCod As String, ByVal psPersCod As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    On Error GoTo ErrorRegistraExoneraAdm
    
    lsSQL = "exec stp_ins_RegistraExoneraAdm '" & psCtaCod & "','" & psExoneraCod & "','" & psNivExoCod & "','" & psPersCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    Set lrDatos = Nothing
    Exit Sub
ErrorRegistraExoneraAdm:
    Err.Raise Err.Number, "Registra Exoneración Adm Créditos", Err.Description
End Sub
'RECO FIN******************************************************************************
'ALPA 20140321**************************************************************************************************************
Public Function VerificaDatosNivAprCV(ByVal psCargoCod As String, ByVal psCargoDesc As String, ByVal psNivAprCod As String) As Boolean
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    VerificaDatosNivAprCV = False
    lsSQL = "exec stp_sel_VerificaDatosNivAprCV '" & psCargoCod & "','" & psCargoDesc & "','" & psNivAprCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
    If Not lrDatos.EOF Then
        VerificaDatosNivAprCV = True
    End If
    oConec.CierraConexion
    Set oConec = Nothing
    Set lrDatos = Nothing
End Function

Public Function dInsertaNivAprCV(ByVal psNivAprCod As String, ByVal psCargoCod As String, ByVal pnTipoReg As String, ByVal pnNumCantFirmas As Integer, _
                                 ByVal pnValidaAg As Integer, ByVal pnMontoDesde As Currency, ByVal pnMontoHasta As Currency, _
                                 ByVal pnTCCmas As Currency, ByVal pnTCVmas As Currency) As String
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

    lsSQL = "exec stp_ins_NivelesAprobacionCompraVenta '" & psNivAprCod & "','" & psCargoCod & "'," & pnTipoReg & "," & pnNumCantFirmas & "," & pnValidaAg & "," & pnMontoDesde & ", " & pnMontoHasta & ", " & pnTCCmas & ", " & pnTCVmas
            
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.Ejecutar (lsSQL)
    Set lrDatos = oConec.CargaRecordSet("Select cRHCargos from NivelesAprobacionCompraVenta where cRHCargos= '" & psCargoCod & "' and nEstado=1 and cNivelCod='" & psNivAprCod & "'")
    dInsertaNivAprCV = lrDatos.Fields(0)
    oConec.CierraConexion
    Set oConec = Nothing
    Set lrDatos = Nothing
End Function
Public Function RecuperaNivAprCV(Optional ByVal psNivAprCod As String = "", Optional ByVal psRHCargosCod As String = "") As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaNivApr
    
    sSql = "exec stp_sel_NivelesAprobacionCompraVentaCV '" & psNivAprCod & "','" & psRHCargosCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaNivAprCV = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaNivApr:
    Err.Raise Err.Number, "Recupera Niveles de Aprobacion ", Err.Description
End Function
Public Function RecuperaNivAprValoresCV(ByVal psNivAprCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaNivAprValores
    
    sSql = "exec stp_sel_RecuperaNivAprValoresCV '" & psNivAprCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaNivAprValoresCV = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaNivAprValores:
    Err.Raise Err.Number, "Recupera Niveles de Aprobacion Valores", Err.Description
End Function
Public Function ObtenerNivelesAprobacionCompraVentaTabla(ByVal psNivAprCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    

    lsSQL = "exec stp_sel_NivelesAprobacionCompraVentaTabla '" & psNivAprCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerNivelesAprobacionCompraVentaTabla = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    Set lrDatos = Nothing
End Function
Public Function EliminarNivelesAprobacionCompraVenta(ByVal psNivAprCod As String, ByVal psCargoCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    

    lsSQL = "exec stp_del_NivelesAprobacionCompraVenta '" & psNivAprCod & "','" & psCargoCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set EliminarNivelesAprobacionCompraVenta = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    Set lrDatos = Nothing
End Function
Public Sub AprobacionMovCompraVenta(ByVal cMovnro As String, ByVal cOpeCod As String, ByVal cNivelCod As String, ByVal nMonto As Currency, ByVal nTipoCambioNormal As Currency, ByVal nTipoCambioSolicitado As Currency, ByVal nTipoEspecial As Integer, ByVal psPersCod As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    lsSQL = "exec stp_sel_AprobacionMovCompraVenta '" & cMovnro & "','" & cOpeCod & "','" & cNivelCod & "','" & nMonto & "'," & nTipoCambioNormal & "," & nTipoCambioSolicitado & "," & nTipoEspecial & ",'" & psPersCod & "'"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.Ejecutar (lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub
Public Function ObtenerNivelesAprobacionCompraVentaxMonto(ByVal pnMonto As Currency, ByVal nTipoEspecial As Integer) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    

    lsSQL = "exec stp_sel_ObtenerNivelesAprobacionCompraVentaxMonto '" & pnMonto & "'," & nTipoEspecial & ""
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerNivelesAprobacionCompraVentaxMonto = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerCantidadAprobacionMovCompraVenta(ByVal psMovNro As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    

    lsSQL = "exec stp_sel_ObtenerCantidadAprobacionMovCompraVenta '" & psMovNro & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerCantidadAprobacionMovCompraVenta = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ActualizarAprobacionMovCompraVenta(ByVal psMovNro As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta

    lsSQL = "exec stp_upd_AprobacionMovCompraVenta '" & psMovNro & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.Ejecutar (lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerAprobacionMovCompraVentaPendiente(ByVal cFecha As String, ByVal psRHCargosCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    

    lsSQL = "exec stp_sel_MostrarAprobacionMovCompraVentaPendiente '" & cFecha & "','" & psRHCargosCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerAprobacionMovCompraVentaPendiente = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Sub AprobacionMovCompraVentaDetalle(ByVal cMovnro As String, ByVal cOpeCod As String, ByVal cNivelCod As String, ByVal cMovNroAprobacion As String, ByVal cRHCargosCod As String, ByVal nEstado As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Set oConec = New COMConecta.DCOMConecta
    lsSQL = "exec stp_ins_AprobacionMovCompraVentaDetalle '" & cMovnro & "','" & cOpeCod & "','" & cNivelCod & "','" & cMovNroAprobacion & "','" & cRHCargosCod & "'," & nEstado
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.Ejecutar (lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub

Public Function ObtenerAprobacionMovCompraVentaPendientexCliente(ByVal cFecha As String, Optional ByVal psPersCod As String = "", Optional ByVal lsOpeCod As String = "") As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    

    lsSQL = "exec stp_sel_MostrarAprobacionMovCompraVentaPendientexCliente '" & cFecha & "','" & psPersCod & "','" & lsOpeCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerAprobacionMovCompraVentaPendientexCliente = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function AprobacionMovCompraVentaxMovimiento(ByVal cMovnro As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    

    lsSQL = "exec stp_sel_AprobacionMovCompraVentaxMovimiento '" & cMovnro & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set AprobacionMovCompraVentaxMovimiento = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'****************************************************************************************************************************
'ALPA 20141201***************************************************************************************************************
Public Function GetCargarColocacPermisoAprobacion(ByVal psFecha As String, ByVal psCodCargo As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSQL = "exec stp_sel_ColocacPermisoAprobacion '" & psFecha & "','" & psCodCargo & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GetCargarColocacPermisoAprobacion = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Sub GetActualizarColocacPermisoAprobacion(ByVal psCtaCod As String, ByVal psMovNroSol As String, ByVal psMovNroApro As String, ByVal psPersCodSol As String, ByVal psPersCodApro As String, ByVal psCodCargo As String, ByVal pnTasaSol As Currency, ByVal pnTasaApr As Currency, ByVal pnEstado As Integer, ByVal pnTipo As Integer)
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSQL = "exec stp_ins_ColocacPermisoAprobacion '" & psCtaCod & "','" & psMovNroSol & "', '" & psMovNroApro & "', '" & psPersCodSol & "', '" & psPersCodApro & "', '" & psCodCargo & "', " & pnTasaSol & ", " & pnTasaApr & ", " & pnEstado & ", " & pnTipo
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub
'MACM INICIO
Public Sub GetActualizarColocacPermisoEstado(ByVal psCtaCod As String, ByVal pnEstado As Integer)
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSQL = "exec stp_upd_ColocacPermisoEstado '" & psCtaCod & "', " & pnEstado
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub
Public Function GetCargarColocacPignoAprobacionPersona(ByVal psCodCred As String, ByVal pnEstado As Integer) As Boolean
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    Dim rsVar As Recordset
    Set oConecta = New COMConecta.DCOMConecta
    If oConecta.AbreConexion = False Then Exit Function
    lsSQL = "exec stp_set_ColocacPignoAprobacion '" & psCodCred & "', " & pnEstado
    Set rsVar = oConecta.CargaRecordSet(lsSQL)
    If rsVar!solicitudes = 0 Then
        GetCargarColocacPignoAprobacionPersona = False
    Else
        GetCargarColocacPignoAprobacionPersona = True
    End If
    
'    If Not rsVar.EOF And Not rsVar.BOF Then
'        GetCargarColocacPignoAprobacionPersona = rsVar!solicitudes
'    End If
    rsVar.Close
    Set rsVar = Nothing
    Set oConecta = Nothing
End Function
Public Function GetCargarColocacPermisoAprobacionCuenta(ByVal psCtaCod As String, ByVal pnEstado As Integer)
    Dim rsVar As Recordset
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
        
    If oConecta.AbreConexion = False Then Exit Function
    lsSQL = "exec stp_sel_ColocacPermisoAprobacionCuenta '" & psCtaCod & "', " & pnEstado
    Set rsVar = New Recordset
    Set rsVar = oConecta.CargaRecordSet(lsSQL)
    GetCargarColocacPermisoAprobacionCuenta = ""
    If Not rsVar.EOF And Not rsVar.BOF Then
        GetCargarColocacPermisoAprobacionCuenta = rsVar("nSolicitudes")
    End If
    rsVar.Close
    Set rsVar = Nothing
    Set oConecta = Nothing
    'Set oConecta = New COMConecta.DCOMConecta
    'oConecta.AbreConexion
    'Set GetCargarColocacPermisoAprobacionCuenta = oConecta.CargaRecordSet(lsSQL)
    'oConecta.CierraConexion
    'Set oConecta = Nothing
End Function
Public Function RecuperaSolTasaPendienteCuenta(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSQL = "exec stp_sel_SolTasaPendienteCuenta '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSolTasaPendienteCuenta = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
Public Function RecuperaSolPendientesAge(ByVal gsCodAge As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSQL = "exec stp_sel_SolTasaPendienteAge '" & gsCodAge & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaSolPendientesAge = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'MACM FIN
Public Sub ActualizaTasaAprobadaPigno(ByVal psCtaCod As String, ByVal pnTasaSol As Currency)
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSQL = "Exec stp_upd_CambiarTasaCredPigno '" & psCtaCod & "'," & pnTasaSol
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub
Public Function GetCargarColocacPermisoAprobacionConsulta(ByVal psCodCred As String, ByVal psFecha As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSQL = "exec stp_sel_ColocacPermisoAprobacionConsulta '" & psCodCred & "','" & psFecha & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GetCargarColocacPermisoAprobacionConsulta = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'*************************************************************************
'JUEZ 20160218 ********************************************************************
Public Function RecuperaNivelesUsuario(ByVal psCodUser As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSQL = "exec stp_sel_VerificaUsuarioSiTieneNivel '" & psCodUser & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaNivelesUsuario = oConecta.CargaRecordSet(lsSQL)
    
    If RecuperaNivelesUsuario.EOF Or RecuperaNivelesUsuario.BOF Then
        lsSQL = "exec stp_sel_VerificaUsuarioSiTieneNivelDelegado '" & psCodUser & "'"
        Set RecuperaNivelesUsuario = oConecta.CargaRecordSet(lsSQL)
    End If
    
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function RecuperaNivelAprUsuarioCredito(ByVal psCodUser As String, ByVal psCtaCod As String) As String
Dim oConecta As COMConecta.DCOMConecta
Dim lsSQL As String
Dim rs As ADODB.Recordset
    lsSQL = "exec stp_sel_RecuperaNivelAprUsuarioCredito '" & psCodUser & "','" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(lsSQL)
    If Not rs.EOF And Not rs.BOF Then
        RecuperaNivelAprUsuarioCredito = rs!cNivAprCod
    Else
        RecuperaNivelAprUsuarioCredito = ""
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'END JUEZ *************************************************************************
'EJVG20160526 ***
Public Function VerificarInformeRiesgoXNivApr(ByVal psCtaCod As String) As Boolean
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    Dim rs As ADODB.Recordset
    lsSQL = "SELECT bVerificar = dbo.fnc_ERS0282016_VerificarInformeRiesgosXNivApr ('" & psCtaCod & "')"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(lsSQL)
    If Not rs.EOF And Not rs.BOF Then
        VerificarInformeRiesgoXNivApr = rs!bVerificar
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'END EJVG *******
'EJVG20160204 ERS002-2016***
Public Function CargarTiposAutorizacionxConfig() As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrCargar
    lsSQL = "EXEC stp_sel_ERS0022016_TiposAutorizacionxConfig"
    Set CargarTiposAutorizacionxConfig = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrCargar:
    Err.Raise Err.Number, "Error en proceso (CargarTiposAutorizacionxConfig)", Err.Description
End Function
Public Function CargarAutorizacionEscalonaConfig() As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrCargar
    lsSQL = "EXEC stp_sel_ERS0022016_AutorizacionEscalonaConfig"
    Set CargarAutorizacionEscalonaConfig = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrCargar:
    Err.Raise Err.Number, "Error en proceso (CargarAutorizacionEscalonaConfig)", Err.Description
End Function
Public Sub dActualizarTipoAutorizacion(ByVal psAutorizaCod As String, ByVal pnEstado As Integer, ByVal psUltimaActualizacion As String)
    Dim lsSQL As String
    On Error GoTo Error
    lsSQL = "EXEC stp_upd_ERS0022016_TipoAutorizacion '" & psAutorizaCod & "'," & pnEstado & ",'" & psUltimaActualizacion & "'"
    coConex.Ejecutar lsSQL
    Exit Sub
Error:
    Err.Raise Err.Number, "Error en proceso (dInsertCredVinculados)", Err.Description
End Sub
Public Sub dActualizarEscalonaConfig(ByVal pnMoraPorcentaje As Double, ByVal pnMontoCuoMayorA As Double, ByVal pnMontoCuoMenorIgual As Double, ByVal pnMontoCreMayorA As Double, ByVal pnMontoCreMenorIgual As Double)
    Dim lsSQL As String
    On Error GoTo Error
    lsSQL = "EXEC stp_upd_ERS0022016_AutorizacionEscalonaConfig " & pnMoraPorcentaje & "," & pnMontoCuoMayorA & "," & pnMontoCuoMenorIgual & "," & pnMontoCreMayorA & "," & pnMontoCreMenorIgual
    coConex.Ejecutar lsSQL
    Exit Sub
Error:
    Err.Raise Err.Number, "Error en proceso (dActualizarEscalonaConfig)", Err.Description
End Sub
Public Function CargarAutorizacionesRequeridas(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrCargar
    lsSQL = "EXEC stp_sel_ERS0022016_AutorizacionesRequeridas '" & psCtaCod & "'"
    Set CargarAutorizacionesRequeridas = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrCargar:
    Err.Raise Err.Number, "Error en proceso (CargarTiposAutorizacionxConfig)", Err.Description
End Function
'END EJVG *******

'FRHU 20160704 ERS002-2016
Public Function HistAutorizacionesExoneraciones(ByVal psCtaCod As String, ByVal pnTipo As Integer) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrCargar
    lsSQL = "EXEC stp_sel_ERS0022016_HistAutorizacionExoneraciones '" & psCtaCod & "'," & pnTipo
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set HistAutorizacionesExoneraciones = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrCargar:
    Err.Raise Err.Number, "Error en proceso (HistAutorizacionesExoneraciones)", Err.Description
End Function
'FIN FRHU 20160704
'RECO20160526 ERS0022016************************************************************
Public Function ListaCredNivExoneracionCta(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    lsSQL = "exec stp_sel_ERS0022016_ListaCredNivExoneracionCta '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ListaCredNivExoneracionCta = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function RegistrarCredNivExoneraCabecera(ByVal psCtaCod As String, ByVal pnEstado As Integer, ByVal psMovNroReg As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim oDR As New ADODB.Recordset
    Dim lsSQL As String
    
    lsSQL = "exec stp_ins_sel_ERS0022016_RegistrarCredNivExoneraCabecera '" & psCtaCod & "'," & pnEstado & ",'" & psMovNroReg & " '"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RegistrarCredNivExoneraCabecera = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Sub RegistrarCredNivExoneraDetalle(ByVal pnId As Long, ByVal pnItem As Integer, ByVal psExoneracion As String, ByVal psDescripcion As String, ByVal psNivAprCod As String _
                                       , ByVal pbManual As Integer, ByVal pnEstado As Integer, ByVal psMovNroReg As String)
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSQL = "Exec stp_ins_ERS0022016_RegistrarCredNivExoneraDetalle " & pnId & "," & pnItem & ",'" & psExoneracion & "','" & psDescripcion & "','" & psNivAprCod & "'," _
            & pbManual & "," & pnEstado & ",'" & psMovNroReg & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub
Public Function ObtieneAutorizacionesPendientes(ByVal psUser As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    lsSQL = "exec stp_sel_ERS0022016_ObtieneAutorizacionesPendientes '" & psUser & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneAutorizacionesPendientes = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function ObtieneResultadoAutoExoNiv(ByVal psCtaCod As String, ByVal psUser As String, ByVal gsCodAge As String) As ADODB.Recordset 'ARLO 20170613 AGREGO PARAMETRO gsCodAge
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    lsSQL = "exec stp_sel_ERS0022016_ObtieneResultadoAutoExoNiv '" & psCtaCod & "','" & psUser & "','" & gsCodAge & "'" 'ARLO 20170613 AGREGO PARAMETRO gsCodAge
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneResultadoAutoExoNiv = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Sub ActualizaResultadoAutoExo(ByVal psCtaCod As String, ByVal psCodigo As String, ByVal psNivCod As String, ByVal pnItem As Integer, ByVal psMovNro As String, _
                                     ByVal pnEstado As Integer, ByVal psGlosa As String, ByVal pnTipoReg As Integer)
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSQL = "Exec stp_upd_ERS0022016_ActualizaResultadoAutoExo '" & psCtaCod & "','" & psCodigo & "','" & psNivCod & "'," & pnItem & ",'" & psMovNro & "'," & pnEstado & ",'" & _
                                                                psGlosa & "'," & pnTipoReg
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub
Public Sub ActualizaEstadoAutoExoCab(ByVal psCtaCod As String, ByVal pnEstado As Integer, ByVal psMovNro As String)
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSQL = "Exec stp_sel_ERS0022016_ActualizaEstadoAutoExoCab '" & psCtaCod & "'," & pnEstado & ",'" & psMovNro & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub
Public Sub RegistroAutorizacionManual(ByVal psCtaCod As String, ByVal psAutorizaCod As String, ByVal psNivAprCod As String, ByVal pnEstado As Integer, ByVal psGlosa As String, _
                                      ByVal psMovNroReg As String, ByVal pnPrdEstado As Integer)
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSQL = "Exec stp_ins_ERS0022016_RegistroAutorizacionManual '" & psCtaCod & "','" & psAutorizaCod & "','" & psNivAprCod & "'," & pnEstado & ",'" & psGlosa & "','" & _
                                                                psMovNroReg & "'," & pnPrdEstado
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub
Public Function ObtieneDatosNivelAutoCta(ByVal psCtaCod As String, ByVal psAutorizaCod As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim lsSQL As String
    
    lsSQL = "exec stp_sel_ERS0022016_ObtieneDatosNivelAutoCta '" & psCtaCod & "','" & psAutorizaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneDatosNivelAutoCta = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'RECO FIN***************************************************************************
'FRHU 20160802 ERS002-2016
Public Function ValidarExisteNivelDeAprobacion(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrCargar
    lsSQL = "EXEC stp_sel_ERS0022016_ValidarExisteNivelAprobacion '" & psCtaCod & "'"
    Set ValidarExisteNivelDeAprobacion = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrCargar:
    Err.Raise Err.Number, "Error en proceso (DCOMNivelAprobacion: ValidarExisteNivelDeAprobacion)", Err.Description
End Function
'FIN FRHU 20160802
'ARLO 20170627 ERS0652016 INICIO
Public Function ValidarSiExiteAutorizacion(ByVal psCtaCod As String, ByVal psCodigo As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrCargar
    lsSQL = "Exec stp_sel_ERS0652016_ValidarExiteAutorizacion '" & psCtaCod & "','" & psCodigo & "'"
    Set ValidarSiExiteAutorizacion = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrCargar:
    Err.Raise Err.Number, "Error en proceso (DCOMNivelAprobacion: ValidarSiExiteAutorizacion)", Err.Description
End Function
'FIN********************

'JOEP20171117 MEJORA PARA EXONERACION
Public Sub EliminaCredNivExonera(ByVal pcCtaCod As String, ByVal pnItem As Integer)
    Dim lsSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSQL = "Exec stp_del_EliminarNivExoneracion '" & pcCtaCod & "'," & pnItem & ""
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub
'JOEP20171117 MEJORA PARA EXONERACION

Public Function ObtieneAutoNoContempladas() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorObtieneAutoNoContempladas
    
    sSql = "exec stp_sel_ObtieneAutorizacionesNoContempladas "
   
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneAutoNoContempladas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorObtieneAutoNoContempladas:
    Err.Raise Err.Number, "Obtiene Autorizaciones No Contempladas", Err.Description
End Function
