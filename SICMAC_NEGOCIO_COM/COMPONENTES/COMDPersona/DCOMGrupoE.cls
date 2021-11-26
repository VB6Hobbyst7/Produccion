VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMGrupoE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

Public Function RegistraPersGrupoE(ByVal pnGrupoCod As Integer, ByVal pcPersCod As String, _
                                ByVal pnTipoVinculacion As Integer, ByVal pcPersCodOtro As String, _
                                ByVal pnPorcenOtro As Double, ByVal pcCargoOtro As String)
    
    If pcPersCodOtro = "" Then
        sSql = "INSERT INTO PersGrupoEconomico(nGrupoCod, cPersCod) " & _
           "VALUES(" & pnGrupoCod & ",'" & pcPersCod & "')"
    Else
        sSql = "INSERT INTO PersGrupoEconomico(nGrupoCod, cPersCod, nTipoVinculacion, cPersCodOtro, nPorcenOtro, cCargoOtro) " & _
           " VALUES(" & pnGrupoCod & ",'" & pcPersCod & "'," & pnTipoVinculacion & ",'" & pcPersCodOtro & "'," & pnPorcenOtro & ",'" & pcCargoOtro & "')"
    End If
    oCon.Ejecutar sSql
End Function

Public Function ModificaPersGrupoE(ByVal pnGrupoCod As Integer, ByVal pcPersCod As String, _
                                ByVal pnTipoVinculacion As Integer, ByVal pcPersCodOtro As String, _
                                ByVal pnPorcenOtro As Double, ByVal pcCargoOtro As String)
                                
    sSql = "UPDATE PersGrupoEconomico  " & _
        " SET nTipoVinculacion=" & pnTipoVinculacion & _
        " , cPersCodOtro='" & pcPersCodOtro & "'" & _
        " , nPorcenOtro=" & pnPorcenOtro & _
        " , cCargoOtro='" & pcCargoOtro & "'" & _
        " WHERE cPersCod='" & pcPersCod & "' AND nGrupoCod=" & pnGrupoCod
    oCon.Ejecutar sSql
End Function

Public Function EliminaPersGrupoE(ByVal pcPersCod As String, ByVal pnGrupoE As Integer)
    sSql = "DELETE FROM PersGrupoEconomico WHERE cPersCod='" & pcPersCod & "' AND nGrupoCod=" & pnGrupoE
    oCon.Ejecutar sSql
End Function

Public Function BuscarPersonasXGrupoEcon(ByVal pnGrupoE As Integer) As ADODB.Recordset

sSql = "SELECT P.cPersCod,P.cPersNombre FROM PersGrupoEconomico PG INNER JOIN Persona P ON PG.cPersCod = P.cPersCod" & _
        " WHERE PG.nGrupoCod=" & pnGrupoE
Set BuscarPersonasXGrupoEcon = oCon.CargaRecordSet(sSql)

End Function

Public Function CargaDatos(ByVal pcPersCod As String, ByVal pnGrupoE As Integer) As ADODB.Recordset

sSql = "SELECT nGrupoCod,cPersCodCli=P.cPersCod,cPersNombreCli=P.cPersNombre , " & _
        " nTipoVinculacion=ISNULL(nTipoVinculacion,0),cPersCodOtro=ISNULL(cPersCodOtro,''),cPersNombreOtro=ISNULL(OTRO.cPersNombre,'')," & _
        " nPorcenOtro=ISNULL(nPorcenOtro,0),cCargoOtro=ISNULL(cCargoOtro,'')" & _
        " FROM PersGrupoEconomico PG " & _
        " INNER JOIN Persona P ON PG.cPersCod = P.cPersCod" & _
        " LEFT JOIN Persona OTRO ON PG.cPersCodOtro=OTRO.cPersCod " & _
        "WHERE PG.nGrupoCod=" & pnGrupoE & " AND PG.cPersCod='" & pcPersCod & "'"

Set CargaDatos = oCon.CargaRecordSet(sSql)

End Function

Private Sub Class_Initialize()
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
End Sub

Private Sub Class_Terminate()
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'ALPA 20090824******************************************************
Public Function ObtenerPersonasXGrupoEcon(ByVal pnGrupoE As Integer) As ADODB.Recordset
    sSql = "exec stp_sel_PersGrupoEconomico " & pnGrupoE
    Set ObtenerPersonasXGrupoEcon = oCon.CargaRecordSet(sSql)
End Function
'PTI1 17/08/2018 INFORME N° 002-2017 AC TI / CMACM******************************************************
Public Function ObtenerSoloPersonasXGrupoEcon(ByVal pnGrupoE As Integer) As ADODB.Recordset
    'SE DEJO DE USAR ObtenerPersonasXGrupoEcon
    sSql = "exec stp_sel_SoloPersGrupoEconomico " & pnGrupoE
    Set ObtenerSoloPersonasXGrupoEcon = oCon.CargaRecordSet(sSql)
End Function

'ADD PTI1 17/08/2018 INFORME N° 002-2017 AC TI / CMACM******************************************************
Public Function ObtenerSoloVinculadoxEmpresa(ByVal pnGrupoE As Integer, cPercod As String) As ADODB.Recordset
    sSql = "exec stp_sel_SoloPersxGrupoEconomico " & pnGrupoE & " , '" & cPercod & "'"
    Set ObtenerSoloVinculadoxEmpresa = oCon.CargaRecordSet(sSql)
End Function
'fin pti1

Public Function ObtenerRelacionVinculadosxPersona(ByVal pnGrupoE As Integer) As ADODB.Recordset
    sSql = "exec stp_sel_RelacionGestionVinculado " & pnGrupoE
    Set ObtenerRelacionVinculadosxPersona = oCon.CargaRecordSet(sSql)
End Function

Public Function ObtenerRelacionGestion() As ADODB.Recordset
    sSql = "exec stp_sel_RelacionGestion "
    Set ObtenerRelacionGestion = oCon.CargaRecordSet(sSql)
End Function
Public Function GrabarPersGrupoEconomico(nGrupoCod As Integer, sPersCod As String, sPersCodOtro As String, nRepresentanteLegal As Integer, nTipoVinculacion As Integer, nPorcenOtro As Currency, nCargo As Integer, nCargoOtro As Integer, sComentarios As String, nEstado As Integer, sPropiedadDirecta As String, sPropiedadIndirecta As String, nForzar As Integer, dfecha As Date) As ADODB.Recordset
    sSql = "exec stp_ins_PersGrupoEconomico  " & nGrupoCod & " , '" & sPersCod & "', '" & sPersCodOtro & "', " & nRepresentanteLegal & ", " & nTipoVinculacion & ", " & nPorcenOtro & ", " & nCargo & ", " & nCargoOtro & ", '" & sComentarios & "', " & nEstado & ", '" & sPropiedadDirecta & "', '" & sPropiedadIndirecta & "', " & nForzar & ",'" & Format(dfecha, "YYYY/MM/DD") & "'"
    Set GrabarPersGrupoEconomico = oCon.CargaRecordSet(sSql)
End Function
'EJVG20120419
Public Function TienePermisoAdicionarPersonaGrupoEco(ByVal nGrupoCod As Integer, ByVal sGrupoUsu As String) As ADODB.Recordset
    sSql = "exec stp_sel_PermisoAddPersonasGrupoEco " & nGrupoCod & ",'" & sGrupoUsu & "'"
    Set TienePermisoAdicionarPersonaGrupoEco = oCon.CargaRecordSet(sSql)
End Function
Public Function ObtenerPersGrupoEconomico(nGrupoCod As Integer, sPersCod As String, sPersCodOtro As String) As ADODB.Recordset
    sSql = "exec stp_sel_ObtenerPersGrupoEconomico  " & nGrupoCod & " , '" & sPersCod & "', '" & sPersCodOtro & "' "
    Set ObtenerPersGrupoEconomico = oCon.CargaRecordSet(sSql)
End Function
Public Function ActualizarPersGrupoEconomicoCargos(nGrupoCod As Integer, sPersCod As String, sPersCodOtro As String, nCargo As Integer, nOtroCargo As Integer, nPorcentaje As Integer) As ADODB.Recordset
    sSql = "exec stp_upd_PersGrupoEconomicoCargos  " & nGrupoCod & " , '" & sPersCod & "', '" & sPersCodOtro & "', " & nCargo & ", " & nOtroCargo & ", " & nPorcentaje
    Set ActualizarPersGrupoEconomicoCargos = oCon.CargaRecordSet(sSql)
End Function
Public Function ActualizarRelacionGestionVinculado(sRelacionGestion As String, sPersCod As String, sPersCodVinc As String, nGrupo As Integer, nEstado As Integer, nContador As Integer) As Integer
    sSql = "exec stp_ins_RelacionGestionVinculado  '" & sRelacionGestion & "','" & sPersCod & "', '" & sPersCodVinc & "', " & nGrupo & ", " & nEstado & "," & nContador
    oCon.CargaRecordSet (sSql)
    ActualizarRelacionGestionVinculado = 1
End Function
Public Function ListarPersGrupoEconomicoGestion(nGrupoCod As Integer, sPersCod As String, sPersCodOtro As String) As ADODB.Recordset
    sSql = "exec stp_sel_RelacionGestionVinculadoListar  " & nGrupoCod & " , '" & sPersCod & "', '" & sPersCodOtro & "' "
    Set ListarPersGrupoEconomicoGestion = oCon.CargaRecordSet(sSql)
End Function
Public Function ListarGrupoEconomico(ByVal nTipoConsulta As Integer) As ADODB.Recordset
    sSql = "exec stp_sel_GrupoEconomico  " & nTipoConsulta
    Set ListarGrupoEconomico = oCon.CargaRecordSet(sSql)
End Function
Public Function ActualizarGrupoEconomico(nGrupoCod As Integer, sGrupoDes As String, nEstado As Integer) As Integer
    sSql = "exec stp_ins_GrupoEconomico  " & nGrupoCod & ",'" & sGrupoDes & "', " & nEstado & ""
    oCon.CargaRecordSet (sSql)
    ActualizarGrupoEconomico = 1
End Function
Public Function ListarDatosGrupoEconomicoxGrupoRepo19y20(nGrupoCod As Integer) As ADODB.Recordset
    sSql = "exec stp_sel_Reporte19y20GrupoEconomico  " & nGrupoCod
    Set ListarDatosGrupoEconomicoxGrupoRepo19y20 = oCon.CargaRecordSet(sSql)
End Function
Public Function ListarDatosGrupoEconomicoxRURepo20(nGrupoCod As Integer, dfecha As Date) As ADODB.Recordset
    sSql = "exec stp_sel_Reporte20GrupoEconomico  " & nGrupoCod & ",'" & Format(dfecha, "YYYY/MM/DD") & "'"
    Set ListarDatosGrupoEconomicoxRURepo20 = oCon.CargaRecordSet(sSql)
End Function
Public Function ObtenerDatosPropiedadDirectaEIndirecta() As ADODB.Recordset
    sSql = "exec stp_sel_ObtenerPropiedadDirectaEIndirecta  "
    Set ObtenerDatosPropiedadDirectaEIndirecta = oCon.CargaRecordSet(sSql)
End Function

'*******************************************************************
'JAME20140303
Public Function DevuelveGrupoEconomicoPersona(psPersCod As String) As ADODB.Recordset
   sSql = " exec stp_sel_DevuelveGrupoEconomicoPersona '" & psPersCod & "'"
   Set DevuelveGrupoEconomicoPersona = oCon.CargaRecordSet(sSql)
End Function
Public Function FiltroBuscarEmpresaGrupoEconomico(psPersNombre As String, pnGrupoId As Integer) As ADODB.Recordset
    sSql = " exec stp_sel_FiltroBuscarEmpresaGrupoEconomico " & pnGrupoId & ",'" & psPersNombre & "'"
    Set FiltroBuscarEmpresaGrupoEconomico = oCon.CargaRecordSet(sSql)
End Function
'PTI1 AGREGADO 17/08/2018 INFORME N° 002-2017 AC TI /CMACM
'Se dejo de usar FiltroBuscarEmpresaGrupoEconomico
Public Function FiltroBuscarSoloEmpresaXGrupoEconomico(psPersNombre As String, pnGrupoId As Integer) As ADODB.Recordset
    sSql = " exec stp_sel_FiltroBuscarSoloEmpresaGrupoEconomico " & pnGrupoId & ",'" & psPersNombre & "'"
    Set FiltroBuscarSoloEmpresaXGrupoEconomico = oCon.CargaRecordSet(sSql)
End Function
'FIN AGREGADO

'*******************************************************************

