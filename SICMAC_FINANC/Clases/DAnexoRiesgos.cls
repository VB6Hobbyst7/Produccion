VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DAnexoRiesgos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sConexion As String
Dim sCentralCom As String
Dim sServidorConsolidada As String

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    Set oImp = Nothing
    
    Dim oIni As New ClasIni
    Set oIni = Nothing
    Set dbConec = New DConecta
    dbConec.AbreConexion
    
    Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
    
    sServidorConsolidada = oConst.LeeConstSistema(gConstSistServCentralRiesgos)
    Set oConst = Nothing
End Sub

Private Sub Class_Terminate()
   dbConec.CierraConexion
   Set dbConec = Nothing
End Sub

Public Function CargaConceptos(psOpeCod As String, Optional psConcep As String = "", Optional psCodigo As String = "") As Recordset
On Error GoTo CargaConceptosErr
psSql = "SELECT RTrim(LTrim(cCodClase)) cCodClase, RTrim(LTrim(cCodGrp)) cCodGrp, RTrim(LTrim(cCodConcep)) cCodConcep, isnull(cDescrip,'')cDescrip, isnull(cTpoCalculo,'') cTpoCalculo, ISNULL(RTRIM(tc.cConsDescripcion),'') cTpoCalculoDesc, ISNULL(cCtaContCod,'') cCtaContCod, ISNULL(cFormula,'') cFormula  " _
      & "FROM AnxRiesgosConcepto a LEFT JOIN (SELECT nConsValor,cConsDescripcion FROM Constante WHERE nConsCod = '" & gAnxTipoCalculo & "') tc " _
      & "       ON RTRIM(tc.nConsValor) = a.cTpoCalculo " _
      & "WHERE cOpeCod = '" & psOpeCod & "'" & IIf(psConcep <> "", " and cCodConcep = '00'", "") _
      & IIf(psCodigo = "", "", " and cCodClase + cCodGrp + cCodConcep = '" & psCodigo & "'") & " Order By cCodClase, cCodGrp, cCodConcep "
Set CargaConceptos = dbConec.CargaRecordSet(psSql)
    Exit Function
CargaConceptosErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaConceptos Method")
End Function

Public Function CargaRangos(psOpeCod As String) As Recordset
On Error GoTo CargaRangosErr
psSql = "SELECT cCodRango, cDescrip, nDesde, nHasta " _
      & "FROM AnxRiesgosRango WHERE cOpeCod = '" & psOpeCod & "'"
Set CargaRangos = dbConec.CargaRecordSet(psSql)
    Exit Function
CargaRangosErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaRangos Method")
End Function

Public Function CargaCriterios(psOpeCod As String, Optional psCodigo As String = "") As Recordset
On Error GoTo CargaCriteriosErr
psSql = "SELECT * FROM AnxRiesgosCriterio WHERE cOpeCod = '" & psOpeCod & "' " & IIf(psCodigo = "", "", " and cCodClase + cCodGrp + cCodConcep = '" & psCodigo & "' ") & "  "
Set CargaCriterios = dbConec.CargaRecordSet(psSql)
    Exit Function
CargaCriteriosErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaCriterios Method")
End Function

Public Function CargaConceptoCriterio(psOpeCod As String) As Recordset
On Error GoTo CargaConceptoCriterioErr
psSql = "SELECT cCodClase, cCodGrp, cCodConcep, isnull(cDescrip,'')cDescrip, cTpoCalculo, ISNULL(RTRIM(tc.cConsDescripcion),'') cTpoCalculoDesc " _
      & "FROM AnxRiesgosConcepto a JOIN (SELECT nConsValor, cConsDescripcion FROM Constante WHERE nConsCod = '" & gAnxTipoCalculo & "') tc " _
      & "     ON tc.nConsValor = a.cTpoCalculo " _
      & "WHERE cOpeCod = '" & psOpeCod & "' Order By cCodClase, cCodGrp, cCodConcep"
Set CargaConceptoCriterio = dbConec.CargaRecordSet(psSql)
    Exit Function
CargaConceptoCriterioErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaConceptoCriterio Method")
End Function

Public Function CargaRangoCriterio(psOpeCod As String, psClase As String, psGrupo As String, psConcepto As String) As Recordset
On Error GoTo CargaRangoCriterioErr
psSql = "SELECT r.cCodRango, r.cDescrip, ISNULL(c.cValor,'') cValor, ISNULL(c.cValor,'') cValor1 FROM AnxRiesgosRango r LEFT JOIN " _
        & "(SELECT cCodRango, cValor FROM AnxRiesgosCriterio " _
        & " WHERE cOpeCod = '" & psOpeCod & "' and cCodClase = '" & psClase & "' and cCodGrp = '" & psGrupo & "' and cCodConcep = '" & psConcepto & "') c ON c.cCodRango = r.cCodRango " _
        & "WHERE r.cOpeCod = '" & psOpeCod & "' "
Set CargaRangoCriterio = dbConec.CargaRecordSet(psSql)
    Exit Function
CargaRangoCriterioErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaRangoCriterio Method")
End Function

Public Function ListaDepositosSFN(ByVal pdFecha As Date, ByVal nTipoCambioAn As Currency) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim psSql As String
    On Error GoTo ErroListaDepositosSFN
    oConec.AbreConexion
    
    psSql = "Exec sp_ERS0792016_sel_DepositosSFN '" & Format(pdFecha, "yyyymmdd") & "'," & nTipoCambioAn & " "
    oConec.AbreConexion
    Set ListaDepositosSFN = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    
    Exit Function
ErroListaDepositosSFN:
    Call RaiseError(MyUnhandledError, "Lista Depositos del Sistema Financiero Nacional")
    
End Function '***********NAGL ERS 079-2016 20170407

Public Function ListaDepositosESTADO(ByVal pdFecha As Date, ByVal nTipoCambioAn As Currency) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim psSql As String
    On Error GoTo ErroListaDepositosEstado
    
    psSql = "Exec sp_ERS0792016_sel_DepositosEstado '" & Format(pdFecha, "yyyymmdd") & "', " & nTipoCambioAn & ""
    oConec.AbreConexion
    Set ListaDepositosESTADO = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    
    Exit Function
ErroListaDepositosEstado:
    Call RaiseError(MyUnhandledError, "Lista Depositos del Estado")
End Function '***********NAGL ERS 079-2016 20170407

Public Function ListaOtrosDepositantes(ByVal pnNroCliente As String, ByVal pdFecDiaProceso As Date, ByVal psMoneda As String, pnTipoCambio As Currency, Optional psTipoResult As String = "") As ADODB.Recordset 'NAGL 2190525 Agregó Optional psTipoResult As String = ""
'Optional pnTipo As Integer = 0 Excluyo este parámetro NAGL 20190627
Dim oConec As New DConecta
Dim psSql As String
On Error GoTo ErroListaOtrosDepositantes
    psSql = "Exec sp_ERS0792016_sel_ListarOtrosDepositantes " & pnNroCliente & ",'" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'," & pnTipoCambio & ",'" & psTipoResult & "'" 'NAGL 2190525 Agregó psTipoResult 'NAGL20190627 Excluyó ''" & pnTipo & ",
    oConec.AbreConexion
    Set ListaOtrosDepositantes = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    Exit Function
ErroListaOtrosDepositantes:
    Call RaiseError(MyUnhandledError, "Lista Otros Depositantes")
End Function '***********NAGL ERS 079-2016 20170407

Public Function MuestraBCRP(ByVal pdFecha As Date, ByVal TipoInstrumento As String) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim psSql As String
    On Error GoTo ErroMuestraBCRP
    oConec.AbreConexion
    psSql = "Exec sp_ERS0792016_sel_ListaTipoInstrumento '" & Format(pdFecha, "yyyymmdd") & "', '" & TipoInstrumento & "'"
    oConec.AbreConexion
    Set MuestraBCRP = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    Exit Function
ErroMuestraBCRP:
    Call RaiseError(MyUnhandledError, "Muestra BCRP")
End Function '***********NAGL ERS 079-2016 20170407

Public Function ObligacionBN(ByVal pdFecha As Date) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim psSql As String
    On Error GoTo ErroObligacionBN
    oConec.AbreConexion
    psSql = "Exec sp_ERS0792016_sel_ObligacionBN '" & Format(pdFecha, "yyyymmdd") & "'"
    oConec.AbreConexion
    Set ObligacionBN = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    Exit Function
ErroObligacionBN:
    Call RaiseError(MyUnhandledError, "Obligacion BN")
End Function '***********NAGL ERS 079-2016 20170407

'***********JIPR20200824
Public Function AdeudadoReactiva(ByVal pdFecha As Date, ByVal pcMoneda As String) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim psSql As String
    On Error GoTo ErroAdeudadoReactiva
    oConec.AbreConexion
    psSql = "Exec stp_sel_AdeudadoReactiva '" & Format(pdFecha, "yyyymmdd") & "', '" & pcMoneda & "'"
    oConec.AbreConexion
    Set AdeudadoReactiva = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    Exit Function
ErroAdeudadoReactiva:
    Call RaiseError(MyUnhandledError, "Adeudado Reactiva")
  End Function
  
  Public Function AdeudadoReactiva30(ByVal pdFecha As Date, ByVal pcMoneda As String) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim psSql As String
    On Error GoTo ErroAdeudadoReactiva30
    oConec.AbreConexion
    psSql = "Exec stp_sel_AdeudadoReactiva30 '" & Format(pdFecha, "yyyymmdd") & "', '" & pcMoneda & "'"
    oConec.AbreConexion
    Set AdeudadoReactiva30 = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    Exit Function
ErroAdeudadoReactiva30:
    Call RaiseError(MyUnhandledError, "Adeudado Reactiva 30 días")
  End Function
'***********JIPR20200824


Public Function DevuelveValorizacionDiariaDet(ByVal pdFecha As Date, ByVal psTipo As String) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim psSql As String
    psSql = "Exec sp_ERS0792016_sel_ReporteValorizacionDiariaDet '" & Format(pdFecha, "yyyymmdd") & "', '" & psTipo & "'"
    oConec.AbreConexion
    Set DevuelveValorizacionDiariaDet = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function '***********NAGL ERS 079-2016 20170407

Public Sub ControlValorDiaria(ByVal dFechaValorizacion As Date, ByVal nObligacionMN As Double, nObligacionME As Double, ByVal lnMovNro As String, ByVal pVDatos As ADODB.Recordset)
    Dim psSql As String
    Dim oConec As New DConecta
    Dim Anexo As New DAnexoRiesgos
    Dim I As Integer
    oConec.AbreConexion
    oConec.BeginTrans
    'psSql = "Exec stp_ERS0642016_ins_ValorizacionDiaria '" & Format(dFechaValorizacion & " " & oConec.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "'," & nObligacionMN & "," & nObligacionME & ""
    psSql = "Exec sp_ERS0792016_ins_ValorizacionDiaria '" & Format(dFechaValorizacion, "mm/dd/yyyy") & "'," & nObligacionMN & "," & nObligacionME & ", '" & lnMovNro & "'"
    oConec.Ejecutar (psSql)
    
    psSql = "Exec sp_ERS0792016_del_ValorizacionDiariaDet '" & Format(dFechaValorizacion, "mm/dd/yyyy") & "'"
    oConec.Ejecutar (psSql)
    
    If Not pVDatos Is Nothing Then
     Do While Not pVDatos.EOF
     psSql = " Exec sp_ERS0792016_ins_ValorizacionDiariaDet '" & Format(dFechaValorizacion, "mm/dd/yyyy") & "', '" & Format(pVDatos!dFechaEmision, "mm/dd/yyyy") & "','" & pVDatos!cTipoInstrumento & "','" & pVDatos!cNemonico & "'," & pVDatos!nTasaInteres & "," & pVDatos!nNroPapeles & "," & pVDatos!nPrecioNominal & "," & pVDatos!nPrecioSucio & "," & pVDatos!nValorRazonable & ""
            oConec.Ejecutar (psSql)
            pVDatos.MoveNext
        Loop
    End If
    oConec.CommitTrans
    oConec.CierraConexion
    Set oConec = Nothing
End Sub '***********NAGL ERS 079-2016 20170407

Public Function CargaObligacionBN(ByVal pdFecha As Date, ByVal psTipo As String) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim psSql As String
    On Error GoTo ErroCargaObligacionBN
    
    psSql = "Exec sp_ERS0792016_sel_CargaObligacionBN '" & Format(pdFecha, "yyyymmdd") & "', '" & psTipo & "'"
    oConec.AbreConexion
    Set CargaObligacionBN = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    Exit Function
ErroCargaObligacionBN:
    Call RaiseError(MyUnhandledError, "Carga Obligacion BN")
End Function '***********NAGL ERS 079-2016 20170407

'TORE ERS041-2018
Public Function GrabaMovIndLiq(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String) As Long
Dim oDMov As DMov
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim sMsgErr As String
On Error GoTo GrabaMovIndLiqErr

Set oDMov = New DMov
oDMov.BeginTrans
lbTrans = True

oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.CommitTrans
lbTrans = False

GrabaMovIndLiq = lnMovNro

Set oDMov = Nothing
Exit Function
GrabaMovIndLiqErr:
    sMsgErr = Err.Description
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    Err.Raise vbObjectError + 100, "GrabaMovIndLiq", sMsgErr
End Function
Public Sub GrabarProcesoGenerandoAnexos(ByVal pnMov As Long, ByVal psOpeCod As String)
    Dim psSql As String
    Dim oConec As New DConecta
    oConec.AbreConexion
    psSql = "exec stp_ins_GenerandoAnexos " & pnMov & ",'" & psOpeCod & "'"
    oConec.Ejecutar (psSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub

Public Function CargarMotivosOperaciones(ByVal psAreaCod As String) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim sql As String
    On Error GoTo CargarMotivosOperacionesErr
    
    sql = "exec stp_sel_ObtenerMotivosOperaciones '" & psAreaCod & "'"
    oConec.AbreConexion
    Set CargarMotivosOperaciones = oConec.CargaRecordSet(sql)
    oConec.CierraConexion
    Exit Function
CargarMotivosOperacionesErr:
    Call RaiseError(MyUnhandledError, "Cargar Motivos Operaciones")
End Function

Public Function MostarFMIncidentes(ByVal pnIdMotivo As Integer) As Boolean
    On Error GoTo MostarFMIncidentesErr
    Dim oRs As ADODB.Recordset
    Dim oConec As DConecta
    Dim sql As String
    Dim mostrar As Boolean
    
    Set oRs = New ADODB.Recordset
    Set oConec = New DConecta
    oConec.AbreConexion
    
    sql = "exec stp_sel_MostarFMIncidentes " & pnIdMotivo
    Set oRs = oConec.CargaRecordSet(sql)
    If Not oRs.BOF And Not oRs.EOF Then
        If oRs!nTipoMot = 1 Then
            mostrar = True
        Else
            mostrar = False
        End If
    End If
    oConec.CierraConexion
    MostarFMIncidentes = mostrar
Exit Function
MostarFMIncidentesErr:
    Call RaiseError(MyUnhandledError, "Cargar Motivos Operaciones")
End Function

'*** Verificar que se que se esta verificar que se esta generando los anexos ***
Public Function VerificarEstadoAnexos15A15B(ByVal pdFechaMov As String) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim sql As String
    On Error GoTo VerificarEstadoAnexos15A15BErr
    
    sql = "exec stp_sel_VerificarEstadoAnexos15A15B '" & pdFechaMov & "'"
    oConec.AbreConexion
    Set VerificarEstadoAnexos15A15B = oConec.CargaRecordSet(sql)
    oConec.CierraConexion
    Exit Function
VerificarEstadoAnexos15A15BErr:
    Call RaiseError(MyUnhandledError, "Cargar Motivos Operaciones")
End Function

'Copia de Seguridad
'Public Function VerificarEstadoAnexos15A15B() As String
'    On Error GoTo VerificarEstadoAnexos15A15BErr
'    Dim oRs As ADODB.Recordset
'    Dim oConec As DConecta
'    Dim sql As String
'    Dim nValor As String
''    Dim HoraSis As Variant
''    Dim HoraAct As String
'    Set oRs = New ADODB.Recordset
'    Set oConec = New DConecta
'    oConec.AbreConexion
'
'    'Hora actual del sistema
''    HoraSis = Time
''    HoraAct = CStr(Hour(HoraSis)) & ":" & Minute(HoraSis)
'
''    sql = "exec stp_sel_VerificarEstadoAnexos15A15B '" & HoraAct & "'"
'    sql = "exec stp_sel_VerificarEstadoAnexos15A15B"
'    Set oRs = oConec.CargaRecordSet(sql)
'    If Not oRs.BOF And Not oRs.EOF Then
'        nValor = oRs!nValorExtorno
'    End If
'    oConec.CierraConexion
'    VerificarEstadoAnexos15A15B = nValor
'Exit Function
'VerificarEstadoAnexos15A15BErr:
'    Call RaiseError(MyUnhandledError, "Cargar Motivos Operaciones")
'End Function

Public Function VerificarFechaOperacion(ByVal dFechaOperacion As String, ByVal dFechaSistema As String) As Boolean
    On Error GoTo VerificarFechaOperacionErr
    Dim oRs As ADODB.Recordset
    Dim oConec As DConecta
    Dim sql As String
    Dim nValor As Integer
    
    Set oRs = New ADODB.Recordset
    Set oConec = New DConecta
    oConec.AbreConexion
    
    sql = "exec stp_sel_VerificarFechaOperaciones '" & dFechaOperacion & "','" & dFechaSistema & "'"
    Set oRs = oConec.CargaRecordSet(sql)
    If Not oRs.BOF And Not oRs.EOF Then
        If oRs!nPermiso = "SI" Then
            VerificarFechaOperacion = True
        Else
            VerificarFechaOperacion = False
        End If
    End If
    oConec.CierraConexion
Exit Function
VerificarFechaOperacionErr:
    Call RaiseError(MyUnhandledError, "VerificarFechaOperacion")
End Function

Public Function MostarControlIndLiq(ByVal pdFechaOpe As String, ByVal pdFechaSicmac As String) As ADODB.Recordset
 On Error GoTo hErr
    Dim oRs As ADODB.Recordset
    Dim oConec As DConecta
    Dim HoraSis As Variant
    Dim HoraAct As String
    Dim bValor As Boolean
    
    'Valor de verificacion
    Set oRs = New ADODB.Recordset
    Set oConec = New DConecta
    oConec.AbreConexion
    
    'Hora actual del sistema
    HoraSis = Time
    HoraAct = CStr(Hour(HoraSis)) & ":" & Minute(HoraSis)
    psSql = "exec stp_sel_VerificarFechaOperaciones '" & pdFechaOpe & "','" & HoraAct & "'" & pdFechaSicmac & "'"
    Set MostarControlIndLiq = oConec.CargaRecordSet(psSql)
  
    oConec.CierraConexion
    Exit Function
hErr:
Call RaiseError(MyUnhandledError, "DAnexoRiesgos:No Registrado Method")
End Function

Public Function VerificarExistenciaIncidente(ByVal psNroIncidente As String) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim psSql As String
    On Error GoTo ErroVerificarExistenciaIncidente
    
    psSql = "Exec stp_sel_VerificarExistenciaIncidente '" & psNroIncidente & "'"
    oConec.AbreConexion
    Set VerificarExistenciaIncidente = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    Exit Function
ErroVerificarExistenciaIncidente:
    Call RaiseError(MyUnhandledError, "Carga Patrimonio Efectivo")
End Function
Public Sub RegistrarMotivoOperacion(ByVal pnMotivoSel As Integer, ByVal psNroMov As String, _
                                    ByVal psNroIncidente As String, ByVal pdFechaOpe As Date, _
                                    ByVal pdFechaSis As Date, Optional ByVal pnMovNro As Long = 0)
    Dim psSql As String
    Dim oConec As New DConecta
    oConec.AbreConexion
    psSql = "exec stp_ins_RegistrarMotivoOperacion " & pnMotivoSel & ",'" & psNroMov & _
            "','" & psNroIncidente & "','" & pdFechaOpe & "','" & pdFechaSis & "'," & pnMovNro
    oConec.Ejecutar (psSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub
'Realiza el extorno del movimiento que se genero al momento de inciar la generacion de los anexos
Public Sub ExtornarMovAnexos(ByVal psOpeCod As String)
    Dim psSql As String
    Dim oConec As New DConecta
    oConec.AbreConexion
    psSql = "exec stp_upd_RealizarExtornoAnexos '" & psOpeCod & "'"
    oConec.Ejecutar (psSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub

'Obtiene los datos del usuario que realizo la operación
Public Function ObtenerDatosUsuarioOperacion(ByVal psUser As String) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim psSql As String
    On Error GoTo ErroObtenerDatosUsuarioOperacion
    
    psSql = "exec stp_sel_ObtenerDatosxUsuario '" & psUser & "'"
    oConec.AbreConexion
    Set ObtenerDatosUsuarioOperacion = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    Exit Function
ErroObtenerDatosUsuarioOperacion:
    Call RaiseError(MyUnhandledError, "Obtener datos del usuario que realizo la operacion")
End Function
'Obtiene los detalles de la operacion
Public Function ObtenerDetalleOperacion(ByVal psMovNro As String, Optional ByVal pnMovNro As Long = 0) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim psSql As String
    On Error GoTo ErroObtenerDetalleOperacion
    
    psSql = "exec stp_sel_ObtenerDetalleMov '" & psMovNro & "'," & pnMovNro
    oConec.AbreConexion
    Set ObtenerDetalleOperacion = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    Exit Function
ErroObtenerDetalleOperacion:
    Call RaiseError(MyUnhandledError, "Obtener detalle de la operacion")
End Function
'END TORE

Public Function ObtenerRegValorizacionDiaria(pdFecDiaProceso As Date) As String
On Error GoTo ErroObtenerRegValorizacionDiaria
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nReg As String
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_RegValorizacionDiaria '" & Format(pdFecDiaProceso, "yyyymmdd") & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
        nReg = oRs!Registrado
   Else
    nReg = "0"
   End If
   oConec.CierraConexion
   ObtenerRegValorizacionDiaria = nReg
Exit Function
ErroObtenerRegValorizacionDiaria:
   Call RaiseError(MyUnhandledError, "DAnexoRiesgos:No Registrado Method")
End Function 'NAGL 20170426

Public Function DevuelveFluctuacionesValoresDet(ByVal pdFecha As Date, ByVal psTipo As String) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim psSql As String
   
    psSql = "Exec sp_ERS0792016_sel_ReporteFluctuacionValoresDet '" & Format(pdFecha, "yyyymmdd") & "', '" & psTipo & "'"
    oConec.AbreConexion
   
    Set DevuelveFluctuacionesValoresDet = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function '***********NAGL 20170426

Public Sub ControlFluctuacionValores(ByVal dFechaFluctuacion As Date, ByVal dFechaPatrimonio As Date, ByVal nPatrimonioMN As Double, nPatrimonioME As Double, ByVal lnMovNro As String, ByVal pVDatos As ADODB.Recordset)
    Dim psSql As String
    Dim oConec As New DConecta
    Dim Anexo As New DAnexoRiesgos
    Dim I As Integer
    oConec.AbreConexion
    oConec.BeginTrans
    
    psSql = "Exec sp_ERS0792016_ins_FluctuacionValores '" & Format(dFechaFluctuacion, "mm/dd/yyyy") & "','" & Format(dFechaPatrimonio, "mm/dd/yyyy") & "', " & nPatrimonioMN & "," & nPatrimonioME & ", '" & lnMovNro & "'"
    oConec.Ejecutar (psSql)
    
    psSql = "Exec sp_ERS0792016_del_FluctuacionValoresDet '" & Format(dFechaFluctuacion, "mm/dd/yyyy") & "'"
    oConec.Ejecutar (psSql)
    
    If Not pVDatos Is Nothing Then
     Do While Not pVDatos.EOF
            psSql = " Exec sp_ERS0792016_ins_FluctuacionValoresDet '" & Format(dFechaFluctuacion, "mm/dd/yyyy") & "', '" & pVDatos!cCtaCod & "', '" & pVDatos!cPerscod & "', '" & Format(pVDatos!dfechaVencimiento, "mm/dd/yyyy") & "', '" & pVDatos!cTpoMoneda & "'," & pVDatos!nSaldo & "," & pVDatos!nFluctuacion & "," & pVDatos!nValorLibroTotal & "" 'Se cambio de pVDatos!cCtaNombre a pVDatos!cPersCod NAGL 20171222
            oConec.Ejecutar (psSql)
            pVDatos.MoveNext
        Loop
    End If
    oConec.CommitTrans
    oConec.CierraConexion
    Set oConec = Nothing
End Sub '************NAGL 20170426 Según RFC1712260004*****

Public Function CargaAsigTipoInstrumento(ByVal psCtaCod As String) As String
    Dim oConec As New DConecta
    Dim psSql As String
    Dim psCtaNomb As String
    Dim rs As New ADODB.Recordset
    On Error GoTo ErroCargaAsigTipoInstrumento
    
    psSql = "Exec stp_sel_AsignaTipoInstrumento '" & psCtaCod & "'"
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(psSql)
    
    If Not rs.BOF And Not rs.EOF Then
        psCtaNomb = rs!TpoInstrum
    Else
        psCtaNomb = ""
    End If
    oConec.CierraConexion
    CargaAsigTipoInstrumento = psCtaNomb
    Exit Function
ErroCargaAsigTipoInstrumento:
    Call RaiseError(MyUnhandledError, "Carga Tipo de Instrumento")
End Function '***********NAGL 20170426 Según RFC1712260004*****

Public Function CargaPatrimonioEfectivo(ByVal pdFecha As Date) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim psSql As String
    On Error GoTo ErroCargaPatrimonioEfectivo
    
    psSql = "Exec sp_ERS0792016_sel_CargaPatrimonioEfectivo '" & Format(pdFecha, "yyyymmdd") & "'"
    oConec.AbreConexion
    Set CargaPatrimonioEfectivo = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    Exit Function
ErroCargaPatrimonioEfectivo:
    Call RaiseError(MyUnhandledError, "Carga Patrimonio Efectivo")
End Function '***********NAGL 20170426

Public Function ObtenerRegFluctuacionValores(pdFecDiaFluctuacion As Date) As String
On Error GoTo ErroObtenerRegFluctuacionValores
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nReg As String
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_RegFluctuacionValores '" & Format(pdFecDiaFluctuacion, "yyyymmdd") & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
        nReg = oRs!Registrado
   Else
    nReg = "0"
   End If
   oConec.CierraConexion
   ObtenerRegFluctuacionValores = nReg
Exit Function
ErroObtenerRegFluctuacionValores:
   Call RaiseError(MyUnhandledError, "DAnexoRiesgos:No Registrado Method")
End Function 'NAGL 20170426

Public Function DevueleReporteVarCambYTipoCambioSBS(ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, ByVal psTipoVar As String, ByVal pnOptBal As Integer) As ADODB.Recordset
Dim psSql As String
Dim oConec As New DConecta
Dim oRs As New ADODB.Recordset
  psSql = "Exec stp_sel_ReporteVarCambYBackTest '" & Format(pdFechaIni, "yyyymmdd") & "','" & Format(pdFechaFin, "yyyymmdd") & "' ,'" & psTipoVar & "', " & pnOptBal & " "
  oConec.AbreConexion
  Set DevueleReporteVarCambYTipoCambioSBS = oConec.CargaRecordSet(psSql)
  oConec.CierraConexion
  Set oConec = Nothing
End Function 'NAGL ERS002-2017 20170815

Public Function ObtieneValRegTipoCambioSBS(ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, ByVal pnTipo As Integer) As String
Dim psSql As String
Dim oConec As New DConecta
Dim oRs As New ADODB.Recordset
Dim ValReg As String
oConec.AbreConexion
   psSql = "Exec stp_sel_ValRegTipoCambioSBS '" & Format(pdFechaIni, "yyyymmdd") & "','" & Format(pdFechaFin, "yyyymmdd") & "', " & pnTipo & ""
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
        ValReg = oRs!Registrado
   Else
        ValReg = "NO"
   End If
   oConec.CierraConexion
   ObtieneValRegTipoCambioSBS = ValReg
   Set oConec = Nothing
End Function 'NAGL ERS002-2017 20170815

Public Function DevuelveTipoCambioSBS(ByVal pdFecCambSBS As Date, ByVal pdFecCambSBSFin As Date, ByVal pnTipo As Integer)
Dim psSql As String
Dim oConec As New DConecta
    psSql = "Exec stp_sel_ObtieneTipoCambioSBS '" & Format(pdFecCambSBS, "yyyymmdd") & "','" & Format(pdFecCambSBSFin, "yyyymmdd") & "', " & pnTipo & " "
    oConec.AbreConexion
    Set DevuelveTipoCambioSBS = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function 'NAGL ERS002-2017 20170815

Public Sub ControlTipoCambioSBS(ByVal lnMovNro As String, ByVal pVDatosTipoSBS As ADODB.Recordset)
    Dim psSql As String
    Dim oConec As New DConecta
    Dim Anexo As New DAnexoRiesgos
    Dim Orden As SortSettings
    Dim I As Integer
    oConec.AbreConexion
    oConec.BeginTrans
    
    If Not pVDatosTipoSBS Is Nothing Then
     pVDatosTipoSBS.Sort = "pdFecCambSBS ASC" 'Para Ordenar el RecordSet
     Do While Not pVDatosTipoSBS.EOF
        psSql = "Exec stp_ins_TipoCambioSBS '" & Format(pVDatosTipoSBS!pdFecCambSBS, "YYYYMMDD") & "', " & pVDatosTipoSBS!TCCompra & "," & pVDatosTipoSBS!TCCVenta & ",'" & lnMovNro & "'"
                oConec.Ejecutar (psSql)
                pVDatosTipoSBS.MoveNext
        Loop
    End If
    oConec.CommitTrans
    oConec.CierraConexion
    Set oConec = Nothing
End Sub '***********NAGL ERS002-2017 20170810

Public Function PermiteDeleteModificarTpoCambSBS(pdFecCambSBS As Date) As Boolean
Dim ldFecCie As Date
   On Error GoTo PermiteDeleteModificarTpoCambSBSErr
   PermiteDeleteModificarTpoCambSBS = False
   Dim oGen As New NConstSistemas
   ldFecCie = CDate(oGen.LeeConstSistema(gConstSistCierreMensualCont))
   Set oGen = Nothing
   If pdFecCambSBS <= ldFecCie Then
      PermiteDeleteModificarTpoCambSBS = True
         'MsgBox "El Tipo de Cambio pertenece a un mes ya Cerrado. Imposible realizar operación...!", vbInformation, "Aviso"
   End If
   Exit Function
PermiteDeleteModificarTpoCambSBSErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:DAnexoRiesgos:PermiteElimModTipoCambSBS Method")
End Function '*******NAGL ERS002-2017 20170807

Public Sub EliminaTipoCambioSBS(pdFecCambSBS As Date)
    Dim psSql As String
    Dim oConec As New DConecta
    oConec.AbreConexion
    psSql = "Exec stp_del_TipoCambioSBS '" & Format(pdFecCambSBS, "yyyymmdd") & "'"
    oConec.Ejecutar (psSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub '********NAGL ERS002-2017 20170807

Public Function DevueleReporteCalculoRCambiario(ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, ByVal psTipo As String) As ADODB.Recordset
Dim psSql As String
Dim oConec As New DConecta
Dim rs As New ADODB.Recordset
If psTipo = "POSCAM" Then
    psSql = "Exec stp_sel_ReporteCalculoRCambiario '" & Format(pdFechaIni, "yyyymmdd") & "','" & Format(pdFechaFin, "yyyymmdd") & "'"
ElseIf psTipo = "NEGCAMB" Then
    psSql = "Exec stp_sel_ReportNegocioCambiario '" & Format(pdFechaIni, "yyyymmdd") & "','" & Format(pdFechaFin, "yyyymmdd") & "'"
ElseIf psTipo = "ResulDifCamb" Or psTipo = "ComVentSpot" Then
    psSql = "Exec stp_sel_ResultadosAcumDifCambio_OperComp_Venta '" & Format(pdFechaIni, "yyyymmdd") & "','" & Format(pdFechaFin, "yyyymmdd") & "', '" & psTipo & "'"
End If
oConec.AbreConexion
Set DevueleReporteCalculoRCambiario = oConec.CargaRecordSet(psSql)
oConec.CierraConexion
Set oConec = Nothing
End Function '*****NAGL ERS002-2017 20170815

Public Function DevueleDetalleResultBalance(ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, Optional cCtaContCod As String = "", Optional DarLimite As String = "") As ADODB.Recordset
Dim psSql As String
Dim oConec As New DConecta
Dim rs As New ADODB.Recordset
    psSql = "Exec stp_sel_ResultadosDetalleBalanceComprob '" & Format(pdFechaIni, "yyyymmdd") & "','" & Format(pdFechaFin, "yyyymmdd") & "', '" & cCtaContCod & "', '" & DarLimite & "'"
    oConec.AbreConexion
    Set DevueleDetalleResultBalance = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
Set oConec = Nothing
End Function 'NAGL ERS002-2017 20170829

Public Function ObtieneCuentaCntCalculoRCamb() As ADODB.Recordset
Dim psSql As String
Dim oConec As New DConecta
Dim rs As New ADODB.Recordset
    psSql = "Exec stp_sel_GetCuentaCntCalculoRCamb"
    oConec.AbreConexion
    Set ObtieneCuentaCntCalculoRCamb = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function '******NAGL ERS002-2017 20170801

Public Function ObtieneReporteControlTOSEFinMes(ByVal pdFecha As Date, ByVal psTipo As String) As ADODB.Recordset
Dim psSql As String
Dim oConec As New DConecta
Dim rs As New ADODB.Recordset
If psTipo = "DatoFinMes" Then
    psSql = "Exec stp_sel_ReporteDatosTOSEFindeMes '" & Format(pdFecha, "yyyymmdd") & "'"
ElseIf psTipo = "PromMesIII" Then
    psSql = "Exec stp_sel_ReporteControlTOSEPromedioyIIIMes '" & Format(pdFecha, "yyyymmdd") & "'"
End If
oConec.AbreConexion
Set ObtieneReporteControlTOSEFinMes = oConec.CargaRecordSet(psSql)
oConec.CierraConexion
Set oConec = Nothing
End Function '******NAGL ERS002-2017 20170912

Public Function ObtieneReporteDatosHistTOSE(ByVal pdFecha As Date) As ADODB.Recordset
Dim psSql As String
Dim oConec As New DConecta
Dim rs As New ADODB.Recordset
psSql = "Exec stp_sel_ReporteDatosHistoricosTOSE '" & Format(pdFecha, "yyyymmdd") & "'"
oConec.AbreConexion
Set ObtieneReporteDatosHistTOSE = oConec.CargaRecordSet(psSql)
oConec.CierraConexion
Set oConec = Nothing
End Function '****NAGL ERS002-2017 20170912

Public Function DevuelveReporteControlTOSEDia(ByVal pdFecha As Date) As ADODB.Recordset
Dim psSql As String
Dim oConec As New DConecta
Dim rs As New ADODB.Recordset
    psSql = "Exec stp_sel_ReporteControlTOSEDia '" & Format(pdFecha, "yyyymmdd") & "'"
oConec.AbreConexion
Set DevuelveReporteControlTOSEDia = oConec.CargaRecordSet(psSql)
oConec.CierraConexion
Set oConec = Nothing
End Function '****NAGL ERS002-2017 20170912

Public Function ObtieneControlDatosRCLHistorico(ByVal pdFecha As Date) As ADODB.Recordset
Dim psSql As String
Dim oConec As New DConecta
    psSql = "Exec stp_sel_Reporte15BRCLHistorico '" & Format(pdFecha, "yyyymmdd") & "'"
oConec.AbreConexion
Set ObtieneControlDatosRCLHistorico = oConec.CargaRecordSet(psSql)
oConec.CierraConexion
Set oConec = Nothing
End Function '***NAGL ERS079-2017 20180118

Public Function DevuelveAnexo15BRatioLiquidezDiario(ByVal pdFecha As Date, Optional psTipo As String = "", Optional psSeccion As String = "") As ADODB.Recordset
Dim psSql As String
Dim oConec As New DConecta
If psTipo = "15B" Then
    psSql = "Exec stp_sel_ObtieneAnexo15BRatioLiquidezDiario '" & Format(pdFecha, "yyyymmdd") & "'"
Else
    psSql = "Exec stp_sel_GetFechasRatiosforAnex15BRCLMens '" & Format(pdFecha, "yyyymmdd") & "', '" & psSeccion & "' "
End If
oConec.AbreConexion
Set DevuelveAnexo15BRatioLiquidezDiario = oConec.CargaRecordSet(psSql)
oConec.CierraConexion
Set oConec = Nothing
End Function '****NAGL ERS079-2017 20180117

Public Function ObtenerSaldEstadistAnx15Ay15B(psCtaContCod As String, psMoneda As String, pdFechaProceso As Date, pdFechaBalanceDiario As Date, pnTipoCambBal As Currency, pnTipoCambioMes As Currency) As Currency
On Error GoTo ObtenerSaldEstadistAnx15Ay15BErr
   Dim oRs As New ADODB.Recordset
   Dim oConec As New DConecta
   Dim nSaldo As Currency
   Dim psSql As String
   oConec.AbreConexion
   psSql = "exec stp_sel_ObtenerSaldosEstadisticosAnexo15Ay15B '" & psCtaContCod & "','" & psMoneda & "','" & Format(pdFechaProceso, "yyyymmdd") & "','" & Format(pdFechaBalanceDiario, "yyyymmdd") & "', " & pnTipoCambBal & ", " & pnTipoCambioMes & ""
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!nSaldoImporte), 0, oRs!nSaldoImporte)
        oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSaldEstadistAnx15Ay15B = nSaldo
Exit Function
ObtenerSaldEstadistAnx15Ay15BErr:
   Call RaiseError(MyUnhandledError, "DAnexoRiesgos:ObtieneSaldosEstadisticos Method")
End Function '**NAGL ERS079-2017 20180122

Public Function CargaSaldosEstadistAnx15Ay15B(pdFechaProceso As Date, pdFechaBalanceDiario As Date, pnTipoCambBal As Currency, pnTipoCambioMes As Currency) As ADODB.Recordset
On Error GoTo CargaDataEstadistAnx15Ay15BErr
   Dim oConec As New DConecta
   Dim nSaldo As Currency
   Dim psSql As String
   oConec.AbreConexion
   psSql = "exec stp_sel_CargarSaldosEstadisticosAnexo15Ay15B '" & Format(pdFechaProceso, "yyyymmdd") & "','" & Format(pdFechaBalanceDiario, "yyyymmdd") & "', " & pnTipoCambBal & ", " & pnTipoCambioMes & ""
   Set CargaSaldosEstadistAnx15Ay15B = oConec.CargaRecordSet(psSql)
   oConec.CierraConexion
   Set oConec = Nothing
Exit Function
CargaDataEstadistAnx15Ay15BErr:
   Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaDataEstadistAnx15Ay15B Method")
End Function '**NAGL ERS006-2019 20190416

Public Function ObtieneRestringidosSFN(pdFecha As Date, psMoneda As String) As Currency
Dim psSql As String
Dim oConec As New DConecta
Dim rs As New ADODB.Recordset
Dim nSaldoRest As Currency
    oConec.AbreConexion
    psSql = "Exec stp_sel_ObtieneRestringidosSFN '" & Format(pdFecha, "yyyymmdd") & "','" & psMoneda & "'  "
    Set rs = oConec.CargaRecordSet(psSql)
    If Not rs.BOF And Not rs.EOF Then
      Do While Not rs.EOF
        nSaldoRest = IIf(IsNull(rs!nSaldo), 0, rs!nSaldo)
        rs.MoveNext
      Loop
   Else
    nSaldoRest = 0
   End If
   oConec.CierraConexion
   ObtieneRestringidosSFN = nSaldoRest
Set oConec = Nothing
End Function '***NAGL ERS079-2017 20180123

Public Function ObtieneReporteConsolEstadisticoBCRPSFNADEUD(ByVal pdFecha As Date, ByVal psCtaContCod As String) As ADODB.Recordset
Dim psSql As String
Dim oConec As New DConecta
    psSql = "Exec stp_sel_ObtieneConsoloEstadistBCRPSFNADEUD '" & Format(pdFecha, "yyyymmdd") & "','" & psCtaContCod & "'"
oConec.AbreConexion
Set ObtieneReporteConsolEstadisticoBCRPSFNADEUD = oConec.CargaRecordSet(psSql)
oConec.CierraConexion
Set oConec = Nothing
End Function '****NAGL ERS079-2017 20180130

Public Function ObtieneSaldoEfectTransitoTotal(ByVal pdFecha As Date, ByVal psMoneda As String) As Currency
Dim psSql As String
Dim rs As New ADODB.Recordset
Dim oConec As New DConecta
Dim nSaldoEfectTrans As Currency
oConec.AbreConexion
    psSql = "Exec stp_sel_GetCtaEfectTransitoConcFondos '" & Format(pdFecha, "yyyymmdd") & "','" & psMoneda & "'"
    Set rs = oConec.CargaRecordSet(psSql)
    If Not rs.BOF And Not rs.EOF Then
        nSaldoEfectTrans = IIf(IsNull(rs!nSaldoEfectTrans), 0, rs!nSaldoEfectTrans)
    Else
        nSaldoEfectTrans = 0
    End If
oConec.CierraConexion
ObtieneSaldoEfectTransitoTotal = nSaldoEfectTrans
Set oConec = Nothing
End Function '****NAGL Según TIC1807210002 20181001

Public Function ObtieneCtaSaldoDiario(psCtaContCod As String, pdFecha As Date, pdFechaBalanceDiario, pnTipoCambBalance As Currency, pnTipoCambioMes As Currency) As Currency
Dim psSql As String
Dim rs As New ADODB.Recordset
Dim oConec As New DConecta
Dim nSaldo As Currency
oConec.AbreConexion
    psSql = "Exec stp_sel_ObtieneCtaSaldoDiario '" & psCtaContCod & "','" & Format(pdFecha, "yyyymmdd") & "','" & Format(pdFechaBalanceDiario, "yyyymmdd") & "'," & pnTipoCambBalance & "," & pnTipoCambioMes & ""
    Set rs = oConec.CargaRecordSet(psSql)
    If Not rs.BOF And Not rs.EOF Then
        nSaldo = IIf(IsNull(rs!nSaldo), 0, rs!nSaldo)
    Else
        nSaldo = 0
    End If
oConec.CierraConexion
ObtieneCtaSaldoDiario = nSaldo
Set oConec = Nothing
End Function '****NAGL Según ERS006-2019 20190430

Public Function ObtieneCuentasContablesToAnexosLiqu(ByVal psTpoCtaCont As String, ByVal psMoneda As String, Optional psTipo As String = "", Optional psTipoDet As String) As ADODB.Recordset
Dim psSql As String
Dim oConec As New DConecta
    psSql = "Exec stp_sel_GetCtasContables '" & psTpoCtaCont & "','" & psMoneda & "','" & psTipo & "','" & psTipoDet & "'"
oConec.AbreConexion
Set ObtieneCuentasContablesToAnexosLiqu = oConec.CargaRecordSet(psSql)
oConec.CierraConexion
Set oConec = Nothing
End Function '****NAGL Según ERS006-2019 20190430

Public Sub GuardarTipoCambioSBS(pdFecCambSBS As Date, pnTpoCambioSBS As Currency, lsMovNro As String)
    Dim psSql As String
    Dim oConec As New DConecta
    oConec.AbreConexion
    psSql = "Exec stp_ins_TipoCambioSBS '" & Format(pdFecCambSBS, "YYYYMMDD") & "', " & pnTpoCambioSBS & "," & pnTpoCambioSBS & ",'" & lsMovNro & "'"
    oConec.Ejecutar (psSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub '********NAGL ERS006-2019 20190502

Public Function ObtieneTipoCambioContableSBS(ByVal pdFecha As Date) As Currency
Dim psSql As String
Dim rs As New ADODB.Recordset
Dim oConec As New DConecta
Dim nTipoCambSBS As Currency
    oConec.AbreConexion
    psSql = "Exec stp_ObtieneTipoCambioContable'" & Format(pdFecha, "yyyymmdd") & "'"
    Set rs = oConec.CargaRecordSet(psSql)
    If Not rs.BOF And Not rs.EOF Then
        nTipoCambSBS = IIf(IsNull(rs!nTipoCambCont), 0, rs!nTipoCambCont)
    Else
        nTipoCambSBS = 0
    End If
oConec.CierraConexion
ObtieneTipoCambioContableSBS = nTipoCambSBS
Set oConec = Nothing
End Function '********NAGL ERS006-2019 20190502

Public Function ValidaRatiosparaInterCambLiqMNyME(ByVal pnRCLMN As Currency, ByVal pnRCLME As Currency) As String
Dim psSql As String
Dim rs As New ADODB.Recordset
Dim oConec As New DConecta
Dim psRegValida As Currency
    oConec.AbreConexion
    psSql = "Exec stp_sel_ValidaRatiosparaIntercambioLiquMNyME " & pnRCLMN & "," & pnRCLME & ""
    Set rs = oConec.CargaRecordSet(psSql)
    If Not rs.BOF And Not rs.EOF Then
        psRegValida = rs!ValidacionInterCambLiqu
    Else
        psRegValida = ""
    End If
oConec.CierraConexion
ValidaRatiosparaInterCambLiqMNyME = psRegValida
Set oConec = Nothing
End Function '********NAGL ERS006-2019 20190506

Public Function ObtieneMontoInterCambioLiquidezMNyME(ByVal pnRCLMN As Currency, ByVal pnRCLME As Currency, ByVal pnTpoCambSBS As Currency, ByVal pnImporteActMN As Currency, ByVal pnImporteActME As Currency, ByVal pnFlujoEntMN As Currency, ByVal pnFlujoEntME As Currency, ByVal pnFlujoSalMN As Currency, ByVal pnFlujoSalME As Currency, ByVal psRegValida As String) As ADODB.Recordset
Dim psSql As String
Dim oConec As New DConecta
    psSql = "Exec stp_sel_ObtieneInterCambioLiquidezMNyME " & pnRCLMN & ", " & pnRCLME & ", " & pnTpoCambSBS & ", " & pnImporteActMN & ", " & pnImporteActME & ", " & pnFlujoEntMN & ", " & pnFlujoEntME & ", " & pnFlujoSalMN & ", " & pnFlujoSalME & ", " & psRegValida & ""
oConec.AbreConexion
Set ObtieneMontoInterCambioLiquidezMNyME = oConec.CargaRecordSet(psSql)
oConec.CierraConexion
Set oConec = Nothing
End Function '****NAGL Según ERS006-2019 20190506

Public Function ObtieneFondeoAnx16ANew(ByVal pdFecha As Date, ByVal pnTipoCambio As Currency, Optional psTipo As String = "") As ADODB.Recordset
Dim psSql As String
Dim oConec As New DConecta
    psSql = "Exec stp_sel_FondeoEncajeAnx16A '" & Format(pdFecha, "yyyymmdd") & "'," & pnTipoCambio & ", '" & psTipo & "'"
oConec.AbreConexion
Set ObtieneFondeoAnx16ANew = oConec.CargaRecordSet(psSql)
oConec.CierraConexion
Set oConec = Nothing
End Function '****NAGL Según ERS006-2019 20190514

Public Function ObtieneOtrasObligaciones15A(pdFecha As Date, psMoneda As String, Optional psTipoOblig As String = "", Optional psTipo As String = "") As ADODB.Recordset
Dim psSql As String
Dim oConec As New DConecta
    psSql = "Exec stp_sel_ObtieneOtrasObligaciones15A '" & Format(pdFecha, "yyyymmdd") & "','" & psMoneda & "','" & psTipoOblig & "','" & psTipo & "'"
oConec.AbreConexion
Set ObtieneOtrasObligaciones15A = oConec.CargaRecordSet(psSql)
oConec.CierraConexion
Set oConec = Nothing
End Function '****NAGL Según Anexo02 ERS006-2019 20190615

Public Function ObtieneDisponibilidadConsolCalif(ByVal pdFecha As Date) As Boolean
Dim psSql As String
Dim rs As New ADODB.Recordset
Dim oConec As New DConecta
Dim psValida As String
    oConec.AbreConexion
    psSql = "Exec stp_sel_GetDisponibilidadConsolCalif'" & Format(pdFecha, "yyyymmdd") & "'"
    Set rs = oConec.CargaRecordSet(psSql)
    psValida = rs!psValidaConsol
    If psValida = "1" Then
        ObtieneDisponibilidadConsolCalif = True
    Else
        ObtieneDisponibilidadConsolCalif = False
    End If
oConec.CierraConexion
Set oConec = Nothing
End Function '********NAGL Según Anx03_ERS006-2019 20191107

Public Sub EjecutaConsolidaIntDevCredFinMes(pdFecha As Date, psConsolida As String)
    Dim psSql As String
    Dim oConec As New DConecta
    oConec.AbreConexion
    psSql = "Exec stp_ins_ColocacSaldosDiariosTramosIntDev1418 '" & Format(pdFecha, "yyyymmdd") & "','" & psConsolida & "'"
    oConec.Ejecutar (psSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub '********NAGL Según Anx03_ERS006-2019 20191107

Public Sub EjecutaConsolidaSaldoCapCredFinMes(pdFecha As Date)
    Dim psSql As String
    Dim oConec As New DConecta
    oConec.AbreConexion
    psSql = "Exec stp_ins_ColocacSaldosDiariosTramosSaldoCap1401 '" & Format(pdFecha, "yyyymmdd") & "'"
    oConec.Ejecutar (psSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub '********NAGL 20191213 Según RFC1912050003

Public Sub EliminaConcepto(psOpeCod As String, psClase As String, psGrupo As String, psConcepto As String)
On Error GoTo EliminaConceptoErr
psSql = "DELETE AnxRiesgosConcepto WHERE cOpeCod = '" & psOpeCod & "' and cCodClase = '" & psClase & "' and cCodGrp = '" & psGrupo & "' and cCodConcep = '" & psConcepto & "'"
dbConec.Ejecutar psSql
Exit Sub
EliminaConceptoErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:EliminaConcepto Method")
End Sub

Public Sub EliminaRango(psOpeCod As String, psRango As String)
On Error GoTo EliminaRangoErr
psSql = "DELETE AnxRiesgosRango WHERE cOpeCod = '" & psOpeCod & "' and cCodRango = '" & psRango & "' "
dbConec.Ejecutar psSql
Exit Sub
EliminaRangoErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:EliminaRango Method")
End Sub

Public Sub InsertaConcepto(psOpeCod As String, psClase As String, psGrupo As String, psConcep As String, psDescrip As String, psTpoCalculo As String, psCtaCod As String, psFormula As String)
On Error GoTo InsertaConceptoErr
  psSql = "INSERT AnxRiesgosConcepto (cOpeCod, cCodClase, cCodGrp, cCodConcep, cDescrip, cTpoCalculo, cCtaContCod, cFormula) " _
        & "VALUES ('" & psOpeCod & "','" & psClase & "','" & psGrupo & "','" & psConcep & "','" & psDescrip & "','" & psTpoCalculo & "','" & psCtaCod & "','" & psFormula & "')"
dbConec.Ejecutar psSql
Exit Sub
InsertaConceptoErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:InsertaConcepto Method")
End Sub

Public Sub ActualizaConcepto(psOpeCod As String, psClase As String, psGrupo As String, psConcep As String, psDescrip As String, psTpoCalculo As String, psCtaCod As String, psFormula As String)
On Error GoTo ActualizaConceptoErr
  psSql = "UPDATE AnxRiesgosConcepto SET cDescrip = '" & psDescrip & "', cTpoCalculo = '" & psTpoCalculo & "', cCtaContCod = '" & psCtaCod & "', cFormula = '" & psFormula & "' WHERE cOpeCod = '" & psOpeCod & "' and cCodClase = '" & psClase & "' and cCodGrp = '" & psGrupo & "' and cCodConcep = '" & psConcep & "' "
   dbConec.Ejecutar psSql
Exit Sub
ActualizaConceptoErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:ActualizaConcepto Method")
End Sub

Public Sub InsertaCriterio(psOpeCod As String, psClase As String, psGrupo As String, psConcep As String, psRango As String, psValor As String)
On Error GoTo InsertaCriterioErr
   psSql = "INSERT AnxRiesgosCriterio (cOpeCod, cCodClase, cCodGrp, cCodConcep, cCodRango, cValor) " _
         & "VALUES ('" & psOpeCod & "','" & psClase & "','" & psGrupo & "','" & psConcep & "','" & psRango & "','" & psValor & "' )"
   dbConec.Ejecutar psSql
Exit Sub
InsertaCriterioErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:InsertaCriterio Method")
End Sub
      
Public Sub ActualizaCriterio(psOpeCod As String, psClase As String, psGrupo As String, psConcep As String, psRango As String, psValor As String)
On Error GoTo ActualizaCriterioErr
   psSql = "UPDATE AnxRiesgosCriterio SET cValor = '" & psValor & "' WHERE cOpeCod = '" & psOpeCod & "' and cCodClase = '" & psClase & "' and cCodGrp = '" & psGrupo & "' and cCodConcep = '" & psConcep & "' and cCodRango = '" & psRango & "'"
   dbConec.Ejecutar psSql
Exit Sub
ActualizaCriterioErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:ActualizaCriterio Method")
End Sub

Public Sub EliminaCriterio(psOpeCod As String, psClase As String, psGrupo As String, psConcep As String, psRango As String)
On Error GoTo EliminaCriterioErr
   psSql = "DELETE AnxRiesgosCriterio WHERE cOpeCod = '" & psOpeCod & "' and cCodClase = '" & psClase & "' and cCodGrp = '" & psGrupo & "' and cCodConcep = '" & psConcep & "' and cCodRango = '" & psRango & "'"
   dbConec.Ejecutar psSql
Exit Sub
EliminaCriterioErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:EliminaCriterio Method")
End Sub

Public Sub InsertaRango(psOpeCod As String, psRango As String, psDescrip As String, pnDesde As Integer, pnHasta As Integer)
On Error GoTo InsertaRangoErr
   psSql = "INSERT AnxRiesgosRango (cOpeCod, cCodRango, cDescrip, nDesde, nHasta) " _
         & "VALUES ('" & psOpeCod & "','" & psRango & "','" & psDescrip & "'," & pnDesde & "," & pnHasta & ")"
   dbConec.Ejecutar psSql
Exit Sub
InsertaRangoErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:InsertaRango Method")
End Sub

Public Sub ActualizaRango(psOpeCod As String, psRango As String, psDescrip As String, pnDesde As Integer, pnHasta As Integer)
On Error GoTo ActualizaRangoErr
   psSql = "UPDATE AnxRiesgosRango SET cDescrip = '" & psDescrip & "', nDesde = " & pnDesde & ", nHasta = " & pnHasta & " WHERE cOpeCod = '" & psOpeCod & "' and cCodRango = '" & psRango & "'"
   dbConec.Ejecutar psSql
Exit Sub
ActualizaRangoErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:ActualizaRango Method")
End Sub

Public Function CargaCriterioTasa(psOpeCod As String, psMoneda As String, psCodigo As String, psCtaCod As String, psFecha As String) As Recordset
On Error GoTo CargaCriterioTasaErr
Dim psMonedaCont As String
If psMoneda = "2" Then
    psMonedaCont = psMoneda
Else
    psMonedaCont = "[16]"
End If
  psSql = "SELECT anx.cCodRango, SUM(bal.nCtaSaldoImporte), SUM(bal.nCtaSaldoImporte)*convert(money,anx.cValor)/100 nValor  " _
        & "FROM CtaSaldo bal, AnxRiesgosCriterio anx " _
        & "WHERE cCtaContCod Like '__" & psMonedaCont & "%' and cCtaContCod like '" & psCtaCod & "%' " _
        & "  and anx.cOpeCod = '" & psOpeCod & "' and dCtaSaldoFecha = (SELECT Max(dCtaSaldoFecha) FROM CtaSaldo WHERE cCtaContCod = bal.cCtaContCod and dCtaSaldoFecha <= '" & psFecha & "' And cCtaContCod Like '__" & psMonedaCont & "%' and cCtaContCod like '" & psCtaCod & "%') " _
        & "  and anx.cCodClase + anx.cCodGrp + anx.cCodConcep = '" & psCodigo & "' GROUP BY anx.cCodRango, anx.cValor " _
        & "ORDER BY anx.cCodRango "
Set CargaCriterioTasa = dbConec.CargaRecordSet(psSql)
    Exit Function
CargaCriterioTasaErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaCriterioTasa Method")
End Function

Public Function CargaCriterioSeries(psOpeCod As String, psMoneda As String, psCodigo As String, psCtaCod As String, psFecha As String) As Recordset
On Error GoTo CargaCriterioSeriesErr
Dim psMonedaCont As String
If psMoneda = "2" Then
    psMonedaCont = psMoneda
Else
    psMonedaCont = "[16]"
End If

psSql = "SELECT  anx.cCodRango, SUM(ISNULL(case when cCtaContCod LIKE '%'+SUBSTRING(anx.cValor,2,22) and LEFT(anx.cValor,1) = 'S' THEN bal.nCtaSaldoImporte " _
     & "when not cCtaContCod LIKE '%'+SUBSTRING(anx.cValor,2,22) and LEFT(anx.cValor,1) = 'N' THEN bal.nCtaSaldoImporte " _
     & "else 0 end,0)) nValor " _
     & "FROM CtaSaldo bal, AnxRiesgosCriterio anx " _
     & "WHERE bal.cctacontcod like '" & Left(psCtaCod, 2) & psMonedaCont & Mid(psCtaCod, 4, 22) & "%' and cOpeCod = '" & psOpeCod & "' and dCtaSaldoFecha = (SELECT Max(dCtaSaldoFecha) FROM CtaSaldo WHERE cCtaContCod = bal.cCtaContCod and dCtaSaldoFecha <= '" & psFecha & "') " _
     & "      and anx.cCodClase + anx.cCodGrp + anx.cCodConcep = '" & psCodigo & "' " _
     & "GROUP BY anx.cCodRango ORDER BY anx.cCodRango "
Set CargaCriterioSeries = dbConec.CargaRecordSet(psSql)
    Exit Function
CargaCriterioSeriesErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaCriterioSeries Method")
End Function

Public Function CargaCriterioEncaje(psOpeCod As String, psMoneda As String, psCodigo As String, psCtaCod As String, psFecha As String, pnEncaje As Currency) As Recordset
On Error GoTo CargaCriterioEncajeErr
Dim psMonedaCont As String
If psMoneda = "2" Then
    psMonedaCont = psMoneda
Else
    psMonedaCont = "[16]"
End If
psSql = "SELECT  anx.cCodRango, SUM(ISNULL(CASE WHEN anx.cValor = '<BCR' and bal.nCtaSaldoImporte < " & pnEncaje & " THEN bal.nCtaSaldoImporte " _
     & "        WHEN anx.cValor = '<BCR' and bal.nCtaSaldoImporte > " & pnEncaje & " THEN " & pnEncaje _
     & "        WHEN anx.cValor = '>BCR' and bal.nCtaSaldoImporte > " & pnEncaje & " THEN bal.nCtaSaldoImporte - " & pnEncaje _
     & "        ELSE 0 END,0)) nValor " _
     & "FROM CtaSaldo bal, AnxRiesgosCriterio anx " _
     & "where cctacontcod like '" & Left(psCtaCod, 2) & psMonedaCont & Mid(psCtaCod, 4, 22) & "%' and anx.cOpeCod = '" & psOpeCod & "' and  dCtaSaldoFecha = (SELECT Max(dCtaSaldoFecha) FROM CtaSaldo WHERE cCtaContCod = bal.cCtaContCod and dCtaSaldoFecha <= '" & psFecha & "') " _
     & "      and anx.cCodClase + anx.cCodGrp + anx.cCodConcep = '" & psCodigo & "' " _
     & "GROUP BY anx.cCodRango ORDER BY anx.cCodRango "

Set CargaCriterioEncaje = dbConec.CargaRecordSet(psSql)
    Exit Function
CargaCriterioEncajeErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaCriterioEncaje Method")
End Function

Public Function CargaCriterioCheque(psOpeCod As String, psMoneda As String, psCodigo As String, psFecha As String) As Recordset
On Error GoTo CargaCriterioChequeErr
psSql = "SELECT anx.cCodRango, SUM(ISNULL(CASE WHEN anx.cValor = 'PL' and a.bPlaza = '" & gChqPlazaLocal & "' THEN nMonto " _
     & "WHEN anx.cValor = 'PO' and a.bPlaza = '" & gChqPlazaOtros & "' THEN nMonto     ELSE 0 END ,0)) nValor " _
      & "FROM DocRec a JOIN MOVDoc MD ON MD.nDocTpo = A.nTpoDoc And MD.cDocNro = A.cNroDoc " _
      & "     JOIN MOV M ON M.nMovNro = MD.nMovNro JOIN DocRecEst drc ON drc.cNroDoc = a.cNroDoc and drc.nTpoDoc = a.nTpoDoc and drc.cPersCod = a.cPersCod and drc.cIFCta = a.cIFCta and drc.cMovNro = m.cMovNro, " _
      & "     AnxRiesgosCriterio anx " _
      & "WHERE anx.cOpeCod = '" & psOpeCod & "' and a.nMoneda = " & psMoneda & " and drc.nEstado in (" & gChqEstRegistrado & "," & gChqEstEnValorizacion & ") and M.nMovFlag not in (1) and datediff(dd, MD.dDocFecha, '" & Format(psFecha, gsFormatoFecha) & "') = 0" _
             & " and (a.cDepIF = '0' or not exists" _
             & "       (select mr.nMovNro from movref mr where mr.nMovNroRef = M.nMovNro" _
             & "            and substring(M.cMovNro,1,8) < '" & Format(psFecha, gsFormatoMovFecha) & "') )" _
      & "  and anx.cCodClase + anx.cCodGrp + anx.cCodConcep = '" & psCodigo & "' " _
      & "GROUP BY anx.cCodRango ORDER BY anx.cCodRango"

Set CargaCriterioCheque = dbConec.CargaRecordSet(psSql)
    Exit Function
CargaCriterioChequeErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaCriterioCheque Method")
End Function

Public Function CargaCriterioFecha(psOpeCod As String, psMoneda As String, psCodigo As String, psCtaCod As String, psFecha As String, psFechaRep As String) As Recordset
On Error GoTo CargaCriterioFechaErr
Dim psMonedaCont As String
If psMoneda = "2" Then
    psMonedaCont = psMoneda
Else
    psMonedaCont = "[16]"
End If

   psSql = "SELECT Top 1 Ran.cCodRango, ISNULL(SUM(bal.nCtaSaldoImporte),0) nValor " _
      & "FROM CtaSaldo bal, AnxRiesgosRango Ran " _
      & "WHERE ran.cOpeCod = '" & psOpeCod & "' and cctacontcod LIKE '" & Left(psCtaCod, 2) & psMonedaCont & Mid(psCtaCod, 4, 22) & "%' and  dCtaSaldoFecha = (SELECT Max(dCtaSaldoFecha) FROM CtaSaldo WHERE cCtaContCod = bal.cCtaContCod and dCtaSaldoFecha <= '" & psFechaRep & "') " _
      & "      and datediff(d,'" & psFechaRep & "','" & psFecha & "') BETWEEN Ran.nDesde and Ran.nHasta " _
      & "      and datediff(d,'" & psFechaRep & "','" & psFecha & "') > 0 GROUP BY Ran.cCodRango"
Set CargaCriterioFecha = dbConec.CargaRecordSet(psSql)
    Exit Function
CargaCriterioFechaErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaCriterioFecha Method")
End Function

Public Function CargaCriterioVencBancos(psOpeCod As String, psMoneda As String, psCodigo As String, psCtaCod As String, psObjCod As String, psFecha As String, pnTipCambio As Currency, pnTipoIF As CGTipoIF) As Recordset
On Error GoTo CargaCriterioVencBancosErr
psSql = "SELECT anx.cCodRango, SUM(ISNULL(" & IIf(psMoneda = "1", "co.nSaldo", "co.nSaldoME") & ",0)) nValor " _
      & "FROM   CtaIF cb JOIN CtaIFSaldo co ON co.cPersCod = cb.cPersCod and co.cIFTpo=cb.cIFTpo and co.cCtaIFCod=cb.cCtaIFCod, " _
      & "       AnxRiesgosRango anx " _
      & "WHERE anx.cOpeCod = '" & psOpeCod & "' and LEFT(cb.cCtaIFCOD,2) = '" & Format(gTpoCtaIFCtaPF, "00") & "' and cb.cIFTpo = '" & Format(pnTipoIF, "00") & "' and co.cCtaContCod LIKE '" & psCtaCod & "' " _
      & "  and co.dCtaIFSaldo = " _
      & "  (SELECT MAX(dCtaIFSaldo) FROM CtaIFSaldo " _
      & "   Where co.cPersCod = cPersCod And co.cIFTpo = cIFTpo And co.cCtaIFCod = cCtaIFCod " _
      & "     and SubString(cb.cCtaIfCod,3,1) = " & psMoneda & " and cCtaContCod = co.cCtaContCod and dCtaIFSaldo <= '" & psFecha & "') " _
      & "  and datediff(d,'" & psFecha & "',dCtaIFVenc) BETWEEN nDesde and nHasta " _
      & "  and datediff(d,'" & psFecha & "',dCtaIFVenc) > 0 " _
      & "GROUP BY anx.cCodRango ORDER BY anx.cCodRango "

Set CargaCriterioVencBancos = dbConec.CargaRecordSet(psSql)
    Exit Function
CargaCriterioVencBancosErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaCriterioVencBancos Method")
End Function

Public Function CargaCriterioVencCredVig(psOpeCod As String, psMoneda As String, psCtaCod As String, psFecha As String, pnTipCambio As Currency) As Recordset
On Error GoTo CargaCriterioVencCredVigErr

If Left(psCtaCod, 1) = "N" Then
   psSql = "SELECT ISNULL(SUM(ROUND(nCtaSaldoImporte/" & IIf(psMoneda = "1", 1, pnTipCambio) & ",2)),0) nSaldo " _
        & "FROM CtaSaldo Bal WHERE  dCtaSaldoFecha = (SELECT Max(dCtaSaldoFecha) FROM CtaSaldo WHERE cCtaContCod = bal.cCtaContCod and dCtaSaldoFecha <= '" & psFecha & "') " _
        & "     and SubString(cCtaContCod,3,1) = '" & psMoneda & "' and cCtaContCod LIKE '" & Mid(psCtaCod, 2, Len(psCtaCod) - 3) & "%' and cCtaContCod NOT LIKE '" & Mid(psCtaCod, 2, 22) & "%'"
Else
   psSql = "SELECT ISNULL(SUM(ROUND(nCtaSaldoImporte/" & IIf(psMoneda = "1", 1, pnTipCambio) & ",2)),0) nSaldo " _
        & "FROM CtaSaldo Bal WHERE  dCtaSaldoFecha = (SELECT Max(dCtaSaldoFecha) FROM CtaSaldo WHERE cCtaContCod = bal.cCtaContCod and dCtaSaldoFecha <= '" & psFecha & "') " _
        & "     and cCtaContCod LIKE '14" & psMoneda & Mid(psCtaCod, 4, 22) & "%'"
End If
  
Set CargaCriterioVencCredVig = dbConec.CargaRecordSet(psSql)
Exit Function
CargaCriterioVencCredVigErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaCriterioVencCredVig Method")
End Function

Public Function GetSumaEstadRiesgos(psMoneda As String, psFecha As String, Optional psProd As String = "") As Currency
On Error GoTo GetSumaEstadRiesgosErr
  psSql = "SELECT ISNULL(SUM(nSaldo),0) nSaldo FROM CredEstadRiesgos " _
        & "WHERE " & IIf(psProd = "", "", "cProd LIKE '" & psProd & "' and ") _
        & "      cEstado IN ('V','P') and cMoneda = '" & psMoneda & "' and " _
        & "      dFecha = '" & psFecha & "' and dFecVenc >= '" & psFecha & "'"
Set prs = dbConec.CargaRecordSet(psSql)
If Not prs.EOF Then
   GetSumaEstadRiesgos = IIf(IsNull(prs!nSaldo), 0, prs!nSaldo)
Else
   GetSumaEstadRiesgos = 0
End If
Exit Function
GetSumaEstadRiesgosErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:GetSumaEstadRiesgos Method")
End Function

Public Function CargaCriterioVencEstadCred(psOpeCod As String, psMoneda As String, psFecha As String, psEstado As String, Optional psProd As String = "", Optional pnTotal As Currency = 1) As Recordset
On Error GoTo CargaCriterioVencEstadCredErr
   psSql = "SELECT cCodRango, SUM(nSaldo) nValor, SUM(nSaldo) / " & pnTotal & " nPorcentaje " _
         & "FROM credestadriesgos r, AnxRiesgosRango ran " _
         & "WHERE  ran.cOpeCod = '" & psOpeCod & "' and " & IIf(psProd = "", "", "r.cProd LIKE '" & psProd & "' and ") _
         & "       cEstado IN (" & psEstado & ") and dFecha = '" & psFecha & "' and " _
         & "       cMoneda = '" & psMoneda & "'  and ( datediff(d,'" & psFecha & "', r.dFecVenc) BETWEEN ran.nDesde and ran.nHasta ) and " _
         & "       dFecVenc >= '" & psFecha & "' " _
         & "group by cCodRango ORDER BY cCodRango"

Set CargaCriterioVencEstadCred = dbConec.CargaRecordSet(psSql)
    Exit Function
CargaCriterioVencEstadCredErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaCriterioVencEstadCred Method")
End Function

Public Function CargaCriterioVencAdeuda(psOpeCod As String, psMoneda As String, psCtaCod As String, psObjCod As String, psFecha As String, Optional pnTipCambio As Currency = 1, Optional pnMontoTotal As Currency = 1) As Recordset
On Error GoTo CargaCriterioVencAdeudaErr

psSql = "SELECT anx.cCodRango, SUM(ROUND(CASE WHEN SubString(co.cCtaIFCod,3,1) = '1' and co.cMonedaPago = '2' THEN ROUND(ccal.nCapital*v.nIndiceVac,2) ELSE cCal.nCapital end ,2) ) nValor, SUM(ROUND(CASE WHEN SubString(co.cCtaIFCod,3,1) = '1' and co.cMonedaPago = '2' THEN ROUND(ccal.nCapital*v.nIndiceVac,2) ELSE cCal.nCapital end ,2)) / " & pnMontoTotal & " nPorcentaje " _
      & "FROM   CtaIF cb JOIN CtaIFAdeudados co ON co.cPersCod = cb.cPersCod and co.cIFTpo=cb.cIFTpo and co.cCtaIFCod=cb.cCtaIFCod " _
      & "                JOIN CtaIFCalendario cCal ON cCal.cPersCod = cb.cPersCod and cCal.cIFTpo=cb.cIFTpo and cCal.cCtaIFCod=cb.cCtaIFCod, " _
      & "       AnxRiesgosRango anx, IndiceVac v " _
      & "WHERE anx.cOpeCod = '" & psOpeCod & "' and LEFT(cb.cCtaIFCOD,2) = '" & Format(gTpoCtaIFCtaAdeud, "00") & "'  and SubString(cb.cCtaIFCod,3,1) = '" & psMoneda & "' " _
      & "  and datediff(d,'" & psFecha & "',cCal.dVencimiento) BETWEEN nDesde and nHasta " _
      & "  and datediff(d,'" & psFecha & "',cCal.dVencimiento) > 0 and cb.dCtaIFAper <= '" & psFecha & "' " _
      & "      and v.dIndiceVac = '" & psFecha & "' " & IIf(psCtaCod = "2[46]", "", " and datediff(d,'" & psFecha & "',cCal.dVencimiento) " _
      & "  " & IIf(Right(psCtaCod, 1) = "4", "<=", ">") & "360") & " and cb.cCtaIFEstado IN (1,2) " _
      & "GROUP BY anx.cCodRango ORDER BY anx.cCodRango"

Set CargaCriterioVencAdeuda = dbConec.CargaRecordSet(psSql)
Exit Function
CargaCriterioVencAdeudaErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaCriterioVencAdeuda Method")
End Function
      
Public Function CargaCriterioMonto(psOpeCod As String, pnMonto As Currency, psCodigo As String) As Recordset
On Error GoTo CargaCriterioMontoErr
  psSql = "SELECT anx.cCodRango, ROUND(" & pnMonto & "*convert(money,anx.cValor)/100,2) nValor " _
        & "FROM AnxRiesgosCriterio anx " _
        & "WHERE anx.cOpeCod = '" & psOpeCod & "' and anx.cCodClase + anx.cCodGrp + anx.cCodConcep = '" & psCodigo & "' " _
        & "ORDER BY anx.cCodRango"
Set CargaCriterioMonto = dbConec.CargaRecordSet(psSql)
    Exit Function
CargaCriterioMontoErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaCriterioMonto Method")
End Function

Public Function CargaCriterioPorcentualSegunCriterio(psOpeCod As String, psMoneda As String, psCodigo As String, psCtaCod As String, psFecha As String) As Recordset
On Error GoTo CargaCriterioTasaErr
Dim psMonedaCont As String
Dim rsCriterios As ADODB.Recordset
Set rsCriterios = New ADODB.Recordset
Dim rsCtaCont As ADODB.Recordset
Set rsCtaCont = New ADODB.Recordset

Dim rsRes As ADODB.Recordset
Set rsRes = New ADODB.Recordset

Dim rsResII As ADODB.Recordset
Set rsResII = New ADODB.Recordset

Dim lnMontoTotalMN As Currency
Dim lnMontoTotalME As Currency

If psMoneda = "2" Then
    psMonedaCont = psMoneda
Else
    psMonedaCont = "[16]"
End If
  
  If psCtaCod = "11_303__02%" Then   'Plazos Fijos de CMACs
     psSql = "SELECT anx.cCodRango, SUM(bal.nCtaSaldoImporte), SUM(bal.nCtaSaldoImporte)*convert(money,anx.cValor)/100 nValor  " _
           & "FROM CtaSaldo bal, AnxRiesgosCriterio anx " _
           & "WHERE SubString(cCtaContCod,3,1) like '" & psMonedaCont & "' and cctacontcod like '" & psCtaCod & "%'   " _
           & "  and anx.cOpeCod = '" & psOpeCod & "' and dCtaSaldoFecha = (SELECT Max(dCtaSaldoFecha) FROM CtaSaldo WHERE cCtaContCod = bal.cCtaContCod and dCtaSaldoFecha <= '" & psFecha & "') " _
           & "  and anx.cCodClase + anx.cCodGrp + anx.cCodConcep = '" & psCodigo & "' GROUP BY anx.cCodRango, anx.cValor " _
           & "ORDER BY anx.cCodRango "
  
    psSql = "Select sum(dbo.getsaldocta('" & psFecha & "',cCtacontCod,1)) TotalMN, sum(dbo.getsaldocta('" & psFecha & "',cCtacontCod,2)) TotalME From ctacont where cctacontcod like '" & Left(psCtaCod, 2) & psMonedaCont & Mid(psCtaCod, 4) & "'"
    Set rsCtaCont = dbConec.CargaRecordSet(psSql)
    
    psSql = " Select Sum(nSaldo) SaldoMN, Sum(nSaldoME) SaldoME" _
          & " from CtaIf a" _
          & " Inner Join CtaIFSaldo b On a.cCtaIfCod = b.cCtaIFCod And a.cIFTpo = b.cIFTpo And a.cPersCod = b.cPersCod" _
          & " where a.cIFTpo = '03' And  a.cCtaIfCod Like '03" & psMoneda & "%' and b.cCtaContCod like '" & psCtaCod & "' and dCtaIFSaldo <= '" & psFecha & "' And DateDiff(Day,'" & psFecha & "',dCtaIFVenc) >= 0 " _
          & " and dCtaIFSaldo = (SELECT Max(dCtaIFSaldo) FROM CtaIFSaldo" _
          & " WHERE cCtaContCod = b.cCtaContCod and cPersCod = a.cPersCod and cIFTpo = a.cIFTpo and cCtaIFCod = a.cCtaIFCod and dCtaIFSaldo <= '" & psFecha & "')"
    Set rsCriterios = dbConec.CargaRecordSet(psSql)
    
    If IsNull(rsCriterios!SaldoMN) Then
        lnMontoTotalMN = 1
    Else
        If rsCriterios!SaldoMN = 0 Then
            lnMontoTotalMN = 1
        Else
            lnMontoTotalMN = rsCriterios!SaldoMN
        End If
    End If
    
    If IsNull(rsCriterios!SaldoME) Then
        lnMontoTotalME = 1
    Else
        If rsCriterios!SaldoME = 0 Then
            lnMontoTotalME = 1
        Else
            lnMontoTotalME = rsCriterios!SaldoME
        End If
    End If
    
    psSql = " Select cCodRango, " & rsCtaCont!TotalMN & " * (Sum(SaldoMN) / " & lnMontoTotalMN & ")  nValor, " & rsCtaCont!TotalME & " * ( Sum(SaldoME) / " & lnMontoTotalME & ") nValorME from  " _
          & " (Select Sum(nSaldo)  SaldoMN, Sum(nSaldoME) SaldoME , Case when DateDiff(Day,'" & psFecha & "',dCtaIFVenc) = 0 then 1 else DateDiff(Day,'" & psFecha & "',dCtaIFVenc) end Dias from CtaIf a" _
          & " Inner Join CtaIFSaldo b On a.cCtaIfCod = b.cCtaIFCod And a.cIFTpo = b.cIFTpo And a.cPersCod = b.cPersCod" _
          & " where a.cIFTpo = '03' And  a.cCtaIfCod Like '03" & psMoneda & "%' and b.cCtaContCod like '" & psCtaCod & "' and dCtaIFSaldo <= '" & psFecha & "' And DateDiff(Day,'" & psFecha & "',dCtaIFVenc) >= 0 " _
          & " and dCtaIFSaldo = (SELECT Max(dCtaIFSaldo) FROM CtaIFSaldo" _
          & " WHERE cCtaContCod = b.cCtaContCod and cPersCod = a.cPersCod and cIFTpo = a.cIFTpo and cCtaIFCod = a.cCtaIFCod and dCtaIFSaldo <= '" & psFecha & "')" _
          & " group by Case when DateDiff(Day,'" & psFecha & "',dCtaIFVenc) = 0 then 1 else DateDiff(Day,'" & psFecha & "',dCtaIFVenc) end) As AAA" _
          & " Inner Join (Select * from AnxRiesgosRango where copecod = '" & psOpeCod & "') As BBB On Dias >= nDesde And Dias <= nHasta" _
          & " group by cCodRango"
  
  ElseIf psCtaCod = "11_301__03%" Then   'Plazos Fijos de Bancos
    psSql = "Select sum(dbo.getsaldocta('" & psFecha & "',cCtacontCod,1)) TotalMN, sum(dbo.getsaldocta('" & psFecha & "',cCtacontCod,2)) TotalME From ctacont where cctacontcod like '" & Left(psCtaCod, 2) & psMonedaCont & Mid(psCtaCod, 4) & "'"
    Set rsCtaCont = dbConec.CargaRecordSet(psSql)
    
    psSql = " Select Sum(nSaldo) SaldoMN, Sum(nSaldoME) SaldoME" _
          & " from CtaIf a" _
          & " Inner Join CtaIFSaldo b On a.cCtaIfCod = b.cCtaIFCod And a.cIFTpo = b.cIFTpo And a.cPersCod = b.cPersCod" _
          & " where a.cIFTpo = '01' And  a.cCtaIfCod Like '03" & psMoneda & "%' and b.cCtaContCod like '" & psCtaCod & "' and dCtaIFSaldo <= '" & psFecha & "' And DateDiff(Day,'" & psFecha & "',dCtaIFVenc) >= 0 " _
          & " and dCtaIFSaldo = (SELECT Max(dCtaIFSaldo) FROM CtaIFSaldo" _
          & " WHERE cCtaContCod = b.cCtaContCod and cPersCod = a.cPersCod and cIFTpo = a.cIFTpo and cCtaIFCod = a.cCtaIFCod and dCtaIFSaldo <= '" & psFecha & "')"
    Set rsCriterios = dbConec.CargaRecordSet(psSql)
    
    If IsNull(rsCriterios!SaldoMN) Then
        lnMontoTotalMN = 1
    Else
        If rsCriterios!SaldoMN = 0 Then
            lnMontoTotalMN = 1
        Else
            lnMontoTotalMN = rsCriterios!SaldoMN
        End If
    End If
    
    If IsNull(rsCriterios!SaldoME) Then
        lnMontoTotalME = 1
    Else
        If rsCriterios!SaldoME = 0 Then
            lnMontoTotalME = 1
        Else
            lnMontoTotalME = rsCriterios!SaldoME
        End If
    End If
    
    psSql = " Select cCodRango, " & rsCtaCont!TotalMN & " * (Sum(SaldoMN) / " & lnMontoTotalMN & ")  nValor, " & rsCtaCont!TotalME & " * ( Sum(SaldoME) / " & lnMontoTotalME & ") nValorME from " _
          & " (Select Sum(nSaldo)  SaldoMN, Sum(nSaldoME) SaldoME , Case when DateDiff(Day,'" & psFecha & "',dCtaIFVenc) = 0 then 1 else DateDiff(Day,'" & psFecha & "',dCtaIFVenc) end Dias from CtaIf a" _
          & " Inner Join CtaIFSaldo b On a.cCtaIfCod = b.cCtaIFCod And a.cIFTpo = b.cIFTpo And a.cPersCod = b.cPersCod" _
          & " where a.cIFTpo = '01' And  a.cCtaIfCod Like '03" & psMoneda & "%' and b.cCtaContCod like '" & psCtaCod & "' and dCtaIFSaldo <= '" & psFecha & "'  And DateDiff(Day,'" & psFecha & "',dCtaIFVenc) >= 0 " _
          & " and dCtaIFSaldo = (SELECT Max(dCtaIFSaldo) FROM CtaIFSaldo" _
          & " WHERE cCtaContCod = b.cCtaContCod and cPersCod = a.cPersCod and cIFTpo = a.cIFTpo and cCtaIFCod = a.cCtaIFCod and dCtaIFSaldo <= '" & psFecha & "')" _
          & " group by Case when DateDiff(Day,'" & psFecha & "',dCtaIFVenc) = 0 then 1 else DateDiff(Day,'" & psFecha & "',dCtaIFVenc) end) As AAA" _
          & " Inner Join (Select * from AnxRiesgosRango where copecod = '" & psOpeCod & "') As BBB On Dias >= nDesde And Dias <= nHasta" _
          & " group by cCodRango"
  
  ElseIf psCtaCod = "23_3%" Then   'PLAZO Instit Financieras
    psSql = "Select sum(dbo.getsaldocta('" & psFecha & "',cCtacontCod,1)) TotalMN, sum(dbo.getsaldocta('" & psFecha & "',cCtacontCod,2)) TotalME From ctacont where cctacontcod like '" & Left(psCtaCod, 2) & psMonedaCont & Mid(psCtaCod, 4) & "'"
    Set rsCtaCont = dbConec.CargaRecordSet(psSql)

    psSql = " Select Sum(nSaldCntPF) Saldo From" _
          & " (Select nSaldCntPF, Case when Datediff(Day,'" & psFecha & "',DateAdd(day,nPlazo,dAperPF)) < 0  then  Datediff(Day,'" & psFecha & "',DateAdd(day,nPlazo,(Select dRenovacion from captacplazofijo where cCtaCod = PC.cCtaCod))) when Datediff(Day,'" & psFecha & "',DateAdd(day,nPlazo,dAperPF)) = 0 then 1 else Datediff(Day,'" & psFecha & "',DateAdd(day,nPlazo,dAperPF)) end  Dias from " & sServidorConsolidada & "PlazoFijoConsol PF" _
          & " Inner Join ProductoPersona PC On PC.cCtaCod = PF.cCtaCod And PC.nPrdPersRelac = 10" _
          & " Inner Join InstitucionFinanc IFI On PC.cPersCod = IFI.cPersCod" _
          & " where nEstCtaPF Not in (1300,1400) and substring(PC.cCtaCod,9,1) = '" & psMoneda & "') As AA"
    Set rsCriterios = dbConec.CargaRecordSet(psSql)
    
    If IsNull(rsCriterios!Saldo) Then
        lnMontoTotalMN = 1
    Else
        If rsCriterios!Saldo = 0 Then
            lnMontoTotalMN = 1
        Else
            lnMontoTotalMN = rsCriterios!Saldo
        End If
    End If
    
    psSql = " Select cCodRango,  " & rsCtaCont!TotalMN & " * (Sum(nSaldCntPF)/ " & lnMontoTotalMN & " )  nValor  From" _
          & " (Select nSaldCntPF, Case when Datediff(Day,'" & psFecha & "',DateAdd(day,nPlazo,dAperPF)) < 0  then  Datediff(Day,'" & psFecha & "',DateAdd(day,nPlazo,(Select dRenovacion from captacplazofijo where cCtaCod = PC.cCtaCod))) when Datediff(Day,'" & psFecha & "',DateAdd(day,nPlazo,dAperPF)) = 0 then 1 else Datediff(Day,'" & psFecha & "',DateAdd(day,nPlazo,dAperPF)) end  Dias from " & sServidorConsolidada & "PlazoFijoConsol PF" _
          & " Inner Join ProductoPersona PC On PC.cCtaCod = PF.cCtaCod And PC.nPrdPersRelac = 10" _
          & " Inner Join InstitucionFinanc IFI On PC.cPersCod = IFI.cPersCod" _
          & " where nEstCtaPF Not in (1300,1400) and substring(PC.cCtaCod,9,1) = '" & psMoneda & "') As AA" _
          & " Inner Join (Select * from AnxRiesgosRango where cOpeCod = '" & psOpeCod & "') As Anx On Dias >= nDesde And Dias <= nHasta group by cCodRango"
    
  ElseIf psCtaCod = "2[46]_[234579]0[12456789]%" Then   'ADEUDADOS
    psSql = "Select sum(dbo.getsaldocta('" & psFecha & "',cCtacontCod,1)) TotalMN, sum(dbo.getsaldocta('" & psFecha & "',cCtacontCod,2)) TotalME From ctacont where cctacontcod like '" & Left(psCtaCod, 5) & psMonedaCont & Mid(psCtaCod, 7) & "'"
    Set rsCtaCont = dbConec.CargaRecordSet(psSql)
    
    psSql = " Select Sum(CASE When CMONEDAPAGO = 2 And Substring(a.cCtaIFCod,3,1) = 1 then nCapital * dbo.GetVAC('" & psFecha & "') else nCapital END) Saldo From CtaIf a" _
          & " Inner Join CtaIfAdeudados b on a.cPersCod = b.cPersCod And a.cIFTpo = b.cIFTpo And a.cCtaIFCod = b.cCtaIFCod" _
          & " Inner Join CtaIfCalendario c on a.cPersCod = c.cPersCod And a.cIFTpo = c.cIFTpo And a.cCtaIFCod = c.cCtaIFCod" _
          & " Where c.dVencimiento > '" & psFecha & "' And a.cCtaIFCod Like '05" & psMoneda & "%'"
    Set rsCriterios = dbConec.CargaRecordSet(psSql)
    
    If IsNull(rsCriterios!Saldo) Then
        lnMontoTotalMN = 1
    Else
        If rsCriterios!Saldo = 0 Then
            lnMontoTotalMN = 1
        Else
            lnMontoTotalMN = rsCriterios!Saldo
        End If
    End If
    
    psSql = " Select cCodRango, Round(" & rsCtaCont!TotalMN & " * (Sum(Saldo) / " & lnMontoTotalMN & " ),2) nValor from" _
          & " ( Select CASE When CMONEDAPAGO = 2 And Substring(a.cCtaIFCod,3,1) = 1 then nCapital * dbo.GetVAC('" & psFecha & "') else nCapital END Saldo, Datediff(day,'" & psFecha & "',c.dVencimiento) Dias From CtaIf a " _
          & "   Inner Join CtaIfAdeudados b on a.cPersCod = b.cPersCod And a.cIFTpo = b.cIFTpo And a.cCtaIFCod = b.cCtaIFCod" _
          & "   Inner Join CtaIfCalendario c on a.cPersCod = c.cPersCod And a.cIFTpo = c.cIFTpo And a.cCtaIFCod = c.cCtaIFCod" _
          & "   Where c.dVencimiento > '" & psFecha & "' And a.cCtaIFCod Like '05" & psMoneda & "%') As AAA" _
          & "   Inner Join (Select * from AnxRiesgosRango where cOpeCod = '" & psOpeCod & "') As Anx On Dias >= nDesde And Dias <= nHasta" _
          & " group by cCodRango"
  
  ElseIf psCtaCod = "71_2%" Then   'CARTA FIANZA
    psSql = "Select sum(dbo.getsaldocta('" & psFecha & "',cCtacontCod,1)) TotalMN, sum(dbo.getsaldocta('" & psFecha & "',cCtacontCod,2)) TotalME From ctacont where cctacontcod like '" & Left(psCtaCod, 2) & psMonedaCont & Mid(psCtaCod, 4) & "'"
    Set rsCtaCont = dbConec.CargaRecordSet(psSql)
    
    psSql = " Select  Sum(nSaldoCap) Saldo from " & sServidorConsolidada & "CartaFianzaConsol a" _
          & " inner join " & sServidorConsolidada & "CartaFianzaSaldoConsol b on a.cctacod = b.cctacod" _
          & " where b.dfecha = '" & psFecha & "' and a.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205)" _
          & " and substring(a.cctacod,9,1) = '" & psMoneda & "'"
    Set rsCriterios = dbConec.CargaRecordSet(psSql)
    
    If IsNull(rsCriterios!Saldo) Then
        lnMontoTotalMN = 1
    Else
        If rsCriterios!Saldo = 0 Then
            lnMontoTotalMN = 1
        Else
            lnMontoTotalMN = rsCriterios!Saldo
        End If
    End If
    
    psSql = " Select cCodRango,  Round(" & rsCtaCont!TotalMN & " * (Sum(Saldo) / " & lnMontoTotalMN & " ),2) nValor from" _
          & " (Select  Sum (nSaldoCap) Saldo, datediff(day,'" & psFecha & "',dVencSol) Dias from " & sServidorConsolidada & "CartaFianzaConsol a" _
          & " inner join " & sServidorConsolidada & "CartaFianzaSaldoConsol b on a.cctacod = b.cctacod" _
          & " where b.dfecha = '" & psFecha & "' and a.nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205) " _
          & " and substring(a.cctacod,9,1) = '" & psMoneda & "' group by datediff(day,'" & psFecha & "',dVencSol)) as aa" _
          & " Inner Join (Select * from AnxRiesgosRango where cOpeCod = '" & psOpeCod & "') As Anx On Dias >= nDesde And Dias <= nHasta" _
          & " group by cCodRango"
  End If

Set rsRes = dbConec.CargaRecordSet(psSql)

If rsRes.EOF And rsRes.BOF And rsCtaCont!TotalMN <> 0 Then
    
    rsResII.Fields.Append "cCodRango", adVarChar, 6, adFldMayBeNull
    rsResII.Fields.Append "nValor", adCurrency, , adFldMayBeNull
    
    rsResII.Open , , adOpenStatic
    
    rsResII.AddNew
    rsResII!cCodRango = "01"
    rsResII!nValor = rsCtaCont!TotalMN
    rsResII.Update
    
    Set CargaCriterioPorcentualSegunCriterio = rsResII
Else
    Set CargaCriterioPorcentualSegunCriterio = rsRes
End If

    Exit Function
CargaCriterioTasaErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaCriterioTasa Method")
End Function

Public Function CargaCriterioEstadistico(psOpeCod As String, psMoneda As String, psCodigo As String, psCtaCod As String, psFecha As String) As Recordset
On Error GoTo CargaCriterioTasaErr
Dim psMonedaCont As String
Dim rsCriterios As ADODB.Recordset
Set rsCriterios = New ADODB.Recordset
Dim rsRes As ADODB.Recordset
Set rsRes = New ADODB.Recordset

Dim rsCtaCont As ADODB.Recordset
Set rsCtaCont = New ADODB.Recordset

Dim lnTotal As Double
Dim lnAcumulaT As Currency

If psMoneda = "2" Then
    psMonedaCont = psMoneda
Else
    psMonedaCont = "[16]"
End If
  
    rsCriterios.Fields.Append "cCodRango", adVarChar, 6, adFldMayBeNull
    rsCriterios.Fields.Append "nValor", adCurrency, , adFldMayBeNull
  
    psSql = ""
    If psCtaCod = "14_1" Then   'COLOCACIONES
        
        psSql = "Select sum(dbo.getsaldocta('" & Format(psFecha, gsFormatoFecha) & "',cctacontcod," & psMoneda & ")) Saldo From CtaCont Where cCtaContCod Like '14" & psMoneda & "[13478]%'"
        Set rsCtaCont = dbConec.CargaRecordSet(psSql)
        
        psSql = ""
        psSql = psSql & " SELECT  SUM(CASE WHEN cMoneda ='1' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 + REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END) AS NMONTOMN,"
        psSql = psSql & " SUM(CASE WHEN cMoneda ='2' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 + REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END) AS NMONTOME,"
        psSql = psSql & " RANGO1MN = SUM((CASE WHEN nDiasParaVenc <=30 AND cMoneda ='1' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO1ME = SUM((CASE WHEN nDiasParaVenc <=30 AND cMoneda ='2' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO2MN = SUM((CASE WHEN nDiasParaVenc BETWEEN 31 AND 60 AND cMoneda ='1' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO2ME = SUM((CASE WHEN nDiasParaVenc BETWEEN 31 AND 60 AND cMoneda ='2' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO3MN = SUM((CASE WHEN nDiasParaVenc BETWEEN 61 AND 90 AND cMoneda ='1' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO3ME = SUM((CASE WHEN nDiasParaVenc BETWEEN 61 AND 90 AND cMoneda ='2' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO4MN = SUM((CASE WHEN nDiasParaVenc BETWEEN 91 AND 180 AND cMoneda ='1' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO4ME = SUM((CASE WHEN nDiasParaVenc BETWEEN 91 AND 180 AND cMoneda ='2' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO5MN = SUM((CASE WHEN nDiasParaVenc BETWEEN 181 AND 360 AND cMoneda ='1' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO5ME = SUM((CASE WHEN nDiasParaVenc BETWEEN 181 AND 360 AND cMoneda ='2' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO6MN = SUM((CASE WHEN nDiasParaVenc >360 AND cMoneda ='1' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO6ME = SUM((CASE WHEN nDiasParaVenc >360 AND cMoneda ='2' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " SUM(CASE WHEN cMoneda ='1' THEN nIntDev ELSE 0 END ) AS NINTDEVSOL,"
        psSql = psSql & " SUM(CASE WHEN cMoneda ='2' THEN nIntDev ELSE 0 END ) AS NINTDEVDOL"
        psSql = psSql & " From"
        psSql = psSql & " (SELECT  cCtaCod, nPrdEstado, cRefinan, nSaldoCap, nCapVencido, nIntDev,"
        psSql = psSql & " nDiasAtraso, SUBSTRING(CCTACOD,9,1) AS cMoneda, nMontoDesemb,"
        psSql = psSql & " nDiasParaVenc = CASE    WHEN SUBSTRING(CCTACOD,6,3) ='305'"
        psSql = psSql & "             THEN DATEDIFF(day, '" & Format(psFecha, gsFormatoFecha) & "',dFecVenc)"
        psSql = psSql & "             ELSE (DATEDIFF(day, '" & Format(psFecha, gsFormatoFecha) & "',(Select MAX(dFecVenc) From " & sServidorConsolidada & "PlandespagConsol Where cCtaCod = C.cCtaCod))) END,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='1' AND nDiasAtraso<=15 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap else 0 end as VigCom,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='2' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap else 0 end as VigMes,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap else 0 end as VigCons,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap else 0 end as VigHip,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nDiasAtraso between 31 and 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap-nCapVencido else 0 end as VigCons31,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso between 31 and 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap-nCapVencido else 0 end as VigHip31,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='1' AND nDiasAtraso>15 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nSaldoCap else 0 end as VencCOM,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='2' AND nDiasAtraso>30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nSaldoCap else 0 end as VencMES,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND SUBSTRING(CCTACOD,6,3)<>'305' AND nDiasAtraso>90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nSaldoCap else 0 end as VencCons,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso>90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nSaldoCap else 0 end as VencHIP,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND SUBSTRING(CCTACOD,6,3)<>'305' AND nDiasAtraso BETWEEN 31 AND 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nCapVencido else 0 end as VencCons31,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso BETWEEN 31 AND 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nCapVencido else 0 end as VencHip31,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,3)='305' AND nDiasAtraso>30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nSaldoCap else 0 end as VencPrend,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='1' AND nDiasAtraso<=15 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap else 0 end as REFCom,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='2' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap else 0 end as REFMes,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap else 0 end as REFCons,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap else 0 end as REFHip,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nDiasAtraso between 31 and 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap-nCapVencido else 0 end as REFCons31,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso between 31 and 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap-nCapVencido else 0 end as REFHip31,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='1' AND nPrdEstado in (2201,2205) then nSaldoCap else 0 end as JUDCOM,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='2' AND nPrdEstado in (2201,2205) then nSaldoCap else 0 end as JUDMES,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nPrdEstado in (2201,2205) then nSaldoCap else 0 end as JUDCONS,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nPrdEstado in (2201,2205) then nSaldoCap else 0 end as JUDHIP"
        psSql = psSql & " FROM    " & sServidorConsolidada & "creditoconsol  C"
        psSql = psSql & " WHERE   nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205)"
        psSql = psSql & " AND cCtaCod NOT IN (SELECT CCTACOD FROM COLOCACCRED WHERE CRFA ='RFA')) AS DAT"
        psSql = psSql & " Where VigCom + VigMes + VigCons + VigHip + VigCons31 + VigHip31 + RefCom + REFMes + REFCons + REFHip + REFCons31 + REFHip31 > 0"
    
        Set rsRes = dbConec.CargaRecordSet(psSql)
        
        If Not (rsRes.EOF And rsRes.BOF) Then
            rsCriterios.Open , , adOpenStatic
            If psMoneda = "1" Then
            
               rsCriterios.AddNew
               rsCriterios!cCodRango = "01"
               rsCriterios!nValor = rsRes!RANGO1MN + Format(((rsRes!RANGO1MN / rsRes!NMONTOMN) * (rsCtaCont!Saldo - rsRes!NMONTOMN)), "#.00")
               rsCriterios.Update
               
               rsCriterios.AddNew
               rsCriterios!cCodRango = "02"
               rsCriterios!nValor = rsRes!RANGO2MN + Format(((rsRes!RANGO2MN / rsRes!NMONTOMN) * (rsCtaCont!Saldo - rsRes!NMONTOMN)), "#.00")
               rsCriterios.Update
               
               rsCriterios.AddNew
               rsCriterios!cCodRango = "03"
               rsCriterios!nValor = rsRes!RANGO3MN + Format(((rsRes!RANGO3MN / rsRes!NMONTOMN) * (rsCtaCont!Saldo - rsRes!NMONTOMN)), "#.00")
               rsCriterios.Update
               
               rsCriterios.AddNew
               rsCriterios!cCodRango = "04"
               rsCriterios!nValor = rsRes!RANGO4MN + Format(((rsRes!RANGO4MN / rsRes!NMONTOMN) * (rsCtaCont!Saldo - rsRes!NMONTOMN)), "#.00")
               rsCriterios.Update
               
               If psOpeCod = "770162" Then
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "05"
                    rsCriterios!nValor = (rsRes!RANGO5MN + Format(((rsRes!RANGO5MN / rsRes!NMONTOMN) * (rsCtaCont!Saldo - rsRes!NMONTOMN)), "#.00")) + (rsRes!RANGO6MN + Format(((rsRes!RANGO6MN / rsRes!NMONTOMN) * (rsCtaCont!Saldo - rsRes!NMONTOMN)), "#.00"))
                    rsCriterios.Update
               Else
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "05"
                    rsCriterios!nValor = rsRes!RANGO5MN + Format(((rsRes!RANGO5MN / rsRes!NMONTOMN) * (rsCtaCont!Saldo - rsRes!NMONTOMN)), "#.00")
                    rsCriterios.Update
                    
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "06"
                    rsCriterios!nValor = rsRes!RANGO6MN + Format(((rsRes!RANGO6MN / rsRes!NMONTOMN) * (rsCtaCont!Saldo - rsRes!NMONTOMN)), "#.00")
                    rsCriterios.Update
                End If
            Else
               rsCriterios.AddNew
               rsCriterios!cCodRango = "01"
               rsCriterios!nValor = rsRes!RANGO1ME + Format(((rsRes!RANGO1ME / rsRes!NMONTOME) * (rsCtaCont!Saldo - rsRes!NMONTOME)), "#.00")
               rsCriterios.Update
               
               rsCriterios.AddNew
               rsCriterios!cCodRango = "02"
               rsCriterios!nValor = rsRes!RANGO2ME + Format(((rsRes!RANGO2ME / rsRes!NMONTOME) * (rsCtaCont!Saldo - rsRes!NMONTOME)), "#.00")
               rsCriterios.Update
               
               rsCriterios.AddNew
               rsCriterios!cCodRango = "03"
               rsCriterios!nValor = rsRes!RANGO3ME + Format(((rsRes!RANGO3ME / rsRes!NMONTOME) * (rsCtaCont!Saldo - rsRes!NMONTOME)), "#.00")
               rsCriterios.Update
               
               rsCriterios.AddNew
               rsCriterios!cCodRango = "04"
               rsCriterios!nValor = rsRes!RANGO4ME + Format(((rsRes!RANGO4ME / rsRes!NMONTOME) * (rsCtaCont!Saldo - rsRes!NMONTOME)), "#.00")
               rsCriterios.Update
               
               If psOpeCod = "770162" Then
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "05"
                    rsCriterios!nValor = (rsRes!RANGO5ME + Format(((rsRes!RANGO5ME / rsRes!NMONTOME) * (rsCtaCont!Saldo - rsRes!NMONTOME)), "#.00")) + (rsRes!RANGO6ME + Format(((rsRes!RANGO6ME / rsRes!NMONTOME) * (rsCtaCont!Saldo - rsRes!NMONTOME)), "#.00"))
                    rsCriterios.Update
               Else
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "05"
                    rsCriterios!nValor = rsRes!RANGO5ME + Format(((rsRes!RANGO5ME / rsRes!NMONTOME) * (rsCtaCont!Saldo - rsRes!NMONTOME)), "#.00")
                    rsCriterios.Update
                    
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "06"
                    rsCriterios!nValor = rsRes!RANGO6ME + Format(((rsRes!RANGO6ME / rsRes!NMONTOME) * (rsCtaCont!Saldo - rsRes!NMONTOME)), "#.00")
                    rsCriterios.Update
               End If
            End If
            rsCriterios.MoveFirst
        End If
    Else
        Dim lsFiltro As String
        Dim lsSql As String
        Dim rsF As ADODB.Recordset
        Set rsF = New ADODB.Recordset
        Dim lnAcum As Currency
        If psCtaCod = "21_3" Then
            lsFiltro = " IN (1,2,3)"
            If psMoneda = 1 Then
                lsSql = "Select dbo.GetSaldoCtaAcumulado('" & psFecha & "','21" & psMoneda & "3%',1) - dbo.GetSaldoCtaAcumulado('" & psFecha & "','21" & psMoneda & "305%',1) "
            Else
                lsSql = "Select (dbo.GetSaldoCtaAcumulado('" & psFecha & "','21" & psMoneda & "3%',1) - dbo.GetSaldoCtaAcumulado('" & psFecha & "','21" & psMoneda & "305%',1) ) / (Select Top 1 nValFijo from tipocambio where dFecCamb = DateAdd(Day,1,'" & psFecha & "') order by dFecCamb Desc) "
            End If
            
        Else
            lsFiltro = " NOT IN (1,2,3)"
            If psMoneda = 1 Then
                lsSql = "Select dbo.GetSaldoCtaAcumulado('" & psFecha & "','23" & psMoneda & "3%',1)"
            Else
                lsSql = "Select dbo.GetSaldoCtaAcumulado('" & psFecha & "','23" & psMoneda & "3%',1) / (Select Top 1 nValFijo from tipocambio where dFecCamb = DateAdd(Day,1,'" & psFecha & "') order by dFecCamb Desc) "
            End If
        End If
        
        
        Set rsRes = dbConec.CargaRecordSet(lsSql)
        lnAcum = rsRes.Fields(0)
        
        psSql = psSql & " SELECT  RANGO1MN = SUM(CASE WHEN NPLAZO <=30 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END ) ,"
        psSql = psSql & " RANGO1ME = SUM(CASE WHEN NPLAZO <=30 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END ) ,"
        psSql = psSql & " RANGO2MN = SUM((CASE WHEN NPLAZO BETWEEN 31 AND 60 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO2ME = SUM((CASE WHEN NPLAZO BETWEEN 31 AND 60 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO3MN = SUM((CASE WHEN NPLAZO BETWEEN 61 AND 90 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO3ME = SUM((CASE WHEN NPLAZO BETWEEN 61 AND 90 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO4MN = SUM((CASE WHEN NPLAZO BETWEEN 91 AND 180 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO4ME = SUM((CASE WHEN NPLAZO BETWEEN 91 AND 180 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO5MN = SUM((CASE WHEN NPLAZO BETWEEN 181 AND 360 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO5ME = SUM((CASE WHEN NPLAZO BETWEEN 181 AND 360 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO6MN = SUM((CASE WHEN NPLAZO>360 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO6ME = SUM((CASE WHEN NPLAZO>360 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " SUM(CASE WHEN SUBSTRING(CCTACOD,9,1) ='1' THEN NSALDO ELSE 0 END) AS NSALDOMN,"
        psSql = psSql & " SUM(CASE WHEN SUBSTRING(CCTACOD,9,1) ='2' THEN NSALDO ELSE 0 END) AS NSALDOME"
        psSql = psSql & " From"
        psSql = psSql & " (SELECT SUBSTRING(A.CCTACOD,6,3) AS CPROD, A.CCTACOD, Case When aBS(DateDiff(Day,'" & Format(CDate(psFecha), gsFormatoFecha) & "',DateAdd(Day,A.NPLAZO,dRenovacion))) >= A.NPLAZO Then A.NPLAZO Else aBS(DateDiff(Day,'" & Format(CDate(psFecha), gsFormatoFecha) & "',DateAdd(Day,A.NPLAZO,dRenovacion))) End NPLAZO, nSaldCNTPF AS NSALDO,"
        psSql = psSql & " TIPOBLOQ=ISNULL((SELECT  TOP 1 NBLQMOTIVO"
        psSql = psSql & "         FROM    " & sServidorConsolidada & "PRODUCTOBLOQUEOSCONSOL PB"
        psSql = psSql & "          Where PB.cCtaCod = A.cCtaCod"
        psSql = psSql & "          AND CMOVNRODBL IS  NULL AND CONVERT(CHAR(8),DCIERRE,112) = '" & Format(CDate(psFecha), gsFormatoMovFecha) & "'"
        psSql = psSql & "          ),0)"
        psSql = psSql & " FROM    " & sServidorConsolidada & "PLAZOFIJOCONSOL A Inner Join captacplazofijo C On A.cCtaCod = C.cCtaCod  "
        psSql = psSql & " WHERE   nPersoneria " & lsFiltro
        psSql = psSql & " Union All"
        psSql = psSql & " SELECT  Top 0 SUBSTRING(CCTACOD,6,3) AS CPROD , CCTACOD, 720 AS NPLAZO, NSALDCNTCTS AS NSALDO,"
        psSql = psSql & "     TIPOBLOQ = 0 "
        psSql = psSql & "              "
        psSql = psSql & "             "
        psSql = psSql & "             "
        psSql = psSql & "             "
        psSql = psSql & " FROM  " & sServidorConsolidada & "CTSCONSOL A"
        psSql = psSql & " WHERE  nPersoneria " & lsFiltro
        psSql = psSql & " ) AS X"
        psSql = psSql & " WHERE TIPOBLOQ NOT IN (3,15)"
    
        Set rsRes = dbConec.CargaRecordSet(psSql)
        
        If Not (rsRes.EOF And rsRes.BOF) Then
            rsCriterios.Open , , adOpenStatic
            If psMoneda = "1" Then
               lnAcum = lnAcum - rsRes!RANGO1MN
               lnAcum = lnAcum - rsRes!RANGO2MN
               lnAcum = lnAcum - rsRes!RANGO3MN
               lnAcum = lnAcum - rsRes!RANGO4MN
               lnAcum = lnAcum - rsRes!RANGO5MN
               lnAcum = lnAcum - rsRes!RANGO6MN
               
               lnAcumulaT = rsRes!RANGO1MN + rsRes!RANGO2MN + rsRes!RANGO3MN + rsRes!RANGO4MN + rsRes!RANGO5MN + rsRes!RANGO6MN
               
               rsCriterios.AddNew
               rsCriterios!cCodRango = "01"
               rsCriterios!nValor = rsRes!RANGO1MN + Format(rsRes!RANGO1MN * lnAcum / lnAcumulaT, "0.00")
               rsCriterios.Update
               
               rsCriterios.AddNew
               rsCriterios!cCodRango = "02"
               rsCriterios!nValor = rsRes!RANGO2MN + Format(rsRes!RANGO2MN * lnAcum / lnAcumulaT, "0.00")
               rsCriterios.Update
               
               rsCriterios.AddNew
               rsCriterios!cCodRango = "03"
               rsCriterios!nValor = rsRes!RANGO3MN + Format(rsRes!RANGO3MN * lnAcum / lnAcumulaT, "0.00")
               rsCriterios.Update
               
               rsCriterios.AddNew
               rsCriterios!cCodRango = "04"
               rsCriterios!nValor = rsRes!RANGO4MN + Format(rsRes!RANGO4MN * lnAcum / lnAcumulaT, "0.00")
               rsCriterios.Update
               
               If psOpeCod = "770162" Then
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "05"
                    rsCriterios!nValor = rsRes!RANGO5MN + rsRes!RANGO6MN + Format(rsRes!RANGO5MN * lnAcum / lnAcumulaT, "0.00") + Format(rsRes!RANGO6MN * lnAcum / lnAcumulaT, "0.00")
                    rsCriterios.Update
               Else
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "05"
                    rsCriterios!nValor = rsRes!RANGO5MN + Format(rsRes!RANGO5MN * lnAcum / lnAcumulaT, "0.00")
                    rsCriterios.Update
                    
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "06"
                    rsCriterios!nValor = rsRes!RANGO6MN + Format(rsRes!RANGO6MN * lnAcum / lnAcumulaT, "0.00")
                    rsCriterios.Update
               End If
            Else
               lnAcum = lnAcum - rsRes!RANGO1ME
               lnAcum = lnAcum - rsRes!RANGO2ME
               lnAcum = lnAcum - rsRes!RANGO3ME
               lnAcum = lnAcum - rsRes!RANGO4ME
               lnAcum = lnAcum - rsRes!RANGO5ME
               lnAcum = lnAcum - rsRes!RANGO6ME
               
               lnAcumulaT = rsRes!RANGO1ME + rsRes!RANGO2ME + rsRes!RANGO3ME + rsRes!RANGO4ME + rsRes!RANGO5ME + rsRes!RANGO6ME
               If lnAcumulaT <> 0 Then
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "01"
                    rsCriterios!nValor = rsRes!RANGO1ME + Format(rsRes!RANGO1ME * lnAcum / lnAcumulaT, "0.00")
                    rsCriterios.Update
                    
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "02"
                    rsCriterios!nValor = rsRes!RANGO2ME + Format(rsRes!RANGO2ME * lnAcum / lnAcumulaT, "0.00")
                    rsCriterios.Update
                    
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "03"
                    rsCriterios!nValor = rsRes!RANGO3ME + Format(rsRes!RANGO3ME * lnAcum / lnAcumulaT, "0.00")
                    rsCriterios.Update
                    
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "04"
                    rsCriterios!nValor = rsRes!RANGO4ME + Format(rsRes!RANGO4ME * lnAcum / lnAcumulaT, "0.00")
                    rsCriterios.Update
                    
                    If psOpeCod = "770162" Then
                         rsCriterios.AddNew
                         rsCriterios!cCodRango = "05"
                         rsCriterios!nValor = rsRes!RANGO5ME + rsRes!RANGO6ME + Format(rsRes!RANGO5ME * lnAcum / lnAcumulaT, "0.00") + Format(rsRes!RANGO6ME * lnAcum / lnAcumulaT, "0.00")
                         rsCriterios.Update
                    Else
                         rsCriterios.AddNew
                         rsCriterios!cCodRango = "05"
                         rsCriterios!nValor = rsRes!RANGO5ME + Format(rsRes!RANGO5ME * lnAcum / lnAcumulaT, "0.00")
                         rsCriterios.Update
                         
                         rsCriterios.AddNew
                         rsCriterios!cCodRango = "06"
                         rsCriterios!nValor = rsRes!RANGO6ME + Format(rsRes!RANGO6ME * lnAcum / lnAcumulaT, "0.00")
                         rsCriterios.Update
                    End If
                End If
            End If
            If Not (rsCriterios.EOF And rsCriterios.BOF) Then rsCriterios.MoveFirst
        End If
    End If

Set CargaCriterioEstadistico = rsCriterios
    Exit Function
CargaCriterioTasaErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaCriterioTasa Method")
End Function

Public Function CargaCriterioEstadistico07(psOpeCod As String, psMoneda As String, psCodigo As String, psCtaCod As String, psFecha As String) As Recordset
On Error GoTo CargaCriterioTasaErr
Dim psMonedaCont As String
Dim rsCriterios As ADODB.Recordset
Set rsCriterios = New ADODB.Recordset
Dim rsRes As ADODB.Recordset
Set rsRes = New ADODB.Recordset

Dim rsCtaCont As ADODB.Recordset
Set rsCtaCont = New ADODB.Recordset

Dim lnAcumulaT As Currency

Dim lnTotal As Double

If psMoneda = "2" Then
    psMonedaCont = psMoneda
Else
    psMonedaCont = "[16]"
End If
  
    rsCriterios.Fields.Append "cCodRango", adVarChar, 6, adFldMayBeNull
    rsCriterios.Fields.Append "nValor", adCurrency, , adFldMayBeNull
  
    psSql = ""
    If psCtaCod = "14_1" Then   'COLOCACIONES
        
        psSql = "Select sum(dbo.getsaldocta('" & Format(psFecha, gsFormatoFecha) & "',cctacontcod," & psMoneda & ")) Saldo From CtaCont Where cCtaContCod Like '14" & psMoneda & "[13478]%'"
        Set rsCtaCont = dbConec.CargaRecordSet(psSql)
        
        psSql = ""
        psSql = psSql & " SELECT  SUM(CASE WHEN cMoneda ='1' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 + REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END) AS NMONTOMN,"
        psSql = psSql & " SUM(CASE WHEN cMoneda ='2' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 + REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END) AS NMONTOME,"
        psSql = psSql & " RANGO1MN = SUM((CASE WHEN nDiasParaVenc <=30 AND cMoneda ='1' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO1ME = SUM((CASE WHEN nDiasParaVenc <=30 AND cMoneda ='2' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO2MN = SUM((CASE WHEN nDiasParaVenc BETWEEN 31 AND 60 AND cMoneda ='1' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO2ME = SUM((CASE WHEN nDiasParaVenc BETWEEN 31 AND 60 AND cMoneda ='2' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO3MN = SUM((CASE WHEN nDiasParaVenc BETWEEN 61 AND 90 AND cMoneda ='1' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO3ME = SUM((CASE WHEN nDiasParaVenc BETWEEN 61 AND 90 AND cMoneda ='2' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO4MN = SUM((CASE WHEN nDiasParaVenc BETWEEN 91 AND 180 AND cMoneda ='1' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO4ME = SUM((CASE WHEN nDiasParaVenc BETWEEN 91 AND 180 AND cMoneda ='2' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO5MN = SUM((CASE WHEN nDiasParaVenc BETWEEN 181 AND 360 AND cMoneda ='1' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO5ME = SUM((CASE WHEN nDiasParaVenc BETWEEN 181 AND 360 AND cMoneda ='2' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO6MN = SUM((CASE WHEN nDiasParaVenc >360 AND cMoneda ='1' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " RANGO6ME = SUM((CASE WHEN nDiasParaVenc >360 AND cMoneda ='2' THEN VigCom+ VigMes+ VigCons+ VigHip + VigCons31 +VigHip31 +REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31 ELSE 0 END)),"
        psSql = psSql & " SUM(CASE WHEN cMoneda ='1' THEN nIntDev ELSE 0 END ) AS NINTDEVSOL,"
        psSql = psSql & " SUM(CASE WHEN cMoneda ='2' THEN nIntDev ELSE 0 END ) AS NINTDEVDOL"
        psSql = psSql & " From"
        psSql = psSql & " (SELECT  cCtaCod, nPrdEstado, cRefinan, nSaldoCap, nCapVencido, nIntDev,"
        psSql = psSql & " nDiasAtraso, SUBSTRING(CCTACOD,9,1) AS cMoneda, nMontoDesemb,"
        psSql = psSql & " nDiasParaVenc = CASE    WHEN SUBSTRING(CCTACOD,6,3) ='305'"
        psSql = psSql & "             THEN DATEDIFF(day, '" & Format(psFecha, gsFormatoFecha) & "',dFecVenc)"
        psSql = psSql & "             ELSE (DATEDIFF(day, '" & Format(psFecha, gsFormatoFecha) & "',(Select MAX(dFecVenc) From " & sServidorConsolidada & "PlandespagConsol Where cCtaCod = C.cCtaCod))) END,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='1' AND nDiasAtraso<=15 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap else 0 end as VigCom,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='2' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap else 0 end as VigMes,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap else 0 end as VigCons,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap else 0 end as VigHip,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nDiasAtraso between 31 and 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap-nCapVencido else 0 end as VigCons31,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso between 31 and 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap-nCapVencido else 0 end as VigHip31,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='1' AND nDiasAtraso>15 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nSaldoCap else 0 end as VencCOM,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='2' AND nDiasAtraso>30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nSaldoCap else 0 end as VencMES,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND SUBSTRING(CCTACOD,6,3)<>'305' AND nDiasAtraso>90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nSaldoCap else 0 end as VencCons,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso>90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nSaldoCap else 0 end as VencHIP,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND SUBSTRING(CCTACOD,6,3)<>'305' AND nDiasAtraso BETWEEN 31 AND 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nCapVencido else 0 end as VencCons31,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso BETWEEN 31 AND 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nCapVencido else 0 end as VencHip31,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,3)='305' AND nDiasAtraso>30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nSaldoCap else 0 end as VencPrend,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='1' AND nDiasAtraso<=15 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap else 0 end as REFCom,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='2' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap else 0 end as REFMes,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap else 0 end as REFCons,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap else 0 end as REFHip,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nDiasAtraso between 31 and 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap-nCapVencido else 0 end as REFCons31,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso between 31 and 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap-nCapVencido else 0 end as REFHip31,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='1' AND nPrdEstado in (2201,2205) then nSaldoCap else 0 end as JUDCOM,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='2' AND nPrdEstado in (2201,2205) then nSaldoCap else 0 end as JUDMES,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nPrdEstado in (2201,2205) then nSaldoCap else 0 end as JUDCONS,"
        psSql = psSql & " CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nPrdEstado in (2201,2205) then nSaldoCap else 0 end as JUDHIP"
        psSql = psSql & " FROM    " & sServidorConsolidada & "creditoconsol  C"
        psSql = psSql & " WHERE   nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205)"
        psSql = psSql & " AND cCtaCod NOT IN (SELECT CCTACOD FROM COLOCACCRED WHERE CRFA ='RFA')) AS DAT"
        psSql = psSql & " Where VigCom + VigMes + VigCons + VigHip + VigCons31 + VigHip31 + RefCom + REFMes + REFCons + REFHip + REFCons31 + REFHip31 > 0"
    
        Set rsRes = dbConec.CargaRecordSet(psSql)
        
        If Not (rsRes.EOF And rsRes.BOF) Then
            rsCriterios.Open , , adOpenStatic
            If psMoneda = "1" Then
               rsCriterios.AddNew
               rsCriterios!cCodRango = "01"
               rsCriterios!nValor = rsRes!RANGO1MN + Format(((rsRes!RANGO1MN / rsRes!NMONTOMN) * (rsCtaCont!Saldo - rsRes!NMONTOMN)), "#.00")
               rsCriterios.Update
               
               rsCriterios.AddNew
               rsCriterios!cCodRango = "02"
               rsCriterios!nValor = rsRes!RANGO2MN + Format(((rsRes!RANGO2MN / rsRes!NMONTOMN) * (rsCtaCont!Saldo - rsRes!NMONTOMN)), "#.00")
               rsCriterios.Update
               
               rsCriterios.AddNew
               rsCriterios!cCodRango = "03"
               rsCriterios!nValor = rsRes!RANGO3MN + Format(((rsRes!RANGO3MN / rsRes!NMONTOMN) * (rsCtaCont!Saldo - rsRes!NMONTOMN)), "#.00")
               rsCriterios.Update
               
               rsCriterios.AddNew
               rsCriterios!cCodRango = "04"
               rsCriterios!nValor = rsRes!RANGO4MN + Format(((rsRes!RANGO4MN / rsRes!NMONTOMN) * (rsCtaCont!Saldo - rsRes!NMONTOMN)), "#.00")
               rsCriterios.Update
               
               If psOpeCod = "770162" Then
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "05"
                    rsCriterios!nValor = (rsRes!RANGO5MN + Format(((rsRes!RANGO5MN / rsRes!NMONTOMN) * (rsCtaCont!Saldo - rsRes!NMONTOMN)), "#.00")) + (rsRes!RANGO6MN + Format(((rsRes!RANGO6MN / rsRes!NMONTOMN) * (rsCtaCont!Saldo - rsRes!NMONTOMN)), "#.00"))
                    rsCriterios.Update
               Else
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "05"
                    rsCriterios!nValor = rsRes!RANGO5MN + Format(((rsRes!RANGO5MN / rsRes!NMONTOMN) * (rsCtaCont!Saldo - rsRes!NMONTOMN)), "#.00")
                    rsCriterios.Update
                    
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "06"
                    rsCriterios!nValor = rsRes!RANGO6MN + Format(((rsRes!RANGO6MN / rsRes!NMONTOMN) * (rsCtaCont!Saldo - rsRes!NMONTOMN)), "#.00")
                    rsCriterios.Update
                End If
            Else
               rsCriterios.AddNew
               rsCriterios!cCodRango = "01"
               rsCriterios!nValor = rsRes!RANGO1ME + Format(((rsRes!RANGO1ME / rsRes!NMONTOME) * (rsCtaCont!Saldo - rsRes!NMONTOME)), "#.00")
               rsCriterios.Update
               
               rsCriterios.AddNew
               rsCriterios!cCodRango = "02"
               rsCriterios!nValor = rsRes!RANGO2ME + Format(((rsRes!RANGO2ME / rsRes!NMONTOME) * (rsCtaCont!Saldo - rsRes!NMONTOME)), "#.00")
               rsCriterios.Update
               
               rsCriterios.AddNew
               rsCriterios!cCodRango = "03"
               rsCriterios!nValor = rsRes!RANGO3ME + Format(((rsRes!RANGO3ME / rsRes!NMONTOME) * (rsCtaCont!Saldo - rsRes!NMONTOME)), "#.00")
               rsCriterios.Update
               
               rsCriterios.AddNew
               rsCriterios!cCodRango = "04"
               rsCriterios!nValor = rsRes!RANGO4ME + Format(((rsRes!RANGO4ME / rsRes!NMONTOME) * (rsCtaCont!Saldo - rsRes!NMONTOME)), "#.00")
               rsCriterios.Update
               
               If psOpeCod = "770162" Then
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "05"
                    rsCriterios!nValor = (rsRes!RANGO5ME + Format(((rsRes!RANGO5ME / rsRes!NMONTOME) * (rsCtaCont!Saldo - rsRes!NMONTOME)), "#.00")) + (rsRes!RANGO6ME + Format(((rsRes!RANGO6ME / rsRes!NMONTOME) * (rsCtaCont!Saldo - rsRes!NMONTOME)), "#.00"))
                    rsCriterios.Update
               Else
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "05"
                    rsCriterios!nValor = rsRes!RANGO5ME + Format(((rsRes!RANGO5ME / rsRes!NMONTOME) * (rsCtaCont!Saldo - rsRes!NMONTOME)), "#.00")
                    rsCriterios.Update
                    
                    rsCriterios.AddNew
                    rsCriterios!cCodRango = "06"
                    rsCriterios!nValor = rsRes!RANGO6ME + Format(((rsRes!RANGO6ME / rsRes!NMONTOME) * (rsCtaCont!Saldo - rsRes!NMONTOME)), "#.00")
                    rsCriterios.Update
               End If
            End If
            rsCriterios.MoveFirst
        End If
    Else
        
        Dim lsFiltro As String
        Dim lsSql As String
        Dim rsF As ADODB.Recordset
        Set rsF = New ADODB.Recordset
        Dim lnAcum As Currency
        
        If psCtaCod = "21_3" Then
            lsFiltro = " IN (1,2,3)"
            If psMoneda = 1 Then
                lsSql = "Select (dbo.GetSaldoCtaAcumulado('" & psFecha & "','21" & psMoneda & "3%',1) - dbo.GetSaldoCtaAcumulado('" & psFecha & "','21" & psMoneda & "305%',1))"
            Else
                lsSql = "Select (dbo.GetSaldoCtaAcumulado('" & psFecha & "','21" & psMoneda & "3%',1) - dbo.GetSaldoCtaAcumulado('" & psFecha & "','21" & psMoneda & "305%',1)  )  / (Select Top 1 nValFijo from tipocambio where dFecCamb = DateAdd(Day,1,'" & psFecha & "') order by dFecCamb Desc) "
            End If
        Else
            lsFiltro = " NOT IN (1,2,3)"
            If psMoneda = 1 Then
                lsSql = "Select dbo.GetSaldoCtaAcumulado('" & psFecha & "','23" & psMoneda & "3%',1)"
            Else
                lsSql = "Select dbo.GetSaldoCtaAcumulado('" & psFecha & "','23" & psMoneda & "3%',1) / (Select Top 1 nValFijo from tipocambio where dFecCamb = DateAdd(Day,1,'" & psFecha & "') order by dFecCamb Desc) "
            End If
        End If
        
        Set rsRes = dbConec.CargaRecordSet(lsSql)
        lnAcum = rsRes.Fields(0)
        
        psSql = psSql & " SELECT  RANGO1MN = SUM(CASE WHEN NPLAZO <=7 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END ) ,"
        psSql = psSql & " RANGO1ME = SUM(CASE WHEN NPLAZO <=7 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END ) ,"
        psSql = psSql & " RANGO2MN = SUM((CASE WHEN NPLAZO BETWEEN 8 AND 15 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO2ME = SUM((CASE WHEN NPLAZO BETWEEN 8 AND 15 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO3MN = SUM((CASE WHEN NPLAZO BETWEEN 16 AND 30 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO3ME = SUM((CASE WHEN NPLAZO BETWEEN 16 AND 30 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO4MN = SUM((CASE WHEN NPLAZO BETWEEN 31 AND 60 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO4ME = SUM((CASE WHEN NPLAZO BETWEEN 31 AND 60 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO5MN = SUM((CASE WHEN NPLAZO BETWEEN 61 AND 90 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO5ME = SUM((CASE WHEN NPLAZO BETWEEN 61 AND 90 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO6MN = SUM((CASE WHEN NPLAZO BETWEEN 91 AND 180 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO6ME = SUM((CASE WHEN NPLAZO BETWEEN 91 AND 180 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO7MN = SUM((CASE WHEN NPLAZO BETWEEN 181 AND 360 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO7ME = SUM((CASE WHEN NPLAZO BETWEEN 181 AND 360 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO8MN = SUM((CASE WHEN NPLAZO BETWEEN 361 AND 720 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO8ME = SUM((CASE WHEN NPLAZO BETWEEN 361 AND 720 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO9MN = SUM((CASE WHEN NPLAZO BETWEEN 721 AND 1080 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO9ME = SUM((CASE WHEN NPLAZO BETWEEN 721 AND 1080 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO10MN = SUM((CASE WHEN NPLAZO BETWEEN 1081 AND 1440 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO10ME = SUM((CASE WHEN NPLAZO BETWEEN 1081 AND 1440 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO11MN = SUM((CASE WHEN NPLAZO BETWEEN 1441 AND 1800 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO11ME = SUM((CASE WHEN NPLAZO BETWEEN 1441 AND 1800 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO12MN = SUM((CASE WHEN NPLAZO BETWEEN 1801 AND 3600 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO12ME = SUM((CASE WHEN NPLAZO BETWEEN 1801 AND 3600 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO13MN = SUM((CASE WHEN NPLAZO BETWEEN 3601 AND 7200 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO13ME = SUM((CASE WHEN NPLAZO BETWEEN 3601 AND 7200 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO14MN = SUM((CASE WHEN NPLAZO BETWEEN 7201 AND 12000 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO14ME = SUM((CASE WHEN NPLAZO BETWEEN 7201 AND 12000 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO15MN = SUM((CASE WHEN NPLAZO>12001 AND SUBSTRING(CCTACOD,9,1)='1' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " RANGO15ME = SUM((CASE WHEN NPLAZO>12001 AND SUBSTRING(CCTACOD,9,1)='2' THEN NSALDO ELSE 0 END )),"
        psSql = psSql & " SUM(CASE WHEN SUBSTRING(CCTACOD,9,1) ='1' THEN NSALDO ELSE 0 END) AS NSALDOMN,"
        psSql = psSql & " SUM(CASE WHEN SUBSTRING(CCTACOD,9,1) ='2' THEN NSALDO ELSE 0 END) AS NSALDOME"
        psSql = psSql & " From"
        psSql = psSql & " (SELECT SUBSTRING(A.CCTACOD,6,3) AS CPROD, A.CCTACOD, Case When aBS(DateDiff(Day,'" & Format(CDate(psFecha), gsFormatoFecha) & "',DateAdd(Day,A.NPLAZO,dRenovacion))) >= A.NPLAZO Then A.NPLAZO Else aBS(DateDiff(Day,'" & Format(CDate(psFecha), gsFormatoFecha) & "',DateAdd(Day,A.NPLAZO,dRenovacion))) End NPLAZO, nSaldCNTPF AS NSALDO,"
        psSql = psSql & " TIPOBLOQ=ISNULL((SELECT  TOP 1 NBLQMOTIVO"
        psSql = psSql & "         FROM    " & sServidorConsolidada & "PRODUCTOBLOQUEOSCONSOL PB"
        psSql = psSql & "          Where PB.cCtaCod = A.cCtaCod"
        psSql = psSql & "          AND CMOVNRODBL IS  NULL AND CONVERT(CHAR(8),DCIERRE,112) = '" & Format(CDate(psFecha), gsFormatoMovFecha) & "'"
        psSql = psSql & "          ),0)"
        psSql = psSql & " FROM    " & sServidorConsolidada & "PLAZOFIJOCONSOL A Inner Join captacplazofijo C On A.cCtaCod = C.cCtaCod  "
        psSql = psSql & " WHERE   nPersoneria " & lsFiltro
        psSql = psSql & " Union All"
        psSql = psSql & " SELECT  Top 0  SUBSTRING(CCTACOD,6,3) AS CPROD , CCTACOD, 720 AS NPLAZO, NSALDCNTCTS AS NSALDO,"
        psSql = psSql & "     TIPOBLOQ = 0 "
        psSql = psSql & " FROM  " & sServidorConsolidada & "CTSCONSOL A"
        psSql = psSql & " WHERE  nPersoneria " & lsFiltro
        psSql = psSql & " ) AS X"
        psSql = psSql & " WHERE TIPOBLOQ NOT IN (3)"
    
        Set rsRes = dbConec.CargaRecordSet(psSql)
        
        If Not (rsRes.EOF And rsRes.BOF) Then
            rsCriterios.Open , , adOpenStatic
            If psMoneda = "1" Then
               lnAcum = lnAcum - rsRes!RANGO1MN: lnAcum = lnAcum - rsRes!RANGO2MN
               lnAcum = lnAcum - rsRes!RANGO3MN: lnAcum = lnAcum - rsRes!RANGO4MN
               lnAcum = lnAcum - rsRes!RANGO5MN: lnAcum = lnAcum - rsRes!RANGO6MN
               lnAcum = lnAcum - rsRes!RANGO7MN: lnAcum = lnAcum - rsRes!RANGO8MN
               lnAcum = lnAcum - rsRes!RANGO9MN: lnAcum = lnAcum - rsRes!RANGO10MN
               lnAcum = lnAcum - rsRes!RANGO11MN: lnAcum = lnAcum - rsRes!RANGO12MN
               lnAcum = lnAcum - rsRes!RANGO13MN: lnAcum = lnAcum - rsRes!RANGO14MN
               
               lnAcumulaT = rsRes!RANGO1MN + rsRes!RANGO2MN + rsRes!RANGO3MN + rsRes!RANGO4MN + rsRes!RANGO5MN + rsRes!RANGO6MN + rsRes!RANGO7MN + rsRes!RANGO8MN + rsRes!RANGO9MN + rsRes!RANGO10MN + rsRes!RANGO11MN + rsRes!RANGO12MN + rsRes!RANGO13MN + rsRes!RANGO14MN
               
               rsCriterios.AddNew: rsCriterios!cCodRango = "01": rsCriterios!nValor = rsRes!RANGO1MN + Format(rsRes!RANGO1MN * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
               rsCriterios.AddNew: rsCriterios!cCodRango = "02": rsCriterios!nValor = rsRes!RANGO2MN + Format(rsRes!RANGO2MN * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
               rsCriterios.AddNew: rsCriterios!cCodRango = "03": rsCriterios!nValor = rsRes!RANGO3MN + Format(rsRes!RANGO3MN * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
               rsCriterios.AddNew: rsCriterios!cCodRango = "04": rsCriterios!nValor = rsRes!RANGO4MN + Format(rsRes!RANGO4MN * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
               rsCriterios.AddNew: rsCriterios!cCodRango = "05": rsCriterios!nValor = rsRes!RANGO5MN + Format(rsRes!RANGO5MN * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
               rsCriterios.AddNew: rsCriterios!cCodRango = "06": rsCriterios!nValor = rsRes!RANGO6MN + Format(rsRes!RANGO6MN * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
               rsCriterios.AddNew: rsCriterios!cCodRango = "07": rsCriterios!nValor = rsRes!RANGO7MN + Format(rsRes!RANGO7MN * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
               rsCriterios.AddNew: rsCriterios!cCodRango = "08": rsCriterios!nValor = rsRes!RANGO8MN + Format(rsRes!RANGO8MN * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
               rsCriterios.AddNew: rsCriterios!cCodRango = "09": rsCriterios!nValor = rsRes!RANGO9MN + Format(rsRes!RANGO9MN * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
               rsCriterios.AddNew: rsCriterios!cCodRango = "10": rsCriterios!nValor = rsRes!RANGO10MN + Format(rsRes!RANGO10MN * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
               rsCriterios.AddNew: rsCriterios!cCodRango = "11": rsCriterios!nValor = rsRes!RANGO11MN + Format(rsRes!RANGO11MN * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
               rsCriterios.AddNew: rsCriterios!cCodRango = "12": rsCriterios!nValor = rsRes!RANGO12MN + Format(rsRes!RANGO12MN * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
               rsCriterios.AddNew: rsCriterios!cCodRango = "13": rsCriterios!nValor = rsRes!RANGO13MN + Format(rsRes!RANGO13MN * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
               rsCriterios.AddNew: rsCriterios!cCodRango = "14": rsCriterios!nValor = rsRes!RANGO14MN + Format(rsRes!RANGO14MN * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
            Else
               lnAcum = lnAcum - rsRes!RANGO1ME: lnAcum = lnAcum - rsRes!RANGO2ME
               lnAcum = lnAcum - rsRes!RANGO3ME: lnAcum = lnAcum - rsRes!RANGO4ME
               lnAcum = lnAcum - rsRes!RANGO5ME: lnAcum = lnAcum - rsRes!RANGO6ME
               lnAcum = lnAcum - rsRes!RANGO7ME: lnAcum = lnAcum - rsRes!RANGO8ME
               lnAcum = lnAcum - rsRes!RANGO9ME: lnAcum = lnAcum - rsRes!RANGO10ME
               lnAcum = lnAcum - rsRes!RANGO11ME: lnAcum = lnAcum - rsRes!RANGO12ME
               lnAcum = lnAcum - rsRes!RANGO13ME: lnAcum = lnAcum - rsRes!RANGO14ME
               
               lnAcumulaT = rsRes!RANGO1ME + rsRes!RANGO2ME + rsRes!RANGO3ME + rsRes!RANGO4ME + rsRes!RANGO5ME + rsRes!RANGO6ME + rsRes!RANGO7ME + rsRes!RANGO8ME + rsRes!RANGO9ME + rsRes!RANGO10ME + rsRes!RANGO11ME + rsRes!RANGO12ME + rsRes!RANGO13ME + rsRes!RANGO14ME
               
               If lnAcumulaT <> 0 Then
                    rsCriterios.AddNew: rsCriterios!cCodRango = "01": rsCriterios!nValor = rsRes!RANGO1ME + Format(rsRes!RANGO1ME * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
                    rsCriterios.AddNew: rsCriterios!cCodRango = "02": rsCriterios!nValor = rsRes!RANGO2ME + Format(rsRes!RANGO2ME * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
                    rsCriterios.AddNew: rsCriterios!cCodRango = "03": rsCriterios!nValor = rsRes!RANGO3ME + Format(rsRes!RANGO3ME * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
                    rsCriterios.AddNew: rsCriterios!cCodRango = "04": rsCriterios!nValor = rsRes!RANGO4ME + Format(rsRes!RANGO4ME * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
                    rsCriterios.AddNew: rsCriterios!cCodRango = "05": rsCriterios!nValor = rsRes!RANGO5ME + Format(rsRes!RANGO5ME * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
                    rsCriterios.AddNew: rsCriterios!cCodRango = "06": rsCriterios!nValor = rsRes!RANGO6ME + Format(rsRes!RANGO6ME * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
                    rsCriterios.AddNew: rsCriterios!cCodRango = "07": rsCriterios!nValor = rsRes!RANGO7ME + Format(rsRes!RANGO7ME * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
                    rsCriterios.AddNew: rsCriterios!cCodRango = "08": rsCriterios!nValor = rsRes!RANGO8ME + Format(rsRes!RANGO8ME * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
                    rsCriterios.AddNew: rsCriterios!cCodRango = "09": rsCriterios!nValor = rsRes!RANGO9ME + Format(rsRes!RANGO9ME * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
                    rsCriterios.AddNew: rsCriterios!cCodRango = "10": rsCriterios!nValor = rsRes!RANGO10ME + Format(rsRes!RANGO10ME * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
                    rsCriterios.AddNew: rsCriterios!cCodRango = "11": rsCriterios!nValor = rsRes!RANGO11ME + Format(rsRes!RANGO11ME * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
                    rsCriterios.AddNew: rsCriterios!cCodRango = "12": rsCriterios!nValor = rsRes!RANGO12ME + Format(rsRes!RANGO12ME * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
                    rsCriterios.AddNew: rsCriterios!cCodRango = "13": rsCriterios!nValor = rsRes!RANGO13ME + Format(rsRes!RANGO13ME * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
                    rsCriterios.AddNew: rsCriterios!cCodRango = "14": rsCriterios!nValor = rsRes!RANGO14ME + Format(rsRes!RANGO14ME * lnAcum / lnAcumulaT, "0.00"): rsCriterios.Update
                    rsCriterios.MoveFirst
               End If
            End If
            
        End If
    End If

Set CargaCriterioEstadistico07 = rsCriterios
    Exit Function
CargaCriterioTasaErr:
    Call RaiseError(MyUnhandledError, "DAnexoRiesgos:CargaCriterioTasa Method")
End Function
'JACA 20110516******************************************************************
Public Function obtenerCreditosVigentesAnx7(ByVal psFecha As Date) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo obtenerCreditosVigentesAnx7Err
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_CreditosVigentesAnx7 '" & Format(psFecha, "yyyymmdd") & "'"

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set obtenerCreditosVigentesAnx7 = rs
    Exit Function
obtenerCreditosVigentesAnx7Err:
    Call RaiseError(MyUnhandledError, "Obtener Creditos VIgentes Anexo 07")
End Function
Public Sub insertarCreditosVigentesAnx7(ByVal psFecha As Date, ByVal prs As Recordset, ByVal sMovNro As String)
Dim sSQL As String
Dim oCon As New DConecta
 Dim I As Integer
On Error GoTo insertarCreditosVigentesAnx7Err
    oCon.AbreConexion
   

    If Not prs.EOF Or Not prs.BOF Then
        For I = 0 To prs.RecordCount - 1
            sSQL = "exec stp_ins_CreditosVigentesAnx7 " & prs(0) & "," & prs(1) & "," & prs(2) & "," & prs(3) & "," & prs(4) & "," & prs(5) & "," & prs(6) & "," & prs(7) & "," & prs(8) & "," & prs(9) & "," & prs(10) & "," & prs(11) & "," & prs(12) & "," & prs(13) & ",'" & Format(psFecha, "yyyymmdd") & "'," & prs(prs.Fields.Count - 1) & ",'" & sMovNro & "'"
            oCon.CargaRecordSet (sSQL)
            prs.MoveNext
        Next I
    End If
    

    'oCon.AbreConexion
    'oCon.CargaRecordSet (sSQL)
    oCon.CierraConexion
    Exit Sub
insertarCreditosVigentesAnx7Err:
    Call RaiseError(MyUnhandledError, "Registrar Creditos VIgentes Anexo 07")
End Sub
Public Sub actualizarCreditosVigentesAnx7(ByVal psFecha As Date)
Dim sSQL As String
Dim oCon As New DConecta

On Error GoTo actualizarCreditosVigentesAnx7Err
    oCon.AbreConexion
     
    sSQL = "exec stp_upd_CreditosVigentesAnx7 '" & Format(psFecha, "yyyymmdd") & "'"
   
    oCon.AbreConexion
    oCon.CargaRecordSet (sSQL)
    oCon.CierraConexion
    Exit Sub
actualizarCreditosVigentesAnx7Err:
    Call RaiseError(MyUnhandledError, "Actualizar Creditos VIgentes Anexo 07")
End Sub

Public Function obtenerCreditosRefinanciadosAnx7(ByVal psFecha As Date) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo obtenerCreditosRefinanciadosAnx7Err
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_CreditosRefinanciadosAnx7 '" & Format(psFecha, "yyyymmdd") & "'"

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set obtenerCreditosRefinanciadosAnx7 = rs
    Exit Function
obtenerCreditosRefinanciadosAnx7Err:
    Call RaiseError(MyUnhandledError, "Obtener Creditos Refinanciados Anexo 07")
End Function
Public Sub insertarCreditosRefinanciadosAnx7(ByVal psFecha As Date, ByVal prs As Recordset, ByVal sMovNro As String)
Dim sSQL As String
Dim oCon As New DConecta
 Dim I As Integer
On Error GoTo insertarCreditosRefinanciadosAnx7Err
    oCon.AbreConexion
   

    If Not prs.EOF Or Not prs.BOF Then
        For I = 0 To prs.RecordCount - 1
            sSQL = "exec stp_ins_CreditosRefinanciadosAnx7 " & prs(0) & "," & prs(1) & "," & prs(2) & "," & prs(3) & "," & prs(4) & "," & prs(5) & "," & prs(6) & "," & prs(7) & "," & prs(8) & "," & prs(9) & "," & prs(10) & "," & prs(11) & "," & prs(12) & "," & prs(13) & ",'" & Format(psFecha, "yyyymmdd") & "'," & prs(prs.Fields.Count - 1) & ",'" & sMovNro & "'"
            oCon.CargaRecordSet (sSQL)
            prs.MoveNext
        Next I
    End If
    

    'oCon.AbreConexion
    'oCon.CargaRecordSet (sSQL)
    oCon.CierraConexion
    Exit Sub
insertarCreditosRefinanciadosAnx7Err:
    Call RaiseError(MyUnhandledError, "Registrar Creditos Refinanciados Anexo 07")
End Sub
Public Sub actualizarCreditosRefinanciadosAnx7(ByVal psFecha As Date)
Dim sSQL As String
Dim oCon As New DConecta

On Error GoTo actualizarCreditosRefinanciadosAnx7Err
    oCon.AbreConexion
     
    sSQL = "exec stp_upd_CreditosRefinanciadosAnx7 '" & Format(psFecha, "yyyymmdd") & "'"
   
    oCon.AbreConexion
    oCon.CargaRecordSet (sSQL)
    oCon.CierraConexion
    Exit Sub
actualizarCreditosRefinanciadosAnx7Err:
    Call RaiseError(MyUnhandledError, "Actualizar Creditos Refinanciados Anexo 07")
End Sub
Public Function obtenerAdeudadosAnx7(ByVal psFecha As Date, ByVal pnMoneda As Integer) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo obtenerAdeudadosAnx7Err
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_AdeudadosAnx7 '" & Format(psFecha, "yyyymmdd") & "'," & pnMoneda & ""

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set obtenerAdeudadosAnx7 = rs
    Exit Function
obtenerAdeudadosAnx7Err:
    Call RaiseError(MyUnhandledError, "Obtener Adeudados Anexo 07")
End Function
Public Sub insertarAdeudadosAnx7(ByVal psFecha As Date, ByVal prs As Recordset, ByVal sMovNro As String, ByVal pnMoneda As Integer)
Dim sSQL As String
Dim oCon As New DConecta
 Dim I As Integer
On Error GoTo insertarAdeudadosAnx7Err
    oCon.AbreConexion
   

    If Not prs.EOF Or Not prs.BOF Then
        For I = 0 To prs.RecordCount - 1
            sSQL = "exec stp_ins_AdeudadosAnx7 " & prs(0) & "," & prs(1) & "," & prs(2) & "," & prs(3) & "," & prs(4) & "," & prs(5) & "," & prs(6) & "," & prs(7) & "," & prs(8) & "," & prs(9) & "," & prs(10) & "," & prs(11) & "," & prs(12) & "," & prs(13) & ",'" & Format(psFecha, "yyyymmdd") & "'," & pnMoneda & ",'" & sMovNro & "'"
            oCon.CargaRecordSet (sSQL)
            prs.MoveNext
        Next I
    End If
    

    'oCon.AbreConexion
    'oCon.CargaRecordSet (sSQL)
    oCon.CierraConexion
    Exit Sub
insertarAdeudadosAnx7Err:
    Call RaiseError(MyUnhandledError, "Registrar Adeudados Anexo 07")
End Sub
Public Sub actualizarAdeudadosAnx7(ByVal psFecha As Date)
Dim sSQL As String
Dim oCon As New DConecta

On Error GoTo actualizarAdeudadosAnx7Err
    oCon.AbreConexion
     
    sSQL = "exec stp_upd_AdeudadosAnx7 '" & Format(psFecha, "yyyymmdd") & "'"
   
    oCon.AbreConexion
    oCon.CargaRecordSet (sSQL)
    oCon.CierraConexion
    Exit Sub
actualizarAdeudadosAnx7Err:
    Call RaiseError(MyUnhandledError, "Actualizar Adeudados Anexo 07")
End Sub
'JACA END***********************************************************************


