VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSeguros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsServerAdmin As String
Dim vsServerComunes As String
Dim vsServerImagenes As Variant
Dim vsServerNegocio As String
Dim vsServerPersonas As String
Dim oConect As DConecta
Dim sSql As String
Dim lbBitControlCuentasUltNivel As Boolean
Dim rsCta As ADODB.Recordset
Public Sub RollbackTrans()
    oConect.RollbackTrans
End Sub

Public Sub CommitTrans()
    oConect.CommitTrans
End Sub
Public Sub BeginTrans()
    oConect.BeginTrans
End Sub
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    Dim oCon As NConstSistemas
    Set oCon = New NConstSistemas
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)

    Set oImp = Nothing
    Set rsCta = New ADODB.Recordset
    
    Dim oIni As ClasIni
    
    Set oIni = New ClasIni
    
    lbBitControlCuentasUltNivel = IIf(oCon.LeeConstSistema(gConstSistBitControCtasContUltNiv) = "1", True, False)
    
    vsServerAdmin = oIni.BaseAdministracion
    vsServerComunes = oIni.BaseComunes
    vsServerImagenes = oIni.BaseImagenes
    vsServerNegocio = oIni.BaseNegocio
    vsServerPersonas = oIni.BasePersonas
    
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then
        Call RaiseError(MyUnhandledError, "DMov:Initialize Method. Error en Conexion a Base de datos")
    End If
    Set oIni = Nothing
End Sub
Private Sub Class_Terminate()
oConect.CierraConexion
Set oConect = Nothing
End Sub
Public Function ObtenerSegTarjetaSolicitudesPendientes(ByVal pnTipo As Integer, Optional ByVal psNumSolicitud As String = "") As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConecta As DConecta
    
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_SegTarjetaSolicitudesPendientes " & pnTipo & ",'" & psNumSolicitud & "'"
    Set oConecta = New DConecta
    
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    Set ObtenerSegTarjetaSolicitudesPendientes = rs
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function RecuperarCtaAhorro(ByVal psNumSolicitud As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConecta As DConecta
    
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_RecuperarCtasTitularSegTarjeta '" & psNumSolicitud & "'"
    Set oConecta = New DConecta
    
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    Set RecuperarCtaAhorro = rs
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Sub InsertaSegTarjetaEstado(ByVal psNumSolicitud As String, ByVal pdPrdEstado As Date, ByVal pnPrdEstado As Integer, ByVal pcDescripcion As String, ByVal pnMonto As Currency, ByVal psUltimaActualizacion As String)
On Error GoTo InsertaErr
    Dim lsSql As String

    lsSql = "exec stp_ins_SegTarjetaEstado '" & psNumSolicitud & "','" & Format(pdPrdEstado, "yyyymmdd") & "'," & pnPrdEstado & ",'" & pcDescripcion & "'," & pnMonto & ",'" & psUltimaActualizacion & "'"
    oConect.Ejecutar lsSql
    
        Exit Sub
InsertaErr:
    Call RaiseError(MyUnhandledError, "DSeguros:InsertaSegTarjeta Method")
End Sub
Public Sub ActualizarSegTarjetaSolicitudAproAcep(ByVal psNumSolicitud As String, ByVal psMovNro As String, ByVal pnEstado As Integer, ByVal psCtaDeposito As String, ByVal pnMonto As Currency)
On Error GoTo ActualizarErr
    Dim lsSql As String

    lsSql = "exec stp_upd_SegTarjetaSolicitudAproAcep '" & psNumSolicitud & "','" & psMovNro & "'," & pnEstado & ",'" & psCtaDeposito & "'," & pnMonto
    oConect.Ejecutar lsSql
    
    Exit Sub
ActualizarErr:
    Call RaiseError(MyUnhandledError, "DSeguros:ActualizarSegTarjeta Method")
End Sub
Public Function RecuperarDepositosPendientes(ByVal psPersCod As String, ByVal psIFTpo As String, psCtaIFCod As String, ByVal pnImporte As Currency, ByVal pdFechaDeposito As Date) As ADODB.Recordset
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConecta As DConecta
    
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_RecuperarDepositosPendientes '" & psPersCod & "','" & psIFTpo & "','" & psCtaIFCod & "'," & pnImporte & ",'" & Format(pdFechaDeposito, "yyyymmdd") & "'"
    Set oConecta = New DConecta
    
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    Set RecuperarDepositosPendientes = rs
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'**********************************************
'JUEZ 20140711 *************************************************************
Public Function RecuperaGeneracionTramaSegTarjeta(ByVal pnMes As Integer, ByVal pnAño As Integer, ByVal psTipoTrama As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    sSql = "exec stp_sel_RecuperaGeneracionTramaSegTarjeta " & pnMes & "," & pnAño & ",'" & psTipoTrama & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGeneracionTramaSegTarjeta = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function ObtenerSegTarjetaParametros(ByVal pnParamVar As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    sSql = "exec stp_sel_RecuperaSegTarjetaParametro " & pnParamVar
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set ObtenerSegTarjetaParametros = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function RecuperaSegTarjetaOpeBancos(ByVal psAnio As String, ByVal psMes As String, ByVal pnTipoOpe As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    sSql = "exec stp_sel_SegTarjetaOpeBancos '" & psAnio & "','" & psMes & "'," & pnTipoOpe
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaSegTarjetaOpeBancos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function RecuperaSegTarjetaOpeBancosExtorno(ByVal psFechaMov As String, ByVal pnTipoOpe As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    sSql = "exec stp_sel_SegTarjetaOpeBancosExtorno '" & psFechaMov & "'," & pnTipoOpe
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaSegTarjetaOpeBancosExtorno = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'END JUEZ ******************************************************************
'APRI20180130 ERS028-2017
Public Function RecuperaGeneracionTramaSeguros(ByVal pnMes As Integer, ByVal pnAño As Integer, ByVal pnTipoTrama As Integer, Optional ByVal pnMoneda As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    sSql = "exec stp_sel_RecuperaGeneracionTramaSeguros " & pnMes & "," & pnAño & "," & pnTipoTrama & "," & pnMoneda & ""
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGeneracionTramaSeguros = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function RecuperaSegurosOpeBancos(ByVal psAnio As String, ByVal psMes As String, ByVal pnMoneda As Integer, ByVal pnTipoSeg As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    sSql = "exec stp_sel_SegurosOpeBancos '" & psAnio & "','" & psMes & "'," & pnMoneda & "," & pnTipoSeg
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaSegurosOpeBancos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function RecuperaSegurosOpeBancosExtorno(ByVal psFechaMov As String, ByVal pnMoneda As Integer, ByVal pnTipoSeg As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    sSql = "exec stp_sel_SeguroOpeBancosExtorno '" & psFechaMov & "'," & pnMoneda & "," & pnTipoSeg
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaSegurosOpeBancosExtorno = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'END APRI
'PASIERS1362014
Public Function RecuperaRetPagoSeguro(ByVal psAnio As String, ByVal psMes As String, ByVal psOpecod As String, ByVal pnMoneda As Integer) As ADODB.Recordset
    Dim oConecta As DConecta
    sSql = "stp_sel_ERS1362014_MovRetPagoSeguro '" & psAnio & "','" & psMes & "','" & psOpecod & "'," & pnMoneda
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaRetPagoSeguro = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function GetMovRetPagoSeguroxExtorno(ByVal pdFecIni As Date, ByVal pdFecFin As Date, ByVal psOpecod As String) As ADODB.Recordset
    Dim oConecta As DConecta
    Set oConecta = New DConecta
    oConecta.AbreConexion
        sSql = "stp_sel_ERS1362014_ObtieneMovRetPagoSeguroxExtorno '" & Format(pdFecIni, "yyyyMMdd") & "','" & Format(pdFecFin, "yyyyMMdd") & "','" & psOpecod & "'"
        Set GetMovRetPagoSeguroxExtorno = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
End Function
Public Function GetCuotaCredxRetPagSeguro(ByVal psCtaCod As String, ByVal pdFecha As Date) As ADODB.Recordset
    Dim oConecta As DConecta
    Set oConecta = New DConecta
    oConecta.AbreConexion
        sSql = "stp_sel_ERS1362014_ObtieneCuotaCredxRetPagSeguro '" & psCtaCod & "','" & Format(pdFecha, "yyyyMMdd") & "'"
        Set GetCuotaCredxRetPagSeguro = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
End Function
Public Function GetMontoSegDesgravamenxRetPagSeguro(ByVal psCtaCod As String, ByVal pnSaldoCap As Currency, ByVal pdFecVig As Date) As Currency
    Dim oConecta As DConecta
    Dim rs As ADODB.Recordset
    Set oConecta = New DConecta
    
    oConecta.AbreConexion
    sSql = "stp_sel_ERS1362014_ObtieneMontoSegDesgravamenxRetPagSeguro '" & psCtaCod & "'," & pnSaldoCap & ",'" & Format(pdFecVig, "yyyyMMdd") & "'"
    Set rs = oConecta.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
         GetMontoSegDesgravamenxRetPagSeguro = rs!nMonto
    Else
        GetMontoSegDesgravamenxRetPagSeguro = 0
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
End Function
'end PASI***************************************
'JUEZ 20150510 ******************************************************************
Public Function RecuperaSegTarjetaAnulaDevPend(ByVal pbRegistrado As Boolean, ByVal pbPagoAseg As Boolean)
Dim sSql As String
Dim oConecta As DConecta
    
    sSql = "exec stp_sel_SegTarjetaAnulaDevPendiente " & IIf(pbRegistrado, 1, 0) & "," & IIf(pbPagoAseg, 1, 0)
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaSegTarjetaAnulaDevPend = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Sub ActualizaSegTarjetaAnulaDevPendiente(ByVal psNumCert As String, ByVal psPersCod As String, ByVal pbRegistrado As Boolean, ByVal pbPagoAseg As Boolean, Optional ByVal pnMovNroPagoAseg As Long = 0)
    Dim lsSql As String

    lsSql = "exec stp_upd_SegTarjetaAnulaDevPendiente '" & psNumCert & "','" & psPersCod & "'," & IIf(pbRegistrado, 1, 0) & "," & IIf(pbPagoAseg, 1, 0) & "," & pnMovNroPagoAseg
    oConect.Ejecutar lsSql
End Sub
Public Sub ExtornaSegTarjetaAnulaDevPendiente(ByVal pnMovNroPagoAseg As Long)
    Dim lsSql As String

    lsSql = "exec stp_upd_ExtornaSegTarjetaAnulaDevPendiente " & pnMovNroPagoAseg
    oConect.Ejecutar lsSql
End Sub
'END JUEZ ***********************************************************************
