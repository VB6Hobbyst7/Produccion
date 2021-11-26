VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLineaCreditoV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function ObtenerLineaCreditoAdeudado() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_LineaCreditoAdeudado "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerLineaCreditoAdeudado = rs
End Function

Public Function ObtenerLineaCreditoAdeudadoPagare(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_LineaCreditoAdeudadoPagare '" & psPersCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerLineaCreditoAdeudadoPagare = rs
End Function

Public Function ObtenerLineaCreditoRRPP() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_LineaCreditoRRPP"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerLineaCreditoRRPP = rs
End Function

Public Function ObtenerLineaCreditoTipoCredito(ByVal psLineaCreditoCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_LineaCreditoTipoCredito '" & psLineaCreditoCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerLineaCreditoTipoCredito = rs
End Function


Public Sub InsertaLineaCreditoTipoCredito(ByVal psLineaCreditoCod As String, ByVal psTpoCredCod As String, ByVal pnLogico As Integer)
    On Error GoTo InsertaLineaCreditoTipoCreditoErr
    Dim sql As String
    Dim oCon As New DConecta
    'InsertaLineaCreditoTipoCredito = 1
    oCon.AbreConexion
    sql = "exec stp_ins_LineaCreditoTipoCredito '" & psLineaCreditoCod & "','" & psTpoCredCod & "'," & pnLogico
    oCon.Ejecutar sql
    oCon.CierraConexion
    'InsertaLineaCreditoTipoCredito = 0
    Exit Sub
InsertaLineaCreditoTipoCreditoErr:
    Call RaiseError(MyUnhandledError, "DLineaCreditoV2:InsertaLineaCreditoTipoCredito Method")
End Sub
Public Sub InsertaLineaCreditoAgencia(ByVal psLineaCreditoCod As String, ByVal psAgeCod As String, ByVal pnLogico As Integer)
    On Error GoTo InsertaLineaCreditoAgenciaErr
    Dim sql As String
    Dim oCon As New DConecta
    'InsertaLineaCreditoAgencia = 1
    oCon.AbreConexion
    sql = "exec stp_ins_LineaCreditoAgencia '" & psLineaCreditoCod & "','" & psAgeCod & "'," & pnLogico
    oCon.Ejecutar sql
    oCon.CierraConexion
    'InsertaLineaCreditoAgencia = 0
    Exit Sub
InsertaLineaCreditoAgenciaErr:
    Call RaiseError(MyUnhandledError, "DLineaCreditoV2:InsertaLineaCreditoAgencia Method")
End Sub

Public Sub InsertarLineaCredito(ByVal psLineaCreditoCod As String, ByVal psLineaCreditoDes As String, ByVal psPersCod As String, ByVal psCodAdeudado As String, ByVal pnCodRRPP As Integer, ByVal pnRanMonDesde As Currency, ByVal pnRanMonHasta As Currency, ByVal pnRanTasDesde As Currency, ByVal pnRanTasHasta As Currency, ByVal pbCalA As Integer, ByVal pbCalB As Integer, ByVal pbCalC As Integer, ByVal pbCalD As Integer, ByVal pbCalE As Integer, ByVal pdFechaMax As Date, ByVal nEstado As Integer, ByVal psPersCodMV As String)
    On Error GoTo InsertarLineaCreditoErr
    Dim sql As String
    Dim oCon As New DConecta
    'InsertaMovCont = 1
    oCon.AbreConexion
    sql = "exec stp_ins_LineaCredito '" & psLineaCreditoCod & "' , '" & psLineaCreditoDes & "','" & psPersCod & "', '" & psCodAdeudado & "'," & pnCodRRPP & ", " & pnRanMonDesde & " , " & pnRanMonHasta & ", " & pnRanTasDesde & ", " & pnRanTasHasta & ", " & pbCalA & ", " & pbCalB & ", " & pbCalC & ", " & pbCalD & ", " & pbCalE & ", '" & Format(pdFechaMax, "YYYY/MM/DD") & "', " & nEstado & ", '" & psPersCodMV & "'"
    oCon.Ejecutar sql
    oCon.CierraConexion
    'InsertaMovCont = 0
    Exit Sub
InsertarLineaCreditoErr:
    Call RaiseError(MyUnhandledError, "DLineaCreditoV2:LineaCreditoAgencia Method")
End Sub

Public Function ObtenerLineaCredito(ByVal psLineaCreditoCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_LineaCredito '" & psLineaCreditoCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerLineaCredito = rs
End Function

Public Function ObtenerLineaCreditoAgencia(ByVal psLineaCreditoCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_LineaCreditoAgencia '" & psLineaCreditoCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerLineaCreditoAgencia = rs
End Function

Public Function ObtenerLineaCreditoPriorizar(ByVal psAgencia As String, ByVal psMoneda As String, ByVal pdFecha As Date) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_LineaCreditoPrioridad '" & psAgencia & "','" & psMoneda & "','" & Format(pdFecha, "YYYY/MM/DD") & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerLineaCreditoPriorizar = rs
End Function

Public Function ObtenerLineaCreditoMoneda() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_LineaCreditoMoneda "
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerLineaCreditoMoneda = rs
End Function
'nPrioridad , cLineaCreditoCod, cLineaCreditoDes, nMontoCol, nSaldoLinea, nSaldoCol, nPorceCol, nDiferenc, nPorMax, nEstado, cAgeCod, cMoneda
Public Sub InsertaLineaCreditoPrioridad(ByVal pnPrioridad As Integer, ByVal psLineaCreditoCod As String, ByVal psLineaCreditoDes As String, ByVal pnMontoCol As Currency, ByVal nSaldoLinea As Currency, ByVal pnSaldoCol As Currency, ByVal pnPorceCol As Currency, ByVal pnDiferenc As Currency, ByVal pnPorMax As Currency, ByVal pnEstado As Integer, ByVal psAgeCod As String, ByVal psMoneda As String)
    On Error GoTo InsertaLineaCreditoPrioridadErr
    Dim sql As String
    Dim oCon As New DConecta
    'InsertaLineaCreditoTipoCredito = 1
    oCon.AbreConexion
    sql = "exec stp_ins_LineaCreditoPrioridad " & pnPrioridad & ",'" & psLineaCreditoCod & "','" & psLineaCreditoDes & "'," & pnMontoCol & "," & nSaldoLinea & "," & pnSaldoCol & ", " & pnPorceCol & "," & pnDiferenc & "," & pnPorMax & "," & pnEstado & ", '" & psAgeCod & "', '" & psMoneda & "'"
    oCon.Ejecutar sql
    oCon.CierraConexion
    'InsertaLineaCreditoTipoCredito = 0
    Exit Sub
InsertaLineaCreditoPrioridadErr:
    Call RaiseError(MyUnhandledError, "DLineaCreditoV2:InsertaLineaCreditoPrioridad Method")
End Sub
Public Function ObtenerLineaCreditoPriorizarImprimir(ByVal psAgencia As String, ByVal psMoneda As String, ByVal pdFecha As Date) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_LineaCreditoPrioridadImprimir '" & psAgencia & "','" & psMoneda & "','" & Format(pdFecha, "YYYY/MM/DD") & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerLineaCreditoPriorizarImprimir = rs
End Function

Public Function ObtenerConstanteLineaCredito(ByVal pnConsCod As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_ConstanteLineaCredito '" & pnConsCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerConstanteLineaCredito = rs
End Function
Public Sub InsertaProductoCrediticio(ByVal psTpoProdCod As String, ByVal pnSoles As Integer, ByVal pnDolares As Integer, ByVal pnPersoneriaN As Integer, ByVal pnPersoneriaJ As Integer, ByVal pnPlazo As Integer, ByVal psPlazoMin As Integer, ByVal psPlazoMax As Integer, ByVal pbCalificacion As Integer, ByVal pbCalA As Integer, ByVal pbCalB As Integer, ByVal pbCalC As Integer, ByVal pbCalD As Integer, ByVal pbCalE As Integer, ByVal pnGenero As Integer, ByVal pnGeneroM As Integer, ByVal pnGeneroF As Integer, ByVal pnCalificacion As Integer, ByVal pnCalificacionSBSPor As Integer, ByVal pnEdad As Integer, ByVal pnEdadMin As Integer, ByVal pnEdadMax As Integer)
    On Error GoTo InsertaProductoCrediticioErr
    Dim sql As String
    Dim oCon As New DConecta
    'InsertaLineaCreditoTipoCredito = 1
    oCon.AbreConexion
    sql = "exec stp_ins_ProductoCrediticio '" & psTpoProdCod & "'," & pnSoles & "," & pnDolares & "," & pnPersoneriaN & "," & pnPersoneriaJ & "," & pnPlazo & "," & psPlazoMin & "," & psPlazoMax & "," & pbCalA & "," & pbCalB & "," & pbCalC & "," & pbCalD & "," & pbCalE & "," & pnGenero & "," & pnGeneroM & "," & pnGeneroF & "," & pnCalificacion & "," & pnCalificacionSBSPor & "," & pnEdad & "," & pnEdadMin & "," & pnEdadMax & ",1"
    oCon.Ejecutar sql
    oCon.CierraConexion
    'InsertaLineaCreditoTipoCredito = 0
    Exit Sub
InsertaProductoCrediticioErr:
    Call RaiseError(MyUnhandledError, "DLineaCreditoV2:InsertaProductoCrediticio Method")
End Sub

Public Sub InsertaProductoCrediticioTasas(ByVal psTpoProdCod As String, ByVal pnMoneda As Integer, ByVal pnOrden As Integer, ByVal pnMontoDesde As Currency, ByVal pnMontoHasta As Currency, ByVal pnTipoMinino As Currency, ByVal pnTipoMaximo As Currency)
    On Error GoTo InsertaProductoCrediticioTasasErr
    Dim sql As String
    Dim oCon As New DConecta
    'InsertaLineaCreditoTipoCredito = 1
    oCon.AbreConexion
    sql = "exec stp_ins_ProductoCrediticioTasas '" & psTpoProdCod & "'," & pnMoneda & "," & pnOrden & "," & pnMontoDesde & "," & pnMontoHasta & "," & pnTipoMinino & "," & pnTipoMaximo
    oCon.Ejecutar sql
    oCon.CierraConexion
    'InsertaLineaCreditoTipoCredito = 0
    Exit Sub
InsertaProductoCrediticioTasasErr:
    Call RaiseError(MyUnhandledError, "DLineaCreditoV2:InsertaProductoCrediticioTasas Method")
End Sub

Public Sub EliminarProductoCreditocioTasas(ByVal psTpoProdCod As String, ByVal pnMoneda As Integer)
    On Error GoTo InsertaProductoCrediticioTasasErr
    Dim sql As String
    Dim oCon As New DConecta
    'InsertaLineaCreditoTipoCredito = 1
    oCon.AbreConexion
    sql = "exec stp_del_ProductoCreditocioTasas '" & psTpoProdCod & "'," & pnMoneda & ""
    oCon.Ejecutar sql
    oCon.CierraConexion
    'InsertaLineaCreditoTipoCredito = 0
    Exit Sub
InsertaProductoCrediticioTasasErr:
    Call RaiseError(MyUnhandledError, "DLineaCreditoV2:InsertaProductoCrediticioTasas Method")
End Sub
Public Sub InsertaProductoCreditocioAgencia(ByVal psTpoProdCod As String, ByVal psAgeCod As String, ByVal pnLogico As Integer)
    On Error GoTo InsertaProductoCreditocioAgenciaErr
    Dim sql As String
    Dim oCon As New DConecta
    'InsertaLineaCreditoAgencia = 1
    oCon.AbreConexion
    sql = "exec stp_ins_ProductoCreditocioAgencia '" & psTpoProdCod & "','" & psAgeCod & "'," & pnLogico
    oCon.Ejecutar sql
    oCon.CierraConexion
    'InsertaLineaCreditoAgencia = 0
    Exit Sub
InsertaProductoCreditocioAgenciaErr:
    Call RaiseError(MyUnhandledError, "DLineaCreditoV2:InsertaProductoCreditocioAgencia Method")
End Sub
Public Sub InsertaProductoCreditocioDestino(ByVal psTpoProdCod As String, ByVal psDestinoCod As String, ByVal pnLogico As Integer)
    On Error GoTo InsertaProductoCreditocioDestinoErr
    Dim sql As String
    Dim oCon As New DConecta
    'InsertaLineaCreditoAgencia = 1
    oCon.AbreConexion
    sql = "exec stp_ins_ProductoCreditocioDestino '" & psTpoProdCod & "','" & psDestinoCod & "'," & pnLogico
    oCon.Ejecutar sql
    oCon.CierraConexion
    'InsertaLineaCreditoAgencia = 0
    Exit Sub
InsertaProductoCreditocioDestinoErr:
    Call RaiseError(MyUnhandledError, "DLineaCreditoV2:InsertaProductoCreditocioDestino Method")
End Sub
Public Function ObtenerProductoCrediticio(ByVal psTpoProdCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_ProductoCrediticio '" & psTpoProdCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerProductoCrediticio = rs
End Function
Public Function ObtenerProductoCrediticioTasas(ByVal psTpoProdCod As String, ByVal pnMoneda As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_ProductoCrediticioTasas '" & psTpoProdCod & "','" & pnMoneda & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerProductoCrediticioTasas = rs
End Function
'Public Function ObtenerProductoCrediticio(ByVal psTpoProdCod As String) As ADODB.Recordset
'    Dim sSql As String
'    Dim oCon As New DConecta
'    Dim rs As New ADODB.Recordset
'    sSql = "exec stp_sel_ProductoCrediticio '" & psTpoProdCod & "'"
'    oCon.AbreConexion
'    Set rs = oCon.CargaRecordSet(sSql)
'    oCon.CierraConexion
'    Set ObtenerConstanteLineaCredito = rs
'End Function
Public Function ObtenerProductoCreditocioAgencia(ByVal psTpoProdCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_ProductoCreditocioAgencia '" & psTpoProdCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerProductoCreditocioAgencia = rs
End Function
Public Function ObtenerProductoCreditocioDestino(ByVal psTpoProdCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_ProductoCreditocioDestino '" & psTpoProdCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerProductoCreditocioDestino = rs
End Function
Public Function CorrelativoAdeudo(ByVal psLinea As String, ByVal psPersCod As String, ByVal psIFCod As String, ByVal psCtaIFCod As String) As String
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset
    
    'Se Aumentaron los 2 digitos del Paquete
    'sSql = "select MAX(SUBSTRING(cLineaCred,12,2)) as nCorrel "
    'sSql = sSql & " From ColocLineaCredito Where cLineaCred like '" & Left(psLinea, 6) & "'+'___'+'" & Right(psLinea, 2) & "%'"
    'sSql = sSql & " and isnull(cCtaIFCod,'')<>'' "
    sSql = "Exec stp_sel_ObtieneCorrelativoAdeudo '" & psLinea & "','" & psPersCod & "','" & psIFCod & "','" & psCtaIFCod & "'" 'Creado by NAGL 20190213 Según INC1902130002
    '**********Creado by NAGL 20190213*********
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If R.RecordCount > 0 And R!nCorrel <> "" Then
        CorrelativoAdeudo = Right("00" & Trim(Str(IIf(IsNull(R!nCorrel), 0, R!nCorrel))), 2)
    Else
        CorrelativoAdeudo = "00"
    End If
    R.Close
    Set R = Nothing
End Function
Public Function Correlativo(ByVal psLinea As String) As String
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset
    
    'Se Aumentaron los 2 digitos del Paquete
    sSql = "select MAX(SUBSTRING(cLineaCred,12,2)) as nCorrel "
    sSql = sSql & " From ColocLineaCredito where cLineaCred like '" & Left(psLinea, 6) & "'+'___'+'" & Right(psLinea, 2) & "%'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If R.RecordCount > 0 And R!nCorrel <> "" Then
        Correlativo = Right("00" & Trim(Str(IIf(IsNull(R!nCorrel), 0, R!nCorrel) + 1)), 2)
    Else
        Correlativo = "01"
    End If
    R.Close
    Set R = Nothing
End Function
Public Function ObtenerSacarCodigoLinea(ByVal psPersCod As String, psIFTpo As String, psCtaIFCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_SacarCodigoLinea '" & psPersCod & "','" & psIFTpo & "','" & psCtaIFCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerSacarCodigoLinea = rs
End Function
Public Function NuevaLinea(ByVal pnLongitud As Integer, ByVal psLineaCred As String, _
    ByVal psDescripcion As String, ByVal pbEstado As Integer, ByVal pnPlazoMax As Double, _
    ByVal pnPlazoMin As Double, ByVal pnMontoMax As Double, ByVal pnMontoMin As Double, _
    ByVal psPersCod As String, ByVal psDescFondo As String, ByVal psDescSubFondo As String, _
    ByVal psDescProduc As String, Optional ByVal psAbrevFondo As String = "", Optional ByVal psAbrevSubFondo As String = "", _
    Optional ByVal pbPreferencial As Boolean = False, _
    Optional ByVal pMatAgencias As Variant, _
    Optional ByVal psIFCod As String = "", Optional ByVal psCtaIFCod As String = "", Optional ByVal psCorrelativo As String = "01") As String


Dim I, j, K As Integer
Dim Abv1, Abv2 As String
Dim psAdeudados As String

    On Error GoTo ErrorNuevaLinea
'    NuevaLinea = ""
'    psAdeudados = CorrelativoAdeudo(psLineaCred, psPersCod, psIFCod, psCtaIFCod)
'    If Trim(psAdeudados) = "00" Then
'        psAdeudados = Correlativo(psLineaCred)
'    Else
'        NuevaLinea = ModificarLinea(5, psLineaCred, psDescripcion, _
'                    pbEstado, pnPlazoMax, pnPlazoMin, pnMontoMax, _
'                    pnMontoMin, psPersCod, pbPreferencial, pMatAgencias)
'
'        Exit Function
'    End If
    'psAdeudados = psCorrelativo
    'psLineaCred = psLineaCred & psAdeudados

    NuevaLinea = ValidaDatos(pnLongitud, psLineaCred, psDescripcion, pbEstado, pnPlazoMax, pnPlazoMin, pnMontoMax, pnMontoMin, psPersCod)
    If NuevaLinea <> "" Then
        Exit Function
    End If
    
    If ExisteLinea(psLineaCred) Then
        NuevaLinea = "Linea de Credito ya Existe"
        Exit Function
    End If
        
    For I = 1 To 4
        Select Case I
            Case 1 'Crear el Fondo
                If Not ExisteLinea(Mid(psLineaCred, 1, 2)) Then
                    Call NuevaLineaCredito(Mid(psLineaCred, 1, 2), psDescFondo, 1, 0, 0, 0#, 0#, psPersCod, psAbrevFondo)
                End If
            Case 2 'Crear el Sub fondo
                If Not ExisteLinea(Mid(psLineaCred, 1, 5)) Then
                    Call NuevaLineaCredito(Mid(psLineaCred, 1, 5), psDescSubFondo, 1, 0, 0, 0#, 0#, psPersCod, psAbrevSubFondo, , , psIFCod, psCtaIFCod)
                End If
            Case 3 'Crear en el Plazo Escojido
                If Not ExisteLinea(Mid(psLineaCred, 1, 6)) Then
                    If Mid(psLineaCred, 6, 1) = "1" Then
                        Call NuevaLineaCredito(Mid(psLineaCred, 1, 6), psDescSubFondo & " - CORTO PL", 1, 0, 0, 0#, 0#, psPersCod, , , , psIFCod, psCtaIFCod)
                    Else
                        Call NuevaLineaCredito(Mid(psLineaCred, 1, 6), psDescSubFondo & " - LARGO PL", 1, 0, 0, 0#, 0#, psPersCod, , , , psIFCod, psCtaIFCod)
                    End If
                End If
            Case 4 'Crear el Producto
                    K = -1
                    Abv1 = Mid(psDescProduc, 1, 3)
                    For j = 1 To Len(psDescProduc)
                        If Mid(psDescProduc, j, 1) = " " Then
                            K = j + 1
                            Exit For
                        End If
                    Next j
                    If K <> -1 Then
                        Abv2 = Mid(psDescProduc, K, 3)
                    End If
                    If Not ExisteLinea(Mid(psLineaCred, 1, 5)) Then
                        Call NuevaLineaCredito(Mid(psLineaCred, 1, 5), Abv1 & "-" & Abv2, 1, 0, 0, 0#, 0#, psPersCod, , , , psIFCod, psCtaIFCod)
                    End If
                    
        End Select
    Next I
    
    Call NuevaLineaCredito(psLineaCred, psDescripcion, pbEstado, pnPlazoMax, pnPlazoMin, pnMontoMax, pnMontoMin, psPersCod, , pbPreferencial, pMatAgencias, psIFCod, psCtaIFCod)
    
    
    Exit Function

ErrorNuevaLinea:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Sub NuevaLineaCredito(ByVal psLineaCredCod As String, ByVal psDescription As String, _
    ByVal pbEstado As Integer, ByVal pnPlazoMax As Integer, ByVal pnPlazoMin As Integer, _
    ByVal pnMontoMax As Double, ByVal pnMontoMin As Double, ByVal psPersCod As String, _
    Optional ByVal psAbrev As String = "", Optional ByVal pbPreferencial As Boolean = False, _
    Optional ByVal pMatAgencias As Variant, _
    Optional ByVal psIFCod As String = "", Optional ByVal psCtaIFCod As String = "")

Dim sSql As String
Dim oConecta As DConecta
Dim I As Integer
    
    On Error GoTo ErrorNuevaLineaCredito
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = "INSERT INTO ColocLineaCredito(cLineaCred, cDescripcion, bEstado, nPlazoMax, nPlazoMin, nMontoMax, nMontoMin,cPersCod,cAbrev,bPreferencial,cIFTpo,cCtaIFCod) "
    sSql = sSql & " VALUES('" & psLineaCredCod & "','" & psDescription & "'," & Trim(Str(pbEstado)) & "," & Format(pnPlazoMax, "#0") & "," & Format(pnPlazoMin, "#0") & "," & Format(pnMontoMax, "#0.00") & "," & Format(pnMontoMin, "#0.00") & ",'" & psPersCod & "','" & Trim(psAbrev) & "'," & IIf(pbPreferencial = True, 1, 0) & ",'" & psIFCod & "','" & psCtaIFCod & "')"
    oConecta.ConexionActiva.Execute sSql
    
    'Para el Saldo de Linea de Credito
    sSql = "INSERT INTO ColocLineaCreditoSaldo(cLineaCred, nMontoTotal, nSaldoCap, nMontoColocado, nMoneda, nMontoReservado)"
    sSql = sSql & " VALUES('" & psLineaCredCod & "',0.00,0.00,0.00,0,0.00)"
    oConecta.ConexionActiva.Execute sSql
    
    '**** Manejo de Lineas de Credito por Agencia ****
    If IsArray(pMatAgencias) Then
        For I = 0 To UBound(pMatAgencias) - 1
            sSql = "INSERT INTO ColocLineaCreditoAgencia(cLineaCred,cAgeCod)"
            sSql = sSql & " VALUES('" & psLineaCredCod & "','" & pMatAgencias(I) & "')"
            oConecta.ConexionActiva.Execute sSql
        Next I
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Sub
    
ErrorNuevaLineaCredito:
    Err.Raise Err.Number, "Nueva LineaCredito", Err.Description
    
End Sub

Public Function ExisteLinea(ByVal psLineaCred As String) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As DConecta

    On Error GoTo ErrorExisteLinea
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = "Select cLineaCred from ColocLineacredito Where cLineaCred = '" & psLineaCred & "'"
    Set R = oConecta.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        ExisteLinea = True
    Else
        ExisteLinea = False
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorExisteLinea:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function ModificarLinea(ByVal pnLongitud As Integer, ByVal psLineaCred As String, ByVal psDescripcion As String, ByVal pbEstado As Integer, ByVal pnPlazoMax As Double, ByVal pnPlazoMin As Double, ByVal pnMontoMax As Double, ByVal pnMontoMin As Double, _
                                ByVal psPersCod As String, Optional ByVal pbPreferencia As Boolean = False, Optional ByVal pMatAgencias As Variant, Optional ByVal psIFTpo As String = "", Optional ByVal psCtaIFCod As String = "") As String


    On Error GoTo ErrorModificar
    ModificarLinea = ""
    ModificarLinea = ValidaDatos(pnLongitud, psLineaCred, psDescripcion, pbEstado, pnPlazoMax, pnPlazoMin, pnMontoMax, pnMontoMin, psPersCod)
    If ModificarLinea <> "" Then
        Exit Function
    End If
    
    If pnLongitud <> 1 Then
        Call ActualizarLinea(psLineaCred, psDescripcion, pbEstado, pnPlazoMax, pnPlazoMin, pnMontoMax, pnMontoMin, RecuperaInstitucion(psLineaCred), pbPreferencia, pMatAgencias, psIFTpo, psCtaIFCod)
    Else
        Call ActualizarLinea(psLineaCred, psDescripcion, pbEstado, pnPlazoMax, pnPlazoMin, pnMontoMax, pnMontoMin, psPersCod, pbPreferencia, pMatAgencias, psIFTpo, psCtaIFCod)
    End If
    

    Exit Function

ErrorModificar:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function
Public Function RecuperaInstitucion(ByVal psLineaCred As String) As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset
Dim sSql As String
    
    On Error GoTo ErrorRecuperaInstitucion
    RecuperaInstitucion = ""
    sSql = "Select cPersCod from ColocLineacredito where cLineaCred = '" & Mid(psLineaCred, 1, 2) & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        RecuperaInstitucion = R!cPerscod
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorRecuperaInstitucion:
    Err.Raise Err.Number, "", Err.Description
End Function
Private Function ValidaDatos(ByVal pnLongitud As Integer, ByVal psLineaCred As String, ByVal psDescripcion As String, ByVal pbEstado As Integer, ByVal pnPlazoMax As Double, ByVal pnPlazoMin As Double, ByVal pnMontoMax As Double, ByVal pnMontoMin As Double, ByVal psPersCod As String) As String
Dim nLongTemp As Integer
Dim oConecta As DConecta
Dim sSql As String
Dim R As ADODB.Recordset
Dim I As Integer

    ValidaDatos = ""
    Select Case pnLongitud
        Case 1
            nLongTemp = 2
        Case 2
            nLongTemp = 5
        Case 3
            nLongTemp = 6
        Case 4
            nLongTemp = 9
        Case 5
            nLongTemp = 13 '11(Nueva Longitud 13)
    End Select
    
    'Valida Existencia de Descendientes
    Set oConecta = New DConecta

    Set oConecta = Nothing
    
    'Valida Longitud
    If Len(Trim(psLineaCred)) <> nLongTemp Then
        ValidaDatos = "La Longitud del Codigo de la Linea de Credito No es Correcta"
        Exit Function
    End If
        
    'Valida Plazos
    If pnPlazoMax <= 0# And pnLongitud = 5 Then
        ValidaDatos = "El Plazo Maximo debe ser mayor que Cero"
        Exit Function
    End If
    If pnPlazoMin <= 0# And pnLongitud = 5 Then
        ValidaDatos = "El Plazo Minimo debe ser mayor que Cero"
        Exit Function
    End If
    
    If pnPlazoMin > pnPlazoMax Then
        ValidaDatos = "El Plazo Minimo no debe ser mayor el Plazo Maximo"
        Exit Function
    End If
    
    'Valida Monto Maximo y Minimo
    If pnMontoMin <= 0# And pnLongitud = 5 Then
        ValidaDatos = "El Monto Minimo debe ser mayor que Cero"
        Exit Function
    End If
    If pnMontoMax <= 0# And pnLongitud = 5 Then
        ValidaDatos = "El Monto Maximo debe ser mayor que Cero"
        Exit Function
    End If
    
    If pnMontoMin > pnMontoMax Then
        ValidaDatos = "El Monto Minimo no debe ser mayor el Monto Maximo"
        Exit Function
    End If
    
    'Valida Institucion
    If Len(Trim(psPersCod)) = 0 And pnLongitud = 1 Then
        ValidaDatos = "Debe Ingresar La Institucion"
        Exit Function
    End If
    
    
End Function

Public Sub ActualizarLinea(ByVal psLineaCredCod As String, ByVal psDescription As String, _
                        ByVal pbEstado As Integer, ByVal pnPlazoMax As Integer, ByVal pnPlazoMin As Integer, _
                        ByVal pnMontoMax As Double, ByVal pnMontoMin As Double, ByVal psPersCod As String, _
                        Optional ByVal pbPreferencial As Boolean = False, Optional ByVal pMatAgencias As Variant, Optional ByVal psIFTpo As String = "", Optional ByVal psCtaIFCod As String = "")
Dim sSql As String
Dim oConecta As DConecta
Dim I As Integer

    On Error GoTo ErrorActualizarLinea
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = "UPDATE ColocLineaCredito SET cDescripcion = '" & psDescription & "',"
    sSql = sSql & " bEstado = " & Trim(Str(pbEstado)) & ","
    sSql = sSql & " nPlazoMax = " & Format(pnPlazoMax, "#0") & ","
    sSql = sSql & " nPlazoMin = " & Format(pnPlazoMin, "#0") & ","
    sSql = sSql & " nMontoMax = " & Format(pnMontoMax, "#0.00") & ","
    sSql = sSql & " nMontoMin = " & Format(pnMontoMin, "#0.00") & ","
    sSql = sSql & " cPersCod = '" & psPersCod & "',"
    sSql = sSql & " bPreferencial=" & IIf(pbPreferencial = True, 1, 0) & ","
    sSql = sSql & " cIFTpo='" & psIFTpo & "',"
    sSql = sSql & " cCtaIFCod='" & psCtaIFCod & "'"
    sSql = sSql & " WHERE cLineaCred = '" & psLineaCredCod & "'"
    oConecta.ConexionActiva.Execute sSql
    
   'CUSCO
   sSql = "DELETE FROM ColocLineaCreditoAgencia WHERE cLineaCred='" & psLineaCredCod & "'"
   oConecta.ConexionActiva.Execute sSql
   
    If IsArray(pMatAgencias) Then
        For I = 0 To UBound(pMatAgencias) - 1
            sSql = "INSERT INTO ColocLineaCreditoAgencia(cLineaCred,cAgeCod)"
            sSql = sSql & " VALUES('" & psLineaCredCod & "','" & pMatAgencias(I) & "')"
            oConecta.ConexionActiva.Execute sSql
        Next I
    End If
    '************************************
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Sub
ErrorActualizarLinea:
    Err.Raise Err.Number, "Actualiza Linea", Err.Description
    
End Sub

Public Function RecuperaSubFondos(ByVal psFondo As String, ByVal nTipo As Integer) As String
Dim sSql As String
Dim oRs As ADODB.Recordset
Dim oConecta As DConecta
Set oRs = New ADODB.Recordset

    sSql = "Select DISTINCT SUBSTRING(cLineaCred,1,4) as cSubFondo, cDescripcion, cAbrev "
    sSql = sSql & " From ColocLineaCredito "
    sSql = sSql & " Where Len(cLineaCred) = 5 and cLineaCred like '" & psFondo & "%'"
    sSql = sSql & " Group by SUBSTRING(cLineaCred,1,4), cDescripcion, cAbrev"
    sSql = sSql & " Order by cDescripcion"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    RecuperaSubFondos = ""
    If Not (oRs.BOF Or oRs.EOF) Then
        If nTipo = 1 Then
            RecuperaSubFondos = oRs!cAbrev
        Else
            RecuperaSubFondos = oRs!cDescripcion
        End If
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

End Function

Public Function RecuperaFondos(ByVal psPersCod As String, ByVal nTipo As Integer) As String
Dim sSql As String
Dim oRs As ADODB.Recordset
Dim oConecta As DConecta
Set oRs = New ADODB.Recordset

    sSql = "Select top 1 L.cLineaCred, P.cPersNombre, L.cAbrev "
    sSql = sSql & " from ColocLineaCredito L Inner Join Persona P ON P.cPersCod = L.cPersCod"
    sSql = sSql & " Where Len(cLineaCred) = 2 "
    sSql = sSql & " and P.cPersCod = '" & psPersCod & "'"
    sSql = sSql & " Order by cPersNombre  "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    RecuperaFondos = ""
    If Not (oRs.BOF Or oRs.EOF) Then
        If nTipo = 1 Then
            RecuperaFondos = oRs!cAbrev
        Else
            RecuperaFondos = oRs!cpersNombre
        End If
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
End Function
Public Function RecuperaProductosDeCredito(Optional ByVal psCab As String = "", Optional ByVal nTipo As Integer) As String
Dim sSql As String
Dim oRs As ADODB.Recordset
Dim oConecta As DConecta
Set oRs = New ADODB.Recordset
    sSql = " Select C.cConsDescripcion, C.nConsValor, CF.cAbrev from Constante C "
    sSql = sSql & " Inner Join ColocCredConsFiltro CF ON C.nConsCod = CF.nConsCod AND C.nConsValor = CF.nConsValor"
    sSql = sSql & " Where CF.nCodFiltro = " & gCredFiltroProd
    sSql = sSql & " AND C.nConsCod =3034"
    If psCab <> "" Then
        sSql = sSql & " AND C.nConsValor like '" & Mid(psCab, 1, 1) & "%'"
    End If
    sSql = sSql & " Order by C.nConsValor  "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set oRs = oConecta.CargaRecordSet(sSql)
    RecuperaProductosDeCredito = ""
    If Not (oRs.BOF Or oRs.EOF) Then
        If nTipo = 1 Then
            RecuperaProductosDeCredito = oRs!cAbrev
        Else
            RecuperaProductosDeCredito = oRs!cConsDescripcion
        End If
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'Public Function RecuperaLineasCredito(ByVal pnLongitud As Integer, Optional ByVal psCriterio As String = "", Optional ByVal pcVer As String = "") As ADODB.Recordset
'Dim sSql As String
'Dim oConecta As COMConecta.DCOMConecta
'
'On Error GoTo ErrorRecuperaLineasCredito
'
'    If UCase$(pcVer) = "ECSA" Then
'        sSql = " Select cFondo = F.cLineaCred, Fondo = F.cDescripcion,"
'        sSql = sSql & " cSubFondo = SF.cLineaCred, SubFondo = SF.cDescripcion,"
'        sSql = sSql & " cLineaCred = LC.cLineaCred, cDescripcion = LC.cLineaCred + '-' + LC.cDescripcion,"
'        sSql = sSql & " Moneda = "
'        sSql = sSql & " Case"
'        sSql = sSql & "    When SubString(LC.cLineaCred, 5, 1) = 1 Then 'Soles'"
'        sSql = sSql & "    When SubString(LC.cLineaCred, 5, 1) = 2 Then 'Dolares'"
'        sSql = sSql & "    When SubString(LC.cLineaCred, 5, 1) = 3 Then 'VAC'"
'        sSql = sSql & "    Else 'No Definido'"
'        sSql = sSql & " End, "
'        sSql = sSql & " LC.bEstado, LC.nPlazoMax, LC.nPlazoMin, LC.nMontoMax,"
'        sSql = sSql & " LC.nMontoMin, LC.cPersCod, LC.cIFTpo, LC.cAbrev, LC.bPreferencial,"
'        sSql = sSql & " LCS.nMontoTotal , LCS.nSaldoCap, LCS.nMontoColocado"
'        sSql = sSql & " From ColocLineaCredito LC"
'        sSql = sSql & " Inner Join ColocLineaCredito SF On SubString(LC.cLineaCred, 1, 4) = SubString(SF.cLineaCred, 1, 4)"
'        sSql = sSql & " Inner Join ColocLineaCredito F On SubString(LC.cLineaCred, 1, 2) = SubString(F.cLineaCred, 1, 2)"
'        sSql = sSql & " Left Join ColocLineaCreditoSaldo LCS On LC.cLineaCred = LCS.cLineaCred"
'        sSql = sSql & " Where Len(LC.cLineaCred) = " & pnLongitud & " And Len(sF.cLineaCred) = 4 And Len(F.cLineaCred) = 2 And LC.bEstado = 1"
'        sSql = sSql & " ORder By 1, 3, 5"
'
'    Else
'         sSql = "Select CL.cLineaCred, CL.cDescripcion, CL.nPlazoMax, CL.nPlazoMin, CL.nMontoMax, CL.nMontoMin, P.cPersNombre + space(50) + P.cPersCod as PersCod, convert(int,CL.bEstado) as nEstado"
'         sSql = sSql & " From " & gConsComunes & "ColocLineacredito CL Inner Join Persona P ON CL.cPersCod = P.cPersCod "
'
'         Select Case pnLongitud
'             Case 1
'                 sSql = sSql & " WHERE LEN(RTrim(CL.cLineaCred))<=2"
'             Case 2
'                 sSql = sSql & " WHERE LEN(RTrim(CL.cLineaCred))<=5"
'             Case 3
'                 sSql = sSql & " WHERE LEN(RTrim(CL.cLineaCred))<=6"
'             Case 4
'                 sSql = sSql & " WHERE LEN(RTrim(CL.cLineaCred))<=9"
'             Case 5
'                 sSql = sSql & " WHERE LEN(RTrim(CL.cLineaCred))=13" 'Ahora son 13
'         End Select
'         If psCriterio <> "" Then
'             sSql = sSql & " AND cLineaCred like '" & psCriterio & "%' "
'         End If
'         sSql = sSql & " AND RTRIM(CL.cLineaCred)<>'' ORDER BY cLineaCred"
'    End If
'    Set oConecta = New COMConecta.DCOMConecta
'    oConecta.AbreConexion
'    Set RecuperaLineasCredito = oConecta.CargaRecordSet(sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorRecuperaLineasCredito:
'    Err.Raise Err.Number, "Linea de Credito", Err.Description
'End Function
Public Function BuscarCodigoLinea(ByVal psPersCod As String, psIFTpo As String, psCtaIFCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_buscar_lineaCredito '" & psPersCod & "','" & psIFTpo & "','" & psCtaIFCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set BuscarCodigoLinea = rs
End Function
Public Function ObtenerListaLineaCredito() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_LineaCreditoListar"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set ObtenerListaLineaCredito = rs
End Function
Public Function ObtenerListaLineaCreditoVer(ByVal psLineaCredCod As String) As Integer
    Dim sSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
    sSql = "exec stp_sel_LineaCreditoVer '" & psLineaCredCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    If Not (rs.BOF Or rs.EOF) Then
        ObtenerListaLineaCreditoVer = 1
    Else
        ObtenerListaLineaCreditoVer = 0
    End If
End Function
Public Sub EliminarLineaCreditoPrioridad(ByVal psAgeCod As String, ByVal psMoneda As String)
    On Error GoTo EliminarLineaCreditoPrioridadErr
    Dim sql As String
    Dim oCon As New DConecta
    'InsertaLineaCreditoTipoCredito = 1
    oCon.AbreConexion
    sql = "exec stp_del_LineaCreditoPrioridad '" & psAgeCod & "', '" & psMoneda & "'"
    oCon.Ejecutar sql
    oCon.CierraConexion
    'InsertaLineaCreditoTipoCredito = 0
    Exit Sub
EliminarLineaCreditoPrioridadErr:
    Call RaiseError(MyUnhandledError, "DLineaCreditoV2:EliminarLineaCreditoPrioridad Method")
End Sub

'JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
Public Function CabCorrelativo(ByVal psLinea As String, ByVal cPerscod As String) As String
Dim sSQL As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset
        
    sSQL = "exec stp_sel_ObtieneCorrelativoCab '" & psLinea & "','" & cPerscod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    CabCorrelativo = R!cLinea
    R.Close
    Set R = Nothing
End Function
'JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
