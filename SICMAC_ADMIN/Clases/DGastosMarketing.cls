VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DGastosMarketing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'VAPI SEGUN ERS 082-2014
    Dim oCon As New DConecta
'fin VAPI
Public Sub InsertaTipoActividad(ByVal psNombre As String, ByVal psDescripcion As String, ByVal pbEstado As Boolean)
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorInsertaTipoActividad
    Set oConecta = New DConecta
    sSql = "Exec stp_ins_MktTipoActividadGasto '" & psNombre & "','" & psDescripcion & "'," & IIf(pbEstado = True, 1, 0)
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInsertaTipoActividad:
    Err.Raise Err.Number, "Inserta Tipo Actividad", Err.Description
End Sub
Public Sub ActualizaTipoActividad(ByVal pnId As Long, ByVal psNombre As String, ByVal psDescripcion As String, ByVal pbEstado As Boolean)
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorInsertaTipoActividad
    Set oConecta = New DConecta
    sSql = "Exec stp_upd_MktTipoActividadGasto " & pnId & ",'" & psNombre & "','" & psDescripcion & "'," & IIf(pbEstado = True, 1, 0)
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInsertaTipoActividad:
    Err.Raise Err.Number, "Actualiza Tipo Actividad", Err.Description
End Sub
Public Function RecuperaTipoActividad() As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorInsertaTipoActividad
    Set oConecta = New DConecta
    sSql = "Exec stp_sel_MktTipoActividadGasto"
    oConecta.AbreConexion
    Set RecuperaTipoActividad = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorInsertaTipoActividad:
    Err.Raise Err.Number, "Recupera Tipo Actividad", Err.Description
End Function
Public Function RecuperaTipoActividadxEstado(ByVal pbEstadoActivo As Boolean) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaTpoAct
    Set oConecta = New DConecta
    sSql = "Exec stp_sel_MktTipoActividadxEstado " & IIf(pbEstadoActivo = True, 1, 0)
    oConecta.AbreConexion
    Set RecuperaTipoActividadxEstado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaTpoAct:
    Err.Raise Err.Number, "Recupera Tipo Actividad", Err.Description
End Function
Public Sub InsertaCategoriaGasto(ByVal psNombre As String, ByVal psDescripcion As String, ByVal pbEstado As Boolean)
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorInsertaCatGasto
    Set oConecta = New DConecta
    sSql = "Exec stp_ins_MktCategoriaGasto '" & psNombre & "','" & psDescripcion & "'," & IIf(pbEstado = True, 1, 0)
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInsertaCatGasto:
    Err.Raise Err.Number, "Inserta Gasto", Err.Description
End Sub
Public Sub ActualizaCategoriaGasto(ByVal pnId As Long, ByVal psNombre As String, ByVal psDescripcion As String, ByVal pbEstado As Boolean)
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorActualizaCatGasto
    Set oConecta = New DConecta
    sSql = "Exec stp_upd_MktCategoriaGasto " & pnId & ",'" & psNombre & "','" & psDescripcion & "'," & IIf(pbEstado = True, 1, 0)
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorActualizaCatGasto:
    Err.Raise Err.Number, "Inserta Gasto", Err.Description
End Sub
Public Function RecuperaCategoriaGasto() As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaCatGasto
    Set oConecta = New DConecta
    sSql = "Exec stp_sel_MktCategoriaGasto"
    oConecta.AbreConexion
    Set RecuperaCategoriaGasto = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaCatGasto:
    Err.Raise Err.Number, "Recupera Tipo Actividad", Err.Description
End Function
Public Function RecuperaCategoriaGastoxEstado(ByVal pbEstadoActivo As Boolean) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaCatGasto
    Set oConecta = New DConecta
    sSql = "Exec stp_sel_MktCategoriaGastoxEstado " & IIf(pbEstadoActivo = True, 1, 0)
    oConecta.AbreConexion
    Set RecuperaCategoriaGastoxEstado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaCatGasto:
    Err.Raise Err.Number, "Recupera Tipo Actividad", Err.Description
End Function
Public Sub InsertaProductoServicio(ByVal psNombre As String, ByVal pnCategoriaGasto As Long, ByVal pnUnidad As Integer)
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorNuevoProducto
    Set oConecta = New DConecta
    sSql = "Exec stp_ins_MktProductoServicio '" & psNombre & "'," & pnCategoriaGasto & "," & pnUnidad
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorNuevoProducto:
    Err.Raise Err.Number, "Nuevo Producto o Servicio", Err.Description
End Sub
Public Sub ActualizaProductoServicio(ByVal pnCodProducto As Long, ByVal psNombre As String, ByVal pnCategoriaGasto As Long, ByVal pnUnidad As Integer)
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorActualizaProducto
    Set oConecta = New DConecta
    sSql = "Exec stp_upd_MktProductoServicio " & pnCodProducto & ",'" & psNombre & "'," & pnCategoriaGasto & "," & pnUnidad
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorActualizaProducto:
    Err.Raise Err.Number, "Actualiza Producto o Servicio", Err.Description
End Sub
Public Function RecuperaProductoServicio() As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaProductoServicio
    Set oConecta = New DConecta
    sSql = "Exec stp_sel_MktProductoServicio"
    oConecta.AbreConexion
    Set RecuperaProductoServicio = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaProductoServicio:
    Err.Raise Err.Number, "Recupera Producto Servicio", Err.Description
End Function
Public Function RecuperaProductoPaArbol() As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaProductoPaArbol
    Set oConecta = New DConecta
    sSql = "Exec stp_sel_MktProductoPaArbol"
    oConecta.AbreConexion
    Set RecuperaProductoPaArbol = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaProductoPaArbol:
    Err.Raise Err.Number, "Recupera Producto Arbol", Err.Description
End Function
Public Function RecuperaActividad() As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaActividad
    Set oConecta = New DConecta
    sSql = "Exec stp_sel_MktActividad"
    oConecta.AbreConexion
    Set RecuperaActividad = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaActividad:
    Err.Raise Err.Number, "Recupera Actividad", Err.Description
End Function
Public Function RecuperaActividadxId(ByVal pnActividad As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaActividad
    Set oConecta = New DConecta
    sSql = "Exec stp_sel_MktActividadxId " & pnActividad
    oConecta.AbreConexion
    Set RecuperaActividadxId = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaActividad:
    Err.Raise Err.Number, "Recupera Actividad", Err.Description
End Function
Public Function RecuperaActividadPaArbol() As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaActividadPaArbol
    Set oConecta = New DConecta
    sSql = "Exec stp_sel_MktActividadPaArbol"
    oConecta.AbreConexion
    Set RecuperaActividadPaArbol = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaActividadPaArbol:
    Err.Raise Err.Number, "Recupera Actividad Arbol", Err.Description
End Function
Public Sub InsertaActividad(ByVal psNombreAct As String, ByVal pbPreProgramado As Boolean, ByVal pdFecIni As Date, ByVal pdFecFin As Date, psDescripcion As String, pnTipoAct As Long)
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorInsertaActividad
    Set oConecta = New DConecta
    sSql = "Exec stp_ins_MktActividad '" & psNombreAct & "'," & IIf(pbPreProgramado = True, 1, 0) & ",'" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & psDescripcion & "'," & pnTipoAct
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInsertaActividad:
    Err.Raise Err.Number, "Inserta Actividad", Err.Description
End Sub
Public Sub ActualizaActividad(ByVal pnActividadId As Long, ByVal psNombreAct As String, ByVal pbPreProgramado As Boolean, ByVal pdFecIni As Date, ByVal pdFecFin As Date, psDescripcion As String, pnTipoAct As Long)
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorActualizaActividad
    Set oConecta = New DConecta
    sSql = "Exec stp_upd_MktActividad " & pnActividadId & ",'" & psNombreAct & "'," & IIf(pbPreProgramado = True, 1, 0) & ",'" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & psDescripcion & "'," & pnTipoAct
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorActualizaActividad:
    Err.Raise Err.Number, "Inserta Actividad", Err.Description
End Sub
Public Function RecuperaUnidadxEstado(ByVal pbEstado As Boolean) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorActualizaActividad
    Set oConecta = New DConecta
    sSql = "Exec stp_sel_MktUnidadxEstado " & IIf(pbEstado = True, 1, 0)
    oConecta.AbreConexion
    Set RecuperaUnidadxEstado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorActualizaActividad:
    Err.Raise Err.Number, "Recupera Unidad", Err.Description
End Function
Public Sub InsertaCompra(ByVal pnProductoCod As Long, ByVal pdFechaCompra As Date, ByVal pnPrecioUnit As Double, ByVal pnMoneda As Integer, ByVal pnCantidad As Integer, ByVal pnTotal As Double, ByVal psUser As String)
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorInsertaCompra
    Set oConecta = New DConecta
    sSql = "Exec stp_ins_MktCompra " & pnProductoCod & ",'" & Format(pdFechaCompra, "yyyymmdd") & "'," & pnMoneda & "," & pnPrecioUnit & "," & pnCantidad & "," & pnTotal & ",'" & psUser & "'"
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInsertaCompra:
    Err.Raise Err.Number, "Inserta Compra", Err.Description
End Sub
Public Function RecuperaProductoxActividad(ByVal psAgeCod As String, ByVal pnActividadId As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaProductoxActividad
    Set oConecta = New DConecta
    sSql = "Exec stp_sel_MktProductoxActividad '" & psAgeCod & "'," & pnActividadId
    oConecta.AbreConexion
    Set RecuperaProductoxActividad = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaProductoxActividad:
    Err.Raise Err.Number, "Recupera Producto x Actividad", Err.Description
End Function
Public Function RecuperaStockProducto(ByVal pnProductoId As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaStockProducto
    Set oConecta = New DConecta
    sSql = "Exec stp_sel_MktStockProducto " & pnProductoId
    oConecta.AbreConexion
    Set RecuperaStockProducto = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRecuperaStockProducto:
    Err.Raise Err.Number, "Recupera Stock Producto", Err.Description
End Function
Public Function RecuperaStockCompraxProducto(ByVal pnProductoId As Long) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrRecuperaStockCompraxProducto
    Set oConecta = New DConecta
    sSql = "Exec stp_sel_MktStockCompraxProducto " & pnProductoId
    oConecta.AbreConexion
    Set RecuperaStockCompraxProducto = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrRecuperaStockCompraxProducto:
    Err.Raise Err.Number, "Recupera Stock Compra", Err.Description
End Function
Public Sub EliminaProductoEnActividad(ByVal pnId As Long)
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrInsertaActividadProducto
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = "Delete from MktProductoActividadDet where nId=" & pnId
    oConecta.Ejecutar (sSql)
    sSql = "Delete from MktProductoActividad where nId=" & pnId
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrInsertaActividadProducto:
    Err.Raise Err.Number, "Inserta Actividad de Producto", Err.Description
End Sub
Public Function InsertaProductoEnActividad(ByVal psAgeCod As String, ByVal pnActividadId As Long, ByVal pnProdServId As Long, ByVal pdFecha As Date, ByVal psComentario As String) As Long
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oConecta As DConecta
    On Error GoTo ErrInsertaActividadProducto
    Set oConecta = New DConecta
    sSql = "Exec stp_ins_MktActividadProducto '" & psAgeCod & "'," & pnActividadId & "," & pnProdServId & ",'" & Format(pdFecha, "yyyymmdd") & "','" & psComentario & "'"
    InsertaProductoEnActividad = 0
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    If Not RSVacio(rs) Then
        InsertaProductoEnActividad = rs!nIdNuevo
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrInsertaActividadProducto:
    Err.Raise Err.Number, "Inserta Actividad de Producto", Err.Description
End Function
Public Sub InsertaProductoEnActividadDetalle(ByVal pnId As Long, ByVal pnCompraId As Long, ByVal pnCantidad As Long)
    Dim sSql As String
    Dim oConecta As DConecta
    On Error GoTo ErrInsertaProductoEnActividadDetalle
    Set oConecta = New DConecta
    sSql = "Exec stp_ins_MktActividadProductoDet " & pnId & "," & pnCompraId & "," & pnCantidad
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrInsertaProductoEnActividadDetalle:
    Err.Raise Err.Number, "Inserta Actividad de Producto Detalle", Err.Description
End Sub


'VAPI SEGUN ERS 082-2014
Public Function getOcon() As DConecta
    Set getOcon = oCon
End Function

Public Function RecuperaBienesServicios(ByVal pCodigoComodin, Optional pConSaldo As Integer, Optional pnAlma As Integer) As ADODB.Recordset
On Error GoTo ErrRecuperaBienesServicios
    Dim sqlA  As String
    If oCon.AbreConexion = False Then Exit Function
    sqlA = "EXEC stp_selBienesServicios " & "'" & pCodigoComodin & "', " & pConSaldo & ", " & pnAlma
    Set RecuperaBienesServicios = oCon.CargaRecordSet(sqlA)
    oCon.CierraConexion
    Exit Function
ErrRecuperaBienesServicios:
    Err.Raise Err.Number, "Recupera los bienes y servicios", Err.Description
End Function

Public Function RecuperaCampanas() As ADODB.Recordset
On Error GoTo ErrRecuperaCampanas
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = "Exec stp_sel_ObtieneCampanas"
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set RecuperaCampanas = rs
    Exit Function
ErrRecuperaCampanas:
    Err.Raise Err.Number, "Recupera las campañas", Err.Description
End Function

Public Function RecuperaCombosXCampana(ByVal nIdCampana As String) As ADODB.Recordset
On Error GoTo ErrRecuperaCampanas
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = "Exec ERS0822014_stp_sel_comboXcampana " & nIdCampana
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set RecuperaCombosXCampana = rs
    Exit Function
ErrRecuperaCampanas:
    Err.Raise Err.Number, "Recupera los combos por campañas", Err.Description
End Function
Public Function InsertaComboCampana(ByVal idCampana As Integer, ByVal cComboDescripcion As String, ByVal bDesembolso As Integer, ByVal bApertura As Integer, ByVal bSoles As Integer, ByVal bDolares As Integer, ByVal nMinSoles As String, ByVal nMaxSoles As String, ByVal nMinDolares As String, ByVal nMaxDolares As String) As Integer
    On Error GoTo ErrInsertaComboCampana
    Dim sql As String
    Dim RsSource  As ADODB.Recordset
    Set RsSource = New ADODB.Recordset
    sql = "EXEC ERS0822014_stp_ins_comboxcampana " _
                    & idCampana & ", " _
             & " '" & cComboDescripcion & "', " _
                    & bDesembolso & ", " _
                    & bApertura & ", " _
                    & bSoles & ", " _
                    & bDolares & ", " _
                    & nMinSoles & ", " _
                    & nMaxSoles & ", " _
                    & nMinDolares & ", " _
                    & nMaxDolares
    Set RsSource = oCon.CargaRecordSet(sql)
    InsertaComboCampana = RsSource!Cod
Exit Function
ErrInsertaComboCampana:
    Err.Raise Err.Number, "Insertar combo de campañas", Err.Description
End Function

Public Function InsertaDetalleCombo(ByVal nIdCombo As String, ByVal cBSCod As String, ByVal nCantidad As String) As Integer
On Error GoTo ErrInsertaDetalleCombo:
    Dim RsSource  As ADODB.Recordset
    Set RsSource = New ADODB.Recordset
    Dim sql As String
    sql = "EXEC ERS0822014_stp_ins_comboBienes " _
                            & nIdCombo & ", " _
                            & " '" & cBSCod & "', " _
                            & nCantidad
   Set RsSource = oCon.CargaRecordSet(sql)
   InsertaDetalleCombo = RsSource!Cod
Exit Function
ErrInsertaDetalleCombo:
    Err.Raise Err.Number, "Insertar los detalles combo de campañas", Err.Description
End Function

Public Function RecuperaComboBienes(ByVal nIdCombo As String) As ADODB.Recordset
On Error GoTo ErrRecuperaComboBienes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If oCon.AbreConexion = False Then Exit Function
    Dim sSql As String
    sSql = "Exec ERS0822014_stp_sel_comboBienes " & nIdCombo
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set RecuperaComboBienes = rs
    Exit Function
ErrRecuperaComboBienes:
    Err.Raise Err.Number, "Recupera los bienes por combos", Err.Description
End Function

Public Sub EliminaBienEnCombo(ByVal nIdComboBienes As String)
On Error GoTo ErrEliminaBienEnCombo:
     If oCon.AbreConexion = False Then Exit Sub
    Dim sql As String
    sql = "EXEC ERS0822014_stp_del_comboBienes " & nIdComboBienes
    oCon.Ejecutar (sql)
    oCon.CierraConexion
Exit Sub
ErrEliminaBienEnCombo:
    Err.Raise Err.Number, "Elimina un bien por combo", Err.Description
End Sub

Public Sub ActualizaComboxCampana(ByVal nIdCombo As Integer, ByVal idCampana As Integer, ByVal cComboDescripcion As String, ByVal bDesembolso As Integer, ByVal bApertura As Integer, ByVal bSoles As Integer, ByVal bDolares As Integer, ByVal nMinSoles As String, ByVal nMaxSoles As String, ByVal nMinDolares As String, ByVal nMaxDolares As String)
    On Error GoTo ErrActualizaComboxCampana:
     If oCon.AbreConexion = False Then Exit Sub
    Dim sql As String
    sql = "EXEC ERS0822014_stp_upd_comboxcampana " _
                    & nIdCombo & ", " _
                    & idCampana & ", " _
             & " '" & cComboDescripcion & "', " _
                    & bDesembolso & ", " _
                    & bApertura & ", " _
                    & bSoles & ", " _
                    & bDolares & ", " _
                    & nMinSoles & ", " _
                    & nMaxSoles & ", " _
                    & nMinDolares & ", " _
                    & nMaxDolares
    oCon.Ejecutar (sql)
    oCon.CierraConexion
Exit Sub
ErrActualizaComboxCampana:
    Err.Raise Err.Number, "Actualiza un combo de la campana", Err.Description
End Sub

Public Function ExisteDescripcionCombo(ByVal nIdCombo As String, ByVal nIdCampana As String, ByVal cComboDescripcion As String) As Integer
On Error GoTo ErrExisteDescripcionCombo:
    Dim RsSource  As ADODB.Recordset
    Set RsSource = New ADODB.Recordset
    If oCon.AbreConexion = False Then Exit Function
    Dim sql As String
    
    If nIdCombo = "" Then
        sql = "EXEC ERS0822014_stp_sel_ExisteDescripcionCombo null," & nIdCampana & "," & "'" & cComboDescripcion & "'"
    Else
        sql = "EXEC ERS0822014_stp_sel_ExisteDescripcionCombo " & nIdCombo & "," & nIdCampana & "," & "'" & cComboDescripcion & "'"
    End If
    
   Set RsSource = oCon.CargaRecordSet(sql)
   ExisteDescripcionCombo = RsSource!existe
   oCon.CierraConexion
Exit Function
ErrExisteDescripcionCombo:
    Err.Raise Err.Number, "Verifica si existe una descripción del combo", Err.Description
End Function

Public Sub EliminaComboCampana(ByVal nIdCombo As String)
On Error GoTo ErrEliminaComboCampana:
     If oCon.AbreConexion = False Then Exit Sub
    Dim sql As String
    sql = "EXEC ERS0822014_stp_del_comboxcampana " & nIdCombo
    oCon.Ejecutar (sql)
    oCon.CierraConexion
Exit Sub
ErrEliminaComboCampana:
    Err.Raise Err.Number, "Elimina un combo de campana", Err.Description
End Sub


Public Function InsertaEntregaCampana(ByVal cAgencia As String, ByVal cPersCodUser As String, ByVal cPersCodUserSol As String, ByVal cPersCodCliente As String, ByVal fecha As String, ByVal cGlosa As String, ByVal sMovNro As String) As Integer
    On Error GoTo ErrInsertaEntregaCampana
    Dim sql As String
    Dim RsSource  As ADODB.Recordset
    Set RsSource = New ADODB.Recordset
    sql = "EXEC ERS0822014_stp_ins_entregacampana " _
                & " '" & cAgencia & "', " _
                & " '" & cPersCodUser & "', " _
                & " '" & cPersCodUserSol & "', " _
                & " '" & cPersCodCliente & "', " _
                & " '" & fecha & "', " _
                & " '" & cGlosa & "'," _
                & " '" & sMovNro & "'," _
                       & "1" & "," _
                       & "NULL"

    Set RsSource = oCon.CargaRecordSet(sql)
    InsertaEntregaCampana = RsSource!Cod
Exit Function
ErrInsertaEntregaCampana:
    Err.Raise Err.Number, "Inserta una entrega de campana", Err.Description
End Function

Public Function InsertaDetalleEntregaCampana(ByVal nIdEntregaCampana As Integer, ByVal cBSCod As String, ByVal nIdCombo As String, ByVal nIdCampana As String, ByVal nCantidad As String, ByVal bCombo As Integer) As Integer
On Error GoTo ErrInsertaEntregaCombobienes:
    Dim RsSource  As ADODB.Recordset
    Set RsSource = New ADODB.Recordset
    Dim sql As String
    sql = "EXEC ERS0822014_stp_ins_detalleentregacampana " _
                           & nIdEntregaCampana & ", " _
                           & " '" & cBSCod & "', " _
                           & nIdCombo & ", " _
                           & nIdCampana & ", " _
                           & nCantidad & ", " _
                           & bCombo
   Set RsSource = oCon.CargaRecordSet(sql)
    InsertaDetalleEntregaCampana = RsSource!Cod
Exit Function
ErrInsertaEntregaCombobienes:
    Err.Raise Err.Number, "Inserta los detalles de la entregas", Err.Description
End Function

Public Function RecuperaComboBienesInserta(ByVal nIdCombo As String) As ADODB.Recordset
On Error GoTo ErrRecuperaComboBienes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sSql As String
    sSql = "Exec ERS0822014_stp_sel_comboBienes " & nIdCombo
    Set rs = oCon.CargaRecordSet(sSql)
    Set RecuperaComboBienesInserta = rs
    Exit Function
ErrRecuperaComboBienes:
    Err.Raise Err.Number, "Recupera los bienes por combos al insertar", Err.Description
End Function


Public Function AcualizaDetalleCombo(ByVal nIdCombbien As String, ByVal cBSCod As String, ByVal nCantidad As String) As Integer
On Error GoTo AcualizaDetalleCombo:
    Dim RsSource  As ADODB.Recordset
    Set RsSource = New ADODB.Recordset
    Dim sql As String
    sql = "EXEC ERS0822014_stp_upd_comboBienes " _
                            & nIdCombbien & ", " _
                            & " '" & cBSCod & "', " _
                            & nCantidad
   Set RsSource = oCon.CargaRecordSet(sql)
Exit Function
AcualizaDetalleCombo:
    Err.Raise Err.Number, "Insertar los detalles combo de campañas", Err.Description
End Function



