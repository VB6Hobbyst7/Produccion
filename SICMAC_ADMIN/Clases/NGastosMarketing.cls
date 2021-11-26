VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NGastosMarketing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oDGastosMarketing As New DGastosMarketing
Public Function FechaEstaEntreFechaActividad(ByVal pnActividad As Long, ByVal pdFecha As Date) As Boolean
    Dim oGasto As New DGastosMarketing
    Dim rs As New ADODB.Recordset
    
    FechaEstaEntreFechaActividad = False
    Set rs = oGasto.RecuperaActividadxId(pnActividad)
    If Not RSVacio(rs) Then
        If pdFecha >= rs!dFechaIni And pdFecha <= rs!dFechaFin Then
            FechaEstaEntreFechaActividad = True
        End If
    End If
    Set oGasto = Nothing
End Function
Public Sub InsertaProductoEnActividad(ByVal psAgeCod As String, ByVal pnActividadId As Long, ByVal pnProdServId As Long, ByVal pnCantidad As Long, ByVal pdFecha As Date, ByVal psComentario As String, ByRef psMsgErr As String)
    Dim oDGasto As New DGastosMarketing
    Dim rsCompra As New ADODB.Recordset
    Dim MatCompra() As Variant
    Dim i As Integer
    Dim lnCantidad As Long, lnStock As Long
    Dim lnIDNuevo As Long

    On Error GoTo ErrInsertaProductoEnActividad
    
    i = 0
    lnStock = 0
    lnIDNuevo = 0
    lnCantidad = pnCantidad
    Set rsCompra = oDGasto.RecuperaStockCompraxProducto(pnProdServId)
    
    If Not RSVacio(rsCompra) Then
        Do While Not rsCompra.EOF
            lnStock = rsCompra!nStock
            ReDim Preserve MatCompra(1, i)
            MatCompra(0, i) = rsCompra!nCompraId
            If lnStock <= lnCantidad Then
                MatCompra(1, i) = lnStock
                lnCantidad = lnCantidad - lnStock
            Else
                MatCompra(1, i) = lnCantidad
                lnCantidad = 0
            End If
            i = i + 1
            rsCompra.MoveNext
            If lnCantidad = 0 Then Exit Do
        Loop
        lnIDNuevo = oDGasto.InsertaProductoEnActividad(psAgeCod, pnActividadId, pnProdServId, pdFecha, psComentario)
        If lnIDNuevo <> 0 Then
            For i = 0 To UBound(MatCompra, 2)
                Call oDGasto.InsertaProductoEnActividadDetalle(lnIDNuevo, MatCompra(0, i), MatCompra(1, i))
            Next
        Else
            psMsgErr = "Ocurrio un Error al Agregar a la Lista, comuniquese con el Dpto. de TI"
        End If
    Else
        psMsgErr = "Ocurrio un Error al Agregar a la Lista, comuniquese con el Dpto. de TI"
    End If
    Exit Sub
ErrInsertaProductoEnActividad:
    psMsgErr = "Ocurrio un Error al Agregar a la Lista, comuniquese con el Dpto. de TI"
    Err.Raise Err.Number, "Inserta Actividad de Producto", Err.Description
End Sub




'VAPI SEGUN ERS 082-2014
Public Function getOcon() As DConecta
    Set getOcon = oDGastosMarketing.getOcon
End Function
Public Function getMaterialesPromocion() As ADODB.Recordset
    Set getMaterialesPromocion = oDGastosMarketing.RecuperaBienesServicios("11101%")
End Function
Public Function RecuperaCampanas() As ADODB.Recordset
    Set RecuperaCampanas = oDGastosMarketing.RecuperaCampanas
End Function

Public Function RecuperaCombosXCampana(ByVal nIdCampana As String) As ADODB.Recordset
    Set RecuperaCombosXCampana = oDGastosMarketing.RecuperaCombosXCampana(nIdCampana)
End Function
'Nota importante: la funcion requiere que se utilice previa llamada a  oDGastosMarketing.getOcon.AbreConexion
Public Function InsertaComboCampana(ByVal idCampana As Integer, ByVal cComboDescripcion As String, ByVal bDesembolso As Integer, ByVal bApertura As Integer, ByVal bSoles As Integer, ByVal bDolares As Integer, ByVal nMinSoles As String, ByVal nMaxSoles As String, ByVal nMinDolares As String, ByVal nMaxDolares As String) As Integer
    InsertaComboCampana = oDGastosMarketing.InsertaComboCampana(idCampana, cComboDescripcion, bDesembolso, bApertura, bSoles, bDolares, nMinSoles, nMaxSoles, nMinDolares, nMaxDolares)
End Function
'Nota importante: el procedimiento requiere que se utilice previa llamada a  oDGastosMarketing.getOcon.AbreConexion
Public Sub InsertaDetalleCombo(ByVal nIdCombo As String, ByVal cBSCod As String, ByVal nCantidad As String)
    Call oDGastosMarketing.InsertaDetalleCombo(nIdCombo, cBSCod, nCantidad)
End Sub

Public Function RecuperaComboBienes(ByVal nIdCombo As String) As ADODB.Recordset
    Set RecuperaComboBienes = oDGastosMarketing.RecuperaComboBienes(nIdCombo)
End Function

Public Sub EliminaBienEnCombo(ByVal nIdComboBienes As String)
    Call oDGastosMarketing.EliminaBienEnCombo(nIdComboBienes)
End Sub

Public Function InsertaDetalleComboCon(ByVal nIdCombo As String, ByVal cBSCod As String, ByVal nCantidad As String) As Integer
    oDGastosMarketing.getOcon.AbreConexion
        InsertaDetalleComboCon = oDGastosMarketing.InsertaDetalleCombo(nIdCombo, cBSCod, nCantidad)
    oDGastosMarketing.getOcon.CierraConexion
End Function

Public Sub ActualizaComboxCampana(ByVal nIdCombo, ByVal idCampana As Integer, ByVal cComboDescripcion As String, ByVal bDesembolso As Integer, ByVal bApertura As Integer, ByVal bSoles As Integer, ByVal bDolares As Integer, ByVal nMinSoles As String, ByVal nMaxSoles As String, ByVal nMinDolares As String, ByVal nMaxDolares As String)
    Call oDGastosMarketing.ActualizaComboxCampana(nIdCombo, idCampana, cComboDescripcion, bDesembolso, bApertura, bSoles, bDolares, nMinSoles, nMaxSoles, nMinDolares, nMaxDolares)
End Sub
Public Function ExisteDescripcionCombo(ByVal nIdCombo As String, ByVal nIdCampana As String, ByVal cComboDescripcion As String) As Integer
    ExisteDescripcionCombo = oDGastosMarketing.ExisteDescripcionCombo(nIdCombo, nIdCampana, cComboDescripcion)
End Function

Public Sub EliminaComboCampana(ByVal nIdCombo As String)
    oDGastosMarketing.EliminaComboCampana nIdCombo
End Sub

Public Function getMaterialesPromocionConSaldoXalmacen(ByVal nAlmacen As Integer) As ADODB.Recordset
    Set getMaterialesPromocionConSaldoXalmacen = oDGastosMarketing.RecuperaBienesServicios("11101%", 1, nAlmacen)
End Function
'Nota importante: la funcion requiere que se utilice previa llamada a  oDGastosMarketing.getOcon.AbreConexion
Public Function InsertaEntregaCampana(ByVal cAgencia As String, ByVal cPersCodUser As String, cPersCodUserSol As String, ByVal cPersCodCliente As String, ByVal fecha As String, ByVal cGlosa As String, ByVal sMovNro As String) As Integer
    
    InsertaEntregaCampana = oDGastosMarketing.InsertaEntregaCampana(cAgencia, cPersCodUser, cPersCodUserSol, cPersCodCliente, fecha, cGlosa, sMovNro)
End Function
'Nota importante: el procedimiento requiere que se utilice previa llamada a  oDGastosMarketing.getOcon.AbreConexion
Public Function InsertaDetalleEntregaCampana(ByVal nIdEntregaCampana As Integer, ByVal cBSCod As String, ByVal nIdCombo As String, ByVal nIdCampana As String, ByVal nCantidad As String, ByVal bCombo As Integer) As Integer
    InsertaDetalleEntregaCampana = oDGastosMarketing.InsertaDetalleEntregaCampana(nIdEntregaCampana, cBSCod, nIdCombo, nIdCampana, nCantidad, bCombo)
End Function

Public Function RecuperaComboBienesInserta(ByVal nIdCombo As String) As ADODB.Recordset
    Set RecuperaComboBienesInserta = oDGastosMarketing.RecuperaComboBienesInserta(nIdCombo)
End Function


Public Function ActualizaDetalleComboCon(ByVal nIdCombobien As String, ByVal cBSCod As String, ByVal nCantidad As String) As Integer
    oDGastosMarketing.getOcon.AbreConexion
        ActualizaDetalleComboCon = oDGastosMarketing.AcualizaDetalleCombo(nIdCombobien, cBSCod, nCantidad)
    oDGastosMarketing.getOcon.CierraConexion
End Function
