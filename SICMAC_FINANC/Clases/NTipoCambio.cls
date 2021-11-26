VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsConexion As String
Dim vsServerCom As String
Dim vsServerPers As String

'Public Function CargaTipoCambio() As adodb.Recordset
'    'Carga en un Data los Registros existentes en la Tabla Tipo Cambio
'    Dim RegTC As New adodb.Recordset
'    Dim lnDiasTipoCambio As Integer
'    Dim oConsSit As NConstSistemas
'    Dim sql As String
'    Dim Rs As adodb.Recordset
'    Dim oConect As DConecta
'
'    Set oConsSit = New NConstSistemas
'    Set oConect = New DConecta
'    lnDiasTipoCambio = 0
'    lnDiasTipoCambio = oConsSit.LeeConstSistema(gConstSistDiasTipoCambioMant)
'
'    sql = "  SET ROWCOUNT " & lnDiasTipoCambio & "  " _
'        & "  Select Convert(VarChar(12), dFecCamb,103)  as  Fecha,  nValVent as Venta, " _
'        & "         nValComp as Compra, nValVentEsp VentaE, nValCompEsp CompraE , nValFijoDia as FijoDia, nValFijo as Fijo, nValPond Ponderado, dFecCamb  " _
'        & "  From " & vsServerCom & "TipoCambio Order BY dFecCamb Desc " _
'        & " SET ROWCOUNT 0 "
'    If oConect.AbreConexion = False Then Exit Function
'    Set Rs = oConect.CargaRecordSet(sql)
'    Set CargaTipoCambio = Rs
'    Set oConsSit = Nothing
'    oConect.CierraConexion
'    Set oConect = Nothing
'End Function

Public Function CargaTipoCambio() As ADODB.Recordset
    'Carga en un Data los Registros existentes en la Tabla Tipo Cambio
    Dim RegTC As New ADODB.Recordset
    Dim lnDiasTipoCambio As Integer
    Dim oConsSit As NConstSistemas
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oConect As DConecta
    
    Set oConsSit = New NConstSistemas
    Set oConect = New DConecta
    lnDiasTipoCambio = 0
    lnDiasTipoCambio = oConsSit.LeeConstSistema(gConstSistDiasTipoCambioMant)
    
    sql = "  SET ROWCOUNT " & lnDiasTipoCambio & "  " _
        & "  Select Convert(VarChar(12), dFecCamb,103)  as  Fecha,  nValVent as Venta, " _
        & "         nValComp as Compra, nValVentEsp VentaE, nValCompEsp CompraE , nValFijoDia as FijoDia, nValFijo as Fijo, nValPond Ponderado, dFecCamb,Isnull(nValPondVenta,0) PonderadoVenta  " _
        & "  From " & vsServerCom & "TipoCambio Order BY dFecCamb Desc " _
        & " SET ROWCOUNT 0 "
    If oConect.AbreConexion = False Then Exit Function
    Set rs = oConect.CargaRecordSet(sql)
    Set CargaTipoCambio = rs
    Set oConsSit = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Private Sub Class_Initialize()
    Dim oIni As ClasIni
    Set oIni = New ClasIni
    vsConexion = oIni.CadenaConexion
    vsServerCom = oIni.BaseComunes
    vsServerPers = oIni.BasePersonas
    Set oIni = Nothing
End Sub
Public Function ExisteTipoCambio(ByVal pdFecha As String) As Boolean
'Verifica si el Tipo de Cambio para la fecha enviada como parámetro ya ha sido ingresada
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
Dim sql As String

Set oCon = New DConecta

If oCon.AbreConexion = False Then Exit Function
sql = "Select dFecCamb from " & vsServerCom & "TipoCambio Where DateDiff(day, dFecCamb, '" & pdFecha & "') = 0"

Set rs = oCon.CargaRecordSet(sql)
If rs.EOF And rs.BOF Then
   ExisteTipoCambio = False
Else
   ExisteTipoCambio = True
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function EmiteTipoCambio(ByVal dFecha As Date, ByVal nTpoTipoCambio As Integer) As Double
    Dim sql As String
    Dim rs As Recordset
    Dim lbCentral As Boolean
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    EmiteTipoCambio = 0
    Set rs = New Recordset
    
    If oCon.AbreConexion = False Then Exit Function
    rs.CursorLocation = adUseClient
    
    sql = " Select nConsSisValor From ConstSistema where nConsSisCod = " & gConstSistBitCentral
    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        lbCentral = True
    Else
       If rs.Fields(0) = 1 Then
           lbCentral = True
       Else
           lbCentral = False
       End If
    End If
    
    rs.Close
    
    If lbCentral Then
        'ALPA 20140331************************************
        'Se agregó tipos de cambio nValSBSDia,nValCompTr,nValVentTr
        'Centralizado
        sql = "Select IsNull(nValFijo,0) nValFijo, IsNull(nValFijoDia,0) nValFijoDia, IsNull(nValVent,0) nValVent, IsNull(nValComp,0) nValComp, IsNull(nValVentEsp,0) nValVentEsp, IsNull(nValCompEsp,0) nValCompEsp, IsNull(nValPond,0) nValPond, IsNull(nValPondVenta,0) nValPondVenta,nValSBSDia=isnull(nValSBSDia,0),nValCompTr=isnull(nValCompTr,0),nValVentTr=isnull(nValVentTr,0) From TipoCambio " _
            & " WHERE dFecCamb = (   Select Max(dFecCamb)" _
            & "                      From TipoCambio " _
            & "                      Where datediff(day,dFecCamb,'" & Format$(dFecha, gsFormatoFecha) & "')=0)"
' sql = "Select IsNull(nValFijo,0) nValFijo, IsNull(nValFijoDia,0) nValFijoDia, IsNull(nValVent,0) nValVent, IsNull(nValComp,0) nValComp, IsNull(nValVentEsp,0) nValVentEsp, IsNull(nValCompEsp,0) nValCompEsp, IsNull(nValPond,0) nValPond, IsNull(nValPondVenta,0) nValPondVenta From TipoCambio " _
'            & " WHERE dFecCamb = (   Select Max(dFecCamb)" _
'            & "                      From TipoCambio " _
'            & "                      Where datediff(day,dFecCamb,'" & Format$(dFecha, gsFormatoFecha) & "')=0)"
'

    Else
        'Distribuido
        sql = "Select nValFijo, nValFijoDia, nValVent, nValComp From DbComunes..TipCambio " _
           & " WHERE dFecCamb = (   Select Max(dFecCamb)" _
           & "                      From DbComunes..TipCambio " _
           & "                      Where datediff(day,dFecCamb,'" & Format$(dFecha, gsFormatoFecha) & "')=0)"
    End If
    
    Set rs = oCon.CargaRecordSet(sql)
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF And Not rs.BOF Then
        Select Case nTpoTipoCambio
            Case TCFijoMes
                EmiteTipoCambio = rs("nValFijo")
            Case TCFijoDia
                EmiteTipoCambio = rs("nValFijoDia")
            Case TCVenta
                EmiteTipoCambio = rs("nValVent")
            Case TCCompra
                EmiteTipoCambio = rs("nValComp")
            Case TCVentaEsp And lbCentral
                EmiteTipoCambio = rs("nValVentEsp")
            Case TCCompraEsp And lbCentral
                EmiteTipoCambio = rs("nValCompEsp")
            Case TCPonderado And lbCentral
                EmiteTipoCambio = rs("nValPond")
            Case 7 And lbCentral
                EmiteTipoCambio = rs("nValPondVenta")
            'ALPA20140226*************************************
            Case 9
                EmiteTipoCambio = rs("nValSBSDia")
            Case 10
                EmiteTipoCambio = rs("nValCompTr")
            Case 11
                EmiteTipoCambio = rs("nValVentTr")
            '*************************************************
            Case Else
                EmiteTipoCambio = 0
        End Select
    End If
    rs.Close
    Set rs = Nothing
End Function

'Public Function GrabaTipoCambio(ByVal psFormatoFecha As String, ByVal pdFecCamb As Date, ByVal pnValVent As Currency, ByVal pnValComp As Currency, _
'                                    ByVal pnValVentEsp As Currency, ByVal pnValCompEsp As Currency, _
'                                    ByVal pnValFijo As Currency, ByVal pnValFijoDia As Currency, ByVal pnValPonderado As Currency, _
'                                    ByVal psUltimaActualizacion As String, _
'                                    Optional ByVal pbEjectBatch As Boolean)
'    Dim oTipoCambio As dTipoCambio
'    Set oTipoCambio = New dTipoCambio
'    oTipoCambio.ActualizaTipoCambioDiario pdFecCamb, psUltimaActualizacion, pnValFijoDia, pbEjectBatch
'    oTipoCambio.InsertaTipoCambio pdFecCamb, pnValVent, pnValComp, pnValVentEsp, pnValCompEsp, pnValFijo, pnValFijoDia, pnValPonderado, psUltimaActualizacion, pbEjectBatch
'
'    If pbEjectBatch Then
'        GrabaTipoCambio = oTipoCambio.EjecutaBatch
'    Else
'        GrabaTipoCambio = 0
'    End If
'End Function

Public Function GrabaTipoCambio(ByVal psFormatoFecha As String, ByVal pdFecCamb As Date, ByVal pnValVent As Currency, ByVal pnValComp As Currency, _
                                    ByVal pnValVentEsp As Currency, ByVal pnValCompEsp As Currency, _
                                    ByVal pnValFijo As Currency, ByVal pnValFijoDia As Currency, ByVal pnValPonderado As Currency, _
                                    ByVal psUltimaActualizacion As String, ByVal pnValPonderadoVenta As Currency, _
                                    Optional ByVal pbEjectBatch As Boolean)
    Dim oTipoCambio As dTipoCambio
    Set oTipoCambio = New dTipoCambio
    oTipoCambio.ActualizaTipoCambioDiario pdFecCamb, psUltimaActualizacion, pnValFijoDia, pbEjectBatch
    oTipoCambio.InsertaTipoCambio pdFecCamb, pnValVent, pnValComp, pnValVentEsp, pnValCompEsp, pnValFijo, pnValFijoDia, pnValPonderado, psUltimaActualizacion, pnValPonderadoVenta, pbEjectBatch
    
    If pbEjectBatch Then
        GrabaTipoCambio = oTipoCambio.EjecutaBatch
    Else
        GrabaTipoCambio = 0
    End If
End Function


'Public Function ActualizaTipoCambio(ByVal psFormatoFecha As String, ByVal pdFecCamb As Date, ByVal pnValVent As Currency, ByVal pnValComp As Currency, _
'                                    ByVal pnValVentEnt As Currency, ByVal pnValCompEst As Currency, _
'                                    ByVal pnValFijo As Currency, ByVal pnValFijoDia As Currency, ByVal pnValPonderado As Currency, ByVal psUltimaActualizacion As String, _
'                                    Optional ByVal pbEjectBatch As Boolean)
'Dim oTipoCambio As dTipoCambio
'Set oTipoCambio = New dTipoCambio
'
'oTipoCambio.ActualizaTipoCambio pdFecCamb, psUltimaActualizacion, pnValVent, pnValComp, pnValVentEnt, pnValCompEst, pnValFijo, pnValFijoDia, pnValPonderado, pbEjectBatch
'
'If pbEjectBatch Then
'    ActualizaTipoCambio = oTipoCambio.EjecutaBatch
'Else
'    ActualizaTipoCambio = 0
'End If
'
'End Function

Public Function ActualizaTipoCambio(ByVal psFormatoFecha As String, ByVal pdFecCamb As Date, ByVal pnValVent As Currency, ByVal pnValComp As Currency, _
                                    ByVal pnValVentEnt As Currency, ByVal pnValCompEst As Currency, _
                                    ByVal pnValFijo As Currency, ByVal pnValFijoDia As Currency, ByVal pnValPonderado As Currency, ByVal psUltimaActualizacion As String, _
                                    ByVal pnValPonderadoVenta As Currency, Optional ByVal pbEjectBatch As Boolean)
Dim oTipoCambio As dTipoCambio
Set oTipoCambio = New dTipoCambio

oTipoCambio.ActualizaTipoCambio pdFecCamb, psUltimaActualizacion, pnValVent, pnValComp, pnValVentEnt, pnValCompEst, pnValFijo, pnValFijoDia, pnValPonderado, pbEjectBatch, pnValPonderadoVenta

If pbEjectBatch Then
    ActualizaTipoCambio = oTipoCambio.EjecutaBatch
Else
    ActualizaTipoCambio = 0
End If

End Function


'funcion que se incremento en edpyme
Public Function SelecTipoCambio(ByVal dFecha As Date) As ADODB.Recordset
    Dim sql As String
    Dim rs As Recordset
    Dim oCon As DConecta
    
    Set oCon = New DConecta
 
    If oCon.AbreConexion = False Then Exit Function
     sql = " Cnt_SelTpoCambio_sp '" & Format$(dFecha, gsFormatoFecha) & "'"
    Set rs = oCon.CargaRecordSet(sql)
    Set SelecTipoCambio = rs
     oCon.CierraConexion
    Set oCon = Nothing
End Function

'ALPA 20140331**************************************************************
Public Function ObtenerNivelesAprobacionCompraVentaxMonto(ByVal pnMonto As Currency, ByVal pnDifTCambio As Currency, ByVal pnTipo As Integer, ByRef pnNivel As String, ByVal nTipoEspecial As Integer) As Integer
'pnTipoCambio           =       1--Compra
'pnTipoCambio           =       2--Venta
    Dim oRs As ADODB.Recordset
    Dim nValidado As Integer
    nValidado = 0
    Set oRs = New ADODB.Recordset
    Dim oNivApr As dTipoCambio
    Set oNivApr = New dTipoCambio
    Set oRs = oNivApr.ObtenerNivelesAprobacionCompraVentaxMonto(pnMonto, nTipoEspecial)
    If Not (oRs.BOF Or oRs.EOF) Then
        Do While Not oRs.EOF
            If pnTipo = 1 Then
                If pnDifTCambio > 0 Then
                    If Abs(pnDifTCambio) <= oRs!nTCCmas Then
                    nValidado = 1
                    pnNivel = oRs!cNivelCod
                    Exit Do
                    End If
                End If
            Else
                If pnDifTCambio < 0 Then
                    If Abs(pnDifTCambio) <= oRs!nTCVmas Then
                    nValidado = 1
                    pnNivel = oRs!cNivelCod
                    Exit Do
                    End If
                End If
            End If
            pnNivel = oRs!cNivelCod
        oRs.MoveNext
        Loop
    End If
    ObtenerNivelesAprobacionCompraVentaxMonto = nValidado
End Function

Public Function ObtenerCantidadAprobacionMovCompraVenta(ByVal psMovNro As String, ByRef nEstado As Integer) As Integer
    Dim oNivApr As dTipoCambio
    Set oNivApr = New dTipoCambio
    Dim oRs As ADODB.Recordset
    Set oRs = oNivApr.ObtenerCantidadAprobacionMovCompraVenta(psMovNro)
    If Not (oRs.BOF Or oRs.EOF) Then
        nEstado = oRs!nEstado
        ObtenerCantidadAprobacionMovCompraVenta = oRs!nNroFirmas - oRs!nContador
    End If
    Set oRs = Nothing
End Function

'***************************************************************************
'****
Public Function ActualizarAprobacionMovCompraVenta(ByVal psMovNro As String)
    Dim lsSql As String
    Dim oConec As DConecta

    lsSql = "exec stp_upd_AprobacionMovCompraVenta '" & psMovNro & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    oConec.Ejecutar (lsSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerAprobacionMovCompraVentaPendiente(ByVal cFecha As String, ByVal psRHCargosCod As String) As ADODB.Recordset
    Dim lsSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    

    lsSql = "exec stp_sel_MostrarAprobacionMovCompraVentaPendiente '" & cFecha & "','" & psRHCargosCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerAprobacionMovCompraVentaPendiente = oConec.CargaRecordSet(lsSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Sub AprobacionMovCompraVentaDetalle(ByVal cMovNro As String, ByVal cOpeCod As String, ByVal cNivelCod As String, ByVal cMovNroAprobacion As String, ByVal cRHCargosCod As String, ByVal nEstado As Integer)
    Dim lsSql As String
    Dim oConec As COMConecta.DCOMConecta
    Set oConec = New COMConecta.DCOMConecta
    lsSql = "exec stp_ins_AprobacionMovCompraVentaDetalle '" & cMovNro & "','" & cOpeCod & "','" & cNivelCod & "','" & cMovNroAprobacion & "','" & cRHCargosCod & "'," & nEstado
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.Ejecutar (lsSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub

Public Function ObtenerAprobacionMovCompraVentaPendientexCliente(ByVal cFecha As String, Optional ByVal psPersCod As String = "", Optional ByVal lsOpeCod As String = "") As ADODB.Recordset
    Dim lsSql As String
    Dim oConec As DConecta 'COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    

    lsSql = "exec stp_sel_MostrarAprobacionMovCompraVentaPendientexCliente '" & cFecha & "','" & psPersCod & "','" & lsOpeCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ObtenerAprobacionMovCompraVentaPendientexCliente = oConec.CargaRecordSet(lsSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function AprobacionMovCompraVentaxMovimiento(ByVal cMovNro As String) As ADODB.Recordset
    Dim lsSql As String
    Dim oConec As DConecta
    Dim lrDatos As ADODB.Recordset
    

    lsSql = "exec stp_sel_AprobacionMovCompraVentaxMovimiento '" & cMovNro & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set AprobacionMovCompraVentaxMovimiento = oConec.CargaRecordSet(lsSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

