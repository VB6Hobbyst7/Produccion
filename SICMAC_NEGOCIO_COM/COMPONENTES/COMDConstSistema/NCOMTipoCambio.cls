VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim vsFormatoFecha As String
Dim vsConexion As String
Dim vsServerCom As String
Dim vsServerPers As String
Dim oCon As New COMConecta.DCOMConecta
Dim oerror As New COMConecta.COMErrorHandling

Public Sub inicio(ByVal psFormatoFecha As String)
    vsFormatoFecha = psFormatoFecha
End Sub

Private Sub Class_Initialize()
Dim oIni As New COMConecta.DCOMClasIni

If oCon.AbreConexion = False Then
    oerror.RaiseError oerror.MyUnhandledError, "DMov:Initialize Method. Error en Conexion a Base de datos"
End If
vsConexion = oIni.CadenaConexion
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing
End Sub
'ALPA 20081003*******************************************************************************
'Se agrego el parametro nValPondREU para el REU
'********************************************************************************************
Public Function CargaTipoCambio() As ADODB.Recordset
    'Carga en un Data los Registros existentes en la Tabla Tipo Cambio
    Dim RegTC As New ADODB.Recordset
    Dim lnDiasTipoCambio As Integer
    Dim oConsSit As NCOMConstSistema
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    
    Set oConsSit = New NCOMConstSistema
    lnDiasTipoCambio = 0
    lnDiasTipoCambio = oConsSit.LeeConstSistema(gConstSistDiasTipoCambioMant)
    
    sql = "  SET ROWCOUNT " & lnDiasTipoCambio & "  " _
        & "  Select Convert(VarChar(12), dFecCamb,103)  as  Fecha,  nValVent as Venta, " _
        & "         nValComp as Compra, nValVentEsp VentaE, nValCompEsp CompraE , nValFijoDia as FijoDia, nValFijo as Fijo, nValPond Ponderado, dFecCamb,nValPondVenta,nValPondREU,nValSBSDia,nValCompTr,nValVentTr" _
        & "  From " & vsServerCom & "TipoCambio Order BY dFecCamb Desc " _
        & " SET ROWCOUNT 0 "
    If oCon.AbreConexion = False Then Exit Function
    Set rs = oCon.CargaRecordSet(sql)
    Set CargaTipoCambio = rs
    Set oConsSit = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ExisteTipoCambio(ByVal pdFecha As String) As Boolean
'Verifica si el Tipo de Cambio para la fecha enviada como parámetro ya ha sido ingresada
Dim rs As New ADODB.Recordset
Dim sql As String

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
'ALPA 20140317****************
'Se agregó tres tipos de cambio: nValSBSDia=isnull(nValSBSDia,0),nValCompTr=isnull(nValCompTr,0),nValVentTr=isnull(nValVentTr,0)
'ALPA 20081003*******************************************************************************
'Se agrego el parametro nValPondREU para el REU
'********************************************************************************************
Public Function EmiteTipoCambio(ByVal dFecha As Date, ByVal nTpoTipoCambio As TipoCambio) As Double
    Dim sql As String
    Dim rs As New ADODB.Recordset

    EmiteTipoCambio = 0

    If oCon.AbreConexion = False Then Exit Function
    rs.CursorLocation = adUseClient
    sql = "Select isnull(nValPondREU,0) nValPondREU,nValFijo, nValFijoDia, nValVent, nValComp, nValVentEsp, nValCompEsp, nValPond,nValPondVenta,nValSBSDia=isnull(nValSBSDia,0),nValCompTr=isnull(nValCompTr,0),nValVentTr=isnull(nValVentTr,0) From TipoCambio " _
        & " WHERE dFecCamb = (   Select Max(dFecCamb)" _
        & "                      From TipoCambio " _
        & "                      Where datediff(day,dFecCamb,'" & Format$(dFecha, "mm/dd/yyyy") & "')=0)"
  
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
            Case TCVentaEsp
                EmiteTipoCambio = rs("nValVentEsp")
            Case TCCompraEsp
                EmiteTipoCambio = rs("nValCompEsp")
            Case TCPonderado
                EmiteTipoCambio = rs("nValPond")
            Case TCPondVenta
                EmiteTipoCambio = IIf(IsNull(rs("nValPondVenta")), 0, rs("nValPondVenta"))
            Case TCPondREU
                EmiteTipoCambio = rs("nValPondREU")
            'ALPA20140226*************************************
            Case TCSBSDia
                EmiteTipoCambio = rs("nValSBSDia")
            Case TCCompraTraider
                EmiteTipoCambio = rs("nValCompTr")
            Case TCVentaTraider
                EmiteTipoCambio = rs("nValVentTr")
            '*************************************************
        End Select
    End If
    rs.Close
    Set rs = Nothing
End Function
'ALPA 20081003*******************************************************************************
'Se agrego el parametro pnPonREU para el REU
'********************************************************************************************
Public Function GrabaTipoCambio(ByVal psFormatoFecha As String, ByVal pdFecCamb As Date, ByVal pnValVent As Currency, ByVal pnValComp As Currency, _
                                    ByVal pnValVentEsp As Currency, ByVal pnValCompEsp As Currency, _
                                    ByVal pnValFijo As Currency, ByVal pnValFijoDia As Currency, ByVal pnValPonderado As Currency, _
                                    ByVal psUltimaActualizacion As String, _
                                    Optional ByVal pbEjectBatch As Boolean, Optional ByVal pnPondVenta As Currency, Optional ByVal pnPonREU As Currency = 0)
    Dim oTipoCambio As DCOMTipoCambio
    Set oTipoCambio = New DCOMTipoCambio
    oTipoCambio.inicio psFormatoFecha
    oTipoCambio.ActualizaTipoCambioDiario pdFecCamb, psUltimaActualizacion, pnValFijoDia, pbEjectBatch
    oTipoCambio.InsertaTipoCambio pdFecCamb, pnValVent, pnValComp, pnValVentEsp, pnValCompEsp, pnValFijo, pnValFijoDia, pnValPonderado, psUltimaActualizacion, pbEjectBatch, pnPondVenta, pnPonREU
    
    If pbEjectBatch Then
        GrabaTipoCambio = oTipoCambio.EjecutaBatch
    Else
        GrabaTipoCambio = 0
    End If
End Function
'ALPA 20081003*******************************************************************************
'Se agrego el parametro pnPonREU para el REU
'********************************************************************************************
Public Function ActualizaTipoCambio(ByVal psFormatoFecha As String, ByVal pdFecCamb As Date, ByVal pnValVent As Currency, ByVal pnValComp As Currency, _
                                    ByVal pnValVentEnt As Currency, ByVal pnValCompEst As Currency, _
                                    ByVal pnValFijo As Currency, ByVal pnValFijoDia As Currency, ByVal pnValPonderado As Currency, ByVal psUltimaActualizacion As String, _
                                    Optional ByVal pbEjectBatch As Boolean, Optional ByVal pnValPondVenta As Currency, Optional ByVal pnPonREU As Currency = 0)
Dim oTipoCambio As DCOMTipoCambio
Set oTipoCambio = New DCOMTipoCambio

oTipoCambio.inicio psFormatoFecha
oTipoCambio.ActualizaTipoCambio pdFecCamb, psUltimaActualizacion, pnValVent, pnValComp, pnValVentEnt, pnValCompEst, pnValFijo, pnValFijoDia, pnValPonderado, pbEjectBatch, pnValPondVenta, pnPonREU

If pbEjectBatch Then
    ActualizaTipoCambio = oTipoCambio.EjecutaBatch
Else
    ActualizaTipoCambio = 0
End If
End Function




