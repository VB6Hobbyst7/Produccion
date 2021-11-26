VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'Public Function DarOrdenCompraXAF(ByVal sDesde As String, ByVal sHasta As String, ByVal sMoneda As String) As ADODB.Recordset
'    Dim oInventario As DCOMInvActivoFijo
'    Set oInventario = New DCOMInvActivoFijo
'    Set DarOrdenCompraXAF = oInventario.ListarOrdenCompraXAF(sDesde, sHasta, sMoneda)
'End Function

'Public Function DarBBSXOC(ByVal sMvoNro As String) As ADODB.Recordset
'    Dim oInventario As DCOMInvActivoFijo
'    Set oInventario = New DCOMInvActivoFijo
'    Set DarBBSXOC = oInventario.ListarBBSXOC(sMvoNro)
'End Function

'Public Function DarDatosBBS(ByVal sMvoNro As String, ByVal sBSCod As String, ByVal iMovItem As Integer) As ADODB.Recordset
'    Dim oInventario As DCOMInvActivoFijo
'    Set oInventario = New DCOMInvActivoFijo
'    Set DarDatosBBS = oInventario.DevolverDatosBBS(sMvoNro, sBSCod, iMovItem)
'End Function
'
'Public Function DarAgencia(ByVal sMvoNro As String) As ADODB.Recordset
'    Dim oInventario As DCOMInvActivoFijo
'    Set oInventario = New DCOMInvActivoFijo
'    Set DarAgencia = oInventario.DevolverAgencia(sMvoNro)
'End Function
'
'Public Function DarCorrelativo(ByVal sBSCod As String) As ADODB.Recordset
'    Dim oInventario As DCOMInvActivoFijo
'    Set oInventario = New DCOMInvActivoFijo
'    Set DarCorrelativo = oInventario.DevolverCorrelativo(sBSCod)
'End Function
'
'Public Function DarTipoBien(ByVal sBSCod As String) As ADODB.Recordset
'    Dim oInventario As DCOMInvActivoFijo
'    Set oInventario = New DCOMInvActivoFijo
'    Set DarTipoBien = oInventario.DevolverTipoBien(sBSCod)
'End Function

'Public Function DarActivadoXBBS(ByVal sMvoNro As String, ByVal sBSCod As String, ByVal iMovItem As Integer) As ADODB.Recordset
'    Dim oInventario As DCOMInvActivoFijo
'    Set oInventario = New DCOMInvActivoFijo
'    Set DarActivadoXBBS = oInventario.DevolverActivadoXBBS(sMvoNro, sBSCod, iMovItem)
'End Function
'
'Public Sub InsertarAF(ByVal sCodInventario As String, ByVal sNombre As String, ByVal sAgencia As String, ByVal sMarca As String, ByVal sModelo As String, ByVal sSerie As String, ByVal sFIngreso As String, ByVal sMovNro As String, ByVal sBSCod As String, ByVal iMovItem As Integer)
'    Dim oInventario As DCOMInvActivoFijo
'    Set oInventario = New DCOMInvActivoFijo
'    oInventario.RegistrarAF sCodInventario, sNombre, sAgencia, sMarca, sModelo, sSerie, sFIngreso, sMovNro, sBSCod, iMovItem
'End Sub

Public Function ObtenerTipoAF() As ADODB.Recordset
    Dim oInventario As COMDAuditoria.DCOMAF
    Set oInventario = New COMDAuditoria.DCOMAF
    Set ObtenerTipoAF = oInventario.CargarTipoAF
End Function

Public Function ObtenerBienXTipo(ByVal sBSCod As String) As ADODB.Recordset
    Dim oInventario As COMDAuditoria.DCOMAF
    Set oInventario = New COMDAuditoria.DCOMAF
    Set ObtenerBienXTipo = oInventario.CargarBienXTipo(sBSCod)
End Function

Public Function ObtenerReporteAF(ByVal sAgeCod As String, ByVal sBSCod As String, ByVal sBSCodBien As String, ByVal sFecha As String) As ADODB.Recordset
    Dim oInventario As COMDAuditoria.DCOMAF
    Set oInventario = New COMDAuditoria.DCOMAF
    Set ObtenerReporteAF = oInventario.CargarReporteAF(sAgeCod, sBSCod, sBSCodBien, sFecha)
End Function

'Public Function ObtenerBienes(ByVal sCodInv As String, ByVal sFI As String, ByVal sFF As String, ByVal sDescripcion As String, ByVal sFDepre As String) As ADODB.Recordset
'    Dim oInventario As DCOMInvActivoFijo
'    Set oInventario = New DCOMInvActivoFijo
'    Set ObtenerBienes = oInventario.DevolverBienes(sCodInv, sFI, sFF, sDescripcion, sFDepre)
'End Function

'Log
Public Function ObtenerReporteLog(ByVal sAgeCod As String, ByVal sDesde As String, ByVal sHasta As String, ByVal sUsuario As String) As ADODB.Recordset
    Dim oInventario As COMDAuditoria.DCOMAF
    Set oInventario = New COMDAuditoria.DCOMAF
    Set ObtenerReporteLog = oInventario.CargarReporteLog(sAgeCod, sDesde, sHasta, sUsuario)
End Function



