VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NInvActivoFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function DarOrdenCompraXAF(ByVal sDesde As String, ByVal sHasta As String, ByVal sMoneda As String) As ADODB.Recordset
    Dim oInventario As DInvActivoFijo
    Set oInventario = New DInvActivoFijo
    Set DarOrdenCompraXAF = oInventario.ListarOrdenCompraXAF(sDesde, sHasta, sMoneda)
End Function

Public Function DarBBSXOC(ByVal sMvoNro As String) As ADODB.Recordset
    Dim oInventario As DInvActivoFijo
    Set oInventario = New DInvActivoFijo
    Set DarBBSXOC = oInventario.ListarBBSXOC(sMvoNro)
End Function

Public Function DarDatosBBS(ByVal sMvoNro As String, ByVal sBSCod As String, ByVal iMovItem As Integer) As ADODB.Recordset
    Dim oInventario As DInvActivoFijo
    Set oInventario = New DInvActivoFijo
    Set DarDatosBBS = oInventario.DevolverDatosBBS(sMvoNro, sBSCod, iMovItem)
End Function

Public Function DarAgencia(ByVal sMvoNro As String) As ADODB.Recordset
    Dim oInventario As DInvActivoFijo
    Set oInventario = New DInvActivoFijo
    Set DarAgencia = oInventario.DevolverAgencia(sMvoNro)
End Function

Public Function DarCorrelativo(ByVal sBSCod As String) As ADODB.Recordset
    Dim oInventario As DInvActivoFijo
    Set oInventario = New DInvActivoFijo
    Set DarCorrelativo = oInventario.DevolverCorrelativo(sBSCod)
End Function

'JIPR20200313 INICIO
Public Function DarCorrelativoAF(ByVal sBSCod As String, Optional sAgencia As String, Optional nTipo As String) As ADODB.Recordset
    Dim oInventario As DInvActivoFijo
    Set oInventario = New DInvActivoFijo
    Set DarCorrelativoAF = oInventario.DevolverCorrelativoAF(sBSCod, sAgencia, nTipo)
End Function
'JIPR20200313 FIN

Public Function DarTipoBien(ByVal sBSCod As String) As ADODB.Recordset
    Dim oInventario As DInvActivoFijo
    Set oInventario = New DInvActivoFijo
    Set DarTipoBien = oInventario.DevolverTipoBien(sBSCod)
End Function

Public Function DarActivadoXBBS(ByVal sMvoNro As String, ByVal sBSCod As String, ByVal iMovItem As Integer) As ADODB.Recordset
    Dim oInventario As DInvActivoFijo
    Set oInventario = New DInvActivoFijo
    Set DarActivadoXBBS = oInventario.DevolverActivadoXBBS(sMvoNro, sBSCod, iMovItem)
End Function

Public Sub InsertarAF(ByVal sCodInventario As String, ByVal sNombre As String, ByVal sAgencia As String, ByVal sMarca As String, ByVal sModelo As String, ByVal sSerie As String, ByVal sFIngreso As String, ByVal sMovNro As String, ByVal sBSCod As String, ByVal iMovItem As Integer)
    Dim oInventario As DInvActivoFijo
    Set oInventario = New DInvActivoFijo
    oInventario.RegistrarAF sCodInventario, sNombre, sAgencia, sMarca, sModelo, sSerie, sFIngreso, sMovNro, sBSCod, iMovItem
End Sub

Public Function ObtenerTipoAF() As ADODB.Recordset
    Dim oInventario As DInvActivoFijo
    Set oInventario = New DInvActivoFijo
    Set ObtenerTipoAF = oInventario.CargarTipoAF
End Function

Public Function ObtenerBienXTipo(ByVal sBSCod As String) As ADODB.Recordset
    Dim oInventario As DInvActivoFijo
    Set oInventario = New DInvActivoFijo
    Set ObtenerBienXTipo = oInventario.CargarBienXTipo(sBSCod)
End Function

Public Function ObtenerReporteAF(ByVal sAgeCod As String, ByVal sBSCod As String, ByVal sBSCodBien As String, ByVal sFecha As String, ByVal psCategoBien As String, ByVal psDel As String, ByVal psAl As String, Optional ByVal psSerie As String = "") As ADODB.Recordset
    '*** PEAC 20100507 - SE AGREGO PARAMETRO (psCategoBien)
    '*** peac 20120425 -SE AGREGO PARAMETRO (psSerie)
    Dim oInventario As DInvActivoFijo
    Set oInventario = New DInvActivoFijo
    Set ObtenerReporteAF = oInventario.CargarReporteAF(sAgeCod, sBSCod, sBSCodBien, sFecha, psCategoBien, psDel, psAl, psSerie)
End Function

Public Function ObtenerBienes(ByVal sCodInv As String, ByVal sFI As String, ByVal sFF As String, ByVal sDescripcion As String, ByVal sFDepre As String) As ADODB.Recordset
    Dim oInventario As DInvActivoFijo
    Set oInventario = New DInvActivoFijo
    Set ObtenerBienes = oInventario.DevolverBienes(sCodInv, sFI, sFF, sDescripcion, sFDepre)
End Function


'*** PEAC 20100511
Public Function DarCorrelativoNoDepre(ByVal sBSCod As String) As ADODB.Recordset
    Dim oInventario As DInvActivoFijo
    Set oInventario = New DInvActivoFijo
    Set DarCorrelativoNoDepre = oInventario.DevolverCorrelativoNoDepre(sBSCod)
End Function

