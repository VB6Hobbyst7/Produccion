VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NInvTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function InsertarTransferencia(ByVal iTipoTransferencia As Integer, ByVal sFecha As String, ByVal sOrigen As String, ByVal sDestino As String, ByVal sOrigenCod As String, ByVal sDestinoCod As String) As Integer
    Dim oInventario As DInvTransferencia
    Set oInventario = New DInvTransferencia
    oInventario.RegistrarTransferencia iTipoTransferencia, sFecha, sOrigen, sDestino, sOrigenCod, sDestinoCod
    InsertarTransferencia = oInventario.MaximoIdTransferencia
End Function

Public Sub InsertarInventarioTransferencia(ByVal iTransferenciaId As Integer, ByVal cBSCod As String)
    Dim oInventario As DInvTransferencia
    Set oInventario = New DInvTransferencia
    oInventario.RegistrarInventarioTransferencia iTransferenciaId, cBSCod
End Sub

Public Sub ActualizarBS(ByVal sMovNro As String, ByVal cBSCod As String, ByVal sAreCod As String, ByVal sAgeCod As String, ByVal cCodBSNvo As String)
    Dim oInventario As DInvTransferencia
    Set oInventario = New DInvTransferencia
    oInventario.ModificarBS sMovNro, cBSCod, sAreCod, sAgeCod, cCodBSNvo
End Sub

'Reporte de Transferencia
Public Function ObtenerReporteTransferencia(ByVal sFI As String, ByVal sFF As String) As ADODB.Recordset
    Dim oInventario As DInvTransferencia
    Set oInventario = New DInvTransferencia
    Set ObtenerReporteTransferencia = oInventario.CargarReporteTransferencia(sFI, sFF)
End Function

'Asiento Transferencia
Public Sub InsertarAsientoTransferencia(ByVal sCodInventario As String, ByVal sNvoCodInventario As String, ByVal cBSValor As Currency, ByVal dFechaRegistro As Date, ByVal sTipo As String, ByVal cValorDepre As Currency)
    Dim oInventario As DInvTransferencia
    Set oInventario = New DInvTransferencia
    oInventario.RegistrarAsientoTransferencia sCodInventario, sNvoCodInventario, cBSValor, dFechaRegistro, sTipo, cValorDepre
End Sub
