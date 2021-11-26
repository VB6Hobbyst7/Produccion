VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMPatrimonioEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function ObtenerPatrimonioEfectivo(ByVal psAno As String, ByVal psMes As String) As Currency
    Dim oDCOMPatrimonioEfectivo As New COMDCredito.DCOMPatrimonioEfectivo
On Error GoTo ErrorObtenerPatrimonioEfectivo
    Dim lrDatos As ADODB.Recordset
    Set lrDatos = oDCOMPatrimonioEfectivo.ObtenerPatrimonioEfectivoAjustado(psAno, psMes)
    
    If lrDatos.RecordCount > 0 Then
        ObtenerPatrimonioEfectivo = lrDatos!nValor
    Else
        ObtenerPatrimonioEfectivo = 0
    End If
    
    Exit Function
ErrorObtenerPatrimonioEfectivo:
    ObtenerPatrimonioEfectivo = 0
    Err.Raise Err.Number, "NCOMPatrimonioEfectivo:ActualizarPatrimonioEfectivo", Err.Description
End Function
Public Function ActualizarPatrimonioEfectivo(ByVal pnAno As Integer, ByVal pnMes As Integer, ByVal pnValor As Currency, ByVal pcReferencia As String) As Boolean
    Dim oDCOMPatrimonioEfectivo As New COMDCredito.DCOMPatrimonioEfectivo
On Error GoTo ErrorActualizarPatrimonioEfectivo
    Call oDCOMPatrimonioEfectivo.ActualizarPatrimonioEfectivo(pnAno, pnMes, pnValor, pcReferencia)
    ActualizarPatrimonioEfectivo = True
    Exit Function
ErrorActualizarPatrimonioEfectivo:
    ActualizarPatrimonioEfectivo = False
    Err.Raise Err.Number, "NCOMPatrimonioEfectivo:ActualizarPatrimonioEfectivo", Err.Description
End Function
Public Function ListaPatrimonioEfectivo(ByVal pnAno As Integer) As ADODB.Recordset
On Error GoTo ErrorListaPatrimonioEfectivo
    Dim objPatrimonioEfectivo As New COMDCredito.DCOMPatrimonioEfectivo
    Set ListaPatrimonioEfectivo = objPatrimonioEfectivo.ListaPatrimonioEfectivo(pnAno)
    Set objPatrimonioEfectivo = Nothing
    Exit Function
ErrorListaPatrimonioEfectivo:
    Err.Raise Err.Number, "NCOMPatrimonioEfectivo:ListaPatrimonioEfectivo", Err.Description
End Function

