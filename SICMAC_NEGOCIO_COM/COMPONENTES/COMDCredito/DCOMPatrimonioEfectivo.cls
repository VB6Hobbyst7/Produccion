VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMPatrimonioEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*****************************************************************************************
'***     Rutina:           DPatrimonioEfectivo
'***     Descripcion:      Clase que permite Administrar Datos relacionados al Patrimonio Efectivo
'***     Creado por:        ORCR
'***     Maquina:
'***     Fecha-Tiempo:         28/02/2014 12:06:08 AM
'***     Ultima Modificacion: Inicio de descripicion
'*****************************************************************************************
Public Function ObtenerPatrimonioEfectivoAjustado(ByVal psAno As String, ByVal psMes As String) As ADODB.Recordset
On Error GoTo ErrorObtenerPatrimonioEfectivoAjustado
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String

    sSql = "exec sp_obtenerPatrimonioEfectivoAjustado '" & psAno & "', '" & psMes & "'"

    oConecta.AbreConexion
    Set ObtenerPatrimonioEfectivoAjustado = oConecta.CargaRecordSet(sSql, adLockReadOnly)
    oConecta.CierraConexion

    Set oConecta = Nothing
    
    Exit Function
ErrorObtenerPatrimonioEfectivoAjustado:
    Err.Raise Err.Number, "DCOMPatrimonioEfectivo:ObtenerPatrimonioEfectivoAjustado"", Err.Description"
End Function
Public Sub ActualizarPatrimonioEfectivo(ByVal pnAno As Integer, ByVal pnMes As Integer, ByVal pnValor As Currency, ByVal pcReferencia As String)
On Error GoTo ErrorActualizarPatrimonioEfectivo
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "exec sp_actualizarPatrimonioEfectivoAjustado " & pnAno & "," & pnMes & "," & pnValor & ",'" & pcReferencia & "'"
        
    oConecta.AbreConexion
    oConecta.CargaRecordSet (sSql)
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    
    Exit Sub
ErrorActualizarPatrimonioEfectivo:
    Err.Raise Err.Number, "DCOMPatrimonioEfectivo:ActualizarPatrimonioEfectivo", Err.Description
End Sub
Public Function ListaPatrimonioEfectivo(ByVal pnAno As Integer) As ADODB.Recordset
On Error GoTo ErrorListaPatrimonioEfectivo
    Dim oConec As New COMConecta.DCOMConecta
    Dim sSql As String
    
    sSql = "exec sp_listarPatrimonioEfectivoAjustado " & pnAno
    
    oConec.AbreConexion
    Set ListaPatrimonioEfectivo = oConec.CargaRecordSet(sSql, adLockReadOnly)
    oConec.CierraConexion
    
    Set oConec = Nothing
    
    Exit Function
ErrorListaPatrimonioEfectivo:
    Err.Raise Err.Number, "DCOMPatrimonioEfectivo:ListaPatrimonioEfectivo"", Err.Description"
End Function
