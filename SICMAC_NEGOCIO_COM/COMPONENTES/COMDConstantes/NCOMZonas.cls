VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMZonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oconec As New COMConecta.DCOMConecta

Public Sub ActualizaZonas(ByVal psCodigo As String, ByVal psDescripcion As String, ByVal pnAccion As Integer)
    If pnAccion = 1 Then
       dInsertZonas psCodigo, psDescripcion, False
    Else
        If pnAccion = 2 Then
            dUpdateZonas psCodigo, psDescripcion, False
        Else
            dDeleteZonas psCodigo, False
        End If
    End If
End Sub

Public Sub dInsertZonas(ByVal psUbiGeoCod As String, ByVal psUbiGeoDescripcion As String, Optional pbEjecBatch As Boolean = False)
    Dim sSQL As String
    
    oconec.AbreConexion
    sSQL = "INSERT INTO UbicacionGeografica(cUbiGeoCod, cUbiGeoDescripcion) "
    sSQL = sSQL & " VALUES('" & psUbiGeoCod & "','" & psUbiGeoDescripcion & "')"
    If pbEjecBatch Then
        oconec.AdicionaCmdBatch sSQL
    Else
        oconec.Ejecutar sSQL
    End If
    oconec.CierraConexion
End Sub

Public Sub dUpdateZonas(ByVal psUbiGeoCod As String, ByVal psUbiGeoDescripcion As String, Optional pbEjecBatch As Boolean = False)
Dim sSQL As String
    oconec.AbreConexion
    sSQL = "UPDATE UbicacionGeografica SET cUbiGeoDescripcion = '" & psUbiGeoDescripcion & "' "
    sSQL = sSQL & " Where cUbiGeoCod = '" & psUbiGeoCod & "'"
    If pbEjecBatch Then
        oconec.AdicionaCmdBatch sSQL
    Else
        oconec.Ejecutar sSQL
    End If
    oconec.CierraConexion
End Sub

Public Sub dDeleteZonas(ByVal psUbiGeoCod As String, Optional pbEjecBatch As Boolean = False)
Dim sSQL As String
    oconec.AbreConexion
    sSQL = "DELETE UbicacionGeografica  Where cUbiGeoCod = '" & psUbiGeoCod & "'"
    If pbEjecBatch Then
        oconec.AdicionaCmdBatch sSQL
    Else
        oconec.Ejecutar sSQL
    End If
    oconec.CierraConexion
End Sub


