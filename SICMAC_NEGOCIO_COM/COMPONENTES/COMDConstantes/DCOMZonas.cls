VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMZonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oConecta As COMConecta.DCOMConecta

Public Function DameZonas(ByVal pnNivel As Integer, ByVal psValor As String) As ADODB.Recordset
Dim sSQL As String
'Dim oConecta As COMConecta.DCOMConecta
    Select Case pnNivel
        Case 0
            sSQL = "select * from UbicacionGeografica where len(cUbiGeoCod)=3 order by cUbiGeoDescripcion"
        Case 1
            sSQL = "select * from UbicacionGeografica where cUbiGeoCod like '1__000000000' order by cUbiGeoDescripcion"
        Case 2
            sSQL = "select * from UbicacionGeografica where cUbiGeoCod like '2" & Mid(psValor, 2, 2) & "__0000000' order by cUbiGeoDescripcion"
        Case 3
            sSQL = "select * from UbicacionGeografica where cUbiGeoCod like '3" & Mid(psValor, 2, 4) & "__00000' order by cUbiGeoDescripcion"
        Case 4
            sSQL = "select * from UbicacionGeografica where cUbiGeoCod like '4" & Mid(psValor, 2, 6) & "%' order by cUbiGeoDescripcion"
    End Select
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DameZonas = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function DameMaximoValorZona(ByVal psValor As String, ByVal pnNivel As Integer) As String
Dim sSQL As String
Dim R As ADODB.Recordset
'Dim oConecta As COMConecta.DCOMConecta
Dim sValorTmp As String

    Select Case pnNivel
        Case 1
            sSQL = "Select MAX(cUbiGeoCod) as sMax From UbicacionGeografica Where cUbiGeoCod like '1__000000000' "
        Case 2
            sSQL = "select MAX(cUbiGeoCod) as sMax from UbicacionGeografica where cUbiGeoCod like '2" & Mid(psValor, 2, 2) & "__0000000' "
        Case 3
            sSQL = "select MAX(cUbiGeoCod) as sMax from UbicacionGeografica where cUbiGeoCod like '3" & Mid(psValor, 2, 4) & "__00000' "
        Case 4
            sSQL = "select MAX(cUbiGeoCod) as sMax from UbicacionGeografica where cUbiGeoCod like '4" & Mid(psValor, 2, 6) & "%' "
    End Select
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If IsNull(R!sMax) Then
        sValorTmp = psValor & String(12, "0")
    Else
        sValorTmp = R!sMax
    End If
    Select Case pnNivel
        Case 1
            DameMaximoValorZona = pnNivel & Right("00" & Trim(Str(CInt(Mid(sValorTmp, 2, 2)) + 1)), 2) & String(9, "0")
        Case 2
            DameMaximoValorZona = pnNivel & Mid(sValorTmp, 2, 2) & Right("00" & Trim(Str(CInt(Mid(sValorTmp, 4, 2)) + 1)), 2) & String(7, "0")
        Case 3
            DameMaximoValorZona = pnNivel & Mid(sValorTmp, 2, 4) & Right("00" & Trim(Str(CInt(Mid(sValorTmp, 6, 2)) + 1)), 2) & String(5, "0")
        Case 4
            DameMaximoValorZona = pnNivel & Mid(sValorTmp, 2, 6) & Right("00000" & Trim(Str(CInt(Mid(sValorTmp, 8, 12)) + 1)), 5)
    End Select

End Function

Public Function DameUnaZona(ByVal psValor As String) As ADODB.Recordset
Dim sSQL As String
'Dim oConecta As COMConecta.DCOMConecta

    sSQL = "select * from UbicacionGeografica where cUbiGeoCod= '" & psValor & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set DameUnaZona = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function


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
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSQL = "INSERT INTO UbicacionGeografica(cUbiGeoCod, cUbiGeoDescripcion) "
    sSQL = sSQL & " VALUES('" & psUbiGeoCod & "','" & psUbiGeoDescripcion & "')"
    If pbEjecBatch Then
        oConecta.AdicionaCmdBatch sSQL
    Else
        oConecta.Ejecutar sSQL
    End If
    oConecta.CierraConexion
End Sub

Public Sub dUpdateZonas(ByVal psUbiGeoCod As String, ByVal psUbiGeoDescripcion As String, Optional pbEjecBatch As Boolean = False)
Dim sSQL As String
Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSQL = "UPDATE UbicacionGeografica SET cUbiGeoDescripcion = '" & psUbiGeoDescripcion & "' "
    sSQL = sSQL & " Where cUbiGeoCod = '" & psUbiGeoCod & "'"
    If pbEjecBatch Then
        oConecta.AdicionaCmdBatch sSQL
    Else
        oConecta.Ejecutar sSQL
    End If
    oConecta.CierraConexion
End Sub

Public Sub dDeleteZonas(ByVal psUbiGeoCod As String, Optional pbEjecBatch As Boolean = False)
Dim sSQL As String
    oConecta.AbreConexion
    sSQL = "DELETE UbicacionGeografica  Where cUbiGeoCod = '" & psUbiGeoCod & "'"
    If pbEjecBatch Then
        oConecta.AdicionaCmdBatch sSQL
    Else
        oConecta.Ejecutar sSQL
    End If
    oConecta.CierraConexion
End Sub




