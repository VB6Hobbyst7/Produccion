VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRHConvenios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Function Cargo(ByVal psPersCod As String) As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Set Co = New DConecta

Sql = " Select R.cPersCod,"
Sql = Sql & " (select cRHCargoDescripcion from RHCargosTabla where cRHCargoCod = R.cRHCargoCod ) Cargo"
Sql = Sql & " from rhcargos R where R.cPersCod = '" & psPersCod & "' and"
Sql = Sql & " dRHCargoFecha =(select max(dRHCargoFecha) dFechaC from rhcargos where cPersCod = '" & psPersCod & "')"

Cargo = ""
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
If Not (Rs.EOF And Rs.BOF) Then
    Cargo = Rs!Cargo
End If
Co.CierraConexion

Set Co = Nothing
Set Rs = Nothing
End Function

Public Function Get_Provincia(ByVal psPersUbigeo As String) As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Set Co = New DConecta

Sql = "select cUbiGeoDescripcion from dbo.UbicacionGeografica "
Sql = Sql & " Where cUbiGeoCod = '2" & Mid(psPersUbigeo, 2, 4) & "0000000'"

Get_Provincia = ""
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
If Not (Rs.EOF And Rs.BOF) Then
    Get_Provincia = Trim(Rs!cUbiGeoDescripcion)
End If
Co.CierraConexion

Set Co = Nothing
Set Rs = Nothing
End Function
Public Function Get_CodUbigeo(ByVal psPersCod As String) As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Set Co = New DConecta

Sql = "select cPersDireccUbiGeo from Persona where cPersCod  = '" & psPersCod & "'"

Get_CodUbigeo = ""
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
If Not (Rs.EOF And Rs.BOF) Then
    Get_CodUbigeo = Rs!cPersDireccUbiGeo
End If
Co.CierraConexion

Set Co = Nothing
Set Rs = Nothing
End Function


Public Function Get_Distrito(ByVal psPersUbigeo As String) As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Set Co = New DConecta

Sql = "select cUbiGeoDescripcion from dbo.UbicacionGeografica "
Sql = Sql & " Where cUbiGeoCod = '3" & Mid(psPersUbigeo, 2, 6) & "00000'"

Get_Distrito = ""
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
If Not (Rs.EOF And Rs.BOF) Then
    Get_Distrito = Rs!cUbiGeoDescripcion
End If
Co.CierraConexion

Set Co = Nothing
Set Rs = Nothing
End Function

Public Function Get_Urb(ByVal psPersUbigeo As String) As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Set Co = New DConecta

Sql = "select cUbiGeoDescripcion from dbo.UbicacionGeografica "
Sql = Sql & " Where cUbiGeoCod = '" & psPersUbigeo & "'"

Get_Urb = ""
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
If Not (Rs.EOF And Rs.BOF) Then
    Get_Urb = Rs!cUbiGeoDescripcion
End If
Co.CierraConexion

Set Co = Nothing
Set Rs = Nothing
End Function
