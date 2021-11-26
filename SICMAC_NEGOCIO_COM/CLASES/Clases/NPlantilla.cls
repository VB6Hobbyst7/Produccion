VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NPlantilla"
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
Dim vsFormatoFecha As String
Public Sub inicio(ByVal psFormatoFecha As String)
vsFormatoFecha = psFormatoFecha
End Sub
Public Function GetPlantillaDoc(ByVal psPlantillaDocID As String) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
    
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
GetPlantillaDoc = ""
sql = "Select * from PlantillaDoc where cPlantillaDocID='" & psPlantillaDocID & "'"
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetPlantillaDoc = rs!mPlantillaDoc
End If
rs.Close: Set rs = Nothing
oConect.CierraConexion: Set oConect = Nothing
End Function

Private Sub Class_Initialize()
Dim oIni As ClasIni
Set oIni = New ClasIni
vsConexion = oIni.CadenaConexion
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing
End Sub
Public Function GetPlantillas(ByVal psPlantillaID As String) As ADODB.Recordset
Dim sql  As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
sql = "select cPlantillaDocID, cPlantillaRef, mPlantillaDoc from PlantillaDoc where cPlantillaDocID like '" & psPlantillaID & "%'"
Set rs = oConect.CargaRecordSet(sql)
Set GetPlantillas = rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function GetNewCodPlantilla(ByVal psPlantillaID As String) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsCodigo As String
Set rs = New ADODB.Recordset
Set oConect = New DConecta

If oConect.AbreConexion = False Then Exit Function
sql = "select Max(Substring(cPlantillaDocID,Len('" & psPlantillaID & "') + 1,10)) as MaxCodPlantID from PlantillaDoc where cPlantillaDocID like '" & psPlantillaID & "%'"
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.EOF Then
    If IsNull(rs!MaxCodPlantID) Then
        lsCodigo = Format(1, String(10 - Len(psPlantillaID), "0"))
    Else
        lsCodigo = Format(Val(rs!MaxCodPlantID) + 1, String(10 - Len(psPlantillaID), "0"))
    End If
Else
    lsCodigo = Format(1, String(10 - Len(psPlantillaID), "0"))
End If
rs.Close: Set rs = Nothing
oConect.CierraConexion: Set oConect = Nothing
GetNewCodPlantilla = psPlantillaID + lsCodigo
End Function
Public Function GetNroPlantillaOpe(ByVal psOpeCod As String) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsCodigo As String
Set rs = New ADODB.Recordset
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
sql = "select Max(cPlantillaDocID) as MaxCodPlantID from PlantillaDoc where cPlantillaDocID like '" & psOpeCod & "%'"
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.EOF Then
    GetNroPlantillaOpe = IIf(IsNull(rs!MaxCodPlantID), "", rs!MaxCodPlantID)
End If
rs.Close: Set rs = Nothing
oConect.CierraConexion: Set oConect = Nothing
End Function

