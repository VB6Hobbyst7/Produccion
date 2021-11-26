VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMPlantilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsConexion As String
Dim vsServerCom As String
Dim vsServerPers As String
Dim vsFormatoFecha As String
Public Function GrabaPlantilla(ByVal psPlantillaDocID As String, ByVal psReferencia As String, ByVal psPlantillaDoc As String) As Integer
Dim Sql As String
Dim Rs As New ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta
    
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
GrabaPlantilla = 1
Sql = "Select cPlantillaDocID from PlantillaDoc where cPlantillaDocID='" & psPlantillaDocID & "'"
Set Rs = oConect.CargaRecordSet(Sql)
If Not Rs.EOF And Not Rs.BOF Then
    Sql = "UPDATE  PlantillaDoc SET mPlantillaDoc ='" & Replace(psPlantillaDoc, "'", "''") & "',cPlantillaRef='" & Replace(psReferencia, "'", "''") & "'  WHERE cPlantillaDocID='" & psPlantillaDocID & "'"
    oConect.Ejecutar Sql
Else
    Sql = "INSERT INTO PLANTILLADOC (cPlantillaDocID, cPlantillaRef,  mPlantillaDoc) " _
        & " VALUES('" & psPlantillaDocID & " ','" & Replace(psReferencia, "'", "''") & "','" & Replace(psPlantillaDoc, "'", "''") & "')"
    oConect.Ejecutar Sql
End If
Rs.Close: Set Rs = Nothing
oConect.CierraConexion: Set oConect = Nothing
GrabaPlantilla = 0
End Function

Public Sub inicio(ByVal psFormatoFecha As String)
vsFormatoFecha = psFormatoFecha
End Sub

Public Function GetPlantillaDoc(ByVal psPlantillaDocID As String) As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta
    
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
GetPlantillaDoc = ""
Sql = "Select * from PlantillaDoc where cPlantillaDocID='" & psPlantillaDocID & "'"
Set Rs = oConect.CargaRecordSet(Sql)
If Not Rs.EOF And Not Rs.BOF Then
    GetPlantillaDoc = Rs!mPlantillaDoc
End If
Rs.Close: Set Rs = Nothing
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetPlantillas(ByVal psPlantillaID As String) As ADODB.Recordset
Dim Sql  As String
Dim Rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta

Set Rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
Sql = "select cPlantillaDocID, cPlantillaRef, mPlantillaDoc from PlantillaDoc where cPlantillaDocID like '" & psPlantillaID & "%'"
Set Rs = oConect.CargaRecordSet(Sql)
Set GetPlantillas = Rs
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetNewCodPlantilla(ByVal psPlantillaID As String) As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta
Dim lsCodigo As String
Set Rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta

If oConect.AbreConexion = False Then Exit Function
Sql = "select Max(Substring(cPlantillaDocID,Len('" & psPlantillaID & "') + 1,10)) as MaxCodPlantID from PlantillaDoc where cPlantillaDocID like '" & psPlantillaID & "%'"
Set Rs = oConect.CargaRecordSet(Sql)
If Not Rs.EOF And Not Rs.EOF Then
    If IsNull(Rs!MaxCodPlantID) Then
        lsCodigo = Format(1, String(10 - Len(psPlantillaID), "0"))
    Else
        lsCodigo = Format(Val(Rs!MaxCodPlantID) + 1, String(10 - Len(psPlantillaID), "0"))
    End If
Else
    lsCodigo = Format(1, String(10 - Len(psPlantillaID), "0"))
End If
Rs.Close: Set Rs = Nothing
oConect.CierraConexion: Set oConect = Nothing
GetNewCodPlantilla = psPlantillaID + lsCodigo
End Function

Public Function GetNroPlantillaOpe(ByVal psOpeCod As String) As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta
Dim lsCodigo As String
Set Rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
Sql = "select Max(cPlantillaDocID) as MaxCodPlantID from PlantillaDoc where cPlantillaDocID like '" & psOpeCod & "%'"
Set Rs = oConect.CargaRecordSet(Sql)
If Not Rs.EOF And Not Rs.EOF Then
    GetNroPlantillaOpe = IIf(IsNull(Rs!MaxCodPlantID), "", Rs!MaxCodPlantID)
End If
Rs.Close: Set Rs = Nothing
oConect.CierraConexion: Set oConect = Nothing
End Function

Private Sub Class_Initialize()
Dim oIni As COMConecta.DCOMClasIni
Set oIni = New COMConecta.DCOMClasIni
vsConexion = oIni.CadenaConexion
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing
End Sub


