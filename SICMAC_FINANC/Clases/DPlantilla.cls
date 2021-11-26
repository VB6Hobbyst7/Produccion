VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dPlantilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String
Dim vsFormatoFecha As String
Public Function GrabaPlantilla(ByVal psPlantillaDocID As String, ByVal psReferencia As String, ByVal psPlantillaDoc As String) As Integer
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
    
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
GrabaPlantilla = 1
sql = "Select cPlantillaDocID from PlantillaDoc where cPlantillaDocID='" & psPlantillaDocID & "'"
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    sql = "UPDATE  PlantillaDoc SET mPlantillaDoc ='" & Replace(psPlantillaDoc, "'", "''") & "' WHERE cPlantillaDocID='" & psPlantillaDocID & "'"
    oConect.Ejecutar sql
Else
    sql = "INSERT INTO PLANTILLADOC (cPlantillaDocID, cPlantillaRef,  mPlantillaDoc) " _
        & " VALUES ('" & psPlantillaDocID & "','','')"
    oConect.Ejecutar sql
    sql = "UPDATE  PlantillaDoc SET mPlantillaDoc ='" & Replace(psPlantillaDoc, "'", "''") & "',cPlantillaRef='" & Replace(psReferencia, "'", "''") & "'  WHERE cPlantillaDocID='" & psPlantillaDocID & "'"
    oConect.Ejecutar sql
End If
rs.Close: Set rs = Nothing
oConect.CierraConexion: Set oConect = Nothing
GrabaPlantilla = 0
End Function
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing
End Sub
