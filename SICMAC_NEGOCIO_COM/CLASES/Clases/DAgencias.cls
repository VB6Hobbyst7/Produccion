VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DAgencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function NombreAgencia(ByVal psAgeCod As String) As String
Dim ssql As String
Dim oconecta As DConecta
Dim R As ADODB.Recordset
    ssql = "select cAgeDescripcion from Agencias Where cAgeCod = '" & psAgeCod & "'"
    Set oconecta = New DConecta
    oconecta.AbreConexion
    Set R = oconecta.CargaRecordSet(ssql)
    oconecta.CierraConexion
    Set oconecta = Nothing
    
    If R.RecordCount > 0 Then
        NombreAgencia = Trim(UCase(R!cAgeDescripcion))
    Else
        NombreAgencia = ""
    End If
    R.Close
End Function

Public Function RecuperaAgencias(Optional lsCodAge As String) As ADODB.Recordset
Dim ssql As String
Dim sSqlFiltro As String
Dim oconecta As DConecta


    If lsCodAge <> "" Then sSqlFiltro = " where cAgeCod = '" & lsCodAge & "'"
    

    ssql = "Select * From Agencias " & sSqlFiltro
    
    Set oconecta = New DConecta
    oconecta.AbreConexion
    Set RecuperaAgencias = oconecta.CargaRecordSet(ssql)
    oconecta.CierraConexion
    Set oconecta = Nothing

End Function

