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

Public Function NombreAgencia(ByVal psAgecod As String) As String
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset
    sSql = "select cAgeDescripcion from Agencias Where cAgeCod = '" & psAgecod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If R.RecordCount > 0 Then
        NombreAgencia = Trim(UCase(R!cAgeDescripcion))
    Else
        NombreAgencia = ""
    End If
    R.Close
End Function

Public Function RecuperaAgencias(Optional lsCodAge As String) As ADODB.Recordset
Dim sSql As String
Dim sSqlFiltro As String
Dim oConecta As DConecta


    If lsCodAge <> "" Then sSqlFiltro = " where cAgeCod = '" & lsCodAge & "'"
    

    sSql = "Select * From Agencias " & sSqlFiltro
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaAgencias = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function

