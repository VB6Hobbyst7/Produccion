VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DFeriado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RecuperaFeriado() As ADODB.Recordset
Dim ssql As String
Dim oConecta As DConecta
    
    ssql = "Select * From Feriado Order by dFeriado "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaFeriado = oConecta.CargaRecordSet(ssql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaDias(ByVal dRangoF As Date) As ADODB.Recordset
Dim VSQL As String
Dim Co As DConecta
Dim Rs As New ADODB.Recordset
Set Co = New DConecta

VSQL = "select dFeriado, cdescrip from Feriado where dferiado > '" & Format(dRangoF, "YYYY/MM/DD") & "' order by dFeriado"
Co.AbreConexion
Set Rs = Co.CargaRecordSet(VSQL)
Co.CierraConexion
Set RecuperaDias = Rs
End Function

Public Function DetallaFeriado(ByVal FecVer As Date) As ADODB.Recordset
Dim VSQL As String
Dim Co As DConecta
Dim Rs As New ADODB.Recordset
Set Co = New DConecta
' Verifica si la Fecha seleccionada ya existe en la tabla feriados
VSQL = "select dFeriado, cDescrip from Feriado where dFeriado = '" & Format(FecVer, "MM/DD/YYYY") & "' "
Co.AbreConexion
Set Rs = Co.CargaRecordSet(VSQL)
Co.CierraConexion
Set DetallaFeriado = Rs
End Function
Public Sub InsertaFecha(ByVal ldFecha As Date, ByVal lsDesFer As String, ByVal pdFecSis As Date, _
ByVal psCodAge As String, ByVal psUser As String)

Dim Co As DConecta
Dim SQL1 As String
Dim sUltMov As String
Set Co = New DConecta

Dim ObjMov As DMov
Set ObjMov = New DMov
sUltMov = ObjMov.GeneraMovNro(pdFecSis, psCodAge, psUser)
SQL1 = "Insert Feriado(dFeriado, cDescrip, cUltimaActualizacion) values ('" & Format(ldFecha, "mm/dd/yyyy") & "', '" & Trim(lsDesFer) & "', '"
SQL1 = SQL1 & sUltMov & "')"
Co.AbreConexion
Co.Ejecutar (SQL1)
Co.CierraConexion

Set Co = Nothing
End Sub
Public Sub EliminaFeriado(ByVal DelFer As Date)
Dim SQL1 As String
Dim Co As DConecta
Set Co = New DConecta
SQL1 = "delete Feriado where dFeriado = '" & Format(DelFer, "MM/DD/YYYY") & "'"
Co.AbreConexion
Co.BeginTrans
Co.Ejecutar (SQL1)
Co.CommitTrans
Co.CierraConexion
End Sub

Public Sub LlenaDomingo(ByVal lsDomingo As Date, ByVal pdFecSis As Date, _
ByVal psCodAge As String, ByVal psUser As String)
Dim Co As DConecta
Dim SQL1 As String
Dim sUltMov As String
Set Co = New DConecta

Dim ObjMov As DMov
Set ObjMov = New DMov
sUltMov = ObjMov.GeneraMovNro(pdFecSis, psCodAge, psUser)

SQL1 = " Insert Feriado values ('" & Format(lsDomingo, "YYYY/MM/DD") & "', 'DOMINGO', '"
SQL1 = SQL1 & sUltMov & "')"

Co.AbreConexion
Co.BeginTrans
Co.Ejecutar (SQL1)
Co.CommitTrans
Co.CierraConexion

Set Co = Nothing
Set ObjMov = Nothing

End Sub

