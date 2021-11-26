VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMFeriado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RecuperaFeriado() As ADODB.Recordset
Dim sSQL As String
Dim oConecta As COMConecta.DCOMConecta
    
    sSQL = "Select * From Feriado Order by dFeriado "
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaFeriado = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
 
Public Function RecuperaFeriadoAgencias(ByVal pdFecha As Date) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As COMConecta.DCOMConecta
    
    sSQL = "Select A.cAgeCod,A.cAgeDescripcion,  (CASE WHEN FA.cCodAge IS NULL THEN 0 ELSE 1 END) as Valor"
    sSQL = sSQL & " from Agencias A"
    'By Capi 06122007 por formato fecha
    'sSQL = sSQL & "  LEFT Join FeriadoAge FA ON A.cAgeCod = FA.cCodAge  and FA.dFeriado= '" & Format(pdFecha, "mm/dd/yyyy") & "'"
    sSQL = sSQL & "  LEFT Join FeriadoAge FA ON A.cAgeCod = FA.cCodAge  and FA.dFeriado= '" & Format(pdFecha, "yyyymmdd") & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaFeriadoAgencias = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function


Public Function RecuperaDias(ByVal dRangoF As Date) As ADODB.Recordset
Dim VSQL As String
Dim Co As COMConecta.DCOMConecta
Dim Rs As New ADODB.Recordset
Set Co = New COMConecta.DCOMConecta

VSQL = "select dFeriado, cdescrip from Feriado where dferiado > '" & Format(dRangoF, "YYYY/MM/DD") & "' order by dFeriado"
Co.AbreConexion
Set Rs = Co.CargaRecordSet(VSQL)
Co.CierraConexion
Set RecuperaDias = Rs
End Function

Public Function DetallaFeriado(ByVal FecVer As Date, ByVal psCodAge As String) As ADODB.Recordset
Dim VSQL As String
Dim Co As COMConecta.DCOMConecta
Dim Rs As New ADODB.Recordset
Set Co = New COMConecta.DCOMConecta
' Verifica si la Fecha seleccionada ya existe en la tabla feriados
'VSQL = "select dFeriado, cDescrip from Feriado where dFeriado = '" & Format(FecVer, "MM/DD/YYYY") & "' "

VSQL = " Select F.dferiado, FA.cCodAge"
VSQL = VSQL & " from Feriado F"
VSQL = VSQL & " Inner Join FeriadoAge FA ON F.dFeriado=FA.dFeriado"
'By Capi 06122007 por formato fecha
'VSQL = VSQL & " Where F.dFeriado='" & Format(FecVer, "MM/DD/YYYY") & "' AND FA.cCodAge = '" & psCodAge & "'"
VSQL = VSQL & " Where F.dFeriado='" & Format(FecVer, "YYYYMMDD") & "' AND FA.cCodAge = '" & psCodAge & "'"

Co.AbreConexion
Set Rs = Co.CargaRecordSet(VSQL)
Co.CierraConexion
Set DetallaFeriado = Rs
End Function
Public Sub InsertaFecha(ByVal ldFecha As Date, ByVal lsDesFer As String, ByVal pdFecSis As Date, _
ByVal psCodAge As String, ByVal psUser As String)

Dim Co As COMConecta.DCOMConecta
Dim SQL1 As String
Dim sUltMov As String
Set Co = New COMConecta.DCOMConecta

Dim ObjMov As COMNContabilidad.NCOMContFunciones  'DMov
Set ObjMov = New COMNContabilidad.NCOMContFunciones 'DMov
sUltMov = ObjMov.GeneraMovNro(pdFecSis, psCodAge, psUser)
'By Capi 06122007 por formato de fecha
'SQL1 = "Insert Feriado(dFeriado, cDescrip, cUltimaActualizacion) values ('" & Format(ldFecha, "mm/dd/yyyy") & "', '" & Trim(lsDesFer) & "', '"
SQL1 = "Insert Feriado(dFeriado, cDescrip, cUltimaActualizacion) values ('" & Format(ldFecha, "yyyymmdd") & "', '" & Trim(lsDesFer) & "', '"
SQL1 = SQL1 & sUltMov & "')"
Co.AbreConexion
Co.Ejecutar (SQL1)
Co.CierraConexion

Set Co = Nothing
End Sub

Public Sub dInsertFeriadoAge(ByVal dFeriado As Date, ByVal sCodAge As String)
Dim Co As COMConecta.DCOMConecta
Set Co = New COMConecta.DCOMConecta

Dim sSQL As String

    sSQL = "INSERT INTO FeriadoAge(dFeriado, cCodAge) "
    
    'By Capi 06122007 para uniformizar formtato de fecha se cambio a yyyy/mm/dd
    'sSQL = sSQL & " VALUES('" & Format(dFeriado, "mm/dd/yyyy") & "','" & sCodAge & "')"
    sSQL = sSQL & " VALUES('" & Format(dFeriado, "yyyymmdd") & "','" & sCodAge & "')"

    Co.AbreConexion
    Co.Ejecutar (sSQL)
    Co.CierraConexion


End Sub

Public Sub EliminaFeriado(ByVal DelFer As Date)
Dim SQL1 As String
Dim SQL2 As String
Dim Co As COMConecta.DCOMConecta
Set Co = New COMConecta.DCOMConecta
'By Capi 06122007 por formato de fecha
'SQL1 = "delete Feriado where dFeriado = '" & Format(DelFer, "MM/DD/YYYY") & "'"
'SQL2 = "delete FeriadoAge where dFeriado = '" & Format(DelFer, "MM/DD/YYYY") & "'"

SQL1 = "delete Feriado where dFeriado = '" & Format(DelFer, "YYYYMMDD") & "'"
SQL2 = "delete FeriadoAge where dFeriado = '" & Format(DelFer, "YYYYMMDD") & "'"

Co.AbreConexion
Co.BeginTrans
Co.Ejecutar (SQL1)
Co.Ejecutar (SQL2)
Co.CommitTrans
Co.CierraConexion
End Sub

Public Sub LlenaDomingo(ByVal lsDomingo As Date, ByVal pdFecSis As Date, _
ByVal psCodAge As String, ByVal psUser As String)
Dim Co As COMConecta.DCOMConecta
Dim SQL1 As String
Dim sUltMov As String
Set Co = New COMConecta.DCOMConecta

Dim ObjMov As COMNContabilidad.NCOMContFunciones  'DMov
Set ObjMov = New COMNContabilidad.NCOMContFunciones 'DMov
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



