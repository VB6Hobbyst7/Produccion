VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMPersGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oConn As COMConecta.DCOMConecta
Dim sql As String

Private Sub Class_Initialize()
Dim oIni As COMConecta.DCOMClasIni
Dim sConn As String

    Set oIni = New COMConecta.DCOMClasIni
    sConn = oIni.CadenaConexion
    Set oIni = Nothing
    
Set oConn = New COMConecta.DCOMConecta
If oConn.AbreConexion(sConn) = False Then
    Err.Raise Err.Number, "Error en Conexion a Base de datos", Err.Description
End If
End Sub
Private Sub Class_Terminate()
    oConn.CierraConexion
    Set oConn = Nothing
End Sub

Public Sub dInsertCodigoPostal(ByVal pnCodPostal, ByVal psDescCodPostal, ByVal psCodZon As String)

On Error GoTo dErrorCodigoPostal

    sql = "INSERT CODIGOPOSTAL (nCodPostal, cDesCodPostal, cCodZon) " _
        & "VALUES (" & pnCodPostal & ",'" & psDescCodPostal & "'," & psCodZon & ")"

    oConn.Ejecutar sql

    Exit Sub
    
dErrorCodigoPostal:
    Err.Raise Err.Number, "Error en <<dInsertCodigoPostal>>" + Err.Description, "Aviso"

End Sub

Public Function dObtieneZona(ByVal psTipoZon As String, Optional ByVal psDep As String = "@", _
                            Optional ByVal psProv As String = "@") As Recordset
    
On Error GoTo dErrorObtieneZona

    sql = "SELECT cUbiGeoCod, cUbiGeoDescripcion FROM UBICACIONGEOGRAFICA " _
        & "WHERE SUBSTRING(cUbiGeoCod,1,1) = '" & psTipoZon & "'" _
        
    If psDep <> "@" Then
        sql = sql & " AND SUBSTRING(cUbiGeoCod,2,2) = '" & psDep & "'"
    End If
    
    If psProv <> "@" Then
        sql = sql & " AND SUBSTRING(cUbiGeoCod, 4,2) = '" & psProv & "'"
    End If

    Set dObtieneZona = oConn.CargaRecordSet(sql)
        
    Exit Function
    
dErrorObtieneZona:
    Err.Raise Err.Number, "Error en <<dObtieneZona>>" + Err.Description, "Aviso"

End Function

Public Function dGeneraCodigoPostal() As Long
Dim rs As Recordset

    On Error GoTo dErrorGeneraCodigoPostal
    sql = "SELECT ISNULL(MAX(nCodPostal),0) + 1 as CodPostal FROM CODIGOPOSTAL "
    
    Set rs = oConn.CargaRecordSet(sql)

    If Not rs.EOF And Not rs.BOF Then
        dGeneraCodigoPostal = rs!CodPostal
    Else
        dGeneraCodigoPostal = 1
    End If
    
    Set rs = Nothing
    
    Exit Function
    
dErrorGeneraCodigoPostal:
    Err.Raise Err.Number, "Error en <<dGeneraCodigoPostal>>" + Err.Description, "Aviso"
    
End Function
Public Function GetCIIU() As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
sql = " Select cCIIUcod, cCIIUdescripcion, " & _
      " (Select cConsDescripcion from Constante where nConsCod=1013 and nConsValor=nSector ) Sector   " & _
      " From CIIU "
Set rs = oConn.CargaRecordSet(sql)
Set GetCIIU = rs
Set rs = Nothing
End Function
Public Function ExisteCIIUGiro(ByVal psDescripcion As String, ByVal psCIIU As String) As Boolean
Dim rs As New ADODB.Recordset
Dim sql As String

sql = "select * from CIIUGiroNeg "
sql = sql & " Where cCIIUCod ='" & Trim(psCIIU) & "' And cDescripcion = '" & psDescripcion & "'"
Set rs = oConn.CargaRecordSet(sql)
If rs.EOF And rs.BOF Then
    ExisteCIIUGiro = False
Else
    ExisteCIIUGiro = True
End If
Set rs = Nothing

End Function

Public Function ExisteCIIU(ByVal psCodCIIU As String) As Boolean
Dim rs As New ADODB.Recordset
Dim sql As String
sql = "Select * from CIIU where  cCIIUcod='" & psCodCIIU & "'"
Set rs = oConn.CargaRecordSet(sql)
If rs.EOF And rs.BOF Then
    ExisteCIIU = False
Else
    ExisteCIIU = True
End If
Set rs = Nothing
End Function
Public Sub dUpdateCIIUGiro(ByVal psCIIU As String, ByVal psDescripcion1 As String, ByVal psDescripcion2 As String)
Dim sql As String
sql = " Update CIIUGiroNeg "
sql = sql & " Set cDescripcion='" & psDescripcion2 & "' "
sql = sql & " Where cCIIUCod ='" & psCIIU & "' And cDescripcion = '" & psDescripcion1 & "'"
oConn.Ejecutar (sql)
End Sub

Public Sub dUpdateCIIU(ByVal psCodCIIU As String, ByVal psDescripcion As String, _
ByVal pnSector As Integer, ByVal pdFecSis As Date, ByVal psUser As String, ByVal psCodAge As String)


Dim sql As String
Dim sUltMov As String
sUltMov = GeneraMovNro(pdFecSis, psCodAge, psUser)


sql = " Update CIIU"
sql = sql & " set cCIIUdescripcion='" & psDescripcion & "' "
sql = sql & " , cUltimaActualizacion ='" & sUltMov & "' "
sql = sql & " , nSector ='" & pnSector & "' "
sql = sql & " where cCIIUcod='" & psCodCIIU & "'"

'oConn.BeginTrans
oConn.Ejecutar (sql)
'oConn.CommitTrans
End Sub
Public Sub dInsertaCIIUGiro(ByVal psCIIU As String, ByVal psDescripcion As String)
Dim sql As String
Dim rs As New ADODB.Recordset
sql = "select isnull(max(cGiroNegCod),0) Nro from CIIUGiroNeg"
Set rs = oConn.CargaRecordSet(sql)
sql = " Insert Into CIIUGiroNeg " & _
      " (cGiroNegCod,cDescripcion,cCIIUCod) values " & _
      " ('" & Format(Val(rs!Nro) + 1, "00000") & "','" & psDescripcion & "','" & psCIIU & "')"
'oConn.BeginTrans
oConn.ConexionActiva.Execute (sql)
'oConn.CommitTrans
Set rs = Nothing
End Sub
Public Function GetGiroDescripcion(ByVal psCodGiro As String) As String
Dim sql As String
Dim rs As New ADODB.Recordset

sql = " Select cDescripcion, cGiroNegCod, cCIIUCod from CIIUGIRONEG "
sql = sql & " where cGiroNegCod='" & psCodGiro & "'"
Set rs = oConn.CargaRecordSet(sql)
If rs.EOF And rs.BOF Then
    GetGiroDescripcion = ""
Else
    GetGiroDescripcion = rs!CDescripcion
End If
Set rs = Nothing
End Function

Public Sub dInsertaCIIU(ByVal psCodCIIU As String, _
 ByVal psDescripcion As String, ByVal pnSector As Integer, _
ByVal pdFecSis As Date, ByVal psUser As String, ByVal psCodAge As String)

Dim sql As String
Dim sUltMov As String

sUltMov = GeneraMovNro(pdFecSis, psCodAge, psUser)



sql = " Insert CIIU"
sql = sql & " (cCIIUcod,cCIIUdescripcion,cUltimaActualizacion,nSector)"
sql = sql & " Values"
sql = sql & " ('" & psCodCIIU & "','" & psDescripcion & "','" & Format(pdFecSis, "dd/mm/yyyy") & "',"
sql = sql & pnSector & ")"

'oConn.BeginTrans
oConn.Ejecutar (sql)
'oConn.CommitTrans

End Sub
Public Function GetCIIUGiro(ByVal psCIIU As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
sql = " select * from CIIUGiroNeg WHERE cCIIUCod = '" & psCIIU & "'"
Set rs = oConn.CargaRecordSet(sql)
Set GetCIIUGiro = rs
Set rs = Nothing
End Function

Public Function dObtieneCodigosPostal() As Recordset
Dim rs As Recordset

    sql = "SELECT nCodPostal, cDesCodPostal, cCodZon, cUbiGeoDescripcion " _
            & " FROM CODIGOPOSTAL CP INNER JOIN UBICACIONGEOGRAFICA UG ON CP.cCodZon = UG.cUbiGeoCod " _
            & "ORDER BY nCodPostal"
            
    Set rs = oConn.CargaRecordSet(sql)
    Set dObtieneCodigosPostal = rs
    Set rs = Nothing
    
End Function

Public Sub dUpdateCodigoPostal(ByVal psCodPostal As Long, ByVal psDesCodPostal As String, psCodZon As String)

    On Error GoTo dErrorUpdateCodigoPostal
    sql = "UPDATE CODIGOPOSTAL SET cDesCodPostal ='" & psDesCodPostal & "', cCodZon= '" & psCodZon & "'" _
        & "WHERE nCodPostal = " & psCodPostal
        
   oConn.Ejecutar sql
   
   Exit Sub
   
dErrorUpdateCodigoPostal:
    Err.Raise Err.Number, "Error en <<dUpdateCodigoPostal>>" + Err.descripction, "Aviso"

End Sub

Public Sub dDeleteCodigoPostal(ByVal psCodPostal As Long)

    On Error GoTo dErrorDeleteCodigoPostal
    sql = "DELETE CODIGOPOSTAL WHERE nCodPostal = " & psCodPostal
    
    oConn.Ejecutar sql
    
    Exit Sub
    
dErrorDeleteCodigoPostal:
    Err.Raise Err.Number, "Error en <<dDeleteCodigoPostal>>" + Err.Description, "Aviso"
    
End Sub

'Para hacer mas rapida la Generacion del Nro de Mov

Private Function GeneraMovNro(Optional ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional ByVal psMovNro As String = "", Optional pnCont As Integer = 0) As String
    On Error GoTo GeneraMovNroErr
    Dim rs As ADODB.Recordset
    Dim sql As String
    Set rs = New ADODB.Recordset
    If pnCont = 0 And psMovNro <> "" Then
        psMovNro = Left(psMovNro, 19) & Format(Val(Mid(psMovNro, 20, 2)) + 1, "00") & Right(psMovNro, 4)
    End If
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & oConn.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & psCodAge & "','" & psUser & "'"
    Else
       sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    Set rs = oConn.Ejecutar(sql)
    If Not rs.EOF Then
        GeneraMovNro = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
GeneraMovNroErr:
    Err.Raise Err.Number, "Genera MovNro", Err.Description
End Function

'MIOL 20120815, SEGUN RQ12134 ************************************************************
Public Function RecuperaEstCivConvenio(ByVal psCuenta As String, ByVal pdFecSist As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_PersEstCivConvenio '" & psCuenta & "','" & pdFecSist & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaEstCivConvenio = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function

Public Function RecuperaIngresoNeto(ByVal psCuenta As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_IngNeto '" & psCuenta & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaIngresoNeto = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
'END MIOL ********************************************************************************
'JUEZ 20140530 **************************************************************
Public Function GetSectorEconomico() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_SectorEconomico"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set GetSectorEconomico = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Sub ActualizarLimiteSector(ByVal psSectorCod As String, ByVal pnLimite As Double)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_upd_ActualizarLimiteSector '" & psSectorCod & "'," & pnLimite

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Call oConecta.Ejecutar(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Sub
Public Function VerificaSuperaUmbralSectorEcon(ByVal psPersCod As String, ByVal pnMontoSol As Double, ByVal pnMonedaSol As Integer) As Boolean
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_VerificaSuperaUmbralSector '" & psPersCod & "'," & pnMontoSol & "," & pnMonedaSol

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set rs = oConecta.CargaRecordSet(sSql)
If Not (rs.EOF And rs.BOF) Then
    If rs!Resultado <> "" Then
        VerificaSuperaUmbralSectorEcon = True
    End If
End If
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Sub InsertarSolicitudAutorizacionRiesgos(ByVal psCtaCod As String, ByVal pnMontoSol As Double)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_ins_SolicitudAutorizacionRiesgos '" & psCtaCod & "'," & pnMontoSol

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Call oConecta.Ejecutar(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Sub
Public Sub EliminarSolicitudAutorizacionRiesgos(ByVal psCtaCod As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_del_SolicitudAutorizacionRiesgos '" & psCtaCod & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Call oConecta.Ejecutar(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Sub
Public Function VerificaSolicitudAutorizacionRiesgos(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_VerificaSolicitudAutorizacionRiesgos '" & psPersCod & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set VerificaSolicitudAutorizacionRiesgos = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Function RecuperaSolicitudAutorizacionRiesgos() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_RecuperaSolicitudAutorizacionRiesgos"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaSolicitudAutorizacionRiesgos = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
Public Sub ActualizarSolicitudAutorizacionRiesgos(ByVal psCtaCod As String, ByVal psGlosa As String, ByVal pnEstado As Integer)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
'0 Solicitado
'1 Autorizado
'2 Rechazado
sSql = "exec stp_upd_ActualizarSolicitudAutorizacionRiesgos '" & psCtaCod & "','" & psGlosa & "'," & pnEstado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Call oConecta.Ejecutar(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Sub
Public Function RecuperaAutorizacionRechazoRiesgos(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
sSql = "exec stp_sel_RecuperaAutorizacionRechazoRiesgos '" & psCtaCod & "'"

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set RecuperaAutorizacionRechazoRiesgos = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
End Function
'END JUEZ *******************************************************************
