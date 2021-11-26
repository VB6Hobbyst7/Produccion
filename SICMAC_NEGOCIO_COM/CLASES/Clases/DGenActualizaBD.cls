VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DGenActualizaBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oConn As DConecta
Dim Sql As String

Private Sub Class_Initialize()
Dim oIni As ClasIni
Dim sConn As String

    Set oIni = New ClasIni
    sConn = oIni.CadenaConexion
    Set oIni = Nothing
    
Set oConn = New DConecta
If oConn.AbreConexion(sConn) = False Then
    Call RaiseError(MyUnhandledError, "DPigContrato:Initialize. Error en Conexion a Base de datos")
End If
End Sub
Private Sub Class_Terminate()
    oConn.CierraConexion
    Set oConn = Nothing
End Sub

Public Sub dInsertCodigoPostal(ByVal pnCodPostal, ByVal psDescCodPostal, ByVal psCodZon As String)

On Error GoTo dErrorCodigoPostal

    Sql = "INSERT CODIGOPOSTAL (nCodPostal, cDesCodPostal, cCodZon) " _
        & "VALUES (" & pnCodPostal & ",'" & psDescCodPostal & "'," & psCodZon & ")"

    oConn.Ejecutar Sql

    Exit Sub
    
dErrorCodigoPostal:
    Err.Raise Err.Number, "Error en <<dInsertCodigoPostal>>" + Err.Description, "Aviso"

End Sub

Public Function dObtieneZona(ByVal psTipoZon As String, Optional ByVal psDep As String = "@", _
                            Optional ByVal psProv As String = "@") As Recordset
    
On Error GoTo dErrorObtieneZona

    Sql = "SELECT cUbiGeoCod, cUbiGeoDescripcion FROM UBICACIONGEOGRAFICA " _
        & "WHERE SUBSTRING(cUbiGeoCod,1,1) = '" & psTipoZon & "'" _
        
    If psDep <> "@" Then
        Sql = Sql & " AND SUBSTRING(cUbiGeoCod,2,2) = '" & psDep & "'"
    End If
    
    If psProv <> "@" Then
        Sql = Sql & " AND SUBSTRING(cUbiGeoCod, 4,2) = '" & psProv & "'"
    End If

    Set dObtieneZona = oConn.CargaRecordSet(Sql)
        
    Exit Function
    
dErrorObtieneZona:
    Err.Raise Err.Number, "Error en <<dObtieneZona>>" + Err.Description, "Aviso"

End Function

Public Function dGeneraCodigoPostal() As Long
Dim Rs As Recordset

    On Error GoTo dErrorGeneraCodigoPostal
    Sql = "SELECT ISNULL(MAX(nCodPostal),0) + 1 as CodPostal FROM CODIGOPOSTAL "
    
    Set Rs = oConn.CargaRecordSet(Sql)

    If Not Rs.EOF And Not Rs.BOF Then
        dGeneraCodigoPostal = Rs!CodPostal
    Else
        dGeneraCodigoPostal = 1
    End If
    
    Set Rs = Nothing
    
    Exit Function
    
dErrorGeneraCodigoPostal:
    Err.Raise Err.Number, "Error en <<dGeneraCodigoPostal>>" + Err.Description, "Aviso"
    
End Function
Public Function GetCIIU() As ADODB.Recordset
Dim Rs As New ADODB.Recordset
Dim Sql As String
Sql = " Select cCIIUcod, cCIIUdescripcion, " & _
      " (Select cConsDescripcion from Constante where nConsCod=1013 and nConsValor=nSector ) Sector   " & _
      " From CIIU "
Set Rs = oConn.CargaRecordSet(Sql)
Set GetCIIU = Rs
Set Rs = Nothing
End Function
Public Function ExisteCIIUGiro(ByVal psDescripcion As String, ByVal psCIIU As String) As Boolean
Dim Rs As New ADODB.Recordset
Dim Sql As String

Sql = "select * from CIIUGiroNeg "
Sql = Sql & " Where cCIIUCod ='" & Trim(psCIIU) & "' And cDescripcion = '" & psDescripcion & "'"
Set Rs = oConn.CargaRecordSet(Sql)
If Rs.EOF And Rs.BOF Then
    ExisteCIIUGiro = False
Else
    ExisteCIIUGiro = True
End If
Set Rs = Nothing

End Function

Public Function ExisteCIIU(ByVal psCodCIIU As String) As Boolean
Dim Rs As New ADODB.Recordset
Dim Sql As String
Sql = "Select * from CIIU where  cCIIUcod='" & psCodCIIU & "'"
Set Rs = oConn.CargaRecordSet(Sql)
If Rs.EOF And Rs.BOF Then
    ExisteCIIU = False
Else
    ExisteCIIU = True
End If
Set Rs = Nothing
End Function
Public Sub dUpdateCIIUGiro(ByVal psCIIU As String, ByVal psDescripcion1 As String, ByVal psDescripcion2 As String)
Dim Sql As String
Sql = " Update CIIUGiroNeg "
Sql = Sql & " Set cDescripcion='" & psDescripcion2 & "' "
Sql = Sql & " Where cCIIUCod ='" & psCIIU & "' And cDescripcion = '" & psDescripcion1 & "'"
oConn.Ejecutar (Sql)
End Sub

Public Sub dUpdateCIIU(ByVal psCodCIIU As String, ByVal psDescripcion As String, _
ByVal pnSector As Integer, ByVal pdFecSis As Date, ByVal psUser As String, ByVal psCodAge As String)


Dim Sql As String
Dim sUltMov As String
Dim ObjMov As DMov
Set ObjMov = New DMov
sUltMov = ObjMov.GeneraMovNro(pdFecSis, psCodAge, psUser)


Sql = " Update CIIU"
Sql = Sql & " set cCIIUdescripcion='" & psDescripcion & "' "
Sql = Sql & " , cUltimaActualizacion ='" & sUltMov & "' "
Sql = Sql & " , nSector ='" & pnSector & "' "
Sql = Sql & " where cCIIUcod='" & psCodCIIU & "'"

oConn.BeginTrans
oConn.Ejecutar (Sql)
oConn.CommitTrans
Set ObjMov = Nothing
End Sub
Public Sub dInsertaCIIUGiro(ByVal psCIIU As String, ByVal psDescripcion As String)
Dim Sql As String
Dim Rs As New ADODB.Recordset
Sql = "select isnull(max(cGiroNegCod),0) Nro from CIIUGiroNeg"
Set Rs = oConn.CargaRecordSet(Sql)
Sql = " Insert Into CIIUGiroNeg " & _
      " (cGiroNegCod,cDescripcion,cCIIUCod) values " & _
      " ('" & Format(Val(Rs!Nro) + 1, "00000") & "','" & psDescripcion & "','" & psCIIU & "')"
oConn.BeginTrans
oConn.Ejecutar (Sql)
oConn.CommitTrans
Set Rs = Nothing
End Sub
Public Function GetGiroDescripcion(ByVal psCodGiro As String) As String
Dim Sql As String
Dim Rs As New ADODB.Recordset

Sql = " Select cDescripcion, cGiroNegCod, cCIIUCod from CIIUGIRONEG "
Sql = Sql & " where cGiroNegCod='" & psCodGiro & "'"
Set Rs = oConn.CargaRecordSet(Sql)
If Rs.EOF And Rs.BOF Then
    GetGiroDescripcion = ""
Else
    GetGiroDescripcion = Rs!cDescripcion
End If
Set Rs = Nothing
End Function

Public Sub dInsertaCIIU(ByVal psCodCIIU As String, _
 ByVal psDescripcion As String, ByVal pnSector As Integer, _
ByVal pdFecSis As Date, ByVal psUser As String, ByVal psCodAge As String)

Dim Sql As String
Dim sUltMov As String
Dim ObjMov As DMov
Set ObjMov = New DMov
sUltMov = ObjMov.GeneraMovNro(pdFecSis, psCodAge, psUser)



Sql = " Insert CIIU"
Sql = Sql & " (cCIIUcod,cCIIUdescripcion,cUltimaActualizacion,nSector)"
Sql = Sql & " Values"
Sql = Sql & " ('" & psCodCIIU & "','" & psDescripcion & "','" & Format(pdFecSis, "dd/mm/yyyy") & "',"
Sql = Sql & pnSector & ")"

oConn.BeginTrans
oConn.Ejecutar (Sql)
oConn.CommitTrans
Set ObjMov = Nothing
End Sub
Public Function GetCIIUGiro(ByVal psCIIU As String) As ADODB.Recordset
Dim Rs As New ADODB.Recordset
Dim Sql As String
Sql = " select * from CIIUGiroNeg WHERE cCIIUCod = '" & psCIIU & "'"
Set Rs = oConn.CargaRecordSet(Sql)
Set GetCIIUGiro = Rs
Set Rs = Nothing
End Function

Public Function dObtieneCodigosPostal() As Recordset
Dim Rs As Recordset

    Sql = "SELECT nCodPostal, cDesCodPostal, cCodZon, cUbiGeoDescripcion " _
            & " FROM CODIGOPOSTAL CP INNER JOIN UBICACIONGEOGRAFICA UG ON CP.cCodZon = UG.cUbiGeoCod " _
            & "ORDER BY nCodPostal"
            
    Set Rs = oConn.CargaRecordSet(Sql)
    Set dObtieneCodigosPostal = Rs
    Set Rs = Nothing
    
End Function

Public Sub dUpdateCodigoPostal(ByVal psCodPostal As Long, ByVal psDesCodPostal As String, psCodZon As String)

    On Error GoTo dErrorUpdateCodigoPostal
    Sql = "UPDATE CODIGOPOSTAL SET cDesCodPostal ='" & psDesCodPostal & "', cCodZon= '" & psCodZon & "'" _
        & "WHERE nCodPostal = " & psCodPostal
        
   oConn.Ejecutar Sql
   
   Exit Sub
   
dErrorUpdateCodigoPostal:
    Err.Raise Err.Number, "Error en <<dUpdateCodigoPostal>>" + Err.descripction, "Aviso"

End Sub

Public Sub dDeleteCodigoPostal(ByVal psCodPostal As Long)

    On Error GoTo dErrorDeleteCodigoPostal
    Sql = "DELETE CODIGOPOSTAL WHERE nCodPostal = " & psCodPostal
    
    oConn.Ejecutar Sql
    
    Exit Sub
    
dErrorDeleteCodigoPostal:
    Err.Raise Err.Number, "Error en <<dDeleteCodigoPostal>>" + Err.Description, "Aviso"
    
End Sub
