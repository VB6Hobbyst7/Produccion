VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogVehiculo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Function GetConductores() As ADODB.Recordset
Dim Co As DConecta
Dim SQL As String
Dim Rs As ADODB.Recordset

SQL = " Select P.cPersCod, cPersNombre,"
SQL = SQL & " (select cPersIDnro from persID where cPersIDTpo = '7' and cPersCod = LC.cPersCod) Brevete,"
SQL = SQL & " (select cPersIDnro from persID where cPersIDTpo = '1' and cPersCod = LC.cPersCod) DNI,"
SQL = SQL & " cEstado"
SQL = SQL & " from LogisticaConductor LC"
SQL = SQL & " Inner Join Persona P on P.cPersCod = LC.cPersCod"

Set Co = New DConecta
Co.AbreConexion
Set Rs = Co.CargaRecordSet(SQL)
Set GetConductores = Rs
Co.CierraConexion
Set Rs = Nothing
Set Co = Nothing
End Function

Function GetBrevete(ByVal psPersCod As String) As String
Dim Co As DConecta
Dim SQL As String
Dim Rs As ADODB.Recordset
SQL = "select cPersIDnro from persID where cPersIDTpo = '7' and cPersCod='" & psPersCod & "'"
Set Co = New DConecta
Co.AbreConexion
Set Rs = Co.CargaRecordSet(SQL)
Co.CierraConexion
If Not (Rs.EOF And Rs.BOF) Then
    GetBrevete = Rs!cPersIDnro
Else
    GetBrevete = ""
End If
Set Rs = Nothing
Set Co = Nothing
End Function

Sub InsertRegVehiculoKm(ByVal pdFecha As Date, ByVal pnKm As Long, ByVal pnKmFin As Long, _
ByVal psBSCod As String, ByVal psBSSerie As String, _
ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String)

Dim Sql1 As String
Dim Sql2 As String
Dim Co As DConecta
Dim DM As DMov
Dim sMovNro As String
Set DM = New DMov
Set Co = New DConecta

sMovNro = DM.GeneraMovNro(pdFecSis, psCodAge, psCodUser)

Sql1 = "Insert LogisticaKilometraje"
Sql1 = Sql1 & " (dFecha,nKm,nKmF,cBSCod,cBSSerie) values"
Sql1 = Sql1 & " ('" & Format(pdFecha, "YYYY/MM/DD") & "','" & pnKm & "'," & pnKmFin & ","
Sql1 = Sql1 & "'" & psBSCod & "','" & psBSSerie & "')"

Sql2 = "Insert Mov"
Sql2 = Sql2 & "(cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) values"
Sql2 = Sql2 & "('" & sMovNro & "','902053','Registro de Kilometraje Vehiculo',13,0)"

Co.AbreConexion
Co.BeginTrans
    Co.CargaRecordSet (Sql2)
    Co.CargaRecordSet (Sql1)
Co.CommitTrans
Co.CierraConexion
End Sub

Function GetAgencias() As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Dim SQL As String
    SQL = "select cAgeCod, cAgeDescripcion from agencias where nEstado = 1 Order by cAgeCod"
    Set Co = New DConecta
    Co.AbreConexion
    Set Rs = Co.CargaRecordSet(SQL)
    Co.CierraConexion
Set GetAgencias = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

Function GetVehiculoPapeleta(ByVal psBSCod As String, ByVal psBSSerie As String) As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Dim SQL As String
    
    SQL = "select cDescripcion, nMonto, dFecha from LogisticaPapeleta"
    SQL = SQL & " where cBSCod =  '" & psBSCod & "' and cBSSerie = '" & psBSSerie & "'"
    SQL = SQL & " and cFlag is null"
    
    Set Co = New DConecta
    Co.AbreConexion
    Set Rs = Co.CargaRecordSet(SQL)
    Co.CierraConexion
Set GetVehiculoPapeleta = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

Function GetVehiculoIncidencia(ByVal psBSCod As String, ByVal psBSSerie As String, Optional psFecha As String = "") As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Dim SQL As String
Dim XQL As String

If Len(psFecha) > 0 And IsDate(psFecha) Then
   XQL = " and dFecha = '" & Format(psFecha, "YYYYMMDD") & "'"
Else
   XQL = ""
End If


    SQL = "select dFecha, cDescripcion,"
    SQL = SQL & " (select LI.cDescripcion from LogisticaTipoincidencia LI where LI.nTipoIncidencia = L.nTipoIncidencia) TipoIncidencia,"
    SQL = SQL & " cLugar"
    SQL = SQL & " from LogisticaIncidencia L"
    SQL = SQL & " Where cFlag Is Null"
    SQL = SQL & " and cBSCod =  '" & psBSCod & "' and cBSSerie = '" & psBSSerie & "' " & XQL
    SQL = SQL & " Order by dFecha desc"
    
    Set Co = New DConecta
    Co.AbreConexion
    Set Rs = Co.CargaRecordSet(SQL)
    Co.CierraConexion
    
Set GetVehiculoIncidencia = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

Sub LiberaAsignacionVehiculo(ByVal psBSCod As String, ByVal psBSSerie As String, ByVal pdFecha As Date, ByVal vPersCod As String)
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Co As New DConecta
Co.AbreConexion
    SQL = ""
    'Sql = " Select cPersCod from LogisticaAsigna "
    'Sql = Sql & " where cBSCod='" & psBSCod & "' and cBSSerie='" & psBSSerie & "'"
    'Sql = Sql & " and cflag is Null and dFechaFin is null and cEstado='1'"
    'Set Rs = Co.CargaRecordSet(Sql)
    
    'Sql = " Update LogisticaAsigna Set dFechaFin='" & Format(pdFecha, "YYYY/MM/DD") & "'"
    'Sql = Sql & " where cBSCod='" & psBSCod & "' and cBSSerie='" & psBSSerie & "'"
    'Sql = Sql & " and cflag is Null and dFechaFin is null and cEstado='1'"
    
    SQL = " Update LogisticaAsigna Set cEstado='1', dFechaFin='" & Format(pdFecha, "YYYY/MM/DD") & "'"
    SQL = SQL & " where cBSCod='" & psBSCod & "' and cBSSerie='" & psBSSerie & "'"
    SQL = SQL & " and cflag is Null and cEstado='2'"
    
    Co.CargaRecordSet (SQL)
    
    SQL = "Update LogisticaConductor set cEstado='1' where cPersCod='" & vPersCod & "'"
    Co.CargaRecordSet (SQL)
    
    SQL = "Update LogisticaVehiculo set nEstado = 1 where "
    SQL = SQL & " cBSCod='" & psBSCod & "' and cBSSerie='" & psBSSerie & "'"
    Co.CargaRecordSet (SQL)
    
Co.CierraConexion
Set Rs = Nothing
Set Co = Nothing
End Sub

Sub AsignacionVehiculo(ByVal psPersCod As String, ByVal psBSCod As String, ByVal psBSSerie As String, _
                       ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, ByVal chk As Integer)
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Dim Sql0 As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql3 As String

Set Co = New DConecta
Co.AbreConexion
    
'    Sql0 = "select * from LogisticaAsigna where cEstado = '2'"
'    Sql0 = Sql0 & " and cBSCod='" & psBSCod & "' and cBSSerie='" & psBSSerie & "'"
'    Sql0 = Sql0 & " and cflag is null"
'    Set Rs = Co.CargaRecordSet(Sql0)
'    While Not Rs.EOF
'        Sql0 = "Update LogisticaConductor Set cEstado='1' where cPersCod='" & Rs!cPersCod & "'"
'        Co.CargaRecordSet (Sql0)
'        Rs.MoveNext
'    Wend

'If chk = 1 Then
 '   Sql0 = "Update LogisticaAsigna set cEstado = '1' where cEstado = '2'"
 '   Sql0 = Sql0 & " and cBSCod='" & psBSCod & "' and cBSSerie='" & psBSSerie & "'"
 '   Sql0 = Sql0 & " and cFlag is null"
        
    Sql1 = "Insert LogisticaAsigna"
    Sql1 = Sql1 & "(cPersCod,cBSCod,cBSSerie,cEstado,dFecha,dFechaFin) values"
    Sql1 = Sql1 & "('" & psPersCod & "','" & psBSCod & "',"
    Sql1 = Sql1 & "'" & psBSSerie & "','2','" & Format(pdFechaIni, "YYYY/MM/DD") & "','" & Format(pdFechaFin, "YYYY/MM/DD") & "')"
    
    Sql3 = "Update LogisticaConductor set cEstado='2' where cPersCod='" & psPersCod & "'"
    
    Sql2 = "Update LogisticaVehiculo set nEstado = 2 where "
    Sql2 = Sql2 & " cBSCod='" & psBSCod & "' and cBSSerie='" & psBSSerie & "'"
    'Co.CargaRecordSet (Sql0)
    Co.CargaRecordSet (Sql1)
    Co.CargaRecordSet (Sql3)
    Co.CargaRecordSet (Sql2)
'Else
'    Sql2 = "Update LogisticaVehiculo set nEstado = 1 where "
'    Sql2 = Sql2 & " cBSCod='" & psBSCod & "' and cBSSerie='" & psBSSerie & "'"
'    Co.CargaRecordSet (Sql2)
'End If
Co.CierraConexion
Set Rs = Nothing
Set Co = Nothing
End Sub

Function GetConductorLibre() As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Dim SQL As String
SQL = "select LC.cPersCod, cPersNombre"
SQL = SQL & " from LogisticaConductor LC"
SQL = SQL & " Inner Join Persona P on P.cPersCod = LC.cPersCod"
SQL = SQL & " where cEstado = '1' Order by cPersNombre"
    Set Co = New DConecta
    Co.AbreConexion
    Set Rs = Co.CargaRecordSet(SQL)
    Co.CierraConexion
Set GetConductorLibre = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

Function GetVehiculoAsignacion(ByVal psBSCod As String, ByVal psBSSerie As String) As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Dim SQL As String
    SQL = ""
    SQL = "Select LA.*, P.cPersNombre"
    SQL = SQL & " from LogisticaAsigna LA"
    SQL = SQL & " Inner Join Persona P on P.cPersCod = LA.cPersCod"
    SQL = SQL & " Where LA.cFlag Is Null and "
    SQL = SQL & " cBSCod =  '" & psBSCod & "' and cBSSerie = '" & psBSSerie & "'"
    SQL = SQL & " order by LA.dFecha,LA.cEstado desc"
    Set Co = New DConecta
    Co.AbreConexion
    Set Rs = Co.CargaRecordSet(SQL)
    Co.CierraConexion
Set GetVehiculoAsignacion = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

Function GetVehiculoSoat(ByVal psBSCod As String, ByVal psBSSerie As String) As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Dim SQL As String
    SQL = " select dInicio,dVencimiento,nMonto from LogisticaVehiculoSoat"
    SQL = SQL & " where cBSCod =  '" & psBSCod & "' and cBSSerie = '" & psBSSerie & "'"
    SQL = SQL & " and cFlag is null"
    Set Co = New DConecta
    Co.AbreConexion
    Set Rs = Co.CargaRecordSet(SQL)
    Co.CierraConexion
Set GetVehiculoSoat = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

Sub InsertRegVehiculoIncidencia(ByVal pdFecha As Date, ByVal psDescripcion As String, _
ByVal pnTipoIncidencia As Integer, ByVal psLugar As String, _
ByVal psBSCod As String, ByVal psBSSerie As String, _
ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String)

Dim Sql1 As String
Dim Sql2 As String
Dim Co As DConecta
Dim DM As DMov
Dim sMovNro As String
Set DM = New DMov
Set Co = New DConecta

sMovNro = DM.GeneraMovNro(pdFecSis, psCodAge, psCodUser)

Sql1 = "Insert LogisticaIncidencia"
Sql1 = Sql1 & " (dFecha,cDescripcion,nTipoIncidencia,cLugar,cBSCod,cBSSerie) values"
Sql1 = Sql1 & " ('" & Format(pdFecha, "YYYY/MM/DD") & "','" & psDescripcion & "'," & pnTipoIncidencia & ","
Sql1 = Sql1 & "'" & psLugar & "','" & psBSCod & "','" & psBSSerie & "')"


Sql2 = "Insert Mov"
Sql2 = Sql2 & "(cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) values"
Sql2 = Sql2 & "('" & sMovNro & "','902058','Registro de Incidencia',13,0)"

Co.AbreConexion
Co.BeginTrans
    Co.CargaRecordSet (Sql2)
    Co.CargaRecordSet (Sql1)
Co.CommitTrans
Co.CierraConexion
End Sub

Sub InsertRegVehiculoPapeleta(ByVal pdFecha As Date, ByVal psDescripcion As String, _
ByVal pnMonto As Currency, ByVal psBSCod As String, ByVal psBSSerie As String, _
ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String)


Dim Sql1 As String
Dim Sql2 As String
Dim Co As DConecta
Dim DM As DMov
Dim sMovNro As String
Set DM = New DMov
Set Co = New DConecta

sMovNro = DM.GeneraMovNro(pdFecSis, psCodAge, psCodUser)

Sql1 = "Insert LogisticaPapeleta"
Sql1 = Sql1 & " (cDescripcion, nMonto, dFecha,cBSCod,cBSSerie) values"
Sql1 = Sql1 & " ('" & psDescripcion & "'," & pnMonto & ",'" & Format(pdFecha, "YYYY/MM/DD") & "',"
Sql1 = Sql1 & "'" & psBSCod & "','" & psBSSerie & "')"

Sql2 = "Insert Mov"
Sql2 = Sql2 & "(cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) values"
Sql2 = Sql2 & "('" & sMovNro & "','902057','Registro de Papeleta',13,0)"

Co.AbreConexion
Co.BeginTrans
    Co.CargaRecordSet (Sql2)
    Co.CargaRecordSet (Sql1)
Co.CommitTrans
Co.CierraConexion
End Sub

Sub InsertRegVehiculoSoat(ByVal pdInicio As Date, ByVal pdVencimiento As Date, _
ByVal psBSCod As String, ByVal psBSSerie As String, ByVal pnMonto As Currency, _
ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String)

Dim Sql1 As String
Dim Sql2 As String
Dim Co As DConecta
Dim DM As DMov
Dim sMovNro As String
Set DM = New DMov
Set Co = New DConecta

sMovNro = DM.GeneraMovNro(pdFecSis, psCodAge, psCodUser)

Sql1 = "Insert LogisticaVehiculoSoat"
Sql1 = Sql1 & " (dInicio,dVencimiento,cBSCod,cBSSerie,nMonto) values"
Sql1 = Sql1 & " ('" & Format(pdInicio, "YYYY/MM/DD") & "','" & Format(pdVencimiento, "YYYY/MM/DD") & "', "
Sql1 = Sql1 & "'" & psBSCod & "','" & psBSSerie & "'," & pnMonto & ")"

Sql2 = "Insert Mov"
Sql2 = Sql2 & "(cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) values"
Sql2 = Sql2 & "('" & sMovNro & "','902052','Registro de SOAT',13,0)"

Co.AbreConexion
Co.BeginTrans
    Co.CargaRecordSet (Sql2)
    Co.CargaRecordSet (Sql1)
Co.CommitTrans
Co.CierraConexion
End Sub

Sub InsertRegVehiculoCarga(ByVal pdFecha As Date, ByVal psDescripcion As String, _
ByVal psDestinoAge As String, ByVal psDestinoDesc As String, _
ByVal psBSCod As String, ByVal psBSSerie As String, _
ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String)

Dim Sql1 As String
Dim Sql2 As String
Dim Co As DConecta
Dim DM As DMov
Dim sMovNro As String
Set DM = New DMov
Set Co = New DConecta

sMovNro = DM.GeneraMovNro(pdFecSis, psCodAge, psCodUser)

Sql1 = "Insert LogisticaCarga"
Sql1 = Sql1 & " (dFecha,cDescripcion,cDestinoAge,cDestinoDesc,cBSCod,cBSSerie) values"
Sql1 = Sql1 & " ('" & Format(pdFecha, "YYYY/MM/DD") & "','" & psDescripcion & "','" & psDestinoAge & "','" & psDestinoDesc & "',"
Sql1 = Sql1 & "'" & psBSCod & "','" & psBSSerie & "')"

Sql2 = "Insert Mov"
Sql2 = Sql2 & "(cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) values"
Sql2 = Sql2 & "('" & sMovNro & "','902054','Registro de Carga de Vehiculo',13,0)"

Co.AbreConexion
Co.BeginTrans
    Co.CargaRecordSet (Sql2)
    Co.CargaRecordSet (Sql1)
Co.CommitTrans
Co.CierraConexion
End Sub

Sub InsertRegConductor(ByVal psPersCod As String, ByVal psBrevete As String, _
ByVal psEstado As String, ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String)

Dim Sql1 As String
Dim Sql2 As String
Dim Co As DConecta
Dim DM As DMov
Dim sMovNro As String
Set DM = New DMov
Set Co = New DConecta

sMovNro = DM.GeneraMovNro(pdFecSis, psCodAge, psCodUser)

Sql1 = "Insert LogisticaConductor"
Sql1 = Sql1 & " (cPersCod,cBrevete,cEstado) values"
Sql1 = Sql1 & " ('" & psPersCod & "','" & psBrevete & "','" & psEstado & "')"

Sql2 = "Insert Mov"
Sql2 = Sql2 & "(cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) values"
Sql2 = Sql2 & "('" & sMovNro & "','902055','Registro de Conductor',13,0)"

Co.AbreConexion
Co.BeginTrans
    Co.CargaRecordSet (Sql2)
    Co.CargaRecordSet (Sql1)
Co.CommitTrans
Co.CierraConexion
End Sub

Sub InsertTipoIncidencia(ByVal psDescripcion As String)
Dim SQL As String
Dim Co As DConecta
Set Co = New DConecta
SQL = "Insert LogisticaTipoIncidencia (cDescripcion) values"
SQL = SQL & "('" & psDescripcion & "')"

Co.AbreConexion
Co.CargaRecordSet (SQL)
Co.CierraConexion
End Sub

Sub InsertaRegVehiculo(ByVal pscBSCod As String, ByVal pscBSSerie As String, ByVal pscModelo As String, _
ByVal pnnAñoFab As Integer, ByVal pnnEstado As Integer, _
ByVal pnnCombustible As Integer, ByVal pscNroMotor As String, _
ByVal pscPlaca As String, ByVal pnnTipoVehiculo As Integer, ByVal pnnMarca As Integer, _
ByVal pnnColor As Integer, ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String)

Dim Sql1 As String
Dim Sql2 As String
Dim Co As DConecta
Dim DM As DMov
Dim sMovNro As String
Set DM = New DMov
Set Co = New DConecta

sMovNro = DM.GeneraMovNro(pdFecSis, psCodAge, psCodUser)

Sql1 = " Insert LogisticaVehiculo "
Sql1 = Sql1 & " (cBSCod,cBSSerie,cModelo,nAñoFab,nEstado,nCombustible,cNroMotor,cPlaca,nTipoVehiculo,nMarca,nColor) values "
Sql1 = Sql1 & " ('" & pscBSCod & "','" & pscBSSerie & "','" & pscModelo & "'," & pnnAñoFab & ","
Sql1 = Sql1 & pnnEstado & "," & pnnCombustible & ",'" & pscNroMotor & "','" & pscPlaca & "',"
Sql1 = Sql1 & pnnTipoVehiculo & "," & pnnMarca & "," & pnnColor & ")"

Sql2 = "Insert Mov"
Sql2 = Sql2 & "(cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) values"
Sql2 = Sql2 & "('" & sMovNro & "','902051','Registro de Vehiculo',13,0)"

Co.AbreConexion
Co.BeginTrans
    Co.CargaRecordSet (Sql2)
    Co.CargaRecordSet (Sql1)
Co.CommitTrans
Co.CierraConexion
End Sub

Function GetVehiculoKm(ByVal psBSCod As String, ByVal psBSSerie As String, Optional psFecha As String = "") As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Dim SQL As String
Dim XQL As String

If Len(psFecha) > 0 And IsDate(psFecha) Then
   XQL = " and dFecha = '" & Format(psFecha, "YYYYMMDD") & "'"
Else
   XQL = ""
End If

    SQL = " select * from LogisticaKilometraje"
    SQL = SQL & " where cBSCod =  '" & psBSCod & "' and cBSSerie = '" & psBSSerie & "'"
    SQL = SQL & " and cFlag is null " & XQL
    
    Set Co = New DConecta
    Co.AbreConexion
    Set Rs = Co.CargaRecordSet(SQL)
    Co.CierraConexion
Set GetVehiculoKm = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

Function GetVehiculoCarga(ByVal psBSCod As String, ByVal psBSSerie As String, Optional psFecha As String = "") As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Dim SQL As String
Dim XQL As String

If Len(psFecha) > 0 And IsDate(psFecha) Then
   XQL = " and dFecha = '" & Format(psFecha, "YYYYMMDD") & "'"
Else
   XQL = ""
End If

    SQL = "select dFecha,cDescripcion,cDestinoAge,"
    SQL = SQL & " (select  cAgeDescripcion from agencias where cAgeCod = cDestinoAge) NombreAgencia,cDestinoDesc"
    SQL = SQL & " from LogisticaCarga where cFlag is null  "
    SQL = SQL & " and cBSCod =  '" & psBSCod & "' and cBSSerie = '" & psBSSerie & "' " & XQL
    
    Set Co = New DConecta
    Co.AbreConexion
    Set Rs = Co.CargaRecordSet(SQL)
    Co.CierraConexion
Set GetVehiculoCarga = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

Function GetColorV() As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Dim SQL As String
    SQL = "select cColor, nColor from Logisticacolor"
    Set Co = New DConecta
    Co.AbreConexion
    Set Rs = Co.CargaRecordSet(SQL)
    Co.CierraConexion
Set GetColorV = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

Function GetMarcaV() As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Dim SQL As String
    SQL = "SELECT CMARCA, NMARCA FROM LOGISTICAMARCAV"
    Set Co = New DConecta
    Co.AbreConexion
    Set Rs = Co.CargaRecordSet(SQL)
    Co.CierraConexion
Set GetMarcaV = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

Function GetTipoIncidencia() As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Dim SQL As String
    SQL = "select nTipoincidencia, cDescripcion from LogisticaTipoIncidencia"
    Set Co = New DConecta
    Co.AbreConexion
    Set Rs = Co.CargaRecordSet(SQL)
    Co.CierraConexion
Set GetTipoIncidencia = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

Function GetTipoVehiculo() As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Dim SQL As String
    SQL = "SELECT cDesCripcion, nTipoVehiculo FROM LogisticaTipoVehiculo"
    Set Co = New DConecta
    Co.AbreConexion
    Set Rs = Co.CargaRecordSet(SQL)
    Co.CierraConexion
Set GetTipoVehiculo = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

Function GetVehiculoSeries(ByVal pnAño As Integer) As ADODB.Recordset
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Set Co = New DConecta
    SQL = "Select  cBSCod,cSerie, cBSCod+cSerie Codigo,cDescripcion from bsactivofijo"
    SQL = SQL & " where not cBSCod+cSerie in (select  cBSCod+cBSSerie  from LogisticaVehiculo where cFlag is null )"
    SQL = SQL & " and  cbscod like '11203%' and nanio = " & pnAño & " and dBaja is null"
    SQL = SQL & " Order by cBSCod,cSerie"
    Co.AbreConexion
    Set Rs = Co.CargaRecordSet(SQL)
    Co.CierraConexion
    Set GetVehiculoSeries = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

Function GetVehiculos(Optional nTipo As Integer = 0) As ADODB.Recordset
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Set Co = New DConecta

SQL = "select cBSCod, cBSSerie, cModelo, nAñoFab, nEstado,"
SQL = SQL & " case when ncombustible = 0 then 'PETROLEO'"
SQL = SQL & "  when ncombustible = 1 then 'GASOLINA'"
SQL = SQL & "  when ncombustible = 2 then 'GAS' END Combustible,"
SQL = SQL & " cNroMotor, cPlaca,"
SQL = SQL & " (select cDescripcion from LogisticaTipoVehiculo where nTipoVehiculo = LV.nTipoVehiculo) TipoV,"
SQL = SQL & " (select cMarca from logisticaMarcav where nMarca = LV.nMarca) Marcar,"
SQL = SQL & " (select cColor from logisticacolor where nColor = LV.nColor) Color"
SQL = SQL & " from LogisticaVehiculo LV where cFlag is null"
If nTipo <> 0 Then
    SQL = SQL & " and nTipoVehiculo=" & nTipo
End If
    Co.AbreConexion
    Set Rs = Co.CargaRecordSet(SQL)
    Co.CierraConexion
    Set GetVehiculos = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

Function VerificaConductor(ByVal psPersCod As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Co As DConecta
SQL = "select * from LogisticaConductor where cPersCod = '" & psPersCod & "'"

Set Co = New DConecta
Co.AbreConexion
Set Rs = Co.CargaRecordSet(SQL)
Co.CierraConexion

If Not (Rs.EOF And Rs.EOF) Then
    VerificaConductor = True
Else
    VerificaConductor = False
End If

Set Co = Nothing
Set Rs = Nothing
End Function

Function GetDatoVehiculo(ByVal psPlaca As String) As ADODB.Recordset
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Set Co = New DConecta

SQL = "select cBSCod, cBSSerie,cModelo,nEstado,cPlaca,"
SQL = SQL & " (select cDescripcion from LogisticaTipoVehiculo where nTipoVehiculo = LV.nTipoVehiculo) TipoV,"
SQL = SQL & " (select cMarca from logisticaMarcav where nMarca = LV.nMarca) Marcar"
SQL = SQL & " from LogisticaVehiculo  LV where cPlaca = '" & psPlaca & "'"

    Co.AbreConexion
    Set Rs = Co.CargaRecordSet(SQL)
    Co.CierraConexion
    Set GetDatoVehiculo = Rs
Set Rs = Nothing
Set Co = Nothing
End Function


