VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogVehiculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim Sql As String

Function ListaVehiculos(ByVal pnTipo As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim Co As DConecta, cCondicion As String
Set Co = New DConecta

If pnTipo > 0 Then
   cCondicion = " WHERE r.nTipo=" & pnTipo & " "
Else
   cCondicion = ""
End If

Sql = "select r.nVehiculoCod,t.cTipoVehiculo,m.cMarca,c.cColor, r.cPlaca, r.nEstado,e.cEstado" & _
"  from LogVehiculo r " & _
" inner join (select nConsValor AS nTipoVehiculo,cConsDescripcion as cTipoVehiculo from Constante where nConsCod=9026 and nconscod<>nconsvalor) t on r.nTipoVehiculo = t.nTipoVehiculo " & _
" inner join (select nConsValor AS nMarca,cConsDescripcion as cMarca from Constante where nConsCod=9022 and nconscod<>nconsvalor) m on r.nMarca = m.nMarca " & _
" inner join (select nConsValor AS nColor,cConsDescripcion as cColor from Constante where nConsCod=9023 and nconscod<>nconsvalor) c on r.nColor = c.nColor " & _
" inner join (select nConsValor AS nEstado,cConsDescripcion as cEstado from Constante where nConsCod=9020 and nconscod<>nconsvalor) e on r.nEstado = e.nEstado " & _
"   " + cCondicion

'left outer join (select * from LogVehiculoAsignacionMov) v on r.nAsignacionNro = v.nAsignacionNro
If Co.AbreConexion Then
   Set rs = Co.CargaRecordSet(Sql)
   Co.CierraConexion
End If

Set ListaVehiculos = rs
Set rs = Nothing
Set Co = Nothing
End Function

Function ListaConductores() As ADODB.Recordset
Dim rs As New ADODB.Recordset, Co As DConecta
Set Co = New DConecta

Set ListaConductores = Nothing
Sql = " Select c.*, p.cPersNombre, cAgencia=coalesce(a.cAgeDescripcion,'') " & _
      " from LogVehiculoConductor c " & _
      " Inner Join Persona p on c.cPersCod = p.cPersCod " & _
      " left Join Agencias a on c.cAgeCod = a.cAgeCod "
      
If Co.AbreConexion Then
   Set rs = Co.CargaRecordSet(Sql)
   Co.CierraConexion
End If
Set ListaConductores = rs
Set rs = Nothing
Set Co = Nothing
End Function

'-------------------------------------------------------
'SECCION DE INSERCIONES
'-------------------------------------------------------
Sub InsertRegConductor(ByVal psPersCod As String, ByVal psCargoCod As String, ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal psBrevete As String, _
ByVal pnEstado As Integer, ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String)
Dim Sql1 As String
Dim Sql2 As String
Dim Co As DConecta
Dim sMovNro As String
Set Co = New DConecta

On Error GoTo Sal_InsertCond
sMovNro = GetLogMovNro

Sql1 = "Insert LogVehiculoConductor"
Sql1 = Sql1 & " (cPersCod,cRHCargoCod,cAreaCod, cAgeCod, cBrevete,cMovNro,nEstado) values"
Sql1 = Sql1 & " ('" & psPersCod & "','" & psCargoCod & "','" & psAreaCod & "','" & psAgeCod & "','" & psBrevete & "','" & sMovNro & "'," & pnEstado & ")"

If Co.AbreConexion Then
   Co.BeginTrans
   Co.CargaRecordSet (Sql1)
   Co.CommitTrans
   Co.CierraConexion
Else
   MsgBox "No se puede establecer la conexión..." + Space(10), vbInformation
End If
Exit Sub
Sal_InsertCond:
End Sub

Function GetVehiculoSoat(ByVal pnVehiculoCod As Integer) As String
Dim rs As ADODB.Recordset
Dim Co As DConecta
Dim Sql As String

GetVehiculoSoat = ""
Sql = "Select max(dFechaVig) as MaxFechaVig from LogVehiculoSoat where nVehiculoCod = '" & pnVehiculoCod & "' "
    
Set Co = New DConecta
If Co.AbreConexion Then
   Set rs = Co.CargaRecordSet(Sql)
   If Not rs.EOF Then
      GetVehiculoSoat = IIf(IsNull(rs!MaxFechaVig), "No Indica", rs!MaxFechaVig)
   End If
   Co.CierraConexion
End If
Set rs = Nothing
Set Co = Nothing
End Function

Sub InsertaRegVehiculo(ByVal pscBSCod As String, ByVal pscBSSerie As String, ByVal pnTipo As Integer, ByVal pscModelo As String, _
ByVal pnnAñoFab As Integer, ByVal pnnEstado As Integer, _
ByVal pnnCombustible As Integer, ByVal pscNroMotor As String, _
ByVal pscPlaca As String, ByVal pnnTipoVehiculo As Integer, ByVal pnnMarca As Integer, _
ByVal pnnColor As Integer, ByVal pdFechaIni As Date, pdFechaFin As Date, _
ByVal pdFecSis As Date, ByVal psCodAge As String, ByVal psCodUser As String)

Dim Sql1 As String, Sql2 As String, Sql3 As String
Dim sMovNro As String, nVehiculoCod As Integer, rs As New ADODB.Recordset
Dim Co As New DConecta

sMovNro = GetLogMovNro

Sql1 = " Insert LogVehiculo "
Sql1 = Sql1 & " (cMovNro,cBSCod,cBSSerie,nTipo,cModelo,nAnioFab,nEstado,nCombustible,cNroMotor,cPlaca,nTipoVehiculo,nMarca,nColor) values "
Sql1 = Sql1 & " ('" & sMovNro & "','" & pscBSCod & "','" & pscBSSerie & "'," & pnTipo & ",'" & pscModelo & "'," & pnnAñoFab & ","
Sql1 = Sql1 & pnnEstado & "," & pnnCombustible & ",'" & pscNroMotor & "','" & pscPlaca & "',"
Sql1 = Sql1 & pnnTipoVehiculo & "," & pnnMarca & "," & pnnColor & ")"


If Co.AbreConexion Then
   Co.BeginTrans
   Co.CargaRecordSet (Sql1)
   Co.CommitTrans
   Set rs = Co.CargaRecordSet(" Select nVehiculoCod from LogVehiculo where cBSCod = '" & pscBSCod & "' and cBSSerie = '" & pscBSSerie & "'")
   nVehiculoCod = rs!nVehiculoCod
   Set rs = Nothing
   
   Sql3 = "Insert into LogVehiculoSOAT ( nVehiculoCod, dFechaIni, dFechaVig ) " & _
          " values (" & nVehiculoCod & ",'" & Format(pdFechaIni, "YYYYMMDD") & "','" & Format(pdFechaFin, "YYYYMMDD") & "')"
   
   Co.Ejecutar Sql3
   Co.CierraConexion
End If
End Sub


Function YaEstaRegistrado(vPersCod As String) As Boolean
Dim rs As ADODB.Recordset
Dim Co As DConecta
Dim Sql As String

YaEstaRegistrado = False
Sql = "select * from LogVehiculoConductor where cPersCod = '" & vPersCod & "'"
Set Co = New DConecta
If Co.AbreConexion Then
   Set rs = Co.CargaRecordSet(Sql)
   If Not rs.EOF Then
      YaEstaRegistrado = True
   End If
   Co.CierraConexion
End If
Set rs = Nothing
Set Co = Nothing
End Function

Function GetTipoIncidencia() As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim Co As DConecta
Dim Sql As String

Sql = "select nConsValor AS nIncidencia,cConsDescripcion as cIncidencia from Constante where nConsCod=9025 and nconscod<>nconsvalor"
Set Co = New DConecta
If Co.AbreConexion Then
   Set rs = Co.CargaRecordSet(Sql)
   Set GetTipoIncidencia = rs
   Co.CierraConexion
End If

Set rs = Nothing
Set Co = Nothing
End Function

Function GetVehiculoSeries(ByVal pnAño As Integer) As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim Co As DConecta
Set Co = New DConecta

    Sql = "Select  cBSCod,cSerie, cBSCod+cSerie Codigo,cDescripcion from bsactivofijo"
    Sql = Sql & " where not cBSCod+cSerie in (select  cBSCod+cBSSerie  from LogVehiculo  )"
    Sql = Sql & " and  cbscod like '11203%' and nanio = " & pnAño & " and dBaja is null"
    Sql = Sql & " Order by cBSCod,cSerie"
    Co.AbreConexion
    Set rs = Co.CargaRecordSet(Sql)
    Co.CierraConexion
    Set GetVehiculoSeries = rs
Set rs = Nothing
Set Co = Nothing
End Function

Function GetVehiculoMovDet(ByVal pnAsignacionNro As Integer) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String, oConn  As New DConecta

Sql = "select dFecha=coalesce(m.dFecha,''), m.nTipoReg,m.cValor0,m.cValor1,m.cValor2,t.cTipoReg as cRegistro," & _
          "       m.cDescripcion,cDesc1=coalesce(o.cUbigeoDescripcion,''), cDesc2=coalesce(d.cUbigeoDescripcion,''), m.nMonto " & _
          "  from LogVehiculoAsignacionDet m " & _
          " inner join (select nConsValor AS nTipoReg,cConsDescripcion as cTipoReg from Constante where nConsCod=9024 and nconscod<>nconsvalor) t on m.nTipoReg=t.nTipoReg " & _
          " left outer join UbicacionGeografica o on m.cValor1=o.cUbigeoCod " & _
          " left outer join UbicacionGeografica d on m.cValor2=d.cUbigeoCod " & _
          " where m.nAsignacionNro = '" & pnAsignacionNro & "' and m.nEstado=1 "
          
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(Sql)
   Set GetVehiculoMovDet = rs
   oConn.CierraConexion
End If
End Function

Function GetMovDocVehiculo(vDocTpo As Integer, vDocNro As String) As ADODB.Recordset
Dim Sql As String, oConn As New DConecta

Sql = "select d.nMovNro,m.cMovDesc,c.nMovImporte " & _
"  from MovDoc d inner join Mov m on d.nMovNro=m.nMovNro " & _
"                inner join MovCta c on d.nMovNro = c.nMovNro " & _
" where nDocTpo=" & vDocTpo & " and d.cDocNro = '" & vDocNro & "' and c.nMovImporte>0 "

If oConn.AbreConexion Then
   Set GetMovDocVehiculo = oConn.CargaRecordSet(Sql)
   oConn.CierraConexion
End If
End Function

Function GetTiposVehiculo(pnTipo As Integer) As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim Co As DConecta
Dim Sql As String
Dim nMax As Integer, nMin As Integer

nMin = pnTipo * 10
nMax = (pnTipo + 1) * 10

    Sql = "select nConsValor AS nTipoVehiculo,cConsDescripcion as cDescripcion from Constante where nConsCod=9026 and nconscod<>nconsvalor and nConsValor>=" & nMin & " and nConsValor<" & nMax & " "
    Set Co = New DConecta
    Co.AbreConexion
    Set rs = Co.CargaRecordSet(Sql)
    Co.CierraConexion
Set GetTiposVehiculo = rs
Set rs = Nothing
Set Co = Nothing
End Function

Function GetMarcasVehiculo(pnTipo As Integer) As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim Co As DConecta
Dim Sql As String
Dim nMax As Integer, nMin As Integer

nMin = pnTipo * 10
nMax = (pnTipo + 1) * 10

Sql = "select nConsValor AS nMarca,cConsDescripcion as cMarca from Constante where nConsCod=9022 and nconscod<>nconsvalor and nConsValor>=" & nMin & " and nConsValor<" & nMax & ""
Set Co = New DConecta
    Co.AbreConexion
    Set rs = Co.CargaRecordSet(Sql)
    Co.CierraConexion
Set GetMarcasVehiculo = rs
Set rs = Nothing
Set Co = Nothing
End Function

Function GetColorV() As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim Co As DConecta
Dim Sql As String
    'SQL = "select cColor, nColor from Logisticacolor"
    Sql = "select nConsValor AS nColor,cConsDescripcion as cColor from Constante where nConsCod=9023 and nconscod<>nconsvalor"
    Set Co = New DConecta
    Co.AbreConexion
    Set rs = Co.CargaRecordSet(Sql)
    Co.CierraConexion
Set GetColorV = rs
Set rs = Nothing
Set Co = Nothing
End Function

Function GetCombustible() As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim Co As DConecta
Dim Sql As String
    Sql = "select nConsValor AS nCombustible,cConsDescripcion as cCombustible from Constante where nConsCod=9027 and nconscod<>nconsvalor"
    Set Co = New DConecta
    Co.AbreConexion
    Set rs = Co.CargaRecordSet(Sql)
    Co.CierraConexion
Set GetCombustible = rs
Set rs = Nothing
Set Co = Nothing
End Function

Function GetBrevete(psPersCod As String) As String
Dim rs As ADODB.Recordset
Dim Co As DConecta
Dim Sql As String

GetBrevete = ""
    Sql = "select cPersIDnro from persID where cPersIDTpo = '7' and cPersCod='" & psPersCod & "'"
    Set Co = New DConecta
    Co.AbreConexion
    Set rs = Co.CargaRecordSet(Sql)
    Co.CierraConexion
If rs.State <> 0 Then
   If Not rs.EOF Then
      GetBrevete = rs!cPersIDnro
   End If
End If
Set rs = Nothing
Set Co = Nothing
End Function


