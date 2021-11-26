VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCaptaTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oConect As COMConecta.DCOMConecta
Dim oError As New COMConecta.COMErrorHandling
Dim oVarPublicas As New COMFunciones.FCOMVarPublicas

Dim sSql As String
Dim Rs As ADODB.Recordset
Private Sub Class_Initialize()
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:Initialize Method. Error en Conexion a Base de datos")
End If

End Sub
Private Sub Class_Terminate()
oConect.CierraConexion
Set oConect = Nothing

End Sub
Public Sub RegistraTarjetaPersona(ByVal sTarjeta As String, ByVal sMovNro As String, _
ByVal sPersona As String)
Dim sSql As String

sSql = "Update Tarjeta set nEstado = " & gCapTarjEstActiva & ", cMovNro ='" & sMovNro & "', "
sSql = sSql & " cPersCod = '" & sPersona & "' where cTarjCod='" & sTarjeta & "'"
oConect.Ejecutar (sSql)

sSql = "INSERT TarjetaEstado (cTarjCod,cMovNro,nEstado,cDescripcion) VALUES " _
    & "('" & sTarjeta & "','" & sMovNro & "'," & gCapTarjEstActiva & ",'AFILIACION TARJETA')"
oConect.Ejecutar (sSql)

End Sub
Public Sub BloqueoTarjeta(ByVal sTarjeta As String, ByVal sMovNro As String)
Dim sSql As String

sSql = "Update Tarjeta set nEstado = " & gCapTarjEstBloqueada & ", cMovNro ='" & sMovNro & "', "
sSql = sSql & "  where cTarjCod='" & sTarjeta & "'"
oConect.Ejecutar (sSql)

sSql = "INSERT TarjetaEstado (cTarjCod,cMovNro,nEstado,cDescripcion) VALUES " _
    & "('" & sTarjeta & "','" & sMovNro & "'," & gCapTarjEstBloqueada & ",'BLOQUEO TARJETA')"
oConect.Ejecutar (sSql)

End Sub
Public Function VerificaTarjetaActivaReg(ByVal psTarjeta As String) As Boolean
Dim sSql As String
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
sSql = "Select count(*) Nro from Tarjeta where cTarjcod = '" & psTarjeta & "' and nEstado =" & CaptacTarjetaEstado.gCapTarjEstRegistrada
 Set Rs = oConect.CargaRecordSet(sSql)
If Rs!Nro = 0 Then
    VerificaTarjetaActivaReg = False
Else
    VerificaTarjetaActivaReg = True
End If

Set Rs = Nothing
End Function

Public Function VerificaTarjetaActiva(ByVal psTarjeta As String) As Boolean
Dim sSql As String
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
sSql = "Select count(*) Nro from Tarjeta where cTarjcod = '" & psTarjeta & "' and nEstado =" & gCapTarjEstActiva
 Set Rs = oConect.CargaRecordSet(sSql)
If Rs!Nro = 0 Then
    VerificaTarjetaActiva = False
Else
    VerificaTarjetaActiva = True
End If

Set Rs = Nothing
End Function
Public Function VerificaPersonaTarjetaActiva(ByVal psTarjeta As String) As Boolean
Dim sSql As String
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
sSql = "Select count(*) Nro from Tarjeta where cPersCod = '" & psTarjeta & "' and nEstado =" & gCapTarjEstActiva
Set Rs = oConect.CargaRecordSet(sSql)
 
If Rs!Nro = 0 Then
    VerificaPersonaTarjetaActiva = False
Else
    VerificaPersonaTarjetaActiva = True
End If

Set Rs = Nothing
End Function

Public Function Get_Tarj_PVV(ByVal psTarjeta As String) As String
Dim sSql As String
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
sSql = "Select cPVV  from Tarjeta where cTarjcod = '" & psTarjeta & "'"
Set Rs = oConect.CargaRecordSet(sSql)
Get_Tarj_PVV = Rs!cPVV
Set Rs = Nothing
End Function
Public Function Get_Tarj_HistorialEst(ByVal psTarjeta As String) As ADODB.Recordset
Dim sSql As String
sSql = "SELECT convert(datetime,substring(TE.cMovNro,1,8),103) Fecha , cConsDescripcion"
sSql = sSql & " From Tarjeta T"
sSql = sSql & " Inner Join TarjetaEstado TE on TE.cTarjcod = T.cTarjCod and T.cTarjCod='" & psTarjeta & "'"
sSql = sSql & " Inner Join Constante C on C.nConsValor = TE.nEstado and nConsCod = 9065"
sSql = sSql & " Order by nConsValor"
Set Get_Tarj_HistorialEst = oConect.CargaRecordSet(sSql)
End Function
Public Function Get_Datos_Tarj(ByVal psTarjeta As String) As ADODB.Recordset
Dim sSql As String
sSql = " select cTarjCod,cPersCod,dCaduca,nEstado,cPVV,cMovNro,Lote,Inicio,Final,agencia"
sSql = sSql & " from Tarjeta where cTarjCod = '" & psTarjeta & "'"
Set Get_Datos_Tarj = oConect.CargaRecordSet(sSql)
End Function
Public Function Get_Tarj_Cod_Titular(ByVal psTarjeta As String) As String
Dim sSql As String
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
sSql = " Select P.cPersCod "
sSql = sSql & " from Tarjeta T"
sSql = sSql & " Inner Join Persona P on T.cPersCod = P.cPersCod"
sSql = sSql & " where T.cTarjCod = '" & psTarjeta & "'"
Set Rs = oConect.CargaRecordSet(sSql)
If Not (Rs.EOF And Rs.BOF) Then
    Get_Tarj_Cod_Titular = Rs!cPersCod
Else
    Get_Tarj_Cod_Titular = ""
End If
Set Rs = Nothing
End Function
Public Sub Put_Tarj_PVV(ByVal psTarjeta As String, ByVal psPVV As String)
Dim sSql As String
sSql = "Update Tarjeta set  cPVV ='" & psPVV & "' where cTarjcod = '" & psTarjeta & "'"
oConect.Ejecutar (sSql)
End Sub

Public Sub RegistraCambioPind(ByVal psTarjeta As String, pdFecha As Date, ByVal psMovNro As String)
Dim sSql As String
sSql = "Insert tarjetaCambioPin"
sSql = sSql & " (cCodTarj,dFecha,cMovNro) values"
sSql = sSql & " ('" & psTarjeta & "','" & Format(pdFecha, "MM/DD/YYYY HH:MM:SS") & "','" & psMovNro & "')"
oConect.Ejecutar (sSql)
End Sub

