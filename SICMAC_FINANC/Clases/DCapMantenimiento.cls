VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCapMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public dbCmact As DConecta
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String

Public Function GetDatosCuentaAho(ByVal sCuenta As String) As Recordset
Dim sSql As String
Dim rsCta As Recordset

sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, " _
    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.bOrdPag, A.dUltContacto, T.cConsDescripcion cEstado, " _
    & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.bInactiva, A.bInmovilizada FROM Producto P " _
    & "INNER JOIN Captaciones C INNER JOIN CaptacAhorros A ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN " & sDBComunes & "Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN " & sDBComunes & "Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor WHERE P.cCtaCod = '" & sCuenta & "' AND " _
    & "T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    & "AND T2.nConsCod = " & gCaptacTipoTasa

Set rsCta = dbCmact.CargaRecordSet(sSql)
Set GetDatosCuentaAho = rsCta
Set rsCta = Nothing
End Function

Public Function GetDatosCuentaPF(ByVal sCuenta As String) As Recordset
Dim sSql As String
Dim rsCta As Recordset

sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, " _
    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.nPlazo, A.nIntPag, A.dRenovacion, A.nApertura, " _
    & "A.nFormaRetiro, A.dAuxiliar, A.nDuplicado, T.cConsDescripcion cEstado, " _
    & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, T3.cConsDescripcion cRetiro, " _
    & "A.dUltCierreAnt, A.dAuxiliarAnt, ISNULL(CI.cCtaCodAbono,'') cCtaCodAbono FROM Producto P " _
    & "INNER JOIN Captaciones C INNER JOIN CaptacPlazoFijo A LEFT JOIN CaptacCtaAboIntPF CI " _
    & "ON A.cCtaCod = CI.cCtaCod ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN " & sDBComunes & "Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN " & sDBComunes & "Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN " & sDBComunes & "Constante T3 ON A.nFormaRetiro = T3.nConsValor WHERE P.cCtaCod = '" & sCuenta & "' AND " _
    & "T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    & "AND T2.nConsCod = " & gCaptacTipoTasa & " AND T3.nConsCod = " & gCaptacPFFormaRetiro

Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaPF = rsCta
Set rsCta = Nothing
End Function

Public Function GetDatosCuentaCTS(ByVal sCuenta As String) As Recordset
Dim sSql As String
Dim rsCta As Recordset

sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, " _
    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.nSaldRetiro, A.nIntSaldo, PE.cPersNombre cInstitucion, " _
    & "T.cConsDescripcion cEstado, T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa " _
    & "FROM Producto P INNER JOIN Captaciones C INNER JOIN CaptacCTS A INNER JOIN " & sDBPersona & "" _
    & "Persona PE ON A.cCodInst = PE.cPersCod ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN " & sDBComunes & "Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN " & sDBComunes & "Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor WHERE P.cCtaCod = '" & sCuenta & "' AND " _
    & "T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    & "AND T2.nConsCod = " & gCaptacTipoTasa

Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaCTS = rsCta
Set rsCta = Nothing
End Function

Public Function GetProductoPersona(ByVal sCuenta As String) As Recordset
Dim sSql As String
Dim rsCta As Recordset

sSql = "Select PP.cPersCod, P.cPersNombre Nombre, (UPPER(RTRIM(T.cConsDescripcion)) " _
    & "+ SPACE(50) + CONVERT(Varchar(2),PP.nPrdPersRelac)) cRelacion " _
    & "FROM " & sDBPersona & "Persona P INNER JOIN ProductoPersona PP ON " _
    & "P.cPersCod = PP.cPersCod INNER JOIN " & sDBComunes & "Constante T " _
    & "ON PP.nPrdPersRelac = T.nConsValor WHERE PP.cCtaCod = '" & sCuenta & "' " _
    & "AND T.nConsCod = " & gCaptacRelacPersona
    
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetProductoPersona = rsCta
Set rsCta = Nothing
End Function

Public Function GetCuentasPersona(ByVal sPers As String, Optional nProd As Producto, _
        Optional bActivas As Boolean = False, Optional bBloqueadas As Boolean = False) As Recordset
Dim sSql As String
Dim rsCta As Recordset

sSql = "Select PP.cCtaCod, (UPPER(T3.cConsDescripcion) + ' - ' + T1.cConsDescripcion) cDescripcion, nNivel = 1, " _
    & "T1.cConsDescripcion cRelacion, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, " _
    & "UPPER(T2.cConsDescripcion) cProducto, UPPER(T3.cConsDescripcion) cMoneda FROM ProductoPersona PP INNER JOIN Producto P INNER " _
    & "JOIN Captaciones C ON P.cCtaCod = C.cCtaCod ON PP.cCtaCod = P.cCtaCod INNER JOIN " _
    & sDBComunes & "Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN " & sDBComunes & "" _
    & "Constante T1 ON PP.nPrdPersRelac = T1.nConsValor INNER JOIN " & sDBComunes & "Constante T2 " _
    & "ON SUBSTRING(PP.cCtaCod,6,3) = CONVERT(Varchar(3),T2.nConsValor) INNER JOIN " & sDBComunes & "" _
    & "Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor) WHERE PP.cPersCod = '" & sPers & "' " _
    & "AND T1.nConsCod = " & gCaptacRelacPersona & " AND T.nConsCod = " & gCaptacEstado & " AND " _
    & "T2.nConsCod = " & gProducto & " AND T3.nConsCod = " & gMoneda

If nProd <> 0 Then
    sSql = sSql & " AND SUBSTRING(PP.cCtaCod,6,3) = '" & nProd & "'"
End If
If bActivas Then
    sSql = sSql & " AND P.nPrdEstado NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ")"
End If
If bBloqueadas Then
    sSql = sSql & " AND P.nPrdEstado NOT IN (" & gCapEstBloqRetiro & "," & gCapEstBloqTotal & ")"
End If
sSql = sSql & " ORDER BY PP.cCtaCod"
Set rsCta = dbCmact.CargaRecordSet(sSql)
Set GetCuentasPersona = rsCta
Set rsCta = Nothing
End Function

Public Sub EliminaCuentaTarjPersona(ByVal sPersona As String, ByVal sTarjeta As String)
Dim sSql As String
sSql = "Delete CuentaTarj Where cPersCod = '" & sPersona & "' AND cTarjCod = '" & sTarjeta & "'"
dbCmact.Ejecutar sSql
End Sub

Public Sub AgregaCuentaTarjPersona(ByVal sCuenta As String, ByVal sPersona As String, ByVal sTarjeta As String)
Dim sSql As String

sSql = "INSERT CuentaTarj (cCtaCod,cPersCod,cTarjCod) " _
    & "VALUES ('" & sCuenta & "','" & sPersona & "','" & sTarjeta & "')"

dbCmact.Ejecutar sSql
End Sub


Public Sub EliminaProductoPersona(ByVal sCuenta As String)
Dim sSql As String
sSql = "Delete ProductoPersona Where cCtaCod = '" & sCuenta & "'"
dbCmact.Ejecutar sSql
End Sub

Public Sub AgregaProductoPersona(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona)
Dim sSql As String

sSql = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
    & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ")"

dbCmact.Ejecutar sSql
End Sub


Public Function GetCapBloqueos(ByVal sCuenta As String, ByVal nTipoBloqueo As CaptacTipoBloqueo, ByVal nConstante As ConstanteCabecera) As Recordset
Dim sSql As String
Dim rsRel As Recordset

Set rsRel = New Recordset
rsRel.CursorLocation = adUseClient

sSql = "Select nValor = CASE WHEN B.cComentario IS NULL THEN 0 WHEN B.cComentario IS NOT NULL THEN 1 END, " _
    & "C.cConsDescripcion, ISNULL(B.cComentario,'') cComentario, ISNULL(B.dFecha,'') dFecha, B.Usu, C.nConsValor FROM " _
    & "(Select nConsValor, cConsDescripcion FROM Constante WHERE nConsCod = " & nConstante & ") C LEFT JOIN " _
    & "(Select nBlqMotivo, cComentario, CONVERT(VARCHAR(10),CONVERT(Datetime,LEFT(cMovNro,8),101),105) dFecha, RIGHT(cMovNro,4) Usu FROM CaptacBloqueos WHERE " _
    & "cCtaCod = '" & sCuenta & "' And cMovNroDbl IS NULL AND nBlqTpo = " & nTipoBloqueo & " ) B ON " _
    & "C.nConsValor = B.nBlqMotivo Order by C.nConsValor"

rsRel.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsRel.ActiveConnection = Nothing
Set GetCapBloqueos = rsRel
End Function


Public Function NuevoBloqueoRetiro(ByVal sCuenta As String, ByVal nMotivo As CaptacMotBloqueoRet, _
        cComentario As String, ByVal cMovNro As String)
Dim sSql As String
sSql = "INSERT CaptacBloqueos (cCtaCod,nBlqTpo,nBlqMotivo,cComentario,cMovNro,cMovNroDbl) " _
    & "VALUES ('" & sCuenta & "'," & gCapTpoBlqRetiro & "," & nMotivo & ",'" & cComentario & "','" & cMovNro & "',NULL)"
dbCmact.Ejecutar sSql
End Function

Public Function NuevoBloqueoTotal(ByVal sCuenta As String, ByVal nMotivo As CaptacMotBloqueoTot, _
        cComentario As String, ByVal cMovNro As String)
Dim sSql As String
sSql = "INSERT CaptacBloqueos (cCtaCod,nBlqTpo,nBlqMotivo,cComentario,cMovNro,cMovNroDbl) " _
    & "VALUES ('" & sCuenta & "'," & gCapTpoBlqTotal & "," & nMotivo & ",'" & cComentario & "','" & cMovNro & "',NULL)"
dbCmact.Ejecutar sSql
End Function

Public Sub ActualizaEstadoCuenta(ByVal sCuenta As String, nEstado As CaptacEstado)
Dim sSql As String
sSql = "UPDATE Producto Set nPrdEstado = " & nEstado & " WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Ejecutar sSql
End Sub

Public Sub ActualizaBloqueoRet(ByVal sCuenta As String, ByVal sComentario As String, _
    ByVal sMovNroDbl As String, ByVal nMotivo As CaptacMotBloqueoRet)
Dim sSql As String
sSql = "UPDATE CaptacBloqueos Set cComentario = '" & sComentario & "', " _
    & "cMovNroDbl = '" & sMovNroDbl & "' WHERE cCtaCod = '" & sCuenta & "' " _
    & "And nBlqTpo = " & gCapTpoBlqRetiro & " AND nBlqMotivo = " & nMotivo
dbCmact.Ejecutar sSql
End Sub

Public Sub ActualizaBloqueoTot(ByVal sCuenta As String, ByVal sComentario As String, _
    ByVal sMovNroDbl As String, ByVal nMotivo As CaptacMotBloqueoTot)

Dim sSql As String
sSql = "UPDATE CaptacBloqueos Set cComentario = '" & sComentario & "', " _
    & "cMovNroDbl = '" & sMovNroDbl & "' WHERE cCtaCod = '" & sCuenta & "' " _
    & "And nBlqTpo = " & gCapTpoBlqTotal & " AND nBlqMotivo = " & nMotivo
dbCmact.Ejecutar sSql
End Sub

Public Function ExisteBloqueoTot(ByVal sCuenta As String, ByVal nMotivo As CaptacMotBloqueoTot) As Boolean
Dim sSql As String
Dim bExiste As Boolean
Dim rsBloq As Recordset
bExiste = False
sSql = "Select cCtaCod FROM CaptacBloqueos WHERE cCtaCod = '" & sCuenta & "' " _
    & "And nBlqTpo = " & gCapTpoBlqTotal & " AND nBlqMotivo = " & nMotivo & " AND " _
    & "cMovNroDbl IS NULL"
Set rsBloq = New Recordset
rsBloq.CursorLocation = adUseClient
rsBloq.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsBloq.ActiveConnection = Nothing
If Not (rsBloq.EOF And rsBloq.BOF) Then
    bExiste = True
End If
Set rsBloq = Nothing
ExisteBloqueoTot = bExiste
End Function

Public Function ExisteBloqueoRet(ByVal sCuenta As String, ByVal nMotivo As CaptacMotBloqueoRet) As Boolean
Dim sSql As String
Dim bExiste As Boolean
Dim rsBloq As Recordset
bExiste = False
sSql = "Select cCtaCod FROM CaptacBloqueos WHERE cCtaCod = '" & sCuenta & "' " _
    & "And nBlqTpo = " & gCapTpoBlqRetiro & " AND nBlqMotivo = " & nMotivo & " AND " _
    & "cMovNroDbl IS NULL"
Set rsBloq = New Recordset
rsBloq.CursorLocation = adUseClient
rsBloq.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsBloq.ActiveConnection = Nothing
If Not (rsBloq.EOF And rsBloq.BOF) Then
    bExiste = True
End If
Set rsBloq = Nothing
ExisteBloqueoRet = bExiste
End Function

Public Sub ActualizaComentBlqRet(ByVal sCuenta As String, ByVal sComentario As String, _
    ByVal nMotivo As CaptacMotBloqueoRet)
Dim sSql As String

sSql = "UPDATE CaptacBloqueos Set cComentario = '" & sComentario & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "' AND nBlqTpo = " & gCapTpoBlqRetiro & " " _
    & "AND nBlqMotivo = " & nMotivo
dbCmact.Ejecutar sSql
End Sub

Public Sub ActualizaComentBlqTot(ByVal sCuenta As String, ByVal sComentario As String, _
    ByVal nMotivo As CaptacMotBloqueoTot)
Dim sSql As String

sSql = "UPDATE CaptacBloqueos Set cComentario = '" & sComentario & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "' AND nBlqTpo = " & gCapTpoBlqTotal & " " _
    & "AND nBlqMotivo = " & nMotivo
dbCmact.Ejecutar sSql
End Sub

Public Function GetCapTasaInteres(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
    ByVal nTipoTasa As CaptacTipoTasa, ByVal nValor As Double) As Double

Dim rsTasa As Recordset
Dim sSql As String

sSql = "SELECT nTasaValor FROM CaptacTasas " _
    & "WHERE cTasaProd = " & nProducto & " AND cTasaMon = " & nMoneda & " AND " _
    & "cTasaTpo = " & nTipoTasa & " And " & nValor & " BETWEEN nValorIni AND nValorFin"

Set rsTasa = New Recordset
rsTasa.CursorLocation = adUseClient
rsTasa.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTasa.ActiveConnection = Nothing
If Not (rsTasa.EOF Or rsTasa.BOF) Then
    GetCapTasaInteres = rsTasa("nTasaValor")
Else
    GetCapTasaInteres = 0
End If
rsTasa.Close
Set rsTasa = Nothing
End Function

Public Function GetPersonaTarj(ByVal sPersona As String) As Recordset
Dim sSql As String
Dim rsTarj As Recordset
sSql = "SELECT CT.cCtaCod, CT.cTarjCod FROM Producto P INNER JOIN ProductoPersona PP INNER JOIN " _
    & "CuentaTarj CT INNER JOIN Tarjeta T ON CT.cTarjCod = T.cTarjCod ON PP.cCtaCod = CT.cCtaCod " _
    & "AND PP.cPersCod = CT.cPersCod ON P.cCtaCod = PP.cCtaCod WHERE CT.cPersCod = '" & sPersona & "' " _
    & "AND T.cEstado = " & gCapTarjEstActiva & " AND P.nPrdEstado NOT IN (" & gCapEstAnulada & "," _
    & gCapEstAnulada & "," & gCapEstCancelada & "," & gCapEstCancelada & ")"
Set rsTarj = New Recordset
rsTarj.CursorLocation = adUseClient
rsTarj.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTarj.ActiveConnection = Nothing
Set GetPersonaTarj = rsTarj
Set rsTarj = Nothing
End Function

Public Function GetTarjetaCuentas(ByVal sTarjeta As String) As Recordset
Dim sSql As String
Dim rsTarj As Recordset

sSql = "SELECT CT.cCtaCod, UPPER(K.cConsDescripcion) Producto, UPPER(K1.cConsDescripcion) Moneda, CT.cPersCod, UPPER(K2.cConsDescripcion) Relacion, " _
    & "T.cEstado FROM Producto P INNER JOIN ProductoPersona PP INNER JOIN CuentaTarj CT INNER JOIN Tarjeta T ON CT.cTarjCod = T.cTarjCod ON PP.cCtaCod = " _
    & "CT.cCtaCod AND PP.cPersCod = CT.cPersCod ON P.cCtaCod = PP.cCtaCod INNER JOIN " & sDBComunes & "Constante K ON SUBSTRING(P.cCtaCod,6,3) " _
    & "= CONVERT(Varchar(3),K.nConsValor) INNER JOIN " & sDBComunes & "Constante K1 ON SUBSTRING(P.cCtaCod,9,1) = CONVERT(Varchar(1),K1.nConsValor) " _
    & "INNER JOIN " & sDBComunes & "Constante K2 ON PP.nPrdPersRelac = K2.nConsValor WHERE " _
    & "T.cEstado NOT IN (" & gCapTarjEstCancelada & ") AND P.nPrdEstado NOT IN (" & gCapEstAnulada & "," & gCapEstAnulada & "," _
    & gCapEstCancelada & "," & gCapEstCancelada & ") AND T.cTarjCod = '" & sTarjeta & "' And K.nConsCod = " & gProducto & " " _
    & "And K1.nConsCod = " & gMoneda & " AND K2.nConsCod = " & gCaptacRelacPersona

Set rsTarj = New Recordset
rsTarj.CursorLocation = adUseClient
rsTarj.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTarj.ActiveConnection = Nothing
Set GetTarjetaCuentas = rsTarj
Set rsTarj = Nothing
End Function


Public Sub AgregaTarjeta(ByVal sTarjeta As String, ByVal sClave As String, ByVal dCaduca As Date, ByVal sMovNro As String)
Dim sSql As String
sSql = "INSERT Tarjeta (cTarjCod,cClave,dCaduca,cEstado) VALUES " _
    & "('" & sTarjeta & "','" & sClave & "','" & Format$(dCaduca, gsFormatoFecha) & "'," & gCapTarjEstActiva & ")"
dbCmact.Ejecutar sSql
sSql = "INSERT TarjetaEstado (cTarjCod,cMovNro,cEstado,cDescripcion) VALUES " _
    & "('" & sTarjeta & "','" & sMovNro & "'," & gCapTarjEstActiva & ",'NUEVA TARJETA')"
dbCmact.Ejecutar sSql
End Sub

Public Sub AgregaCuentaTarj(ByVal sTarjeta As String, ByVal sCuenta As String, ByVal sPersona As String)
Dim sSql As String
sSql = "INSERT CuentaTarj (cCtaCod,cPersCod,cTarjCod) VALUES " _
    & "('" & sCuenta & "','" & sPersona & "','" & sTarjeta & "')"
dbCmact.Ejecutar sSql
End Sub

Public Function ExisteTarjeta(ByVal sTarjeta As String) As Boolean
Dim sSql As String
Dim rsTarj As Recordset
rsTarj.CursorLocation = adUseClient
sSql = "SELECT cTarjCod FROM Tarjeta WHERE cTarjCod = '" & sTarjeta & "'"
Set rsTarj = dbCmact.CargaRecordSet(sSql)
If rsTarj.EOF And rsTarj.BOF Then
    ExisteTarjeta = False
Else
    ExisteTarjeta = True
End If
rsTarj.Close
Set rsTarj = Nothing
End Function

Public Function ExisteProdPersonaTarj(ByVal sPersona As String, ByVal sCuenta As String) As Boolean
Dim sSql As String
Dim rsTarj As Recordset
Set rsTarj = New Recordset
rsTarj.CursorLocation = adUseClient
sSql = "SELECT cTarjCod FROM CuentaTarj WHERE cPersCod = '" & sPersona & "' AND cCtaCod = '" & sCuenta & "'"
rsTarj.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTarj.ActiveConnection = Nothing
If rsTarj.EOF And rsTarj.BOF Then
    ExisteProdPersonaTarj = False
Else
    ExisteProdPersonaTarj = True
End If
rsTarj.Close
Set rsTarj = Nothing
End Function

Public Function ExisteCuentaTarj(ByVal sTarjeta As String, ByVal sPersona As String, ByVal sCuenta As String) As Boolean
Dim sSql As String
Dim rsTarj As Recordset
rsTarj.CursorLocation = adUseClient
sSql = "SELECT cTarjCod FROM CuentaTarj WHERE cTarjCod = '" & sTarjeta & "' AND " _
    & "cPersCod = '" & sPersona & "' AND cCtaCod = '" & sCuenta & "'"
rsTarj.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTarj = dbCmact.CargaRecordSet(sSql)
If rsTarj.EOF And rsTarj.BOF Then
    ExisteCuentaTarj = False
Else
    ExisteCuentaTarj = True
End If
rsTarj.Close
Set rsTarj = Nothing
End Function

Public Function GetBeneficiarios(ByVal sPersona As String) As Recordset
Dim rsPers As Recordset
Dim sSql As String

sSql = "SELECT PR.cPersRelacPersCod, P.cPersNombre, T.cConsDescripcion cParentesco, " _
    & "DATEDIFF(yy,dPersNacCreac,Getdate()) nEdad, nPersRelacBenefPorc, PR.cPersRelac " _
    & "FROM " & sDBPersona & "Persona P INNER JOIN PersRelaciones PR INNER JOIN Constante T " _
    & "ON PR.cPersRelac = T.nConsValor ON P.cPersCod = PR.cPersRelacPersCod WHERE " _
    & "bPersRelacBenef = 1 AND PR.cPersCod = '" & sPersona & "' And T.nConsCod = " & gPersRelacion
    
Set rsPers = dbCmact.CargaRecordSet(sSql)
Set GetBeneficiarios = rsPers
Set rsPers = Nothing
End Function

Public Function GetDatosPersona(ByVal sPersona As String) As Recordset
Dim sSql As String
Dim rsCta As Recordset
sSql = "Select P.cPersNombre Nombre, P.cPersDireccDomicilio Direccion, U.cUbiGeoDescripcion Zona, " _
    & "P.cPersTelefono Fono, T.cConsDescripcion ID, I.cPersIDnro [ID N°], P.cPersCod, P.nPersPersoneria, " _
    & "I.cPersIDtpo FROM " & sDBPersona & "PersId I RIGHT JOIN " & sDBPersona & "Persona P LEFT JOIN " _
    & sDBComunes & "UbicacionGeografica U ON P.cPersDireccUbiGeo = U.cUbiGeoCod ON I.cPersCod = P.cPersCod " _
    & "INNER JOIN " & sDBComunes & "Constante T ON I.cPersIDTpo = T.nConsValor WHERE P.cPersCod = '" & sPersona & "' " _
    & "AND T.nConsCod = " & gPersIdTipo

Set rsCta = dbCmact.CargaRecordSet(sSql)
Set GetDatosPersona = rsCta
Set rsCta = Nothing
End Function

Public Function GetDatosOrdenPago(ByVal sCuenta As String, ByVal nNumOP As Long) As Recordset
Dim sSql As String
Dim rsCta As Recordset
sSql = "Select D.cNroDoc, K.cConsDescripcion cDescripcion,  D.nMonto, CONVERT(Varchar(10),CONVERT(Datetime,LEFT(D.cMovNro,8)),103) Fecha, " _
    & "RIGHT(D.cMovNro,4) Usuario, D.cEstado FROM DocRecOP OP INNER JOIN DocRecOPEst D INNER JOIN " & sDBComunes & "Constante K ON D.cEstado = K.nConsValor " _
    & "ON OP.nTpoDoc = D.nTpoDoc AND OP.cNroDoc = D.cNroDoc AND OP.cCtaCod = D.cCtaCod WHERE D.cCtaCod = '" & sCuenta & "' AND D.cNroDoc = '" & nNumOP & "' And K.nConsCod = " & gCaptacOrdPagoEstado & " ORDER BY D.cMovNro DESC"
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosOrdenPago = rsCta
Set rsCta = Nothing
End Function

Public Function GetPersonaCuenta(ByVal sCuenta As String) As Recordset
Dim sSql As String
Dim rsCta As Recordset
sSql = "Select P.cPersNombre Nombre, T1.cConsDescripcion Relacion, P.cPersDireccDomicilio Direccion, ISNULL(U.cUbiGeoDescripcion,'') Zona, " _
    & "P.cPersTelefono Fono, T.cConsDescripcion ID, I.cPersIDnro [ID N°], P.cPersCod, P.nPersPersoneria, " _
    & "I.cPersIDtpo, PP.nPrdPersRelac, PP.cCtaCod FROM " & sDBPersona & "PersId I INNER JOIN (" & sDBPersona & "Persona P LEFT JOIN " _
    & sDBComunes & "UbicacionGeografica U ON P.cPersDireccUbiGeo = U.cUbiGeoCod) INNER JOIN ProductoPersona PP " _
    & "ON P.cPersCod = PP.cPersCod ON I.cPersCod = P.cPersCod INNER JOIN " & sDBComunes & "Constante T ON I.cPersIDTpo = T.nConsValor " _
    & "INNER JOIN " & sDBComunes & "Constante T1 ON PP.nPrdPersRelac = T1.nConsValor WHERE PP.cCtaCod = '" & sCuenta & "' " _
    & "AND T.nConsCod = " & gPersIdTipo & " AND T1.nConsCod = " & gCaptacRelacPersona & " ORDER BY Nombre, I.cPersIDtpo"

Set rsCta = dbCmact.CargaRecordSet(sSql)
Set GetPersonaCuenta = rsCta
Set rsCta = Nothing
End Function


Public Sub EliminaBeneficiarios(ByVal sPersona As String)
Dim sSql As String
sSql = "DELETE PersRelaciones WHERE cPersCod = '" & sPersona & "' And bPersRelacBenef = 1"
dbCmact.Ejecutar sSql
End Sub

Public Sub AgregaBeneficiario(ByVal sPersona As String, ByVal sPersRelac As String, _
    ByVal nRelacion As PersRelacion, ByVal nPorcentaje As Double, ByVal sMovNro As String)

Dim sSql As String
sSql = "INSERT PersRelaciones (cPersCod,cPersRelacPersCod,cPersRelac,bPersRelacBenef,nPersRelacBenefPorc,bPersRelacAMP, cMovNro) " _
    & "VALUES ('" & sPersona & "','" & sPersRelac & "'," & nRelacion & ",1," & nPorcentaje & ",NULL,'" & sMovNro & "')"
dbCmact.Ejecutar sSql
End Sub

Public Sub ActualizaDatosCuenta(ByVal sCuenta As String, ByVal bOrdPag As Boolean, _
        ByVal nFirmas As Integer, nTipoCuenta As ProductoCuentaTipo, ByVal sInstitucion As String)

Dim sSql As String
Dim nProd As Producto

nProd = CLng(Mid(sCuenta, 6, 3))
If nProd = gCapAhorros Then
    sSql = "Update CaptacAhorros Set bOrdPag = " & IIf(bOrdPag, 1, 0) & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
    dbCmact.Ejecutar sSql
ElseIf nProd = gCapCTS Then
    sSql = "Update CaptacCTS Set cCodInst = '" & sInstitucion & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
    dbCmact.Ejecutar sSql
End If
sSql = "Update Captaciones Set nFirmas = " & nFirmas & ", " _
    & "nPrdCtaTpo = " & nTipoCuenta & " WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Ejecutar sSql
End Sub

Public Function GetOrdenPagoEmitidas(ByVal sCuenta As String) As Recordset
Dim sSql As String
Dim rsCta As Recordset
rsCta.CursorLocation = adUseClient
sSql = "SELECT nInicio, nFin, Convert(Varchar(10),Convert(Datetime,LEFT(cMovNro,8)),103) Fecha, " _
    & "RIGHT(cMovNro,4) Usuario FROM MovDocEmitidoRango WHERE cCtaCod = '" & sCuenta & "' " _
    & "ORDER BY Fecha DESC"
Set rsCta = dbCmact.CargaRecordSet(sSql)
Set GetOrdenPagoEmitidas = rsCta
Set rsCta = Nothing
End Function

Public Function ExisteOrdenPagoEmitidas(ByVal sCuenta As String, ByVal nInicio As Long, _
        ByVal nFin As Long) As String
Dim sSql As String, sRango As String
Dim rsCta As Recordset
rsCta.CursorLocation = adUseClient
sSql = "Select Rango = Convert(Varchar(10),nInicio) + ' - ' + Convert(Varchar(10),nFin) " _
    & "FROM MovDocEmitidoRango Where (" & nInicio & " Between nInicio And nFin Or " & nFin & " " _
    & "BETWEEN nInicio And nFin) And cCtaCod = '" & sCuenta & "'"
Set rsCta = dbCmact.CargaRecordSet(sSql)
If rsCta.EOF And rsCta.BOF Then
    ExisteOrdenPagoEmitidas = ""
Else
    sRango = ""
    Do While Not rsCta.EOF
        sRango = sRango & rsCta("Rango") & Chr$(13)
        rsCta.MoveNext
    Loop
    ExisteOrdenPagoEmitidas = sRango
End If
rsCta.Close
Set rsCta = Nothing
End Function

Public Function GetEstadoOrdenPagoEmitida(ByVal sCuenta As String, ByVal nNumOP As Long) As Recordset
Dim sSql As String
Dim rsOP As Recordset
sSql = "Select Top 1 cEstado, K.cConsDescripcion Estado FROM DocRecOPEst E INNER JOIN DocRecOP D ON E.cTpoDoc = D.cTpoDoc " _
    & "AND E.cNroDoc = D.cNroDoc INNER JOIN " & sDBComunes & "Constante K ON E.cEstado = K.nConsValor Where " _
    & "D.cCtaCod = '" & sCuenta & "' And D.cTpoDoc = " & TpoDocOrdenPago & " And E.cNroDoc = '" & nNumOP & "' " _
    & "And K.nConsCod = " & gCaptacOrdPagoEstado & " Order by E.cMovNro DESC"
Set rsOP = dbCmact.CargaRecordSet(sSql)
Set GetEstadoOrdenPagoEmitida = rsOP
Set rsOP = Nothing
End Function

Public Function AgregaMovNoContable(ByVal sMovNro As String, ByVal sDesc As String, ByVal nOperacion As CaptacOperacion)
Dim sSql As String
sSql = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,cMovEstado,cMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sDesc & "','" & gMovEstContabNoContable & "',NULL)"
dbCmact.Ejecutar sSql
End Function

Public Function AgregaOrdenPagoRecibida(ByVal sCuenta As String, ByVal nNumOP As Long, _
        ByVal nMonto As Double, Optional cInstitucion As String = "")
Dim sSql As String
If cInstitucion = "" Then
    sSql = "INSERT DocRecOP (cTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers) " _
        & "VALUES (" & TpoDocOrdenPago & ",'" & nNumOP & "','" & sCuenta & "'," & nMonto & ",NULL)"
Else
    sSql = "INSERT DocRecOP (cTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers) " _
        & "VALUES ('" & TpoDocOrdenPago & "','" & nNumOP & "','" & sCuenta & "'," & nMonto & ",'" & cInstitucion & "')"
End If
dbCmact.Ejecutar sSql
End Function

Public Function AnulaOrdenPagoEmitida(ByVal sCuenta As String, ByVal nNumOP As Long, _
            ByVal sMovNro As String, ByVal nMonto As Double)
Dim sSql As String
sSql = "INSERT DocRecOpEst (cTpoDoc,cNroDoc,cCtaCod,cMovNro,nMonto,cEstado) " _
    & "VALUES (" & TpoDocOrdenPago & ",'" & nNumOP & "','" & sCuenta & "','" & sMovNro & "'," & nMonto & "," & gCapOPEstAnulada & ")"
dbCmact.Ejecutar sSql
End Function

Public Sub AgregaOrdenPagoEmitidas(ByVal sCuenta As String, ByVal nInicio As Long, _
        ByVal nFin As Long, ByVal sMovNro As String)

Dim sSql As String, sDesc As String
sDesc = "Emisión Orden Pago Del " & nInicio & " Al " & nFin & ". Cuenta " & sCuenta
sSql = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,cMovEstado,cMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & gAhoOPEmision & "','" & sDesc & "','" & gMovEstContabNoContable & "',NULL)"
dbCmact.Ejecutar sSql
sSql = "INSERT MovDocEmitidoRango (cCtaCod,nInicio,nFin,cMovNro) " _
    & "VALUES ('" & sCuenta & "'," & nInicio & "," & nFin & ",'" & sMovNro & "')"
dbCmact.Ejecutar sSql
End Sub

Public Sub AgregaTarjetaEstado(ByVal sTarjeta As String, ByVal sMovNro As String, _
        ByVal nEstado As CaptacTarjetaEstado, ByVal sGlosa As String)

Dim sSql As String
sSql = "INSERT TarjetaEstado (cTarjCod,cMovNro,cEstado,cDescripcion) " _
    & "VALUES ('" & sTarjeta & "','" & sMovNro & "'," & nEstado & ",'" & sGlosa & "')"
dbCmact.Ejecutar sSql
End Sub

Public Sub ActualizaTarjetaEstado(ByVal sTarjeta As String, ByVal nEstado As CaptacTarjetaEstado)
Dim sSql As String
sSql = "Update Tarjeta Set cEstado = " & nEstado & " " _
    & "WHERE cTarjCod = '" & sTarjeta & "'"
dbCmact.Ejecutar sSql
End Sub

Public Function GetTarjetaEstadoHist(ByVal sTarjeta As String) As Recordset
Dim sSql As String
Dim rsTarj As Recordset

sSql = "SELECT CONVERT(Varchar(10),CONVERT(Datetime,LEFT(cMovNro,8)),103) Fecha, K.cConsDescripcion Estado, cDescripcion, " _
    & "RIGHT(cMovNro,4) Usu FROM Tarjeta T INNER JOIN TarjetaEstado E ON T.cTarjCod = E.cTarjCod INNER JOIN " _
    & sDBComunes & "Constante K ON E.cEstado = K.nConsValor WHERE T.cTarjCod = '" & sTarjeta & "' AND K.nConsCod = " _
    & gCaptacTarjetaEstado & " ORDER BY Fecha DESC"

Set rsTarj = dbCmact.CargaRecordSet(sSql)
Set GetTarjetaEstadoHist = rsTarj
Set rsTarj = Nothing
End Function

Public Function GetNombreTitulares(ByVal sCuenta As String) As String
Dim sSql As String
Dim rsTit As ADODB.Recordset
Dim sCadena As String
Dim sCadAux As String
Dim sCadAux1 As String
Dim sCadAux2 As String

Dim nPosApe As Long
Dim nPosNom  As Long
Dim sCadAux3 As String
Dim nContador As Long

Dim lstOpe As String * 30

sSql = "SELECT P.cPersNombre FROM " & sDBPersona & "Persona P INNER JOIN ProductoPersona PP " _
    & "INNER JOIN Producto D ON PP.cCtaCod = D.cCtaCod ON P.cPersCod = PP.cPersCod " _
    & "WHERE D.cCtaCod = '" & sCuenta & "' AND PP.nPrdPersRelac = " & gCapRelPersTitular
Set rsTit = dbCmact.CargaRecordSet(sSql)
If rsTit.EOF And rsTit.BOF Then
    GetNombreTitulares = ""
Else
    nContador = 1
    lstOpe = PstaNombre(rsTit("cPersNombre"), False)
    sCadAux = Trim(lstOpe)
    sCadAux3 = rsTit("cPersNombre")
    rsTit.MoveNext
    If Not rsTit.EOF Then sCadAux = ""
    Do While Not rsTit.EOF And nContador <> 4
        If sCadAux <> "" Then
            sCadAux3 = rsTit("cPersNombre")
        End If
        nPosApe = InStr(1, sCadAux3, "/", vbTextCompare)
        nPosNom = InStr(1, sCadAux3, ",", vbTextCompare)
        sCadAux1 = Left(Left(sCadAux3, nPosApe), 18)
        sCadAux2 = Left(Mid(sCadAux3, nPosNom + 1), 18)
        nPosNom = InStr(1, Trim(sCadAux2), " ", vbTextCompare)
        If nPosNom <> 0 Then
            sCadAux2 = Mid(sCadAux2, 1, nPosNom)
        End If
        sCadAux2 = Trim(sCadAux2)
        If sCadAux <> "" Then
            If nContador Mod 2 = 0 Then
                sCadAux = sCadAux & "*" & Left(sCadAux1 & sCadAux2, 20) & Space(20 - Len(Left(sCadAux1 & sCadAux2, 20)))
            Else
                sCadAux = sCadAux & "," & Left(sCadAux1 & sCadAux2, 20) & Space(20 - Len(Left(sCadAux1 & sCadAux2, 20)))
            End If
            nContador = nContador + 1
            rsTit.MoveNext
        Else
            sCadAux = Left(sCadAux1 & sCadAux2, 20) & Space(23 - Len(Left(sCadAux1 & sCadAux2, 20)))
        End If
    Loop
    GetNombreTitulares = sCadAux
End If
rsTit.Close
Set rsTit = Nothing
End Function

Public Function GetCTSPeriodo() As Recordset
Dim sSql As String
Dim rsPer As Recordset
sSql = "Select nItem,nMesInicio,nMesFin,cDescripcion,nPorcentaje FROM CaptacCTSPeriodo " _
    & "Order by nItem"
Set rsPer = dbCmact.CargaRecordSet(sSql)
Set GetCTSPeriodo = rsPer
Set rsPer = Nothing
End Function

Public Function GetSaldoFecha(ByVal sCuenta As String, ByVal dFecha As Date) As Recordset
Dim sSql As String, sNumMov As String
Dim rsCta As Recordset
sSql = "Select TOP 1 CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha, " _
    & "C.nSaldoDisponible, C.nSaldoContable FROM Mov M INNER JOIN MovCap C " _
    & "INNER JOIN MovCapDet CD INNER JOIN OpeTpo O INNER JOIN CapMovTipo CT ON " _
    & "O.cOpeCod = CT.cOpeCod ON CD.cOpeCod = O.cOpeCod ON C.nMovNro = CD.nMovNro ON " _
    & "M.nMovNro = C.nMovNro WHERE C.cCtaCod = '" & sCuenta & "' AND CD.nConceptoCod " _
    & "IN (" & gConcCapital & ") AND M.cMovNro <= '" & Format$(dFecha, "yyyymmdd") & "' " _
    & "AND M.nMovFlag NOT IN (" & gMovFlagExtornado & ") ORDER BY M.nMovNro DESC"
Set rsCta = dbCmact.CargaRecordSet(sSql)
Set GetSaldoFecha = rsCta
Set rsCta = Nothing
End Function
'By Capi 04022008 para obtener los plazos fijos interes cash
Public Function ObtenerPlazoFijoIntCash(ByVal psOpcion As String, ByVal psMoneda As String)
    Dim sSql As String
    Dim oCon As New DConecta
    Set oCon = New DConecta
    oCon.AbreConexion
    sSql = " Exec stp_sel_ReporteDetPlazoFijoInteresCash '" & psOpcion & "','" & psMoneda & "'"
    Set ObtenerPlazoFijoIntCash = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
 End Function


Public Function GetMovimientosCuenta(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0) As Recordset

Dim sSql As String, sNumMov As String
sNumMov = ""
If nNumMov > 0 Then sNumMov = " TOP " & nNumMov + 1 & " "
sSql = "Select " & sNumMov & " CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha, " _
    & "O.cOpeDesc Operacion, ISNULL(MD.cDocNro,'') cDocumento , nAbono = CASE " _
    & "WHEN CT.nCapMovTpo IN (" & gCapMovApertura & "," & gCapMovDeposito & "," & gCapMovIntCap & ") THEN ABS(CD.nMonto) " _
    & "ELSE 0 END, " _
    & "nCargo = CASE WHEN CT.nCapMovTpo IN (" & gCapMovRetiroInt & "," & gCapMovRetiroInac & "," & gCapMovCancelAct & ") THEN ABS(CD.nMonto)*-1 " _
    & "ELSE 0 END, " _
    & "C.nSaldoContable, A.cAgeDescripcion cAgencia " _
    & "FROM MovDoc MD RIGHT JOIN Mov M INNER JOIN MovCap C INNER JOIN MovCapDet CD INNER JOIN " _
    & "OpeTpo O INNER JOIN CapMovTipo CT ON O.cOpeCod = CT.cOpeCod ON " _
    & "CD.cOpeCod = O.cOpeCod ON C.nMovNro = CD.nMovNro ON M.nMovNro = C.nMovNro " _
    & "ON MD.nMovNro = M.nMovNro INNER JOIN Agencias A ON SUBSTRING(M.cMovNro,18,2) = A.cAgeCod " _
    & "WHERE C.cCtaCod = '" & sCuenta & "' AND CD.nConceptoCod IN (" & gConcCapital & ")"
If nNumMov > 0 Then
    sSql = sSql & " ORDER BY M.nMovNro DESC"
Else
    sSql = sSql & " AND LEFT(M.cMovNro,8) BETWEEN '" & Format$(dFecIni, "yyyymmdd") & "' " _
        & "AND '" & Format$(dFecFin, "yyyymmdd") & "'"
End If
Set GetMovimientosCuenta = dbCmact.CargaRecordSet(sSql)

End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim sConn As String
Dim ClsIni As DClassIni
Set ClsIni = New DClassIni
sDBComunes = ClsIni.BaseComunes
sDBPersona = ClsIni.BasePersonas
sDBImagenes = ClsIni.BaseImagenes
Set ClsIni = Nothing
Set dbCmact = New DConecta
dbCmact.AbreConexion
End Sub

Private Sub Class_Terminate()
dbCmact.CierraConexion
Set dbCmact = Nothing
End Sub

'EJVG20131122 ***
Public Function CuentaEsValida(ByVal psCtaCod As String, ByVal psPersCod As String) As Boolean
    Dim oConectx As New DConecta
    Dim sSql As String
    Dim rsCta As New ADODB.Recordset
    sSql = "Exec stp_sel_ERS0622013_CtaEsValida '" & psCtaCod & "','" & psPersCod & "'"
    oConectx.AbreConexion
    Set rsCta = oConectx.CargaRecordSet(sSql)
    oConectx.CierraConexion
    If Not rsCta.EOF Then
        CuentaEsValida = rsCta!bCtaEsValida
    End If
    Set rsCta = Nothing
    Set oConectx = Nothing
End Function
'END EJVG *******
'PASIERS1242014
Public Function GetMovNegxExtornoPagoProveedor(ByVal psOpeCod As String, ByVal pnMovNro As Long) As Recordset
Dim sSql As String
Dim rsCta As Recordset
    sSql = "stp_sel_ERS1242014_ObtenerMovNegxExtornoPagoProveedor '" & psOpeCod & "'," & pnMovNro
    Set rsCta = dbCmact.CargaRecordSet(sSql)
    Set GetMovNegxExtornoPagoProveedor = rsCta
    Set rsCta = Nothing
End Function
Public Function ObtieneMovIniDevSUNAT(ByVal pnMovNro As Long) As Recordset
    Dim sSql As String
    Dim rsMov As Recordset
    sSql = "stp_sel_ERS1242014_ObtieneMovIniDevSUNAT " & pnMovNro
    Set rsMov = dbCmact.CargaRecordSet(sSql)
    If Not rsMov.EOF And Not rsMov.BOF Then
        Set ObtieneMovIniDevSUNAT = rsMov
        Exit Function
    End If
    Set ObtieneMovIniDevSUNAT = Nothing
    Set rsMov = Nothing
End Function
'end PASI***********
