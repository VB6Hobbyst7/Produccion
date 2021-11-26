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

Public dbCmact As Connection
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String

Public Function GetnMovNro(ByVal sMovNro As String) As Long
Dim rsMov As Recordset
Dim sSql As String
sSql = "Select nMovNro From Mov Where cMovNro = '" & sMovNro & "'"
Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
If rsMov.EOF And rsMov.BOF Then
    GetnMovNro = 0
Else
    GetnMovNro = rsMov("nMovNro")
End If
rsMov.Close
Set rsMov = Nothing
End Function

Public Function GetNroOPeradoras(ByVal sCodage As String) As Long
Dim rsMov As Recordset
Dim sSql As String

sSql = " SELECT NRP=IsNULL(NRP,0) FROM AGENCIAS WHERE CAGECOD='" & sCodage & "'"

Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
If rsMov.EOF And rsMov.BOF Then
    GetNroOPeradoras = 0
Else
    GetNroOPeradoras = rsMov("NRP")
End If
rsMov.Close
Set rsMov = Nothing

End Function


Public Function GetTpoActI() As Recordset
Dim RSTEMP As Recordset
Dim sSql As String


    sSql = " Select nconsvalor,cconsdescripcion from constante where nconscod='2027' and nconsvalor<>'2027' order by cconsdescripcion "
    Set RSTEMP = New Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    RSTEMP.ActiveConnection = Nothing
    Set GetTpoActI = RSTEMP
    Set RSTEMP = Nothing

End Function
Public Function GetTpoDocActI() As Recordset
Dim RSTEMP As Recordset
Dim sSql As String


    sSql = " Select nconsvalor,cconsdescripcion from constante where nconscod='2028' and nconsvalor<>'2028' order by cconsdescripcion "
    Set RSTEMP = New Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    RSTEMP.ActiveConnection = Nothing
    Set GetTpoDocActI = RSTEMP
    Set RSTEMP = Nothing

End Function

Public Function GetDPTO() As Recordset
Dim RSTEMP As Recordset
Dim sSql As String


    sSql = " select cubigeocod,cubigeodes from ubigeo where cubigeocod like '1%' order by cubigeodes "
    Set RSTEMP = New Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    RSTEMP.ActiveConnection = Nothing
    Set GetDPTO = RSTEMP
    Set RSTEMP = Nothing

End Function

Public Function GetPROV(Optional sDepto As String) As Recordset
Dim RSTEMP As Recordset
Dim sSql As String


    sSql = " select cubigeocod,cubigeodes from ubigeo where cubigeocod like '2" & sDepto & "%' order by cubigeodes"
    Set RSTEMP = New Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    RSTEMP.ActiveConnection = Nothing
    Set GetPROV = RSTEMP
    Set RSTEMP = Nothing

End Function



Public Function GetDatosOpeLLamada(ByVal nMovNro As Long, ByVal sOperacion As String) As Recordset
Dim RSTEMP As Recordset
Dim sSql As String


    sSql = "Select mc.cctacod,mc.nmonto,mr.nmovnroref from movref mr join movcap mc on mc.nmovnro=mr.nmovnroref where mr.nmovnro=" & nMovNro & " and copecod='" & sOperacion & "'"
    Set RSTEMP = New Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    RSTEMP.ActiveConnection = Nothing
    Set GetDatosOpeLLamada = RSTEMP
    Set RSTEMP = Nothing
        
End Function


Public Function GetCuentasPersonaAI(Optional ByVal sPersCod As String) As Recordset
 Dim RSTEMP As Recordset
 Dim sSql As String
     sSql = " select pp.cctacod,PROD=CASE WHEN SUBSTRING(PP.CCTACOD,6,3)='232'  THEN 'AHORROS' WHEN SUBSTRING(PP.CCTACOD,6,3)='233' THEN 'PLAZO FIJO' ELSE 'CTS' END "
     sSql = sSql & " ,estado=k1.cconsdescripcion,relacion=k2.cconsdescripcion  "
     sSql = sSql & " from productopersona pp "
     sSql = sSql & " join producto p on p.cctacod=pp.cctacod "
     sSql = sSql & " join ( select * from constante where nconscod='2001' ) k1 on k1.nconsvalor=p.nprdestado "
     sSql = sSql & " join ( select * from constante where nconscod='2005' ) k2 on k2.nconsvalor=pp.nprdpersrelac "
     sSql = sSql & " where pp.cperscod='" & sPersCod & "' AND PP.CCTACOD LIKE '108__23[234]%' "
     sSql = sSql & " and p.nprdestado not in (1300,1400) "
     sSql = sSql & " UNION "
     sSql = sSql & " select pp.cctacod,PROD=K3.cconsdescripcion  "
     sSql = sSql & " ,estado=k1.cconsdescripcion,relacion=k2.cconsdescripcion "
     sSql = sSql & " from productopersona pp "
     sSql = sSql & " join producto p on p.cctacod=pp.cctacod "
     sSql = sSql & " join ( select * from constante where nconscod='3001' ) k1 on k1.nconsvalor=p.nprdestado "
     sSql = sSql & " join ( select * from constante where nconscod='3002' ) k2 on k2.nconsvalor=pp.nprdpersrelac "
     sSql = sSql & " JOIN ( select * from constante where nconscod='1001' ) k3 on k3.nconsvalor=SUBSTRING(PP.CCTACOD,6,3) "
     sSql = sSql & " where pp.cperscod='" & sPersCod & "' AND PP.CCTACOD LIKE '108__[1234][012]%' "
     sSql = sSql & " AND PP.CCTACOD NOT LIKE  '108__23[234]%' AND PP.CCTACOD NOT LIKE  '108__305%' "
     sSql = sSql & " AND PP.nprdpersrelac  IN (20,21,22,23,24,25) "
     sSql = sSql & " and p.nprdestado  in (2020,2021,2022,2201,2030,2031,2032,2205)  "
     sSql = sSql & " UNION "
     sSql = sSql & "  select pp.cctacod,PROD='PRENDARIO' "
     sSql = sSql & " ,estado=k1.cconsdescripcion,relacion=k2.cconsdescripcion "
     sSql = sSql & " from productopersona pp "
     sSql = sSql & "   join producto p on p.cctacod=pp.cctacod "
     sSql = sSql & " join ( select * from constante where nconscod='3001' ) k1 on k1.nconsvalor=p.nprdestado "
     sSql = sSql & " join ( select * from constante where nconscod='3002' ) k2 on k2.nconsvalor=pp.nprdpersrelac "
     sSql = sSql & " where pp.cperscod='" & sPersCod & "' AND PP.CCTACOD LIKE  '108__305%' "
     sSql = sSql & " AND PP.nprdpersrelac=20 "
     sSql = sSql & " and p.nprdestado  in (2101,2104,2106,2107) "

             
             
             
             
     Set RSTEMP = New Recordset
     RSTEMP.CursorLocation = adUseClient
     RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
     RSTEMP.ActiveConnection = Nothing
     Set GetCuentasPersonaAI = RSTEMP
     Set RSTEMP = Nothing
 
  
End Function


Public Function GetDatosCuentaAho(ByVal sCuenta As String) As Recordset
Dim sSql As String
Dim rsCta As Recordset

sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, AG.cAgeDescripcion, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.bOrdPag, A.dUltContacto, T.cConsDescripcion cEstado, " _
    & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.bInactiva, A.bInmovilizada, A.nSobregiro,c.nfirmasmin,c.calias,nTpoPrograma=isnull(c.nTpoPrograma,0) FROM Producto P " _
    & "INNER JOIN Captaciones C INNER JOIN CaptacAhorros A ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod " _
    & "WHERE P.cCtaCod = '" & sCuenta & "' AND T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    & "AND T2.nConsCod = " & gCaptacTipoTasa

Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set GetDatosCuentaAho = rsCta

Set rsCta.ActiveConnection = Nothing
Set rsCta = Nothing
End Function

Public Function GetDatosCuentaPF(ByVal sCuenta As String) As Recordset
Dim sSql As String
Dim rsCta As Recordset

sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, AG.cAgeDescripcion, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.nPlazo, A.nIntPag, A.dRenovacion, A.nApertura, " _
    & "A.nFormaRetiro, A.dAuxiliar, A.nDuplicado, T.cConsDescripcion cEstado, A.bBusCredPend, " _
    & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, T3.cConsDescripcion cRetiro, " _
    & "A.dUltCierreAnt, A.dAuxiliarAnt, ISNULL(CI.cCtaCodAbono,'') cCtaCodAbono,a.nformaretiro,c.calias,c.nfirmasmin FROM Producto P " _
    & "INNER JOIN Captaciones C INNER JOIN CaptacPlazoFijo A LEFT JOIN CaptacCtaAboIntPF CI " _
    & "ON A.cCtaCod = CI.cCtaCod ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Constante T3 ON A.nFormaRetiro = T3.nConsValor " _
    & "INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod WHERE P.cCtaCod = '" & sCuenta & "' AND " _
    & "T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    & "AND T2.nConsCod = " & gCaptacTipoTasa & " AND T3.nConsCod = " & gCaptacPFFormaRetiro


Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaPF = rsCta
Set rsCta = Nothing
End Function

Public Function GetCuentasNoAsociadas(sPersCod) As Recordset
Dim sSql As String
Dim rsCta As Recordset

sSql = "    SELECT DISTINCT CUENTA=PP.CCTACOD,PRODUCTO=case when substring(PP.CCTACOD,6,3)='232' then 'AHORROS' "
sSql = sSql & "    WHEN  substring(PP.CCTACOD,6,3)='233' THEN 'PLAZO FIJO'  "
sSql = sSql & "    WHEN  substring(PP.CCTACOD,6,3)='234' THEN 'CTS' END, "
sSql = sSql & "    MONEDA=CASE WHEN SUBSTRING(PP.CCTACOD,9,1)=1 THEN 'SOLES' ELSE 'DOLARES' END , "
sSql = sSql & "    ESTADO=k.CCONSDESCRIPCION,PERSONERIA=C.NPERSONERIA  "
sSql = sSql & "    FROM PRODUCTOPERSONA PP "
sSql = sSql & "    JOIN PRODUCTO P ON P.CCTACOD=PP.CCTACOD "
sSql = sSql & "    JOIN CAPTACIONES C ON C.CCTACOD=PP.CCTACOD "
sSql = sSql & "    JOIN (SELECT * FROM CONSTANTE WHERE NCONSCOD=2001) K ON K.NCONSVALOR=P.NPRDESTADO "
sSql = sSql & "    LEFT JOIN CUENTATARJ  CT ON CT.CCTACOD=PP.CCTACOD "
sSql = sSql & "    WHERE CT.CCTACOD IS NULL AND "
sSql = sSql & "    PP.CPERSCOD='" & sPersCod & "' AND PP.NPRDPERSRELAC=10 AND "
sSql = sSql & "    PP.CCTACOD LIKE '108__23[234]%'AND "
sSql = sSql & "    PP.NPRDPERSRELAC in (10,12)  AND P.NPRDESTADO NOT IN (1300,1400)"
sSql = sSql & "    ORDER BY PP.CCTACOD "
    
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetCuentasNoAsociadas = rsCta
Set rsCta = Nothing
End Function

Public Function GetPersonasCuentaNA(ByVal sCuenta As String, ByVal sPersCod As String, Optional ByVal nPersoneria As Integer) As Recordset
Dim sSql As String, sqlaux As String
Dim rsCta As Recordset

If nPersoneria = 1 Then
    sqlaux = " and pP.nprdpersrelac=" & CaptacRelacPersona.gCapRelPersTitular
Else
    'sqlaux = " and pP.nprdpersrelac in (" & CaptacRelacPersona.gCapRelPersTitular & "," & CaptacRelacPersona.gCapRelPersRepTitular & ")"
    sqlaux = " and pP.nprdpersrelac in (" & CaptacRelacPersona.gCapRelPersTitular & ",12)"
End If

sSql = " Select CODIGO=PP.cPersCod, NOMBRE=P.cPersNombre , "
sSql = sSql & " RELACION=K.CCONSDESCRIPCION, "
sSql = sSql & " PERSONERIA=P.nPersPersoneria "
'ssql = ssql & " Obligatorio=case when pp.cobligatorio='S' then 'SI' when (pp.cobligatorio='N' or pp.cobligatorio is null) then 'NO' else 'OPCIONAL' end "
sSql = sSql & " FROM Persona P "
sSql = sSql & " INNER JOIN ProductoPersona PP ON PP.CPERSCOD=P.CPERSCOD "
sSql = sSql & " JOIN (SELECT * FROM CONSTANTE WHERE NCONSCOD=2005 ) K ON K.NCONSVALOR=PP.NPRdpersrelac "
sSql = sSql & " WHERE PP.cCtaCod = '" & sCuenta & "'  AND PP.CPERSCOD<>'" & sPersCod & "' " & sqlaux
sSql = sSql & " Order by PP.nPrdPersRelac "
    
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set GetPersonasCuentaNA = rsCta
Set rsCta.ActiveConnection = Nothing
Set rsCta = Nothing
End Function


Public Function EsCtaConvenio(ByVal sCuenta As String) As Boolean
  Dim sSql As String, rsCta As Recordset
  
EsCtaConvenio = False
  
  sSql = " select cperscod,cctacod from captacconveniocta " _
         & " where ntpocuenta=1 and cctacod='" & sCuenta & "'"
    
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
 If Not rsCta.EOF Then
   If Trim(rsCta!cCtaCod) = Trim(sCuenta) Then EsCtaConvenio = True
 End If

Set rsCta.ActiveConnection = Nothing
Set rsCta = Nothing
End Function


Public Function GetDatosCuentaCTS(ByVal sCuenta As String) As Recordset
Dim sSql As String
Dim rsCta As Recordset

sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas,c.nfirmasmin ,C.nIntAcum, dUltCierre=( case when dbo.ultimodiames(dUltCierre)=1 then dUltCierre else dateadd(dd,1, dUltCierre) end )  , C.dApertura, AG.cAgeDescripcion, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.nSaldRetiro, A.nIntSaldo, PE.cPersNombre cInstitucion, " _
    & "T.cConsDescripcion cEstado, T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.cCodInst,C.cAlias " _
    & "FROM Producto P INNER JOIN Captaciones C INNER JOIN CaptacCTS A INNER JOIN " & sDBPersona & "" _
    & "Persona PE ON A.cCodInst = PE.cPersCod ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod " _
    & "WHERE P.cCtaCod = '" & sCuenta & "' AND T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
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

sSql = " Select PP.cPersCod, P.cPersNombre Nombre, (UPPER(RTRIM(T.cConsDescripcion)) " _
    & " + SPACE(75 - len((UPPER(RTRIM(T.cConsDescripcion))))) + CONVERT(Varchar(2),PP.nPrdPersRelac)) cRelacion, P.nPersPersoneria, Obligatorio=case when pp.cobligatorio='S' then 'SI' when (pp.cobligatorio='N' or pp.cobligatorio is null) then 'NO' else 'OPCIONAL' end   " _
    & " FROM " & sDBPersona & "Persona P INNER JOIN ProductoPersona PP ON " _
    & " P.cPersCod = PP.cPersCod INNER JOIN " & sDBComunes & "Constante T " _
    & " ON PP.nPrdPersRelac = T.nConsValor WHERE PP.cCtaCod = '" & sCuenta & "' " _
    & " AND T.nConsCod = " & gCaptacRelacPersona & "" _
    & " Order by PP.nPrdPersRelac "
    
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetProductoPersona = rsCta
Set rsCta = Nothing
End Function

Public Function GetCuentasPersona(ByVal sPers As String, Optional nProd As Producto, _
        Optional bActivas As Boolean = False, Optional bBloqueadas As Boolean = False, _
        Optional nMoneda As Moneda, Optional bOrdPag As Boolean = False, _
        Optional sAgencia As String = "") As Recordset
Dim sSql As String
Dim rsCta As Recordset

If nProd = gCapAhorros Then
    sSql = "Select PP.cCtaCod, (UPPER(T3.cConsDescripcion) + ' - ' + T1.cConsDescripcion) cDescripcion, nNivel = 1, " _
        & "T1.cConsDescripcion cRelacion, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, c.npersoneria, " _
        & "UPPER(T2.cConsDescripcion) cProducto, UPPER(T3.cConsDescripcion) cMoneda FROM ProductoPersona PP INNER JOIN Producto P INNER " _
        & "JOIN Captaciones C INNER JOIN CaptacAhorros A ON C.cCtaCod = A.cCtaCod ON P.cCtaCod = C.cCtaCod ON PP.cCtaCod = P.cCtaCod INNER JOIN " _
        & sDBComunes & "Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN " & sDBComunes & "" _
        & "Constante T1 ON PP.nPrdPersRelac = T1.nConsValor INNER JOIN " & sDBComunes & "Constante T2 " _
        & "ON SUBSTRING(PP.cCtaCod,6,3) = CONVERT(Varchar(3),T2.nConsValor) INNER JOIN " & sDBComunes & "" _
        & "Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor) WHERE PP.cPersCod = '" & sPers & "' " _
        & "AND T1.nConsCod = " & gCaptacRelacPersona & " AND T.nConsCod = " & gCaptacEstado & " AND " _
        & "T2.nConsCod = " & gProducto & " AND T3.nConsCod = " & gMoneda
Else
    sSql = "Select PP.cCtaCod, (UPPER(T3.cConsDescripcion) + ' - ' + T1.cConsDescripcion) cDescripcion, nNivel = 1, " _
        & "T1.cConsDescripcion cRelacion, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, c.npersoneria, " _
        & "UPPER(T2.cConsDescripcion) cProducto, UPPER(T3.cConsDescripcion) cMoneda FROM ProductoPersona PP INNER JOIN Producto P INNER " _
        & "JOIN Captaciones C ON P.cCtaCod = C.cCtaCod ON PP.cCtaCod = P.cCtaCod INNER JOIN " _
        & sDBComunes & "Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN " & sDBComunes & "" _
        & "Constante T1 ON PP.nPrdPersRelac = T1.nConsValor INNER JOIN " & sDBComunes & "Constante T2 " _
        & "ON SUBSTRING(PP.cCtaCod,6,3) = CONVERT(Varchar(3),T2.nConsValor) INNER JOIN " & sDBComunes & "" _
        & "Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor) WHERE PP.cPersCod = '" & sPers & "' " _
        & "AND T1.nConsCod = " & gCaptacRelacPersona & " AND T.nConsCod = " & gCaptacEstado & " AND " _
        & "T2.nConsCod = " & gProducto & " AND T3.nConsCod = " & gMoneda
End If

If nMoneda <> 0 Then
    sSql = sSql & " And PP.cCtaCod LIKE '________" & Trim(nMoneda) & "%'"
End If
If nProd <> 0 Then
    sSql = sSql & " AND SUBSTRING(PP.cCtaCod,6,3) = '" & nProd & "'"
End If
If bActivas Then
    sSql = sSql & " AND P.nPrdEstado NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ")"
End If
If bBloqueadas Then
    sSql = sSql & " AND P.nPrdEstado NOT IN (" & gCapEstBloqRetiro & "," & gCapEstBloqTotal & ")"
End If
If bOrdPag Then
    sSql = sSql & " AND A.bOrdPag = " & IIf(bOrdPag, 1, 0) & " "
End If
If sAgencia <> "" Then
    sSql = sSql & " AND P.cCtaCod LIKE '___" & sAgencia & "%' "
End If
sSql = sSql & " ORDER BY PP.cCtaCod"
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetCuentasPersona = rsCta
Set rsCta = Nothing
End Function

Public Function GetCuentasAhorro(ByVal bTodas As Boolean, Optional cAgeCod As String = "", Optional cCtaCod As String = "") As Recordset
Dim sSql As String, sSqlaux As String
Dim rsCta As Recordset

sSqlaux = ""
    If bTodas = True And cAgeCod <> "" Then
        sSqlaux = " and substring(p.cctacod,4,2)='" & cAgeCod & "'"
    End If
    
    If cCtaCod <> "" Then
        sSqlaux = " and p.cctacod='" & cCtaCod & "'"
    End If
    
    

  sSql = " Select p.Cctacod,  "
  sSql = sSql & " cpersnombre=( Select top  1  f.cpersnombre from persona f join  productopersona pp  on pp.cperscod=f.cperscod    where pp.nprdpersrelac=" & gCapRelPersTitular & "  and pp.cctacod=p.cctacod ),  "
  sSql = sSql & " exonerada=(case when ca.bnocobroinactivas is null then 0 else case when ca.bnocobroinactivas=0 then 0 else 1 end end) "
  sSql = sSql & " from Producto p join Captaciones c on c.cctacod=p.cctacod join CaptacAhorros ca on ca.cctacod=p.cctacod  "
  sSql = sSql & " where p.nprdestado not in (" & gCapEstAnulada & "," & gCapEstCancelada & " ) " & sSqlaux
  sSql = sSql & " ORDER BY CPERSNOMBRE "

  Set rsCta = New Recordset
  rsCta.CursorLocation = adUseClient
  rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
  Set rsCta.ActiveConnection = Nothing
  Set GetCuentasAhorro = rsCta
  Set rsCta = Nothing
  
End Function
Public Sub ActualizaCtaDescInact(ByVal psCuenta As String, ByVal bdescinactiva As Boolean, ByVal psMovNro As String)
Dim sSql As String
     
        sSql = "Update Captacahorros set bnocobroinactivas=" & IIf(bdescinactiva = False, 0, 1) & " , cnocobroinactivas='" & psMovNro & "'"
        sSql = sSql & " where cctacod='" & psCuenta & "'"
        dbCmact.Execute sSql

End Sub


Public Sub EliminaCuentaTarjPersona(ByVal sPersona As String, ByVal sTarjeta As String)
Dim sSql As String
sSql = "Delete CuentaTarj Where cPersCod = '" & sPersona & "' AND cTarjCod = '" & sTarjeta & "'"
dbCmact.Execute sSql
End Sub

Public Sub AgregaCuentaTarjPersona(ByVal sCuenta As String, ByVal sPersona As String, ByVal sTarjeta As String)
Dim sSql As String

sSql = "INSERT CuentaTarj (cCtaCod,cPersCod,cTarjCod) " _
    & "VALUES ('" & sCuenta & "','" & sPersona & "','" & sTarjeta & "')"

dbCmact.Execute sSql
End Sub

Public Sub EliminaProductoPersona(ByVal sCuenta As String)
Dim sSql As String
sSql = "Delete ProductoPersona Where cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub AgregaProductoPersona(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona, Optional ByVal COBLIGATORIO As String = "N")
Dim sSql As String

sSql = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac,cObligatorio) " _
    & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ",'" & COBLIGATORIO & "')"

dbCmact.Execute sSql
End Sub

Public Function GetCapBloqueos(ByVal sCuenta As String, ByVal nTipoBloqueo As CaptacTipoBloqueo, ByVal nConstante As ConstanteCabecera) As Recordset
Dim sSql As String
Dim rsRel As Recordset

Set rsRel = New Recordset
rsRel.CursorLocation = adUseClient

sSql = "Select nEstado = CASE WHEN B.cMovNroDbl IS NULL THEN 1 ELSE 0 END, " _
    & "CONVERT(VARCHAR(10),CONVERT(Datetime,LEFT(cMovNro,8),101),105) dFecha, " _
    & "(C.cConsDescripcion + Space(75) + Convert(Varchar(3),C.nConsValor)) Motivo, " _
    & "RIGHT(B.cMovNro,4) Usu, dFechaDbl = CASE WHEN B.cMovNroDbl IS NULL THEN '' ELSE " _
    & "CONVERT(VARCHAR(10),CONVERT(Datetime,LEFT(B.cMovNroDbl,8),101),105) END, " _
    & "UsuDbl = CASE WHEN B.cMovNroDbl IS NULL THEN '' ELSE RIGHT(B.cMovNroDbl,4) END, " _
    & "ISNULL(B.cComentario,'') cComentario, B.cMovNro, Flag = CASE WHEN B.cMovNroDbl IS NULL THEN 'A' ELSE '' END " _
    & "FROM ProductoBloqueos B INNER JOIN Constante C ON B.nBlqMotivo = C.nConsValor WHERE " _
    & "C.nConsCod = " & nConstante & " And B.cCtaCod = '" & sCuenta & "' AND B.nBlqTpo = " & nTipoBloqueo _
    & " Order by B.cMovNro"

rsRel.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsRel.ActiveConnection = Nothing
Set GetCapBloqueos = rsRel
End Function

Public Function NuevoBloqueoRetiro(ByVal sCuenta As String, ByVal nMotivo As CaptacMotBloqueoRet, _
        cComentario As String, ByVal cMovNro As String)
Dim sSql As String
sSql = "INSERT ProductoBloqueos (cCtaCod,nBlqTpo,nBlqMotivo,cComentario,cMovNro,cMovNroDbl) " _
    & "VALUES ('" & sCuenta & "'," & gCapTpoBlqRetiro & "," & nMotivo & ",'" & cComentario & "','" & cMovNro & "',NULL)"
dbCmact.Execute sSql
End Function

Public Function NuevoBloqueoTotal(ByVal sCuenta As String, ByVal nMotivo As CaptacMotBloqueoTot, _
        cComentario As String, ByVal cMovNro As String)
Dim sSql As String
sSql = "INSERT ProductoBloqueos (cCtaCod,nBlqTpo,nBlqMotivo,cComentario,cMovNro,cMovNroDbl) " _
    & "VALUES ('" & sCuenta & "'," & gCapTpoBlqTotal & "," & nMotivo & ",'" & cComentario & "','" & cMovNro & "',NULL)"
dbCmact.Execute sSql
End Function

Public Sub ActualizaEstadoCuenta(ByVal sCuenta As String, nEstado As CaptacEstado)
Dim sSql As String
sSql = "UPDATE Producto Set nPrdEstado = " & nEstado & " WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaBloqueoRet(ByVal sCuenta As String, ByVal sComentario As String, _
    ByVal sMovNroDbl As String, ByVal nMotivo As CaptacMotBloqueoRet, ByVal sMovNro As String)
Dim sSql As String
sSql = "UPDATE ProductoBloqueos Set cComentario = '" & sComentario & "', " _
    & "cMovNroDbl = '" & sMovNroDbl & "' WHERE cCtaCod = '" & sCuenta & "' " _
    & "And nBlqTpo = " & gCapTpoBlqRetiro & " AND nBlqMotivo = " & nMotivo & " AND cMovNro = '" & sMovNro & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaBloqueoTot(ByVal sCuenta As String, ByVal sComentario As String, _
    ByVal sMovNroDbl As String, ByVal nMotivo As CaptacMotBloqueoTot, ByVal sMovNro As String)

Dim sSql As String
sSql = "UPDATE ProductoBloqueos Set cComentario = '" & sComentario & "', " _
    & "cMovNroDbl = '" & sMovNroDbl & "' WHERE cCtaCod = '" & sCuenta & "' " _
    & "And nBlqTpo = " & gCapTpoBlqTotal & " AND nBlqMotivo = " & nMotivo & " AND cMovNro = '" & sMovNro & "'"
dbCmact.Execute sSql
End Sub

Public Function ExisteBloqueoTot(ByVal sCuenta As String, ByVal nMotivo As CaptacMotBloqueoTot) As Boolean
Dim sSql As String
Dim bExiste As Boolean
Dim rsBloq As Recordset
bExiste = False
sSql = "Select cCtaCod FROM ProductoBloqueos WHERE cCtaCod = '" & sCuenta & "' " _
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
sSql = "Select cCtaCod FROM ProductoBloqueos WHERE cCtaCod = '" & sCuenta & "' " _
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
    ByVal nMotivo As CaptacMotBloqueoRet, ByVal sMovNro As String)
Dim sSql As String

sSql = "UPDATE ProductoBloqueos Set cComentario = '" & sComentario & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "' AND nBlqTpo = " & gCapTpoBlqRetiro & " " _
    & "AND nBlqMotivo = " & nMotivo & " AND cMovNro = '" & sMovNro & "'"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaComentBlqTot(ByVal sCuenta As String, ByVal sComentario As String, _
    ByVal nMotivo As CaptacMotBloqueoTot, ByVal sMovNro As String)
Dim sSql As String

sSql = "UPDATE ProductoBloqueos Set cComentario = '" & sComentario & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "' AND nBlqTpo = " & gCapTpoBlqTotal & " " _
    & "AND nBlqMotivo = " & nMotivo & " AND cMovNro = '" & sMovNro & "'"
dbCmact.Execute sSql
End Sub

Public Function GetCapTasaInteres(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
    ByVal nTipoTasa As CaptacTipoTasa, Optional nPlazo As Double = 0, Optional nValor As Double = 0, _
    Optional sCodage As String = "", Optional bOrdPag As Boolean = False) As Double

Dim rsTasa As Recordset
Dim sSql As String
Dim cOrdPag As String
cOrdPag = IIf(bOrdPag, "1", "0")

sSql = "SELECT nTasaValor FROM CaptacTasas WHERE nTasaProd = " & nProducto & " AND " _
    & "nTasaMon = " & nMoneda & " AND nTasaTpo = " & nTipoTasa & " And " & nPlazo & " " _
    & "BETWEEN nPlazoIni AND nPlazoFin And " & nValor & " BETWEEN nValorIni And nValorFin " _
    & "And cOrdPag = '" & cOrdPag & "' And cCodAge = '" & sCodage & "'"

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
    & "AND T.nEstado NOT IN (" & gCapTarjEstCancelada & ") AND P.nPrdEstado NOT IN (" & gCapEstAnulada & "," _
    & gCapEstAnulada & "," & gCapEstCancelada & "," & gCapEstCancelada & ")"
Set rsTarj = New Recordset
rsTarj.CursorLocation = adUseClient
rsTarj.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTarj.ActiveConnection = Nothing
Set GetPersonaTarj = rsTarj
Set rsTarj = Nothing
End Function

Public Function GetTarjetaCuentas(ByVal sTarjeta As String, Optional nProducto As Producto = gCapTodoAhorro) As Recordset
Dim sSql As String
Dim rsTarj As Recordset

Dim sProducto As String
If nProducto = 99 Then
    sProducto = "23[234]"
Else
    sProducto = CStr(nProducto)
End If

sSql = "SELECT distinct CT.cCtaCod, UPPER(K.cConsDescripcion) Producto, UPPER(K1.cConsDescripcion) Moneda, CT.cPersCod, UPPER(K2.cConsDescripcion) Relacion, " _
    & "T.nEstado FROM Producto P INNER JOIN  ProductoPersona PP INNER JOIN CuentaTarj CT INNER JOIN Tarjeta T ON CT.cTarjCod = T.cTarjCod ON PP.cCtaCod = " _
    & "CT.cCtaCod AND PP.cPersCod = CT.cPersCod ON P.cCtaCod = PP.cCtaCod INNER JOIN " & sDBComunes & "Constante K ON SUBSTRING(P.cCtaCod,6,3) " _
    & "= CONVERT(Varchar(3),K.nConsValor) INNER JOIN " & sDBComunes & "Constante K1 ON SUBSTRING(P.cCtaCod,9,1) = CONVERT(Varchar(1),K1.nConsValor) " _
    & "INNER JOIN " & sDBComunes & "Constante K2 ON PP.nPrdPersRelac = K2.nConsValor WHERE " _
    & "P.nPrdEstado NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ") " _
    & "AND T.cTarjCod = '" & sTarjeta & "' And K.nConsCod = " & gProducto & " " _
    & "And K1.nConsCod = " & gMoneda & " AND K2.nConsCod = " & gCaptacRelacPersona & " " _
    & "And Substring(P.cCtaCod,6,3) like '" & IIf(nProducto = gCapTodoAhorro, "%", nProducto) & "'"



'sSql = " SELECT distinct T.cCtaCod, UPPER(K.cConsDescripcion) Producto, UPPER(K1.cConsDescripcion) Moneda, " _
'      & " CPERSCOD=(SELECT TOP 1  CT.cPersCod  FROM CUENTATARJ CT " _
'      & " JOIN PRODUCTOPERSONA PP ON PP.CCTACOD=CT.CCTACOD AND PP.CPERSCOD=CT.CPERSCOD " _
'      & "  WHERE PP.NPRDPERSRELAC in (10,12,11,13) AND CT.CCTACOD=P.CCTACOD) " _
'      & " , UPPER(K2.cConsDescripcion) Relacion, T.nEstado " _
'      & "  FROM Producto P " _
'      & "  INNER JOIN (SELECT DISTINCT CT.CCTACOD, A.CTARJCOD,A.NESTADO FROM  Tarjeta A JOIN CUENTATARJ CT ON CT.CTARJCOD=A.CTARJCOD  ) T  ON T.CCTACOD=P.CCTACOD " _
'      & "  INNER JOIN Constante K ON SUBSTRING(P.cCtaCod,6,3) = CONVERT(Varchar(3),K.nConsValor) " _
'      & "  INNER JOIN Constante K1 ON SUBSTRING(P.cCtaCod,9,1) = CONVERT(Varchar(1),K1.nConsValor) " _
'      & "  INNER JOIN Constante K2 ON  K2.nConsValor in (10)  " _
'      & "      WHERE P.nPrdEstado NOT IN (1300,1400) " _
'      & "     AND T.cTarjCod = '" & sTarjeta & "' And K.nConsCod = 1001 " _
'      & "      And K1.nConsCod = " & gMoneda & " AND K2.nConsCod = 2005 " _
'      & "     And Substring(P.cCtaCod,6,3) like  '" & sProducto & "' "


Set rsTarj = New Recordset
rsTarj.CursorLocation = adUseClient
rsTarj.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTarj.ActiveConnection = Nothing
Set GetTarjetaCuentas = rsTarj
Set rsTarj = Nothing
End Function


Public Sub AgregaTarjeta(ByVal sTarjeta As String, ByVal sClave As String, ByVal dCaduca As Date, ByVal sMovNro As String)
Dim sSql As String
sSql = "INSERT Tarjeta (cTarjCod,cClave,dCaduca,nEstado,cMovNro) VALUES " _
    & "('" & sTarjeta & "','" & sClave & "','" & Format$(dCaduca, "mm/dd/yyyy") & "'," & gCapTarjEstActiva & ",'" & sMovNro & "')"
dbCmact.Execute sSql
sSql = "INSERT TarjetaEstado (cTarjCod,cMovNro,nEstado,cDescripcion) VALUES " _
    & "('" & sTarjeta & "','" & sMovNro & "'," & gCapTarjEstActiva & ",'NUEVA TARJETA')"
dbCmact.Execute sSql
End Sub

Public Sub AgregaCuentaTarj(ByVal sTarjeta As String, ByVal sCuenta As String, ByVal sPersona As String)
Dim sSql As String
sSql = "INSERT CuentaTarj (cCtaCod,cPersCod,cTarjCod) VALUES " _
    & "('" & sCuenta & "','" & sPersona & "','" & sTarjeta & "')"
dbCmact.Execute sSql
End Sub

Public Function ExisteTarjeta(ByVal sTarjeta As String) As Boolean
Dim sSql As String
Dim rsTarj As Recordset
Set rsTarj = New Recordset
rsTarj.CursorLocation = adUseClient
sSql = "SELECT cTarjCod FROM Tarjeta WHERE cTarjCod = '" & sTarjeta & "'"
rsTarj.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTarj.ActiveConnection = Nothing
If rsTarj.EOF And rsTarj.BOF Then
    ExisteTarjeta = False
Else
    ExisteTarjeta = True
End If
rsTarj.Close
Set rsTarj = Nothing
End Function

Public Function EsTarjetaEmitida(ByVal sTarjeta As String) As Boolean
Dim sSql As String
Dim rsTarj As Recordset
Set rsTarj = New Recordset
rsTarj.CursorLocation = adUseClient
sSql = "SELECT cTarjCod FROM TarjetaAux WHERE cTarjCod = '" & sTarjeta & "'"
rsTarj.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTarj.ActiveConnection = Nothing
If rsTarj.EOF And rsTarj.BOF Then
    EsTarjetaEmitida = False
Else
    EsTarjetaEmitida = True
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
Set rsTarj = New Recordset
rsTarj.CursorLocation = adUseClient
sSql = "SELECT cTarjCod FROM CuentaTarj WHERE cTarjCod = '" & sTarjeta & "' AND " _
    & "cPersCod = '" & sPersona & "' AND cCtaCod = '" & sCuenta & "'"
rsTarj.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTarj.ActiveConnection = Nothing
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
    & "DATEDIFF(yy,dPersNacCreac,Getdate()) nEdad, nPersRelacBenefPorc, PR.nPersRelac " _
    & "FROM " & sDBPersona & "Persona P INNER JOIN PersRelaciones PR INNER JOIN Constante T " _
    & "ON PR.nPersRelac = T.nConsValor ON P.cPersCod = PR.cPersRelacPersCod WHERE " _
    & "bPersRelacBenef = 1 AND PR.cPersCod = '" & sPersona & "' And T.nConsCod = " & gPersRelacion
    
Set rsPers = New Recordset
rsPers.CursorLocation = adUseClient
rsPers.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsPers.ActiveConnection = Nothing
Set GetBeneficiarios = rsPers
Set rsPers = Nothing
End Function

Public Function GetDatosPersona(ByVal sPersona As String) As Recordset
Dim sSql As String
Dim rsCta As Recordset
sSql = "Select P.cPersNombre Nombre, P.cPersDireccDomicilio Direccion, ISNUll(U.cUbiGeoDescripcion,'') Zona, " _
    & "ISNUll(P.cPersTelefono,'') Fono, ISNUll(T.cConsDescripcion,'') ID, ISNull(I.cPersIDnro,'') [ID N°], P.cPersCod, P.nPersPersoneria, " _
    & "I.cPersIDtpo FROM " & sDBPersona & "PersId I RIGHT JOIN " & sDBPersona & "Persona P LEFT JOIN " _
    & sDBComunes & "UbicacionGeografica U ON P.cPersDireccUbiGeo = U.cUbiGeoCod ON I.cPersCod = P.cPersCod " _
    & "LEFT JOIN " & sDBComunes & "Constante T ON I.cPersIDTpo = T.nConsValor And T.nConsCod = " & gPersIdTipo & "  WHERE P.cPersCod = '" & sPersona & "' "
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosPersona = rsCta
Set rsCta = Nothing
End Function

Public Function GetDatosOrdenPago(ByVal sCuenta As String, ByVal nNumOP As Long) As Recordset
Dim sSql As String
Dim rsCta As Recordset
sSql = "Select D.cNroDoc, K.cConsDescripcion cDescripcion,  D.nMonto, CONVERT(Varchar(10),CONVERT(Datetime,LEFT(D.cMovNro,8)),103) Fecha, " _
    & "RIGHT(D.cMovNro,4) Usuario, D.nEstado FROM DocRecOP D INNER JOIN " & sDBComunes & "Constante K ON D.nEstado = K.nConsValor " _
    & "WHERE D.cCtaCod = '" & sCuenta & "' AND D.cNroDoc = '" & nNumOP & "' And K.nConsCod = " & gCaptacOrdPagoEstado _
    & " And D.cMovNro IN (Select MAX(cMovNro) From DocRecOP D1 Where D1.cCtaCod = D.cCtaCod And D1.cNroDoc = D.cNroDoc) " _
    & "ORDER BY D.cMovNro DESC"
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosOrdenPago = rsCta
Set rsCta = Nothing
End Function

Public Function GetPersonaCuenta(ByVal sCuenta As String, Optional nTipRel As CaptacRelacPersona = 0) As Recordset

Dim sSql As String
Dim rsCta As Recordset
Dim lbflag As Boolean


sSql = " Select P.cPersNombre Nombre,  ISNULL(T1.cConsDescripcion,'') +  Space(50) +  str(T1.nConsValor) Relacion, ISNULL(P.cPersDireccDomicilio,'') Direccion, ISNULL(U.cUbiGeoDescripcion,'') Zona, " _
    & " P.cPersTelefono Fono, ISNULL(T.cConsDescripcion,'') ID, ISNULL(I.cPersIDnro,'') [ID N°], P.cPersCod, P.nPersPersoneria, " _
    & " I.cPersIDtpo, PP.nPrdPersRelac, PP.cCtaCod,pp.bobligatorio,pp.cobligatorio FROM " & sDBPersona & "PersId I RIGHT JOIN (" & sDBPersona & "Persona P LEFT JOIN " _
    & sDBComunes & "UbicacionGeografica U ON P.cPersDireccUbiGeo = U.cUbiGeoCod) RIGHT JOIN ProductoPersona PP " _
    & "ON P.cPersCod = PP.cPersCod ON I.cPersCod = P.cPersCod LEFT JOIN " & sDBComunes & "Constante T ON I.cPersIDTpo = T.nConsValor   AND T.nConsCod = " & gPersIdTipo & "  " _
    & "LEFT JOIN " & sDBComunes & "Constante T1 ON PP.nPrdPersRelac = T1.nConsValor AND T1.nConsCod = " & gCaptacRelacPersona & " WHERE PP.cCtaCod = '" & sCuenta & "'  and  (PP.nPrdPersRelac = '" & nTipRel & "' or  0=" & nTipRel & ")   " _
    & " ORDER BY PP.nPrdPersRelac,Nombre, I.cPersIDtpo"

Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetPersonaCuenta = rsCta
Set rsCta = Nothing
End Function
Public Function GetDatosOP(Optional ByVal sAgecod As String = "") As Recordset
Dim sSql As String, sqlaux
Dim rsCta As Recordset
sqlaux = ""

If sAgecod <> "" Then
sqlaux = " and substring(p.cctacod,4,2)='" & sAgecod & "'"

End If

sSql = "SELECT  P.CCTACOD,CCTACODANT=ISNULL(R.CCTACODANT,''),NOMBRE=(SELECT TOP 1 E.CPERSNOMBRE FROM PERSONA E JOIN PRODUCTOPERSONA PP ON PP.CPERSCOD=E.CPERSCOD "
sSql = sSql & " WHERE PP.NPRDPERSRELAC=10 AND PP.CCTACOD=P.CCTACOD),"
sSql = sSql & " DIRECCION=ISNULL((SELECT TOP 1 E.CPERSDIRECCDOMICILIO FROM PERSONA E JOIN PRODUCTOPERSONA PP ON PP.CPERSCOD=E.CPERSCOD"
sSql = sSql & " WHERE PP.NPRDPERSRELAC=10 AND PP.CCTACOD=P.CCTACOD),'')"
sSql = sSql & " FROM PRODUCTO P"
sSql = sSql & " LEFT JOIN RELCUENTAS R ON R.CCTACOD=P.CCTACOD"
sSql = sSql & " JOIN CAPTACIONES C ON C.CCTACOD=P.CCTACOD"
sSql = sSql & " JOIN CAPTACAHORROS CA ON CA.CCTACOD=P.CCTACOD"
sSql = sSql & " WHERE  CA.BORDPAG=1  AND P.NPRDESTADO NOT IN (1400,1300) " & sqlaux
 

Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosOP = rsCta
Set rsCta = Nothing

End Function



Public Sub EliminaBeneficiarios(ByVal sPersona As String)
Dim sSql As String
sSql = "DELETE PersRelaciones WHERE cPersCod = '" & sPersona & "' And bPersRelacBenef = 1"
dbCmact.Execute sSql
End Sub

Public Sub AgregaBeneficiario(ByVal sPersona As String, ByVal sPersRelac As String, _
    ByVal nRelacion As PersRelacion, ByVal nPorcentaje As Double, ByVal sMovNro As String)

Dim sSql As String
sSql = "INSERT PersRelaciones (cPersCod,cPersRelacPersCod,nPersRelac,bPersRelacBenef,nPersRelacBenefPorc,bPersRelacAMP, cMovNro) " _
    & "VALUES ('" & sPersona & "','" & sPersRelac & "'," & nRelacion & ",1," & nPorcentaje & ",NULL,'" & sMovNro & "')"
dbCmact.Execute sSql
End Sub

Public Sub AgregaCartaAIP(ByVal cMovNro As String, ByVal cNroRef As String, ByVal cReferencia As String, ByVal NTipoDelito As Integer, ByVal Entidad As String, ByVal Secretario As String, ByVal nTipoDoc As Integer, ByVal cNroDoc As String, ByVal Ubigeo As String, ByRef NroCarta As String)
Dim sSql As String
Dim rsCta As Recordset

On Error GoTo ErrGraba

dbCmact.BeginTrans

    sSql = " insert into Cartasai(NroCarta,cMovnro,cNroRef,cReferencia,NTipoDelito,Entidad,Secretario,nTipoDoc,cNroDoc,Ubigeo) " _
         & " values(dbo.GetNroCartaAI(), '" & cMovNro & "' , '" & cNroRef & "' , '" & cReferencia & "' , " & NTipoDelito & " , '" & Entidad & "','" & Secretario & "' ," & nTipoDoc & " ,'" & cNroDoc & "' ,'" & Ubigeo & "' ) "
    
    dbCmact.Execute sSql
    
    sSql = "Select nrocarta from cartasai where cmovnro='" & cMovNro & "'"
    Set rsCta = New Recordset
    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsCta.ActiveConnection = Nothing
        NroCarta = rsCta!NroCarta
    Set rsCta = Nothing
    
dbCmact.CommitTrans

Exit Sub

ErrGraba:
    dbCmact.RollbackTrans
    Err.Raise Err.Number, "", Err.Description
    

End Sub


Public Sub AgregaCartaAI(ByVal cMovNro As String, ByVal cNroRef As String, ByVal cReferencia As String, ByVal NTipoDelito As Integer, ByVal Entidad As String, ByVal Secretario As String, ByVal nTipoDoc As Integer, ByVal cNroDoc As String, ByVal Ubigeo As String, ByVal NroCarta As String)
Dim sSql As String

    sSql = " insert into Cartasai(NroCarta,cMovnro,cNroRef,cReferencia,NTipoDelito,Entidad,Secretario,nTipoDoc,cNroDoc,Ubigeo) " _
         & " values('" & NroCarta & "', '" & cMovNro & "' , '" & cNroRef & "' , '" & cReferencia & "' , " & NTipoDelito & " , '" & Entidad & "','" & Secretario & "' ," & nTipoDoc & " ,'" & cNroDoc & "' ,'" & Ubigeo & "' ) "
         
dbCmact.Execute sSql

End Sub

Public Sub AgregaBitacoraPPer(ByVal cMovNro As String, ByVal cCtaCod As String, ByVal cPersCod As String, ByVal NPRDPERSRELAC As String, ByVal COBLIGATORIO As String)
Dim sSql As String

    sSql = " insert into BITACORAPRODPERSONA(CMOVNRO,CCTACOD,CPERSCOD,NPRDPERSRELAC,COBLIGATORIO) " _
         & " values('" & cMovNro & "','" & cCtaCod & "','" & cPersCod & "'," & NPRDPERSRELAC & ",'" & COBLIGATORIO & "' ) "
         
dbCmact.Execute sSql

End Sub


Public Sub ActualizaDatosCuenta(ByVal sCuenta As String, ByVal bOrdPag As Boolean, _
        ByVal nFirmas As Integer, nTipoCuenta As ProductoCuentaTipo, _
        ByVal sInstitucion As String, Optional sCuentaAboInt As String = "", Optional ByVal nFirmasMin As Integer, Optional ByVal sAlias As String = "", _
        Optional ByVal nTpoPrograma As Integer = 0)

Dim sSql As String
Dim nProd As Producto

nProd = CLng(Mid(sCuenta, 6, 3))
If nProd = gCapAhorros Then
    sSql = "Update CaptacAhorros Set bOrdPag = " & IIf(bOrdPag, 1, 0) & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
ElseIf nProd = gCapCTS Then
    sSql = "Update CaptacCTS Set cCodInst = '" & sInstitucion & "' " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
ElseIf nProd = gCapPlazoFijo Then
    sSql = "DELETE CaptacCtaAboIntPF Where cCtaCod = '" & sCuenta & "'"
    If sCuentaAboInt <> "" Then
        dbCmact.Execute sSql
        sSql = "INSERT CaptacCtaAboIntPF (cCtaCod,cCtaCodAbono) VALUES ('" & sCuenta & "','" & sCuentaAboInt & "')"
    End If
End If
dbCmact.Execute sSql

If nProd = gCapAhorros Then
    sSql = "Update Captaciones Set nFirmas = " & nFirmas & ", nfirmasmin=" & nFirmasMin & ", calias='" & sAlias & "' , nTpoPrograma=" & nTpoPrograma & "," _
          & " nPrdCtaTpo = " & nTipoCuenta & " WHERE cCtaCod = '" & sCuenta & "'"
          dbCmact.Execute sSql
Else
    sSql = "Update Captaciones Set nFirmas = " & nFirmas & ", nfirmasmin=" & nFirmasMin & ", calias='" & sAlias & "' , " _
          & " nPrdCtaTpo = " & nTipoCuenta & " WHERE cCtaCod = '" & sCuenta & "'"
          dbCmact.Execute sSql
End If

End Sub

Public Function GetOrdenPagoEmitidas(ByVal sCuenta As String) As Recordset
Dim sSql As String
Dim rsCta As Recordset
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
sSql = "SELECT nInicio, nFin, Convert(Varchar(10),Convert(Datetime,LEFT(cMovNro,8)),103) Fecha, " _
    & "RIGHT(cMovNro,4) Usuario FROM CapOrdPagEmision WHERE cCtaCod = '" & sCuenta & "' " _
    & "And nEstado = " & gCapTalOrdPagEstEntregado & " ORDER BY Fecha DESC"
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetOrdenPagoEmitidas = rsCta
Set rsCta = Nothing
End Function

Public Function ExisteOrdenPagoEmitidas(ByVal sCuenta As String, ByVal nInicio As Long, _
        ByVal nFin As Long) As String
Dim sSql As String, sRango As String
Dim rsCta As Recordset
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
sSql = "Select Rango = Convert(Varchar(10),nInicio) + ' - ' + Convert(Varchar(10),nFin) " _
    & "FROM CapOrdPagEmision Where (" & nInicio & " Between nInicio And nFin Or " & nFin & " " _
    & "BETWEEN nInicio And nFin) And cCtaCod = '" & sCuenta & "' And nEstado = " & gCapTalOrdPagEstEntregado
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
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
sSql = "Select Top 1 E.nEstado, K.cConsDescripcion Estado FROM DocRecOPEst E INNER JOIN DocRecOP D ON E.nTpoDoc = D.nTpoDoc " _
    & "AND E.cNroDoc = D.cNroDoc INNER JOIN " & sDBComunes & "Constante K ON E.nEstado = K.nConsValor Where " _
    & "D.cCtaCod = '" & sCuenta & "' And D.nTpoDoc = " & TpoDocOrdenPago & " And E.cNroDoc = '" & nNumOP & "' " _
    & "And K.nConsCod = " & gCaptacOrdPagoEstado & " Order by E.cMovNro DESC"
Set rsOP = New Recordset
rsOP.CursorLocation = adUseClient
rsOP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsOP.ActiveConnection = Nothing
Set GetEstadoOrdenPagoEmitida = rsOP
Set rsOP = Nothing
End Function

Public Function AgregaMovNoContable(ByVal sMovNro As String, ByVal sDesc As String, ByVal nOperacion As CaptacOperacion)
Dim sSql As String
sSql = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sDesc & "'," & gMovEstContabNoContable & "," & gMovFlagVigente & ")"
dbCmact.Execute sSql
End Function

Public Function AgregaOrdenPagoRecibida(ByVal sCuenta As String, ByVal nNumOP As Long, _
        ByVal nMonto As Double, Optional cInstitucion As String = "")
Dim sSql As String
If cInstitucion = "" Then
    sSql = "INSERT DocRecOP (nTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers) " _
        & "VALUES (" & TpoDocOrdenPago & ",'" & nNumOP & "','" & sCuenta & "'," & nMonto & ",NULL)"
Else
    sSql = "INSERT DocRecOP (nTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers) " _
        & "VALUES ('" & TpoDocOrdenPago & "','" & nNumOP & "','" & sCuenta & "'," & nMonto & ",'" & cInstitucion & "')"
End If
dbCmact.Execute sSql
End Function

Public Function AnulaOrdenPagoEmitida(ByVal sCuenta As String, ByVal nNumOP As Long, _
            ByVal sMovNro As String, ByVal nMonto As Double)
Dim sSql As String
sSql = "INSERT DocRecOpEst (nTpoDoc,cNroDoc,cCtaCod,cMovNro,nMonto,nEstado) " _
    & "VALUES (" & TpoDocOrdenPago & ",'" & nNumOP & "','" & sCuenta & "','" & sMovNro & "'," & nMonto & "," & gCapOPEstAnulada & ")"
dbCmact.Execute sSql
End Function

Public Sub AgregaOrdenPagoEmitidas(ByVal sCuenta As String, ByVal nInicio As Long, _
        ByVal nFin As Long, ByVal sMovNro As String)

Dim sSql As String, sDesc As String
Dim nMovNro As Long
sDesc = "Emisión Orden Pago Del " & nInicio & " Al " & nFin & ". Cuenta " & sCuenta
sSql = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & gAhoOPEmision & "','" & sDesc & "'," & gMovEstContabNoContable & "," & gMovFlagVigente & ")"
dbCmact.Execute sSql
nMovNro = GetnMovNro(sMovNro)
sSql = "INSERT MovDocEmitidoRango (nMovNro,cCtaCod,nInicio,nFin) " _
    & "VALUES (" & nMovNro & ",'" & sCuenta & "'," & nInicio & "," & nFin & ")"
dbCmact.Execute sSql
End Sub

Public Sub AgregaTarjetaEstado(ByVal sTarjeta As String, ByVal sMovNro As String, _
        ByVal nEstado As CaptacTarjetaEstado, ByVal sGlosa As String)

Dim sSql As String
sSql = "INSERT TarjetaEstado (cTarjCod,cMovNro,nEstado,cDescripcion) " _
    & "VALUES ('" & sTarjeta & "','" & sMovNro & "'," & nEstado & ",'" & sGlosa & "')"
dbCmact.Execute sSql
End Sub

Public Sub ActualizaTarjetaEstado(ByVal sTarjeta As String, ByVal nEstado As CaptacTarjetaEstado, _
        Optional ByVal sClave As String = "")
Dim sSql As String
If sClave <> "" Then
    sSql = "Update Tarjeta Set nEstado = " & nEstado & ", cClave = '" & sClave & "' " _
        & "WHERE cTarjCod = '" & sTarjeta & "'"
Else
    sSql = "Update Tarjeta Set nEstado = " & nEstado & " " _
        & "WHERE cTarjCod = '" & sTarjeta & "'"
End If
dbCmact.Execute sSql
End Sub

Public Function GetTarjetaEstadoHist(ByVal sTarjeta As String) As Recordset
Dim sSql As String
Dim rsTarj As Recordset

sSql = "SELECT CONVERT(Varchar(10),CONVERT(Datetime,LEFT(E.cMovNro,8)),103) Fecha, K.cConsDescripcion Estado, cDescripcion, " _
    & "RIGHT(E.cMovNro,4) Usu FROM Tarjeta T INNER JOIN TarjetaEstado E ON T.cTarjCod = E.cTarjCod INNER JOIN " _
    & sDBComunes & "Constante K ON E.nEstado = K.nConsValor WHERE T.cTarjCod = '" & sTarjeta & "' AND K.nConsCod = " _
    & gCaptacTarjetaEstado & " ORDER BY Fecha DESC"

Set rsTarj = New Recordset
rsTarj.CursorLocation = adUseClient
rsTarj.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTarj.ActiveConnection = Nothing
Set GetTarjetaEstadoHist = rsTarj
Set rsTarj = Nothing
End Function

Public Function GetNombreTitulares(ByVal sCuenta As String, Optional pbNombreCompleto As Boolean = False, Optional pnTitulares As Integer = 0, Optional pnMargenIzquierdo As Integer = 0) As String
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

Dim lnNumTit As Integer


sSql = "SELECT P.cPersNombre FROM " & sDBPersona & "Persona P INNER JOIN ProductoPersona PP " _
    & "INNER JOIN Producto D ON PP.cCtaCod = D.cCtaCod ON P.cPersCod = PP.cPersCod " _
    & "WHERE D.cCtaCod = '" & sCuenta & "' AND PP.nPrdPersRelac IN (" & gCapRelPersTitular & "," & gGiroRelPersRemitente & ")"
Set rsTit = New ADODB.Recordset
rsTit.CursorLocation = adUseClient
rsTit.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsTit.ActiveConnection = Nothing

If pbNombreCompleto Then
    If Not (rsTit.EOF And rsTit.BOF) Then
        lnNumTit = 0
        sCadena = ""
        While Not rsTit.EOF And lnNumTit < pnTitulares
            If sCadena = "" Then
                sCadena = sCadena & PstaNombre(rsTit("cPersNombre"), False) & Chr(10)
            Else
                sCadena = sCadena & Space(pnMargenIzquierdo) & PstaNombre(rsTit("cPersNombre"), False) & Chr(10)
            End If
            lnNumTit = lnNumTit + 1
            rsTit.MoveNext
        Wend
    End If
    GetNombreTitulares = sCadena
    
    rsTit.Close
    Set rsTit = Nothing
    Exit Function
End If


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
Set rsPer = New Recordset
rsPer.CursorLocation = adUseClient
rsPer.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsPer.ActiveConnection = Nothing
Set GetCTSPeriodo = rsPer
Set rsPer = Nothing
End Function

'Public Function GetSaldoFecha(ByVal sCuenta As String, ByVal dFecha As Date) As Recordset
'Dim sSql As String, sNumMov As String
'Dim rsCta As Recordset
'sSql = "Select TOP 1 CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha, " _
'    & "C.nSaldoDisponible, C.nSaldoContable FROM Mov M INNER JOIN MovCap C " _
'    & "INNER JOIN MovCapDet CD INNER JOIN OpeTpo O INNER JOIN CapMovTipo CT ON " _
'    & "O.cOpeCod = CT.cOpeCod ON CD.cOpeCod = O.cOpeCod ON C.nMovNro = CD.nMovNro ON " _
'    & "M.nMovNro = C.nMovNro WHERE C.cCtaCod = '" & sCuenta & "' AND CD.nConceptoCod " _
'    & "IN (" & gConcCapital & ") AND LEFT(M.cMovNro,8) <= '" & Format$(dFecha, "yyyymmdd") & "' " _
'    & "AND M.nMovFlag IN (" & gMovFlagVigente & ") ORDER BY M.nMovNro DESC"
'Set rsCta = New Recordset
'rsCta.CursorLocation = adUseClient
'rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'Set rsCta.ActiveConnection = Nothing
'Set GetSaldoFecha = rsCta
'Set rsCta = Nothing
'End Function

Public Function GetSaldoFecha(ByVal sCuenta As String, ByVal dFecha As Date) As Recordset

Dim sSql As String, sNumMov As String

Dim rsCta As Recordset

sSql = "Select TOP 1 CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha, " _
    & "C.nSaldoDisponible, C.nSaldoContable FROM Mov M INNER JOIN MovCap C " _
    & "INNER JOIN MovCapDet CD INNER JOIN OpeTpo O INNER JOIN CapMovTipo CT ON " _
    & "O.cOpeCod = CT.cOpeCod ON CD.cOpeCod = O.cOpeCod ON C.nMovNro = CD.nMovNro ON " _
    & "M.nMovNro = C.nMovNro WHERE C.cCtaCod = '" & sCuenta & "' AND CD.nConceptoCod " _
    & "= " & gConcCapital & " AND LEFT(M.cMovNro,8) <= '" & Format$(dFecha, "yyyymmdd") & "' " _
    & "AND M.nMovFlag =" & gMovFlagVigente & " ORDER BY M.nMovNro DESC"

Set rsCta = New Recordset

rsCta.CursorLocation = adUseClient

'dbCmact.ConnectionTimeout = 120

rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsCta.ActiveConnection = Nothing

Set GetSaldoFecha = rsCta

Set rsCta = Nothing

End Function




Public Function GetMovimientosCuenta(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0, _
        Optional bMuestraExtornos As Boolean = True) As Recordset
Dim sSql As String, sNumMov As String, sMuestraExtornos As String
Dim rsCta As Recordset

sNumMov = ""
sMuestraExtornos = ""
If nNumMov > 0 Then sNumMov = " TOP " & nNumMov + 1 & " "
If Not bMuestraExtornos Then
    sMuestraExtornos = " And M.nMovFlag IN (" & gMovFlagVigente & ") "
End If

sSql = "Select " & sNumMov & " CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha, " _
    & "O.cOpeDesc Operacion, ISNULL(MD.cDocNro,'') cDocumento , nAbono = CASE " _
    & "WHEN CT.nCapMovTpo IN (" & gCapMovApertura & "," & gCapMovDeposito & "," & gCapMovIntCap & ") THEN ABS(C.nMonto) " _
    & "ELSE 0 END, " _
    & "nCargo = CASE WHEN CT.nCapMovTpo IN (" & gCapMovRetiro & "," & gCapMovRetiroInac & "," & gCapMovCancelAct & "," & gCapMovCancelInact & "," & gCapMovRetiroInt & ") THEN ABS(C.nMonto)*-1 " _
    & "ELSE 0 END, " _
    & "C.nSaldoContable, A.cAgeDescripcion cAgencia, CAST(left(cmovnro,14) AS BIGINT) cmovnronue " _
    & "FROM MovDoc MD RIGHT JOIN Mov M INNER JOIN MovCapDet CD INNER JOIN MovCap C INNER JOIN " _
    & "OpeTpo O INNER JOIN CapMovTipo CT ON O.cOpeCod = CT.cOpeCod ON " _
    & "C.cOpeCod = O.cOpeCod ON C.nMovNro = CD.nMovNro and c.copecod=cd.copecod And C.cCtaCod = CD.cCtaCod ON M.nMovNro = C.nMovNro " _
    & "ON MD.nMovNro = M.nMovNro INNER JOIN Agencias A ON SUBSTRING(M.cMovNro,18,2) = A.cAgeCod " _
    & "WHERE C.cCtaCod = '" & sCuenta & "'" & sMuestraExtornos

If CLng(Mid(sCuenta, 6, 3)) = gCapPlazoFijo Then
     sSql = sSql & " AND CD.nConceptoCod IN (" & gConcCapital & "," & gConcInteres & ", " & gConcITFCliente & ")"
Else
     sSql = sSql & " AND CD.nConceptoCod IN (" & gConcCapital & ", " & gConcITFCliente & ")"
End If
'
'If nNumMov > 0 Then
'    ssql = ssql & " ORDER BY M.nMovNro asc, C.cOpeCod asc"
'Else
'    ssql = ssql & " AND LEFT(M.cMovNro,8) BETWEEN '" & Format$(dFecIni, "yyyymmdd") & "' " _
'        & "AND '" & Format$(dFecFin, "yyyymmdd") & "' ORDER BY M.nMovNro asc, C.cOpeCod asc"
'End If

If nNumMov > 0 Then
    sSql = sSql & " ORDER BY CAST(left(cmovnro,14) AS BIGINT) DESC,cd.copecod asc"
Else
    sSql = sSql & " AND LEFT(M.cMovNro,8) BETWEEN '" & Format$(dFecIni, "yyyymmdd") & "' " _
        & "AND '" & Format$(dFecFin, "yyyymmdd") & "'"
        
        sSql = sSql & " ORDER BY CAST(left(cmovnro,14) AS BIGINT) desc,cd.copecod asc"
End If

Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetMovimientosCuenta = rsCta
Set rsCta = Nothing
End Function


Public Function GetMovimientosCuenta2(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0, _
        Optional bMuestraExtornos As Boolean = True, Optional cUser As String, Optional ByRef RSTEMP As Recordset) As Recordset

Dim sSql As String, sNumMov As String, sMuestraExtornos As String
Dim rsCta As Recordset

sNumMov = ""
sMuestraExtornos = ""
If nNumMov > 0 Then sNumMov = " TOP " & nNumMov + 1 & " "
If Not bMuestraExtornos Then
    sMuestraExtornos = " And M.nMovFlag IN (" & gMovFlagVigente & ") "
End If

sSql = "Create Table #TMP" & cUser & "(nmovnro int ,dfecha varchar(30),  cOperacion varchar(100),  cDocumento varchar(20),  nAbono money,  nCargo money, nSaldoDisponible money, nSaldoContable money,  cAgencia varchar(50),  cMovNroNue varchar(14),nMovflag int,nMovestado int,nConceptoCod int,Vuser  char(4) ) "
dbCmact.Execute sSql

sSql = " Insert into #TMP" & cUser
sSql = sSql & " Select " & sNumMov & " m.nmovnro,CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha, " _
    & "O.cOpeDesc Operacion, ISNULL(MD.cDocNro,'') cDocumento , nAbono = CASE " _
    & "WHEN CT.nCapMovTpo IN (" & gCapMovApertura & "," & gCapMovDeposito & "," & gCapMovIntCap & ") THEN ABS(C.nMonto) " _
    & "ELSE 0 END, " _
    & "nCargo = CASE WHEN CT.nCapMovTpo IN (" & gCapMovRetiro & "," & gCapMovRetiroInac & "," & gCapMovCancelAct & "," & gCapMovCancelInact & "," & gCapMovRetiroInt & ") THEN ABS(C.nMonto)*-1 " _
    & "ELSE 0 END, " _
    & "C.nSaldoDisponible,C.nSaldoContable, A.cAgeDescripcion cAgencia, CAST(left(cmovnro,14) AS BIGINT) cmovnronue,nmovflag,nmovestado,cd.nconceptocod,right(m.cmovnro,4) as VUSUARIO " _
    & "FROM MovDoc MD RIGHT JOIN Mov M INNER JOIN MovCapDet CD INNER JOIN MovCap C INNER JOIN " _
    & "OpeTpo O INNER JOIN CapMovTipo CT ON O.cOpeCod = CT.cOpeCod ON " _
    & "C.cOpeCod = O.cOpeCod ON C.nMovNro = CD.nMovNro and c.copecod=cd.copecod And C.cCtaCod = CD.cCtaCod ON M.nMovNro = C.nMovNro " _
    & "ON MD.nMovNro = M.nMovNro INNER JOIN Agencias A ON SUBSTRING(M.cMovNro,18,2) = A.cAgeCod " _
    & "WHERE C.cCtaCod = '" & sCuenta & "'" & sMuestraExtornos

If CLng(Mid(sCuenta, 6, 3)) = gCapPlazoFijo Then
     sSql = sSql & " AND CD.nConceptoCod IN (" & gConcCapital & "," & gConcInteres & ", " & gConcITFCliente & ")"
Else
     sSql = sSql & " AND CD.nConceptoCod IN (" & gConcCapital & ", " & gConcITFCliente & ")"
End If
'
'If nNumMov > 0 Then
'    ssql = ssql & " ORDER BY M.nMovNro asc, C.cOpeCod asc"
'Else
'    ssql = ssql & " AND LEFT(M.cMovNro,8) BETWEEN '" & Format$(dFecIni, "yyyymmdd") & "' " _
'        & "AND '" & Format$(dFecFin, "yyyymmdd") & "' ORDER BY M.nMovNro asc, C.cOpeCod asc"
'End If

If nNumMov > 0 Then
    sSql = sSql & " ORDER BY CAST(left(cmovnro,14) AS BIGINT) desc ,cd.copecod asc"
   ' ssql = ssql & " ORDER BY m.nmovnro desc ,cd.copecod asc"
   
Else
    sSql = sSql & " AND LEFT(M.cMovNro,8) BETWEEN '" & Format$(dFecIni, "yyyymmdd") & "' " _
        & "AND '" & Format$(dFecFin, "yyyymmdd") & "'"
        
        sSql = sSql & " ORDER BY CAST(left(cmovnro,14) AS BIGINT) asc ,cd.copecod ASC "
     '   ssql = ssql & " ORDER BY m.nmovnro asc ,cd.copecod asc"
       
End If
  dbCmact.Execute sSql
  

  sSql = "  select dfecha,coperacion,cdocumento,nabono,ncargo,nsaldocontable,cagencia,vUser,nSaldoDisponible  from #TMP" & cUser
  
  If nNumMov > 0 Then
        sSql = sSql & " order by  nmovnro  asc, coperacion asc "
  End If

Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetMovimientosCuenta2 = rsCta

 sSql = sSql & " Select TOP 1 dFecha, " _
        & "nSaldoDisponible, nSaldoContable from #TMP" & cUser _
        & " WHERE nConceptoCod=" & gConcCapital & " AND left(cMovNroNue,8) <= '" & Format$(dFecFin, "yyyymmdd") & "' " _
        & "AND nMovFlag =" & gMovFlagVigente & " ORDER BY nMovNro asc  "
    
    Set RSTEMP = New Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set RSTEMP.ActiveConnection = Nothing
        
  sSql = " drop table #TMP" & cUser
 
  dbCmact.Execute sSql
    
 
 
Set rsCta = Nothing


End Function

Public Function GetMovimientosCuenta4(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0, _
        Optional bMuestraExtornos As Boolean = True, Optional cUser As String, Optional ByRef RSTEMP As Recordset) As Recordset

Dim sSql As String, sNumMov As String, sMuestraExtornos As String
Dim rsCta As Recordset

sNumMov = ""
sMuestraExtornos = ""
If nNumMov > 0 Then sNumMov = " TOP " & nNumMov + 1 & " "
If Not bMuestraExtornos Then
    sMuestraExtornos = " And M.nMovFlag IN (" & gMovFlagVigente & ") "
End If

sSql = "Create Table #TMP" & cUser & "(nmovnro int ,dfecha varchar(30),cOperacion varchar(100),  cDocumento varchar(20),sPersona  varchar(70), nAbono money,  nCargo money, nSaldoDisponible money, nSaldoContable money,  cAgencia varchar(50),  cMovNroNue varchar(14),nMovflag int,nMovestado int,nConceptoCod int,Vuser  char(4) ) "
dbCmact.Execute sSql

sSql = " Insert into #TMP" & cUser
sSql = sSql & " Select " & sNumMov & " m.nmovnro,CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha, " _
    & " O.cOpeDesc Operacion, ISNULL(MD.cDocNro,'') cDocumento, " _
    & " case when mb.creferencia is null or mb.creferencia='' then isnull(ltrim(rtrim(left(m.cmovdesc,60))),'') else ltrim(rtrim(left(mb.creferencia,40))) + ' ' + ltrim(rtrim(right(mb.creferencia,20))) end Persona, " _
    & " nAbono = CASE " _
    & " WHEN CT.nCapMovTpo IN (" & gCapMovApertura & "," & gCapMovDeposito & "," & gCapMovIntCap & ") THEN ABS(C.nMonto) " _
    & " ELSE 0 END, " _
    & " nCargo = CASE WHEN CT.nCapMovTpo IN (" & gCapMovRetiro & "," & gCapMovRetiroInac & "," & gCapMovCancelAct & "," & gCapMovCancelInact & "," & gCapMovRetiroInt & ") THEN ABS(C.nMonto)*-1 " _
    & " ELSE 0 END, " _
    & " C.nSaldoDisponible,C.nSaldoContable, A.cAgeDescripcion cAgencia, CAST(left(cmovnro,14) AS BIGINT) cmovnronue,nmovflag,nmovestado,cd.nconceptocod,right(m.cmovnro,4) as VUSUARIO " _
    & " FROM MovDoc MD RIGHT JOIN Mov M INNER JOIN MovCapDet CD INNER JOIN MovCap C INNER JOIN " _
    & " OpeTpo O INNER JOIN CapMovTipo CT ON O.cOpeCod = CT.cOpeCod ON " _
    & " C.cOpeCod = O.cOpeCod ON C.nMovNro = CD.nMovNro and c.copecod=cd.copecod And C.cCtaCod = CD.cCtaCod ON M.nMovNro = C.nMovNro " _
    & " ON MD.nMovNro = M.nMovNro INNER JOIN Agencias A ON SUBSTRING(M.cMovNro,18,2) = A.cAgeCod " _
    & " left join  movconvcobranza mb on mb.nmovnro=c.nmovnro   " _
    & " WHERE C.cCtaCod = '" & sCuenta & "' and ct.ncapmovtpo=7   " & sMuestraExtornos

If CLng(Mid(sCuenta, 6, 3)) = gCapPlazoFijo Then
     sSql = sSql & " AND CD.nConceptoCod IN (" & gConcCapital & "," & gConcInteres & ", " & gConcITFCliente & ")"
Else
     sSql = sSql & " AND CD.nConceptoCod IN (" & gConcCapital & ", " & gConcITFCliente & ")"
End If
'
'If nNumMov > 0 Then
'    ssql = ssql & " ORDER BY M.nMovNro asc, C.cOpeCod asc"
'Else
'    ssql = ssql & " AND LEFT(M.cMovNro,8) BETWEEN '" & Format$(dFecIni, "yyyymmdd") & "' " _
'        & "AND '" & Format$(dFecFin, "yyyymmdd") & "' ORDER BY M.nMovNro asc, C.cOpeCod asc"
'End If

If nNumMov > 0 Then
    'sSql = sSql & " ORDER BY CAST(left(cmovnro,14) AS BIGINT) asc ,cd.copecod asc"
      sSql = sSql & " ORDER BY m.nmovnro desc ,cd.copecod asc"
Else
    sSql = sSql & " AND LEFT(M.cMovNro,8) BETWEEN '" & Format$(dFecIni, "yyyymmdd") & "' " _
        & "AND '" & Format$(dFecFin, "yyyymmdd") & "'"
        
     '   sSql = sSql & " ORDER BY CAST(left(cmovnro,14) AS BIGINT) asc ,cd.copecod ASC "
      sSql = sSql & " ORDER BY m.nmovnro asc ,cd.copecod asc"
End If
  dbCmact.Execute sSql
  

  sSql = "  select dfecha,coperacion,cdocumento,sPersona,nabono,nsaldocontable,cagencia,vUser,nSaldoDisponible  from #TMP" & cUser


Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetMovimientosCuenta4 = rsCta

 sSql = sSql & " Select TOP 1 dFecha, " _
        & "nSaldoDisponible, nSaldoContable from #TMP" & cUser _
        & " WHERE nConceptoCod=" & gConcCapital & " AND left(cMovNroNue,8) <= '" & Format$(dFecFin, "yyyymmdd") & "' " _
        & "AND nMovFlag =" & gMovFlagVigente & " ORDER BY nMovNro asc  "
    
    Set RSTEMP = New Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set RSTEMP.ActiveConnection = Nothing
        
  sSql = " drop table #TMP" & cUser
 
  dbCmact.Execute sSql
    
 
 
Set rsCta = Nothing


End Function




Public Function GetSaldoFecha2(ByVal cUser As String, ByVal dFecFin As Date) As Recordset
Dim rsCta As Recordset
Dim sSql As String

sSql = sSql & " Select TOP 1 dFecha, " _
        & "nSaldoDisponible, nSaldoContable from #TMP" & cUser _
        & " WHERE nConceptoCod=" & gConcCapital & " AND left(cMovNroNue,8) <= '" & Format$(dFecFin, "yyyymmdd") & "' " _
        & "AND nMovFlag =" & gMovFlagVigente & " ORDER BY nMovNro DESC"
    
    Set rsCta = New Recordset
    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsCta.ActiveConnection = Nothing
    Set GetSaldoFecha2 = rsCta
        
  sSql = " drop table #TMP" & cUser
 
  dbCmact.Execute sSql
    
Set rsCta = Nothing
End Function






Public Sub ActualizaSobregiro(ByVal sCuenta As String, ByVal nSobregiro As Long)
Dim sSql As String
sSql = "Update CaptacAhorros Set nSobreGiro = " & nSobregiro & " Where cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSql
End Sub

Public Sub AgregaCaptacEstado(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, _
        ByVal sMovNro As String)
Dim sSql As String
'Agrega un registro mas de estado a la historia de estados de la cuenta
sSql = "INSERT CaptacEstado (cCtaCod,nPrdEstado,cMovNro) " _
    & "VALUES ('" & sCuenta & "'," & nEstado & ",'" & sMovNro & "')"
dbCmact.Execute sSql
End Sub

Public Function GetTarifaOrdenPago(ByVal nMoneda As Moneda) As Recordset
Dim sSql As String
Dim rsOP As Recordset
sSql = "Select nNumOP, nCosto FROM CapOrdPagTarifa Where nMoneda = " & nMoneda & " " _
    & "Order by nNumOP"
Set rsOP = New Recordset
rsOP.CursorLocation = adUseClient
rsOP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsOP.ActiveConnection = Nothing
Set GetTarifaOrdenPago = rsOP
Set rsOP = Nothing
End Function

Public Function GetMaxOrdPagEmitida(ByVal sCuenta As String) As Long
Dim rsNum As Recordset
Dim sSql As String

Set rsNum = New Recordset
rsNum.CursorLocation = adUseClient
'sSql = "Select ISNULL(MAX(nFin),0) nNum From CapOrdPagEmision Where cCtaCod LIKE '" & Mid(sCuenta, 1, 9) & "%' " _
'    & "And nEstado NOT IN (" & gCapTalOrdPagEstExtornado & ")"

sSql = " Select ISNULL(MAX(nFin),0) nNum From CapOrdPagEmision Where cCtaCod ='" & sCuenta & "' " _
    & " And nEstado NOT IN (" & gCapTalOrdPagEstExtornado & ") and bantiguo is null "


rsNum.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsNum.ActiveConnection = Nothing
If Not (rsNum.EOF And rsNum.BOF) Then
    GetMaxOrdPagEmitida = rsNum("nNum")
Else
    GetMaxOrdPagEmitida = 0
End If
rsNum.Close
Set rsNum = Nothing
End Function

Public Function GetHistOrdPagEmision(ByVal sCuenta As String) As Recordset
Dim rsHist As Recordset
Dim sSql As String
sSql = "Select nInicio, nFin, dFecha = CONVERT(Varchar(12),CONVERT(Datetime,SUBSTRING(cMovNro,1,8)),103), " _
    & "cEstado = K.cConsDescripcion From CapOrdPagEmision O INNER JOIN Constante K ON O.nEstado = K.nConsValor " _
    & "Where O.cCtaCod = '" & sCuenta & "' And K.nConsCod = " & gCapOrdPagTalEstado & " Order by cMovNro"
Set rsHist = New Recordset
rsHist.CursorLocation = adUseClient
rsHist.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsHist.ActiveConnection = Nothing
Set GetHistOrdPagEmision = rsHist
Set rsHist = Nothing
End Function
            
Public Sub AgregaCapOrdPagEmision(ByVal sCuenta As String, ByVal sMovNro As String, _
        ByVal nEstado As CapOrdPagTalEstado, ByVal nInicio As Long, ByVal nFin As Long, _
        ByVal nNumTal As Long)
Dim sSql As String
sSql = "Insert CapOrdPagEmision (cCtaCod,cMovNro,nEstado,nInicio,nFin,nNumTal) " _
    & "Values ('" & sCuenta & "','" & sMovNro & "'," & Trim(nEstado) & "," & Trim(nInicio) & "," & Trim(nFin) & "," & Trim(nNumTal) & ")"
dbCmact.Execute sSql
End Sub
            
Public Sub AgregaCapOrdPagEstado(ByVal sCuenta As String, ByVal sMovNro As String, _
        ByVal nEstado As CapOrdPagTalEstado, ByVal nInicio As Long)
Dim sSql As String
sSql = "Insert CapOrdPagEstado (cCtaCod,cMovNro,nEstado,nInicio) " _
    & "Values ('" & sCuenta & "','" & sMovNro & "'," & Trim(nEstado) & "," & Trim(nInicio) & ")"
dbCmact.Execute sSql
End Sub
            
Public Sub ActualizaCapOrdPagEmision(ByVal sCuenta As String, ByVal sMovNroNuevo As String, _
        ByVal nEstadoNuevo As CapOrdPagTalEstado, ByVal sMovNroAnt As String, _
        ByVal nEstadoAnt As CapOrdPagTalEstado)

Dim sSql As String
sSql = "Update CapOrdPagEmision Set nEstado = " & nEstadoNuevo & ", cMovNro = '" & sMovNroNuevo & "' " _
    & "Where cCtaCod = '" & sCuenta & "' And nEstado = " & nEstadoAnt & " And cMovNro = '" & sMovNroAnt & "'"
dbCmact.Execute sSql
End Sub
    
Public Function GetCapOrdPagEmision(ByVal nEstado As CapOrdPagTalEstado) As Recordset
Dim rsOrden As Recordset
Dim sSql As String

sSql = "Select OP.cCtaCod, OP.nInicio, OP.nFin, dFecha = CONVERT(Varchar(12),CONVERT(Datetime,SUBSTRING(OP.cMovNro,1,8)),103), " _
    & "P.cPersNombre, T.nTipo, OP.nNumTal, OP.cMovNro From CapOrdPagEmision OP INNER JOIN ProductoPersona PC INNER JOIN Persona P " _
    & "ON PC.cPersCod = P.cPersCod ON OP.cCtaCod = PC.cCtaCod INNER JOIN CapOrdPagTarifa T ON " _
    & "CONVERT(INT,SUBSTRING(OP.cCtaCod,9,1)) = T.nMoneda And (OP.nFin - OP.nInicio + 1)/OP.nNumTal = T.nNumOP " _
    & "Where OP.nEstado = " & nEstado & " And PC.nPrdPersRelac = " & gCapRelPersTitular & " And P.cPersCod IN " _
    & "(Select MAX(PC1.cPersCod) From ProductoPersona PC1 Where PC1.cCtaCod = OP.cCtaCod And PC1.nPrdPersRelac = " & gCapRelPersTitular & ") Order by dFecha"
Set rsOrden = New Recordset
rsOrden.CursorLocation = adUseClient
rsOrden.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsOrden.ActiveConnection = Nothing
Set GetCapOrdPagEmision = rsOrden
Set rsOrden = Nothing
End Function
        
Public Function GetCuentaAbonoIF(ByVal sPersCod As String, ByVal nMoneda As Moneda) As String
Dim rsCta As Recordset
Dim sSql As String
sSql = "Select C.cCtaCod FROM ProductoPersona PP INNER JOIN CuentaIFEspecial C ON " _
    & " PP.cCtaCod = C.cCtaCod Where PP.cPersCod = '" & sPersCod & "' And C.bAbono = 1 " _
    & " And C.cCtaCod LIKE '________" & Trim(nMoneda) & "%'"

Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
If Not (rsCta.EOF And rsCta.BOF) Then
    GetCuentaAbonoIF = rsCta("cCtaCod")
Else
    GetCuentaAbonoIF = ""
End If
Set rsCta = Nothing
End Function

Public Function BuscaCreditosPendientesPago(ByVal sCuenta As String) As Boolean
Dim rsCta As Recordset
Dim sSql As String
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
sSql = "Select bBusqCredPend From CaptacPlazoFijo Where cCtaCod = '" & sCuenta & "'"
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
If rsCta.EOF And rsCta.BOF Then
    BuscaCreditosPendientesPago = True
Else
    BuscaCreditosPendientesPago = rsCta("bBusqCredPend")
End If
rsCta.Close
Set rsCta = Nothing
End Function

Public Function GetProductoPersonaConsol(ByVal sCuenta As String) As Recordset
Dim sSql As String
Dim rsCta As Recordset

sSql = "Select PP.cCtaCod FROM " & sDBPersona & "Persona P INNER JOIN ProductoPersona PP ON " _
    & "P.cPersCod = PP.cPersCod WHERE PP.cCtaCod = '" & sCuenta & "' Order by PP.cCtaCod"
    
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetProductoPersonaConsol = rsCta
Set rsCta = Nothing
End Function

Public Function GetChequesOperaciones(Optional sCuenta As String = "", Optional sNumCheque As String = "", Optional ByVal nEstado As Integer = -1) As Recordset
Dim sSql As String, sqlaux As String
Dim rsCta As Recordset
sqlaux = ""
    

If sCuenta <> "" Then
   If nEstado <> -1 Then
        sqlaux = " and k.nconsvalor=" & nEstado & " "
      
    sSql = "Select D.cNroDoc, D.cIFCta, (P.cPersNombre + Space(75) + D.cIFTpo + D.cPersCod) Banco, " _
        & "K.cConsDescripcion + Space(75) + Convert(Varchar(1),D.nEstado) Estado, " _
        & "dFecReg = CONVERT(Varchar(12),CONVERT(Datetime,SUBSTRING(M.cMovNro,1,8)),103), " _
        & "dValor = CASE WHEN D.nEstado = " & gChqEstValorizado & " THEN " _
        & "Convert(Varchar(12),D.dValorizacion,103) ELSE Convert(Varchar(12),D.dValorizaRef,103) End, " _
        & "sSimbolo = CASE WHEN D.nMoneda = " & gMonedaNacional & " THEN 'S/.' ELSE 'US$' END, " _
        & " C.nMonto From Persona P INNER JOIN DocRec D INNER JOIN DocRecCapta C INNER JOIN MovCap MC INNER JOIN " _
        & "Mov M ON MC.nMovNro = M.nMovNro ON C.nMovNro = MC.nMovNro ON " _
        & "D.nTpoDoc = C.nTpoDoc And D.cNroDoc = C.cNroDoc And D.cPersCod = C.cPersCod And " _
        & "D.cIFTpo = C.cIFTpo ON P.cPersCod = D.cPersCod INNER JOIN Constante K ON " _
        & "D.nEstado = K.nConsValor and mc.copecod not like '99%'  Where    C.cCtaCod = '" & sCuenta & "' And K.nConsCod = " & gChequeEstado & "  " & sqlaux _
        & "Order by dFecReg"
        
   Else
    sSql = " Select D.cNroDoc, D.cIFCta, (P.cPersNombre + Space(75) + D.cIFTpo + D.cPersCod) Banco," _
           & " K.cConsDescripcion + Space(75) + Convert(Varchar(1),D.nEstado) Estado, " _
           & " dFecReg = CONVERT(Varchar(12),CONVERT(Datetime,SUBSTRING(X.cMovNro,1,8)),103), " _
           & " dValor = CASE WHEN D.nEstado = 2 THEN Convert(Varchar(12),D.dValorizacion,103) ELSE Convert(Varchar(12), " _
           & " D.dValorizaRef,103) End, sSimbolo = CASE WHEN D.nMoneda = 1 THEN 'S/.' ELSE 'US$' END, " _
           & "  C.nMonto " _
           & "  From Persona P " _
           & " INNER JOIN DocRec D " _
           & " INNER JOIN DocRecCapta C " _
           & " INNER join ( select DISTINCT m.cmovnro, mc.cctacod,MC.NMOVNRO from MOV M JOIN MOVCAP MC ON MC.NMOVNRO=M.NMOVNRO WHERE mc.copecod not like '99%' ) X " _
           & " ON X.NMOVNRO=C.NMOVNRO " _
           & " on D.nTpoDoc = C.nTpoDoc And D.cNroDoc = C.cNroDoc And D.cPersCod = C.cPersCod And D.cIFTpo = C.cIFTpo ON P.cPersCod = D.cPersCod " _
           & " INNER JOIN Constante K ON D.nEstado = K.nConsValor " _
           & " Where    C.cCtaCod = '" & sCuenta & "' And K.nConsCod = " & gChequeEstado & "  and k.nconsvalor=1 Order by dFecReg "
  
        
   End If

ElseIf sNumCheque <> "" Then
    If nEstado <> -1 Then
        sqlaux = " and E.nEstado =" & nEstado & " "
        
        sSql = "Select D.cNroDoc, D.cIFCta, (P.cPersNombre + Space(75) + D.cIFTpo + D.cPersCod) Banco, " _
        & "K.cConsDescripcion + Space(75) + Convert(Varchar(1),D.nEstado) Estado, " _
        & "dFecReg = CONVERT(Varchar(12),CONVERT(Datetime,SUBSTRING(E.cMovNro,1,8)),103), " _
        & "dValor = CASE WHEN D.nEstado = " & gChqEstValorizado & " THEN " _
        & "Convert(Varchar(12),D.dValorizacion,103) ELSE Convert(Varchar(12),D.dValorizaRef,103) End, " _
        & "sSimbolo = CASE WHEN D.nMoneda = " & gMonedaNacional & " THEN 'S/.' ELSE 'US$' END, " _
        & " D.nMonto From Persona P INNER JOIN DocRec D INNER JOIN DocRecEst E ON D.nTpoDoc = E.nTpoDoc " _
        & "And D.cNroDoc = E.cNroDoc And D.cPersCod = E.cPersCod And D.cIFTpo = E.cIFTpo " _
        & "ON P.cPersCod = D.cPersCod INNER JOIN Constante K ON D.nEstado = K.nConsValor " _
        & "Where D.cNroDoc = '" & sNumCheque & "' And K.nConsCod = " & gChequeEstado & " " & sqlaux _
        & "And E.cMovNro IN (Select MAX(E1.cMovNro) From DocRecEst E1 Where E1.nTpoDoc = E.nTpoDoc " _
        & "And E1.cNroDoc = E.cNroDoc And E1.cPersCod = E.cPersCod And E1.cIFTpo = E.cIFTpo) " _
        & "Order by dFecReg"
        
    Else
        sSql = " Select D.cNroDoc, D.cIFCta, (P.cPersNombre + Space(75) + D.cIFTpo + D.cPersCod) Banco, " _
               & " K.cConsDescripcion + Space(75) + Convert(Varchar(1),D.nEstado) Estado, " _
               & " dFecReg = CONVERT(Varchar(12),CONVERT(Datetime,SUBSTRING(E.cMovNro,1,8)),103), " _
               & " dValor = CASE WHEN D.nEstado = 2 THEN Convert(Varchar(12),D.dValorizacion,103) ELSE Convert(Varchar(12),D.dValorizaRef,103) End, " _
               & " sSimbolo = CASE WHEN D.nMoneda = 1 THEN 'S/.' ELSE 'US$' END, " _
               & "    D.nMonto " _
               & "   From Persona P " _
               & "  INNER JOIN DocRec D " _
               & "  INNER JOIN  DocRecEst E ON D.nTpoDoc = E.nTpoDoc And D.cNroDoc = E.cNroDoc And D.cPersCod = E.cPersCod And D.cIFTpo = E.cIFTpo ON P.cPersCod = D.cPersCod " _
               & "    INNER JOIN Constante K ON D.nEstado = K.nConsValor " _
               & " Where D.cNroDoc = '" & sNumCheque & "' And K.nConsCod =  " & gChequeEstado & "  and e.nestado=1 " _
               & "   Order by dFecReg "
        
    End If
        
        '& " E.nEstado IN (" & gChqEstEnValorizacion & "," & gChqEstRegistrado & ") "
End If
   
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetChequesOperaciones = rsCta
Set rsCta = Nothing
End Function


Public Function GetChequesRegistrados(ByVal sNumCheque As String) As Recordset
Dim sSql As String
Dim rsCta As Recordset
sSql = "Select D.cNroDoc, D.cIFCta, (P.cPersNombre + Space(75) + D.cIFTpo + D.cPersCod) Banco, " _
    & "K.cConsDescripcion + Space(75) + Convert(Varchar(1),D.nEstado) Estado, " _
    & "dFecReg = CONVERT(Varchar(12),CONVERT(Datetime,SUBSTRING(E.cMovNro,1,8)),103), " _
    & "dValor = Convert(Varchar(12),D.dValorizaRef,103), " _
    & "sSimbolo = CASE WHEN D.nMoneda = " & gMonedaNacional & " THEN 'S/.' ELSE 'US$' END, " _
    & "D.nMonto, M.nMovNro From Persona P INNER JOIN DocRec D INNER JOIN DocRecEst E JOIN Mov M ON " _
    & "E.cMovNro = M.cMovNro ON D.nTpoDoc = E.nTpoDoc " _
    & "And D.cNroDoc = E.cNroDoc And D.cPersCod = E.cPersCod And D.cIFTpo = E.cIFTpo " _
    & "ON P.cPersCod = D.cPersCod INNER JOIN Constante K ON D.nEstado = K.nConsValor " _
    & "Where D.cNroDoc = '" & sNumCheque & "' And K.nConsCod = " & gChequeEstado & " And E.nEstado = " & gChqEstRegistrado & " " _
    & "And D.nEstado = " & gChqEstRegistrado & " Order by dFecReg"
   
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetChequesRegistrados = rsCta
Set rsCta = Nothing
End Function

Public Function GetChequesValorizadosInmediato(ByVal dFecha As Date, Optional sCuenta As String = "", _
            Optional sNumCheque As String = "") As Recordset

Dim sSql As String
Dim rsCta As Recordset

If sNumCheque <> "" Then
    sSql = "Select D.cNroDoc, D.cIFCta, (P.cPersNombre + Space(75) + D.cIFTpo + D.cPersCod) Banco, " _
        & "K.cConsDescripcion + Space(75) + Convert(Varchar(1),D.nEstado) Estado, " _
        & "dFecReg = CONVERT(Varchar(12),CONVERT(Datetime,SUBSTRING(E.cMovNro,1,8)),103), " _
        & "dValor = Convert(Varchar(12),D.dValorizacion,103), " _
        & "sSimbolo = CASE WHEN D.nMoneda = " & gMonedaNacional & " THEN 'S/.' ELSE 'US$' END, " _
        & "D.nMonto, M.nMovNro From Persona P INNER JOIN DocRec D INNER JOIN DocRecEst E JOIN Mov M ON " _
        & "E.cMovNro = M.cMovNro ON D.nTpoDoc = E.nTpoDoc " _
        & "And D.cNroDoc = E.cNroDoc And D.cPersCod = E.cPersCod And D.cIFTpo = E.cIFTpo " _
        & "ON P.cPersCod = D.cPersCod INNER JOIN Constante K ON D.nEstado = K.nConsValor " _
        & "Where D.cNroDoc = '" & sNumCheque & "' And K.nConsCod = " & gChequeEstado & " And " _
        & "E.nEstado = " & gChqEstValorizado & " And M.cOpeCod = '" & gChqOpeValorInmediata & "' " _
        & "And M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' " _
        & "And D.nEstado = " & gChqEstValorizado & " Order by dFecReg"
        
ElseIf sCuenta <> "" Then
    sSql = "Select D.cNroDoc, D.cIFCta, (P.cPersNombre + Space(75) + D.cIFTpo + D.cPersCod) Banco, " _
        & "K.cConsDescripcion + Space(75) + Convert(Varchar(1),D.nEstado) Estado, " _
        & "dFecReg = CONVERT(Varchar(12),CONVERT(Datetime,SUBSTRING(M.cMovNro,1,8)),103), " _
        & "dValor = Convert(Varchar(12),D.dValorizacion,103), " _
        & "sSimbolo = CASE WHEN D.nMoneda = " & gMonedaNacional & " THEN 'S/.' ELSE 'US$' END, " _
        & "MC.nMonto, M.nMovNro From Persona P INNER JOIN DocRec D INNER JOIN DocRecEst E JOIN Mov M JOIN MovCap MC " _
        & "ON M.nMovNro = MC.nMovNro ON E.cMovNro = M.cMovNro ON D.nTpoDoc = E.nTpoDoc " _
        & "And D.cNroDoc = E.cNroDoc And D.cPersCod = E.cPersCod And D.cIFTpo = E.cIFTpo " _
        & "ON P.cPersCod = D.cPersCod INNER JOIN Constante K ON " _
        & "D.nEstado = K.nConsValor Where MC.cCtaCod = '" & sCuenta & "' And K.nConsCod = " & gChequeEstado & " And " _
        & "M.cOpeCod = '" & gChqOpeValorInmediata & "' And E.nEstado = " & gChqEstValorizado & " And " _
        & "M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "%' " _
        & "And D.nEstado = " & gChqEstValorizado & " Order by dFecReg"

End If
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetChequesValorizadosInmediato = rsCta
Set rsCta = Nothing
End Function

Public Function GetChequesEnValorizacion(ByVal sCuenta As String) As Recordset

Dim sSql As String
Dim rsCta As Recordset

    sSql = " Select Count(*) Num,  IsNull(Sum(nMonto),0) Monto from DocRecCapta DRC" _
         & " Inner Join DocRecEst DRE On DRC.nTpoDoc = DRE.nTpoDoc And DRC.cNroDoc = DRE.cNroDoc And DRC.cPersCod = DRE.cPersCod" _
         & " Where DRC.cCtaCod = '" & sCuenta & "' And DRE.cMovNro = (Select Max(cMovNro) Mov From DocRecEst DREAdd Where DREAdd.nTpoDoc = DRE.nTpoDoc And DREAdd.cNroDoc = DRE.cNroDoc And DREAdd.cPersCod = DRE.cPersCod)" _
         & " And DRE.nEstado = " & gChqEstEnValorizacion

Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetChequesEnValorizacion = rsCta
Set rsCta = Nothing
End Function

Public Function GetCapEstadMovTipo(ByVal sCondicion As String, ByVal sFecha As String) As Recordset
Dim sSql As String
Dim rsMov As Recordset
sSql = ""

sSql = "Select nNumAper = ISNULL(COUNT(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 1 and not c.copecod like '99%'  THEN C.nMovNro END),0), " _
    & "nMonAper = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 1 and not c.copecod like '99%'  THEN Abs(CD.nMonto) END),0), " _
    & "nNumCancAct = ISNULL(COUNT(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 2 and not c.copecod like '99%'  THEN C.nMovNro END),0), " _
    & "nMonCancAct = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 2 and not c.copecod like '99%'  THEN Abs(CD.nMonto) END)*-1,0), " _
    & "nNumCancInact = ISNULL(COUNT(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 3 and not c.copecod like '99%'  THEN C.nMovNro END),0), " _
    & "nMonCancInact = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 3 and not c.copecod like '99%'  THEN Abs(CD.nMonto) END)*-1,0), " _
    & "nMonRetInact = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 4 and not c.copecod like '99%'  THEN Abs(CD.nMonto) END)*-1,0), " _
    & "nMonRetInt = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 2 And EO.nCapMovTpo = 5 and not c.copecod like '99%' THEN Abs(CD.nMonto) END)*-1,0), " _
    & "nNumRet = ISNULL(COUNT(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 6 and not c.copecod like '99%'  THEN C.nMovNro END),0), " _
    & "nMonRet = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 6 and not c.copecod like '99%'  THEN Abs(CD.nMonto) END)*-1,0), " _
    & "nNumDep = ISNULL(COUNT(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 7 and not c.copecod like '99%'  THEN C.nMovNro END),0), " _
    & "nMonDep = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 7 and not c.copecod like '99%' THEN Abs(CD.nMonto) END),0), " _
    & "nMonIntCap = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 8 and not c.copecod like '99%'  THEN Abs(CD.nMonto) END),0), " _
    & "nNumITF= ISNULL(COUNT(CASE WHEN CD.nConceptoCod = 20 and c.copecod in ('990101','990301') And EO.nCapMovTpo = 6 THEN C.nMovNro END),0),  " _
    & "nITF= ISNULL(SUM(CASE WHEN CD.nConceptoCod = 20  and c.copecod in ('990101','990301') And EO.nCapMovTpo = 6 THEN Abs(CD.nMonto) END)*-1,0)  " _
    & "From  (Select nMovNro, cOpeCod From Mov Where cMovNro LIKE '" & sFecha & "' And " _
    & "         nMovFlag = 0) M   " _
    & "     INNER JOIN MovCap C ON M.nMovNro = C.nMovNro " _
    & "     INNER JOIN MovCapDet CD ON CD.nMovNro = C.nMovNro And Cd.cCtaCod = C.cCtaCod and CD.cOpeCod = C.cOpeCod " _
    & "     INNER JOIN CapMovTipo  EO on EO.cOpeCod = C.cOpeCod   " _
    & "Where C.cCtaCod LIKE '" & sCondicion & "' And EO.bEstadistica = 1"

Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetCapEstadMovTipo = rsMov
Set rsMov = Nothing
End Function

Public Function GetCapEstadSaldoProd(ByVal sCondicion As String)
Dim sSql As String
Dim rsMov As Recordset
'Calculamos El Saldo por Producto
sSql = "Select nSaldo = ISNULL(SUM(nSaldo),0) From Producto Where nPrdEstado NOT IN (1300,1400) " _
    & "And cCtaCod LIKE '" & sCondicion & "'"

Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetCapEstadSaldoProd = rsMov
Set rsMov = Nothing
End Function

Public Function GetCapEstadMovCheques(ByVal sFecha As String, Optional sCondicion As String)
Dim sSql As String
Dim rsMov As Recordset
'Calculamos el monto en cheques movidos en la fecha indicada
sSql = "Select nMonChq = ISNULL(SUM(D.nMonto),0) From DocRec D INNER JOIN DocRecCapta DC INNER JOIN MovCap C INNER JOIN " _
    & "Mov M ON C.nMovNro = M.nMovNro ON DC.nMovNro = C.nMovNro And DC.cCtaCod = C.cCtaCod ON " _
    & "D.nTpoDoc = DC.nTpoDoc And D.cNroDoc = DC.cNroDoc And D.cPersCod = DC.cPersCod And " _
    & "D.cIFTpo = D.cIFTpo Where D.nEstado IN (1,2) And M.cMovNro LIKE '" & sFecha & "'"

'Agregado
If sCondicion <> "" Then
    sSql = sSql & "And C.cCtaCod LIKE '" & sCondicion & "'"
End If
'Fin Agregado

Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov.ActiveConnection = Nothing
Set GetCapEstadMovCheques = rsMov
Set rsMov = Nothing
End Function

Public Function GetITFReporteContribuyentes(ByVal dInicio As Date, ByVal dFin As Date, _
        Optional ByVal bNacional As Boolean = True) As ADODB.Recordset
Dim sSql As String
Dim sNacional As String
Dim rsItf As ADODB.Recordset
Dim sMes As String

sMes = Format$(dFin, "yyyymm")
If bNacional Then
    sNacional = " ID.cPersIDTpo <> '" & gPersIdExtranjeria & "' "
Else
    sNacional = " ID.cPersIDTpo = '" & gPersIdExtranjeria & "' "
End If

'ssql = "Select Distinct P.cPersCod, P.nPersPersoneria, P.cPersIDTpo, P.cPersIDnro From " _
'    & "(Select PP.cPersCod From " _
'    & "(Select MC.cCtaCod From Mov M JOIN MovCap MC ON M.nMovNro = MC.nMovNro " _
'    & "Where M.cMovNro LIKE '" & sMes & "%' And M.nMovFlag = " & gMovFlagVigente & " And MC.cOpeCod LIKE '990[13]%' " _
'    & "UNION Select MC.cCtaCod From Mov M JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
'    & "Where M.cMovNro LIKE '" & sMes & "' And M.nMovFlag = " & gMovFlagVigente & "  And MC.cOpeCod LIKE '990[13]%' " _
'    & "UNION Select cCtaCod = V.cNumRecibo From (Select nMovNro, cMovNro, cOpeCod From Mov " _
'    & "Where cMovNro LIKE '" & sMes & "%' And nMovFlag = " & gMovFlagVigente & ") M JOIN MovServicios V ON " _
'    & "M.nMovNro = V.nMovNro Where V.cOpeCod NOT LIKE '990[13]%') M JOIN ProductoPersona PP ON " _
'    & "M.cCtaCod = PP.cCtaCod Where PP.nPrdPersRelac IN (10,20) Group by PP.cPersCod " _
'    & "UNION Select cPersCod = RTRIM(V.cReferencia) From (Select nMovNro, cMovNro, cOpeCod " _
'    & "From Mov Where cMovNro LIKE '" & sMes & "%' And nMovFlag = " & gMovFlagVigente & ") M JOIN MovOpeVarias V " _
'    & "ON M.nMovNro = V.nMovNro Where M.cOpeCod NOT LIKE '990[13]%' Group by RTRIM(V.cReferencia)) PP " _
'    & "JOIN (Select P.cPersCod, P.cPersNombre, P.nPersPersoneria, cPersIDTpo = ISNULL(ID.cPersIDTpo,''), " _
'    & "cPersIDnro = ISNULL(ID.cPersIDnro,'') From Persona P LEFT JOIN PersID ID ON P.cPersCod = ID.cPersCod " _
'    & "Where " & sNacional & ")  P ON PP.cPersCod = P.cPersCod where cPersIDnro <> ''  "




sSql = " Select  P.cPersCod, P.nPersPersoneria, P.cPersIDTpo, P.cPersIDnro"
sSql = sSql & " From"
sSql = sSql & " (Select distinct P.cPersCod, P.cPersNombre, P.nPersPersoneria, cPersIDTpo = ISNULL(ID.cPersIDTpo,''),"
sSql = sSql & "  cPersIDnro = ISNULL(ID.cPersIDnro,'') From Persona P LEFT JOIN PersID ID ON P.cPersCod = ID.cPersCod"
sSql = sSql & "  Where " & sNacional & " )P "
sSql = sSql & " Join "
sSql = sSql & " (Select distinct PP.cPersCod From "
sSql = sSql & " (Select MC.cCtaCod From Mov M JOIN MovCap MC ON M.nMovNro = MC.nMovNro "
sSql = sSql & "  Where M.cMovNro LIKE '" & sMes & "%' And MC.cOpeCod LIKE '990[13]%' And M.nMovFlag = " & gMovFlagVigente & " "
sSql = sSql & "  UNION Select MC.cCtaCod From Mov M JOIN MovCol MC ON M.nMovNro = MC.nMovNro"
sSql = sSql & "  Where M.cMovNro LIKE '" & sMes & "%' And MC.cOpeCod LIKE '990[13]%'  And M.nMovFlag = " & gMovFlagVigente & " "
sSql = sSql & "  UNION Select cCtaCod = V.cNumRecibo From (Select nMovNro, cMovNro, cOpeCod From Mov"
sSql = sSql & "  Where cMovNro LIKE '" & sMes & "%' And nMovFlag = " & gMovFlagVigente & " ) M"
sSql = sSql & "  JOIN MovServicios V ON"
sSql = sSql & "  M.nMovNro = V.nMovNro Where V.cOpeCod NOT LIKE '990[13]%') M"
sSql = sSql & "  JOIN ProductoPersona PP ON  M.cCtaCod = PP.cCtaCod"
sSql = sSql & " Where PP.nPrdPersRelac IN (10,20)"
sSql = sSql & "  Group by PP.cPersCod"
sSql = sSql & "  Union"
sSql = sSql & "  Select cPersCod = RTRIM(V.cReferencia)"
sSql = sSql & " From ( Select nMovNro, cMovNro, cOpeCod"
sSql = sSql & " From Mov Where cMovNro LIKE '" & sMes & "%' And nMovFlag =" & gMovFlagVigente & " ) M JOIN MovOpeVarias V"
sSql = sSql & " ON M.nMovNro = V.nMovNro"
sSql = sSql & " Where M.cOpeCod NOT LIKE '990[13]%'"
sSql = sSql & " Group by RTRIM(V.cReferencia)) PP ON PP.cPersCod = P.cPersCod"
sSql = sSql & " where cPersIDnro <> '' "
sSql = sSql & " order by p.cpersidtpo "







Set rsItf = New Recordset
rsItf.CursorLocation = adUseClient
rsItf.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsItf.ActiveConnection = Nothing
Set GetITFReporteContribuyentes = rsItf
Set rsItf = Nothing
End Function

Public Function GetITFReporteSustentacionMov(ByVal dInicio As Date, ByVal dFin As Date) As ADODB.Recordset
Dim sSql As String
Dim rsItf As ADODB.Recordset
Dim sInicio As String, sFin As String

sInicio = Format(dInicio, "yyyymmdd")
sFin = Format(dFin, "yyyymmdd")

sSql = " SELECT  cfecha, CCODAGE=Substring(cta.cuenta,4,2) , "
sSql = sSql & "    cta.COPECOD, COPEDESC, nPrdConceptoCod, cDescripcion, "
sSql = sSql & "  sum(nMontDet) nMontDet "
sSql = sSql & " From "
sSql = sSql & " (   SELECT  MC.CCTACOD AS CUENTA, LEFT(M.CMOVNRO,8) as cfecha, MC.COPECOD, SUBSTRING(MC.CCTACOD,9,1) AS NMONEDA, "
sSql = sSql & "      O.COPEDESC, MD.nPrdConceptoCod, PC.cDescripcion, sum(case when SUBSTRING(MC.CCTACOD,9,1)=2 then round(MD.nMonto*dbo.GetTCPonderado( LEFT(M.CMOVNRO,8)),2) else round(MD.nMonto,2) end) as nMontDet, "
sSql = sSql & "  (SELECT max(cperscod) FROM ProductoPersona pp where pp.cctacod = MC.cctacod And nPrdPersRelac = 20) CodPers "
sSql = sSql & "    FROM    MOV M"
sSql = sSql & "        JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO"
sSql = sSql & "        JOIN MOVCOLDET MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.COPECOD = MC.COPECOD AND MD.CCTACOD = MC.CCTACOD"
sSql = sSql & "        JOIN OPETPO O ON O.COPECOD = MC.COPECOD"
sSql = sSql & "        JOIN PRODUCTOCONCEPTO PC ON PC.nPrdConceptoCod = MD.nPrdConceptoCod"
sSql = sSql & "    WHERE   MC.COPECOD LIKE '99%' AND LEFT(M.CMOVNRO,8) BETWEEN  '" & sInicio & "' AND '" & sFin & "' AND M.NMOVFLAG = 0"
sSql = sSql & "    GROUP BY MC.CCTACOD, LEFT(M.CMOVNRO,8), MC.COPECOD, SUBSTRING(MC.CCTACOD,9,1),O.COPEDESC, MD.nPrdConceptoCod, PC.cDescripcion"
sSql = sSql & "    Union"
sSql = sSql & "    SELECT  MC.CCTACOD AS CUENTA,LEFT(M.CMOVNRO,8) as cfecha, MC.COPECOD, SUBSTRING(MC.CCTACOD,9,1) AS NMONEDA,"
sSql = sSql & "        O.COPEDESC, MD.nConceptoCod, PC.cDescripcion, sum(case when SUBSTRING(MC.CCTACOD,9,1)=2 then round(MD.nMonto*dbo.GetTCPonderado( LEFT(M.CMOVNRO,8)),2) else round(MD.nMonto,2) end) as nMontDet,"
sSql = sSql & "        (SELECT max(cperscod) FROM ProductoPersona pp where pp.cctacod = MC.cctacod And nPrdPersRelac = 10) CodPers"
sSql = sSql & "     FROM    MOV M"
sSql = sSql & "        JOIN MOVCAP MC ON MC.NMOVNRO = M.NMOVNRO"
sSql = sSql & "        JOIN MOVCAPDET MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.COPECOD = MC.COPECOD AND MD.CCTACOD = MC.CCTACOD"
sSql = sSql & "        JOIN OPETPO O ON O.COPECOD = MC.COPECOD"
sSql = sSql & "        JOIN PRODUCTOCONCEPTO PC ON PC.nPrdConceptoCod = MD.nConceptoCod"
sSql = sSql & "    WHERE   MC.COPECOD LIKE '99%' AND LEFT(M.CMOVNRO,8) BETWEEN '" & sInicio & "' AND '" & sFin & "' AND M.NMOVFLAG = 0"
sSql = sSql & "        AND NOT M.COPECOD IN (100103, 100104)"
sSql = sSql & "    GROUP BY MC.CCTACOD, LEFT(M.CMOVNRO,8), MC.COPECOD, SUBSTRING(MC.CCTACOD,9,1),O.COPEDESC, MD.nConceptoCod, PC.cDescripcion"
sSql = sSql & "    Union"
sSql = sSql & "    SELECT  '108'+right(MC.cCtaContCod,2) AS cuenta , LEFT(M.CMOVNRO,8) AS CFECHA, M.COPECOD, SUBSTRING(MC.cCtaContCod,3,1) AS NMONEDA,"
sSql = sSql & "        O.COPEDESC, 99,'ITF CUENTA PROPIA',"
sSql = sSql & "        SUM(CASE WHEN SUBSTRING(MC.cCtaContCod,3,1) ='2' THEN round(ABS(ME.NMOVMEIMPORTE)*dbo.GetTCPonderado( LEFT(M.CMOVNRO,8)),2)  ELSE round(ABS(MC.nMovImporte),2) END) as nMontDet ,"
sSql = sSql & "        '1089800000272' as  CodPers"
sSql = sSql & "    FROM    MOV M"
sSql = sSql & "        JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO"
sSql = sSql & "        LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO and ME.nMovItem = MC.nMovItem"
sSql = sSql & "        JOIN OPETPO O ON O.COPECOD = M.COPECOD"
sSql = sSql & "    WHERE   LEFT(M.CMOVNRO,8) BETWEEN '" & sInicio & "' AND '" & sFin & "' AND M.NMOVFLAG = 0"
sSql = sSql & "        AND MC.cCtaContCod LIKE '21_4010901%' and MC.nMovImporte<0"
sSql = sSql & "    GROUP BY MC.cCtaContCod,LEFT(M.CMOVNRO,8), M.COPECOD, SUBSTRING(MC.cCtaContCod,3,1),O.COPEDESC"
sSql = sSql & " ) AS CTA"
sSql = sSql & "     left join persona pe on pe.cperscod =cta.codpers"
sSql = sSql & "     left join persid bb on bb.cperscod=cta.codpers    and bb.cPersIDTpo = 1"
sSql = sSql & "     left join persid cc on cc.cperscod=cta.codpers    and cc.cPersIDTpo = 2"
sSql = sSql & "     left join persid dd on dd.cperscod=cta.codpers    and dd.cPersIDTpo = 4"
sSql = sSql & "     left join persid ee on ee.cperscod=cta.codpers    and ee.cPersIDTpo = 11 "
sSql = sSql & "     left join persid ff on ff.cperscod=cta.codpers    and ff.cPersIDTpo = 10 "
'ssql = ssql & " where  isnull(bb.cPersIDTpo, isnull(cc.cPersIDTpo,isnull(  dd.cPersIDTpo, ff.cPersIDTpo))) in (1,2,4,10)"
'ssql = ssql & " where  isnull(bb.cPersIDTpo, isnull( cc.cPersIDTpo,isnull(dd.cPersIDTpo, isnull(ee.cPersIDTpo,ff.cPersIDTpo))))  in (1,2,4,10,11) "
sSql = sSql & " group by cta.COPECOD, cfecha,Substring(cta.cuenta,4,2), COPEDESC, nPrdConceptoCod, cDescripcion"
sSql = sSql & " order by Substring(cta.cuenta,4,2), nPrdConceptocod, cFecha, cta.copecod"

Set rsItf = New Recordset
rsItf.CursorLocation = adUseClient
rsItf.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsItf.ActiveConnection = Nothing
Set GetITFReporteSustentacionMov = rsItf
Set rsItf = Nothing

End Function

Public Function GetITFReporteMovimientosAcumulados(ByVal dInicio As Date, ByVal dFin As Date) As ADODB.Recordset
Dim sSql As String
Dim rsItf As ADODB.Recordset
Dim sMes As String

sMes = Format$(dFin, "yyyymm")

sSql = "Select Distinct P.cPersCod, P.nPersPersoneria, P.cPersIDTpo, P.cPersIDnro From " _
    & "(Select PP.cPersCod From " _
    & "(Select MC.cCtaCod From Mov M JOIN MovCap MC ON M.nMovNro = MC.nMovNro " _
    & "Where M.cMovNro LIKE '" & sMes & "%' And M.nMovFlag = " & gMovFlagVigente & " And MC.cOpeCod LIKE '990[13]%' " _
    & "UNION Select MC.cCtaCod From Mov M JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
    & "Where M.cMovNro LIKE '" & sMes & "' And M.nMovFlag = " & gMovFlagVigente & "  And MC.cOpeCod LIKE '990[13]%' " _
    & "UNION Select cCtaCod = V.cNumRecibo From (Select nMovNro, cMovNro, cOpeCod From Mov " _
    & "Where cMovNro LIKE '" & sMes & "%' And nMovFlag = " & gMovFlagVigente & ") M JOIN MovServicios V ON " _
    & "M.nMovNro = V.nMovNro Where V.cOpeCod NOT LIKE '990[13]%') M JOIN ProductoPersona PP ON " _
    & "M.cCtaCod = PP.cCtaCod Where PP.nPrdPersRelac IN (10,20) Group by PP.cPersCod " _
    & "UNION Select cPersCod = RTRIM(V.cReferencia) From (Select nMovNro, cMovNro, cOpeCod " _
    & "From Mov Where cMovNro LIKE '" & sMes & "%' And nMovFlag = " & gMovFlagVigente & ") M JOIN MovOpeVarias V " _
    & "ON M.nMovNro = V.nMovNro Where M.cOpeCod NOT LIKE '990[13]%' Group by RTRIM(V.cReferencia)) PP " _
    & "JOIN (Select P.cPersCod, P.cPersNombre, P.nPersPersoneria, cPersIDTpo = ISNULL(ID.cPersIDTpo,''), " _
    & "cPersIDnro = ISNULL(ID.cPersIDnro,'') From Persona P LEFT JOIN PersID ID ON P.cPersCod = ID.cPersCod) P " _
    & "ON PP.cPersCod = P.cPersCod "

Set rsItf = New Recordset
rsItf.CursorLocation = adUseClient
rsItf.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsItf.ActiveConnection = Nothing
Set GetITFReporteMovimientosAcumulados = rsItf
Set rsItf = Nothing
End Function

Public Sub AgregaITFControlContribuyentes(ByVal dFecha As Date, ByVal sMes As String, ByVal sPersCod As String, _
            ByVal sIDTipo As String, ByVal sIDNro As String)
Dim sSql As String
sSql = "INSERT ITFControlContribuyentes (dFecha,cMes,cPersCod,cIDTipo,cIDNro) " _
    & "VALUES ('" & Format$(dFecha, "mm/dd/yyyy hh:mm:ss") & "','" & sMes & "','" & sPersCod & "','" & sIDTipo & "','" & sIDNro & "')"
dbCmact.Execute sSql
End Sub

Public Function GetITFExonerados(ByVal dInicio As Date, ByVal dFin As Date) As ADODB.Recordset
Dim sSql As String
Dim sInicio As String, sFin As String
Dim rsItf As ADODB.Recordset

sInicio = Format$(dInicio, "yyyymmdd")
sFin = Format$(dFin, "yyyymmdd")

sSql = "Select Distinct PP.cCtaCod, P.cPersCod, P.nPersPersoneria, P.cPersIDTpo, P.cPersIDnro, TipoDecla = '1', E.cRegistro, "
sSql = sSql & " E.nExoTpo, CodOpe = CASE WHEN (E.nExoTpo = 2 or E.nExoTpo = 1 or E.nExoTpo = 14 ) then 'b' "
sSql = sSql & " WHEN E.nExoTpo = 15 THEN 'o' WHEN E.nExoTpo = 3 then 'c' "
sSql = sSql & " WHEN E.nExoTpo = 4 THEN 'd' WHEN substring(PP.cCtaCod, 6, 3) = '423' then 'w' WHEN E.nExoTpo = 6 THEN 'm'  End  , TipoOper = '1'       From "
sSql = sSql & " ITFExoneracionCta E JOIN ProductoPersona PP JOIN (Select P.cPersCod, P.cPersNombre, P.nPersPersoneria, cPersIDTpo = ISNULL(ID.cPersIDTpo,''), "
sSql = sSql & " cPersIDnro = ISNULL(ID.cPersIDnro,'') From Persona P LEFT JOIN PersID ID ON P.cPersCod = ID.cPersCod  where  (cPersIDnro<>'' AND cPersIDnro IS NOT NULL )  ) P "
sSql = sSql & " ON PP.cPersCod = P.cPersCod ON E.cCtaCod = PP.cCtaCod Where PP.nPrdPersRelac IN (" & gCapRelPersTitular & "," & gColRelPersTitular & ") "
sSql = sSql & " And E.nExoTpo not in (0,21,24) AND E.CREGISTRO <='" & sFin & "'    "

Set rsItf = New Recordset
rsItf.CursorLocation = adUseClient
rsItf.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsItf.ActiveConnection = Nothing
Set GetITFExonerados = rsItf
Set rsItf = Nothing
End Function

Public Sub AgregaITFControlExonerados(ByVal dFecha As Date, ByVal sMes As String, ByVal sPersCod As String, _
            ByVal sIDTipo As String, ByVal sIDNro As String, ByVal sCtaCod As String, ByVal sTipoDecl As String, _
            ByVal sCodExon As String, ByVal sTipoOpe As String, ByVal sMovNro As String)
Dim sSql As String
sSql = "INSERT ITFControlExonerados (dFecha,cMes,cPersCod,cIDTipo,cIDNro,cCtaCod,cTipoDecl,cCodExon,cTipoOpe,cMovNro) " _
    & "VALUES ('" & Format$(dFecha, "mm/dd/yyyy hh:mm:ss") & "','" & sMes & "','" & sPersCod & "','" & sIDTipo & "','" _
    & sIDNro & "','" & sCtaCod & "','" & sTipoDecl & "','" & sCodExon & "','" & sTipoOpe & "','" & sMovNro & "')"
dbCmact.Execute sSql
End Sub


Private Sub Class_Initialize()
Dim sConn As String
Dim ClsIni As COMConecta.DCOMClasIni
Set ClsIni = New COMConecta.DCOMClasIni
sConn = ClsIni.CadenaConexion
sDBComunes = ClsIni.BaseComunes
sDBPersona = ClsIni.BasePersonas
sDBImagenes = ClsIni.BaseImagenes
Set ClsIni = Nothing
Set dbCmact = New Connection
dbCmact.Open sConn
dbCmact.Execute "SET DATEFORMAT MDY"
End Sub

Private Sub Class_Terminate()
dbCmact.Close
Set dbCmact = Nothing
End Sub



