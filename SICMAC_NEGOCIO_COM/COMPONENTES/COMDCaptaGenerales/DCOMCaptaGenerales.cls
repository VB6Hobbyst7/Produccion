VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMCaptaGenerales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public dbCmact As ADODB.Connection
Dim oConecta As New COMConecta.DCOMConecta 'GIPO 13-03-2017 ERS070-2016
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String
Public Function GetnMovNro(ByVal sMovNro As String) As Long
Dim rsMov As ADODB.Recordset
Dim sSql As String
sSql = "Select nMovNro From Mov Where cMovNro = '" & sMovNro & "'"
Set rsMov = New ADODB.Recordset
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

Public Function GetNroOPeradoras(ByVal sCodAge As String) As Long
    Dim rsMov As ADODB.Recordset
    Dim sSql As String

    sSql = " SELECT NRP=IsNULL(NRP,0) FROM AGENCIAS WHERE CAGECOD='" & sCodAge & "'"

    Set rsMov = New ADODB.Recordset
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

Public Function GetTpoActI() As ADODB.Recordset
    Dim rstemp As ADODB.Recordset
    Dim sSql As String

    sSql = " Select nconsvalor,cconsdescripcion from constante where nconscod='2027' and nconsvalor<>'2027' order by cconsdescripcion "
    Set rstemp = New ADODB.Recordset
    rstemp.CursorLocation = adUseClient
    rstemp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText




    rstemp.ActiveConnection = Nothing
    Set GetTpoActI = rstemp
    Set rstemp = Nothing

End Function
Public Function GetTpoDocActI() As ADODB.Recordset
Dim rstemp As ADODB.Recordset
Dim sSql As String


    sSql = " Select nconsvalor,cconsdescripcion from constante where nconscod='2028' and nconsvalor<>'2028' order by cconsdescripcion "
    Set rstemp = New ADODB.Recordset
    rstemp.CursorLocation = adUseClient
    rstemp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText




    rstemp.ActiveConnection = Nothing
    Set GetTpoDocActI = rstemp
    Set rstemp = Nothing

End Function

Public Function GetDPTO() As ADODB.Recordset
Dim rstemp As ADODB.Recordset
Dim sSql As String


    sSql = " select cubigeocod,cubigeodes from ubigeo where cubigeocod like '1%' order by cubigeodes "
    Set rstemp = New ADODB.Recordset
    rstemp.CursorLocation = adUseClient
    rstemp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText




    rstemp.ActiveConnection = Nothing
    Set GetDPTO = rstemp
    Set rstemp = Nothing

End Function

Public Function GetPROV(Optional sDepto As String) As ADODB.Recordset
Dim rstemp As ADODB.Recordset
Dim sSql As String


    sSql = " select cubigeocod,cubigeodes from ubigeo where cubigeocod like '2" & sDepto & "%' order by cubigeodes"
    Set rstemp = New ADODB.Recordset
    rstemp.CursorLocation = adUseClient
    rstemp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText




    rstemp.ActiveConnection = Nothing
    Set GetPROV = rstemp
    Set rstemp = Nothing

End Function



Public Function GetDatosOpeLLamada(ByVal nMovNro As Long, ByVal sOperacion As String) As ADODB.Recordset
Dim rstemp As ADODB.Recordset
Dim sSql As String


    sSql = "Select mc.cctacod,mc.nmonto,mr.nmovnroref from movref mr join movcap mc on mc.nmovnro=mr.nmovnroref where mr.nmovnro=" & nMovNro & " and copecod='" & sOperacion & "'"
    Set rstemp = New ADODB.Recordset
    rstemp.CursorLocation = adUseClient
    rstemp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText




    rstemp.ActiveConnection = Nothing
    Set GetDatosOpeLLamada = rstemp
    Set rstemp = Nothing
        
End Function


Public Function GetCuentasPersonaAI(Optional ByVal sPersCod As String) As ADODB.Recordset
 Dim rstemp As ADODB.Recordset
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

     Set rstemp = New ADODB.Recordset
     rstemp.CursorLocation = adUseClient
     rstemp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText




     rstemp.ActiveConnection = Nothing
     Set GetCuentasPersonaAI = rstemp
     Set rstemp = Nothing
 
  
End Function



Public Function GetDatosCuentaAho(ByVal sCuenta As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

'**Modificado por DAOR 20081110, nSaldoDisp=nSaldoDisp-nRetencion (por cajeros) ********
'sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
'    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, AG.cAgeDescripcion, " _
'    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.bOrdPag, A.dUltContacto, T.cConsDescripcion cEstado, " _
'    & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.bInactiva, A.bInmovilizada, A.nSobregiro, " _
'    & "C.nFirmasMin, C.cAlias, nTpoPrograma=isnull(c.nTpoPrograma,0), ISNULL(C.nBloqueoParcial,0) nBloqueoParcial,nPlazo,nMontoAbono,dRenoPanderito FROM Producto P " _
'    & "INNER JOIN Captaciones C INNER JOIN CaptacAhorros A ON C.cCtaCod = A.cCtaCod ON " _
'    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
'    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
'    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod " _
'    & "WHERE P.cCtaCod = '" & sCuenta & "' AND T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
'    & "AND T2.nConsCod = " & gCaptacTipoTasa

' *** COMENTADO POR RIRO RIRO 20131102 SEGUN "CAMBIOS EN PODERES"

    'sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
    '& "(C.nSaldoDisp-Isnull(A.nRetencion,0)) as nSaldoDisp,A.nRetencion, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, AG.cAgeDescripcion, " _
    '& "C.nPrdCtaTpo, C.nPrdTasaInteres, A.bOrdPag, A.dUltContacto, T.cConsDescripcion cEstado, " _
    '& "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.bInactiva, A.bInmovilizada, A.nSobregiro, " _
    '& "C.nFirmasMin, C.cAlias, nTpoPrograma=isnull(c.nTpoPrograma,0), ISNULL(C.nBloqueoParcial,0) nBloqueoParcial,nPlazo,nMontoAbono,dRenoPanderito FROM Producto P " _
    '& "INNER JOIN Captaciones C INNER JOIN CaptacAhorros A ON C.cCtaCod = A.cCtaCod ON " _
    '& "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    '& "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
    '& "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod " _
    '& "WHERE P.cCtaCod = '" & sCuenta & "' AND T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    '& "AND T2.nConsCod = " & gCaptacTipoTasa

' *** Modificado a Store por AMDO20140207 *********************************
' *** AGREGADO POR RIRO RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
'    sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
'    & "(C.nSaldoDisp-Isnull(A.nRetencion,0)) as nSaldoDisp,A.nRetencion, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, AG.cAgeDescripcion, " _
'    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.bOrdPag, A.dUltContacto, T.cConsDescripcion cEstado, " _
'    & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.bInactiva, A.bInmovilizada, A.nSobregiro, " _
'    & "C.nFirmasMin, C.cAlias, nTpoPrograma=isnull(c.nTpoPrograma,0), ISNULL(C.nBloqueoParcial,0) nBloqueoParcial,nPlazo,nMontoAbono,dRenoPanderito,isnull(cReglas,'') cReglas FROM Producto P " _
'    & "INNER JOIN Captaciones C INNER JOIN CaptacAhorros A ON C.cCtaCod = A.cCtaCod ON " _
'    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
'    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
'    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod " _
'    & "WHERE P.cCtaCod = '" & sCuenta & "' AND T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
'    & "AND T2.nConsCod = " & gCaptacTipoTasa
    
' *** END RIRO
    sSql = "exec stp_sel_DevolverDatosCuentaAhorro '" & sCuenta & "'" 'AMDO 20140107
' ***END Modificado a Store por AMDO20140207 *********************************

Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set GetDatosCuentaAho = rsCta

Set rsCta.ActiveConnection = Nothing
Set rsCta = Nothing
End Function
'madm 20111121 - 20110429 - madm 20100512 -----------------------------------------------------------------------------------------------------
Public Function GetDatosCuentaPFCodPer(ByVal cPersCod As String, Optional bFiltro As Integer = 0) As ADODB.Recordset
    Dim sSql As String
    Dim rsCta As ADODB.Recordset
    
'    sSql = "Select P.cCtaCod cPersNombre,P.cCtaCod cPersCod  from producto P " _
'           & " inner join ProductoPersona PP on P.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac=10 " _
'           & " inner join Constante C on C.nConsValor = P.nPrdEstado And C.nConsCod = 2001" _
'           & " And  substring(P.cCtaCod,6,3) = '233' And nPrdEstado = 1000 " _
'           & " And P.CctaCod not in (select cNroDoc from garantias G where G.cNroDoc = P.cCtaCod ) " _
'           & "where cPersCod = '" & cPersCod & "'"
    If bFiltro = 1 Then
        sSql = " exec stp_sel_ValidaPlazoFijoGarPerCF '" & cPersCod & "'"
    Else
        sSql = " exec stp_sel_ValidaPlazoFijoGarPer '" & cPersCod & "'"
    End If
    
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    Set rsCta.ActiveConnection = Nothing
    Set GetDatosCuentaPFCodPer = rsCta
    Set rsCta = Nothing
    
End Function
'end madm -----------------------------------------------------------------------------------------------------------

Public Function GetDatosCuentaPF(ByVal sCuenta As String) As ADODB.Recordset
    Dim sSql As String
    Dim rsCta As ADODB.Recordset
    'By Capi 07082008 se adiciono columna bTasaPactada
    'By BRGO 20111220 Se adicionó la columna "nTpoPrograma" y "nMontoAbono"
    '***Modificado por ELRO el 20121030, según OYP-RFC098-2012
    'sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
    '    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, AG.cAgeDescripcion, " _
    '    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.nPlazo, A.nIntPag, A.dRenovacion, A.nApertura, A.nIntAdelRev, " _
    '    & "A.nFormaRetiro, A.dAuxiliar, A.nDuplicado, T.cConsDescripcion cEstado, A.bBusCredPend, " _
    '    & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, T3.cConsDescripcion cRetiro, " _
    '    & "A.dUltCierreAnt, A.dAuxiliarAnt, ISNULL(CI.cCtaCodAbono,'') cCtaCodAbono, a.nformaretiro,c.calias, " _
    '    & "c.nfirmasmin,nIntDevChq,C.cUltimaActualizacion,A.bTasaPactada,nTpoPrograma=IsNull(C.nTpoPrograma,0), nMontoAbono = IsNull(A.nMontoAbono,0.00) FROM Producto P " _
    '    & "INNER JOIN Captaciones C INNER JOIN CaptacPlazoFijo A LEFT JOIN CaptacCtaAboIntPF CI " _
    '    & "ON A.cCtaCod = CI.cCtaCod ON C.cCtaCod = A.cCtaCod ON " _
    '    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    '    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
    '    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Constante T3 ON A.nFormaRetiro = T3.nConsValor " _
    '    & "INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod WHERE P.cCtaCod = '" & sCuenta & "' AND " _
    '    & "T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    '    & "AND T2.nConsCod = " & gCaptacTipoTasa & " AND T3.nConsCod = " & gCaptacPFFormaRetiro
    sSql = "exec stp_sel_RFC0982012_ObtenerDatosCuentaPF '" & sCuenta & "'" 'MIOL 20121113, Se acualizo el Store Procedure
    '***Fin Modificado por ELRO el 20121030*******************
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    Set rsCta.ActiveConnection = Nothing
    Set GetDatosCuentaPF = rsCta
    Set rsCta = Nothing
End Function

Public Function GetCuentasNoAsociadas(sPersCod) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

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
    
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText




Set rsCta.ActiveConnection = Nothing
Set GetCuentasNoAsociadas = rsCta
Set rsCta = Nothing
End Function

Public Function GetPersonasCuentaNA(ByVal sCuenta As String, ByVal sPersCod As String, Optional ByVal nPersoneria As Integer) As ADODB.Recordset
Dim sSql As String, sqlaux As String
Dim rsCta As ADODB.Recordset

If nPersoneria = 1 Then
    sqlaux = " and pP.nprdpersrelac=" & CaptacRelacPersona.gCapRelPersTitular
Else
    sqlaux = " and pP.nprdpersrelac in (" & CaptacRelacPersona.gCapRelPersTitular & "," & CaptacRelacPersona.gCapRelPersRepTitular & ")"
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
    
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText




Set GetPersonasCuentaNA = rsCta
Set rsCta.ActiveConnection = Nothing
Set rsCta = Nothing
End Function


Public Function EsCtaConvenio(ByVal sCuenta As String) As Boolean
  Dim sSql As String, rsCta As ADODB.Recordset
  
EsCtaConvenio = False
  
  sSql = " select cperscod,cctacod from captacconveniocta " _
         & " where ntpocuenta=1 and cctacod='" & sCuenta & "'"
    
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText




 If Not rsCta.EOF Then
   If Trim(rsCta!cCtaCod) = Trim(sCuenta) Then EsCtaConvenio = True
 End If

Set rsCta.ActiveConnection = Nothing
Set rsCta = Nothing
End Function


Public Function GetDatosCuentaCTS(ByVal sCuenta As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

'**Modificado por DAOR 20081110, nSaldoDisp=nSaldoDisp-nRetencion (por cajeros) ********
'sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado,RetAdiCTS=Isnull(A.nRetAdiCTS,0), " _
'    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas,c.nfirmasmin ,C.nIntAcum, dUltCierre=( case when dbo.ultimodiames(dUltCierre)=1 then dUltCierre else dateadd(dd,1, dUltCierre) end )  , C.dApertura, AG.cAgeDescripcion, " _
'    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.nSaldRetiro, A.nIntSaldo, PE.cPersNombre cInstitucion, " _
'    & "T.cConsDescripcion cEstado, T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.cCodInst,C.cAlias,ISNULL(C.nBloqueoParcial,0) nBloqueoParcial, IsNull(PEJUR.cPersJurSigla,'') EmpSiglas " _
'    & "FROM Producto P INNER JOIN Captaciones C INNER JOIN CaptacCTS A INNER JOIN " & sDBPersona & "" _
'    & "Persona PE ON A.cCodInst = PE.cPersCod ON C.cCtaCod = A.cCtaCod ON " _
'    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
'    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
'    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod " _
'    & "Left JOIN " & sDBPersona & "PersonaJur PEJUR ON PE.cPersCod = PEJUR.cPersCod " _
'    & "WHERE P.cCtaCod = '" & sCuenta & "' AND T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
'    & "AND T2.nConsCod = " & gCaptacTipoTasa

'Mody by Gitu 08-09-2010 para corregir los dias transcurridos para el calculo de interes de CTS
'dUltCierre=( case when dbo.ultimodiames(dUltCierre)=1 then dUltCierre else dateadd(dd,1, dUltCierre) end )
'*** Modif por BRGO 20111220 - Se agregó la columna "nTpoPrograma"
'***Modificado por ELRO el 20121026, según OYP-RFC101-2012
'sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado,RetAdiCTS=Isnull(A.nRetAdiCTS,0), " _
'    & "(C.nSaldoDisp-Isnull(A.nRetencion,0)) as nSaldoDisp,A.nRetencion, C.nPersoneria, C.nFirmas,c.nfirmasmin ,C.nIntAcum, dUltCierre , C.dApertura, AG.cAgeDescripcion, " _
'    & "C.nPrdCtaTpo, C.nPrdTasaInteres, (A.nSaldRetiro-Isnull(A.nRetencion,0)) as nSaldRetiro, A.nIntSaldo, PE.cPersNombre cInstitucion, " _
'    & "T.cConsDescripcion cEstado, T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.cCodInst,C.cAlias,ISNULL(C.nBloqueoParcial,0) nBloqueoParcial, IsNull(PEJUR.cPersJurSigla,'') EmpSiglas,ISNULL(C.nTpoPrograma,0) nTpoPrograma " _
'    & "FROM Producto P INNER JOIN Captaciones C INNER JOIN CaptacCTS A INNER JOIN " & sDBPersona & "" _
'    & "Persona PE ON A.cCodInst = PE.cPersCod ON C.cCtaCod = A.cCtaCod ON " _
'    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
'    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
'    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod " _
'    & "Left JOIN " & sDBPersona & "PersonaJur PEJUR ON PE.cPersCod = PEJUR.cPersCod " _
'    & "WHERE P.cCtaCod = '" & sCuenta & "' AND T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
'    & "AND T2.nConsCod = " & gCaptacTipoTasa
sSql = "exec stp_sel_RFC1012012_DevolverDatosCuentaCTS '" & sCuenta & "'"
'***Fin Modificado por ELRO el 20121026*******************
'******************************************************************************************
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText




Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaCTS = rsCta
Set rsCta = Nothing
End Function

' *** RIRO20131102 SE AGREGO PARAMETRO, sFilter
Public Function GetProductoPersona(ByVal sCuenta As String, Optional sFilter As String = "") As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

' *** COMENTADO POR RIRO20131102 SEGUN "CAMBIOS EN PODERES"

    'sSql = " Select PP.cPersCod, P.cPersNombre Nombre, (UPPER(RTRIM(T.cConsDescripcion)) " _
    '& " + SPACE(75 - len((UPPER(RTRIM(T.cConsDescripcion))))) + CONVERT(Varchar(2),PP.nPrdPersRelac)) cRelacion, P.nPersPersoneria, Obligatorio=case when pp.cobligatorio='S' then 'SI' when (pp.cobligatorio='N' or pp.cobligatorio is null) then 'NO' else 'OPCIONAL' end   " _
    '& " FROM " & sDBPersona & "Persona P INNER JOIN ProductoPersona PP ON " _
    '& " P.cPersCod = PP.cPersCod INNER JOIN " & sDBComunes & "Constante T " _
    '& " ON PP.nPrdPersRelac = T.nConsValor WHERE PP.cCtaCod = '" & sCuenta & "' " _
    '& " AND T.nConsCod = " & gCaptacRelacPersona & "" _
    '& " Order by PP.nPrdPersRelac "
    
' *** RIRO20131102 SEGUN "CAMBIOS EN PODERES"
sSql = " Select PP.cPersCod, P.cPersNombre Nombre, (UPPER(RTRIM(T.cConsDescripcion)) " _
    & " + SPACE(75 - len((UPPER(RTRIM(T.cConsDescripcion))))) + CONVERT(Varchar(2),PP.nPrdPersRelac)) cRelacion, P.nPersPersoneria, Obligatorio=case when pp.cobligatorio='S' then 'SI' when (pp.cobligatorio='N' or pp.cobligatorio is null) then 'NO' else 'OPCIONAL' end  , cGrupo as Grupo , pp.nPrdPersRelac" _
    & " FROM " & sDBPersona & "Persona P INNER JOIN ProductoPersona PP ON " _
    & " P.cPersCod = PP.cPersCod INNER JOIN " & sDBComunes & "Constante T " _
    & " ON PP.nPrdPersRelac = T.nConsValor WHERE PP.cCtaCod = '" & sCuenta & "' " _
    & " AND T.nConsCod = " & gCaptacRelacPersona & ""
       
    If Len(Trim(sFilter)) <> 0 Then
         sSql = sSql & "and pp.nPrdPersRelac not in (" & sFilter & ")"
    End If
    
    sSql = sSql & " Order by PP.nPrdPersRelac "
    
' *** END RIRO
    
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText




Set rsCta.ActiveConnection = Nothing
Set GetProductoPersona = rsCta
Set rsCta = Nothing
End Function

Public Function GetCuentasPersona(ByVal sPers As String, Optional nProd As Producto, _
        Optional bActivas As Boolean = False, Optional bBloqueadas As Boolean = False, _
        Optional nMoneda As Moneda, Optional bOrdPag As Boolean = False, _
        Optional sAgencia As String = "", Optional sPrograma As String = "", _
        Optional bCtasDesemb = False, Optional nTiposCta As String = "", _
        Optional psCadPrd As String = "", Optional ByVal bBloqueadaTotal As Boolean) As ADODB.Recordset  '** Juez 20120723 Se agregó bCtasDesemb
        'BRGO 20111118 Se agregó el parámetro "nPrograma"
        'JUEZ 20140520 Se cambio tipo de dato a nTiposCta de ProductoCuentaTipo a string
Dim sSql As String
Dim rsCta As ADODB.Recordset

'By Capi 11012008 se modifico la consulta para que verifique si la cuenta esta exonerada de ITF
'***Condicion ITF.nExoTpo > 0 agregado por ELRO el 20130726, SATI INC1306180008
'JUEZ 20141114 Se agregó nTpoPrograma
'RIRO20150527 ERS162-2015, SEe agregó "Constante t4"
If nProd = gCapAhorros Then
    sSql = "Select PP.cCtaCod, (UPPER(T3.cConsDescripcion) + ' - ' + T1.cConsDescripcion) cDescripcion, nNivel = 1, " _
        & "T1.cConsDescripcion cRelacion, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, c.npersoneria,nTpoPrograma=ISNULL(nTpoPrograma,0), " _
        & "UPPER(T2.cConsDescripcion) cProducto, UPPER(T3.cConsDescripcion) cMoneda,P.nTasaInteres,C.dApertura,C.nfirmas,Exonerada=Isnull(ITF.cCtaCod,'')  " _
        & ",t4.cConsDescripcion SubProducto " _
        & " FROM ProductoPersona PP INNER JOIN Producto P INNER " _
        & "JOIN Captaciones C INNER JOIN CaptacAhorros A ON C.cCtaCod = A.cCtaCod ON P.cCtaCod = C.cCtaCod ON PP.cCtaCod = P.cCtaCod INNER JOIN " _
        & sDBComunes & "Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN " & sDBComunes & "" _
        & "Constante T1 ON PP.nPrdPersRelac = T1.nConsValor INNER JOIN " & sDBComunes & "Constante T2 " _
        & "ON SUBSTRING(PP.cCtaCod,6,3) = CONVERT(Varchar(3),T2.nConsValor) INNER JOIN " & sDBComunes & "" _
        & "Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor) LEFT JOIN ITFExoneracionCta ITF ON " _
        & "PP.cCtaCod=itf.cCtaCod and  ITF.nExoTpo > 0 " _
        & "left join Constante t4 on t4.nConsValor = c.nTpoPrograma and t4.nConsCod = case SUBSTRING(c.cCtaCod,6,3) when '232' then 2030 when '233' then 2032 when '234' then 2033 end" _
        & " WHERE PP.cPersCod = '" & sPers & "' " _
        & "AND T1.nConsCod = " & gCaptacRelacPersona & " AND T.nConsCod = " & gCaptacEstado & " AND " _
        & "T2.nConsCod = " & gProducto & " AND T3.nConsCod = " & gMoneda
Else
    sSql = "Select PP.cCtaCod, (UPPER(T3.cConsDescripcion) + ' - ' + T1.cConsDescripcion) cDescripcion, nNivel = 1, " _
        & "T1.cConsDescripcion cRelacion, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, c.npersoneria, " _
        & "UPPER(T2.cConsDescripcion) cProducto, UPPER(T3.cConsDescripcion) cMoneda,P.nTasaInteres,C.dApertura,C.nfirmas FROM ProductoPersona PP INNER JOIN Producto P INNER " _
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

If psCadPrd <> "" Then
    sSql = sSql & " AND SUBSTRING(PP.cCtaCod,6,3) IN (SELECT Valor FROM dbo.fnc_getTblValoresTexto ('" & psCadPrd & "'))" '" & nProd & "'"
End If

'*** BRGO 20111118 *******************************************
If sPrograma <> "" Then
    'sSql = sSql & " AND nTpoPrograma in ('" & sPrograma & "')"
    sSql = sSql & " AND ISNULL(nTpoPrograma,0) in (" & sPrograma & ")" 'EJVG20130920
End If
'*** END BRGO ************************************************
If bActivas Then
    sSql = sSql & " AND P.nPrdEstado NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ")"
End If
If bBloqueadas Then
    sSql = sSql & " AND P.nPrdEstado NOT IN (" & gCapEstBloqRetiro & "," & gCapEstBloqTotal & ")"
End If
If bBloqueadaTotal Then
    sSql = sSql & " AND P.nPrdEstado NOT IN (" & gCapEstBloqTotal & ")"
End If
If bOrdPag Then
    sSql = sSql & " AND A.bOrdPag = " & IIf(bOrdPag, 1, 0) & " "
End If
'By Capi 10042008 para que jale solo la cuenta elegida
If Len(Trim(sAgencia)) > 2 Then
    sSql = sSql & " AND P.cCtaCod = '" & sAgencia & "'"
Else
    If sAgencia <> "" Then
        sSql = sSql & " AND P.cCtaCod LIKE '___" & sAgencia & "%' "
    End If
End If

If bCtasDesemb Then '** Juez 20120723
    sSql = sSql & " AND PP.nPrdPersRelac=10 "
End If
'If nTiposCta <> -1 Then 'JUEZ 20131227
If nTiposCta <> "" Then 'JUEZ 20140520
    sSql = sSql & " AND nPrdCtaTpo in (" & nTiposCta & ")"
End If

sSql = sSql & " ORDER BY PP.cCtaCod"
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsCta.ActiveConnection = Nothing
Set GetCuentasPersona = rsCta
Set rsCta = Nothing
End Function
'*** BRGO 14/06/2011 **************************************************
Public Function GetCuentaAhorroTitularesPF(ByVal nTpoCta As Integer, ByVal cPersCod As String, ByVal cMoneda As String) As ADODB.Recordset
    Dim sSql As String
    Dim rsCta As ADODB.Recordset
    
    sSql = " exec stp_sel_CuentasAhorroTitularesPF " & nTpoCta & ",'" & cPersCod & "','" & cMoneda & "'"
    
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    Set rsCta.ActiveConnection = Nothing
    Set GetCuentaAhorroTitularesPF = rsCta
    Set rsCta = Nothing
End Function
'***********************************************************************
Public Function GetCuentasAhorro(ByVal bTodas As Boolean, Optional cAgeCod As String = "", Optional cCtaCod As String = "") As ADODB.Recordset
Dim sSql As String, sSqlaux As String
Dim rsCta As ADODB.Recordset

sSqlaux = ""
    If bTodas = True And cAgeCod <> "" Then
        sSqlaux = " and substring(p.cctacod,4,2)='" & cAgeCod & "'"
    End If
    
    If cCtaCod <> "" Then
        sSqlaux = " and p.cctacod='" & cCtaCod & "'"
    End If
    
  '**Modificado por DAOR 20081110, nSaldoDisp=nSaldoDisp-nRetencion (por cajeros) ********
  '  'By Capi 20012008 se agrego cNoCobroInactivas
  'sSql = " Select p.Cctacod,  "
  'sSql = sSql & " cpersnombre=( Select top  1  f.cpersnombre from persona f join  productopersona pp  on pp.cperscod=f.cperscod    where pp.nprdpersrelac=" & gCapRelPersTitular & "  and pp.cctacod=p.cctacod ),  "
  'sSql = sSql & " exonerada=(case when ca.bnocobroinactivas is null then 0 else case when ca.bnocobroinactivas=0 then 0 else 1 end end), "
  'sSql = sSql & " SaldoCnt=p.nsaldo,SaldoDisp=c.nsaldodisp,cNoCobroInactivas"
  'sSql = sSql & " from Producto p join Captaciones c on c.cctacod=p.cctacod join CaptacAhorros ca on ca.cctacod=p.cctacod  "
  'sSql = sSql & " where p.nprdestado not in (" & gCapEstAnulada & "," & gCapEstCancelada & " ) " & sSqlaux
  'sSql = sSql & " ORDER BY CPERSNOMBRE "
  sSql = " Select p.Cctacod,  "
  sSql = sSql & " cpersnombre=( Select top  1  f.cpersnombre from persona f join  productopersona pp  on pp.cperscod=f.cperscod    where pp.nprdpersrelac=" & gCapRelPersTitular & "  and pp.cctacod=p.cctacod ),  "
  sSql = sSql & " exonerada=(case when ca.bnocobroinactivas is null then 0 else case when ca.bnocobroinactivas=0 then 0 else 1 end end), "
  sSql = sSql & " SaldoCnt=p.nsaldo,(c.nSaldoDisp-ca.nRetencion) as SaldoDisp,cNoCobroInactivas"
  sSql = sSql & " from Producto p join Captaciones c on c.cctacod=p.cctacod join CaptacAhorros ca on ca.cctacod=p.cctacod  "
  sSql = sSql & " where p.nprdestado not in (" & gCapEstAnulada & "," & gCapEstCancelada & " ) " & sSqlaux
  sSql = sSql & " ORDER BY CPERSNOMBRE "
  '**********************************************************************************************
  
  Set rsCta = New ADODB.Recordset
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

Public Sub EliminaProductoPersona(ByVal sCuenta As String, Optional ByVal pnPrdPersRelac = 99) 'By Capi 06122007 se agrego prm opc para que elimine condicionalmente
Dim sSql As String
Dim pSql As String
Dim rs As ADODB.Recordset

sSql = "Delete ProductoPersona Where cCtaCod = '" & sCuenta & "'"
'By Capi 06122007 para controlar la condicion
If pnPrdPersRelac <> 99 Then
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    pSql = "Select cCtaCod From GiroDestinatario Where cCtaCod = '" & sCuenta & "'"
    rs.Open pSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    rs.ActiveConnection = Nothing
    If rs.RecordCount >= 1 Then
        sSql = sSql & " And nPrdPersRelac <> '" & pnPrdPersRelac & "'"
    Else
        Exit Sub
    End If
End If
'End By
dbCmact.Execute sSql
End Sub

Public Sub AgregaProductoPersona(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona, Optional ByVal COBLIGATORIO As String = "N", Optional ByVal sGrupo As String)
Dim sSql As String

' *** RIRO 20131102 SEGUN "CAMBIOS EN PODERES", SE AGREGO PARAMETRO "sGrupo"
sSql = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac,cObligatorio,cGrupo) " _
    & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ",'" & COBLIGATORIO & "','" & sGrupo & "')"

dbCmact.Execute sSql
End Sub

Public Sub AgregaMovLavDineroPersona(ByVal nMovNro As Long, ByVal cPersCod As String, ByVal nRelacion As MovLavDineroRelacion, ByVal nNoPresencial As Integer)
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    'sSql = "Exec stp_ins_RelacionMovLavDineroPersona " & nMovNro & ",'" & cPersCod & "'," & nRelacion NDX ERS0062020
    sSql = "Exec stp_ins_RelacionMovLavDineroPersona " & nMovNro & ",'" & cPersCod & "'," & nRelacion & "," & nNoPresencial
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Function GetCapBloqueos(ByVal sCuenta As String, ByVal nTipoBloqueo As CaptacTipoBloqueo, ByVal nConstante As ConstanteCabecera) As ADODB.Recordset
Dim sSql As String
Dim rsRel As ADODB.Recordset

Set rsRel = New ADODB.Recordset
rsRel.CursorLocation = adUseClient

sSql = "Select nEstado = CASE WHEN B.cMovNroDbl IS NULL THEN 1 ELSE 0 END, " _
    & "CONVERT(VARCHAR(10),CONVERT(Datetime,LEFT(cMovNro,8),101),105) dFecha, " _
    & "substring(cMovNro,9,2)+':'+substring(cMovNro,11,2)+':'+substring(cMovNro,13,2) dHora," _
    & "(C.cConsDescripcion + Space(75) + Convert(Varchar(3),C.nConsValor)) Motivo, "
    '** Juez 20120411 Se agregó Campo dHora
If nTipoBloqueo = gCapTpoBlqParcial Then
    sSql = sSql & "B.nMonto, "
End If
sSql = sSql & "RIGHT(B.cMovNro,4) Usu, dFechaDbl = CASE WHEN B.cMovNroDbl IS NULL THEN '' ELSE " _
    & "CONVERT(VARCHAR(10),CONVERT(Datetime,LEFT(B.cMovNroDbl,8),101),105) END, " _
    & "dHoraDbl = CASE WHEN B.cMovNroDbl IS NULL THEN '' " _
    & "ELSE substring(cMovNroDbl,9,2)+':'+substring(cMovNroDbl,11,2)+':'+substring(cMovNroDbl,13,2)  END," _
    & "UsuDbl = CASE WHEN B.cMovNroDbl IS NULL THEN '' ELSE RIGHT(B.cMovNroDbl,4) END, " _
    & "ISNULL(B.cComentario,'') cComentario, B.cMovNro, Flag = CASE WHEN B.cMovNroDbl IS NULL THEN 'A' ELSE '' END " _
    & "FROM ProductoBloqueos B INNER JOIN Constante C ON B.nBlqMotivo = C.nConsValor WHERE " _
    & "C.nConsCod = " & nConstante & " And B.cCtaCod = '" & sCuenta & "' AND B.nBlqTpo = " & nTipoBloqueo _
    & " Order by B.cMovNro"
    '** Juez 20120411 Se agregó Campo dHoraDbl

rsRel.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsRel.ActiveConnection = Nothing
Set GetCapBloqueos = rsRel
End Function

Public Function NuevoBloqueoRetiro(ByVal sCuenta As String, ByVal nMotivo As CaptacMotBloqueoRet, _
        cComentario As String, ByVal cMovnro As String)
Dim sSql As String
sSql = "INSERT ProductoBloqueos (cCtaCod,nBlqTpo,nBlqMotivo,cComentario,cMovNro,cMovNroDbl) " _
    & "VALUES ('" & sCuenta & "'," & gCapTpoBlqRetiro & "," & nMotivo & ",'" & cComentario & "','" & cMovnro & "',NULL)"

dbCmact.Execute sSql
End Function

Public Function NuevoBloqueoParcial(ByVal sCuenta As String, ByVal nMotivo As CaptacMotBloqueoParcial, _
        cComentario As String, ByVal cMovnro As String, ByVal nMonto As Double)
Dim sSql As String
sSql = "INSERT ProductoBloqueos (cCtaCod,nBlqTpo,nBlqMotivo,cComentario,cMovNro,cMovNroDbl,nMonto) " _
    & "VALUES ('" & sCuenta & "'," & gCapTpoBlqParcial & "," & nMotivo & ",'" & cComentario & "','" & cMovnro & "',NULL," & nMonto & ")"

dbCmact.Execute sSql
End Function

Public Function NuevoBloqueoTotal(ByVal sCuenta As String, ByVal nMotivo As CaptacMotBloqueoTot, _
        cComentario As String, ByVal cMovnro As String)
Dim sSql As String
sSql = "INSERT ProductoBloqueos (cCtaCod,nBlqTpo,nBlqMotivo,cComentario,cMovNro,cMovNroDbl) " _
    & "VALUES ('" & sCuenta & "'," & gCapTpoBlqTotal & "," & nMotivo & ",'" & cComentario & "','" & cMovnro & "',NULL)"

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

Public Sub ActualizaBloqueoParcial(ByVal sCuenta As String, ByVal sComentario As String, _
    ByVal sMovNroDbl As String, ByVal nMotivo As CaptacMotBloqueoParcial, ByVal sMovNro As String)
Dim sSql As String
sSql = "UPDATE ProductoBloqueos Set cComentario = '" & sComentario & "', " _
    & "cMovNroDbl = '" & sMovNroDbl & "' WHERE cCtaCod = '" & sCuenta & "' " _
    & "And nBlqTpo = " & gCapTpoBlqParcial & " AND nBlqMotivo = " & nMotivo & " AND cMovNro = '" & sMovNro & "'"

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
Dim rsBloq As ADODB.Recordset
bExiste = False
sSql = "Select cCtaCod FROM ProductoBloqueos WHERE cCtaCod = '" & sCuenta & "' " _
    & "And nBlqTpo = " & gCapTpoBlqTotal & " AND nBlqMotivo = " & nMotivo & " AND " _
    & "cMovNroDbl IS NULL"
Set rsBloq = New ADODB.Recordset
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
Dim rsBloq As ADODB.Recordset
bExiste = False
sSql = "Select cCtaCod FROM ProductoBloqueos WHERE cCtaCod = '" & sCuenta & "' " _
    & "And nBlqTpo = " & gCapTpoBlqRetiro & " AND nBlqMotivo = " & nMotivo & " AND " _
    & "cMovNroDbl IS NULL"
Set rsBloq = New ADODB.Recordset
rsBloq.CursorLocation = adUseClient
rsBloq.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsBloq.ActiveConnection = Nothing
If Not (rsBloq.EOF And rsBloq.BOF) Then
    bExiste = True
End If
Set rsBloq = Nothing
ExisteBloqueoRet = bExiste
End Function

Public Sub ActualizaMontoBlqParcial(ByVal sCuenta As String, ByVal nMonto As Double)
Dim sSql As String

sSql = "UPDATE Captaciones Set nBloqueoParcial = " & nMonto & " " _
    & "WHERE cCtaCod = '" & sCuenta & "'"

dbCmact.Execute sSql
End Sub

Public Sub ActualizaComentBlqParcial(ByVal sCuenta As String, ByVal sComentario As String, _
    ByVal nMotivo As CaptacMotBloqueoParcial, ByVal sMovNro As String)
Dim sSql As String

sSql = "UPDATE ProductoBloqueos Set cComentario = '" & sComentario & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "' AND nBlqTpo = " & gCapTpoBlqParcial & " " _
    & "AND nBlqMotivo = " & nMotivo & " AND cMovNro = '" & sMovNro & "'"

dbCmact.Execute sSql
End Sub


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
    Optional sCodAge As String = "", Optional bOrdPag As Boolean = False) As Double

Dim rsTasa As ADODB.Recordset
Dim sSql As String
Dim cOrdPag As String
cOrdPag = IIf(bOrdPag, "1", "0")

sSql = "SELECT nTasaValor FROM CaptacTasas WHERE nTasaProd = " & nProducto & " AND " _
    & "nTasaMon = " & nMoneda & " AND nTasaTpo = " & nTipoTasa & " And " & nPlazo & " " _
    & "BETWEEN nPlazoIni AND nPlazoFin And " & nValor & " BETWEEN nValorIni And nValorFin " _
    & "And cOrdPag = '" & cOrdPag & "' And cCodAge = '" & sCodAge & "'"

Set rsTasa = New ADODB.Recordset
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

Public Function GetPersonaTarj(ByVal sPersona As String) As ADODB.Recordset
Dim sSql As String
Dim rsTarj As ADODB.Recordset
sSql = "SELECT CT.cCtaCod, CT.cTarjCod FROM Producto P INNER JOIN ProductoPersona PP INNER JOIN " _
    & "CuentaTarj CT INNER JOIN Tarjeta T ON CT.cTarjCod = T.cTarjCod ON PP.cCtaCod = CT.cCtaCod " _
    & "AND PP.cPersCod = CT.cPersCod ON P.cCtaCod = PP.cCtaCod WHERE CT.cPersCod = '" & sPersona & "' " _
    & "AND T.nEstado NOT IN (" & gCapTarjEstCancelada & ") AND P.nPrdEstado NOT IN (" & gCapEstAnulada & "," _
    & gCapEstAnulada & "," & gCapEstCancelada & "," & gCapEstCancelada & ")"
Set rsTarj = New ADODB.Recordset
rsTarj.CursorLocation = adUseClient
rsTarj.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsTarj.ActiveConnection = Nothing
Set GetPersonaTarj = rsTarj
Set rsTarj = Nothing
End Function

Public Function GetTarjetaCuentas(ByVal sTarjeta As String, Optional nProducto As Producto = gCapTodoAhorro) As ADODB.Recordset
Dim sSql As String
Dim rsTarj As ADODB.Recordset

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


Set rsTarj = New ADODB.Recordset
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
Dim rsTarj As ADODB.Recordset
Set rsTarj = New ADODB.Recordset
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
Dim rsTarj As ADODB.Recordset
Set rsTarj = New ADODB.Recordset
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
Dim rsTarj As ADODB.Recordset
Set rsTarj = New ADODB.Recordset
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
Dim rsTarj As ADODB.Recordset
Set rsTarj = New ADODB.Recordset
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

Public Function GetBeneficiarios(ByVal sPersona As String) As ADODB.Recordset
Dim rsPers As ADODB.Recordset
Dim sSql As String

sSql = "SELECT PR.cPersRelacPersCod, P.cPersNombre, T.cConsDescripcion cParentesco, " _
    & "DATEDIFF(yy,dPersNacCreac,Getdate()) nEdad, nPersRelacBenefPorc, PR.nPersRelac " _
    & "FROM " & sDBPersona & "Persona P INNER JOIN PersRelaciones PR INNER JOIN Constante T " _
    & "ON PR.nPersRelac = T.nConsValor ON P.cPersCod = PR.cPersRelacPersCod WHERE " _
    & "bPersRelacBenef = 1 AND PR.cPersCod = '" & sPersona & "' And T.nConsCod = " & gPersRelacion
    
Set rsPers = New ADODB.Recordset
rsPers.CursorLocation = adUseClient
rsPers.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsPers.ActiveConnection = Nothing
Set GetBeneficiarios = rsPers
Set rsPers = Nothing
End Function

'ARCV se agrego dPersNacCreac y EstadoCivil
Public Function GetDatosPersona(ByVal sPersona As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset
'FRHU 20140926 ERS099-2014
'sSql = "Select P.cPersNombre Nombre, P.cPersDireccDomicilio Direccion, ISNUll(U.cUbiGeoDescripcion,'') Zona, " _
'    & "ISNUll(P.cPersTelefono,'') Fono, ISNUll(T.cConsDescripcion,'') ID, ISNull(I.cPersIDnro,'') [ID N°], P.cPersCod, P.nPersPersoneria, " _
'    & "I.cPersIDtpo " & _
'    " ,dPersNacCreac,cEstadoCivil=ISNULL(EC.cConsDescripcion,''),cCIIUdescripcion=ISNULL(cCIIUdescripcion,'')  " & _
'    " FROM " & sDBPersona & "PersId I RIGHT JOIN " & sDBPersona & "Persona P LEFT JOIN " _
'    & sDBComunes & "UbicacionGeografica U ON P.cPersDireccUbiGeo = U.cUbiGeoCod ON I.cPersCod = P.cPersCod " _
'    & "LEFT JOIN " & sDBComunes & "Constante T ON I.cPersIDTpo = T.nConsValor And T.nConsCod = " & gPersIdTipo & _
'    " LEFT JOIN " & sDBPersona & "PersonaNat PNat ON PNat.cPersCod=P.cPersCod LEFT JOIN Constante EC ON EC.nConsValor=PNat.nPersNatEstCiv AND EC.nConsCod=1020" & _
'    " LEFT JOIN CIIU CI on CI.cCIIUCod= P.cPersCIIU  " & _
'    "  WHERE P.cPersCod = '" & sPersona & "' "
sSql = "stp_sel_GetDatosPersona '" & sPersona & "'"
'FIN FRHU 20140926
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsCta.ActiveConnection = Nothing
Set GetDatosPersona = rsCta
Set rsCta = Nothing
End Function

Public Function GetDatosOrdenPago(ByVal sCuenta As String, ByVal nNumOP As Long) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset
sSql = "Select D.cNroDoc, K.cConsDescripcion cDescripcion,  D.nMonto,dbo.FechaMov(D.cMovNro) Fecha, " _
    & "RIGHT(D.cMovNro,4) Usuario, D.nEstado FROM DocRecOP D INNER JOIN " & sDBComunes & "Constante K ON D.nEstado = K.nConsValor " _
    & "WHERE D.cCtaCod = '" & sCuenta & "' AND D.cNroDoc = '" & nNumOP & "' And K.nConsCod = " & COMDConstantes.gCaptacOrdPagoEstado _
    & " And D.cMovNro IN (Select MAX(cMovNro) From DocRecOP D1 Where D1.cCtaCod = D.cCtaCod And D1.cNroDoc = D.cNroDoc) " _
    & "ORDER BY D.cMovNro DESC"
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsCta.ActiveConnection = Nothing
Set GetDatosOrdenPago = rsCta
Set rsCta = Nothing
End Function

Public Function GetPersonaCuenta(ByVal sCuenta As String, Optional nTipRel As CaptacRelacPersona = 0) As ADODB.Recordset

Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim lbflag As Boolean

'***Modificado por ELRO el 20121123, según SATI INC1211220007
'sSql = " Select P.cPersNombre Nombre,  ISNULL(T1.cConsDescripcion,'') +  Space(50) +  str(T1.nConsValor) Relacion, ISNULL(P.cPersDireccDomicilio,'') Direccion, ISNULL(U.cUbiGeoDescripcion,'') Zona, " _
'    & " P.cPersTelefono Fono, ISNULL(T.cConsDescripcion,'') ID, ISNULL(I.cPersIDnro,'') [ID N°], P.cPersCod, P.nPersPersoneria, " _
'    & " I.cPersIDtpo, PP.nPrdPersRelac, PP.cCtaCod,pp.bobligatorio,pp.cobligatorio FROM " & sDBPersona & "PersId I RIGHT JOIN (" & sDBPersona & "Persona P LEFT JOIN " _
'    & sDBComunes & "UbicacionGeografica U ON P.cPersDireccUbiGeo = U.cUbiGeoCod) RIGHT JOIN ProductoPersona PP " _
'    & "ON P.cPersCod = PP.cPersCod ON I.cPersCod = P.cPersCod LEFT JOIN " & sDBComunes & "Constante T ON I.cPersIDTpo = T.nConsValor   AND T.nConsCod = " & gPersIdTipo & "  " _
'    & "LEFT JOIN " & sDBComunes & "Constante T1 ON PP.nPrdPersRelac = T1.nConsValor AND T1.nConsCod = " & gCaptacRelacPersona & " WHERE PP.cCtaCod = '" & sCuenta & "'  and  (PP.nPrdPersRelac = '" & nTipRel & "' or  0=" & nTipRel & ") and PP.nPrdPersRelac not in('" & gCapRelPersPromotor & "')   " _
'    & " ORDER BY PP.nPrdPersRelac,Nombre, I.cPersIDtpo"
sSql = "exec stp_sel_DevolverPersonaCuenta '" & sCuenta & "', " & nTipRel & ""
'***Modificado por ELRO el 20121123**************************
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsCta.ActiveConnection = Nothing
Set GetPersonaCuenta = rsCta
Set rsCta = Nothing
End Function
Public Function GetDatosOP(Optional ByVal sAgecod As String = "") As Recordset
Dim sSql As String, sqlaux
Dim rsCta As ADODB.Recordset
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
 

Set rsCta = New ADODB.Recordset
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
Public Sub AgregaCartaAIP(ByVal cMovnro As String, ByVal cNroRef As String, ByVal cReferencia As String, ByVal NTipoDelito As Integer, ByVal Entidad As String, ByVal Secretario As String, ByVal nTipoDoc As Integer, ByVal cNroDoc As String, ByVal Ubigeo As String, ByRef NroCarta As String)
Dim sSql As String
Dim rsCta As ADODB.Recordset

On Error GoTo ErrGraba

dbCmact.BeginTrans

    sSql = " insert into Cartasai(NroCarta,cMovnro,cNroRef,cReferencia,NTipoDelito,Entidad,Secretario,nTipoDoc,cNroDoc,Ubigeo) " _
         & " values(dbo.GetNroCartaAI(), '" & cMovnro & "' , '" & cNroRef & "' , '" & cReferencia & "' , " & NTipoDelito & " , '" & Entidad & "','" & Secretario & "' ," & nTipoDoc & " ,'" & cNroDoc & "' ,'" & Ubigeo & "' ) "
    
    dbCmact.Execute sSql
    
    sSql = "Select nrocarta from cartasai where cmovnro='" & cMovnro & "'"
    Set rsCta = New ADODB.Recordset
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


Public Sub AgregaCartaAI(ByVal cMovnro As String, ByVal cNroRef As String, ByVal cReferencia As String, ByVal NTipoDelito As Integer, ByVal Entidad As String, ByVal Secretario As String, ByVal nTipoDoc As Integer, ByVal cNroDoc As String, ByVal Ubigeo As String, ByVal NroCarta As String)
Dim sSql As String

    sSql = " insert into Cartasai(NroCarta,cMovnro,cNroRef,cReferencia,NTipoDelito,Entidad,Secretario,nTipoDoc,cNroDoc,Ubigeo) " _
         & " values('" & NroCarta & "', '" & cMovnro & "' , '" & cNroRef & "' , '" & cReferencia & "' , " & NTipoDelito & " , '" & Entidad & "','" & Secretario & "' ," & nTipoDoc & " ,'" & cNroDoc & "' ,'" & Ubigeo & "' ) "
         
dbCmact.Execute sSql

End Sub

Public Sub AgregaBitacoraPPer(ByVal cMovnro As String, ByVal cCtaCod As String, ByVal cPersCod As String, ByVal NPRDPERSRELAC As String, ByVal COBLIGATORIO As String, Optional cGrupo As String = "")
Dim sSql As String

    ' *** RIRO 20131102 SEGUN "CAMBIOS EN PODERES", SE AGREGO PARAMETRO "cGrupo"
    sSql = " insert into BITACORAPRODPERSONA(CMOVNRO,CCTACOD,CPERSCOD,NPRDPERSRELAC,COBLIGATORIO, cGrupo) " _
         & " values('" & cMovnro & "','" & cCtaCod & "','" & cPersCod & "'," & NPRDPERSRELAC & ",'" & COBLIGATORIO & "', '" & cGrupo & "' ) "
         

dbCmact.Execute sSql

End Sub

Public Sub AgregaExtornoCapAumDisPF(ByVal sCuenta As String, ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double, ByVal dUltMov As Date, ByVal nTasa As Double, _
                                    ByVal nIntAcum As Double, ByVal dAuxiliar As Date, ByVal nExtracto As Integer, ByVal nIntPag As Double, _
                                    ByVal nIntDevChq As Double, ByVal nPlazo As Integer, ByVal sUltActualizacion As String, ByVal dRenovacion As Date, _
                                    ByVal nMovNro As Long)

Dim sSql As String
sSql = "INSERT CaptacExtornoCapAumDisPF (cCtacod,nSaldoDisp,nSaldoCnt,dUltimoCierre,nTasaInteres,nIntAcum,dAuxiliar,nTransacc,nIntPag,nIntDevChq,nPlazo,cUltActualizacion,dRenovacion,nMovNro) " _
    & "VALUES ('" & sCuenta & "','" & nSaldoDisp & "','" & nSaldoCnt & "','" & Format(dUltMov, "mm/dd/yyyy") & "','" & nTasa & "','" & nIntAcum & "','" & Format(dAuxiliar, "mm/dd/yyyy") & "','" & nExtracto & "','" & nIntPag & "','" & nIntDevChq & "','" & nPlazo & "','" & sUltActualizacion & "','" & Format(dRenovacion, "mm/dd/yyyy") & "','" & nMovNro & "')"

dbCmact.Execute sSql
End Sub


Public Sub ActualizaDatosCuenta(ByVal sCuenta As String, ByVal bOrdPag As Boolean, _
        ByVal nFirmas As Integer, nTipoCuenta As ProductoCuentaTipo, _
        ByVal sInstitucion As String, Optional sCuentaAboInt As String = "", Optional ByVal nFirmasMin As Integer, Optional ByVal sAlias As String = "", _
        Optional ByVal nTpoPrograma As Integer = 0, Optional ByVal pnPorRetAdiCTS As Integer = 0, Optional ByVal psMovNro As String, Optional ByVal nSubasta As Integer, Optional ByVal sReglas As String)
        'MIOL 20121109, SEGUN RFC098-2012-B SE AGREGO: nSubasta
        'Parametros pnPorRetAdiCTS y psMovNro by capi Agosto-2007(Acta 014-2007)
        'RIRO20131102 SEGUN "CAMBIOS EN PODERES", SE AGREGO PARAMETR "sReglas"

Dim sSql As String
Dim nProd As Producto

nProd = CLng(Mid(sCuenta, 6, 3))
If nProd = gCapAhorros Then
    sSql = "Update CaptacAhorros Set bOrdPag = " & IIf(bOrdPag, 1, 0) & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
ElseIf nProd = gCapCTS Then
    If pnPorRetAdiCTS <> 0 Then
         'By Capi Agosto 2007 (Acta 014-2007)
         sSql = "Update CaptacCTS Set nRetAdiCTS=Isnull(nRetAdiCTS+1,1),nSaldRetiro =Round(nSaldRetiro +(Select nSaldoDisp * " & pnPorRetAdiCTS & "/100 From Captaciones WHERE cCtaCod = '" & sCuenta & "'" & " ),2)" & "WHERE cCtaCod = '" & sCuenta & "'"
         dbCmact.Execute sSql
         'sSql = "Insert Mov (nMovNro,cMovNro,cOpeCod,cMovDesc) Values ()"
         'dbCmact.Execute sSql
    End If
    sSql = "Update CaptacCTS Set cCodInst = '" & sInstitucion & "' " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
ElseIf nProd = gCapPlazoFijo Then
    sSql = "DELETE CaptacCtaAboIntPF Where cCtaCod = '" & sCuenta & "'"
    If sCuentaAboInt <> "" Then
        dbCmact.Execute sSql
        sSql = "INSERT CaptacCtaAboIntPF (cCtaCod,cCtaCodAbono) VALUES ('" & sCuenta & "','" & sCuentaAboInt & "')"
    End If
    dbCmact.Execute sSql
    'MIOL 20121109, SEGUN RFC098-B ***********************************
    sSql = "UPDATE CaptacPlazoFijo SET bAbonoIntCtaAho = " & IIf(sCuentaAboInt <> "", 1, 0) & ", Subasta=" & nSubasta & " WHERE cCtaCod='" & sCuenta & "'"
    'sSql = "UPDATE CaptacPlazoFijo SET bAbonoIntCtaAho = " & IIf(sCuentaAboInt <> "", 1, 0) & " WHERE cCtaCod='" & sCuenta & "'"
    'END MIOL ********************************************************
End If
dbCmact.Execute sSql

If nProd = gCapAhorros Then
    
    'by capi 09012009 porque actualizaba el tpoprograma equivocadamente.
    
'    sSql = "Update Captaciones Set nFirmas = " & nFirmas & ", nfirmasmin=" & nFirmasMin & ", calias='" & sAlias & "' , nTpoPrograma=" & nTpoPrograma & "," _
'          & " nPrdCtaTpo = " & nTipoCuenta & " WHERE cCtaCod = '" & sCuenta & "'"
'          dbCmact.Execute sSql
          
    ' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            'sSql = "Update Captaciones Set nFirmas = " & nFirmas & ", nfirmasmin=" & nFirmasMin & ", calias='" & sAlias & "'," _
            '& " nPrdCtaTpo = " & nTipoCuenta & " WHERE cCtaCod = '" & sCuenta & "'"
        
    ' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            sSql = "Update Captaciones Set nFirmas = " & nFirmas & ", nfirmasmin=" & nFirmasMin & ", calias='" & sAlias & "'," _
                & " nPrdCtaTpo = " & nTipoCuenta & ", cReglas = '" & sReglas & "' WHERE cCtaCod = '" & sCuenta & "'"
    
    ' *** END RIRO
    
          dbCmact.Execute sSql
' End By
Else
    
    ' *** COMENTADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            'sSql = "Update Captaciones Set nFirmas = " & nFirmas & ", nfirmasmin=" & nFirmasMin & ", calias='" & sAlias & "' , " _
            '& " nPrdCtaTpo = " & nTipoCuenta & " WHERE cCtaCod = '" & sCuenta & "'"
        
    ' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
            sSql = "Update Captaciones Set nFirmas = " & nFirmas & ", nfirmasmin=" & nFirmasMin & ", calias='" & sAlias & "' , " _
                  & " nPrdCtaTpo = " & nTipoCuenta & ", cReglas = '" & sReglas & "' WHERE cCtaCod = '" & sCuenta & "'"
    
    ' *** END RIRO
    
            dbCmact.Execute sSql

End If

End Sub

Public Function GetOrdenPagoEmitidas(ByVal sCuenta As String) As Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset
Set rsCta = New ADODB.Recordset
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
Dim rsCta As ADODB.Recordset
Set rsCta = New ADODB.Recordset
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

Public Function GetEstadoOrdenPagoEmitida(ByVal sCuenta As String, ByVal nNumOP As Long) As ADODB.Recordset
Dim sSql As String
Dim rsOP As ADODB.Recordset
sSql = "Select Top 1 E.nEstado, K.cConsDescripcion Estado FROM DocRecOPEst E INNER JOIN DocRecOP D ON E.nTpoDoc = D.nTpoDoc " _
    & "AND E.cNroDoc = D.cNroDoc INNER JOIN " & sDBComunes & "Constante K ON E.nEstado = K.nConsValor Where " _
    & "D.cCtaCod = '" & sCuenta & "' And D.nTpoDoc = " & TpoDocOrdenPago & " And E.cNroDoc = '" & nNumOP & "' " _
    & "And K.nConsCod = " & gCaptacOrdPagoEstado & " Order by E.cMovNro DESC"
Set rsOP = New ADODB.Recordset
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
sSql = "INSERT DocRecOp (nTpoDoc,cNroDoc,cCtaCod,cMovNro,nMonto,nEstado) " _
    & "VALUES (" & TpoDocOrdenPago & ",'" & nNumOP & "','" & sCuenta & "','" & sMovNro & "'," & nMonto & "," & COMDConstantes.gCapOPEstAnulada & ")"

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

Public Function GetTarjetaEstadoHist(ByVal sTarjeta As String) As ADODB.Recordset
Dim sSql As String
Dim rsTarj As ADODB.Recordset

sSql = "SELECT CONVERT(Varchar(10),CONVERT(Datetime,LEFT(E.cMovNro,8)),103) Fecha, K.cConsDescripcion Estado, cDescripcion, " _
    & "RIGHT(E.cMovNro,4) Usu FROM Tarjeta T INNER JOIN TarjetaEstado E ON T.cTarjCod = E.cTarjCod INNER JOIN " _
    & sDBComunes & "Constante K ON E.nEstado = K.nConsValor WHERE T.cTarjCod = '" & sTarjeta & "' AND K.nConsCod = " _
    & gCaptacTarjetaEstado & " ORDER BY Fecha DESC"

Set rsTarj = New ADODB.Recordset
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
Dim oFuncion As New COMFunciones.FCOMCadenas


'sSql = "SELECT Pe.cPersNombre "
'sSql = sSql & " FROM (SELECT cPersCod FROM ProductoPersona WHERE P.cCtaCod = '" & sCuenta & "'" AND nPrdPersRelac IN (" & gCapRelPersTitular & "," & gGiroRelPersRemitente & ")"
'sSql = sSql & " INNER JOIN " & sDBPersona & "Persona Pe ON PP.cPersCod = Pe.cPersCod "
'sSql = sSql & " WHERE P.cCtaCod = '" & sCuenta & "'"
sSql = "SELECT P.cPersNombre "
sSql = sSql & " FROM Persona P"
sSql = sSql & " INNER JOIN ("
sSql = sSql & "     SELECT cPersCod FROM ProductoPersona WHERE cCtaCod = '" & sCuenta & "' AND nPrdPersRelac IN (" & gCapRelPersTitular & "," & gGiroRelPersRemitente & ")"
sSql = sSql & " ) PP ON P.cPersCod = PP.cPersCod"
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
                sCadena = sCadena & oFuncion.PstaNombre(rsTit("cPersNombre"), False) & Chr(10)
            Else
                sCadena = sCadena & Space(pnMargenIzquierdo) & oFuncion.PstaNombre(rsTit("cPersNombre"), False) & Chr(10)
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
    lstOpe = oFuncion.PstaNombre(rsTit("cPersNombre"), False)
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
'rsTit.Close
Set oFuncion = Nothing
Set rsTit = Nothing
End Function

Public Function GetCTSPeriodo() As ADODB.Recordset
Dim sSql As String
Dim rsPer As ADODB.Recordset
sSql = "Select nItem,nMesInicio,nMesFin,cDescripcion,nPorcentaje FROM CaptacCTSPeriodo " _
    & "Order by nItem"
Set rsPer = New ADODB.Recordset
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

Public Function GetSaldoFecha(ByVal sCuenta As String, ByVal dFecha As Date) As ADODB.Recordset

Dim sSql As String, sNumMov As String

Dim rsCta As ADODB.Recordset

sSql = "Select TOP 1 CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha, " _
    & "C.nSaldoDisponible, C.nSaldoContable FROM Mov M INNER JOIN MovCap C " _
    & "INNER JOIN MovCapDet CD INNER JOIN OpeTpo O INNER JOIN CapMovTipo CT ON " _
    & "O.cOpeCod = CT.cOpeCod ON CD.cOpeCod = O.cOpeCod ON C.nMovNro = CD.nMovNro ON " _
    & "M.nMovNro = C.nMovNro WHERE C.cCtaCod = '" & sCuenta & "' AND CD.nConceptoCod " _
    & "= " & gConcCapital & " AND LEFT(M.cMovNro,8) <= '" & Format$(dFecha, "yyyymmdd") & "' " _
    & "AND M.nMovFlag =" & gMovFlagVigente & " ORDER BY M.nMovNro DESC"

Set rsCta = New ADODB.Recordset

rsCta.CursorLocation = adUseClient

'dbCmact.ConnectionTimeout = 120

rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText





Set rsCta.ActiveConnection = Nothing

Set GetSaldoFecha = rsCta

Set rsCta = Nothing

End Function

Public Function GetMovimientosCuenta(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0, _
        Optional bMuestraExtornos As Boolean = True) As ADODB.Recordset


   Dim sSql As String ', sNumMov As String, sMuestraExtornos As String'WIOR 20140625 COMENTO
   'Dim rsCta As ADODB.Recordset'WIOR 20140625 COMENTO
   Dim oConecta As COMConecta.DCOMConecta 'WIOR 20140625
   
   On Error GoTo ErrorGetMovimientos 'JUEZ 20150205
   
'***Modificado por ELRO el 20120814, según RQ12238-2012
'    sNumMov = ""
'    sMuestraExtornos = ""
'    If nNumMov > 0 Then sNumMov = " TOP " & nNumMov + 1 & " "
'    If Not bMuestraExtornos Then
'        sMuestraExtornos = " And M.nMovFlag IN (" & gMovFlagVigente & ") "
'    End If
'
'    sSql = " Select Distinct " & sNumMov & " CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(IsNull(MOVOH.CMOVNRO,M.CMOVNRO),1,8)),103) + ' ' + SUBSTRING(IsNull(MOVOH.CMOVNRO,M.CMOVNRO),9,2)+ ':' + SUBSTRING(IsNull(MOVOH.CMOVNRO,M.CMOVNRO),11,2) + ':' + SUBSTRING(IsNull(MOVOH.CMOVNRO,M.CMOVNRO),13,2) Fecha, " _
'        & " O.cOpeDesc Operacion, ISNULL(MD.cDocNro,'') cDocumento , nAbono = CASE " _
'        & " WHEN CT.nCapMovTpo IN (" & gCapMovApertura & "," & gCapMovDeposito & "," & gCapMovIntCap & ") THEN ABS(C.nMonto) " _
'        & " ELSE 0 END, " _
'        & " nCargo = CASE WHEN CT.nCapMovTpo IN (" & gCapMovRetiro & "," & gCapMovRetiroInac & "," & gCapMovCancelAct & "," & gCapMovCancelInact & "," & gCapMovRetiroInt & ") THEN ABS(C.nMonto)*-1 " _
'        & " ELSE 0 END, " _
'        & " C.nSaldoContable, A.cAgeDescripcion cAgencia, cUsu = RIGHT(IsNull(MOVOH.CMOVNRO,M.CMOVNRO),4), CAST(LEFT(M.cMovNro,14) AS BigInt) cMovNroNue,M.cOpeCod  , IsNull(MOVOH.CMOVNRO,M.CMOVNRO) cMovNro " _
'        & " FROM MovDoc MD " _
'        & " RIGHT JOIN Mov M  ON MD.nMovNro = M.nMovNro " _
'        & " INNER JOIN MovCap C ON M.nMovNro = C.nMovNro    " _
'        & " INNER JOIN MovCapDet CD ON c.copecod=cd.copecod And C.cCtaCod = CD.cCtaCod AND CD.nMovNro = C.nMovNro " _
'        & " INNER JOIN OpeTpo O ON C.cOpeCod = O.cOpeCod" _
'        & " INNER JOIN CapMovTipo CT ON O.cOpeCod = CT.cOpeCod " _
'        & " LEFT  JOIN DBTARJETA..ATMMOV ATMMOV ON ATMMOV.nMovNro = M.nMovNro " _
'        & " LEFT  JOIN DBTARJETA..MOVOFFHOST MOVOH ON MOVOH.nMovNro = ATMMOV.nMovNroOffHost " _
'        & "INNER JOIN Agencias A ON SUBSTRING(M.cMovNro,18,2) = A.cAgeCod " _
'        & "WHERE C.cCtaCod in (Select Valor from fnc_getTblValoresTexto('" & sCuenta & "'))" & sMuestraExtornos 'BRGO 20111220 Considera mas de una cuenta para casos de migración
'
'If CLng(Mid(sCuenta, 6, 3)) = gCapPlazoFijo Then
'     sSql = sSql & " AND (CD.nConceptoCod IN (" & gConcCapital & "," & gConcInteres & ", " & gConcITFCliente & " , " & gConcProvision & ", 12) or C.cOpeCod in ('210501'))"
'     'sSql = sSql & " AND CD.nConceptoCod IN (" & gConcCapital & ", " & gConcITFCliente & ", " & gConcProvision & ")" 'ARCV 26-06-2007
'Else
'     sSql = sSql & " AND CD.nConceptoCod IN (" & gConcCapital & ", " & gConcITFCliente & ", " & gConcOpeRetOtraAge & "," & gConcOpeDepOtraAge & "," & gConcOpeRetxMaxOpe & ")"
'End If
'
'If nNumMov > 0 Then
'    sSql = sSql & " ORDER BY CAST(LEFT(M.cMovNro,14) AS BIGINT) DESC"
'Else
'    sSql = sSql & " AND LEFT(IsNull(MOVOH.CMOVNRO,M.CMOVNRO),8) BETWEEN '" & Format$(dFecIni, "yyyymmdd") & "' " _
'        & "AND '" & Format$(dFecFin, "yyyymmdd") & "'"
'    sSql = sSql & " ORDER BY cMovNro,M.cOpeCod     "
'   'sSql = sSql & " ORDER BY CAST(LEFT(M.cMovNro,14) AS BIGINT)"
'End If

If nNumMov > 0 Then
    sSql = "stp_sel_devolverMovimientosXNroMovCuenta '" & sCuenta & "', " & nNumMov & ", " & IIf(bMuestraExtornos, 1, 0) & ""
Else
    sSql = "stp_sel_devolverMovimientosXFechaCuentas '" & sCuenta & "', '" & Format$(dFecIni, "yyyymmdd") & "', '" & Format$(dFecFin, "yyyymmdd") & "', " & IIf(bMuestraExtornos, 1, 0) & ""
End If

'WIOR 20140625 *****************************
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set GetMovimientosCuenta = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
'WIOR FIN **********************************

'WIOR 20140625 (COMENTO TODO ESTA PARTE) ********************
'Set rsCta = New ADODB.Recordset
'rsCta.CursorLocation = adUseClient
'rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'
'
'Set rsCta.ActiveConnection = Nothing
'Set GetMovimientosCuenta = rsCta
'Set rsCta = Nothing
'WIOR FIN (COMENTARIO) ***************************************
Exit Function
ErrorGetMovimientos: 'JUEZ 20150205
    Err.Raise Err.Number, "Error", Err.Description
End Function
'By Capi 11112008 caj aut pendientes
Public Function GetMovimientosPendientesCuenta(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0, _
        Optional bMuestraExtornos As Boolean = True) As ADODB.Recordset

    Dim sSql As String, sNumMov As String, sMuestraExtornos As String
    'Dim rsCta As ADODB.Recordset 'WIOR 20140625 COMENTO
    Dim oConecta As COMConecta.DCOMConecta 'WIOR 20140625
    
    On Error GoTo ErrorGetMovimientos 'JUEZ 20150205
    
'    sNumMov = ""
'    sMuestraExtornos = ""
'    If nNumMov > 0 Then sNumMov = " TOP " & nNumMov + 1 & " "
'    If Not bMuestraExtornos Then
'        sMuestraExtornos = " And M.nMovFlag IN (" & gMovFlagVigente & ") "
'    End If
'    '** Modify By GITU 2012-09-03 se agrego la descripcion de la scompras en POS VISA
'    sSql = "Select  Distinct " & sNumMov & " CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha, " _
'            & "Case When M.cOpeCod = '208035' Then 'COMPRA POS VISA ' + IsNull(TRA.cNomTerm,'') Else O.cOpeDesc  End Operacion, " _
'            & "ISNULL(MD.cDocNro,'') cDocumento , nAbono = CASE " _
'            & "WHEN CT.nCapMovTpo IN (" & gCapMovApertura & "," & gCapMovDeposito & "," & gCapMovIntCap & ") THEN ABS(C.nMonto) " _
'            & "ELSE 0 END, " _
'            & "nCargo = CASE WHEN CT.nCapMovTpo IN (" & gCapMovRetiro & "," & gCapMovRetiroInac & "," & gCapMovCancelAct & "," & gCapMovCancelInact & "," & gCapMovRetiroInt & ") THEN ABS(C.nMonto)*-1 " _
'            & "ELSE 0 END, " _
'            & "C.nSaldoContable, A.cAgeDescripcion cAgencia, cUsu = RIGHT(cMovNro,4), CAST(LEFT(cMovNro,14) AS BigInt) cMovNroNue,M.cOpeCod  ,M.cMovNro " _
'            & "FROM MovDoc MD " _
'            & "RIGHT JOIN DBTARJETA..ATMMov AM on MD.nMovNro = AM.nMovNroOffHost " _
'            & "INNER JOIN DBTARJETA..MovOFFHOST M  ON AM.nMovNroOFFHost = M.nMovNro " _
'            & "INNER JOIN DBTARJETA..MovCapOFFHOST C ON M.nMovNro = C.nMovNro " _
'            & "LEFT JOIN DBTARJETA..Tramas TRA ON TRA.cIDTrama = AM.trace And AM.cPan = TRA.Pan " _
'            & "INNER JOIN DBTARJETA..MovCapDetOFFHOST CD ON c.copecod=cd.copecod And C.cCtaCod = CD.cCtaCod AND CD.nMovNro = C.nMovNro " _
'            & "INNER JOIN OpeTpo O ON C.cOpeCod = O.cOpeCod " _
'            & "INNER JOIN CapMovTipo CT ON O.cOpeCod = CT.cOpeCod " _
'            & "INNER JOIN Agencias A ON SUBSTRING(M.cMovNro,18,2) = A.cAgeCod " _
'            & "WHERE C.cCtaCod COLLATE Modern_Spanish_CI_AS in (Select Valor from fnc_getTblValoresTexto('" & sCuenta & "'))" & sMuestraExtornos 'BRGO 20111220 Considera mas de una cuenta para casos de migración
'    sSql = sSql & "AND CD.nConceptoCod IN (" & gConcCapital & ", " & gConcITFCliente & ", " & gConcOpeRetOtraAge & " ," & gConcOpeRetxMaxOpe & " )"
'    sSql = sSql & "AND AM.nMovNro=AM.nMovNroOffHost And C.nMonto>0 And TRA.cIDTrama Is Not Null"
'    If nNumMov > 0 Then
'        sSql = sSql & " ORDER BY CAST(LEFT(M.cMovNro,14) AS BIGINT) DESC"
'    Else
'        sSql = sSql & " AND LEFT(M.cMovNro,8) BETWEEN '" & Format$(dFecIni, "yyyymmdd") & "' " _
'            & "AND '" & Format$(dFecFin, "yyyymmdd") & "'"
'        sSql = sSql & " ORDER BY M.cMovNro,M.cOpeCod     "
'        'sSql = sSql & " ORDER BY CAST(LEFT(M.cMovNro,14) AS BIGINT)"
'    End If

If nNumMov > 0 Then
    sSql = "stp_sel_devolverMovimientosPendientesXNroMovCuenta '" & sCuenta & "', " & nNumMov & ", " & IIf(bMuestraExtornos, 1, 0) & ""
Else
    sSql = "stp_sel_devolverMovimientosPendientesXFechaCuentas '" & sCuenta & "', '" & Format$(dFecIni, "yyyymmdd") & "', '" & Format$(dFecFin, "yyyymmdd") & "', " & IIf(bMuestraExtornos, 1, 0) & ""
End If

'WIOR 20140625 *****************************
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Set GetMovimientosPendientesCuenta = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
'WIOR FIN **********************************

'WIOR 20140625 (COMENTO TODO ESTA PARTE) ********************
'Set rsCta = New ADODB.Recordset
'rsCta.CursorLocation = adUseClient
'rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'
'
'
'Set rsCta.ActiveConnection = Nothing
'Set GetMovimientosPendientesCuenta = rsCta
'Set rsCta = Nothing
'WIOR FIN (COMENTARIO) ***************************************
Exit Function
ErrorGetMovimientos: 'JUEZ 20150205
    Err.Raise Err.Number, "Error", Err.Description
End Function
'JUEZ 20150205 *********************************************************
Public Function GetMovimientosDebitosAutoCuenta(ByVal sCuenta As String, Optional dFecIni As Date, _
                                                Optional dFecFin As Date, Optional nNumMov As Long = 0, _
                                                Optional bMuestraExtornos As Boolean = True) As ADODB.Recordset

Dim sSql As String, sNumMov As String, sMuestraExtornos As String
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorGetMovimientos

    If nNumMov > 0 Then
        sSql = "stp_sel_devolverMovimientosDebitosAutoXNroMovCuenta '" & sCuenta & "', " & nNumMov & ", " & IIf(bMuestraExtornos, 1, 0) & ""
    Else
        sSql = "stp_sel_devolverMovimientosDebitosAutoXFechaCuentas '" & sCuenta & "', '" & Format$(dFecIni, "yyyymmdd") & "', '" & Format$(dFecFin, "yyyymmdd") & "', " & IIf(bMuestraExtornos, 1, 0) & ""
    End If
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GetMovimientosDebitosAutoCuenta = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
ErrorGetMovimientos:
    Err.Raise Err.Number, "Error", Err.Description
End Function
'END JUEZ **************************************************************
Public Function GetMovimientosCuenta2(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0, _
        Optional bMuestraExtornos As Boolean = True, Optional cUser As String, Optional ByRef rstemp As ADODB.Recordset) As ADODB.Recordset

Dim sSql As String, sNumMov As String, sMuestraExtornos As String
Dim rsCta As ADODB.Recordset

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

Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText



Set rsCta.ActiveConnection = Nothing
Set GetMovimientosCuenta2 = rsCta

 sSql = sSql & " Select TOP 1 dFecha, " _
        & "nSaldoDisponible, nSaldoContable from #TMP" & cUser _
        & " WHERE nConceptoCod=" & gConcCapital & " AND left(cMovNroNue,8) <= '" & Format$(dFecFin, "yyyymmdd") & "' " _
        & "AND nMovFlag =" & gMovFlagVigente & " ORDER BY nMovNro asc  "
    
    Set rstemp = New ADODB.Recordset
    rstemp.CursorLocation = adUseClient
    rstemp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText



    Set rstemp.ActiveConnection = Nothing
        
  sSql = " drop table #TMP" & cUser
 
  dbCmact.Execute sSql
    
 
 
Set rsCta = Nothing


End Function

Public Function GetMovimientosCuenta4(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0, _
        Optional bMuestraExtornos As Boolean = True, Optional cUser As String, Optional ByRef rstemp As ADODB.Recordset) As ADODB.Recordset

Dim sSql As String, sNumMov As String, sMuestraExtornos As String
Dim rsCta As ADODB.Recordset

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


Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText



Set rsCta.ActiveConnection = Nothing
Set GetMovimientosCuenta4 = rsCta

 sSql = sSql & " Select TOP 1 dFecha, " _
        & "nSaldoDisponible, nSaldoContable from #TMP" & cUser _
        & " WHERE nConceptoCod=" & gConcCapital & " AND left(cMovNroNue,8) <= '" & Format$(dFecFin, "yyyymmdd") & "' " _
        & "AND nMovFlag =" & gMovFlagVigente & " ORDER BY nMovNro asc  "
    
    Set rstemp = New ADODB.Recordset
    rstemp.CursorLocation = adUseClient
    rstemp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set rstemp.ActiveConnection = Nothing
        
  sSql = " drop table #TMP" & cUser
 
  dbCmact.Execute sSql
    
 
 
Set rsCta = Nothing


End Function
'*** BRGO 20111228 *****************************************************
Public Function GetMovimientosAumCapitalPF(ByVal sCuenta As String) As ADODB.Recordset
    Dim rsCta As ADODB.Recordset
    Dim sSql As String
    
    sSql = "Exec ObtieneMovimientosAumCapitalPF '" & sCuenta & "'"
        Set rsCta = New ADODB.Recordset
        rsCta.CursorLocation = adUseClient
        rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
        Set rsCta.ActiveConnection = Nothing
        Set GetMovimientosAumCapitalPF = rsCta
    Set rsCta = Nothing
End Function
'*** END BRGO ***********************************************************

Public Function GetSaldoFecha2(ByVal cUser As String, ByVal dFecFin As Date) As ADODB.Recordset
Dim rsCta As ADODB.Recordset
Dim sSql As String

sSql = sSql & " Select TOP 1 dFecha, " _
        & "nSaldoDisponible, nSaldoContable from #TMP" & cUser _
        & " WHERE nConceptoCod=" & gConcCapital & " AND left(cMovNroNue,8) <= '" & Format$(dFecFin, "yyyymmdd") & "' " _
        & "AND nMovFlag =" & gMovFlagVigente & " ORDER BY nMovNro DESC"
    
    Set rsCta = New ADODB.Recordset
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

Public Function GetTarifaOrdenPago(ByVal nMoneda As Moneda) As ADODB.Recordset
Dim sSql As String
Dim rsOP As ADODB.Recordset
'sSql = "Select nNumOP, nCosto FROM CapOrdPagTarifa Where nMoneda = " & nMoneda & " " _
'   & "Order by nNumOP"
sSql = "exec stp_sel_CapOrdPagTarifa " & nMoneda 'Add Gemo20210716
Set rsOP = New ADODB.Recordset
rsOP.CursorLocation = adUseClient
rsOP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsOP.ActiveConnection = Nothing
Set GetTarifaOrdenPago = rsOP
Set rsOP = Nothing
End Function



Public Function GetObtenerMontoCheque(ByVal psDocNro As String) As Double

Dim sSql As String
Dim sCtaCod As String
Dim R As ADODB.Recordset
   
    sSql = ""
    sSql = sSql & " SELECT ISNULL(sum(nMonto),0) Total FROM Mov M"
    sSql = sSql & " JOIN MovCap MC on MC.nMovNro=M.nMovNro"
    sSql = sSql & " Join MovDoc MD on MD.nMovNro=MC.nMovNro"
    sSql = sSql & " where cDocNro='" & psDocNro & "'  and nMovFlag=0 and MC.cOpeCod='200202'"
    Set R = New ADODB.Recordset
    R.CursorLocation = adUseClient
    R.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText


    Set R.ActiveConnection = Nothing
    If Not R.EOF And Not R.BOF Then
        GetObtenerMontoCheque = R!Total
    Else
        GetObtenerMontoCheque = 0
    End If
    
End Function


Public Function GetMaxOrdPagEmitida(ByVal sCuenta As String) As Long
Dim rsNum As ADODB.Recordset
Dim sSql As String

Set rsNum = New ADODB.Recordset
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

Public Function GetHistOrdPagEmision(ByVal sCuenta As String) As ADODB.Recordset
Dim rsHist As ADODB.Recordset
Dim sSql As String
sSql = "Select nInicio, nFin, dFecha = CONVERT(Varchar(12),CONVERT(Datetime,SUBSTRING(cMovNro,1,8)),103), " _
    & "cEstado = K.cConsDescripcion From CapOrdPagEmision O INNER JOIN Constante K ON O.nEstado = K.nConsValor " _
    & "Where O.cCtaCod = '" & sCuenta & "' And K.nConsCod = " & gCapOrdPagTalEstado & " Order by cMovNro"
Set rsHist = New ADODB.Recordset
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
Public Function GetCapOrdPagEmision_Cuenta(ByVal nEstado As CapOrdPagTalEstado, ByVal psCodCta As String) As ADODB.Recordset
Dim rsOrden As ADODB.Recordset
Dim sSql As String

sSql = "Select OP.cCtaCod, OP.nInicio, OP.nFin, dFecha = CONVERT(Varchar(12),CONVERT(Datetime,SUBSTRING(OP.cMovNro,1,8)),103), " _
    & "P.cPersNombre, T.nTipo, OP.nNumTal, OP.cMovNro From CapOrdPagEmision OP INNER JOIN ProductoPersona PC INNER JOIN Persona P " _
    & "ON PC.cPersCod = P.cPersCod ON OP.cCtaCod = PC.cCtaCod INNER JOIN CapOrdPagTarifa T ON " _
    & "CONVERT(INT,SUBSTRING(OP.cCtaCod,9,1)) = T.nMoneda And (OP.nFin - OP.nInicio + 1)/OP.nNumTal = T.nNumOP " _
    & "Where OP.nEstado = " & nEstado & " And PC.nPrdPersRelac = " & gCapRelPersTitular & " And P.cPersCod IN " _
    & "(Select MAX(PC1.cPersCod) From ProductoPersona PC1 Where PC1.cCtaCod = OP.cCtaCod And PC1.nPrdPersRelac = " & gCapRelPersTitular & ") " _
    & " and OP.cCtaCod ='" & psCodCta & "' Order by dFecha"
Set rsOrden = New ADODB.Recordset
rsOrden.CursorLocation = adUseClient
rsOrden.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsOrden.ActiveConnection = Nothing
Set GetCapOrdPagEmision_Cuenta = rsOrden
Set rsOrden = Nothing
End Function
Public Function GetCapOrdPagEmision(ByVal nEstado As CapOrdPagTalEstado) As ADODB.Recordset
Dim rsOrden As ADODB.Recordset
Dim sSql As String

sSql = "Select OP.cCtaCod, OP.nInicio, OP.nFin, dFecha = CONVERT(Varchar(12),CONVERT(Datetime,SUBSTRING(OP.cMovNro,1,8)),103), " _
    & "P.cPersNombre, T.nTipo, OP.nNumTal, OP.cMovNro From CapOrdPagEmision OP INNER JOIN ProductoPersona PC INNER JOIN Persona P " _
    & "ON PC.cPersCod = P.cPersCod ON OP.cCtaCod = PC.cCtaCod INNER JOIN CapOrdPagTarifa T ON " _
    & "CONVERT(INT,SUBSTRING(OP.cCtaCod,9,1)) = T.nMoneda And (OP.nFin - OP.nInicio + 1)/OP.nNumTal = T.nNumOP " _
    & "Where OP.nEstado = " & nEstado & " And PC.nPrdPersRelac = " & gCapRelPersTitular & " And P.cPersCod IN " _
    & "(Select MAX(PC1.cPersCod) From ProductoPersona PC1 Where PC1.cCtaCod = OP.cCtaCod And PC1.nPrdPersRelac = " & gCapRelPersTitular & ") Order by dFecha"
Set rsOrden = New ADODB.Recordset
rsOrden.CursorLocation = adUseClient
rsOrden.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsOrden.ActiveConnection = Nothing
Set GetCapOrdPagEmision = rsOrden
Set rsOrden = Nothing
End Function
        
Public Function GetCuentaAbonoIF(ByVal sPersCod As String, ByVal nMoneda As Moneda) As String
Dim rsCta As ADODB.Recordset
Dim sSql As String
sSql = "Select C.cCtaCod FROM ProductoPersona PP INNER JOIN CuentaIFEspecial C ON " _
    & " PP.cCtaCod = C.cCtaCod Where PP.cPersCod = '" & sPersCod & "' And C.bAbono = 1 " _
    & " And C.cCtaCod LIKE '________" & Trim(nMoneda) & "%' AND SUBSTRING(C.cCtaCod,6,3)='232' "

Set rsCta = New ADODB.Recordset
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
Dim rsCta As ADODB.Recordset
Dim sSql As String
Set rsCta = New ADODB.Recordset
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

Public Function GetProductoPersonaConsol(ByVal sCuenta As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

sSql = "Select PP.cCtaCod FROM " & sDBPersona & "Persona P INNER JOIN ProductoPersona PP ON " _
    & "P.cPersCod = PP.cPersCod WHERE PP.cCtaCod = '" & sCuenta & "' Order by PP.cCtaCod"
    
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsCta.ActiveConnection = Nothing
Set GetProductoPersonaConsol = rsCta
Set rsCta = Nothing
End Function

Public Function GetChequesOperaciones(Optional sCuenta As String = "", Optional sNumCheque As String = "", Optional ByVal nEstado As Integer = -1) As ADODB.Recordset
Dim sSql As String, sqlaux As String
Dim rsCta As ADODB.Recordset
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
   
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsCta.ActiveConnection = Nothing
Set GetChequesOperaciones = rsCta
Set rsCta = Nothing
End Function


Public Function GetChequesRegistrados(ByVal sNumCheque As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset
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
    & "And D.nEstado IN (" & gChqEstRegistrado & "," & gChqEstEnValorizacion & " )Order by dFecReg"
   
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsCta.ActiveConnection = Nothing
Set GetChequesRegistrados = rsCta
Set rsCta = Nothing
End Function

Public Function GetChequesValorizadosInmediato(ByVal dFecha As Date, Optional sCuenta As String = "", _
            Optional sNumCheque As String = "") As ADODB.Recordset

Dim sSql As String
Dim rsCta As ADODB.Recordset

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
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsCta.ActiveConnection = Nothing
Set GetChequesValorizadosInmediato = rsCta
Set rsCta = Nothing
End Function

Public Function GetChequesEnValorizacion(ByVal sCuenta As String) As ADODB.Recordset

Dim sSql As String
Dim rsCta As ADODB.Recordset

    sSql = " Select Count(*) Num,  IsNull(Sum(nMonto),0) Monto from DocRecCapta DRC" _
         & " Inner Join DocRecEst DRE On DRC.nTpoDoc = DRE.nTpoDoc And DRC.cNroDoc = DRE.cNroDoc And DRC.cPersCod = DRE.cPersCod" _
         & " Where DRC.cCtaCod = '" & sCuenta & "' And DRE.cMovNro = (Select Max(cMovNro) Mov From DocRecEst DREAdd Where DREAdd.nTpoDoc = DRE.nTpoDoc And DREAdd.cNroDoc = DRE.cNroDoc And DREAdd.cPersCod = DRE.cPersCod)" _
         & " And DRE.nEstado = " & gChqEstEnValorizacion

Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsCta.ActiveConnection = Nothing
Set GetChequesEnValorizacion = rsCta
Set rsCta = Nothing
End Function

Public Function GetCapEstadMovTipo(ByVal sCondicion As String, ByVal sFecha As String) As ADODB.Recordset
Dim sSql As String
Dim rsMov As ADODB.Recordset
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

Set rsMov = New ADODB.Recordset
rsMov.CursorLocation = adUseClient
'ALPA 20080723*******************************************************
dbCmact.CommandTimeout = 0
'********************************************************************
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsMov.ActiveConnection = Nothing
Set GetCapEstadMovTipo = rsMov
Set rsMov = Nothing
End Function

Public Function GetCapEstadSaldoProd(ByVal sCondicion As String)
Dim sSql As String
Dim rsMov As ADODB.Recordset
'Calculamos El Saldo por Producto
sSql = "Select nSaldo = ISNULL(SUM(nSaldo),0) From Producto Where nPrdEstado NOT IN (1300,1400) " _
    & "And cCtaCod LIKE '" & sCondicion & "'"

Set rsMov = New ADODB.Recordset
rsMov.CursorLocation = adUseClient
'ALPA 20080723*******************************************************
dbCmact.CommandTimeout = 0
'********************************************************************
rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsMov.ActiveConnection = Nothing
Set GetCapEstadSaldoProd = rsMov
Set rsMov = Nothing
End Function

Public Function GetCapEstadMovCheques(ByVal sFecha As String, Optional sCondicion As String)
Dim sSql As String
Dim rsMov As ADODB.Recordset
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

Set rsMov = New ADODB.Recordset
rsMov.CursorLocation = adUseClient
'ALPA 20080723*******************************************************
dbCmact.CommandTimeout = 0
'********************************************************************
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

Set rsItf = New ADODB.Recordset
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
'-----------------------Comentado por AVMM (hace muy lenta la consulta)-----------------------
'sSql = sSql & "     left join persona pe on pe.cperscod =cta.codpers"
'sSql = sSql & "     left join persid bb on bb.cperscod=cta.codpers    and bb.cPersIDTpo = 1"
'sSql = sSql & "     left join persid cc on cc.cperscod=cta.codpers    and cc.cPersIDTpo = 2"
'sSql = sSql & "     left join persid dd on dd.cperscod=cta.codpers    and dd.cPersIDTpo = 4"
'sSql = sSql & "     left join persid ee on ee.cperscod=cta.codpers    and ee.cPersIDTpo = 11 "
'sSql = sSql & "     left join persid ff on ff.cperscod=cta.codpers    and ff.cPersIDTpo = 10 "
'---------------------------------------------------------------------------------------------

'ssql = ssql & " where  isnull(bb.cPersIDTpo, isnull(cc.cPersIDTpo,isnull(  dd.cPersIDTpo, ff.cPersIDTpo))) in (1,2,4,10)"
'ssql = ssql & " where  isnull(bb.cPersIDTpo, isnull( cc.cPersIDTpo,isnull(dd.cPersIDTpo, isnull(ee.cPersIDTpo,ff.cPersIDTpo))))  in (1,2,4,10,11) "
sSql = sSql & " group by cta.COPECOD, cfecha,Substring(cta.cuenta,4,2), COPEDESC, nPrdConceptoCod, cDescripcion"
sSql = sSql & " order by Substring(cta.cuenta,4,2), nPrdConceptocod, cFecha, cta.copecod"

Set rsItf = New ADODB.Recordset
rsItf.CursorLocation = adUseClient
dbCmact.CommandTimeout = 8000
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

Set rsItf = New ADODB.Recordset
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
sSql = sSql & " E.nExoTpo, CodOpe = CASE WHEN (E.nExoTpo = 2 or E.nExoTpo = 1 or E.nExoTpo = 14 or E.nExoTpo = 22 or E.nExoTpo = 23) then 'b' "
sSql = sSql & " WHEN E.nExoTpo = 15 THEN 'o' WHEN E.nExoTpo = 3 then 'c' "
sSql = sSql & " WHEN E.nExoTpo = 4 THEN 'd' WHEN substring(PP.cCtaCod, 6, 3) = '423' then 'w' WHEN E.nExoTpo = 6 THEN 'm'  End  , TipoOper = '1'       From "
sSql = sSql & " ITFExoneracionCta E JOIN ProductoPersona PP JOIN (Select P.cPersCod, P.cPersNombre, P.nPersPersoneria, cPersIDTpo = ISNULL(ID.cPersIDTpo,''), "
sSql = sSql & " cPersIDnro = ISNULL(ID.cPersIDnro,'') From Persona P LEFT JOIN PersID ID ON P.cPersCod = ID.cPersCod  where  (cPersIDnro<>'' AND cPersIDnro IS NOT NULL )  ) P "
sSql = sSql & " ON PP.cPersCod = P.cPersCod ON E.cCtaCod = PP.cCtaCod Where PP.nPrdPersRelac IN (" & gCapRelPersTitular & "," & gColRelPersTitular & ") "
sSql = sSql & " And E.nExoTpo not in (0,21,24) AND E.CREGISTRO <='" & sFin & "'    "

Set rsItf = New ADODB.Recordset
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

Public Sub InsertaOrdenIngOp(ByVal pscctacod As String, ByVal psFecha As String, ByVal psIni As String, ByVal psFin As String, psNro As Integer)
    Dim sSql As String

     sSql = "insert into capordpagemision "
     sSql = sSql & " values('" & pscctacod & "','" & Format(psFecha, "yyyymmdd") & "'+'00000010801009999',4," & CLng(psIni) & "," & CLng(psFin) & "," & Val(psNro) & ",1  ) "
     
     dbCmact.Execute (sSql)
     
     sSql = "insert into capordpagestado "
     sSql = sSql & "values('" & pscctacod & "','" & Format(psFecha, "yyyymmdd") & "'+'00000010801009999',4," & CLng(psIni) & " ) "
     
     dbCmact.Execute (sSql)
End Sub

Public Function ObtenerOrdenIngOp(ByVal psCtaCod As String, ByVal pnIni As String, ByVal pnFin As String) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    sSql = "Select count(*) from capordpagemision where cctacod='" & psCtaCod & "' and ninicio=" & CLng(pnIni) & " and nfin=" & CLng(pnFin) & " and bantiguo=1 "
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerOrdenIngOp = rs
    Else
        Set ObtenerOrdenIngOp = Nothing
    End If

End Function



Public Function ObtenerTotalOperacones() As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    
    sSql = "Select c.copecod,Upper(o.operacion),rct.crhcargodescripcion + space(150-len(rct.crhcargodescripcion))+rct.crhcargocod Cargo,c.cnivel,c.nMontoIniSol,c.nMontoFinSol,c.nMontoIniDol,nMontoFinDol,'E' as Estado,cast(case when c.bactivo=1 then 1 else 0 end as char(1)) as bactivo "
    sSql = sSql & " from capautorizacionrango c "
    sSql = sSql & " inner join  rhcargostabla  rct on rct.crhcargocod=c.crhcargocod "
    sSql = sSql & " inner join  (Select ch.copecod,f.copedesc +':'+ ch.copedesc as operacion "
    sSql = sSql & " from opetpo ch "
    sSql = sSql & " inner join opetpo f on f.copecod=case when left(ch.copecod,1)='2' then  left(ch.copecod,2)+'0000' when  left(ch.copecod,2)='30' then  left(ch.copecod,2)+'0000'  when left(ch.copecod,2)='31' then  left(ch.copecod,2)+'0000'      when  left(ch.copecod,5)='90002' then '900020'  when left(ch.copecod,5)='90100' then  '901001' when left(ch.copecod,5)='90101' then  '901011' when  left(ch.copecod,5)='90003'  then   '900030'   end "
    sSql = sSql & " where ch.copecod like  '2%[1-9]00' or ch.copecod like  '3[0-1]%00' or ch.copecod like '90002[1-6]' or ch.copecod like  '9010[01][2-9]' or ch.copecod like '90003[1-5]') o on o.copecod=c.copecod "
    sSql = sSql & "   Where Len(rct.cRHCargoCod) > 3 "
    sSql = sSql & " order by c.copecod "
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerTotalOperacones = rs
        Set rs.ActiveConnection = Nothing
    Else
        Set ObtenerTotalOperacones = Nothing
    End If
    
End Function

Public Function ObtenerListaCuentas(ByVal psUsuario As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    sSql = " Select o.operacion + space(150-len(o.operacion))+ m.copecodori as Opera,m.cctacod as Cuenta,p.cpersnombre + space(150-len(p.cpersnombre))+m.cperscodcli as persona,m.nmontosolicitado,m.nmontoaprobado,(case when m.nmoneda=1 then 'S/.' else 'US$' end) + space(147)+ cast(m.nmoneda as char(1)) as moneda ,m.cautobs,m.cautestado,m.nidaut,c.nmontofinsol,c.nmontofindol "
    sSql = sSql & " from capautorizacionope m "
    sSql = sSql & "  inner join (Select ch.copecod,f.copedesc +':'+ ch.copedesc as operacion "
    sSql = sSql & " from opetpo ch "
    sSql = sSql & " inner join opetpo f on f.copecod=case when left(ch.copecod,1)='2' then  left(ch.copecod,2)+'0000' when  left(ch.copecod,2)='30' then  left(ch.copecod,2)+'0000'  when left(ch.copecod,2)='31' then  left(ch.copecod,2)+'0000'      when  left(ch.copecod,5)='90002' then '900020'  when left(ch.copecod,5)='90100' then  '901001' when left(ch.copecod,5)='90101' then  '901011' when  left(ch.copecod,5)='90003'  then   '900030'   end "
    sSql = sSql & " where ch.copecod like  '2%[1-9]00' or ch.copecod like  '3[0-1]%00' or ch.copecod like '90002[1-6]' or ch.copecod like  '9010[01][2-9]' or ch.copecod like '90003[1-5]' ) o on o.copecod=m.copecodori "
    sSql = sSql & " inner join persona p on p.cperscod=m.cperscodcli "
    sSql = sSql & "  inner join (Select nconsvalor,cconsdescripcion from constante where nconscod='1011') as tm on tm.nconsvalor=m.nmoneda "
    sSql = sSql & "  inner join capautorizacionrango c on c.copecod=m.copecodori "
    sSql = sSql & "  inner join  (Select top 1 rhc1.crhcargocod,rhc1.cperscod from rhcargos rhc1 inner join rrhh rh1 on rh1.cperscod=rhc1.cperscod where rh1.cuser='" & psUsuario & "' and  rh1.nrhestado='201' "
    sSql = sSql & "  order by drhcargoFecha desc) rhc on rhc.crhcargocod=c.crhcargocod "
    sSql = sSql & "  where left(m.cultimaactualizacion,8)=convert(char(8),getdate(),112) and m.cautestado='P' and (case when nmoneda=1 then c.nMontoFinsol else c.nMontoFinDol end)>=m.nmontosolicitado "
    sSql = sSql & " order by m.nidaut  "
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerListaCuentas = rs
        Set rs.ActiveConnection = Nothing
    Else
        Set ObtenerListaCuentas = Nothing
    End If
End Function

'Modificado por RIRO el 20130411
Public Function ObtenerListaSolicitudes(ByVal psCodPers As String, ByVal psTipo As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    If psTipo = "0" Then
        
        sSql = " SELECT [Nro Solicitud]=nnumsolicitud,k.cConsDescripcion,Moneda=CASE WHEN NMONEDA=1 THEN 'SOLES' ELSE 'DOLARES' END, "
        sSql = sSql & " Plazo=nPlazo,Tasa=dbo.ConvierteTNAaTEA(ntasa),Monto=nmonto,Estado=case when nestado=0 then 'SOLICITADO' when nestado=1 then 'APROBADO' END   ,nmoneda,nestado,nproducto,Persona=p.cpersnombre,p.cperscod,bpermanente=isnull(bpermanente,0), usuario=RIGHT(c.cMovNro, 4), c.cSubProducto, c.nTasaTarif, c.nTasaSolicitada " 'MODIFICADO POR "RIRO" 23/11/2012
        sSql = sSql & " FROM  CAPTASAESPECIAL c "
        sSql = sSql & " JOIN Persona P ON p.cPersCod = c.cPersCod left join Constante k on k.nConsCod = LEFT(c.cSubProducto, 4) and k.nConsValor = right(c.cSubProducto,1)" 'MODIFICADO POR "RIRO" 23/11/2012
        sSql = sSql & " Where c.cperscod='" & psCodPers & "' and "
        sSql = sSql & " left(c.cmovnro,14)+cast(c.nnumsolicitud as varchar(12))  in "
        sSql = sSql & " ( select max(left(c.cmovnro,14))+cast(c.nnumsolicitud as varchar(12)) from CAPTASAESPECIAL c "
        sSql = sSql & "   group by c.nnumsolicitud  ) "
        sSql = sSql & "  and (nestado=0 or (nestado=3 and nextorno=1) or (nestado=3 and nextorno=4) ) "
        sSql = sSql & "  order by nnumsolicitud desc,left(cmovnro,14) desc "

'        sSql = " SELECT [Nro Solicitud]=nnumsolicitud,Producto=case when nproducto=233 then 'AHORROS' ELSE 'PLAZO FIJO' END,Moneda=CASE WHEN NMONEDA=1 THEN 'SOLES' ELSE 'DOLARES' END, "
'        sSql = sSql & " Plazo=nPlazo,Tasa=dbo.ConvierteTNAaTEA(ntasa),Monto=nmonto,Estado=case when nestado=0 then 'SOLICITADO' when nestado=1 then 'APROBADO' END   ,nmoneda,nestado,nproducto,Persona=p.cpersnombre,p.cperscod,bpermanente=isnull(bpermanente,0) "
'        sSql = sSql & " FROM  CAPTASAESPECIAL c "
'        sSql = sSql & " JOIN Persona P ON p.cPersCod = c.cPersCod "
'        sSql = sSql & " Where c.cperscod='" & psCodPers & "' and "
'        sSql = sSql & " left(c.cmovnro,14)+cast(c.nnumsolicitud as varchar(12))  in "
'        sSql = sSql & " ( select max(left(c.cmovnro,14))+cast(c.nnumsolicitud as varchar(12)) from CAPTASAESPECIAL c "
'        sSql = sSql & "   group by c.nnumsolicitud  ) "
'        sSql = sSql & "  and (nestado=0 or (nestado=3 and nextorno=1) or (nestado=3 and nextorno=4) ) "
'        sSql = sSql & "  order by nnumsolicitud desc,left(cmovnro,14) desc "
    
    ElseIf psTipo = "1" Then

        sSql = " SELECT [Nro Solicitud]=nnumsolicitud,k.cConsDescripcion,Moneda=CASE WHEN NMONEDA=1 THEN 'SOLES' ELSE 'DOLARES' END, "
        sSql = sSql & " Plazo=nPlazo,Tasa=case when (nestado=3 and nextorno=2) then ntasa else  dbo.ConvierteTNAaTEA(ntasa) end ,Monto=nmonto,Estado=case when nestado=0 then 'SOLICITADO' when nestado=1 then 'APROBADO' END   ,nmoneda,nestado,nproducto,Persona=p.cpersnombre,p.cperscod,bpermanente=isnull(bpermanente,0), usuario=RIGHT(c.cMovNro, 4), c.cSubProducto, c.nTasaTarif, nTasaSolicitada" 'MODIFICADO POR "RIRO" 23/11/2012
        sSql = sSql & " FROM  CAPTASAESPECIAL c "
        sSql = sSql & " JOIN Persona P ON p.cPersCod = c.cPersCod left join Constante k on k.nConsCod = LEFT(c.cSubProducto, 4) and k.nConsValor = right(c.cSubProducto,1)" 'MODIFICADO POR "RIRO" 23/11/2012
        sSql = sSql & " Where c.cperscod='" & psCodPers & "' and "
        sSql = sSql & " left(c.cmovnro,14)+cast(c.nnumsolicitud as varchar(12))  in "
        sSql = sSql & " ( select max(left(c.cmovnro,14))+cast(c.nnumsolicitud as varchar(12)) from CAPTASAESPECIAL c "
        sSql = sSql & " group by c.nnumsolicitud  ) "
        sSql = sSql & " and (nestado=1 or (nestado=3 and nextorno=2) ) "
        sSql = sSql & "  order by nnumsolicitud desc,left(cmovnro,14) desc "
        
'        sSql = " SELECT [Nro Solicitud]=nnumsolicitud,Producto=case when nproducto=233 then 'AHORROS' ELSE 'PLAZO FIJO' END,Moneda=CASE WHEN NMONEDA=1 THEN 'SOLES' ELSE 'DOLARES' END, "
'        sSql = sSql & " Plazo=nPlazo,Tasa=case when (nestado=3 and nextorno=2) then ntasa else  dbo.ConvierteTNAaTEA(ntasa) end ,Monto=nmonto,Estado=case when nestado=0 then 'SOLICITADO' when nestado=1 then 'APROBADO' END   ,nmoneda,nestado,nproducto,Persona=p.cpersnombre,p.cperscod,bpermanente=isnull(bpermanente,0) "
'        sSql = sSql & " FROM  CAPTASAESPECIAL c "
'        sSql = sSql & " JOIN Persona P ON p.cPersCod = c.cPersCod "
'        sSql = sSql & " Where c.cperscod='" & psCodPers & "' and "
'        sSql = sSql & " left(c.cmovnro,14)+cast(c.nnumsolicitud as varchar(12))  in "
'        sSql = sSql & " ( select max(left(c.cmovnro,14))+cast(c.nnumsolicitud as varchar(12)) from CAPTASAESPECIAL c "
'        sSql = sSql & " group by c.nnumsolicitud  ) "
'        sSql = sSql & " and (nestado=1 or (nestado=3 and nextorno=2) ) "
'        sSql = sSql & "  order by nnumsolicitud desc,left(cmovnro,14) desc "
        
    ElseIf psTipo = "2" Then
    
        sSql = " SELECT [Nro Solicitud]=nnumsolicitud,k.cConsDescripcion,Moneda=CASE WHEN NMONEDA=1 THEN 'SOLES' ELSE 'DOLARES' END, "
        sSql = sSql & " Plazo=nPlazo,Tasa= dbo.ConvierteTNAaTEA(ntasa) ,Monto=nmonto,Estado=case when nestado=0 then 'SOLICITADO' when nestado=1 then 'APROBADO' END   ,nmoneda,nestado,nproducto,Persona=p.cpersnombre,p.cperscod,bpermanente=isnull(bpermanente,0), usuario=RIGHT(c.cMovNro, 4), c.cSubProducto, c.nTasaTarif, nTasaSolicitada" 'MODIFICADO POR "RIRO" EL 23/11/2012
        sSql = sSql & " FROM  CAPTASAESPECIAL c "
        sSql = sSql & " JOIN Persona P ON p.cPersCod = c.cPersCod left join Constante k on k.nConsCod = LEFT(c.cSubProducto, 4) and k.nConsValor = right(c.cSubProducto,1)" 'MODIFICADO POR "RIRO" EL 23/11/2012
        sSql = sSql & " Where c.cperscod='" & psCodPers & "' and "
        sSql = sSql & " left(c.cmovnro,14)+cast(c.nnumsolicitud as varchar(12))  in "
        sSql = sSql & " ( select max(left(c.cmovnro,14))+cast(c.nnumsolicitud as varchar(12)) from CAPTASAESPECIAL c "
        sSql = sSql & " group by c.nnumsolicitud  ) "
        sSql = sSql & " and (nestado=4 ) "
        sSql = sSql & "  order by nnumsolicitud desc,left(cmovnro,14) desc "
    
'        sSql = " SELECT [Nro Solicitud]=nnumsolicitud,Producto=case when nproducto=233 then 'AHORROS' ELSE 'PLAZO FIJO' END,Moneda=CASE WHEN NMONEDA=1 THEN 'SOLES' ELSE 'DOLARES' END, "
'        sSql = sSql & " Plazo=nPlazo,Tasa= dbo.ConvierteTNAaTEA(ntasa) ,Monto=nmonto,Estado=case when nestado=0 then 'SOLICITADO' when nestado=1 then 'APROBADO' END   ,nmoneda,nestado,nproducto,Persona=p.cpersnombre,p.cperscod,bpermanente=isnull(bpermanente,0) "
'        sSql = sSql & " FROM  CAPTASAESPECIAL c "
'        sSql = sSql & " JOIN Persona P ON p.cPersCod = c.cPersCod "
'        sSql = sSql & " Where c.cperscod='" & psCodPers & "' and "
'        sSql = sSql & " left(c.cmovnro,14)+cast(c.nnumsolicitud as varchar(12))  in "
'        sSql = sSql & " ( select max(left(c.cmovnro,14))+cast(c.nnumsolicitud as varchar(12)) from CAPTASAESPECIAL c "
'        sSql = sSql & " group by c.nnumsolicitud  ) "
'        sSql = sSql & " and (nestado=4 ) "
'        sSql = sSql & "  order by nnumsolicitud desc,left(cmovnro,14) desc "
    
    End If
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerListaSolicitudes = rs
        Set rs.ActiveConnection = Nothing
    Else
        Set ObtenerListaSolicitudes = Nothing
    End If

End Function

Public Sub InsertarITFExoneracion(ByVal psNroCuenta As String, ByVal psvalor As String, ByVal psVCMovNro As String)
 Dim sSql As String
    sSql = "insert into ITFEXONERACIONCTA(cCtaCod,nExoTpo,cRegistro) "
    sSql = sSql & " values('" & psNroCuenta & "','" & psvalor & "','" & psVCMovNro & "') "

    dbCmact.Execute sSql
End Sub

Public Sub ActualizarITFExoneracion(ByVal psNroCuenta As String, ByVal psvalor As String, ByVal psVCMovNro As String, ByVal psAux As String)
 Dim sSql As String
    sSql = "Update ITFExoneracionCta set nExoTpo=" & psvalor & ", cRegistro='" & psVCMovNro & "' " & psAux & " where cctacod='" & psNroCuenta & "'"

    dbCmact.Execute sSql
End Sub

Public Sub InsertarITFExoneracionDET(ByVal psNroCuenta As String, ByVal psvalor As String, ByVal psVCMovNro As String)
 Dim sSql As String
   sSql = "insert into ITFEXONERACIONCTADET(cMovNro,cCtaCod,nExoTpo) "
   sSql = sSql & " VALUES('" & psVCMovNro & "','" & psNroCuenta & "'," & psvalor & ") "

dbCmact.Execute sSql
End Sub

Public Function TieneExoneracion(ByVal sCuenta As String) As Integer
 Dim sSql As String
 Dim rstemp As ADODB.Recordset
 sSql = "SELECT count(*) FROM ITFExoneracionCta where cctacod='" & sCuenta & "'"
 Set rstemp = New ADODB.Recordset
 rstemp.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

 If Not (rstemp.EOF And rstemp.BOF) Then
   If rstemp.State = 1 Then
     If rstemp.RecordCount > 0 Then
         TieneExoneracion = rstemp.Fields(0).Value
     End If
   End If
 End If
Set rstemp = Nothing
End Function

Public Function ObtenerHistorialExonearcion(ByVal psCuenta As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    sSql = "SELECT CMovNro=left(i.CmovNro,8),i.CCtacod,i.NExoTpo,c.cconsdescripcion FROM ITFExoneracionCtaDet I "
    sSql = sSql & " inner join (SELECT nconsvalor,cconsdescripcion FROM CONSTANTE WHERE NCONSCOD=1044) c on c.nconsvalor=i.NExoTpo "
    sSql = sSql & " where cCtacod='" & psCuenta & "'"
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF And Not rs.BOF Then
       Set ObtenerHistorialExonearcion = rs
       Set rs.ActiveConnection = Nothing
    Else
       Set ObtenerHistorialExonearcion = Nothing
    End If
    
End Function

Public Function ObtenerDctoInactivasFecha() As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    sSql = "select nDia,nMes from dbo.DctoInactivasFechas"
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF And Not rs.BOF Then
       Set ObtenerDctoInactivasFecha = rs
       Set rs.ActiveConnection = Nothing
    Else
       Set ObtenerDctoInactivasFecha = Nothing
    End If
End Function

Public Function BloqueoDesPlazoFijo(ByVal psCtaCod As String, ByVal pnTipoBloqueo As Integer)
    Dim sql As String
    Dim rs As New ADODB.Recordset
     
    sql = " SELECT count (*) NroBloqueo FROM CaptaPlazoFijoBloqueo " & _
          " WHERE cCtaCod='" & psCtaCod & "'"
    rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF And rs.BOF) Then
        If rs!NroBloqueo = 0 Then
            sql = " INSERT CaptaPlazoFijoBloqueo(cCtaCod,nBloqueo)" & _
                  " VALUES ('" & psCtaCod & "', " & pnTipoBloqueo & ")"
        Else
            sql = " UPDATE CaptaPlazoFijoBloqueo SET nBloqueo= " & pnTipoBloqueo & " " & _
                  " WHERE cCtaCod='" & psCtaCod & "'"
        End If
        dbCmact.Execute sql
    End If
    
End Function

Public Function VerificaBloqueoPlazoFijo(ByVal psCtaCod As String) As Boolean
    Dim sql As String
    Dim rs As New ADODB.Recordset
     
    sql = " SELECT count (*) NroBloqueo FROM CaptaPlazoFijoBloqueo " & _
          " WHERE cCtaCod='" & psCtaCod & "' and nBloqueo =1 "
    rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText


    If Not (rs.EOF And rs.BOF) Then
        If rs!NroBloqueo > 0 Then
            VerificaBloqueoPlazoFijo = True
        Else
            VerificaBloqueoPlazoFijo = False
        End If
    End If
    
End Function

Public Function ObtenerNroMovimientosRet(ByVal psCtaCod As String, ByVal psFecha As String) As Integer
    Dim sql As String
    Dim rs As New ADODB.Recordset
     
    sql = " select count(MC.nMovNro) NroMovRet from Mov M " & _
          " Join MovCap MC  on M.nMovNro=MC.nMovNro " & _
          " where MC.cCtaCod = '" & psCtaCod & "' and   MC.copecod ='200301' and cMovNro like '" & psFecha & "%" & "' "
    rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF And rs.BOF) Then
       ObtenerNroMovimientosRet = rs!NroMovRet
    Else
       ObtenerNroMovimientosRet = 0
    End If
    
End Function
'By Capi 05032008
Public Function ObtenerSiEsCuentaRRHH(ByVal psCtaCod As String) As Boolean
    Dim lsql As String
    Dim lrs As New ADODB.Recordset
    lsql = " Select cCtaCod from RhCuentas " & _
          " where cCtaCod = '" & psCtaCod & "'"
    lrs.Open lsql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText


    If Not (lrs.EOF And lrs.BOF) Then
       ObtenerSiEsCuentaRRHH = True
    Else
       ObtenerSiEsCuentaRRHH = False
    End If
End Function


Public Function GetPromotores(ByVal psCodAge As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim oGen As COMDConstSistema.DCOMGeneral
    Dim sPromotor As String
    
    Set oGen = New COMDConstSistema.DCOMGeneral
    sPromotor = oGen.LeeConstSistema(gConstSistRHCargoCodPromotores)
    rs.CursorLocation = adUseClient
    
    sSql = "Select R.cUser, R.cPersCod, P.cPersNombre from RRHH R inner join Persona P ON R.cPersCod = P.cpersCod "
    sSql = sSql & " AND nRHEstado in (201,301) "
    'By Capi Agosto-2007(Acta 014-2007/TI-D)
    'sSQl = sSQl & " where  RC.cRHCargoCod in (" & sPromotor & ") AND RC.dRHCargoFecha = (select MAX(dRHCargoFecha) from RHCargos RHC2 where RHC2.cPersCod = RC.cPersCod) "
    sSql = sSql & " Where R.cAgenciaActual='" & psCodAge & "'"
    'End By Capi
    sSql = sSql & " order by P.cPersNombre "
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    Set GetPromotores = rs
    Set rs.ActiveConnection = Nothing
    Exit Function
End Function

'Add Gitu 23-09-2009
'Recupera la cartera de ahorros por gestor
Public Function CarteraGestor(ByVal psCodGestor As String, ByVal psCodAge As String, ByVal psCtasPF As String, ByVal psCtasCTS As String) As ADODB.Recordset
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim lsCondProd As String
    
    If psCtasPF <> "" And psCtasCTS <> "" Then
        lsCondProd = "In ('233','234')"
    ElseIf psCtasPF <> "" And psCtasCTS = "" Then
        lsCondProd = "= '233'"
    ElseIf psCtasPF = "" And psCtasCTS <> "" Then
        lsCondProd = "= '234'"
    Else
        lsCondProd = "= ''"
    End If
    
    rs.CursorLocation = adUseClient
    
    If Trim(psCodGestor) <> "00000000" Then
        sSql = "Select PE.cPersNombre, Count(P.cCtaCod) NumCtas, PP.cPerscod " _
             & "From(Select PP1.cCtaCod, PP1.cPerscod" _
             & "     From ProductoPersona PP1" _
             & "         Join Producto P1 On PP1.cCtaCod = P1.cCtaCod and P1.nPrdEstado = 1000" _
             & "     Where PP1.cPersCod = '" & psCodGestor & "' And PP1.nPrdPersRelac = 14" _
             & "         and Substring(PP1.cCtacod,6,3) " & lsCondProd _
             & "         and Substring(PP1.cCtacod,4,2) = '" & psCodAge & "') P" _
             & "    Join ProductoPersona PP On P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 10" _
             & "    Join Persona PE On PP.cPerscod = PE.cPerscod " _
             & "Group by PE.cPersNombre, PP.cPerscod " _
             & "Order by PE.cPersNombre"
    Else
        sSql = "Select PE.cperscod, PE.cPersNombre, P.cCtaCod NumCtas " _
             & "From Producto P " _
             & "    join ProductoPersona PP On P.cCtaCod = PP.cCtaCod and PP.nPrdPersRelac = 10 " _
             & "    join Persona PE On PP.cPerscod = PE.cPerscod " _
             & "Where P.nPrdEstado =  1000 And SubString(P.cCtaCod,6,3)" & lsCondProd _
             & "    And Substring(P.cCtacod,4,2) = '" & psCodAge & "' " _
             & "    And P.cCtaCod Not In (Select cctacod  from ProductoPersona where nPrdPersRelac = 14) " _
             & "order by PE.cPersNombre, PE.cperscod"
    End If
    
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set CarteraGestor = rs
    Set rs.ActiveConnection = Nothing

End Function

Public Function ReasignaCarteraGestor(ByVal pMatCuentas As Variant, _
                                        ByVal psPerscodActual As String, _
                                        ByVal psPersCodAnterior As String, _
                                        ByVal pdFecSis As Date, _
                                        ByVal psTipoPF As String, _
                                        ByVal psTipoCTS As String, _
                                        ByVal psCodAge As String)
    
Dim i As Integer

On Error GoTo ErrorReasignaCarteraGestor

For i = 1 To UBound(pMatCuentas)
    Call ActualizaCuentasGestor(pMatCuentas(i), psPerscodActual, Trim(psPersCodAnterior), pdFecSis, psTipoPF, psTipoCTS, psCodAge)
Next i

Exit Function

ErrorReasignaCarteraGestor:
    Err.Raise Err.Number, "Error", Err.Description

End Function
Public Sub ActualizaCuentasGestor(ByVal psPersCodCli As String, ByVal psPerscodActual As String, _
                                    Optional ByVal psPersCodAnterior As String = "", Optional ByVal pdFecSis As Date = 0, _
                                    Optional ByVal psTipPF As String, Optional ByVal psTipCTS As String, _
                                    Optional ByVal psCodAge As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim oFechas As COMDConstSistema.DCOMGeneral
Dim lsCondProd As String

    On Error GoTo ErrorActualizaCuentasGestor
    Set oConecta = New COMConecta.DCOMConecta
    
    If psTipPF <> "" And psTipCTS <> "" Then
        lsCondProd = "In ('233','234')"
    ElseIf psTipPF <> "" And psTipCTS = "" Then
        lsCondProd = "= '233'"
    ElseIf psTipPF = "" And psTipCTS <> "" Then
        lsCondProd = "= '234'"
    Else
        lsCondProd = "= ''"
    End If
    
    If psPersCodAnterior = "00000000" Then
        sSql = "Insert Into ProductoPersona (cCtaCod, cPersCod, nPrdPersRelac)" _
             & "Values('" & psPersCodCli & "','" & psPerscodActual & "'," & Trim(Str(COMDConstantes.gCapRelPersPromotor)) & ")"
    Else
        sSql = "Update ProductoPersona Set cPerscod = '" & psPerscodActual & "' " _
             & "From ProductoPersona P1 " _
             & "Join (Select PP1.cCtaCod, PP1.cPerscod " _
             & "      From ProductoPersona PP1 " _
             & "        Join Producto P1 On PP1.cCtaCod = P1.cCtaCod and P1.nPrdEstado = " & Trim(Str(COMDConstantes.gCapEstActiva)) _
             & "      Where PP1.cPersCod = '" & psPersCodCli & "' And PP1.nPrdPersRelac = " & Trim(Str(COMDConstantes.gCapRelPersTitular)) _
             & "        and Substring(PP1.cCtacod,6,3) " & lsCondProd _
             & "        and Substring(PP1.cCtacod,4,2) = '" & psCodAge & "'" _
             & "      ) B On P1.cCtaCod = B.cCtaCod  And P1.nPrdPersRelac = " & Trim(Str(COMDConstantes.gCapRelPersPromotor))
    End If
    oConecta.AbreConexion
    oConecta.ConexionActiva.Execute sSql
'    If psPersCodAnterior <> "" Then
'        Set oFechas = New COMDConstSistema.DCOMGeneral
'        sSql = "INSERT INTO ColocacReasignaCartera(cCtaCod,dReasignacion,cPersCodAnterior,cPersCodActual)" & _
'               " VALUES ('" & psCtaCod & "','" & Format(oFechas.FechaHora(pdFecSis), "mm/dd/yyyy hh:mm:ss") & "','" & psPersCodAnterior & "','" & psPerscodActual & "')"
'        Set oFechas = Nothing
'        oConecta.ConexionActiva.Execute sSql
'    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorActualizaCuentasGestor:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso ActualizaCreditoAnalista", Err.Description

End Sub

Public Function GetAbonoPactadoSubProductoAho(ByVal psCtaCod As String, ByVal nMonto As Double, ByVal pdFecSis As Date, ByVal pnTpoProgram As Integer, ByRef psMensaje As String)
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim sPromotor As String
    Dim dFecha  As String
    'pnTpoProgram =0 -> Aho Corriente/pnTpoProgram =1 -> Aho Ñañito/pnTpoProgram =2 -> Aho Panderito
    'pnTpoProgram =3 -> Aho Pandero/pnTpoProgram =4 -> Aho Destino
   
    sSql = "Select nMontoAbono from CaptacAhorros where cCtaCod='" & psCtaCod & "'"
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText


    If Not (rs.EOF And rs.BOF) Then
        If nMonto < rs("nMontoAbono") Then
           psMensaje = " El Monto de Abono es menor que el Monto Pactado ... " & rs("nMontoAbono")
        Else
           psMensaje = ""
        End If
    End If
    rs.Close
    
    'JUEZ 20141014 Comentado para nuevos parámetros ***********
    'Set rs = New ADODB.Recordset
    'If pnTpoProgram = 4 Then
    '    dFecha = Format(pdFecSis, "YYYYMMDD") & "%"
    '    sSql = "select count(*) nNroAbonos  from Mov M  JOIN MovCap MC on M.nMovNro = MC.nMovNro " _
    '           & " where cMovNro like '" & dFecha & "' and cCtaCod='" & psCtaCod & "'" _
    '           & " and nMovFlag=0 and  mc.cOpeCod like '2002%' "
    '    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText


    '    If Not (rs.EOF And rs.BOF) Then
    '        If rs("nNroAbonos") > 0 Then
    '            If psMensaje = "" Then
    '                psMensaje = "Solo se permite un solo Abono para el Mes"
    '            Else
    '                psMensaje = psMensaje & " y ya se realizo el Abono para el Mes"
    '            End If
    '        End If
    '    End If
    'End If
    'rs.Close
    'END JUEZ *************************************************
    
    Exit Function
End Function

Public Function GetVisualizaSaldoPosicion(ByVal pCargo As String) As Integer
    
    Dim oGen As COMDConstSistema.DCOMGeneral
    Dim sPosicion As String
    Dim sCod() As String
    Dim i As Integer
    
    Set oGen = New COMDConstSistema.DCOMGeneral
    sPosicion = oGen.LeeConstSistema(gConstSistRHCargoCodPosicionCli)
    
    sCod = Split(sPosicion, ",")
    GetVisualizaSaldoPosicion = 0
    For i = 0 To UBound(sCod)
        If pCargo = sCod(i) Then
            GetVisualizaSaldoPosicion = 1
        End If
    Next i
    
    Exit Function
    
End Function
'Add By GITU 23-08-2011
Public Function GetConsultaSaldoSinTarjeta(ByVal pCargo As String) As Integer
    
    Dim oGen As COMDConstSistema.DCOMGeneral
    Dim sPosicion As String
    Dim sCod() As String
    Dim i As Integer
    
    Set oGen = New COMDConstSistema.DCOMGeneral
    sPosicion = oGen.LeeConstSistema(413)
    
    sCod = Split(sPosicion, ",")
    GetConsultaSaldoSinTarjeta = 0
    For i = 0 To UBound(sCod)
        If pCargo = sCod(i) Then
            GetConsultaSaldoSinTarjeta = 1
        End If
    Next i
    
    Exit Function
    
End Function
'End GITU

'By Capi 07082008 se adiciono parametro opcional bTasaPactada
Public Sub CambiaTasaInteres(ByVal sCuenta As String, ByVal nTasaAnt As Double, ByVal nTasa As Double, ByVal cMovnro As String, Optional ByVal pbTasaPactada = True)

    Dim sSql As String

    '*** Creado por GRVA 11/07/2007 ***
    sSql = "Update Producto Set nTasaInteres = " & nTasa & " Where cCtaCod = '" & sCuenta & "'"
    dbCmact.Execute sSql
    
    'By Capi 15/01/2008
    sSql = "Update CaptacPlazoFijo Set bTasaPactada = " & IIf(pbTasaPactada = True, 1, 0) & " Where cCtaCod = '" & sCuenta & "'"
    dbCmact.Execute sSql
    
    sSql = "Update CapCambioTasa Set cActiva = 'N' Where cCtaCod = '" & sCuenta & "'"
    dbCmact.Execute sSql
    
    sSql = "Insert Into CapCambioTasa(cCtaCod, nTasaAnterior, nTasaCambio, cRegistro,cActiva)"
    sSql = sSql & "Values('" & sCuenta & "'," & nTasaAnt & "," & nTasa & ",'" & cMovnro & "','S'" & ")"
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
    Set dbCmact = New ADODB.Connection
    dbCmact.Open sConn
    dbCmact.Execute "SET DATEFORMAT MDY"
End Sub

Private Sub Class_Terminate()
    dbCmact.Close
    Set dbCmact = Nothing
End Sub
'ALPA 20100111*********************************************************
Public Function ObtenerCtaPorPersonas(ByVal sPersCod As String, ByVal sCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
        
    oCon.AbreConexion
    sSql = "exec stp_sel_ObtenerCtaPorPersonas '" & sPersCod & "'," & sCtaCod & ""
    Set ObtenerCtaPorPersonas = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function ObtenerCtasParaDepositoEnLote(ByVal sInstCod As String, ByVal sMoneda As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
        
    oCon.AbreConexion
    sSql = "exec stp_sel_ListaCtasParaDepositoEnLote '" & sInstCod & "','" & sMoneda & "'"
    Set ObtenerCtasParaDepositoEnLote = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'***********************************************************************
'MADM 20101122
Public Function GetCuentaPagoServicio(ByVal sPersCod As String, ByVal nMoneda As Integer) As String
Dim rsCta As ADODB.Recordset
Dim sSql As String
sSql = "exec stp_Sel_GetDatosCtaAhoPagoServicio '" & sPersCod & "', " & nMoneda & " "

Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsCta.ActiveConnection = Nothing
If Not (rsCta.EOF And rsCta.BOF) Then
    GetCuentaPagoServicio = rsCta("cCtaCod")
Else
    GetCuentaPagoServicio = ""
End If
Set rsCta = Nothing
End Function
'END MADM
'***********************************************************************
''ALPA 20110315**************************************************************
'Public Function GetDatosCuentaAhorroCodPer(ByVal cPersCod As String, ByVal nTipoPrograma As Integer) As ADODB.Recordset
'    Dim sSql As String
'    Dim rsCta As ADODB.Recordset
'
'    sSql = "Select P.cCtaCod cPersNombre,P.cCtaCod cPersCod  from producto P " _
'           & " inner join ProductoPersona PP on P.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac=10 " _
'           & " inner join Captaciones Cap on P.cCtaCod=Cap.cCtaCod" _
'           & " inner join Constante C on C.nConsValor = P.nPrdEstado And C.nConsCod = 2001" _
'           & " And  substring(P.cCtaCod,6,3) = '232' And nPrdEstado = 1000 " _
'           & " And Cap.nTpoPrograma in (" & nTipoPrograma & ") " _
'           & " And P.CctaCod not in (select cNroDoc from garantias G where G.cNroDoc = P.cCtaCod ) " _
'           & "where cPersCod = '" & cPersCod & "'" _
'           & " and P.cCtaCod not in (   Select cCtaCod " _
'           & "                          From RHCuentas " _
'           & "                          where cPersCod='" & cPersCod & "' " _
'           & "                              and substring(cCtaCod,6,3) in ('232')" _
'           & "                       ) "
'
'
'     Set rsCta = New ADODB.Recordset
'    rsCta.CursorLocation = adUseClient
'    rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'
'    Set rsCta.ActiveConnection = Nothing
'    Set GetDatosCuentaAhorroCodPer = rsCta
'    Set rsCta = Nothing
'
'End Function
'Public Function GetDatosCuentaAhorro(ByVal sCuenta As String) As ADODB.Recordset
'    Dim sSql As String
'    Dim rsCta As ADODB.Recordset
'    'By Capi 07082008 se adiciono columna bTasaPactada
'    sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
'        & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, AG.cAgeDescripcion, " _
'        & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.nPlazo, 0 nIntPag, '1990/01/01' dRenovacion, nMonApert nApertura, 0 nIntAdelRev, " _
'        & "3 nFormaRetiro, '1990/01/01' dAuxiliar, 0 nDuplicado, T.cConsDescripcion cEstado, 0 bBusCredPend, " _
'        & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, T3.cConsDescripcion cRetiro, " _
'        & "dUltCierre dUltCierreAnt, dUltCierre dAuxiliarAnt, ISNULL(CI.cCtaCodAbono,'') cCtaCodAbono, 3 nformaretiro,c.calias,c.nfirmasmin,0 nIntDevChq,C.cUltimaActualizacion,0 bTasaPactada FROM Producto P " _
'        & "INNER JOIN Captaciones C INNER JOIN CaptacAhorros A LEFT JOIN CaptacCtaAboIntPF CI " _
'        & "ON A.cCtaCod = CI.cCtaCod ON C.cCtaCod = A.cCtaCod ON " _
'        & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
'        & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
'        & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Constante T3 ON 3 = T3.nConsValor " _
'        & "INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod WHERE P.cCtaCod = '" & sCuenta & "' AND " _
'        & "T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
'        & "AND T2.nConsCod = " & gCaptacTipoTasa & " AND T3.nConsCod = " & gCaptacPFFormaRetiro
'
'
'    Set rsCta = New ADODB.Recordset
'    rsCta.CursorLocation = adUseClient
'    rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'
'    Set rsCta.ActiveConnection = Nothing
'    Set GetDatosCuentaAhorro = rsCta
'    Set rsCta = Nothing
'End Function
''****************************************************************************

'JACA 20111018*************************************************************************

Public Function getTarifasOpeEnOtrasAgencias(psAgencia As String, psTipoAhorro As String, pnMoneda As Integer) As Recordset
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
        
    oCon.AbreConexion
    sSql = "exec stp_sel_TarifasOpeEnOtrasAgencias '" & psAgencia & "','" & psTipoAhorro & "'," & pnMoneda & ""
    Set getTarifasOpeEnOtrasAgencias = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Sub insertarTarifaOpeEnOtrasAge(psAgeCod As String, pnTpoPrograma As Integer, pnMoneda As Integer, pnComision As Double, pnMontoMin As Currency, pnMontoMax As Currency, psMovNro As String)
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
        
    oCon.AbreConexion
    sSql = "exec stp_ins_TarifaOpeEnOtrasAge '" & psAgeCod & "'," & pnTpoPrograma & "," & pnMoneda & "," & pnComision & "," & pnMontoMin & "," & pnMontoMax & ",'" & psMovNro & "'"
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Sub updEstadoTarifaOpeEnOtrasAge(psAgeCod As String, pnTpoPrograma As String, pnMoneda As Integer)
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
        
    oCon.AbreConexion
    sSql = "exec stp_upd_EstadoTarifaOpeEnOtrasAge '" & psAgeCod & "','" & pnTpoPrograma & "'," & pnMoneda & ""
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'Public Function getValTarifaOpeEnOtrasAge(psAgencia As String, psTipoAhorro As Integer, pnMoneda As Integer) As Recordset
Public Function getValTarifaOpeEnOtrasAge(psCtaCod As String, Optional ByVal nOperacion As Integer = 0) As Recordset 'APRI20190109 ERS077-2018
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
        
    oCon.AbreConexion
    'sSql = "exec stp_sel_ValTarifaOpeEnOtrasAge '" & psAgencia & "'," & psTipoAhorro & "," & pnMoneda & ""
    'sSql = "exec stp_sel_ValTarifaOpeEnOtrasAge '" & psCtaCod & "'" 'APRI20190109 ERS077-2018
    sSql = "exec stp_sel_ValTarifaOpeEnOtrasAge '" & psCtaCod & "'," & nOperacion  'APRI20210403 ADECUACION
    Set getValTarifaOpeEnOtrasAge = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function getOperaciones(ByVal psFiltro As String) As Recordset
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
        
    oCon.AbreConexion
    sSql = "exec stp_sel_OpeTarifa '" & psFiltro & "'"
    Set getOperaciones = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function getTarifasOperaciones(psOpeCod As String, psAgencia As String, psTipoAhorro As String) As Recordset
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
        
    oCon.AbreConexion
    sSql = "exec stp_sel_TarifasOperaciones '" & psOpeCod & "','" & psAgencia & "','" & psTipoAhorro & "'"
    Set getTarifasOperaciones = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Sub updEstadoTarifaOperaciones(psOpeCod As String, psAgeCod As String, pnTpoPrograma As String)
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
        
    oCon.AbreConexion
    sSql = "exec stp_upd_EstadoTarifaOperaciones '" & psOpeCod & "','" & psAgeCod & "','" & pnTpoPrograma & "'"
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Sub insertarTarifaOperaciones(psOpeCod As String, psAgeCod As String, pnTpoPrograma As Integer, pnMontoMN As Currency, pnMontoME As Currency, pnOpeMax As Integer, psMovNro As String)
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
        
    oCon.AbreConexion
    sSql = "exec stp_ins_TarifaOperaciones '" & psOpeCod & "','" & psAgeCod & "'," & pnTpoPrograma & "," & pnMontoMN & "," & pnMontoME & "," & pnOpeMax & ",'" & psMovNro & "'"
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
Public Function getNroMovimientos(ByVal psCtaCod As String, ByVal psFecha As Date) As Integer
    Dim sql As String
    Dim rs As New ADODB.Recordset
     
    
    sql = "exec stp_sel_NroMovCta '" & psCtaCod & "','" & Format(psFecha, "yyyymmdd") & "'"
    rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF And rs.BOF) Then
       getNroMovimientos = rs!NroMov
    Else
       getNroMovimientos = 0
    End If
    
End Function
Public Function getComisionxNroOperaciones(psOpeCod As String, psAgencia As String, psTipoAhorro As Integer) As Recordset
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
        
    oCon.AbreConexion
    sSql = "exec stp_sel_ComisionxNroOperaciones '" & psOpeCod & "','" & psAgencia & "'," & psTipoAhorro & ""
    Set getComisionxNroOperaciones = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'JACA END*****************************************************************************
'*** BRGO 20111020 *****************************************************
Public Function GetCuentasCaptacxSubProducto(ByVal sPersCod As String, ByVal nProd As Integer, ByVal nTpoPrograma As Integer, Optional ByVal psCodInst As String = "") As ADODB.Recordset 'JUEZ 20140319 Se agregó psCodInst
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
        
    oCon.AbreConexion
    sSql = "exec stp_sel_ObtenerCtasCaptacPorSubProducto '" & sPersCod & "'," & nProd & "," & nTpoPrograma & ",'" & psCodInst & "'"
    Set GetCuentasCaptacxSubProducto = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'*** END BRGO *****************************************************
'***Agregado por ELRO el 20120813, según OYP-RFC024-2012
Public Function devolverVoucherSinOperacion(ByVal psAgeCod As String, _
                                            ByVal psMoneda As String) As ADODB.Recordset

    Dim sSql As String
    Dim oDCOMConecta  As COMConecta.DCOMConecta
    Set oDCOMConecta = New COMConecta.DCOMConecta

    oDCOMConecta.AbreConexion
    sSql = "exec stp_sel_DevolverVoucherSinOperacion '" & psAgeCod & "', '" & psMoneda & "'"
    Set devolverVoucherSinOperacion = oDCOMConecta.CargaRecordSet(sSql)
    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing

End Function

Public Sub confirmarCapVoucherDeposito(ByVal pnConfirmado As Integer, _
                                       ByVal psMov As String, _
                                       ByVal pnMovNroRVD As Long, Optional ByVal pnId As Long = 0)
    Dim sSql As String
    Dim oDCOMConecta  As COMConecta.DCOMConecta
    Set oDCOMConecta = New COMConecta.DCOMConecta

    oDCOMConecta.AbreConexion
    'sSql = "exec stp_upd_ActualizarCapVoucherDeposito " & pnConfirmado & ", '" & psMov & "', " & pnMovNroRVD & ""
    sSql = "exec stp_upd_ActualizarCapVoucherDeposito " & pnConfirmado & ", '" & psMov & "', " & pnMovNroRVD & "," & pnId 'EJVG20130911
    oDCOMConecta.CargaRecordSet (sSql)
    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
End Sub

Public Function devolverCajaBancosOperacionesPendientes(ByVal psFechaVoucher As String, _
                                                        ByVal psIFTpo As String, _
                                                        ByVal psPersCod As String, _
                                                        ByVal psCtaIFCod As String, _
                                                        ByVal pnMovImporte As Currency, _
                                                        ByVal psMoneda As String) As ADODB.Recordset
    Dim sSql As String
    Dim oDCOMConecta  As COMConecta.DCOMConecta
    Set oDCOMConecta = New COMConecta.DCOMConecta

    oDCOMConecta.AbreConexion
    sSql = "exec stp_sel_DevolcerCajaBancosOperaciones '" & psFechaVoucher & "', '" & psIFTpo & "', '" & psPersCod & "', '" & psCtaIFCod & "', " & pnMovImporte & ", '" & psMoneda & "'"
    Set devolverCajaBancosOperacionesPendientes = oDCOMConecta.CargaRecordSet(sSql)
    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
End Function

Public Function devolverVoucherIFSinOperacion(ByVal psAgeCod As String, _
                                              ByVal psMoneda As String, _
                                              ByVal pnTipMot As Integer, Optional ByVal psDetalle As String = "") As ADODB.Recordset
    Dim sSql As String
    Dim oDCOMConecta  As COMConecta.DCOMConecta
    Set oDCOMConecta = New COMConecta.DCOMConecta

    oDCOMConecta.AbreConexion
    'sSql = "exec stp_sel_DevolverVoucherIFSinOperacion '" & psAgeCod & "', '" & psMoneda & "', " & pnTipMot & ""
    sSql = "exec stp_sel_DevolverVoucherIFSinOperacion '" & psAgeCod & "', '" & psMoneda & "', " & pnTipMot & ",'" & psDetalle & "'"
    Set devolverVoucherIFSinOperacion = oDCOMConecta.CargaRecordSet(sSql)
    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
End Function

Public Function devolverVoucherNroIFSinOperacion(ByVal psAgeCod As String, _
                                                 ByVal psMoneda As String, _
                                                 ByVal pnTipMot As Integer, _
                                                 ByVal psIFTpo As String, _
                                                 ByVal psPersCodIF As String, Optional ByVal psDetalle As String = "") As ADODB.Recordset
    Dim sSql As String
    Dim oDCOMConecta  As COMConecta.DCOMConecta
    Set oDCOMConecta = New COMConecta.DCOMConecta

    oDCOMConecta.AbreConexion
    'sSql = "exec stp_sel_DevolverVoucherNroIFSinOperacion '" & psAgeCod & "', '" & psMoneda & "', " & pnTipMot & ", '" & psIFTpo & "', '" & psPersCodIF & "'"
    sSql = "exec stp_sel_DevolverVoucherNroIFSinOperacion '" & psAgeCod & "', '" & psMoneda & "', " & pnTipMot & ", '" & psIFTpo & "', '" & psPersCodIF & "','" & psDetalle & "'"
    Set devolverVoucherNroIFSinOperacion = oDCOMConecta.CargaRecordSet(sSql)
    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
End Function

Public Function devolverVoucherClienteNroIFSinOperacion(ByVal psAgeCod As String, _
                                                        ByVal psMoneda As String, _
                                                        ByVal pnTipMot As Integer, _
                                                        ByVal psIFTpo As String, _
                                                        ByVal psPersCodIF As String, _
                                                        ByVal psCtaIFCod As String, _
                                                        ByVal psNroVou As String)

    Dim sSql As String
    Dim oDCOMConecta  As COMConecta.DCOMConecta
    Set oDCOMConecta = New COMConecta.DCOMConecta

    oDCOMConecta.AbreConexion
    sSql = "exec stp_sel_DevolverVoucherClienteNroIFSinOperacion '" & psAgeCod & "', '" & psMoneda & "', " & pnTipMot & ", '" & psIFTpo & "', '" & psPersCodIF & "', '" & psCtaIFCod & "', '" & psNroVou & "'"
    Set devolverVoucherClienteNroIFSinOperacion = oDCOMConecta.CargaRecordSet(sSql)
    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
End Function

Public Function devolverMotivoCapVoucherDeposito() As ADODB.Recordset
Dim sSql As String
    Dim oDCOMConecta  As COMConecta.DCOMConecta
    Set oDCOMConecta = New COMConecta.DCOMConecta

    oDCOMConecta.AbreConexion
    sSql = "exec stp_sel_DevolverMotivoCapVoucherDeposito"
    Set devolverMotivoCapVoucherDeposito = oDCOMConecta.CargaRecordSet(sSql)
    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
End Function

Public Sub registrarCapVoucherDeposito(ByVal pnMovNroRVD As Long, _
                                       ByVal psFecReg As String, _
                                       ByVal psPersCodCli As String, _
                                       ByVal psIFTpo As String, _
                                       ByVal psPersCodIF As String, _
                                       ByVal psCtaIFCod As String, _
                                       ByVal pnTipMot As Integer, _
                                       ByVal pnConfirmado As Integer, _
                                       ByVal psNroVou As String, _
                                       ByVal pnMonVou As Currency, _
                                       ByVal psFecVou As String, _
                                       ByVal psAgeCod As String, _
                                       ByVal psMoneda As String, _
                                       ByVal psMov As String, _
                                       ByVal psCtaContBanco As String, _
                                       ByVal psSubCtaContBanco As String)
Dim lsSQL As String
lsSQL = "exec stp_ins_RegistrarCapRegVouDep " & pnMovNroRVD & ", '" & psFecReg & "', '" & psPersCodCli & "', '" & psIFTpo & "', '" & psPersCodIF & "', '" & psCtaIFCod & "', " & pnTipMot & ", " & pnConfirmado & ", '" & psNroVou & "', " & pnMonVou & ", '" & psFecVou & "', '" & psAgeCod & "', '" & psMoneda & "', '" & psMov & "', '" & psCtaContBanco & "', '" & psSubCtaContBanco & "'"

dbCmact.Execute lsSQL
End Sub

Public Sub registrarMovRef(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long)
Dim lsSQL As String
lsSQL = "exec stp_ins_RegistrarMovRef " & pnMovNro & ", " & pnMovNroRef & ""

dbCmact.Execute lsSQL
End Sub

'Public Sub eliminarCapVoucherDeposito(ByVal psMovNro As String, ByVal pnMovNroRVD As Long)
Public Sub eliminarCapVoucherDeposito(ByVal psMovNro As String, ByVal pnMovNroRVD As Long, Optional ByVal pnId As Long = 0)
Dim lsSQL As String
'lsSQL = "exec stp_upd_EliminarCapVoucherDeposito '" & psMovNro & "', " & pnMovNroRVD & ""
lsSQL = "exec stp_upd_EliminarCapVoucherDeposito '" & psMovNro & "', " & pnMovNroRVD & "," & pnId 'EJVG20130911

dbCmact.Execute lsSQL
End Sub

'***Fin Agregado por ELRO*******************************
'***Agregado por ELRO el 20120813, según OYP-RFC101-2012
Public Function devolverHistorialCaptacSueldosCTS(ByVal psCuenta As String, Optional ByRef psUltimaRemuneracion As Currency = 0#) As ADODB.Recordset

Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim oDCOMConecta  As New COMConecta.DCOMConecta

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_RFC1012012_DevolverHistorialCaptacSueldosCTS '" & psCuenta & "'"
Set rsCta = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

Set devolverHistorialCaptacSueldosCTS = rsCta
If Not (rsCta.BOF And rsCta.EOF) Then
    rsCta.MoveFirst
    psUltimaRemuneracion = rsCta!nSueldoTotal
Else
    psUltimaRemuneracion = 0
End If

Set rsCta = Nothing
End Function

Public Function devolverListadoCuentasCTS(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim rsCTS As New ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Dim oDCOMConecta  As New COMConecta.DCOMConecta

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_RFC1012012_DevolverListadoCuentasCTS '" & psPersCod & "'"
Set devolverListadoCuentasCTS = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

Set oDCOMConecta = Nothing
Set oDCOMCaptaGenerales = Nothing
End Function
'***Fin Agregado por ELRO el 20120813*******************
'***Agregado por ELRO el 20130111, según OYP-RFC123-2012
Public Function devolverHistoricoRenovacionesDPF(ByVal psCuenta As String) As ADODB.Recordset

Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim oDCOMConecta  As New COMConecta.DCOMConecta

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_RFC1232012_DevolverRenovacionesDPF '" & psCuenta & "'"
Set rsCta = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

Set devolverHistoricoRenovacionesDPF = rsCta
Set rsCta = Nothing
End Function
'***Fin Agregado por ELRO el 20130111*******************
'***Agregado por ELRO el 20130323, según TI-ERS011-2013***
Public Function devolverCuentasPersona(ByVal psPersCod As String) As ADODB.Recordset

Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim oDCOMConecta  As New COMConecta.DCOMConecta

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_ERS0112013_DevolverCuentasPersona '" & psPersCod & "'"
Set rsCta = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

Set devolverCuentasPersona = rsCta
Set rsCta = Nothing
End Function

Public Function devolverDatosCuentaPersona(ByVal psCtaCod As String) As ADODB.Recordset

Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim oDCOMConecta  As New COMConecta.DCOMConecta

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_ERS0112013_DevolverDatosCuentaPersona '" & psCtaCod & "'"
Set rsCta = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

Set devolverDatosCuentaPersona = rsCta
Set rsCta = Nothing
End Function

Public Function devolverSubProductoDestino() As ADODB.Recordset

Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim oDCOMConecta  As New COMConecta.DCOMConecta

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_ERS0112013_devolverSubProductoDestino "
Set rsCta = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

Set devolverSubProductoDestino = rsCta
Set rsCta = Nothing
End Function

Public Function devolverTitularCuenta(ByVal psCtaCod As String) As ADODB.Recordset

Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim oDCOMConecta  As New COMConecta.DCOMConecta

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_ERS0112013_DevolverTitularCuenta '" & psCtaCod & "'"
Set rsCta = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

Set devolverTitularCuenta = rsCta
Set rsCta = Nothing
End Function

'***Fin Agregado por ELRO el 20130323, según TI-ERS011-2013
'JUEZ 20130520 ****************************************************************
Public Function RecuperaCtasParaDebitoEnvioEstadoCta(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim oConec As New COMConecta.DCOMConecta

oConec.AbreConexion
sSql = "exec stp_sel_RecuperaCtasParaDebitoEnvioEstadoCta '" & psPersCod & "'"
Set rsCta = oConec.CargaRecordSet(sSql)
oConec.CierraConexion

Set RecuperaCtasParaDebitoEnvioEstadoCta = rsCta
Set rsCta = Nothing
End Function
Public Sub InsertarRegistroEnvioEstadoCta(ByVal psCtaCod As String, ByVal psPersCod As String, ByVal pnModoEnvio As Integer, _
                                          ByVal psCtaCodEnvio As String, ByVal psMovNro As String, Optional ByVal pnTipo As Integer = 0, Optional ByVal pnCanal As Integer)
                                          'APRI2018 ERS036-2017 ADD -> pnTipo
                                          'APRI20190615 ADD pnCanal
Dim sSql As String
Dim oConec  As COMConecta.DCOMConecta
Set oConec = New COMConecta.DCOMConecta
oConec.AbreConexion
'sSql = "exec stp_ins_RegistroEnvioEstadoCta '" & psCtaCod & "','" & psPersCod & "'," & pnModoEnvio & ",'" & psCtaCodEnvio & "','" & psMovNro & "'"
sSql = "exec stp_ins_RegistroEnvioEstadoCta '" & psCtaCod & "','" & psPersCod & "'," & pnModoEnvio & ",'" & psCtaCodEnvio & "','" & psMovNro & "'," & pnTipo & "," & pnCanal
oConec.CargaRecordSet (sSql)

oConec.CierraConexion
Set oConec = Nothing
End Sub
Public Sub EliminarRegistroEnvioEstadoCta(ByVal psCtaCod As String, Optional ByVal psPersCod As String = "", Optional ByVal pnTipo As Integer = 0, Optional ByVal psMovNro As String = "", _
                                          Optional ByVal pnCanal As Integer = 0)
'APRI20180405 ERS036-2017 ADD -> psPersCod,pnTipo,psMovNro
'APRI20190615 ADD pnCanal
Dim sSql As String
Dim oConec  As COMConecta.DCOMConecta
Set oConec = New COMConecta.DCOMConecta
    
oConec.AbreConexion
'sSql = "exec stp_upd_EliminaEnvioEstadoCta '" & psCtaCod & "'"
sSql = "exec stp_upd_EliminaEnvioEstadoCta '" & psCtaCod & "','" & psPersCod & "'," & pnTipo & ",'" & psMovNro & "'," & pnCanal
oConec.CargaRecordSet (sSql)
oConec.CierraConexion
Set oConec = Nothing
End Sub
Public Function RecuperaDatosEnvioEstadoCta(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim oConec As New COMConecta.DCOMConecta

oConec.AbreConexion
sSql = "exec stp_sel_RecuperaDatosEnvioEstadoCta '" & psCtaCod & "'"
Set rsCta = oConec.CargaRecordSet(sSql)

oConec.CierraConexion

Set RecuperaDatosEnvioEstadoCta = rsCta
Set rsCta = Nothing
End Function
Public Function RecuperaDatosEnvioEstadoCtaReporte(ByVal pnTipo As Integer, ByVal pnAnioMes As String, ByVal pnPeriodo As Integer, _
                                                   ByVal pnRango As Integer, ByVal psCodAge As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim oConec As New COMConecta.DCOMConecta

oConec.AbreConexion
sSql = "exec stp_sel_RecuperaDatosEnvioEstadoCtaReporte " & pnTipo & ",'" & pnAnioMes & "'," & pnPeriodo & "," & pnRango & ",'" & psCodAge & "'"
Set rsCta = oConec.CargaRecordSet(sSql)

oConec.CierraConexion

Set RecuperaDatosEnvioEstadoCtaReporte = rsCta
Set rsCta = Nothing
End Function
Public Function RecuperaDatosFormatoEstadoCta(ByVal psCtaCod As String, ByVal pnTipo As Integer, Optional ByVal psFecha As String = "", Optional ByVal psPersCod As String = "") As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim oConec As New COMConecta.DCOMConecta

oConec.AbreConexion
If pnTipo = 1 Then
    sSql = "exec stp_sel_RecuperaDatosFormatoEstadoCtaAhorros '" & psCtaCod & "','" & psFecha & "','" & psPersCod & "'"
Else
    sSql = "exec stp_sel_RecuperaDatosFormatoEstadoCtaCreditos '" & psCtaCod & "'"
End If
Set rsCta = oConec.CargaRecordSet(sSql)

oConec.CierraConexion

Set RecuperaDatosFormatoEstadoCta = rsCta
Set rsCta = Nothing
End Function
'APRI20180409 ERS036-2017
Public Function RecuperaHistorialGeneracionEnvioEstadoCta(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim oConec As New COMConecta.DCOMConecta

oConec.AbreConexion
sSql = "exec stp_sel_ObtenerHistorialGeneracionEnvioEstadoCta '" & psCtaCod & "'"
Set rsCta = oConec.CargaRecordSet(sSql)

oConec.CierraConexion

Set RecuperaHistorialGeneracionEnvioEstadoCta = rsCta
Set rsCta = Nothing
End Function
Public Function CargarProductosVigentes(ByVal psCodPers As String, ByVal pnTipo As Integer, ByVal pnTipoOpe As Integer) As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    sSql = "exec stp_sel_ProductosVigentesEnvioEstadoCta '" & Trim(psCodPers) & "'," & pnTipo & "," & pnTipoOpe
    Set CargarProductosVigentes = oCon.CargaRecordSet(sSql)
      
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function RecuperaCiudadAgencia(ByVal psAgeCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    sSql = "Exec stp_sel_ObtenerCiudadAgencia '" & psAgeCod & "'"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaCiudadAgencia = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function GeneraExtractoCta(ByVal psCuenta As String, ByVal psPeriodo As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    sSql = "Exec stp_sel_GeneraExtractoEmisionEstadoCta '" & psCuenta & "','" & psPeriodo & "'"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set GeneraExtractoCta = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function ObtenerDatosAhorros(ByVal psCuenta As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    sSql = "Exec stp_sel_datosCuentaAhorros '" & psCuenta & "'"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerDatosAhorros = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function MostarTextoComisionEnvioEstAhorros(ByVal pnTipo As Integer, ByVal pnTpoPrograma As Integer) As String
    Dim rsVar As New ADODB.Recordset
    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    If oCon.AbreConexion = False Then Exit Function
    
    sSql = "EXEC stp_sel_MostarTextoComisionEnvioEstAhorros " & pnTipo & "," & pnTpoPrograma
    
    MostarTextoComisionEnvioEstAhorros = ""
    Set rsVar = oCon.CargaRecordSet(sSql)
    If Not (rsVar.EOF And rsVar.BOF) Then
        MostarTextoComisionEnvioEstAhorros = rsVar("Texto")
    End If
    rsVar.Close
    Set rsVar = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'FIN APRI
'APRI20190615 SATI RFC1902040001
Public Function ObtenerNroItemEstadoCta(ByVal psPersCod As String, ByVal psMovNro As String, ByVal pnTipo As Integer) As String
Dim sSql As String
Dim oDCOMConecta  As New COMConecta.DCOMConecta
Dim rs As New ADODB.Recordset

oDCOMConecta.AbreConexion

sSql = "exec stp_sel_ObtenerNroItemEstadoCta '" & psPersCod & "','" & psMovNro & "'," & pnTipo

Set rs = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

If Not (rs.BOF And rs.EOF) Then
    ObtenerNroItemEstadoCta = rs!cNroItem
Else
    ObtenerNroItemEstadoCta = ""
End If

Set oDCOMConecta = Nothing
Set rs = Nothing
End Function
'END APRI
Public Function RecuperaDatosFormatoEstadoCtaAhorrosMov(ByVal psCtaCod As String, ByVal psFecha As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim oConec As New COMConecta.DCOMConecta

oConec.AbreConexion
sSql = "exec stp_sel_RecuperaDatosFormatoEstadoCtaAhorrosMov '" & psCtaCod & "','" & psFecha & "'"
Set rsCta = oConec.CargaRecordSet(sSql)

oConec.CierraConexion

Set RecuperaDatosFormatoEstadoCtaAhorrosMov = rsCta
Set rsCta = Nothing
End Function
Public Sub InsertarGeneracionEnvioEstadoCta(ByVal psCtaCod As String, ByVal psPersCod As String, ByVal psAnio As String, ByVal psMes As String, ByVal pnPeriodo As Integer, ByVal psMovNro As String)
Dim sSql As String
Dim oConec  As COMConecta.DCOMConecta
Set oConec = New COMConecta.DCOMConecta
    
oConec.AbreConexion
sSql = "exec stp_ins_GeneracionEnvioEstadoCta '" & psCtaCod & "','" & psPersCod & "','" & psAnio & "','" & psMes & "'," & pnPeriodo & ",'" & psMovNro & "'"
oConec.CargaRecordSet (sSql)
oConec.CierraConexion
Set oConec = Nothing
End Sub
'END JUEZ *********************************************************************
'***Agregado por ELRO el 20130701, según RFC1306270002****
Public Function devolverCuentasPersonaServicioPago(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim oDCOMConecta  As New COMConecta.DCOMConecta

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_DevolverCuentasPersonaServicioPago '" & psPersCod & "'"
Set rsCta = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

Set devolverCuentasPersonaServicioPago = rsCta
Set oDCOMConecta = Nothing
Set rsCta = Nothing
End Function

Public Function registrarServicioPago(ByVal psPersCodSerPag As String, ByVal psNomSerPag As String, _
                                      ByVal psDesSerPag As String, ByVal psCtaCod As String, _
                                      ByVal pnComision As Currency, ByVal pnPenalidad As Currency, _
                                      ByVal pnLimMaxOpeDia As Long, psFecReg As String, _
                                      ByVal psMovNro As String, Optional pnComisionTransf As Currency = 0) As Long
                                      'RIRO20150430 ERS146-2014 ADD pnComisionTransf
                                      'RIRO20150814 Se cambió pnLimMaxOpeDia por Long
Dim sSql As String
Dim oDCOMConecta  As New COMConecta.DCOMConecta
Dim rsServicioPago As New ADODB.Recordset

oDCOMConecta.AbreConexion

'RIRO20150430 Comentado
'sSql = "exec stp_ins_RegistrarServicioPago '" & psPersCodSerPag & "', '" & psNomSerPag & "', '" & psDesSerPag & "', '" & psCtaCod & "', " & pnComision & ", " & pnPenalidad & ", " & pnLimMaxOpeDia & ", '" & psFecReg & "', '" & psMovNro & "'"

'RIRO20150430 ERS146-2014 ADD pnComisionTransf
sSql = "exec stp_ins_RegistrarServicioPago '" & psPersCodSerPag _
                                              & "', '" & psNomSerPag _
                                              & "', '" & psDesSerPag _
                                              & "', '" & psCtaCod _
                                              & "', " & pnComision _
                                              & ", " & pnPenalidad _
                                              & ", " & pnLimMaxOpeDia _
                                              & ", '" & psFecReg _
                                              & "', '" & psMovNro _
                                              & "', " & pnComisionTransf

Set rsServicioPago = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

If Not (rsServicioPago.BOF And rsServicioPago.EOF) Then
    registrarServicioPago = rsServicioPago!IdSerPag
Else
    registrarServicioPago = 0
End If

Set oDCOMConecta = Nothing
Set rsServicioPago = Nothing
End Function

Public Function modificarServicioPago(ByVal psNomSerPag As String, _
                                 ByVal psDesSerPag As String, ByVal psCtaCod As String, _
                                 ByVal pnComision As Currency, ByVal pnPenalidad As Currency, _
                                 ByVal pnLimMaxOpeDia As Long, ByVal psMovNro As String, _
                                 ByVal pnIdSerPag As Long, Optional pnComisionTransf As Currency = 0) As Long
                                 'RIRO20150430 ERS146-2014 ADD pnComisionTransf
Dim sSql As String
Dim oDCOMConecta  As New COMConecta.DCOMConecta
Dim rsServicioPago As New ADODB.Recordset

oDCOMConecta.AbreConexion

'RIRO20150430 ERS146-2014 Comentado
'sSql = "exec stp_upd_ModificarServicioPago '" & psNomSerPag & "', '" & psDesSerPag & "', '" & psCtaCod & "', " & pnComision & ", " & pnPenalidad & ", " & pnLimMaxOpeDia & ", '" & psMovNro & "', " & pnIdSerPag & ""

'RIRO20150430 ERS146-2014 ADD pnComisionTransf
sSql = "exec stp_upd_ModificarServicioPago '" & psNomSerPag _
                                              & "', '" & psDesSerPag _
                                              & "', '" & psCtaCod _
                                              & "', " & pnComision _
                                              & ", " & pnPenalidad _
                                              & ", " & pnLimMaxOpeDia _
                                              & ", '" & psMovNro _
                                              & "', " & pnIdSerPag _
                                              & ", " & pnComisionTransf

Set rsServicioPago = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

If Not (rsServicioPago.BOF And rsServicioPago.EOF) Then
    modificarServicioPago = rsServicioPago!IdSerPag
Else
    modificarServicioPago = 0
End If

Set oDCOMConecta = Nothing
Set rsServicioPago = Nothing
End Function

Public Function devolverConveniosPersona(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim rsConvenios As ADODB.Recordset
Dim oDCOMConecta  As New COMConecta.DCOMConecta

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_DevolverConveniosPersona '" & psPersCod & "'"
Set rsConvenios = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

Set devolverConveniosPersona = rsConvenios
Set oDCOMConecta = Nothing
Set rsConvenios = Nothing
End Function

Public Function devolverConvenio(ByVal pnIdSerPag As Long) As ADODB.Recordset
Dim sSql As String
Dim rsConvenio As ADODB.Recordset
Dim oDCOMConecta  As New COMConecta.DCOMConecta

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_DevolverConverio " & pnIdSerPag
Set rsConvenio = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

Set devolverConvenio = rsConvenio
Set oDCOMConecta = Nothing
Set rsConvenio = Nothing
End Function

Public Function verificarNombreConvenio(ByVal psPersCodSerPag As String, ByVal psNomSerPag As String) As Long
Dim sSql As String
Dim rsConvenio As ADODB.Recordset
Dim oDCOMConecta  As New COMConecta.DCOMConecta

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_verificarNombreConvenio '" & psPersCodSerPag & "', '" & psNomSerPag & "'"
Set rsConvenio = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

If Not (rsConvenio.BOF And rsConvenio.EOF) Then
    verificarNombreConvenio = rsConvenio!Id_SerPag
Else
    verificarNombreConvenio = -1
End If
Set oDCOMConecta = Nothing
Set rsConvenio = Nothing
End Function

Public Function darBajaServicioPago(ByVal pnIdSerPag As Long, psMovNro As String) As Long
Dim sSql As String
Dim oDCOMConecta  As New COMConecta.DCOMConecta
Dim rsServicioPago As New ADODB.Recordset

oDCOMConecta.AbreConexion
sSql = "exec stp_upd_DarBajaServicioPago " & pnIdSerPag & ", '" & psMovNro & "'"
Set rsServicioPago = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

If Not (rsServicioPago.BOF And rsServicioPago.EOF) Then
    darBajaServicioPago = rsServicioPago!IdSerPag
Else
    darBajaServicioPago = 0
End If

Set oDCOMConecta = Nothing
Set rsServicioPago = Nothing
End Function

Public Function devolverConvenioVigente(ByVal psPersCodSerPag As String, ByVal pnNroSerPag As Integer) As ADODB.Recordset
Dim sSql As String
Dim rsConvenios As New ADODB.Recordset
Dim oDCOMConecta  As New COMConecta.DCOMConecta

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_DevolverConvenioVigente '" & psPersCodSerPag & "', " & pnNroSerPag & ""
Set rsConvenios = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

Set devolverConvenioVigente = rsConvenios
Set oDCOMConecta = Nothing
Set rsConvenios = Nothing
End Function

Public Function registrarArchivoServicioPago(ByVal pnIdSerPag As Long, ByVal psRefArc As String, _
                                             ByVal psNomArc As String, pnMonto As Currency, _
                                             ByVal psFecReg As String, psMovNro As String) As Long
Dim sSql As String
Dim rsServicioPago As New ADODB.Recordset

sSql = "exec stp_ins_RegistrarArchivoServicioPago " & pnIdSerPag & ", '" & psRefArc & "', '" & psNomArc & "', " & pnMonto & ", '" & psFecReg & "', '" & psMovNro & "'"
Set rsServicioPago = dbCmact.Execute(sSql)

If Not (rsServicioPago.BOF And rsServicioPago.EOF) Then
    registrarArchivoServicioPago = rsServicioPago!IdSerPagArc
Else
    registrarArchivoServicioPago = 0
End If

Set rsServicioPago = Nothing
End Function

Public Function registrarDebitoServicioPago(ByVal pnIdSerPagArc As Long, psPersCod As String, ByVal pnMonto As Currency) As Long
Dim sSql As String
Dim rsDebito As New ADODB.Recordset

sSql = "exec stp_ins_RegistrarDebitoServicioPago " & pnIdSerPagArc & ", '" & psPersCod & "', " & pnMonto & ""
Set rsDebito = dbCmact.Execute(sSql)

If Not (rsDebito.BOF And rsDebito.EOF) Then
    registrarDebitoServicioPago = rsDebito!IdSerPagDeb
Else
    registrarDebitoServicioPago = 0
End If

Set rsDebito = Nothing
End Function

Public Function devolverTramaConvenioServicioPago(ByVal pnIdSerPag As Long) As ADODB.Recordset
Dim sSql As String
Dim rsTramas As New ADODB.Recordset
Dim oDCOMConecta  As New COMConecta.DCOMConecta

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_DevolverTramaConvenioServicioPago " & pnIdSerPag & ""
Set rsTramas = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

Set devolverTramaConvenioServicioPago = rsTramas
Set oDCOMConecta = Nothing
Set rsTramas = Nothing
End Function

Public Function darBajaTramaConvenioServicioPago(ByVal pnIdSerPagArc As Long, ByVal psMovNro As String, _
                                                 ByVal psCtaCod As String, ByVal pnMonto As Currency) As Long
Dim sSql As String
Dim oDCOMConecta  As New COMConecta.DCOMConecta
Dim rsTrama As New ADODB.Recordset

oDCOMConecta.AbreConexion
sSql = "exec stp_upd_DarBajaTramaConvenioServicioPago " & pnIdSerPagArc & ", '" & psMovNro & "', '" & psCtaCod & "', " & pnMonto & ""
Set rsTrama = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

If Not (rsTrama.BOF And rsTrama.EOF) Then
    darBajaTramaConvenioServicioPago = rsTrama!IdSerPagArc
Else
    darBajaTramaConvenioServicioPago = 0
End If

Set oDCOMConecta = Nothing
Set rsTrama = Nothing
End Function

Public Function buscarConvenioServicioPago(ByVal pnTipo As Integer, ByVal psBusqueda As String) As ADODB.Recordset
Dim sSql As String
Dim rsConvenios As New ADODB.Recordset
Dim oDCOMConecta  As New COMConecta.DCOMConecta

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_BuscarConvenioServicioPago " & pnTipo & ", '" & psBusqueda & "'"
Set rsConvenios = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

Set buscarConvenioServicioPago = rsConvenios
Set oDCOMConecta = Nothing
Set rsConvenios = Nothing
End Function

Public Function devolverBeneficiariosNoCobroTramaBajaConvenioServicioPago(ByVal pnIdSerPagArc As Long) As ADODB.Recordset
Dim sSql As String
Dim rsBeneficiarios As New ADODB.Recordset
Dim oDCOMConecta  As New COMConecta.DCOMConecta

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_devolverBeneficiariosNoCobroTramaBajaConvenioServicioPago " & pnIdSerPagArc & ""
Set rsBeneficiarios = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

Set devolverBeneficiariosNoCobroTramaBajaConvenioServicioPago = rsBeneficiarios
Set oDCOMConecta = Nothing
Set rsBeneficiarios = Nothing
End Function

Public Function devolverBeneficiarioConvenioServicioPago(ByVal psPersIDnro As String, ByVal pnIdSerPag As Long) As ADODB.Recordset
Dim sSql As String
Dim rsBeneficiarios As New ADODB.Recordset
Dim oDCOMConecta  As New COMConecta.DCOMConecta

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_DevolverBeneficiarioConvenioServicioPago '" & psPersIDnro & "', " & pnIdSerPag & ""
Set rsBeneficiarios = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

Set devolverBeneficiarioConvenioServicioPago = rsBeneficiarios
Set oDCOMConecta = Nothing
Set rsBeneficiarios = Nothing
End Function
'***Fin Agregado por ELRO el 20130701, según RFC1306270002

'JUEZ 20130723 *******************************************
Public Function VerificaExisteCuentaCTS(ByVal psPersCod As String, ByVal psInstitucion As String, ByVal psMoneda As Integer, Optional ByVal psCtaCod As String = "") As Boolean
Dim sSql As String
Dim rs As New ADODB.Recordset

VerificaExisteCuentaCTS = False
sSql = "exec stp_sel_VerificaExisteCuentaCTS '" & psPersCod & "','" & psInstitucion & "'," & psMoneda & ",'" & psCtaCod & "'"
Set rs = dbCmact.Execute(sSql)

If Not (rs.BOF And rs.EOF) Then
    VerificaExisteCuentaCTS = True
End If
End Function
Public Function ObtenerCuentaCTS(ByVal psPersCod As String, ByVal psInstitucion As String, ByVal pnMoneda As Integer) As ADODB.Recordset
Dim sSql As String
Dim rs As ADODB.Recordset
Dim oConec As New COMConecta.DCOMConecta

oConec.AbreConexion
sSql = "exec stp_sel_ObtenerCuentaCTS '" & psPersCod & "','" & psInstitucion & "'," & pnMoneda
Set rs = oConec.CargaRecordSet(sSql)
oConec.CierraConexion

Set ObtenerCuentaCTS = rs
Set oConec = Nothing
End Function
'END JUEZ ************************************************
'EJVG20130908 ***
Public Function InsertarCapVoucherDepositoCab(ByVal psFecReg As String, _
                                       ByVal psPersCodCli As String, _
                                       ByVal psIFTpo As String, _
                                       ByVal psPersCodIF As String, _
                                       ByVal psCtaIFCod As String, _
                                       ByVal pnConfirmado As Integer, _
                                       ByVal psNroVou As String, _
                                       ByVal pnMonVou As Currency, _
                                       ByVal psFecVou As String, _
                                       ByVal psAgeCod As String, _
                                       ByVal psMoneda As String, _
                                       ByVal psCtaContBanco As String, _
                                       ByVal psSubCtaContBanco As String) As Long
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "Exec stp_ins_RegistrarCapRegVouDepCab '" & psFecReg & "', '" & psPersCodCli & "', '" & psIFTpo & "', '" & psPersCodIF & "', '" & psCtaIFCod & "', " & pnConfirmado & ", '" & psNroVou & "', " & pnMonVou & ", '" & psFecVou & "', '" & psAgeCod & "', '" & psMoneda & "', '" & psCtaContBanco & "', '" & psSubCtaContBanco & "'"
    Set rs = dbCmact.Execute(lsSQL)
    If Not (rs.BOF And rs.EOF) Then
        InsertarCapVoucherDepositoCab = rs!nID
    End If
End Function
Public Sub InsertarCapVoucherDepositoApert(ByVal pnMovNroRVD As Long, ByVal psPersCod As String)
    Dim lsSQL As String
    lsSQL = "Exec stp_ins_RegistrarCapRegVouDepApert " & pnMovNroRVD & ",'" & psPersCod & "'"
    dbCmact.Execute (lsSQL)
End Sub
Public Sub ActualizarCapVoucherDeposito(ByVal pnMovNroRVD As Long, Optional ByVal psCtaCodDepAumCap As String = "", Optional ByVal pnNroCliCTSDepLote As Long = 0, Optional ByVal psListaPersCod As String = "")
    Dim lsSQL As String
    Dim lsActualiza As String
    If psCtaCodDepAumCap <> "" Then
        lsActualiza = lsActualiza & " V.cCtaCodDepAumCap = '" & psCtaCodDepAumCap & "',"
    End If
    If pnNroCliCTSDepLote > 0 Then
        lsActualiza = lsActualiza & " V.nNroCliCTSDepLote = " & pnNroCliCTSDepLote & ","
    End If
    If Len(psListaPersCod) > 0 Then
        lsActualiza = lsActualiza & " V.cPersCodListaApert = '" & psListaPersCod & "',"
    End If
    If Len(lsActualiza) > 0 Then
        lsActualiza = Mid(lsActualiza, 1, Len(lsActualiza) - 1)
        lsSQL = "UPDATE V SET " & lsActualiza & " FROM CapRegVouDep V WHERE nMovNroRVD = " & pnMovNroRVD
        dbCmact.Execute (lsSQL)
    End If
End Sub
Public Sub InsertarCapVoucherDepositoRel(ByVal pnId As Long, ByVal pnMovNroRVD As Long)
    Dim lsSQL As String
    lsSQL = "Exec stp_ins_RegistrarCapRegVouDepRel " & pnId & "," & pnMovNroRVD
    dbCmact.Execute (lsSQL)
End Sub
Public Function GeneraMovNro(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "01", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String
    Dim rs As New ADODB.Recordset
    Dim sql As String

    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & Format(Time, "hh:mm:ss"), "mm/dd/yyyy hh:mm:ss") & "','" & Right(psCodAge, 2) & "','" & psUser & "'"
    Else
       sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    Set rs = dbCmact.Execute(sql)
    If Not rs.EOF Then
        GeneraMovNro = rs.Fields(0)
    End If
    Set rs = Nothing
    Exit Function
End Function
Public Function ExisteVoucherDeposito(ByVal psIFTpo As String, ByVal psPersCod As String, ByVal psCtaIFCod As String, ByVal psVoucherNro As String) As Boolean
    Dim oConn As New DCOMConecta
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "Exec stp_sel_ExisteVoucherDeposito '" & psIFTpo & "','" & psPersCod & "','" & psCtaIFCod & "','" & psVoucherNro & "'"
    oConn.AbreConexion
    Set rs = oConn.CargaRecordSet(sql)
    If Not rs.EOF Then
        ExisteVoucherDeposito = rs!bExiste
    End If
    oConn.CierraConexion
    Set rs = Nothing
    Exit Function
End Function
Public Function RecuperaVoucherxId(ByVal pnId As Long) As ADODB.Recordset
    Dim oConn As New DCOMConecta
    Dim sql As String
    sql = "Exec stp_sel_CapRegVouDepxId " & pnId
    oConn.AbreConexion
    Set RecuperaVoucherxId = oConn.CargaRecordSet(sql)
    oConn.CierraConexion
    Exit Function
End Function
Public Function RecuperaVoucherOpexId(ByVal pnId As Long) As ADODB.Recordset
    Dim oConn As New DCOMConecta
    Dim sql As String
    sql = "Exec stp_sel_CapRegVouDepOpexId " & pnId
    oConn.AbreConexion
    Set RecuperaVoucherOpexId = oConn.CargaRecordSet(sql)
    oConn.CierraConexion
    Exit Function
End Function
Public Sub MigrarDetCapVoucherDeposito(ByVal pnIdAnt As Long, ByVal pnIdNew As Long)
    Dim lsSQL As String
    lsSQL = "Exec stp_ins_MigraDetCapVoucherDeposito " & pnIdAnt & "," & pnIdNew
    dbCmact.Execute (lsSQL)
End Sub
'END EJVG *******
'RIRO20140530 ERS017 ********************
Public Function ObtenerSobranteVoucherPersona(ByVal psPersCod As String) As ADODB.Recordset
    Dim oConn As New DCOMConecta
    Dim sql As String
    sql = "Exec STP_SEL_SOBRANTES_VOUCHER '" & psPersCod & "'"
    oConn.AbreConexion
    Set ObtenerSobranteVoucherPersona = oConn.CargaRecordSet(sql)
    oConn.CierraConexion
    Exit Function
End Function
'RIRO20140530 ERS017 ********************
'*************RECO 20131023 ERS141*******************
Public Function RecuperaSubTipoProducto(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConn As New DCOMConecta
    Dim sql As String
    sql = "Exec stp_sel_ObtenerSubProductoCaptaciones '" & psCtaCod & "'"
    oConn.AbreConexion
    Set RecuperaSubTipoProducto = oConn.CargaRecordSet(sql)
    oConn.CierraConexion
    Exit Function
End Function

Public Function ValidaFondoGarantEcotaxi(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConn As New DCOMConecta
    Dim sql As String
    sql = "Exec stp_sel_ValidaFondoGarantEcotaxi '" & psCtaCod & "'"
    oConn.AbreConexion
    Set ValidaFondoGarantEcotaxi = oConn.CargaRecordSet(sql)
    oConn.CierraConexion
    Exit Function
End Function

'**************************END RECO******************

'********** Add By GITU 09/12/2013 **********'
Public Function ValidaTarjetizacionDat(ByVal sCuenta As String) As ADODB.Recordset

Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim lbflag As Boolean

sSql = "exec stp_sel_ValidaTarjetizacion '" & sCuenta & "'"

Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsCta.ActiveConnection = Nothing
Set ValidaTarjetizacionDat = rsCta
Set rsCta = Nothing
End Function
'********** END GITU ************'
'*** FRHU 20140227 RQ14006
Public Function GetCuentaAhorroPersona(ByVal cPersCod As String, ByVal cMoneda As String) As ADODB.Recordset
    Dim sSql As String
    Dim rsCta As ADODB.Recordset
    
    sSql = " exec stp_sel_CargarCtaAhorros '" & cPersCod & "','" & cMoneda & "'"
    
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    Set rsCta.ActiveConnection = Nothing
    Set GetCuentaAhorroPersona = rsCta
    Set rsCta = Nothing
End Function
'FIN FRHU 20140227
'*** FRHU 20140308 RQ14008
Public Function ObtenerRelacionAhoCred(ByVal cCtaAho As String) As ADODB.Recordset
    Dim sSql As String
    Dim rsCta As ADODB.Recordset
    
    sSql = " exec stp_sel_RelacionProductoAhorroCredito '" & cCtaAho & "'"
    
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    Set rsCta.ActiveConnection = Nothing
    Set ObtenerRelacionAhoCred = rsCta
    Set rsCta = Nothing
End Function
Public Sub ActualizaRelacionAhoCred(ByVal psCuentaAho As String, ByVal psMovNroDes As String, ByVal pbEstado As Boolean)
Dim sSql As String
'sSql = "DELETE FROM RelaAhoCred WHERE cCtaCodAho = '" & psCuentaAho & "'"
sSql = "UPDATE RelaAhoCred SET cMovNroDes = '" & psMovNroDes & "', bEstado = " & IIf(pbEstado = True, 1, 0) & _
       " WHERE cCtaCodAho = '" & psCuentaAho & "' AND bEstado = 1"
dbCmact.Execute sSql
End Sub
'FIN FRHU 20140308
'JUEZ 20140304 **********************************************************
Public Sub CambioTasaCTSCtasVig(ByVal pnMoneda As Moneda, ByVal psMovNro As String, ByVal nTNACTSCajaSueldo As Double, ByVal nTNACTSActivo As Double, ByVal nTNACTSNoActivo As Double)
    Dim lsSQL As String
    lsSQL = "Exec stp_upd_CambioTasaCTSCtasVig " & pnMoneda & ",'" & psMovNro & "'," & nTNACTSCajaSueldo & "," & nTNACTSActivo & "," & nTNACTSNoActivo
    dbCmact.Execute (lsSQL)
End Sub
Public Sub CambioTasaCTS(ByVal psCuenta As String, ByVal pnTasaAnt As Double, ByVal pnTasa As Double, ByVal psMovNro As String)
    Dim lsSQL As String
    lsSQL = "Exec stp_ins_CambioTasaCTS '" & psCuenta & "'," & pnTasaAnt & "," & pnTasa & ",'" & psMovNro & "'"
    dbCmact.Execute (lsSQL)
End Sub
Public Function RecuperaDatosTasaHistorica(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConn As New DCOMConecta
    Dim sSql As String
    sSql = "Exec stp_sel_RecuperaDatosTasaHistorica '" & psCtaCod & "'"
    oConn.AbreConexion
    Set RecuperaDatosTasaHistorica = oConn.CargaRecordSet(sSql)
    oConn.CierraConexion
    Exit Function
End Function
Public Function RecuperaDatosCuentasCTSInstitucion(ByVal psPersCod As String, ByVal psInstitucion As String) As ADODB.Recordset
    Dim oConn As New DCOMConecta
    Dim sSql As String
    sSql = "Exec stp_sel_RecuperaDatosCuentasCTSInstitucion '" & psPersCod & "','" & psInstitucion & "'"
    oConn.AbreConexion
    Set RecuperaDatosCuentasCTSInstitucion = oConn.CargaRecordSet(sSql)
    oConn.CierraConexion
    Exit Function
End Function
Public Function RecuperaDatosCambioTasas(ByVal psMovNro As String) As ADODB.Recordset
    Dim oConn As New DCOMConecta
    Dim sSql As String
    sSql = "Exec stp_sel_RecuperaDatosCambioTasas '" & psMovNro & "'"
    oConn.AbreConexion
    Set RecuperaDatosCambioTasas = oConn.CargaRecordSet(sSql)
    oConn.CierraConexion
    Exit Function
End Function
'END JUEZ ***************************************************************
'*********************JAME 20140319, Segun TI-ERS052-2014 ****************
Public Function DevuelveTasaAperturaPlazoFijo(ByVal psCtaCod As String) As Double
    Dim oConn As New DCOMConecta
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "Exec stp_sel_obtenerTasaAperturaPlazoFijo '" & psCtaCod & "'"
    oConn.AbreConexion
    Set rs = oConn.CargaRecordSet(sql)
    If Not rs.EOF Then
        DevuelveTasaAperturaPlazoFijo = rs!nTasaApertura
    End If
    oConn.CierraConexion
    Set rs = Nothing
    Exit Function
End Function
'********END JAME************
'FRHU 20140926 TI-ERS099-2014
Public Function GetDatosPersonaSolicitudApertura(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim rsCta As ADODB.Recordset
    sSql = "exec stp_sel_GetDatosPersonaSolicitudApertura '" & psPersCod & "'"
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set rsCta.ActiveConnection = Nothing
    Set GetDatosPersonaSolicitudApertura = rsCta
    Set rsCta = Nothing
End Function
Public Function GetPersCodJuridicaSolicitudApertura(ByVal psCuenta As String) As String
    Dim oCon As New DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_GetPersonaJuridicaSolicitudApertura '" & psCuenta & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
        GetPersCodJuridicaSolicitudApertura = rs!cPersCod
    Else
        GetPersCodJuridicaSolicitudApertura = ""
    End If
    oCon.CierraConexion
    Set rs = Nothing
    
    Exit Function
End Function
Public Function GetInstitucionSolicitudApertura(ByVal psCuenta As String, ByVal pnTpoPrograma As Integer) As String
    Dim oCon As New DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_GetInstitucionSolicitudApertura '" & psCuenta & "'," & pnTpoPrograma
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
        GetInstitucionSolicitudApertura = rs!cpersnombre
    Else
        GetInstitucionSolicitudApertura = ""
    End If
    oCon.CierraConexion
    Set rs = Nothing
    
    Exit Function
End Function
Public Function GetDatosParaCronogramaPlazoFijo(ByVal psCuenta As String) As ADODB.Recordset
    Dim sSql As String
    Dim rsCta As ADODB.Recordset
    sSql = "exec stp_sel_GetDatosParaCronogramaPlazoFijo '" & psCuenta & "'"
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set rsCta.ActiveConnection = Nothing
    Set GetDatosParaCronogramaPlazoFijo = rsCta
    Set rsCta = Nothing
End Function
'FIN FRHU 20140926
'JUEZ 20141014 *******************************************************
Public Function ObtenerCantidadOperaciones(ByVal psCtaCod As String, ByVal pnTipoOpe As CaptacMovTipo, ByVal psPeriodo As String) As ADODB.Recordset
    Dim oConn As New DCOMConecta
    Dim sSql As String
    sSql = "Exec stp_sel_CaptacObtenerCantOperaciones '" & psCtaCod & "'," & pnTipoOpe & ",'" & psPeriodo & "'"
    oConn.AbreConexion
    Set ObtenerCantidadOperaciones = oConn.CargaRecordSet(sSql)
    oConn.CierraConexion
    Exit Function
End Function
'END JUEZ ************************************************************
'FRHU 20141201 ERS048-2014*******************************************************
Public Function GetProductosAhorros(ByVal psPersCod As String) As ADODB.Recordset
    Dim oConn As New DCOMConecta
    Dim sSql As String
    sSql = "Exec stp_sel_GetProductosAhorros '" & psPersCod & "'"
    oConn.AbreConexion
    Set GetProductosAhorros = oConn.CargaRecordSet(sSql)
    oConn.CierraConexion
    Exit Function
End Function
Public Function GetMovProductoConcepto(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

    sSql = "exec stp_sel_GetMovProductoConcepto " & pnMovNro

Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set GetMovProductoConcepto = rsCta

Set rsCta.ActiveConnection = Nothing
Set rsCta = Nothing
End Function
'END FRHU ************************************************************
'JUEZ 20150205 *****************************************************************
Public Function ValidaServCobDebitoAutoDet(ByVal psCtaCod As String, ByVal pnTipoDebito As CaptacServDebitoAutomaticoTipo, ByVal psValor1 As String, ByVal psValor2 As String) As Boolean
    Dim oConn As New DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    ValidaServCobDebitoAutoDet = False
    sSql = "Exec stp_sel_ValidaServCobDebitoAutoDet '" & psCtaCod & "'," & pnTipoDebito & ",'" & psValor1 & "','" & psValor2 & "'"
    oConn.AbreConexion
    Set rs = oConn.CargaRecordSet(sSql)
    If Not rs.EOF Then
        ValidaServCobDebitoAutoDet = True
    End If
    oConn.CierraConexion
    Exit Function
End Function
Public Function ValidaServCobDebitoAuto(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConn As New DCOMConecta
    Dim sSql As String
    
    sSql = "Exec stp_sel_ValidaServCobDebitoAuto '" & psCtaCod & "'"
    oConn.AbreConexion
    Set ValidaServCobDebitoAuto = oConn.CargaRecordSet(sSql)
    oConn.CierraConexion
    Exit Function
End Function
Public Function ObtenerServCobDebitoAuto(ByVal pnMovNro As Long, ByVal psUsuario As String, ByVal psFecha As String, ByVal psAgeCod As String, Optional ByVal pbMuestraExtorno As Boolean = False) As ADODB.Recordset
Dim oConn As New DCOMConecta
Dim sSql As String
    
    sSql = "Exec stp_sel_ObtenerServCobDebitoAuto " & pnMovNro & ",'" & psUsuario & "','" & psFecha & "','" & psAgeCod & "'," & IIf(pbMuestraExtorno, 1, 0)
    oConn.AbreConexion
    Set ObtenerServCobDebitoAuto = oConn.CargaRecordSet(sSql)
    oConn.CierraConexion
    Exit Function
End Function
Public Function ObtenerServCobDebitoAutoDet(ByVal pnMovNro As Long, ByVal pnTipoDebito As CaptacServDebitoAutomaticoTipo) As ADODB.Recordset
Dim oConn As New DCOMConecta
Dim sSql As String
    
    sSql = "Exec stp_sel_ObtenerServCobDebitoAutoDet " & pnMovNro & "," & pnTipoDebito
    oConn.AbreConexion
    Set ObtenerServCobDebitoAutoDet = oConn.CargaRecordSet(sSql)
    oConn.CierraConexion
    Exit Function
End Function
Public Function ObtenerDebitoAutomaticoCredServ(ByVal pnTipoDebito As CaptacServDebitoAutomaticoTipo, ByVal pdFecSis As Date)
Dim oConn As New DCOMConecta
Dim sSql As String
    
    If pnTipoDebito = gServConvenio Then
        sSql = "Exec stp_sel_ObtenerDebitoAutomaticoPagoServicios '" & Format(pdFecSis, "yyyyMMdd") & "'"
    ElseIf pnTipoDebito = gServCredito Then
        sSql = "Exec stp_sel_ObtenerDebitoAutomaticoPagoCreditos '" & Format(pdFecSis, "yyyyMMdd") & "'"
    End If
    oConn.AbreConexion
    Set ObtenerDebitoAutomaticoCredServ = oConn.CargaRecordSet(sSql)
    oConn.CierraConexion
    Exit Function
End Function
'END JUEZ **********************************************************************
'JUEZ 20151229 *************************************************************************
Public Sub RegistrarPendComisionMantPoderes(ByVal psCuenta As String, ByVal pnMovNro As Long, ByVal psMovNro As String, ByVal pnComiMantPoderes As Double)
Dim lsSQL As String
lsSQL = "exec stp_ins_PendComisionMantPoderes '" & psCuenta & "', " & pnMovNro & ", '" & psMovNro & "', " & pnComiMantPoderes

dbCmact.Execute lsSQL
'END JUEZ ******************************************************************************
End Sub
'RIRO20160311 **************************************************
Public Function GetDoiCta(ByVal psCtaCod As String) As String
    Dim oCon As New DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    sSql = "exec STP_SEL_DOI_CCTACOD '" & psCtaCod & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
        GetDoiCta = rs!DOI
    Else
        GetDoiCta = ""
    End If
    oCon.CierraConexion
    Set rs = Nothing
    
    Exit Function
End Function

'GIPO ERS070-2016 Sobrantes de Adjudicados
Public Function devolverBeneficiarioSobrantesDeAdjudicados(ByVal psPersIDnro As String, ByVal psTipoDOI As String) As ADODB.Recordset
    Dim sSql As String
    Dim rsBeneficiarios As New ADODB.Recordset
    Dim oDCOMConecta  As New COMConecta.DCOMConecta
    
    oDCOMConecta.AbreConexion
    sSql = "exec stp_sel_ERS0702016_ObtienePersonaSobrantesDeAdjudicado '" & psPersIDnro & "'," & psTipoDOI
    Set rsBeneficiarios = oDCOMConecta.CargaRecordSet(sSql)
    oDCOMConecta.CierraConexion
    Set devolverBeneficiarioSobrantesDeAdjudicados = rsBeneficiarios
    Set oDCOMConecta = Nothing
    Set rsBeneficiarios = Nothing
End Function
'GIPO ERS070-2016
Public Function devolverTipoDOI() As ADODB.Recordset
    Dim sSql As String
    Dim rsTiposDOI As New ADODB.Recordset
    Dim oDCOMConecta  As New COMConecta.DCOMConecta
    oDCOMConecta.AbreConexion
    sSql = "exec stp_sel_ERS0702016_TiposDOI"
    Set rsTiposDOI = oDCOMConecta.CargaRecordSet(sSql)
    oDCOMConecta.CierraConexion
    Set devolverTipoDOI = rsTiposDOI
    Set oDCOMConecta = Nothing
    Set rsTiposDOI = Nothing
End Function
'GIPO ERS070-2016 20-04-2017
Public Function obtenerFechaGeneracionCarta() As ADODB.Recordset
    Dim sSql As String
    Dim rsFechas As New ADODB.Recordset
    Dim oDCOMConecta  As New COMConecta.DCOMConecta
    oDCOMConecta.AbreConexion
    sSql = "exec sp_sel_ERS0702016_GetFechaGeneracionCartasAll "
    Set rsFechas = oDCOMConecta.CargaRecordSet(sSql)
    oDCOMConecta.CierraConexion
    Set obtenerFechaGeneracionCarta = rsFechas
    Set oDCOMConecta = Nothing
    Set rsFechas = Nothing
End Function

'GIPO ERS070-2016 20-04-2017
Public Function obtenerAutorizacionDevolucion(ByVal psPersIDnro As String, ByVal pscAgeCod As String, _
        ByVal psUser As String, ByVal fechaHoy As Date) As ADODB.Recordset
    Dim sSql As String
    Dim rsAutorizacion As New ADODB.Recordset
    Dim oDCOMConecta  As New COMConecta.DCOMConecta
    oDCOMConecta.AbreConexion
    Dim fechaString As String
    fechaString = Format(fechaHoy, "yyyymmdd")
    sSql = "exec stp_sel_ERS0702016_AutorizacionSobranteAdj '" & psPersIDnro & "','" & pscAgeCod & "','" & psUser & "','" & fechaString & "'"
    Set rsAutorizacion = oDCOMConecta.CargaRecordSet(sSql)
    oDCOMConecta.CierraConexion
    Set obtenerAutorizacionDevolucion = rsAutorizacion
    Set oDCOMConecta = Nothing
    Set rsAutorizacion = Nothing
End Function


'GIPO 07-03-2017 ERS070-2016
Public Function RegistrarDevolucionSobrantesAdj(ByVal fecha As String, ByVal codAge As String, ByVal user As String, _
                                           ByVal cUsuarioVisto As String, ByVal psOpeCod As String, ByVal prsDebitos As ADODB.Recordset, _
                                           ByVal psTotal As Currency, ByVal ITF As Currency, ByVal cDoi As String) As String
    On Error GoTo Error
    Dim sql As String

'    Dim oConecta As New COMConecta.DCOMConecta
    'iniciamos la transacción
    oConecta.AbreConexion
    'dbCmact.BeginTrans
    oConecta.BeginTrans
    Dim nMovNro As Long
    Dim cMovnro As String
    Dim lsMovNro As String
    Dim lsMovNroITF As String
    
    cMovnro = GeneraMovNro(fecha, codAge, user)
   
    If ITF <> 0# Then
        lsMovNroITF = GeneraMovNro(fecha, codAge, user, cMovnro)
    Else
        lsMovNroITF = ""
    End If
    nMovNro = registrarMontosDevolucion(cMovnro, lsMovNroITF, psOpeCod, psTotal, ITF)
    registrarDevolucionesCreditosPing prsDebitos, cUsuarioVisto, nMovNro, cDoi, ITF
    'después de la transacción
    oConecta.CommitTrans
    oConecta.CierraConexion
    RegistrarDevolucionSobrantesAdj = cMovnro
    Exit Function
    
Error:
    ''revertimos el proceso
    'oConecta.CierraConexion
    oConecta.RollbackTrans
    Err.Raise Err.Number, "", Err.Description
End Function

'GIPO 07-03-2017 ERS070-2016
Public Function registrarMontosDevolucion(ByVal lsMovNro As String, ByVal lsMovNroITF As String, _
                                           ByVal psOpeCod As String, ByVal psTotal As Currency, _
                                           ByVal ITF As Currency) As Long
                                           
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    Dim nMovNro As Long
    Dim codigoOpITF As String
    'codigoOpITF = gITFCobroEfectivo
    codigoOpITF = "990307"
   
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cMovNro", adVarChar, adParamInput, 25, lsMovNro)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cMovNroITF", adVarChar, adParamInput, 25, lsMovNroITF)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cOpeCod", adVarChar, adParamInput, 20, psOpeCod)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cOpeCodITF", adVarChar, adParamInput, 20, codigoOpITF)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nImporte", adCurrency, adParamInput, , psTotal)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nImporteITF", adCurrency, adParamInput, 20)
    Prm.Value = ITF
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nMoneda", adInteger, adParamInput, 20)
    Prm.Value = 1
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nMovNro", adInteger, adParamOutput, 20)
    Cmd.Parameters.Append Prm
    
    oConecta.AbreConexion

    Cmd.ActiveConnection = oConecta.ConexionActiva
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "stp_ins_ERS0702016_MovDevolucionSobrantesAdj"
    Cmd.Execute
    
    nMovNro = Cmd.Parameters(7).Value
    
'    oConecta.CierraConexion
    
    registrarMontosDevolucion = nMovNro
    
End Function

'Created by GIPO 10-03-2017 ERS070-2016
Public Sub registrarDevolucionesCreditosPing(ByVal prsDebitos As ADODB.Recordset, ByVal cUsuarioVisto As String, _
                            ByVal nMovNro As Long, ByVal cDoi As String, ByVal ITF As Currency)
    Dim Prm As New ADODB.Parameter
    Dim Cmd As New Command
    Dim bAplicaITF As Boolean
    Dim nIdDevolucion As Long
    
    If ITF <> 0 Then
        bAplicaITF = True
    Else
        bAplicaITF = False
    End If
    
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cDOI", adVarChar, adParamInput, 25, cDoi)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@bAplicaITF", adBoolean, adParamInput, 25, bAplicaITF)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nMovNro", adInteger, adParamInput, 25, nMovNro)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cUsuarioSuperv", adVarChar, adParamInput, 25, cUsuarioVisto)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nIdDevolucion", adInteger, adParamOutput, 25)
    Cmd.Parameters.Append Prm
    
    oConecta.AbreConexion
    Cmd.ActiveConnection = oConecta.ConexionActiva
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "stp_ins_ERS0702016_DevolucionSobrantePignoraticio"
    Cmd.Execute
    
    nIdDevolucion = Cmd.Parameters(4).Value
'    oConecta.CierraConexion
    
    prsDebitos.MoveFirst
    Do While Not prsDebitos.EOF
        Set Cmd = New ADODB.Command
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@nIdDevolucion", adInteger, adParamInput, 25, nIdDevolucion)
        Cmd.Parameters.Append Prm
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@cCtaCod", adVarChar, adParamInput, 25, prsDebitos!NroContrato)
        Cmd.Parameters.Append Prm
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@nMontoDevuelto", adCurrency, adParamInput, 25, prsDebitos!Importe)
        Cmd.Parameters.Append Prm
        
         
        oConecta.AbreConexion
        Cmd.ActiveConnection = oConecta.ConexionActiva
        Cmd.CommandType = adCmdStoredProc
        Cmd.CommandText = "stp_ins_ERS0702016_DevolucionSobrantePignoraticioDet"
        Cmd.Execute
        
'        oConecta.CierraConexion
        prsDebitos.MoveNext
    Loop
End Sub

'GIPO 26-04-2017 ERS070-2016
Public Sub AgregarAutorizacionDevolucionSobrante(ByVal lsMovNro As String, ByVal nMontoDevolucion As Double, ByVal cPersCodCliente As String)
Dim sSql As String
sSql = "Exec stp_ins_ERS0702016_DevolucionPignoraticioAutoriza '" & lsMovNro & "'," & nMontoDevolucion & ",'" & cPersCodCliente & "'"
dbCmact.Execute sSql
End Sub

'END RIRO ******************************************************
'PASI20160616 CCE
'Public Function CCE_GetCuentasPersona(ByVal psPersCod As String) As ADODB.Recordset
'Dim sSql As String
'Dim rsCta As ADODB.Recordset
'    sSql = "stp_sel_CCE_CargaCuentasPersona '" & psPersCod & "'"
'Set rsCta = New ADODB.Recordset
'rsCta.CursorLocation = adUseClient
'rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'
'Set rsCta.ActiveConnection = Nothing
'Set CCE_GetCuentasPersona = rsCta
'Set rsCta = Nothing
'End Function
'Public Function CCE_ObtieneCCI(ByVal psCtaCod As String) As ADODB.Recordset
'Dim sSql As String
'Dim rsCta As ADODB.Recordset
'    sSql = "stp_sel_CCE_ObtieneCCI '" & psCtaCod & "'"
'Set rsCta = New ADODB.Recordset
'rsCta.CursorLocation = adUseClient
'rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'
'Set rsCta.ActiveConnection = Nothing
'Set CCE_ObtieneCCI = rsCta
'Set rsCta = Nothing
'End Function
'END PASI

'************** ANDE 20170615 ARQUEO DE EXPEDIENTES DE AHORRO ***************
Public Function ObtenerDatosParticipantesArqueoExpAho(ByVal pcUsuario As String) As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    Dim R As New ADODB.Recordset
    
    On Error GoTo ErrorExtraccionData
    
    oCon.AbreConexion
    cSql = "exec stp_Sel_ParticipantesArqueoAhorro_ERS021_2017 '" & pcUsuario & "'"
    Set R = oCon.CargaRecordSet(cSql)
    oCon.CierraConexion
    
    Set ObtenerDatosParticipantesArqueoExpAho = R
        
    Set R = Nothing
    Exit Function
ErrorExtraccionData:
    Err.Raise Err.Number, "Error en el proceso", Err.Description
End Function

Public Function ObtenerCuentasAperturadas(ByVal pcDel As String, ByVal pcAl As String, ByVal pcAgeCod As String) As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    Dim R As ADODB.Recordset
    On Error GoTo ErrorObtenerCuentas
     
    cSql = "exec  stp_sel_DatosCuentasAperturadas_ERS021_2017 '" & Format(pcDel, "YYYY-MM-DD") & "','" & Format(pcAl, "YYYY-MM-DD") & "','" & pcAgeCod & "'"
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(cSql)
    oCon.CierraConexion
    
    Set ObtenerCuentasAperturadas = R
    Set R = Nothing
    
    Exit Function
    
ErrorObtenerCuentas:
     Err.Raise Err.Number, "Error en el proceso", Err.Description
End Function

Public Function ObtenerTotalCuentasAperturadas(ByVal pcDel As String, ByVal pcAl As String, ByVal pcAgeCod As String)
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    Dim R As ADODB.Recordset
    
    On Error GoTo ErrorObtenerTotalAperturadas
    cSql = "exec stp_sel_TotalCuentasAperturadas_ERS021_2017 '" & Format(pcDel, "YYYY-MM-DD") & "','" & Format(pcAl, "YYYY-MM-DD") & "','" & pcAgeCod & "'"
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(cSql)
    oCon.CierraConexion
    Set ObtenerTotalCuentasAperturadas = R
    
    Set R = Nothing
    Exit Function
    
ErrorObtenerTotalAperturadas:
     Err.Raise Err.Number, "Error en el proceso", Err.Description
End Function

Public Sub GuardarPistaArqueoExpAho(ByVal pcCodArqueo As String, ByVal pcPersCod As String, ByVal pnTipoArqueo As Integer, ByVal pcDesTipoArqueo As String, ByVal pnPerfil As Integer, ByVal pcAgeCod As String, ByVal pcArqueo As String)
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    
    cSql = "exec stp_ins_ArqueoExpedienteAho_ERS021_2017 '" & pcCodArqueo & "','" & pcPersCod & "','" & pnTipoArqueo & "','" & pcDesTipoArqueo & "','" & pnPerfil & "','" & pcAgeCod & "','" & pcArqueo & "'"
    oCon.AbreConexion
    oCon.Ejecutar (cSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
End Sub

Public Function ObtenerArqueosDeldia(ByVal pdDia As Date) As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    Dim R As ADODB.Recordset
    
    cSql = "exec stp_sel_ArqueosDelAdia_ERS021_2017 '" & Format(pdDia, "YYYY-MM-DD") & "'"
    On Error GoTo ErrorObtenerArqueoDelDia
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(cSql)
    oCon.CierraConexion
    Set ObtenerArqueosDeldia = R
    Set R = Nothing
    Exit Function
    
ErrorObtenerArqueoDelDia:
    Err.Raise Err.Number, "Error al obtener arqueos del día", Err.Description
End Function

Public Sub ExtornarArqueoExpAho(ByVal pcIdArqueo As String)
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    
    cSql = "exec stp_upd_ExtornarArqueoExpAho_ERS021_2017 '" & pcIdArqueo & "'"
    
    oCon.AbreConexion
    oCon.Ejecutar (cSql)
    oCon.CierraConexion
End Sub

Public Function ObtenerCiudad(ByVal pcAgeCod As String) As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    Dim R As ADODB.Recordset
    
    cSql = "exec stp_sel_CiudadOperacion_ERS021_2017 '" & pcAgeCod & "'"
        
    On Error GoTo ErrorObtenerCiudad
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(cSql)
    oCon.CierraConexion
    
    Set ObtenerCiudad = R
    Set R = Nothing
    
    Exit Function
ErrorObtenerCiudad:
    Err.Raise Err.Number, "Error al obtener ciudad de operación", Err.Description
End Function

Public Sub RegistrarDeficientesYFaltantes(ByVal pcCodArqueo As String, ByVal pcCuenta As String, ByVal pdFecha As Date, ByVal pcComentario As String, ByVal pcMotivo As String)
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    
    On Error GoTo ErrorRegistrarDF
    cSql = "exec stp_ins_arqueosDeficientesYFaltantes_ERS021_2017 '" & pcCodArqueo & "','" & pcCuenta & "','" & Format(pdFecha, "yyyy-MM-dd") & "','" & pcComentario & "','" & pcMotivo & "'"
    
    oCon.AbreConexion
    oCon.Ejecutar (cSql)
    oCon.CierraConexion
    
    Set oCon = Nothing
    Exit Sub
ErrorRegistrarDF:
    Err.Raise Err.Number, "Error al registrar Deficientes y/o Faltantes", Err.Description
End Sub

Public Function ObtenerCodigoArqueo(ByVal pcFecha As String) As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    Dim R As ADODB.Recordset
    On Error GoTo ErrorObtenerCodigo
    
    cSql = "exec stp_sel_codigoArqueo_ERS021_2017 '" & pcFecha & "'"
    
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(cSql)
    oCon.CierraConexion
    
    Set ObtenerCodigoArqueo = R
    Set R = Nothing
    Exit Function
ErrorObtenerCodigo:
    Err.Raise Err.Number, "Error al generar código d arqueo", Err.Description
End Function

'***************** END ANDE **********************

'***MARG ERS065-2017******************************
Public Function InsertarCapAutSinTarjetaVisto(pcUserSolicitud As String, pcCodAgeSolicitud As String, pcMotivoAutorizacion As String, pcCtaCodCliente As String, pcPersCodCliente As String, pcOpeCod As String, nMonto As Currency) As Integer
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim oRS As New ADODB.Recordset

    On Error GoTo ErrorInsertarCapAutSinTarjetaVisto

    sSql = " exec stp_ins_ERS0652017_InsertarCapAutSinTarjetaVisto '" & pcUserSolicitud & "','" & pcCodAgeSolicitud & "','" & pcMotivoAutorizacion & "','" & pcCtaCodCliente & "','" & pcPersCodCliente & "','" & pcOpeCod & "'," & nMonto & ""
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set oRS = oCon.Ejecutar(sSql)
    If Not oRS.BOF And Not oRS.EOF Then
        If oRS.RecordCount > 0 Then
            InsertarCapAutSinTarjetaVisto = oRS!nID
        Else
            InsertarCapAutSinTarjetaVisto = 0
        End If
    Else
        InsertarCapAutSinTarjetaVisto = 0
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ErrorInsertarCapAutSinTarjetaVisto:
    Err.Raise Err.Number, "InsertarCapAutSinTarjetaVisto ", Err.Description
    InsertarCapAutSinTarjetaVisto = 0
End Function

Public Function ListarCapAutSinTarjetaVisto(ByVal cAgeCod As String) As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    Dim R As ADODB.Recordset
    On Error GoTo ErrorListarCapAutSinTarjetaVisto
    
    cSql = "exec stp_sel_ERS065_2017_ListarCapAutSinTarjetaVisto '" & cAgeCod & "'"
    
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(cSql)
    oCon.CierraConexion
    
    Set ListarCapAutSinTarjetaVisto = R
    Set R = Nothing
    Exit Function
ErrorListarCapAutSinTarjetaVisto:
    Err.Raise Err.Number, "Error al ListarCapAutSinTarjetaVisto", Err.Description
End Function

Public Sub ActualizarCapAutSinTarjetaVisto(ByVal pnIdVisto As Integer, ByVal pcUserVisto As String, ByVal bAceptado As Boolean)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

    On Error GoTo ErrorActualizarCapAutSinTarjetaVisto

    sSql = " exec stp_upd_ERS065_2017_ActualizarCapAutSinTarjetaVisto " & pnIdVisto & ",'" & pcUserVisto & "'," & IIf(bAceptado = True, 1, 0)

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub
ErrorActualizarCapAutSinTarjetaVisto:
    Err.Raise Err.Number, "ActualizarCapAutSinTarjetaVisto ", Err.Description
End Sub

Public Sub ActualizarCapAutSinTarjetaVisto_nMovNro(ByVal cUserSolicitud As String, ByVal cCodAgeSolicitud As String, ByVal cCtaCodCliente As String, ByVal cPersCodCliente As String, ByVal nMovNro As Long, pcOpeCod As String)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

    On Error GoTo ErrorActualizarCapAutSinTarjetaVisto_nMovNro

    sSql = " exec stp_upd_ERS065_2017_ActualizarCapAutSinTarjetaVisto_nMovNro '" & cUserSolicitud & "','" & cCodAgeSolicitud & "','" & cCtaCodCliente & "','" & cPersCodCliente & "'," & nMovNro & ",'" & pcOpeCod & "'"

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.Ejecutar (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub
ErrorActualizarCapAutSinTarjetaVisto_nMovNro:
    Err.Raise Err.Number, "ActualizarCapAutSinTarjetaVisto_nMovNro ", Err.Description
End Sub

Public Function SolicitarVistoAtencionSinTarjeta(pcUserSolicitud As String, pcCodAgeSolicitud As String, pcCtaCodCliente As String, pcPersCodCliente As String, pcOpeCod As String) As Integer
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim oRS As New ADODB.Recordset

    On Error GoTo ErrorSolicitarVistoAtencionSinTarjeta

    sSql = " exec stp_sel_ERS065_2017_SolicitarVistoAtencionSinTarjeta '" & pcUserSolicitud & "','" & pcCodAgeSolicitud & "','" & pcCtaCodCliente & "','" & pcPersCodCliente & "','" & pcOpeCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set oRS = oCon.Ejecutar(sSql)
    If Not oRS.BOF And Not oRS.EOF Then
        If oRS.RecordCount > 0 Then
            SolicitarVistoAtencionSinTarjeta = oRS!nRes
        Else
            SolicitarVistoAtencionSinTarjeta = 0
        End If
    Else
        SolicitarVistoAtencionSinTarjeta = 0
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ErrorSolicitarVistoAtencionSinTarjeta:
    Err.Raise Err.Number, "SolicitarVistoAtencionSinTarjeta ", Err.Description
    SolicitarVistoAtencionSinTarjeta = 0
End Function

'GIPO 20180815 SEGÚN MEMO
Public Sub RegistrarVistoDeUsuario(pcUserSolicitud As String, pcCodAgeSolicitud As String, pcCtaCodCliente As String, pcPersCodCliente As String, pcOpeCod As String)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim oRS As New ADODB.Recordset

    On Error GoTo ErrorRegistrarVistoDeUsuario

    sSql = " exec stp_upd_RegistrarVistoDeUsuario '" & pcUserSolicitud & "','" & pcCodAgeSolicitud & "','" & pcCtaCodCliente & "','" & pcPersCodCliente & "','" & pcOpeCod & "'"
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set oRS = oCon.Ejecutar(sSql)
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub
ErrorRegistrarVistoDeUsuario:
    Err.Raise Err.Number, "RegistrarVistoDeUsuario ", Err.Description
End Sub

Public Function ObtenerCuentasIntitucion(ByVal cPersCod As String) As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    Dim R As ADODB.Recordset
    On Error GoTo ErrorObtenerCodigo
    
    cSql = "exec stp_sel_ERS065_2017_ObtenerCuentasInstitucion '" & cPersCod & "'"
    
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(cSql)
    oCon.CierraConexion
    
    Set ObtenerCuentasIntitucion = R
    Set R = Nothing
    Exit Function
ErrorObtenerCodigo:
    Err.Raise Err.Number, "Error al ObtenerCuentasIntitucion", Err.Description
End Function
'END MARG******************************************
'ANDE 20180417 ERS021-2018 campaña mundialito
Public Function ExtraerCuentasCanceladas(ByVal pcPersCod As String) As ADODB.Recordset
    Dim cSql As String, oCon As New COMConecta.DCOMConecta
    On Error GoTo ErrorExtraerCuentasCanceladas
    Dim R As ADODB.Recordset
    cSql = "exec stp_sel_CuentasCanceladasXCliente '" & pcPersCod & "'"
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(cSql)
    oCon.CierraConexion
    Set ExtraerCuentasCanceladas = R
    Set R = Nothing
    Set oCon = Nothing
    Exit Function
ErrorExtraerCuentasCanceladas:
    Err.Raise Err.Number, "He detectado un problema. No se pudo extraer datos.", Err.Description
End Function

Public Sub RegistrarParticipanteCampanaAho(ByVal pcCtaCod As String, ByVal pnMonto As Currency, ByVal pnProceso As Integer, ByVal pnPunto As Integer, ByVal pcPersCod As String, ByVal pcMovNro As String)
    Dim cSql As String, oCon As New COMConecta.DCOMConecta

On Error GoTo ErrorRegistrarParticipanteCamp
    cSql = "stp_ins_SorteoCuentaParticipante '" & pcCtaCod & "','" & pnMonto & "','" & pnProceso & "','" & pnPunto & "','" & pcPersCod & "','" & pcMovNro & "'"
    oCon.AbreConexion
    oCon.Ejecutar cSql
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub
ErrorRegistrarParticipanteCamp:
    Err.Raise Err.Number, "He detectado un problema. No se pudo registrar al participante.", Err.Description
End Sub

Public Function ExtraerInfoParticipantesCamp(ByVal pcPersCod As String) As ADODB.Recordset
    Dim cSql As String, oCon As New COMConecta.DCOMConecta, R As ADODB.Recordset
    On Error GoTo ErrorExtraerInfoParticipantesCamp
    
    cSql = "exec stp_sel_SorteoInfoParticipante '" & pcPersCod & "'"
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(cSql)
    oCon.CierraConexion
    Set ExtraerInfoParticipantesCamp = R
    Set oCon = Nothing
    Set R = Nothing
    Exit Function
ErrorExtraerInfoParticipantesCamp:
    Err.Raise Err.Number, "He detectado un problema. No se pudo obtener información del participante"
End Function

Public Function getVerificarPersonaNatJur(ByVal cPersCod As String) As ADODB.Recordset
    Dim sSql As String, oConex As COMConecta.DCOMConecta
    sSql = "stp_sel_VerificarNaturalJuridica '" & cPersCod & "'"
    
    On Error GoTo ErrorAlObtenerPersonaNatJur
    Set oConex = New COMConecta.DCOMConecta
    oConex.AbreConexion
    Set getVerificarPersonaNatJur = oConex.CargaRecordSet(sSql)
    oConex.CierraConexion
    Set oConex = Nothing
    
    Exit Function
ErrorAlObtenerPersonaNatJur:
    Err.Raise Err.Number, "Error al tratar de verificar persona natural o jurídica.", Err.Description

End Function

Public Function CampaniaVigentes(ByVal pdFecSis As Date) As ADODB.Recordset
    Dim R As ADODB.Recordset, oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    On Error GoTo ErrorObtenerCampVigentes
    cSql = "stp_sel_ExisteCampaniaAho '" & Format(pdFecSis, "YYYY-MM-DD") & "'"
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(cSql)
    oCon.CierraConexion
    Set CampaniaVigentes = R
    Set R = Nothing
    Set oCon = Nothing
    Exit Function
ErrorObtenerCampVigentes:
    Err.Raise Err.Number, "Error al tratar de obtener campañas vigentes.", Err.Description
End Function

Public Function VerificarCtaCodCamp(ByVal cCtaCod As String) As ADODB.Recordset
    Dim cSql As String, oCon As New COMConecta.DCOMConecta, rData As ADODB.Recordset
    
    On Error GoTo ErrorVerificarCtaCodCamp
        cSql = "stp_sel_VerificarCtaEnCampania '" & cCtaCod & "'"
        oCon.AbreConexion
        Set rData = oCon.CargaRecordSet(cSql)
        oCon.CierraConexion
        Set VerificarCtaCodCamp = rData
        Set rData = Nothing
        Set oCon = Nothing
    Exit Function
ErrorVerificarCtaCodCamp:
    Err.Raise Err.Number, "No se pudo verificar la cuenta, comuniques con TI-desarrollo.", Err.Description
End Function

Public Function ExtraerPuntoCuenta(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim cSql As String, oCon As New COMConecta.DCOMConecta, rData As ADODB.Recordset
    On Error GoTo ErrorExtraePuntoCuenta
    cSql = "stp_sel_PuntoCuenta '" & pcCtaCod & "'"
    oCon.AbreConexion
        Set rData = oCon.CargaRecordSet(cSql)
        Set ExtraerPuntoCuenta = rData
    oCon.CierraConexion
    Set oCon = Nothing
    Set rData = Nothing
    Exit Function
ErrorExtraePuntoCuenta:
    Err.Raise Err.Number, "Error al extraer punto de la cuenta, comuniques con TI-desarrollo.", Err.Description
End Function
'***************** END ANDE **********************

'--=========================================
'--INICIO EAAS20180511 SEGUN ERS TI-006-2018
'--=========================================
Public Function ObtieneCCI(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim cSql As String, oCon As New COMConecta.DCOMConecta, rData As ADODB.Recordset
    On Error GoTo ErrorObtieneCCI
    cSql = "SEL_CCE_ObtieneCCIDesdeCuenta '" & pcCtaCod & "'"
    oCon.AbreConexion
        Set rData = oCon.CargaRecordSet(cSql)
        Set ObtieneCCI = rData
    oCon.CierraConexion
    Set oCon = Nothing
    Set rData = Nothing
    Exit Function
ErrorObtieneCCI:
    Err.Raise Err.Number, "Error al Obtener CCI, comuniques con TI-desarrollo.", Err.Description
End Function
'--=========================================
'--INICIO EAAS20180511 SEGUN ERS TI-006-2018
'--=========================================

'*************** ANDE 20171107 habilitacion de operaciones  -- iniciativa TI-Desarrollo ****************************
Public Sub RegistrarLogHabilitacion(ByVal pnId As Integer, ByVal pcUserH As String, ByVal pdFechaOpe As Date, ByVal pnCondicion As Integer)
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    On Error GoTo ErrorLogHablitacion
    cSql = "exec stp_ins_logHabilitacionOpe '" & pnId & "','" & pcUserH & "','" & pdFechaOpe & "','" & pnCondicion & "'"
    oCon.AbreConexion
        oCon.Ejecutar (cSql)
    oCon.CierraConexion
    Exit Sub
ErrorLogHablitacion:
    Err.Raise Err.Number, "He detectado un problema. No se pudo registrar la pista.", Err.Description
End Sub


Public Function ObtenerOperacionesHabilitar() As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    Dim R As ADODB.Recordset
    On Error GoTo ErrorObtenerOperacionesH
    cSql = "exec stp_sel_ListarOperacionesH"
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(cSql)
    oCon.CierraConexion
    Set ObtenerOperacionesHabilitar = R
    Exit Function
ErrorObtenerOperacionesH:
    Err.Raise Err.Number, "He detectado un problema. No se pudo obtener las operaciones. Por favor intentelo de nuevo más tarde.", Err.Description
End Function

Public Function ObtenerUsuarios_RF_AS_SUP(ByVal pcUser As String) As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    Dim R As ADODB.Recordset
    cSql = "exec stp_sel_UsuarioActivo '" & pcUser & "'"
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(cSql)
    oCon.CierraConexion
    Set ObtenerUsuarios_RF_AS_SUP = R
    Set R = Nothing
End Function

Public Function ObtenerOperacionesHabilitadasxUsuario(ByVal pcUser As String, ByVal pdFecSis As Date) As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    Dim R As New ADODB.Recordset
    On Error GoTo ErrorResumenOpe
    cSql = "exec stp_sel_OperacionesHabilitadasxUsuario '" & pcUser & "','" & Format(pdFecSis, "YYYYMMDD") & "'"
    oCon.AbreConexion
    Set R = oCon.CargaRecordSet(cSql)
    oCon.CierraConexion
    Set ObtenerOperacionesHabilitadasxUsuario = R
    Exit Function
ErrorResumenOpe:
    Err.Raise Err.Number, "He detectado un problema. No se pudo obtener las operaciones habilitadas del usuario.", Err.Description
End Function

Public Function RegistrarHabilitacionDeOperacion(ByVal pcUserH As String, ByVal pcOpeCod As String, ByVal pcHabilitacionSist As String, Optional ByVal pcUserQH As String = "", Optional ByVal pnCantOpe As Integer = 0, Optional ByVal pnEstado As Integer = 0) As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim R As ADODB.Recordset
    Dim cSql As String
    cSql = "exec stp_ins_upd_HablitacionOpcionesOpe '" & pcUserQH & "','" & pcUserH & "','" & pcOpeCod & "','" & pnCantOpe & "','" & pcHabilitacionSist & "','" & pnEstado & "'"
    oCon.AbreConexion
        Set R = oCon.CargaRecordSet(cSql)
    oCon.CierraConexion
    Set RegistrarHabilitacionDeOperacion = R
End Function

Public Function RegistrarActividad(ByVal pcOpeCod As String, ByVal pcUser As String, ByVal pdFecha As Date) As Integer
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String, R As Recordset
    
    On Error GoTo ErrorRegistrarActividad
    cSql = "exec stp_upd_ActividadOperaciones '" & pcOpeCod & "','" & pcUser & "','" & Format(pdFecha, "YYYYMMDD") & "'"
    oCon.AbreConexion
        Set R = oCon.Ejecutar(cSql)
        If Not (R.BOF And R.EOF) Then
            RegistrarActividad = R!Resultado
        Else
            RegistrarActividad = 1
        End If
    oCon.CierraConexion
    Exit Function
ErrorRegistrarActividad:
    RegistrarActividad = 1
End Function
Public Function VerificarOperaciones(ByVal pcUser As String, ByVal pdFecha As Date, ByVal pcOpeCod As String) As ADODB.Recordset
    Dim oCon As New COMConecta.DCOMConecta
    Dim cSql As String
    Dim R As ADODB.Recordset
    On Error GoTo ErrorhabilitarOperacion
        cSql = "exec stp_sel_VerificarOperaciones '" & pcUser & "','" & Format(pdFecha, "YYYYMMDD") & "','" & pcOpeCod & "'"
        oCon.AbreConexion
            Set R = oCon.CargaRecordSet(cSql)
        oCon.CierraConexion
        Set VerificarOperaciones = R
        Set oCon = Nothing
        Set R = Nothing
    Exit Function
ErrorhabilitarOperacion:
    Err.Raise Err.Number, "He detectado un problema. No se pudo registrar la habilitación. Por favor intentelo de nuevo más tarde.", Err.Description
End Function
'*************************CTI3 TI-ERS082-2018
Public Function ConsultaTpoCta(ByVal pcCtaCod As String) As ADODB.Recordset
    Dim conSql As String
    Dim oCon As New COMConecta.DCOMConecta
    Dim rData As ADODB.Recordset
    On Error GoTo ErrorConsulta
    conSql = "stp_sel_mostrartpocta '" & pcCtaCod & "'"
    oCon.AbreConexion
        Set rData = oCon.CargaRecordSet(conSql)
        Set ConsultaTpoCta = rData
    oCon.CierraConexion
    Set oCon = Nothing
    Set rData = Nothing
    Exit Function
ErrorConsulta:
    Err.Raise Err.Number, "Error al obtener el Tipo de Cuenta, comuniquese con TI-desarrollo.", Err.Description
End Function
'***********************************************


'APRI20190601 RFC1902040001
Public Sub GuardarMotivoCancelacion(ByVal pcCtaCod As String, ByVal pnMotivo As Integer, ByVal pcGlosa As String)
    Dim cSql As String, oCon As New COMConecta.DCOMConecta

    On Error GoTo ErrorGuardarMotivoCancelacion
    cSql = "stp_upd_CapMotivoCancelacion '" & pcCtaCod & "'," & pnMotivo & ",'" & pcGlosa & "'"
    oCon.AbreConexion
    oCon.Ejecutar cSql
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub
ErrorGuardarMotivoCancelacion:
    Err.Raise Err.Number, "He detectado un problema. No se pudo GuardarMotivoCancelacion.", Err.Description
End Sub
'END APRI
'APRI20190109 ERS077-2018
Public Function ObtenerCapValorComision(ByVal pcCtaCod As String, ByVal pnCantOpe As Integer, ByVal pnTpoOpe As Integer, ByVal pnTpoCanal As Integer) As Currency
Dim sSql As String
Dim oDCOMConecta  As New COMConecta.DCOMConecta
Dim rsCom As New ADODB.Recordset

oDCOMConecta.AbreConexion
sSql = "exec stp_sel_ValorCapOpeComision '" & pcCtaCod & "'," & pnCantOpe & "," & pnTpoOpe & "," & pnTpoCanal
Set rsCom = oDCOMConecta.CargaRecordSet(sSql)
oDCOMConecta.CierraConexion

If Not (rsCom.BOF And rsCom.EOF) Then
    ObtenerCapValorComision = rsCom!nValor
Else
    ObtenerCapValorComision = 0
End If

Set oDCOMConecta = Nothing
Set rsCom = Nothing
End Function
Public Function CargarCuentasMantCuentaPendiente(Optional ByVal psCodPers As String = "", Optional psCtaCod As String = "") As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    sSql = "exec stp_sel_CuentaPendienteMantenimiento '" & Trim(psCodPers) & "','" & Trim(psCtaCod) & "'"
    Set CargarCuentasMantCuentaPendiente = oCon.CargaRecordSet(sSql)
      
    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function AplicaComunicacionCaptaciones(ByVal cPersCod As String, ByVal cDoi As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
        lsSQL = "Exec stp_sel_AplicaComunicacionCaptaciones '" & cPersCod & "','" & cDoi & "'"
        
        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set AplicaComunicacionCaptaciones = lrDatos
        Set lrDatos = Nothing
End Function
Public Function ObtenerFechaAplicacionTarifario() As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

        lsSQL = "Exec stp_sel_ObtenerFechaAplicacionTarifario "

        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set ObtenerFechaAplicacionTarifario = lrDatos
        Set lrDatos = Nothing
End Function
Public Sub ActualizarFechaComunicacion(ByVal pcPersCod As String)
    Dim cSql As String, oCon As New COMConecta.DCOMConecta

    On Error GoTo ErrorActualizarFechaComunicacion
    cSql = "Exec stp_upd_ActualizarFechaComunicacion '" & pcPersCod & "'"
    oCon.AbreConexion
    oCon.Ejecutar cSql
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub
ErrorActualizarFechaComunicacion:
    Err.Raise Err.Number, "He detectado un problema. No se pudo ActualizarFechaComunicacion.", Err.Description
End Sub
'END APRI

'JIPR20190916 INICIO
Public Function ObtieneDatosCuentas(ByVal pcPersCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

        lsSQL = "Exec stp_sel_ObtieneDatosCuentas  '" & pcPersCod & "'"

        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set ObtieneDatosCuentas = lrDatos
        Set lrDatos = Nothing
End Function

Public Function DatosPersCuenta(ByVal pcPersCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

        lsSQL = "Exec stp_sel_DatosPersCuenta  '" & pcPersCod & "'"

        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set DatosPersCuenta = lrDatos
        Set lrDatos = Nothing
End Function

Public Sub RegistraTarjetaCuenta(ByVal lsNumTarj As String, ByVal lsCuenta As String, ByVal lsFecha As Date, ByVal lnPrior As Integer, ByVal lnRelac As Integer, ByVal lnConsulta As Integer, ByVal lnRetiro As Integer)
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
        
    oCon.AbreConexion
    sSql = "exec stp_ins_RegistraTarjetaCuenta '" & lsNumTarj & "','" & lsCuenta & "','" & Format(lsFecha, "yyyymmdd") & "'," & lnPrior & "," & lnRelac & "," & lnConsulta & "," & lnRetiro & ""
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub RegistraHistorPrioridad(ByVal lsNumTarj As String, ByVal lsCuenta As String, ByVal lsFecha As Date, ByVal lnPrior As Integer)
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
        
    oCon.AbreConexion
    sSql = "exec stp_ins_RegistraHistorPrioridad '" & lsNumTarj & "','" & lsCuenta & "','" & Format(lsFecha, "yyyymmdd") & "', " & lnPrior & ""
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub ActualizaTarjetaCuenta(ByVal lsNumTarj As String, ByVal lsCuenta As String, ByVal lnRelac As Integer, ByVal lnPrior As Integer, ByVal lnConsulta As Integer, ByVal lnRetiro As Integer)
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
        
    oCon.AbreConexion
    sSql = "exec stp_upd_ActualizaTarjetaCuenta '" & lsNumTarj & "','" & lsCuenta & "'," & lnRelac & "," & lnPrior & "," & lnConsulta & "," & lnRetiro & ""
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Sub DesafiliarTarjetaCuenta(ByVal lsNumTarj As String, ByVal lsCuenta As String)
    Dim sSql As String
    Dim oCon  As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
        
    oCon.AbreConexion
    sSql = "exec stp_upd_DesafiliarTarjetaCuenta '" & lsNumTarj & "','" & lsCuenta & "'"
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Function DatosPDF(ByVal lsNumTarj As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

        lsSQL = "Exec stp_sel_DatosPDF  '" & lsNumTarj & "'"

        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set DatosPDF = lrDatos
        Set lrDatos = Nothing
End Function

Public Function DocumentoPDF() As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

        lsSQL = "Exec stp_sel_DOCUMENTOPDF "

        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set DocumentoPDF = lrDatos
        Set lrDatos = Nothing
End Function

'JIPR20190916 FIN
'CTI4 ERS0112020
Public Function ObtieneDOI_Titular(ByVal psCuenta As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

        lsSQL = "Exec stp_sel_ObtieneDOI_Titular '" & psCuenta & "'"

        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set ObtieneDOI_Titular = lrDatos
        Set lrDatos = Nothing
End Function
Public Function ObtieneDatosTrBanco(ByVal pnMovNro As Long) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset

        lsSQL = "Exec stp_sel_ObtieneDatosTrBanco " & pnMovNro

        Set oConec = New COMConecta.DCOMConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set ObtieneDatosTrBanco = lrDatos
        Set lrDatos = Nothing
End Function
'end CTI4
'APRI20210621 ERS031-2021
Public Function AplicaDineroNuevoCaptaciones(ByVal psPersCod As String, ByVal pnMonto As Currency) As Boolean
Dim sSql As String
Dim rsDineroNuevo As New ADODB.Recordset

sSql = "exec stp_sel_DineroNuevo '" & psPersCod & "', " & pnMonto & ""
Set rsDineroNuevo = dbCmact.Execute(sSql)

If Not (rsDineroNuevo.BOF And rsDineroNuevo.EOF) Then
    AplicaDineroNuevoCaptaciones = rsDineroNuevo!nAplica
Else
    AplicaDineroNuevoCaptaciones = 0
End If

Set rsDineroNuevo = Nothing
End Function
'END APRI

