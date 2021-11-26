VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMGeneraAcl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim Conecta  As COMConecta.DCOMConecta

Public Sub Coloc_ACL_CartaFianza()

Dim conexiondbf As New ADODB.Connection
Dim cad As String
Dim lsArchivo As String
Dim sql8 As String  ' CartaFianza
Dim Sql As String
Dim m As Double
Dim Rs As New ADODB.Recordset
Dim sSql As String
Dim RCD As NCOMRCDReportes
Dim lsServerCons As String
Set RCD = New NCOMRCDReportes
lsServerCons = RCD.GetServerConsol

lsArchivo = App.Path & "\SPOOLER\" '+ vNombreFile
ChDir lsArchivo

'Set fs = New Scripting.FileSystemObject
'If fs.FolderExists(lsArchivo) = False Then
'    fs.CreateFolder (lsArchivo)
'End If
 
Set conexiondbf = New ADODB.Connection
conexiondbf.Open "DSN=dsnAuditoriadbf;uid=;pwd=;sourceDB=" & lsArchivo & ";SOURCETYPE=DBF"
sql8 = " CREATE TABLE CartaF ( " & _
       "cCodCta  C(12) NOT NULL ," & _
       "cEstado I(4) NOT NULL ," & _
       "dAsigna T(8) NULL ," & _
       "cCondCF I(4) NULL ," & _
       "cModalid I(4) NOT NULL ," & _
       "dFecApr T(8) NULL ," & _
       "nMontoA N(6,2) NULL ," & _
       "dVencA T(8) NULL ," & _
       "cCodAna C(4) NOT NULL ," & _
       "cApodera C(4) NOT NULL ," & _
       "dFecVig T(8) NULL " & _
      " ) "
conexiondbf.Execute sql8

Sql = Sql & " Select  cCtaCod,   nPrdEstado,    dAsignacion,"
Sql = Sql & "    nCondCF,   nModalidad, dFecApr,"
Sql = Sql & "    nMontoApr, dVencApr,   cCodAnalista,"
Sql = Sql & "    cApoderado , dFecVig"
Sql = Sql & " From " & lsServerCons & "CartaFianzaConsol"
Sql = Sql & " Where nPrdEstado =" & gColocEstVigMor
Sql = Sql & " Order by cCtaCod"

Conecta.AbreConexion
    Set Rs = Conecta.CargaRecordSet(Sql)
Conecta.CierraConexion

'If Rs.EOF And Rs.BOF Then
'    MsgBox "NO EXISTEN DATOS POR CONSOLIDAR"
'Else
 'With Rs
 While Not Rs.EOF

        sSql = " INSERT INTO CartaF (cCodCta, cEstado, dAsigna, cCondCF, cModalid," & _
               " dFecApr, nMontoA, dVencA, cCodAna, cApodera, dFecVig )" & _
               " VALUES  "
        With Rs
        sSql = sSql & " ('" & !cCtaCod & "'," & !nPrdEstado & ",ctod('" & Format(!dAsignacion, "mm/dd/yyyy") & "') , " & !nCondCF & "," & !nModalidad & ","
        sSql = sSql & " ctod('" & Format(!dAsignacion, "mm/dd/yyyy") & "') , " & !nMontoApr & ", ctod('" & Format(!dVencApr, "mm/dd/yyyy") & "') ,'" & !cCodAnalista & "', '" & !cApoderado & " ',"
        sSql = sSql & " ctod('" & Format(!dFecVig, "mm/dd/yyyy") & "') ) "
        End With
        conexiondbf.Execute sSql
        Rs.MoveNext
 Wend
 
Set Rs = Nothing
Set RCD = Nothing
End Sub

Public Sub Coloc_ACL_Creditos()
Dim fs As Scripting.FileSystemObject
Dim Rs As New ADODB.Recordset
Dim conexiondbf As New ADODB.Connection
Dim cad As String
Dim lsArchivo As String
Dim sql5 As String  ' CREDITO
Dim Sql As String
Dim m As Double
Dim sSql As String
Dim lsCreditosVigentes As String
Dim lsPignoraticio As String
Dim RCD As NCOMRCDProceso
Dim lsServerCons As String
Set RCD = New NCOMRCDProceso
'lsServerCons = Rcd.GetServerConsol
lsCreditosVigentes = gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc & "," & gColocEstRefNorm & "," & gColocEstRefMor & "," & gColocEstRefVenc
lsPignoraticio = gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov


lsArchivo = App.Path & "\SPOOLER\" '+ vNombreFile
ChDir lsArchivo

'Set fs = New Scripting.FileSystemObject
'If fs.FolderExists(lsArchivo) = False Then
'    fs.CreateFolder (lsArchivo)
'End If
    
Set conexiondbf = New ADODB.Connection
conexiondbf.Open "DSN=dsnAuditoriadbf;uid=;pwd=;sourceDB=" & lsArchivo & ";SOURCETYPE=DBF"
        
sql5 = " CREATE TABLE Credito (" & _
       "cCodCta C(18) NOT NULL ," & _
       "cEstado I(4) NOT NULL , " & _
       "dAsigna T(8) NOT NULL , " & _
       "nMontoSol N(6,2) NOT NULL ," & _
       "cNumFuente C(8) NOT NULL , " & _
       "cCondCre I(4) NOT NULL ," & _
       "cDestCre I(4) NOT NULL , " & _
       "cCodLin C(11) NULL , " & _
       "dFecApr T(8) NULL , " & _
       "nTasaInt N(6,2) NULL ," & _
       "nMontoApr N(6,2) NULL ," & _
       "nCuotasApr N(6,2) NULL , " & _
       "nPlazoApr I(4) NULL , " & _
       "nGraciaApr I(4) NULL , " & _
       "nCuotaApr N(6,2) NULL ," & _
       "ctipCuota I(4) NULL , "
sql5 = sql5 & "" & _
       "cCodAnali C(4) NULL , " & _
       "cApoderado C(4) NULL , " & _
       "dFecVig T(8) NULL , " & _
       "nMontoDes N(6,2) NULL , " & _
       "nSaldoCap N(6,2) NULL ," & _
       "nDiasAtr I(4) NULL , " & _
       "nIntMorCal N(6,2) NULL , " & _
       "nProxCuo I(4) NULL , " & _
       "dUltPago T(8) NULL , " & _
       "cRefinan C(1) NULL , " & _
       "cNota1 C(1) NULL , "
sql5 = sql5 & "" & _
       "nDiaFijo I(4) NULL , " & _
       "nNroRep I(4) NULL ," & _
       "cNomCli C(40) NULL  " & _
       " ) "
conexiondbf.Execute sql5

'----------------

'sql = "Select  CC.cCtaCod , Pro.nPrdEstado,"
'sql = sql & " dAsigna=(Select  dPrdEstado from ColocacEstado CE where CE.cCtaCod=CC.cCtaCod and CE.nPrdEstado=2000),"
'sql = sql & " nMontoSol=(Select  nMonto from ColocacEstado CE where CE.cCtaCod=CC.cCtaCod and CE.nPrdEstado=2000),"'
'sql = sql & " (Select CFI.cNumFuente From ColocFteIngreso CFI Where CFI.cCtaCod=CC.cCtaCod"
'sql = sql & " and CFI.dPersEval=(Select Max(dPersEval) from ColocFteIngreso Where cCtaCod=CC.cCtaCod)) as nFuenteIngreso,"
'sql = sql & " CC.nColocCondicion, CC.nColocDestino, Co.cLineaCred,"
'sql = sql & " (Select dPrdEstado from ColocacEstado Where cCtaCod=CC.cCtaCod and nPrdEstado=2002) as dFecApr,"
'sql = sql & " Pro.nTasainteres,"
'sql = sql & " (Select nMonto from ColocacEstado Where cCtaCod=CC.cCtaCod and nPrdEstado=2002) as nMontoApr,"
'sql = sql & "  (Select nCuotas from ColocacEstado Where cCtaCod=CC.cCtaCod and nPrdEstado=2002) as nCuotasApr,"
'sql = sql & " (Select nPlazo from ColocacEstado Where cCtaCod=CC.cCtaCod and nPrdEstado=2002) as nPlazoApr,"
'sql = sql & "  (Select nPeriodoGracia from ColocacEstado Where cCtaCod=CC.cCtaCod and nPrdEstado=2002) as nGraciaApr,"
'sql = sql & " nCuotaApr=0,"
'sql = sql & " ctipCuota=0,"
'sql = sql & "  (Select nPeriodoFechaFija from ColocacEstado Where cCtaCod=CC.cCtaCod and nPrdEstado=2002) as cPeriodo,"
'sql = sql & " (select RH.cUser from RRHH RH where RH.cPersCod=PP1.cPersCod) as cCodAnalista,"
'sql = sql & " (select RH.cUser from RRHH RH where RH.cPersCod=PP3.cPersCod) as Apoderado,"
'sql = sql & " Co.dVigencia,"
'sql = sql & " (Select sum(abs(nMonto)) from MovColDet Where cOpeCod=100101 and cCtaCod=CC.cCtaCod and nPrdConceptoCod=1000 ) AS nMontoDes,"
'sql = sql & "   Pro.nSaldo,"
'sql = sql & " (CC.nDiasAtraso) nDiasAtr,"
'sql = sql & " (Select sum(nmonto) from ColocCalendDet where cCtaCod=CC.cCtaCod and nPrdConceptoCod=1101 and CC.nNroCalen=nNroCalen)as nInteresMoraCal,"
'sql = sql & "  Co.dVigencia,"
'sql = sql & " CC.nNroProxCuota as nProxCou,"
'sql = sql & " Pro.dPrdEstado dUltPago,"
'sql = sql & "  (Select nColocNota from ColocCalificacionAnalista where cCtaCod = CC.cCtaCod and dColocNotaFecha="
'sql = sql & " (Select Max(dColocNotaFecha) from ColocCalificacionAnalista where cCtaCod = CC.cCtaCod ) ) as cNota1,"
'sql = sql & " (Select nPeriodoFechaFija from ColocacEstado Where cCtaCod=CC.cCtaCod and nPrdEstado=2002) as nDiaFijo,"
'sql = sql & "  (Select (Max(nNroCalen)-Min(nNroCalen)) from ColocCalendDet where cCtaCod=CC.cCtaCod) as nNroRep,"
'sql = sql & " (Select cPersNombre from Persona where cpersCod=PP1.cPersCod) as cNomCli"
'sql = sql & " from ColocacCred CC"
'sql = sql & " Inner Join Colocaciones Co on Co.cCtaCod=CC.cCtaCod"
'sql = sql & " Inner Join Producto Pro on Pro.cCtaCod=CC.cCtaCod"
'sql = sql & " Inner Join ProductoPersona PP1 on PP1.cCtaCod=CC.cCtaCod"
'sql = sql & " Inner Join ProductoPersona PP2 on PP2.cCtaCod=CC.cCtaCod"
'sql = sql & " Inner Join ProductoPersona PP3 on PP3.cCtaCod=CC.cCtaCod"
'sql = sql & " Where Pro.nPrdEstado in (2020,2021,2022,2030,2031,2032)  and PP1.nPrdPersRelac=20"
'sql = sql & " and PP2.nPrdPersRelac=28 and PP3.nPrdPersRelac=29"'
'************************************
Sql = "SELECT c.cCtaCod, nPrdEstado , dAsignacion as dAsigna , nMontoSol,  cNumFuente,"
Sql = Sql & "         nCondCre, nDestCre,  cLineaCred as cCodLin, dFecApr, nTasaInt,"
Sql = Sql & "          nMontoApr, nCuotasApr, nPlazoApr, nGraciaApr, nCuotaApr,ntipCuota,"
Sql = Sql & "           cCodAnalista as cCodAnali, cApoderado, dFecVig, nMontoDesemb as nMontoDes,"
Sql = Sql & "        nSaldoCap, nDiasAtraso as nDiasAtr,  nMoraCalc, nNroProxCuota as nProxCuo,"
Sql = Sql & "        dFecUltPago as dUltPago, cNota1,  nNroRepro,  cRefinan, nDiaFijo, "
Sql = Sql & "   cNomCli = (Select IsNull(Max(P.cPersNombre),'')"
Sql = Sql & "                     From Persona P join " & lsServerCons & "ProductoPersonaConsol pc on p.cPersCod = p.cPersCod"
Sql = Sql & "             Where pc.cCtaCod = c.cCtaCod and pc.nPrdPersRelaC =20)"
Sql = Sql & " FROM " & lsServerCons & "CreditoConsol c"
Sql = Sql & " Where ( nPrdEstado in (" & lsCreditosVigentes & "," & lsPignoraticio & ") or (nPrdEstado =" & gColocEstRecVigJud & ") )"
Sql = Sql & " And nSaldoCap > 0"

Conecta.AbreConexion
    Set Rs = Conecta.CargaRecordSet(Sql)
Conecta.CierraConexion
'*************************************
 If Rs.EOF And Rs.BOF Then
    MsgBox "NO EXISTEN DATOS POR CONSOLIDAR"
 Else
    With Rs
    Do While Not .EOF
        sSql = "INSERT INTO Credito (cCodCta,cEstado,dAsigna,nMontoSol, " & _
        "cNumFuente,cCondCre,cDestCre,cCodLin, " & _
        "dFecApr , nTasaInt, nMontoApr, nCuotasApr, " & _
        "nPlazoApr, nGraciaApr, nCuotaApr, ctipCuota, " & _
        "cCodAnali ,cApoderado, " & _
        "dFecVig, " & _
        "nMontoDes," & _
        "nSaldoCap, " & _
        "nDiasAtr," & _
        "nIntMorCal,nProxCuo," & _
        "dUltPago, " & _
        "cNota1, " & _
        "nDiaFijo,nNroRep,cNomCli,cRefinan ) "
        
        sSql = sSql & "VALUES "
        sSql = sSql & " ('" & !cCtaCod & "'," & !nPrdEstado & ",ctod('" & Format(!dAsigna, "mm/dd/yyyy") & "')," & !nMontoSol & ",'" & !cNumFuente & " '," & !nCondCre & "," & !nDestCre & ",'" & !cCodLin & "'"
        sSql = sSql & ", ctod('" & Format(!dFecApr, "mm/dd/yyyy") & "')," & !nTasaInt & "," & !nMontoApr & "," & !nCuotasApr
        sSql = sSql & "," & !nPlazoApr & "," & !nGraciaApr & "," & !nCuotaApr & "," & !nTipCuota
        sSql = sSql & ",'" & !cCodAnali & "','" & !cApoderado & "', ctod('" & Format(!dFecVig, "mm/dd/yyyy") & "')"
        sSql = sSql & "," & IIf(IsNull(!nMontoDes), 0, !nMontoDes)
        sSql = sSql & "," & !nSaldoCap & "," & !nDiasAtr
        sSql = sSql & "," & IIf(IsNull(!nMoraCalc), 0, !nMoraCalc)
        sSql = sSql & "," & IIf(IsNull(!nProxCuo), 0, !nProxCuo)
        sSql = sSql & "," & " ctod('" & Format(!dUltPago, "mm/dd/yyyy") & "')"
        sSql = sSql & ",'" & !cNota1 & "'"
        sSql = sSql & "," & IIf(IsNull(!nDiaFijo), 0, !nDiaFijo) & "," & IIf(IsNull(!nNroRepro), 0, !nNroRepro)
        sSql = sSql & ",'" & !cNomCli & "','" & !cRefinan & "') "
        conexiondbf.Execute sSql
        .MoveNext
    Loop
    End With
  End If
  Set Rs = Nothing
  Set RCD = Nothing
End Sub


Public Function GetPlazoFijo() As ADODB.Recordset
Dim sqlPF As String
Dim RCD As NCOMRCDProceso
Dim lsServerCons As String
Set RCD = New NCOMRCDProceso
'lsServerCons = Rcd.GetServerConsol
Dim Rs As ADODB.Recordset
sqlPF = " SELECT T.cCtaCod, T.dApertura, T.nPlazo, T.nCapital, T.nSaldCnt, T.nTasa,T.dUltMov" _
        & " FROM (Select cCtaCod, Convert(varchar(10),dAperPF,101) dApertura ,nPlazo,nCapital," _
        & " nSaldCntPF nSaldCnt,nTasaIntPF nTasa,Convert(varchar(10), dUltMovPF,101) dUltMov" _
        & " from " & lsServerCons & "PlazoFijoConsol Where nEstCtaPF NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ")" _
        & " Union" _
        & " Select cCtaCod,Convert(varchar(10),dAperCTS,101) dApertura, nPlazo = 720," _
        & " nCapitalCTS nCapital, nSaldCntCTS nSaldCnt, nTasaIntCTS nTasa," _
        & " Convert(varchar(10),dUltMovCTS,101) dUltMov" _
        & " from " & lsServerCons & "CTSConsol Where nEstCtaCTS NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ")) T order by T.cCtaCod "
 Conecta.AbreConexion
Set Rs = Conecta.CargaRecordSet(sqlPF)
Set GetPlazoFijo = Rs
Set Rs = Nothing
Set RCD = Nothing
End Function

Public Sub GenerarPF_Movimiento(ByVal dFecSis As Date)
Dim conexiondbf As ADODB.Connection
Dim Rs As ADODB.Recordset
Dim Conecta As COMConecta.DCOMConecta

Dim Sql_Dbf As String
Dim sqlPF As String
Dim SqlTabla As String
Dim vNomser As String
Dim vNomBase As String
Dim vRFileAuditor As String
Dim vFechaSist As String
Dim vMes As String
Dim vYear As String
Dim vNom_cadena As String
Dim lsArchivo As String
Dim SqlPFMov As String
Dim RCD As NCOMRCDProceso
Dim lsServerCons As String
Set RCD = New NCOMRCDProceso
'lsServerCons = Rcd.GetServerConsol
lsArchivo = App.Path & "\SPOOLER\"
ChDir lsArchivo

vMes = Mid(dFecSis, 4, 2) 'vFechaSist
vYear = Mid(dFecSis, 7, 4) 'vFechaSist
vNom_cadena = "PFMov_" & vYear & vMes & ".Dbf"

'sqlPF = "Select * from ("
'sqlPF = sqlPF & " Select CPF.cCtaCod, convert(datetime,convert(varchar(8),M.cMovNro,101)) as dFecha,"
'sqlPF = sqlPF & " OP.cOpeDesc, abs(MC.nMonto) as nMonto"
'sqlPF = sqlPF & " from CaptacPlazoFijo CPF"
'sqlPF = sqlPF & " Inner Join Producto Pro on Pro.cCtaCod=CPF.cCtaCod"
'sqlPF = sqlPF & " Inner Join MovCap MC on CPF.cCtaCod=MC.cCtaCod"
'sqlPF = sqlPF & " Inner Join OpeTPO OP on MC.cOpeCod=OP.cOpeCod"
'sqlPF = sqlPF & " Inner Join Mov M on M.nMovNro=MC.nMovNro"
'sqlPF = sqlPF & " where Pro.nPrdEstado  not in (1300,1400)"
'sqlPF = sqlPF & " Union"
'sqlPF = sqlPF & " Select CPF.cCtaCod, convert(datetime,convert(varchar(8),M.cMovNro,101)) as dFecha,"
'sqlPF = sqlPF & " OP.cOpeDesc, abs(MC.nMonto) as nMonto"
'sqlPF = sqlPF & " from CaptacCTS CPF"
'sqlPF = sqlPF & " Inner Join Producto Pro on Pro.cCtaCod=CPF.cCtaCod"
'sqlPF = sqlPF & " Inner Join MovCap MC on CPF.cCtaCod=MC.cCtaCod"
'sqlPF = sqlPF & " Inner Join OpeTPO OP on MC.cOpeCod=OP.cOpeCod"
'sqlPF = sqlPF & " Inner Join Mov M on M.nMovNro=MC.nMovNro"
'sqlPF = sqlPF & " where Pro.nPrdEstado  not in (1300,1400)"
'sqlPF = sqlPF & " ) T Order by T.cCtaCod, t.dFecha"


'SqlPFMov = "SELECT TA.cCodCta, TA.dFecha, TA.Operacion, TA.nMonto " _
'           & " FROM (Select T.cCodCta, Convert(varchar(10),dFecTran,101) dFecha," _
'           & " LTRIM(RTRIM(O.cNomOpe)) Operacion, ABS(T.nMonTran) nMonto" _
'           & " From DBComunes..Operacion O INNER JOIN TRansAhoConsol T INNER JOIN PlazoFijoConsol PF" _
'           & " ON T.cCodCta = PF.cCodCta ON O.cCodOpe = T.cCodOpe Where cEstCtaPF NOT IN ('C','U')" _
'           & " Union " _
'           & " Select T.cCodCta, Convert(varchar(10),dFecTran,101) dFecha, LTRIM(RTRIM(O.cNomOpe))," _
'           & " ABS(T.nMonTran) nMonto" _
'           & " From DBComunes..Operacion O INNER JOIN TRansAhoConsol T INNER JOIN CTSConsol CTS" _
'           & " ON T.cCodCta = CTS.cCodCta ON O.cCodOpe = T.cCodOpe Where cEstCtaCTS NOT IN ('C','U')" _
'           & " ) TA ORDER BY TA.cCodCta,TA.dFecha"


SqlPFMov = " SELECT TA.cCtaCod, TA.dFecha, TA.Operacion, TA.nMonto "
SqlPFMov = SqlPFMov & " FROM ( "
SqlPFMov = SqlPFMov & "  Select T.cCtaCod, convert(datetime,Convert(varchar(8),M.cMovNro,101),101) dFecha, "
SqlPFMov = SqlPFMov & "  LTRIM(RTRIM(O.cOpeDesc)) Operacion, ABS(T.nMonto) nMonto "
SqlPFMov = SqlPFMov & "  From OpeTpo O"
SqlPFMov = SqlPFMov & "  INNER JOIN Mov M on M.cOpeCod=O.cOpeCod "
SqlPFMov = SqlPFMov & "  INNER JOIN MovCap T on T.nMovNro=M.nMovNro "
SqlPFMov = SqlPFMov & "  INNER JOIN " & lsServerCons & "PlazoFijoConsol PF  on PF.cCtaCod=T.cCtaCod "
SqlPFMov = SqlPFMov & "  Where nEstCtaPF NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ")"
SqlPFMov = SqlPFMov & "  Union "
SqlPFMov = SqlPFMov & "  Select T.cCtaCod, convert(datetime,Convert(varchar(8),M.cMovNro,101),101) dFecha, "
SqlPFMov = SqlPFMov & "  LTRIM(RTRIM(O.cOpeDesc)) Operacion, ABS(T.nMonto) nMonto "
SqlPFMov = SqlPFMov & "  From OpeTpo O "
SqlPFMov = SqlPFMov & "  INNER JOIN Mov M on M.cOpeCod=O.cOpeCod"
SqlPFMov = SqlPFMov & "  INNER JOIN MovCap T on T.nMovNro=M.nMovNro "
SqlPFMov = SqlPFMov & "  INNER JOIN " & lsServerCons & "CTSConsol CTS  on CTS.cCtaCod=T.cCtaCod"
SqlPFMov = SqlPFMov & "  Where nEstCtaCTS NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ") "
SqlPFMov = SqlPFMov & "  ) TA ORDER BY TA.cCtaCod,TA.dFecha "
Set Conecta = New COMConecta.DCOMConecta
Set Rs = New ADODB.Recordset
Conecta.AbreConexion
    Set Rs = Conecta.CargaRecordSet(SqlPFMov)
Conecta.CierraConexion
SqlTabla = "CREATE TABLE " & vNom_cadena & " (" & _
               "Cuenta C(12) NOT NULL , " & _
               "Fecha D(8)  NOT NULL , " & _
               "Operacion C(50) NOT NULL , " & _
               "Monto Y(8) NOT NULL )"

Set conexiondbf = New ADODB.Connection
conexiondbf.Open "DSN=dsnAuditoriadbf;uid=;pwd=;sourceDB=" & lsArchivo & ";SOURCETYPE=DBF"
conexiondbf.Execute SqlTabla
Do While Not Rs.EOF
   Sql_Dbf = "INSERT INTO " & vNom_cadena & " (Cuenta,Fecha,Operacion,Monto) VALUES('" & _
               Rs!cCtaCod & "',ctod('" & Format(Rs!dFecha, "MM/dd/yyyy") & "'),'" & Rs!Operacion & "'," & Rs!nMonto & ")"
    conexiondbf.Execute Sql_Dbf
    Rs.MoveNext
Loop
conexiondbf.Close
Rs.Close
Set Rs = Nothing
Set conexiondbf = Nothing
Set Conecta = Nothing
Set RCD = Nothing
End Sub

Public Sub Generar_Inactivas(ByVal dFecSis As Date)
Dim conexiondbf As ADODB.Connection
Dim Rs As ADODB.Recordset
'Dim Conecta As DConecta

Dim sqlPF As String
Dim SqlTabla As String
Dim vNomser As String
Dim vNomBase As String
Dim vRFileAuditor As String
Dim vFechaSist As String
Dim vMes As String
Dim vYear As String
Dim vNom_cadena As String
Dim lsArchivo As String
Dim Sql_Dbf As String
'Screen.MousePointer = 11
lsArchivo = App.Path & "\SPOOLER\"


ChDir lsArchivo
vMes = Mid(dFecSis, 4, 2) 'vFechaSist
vYear = Mid(dFecSis, 7, 4) 'vFechaSist
vNom_cadena = "Inac_" & vYear & vMes & ".Dbf"

'sqlPF = "  Select CA.cCtaCod, convert(varchar(10),Cap.dApertura,101) as Apertura,"
'sqlPF = sqlPF & " Pro.nTasaInteres as nTasa, Pro.nSaldo, CA.nSaldoAnterior, "
'sqlPF = sqlPF & " convert(datetime,convert(varchar(8),cUltimaActualizacion,102)) as dUltMov,"
'sqlPF = sqlPF & " ca.bOrdPag From"
'sqlPF = sqlPF & " CaptacAhorros CA Inner Join Captaciones CAp on CA.cCtaCod=Cap.cCtaCod"
'sqlPF = sqlPF & " Inner Join Producto Pro on Pro.cCtaCod=CA.CCtaCod"
'sqlPF = sqlPF & " where Pro.nPrdEstado Not in (1300,1400) and CA.bInactiva=1"

sqlPF = " Select cCtaCod,"
sqlPF = sqlPF & " Convert(varchar(10),dAperAC,101) as dAperAc,"
sqlPF = sqlPF & " nTasaIntAC,nSaldDispAC,nSaldCntAC,"
sqlPF = sqlPF & " Convert(varchar(10),dUltCntAC,101) as dUltCntAc,nOrdPag"
sqlPF = sqlPF & " From dbCmactConsolidada..AhorrocConsol Where nEstCtaAC NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ") And bInactiva = 1 Order by cCtaCod"

Set Rs = New ADODB.Recordset

Conecta.AbreConexion
    Set Rs = Conecta.CargaRecordSet(sqlPF)
Conecta.CierraConexion


Set conexiondbf = New ADODB.Connection
SqlTabla = "CREATE TABLE " & vNom_cadena & " (" & _
               "Cuenta C(12) NOT NULL , " & _
               "Apertura D(8)  NOT NULL , " & _
               "Tasa N(8,2) NOT NULL , " & _
               "Salddisp Y(8) NOT NULL, " & _
               "SaldCnt Y(8) NOT NULL, " & _
               "UltCnt D(8)  NOT NULL , " & _
               "OrdPag C(1) NOT NULL)"
conexiondbf.Open "DSN=dsnAuditoriadbf;uid=;pwd=;sourceDB=" & lsArchivo & ";SOURCETYPE=DBF"
conexiondbf.Execute SqlTabla

Do While Not Rs.EOF
   Sql_Dbf = "INSERT INTO " & lsArchivo & vNom_cadena & " (Cuenta,Apertura,Tasa,Salddisp,SaldCnt,UltCnt,OrdPag) VALUES('" & _
               Rs!cCtaCod & "',ctod('" & Format(Rs!dAperAC, "MM/dd/yyyy") & "')," & Rs!nTasaIntAC & "," & Rs!nSaldDispAC & "," & Rs!nSaldcntAc & ",ctod('" & Format(Rs!dUltcntAC, "MM/dd/yyyy") & "'),'" & CStr(Rs!nOrdPag) & "')"
    conexiondbf.Execute Sql_Dbf

    Rs.MoveNext
Loop

End Sub


Public Sub Generar_Activas(ByVal dFecSis As Date)
Dim conexiondbf As ADODB.Connection
Dim Rs As ADODB.Recordset
Dim sqlPF As String
Dim SqlTabla As String
Dim vNomser As String
Dim vNomBase As String
Dim vRFileAuditor As String
Dim vFechaSist As String
Dim vMes As String
Dim vYear As String
Dim vNom_cadena As String
Dim lsArchivo As String
Dim Sql_Dbf As String

lsArchivo = App.Path & "\SPOOLER\"
ChDir lsArchivo

vMes = Mid(dFecSis, 4, 2) 'vFechaSist
vYear = Mid(dFecSis, 7, 4) 'vFechaSist
vNom_cadena = "AHAct_" + vYear + vMes + ".Dbf"

sqlPF = "  Select CA.cCtaCod, convert(varchar(10),Cap.dApertura,101) as Apertura,"
sqlPF = sqlPF & " Pro.nTasaInteres as nTasa, Pro.nSaldo, CA.nSaldoAnterior, "
sqlPF = sqlPF & " convert(datetime,convert(varchar(8),cUltimaActualizacion,102)) as dUltMov,"
sqlPF = sqlPF & " ca.bOrdPag From"
sqlPF = sqlPF & " CaptacAhorros CA Inner Join Captaciones CAp on CA.cCtaCod=Cap.cCtaCod"
sqlPF = sqlPF & " Inner Join Producto Pro on Pro.cCtaCod=CA.CCtaCod"
sqlPF = sqlPF & " where Pro.nPrdEstado Not in (" & gCapEstAnulada & "," & gCapEstCancelada & ") and CA.bInactiva=0"

'Set Conecta = New DConecta
Set Rs = New ADODB.Recordset
Conecta.AbreConexion
    Set Rs = Conecta.CargaRecordSet(sqlPF)
Conecta.CierraConexion
Set conexiondbf = New ADODB.Connection
SqlTabla = "CREATE TABLE " & vNom_cadena & " (" & _
               "Cuenta C(12) NOT NULL , " & _
               "Apertura D(8)  NOT NULL , " & _
               "Tasa N(8,2) NOT NULL , " & _
               "Salddisp Y(8) NOT NULL, " & _
               "SaldCnt Y(8) NOT NULL, " & _
               "UltCnt D(8)  NOT NULL , " & _
               "OrdPag C(1) NOT NULL)"
conexiondbf.Open "DSN=dsnAuditoriadbf;uid=;pwd=;sourceDB=" & lsArchivo & ";SOURCETYPE=DBF"
conexiondbf.Execute SqlTabla
Do While Not Rs.EOF
   Sql_Dbf = "INSERT INTO " & lsArchivo & vNom_cadena & " (Cuenta,Apertura,Tasa,Salddisp,SaldCnt,UltCnt,OrdPag) VALUES('" & _
               Rs!cCtaCod & "',ctod('" & Format(Rs!Apertura, "MM/dd/yyyy") & "')," & Rs!nTasa & "," & Rs!nSaldo & "," & Rs!nSaldoAnterior & ",ctod('" & Format(Rs!dUltMov, "MM/dd/yyyy") & "'),'" & CStr(Rs!bOrdPag) & "')"
    conexiondbf.Execute Sql_Dbf
    Rs.MoveNext
Loop
End Sub
Public Sub GeneraPF(ByVal FecSis As Date)
Dim conexiondbf As ADODB.Connection
Dim Rs As ADODB.Recordset
Dim Conecta As COMConecta.DCOMConecta
Dim Sql_Dbf As String
Dim sqlPF As String
Dim SqlTabla As String
Dim vNomser As String
Dim vNomBase As String
Dim vRFileAuditor As String
Dim vFechaSist As String
Dim vMes As String
Dim vYear As String
Dim vNom_cadena As String
Dim Rser As ADODB.Recordset
Dim lsArchivo As String
Dim RCD As NCOMRCDReportes
Dim lsServerCons As String
Set RCD = New NCOMRCDReportes
'lsServerCons = Rcd.GetServerConsol
lsArchivo = App.Path & "\SPOOLER\"

ChDir lsArchivo
vMes = Mid(FecSis, 4, 2) 'vFechaSist
vYear = Mid(FecSis, 7, 4) 'vFechaSist
vNom_cadena = "Plazo_" & vYear & vMes & ".Dbf"

'sqlPF = " Select T.cCtaCod, T.Apertura, T.nPlazo, T.nSaldo, T.nSaldoAnt, T.nTasa, T.dUltMov"
'sqlPF = sqlPF & " from ("
'sqlPF = sqlPF & " Select CTS.cCtaCod, convert(varchar(10),Cap.dApertura,101) as Apertura, nPlazo=720,"
'sqlPF = sqlPF & " Pro.nSaldo, Cap.nIntAcum as nSaldoAnt,"
'sqlPF = sqlPF & " Pro.nTasaInteres as nTasa , convert(datetime,convert(varchar(8),cUltimaActualizacion,102)) as dUltMov  from CaptacCTS CTS"
'sqlPF = sqlPF & " Inner Join Captaciones Cap on CTS.cCtaCod=Cap.cCtaCod"
'sqlPF = sqlPF & " Inner Join Producto Pro on Pro.cCtaCod=CTS.CCtaCod"
'sqlPF = sqlPF & " where Pro.nPrdEstado Not in (1300,1400)"
'sqlPF = sqlPF & " Union"
'sqlPF = sqlPF & " Select CPF.cCtaCod, convert(varchar(10),Cap.dApertura,101) as Apertura, CPF.nPlazo,"
'sqlPF = sqlPF & " Pro.nSaldo,CPF.nApertura as nSaldoAnt,"
'sqlPF = sqlPF & " Pro.nTasaInteres as nTasa , convert(datetime,convert(varchar(8),cUltimaActualizacion,102)) as dUltMov"
'sqlPF = sqlPF & " from CaptacPlazoFijo CPF"
'sqlPF = sqlPF & " Inner Join Producto Pro on Pro.cCtaCod=CPF.CCtaCod"
'sqlPF = sqlPF & " Inner Join Captaciones Cap on CPF.cCtaCod=Cap.cCtaCod"
'sqlPF = sqlPF & " where Pro.nPrdEstado Not in (1300,1400)) T Order by T.cCtaCod"

sqlPF = " SELECT T.cCtaCod, T.dApertura, T.nPlazo, T.nCapital, T.nSaldCnt, T.nTasa,T.dUltMov" _
        & " FROM (Select cCtaCod, Convert(varchar(10),dAperPF,101) dApertura ,nPlazo,nCapital," _
        & " nSaldCntPF nSaldCnt,nTasaIntPF nTasa,Convert(varchar(10), dUltMovPF,101) dUltMov" _
        & " from " & lsServerCons & "PlazoFijoConsol Where nEstCtaPF NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ")" _
        & " Union" _
        & " Select cCtaCod,Convert(varchar(10),dAperCTS,101) dApertura, nPlazo = 720," _
        & " nCapitalCTS nCapital, nSaldCntCTS nSaldCnt, nTasaIntCTS nTasa," _
        & " Convert(varchar(10),dUltMovCTS,101) dUltMov" _
        & " from " & lsServerCons & "CTSConsol Where nEstCtaCTS NOT IN (" & gCapEstAnulada & "," & gCapEstCancelada & ")) T order by T.cCtaCod "
Set Rs = New ADODB.Recordset
Set Rs = GetPlazoFijo
Set conexiondbf = New ADODB.Connection
SqlTabla = "CREATE TABLE " & vNom_cadena & " (" & _
               "Cuenta C(18) NOT NULL , " & _
               "Apertura D(10)  NOT NULL , " & _
               "Plazo I(4,0) NOT NULL , " & _
               "Capital Y(8) NOT NULL , " & _
               "SaldCnt Y(8) NOT NULL , " & _
               "Tasa N(3,2) NOT NULL , " & _
               "UltMov D(8) NOT NULL )"

conexiondbf.Open "DSN=dsnAuditoriadbf;uid=;pwd=;sourceDB=" & lsArchivo & ";SOURCETYPE=DBF"
conexiondbf.CursorLocation = adUseClient
conexiondbf.Execute SqlTabla

Do While Not Rs.EOF
Sql_Dbf = ""
Sql_Dbf = "INSERT INTO " & vNom_cadena & " (Cuenta,Apertura,Plazo,Capital,SaldCnt,Tasa,UltMov) VALUES('" & _
               Rs!cCtaCod & "',ctod('" & Format(Rs!dApertura, "MM/dd/yyyy") & "')," & Rs!nPlazo & "," & Rs!nCapital & "," & Rs!nSaldCnt & ", " & Rs!nTasa & ",ctod('" & Format(Rs!dApertura, "MM/dd/yyyy") & "'))"
    conexiondbf.Execute Sql_Dbf
    Rs.MoveNext
Loop
Rs.Close
conexiondbf.Close

Set Rser = Nothing
Set RCD = Nothing
End Sub

Private Sub Class_Initialize()
    Set Conecta = New COMConecta.DCOMConecta
End Sub

Private Sub Class_Terminate()
Set Conecta = Nothing
End Sub

'-----------------------------------------------------------------------------------------------------------

