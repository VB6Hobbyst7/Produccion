VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCajaChica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim lsBaseCom As String
Dim lsBasePers As String
Enum DatosCajaChica
    SaldoActual = 0
    SaldoAnterior = 1
    MontoTope = 2
    MontoDesemb = 3
    MontoAsig = 4
    NroCajaChica = 5
End Enum
Enum RendicionTipo
    Exacta = 0
    ConIngreso = 1
    ConEgreso = 2
    '***Agregado por ELRO el 20120705, según OYP-RFC47-2012
    Todo = 3
    '***Fin Agregado por ELRO******************************
End Enum
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni
lsBaseCom = oIni.BaseComunes
lsBasePers = oIni.BasePersonas
Set oIni = Nothing
End Sub
Public Function GetDatosCajaChica(ByVal psAreaCh As String, ByVal psAgeCh As String, Optional ByVal DatosCajaChica As DatosCajaChica = SaldoActual) As Currency
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function
'***Modificado por ELRO el 20120605, según OYP-RFC047-2012
'sql = "SELECT  cAreaCod, cAgeCod, nProcNro, nMontoAsig, nTopeEgresos,nMontoDesem,nSaldoAnt,nSaldo " _
'    & " FROM   CAJACHICA  C " _
'    & " WHERE  C.nProcNro = (SELECT MAX(nProcNro) " _
'    & "                      FROM CAJACHICA C1 " _
'    & "                      WHERE C1.cAreaCod = C.cAreaCod and C1.cAgeCod = C.cAgeCod ) " _
'    & "        And cAreaCod='" & psAreaCh & "' and cAgeCod= '" & psAgeCh & "'"
sql = "exec stp_sel_DevolverDatosCajaChica '" & psAreaCh & "', '" & psAgeCh & "'"
'***Fin Modificado por ELRO*******************************
GetDatosCajaChica = 0
Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    Select Case DatosCajaChica
        Case MontoAsig
            GetDatosCajaChica = rs!nMontoAsig
        Case MontoDesemb
            GetDatosCajaChica = rs!nMontoDesem
        Case MontoTope
            GetDatosCajaChica = rs!nTopeEgresos
        Case SaldoActual
            GetDatosCajaChica = rs!nSaldo
        Case SaldoAnterior
            GetDatosCajaChica = rs!nSaldoAnt
        Case NroCajaChica
            GetDatosCajaChica = rs!nProcNro
    End Select
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing

End Function
'Public Function VerificaCHRendida(ByVal psAreaCH As String, ByVal psAgeCH As String, ByVal pnNroProc As Integer) As Boolean
'Dim Sql As String
'Dim rs As ADODB.Recordset
'Dim oCon As DConecta
'Set oCon = New DConecta
'Set rs = New ADODB.Recordset
'
'If oCon.AbreConexion = False Then Exit Function
'
'Sql = "SELECT  MCH.CAREACOD,MCH.CAGECOD, nProcNro " _
'    & " FROM   MOV M JOIN MOVCAJACHICA MCH ON M.CMOVNRO=MCH.CMOVNRO  " _
'    & " WHERE  MCH.cProcTpo IN('" & gCHTipoProcRendicion & "') " _
'    & "        And MCH.cAreaCod='" & psAreaCH & "' and MCH.cAgeCod= '" & psAgeCH & "' " _
'    & "        And MCH.nProcNro=" & pnNroProc & " " _
'    & "        AND M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "') "
'
'VerificaCHRendida = False
'Set rs = oCon.CargaRecordSet(Sql)
'If Not rs.EOF And Not rs.BOF Then
'    VerificaCHRendida = True
'End If
'rs.Close
'Set rs = Nothing
'oCon.CierraConexion
'Set oCon = Nothing
'End Function
'Public Function VerificaCHAutorizada(ByVal psAreaCH As String, ByVal psAgeCH As String, ByVal pnNroProc As Integer) As Boolean
'Dim Sql As String
'Dim rs As ADODB.Recordset
'Dim oCon As DConecta
'Set oCon = New DConecta
'Set rs = New ADODB.Recordset
'
'If oCon.AbreConexion = False Then Exit Function
'
'Sql = "SELECT  MCH.CMOVNRO ,  MCH.CAREACOD,MCH.CAGECOD, nProcNro  " _
'    & " FROM   MOV M JOIN MOVCAJACHICA MCH ON M.CMOVNRO=MCH.CMOVNRO  " _
'    & " WHERE  MCH.cProcTpo IN('" & gCHTipoProcHabilitacion & "') " _
'    & "        And MCH.cAreaCod='" & psAreaCH & "' and MCH.cAgeCod= '" & psAgeCH & "' " _
'    & "        And MCH.nProcNro=" & pnNroProc & " " _
'    & "        AND M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "') "
'
'VerificaCHAutorizada = False
'Set rs = oCon.CargaRecordSet(Sql)
'If Not rs.EOF And Not rs.BOF Then
'    VerificaCHAutorizada = True
'End If
'rs.Close
'Set rs = Nothing
'oCon.CierraConexion
'Set oCon = Nothing
'End Function


Public Function VerificaTopeCajaChica(ByVal psAreaCod As String, ByVal psAgeCod As String) As Boolean
Dim lnMontoAsig As Currency
Dim oGen As DGeneral
Dim lnTasaCajaCh As Currency
Dim lnSaldo As String
Set oGen = New DGeneral

VerificaTopeCajaChica = False
lnSaldo = GetDatosCajaChica(psAreaCod, psAgeCod, SaldoActual)
lnMontoAsig = GetDatosCajaChica(psAreaCod, psAgeCod, MontoAsig)
lnTasaCajaCh = oGen.GetParametro(4000, 1000)
If lnSaldo <= lnMontoAsig * (lnTasaCajaCh / 100) Then
    VerificaTopeCajaChica = True
End If
Set oGen = Nothing
End Function
Public Function GetSolicitudesArendirCH(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnProcCH As Integer) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New Recordset

If oCon.AbreConexion = False Then Exit Function

sql = " SELECT  MD.CDOCNRO, Convert(varchar(12), MD.dDocFecha,103) as dDocFecha, P.cPersNombre,MC.nMovMonto, P.cPersCod, M.cMovDesc, " _
    & "         ISNULL(A.CAREADESCRIPCION,'') AS AREADESC , MAR.nMOVNRO, A.CAREACOD, AG.CAGECOD , D.CDOCABREV   " _
    & " FROM    MOV M " _
    & "         JOIN MOVCONT MC ON MC.nMOVNRO = M.nMOVNRO " _
    & "         JOIN MOVARENDIR  MAR ON MAR.nMOVNRO =M.nMOVNRO " _
    & "         JOIN MOVCAJACHICA MCH ON MCH.nMOVNRO = MAR.nMOVNRO " _
    & "         JOIN PERSONA P ON P.CPERSCOD = MAR.CPERSCOD " _
    & "         LEFT JOIN AREAS A ON A.CAREACOD = MAR.CAREACOD " _
    & "         LEFT JOIN AGENCIAS AG ON AG.CAGECOD = MAR.CAGECOD " _
    & "         JOIN MOVDOC MD ON MD.nMOVNRO = M.nMOVNRO " _
    & "         JOIN DOCUMENTO D ON D.nDocTpo = MD.nDocTpo " _
    & "  WHERE  MAR.cTpoArendir ='" & gArendirTipoCajaChica & "' " _
    & "         AND NOT EXISTS (SELECT  M1.CMOVNRO " _
    & "                         FROM    MOV M1 JOIN MOVREF MR ON MR.nMOVNRO = M1.nMOVNRO " _
    & "                         WHERE   MR.nMOVNROREF=M.nMOVNRO AND M1.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') ) " _
    & "         AND M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "') " _
    & "         AND MCH.cAreaCod='" & psAreaCh & "' AND MCH.cAgeCod='" & psAgeCh & "' "

Set rs = oCon.CargaRecordSet(sql)
Set GetSolicitudesArendirCH = rs

oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetAtencionesArendir(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnProcCH As Integer, ByVal psCtaArendir As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New Recordset

If oCon.AbreConexion = False Then Exit Function

'****Modificado por ELRO el 20120903, según OYP-RFC104-2012
'sql = " SELECT  MD.CDOCNRO, CONVERT(CHAR(12), MD.dDocFecha ,103  ) AS Fecha ,P.CPERSNOMBRE, ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS IMPORTE,  "
'    & "         nMovSaldo, M1.CMOVDESC AS DESCRIPCION , ISNULL(A.CAREADESCRIPCION,'')  AS CAREADESC,  "
'    & "         M.nMOVNRO AS CMOVNROSOL, M1.nMOVNRO AS CMOVNROATENC , MAR.CAREACOD, "
'    & "         ISNULL(AG.CAGEDESCRIPCION,'') AS CAGEDESC, MAR.CAGECOD , MAR.CPERSCOD "
'    & "  FROM   MOV  M " _
'    & "         JOIN MOVCONT MCT ON MCT.nMOVNRO = M.nMOVNRO "
'    & "         JOIN MOVARENDIR MAR ON MAR.nMOVNRO = M.nMOVNRO "
'    & "         JOIN PERSONA P     ON P.CPERSCOD = MAR.CPERSCOD "
'    & "         LEFT JOIN AREAS A ON A.CAREACOD = MAR.CAREACOD "
'    & "         LEFT JOIN AGENCIAS AG ON AG.CAGECOD=MAR.CAGECOD "
'    & "         JOIN MOVCAJACHICA MCH ON MCH.nMOVNRO = MAR.nMOVNRO "
'    & "         JOIN MOVDOC MD ON MD.nMOVNRO = M.nMOVNRO "
'    & "         JOIN DOCUMENTO  D ON D.nDocTpo = MD.nDocTpo "
'    & "         JOIN MOVREF MR ON MR.nMOVNROREF= M.nMOVNRO "
'    & "         JOIN MOV M1 ON M1.nMOVNRO = MR.nMOVNRO "
'    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M1.nMOVNRO "
'    & "         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM "
'    & "  WHERE  MAR.cTpoArendir='" & gArendirTipoCajaChica & "' AND cCtaContCod in ('" & psCtaArendir & "') "
'    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") "
'    & "         AND M1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") "
'    & "         AND MCH.cAreaCod='" & psAreaCh & "' AND MCH.cAgeCod='" & psAgeCh & "' AND nProcNro=" & pnProcCH & " "
'    & "         AND NOT EXISTS ( SELECT M2.CMOVNRO FROM MOV M2 JOIN MOVREF MR2 ON M2.nMOVNRO = MR2.nMOVNRO "
'    & "                          WHERE  MR2.nMOVNROREF =M1.nMOVNRO AND SUBSTRING(M2.COPECOD,1,5) IN ('" & Mid(gCHArendirCtaSustMN, 1, 5) & "','" & Mid(gCHArendirCtaSustME, 1, 5) & "')  "
'    & "                                 AND M2.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ")) "
sql = " exec stp_sel_DevolverAtencionEntregaEfectivo  '" & psAreaCh & "', '" & psAgeCh & "', " & pnProcCH & ", '" & psCtaArendir & "'"
'****Fin Modificado por ELRO el 20120903*******************
Set rs = oCon.CargaRecordSet(sql)
Set GetAtencionesArendir = rs

oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetCHSustSinRendicion(ByVal psAreaCh As String, _
                                      ByVal psAgeCh As String, _
                                      ByVal pnProcCH As Integer, _
                                      ByVal psCtaArendir As String, _
                                      Optional ByVal pbTodo As Boolean = False, _
                                      Optional ByVal pcPersCod As String = "") As ADODB.Recordset
                                      '***Agregado por ELRO los parametros pbTodo y pcPersCod  el 20120616, según OYP-RFC047-2012
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New Recordset

If oCon.AbreConexion = False Then Exit Function

'***Modificado por ELRO el 20120616, según OYP-RFC047-2012
'sql = " SELECT  MD.CDOCNRO, CONVERT(CHAR(12),MD.DDOCFECHA,103 ) AS Fecha ,P.CPERSNOMBRE, ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS IMPORTE,  " _
'    & "         nMovSaldo, M1.CMOVDESC AS DESCRIPCION , ISNULL(A.CAREADESCRIPCION,'')  AS CAREADESC,  " _
'    & "         M.nMOVNRO AS CMOVNROSOL, M1.nMOVNRO AS CMOVNROATENC , RH.CAREACODActual CAREACOD, " _
'    & "         ISNULL(AG.CAGEDESCRIPCION,'') AS CAGEDESC, ISNULL((SELECT ISNULL(cAgeCod,'') FROM AreaAgencia WHERE cAreaCod = aa.cAreaCod and cAgeCod = ag.cAgeCod),'') CAGECOD, MAR.CPERSCOD, MCH.nProcNro " _
'    & "  FROM   MOV  M " _
'    & "         JOIN MOVCONT MCT ON MCT.nMOVNRO = M.nMOVNRO " _
'    & "         JOIN MOVARENDIR MAR ON MAR.nMOVNRO = M.nMOVNRO " _
'    & "         JOIN PERSONA P     ON P.CPERSCOD = MAR.CPERSCOD JOIN RRHH RH ON RH.cPersCod = Mar.cPersCod " _
'    & "         LEFT JOIN AREAS A ON A.CAREACOD = RH.CAREACODActual " _
'    & "         LEFT JOIN AGENCIAS AG ON AG.CAGECOD= RH.CAGENCIAActual  " _
'    & "         LEFT JOIN (SELECT Distinct cAreaCod FROM AreaAgencia ) AA ON AA.cAreaCod = A.cAreaCod " _
'    & "         JOIN MOVCAJACHICA MCH ON MCH.nMOVNRO = MAR.nMOVNRO " _
'    & "         JOIN MOVDOC MD ON MD.nMOVNRO = M.nMOVNRO " _
'    & "         JOIN DOCUMENTO  D ON D.nDocTpo = MD.nDocTpo " _
'    & "         JOIN MOVREF MR ON MR.nMOVNROREF= M.nMOVNRO " _
'    & "         JOIN MOV M1 ON M1.nMOVNRO = MR.nMOVNRO " _
'    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M1.nMOVNRO " _
'    & "         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
'    & "  WHERE  MAR.cTpoArendir='" & gArendirTipoCajaChica & "' AND cCtaContCod in ('" & psCtaArendir & "')  AND nProcNro <=" & pnProcCH & " " _
'    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
'    & "         AND M1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
'    & "         AND MCH.cAreaCod='" & psAreaCh & "' AND MCH.cAgeCod='" & psAgeCh & "'  " _
'    & "         AND NOT EXISTS ( SELECT M2.CMOVNRO FROM MOV M2 JOIN MOVREF MR2 ON M2.nMOVNRO = MR2.nMOVNRO " _
'    & "                          WHERE  MR2.nMOVNROREF =M1.nMOVNRO AND SUBSTRING(M2.COPECOD,1,5) IN ('" & Mid(gCHArendirCtaRendMN, 1, 5) & "','" & Mid(gCHArendirCtaRendME, 1, 5) & "')  " _
'    & "                                 AND M2.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ")) "
sql = "exec stp_sel_DevolverCajaChicaSustSinRendicion '" & psAreaCh & "','" & psAgeCh & "'," & pnProcCH & ", '" & psCtaArendir & "', " & IIf(pbTodo, 1, 0) & ", '" & pcPersCod & "'"
'***Fin Modificado por ELRO*******************************
Set rs = oCon.CargaRecordSet(sql)
Set GetCHSustSinRendicion = rs

oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetCHRendiciones(ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                 ByVal pnProcCH As Integer, ByVal psCtaArendir As String, _
                                ByVal psCtaPendiente As String, ByVal pnRend As RendicionTipo) As ADODB.Recordset

Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lsFiltro As String
Set oCon = New DConecta
Set rs = New Recordset

If oCon.AbreConexion = False Then Exit Function
Select Case pnRend
       
    Case Exacta
        lsFiltro = " and Rend.MontoRendido=0"
    Case ConIngreso
        lsFiltro = " and Rend.MontoRendido<0"
    Case ConEgreso
        lsFiltro = " and Rend.MontoRendido>0"
    '***Agregado por ELRO el 20120705, según OYP-RFC047-2012
    Case Todo 'Exacta y ConIngreso
        lsFiltro = "and (Rend.MontoRendido=0 or Rend.MontoRendido<0)"
    '*** Fin Agregado por ELRO******************************
    
End Select

'***Modificado por ELRO el 20130221, según SATI INC1301300007
'sql = " SELECT  MD.CDOCNRO, CONVERT(CHAR(12),MD.DDOCFECHA,103 ) AS Fecha, P.CPERSNOMBRE, " _
'    & "         Rend.MontoRendido  as MontoRealRend, abs(Rend.MontoRendido)  as MontoRend, " _
'    & "         Rend.CMOVDESC AS DESCRIPCION , ISNULL(A.CAREADESCRIPCION,'')  AS CAREADESC, " _
'    & "         mar.nMovNro nMovNroSol, M.nMOVNRO AS nMOVNROAtenc, MAR.CAREACOD, " _
'    & "         ISNULL(AG.CAGEDESCRIPCION,'') AS CAGEDESC, MAR.CAGECOD , MAR.CPERSCOD, Rend.nMOVNROREND " _
'    & " FROM    MOV  M " _
'    & "         JOIN MOVCONT MCT    ON MCT.nMOVNRO = M.nMOVNRO " _
'    & "         JOIN MovRef mrS     ON mrs.nMovNro = m.nMovNro JOIN MOVARENDIR MAR     ON MAR.nMOVNRO = MrS.nMOVNRORef " _
'    & "         JOIN PERSONA P      ON P.CPERSCOD = MAR.CPERSCOD " _
'    & "         LEFT JOIN AREAS A   ON A.CAREACOD = MAR.CAREACOD " _
'    & "         LEFT JOIN AGENCIAS AG   ON AG.CAGECOD=MAR.CAGECOD " _
'    & "         JOIN MOVCAJACHICA MCH   ON MCH.nMOVNRO = M.nMOVNRO " _
'    & "         JOIN MOVDOC MD      ON MD.nMOVNRO = M.nMOVNRO " _
'    & "         " _
'    & "         JOIN MOVREF MR      ON MR.nMOVNROREF= M.nMOVNRO " _
'    & "         JOIN MOV M1         ON M1.nMOVNRO = MR.nMOVNRO " _
'    & "         " _
'    & "         " _
'    & "         JOIN (  SELECT  M2.nMOVNRO  AS nMOVNROREND, MR2.nMOVNROREF as nMovNroAtenc, " _
'    & "                         IsNull(SUM(ISNULL(ME1.NMOVMEIMPORTE, MC2.NMOVIMPORTE)),0) AS MontoRendido, M2.cMovDesc " _
'    & "                 FROM    MOV M2  JOIN MOVCAJACHICA MCH2   ON MCH2.nMOVNRO = M2.nMOVNRO" _
'    & "                         JOIN MOVREF MR2 ON M2.nMOVNRO = MR2.nMOVNRO " _
'    & "                         LEFT JOIN MOVCTA MC2 ON MC2.nMOVNRO = M2.nMOVNRO  AND MC2.cCtaContCod IN ('" & psCtaArendir & "','" & psCtaPendiente & "') " _
'    & "                         LEFT JOIN MOVME ME1 ON ME1.nMOVNRO = MC2.nMOVNRO AND ME1.nMOVITEM = MC2.nMOVITEM " _
'    & "                 WHERE   SUBSTRING(M2.COPECOD,1,5) IN ('" & Mid(gCHArendirCtaRendMN, 1, 5) & "','" & Mid(gCHArendirCtaRendME, 1, 5) & "') "
'sql = sql + "                    AND M2.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") AND nProcNro=" & pnProcCH & " " _
'    & "                 GROUP BY M2.nMOVNRO , MR2.nMOVNROREF, M2.cMovDesc ) as REND " _
'    & "         ON REND.nMovNroRend = M1.nMOVNRO " _
'    & " WHERE   MAR.cTpoArendir='" & gArendirTipoCajaChica & "' " _
'    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
'    & "         AND M1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
'    & "         AND MCH.cAreaCod='" & psAreaCh & "' AND MCH.cAgeCod='" & psAgeCh & "'  " & lsFiltro
sql = "exec stp_sel_DevolverRendiciones '" & psAreaCh & "', '" & psAgeCh & "', '" & pnProcCH & "', '" & psCtaArendir & "', '" & psCtaPendiente & "', " & pnRend & ""
'***Fin Modificado por ELRO el 20130221**********************
Set rs = oCon.CargaRecordSet(sql)
Set GetCHRendiciones = rs

oCon.CierraConexion
Set oCon = Nothing
End Function

Public Sub GrabaSolEgresoDirecto(ByVal pbNewProv As Boolean, _
                                 ByVal psPersCod As String, ByVal psProvEstado As String, ByVal psMovUltActProv As String, _
                                 ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, ByVal pnDestino As Integer, _
                                 ByVal psCtaProvision As String, _
                                 ByVal psDocNro As String, ByVal pnDocTpo As TpoDoc, ByVal pdDocFecha As Date, _
                                 ByVal pnTotal As Currency, _
                                 ByVal psAgeCh As String, ByVal psAreaCh As String, ByVal pnNroProc As Integer, _
                                 ByVal rsCtasCont As ADODB.Recordset, ByVal rsObjetos As ADODB.Recordset, _
                                 ByVal rsImp As ADODB.Recordset, Optional pnTCVariable As Currency, _
                                 Optional ByVal lsCtaITFD As String, Optional ByVal lsCtaITFH As String, Optional ByVal nTasaITF As Double = 0, Optional ByRef psCtasInexis As String = "", _
                                 Optional ByVal psMatrizViaticos As Variant = "", Optional ByVal pnNumAgesViaticos As Integer = 0, Optional ByRef pnMovNro As Long = 0, _
                                 Optional ByVal pnRetencSistPensProvMontoBase As Currency = 0#, Optional ByVal pnRetencSistPensProvTotal As Currency = 0#, _
                                Optional ByVal pnTpoProvSistPens As Integer = 0, Optional ByVal poMatDatosRetencSistPensProv As Variant = "") ' madm 20110507 - pnMovNro
Dim sql As String
Dim oMov As DMov
Dim lsSubCta As String
Dim lnItem As Integer
Dim lnOrdenObj As Integer
Dim lsAgeCod As String
Dim lsAreaCod As String
Dim lnImporte As Currency
Dim lnImportePend As Currency
Dim lnSaldoArendir As Currency
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim I As Integer
Dim lnMontoNoGrabado As Currency
Dim lcCodObj As String
Dim lnItemDistri As Integer
Dim rsAgesDistrib As ADODB.Recordset
Dim lssql As String
Dim lnMonto As Currency
Dim lnRegImporte As Currency
Dim lsCtasInexis As String
Dim lcCtaCont As String
Dim rsBuscaCuenta As ADODB.Recordset
Dim lnItemPrin As Integer
Dim lnImpoPrin As Currency
Dim lsSubCtaCorto As String
Dim lnAbortaDistri As Integer

'*** PEAC 20110315
Dim lnMontoCtaOrd As Double, sTexto As String, N As Integer, m As Integer
Dim lcAge As String
Dim rsCtasOrden As ADODB.Recordset

On Error GoTo GrabaSolEgresoDirectoErr
Set oMov = New DMov

oMov.BeginTrans

lbTrans = True

'If pbNewProv Then
'    oMov.InsertProv psPersCod, psProvEstado, psMovUltActProv
'End If
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabPendiente, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
pnMovNro = lnMovNro
oMov.InsertaMovCont lnMovNro, pnTotal, "0", "0"
oMov.InsertaMovGasto lnMovNro, psPersCod, IIf(pnDestino = -1, "", pnDestino)
'Grabamos en MovCta, MovObj y MovCant
lnItem = 0
lnItemDistri = 0
If Not rsCtasCont Is Nothing Then
    Do While Not rsCtasCont.EOF
       lsSubCta = ""
       If rsCtasCont!Código <> "" And rsCtasCont!Monto <> 0 Then
            lnItem = lnItem + 1
            'generamos la subcta recorriendo el recordset de los objetos
            If Not rsObjetos Is Nothing Then
                rsObjetos.MoveFirst
                Do While Not rsObjetos.EOF
                    If rsObjetos!CtaCont = rsCtasCont!Código And rsObjetos!ItemCtaCont = rsCtasCont!ItemCtaCont Then
                        lsSubCta = lsSubCta + rsObjetos![SubCta]
                    End If
                    rsObjetos.MoveNext
                Loop
            End If
            
            If pnDestino = 3 Or pnDestino = 2 Or (pnDestino = 1 And rsCtasCont.Fields(4) = 0) Then
                lnImporte = rsCtasCont!Monto
                For I = 6 To rsCtasCont.Fields.Count - 1
                    rsImp.Move I - 6, adBookmarkFirst
                    If rsImp.Fields(10) = 0 Or Trim(rsCtasCont.Fields(I).Name) = "AJUSTE" Then
                        lnImporte = lnImporte + rsCtasCont.Fields(I).value
                    End If
                    If rsCtasCont.Fields(I).value <> 0 Then
                        oMov.InsertaMovOtrosItem lnMovNro, lnItem, rsImp!CtaCont, rsCtasCont.Fields(I).value, 3
                    End If
                Next
                
                '*********************** PEAC 20100903 --- DISTRIBUCION
                'If psOpeCod = "401371" And lsSubCta = "01" And psAreaCh <> "026" Then
                '*** PEAC 20110219 - SE AGREGO AL IF (Left(rsCtasCont!Código, 2) = "45") SOLO DEBE DISTRIBUIR LA CTA 45 (GASTOS)
                '***Modificado por ELRO el 20120921, según OYP-RFC111-2012
                If (psOpeCod = "401371" Or psOpeCod = "401344") And Right(lsSubCta, 2) = "01" And Left(rsCtasCont!Código, 2) = "45" Then
                '***Fin Modificado por ELRO el 20120921*******************
                    If psAreaCh = "026" Or psAreaCh = "041" Or psAreaCh = "067" Then '*** PEAC 20110304
'                        lnItem = IIf(lnItemDistri > 0, lnItemDistri, lnItem) + 1
'                        lnItemDistri = lnItem
                        
                        lnItem = IIf(lnItemDistri > 0, lnItemDistri + 1, IIf(lnItem > 1, lnItem + 1, lnItem))
                        lnItemDistri = lnItem
                        
                        oMov.InsertaMovCta lnMovNro, lnItem, rsCtasCont!Código + lsSubCta, IIf(rsCtasCont!DH = "D", lnImporte, lnImporte * -1) * IIf(pnDocTpo = 7, -1, 1)
                    Else
                        lsSubCtaCorto = Mid(Trim(lsSubCta), 1, Len(Trim(lsSubCta)) - 2)
                        If Not rsObjetos Is Nothing Then
                            rsObjetos.MoveFirst
                            Do While Not rsObjetos.EOF
                                If rsObjetos!CtaCont = rsCtasCont!Código And rsObjetos!ItemCtaCont = rsCtasCont!ItemCtaCont Then
                                    lcCodObj = rsObjetos!Código
                                End If
                                rsObjetos.MoveNext
                            Loop
                        End If
    
                        'If Left(lcCodObj, 3) <> "02601" Then
                        If Left(lcCodObj, 3) = "026" Or Left(lcCodObj, 3) = "041" Or Left(lcCodObj, 3) = "067" Then
                            lnItemDistri = lnItemDistri + 1
                            oMov.InsertaMovCta lnMovNro, lnItemDistri, rsCtasCont!Código + lsSubCta, IIf(rsCtasCont!DH = "D", lnImporte, lnImporte * -1) * IIf(pnDocTpo = 7, -1, 1)
                        Else
                            If pnNumAgesViaticos > 0 Then
                                
'                                Set rsAgesDistrib = New ADODB.Recordset
'                                lsSql = " exec stp_sel_AgenciaPorcentajeGastos "
'                                Set rsAgesDistrib = oMov.CargaRecordSet(lsSql)
            
                                lnMonto = 0
                                lnRegImporte = 0
                                lsCtasInexis = ""
                                
                                For I = 1 To pnNumAgesViaticos
                                
                                'Do While Not rsAgesDistrib.EOF
                                    'lcCtaCont = rsCtasCont!Código + lsSubCtaCorto + rsAgesDistrib!cAgeCod
                                    lcCtaCont = rsCtasCont!Código + lsSubCtaCorto + psMatrizViaticos(1, I)
                                    
                                    Set rsBuscaCuenta = oMov.CargaRecordSet("select cCtaContCod From CtaCont where cCtaContCod like '" & lcCtaCont & "%' And nCtaEstado = 1 ")
                                                                                                    
                                    If Not (rsBuscaCuenta.EOF And rsBuscaCuenta.BOF) Then
                                        lnItemDistri = lnItemDistri + 1
                                        
                                        lnRegImporte = Round((IIf(rsCtasCont!DH = "D", psMatrizViaticos(3, I), psMatrizViaticos(3, I) * -1) * IIf(pnDocTpo = 7, -1, 1)), 2)
                                        'lnRegImporte = Round((IIf(prsCtasCont!DH = "D", psMatrizViaticos(3, I), psMatrizViaticos(3, I) * -1) * IIf(pnDocTpo = 7, -1, 1)), 2)
                                        
                                        If I = 1 Then
                                            lnItemPrin = lnItemDistri
                                            lnImpoPrin = lnRegImporte
                                        End If
                                        
'                                        If rsAgesDistrib!cAgeCod = "01" Then
'                                            lnItemPrin = lnItemDistri
'                                            lnImpoPrin = lnRegImporte
'                                        End If
                                        
                                        oMov.InsertaMovCta lnMovNro, lnItemDistri, lcCtaCont, lnRegImporte
                                        lnMonto = lnMonto + lnRegImporte
                                        
                                    Else
                                        lsCtasInexis = lsCtasInexis + lcCtaCont + "-"
                                    End If
                                '    rsAgesDistrib.MoveNext
                                'Loop
                                Next I
                            Else
                        
                                Set rsAgesDistrib = New ADODB.Recordset
                                lssql = " exec stp_sel_AgenciaPorcentajeGastos "
                                Set rsAgesDistrib = oMov.CargaRecordSet(lssql)
            
                                lnMonto = 0
                                lnRegImporte = 0
                                lsCtasInexis = ""
                                
                                Do While Not rsAgesDistrib.EOF
                                    lcCtaCont = rsCtasCont!Código + lsSubCtaCorto + rsAgesDistrib!cAgeCod
                                    'Set rsBuscaCuenta = oMov.CargaRecordSet("select cCtaContCod From CtaCont where cCtaContCod like '" & lcCtaCont & "%' and cCtaContCod <> '" & lcCtaCont & "' And nCtaEstado = 1 ")
                                    Set rsBuscaCuenta = oMov.CargaRecordSet("select cCtaContCod From CtaCont where cCtaContCod like '" & lcCtaCont & "%' And nCtaEstado = 1 ")
                                                                
        '                            If Not (rsBuscaCuenta.EOF And rsBuscaCuenta.BOF) Then
        '                                lnAbortaDistri = 1
        '                                Exit Do
        '                            End If
                                    
                                    If Not (rsBuscaCuenta.EOF And rsBuscaCuenta.BOF) Then
                                        lnItemDistri = lnItemDistri + 1
                                        lnRegImporte = Round((IIf(rsCtasCont!DH = "D", lnImporte, lnImporte * -1) * IIf(pnDocTpo = 7, -1, 1)) * rsAgesDistrib!nAgePorcentaje / 100, 2)
                                        
                                        If rsAgesDistrib!cAgeCod = "01" Then
                                            lnItemPrin = lnItemDistri
                                            lnImpoPrin = lnRegImporte
                                        End If
                                        
                                        oMov.InsertaMovCta lnMovNro, lnItemDistri, lcCtaCont, lnRegImporte
                                        lnMonto = lnMonto + lnRegImporte
                                        
                                    Else
                                        lsCtasInexis = lsCtasInexis + lcCtaCont + "-"
                                    End If
                                    rsAgesDistrib.MoveNext
                                Loop
                                
                             End If
                             
                                If lnMonto > lnImporte Then
                                    'lcCtaCont = rsCtasCont!Código + "01" '***PEAC 20110129
                                    lcCtaCont = rsCtasCont!Código + lsSubCtaCorto + "01"
                                    
                                    Set rsBuscaCuenta = oMov.CargaRecordSet("select * From CtaCont where cCtaContCod like '" & lcCtaCont & "' ")
                                    If Not (rsBuscaCuenta.EOF And rsBuscaCuenta.BOF) Then
                                        oMov.ActualizaMovCta lnMovNro, lnItemPrin, , lnImpoPrin - (lnMonto - lnImporte) '' la dif manda a la ag princ
                                    Else
                                        oMov.ActualizaMovCta lnMovNro, lnItemDistri, , lnRegImporte - (lnMonto - lnImporte) '' si noexiste la age prin manda al ultimo age registrado
                                    End If
                                
                                    'oMov.ActualizaMovCta lnMovNro, lnItemDistri, , lnRegImporte - (lnMonto - lnImporte)
                                    
                                ElseIf lnImporte > lnMonto Then
        
                                    'lcCtaCont = rsCtasCont!Código + "01" '*** PEAC 20110129
                                    lcCtaCont = rsCtasCont!Código + lsSubCtaCorto + "01"
                                                                
                                    Set rsBuscaCuenta = oMov.CargaRecordSet("select * From CtaCont where cCtaContCod like '" & lcCtaCont & "' ")
                                    If Not (rsBuscaCuenta.EOF And rsBuscaCuenta.BOF) Then
                                        oMov.ActualizaMovCta lnMovNro, lnItemPrin, , lnImpoPrin + (lnImporte - lnMonto) '' la dif manda a la ag princ
                                    Else
                                        oMov.ActualizaMovCta lnMovNro, lnItemDistri, , lnRegImporte + (lnImporte - lnMonto) '' si noexiste la age prin manda al ultimo age registrado
                                    End If
        
        '                            oMov.ActualizaMovCta lnMovNro, lnItemDistri, , lnRegImporte + (lnImporte - lnMonto)
                                End If
                                
                                RSClose rsAgesDistrib
            
                                Set rsAgesDistrib = Nothing
                                
        '                        pnItemDistri = lnItemDistri
                                psCtasInexis = lsCtasInexis

'                            End If
                        End If
                    End If
                Else
                    lnItem = IIf(lnItemDistri > 0, lnItemDistri + 1, IIf(lnItem > 1, lnItem + 1, lnItem))
                    lnItemDistri = lnItem
                    oMov.InsertaMovCta lnMovNro, lnItem, rsCtasCont!Código + lsSubCta, IIf(rsCtasCont!DH = "D", lnImporte, lnImporte * -1) * IIf(pnDocTpo = 7, -1, 1)
                End If
                '******************************* FIN PEAC
            Else
                lnImporte = rsCtasCont!Monto
                For I = 6 To rsCtasCont.Fields.Count - 1
                    rsImp.Move I - 6, adBookmarkFirst
                    If rsCtasCont.Fields(I) <> 0 Then
                        If Trim(rsCtasCont.Fields(I).Name) = "AJUSTE" Then
                            lnImporte = lnImporte + rsCtasCont.Fields(I)
                            oMov.InsertaMovOtrosItem lnMovNro, lnItem, rsImp!CtaCont, rsCtasCont.Fields(I), 1
                            'oMov.InsertaMovCta lnMovNro, lnItem, rsCtasCont!Código + lsSubCta, IIf(rsCtasCont!DH = "D", lnImporte, lnImporte * -1)
                        Else
                            'oMov.InsertaMovCta lnMovNro, lnItem, rsImp!CtaCont, IIf(rsImp!cDocImpDH = "D", lnImporte, lnImporte * -1)
                            oMov.InsertaMovOtrosItem lnMovNro, lnItem, rsImp!CtaCont, rsCtasCont.Fields(I), 1
                        End If
                    End If
                Next
             
                oMov.InsertaMovCta lnMovNro, lnItem, rsCtasCont!Código + lsSubCta, IIf(rsCtasCont!DH = "D", lnImporte, lnImporte * -1) * IIf(pnDocTpo = 7, -1, 1)
                
            End If
            
            lnOrdenObj = 0
            If Not rsObjetos Is Nothing Then
                rsObjetos.MoveFirst
                Do While Not rsObjetos.EOF
                    If rsObjetos!CtaCont = rsCtasCont!Código And rsObjetos!ItemCtaCont = rsCtasCont!ItemCtaCont Then
                    Select Case rsObjetos!ObjPadre
                        Case ObjCMACAgencias
                            lnOrdenObj = lnOrdenObj + 1
                            oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!ObjPadre
                            oMov.InsertaMovObjAgenciaArea lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código, ""
                        Case ObjCMACArea
                            lnOrdenObj = lnOrdenObj + 1
                            oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!ObjPadre
                            oMov.InsertaMovObjAgenciaArea lnMovNro, lnItem, lnOrdenObj, "", rsObjetos!Código
                        Case ObjCMACAgenciaArea
                            lsAreaCod = Mid(rsObjetos!Código, 1, 3)
                            lsAgeCod = Mid(rsObjetos!Código, 4, 2)
                            lnOrdenObj = lnOrdenObj + 1
                            oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!ObjPadre
                            oMov.InsertaMovObjAgenciaArea lnMovNro, lnItem, lnOrdenObj, lsAgeCod, lsAreaCod
                        Case ObjDescomEfectivo
                            lnOrdenObj = lnOrdenObj + 1
                            oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!ObjPadre
                            oMov.InsertaMovObjEfectivo lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código
                        Case ObjEntidadesFinancieras
                            lnOrdenObj = lnOrdenObj + 1
                            oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!ObjPadre
                            oMov.InsertaMovObjIF lnMovNro, lnItem, lnOrdenObj, Mid(rsObjetos!Código, 4, 13), Mid(rsObjetos!Código, 1, 2), Mid(rsObjetos!Código, 18, 10)
                            'falta insertar movObjIF
                        Case Else
                            lnOrdenObj = lnOrdenObj + 1
                            oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código
                    End Select
                    End If
                    rsObjetos.MoveNext
                Loop
            End If
            'For I = 6 To rsCtasCont.Fields.Count - 1
            '    rsImp.MoveFirst
            '    If rsCtasCont.Fields(I).value <> 0 Then
            '        rsImp.Find "Impuesto='" & rsCtasCont.Fields(I).Name & "'"
            '        If Not rsImp.EOF Then
            '           oMov.InsertaMovOtrosItem lnMovNro, lnItem, rsImp!CTACONT, rsCtasCont.Fields(I).value, ""
            '        End If
            '    End If
            'Next
            rsCtasCont.MoveNext
       End If
    Loop
End If
If Not rsImp Is Nothing Then
    rsImp.MoveFirst
    Do While Not rsImp.EOF
        If rsImp!Ok = "1" And (pnDestino = -1 Or pnDestino = 0) Then
            lnImporte = rsImp!Monto * IIf(rsImp!cDocImpDH = "D", 1, -1)
            
            If lnImporte <> 0 And rsImp!Impuesto <> "AJUSTE" Then
                lnItem = lnItem + 1
                oMov.InsertaMovCta lnMovNro, lnItem, rsImp!CtaCont, lnImporte * IIf(pnDocTpo = 7, -1, 1)
            End If
        ElseIf rsImp!Ok = "1" And (pnDestino = 1) Then
            lnImporte = rsImp!Monto * IIf(rsImp!cDocImpDH = "D", 1, -1)
            
            lnMontoNoGrabado = 0
            rsCtasCont.MoveFirst
            While Not rsCtasCont.EOF
                If rsCtasCont.Fields(4) = 1 Then
                    lnMontoNoGrabado = lnMontoNoGrabado + rsCtasCont.Fields(rsImp.Bookmark + 5)
                End If
                rsCtasCont.MoveNext
            Wend
            
            If rsImp.Fields(10) = 0 Then
                lnImporte = lnImporte - lnMontoNoGrabado
            End If
            
            If lnImporte <> 0 And rsImp!Impuesto <> "AJUSTE" Then
                lnItem = lnItem + 1
                oMov.InsertaMovCta lnMovNro, lnItem, rsImp!CtaCont, lnImporte * IIf(pnDocTpo = 7, -1, 1)
            End If
        ElseIf rsImp!Ok = "1" And rsImp.Fields(10) = 1 Then
            lnImporte = rsImp!Monto * IIf(rsImp!cDocImpDH = "D", 1, -1)
            
            If lnImporte <> 0 And rsImp!Impuesto <> "AJUSTE" Then
                lnItem = lnItem + 1
                oMov.InsertaMovCta lnMovNro, lnItem, rsImp!CtaCont, lnImporte * IIf(pnDocTpo = 7, -1, 1)
            End If
        End If
        rsImp.MoveNext
    Loop
End If
'Cuenta de Provisión / Caja Chica / A rendir Cuenta
lnItem = lnItem + 1
Dim oCont As NContFunciones
Set oCont = New NContFunciones

psCtaProvision = psCtaProvision + oCont.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaProvision, psAreaCh + psAgeCh, False)
Set oCont = Nothing

'***Modificado por ELRO el 20120921, según OYP-RFC111-2012
'lnItemDistri = IIf(psOpeCod = "401371" And Right(lsSubCta, 2) = "01", IIf(lnItemDistri = 0, lnItem, lnItemDistri + 1), lnItem)
lnItemDistri = IIf((psOpeCod = "401371" Or psOpeCod = "401344") And Right(lsSubCta, 2) = "01", IIf(lnItemDistri = 0, lnItem, lnItemDistri + 1), lnItem)
'***Fin Modificado por ELRO el 20120921*******************

'EJVG20140729 ***
If pnRetencSistPensProvMontoBase > 0 Then 'Entra al calculo acumulativo
    oMov.InsertaMovProvRetencSistPens lnMovNro, pnTpoProvSistPens, pnRetencSistPensProvMontoBase, pnRetencSistPensProvTotal
End If
If IsArray(poMatDatosRetencSistPensProv) Then
    If poMatDatosRetencSistPensProv(1) > 0 Then
        oMov.InsertaMovCta lnMovNro, lnItemDistri, poMatDatosRetencSistPensProv(2), poMatDatosRetencSistPensProv(1) * -1
        oMov.InsertaMovProvRetencSistPensDet lnMovNro, lnItemDistri, ConceptoRetencionSistemaPensionProveedor.Aporte, poMatDatosRetencSistPensProv(1), poMatDatosRetencSistPensProv(1)
        lnItemDistri = lnItemDistri + 1
    End If
    If poMatDatosRetencSistPensProv(3) > 0 Then
        oMov.InsertaMovCta lnMovNro, lnItemDistri, poMatDatosRetencSistPensProv(4), poMatDatosRetencSistPensProv(3) * -1
        oMov.InsertaMovProvRetencSistPensDet lnMovNro, lnItemDistri, ConceptoRetencionSistemaPensionProveedor.SeguroAFP, poMatDatosRetencSistPensProv(3), poMatDatosRetencSistPensProv(3)
        lnItemDistri = lnItemDistri + 1
    End If
    If poMatDatosRetencSistPensProv(5) > 0 Then
        oMov.InsertaMovCta lnMovNro, lnItemDistri, poMatDatosRetencSistPensProv(6), poMatDatosRetencSistPensProv(5) * -1
        oMov.InsertaMovProvRetencSistPensDet lnMovNro, lnItemDistri, ConceptoRetencionSistemaPensionProveedor.ComsionAFP, poMatDatosRetencSistPensProv(5), poMatDatosRetencSistPensProv(5)
        lnItemDistri = lnItemDistri + 1
    End If
End If
'END EJVG *******

oMov.InsertaMovCta lnMovNro, lnItemDistri, psCtaProvision, pnTotal * -1 * IIf(pnDocTpo = 7, -1, 1)

oMov.InsertaMovObj lnMovNro, lnItem, 1, ObjCMACAgenciaArea
oMov.InsertaMovObjAgenciaArea lnMovNro, lnItem, 1, psAgeCh, psAreaCh
''Grabamos Documento en MovDoc
oMov.InsertaMovDoc lnMovNro, pnDocTpo, psDocNro, Format(pdDocFecha, gsFormatoFecha)
'Grabamos la Referencia al Movimiento que se regulariza
oMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, pnNroProc, gCHTipoProcEgresoDirecto


'--Calculo del ITF
'Dim lnMontoITF As Currency
Dim lnMontoITF As Double '*** PEAC 20110331

'If nTasaITF <> 0 And Not pnDocTpo = TpoDocRecEgreso Then
'SE ELIMINO EL ITF EN EL SUSTENTO PARA CUSCO
'If nTasaITF <> 0 Then
'   lnMontoITF = Round(pnTotal * nTasaITF, 2)
'   If lsCtaITFD <> "" And lsCtaITFH <> "" And nTasaITF > 0 Then
'       lnItem = lnItem + 1
'       oMov.InsertaMovCta lnMovNro, Format(lnItem, "#0"), lsCtaITFD, lnMontoITF
'       lnItem = lnItem + 1
'       oMov.InsertaMovCta lnMovNro, Format(lnItem, "#0"), lsCtaITFH, lnMontoITF * -1
'   End If
'End If

If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
    oMov.GeneraMovME lnMovNro, psMovNro, pnTCVariable
End If

'*** PEAC 20110314 - Grabamos el asiento de cuentas de orden de los Bienes No Depreciables
lnMontoCtaOrd = 0
rsCtasCont.MoveFirst
Do While Not rsCtasCont.EOF
    sTexto = ""
    If Not rsObjetos Is Nothing Then
        rsObjetos.MoveFirst
        Do While Not rsObjetos.EOF
            If rsObjetos!CtaCont = rsCtasCont!Código And rsObjetos!ItemCtaCont = rsCtasCont!ItemCtaCont Then
                sTexto = sTexto + rsObjetos![SubCta]
            End If
            rsObjetos.MoveNext
        Loop
    End If

    If Len(sTexto) = 2 Then
        If rsCtasCont!Código = "45" & Mid(psOpeCod, 3, 1) & "3011101" Then
            lnMontoCtaOrd = lnMontoCtaOrd + rsCtasCont!Monto
        End If
    Else
        If rsCtasCont!Código + Left(sTexto, 2) = "45" & Mid(psOpeCod, 3, 1) & "3011101" Then
            lnMontoCtaOrd = lnMontoCtaOrd + rsCtasCont!Monto
        End If
    End If
    rsCtasCont.MoveNext
Loop

If lnMontoCtaOrd > 0 Then
    Set rsCtasOrden = oMov.GetCuentaOrdenBienNoDepre(Mid(psOpeCod, 3, 1), Right(lsSubCta, 2))
    If Not (rsCtasOrden.EOF And rsCtasOrden.BOF) Then
        lnItemDistri = lnItemDistri + 1
        oMov.InsertaMovCta lnMovNro, lnItemDistri, rsCtasOrden!cCtaDebe, lnMontoCtaOrd
        lnItemDistri = lnItemDistri + 1
        oMov.InsertaMovCta lnMovNro, lnItemDistri, rsCtasOrden!cCtaHaber, lnMontoCtaOrd * -1
    End If
End If

''*** FIN PEAC

oMov.CommitTrans

lbTrans = False
Set oMov = Nothing
Exit Sub
GrabaSolEgresoDirectoErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function GetSolEgresoDirectoPend(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnProcCH As Integer, ByVal psCtaFondFijo As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New Recordset

If oCon.AbreConexion = False Then Exit Function
 
sql = "Select   0 as OK , D.CDOCABREV , MD.CDOCNRO AS cDocNro,  " _
    & "         CONVERT(VARCHAR(12),MD.dDocFecha,103) AS dDocFecha , P.CPERSNOMBRE, ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)) AS IMPORTE , " _
    & "         M.CMOVDESC , M.nMovNro, MD.nDocTpo, M.cMovNro " _
    & " From    MovCajaChica  MCH " _
    & "         JOIN MOV M          ON M.nMOVNRO = MCH.nMOVNRO " _
    & "         JOIN CajaChica CH   ON CH.CAREACOD=MCH.CAREACOD " _
    & "                             AND CH.CAGECOD = MCH.CAGECOD AND MCH.nProcNro = CH.nProcNro " _
    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO " _
    & "         LEFT JOIN  MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "         JOIN MOVDOC MD ON MD.nMOVNRO  = M.nMOVNRO " _
    & "         JOIN DOCUMENTO D ON D.nDocTpo = MD.nDocTpo " _
    & "         JOIN MOVGASTO MG ON MG.nMOVNRO =M.nMOVNRO " _
    & "         JOIN PERSONA  P ON P.CPERSCOD = MG.CPERSCOD " _
    & " Where   MCH.cProcTpo ='" & gCHTipoProcEgresoDirecto & "' AND M.nMovEstado IN (" & gMovEstContabNoContable & ", " & gMovEstContabPendiente & " ) " _
    & "         AND MC.CCTACONTCOD LIKE '" & psCtaFondFijo & "%' AND m.cOpeCOD in ('" & gCHEgreDirectoSolMN & "','" & gCHEgreDirectoAtencME & "') " _
    & "         AND NOT EXISTS (SELECT  M1.CMOVNRO " _
    & "                         FROM    MOV M1 JOIN MOVREF MR ON MR.nMOVNROREF = M1.nMOVNRO " _
    & "                         WHERE   MR.nMOVNRO=M.nMOVNRO AND M1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ")) " _
    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
    & "         AND MCH.cAreaCod='" & psAreaCh & "' AND MCH.cAgeCod='" & psAgeCh & "'  AND MCH.nProcNro <=" & pnProcCH & "" _
    & " ORDER BY M.nMovNro"

Set rs = oCon.CargaRecordSet(sql)
Set GetSolEgresoDirectoPend = rs

oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GrabaRechazoSolEgrDir(ByVal psMovNroRechazo As String, ByVal pnMovNroSol As Long, _
                                      ByVal psFormatoFecha As String, _
                                      ByVal psOpeCod As String, ByVal psGlosa As String) As Integer


 On Error GoTo GrabaRechazoSolEgrDirErr
    Dim oMov As DMov
    Dim sql As String
    Dim lnSaldoArendir As Currency
    Dim lbTrans As Boolean
    Dim lnMovNro As Long
    Set oMov = New DMov
    GrabaRechazoSolEgrDir = 1
    oMov.BeginTrans
    lbTrans = True
    oMov.InsertaMov psMovNroRechazo, psOpeCod, psGlosa, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(psMovNroRechazo)
    oMov.InsertaMovCont lnMovNro, 0, "0", "0"
    oMov.ActualizaMov pnMovNroSol, , gMovEstContabRechazado, gMovFlagEliminado
    oMov.InsertaMovRef lnMovNro, pnMovNroSol
    oMov.CommitTrans
    lbTrans = False
    GrabaRechazoSolEgrDir = 0
    
    Exit Function
GrabaRechazoSolEgrDirErr:
    If lbTrans Then
        lbTrans = False
        oMov.RollbackTrans
    End If
    Call RaiseError(MyUnhandledError, "NARendir:GrabaRechazoSolEgrDir Method")
End Function
Public Function GrabaAtenEgresoDirec(ByVal psFormatoFecha As String, _
                                     ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                     ByVal pnMovNroSol As Long, _
                                     ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                     ByVal pnNroProc As Integer, ByVal pnImporte As Currency) As Integer
Dim oDMov As DMov
Dim lnSaldo As Currency
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov

On Error GoTo GrabaAtenEgresoDirecErr
oDMov.BeginTrans
lbTrans = True
oDMov.ActualizaMovimiento psMovNro, pnMovNroSol, gMovEstContabNoContable, , True, False, psOpeCod, , , False
'oDMov.ActualizaMov pnMovNroSol, , gMovEstContabNoContable, , psMovNro, False
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovRef lnMovNro, pnMovNroSol
oDMov.ActualizaMovCajaChicaNroProc lnMovNro, pnNroProc

lnSaldo = GetDatosCajaChica(psAreaCh, psAgeCh)
lnSaldo = lnSaldo - pnImporte
oDMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProc, , , , , lnSaldo
oDMov.CommitTrans
lbTrans = False
GrabaAtenEgresoDirec = 0
Set oDMov = Nothing
Exit Function
GrabaAtenEgresoDirecErr:
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    Err.Raise vbObjectError + 100, "GrabaAtenEgresoDirec", "Error Grabar Atencion de Egreso Directo"
End Function
Public Function GetSolEgresoDirectoAtend(ByVal psAreaCh As String, _
                                        ByVal psAgeCh As String, _
                                        ByVal pnProcCH As Integer, _
                                        ByVal psCtaFondFijo As String, _
                                        Optional ByVal pbMuestraCheck As Boolean = True) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New Recordset

If oCon.AbreConexion = False Then Exit Function
 ', ATENC.cMovNroSol
sql = "Select   " & IIf(pbMuestraCheck, " 0 as OK,", "") & " D.CDOCABREV , MD.CDOCNRO AS DOCNRO,  " _
    & "         CONVERT(VARCHAR(12),MD.dDocFecha,103) AS dDocFecha , P.CPERSNOMBRE, ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)) AS IMPORTE , " _
    & "         M.CMOVDESC, M.nMovNro AS cMovNroAtenc,MD.nDocTpo, M1.CMOVNRO AS cMovnroSol, MCH.nProcNro  " _
    & " From    MovCajaChica  MCH " _
    & "         JOIN MOV M          ON M.nMOVNRO = MCH.nMOVNRO " _
    & "         JOIN CajaChica CH   ON CH.CAREACOD=MCH.CAREACOD " _
    & "                             AND CH.CAGECOD = MCH.CAGECOD AND MCH.nProcNro = CH.nProcNro " _
    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO " _
    & "         LEFT JOIN  MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "         JOIN MOVDOC MD ON MD.nMOVNRO  = M.nMOVNRO " _
    & "         JOIN DOCUMENTO D ON D.nDocTpo = MD.nDocTpo " _
    & "         JOIN MOVGASTO MG ON MG.nMOVNRO =M.nMOVNRO " _
    & "         LEFT JOIN PERSONA  P ON P.CPERSCOD = MG.CPERSCOD  " _
    & "         JOIN MOVREF MR ON MR.nMOVNRO = M.nMOVNRO  JOIN MOV M1 ON M1.NMOVNRO =MR.NMOVNROREF " _
    & " Where   MCH.cProcTpo ='" & gCHTipoProcEgresoDirecto & "' " _
    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
    & "         AND MC.CCTACONTCOD LIKE '" & psCtaFondFijo & "%'  AND M.COPECOD IN ('" & gCHEgreDirectoAtencMN & "','" & gCHEgreDirectoAtencMN & "') " _
    & "         AND MCH.cAreaCod='" & psAreaCh & "' AND MCH.cAgeCod='" & psAgeCh & "' and MCH.nProcNro=" & pnProcCH & " " _
    & "ORDER BY M.nMovNro "

Set rs = oCon.CargaRecordSet(sql)
Set GetSolEgresoDirectoAtend = rs

oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GrabaExtornoAtenEgresoDirec(ByVal psFormatoFecha As String, _
                                     ByVal psMovNroExt As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                     ByVal pnMovNroAtenc As Long, ByVal psMovNroSol As String, _
                                     ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                     ByVal pnNroProc As Integer, ByVal pnImporte As Currency) As Integer
Dim oDMov As DMov
Dim lnSaldo As Currency
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov

On Error GoTo GrabaExtornoAtenEgresoDirecErr

GrabaExtornoAtenEgresoDirec = 1

oDMov.BeginTrans
lbTrans = True

oDMov.InsertaMov psMovNroExt, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagDeExtorno
lnMovNro = oDMov.GetnMovNro(psMovNroExt)
oDMov.InsertaMovCont lnMovNro, pnImporte, 0, 0
'se actualiza la atencion con la solicitud de nuevo
'los movimientos de la atencion pasan a ser de la solicitud
oDMov.ActualizaMovimiento psMovNroSol, pnMovNroAtenc, gMovEstContabNoContable, , False, True, , , False
oDMov.InsertaMovRef lnMovNro, pnMovNroAtenc

lnSaldo = GetDatosCajaChica(psAreaCh, psAgeCh)
lnSaldo = lnSaldo + pnImporte
oDMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProc, , , , , lnSaldo
oDMov.CommitTrans
lbTrans = False
GrabaExtornoAtenEgresoDirec = 0
Set oDMov = Nothing
Exit Function
GrabaExtornoAtenEgresoDirecErr:
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    
    Err.Raise vbObjectError + 100, "GrabaExtornoAtenEgresoDirec", "Error Grabar Extorno de Atencion de Egreso Directo"
End Function
Public Function GetRSDatosCajaChica(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnProcCH As Integer, Optional ByRef oMov As DMov) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New Recordset

If oCon.AbreConexion = False Then Exit Function
 
'*** PEAC 20110829
sql = "SELECT   CH.cAreaCod, CH.cAgeCod, CH.nProcNro, nMontoAsig , nSaldo,  " _
    & "         CH.cPersCod , P.cPersNombre, CONVERT(VARCHAR(12),CONVERT(DATETIME, SUBSTRING(M.cMovNro,1,8)),103) AS dDesembolso, " _
    & "         nTopeEgresos,nMontoDesem,nSaldoAnt,nSaldo " _
    & " FROM    CAJACHICA CH " _
    & "         LEFT JOIN " & lsBaseCom & "PERSONA P ON P.CPERSCOD = CH.CPERSCOD " _
    & "         LEFT JOIN MOVCAJACHICA MCH ON MCH.cAreaCod = CH.cAreaCod AND MCH.cAgeCod= CH.cAgeCod " _
    & "                                 AND MCH.nProcNro = CH.nProcNro " _
    & "         JOIN MOV M ON M.NMOVNRO =MCH.NMOVNRO  " _
    & " WHERE   CH.cAreaCod = '" & psAreaCh & "' AND CH.cAgeCod = '" & psAgeCh & "' AND CH.nProcNro  = " & pnProcCH & ""

'    sql = "EXEC stp_sel_GetRSDatosCajaChica '" & psAreaCh & "','" & psAgeCh & "'," & pnProcCH

'AND MCH.cProcTpo in ('" & gCHTipoProcDesembolso & "','" & gCHTipoProcHabilitacion & "')"
If oMov Is Nothing Then
    Set rs = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    Set oCon = Nothing
Else
    Set rs = oMov.CargaRecordSet(sql)
End If
Set GetRSDatosCajaChica = rs

End Function
Public Function GetCHTodosAtenArendir(ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                      ByVal pnProcCH As Integer, ByVal psCtaArendir As String, _
                                      ByVal psCtaPendiente As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lnMovNroRendAnt As Long
Set oCon = New DConecta
Set rs = New Recordset

If oCon.AbreConexion = False Then Exit Function
sql = "SELECT ISNULL(MAX(nMovNro),0) nMovNroRend FROM MovCajaChica WHERE cAreaCod = '" & psAreaCh & "' and cAgeCod = '" & psAgeCh & "' " _
     & " and cProcTpo = '" & gCHTipoProcRendicion & "' and nProcNro < " & pnProcCH & " "
Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF Then
    lnMovNroRendAnt = rs!nMovNroRend
End If
'***Modificado por ELRO el 20130221, según SATI INC1301300007
'sql = "SELECT   MD.CDOCNRO, P.CPERSNOMBRE, CONVERT(CHAR(12), MD.dDocFecha ,103  ) AS Fecha ,          " _
'    & "         ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS IMPORTE, CASE WHEN REND.cMovNRoRend IS NULL THEN ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) ELSE nMovSaldo END nMovSaldo, M1.CMOVDESC AS DESCRIPCION, M1.CMOVNRO AS CMOVNROATENC, " _
'    & "         ISNULL(REND.CMOVNROREND,'') AS CMOVNROREND , ISNULL(Rend.MontoRendido,0), MCH.nProcNro " _
'    & " FROM    MOV  M JOIN MOVCONT MCT ON MCT.nMOVNRO = M.nMOVNRO " _
'    & "         JOIN MOVARENDIR MAR ON MAR.nMOVNRO = M.nMOVNRO " _
'    & "         JOIN PERSONA P     ON P.CPERSCOD = MAR.CPERSCOD " _
'    & "         JOIN MOVCAJACHICA MCH ON MCH.nMOVNRO = MAR.nMOVNRO " _
'    & "         JOIN MOVDOC MD ON MD.nMOVNRO = M.nMOVNRO " _
'    & "         JOIN DOCUMENTO  D ON D.nDocTpo = MD.nDocTpo " _
'    & "         JOIN MOVREF MR ON MR.nMOVNROREF= M.nMOVNRO " _
'    & "         JOIN MOV M1 ON M1.nMOVNRO = MR.nMOVNRO " _
'    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M1.nMOVNRO " _
'    & "         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
'    & "         LEFT JOIN ( SELECT  M2.cMOVNRO AS CMOVNROREND, M2.nMOVNRO AS nMOVNROREND, MR2.nMOVNROREF as nMovNroAtenc, " _
'    & "                             IsNull(Sum(IsNull(ME1.NMOVMEIMPORTE, IsNull(MC2.NMOVIMPORTE, 0))), 0) As MontoRendido " _
'    & "                     FROM    MOV M2 JOIN MOVREF MR2 ON M2.nMOVNRO = MR2.nMOVNRO " _
'    & "                             LEFT JOIN MOVCTA MC2 ON MC2.nMOVNRO = M2.nMOVNRO  AND MC2.cCtaContCod IN ('" & psCtaArendir & "','" & psCtaPendiente & "') " _
'    & "                             LEFT JOIN MOVME ME1 ON ME1.nMOVNRO = MC2.nMOVNRO AND ME1.nMOVITEM = MC2.nMOVITEM " _
'    & "                     WHERE   SUBSTRING(M2.COPECOD,1,5) IN ('" & Mid(gCHArendirCtaRendMN, 1, 5) & "','" & Mid(gCHArendirCtaRendME, 1, 5) & "') " _
'    & "                             AND M2.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
'    & "                     GROUP BY    M2.cMOVNRO, M2.nMOVNRO ,  MR2.nMOVNROREF ) as REND " _
'    & "         ON REND.nMovNroAtenc = M1.nMOVNRO " _
'    & " WHERE   MAR.cTpoArendir='" & gArendirTipoCajaChica & "'  AND cCtaContCod IN ('" & psCtaArendir & "')  " _
'    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
'    & "         AND M1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") "
'sql = sql + "      AND MCH.cAreaCod='" & psAreaCh & "' AND MCH.cAgeCod='" & psAgeCh & "' AND MCH.nProcNro=" & pnProcCH & " " _
'    & " Union "
'sql = sql + "SELECT   MD.CDOCNRO, P.CPERSNOMBRE, CONVERT(CHAR(12), MD.dDocFecha ,103  ) AS Fecha ,          " _
'    & "         ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS IMPORTE, CASE WHEN REND.cMovNRoRend IS NULL THEN ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) ELSE nMovSaldo END nMovSaldo, M1.CMOVDESC AS DESCRIPCION, M1.CMOVNRO AS CMOVNROATENC, " _
'    & "         ISNULL(REND.CMOVNROREND,'') AS CMOVNROREND , ISNULL(Rend.MontoRendido,0), MCH.nProcNro " _
'    & " FROM    MOV  M JOIN MOVCONT MCT ON MCT.nMOVNRO = M.nMOVNRO " _
'    & "         JOIN MOVARENDIR MAR ON MAR.nMOVNRO = M.nMOVNRO " _
'    & "         JOIN PERSONA P     ON P.CPERSCOD = MAR.CPERSCOD " _
'    & "         JOIN MOVCAJACHICA MCH ON MCH.nMOVNRO = MAR.nMOVNRO " _
'    & "         JOIN MOVDOC MD ON MD.nMOVNRO = M.nMOVNRO " _
'    & "         JOIN DOCUMENTO  D ON D.nDocTpo = MD.nDocTpo " _
'    & "         JOIN MOVREF MR ON MR.nMOVNROREF= M.nMOVNRO " _
'    & "         JOIN MOV M1 ON M1.nMOVNRO = MR.nMOVNRO " _
'    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M1.nMOVNRO " _
'    & "         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
'    & "         LEFT JOIN ( SELECT  M2.cMOVNRO AS CMOVNROREND, M2.nMOVNRO AS nMOVNROREND, mch2.nProcNro, MR2.nMOVNROREF as nMovNroAtenc, " _
'    & "                             IsNull(Sum(IsNull(ME1.NMOVMEIMPORTE, IsNull(MC2.NMOVIMPORTE, 0))), 0) As MontoRendido " _
'    & "                     FROM    MOV M2 JOIN MovCajaChica mch2 ON mch2.nMovNro = m2.nMovNro JOIN MOVREF MR2 ON M2.nMOVNRO = MR2.nMOVNRO " _
'    & "                             LEFT JOIN MOVCTA MC2 ON MC2.nMOVNRO = M2.nMOVNRO  AND MC2.cCtaContCod IN ('" & psCtaArendir & "','" & psCtaPendiente & "') " _
'    & "                             LEFT JOIN MOVME ME1 ON ME1.nMOVNRO = MC2.nMOVNRO AND ME1.nMOVITEM = MC2.nMOVITEM " _
'    & "                     WHERE   SUBSTRING(M2.COPECOD,1,5) IN ('" & Mid(gCHArendirCtaRendMN, 1, 5) & "','" & Mid(gCHArendirCtaRendME, 1, 5) & "') " _
'    & "                             AND M2.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
'    & "                     GROUP BY    M2.cMOVNRO , M2.nMOVNRO, mch2.nProcNro, MR2.nMOVNROREF ) as REND " _
'    & "         ON REND.nMovNroAtenc = M1.nMOVNRO " _
'    & " WHERE   MAR.cTpoArendir='" & gArendirTipoCajaChica & "'  AND cCtaContCod IN ('" & psCtaArendir & "')  " _
'    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
'    & "         AND M1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") "
'sql = sql + "      AND MCH.cAreaCod='" & psAreaCh & "' AND MCH.cAgeCod='" & psAgeCh & "' AND MCH.nProcNro < " & pnProcCH & " " _
'    & "    And (REND.nProcNro = " & pnProcCH & " OR REND.nMOVNROREND Is Null) "
sql = "exec stp_sel_DevolverTodasRendiciones '" & psAreaCh & "', '" & psAgeCh & "', " & pnProcCH & ", '" & psCtaArendir & "', '" & psCtaPendiente & "'"
'***Fin Modificado por ELRO el 20130221***********************
Set rs = oCon.CargaRecordSet(sql)
Set GetCHTodosAtenArendir = rs

oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetCHSustAtenArendir(ByVal psCtaArendir As String, _
                                     ByVal psCtaPendiente As String, _
                                     ByVal psMovNroAtenc As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New Recordset

If oCon.AbreConexion = False Then Exit Function

sql = "SELECT   D.CDOCABREV, MD.CDOCNRO, CONVERT(VARCHAR(12),MD.DDOCFECHA,103) AS FECHADOC," _
    & "         isnull(P.CPERSNOMBRE,'') CPERSNOMBRE, ABS(SUM(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE))) AS Importe, " _
    & "         M.cMovNro , M1.CMOVNRO AS CMOVNROREF , M.cMovDesc AS Glosa" _
    & " FROM    MOV M " _
    & "         JOIN MOVREF MR ON MR.nMOVNRO = M.nMOVNRO  JOIN MOV M1 ON M1.NMOVNRO = MR.NMOVNROREF " _
    & "         LEFT JOIN MOVGASTO  MG ON MG.nMOVNRO = M.nMOVNRO " _
    & "         LEFT JOIN PERSONA P ON P.CPERSCOD = MG.CPERSCOD " _
    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.nMOVNRO=MC.nMOVNRO AND MC.nMOVITEM = ME.nMOVITEM " _
    & "         JOIN MOVDOC MD ON MD.nMOVNRO =M.nMOVNRO " _
    & "         JOIN DOCUMENTO D ON D.nDocTpo = MD.nDocTpo " _
    & " WHERE   M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
    & "         AND M1.CMOVNRO='" & psMovNroAtenc & "' " _
    & "         AND MC.CCTACONTCOD IN('" & psCtaArendir & "','" & psCtaPendiente & "') " _
    & " GROUP BY D.CDOCABREV, MD.CDOCNRO, MD.DDOCFECHA , isnull(P.CPERSNOMBRE,'') , M.CMOVNRO, M1.CMOVNRO , M.cMovDesc "

Set rs = oCon.CargaRecordSet(sql)
Set GetCHSustAtenArendir = rs

oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetDatosRendArendirCH(ByVal psMovNroAtenc As String, ByVal psCtaFondoFijo As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New Recordset

If oCon.AbreConexion = False Then Exit Function

'***Modificado por ELRO el 20130222, según SATI INC1301300007
'sql = " Select  CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) AS FechaRend ,  " _
'    & "         M.CMOVNRO, M.cMovDesc, ISNULL(ME.NMOVMEIMPORTE,MC.nMovImporte) AS nMovImporte " _
'    & " FROM    MOV M JOIN MOVREF MR ON MR.nMOVNRO = M.nMOVNRO JOIN MOV M1 ON M1.NMOVNRO=MR.NMOVNROREF " _
'    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO " _
'    & "         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
'    & " WHERE   M1.CMOVNRO ='" & psMovNroAtenc & "' AND MC.CCTACONTCOD LIKE '" & psCtaFondoFijo & "%'"
sql = "exec stp_sel_DevolverIngresosPorRendicion '" & psMovNroAtenc & "', '" & psCtaFondoFijo & "'"
'***Fin Modificado por ELRO el 20130222**********************
Set rs = oCon.CargaRecordSet(sql)
Set GetDatosRendArendirCH = rs

oCon.CierraConexion
Set oCon = Nothing

End Function
Public Function GetAsientoCH(ByVal psAreaCh As String, ByVal psAgeCh As String, _
                              ByVal pnProcCH As Integer) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim oCon As New DConecta
Set oCon = New DConecta
If oCon.AbreConexion = False Then Exit Function

'***Modificado por ELRO el 20120709, según OYP-RFC047
'sql = "SELECT  MC.CCTACONTCOD, MIN(dbo.GetCtaContDesc(mc.cCtaContCod,2,1)) cCtaContDesc, " _
'    & "         SUM(CASE WHEN MC.NMOVIMPORTE>0 THEN ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) ELSE 0 END ) AS DEBE, " _
'    & "         SUM(CASE WHEN MC.NMOVIMPORTE<0 THEN ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)) ELSE 0 END)  AS HABER " _
'    & " FROM    MOV M JOIN MOVCAJACHICA MCH ON M.NMOVNRO = MCH.NMOVNRO " _
'    & "         JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO " _
'    & "         LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO AND ME.NMOVITEM = MC.NMOVITEM " _
'    & " WHERE   SUBSTRING(M.COPECOD,1,5) NOT IN ('" & Mid(gCHDesembOrdenPagoMN, 1, 5) & "','" & Mid(gCHDesembOrdenPagoME, 1, 5) & "','" & Mid(gCHArendirCtaRendMN, 1, 5) & "','" & Mid(gCHArendirCtaRendME, 1, 5) & "') AND M.COPECOD NOT IN ('" & gCHArendirCtaSustMN & "','" & gCHArendirCtaSustME & "') AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
'    & "         AND MCH.CAREACOD='" & psAreaCh & "' AND MCH.CAGECOD='" & psAgeCh & "' AND nProcNro = " & pnProcCH & " and not m.nMovEstado = " & gMovEstContabPendiente & " and not mc.cCtaContCod = '' " _
'    & " GROUP BY MC.CCTACONTCOD UNION ALL " _
'    & "SELECT  MC.CCTACONTCOD, MIN(dbo.GetCtaContDesc(mc.cCtaContCod,2,1)) cCtaContDesc, " _
'    & "         SUM(CASE WHEN MC.NMOVIMPORTE>0 THEN ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) ELSE 0 END ) AS DEBE, " _
'    & "         SUM(CASE WHEN MC.NMOVIMPORTE<0 THEN ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)) ELSE 0 END)  AS HABER " _
'    & " FROM    MOVCAJACHICA MCH JOIN MovRef Ate ON Ate.nMovNroRef = mch.nMovNro JOIN MovRef mr ON mr.nMovNroRef = Ate.nMovNro " _
'    & "        JOIN Mov m ON m.nMovNro = mr.nMovNro JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO " _
'    & "        LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO AND ME.NMOVITEM = MC.NMOVITEM JOIN MovCajaChica mchS ON mchS.nMovNro = m.nMovNro " _
'    & " WHERE   SUBSTRING(M.COPECOD,1,5) NOT IN ('" & Mid(gCHDesembOrdenPagoMN, 1, 5) & "','" & Mid(gCHDesembOrdenPagoME, 1, 5) & "' ) AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
'    & "         AND mchS.cProctpo = 6 and not cCtaContCod = '' " _
'    & "         AND MCH.CAREACOD='" & psAreaCh & "' AND MCH.CAGECOD='" & psAgeCh & "' AND mchS.nProcNro = " & pnProcCH & " and not m.nMovEstado IN (" & gMovEstContabPendiente & "," & gMovEstContabSustPendRendir & ") " _
'    & "GROUP BY MC.CCTACONTCOD ORDER BY mc.cCtaContCod "
sql = "exec stp_sel_DevolverAsientoCH '" & psAreaCh & "', '" & psAgeCh & "', '" & pnProcCH & "'"
'***Fin Modificado por ELRO**************************

Set rs = oCon.CargaRecordSet(sql)
Set GetAsientoCH = rs
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function GetDocumentosAsientoCH(ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                     ByVal pnProcCH As Integer, ByVal psCtaFondoFijo As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim oCon As New DConecta
Set oCon = New DConecta
If oCon.AbreConexion = False Then Exit Function

'***Modificado por ELRO el 20120626, según OYP-RFC047-2012
'sql = "SELECT   MC.CCTACONTCOD, dbo.GetCtaContDesc(mc.cCtaContCod,2,1) cCtaContDesc, ISNULL(D.CDOCABREV,'') AS CDOCABREV, ISNULL(MD.CDOCNRO,'') AS CDOCNRO, " _
'    & "         ISNULL(Convert(Varchar(10),MD.dDocFecha,103),'') as dDocFecha, M.CMOVDESC, " _
'    & "         CASE WHEN MC.NMOVIMPORTE>0 THEN ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE ) ELSE 0 END  AS DEBE, " _
'    & "         CASE WHEN MC.NMOVIMPORTE<0 THEN ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE )) ELSE 0 END  AS HABER " _
'    & " FROM    MOV  M JOIN MOVCAJACHICA MCH ON M.NMOVNRO = MCH.NMOVNRO " _
'    & "         JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO AND ME.NMOVITEM = MC.NMOVITEM " _
'    & "         LEFT JOIN MOVDOC MD ON MD.NMOVNRO = M.NMOVNRO " _
'    & "         LEFT JOIN DOCUMENTO D ON D.nDocTpo = MD.nDocTpo " _
'    & " WHERE   SUBSTRING(M.COPECOD,1,5) NOT IN ('30059','" & Mid(gCHDesembOrdenPagoMN, 1, 5) & "','" & Mid(gCHDesembOrdenPagoME, 1, 5) & "','" & Mid(gCHArendirCtaRendMN, 1, 5) & "','" & Mid(gCHArendirCtaRendME, 1, 5) & "') AND M.COPECOD NOT IN ('" & gCHArendirCtaSustMN & "','" & gCHArendirCtaSustME & "') AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
'    & "         AND NOT MC.cCtaContCod LIKE '" & psCtaFondoFijo & "%' and not m.nMovEstado = " & gMovEstContabPendiente & " " _
'    & "         AND MCH.CAREACOD='" & psAreaCh & "' AND MCH.CAGECOD='" & psAgeCh & "' AND nProcNro = " & pnProcCH & "  and NOT mc.cCtaContCod = '' UNION ALL " _
'    & "SELECT   MC.CCTACONTCOD, dbo.GetCtaContDesc(mc.cCtaContCod,2,1) cCtaContDesc, ISNULL(D.CDOCABREV,'') AS CDOCABREV, ISNULL(MD.CDOCNRO,'') AS CDOCNRO, " _
'    & "         ISNULL(Convert(Varchar(10),MD.dDocFecha,103),'') as dDocFecha, M.CMOVDESC, " _
'    & "         CASE WHEN MC.NMOVIMPORTE>0 THEN ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE ) ELSE 0 END  AS DEBE, " _
'    & "         CASE WHEN MC.NMOVIMPORTE<0 THEN ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE )) ELSE 0 END  AS HABER " _
'    & " FROM    MOVCAJACHICA MCH JOIN MovRef Ate ON Ate.nMovNroRef = mch.nMovNro JOIN MovRef mr ON mr.nMovNroRef = Ate.nMovNro " _
'    & "         JOIN Mov m ON m.nMovNro = mr.nMovNro JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO " _
'    & "         LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO AND ME.NMOVITEM = MC.NMOVITEM " _
'    & "         LEFT JOIN MOVDOC MD ON MD.NMOVNRO = M.NMOVNRO " _
'    & "         LEFT JOIN DOCUMENTO D ON D.nDocTpo = MD.nDocTpo JOIN MovCajaChica mchS ON mchS.nMovNro = m.nMovNro " _
'    & " WHERE   SUBSTRING(M.COPECOD,1,5) NOT IN ('300591','" & Mid(gCHDesembOrdenPagoMN, 1, 5) & "','" & Mid(gCHDesembOrdenPagoME, 1, 5) & "') AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
'    & "         AND NOT MC.cCtaContCod LIKE '" & psCtaFondoFijo & "%' and not m.nMovEstado in (" & gMovEstContabPendiente & "," & gMovEstContabSustPendRendir & ") " _
'    & "         AND mchS.cProctpo = 6  and NOT mc.cCtaContCod = '' " _
'    & "         AND MCH.CAREACOD='" & psAreaCh & "' AND MCH.CAGECOD='" & psAgeCh & "' AND mchS.nProcNro = " & pnProcCH & " "
'    '& " ORDER BY MC.cCtaContCod, cDocAbrev, MD.cDocNro "
sql = "exec stp_sel_DevolverDocumentosAsientoCH '" & psAreaCh & "', '" & psAgeCh & "', " & pnProcCH & ", '" & psCtaFondoFijo & "'"
'***Modificado por ELRO***********************************
Set rs = oCon.CargaRecordSet(sql)
Set GetDocumentosAsientoCH = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GrabaRendicionCh(ByVal psMovNroRend As String, _
                            ByVal psFormatoFecha As String, _
                            ByVal psOpeCod As String, ByVal psGlosa As String, _
                            ByVal pnMontoRendido As String, ByVal pnSaldo As Currency, _
                            ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal pnNroProc As Integer, _
                            ByVal pdFecsis As Date, ByVal psCodUser As String, ByVal psAgencia As String) As Integer

 On Error GoTo GrabaRendicionChErr
 Dim rs As ADODB.Recordset
 Dim lnNuevoProcCH As Integer
 Dim lbTrans As Boolean
 Dim lnMovNro As Long
 Dim oMov As DMov
 Dim sql As String
 Dim lnSaldoArendir As Currency
 Set oMov = New DMov
 Dim lnTotalEgresos As Currency '***Agregado por ELRO el 20120828, según OYP-RFC104-2012
 
    GrabaRendicionCh = 1
    oMov.BeginTrans
    lbTrans = True
    oMov.InsertaMov psMovNroRend, psOpeCod, psGlosa
    lnMovNro = oMov.GetnMovNro(psMovNroRend)
    oMov.InsertaMovCont lnMovNro, pnMontoRendido, "0", "0"
    oMov.InsertaMovCajaChica lnMovNro, psAreaCod, psAgeCod, pnNroProc, gCHTipoProcRendicion
    'ver***************************************
    ActualizaMovCajaChica oMov, lnMovNro, psAreaCod, psAgeCod, pnNroProc, pdFecsis, psCodUser, psAgencia
    Set rs = New ADODB.Recordset
    Set rs = GetRSDatosCajaChica(psAreaCod, psAgeCod, pnNroProc, oMov)
    If Not rs.EOF And Not rs.BOF Then
        lnNuevoProcCH = GetNuevaCajaChica(psAreaCod, psAgeCod)
        '***Agregado por ELRO el 20120828, según OYP-RFC104-2012
        lnTotalEgresos = oMov.devolverTotalEgresos_2(psAreaCod, psAgeCod, lnNuevoProcCH)
        '***Fin Agregado por ELRO el 20120828*******************
        '***Modificado por ELRO el 20120828, según OYP-RFC104-2012
        'oMov.InsertaCajaChica psAreaCod, psAgeCod, lnNuevoProcCH, rs!cPerscod, rs!nMontoAsig, _
        '            rs!nTopeEgresos, rs!nMontoDesem, pnSaldo, pnSaldo
        oMov.InsertaCajaChica psAreaCod, psAgeCod, lnNuevoProcCH, rs!cPerscod, rs!nMontoAsig, _
                    rs!nTopeEgresos, rs!nMontoDesem, pnSaldo, IIf(lnTotalEgresos = 0, pnSaldo, pnSaldo - lnTotalEgresos)
        '***Fin Modificado por ELRO el 20120828*******************
        oMov.InsertaMovCajaChica lnMovNro, psAreaCod, psAgeCod, lnNuevoProcCH, gCHTipoProcNuevaPorRendicion
    End If
    rs.Close
    Set rs = Nothing
    oMov.CommitTrans
    lbTrans = False
    GrabaRendicionCh = 0
    
    Exit Function
GrabaRendicionChErr:
    If lbTrans Then
        lbTrans = False
        oMov.RollbackTrans
    End If
    Call RaiseError(MyUnhandledError, "NARendir:GrabaRendicionCh Method")
End Function
Public Function GrabaAutorizacionCh(ByVal psMovNroAut As String, ByVal psMovNroRend As String, _
                                    ByVal psFormatoFecha As String, _
                                    ByVal psOpeCod As String, ByVal psGlosa As String, _
                                    ByVal pnMontoAutorizado As Currency, ByVal pnSaldo As Currency, _
                                    ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal pnNroProc As Integer, _
                                    Optional ByRef pnMovNroAut As Long = 0) As Integer
                                    'Agregado por ELRO el Parametro pnMovNroAut el 20120618, según OYP-RFC047-2012


 On Error GoTo GrabaAutorizacionChErr
 Dim rs As ADODB.Recordset
 Dim lnNuevoProcCH As Integer
 Dim lbTrans As Boolean
 Dim lnMovNro As Long
    Dim oMov As DMov
    Dim sql As String
    
    Dim lnSaldoArendir As Currency
    Set oMov = New DMov
    
    GrabaAutorizacionCh = 1
    oMov.BeginTrans
    lbTrans = True
    oMov.InsertaMov psMovNroAut, psOpeCod, psGlosa, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(psMovNroAut)
    oMov.InsertaMovCont lnMovNro, pnMontoAutorizado, "0", "0"
    oMov.InsertaMovCajaChica lnMovNro, psAreaCod, psAgeCod, pnNroProc, gCHTipoProcHabilitacion
    oMov.InsertaMovRef lnMovNro, oMov.GetnMovNro(psMovNroRend)
    oMov.CommitTrans
    lbTrans = False
    '***Agregado por ELRO el 20120618, según OYP-RFC047
    pnMovNroAut = lnMovNro
    '***Fin Agregado por ELRO**************************
    GrabaAutorizacionCh = 0
    Exit Function
GrabaAutorizacionChErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Call RaiseError(MyUnhandledError, "NARendir:GrabaAutorizacionCh Method")
End Function
Private Sub ActualizaMovCajaChica(ByRef oMov As DMov, ByVal pnMovNroRend As Long, ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnNroProc As Integer, _
                                  ByVal pdFecsis As Date, ByVal psCodUser As String, ByVal psAgencia As String)
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lsMovNroNew As String
Dim lnMovNronew As Long
Dim lsMovNroActualizaSaldo As String
Set rs = New ADODB.Recordset

'sql = "SELECT   M.NMOVNRO"'***Comentado por ELRO según OYP-RFC104-2012
 sql = "SELECT   M.NMOVNRO, M.CMOVNRO" _
    & " FROM    MOV M     " _
    & "         JOIN MOVCAJACHICA MCH ON MCH.NMOVNRO=M.NMOVNRO JOIN MOVCTA MC ON MC.NMOVNRO=M.NMOVNRO   " _
    & "  WHERE  M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
    & "         AND MCH.CAREACOD='" & psAreaCh & "' AND MC.NMOVIMPORTE>0 and m.nMovEstado IN (" & gMovEstContabNoContable & ") " _
    & "         AND CAGECOD='" & psAgeCh & "' AND nProcNro =" & pnNroProc & " " _
    & "         AND (SUBSTRING(M.COPECOD,1,5)>= '" & Mid(gCHArendirCtaAtencMN, 1, 5) & "' OR " _
    & "             SUBSTRING(M.COPECOD,1,5)>= '" & Mid(gCHArendirCtaAtencME, 1, 5) & "')  GROUP BY M.NMOVNRO,M.CMOVNRO "
'& "              SUBSTRING(M.COPECOD,1,5)>= '" & Mid(gCHArendirCtaAtencME, 1, 5) & "')  GROUP BY M.NMOVNRO "'***Comentado por ELRO según OYP-RFC104-2012

Set rs = oMov.CargaRecordSet(sql)
Do While Not rs.EOF
    '***Modificado por ELRO el 20120829, según OYP-RFC104-2012
    'lsMovNroNew = oMov.GeneraMovNro(pdFecsis, psAgencia, psCodUser)
    lsMovNroNew = oMov.GeneraMovNro(pdFecsis, Mid(rs!cMovNro, 18, 2), Right(rs!cMovNro, 4))
    '***Fin Modificado por ELRO el 20120829*******************
    'Debug.Print lsMovNroNew, rs!cMovNro
    'oMov.ActualizaMovimiento lsMovNroNew, rs!nMovNro, gMovEstContabMovContable, True, True, True
    oMov.ActualizaMov rs!nMovNro, , gMovEstContabMovContable, , lsMovNroNew, False
    lsMovNroActualizaSaldo = lsMovNroActualizaSaldo & "'" & lsMovNroNew & "',"

    
    'lnMovNronew = oMov.GetnMovNro(lsMovNroNew)
    'oMov.InsertaMovCajaChicaRend pnMovNroRend, rs!nMovNro, lnMovNronew
    rs.MoveNext
Loop
RSClose rs

If lsMovNroActualizaSaldo <> "" Then
    lsMovNroActualizaSaldo = Left(lsMovNroActualizaSaldo, Len(lsMovNroActualizaSaldo) - 1)
    oMov.ActualizaSaldoMovimiento lsMovNroActualizaSaldo, "+"
End If

End Sub
Public Function GetNuevaCajaChica(ByVal psAreaCh As String, ByVal psAgeCh As String) As Integer
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New Recordset

If oCon.AbreConexion = False Then Exit Function
 
 
sql = "SELECT ISNULL(MAX(nProcNro),0) + 1 AS MaxCajaChica " _
    & " From CAJACHICA " _
    & " WHERE cAreaCod = '" & psAreaCh & "' AND cAgeCod='" & psAgeCh & "' "
GetNuevaCajaChica = 1
Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetNuevaCajaChica = rs!MaxCajaChica
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetCHRendidas(ByVal psAreaCh As String, ByVal psAgeCh As String) As ADODB.Recordset

Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New Recordset

If oCon.AbreConexion = False Then Exit Function
 
 
sql = "SELECT   MCH.nProcNro " _
    & "FROM     MOV M JOIN MOVCAJACHICA MCH ON M.nMOVNRO=MCH.nMOVNRO " _
    & " WHERE   MCH.cProcTpo IN ('" & gCHTipoProcRendicion & "')         " _
    & "         And MCH.cAreaCod='" & psAreaCh & "' and MCH.cAgeCod= '" & psAgeCh & "' " _
    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
    & " group by  MCH.CAREACOD,MCH.CAGECOD, nProcNro "
    
Set rs = oCon.CargaRecordSet(sql)
Set GetCHRendidas = rs
oCon.CierraConexion
Set oCon = Nothing

End Function
Public Function GetCHRendSinAutorizacion(ByVal psAreaCh As String, ByVal psAgeCh As String) As ADODB.Recordset

Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New Recordset

If oCon.AbreConexion = False Then Exit Function
 
 
sql = "SELECT   MCH.nProcNro " _
    & "FROM     MOV M JOIN MOVCAJACHICA MCH ON M.NMOVNRO=MCH.NMOVNRO " _
    & " WHERE   MCH.cProcTpo IN ('" & gCHTipoProcRendicion & "')         " _
    & "         And MCH.cAreaCod='" & psAreaCh & "' and MCH.cAgeCod= '" & psAgeCh & "' " _
    & "         AND M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
    & "         AND NOT EXISTS ( SELECT M1.CMOVNRO " _
    & "                          FROM   MOV M1 JOIN MOVREF MR ON MR.NMOVNRO = M1.NMOVNRO" _
    & "                          WHERE  MR.NMOVNROREF=M.NMOVNRO AND M1.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "')) " _
    & " Group BY MCH.CAREACOD,MCH.CAGECOD, nProcNro "
    
Set rs = oCon.CargaRecordSet(sql)
Set GetCHRendSinAutorizacion = rs
oCon.CierraConexion
Set oCon = Nothing

End Function
Public Function GetMovCHProceso(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnNroProc As Integer, ByVal pnTipoProc As CHTipoProc) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function

sql = "SELECT  MCH.CAREACOD,MCH.CAGECOD, nProcNro, M.cMovNro, MCH.NMovNro " _
    & " FROM   MOV M JOIN MOVCAJACHICA MCH ON M.NMOVNRO=MCH.NMOVNRO  " _
    & " WHERE  MCH.cProcTpo IN ('" & pnTipoProc & "') " _
    & "        And MCH.cAreaCod='" & psAreaCh & "' and MCH.cAgeCod= '" & psAgeCh & "' " _
    & "        And MCH.nProcNro=" & pnNroProc & " " _
    & "        AND M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') "
GetMovCHProceso = ""
Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetMovCHProceso = rs!cMovNro
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing
End Function



Public Function GetCHAutorizadas(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnNroProc As Integer) As ADODB.Recordset

Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New Recordset

If oCon.AbreConexion = False Then Exit Function
 
sql = " SELECT  0 as Ok , CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) AS FechaAut,  " _
    & "         RIGHT(M.CMOVNRO,4) AS USUARIOAUT, P.CPERSNOMBRE,MC.NMOVMONTO , MC.NMOVMONTO, m.NMovNro, CH.cPersCod , ISNULL(RENDCONT.nProcNro,MCH.nProcNro) AS nProcNro  " _
    & " FROM    CAJACHICA CH " _
    & "         JOIN MOVCAJACHICA MCH ON MCH.CAREACOD = CH.CAREACOD AND MCH.CAGECOD=CH.CAGECOD AND CH.nProcNro = MCH.nProcNro " _
    & "         JOIN MOV M ON M.NMOVNRO = MCH.NMOVNRO " _
    & "         JOIN MOVCONT MC ON MC.NMOVNRO = M.NMOVNRO " _
    & "         JOIN PERSONA P ON P.CPERSCOD = CH.CPERSCOD " _
    & "         LEFT JOIN ( SELECT  nProcNro , MR.NMOVNRO " _
    & "                     FROM    MOVREF MR " _
    & "                             JOIN MOVCAJACHICA MC ON MC.NMOVNRO=MR.NMOVNROREF " _
    & "                             JOIN MOV M1 ON M1.NMOVNRO = MR.NMOVNROREF " _
    & "                      Where  M1.nMovFlag =" & gMovFlagVigente & " And MC.cProcTpo =" & gCHTipoProcRendicion & " ) AS RENDCONT  " _
    & "         ON RENDCONT.NMOVNRO = M.NMOVNRO " _
    & " WHERE   CH.CAREACOD='" & psAreaCh & "' AND CH.CAGECOD= '" & psAgeCh & "' " _
    & "         AND MCH.cProcTpo IN ('" & gCHTipoProcHabilitacion & "') AND " _
    & "         M.nMovFlag =" & gMovFlagVigente & "  " _
    & "         AND NOT EXISTS  (SELECT  M1.CMOVNRO " _
    & "                          FROM    MOV M1 JOIN MOVREF MR ON MR.NMOVNRO = M1.NMOVNRO " _
    & "                          WHERE   M1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
    & "                                  AND MR.NMOVNROREF=M.NMOVNRO ) " _
    & " ORDER BY  M.cMovNro "

'AND CH.nProcNro=" & pnNroProc & "
Set rs = oCon.CargaRecordSet(sql)
Set GetCHAutorizadas = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetCHDesembolsadas(ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                    ByVal pnNroProc As Integer, ByVal psCtaFondoFijo As String, _
                                    ByVal pbEfectivo As Boolean, Optional psTpoDoc As String = "") As ADODB.Recordset

Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lsFiltroDoc As String

Set oCon = New DConecta
Set rs = New Recordset
If pbEfectivo Then
    lsFiltroDoc = " AND MD.CDOCNRO Is Null"
Else
    lsFiltroDoc = " AND NOT MD.CDOCNRO Is Null"
End If

If oCon.AbreConexion = False Then Exit Function
 
 '***Modificado por ELRO el 20120602, según OYP-RFC047-2012
 'sql = "SELECT  0 as Ok, ISNULL(D.CDOCABREV,'EFEC') AS DocAbrev, ISNULL(MD.CDOCNRO,'EFECTIVO') as NroDoc , " _
 '   & "         CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHADESEMB, " _
 '   & "         M.CMOVDESC," & IIf(psTpoDoc <> "", "-1*", "") & "SUM(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)) AS MONTOAUTORIZADO, " _
 '   & "         " & IIf(psTpoDoc <> "", "-1*", "") & "SUM(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)) AS MONTOAUTORIZADO, CH.nSaldo,   " _
 '   & "         M.NMOVNRO AS NMovNroDesemb " _
 '   & " FROM    MOVCAJACHICA MCH " _
 '   & "         JOIN CAJACHICA CH ON CH.cAreaCod= MCH.cAreaCod AND CH.cAgeCod = MCH.cAgeCod  AND CH.nProcNro= MCH.nProcNro " _
 '   & "         JOIN MOV M ON M.nMOVNRO = MCH.nMOVNRO " _
 '   & "         JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO " _
 '   & "         LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM = ME.NMOVITEM " _
 '   & "         LEFT JOIN MOVDOC MD ON MD.NMOVNRO=M.NMOVNRO AND MD.nDocTpo " & IIf(psTpoDoc = "", "NOT IN (" & TpoDocVoucherEgreso, " IN (" & 80) & ")  " _
 '   & "         LEFT JOIN DOCUMENTO D ON D.nDocTpo = MD.nDocTpo " _
 '   & " WHERE   MCH.CAREACOD='" & psAreaCh & "' AND MCH.CAGECOD= '" & psAgeCh & "' AND MCH.nProcNro=" & pnNroProc & " " _
 '   & "         AND MCH.cProcTpo IN ('" & gCHTipoProcDesembolso & "')  AND MC.CCTACONTCOD LIKE '" & psCtaFondoFijo & "%' " & lsFiltroDoc _
 '   & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
 '  & "         AND NOT EXISTS (SELECT  M1.NMOVNRO " _
 '   & "                         FROM    MOV M1 JOIN MOVCAJACHICA MCH1 ON MCH1.NMOVNRO =M1.NMOVNRO  " _
 '   & "                         WHERE   MCH1.cAreaCod= MCH.cAreaCod AND MCH1.cAgeCod = MCH.cAgeCod  AND MCH1.nProcNro= MCH.nProcNro " _
 '   & "                                 AND M1.NMOVFLAG =" & gMovFlagVigente & " AND  MCH1.cProcTpo IN ('" & gCHTipoProcRendicion & "')) " _
 '  & " GROUP BY D.CDOCABREV, MD.CDOCNRO, M.NMOVNRO , M.CMOVNRO  , CH.nSaldo, CH.nMontoDesem, M.CMOVDESC "
 '***Modificado por ELRO***********************************
    sql = "exec stp_sel_DevolverCajasChicasDesembolsadas '" & psAreaCh & "', '" & psAgeCh & "', " & pnNroProc & ", '" & psCtaFondoFijo & "', " & IIf(pbEfectivo, 1, 0) & ", '" & psTpoDoc & "'"
Set rs = oCon.CargaRecordSet(sql)
Set GetCHDesembolsadas = rs
oCon.CierraConexion
Set oCon = Nothing
End Function

 '***Modificado por ELRO el 20120602, según OYP-RFC047-2012
'Public Function GrabaDesembolsoCH(ByVal psMovNro As String, ByVal pnMovNroAut As Long,
'                                        ByVal psPersCod As String, ByVal psFormatoFecha As String, ByVal rsBilletaje As ADODB.Recordset,
'                                        ByVal psOpeCod As String, ByVal psMovDesc As String, ByVal pnImporte As Currency,
'                                        ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnNroProc As Integer,
'                                        ByVal psTpoDoc As String, ByVal psNroDoc As String, ByVal psFechaDoc As String,
'                                        ByVal psNroVoucher As String, ByVal rsAut As ADODB.Recordset,
'                                        ByVal psCtaDiferencia As String, ByVal pnMontoDiferencia As Currency,
'                                        ByVal psCtaDebe As String, ByVal psCtaHaber As String, ByVal pnSaldoCh As Currency, ByVal psCuentaAho As String, Optional gbBitCentral = True, Optional pbGrabaOpeNegocio As Boolean = True,
'                                        Optional ByVal psCtaDebeITF As String, Optional ByVal psCtaHaberITF As String, Optional pnValorITF As Double) As Integer
'***Modificado por ELRO***********************************
Public Function GrabaDesembolsoCH(ByVal psMovNro As String, ByVal pnMovNroAut As Long, _
                             ByVal psPersCod As String, ByVal psFormatoFecha As String, _
                             ByVal rsBilletaje As ADODB.Recordset, _
                             ByVal psOpeCod As String, ByVal psMovDesc As String, _
                             ByVal pnImporte As Currency, _
                             ByVal psAreaCh As String, ByVal psAgeCh As String, _
                             ByVal pnNroProc As Integer, ByVal psTpoDoc As String, _
                             ByVal psNroDoc As String, ByVal psFechaDoc As String, _
                             ByVal psNroVoucher As String, _
                             ByVal rsAut As ADODB.Recordset, _
                             ByVal psCtaDiferencia As String, _
                             ByVal pnMontoDiferencia As Currency, _
                             ByVal psCtaDebe As String, ByVal psCtaHaber As String, _
                             ByVal pnSaldoCh As Currency, _
                             ByVal psCuentaAho As String, _
                             Optional gbBitCentral = True, _
                             Optional pbGrabaOpeNegocio As Boolean = True, _
                             Optional ByVal psCtaDebeITF As String, _
                             Optional ByVal psCtaHaberITF As String, _
                             Optional pnValorITF As Double, _
                             Optional ByRef pnMovNroApro As Long = 0, _
                             Optional ByVal pnTotalEgresos As Currency = 0) As Integer
'***Modificado por ELRO***********************************
'***Agregado por ELRO el parametro pnSaldoActual el 20120828, según OYP-RFC104-2012
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim oContF As NContFunciones
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov
On Error GoTo GrabaDesembolsoCHErr

Set oContF = New NContFunciones

GrabaDesembolsoCH = 1
'***Modificado por ELRO el 20120601, según OYP-RFC047-2012
pnMovNroApro = 0
'***Fin Modificado por ELRO*******************************
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1:
lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaDebe, psAreaCh + psAgeCh, False)
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe + lsSubCta, pnImporte
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, psAgeCh, psAreaCh

'guardamos el billetaje si este posee
If Not rsBilletaje Is Nothing Then
    If rsBilletaje.State = adStateOpen Then
        If Not rsBilletaje.EOF And Not rsBilletaje.BOF Then
            Do While Not rsBilletaje.EOF
                If rsBilletaje!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, rsBilletaje!Monto * -1
                    lnMovOrden = lnMovOrden + 1
                    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oDMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsBilletaje!cEfectivoCod
                End If
                rsBilletaje.MoveNext
            Loop
        End If
        rsBilletaje.Close: Set rsBilletaje = Nothing
    Else
        lnMovItem = lnMovItem + 1
        oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnImporte * -1
    End If
Else
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnImporte * -1
End If
If pnMontoDiferencia <> 0 Then
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDiferencia, pnMontoDiferencia * -1
End If
If Not rsAut Is Nothing Then
    Do While Not rsAut.EOF
        If rsAut!Ok = "1" Then
            oDMov.InsertaMovRef lnMovNro, rsAut!cMovNroAut
            If rsAut!nProcNro <> pnNroProc Then
                oDMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, rsAut!nProcNro, gCHTipoProcCancelada
            End If
        End If
        rsAut.MoveNext
    Loop
End If
oDMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, pnNroProc, gCHTipoProcDesembolso
If psTpoDoc <> "" And psNroDoc <> "" Then
    oDMov.InsertaMovDoc lnMovNro, psTpoDoc, psNroDoc, Format(psFechaDoc, gsFormatoFecha)
End If
If psNroVoucher <> "" Then
'    If psTpoDoc = 80 Then
'    oDMov.InsertaMovDoc lnMovNro, 80, psNroVoucher, Format(psFechaDoc, gsFormatoFecha)
'    Else
    If psTpoDoc = gnDocCuentaPendiente Then
        oDMov.InsertaMovDoc lnMovNro, gnDocCuentaPendiente, psNroVoucher, Format(psFechaDoc, gsFormatoFecha)
    Else
        oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, Format(psFechaDoc, gsFormatoFecha)
    End If
   ' End If
End If
If nVal(psTpoDoc) = TpoDocNotaAbono Or nVal(psTpoDoc) = TpoDocNotaCargo Then
    oDMov.InsertaRegDocCuenta psTpoDoc, psNroDoc, psCuentaAho
    If pbGrabaOpeNegocio Then
        Dim oCapta As NCapMovimientos
        Dim lnSaldo As Double
        Set oCapta = New NCapMovimientos
        If gbBitCentral Then
           lnSaldo = oCapta.CapAbonoCuentaAho(psCuentaAho, pnImporte, gCGPagProvAbonoCent, psMovNro, psMovDesc, TpoDocNotaAbono, psNroDoc, , , , , , , , , , True, False, , , oDMov.GetConexion, , True, CDbl(pnValorITF))
        Else
            Dim oCon As New DConecta
            Dim oDis As New NRHProcesosCierre
            
            If oCon.AbreConexion(Left(psCuentaAho, 2)) Then
                oCon.BeginTrans
                lnSaldo = oDis.Abono(psCuentaAho, pnImporte, gCGPagProvAbonoDist, gCGPagProvAbonoDist, "112" & Left(psCuentaAho, 2), Right(psMovNro, 4), psNroDoc, "PAGO : " & psMovDesc, oCon, GetFechaMov(psMovNro, True))
                oCon.CommitTrans
            End If
            oCon.CierraConexion
            Set oCon = Nothing
            Set oDis = Nothing
        End If
    End If
End If

If pnValorITF <> 0 Then
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebeITF, pnValorITF
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaberITF, pnValorITF * -1
End If

'***Modificado por ELRO el 20120828, según OYP-RFC104-2012
'pnSaldoCh = pnSaldoCh + pnImporte
If pnTotalEgresos > 0 Then
   pnSaldoCh = (pnSaldoCh + pnImporte) - pnTotalEgresos
Else
    pnSaldoCh = pnSaldoCh + pnImporte
End If
'***Fin Modificado por ELRO el 20120828*******************

'***Modificado por ELRO el 20120601, según OYP-RFC047-2012
'If psTpoDoc = gnDocCuentaPendiente Then
If IIf(psTpoDoc = "", 0, psTpoDoc) = gnDocCuentaPendiente Then
'***Fin Modificado por ELRO*******************************
    oDMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProc, , , pnImporte
Else
    oDMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProc, , , , , pnSaldoCh
End If


oDMov.InsertaMovGasto lnMovNro, psPersCod, "0"
oDMov.CommitTrans

GrabaDesembolsoCH = 0
'***Modificado por ELRO el 20120601, según OYP-RFC047-2012
pnMovNroApro = lnMovNro
'***Fin Agregado por ELRO*******************************
lbTrans = False
Set oDMov = Nothing
Set oContF = Nothing
Exit Function
GrabaDesembolsoCHErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaDesembolsoCH", Err.Description & " Error Graba Rendicion a rendir Efectivo "
End Function
Public Function GrabaExtornoDesembolso(ByVal pdFecsis As Date, ByVal pdFechaDesemb As Date, ByVal psFormatoFecha As String, _
                                     ByVal pnMovNroDesemb As Long, ByVal psMovNroExt As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                     ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                     ByVal pnNroProc As Integer, ByVal pnImporte As Currency, _
                                     Optional psOpeCodAut As String = "") As Integer
                                     '***Agregado por ELRO el psOpeCodAut el 20120619, según OYP-RFC047-2012
Dim oDMov As DMov
Dim lnSaldo As Currency
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov

On Error GoTo GrabaExtornoDesembolsoErr
GrabaExtornoDesembolso = 1

oDMov.BeginTrans
lbTrans = True

If pdFecsis <> pdFechaDesemb Then
    oDMov.ExtornaMovimiento psMovNroExt, pnMovNroDesemb
Else
    oDMov.InsertaMov psMovNroExt, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagDeExtorno
    lnMovNro = oDMov.GetnMovNro(psMovNroExt)
    oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
    oDMov.ActualizaMov pnMovNroDesemb, , , gMovFlagEliminado
    oDMov.InsertaMovRef pnMovNroDesemb, lnMovNro
    '***Agregado por ELRO el 20120619, según OYP-RFC047-2012
    If pnNroProc > 1 Then
        Dim lsMovNroExtAut As String
        Dim lnMovNroExtAut As Long
        Dim lnMovNroAut As Long
        lsMovNroExtAut = oDMov.GeneraMovNro(pdFecsis, Left(Right(psMovNroExt, 8), 2), Right(psMovNroExt, 4))
        oDMov.InsertaMov lsMovNroExtAut, "401311", "Extorno Autorizacion Reembolso", gMovEstContabNoContable, gMovFlagDeExtorno
        lnMovNroExtAut = oDMov.GetnMovNro(lsMovNroExtAut)
        lnMovNroAut = oDMov.obtenerMovRef(pnMovNroDesemb)
        oDMov.InsertaMovCont lnMovNroExtAut, pnImporte, "0", "0"
        oDMov.ActualizaMov lnMovNroAut, , , gMovFlagEliminado
        oDMov.InsertaMovRef lnMovNroExtAut, lnMovNroAut
    End If
    '***Fin Agregado por ELRO*******************************
    
End If
lnSaldo = GetDatosCajaChica(psAreaCh, psAgeCh)
lnSaldo = lnSaldo - pnImporte
oDMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProc, , , , , lnSaldo
GrabaExtornoDesembolso = 0
oDMov.CommitTrans
lbTrans = False
Set oDMov = Nothing
Exit Function
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
GrabaExtornoDesembolsoErr:
    Err.Raise vbObjectError + 100, "GrabaExtornoDesembolso", "Error Grabar Extorno de Desembolso de Caja chica"
End Function
'***Modificado por ELRO el 20120529 según OYP-RFC047-2012
'Public Function GrabaHabNuevaCH(ByVal psFormatoFecha As String,
'                                ByVal psMovNroHab As String, ByVal psOpeCod As String, ByVal psMovDesc As String,
'                                ByVal psAreaCh As String, ByVal psAgeCh As String,
'                                ByVal pnNroProc As Integer, ByVal pnImporte As Currency,
'                                ByVal psPersCod As String, ByVal pnMontoAsig As Currency, ByVal pnMontoTope As Currency) As Integer
Public Function GrabaHabNuevaCH(ByVal psFormatoFecha As String, _
                                ByVal psMovNroHab As String, _
                                ByVal psOpeCod As String, _
                                ByVal psMovDesc As String, _
                                ByVal psAreaCh As String, _
                                ByVal psAgeCh As String, _
                                ByVal pnNroProc As Integer, _
                                ByVal pnImporte As Currency, _
                                ByVal psPersCod As String, _
                                ByVal pnMontoAsig As Currency, _
                                ByVal pnMontoTope As Currency) As Long
'***Fin Modificado por ELRO******************************
Dim oDMov As DMov
Dim lnSaldo As Currency
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov

On Error GoTo GrabaHabNuevaCHErr
'***Modificado por ELRO el 20120529 según OYP-RFC047-2012
'GrabaHabNuevaCH = 1
GrabaHabNuevaCH = 0
'***Fin Modificado por ELRO******************************

oDMov.BeginTrans
oDMov.InsertaMov psMovNroHab, psOpeCod, psMovDesc, gMovEstContabNoContable
lnMovNro = oDMov.GetnMovNro(psMovNroHab)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
oDMov.InsertaCajaChica psAreaCh, psAgeCh, pnNroProc, psPersCod, pnMontoAsig, pnMontoTope, 0, 0, 0
oDMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, pnNroProc, gCHTipoProcHabilitacion
oDMov.CommitTrans
lbTrans = False
'***Modificado por ELRO el 20120529 según OYP-RFC047-2012
'GrabaHabNuevaCH = 0
GrabaHabNuevaCH = lnMovNro
'***Fin Modificado por ELRO******************************
Set oDMov = Nothing
Exit Function
GrabaHabNuevaCHErr:
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    Err.Raise vbObjectError + 100, "GrabaHabNuevaCH", "Error Grabar nueva Caja chica"
End Function

Public Function ActualizaCH(ByVal psFormatoFecha As String, _
                           ByVal psAreaCh As String, ByVal psAgeCh As String, _
                           ByVal pnNroProc As Integer, _
                           ByVal psPersCod As String, ByVal pnMontoAsig As Currency, ByVal pnMontoTope As Currency) As Integer
Dim oDMov As DMov
Dim lnSaldo As Currency
Dim lbTrans As Boolean
Set oDMov = New DMov

On Error GoTo ActualizaCHErr
ActualizaCH = 1

oDMov.BeginTrans
lbTrans = True
oDMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProc, psPersCod, pnMontoAsig, pnMontoTope
oDMov.CommitTrans
lbTrans = False
ActualizaCH = 0

Set oDMov = Nothing
Exit Function
ActualizaCHErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "ActualizaCH", "Error Actualizacion de Caja chica"
End Function

Public Function GetMontoNoDesembolsado(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnNroProc As Integer, Optional pnTpoProc As CHTipoProc = gCHTipoProcHabilitacion) As Currency
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New Recordset

If oCon.AbreConexion = False Then Exit Function
'***Modificado por ELRO el 20120114, según SATI INC1301120002
'sql = " SELECT  ISNULL(Sum(MC.NMOVMONTO),0) as Total " _
'    & " FROM    CAJACHICA CH " _
'    & "         JOIN MOVCAJACHICA MCH ON MCH.CAREACOD = CH.CAREACOD AND MCH.CAGECOD=CH.CAGECOD AND CH.nProcNro = MCH.nProcNro " _
'    & "         JOIN MOV M ON M.NMOVNRO = MCH.NMOVNRO " _
'    & "         JOIN MOVCONT MC ON MC.NMOVNRO = M.NMOVNRO " _
'    & "         JOIN PERSONA P ON P.CPERSCOD = CH.CPERSCOD " _
'    & " WHERE   CH.CAREACOD='" & psAreaCh & "' AND CH.CAGECOD= '" & psAgeCh & "' " _
'    & "         AND MCH.cProcTpo IN ('" & pnTpoProc & "') AND " _
'    & "         M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
'    & "         AND NOT EXISTS  (SELECT  M1.CMOVNRO " _
'    & "                          FROM    MOV M1 JOIN MOVREF MR ON MR.NMOVNRO = M1.NMOVNRO " _
'    & "                          WHERE   M1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
'    & "                                  AND MR.NMOVNROREF=M.NMOVNRO ) "
'***Fin Modificado por ELRO el 20120114**********************
sql = "exec stp_sel_DevolverMontoNoDesembolsados '" & psAreaCh & "', '" & psAgeCh & "', '" & pnTpoProc & "'"
Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetMontoNoDesembolsado = rs!Total
End If
RSClose rs
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function GetMontoDesembolsadoNext(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnNroProc As Integer, Optional pnTpoProc As CHTipoProc = gCHTipoProcHabilitacion) As Currency
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta
Set rs = New Recordset

If oCon.AbreConexion = False Then Exit Function
sql = ""
'***Modificado por ELRO el 20120114, según SATI INC1301120002
'sql = sql & "Select ISNULL(SUM(mc.nMovImporte),0) Total "
'sql = sql & "from mov m join movcajachica mch on mch.nmovnro = m.nmovnro "
'sql = sql & "           join MovCta mc ON mc.nMovNro = m.nMovNro and mc.nMovImporte > 0 "
'sql = sql & "where mch.cAreaCod = '" & psAreaCh & "' and mch.cagecod = '" & psAgeCh & "' and m.nMovNro > (SELECT a.nMovNro FROM MovCajaChica  a inner join mov b on a.nmovnro = b.nmovnro WHERE cAreaCod = '" & psAreaCh & "' and cAgeCod = '" & psAgeCh & "' and cProcTpo = '" & gCHTipoProcRendicion & "' and nProcNro = " & pnNroProc & " And b.nmovflag <> " & MovFlag.gMovFlagEliminado & ")  And m.nmovflag <> " & MovFlag.gMovFlagEliminado
'sql = sql & "      and cProcTpo = " & gCHTipoProcDesembolso
'***Fin Modificado por ELRO el 20120114**********************
sql = "exec stp_sel_DevolverProximoMontoDesembolsado '" & psAreaCh & "','" & psAgeCh & "'," & pnNroProc & ""
Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetMontoDesembolsadoNext = rs!Total
End If
RSClose rs
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function GetCHArqueoRendidos(ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                    ByVal pnNroProc As Integer) As ADODB.Recordset

Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lsFiltroDoc As String

Set oCon = New DConecta
Set rs = New Recordset


If oCon.AbreConexion = False Then Exit Function
 
 sql = "SELECT  DISTINCT CH.nProcNro  " _
    & " FROM    CAJACHICA CH " _
    & "         JOIN MOVCAJACHICA MCH   ON MCH.CAREACOD = CH.CAREACOD AND MCH.CAGECOD=CH.CAGECOD AND CH.nProcNro = MCH.nProcNro " _
    & "         JOIN MOV M      ON M.NMOVNRO = MCH.NMOVNRO " _
    & "         JOIN MOVCONT MC     ON MC.NMOVNRO = M.NMOVNRO " _
    & " WHERE   CH.CAREACOD='" & psAreaCh & "' AND CH.CAGECOD= '" & psAgeCh & "' AND CH.nProcNro<>" & pnNroProc & " " _
    & "         AND MCH.cProcTpo IN ('" & gCHTipoProcHabilitacion & "') AND " _
    & "         M.nMovFlag =" & gMovFlagVigente & "" _
    & "         AND NOT EXISTS  ( SELECT M1.CMOVNRO " _
    & "                           FROM   MOV M1 " _
    & "                                  JOIN MOVCAJACHICA MCH1 ON MCH1.CAREACOD = MCH.CAREACOD AND MCH1.CAGECOD=MCH.CAGECOD AND MCH1.nProcNro = MCH.nProcNro " _
    & "                           WHERE  M1.NMOVFLAG =" & gMovFlagVigente & " AND MCH1.cProcTpo in ('" & gCHTipoProcCancelada & "','" & gCHTipoProcArqueo & "'))" _

Set rs = oCon.CargaRecordSet(sql)
Set GetCHArqueoRendidos = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetOrdenesNoCobradas(ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                     ByVal psNroCajasChicas As String) As ADODB.Recordset

Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lsFiltroDoc As String

Set oCon = New DConecta
Set rs = New Recordset


If oCon.AbreConexion = False Then Exit Function

sql = "Select  MD.cDocNro , MCT.nMovMonto" _
    & " From    Mov M " _
    & "         JOIN MOVCAJACHICA MC ON MC.NMOVNRO =M.NMOVNRO " _
    & "         JOIN MOVCONT MCT ON MCT.NMOVNRO =M.NMOVNRO " _
    & "         JOIN MOVDOC MD ON MD.NMOVNRO=M.NMOVNRO AND MD.nDocTpo ='" & TpoDocOrdenPago & "' " _
    & " WHERE   MC.cProcTpo = '" & gCHTipoProcDesembolso & "' AND MC.CAREACOD='" & psAreaCh & "' " _
    & "         AND MC.CAGECOD='" & psAgeCh & "' AND M.NMOVFLAG =" & gMovFlagVigente & "  and MC.nProcNro in (" & psNroCajasChicas & ") " _
    & "         AND NOT EXISTS (SELECT M1.NMOVNRO " _
    & "                         FROM MOV M1 JOIN MOVREF MR ON MR.NMOVNRO=M1.NMOVNRO " _
    & "                         WHERE M1.NMOVFLAG=" & gMovFlagVigente & "  AND MR.NMOVNROREF=M.NMOVNRO) "

Set rs = oCon.CargaRecordSet(sql)
Set GetOrdenesNoCobradas = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetTotalDocArqueos(ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                  ByVal psNroCajasChicas As String) As Currency

Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lsFiltroDoc As String

Set oCon = New DConecta
Set rs = New Recordset


If oCon.AbreConexion = False Then Exit Function

sql = " SELECT SUM(TOTALDOC) AS TOTALDOC FROM ( " _
    & " SELECT  ISNULL(SUM(MC.NMOVIMPORTE),0) AS TOTALDOC " _
    & " FROM    MOV M JOIN MOVCAJACHICA MCH ON MCH.NMOVNRO = M.NMOVNRO " _
    & "         JOIN MOVCTA MC ON MC.NMOVNRO =M.NMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.NMOVNRO =MC.NMOVNRO AND ME.NMOVITEM = MC.NMOVITEM " _
    & "         JOIN MOVDOC MD ON MD.NMOVNRO = M.NMOVNRO " _
    & "         JOIN MOVGASTO MG ON MG.NMOVNRO = M.NMOVNRO " _
    & "         JOIN MOVREF MR ON MR.NMOVNRO = M.NMOVNRO " _
    & " WHERE   M.NMOVFLAG =0 AND MCH.CAREACOD='" & psAreaCh & "' AND MCH.CAGECOD='" & psAgeCh & "' " _
    & "         AND MC.NMOVIMPORTE>0  and MCH.nProcNro in (" & psNroCajasChicas & ")" _
    & "         AND MCH.cProcTpo=" & gCHTipoProcEgresoDirecto & " " _
    & " UNION " _
    & " SELECT  ISNULL(SUM(MC.NMOVIMPORTE),0) AS TOTALDOC " _
    & " FROM    MOV M JOIN MOVCAJACHICA MCH ON MCH.NMOVNRO = M.NMOVNRO " _
    & "         JOIN MOVCTA MC ON MC.NMOVNRO =M.NMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.NMOVNRO =MC.NMOVNRO AND ME.NMOVITEM = MC.NMOVITEM " _
    & "         JOIN MOVDOC MD ON MD.NMOVNRO = M.NMOVNRO " _
    & "         JOIN MOVGASTO MG ON MG.NMOVNRO = M.NMOVNRO " _
    & "         JOIN MOVREF MR ON MR.NMOVNRO = M.NMOVNRO " _
    & "         JOIN MOVARENDIRSUST MAR ON MR.NMOVNRO = nMovNroRend     " _
    & " WHERE   M.NMOVFLAG =0 AND MCH.CAREACOD='" & psAreaCh & "' AND MCH.CAGECOD='" & psAgeCh & "' " _
    & "         AND MC.NMOVIMPORTE>0  and MCH.nProcNro in (" & psNroCajasChicas & ") " _
    & "         AND MCH.cProcTpo=" & gCHTipoProcArendir & " AND MAR.NMOVNROREND IS NOT NULL ) AS TOTALDOC"

Set rs = oCon.CargaRecordSet(sql)
GetTotalDocArqueos = 0
If Not rs.EOF And Not rs.BOF Then
    GetTotalDocArqueos = rs!TotalDoc
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function GetTotalArendirArqueos(ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                  ByVal psNroCajasChicas As String) As Currency

Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lsFiltroDoc As String

Set oCon = New DConecta
Set rs = New Recordset


If oCon.AbreConexion = False Then Exit Function

sql = " SELECT  ISNULL(SUM(MC.NMOVIMPORTE),0) AS TOTALARENDIR  " _
    & " FROM    MOV M JOIN MOVCAJACHICA MCH ON MCH.NMOVNRO = M.NMOVNRO " _
    & "         JOIN MOVCTA MC ON MC.NMOVNRO =M.NMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.NMOVNRO =MC.NMOVNRO AND ME.NMOVITEM = MC.NMOVITEM " _
    & " WHERE   M.NMOVFLAG =0 AND MCH.CAREACOD='" & psAreaCh & "' AND MCH.CAGECOD='" & psAgeCh & "' " _
    & "         AND MC.NMOVIMPORTE>0  and MCH.nProcNro in (" & psNroCajasChicas & ") " _
    & "         AND MCH.cProcTpo IN ('" & gCHTipoProcArendir & "') AND M.COPECOD in ('" & gCHArendirCtaAtencMN & "','" & gCHArendirCtaAtencME & "') "
    
Set rs = oCon.CargaRecordSet(sql)
GetTotalArendirArqueos = 0
If Not rs.EOF And Not rs.BOF Then
    GetTotalArendirArqueos = rs!TOTALARENDIR
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function GetTotalArendirArqueosSustentados(ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                  ByVal psNroCajasChicas As String) As Currency

Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lsFiltroDoc As String

Set oCon = New DConecta
Set rs = New Recordset


If oCon.AbreConexion = False Then Exit Function

sql = " SELECT  ISNULL(SUM(MC.NMOVIMPORTE),0) AS TOTALARENDIR  " _
    & " FROM    MOV M JOIN MOVCAJACHICA MCH ON MCH.NMOVNRO = M.NMOVNRO " _
    & "         JOIN MOVCTA MC ON MC.NMOVNRO =M.NMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.NMOVNRO =MC.NMOVNRO AND ME.NMOVITEM = MC.NMOVITEM " _
    & " WHERE   M.NMOVFLAG =0 AND MCH.CAREACOD='" & psAreaCh & "' AND MCH.CAGECOD='" & psAgeCh & "' " _
    & "         AND MC.NMOVIMPORTE>0  and MCH.nProcNro in (" & psNroCajasChicas & ") " _
    & "         AND MCH.cProcTpo IN ('" & gCHTipoProcArendir & "') AND M.COPECOD in ('" & gCHArendirCtaSustMN & "','" & gCHArendirCtaSustME & "') "
    
Set rs = oCon.CargaRecordSet(sql)
GetTotalArendirArqueosSustentados = 0
If Not rs.EOF And Not rs.BOF Then
    GetTotalArendirArqueosSustentados = rs!TOTALARENDIR
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function GetTotalArendirArqueosNOSustentados(ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                  ByVal psNroCajasChicas As String) As Currency

Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lsFiltroDoc As String

Set oCon = New DConecta
Set rs = New Recordset


If oCon.AbreConexion = False Then Exit Function

sql = " SELECT  ISNULL(SUM(MC.NMOVIMPORTE),0) AS TOTALARENDIR  " _
    & " FROM    MOV M JOIN MOVCAJACHICA MCH ON MCH.NMOVNRO = M.NMOVNRO " _
    & "         JOIN MOVCTA MC ON MC.NMOVNRO =M.NMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.NMOVNRO =MC.NMOVNRO AND ME.NMOVITEM = MC.NMOVITEM " _
    & " WHERE   M.NMOVFLAG =0 AND MCH.CAREACOD='" & psAreaCh & "' AND MCH.CAGECOD='" & psAgeCh & "' " _
    & "         AND MC.NMOVIMPORTE>0  and MCH.nProcNro in (" & psNroCajasChicas & ") " _
    & "         AND MCH.cProcTpo IN ('" & gCHTipoProcArendir & "') AND M.COPECOD in ('" & gCHArendirCtaAtencMN & "','" & gCHArendirCtaAtencME & "') AND M.NMOVNRO NOT IN ( SELECT MR.NMOVNROREF FROM MOVREF MR INNER JOIN MOV M1 ON MR.NMOVNRO=M1.NMOVNRO " _
    & "          WHERE M1.COPECOD in ('" & gCHArendirCtaSustMN & "','" & gCHArendirCtaSustME & "') AND MR.NMOVNROREF=M.NMOVNRO)"
    
Set rs = oCon.CargaRecordSet(sql)
GetTotalArendirArqueosNOSustentados = 0
If Not rs.EOF And Not rs.BOF Then
    GetTotalArendirArqueosNOSustentados = rs!TOTALARENDIR
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing
End Function


Public Function GrabaArqueoCajaChica(ByVal psFormatoFecha As String, _
                                    ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                    ByVal psCtaDebe As String, ByVal psCtaHaber As String, _
                                    ByVal rsBillIngreso As ADODB.Recordset, ByVal rsMonIng As ADODB.Recordset, _
                                    ByVal rsTotales As ADODB.Recordset, ByVal rsOP As ADODB.Recordset, ByVal rsCh As ADODB.Recordset, _
                                    ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                    ByVal psAgeCod As String, ByVal psCodPers As String, _
                                    ByVal psHoraIni As String, ByVal psHoraFin As String, ByVal pnTotalOP As Currency) As Integer
            
Dim oDMov As DMov
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim pnTotal As Currency
Dim lnTotalEfec As Currency
Dim lnTotalDoc As Currency
Dim lnDocContab As Currency
Dim lnARendir As Currency
Dim lnTotalFondo As Currency
Dim lnFondoAsig As Currency
Dim lnDiferencia As Currency


Set oDMov = New DMov
On Error GoTo GrabaArqueoCajaChicaErr

lnMovItem = 0: lnMovOrden = 0
GrabaArqueoCajaChica = 1
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc
lnMovNro = oDMov.GetnMovNro(psMovNro)

If Not rsTotales Is Nothing Then
    If rsTotales.State = adStateOpen Then
        If Not rsTotales.EOF And Not rsTotales.BOF Then
            Do While Not rsTotales.EOF
                Select Case rsTotales.Bookmark
                    Case 1
                        lnTotalEfec = rsTotales!Importe
                    Case 2
                        lnTotalDoc = rsTotales!Importe
                    Case 3
                        lnDocContab = rsTotales!Importe
                    Case 4
                        lnARendir = rsTotales!Importe
                    Case 7
                        lnTotalFondo = rsTotales!Importe
                    Case 8
                        lnFondoAsig = rsTotales!Importe
                    Case 9
                        lnDiferencia = rsTotales!Importe
                        pnTotal = lnDiferencia
                End Select
                rsTotales.MoveNext
            Loop
        End If
        rsTotales.Close: Set rsTotales = Nothing
        oDMov.InsertaMovArqueo lnMovNro, psAgeCod, psCodPers, psHoraIni, psHoraFin, lnTotalEfec, _
                        pnTotalOP, lnTotalDoc, lnDocContab, lnARendir, lnDiferencia
    End If
End If
'guardamos el asiento de diferencia
If pnTotal <> 0 Then
    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, Abs(pnTotal)
    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, Abs(pnTotal) * -1
End If
'guardamos el efectivo de arqueo
If Not rsBillIngreso Is Nothing Then
    If rsBillIngreso.State = adStateOpen Then
        If Not rsBillIngreso.EOF And Not rsBillIngreso.BOF Then
            Do While Not rsBillIngreso.EOF
                If rsBillIngreso!Monto > 0 Then
                    oDMov.InsertaMovArqueoEfectivo lnMovNro, rsBillIngreso!cEfectivoCod, rsBillIngreso!Monto
                End If
                rsBillIngreso.MoveNext
            Loop
        End If
        'rsBillIngreso.Close: Set rsBillIngreso = Nothing
    End If
End If
If Not rsMonIng Is Nothing Then
    If rsMonIng.State = adStateOpen Then
        If Not rsMonIng.EOF And Not rsMonIng.BOF Then
            Do While Not rsMonIng.EOF
                If rsMonIng!Monto > 0 Then
                    oDMov.InsertaMovArqueoEfectivo lnMovNro, rsMonIng!cEfectivoCod, rsMonIng!Monto
                End If
                rsMonIng.MoveNext
            Loop
        End If
        'rsMonIng.Close: Set rsMonIng = Nothing
    End If
End If
'guardamos las ordenes de pago en arqueo
If Not rsOP Is Nothing Then
    If rsOP.State = adStateOpen Then
        If Not rsOP.EOF And Not rsOP.BOF Then
            Do While Not rsOP.EOF
                oDMov.InsertaMovArqueoOP lnMovNro, rsOP!Orden, rsOP!Importe
                rsOP.MoveNext
            Loop
        End If
        rsOP.Close: Set rsOP = Nothing
    End If
End If
'guardamos el movimiento de arqueo por cada  caja chica
If Not rsCh Is Nothing Then
    If rsCh.State = adStateOpen Then
        If Not rsCh.EOF And Not rsCh.BOF Then
            Do While Not rsCh.EOF
                oDMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, rsCh(0), gCHTipoProcArqueo
                rsCh.MoveNext
            Loop
        End If
        rsCh.Close: Set rsCh = Nothing
    End If
End If

oDMov.CommitTrans
lbTrans = False
GrabaArqueoCajaChica = 0
Set oDMov = Nothing
Exit Function
GrabaArqueoCajaChicaErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaArqueoCajaChica", Err.Description & " Error Devolucion de Efectivo de Cajero a Boveda"
End Function

Public Function GrabaRendicionChExt(ByVal psAgencia As String, ByVal psAreaCod As String, ByVal pnProcesoNro As Integer) As Integer
 On Error GoTo GrabaRendicionChExtErr
    Dim oMov As DMov
    Dim sql As String
    Set oMov = New DMov
    Dim lbTrans As Boolean
    
    GrabaRendicionChExt = 1
    oMov.BeginTrans
        lbTrans = True
        oMov.ExtonrnaRendCH psAgencia, psAreaCod, pnProcesoNro
        'ver***************************************
    oMov.CommitTrans
    lbTrans = False
    GrabaRendicionChExt = 0
    Exit Function
GrabaRendicionChExtErr:
    If lbTrans Then
        lbTrans = False
        oMov.RollbackTrans
    End If
    Call RaiseError(MyUnhandledError, "NARendir:GrabaRendicionCh Method")
End Function

'***Agregado por ELRO el 20120529, según OYP-RFC047-2012
Public Function verificarAreasAgenciasCajaChica(ByVal psAreaCod As String, ByVal psAgeCod As String) As Integer
    Dim sql As String
    Dim rsAreasAgencias As ADODB.Recordset
    Dim oDConecta As DConecta
    
    Set oDConecta = New DConecta
    Set rsAreasAgencias = New ADODB.Recordset
    
    If oDConecta.AbreConexion = False Then Exit Function
     
    sql = "exec stp_sel_VerificarAreasAgenciasCajaChica '" & psAreaCod & "', '" & psAgeCod & "'"
    
    
    Set rsAreasAgencias = oDConecta.CargaRecordSet(sql)
    
    If Not rsAreasAgencias.BOF And Not rsAreasAgencias.EOF Then
        verificarAreasAgenciasCajaChica = rsAreasAgencias.Fields(0)
    Else
        verificarAreasAgenciasCajaChica = -1
    End If
       
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsAreasAgencias = Nothing
End Function
Public Function verificarAprobacionCajaChica(ByVal psAreaCod As String, ByVal psAgeCod As String) As Integer
    Dim sql As String
    Dim rsAreasAgencias As ADODB.Recordset
    Dim oDConecta As DConecta
    
    Set oDConecta = New DConecta
    Set rsAreasAgencias = New ADODB.Recordset
    
    If oDConecta.AbreConexion = False Then Exit Function
     
    sql = "exec stp_sel_VerificarAprobacionCajaChica '" & psAreaCod & "', '" & psAgeCod & "'"
    
    
    Set rsAreasAgencias = oDConecta.CargaRecordSet(sql)
    
    If Not rsAreasAgencias.BOF And Not rsAreasAgencias.EOF Then
       verificarAprobacionCajaChica = rsAreasAgencias.Fields(0)
    Else
        verificarAprobacionCajaChica = 0
    End If
       
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsAreasAgencias = Nothing
End Function

Public Function devolverDatosCajaChicaSinAprobar(ByVal psAreaCh As String, ByVal psAgeCh As String) As ADODB.Recordset
Dim sql As String
Dim rsAreaAgencia As ADODB.Recordset
Dim oDConecta As DConecta

Set oDConecta = New DConecta
Set rsAreaAgencia = New Recordset

If oDConecta.AbreConexion = False Then Exit Function
 
sql = "exec stp_sel_DevolverCajaChicaSinAprobar '" & psAreaCh & "', '" & psAgeCh & "'"
Set rsAreaAgencia = oDConecta.CargaRecordSet(sql)
Set devolverDatosCajaChicaSinAprobar = rsAreaAgencia
oDConecta.CierraConexion
Set oDConecta = Nothing
Set rsAreaAgencia = Nothing

End Function

Public Function devolverDatosCajaChicaSinAprobar_2(ByVal pnMovNro As Long, ByVal psAreaCh As String, ByVal psAgeCh As String) As ADODB.Recordset
Dim sql As String
Dim rsAreaAgencia As ADODB.Recordset
Dim oDConecta As DConecta

Set oDConecta = New DConecta
Set rsAreaAgencia = New Recordset

If oDConecta.AbreConexion = False Then Exit Function
 
sql = "exec stp_sel_DevolverCajaChicaSinAprobar_2 '" & pnMovNro & "', '" & psAreaCh & "', '" & psAgeCh & "'"
Set rsAreaAgencia = oDConecta.CargaRecordSet(sql)
Set devolverDatosCajaChicaSinAprobar_2 = rsAreaAgencia
oDConecta.CierraConexion
Set oDConecta = Nothing
Set rsAreaAgencia = Nothing

End Function

Public Function devolverDatosCajaChicaMantenimiento(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnProcNro As Integer) As ADODB.Recordset
Dim sql As String
Dim rsAreaAgencia As ADODB.Recordset
Dim oDConecta As DConecta

Set oDConecta = New DConecta
Set rsAreaAgencia = New Recordset

If oDConecta.AbreConexion = False Then Exit Function
 
sql = "exec stp_sel_DevolverDatosCajaChicaMantenimiento '" & psAreaCh & "', '" & psAgeCh & "', " & pnProcNro & ""
Set rsAreaAgencia = oDConecta.CargaRecordSet(sql)
Set devolverDatosCajaChicaMantenimiento = rsAreaAgencia
oDConecta.CierraConexion
Set oDConecta = Nothing
Set rsAreaAgencia = Nothing

End Function

'MIOL 20130422, SEGUN RQ13152 ********************************************************
Public Function devolverDatosCajaChicaxCambiar(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnProcNro As Integer) As ADODB.Recordset
Dim sql As String
Dim rsAreaAgencia As ADODB.Recordset
Dim oDConecta As DConecta

Set oDConecta = New DConecta
Set rsAreaAgencia = New Recordset

If oDConecta.AbreConexion = False Then Exit Function
 
sql = "exec stp_sel_DevolverDatosCajaChicaxCambiarResponsable '" & psAreaCh & "', '" & psAgeCh & "', " & pnProcNro & ""
Set rsAreaAgencia = oDConecta.CargaRecordSet(sql)
Set devolverDatosCajaChicaxCambiar = rsAreaAgencia
oDConecta.CierraConexion
Set oDConecta = Nothing
Set rsAreaAgencia = Nothing

End Function

Public Function GrabaResponsableNuevoCH(ByVal psMovNroHab As String, _
                                        ByVal psOpeCod As String, _
                                        ByVal psMovDesc As String, _
                                        ByVal psAreaCh As String, _
                                        ByVal psAgeCh As String, _
                                        ByVal pnNroProc As Integer, _
                                        ByVal pscPersActual As String, _
                                        ByVal pscPersNueva As String, _
                                        ByVal pnSaldo As Currency, _
                                        ByVal pnEstado As Integer) As Long

Dim oDMov As DMov
Dim lnSaldo As Currency
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov

On Error GoTo GrabaResponsableNuevoCHErr

GrabaResponsableNuevoCH = 0

oDMov.BeginTrans
oDMov.InsertaMov psMovNroHab, psOpeCod, psMovDesc, gMovEstContabNoContable
lnMovNro = oDMov.GetnMovNro(psMovNroHab)
oDMov.InsertaMovCajaChicaCambResp lnMovNro, psAreaCh, psAgeCh, pnNroProc, pscPersActual, pscPersNueva, pnSaldo, psMovNroHab, pnEstado
If pnEstado = 1 Then
    oDMov.ActualizaMovCajaChicaCambResp pscPersNueva, psAreaCh, psAgeCh, pnNroProc
End If
oDMov.CommitTrans
lbTrans = False

GrabaResponsableNuevoCH = lnMovNro

Set oDMov = Nothing
Exit Function
GrabaResponsableNuevoCHErr:
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    Err.Raise vbObjectError + 100, "Error Grabar Cambio nuevo responsable Caja chica"
End Function
'END MIOL ****************************************************************************

'MIOL 20130422, SEGUN RQ13153 ********************************************************
Public Function devolverDatosRespCajaChica(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal psFecSist As String) As ADODB.Recordset
Dim sql As String
Dim rsAreaAgencia As ADODB.Recordset
Dim oDConecta As DConecta

Set oDConecta = New DConecta
Set rsAreaAgencia = New Recordset

If oDConecta.AbreConexion = False Then Exit Function
 
sql = "exec stp_Sel_RespDefinidoCambioCajaChicaUltimo '" & psAreaCh & "', '" & psAgeCh & "', '" & psFecSist & "'"
Set rsAreaAgencia = oDConecta.CargaRecordSet(sql)
Set devolverDatosRespCajaChica = rsAreaAgencia
oDConecta.CierraConexion
Set oDConecta = Nothing
Set rsAreaAgencia = Nothing

End Function
'END MIOL ****************************************************************************

Public Function GrabaEdicionAperturaCajaChica(ByVal pnMovNroCajaChica As Long, _
                                              ByVal pnMovMonto As Currency, _
                                              ByVal psPersCod As String, _
                                              ByVal psAreCod As String, _
                                              ByVal psAgeCod As String, _
                                              ByVal pnTopeEgreso As Currency, _
                                              ByVal pnProcNro As Integer) As Long
Dim oDMov As DMov
Dim lnSaldo As Currency
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov

On Error GoTo GrabaEdicionAperturaCajaChicaErr
GrabaEdicionAperturaCajaChica = 0

lbTrans = True

oDMov.BeginTrans
oDMov.actualizarMovCont pnMovNroCajaChica, pnMovMonto
oDMov.actualizarCajaChicaSinAprobar psPersCod, pnMovMonto, pnTopeEgreso, psAreCod, psAgeCod, pnProcNro
oDMov.CommitTrans

lbTrans = False
GrabaEdicionAperturaCajaChica = pnMovNroCajaChica

Set oDMov = Nothing
Exit Function
GrabaEdicionAperturaCajaChicaErr:
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    Err.Raise vbObjectError + 100, "GrabaEdicionAperturaCajaChica", "Error al Editar Nueva Caja chica"
End Function


Public Function devolverPorcetajeTopeEgreso() As Currency
Dim sql As String
Dim rsPorcentaje As ADODB.Recordset
Dim oDConecta As DConecta

Set oDConecta = New DConecta
Set rsPorcentaje = New Recordset

If oDConecta.AbreConexion = False Then Exit Function
 
sql = "exec stp_sel_DevolverPorcetajeTopeEgreso "
Set rsPorcentaje = oDConecta.CargaRecordSet(sql)

If Not rsPorcentaje.BOF And Not rsPorcentaje.EOF Then
    devolverPorcetajeTopeEgreso = CCur(rsPorcentaje!cConsDescripcion)
Else
    devolverPorcetajeTopeEgreso = 0#
End If

oDConecta.CierraConexion
Set oDConecta = Nothing
Set rsPorcentaje = Nothing

End Function

Public Function devolverTopeEgreso() As Currency
Dim sql As String
Dim rsPorcentaje As ADODB.Recordset
Dim oDConecta As DConecta

Set oDConecta = New DConecta
Set rsPorcentaje = New Recordset

If oDConecta.AbreConexion = False Then Exit Function
 
sql = "exec stp_sel_DevolverTopeEgreso"
Set rsPorcentaje = oDConecta.CargaRecordSet(sql)

If Not rsPorcentaje.BOF And Not rsPorcentaje.EOF Then
    devolverTopeEgreso = CCur(rsPorcentaje!cConsDescripcion)
Else
    devolverTopeEgreso = 0#
End If

oDConecta.CierraConexion
Set oDConecta = Nothing
Set rsPorcentaje = Nothing

End Function

Public Function devolverCHAprobadas(ByVal psAreaCh As String, _
                                    ByVal psAgeCh As String, _
                                    ByVal pnNroProc As Integer, _
                                    ByVal psCtaFondoFijo As String, _
                                    ByVal pbEfectivo As Boolean, _
                                    Optional psTpoDoc As String = "") As ADODB.Recordset

Dim sql As String
Dim rsAprobadas As ADODB.Recordset
Dim oDConecta As DConecta

Set oDConecta = New DConecta
Set rsAprobadas = New Recordset

If oDConecta.AbreConexion = False Then Exit Function
    sql = "exec stp_sel_DevolverCajasChicasAprobadoApertura '" & psAreaCh & "', '" & psAgeCh & "', " & pnNroProc & ", '" & psCtaFondoFijo & "', " & IIf(pbEfectivo, 1, 0) & ", '" & psTpoDoc & "'"
Set rsAprobadas = oDConecta.CargaRecordSet(sql)
Set devolverCHAprobadas = rsAprobadas
oDConecta.CierraConexion
Set oDConecta = Nothing
End Function

Public Function verificarEncargadoCH(ByVal psPersCod As String) As ADODB.Recordset
Dim sql As String
Dim rsEncargado As ADODB.Recordset
Dim oDConecta As DConecta

Set oDConecta = New DConecta
Set rsEncargado = New Recordset

If oDConecta.AbreConexion = False Then Exit Function
    sql = "exec stp_sel_VerificarEncargadoCH '" & psPersCod & "'"
Set rsEncargado = oDConecta.CargaRecordSet(sql)
Set verificarEncargadoCH = rsEncargado
oDConecta.CierraConexion
Set oDConecta = Nothing
End Function

Public Function grabarEntregaEfectivoCH(ByVal psMovNro As String, _
                                        ByVal psOpeCod As String, _
                                        ByVal psPersCod As String, _
                                        ByVal psAreCod As String, _
                                        ByVal psAgeCod As String, _
                                        ByVal pnProcNro As Integer, _
                                        ByVal pnImporte As Currency) As Long
Dim oDMov As DMov
Dim lnSaldo As Currency
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New DMov

On Error GoTo grabarEntregaEfectivoCHErr
grabarEntregaEfectivoCH = 0

lbTrans = True

oDMov.BeginTrans
oDMov.InsertaMov psMovNro, psOpeCod, "Entrega de Efectivo por Caja Chica", gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.registrarEntregaEfectivoCH lnMovNro, psPersCod, psAreCod, psAgeCod, pnProcNro, pnImporte
oDMov.CommitTrans
lbTrans = False
grabarEntregaEfectivoCH = lnMovNro
Set oDMov = Nothing
Exit Function
grabarEntregaEfectivoCHErr:
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    Err.Raise vbObjectError + 100, "grabarEntregaEfectivoCH", "Error al Grabar la Entrega de Efectivo de Caja Chica"
End Function

Public Function devolverCHSustAtenArendirTodo(ByVal psMovNro As String) As ADODB.Recordset
Dim sql As String
Dim rsSustentacion As ADODB.Recordset
Dim oDConecta As DConecta

Set oDConecta = New DConecta
Set rsSustentacion = New Recordset

If oDConecta.AbreConexion = False Then Exit Function
    sql = "exec stp_sel_DevolverCHSustAtenArendirTodo '" & psMovNro & "'"
Set rsSustentacion = oDConecta.CargaRecordSet(sql)
Set devolverCHSustAtenArendirTodo = rsSustentacion
oDConecta.CierraConexion
Set rsSustentacion = Nothing
Set oDConecta = Nothing
End Function

'***Fin Agregado por ELRO*******************************
'***Agregado por ELRO el 20120901, según OYP-RFC104-2012
Public Function devolverTotalEgresos(ByVal psAreaCod As String, _
                                     ByVal psAgeCod As String, _
                                     ByVal pnProcNro As Integer) As Currency
Dim sql As String
Dim rsTE As ADODB.Recordset
Dim oDConecta As DConecta

Set oDConecta = New DConecta
Set rsTE = New Recordset

If oDConecta.AbreConexion = False Then Exit Function
    sql = "exec stp_sel_DevolverTotalEgresos '" & psAreaCod & "', '" & psAgeCod & "', " & pnProcNro & ""
Set rsTE = oDConecta.CargaRecordSet(sql)

If Not rsTE.BOF And Not rsTE.EOF Then
    devolverTotalEgresos = rsTE!nTEEE
Else
    devolverTotalEgresos = 0
End If

oDConecta.CierraConexion
Set rsTE = Nothing
Set oDConecta = Nothing
End Function
'***Fin Agregado por ELRO el 20120901*******************
'***Agregado por ELRO el 20120927, según OYP-RFC111-2012
Public Sub extornarAutorizacionCajaChica(ByVal pnMovNroAut As Long)
Dim oMov As DMov
Set oMov = New DMov
Dim lbTrans As Boolean
On Error GoTo extornarAutorizacionCajaChicaErr
oMov.BeginTrans
lbTrans = True
oMov.ActualizaMov pnMovNroAut, , gMovEstContabRechazado, gMovFlagEliminado
lbTrans = False
oMov.CommitTrans
Exit Sub

extornarAutorizacionCajaChicaErr:
    If lbTrans Then
        lbTrans = False
        oMov.RollbackTrans
    End If
    Call RaiseError(MyUnhandledError, "NARendir:extornarAutorizacionCajaChica Method")
End Sub

'***Fin Agregado por ELRO el 20120927*******************
