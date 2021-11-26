VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NARendir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsBaseComunes As String
Dim vsBasePesonas As String
Dim lsMsgErr      As String

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As ClasIni
Set oIni = New ClasIni
vsBaseComunes = oIni.BaseComunes
vsBasePesonas = oIni.BasePersonas
Set oIni = Nothing
End Sub
Public Function EmiteCajasChicas() As ADODB.Recordset
On Error GoTo ErrorEmiteCajasChicas
Dim rs As ADODB.Recordset
Dim sSQL As String
Set rs = New ADODB.Recordset

Dim oConect As DConecta
Set oConect = New DConecta

If oConect.AbreConexion = False Then Exit Function

'Modificado xPASI20150514**********************************
'sSql = " SELECT  CH.CAREACOD + CH.CAGECOD AS Codigo, ISNULL(AG.CAGEDESCRIPCION ,A.CAREADESCRIPCION ) as Descripcion  , 1 AS NIVEL " _
'    & "  FROM    CAJACHICA CH JOIN " & vsBaseComunes & "AREAAGENCIA AA ON AA.CAREACOD = CH.CAREACOD " _
'    & "          JOIN " & vsBaseComunes & "AREAS A ON A.CAREACOD = CH.CAREACOD " _
'    & "          LEFT JOIN " & vsBaseComunes & "AGENCIAS AG ON AG.CAGECOD = CH.CAGECOD " _
'    & "  GROUP BY CH.CAREACOD, CH.CAGECOD , A.CAREADESCRIPCION, AG.CAGEDESCRIPCION " _
'    & "  ORDER BY CH.CAREACOD, CH.CAGECOD"

sSQL = "stp_sel_EmiteCajaChica"

'end PASI***********************************************

Set rs = oConect.CargaRecordSet(sSQL)
Set EmiteCajasChicas = rs

oConect.CierraConexion
Set oConect = Nothing
Exit Function
ErrorEmiteCajasChicas:
    Err.Raise vbObjectError + 100, "EmiteCajasChicas", "Error emitir Cajas Chicas"
End Function

Public Function ARendirPendientesTotal(pnMoneda As Moneda, psCtaCod As String, pdFecha As Date, Optional pbCajaChica As Boolean = True) As ADODB.Recordset
Dim sSQL As String
Dim lsTipoArendir As String

If pbCajaChica Then
    lsTipoArendir = gArendirTipoAgencias & "," & gArendirTipoCajaChica & "," & gArendirTipoCajaGeneral
Else
    lsTipoArendir = gArendirTipoAgencias & "," & gArendirTipoCajaGeneral
End If
'***Modificado por ELRO el 20121114, según SATI INC1211100007 y INC1211100008
'       sSql = "SELECT  Convert(datetime, LEFT(M.cMovNro,8), 103) dFecTran, M.NMOVNRO as NMovNroAtenc, M.cMovNro, M.cMovDesc, ISNULL(CONVERT(CHAR(10),MD.dDocFecha,103),SUBSTRING(M.CMOVNRO,7,2) + '/' + SUBSTRING(M.CMOVNRO,5,2)+ '/' + SUBSTRING(M.CMOVNRO,1,4)) AS dDocFecha , mc.cCtaContCod, dbo.GetCtaContDesc(mc.cCtaContCod,2,2) cCtaContDesc, ISNULL(D1.cDocAbrev,'EFEC') AS cDocAbrev, ISNULL(MD.cDocNro,'Efectivo') as cDocNro, " _
'            & "         CONVERT(CHAR(50),P.CPERSNOMBRE) AS cPersona, ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS NMOVIMPORTE,  " _
'            & "         P.cPersCod, ISNULL(MD.nDocTpo,'') as cTpoDocAtenc, ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) + ISNULL(Ref.nMovImporte,0) AS nSaldo ,1 Rendir" _
'            & "  FROM   MOV M " _
'            & "         JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO " _
'            & "         LEFT JOIN MOVME ME ON ME.NMOVNRO =MC.NMOVNRO AND MC.NMOVITEM = ME.NMOVITEM " _
'            & "         LEFT JOIN MOVDOC MD ON MD.NMOVNRO = M.NMOVNRO  AND MD.nDocTpo NOT IN ('" & Format(TpoDocVoucherEgreso, "00") & "') " _
'            & "         LEFT JOIN DOCUMENTO D1 ON D1.nDocTpo = MD.nDocTpo " _
'            & "         JOIN MOVREF MR ON MR.NMOVNRO = M.NMOVNRO and MR.cAgeCodRef = '' " _
'            & "         JOIN MOVARENDIR MAR ON MAR.NMOVNRO=MR.NMOVNROREF " _
'            & "         JOIN PERSONA P ON P.CPERSCOD  = MAR.CPERSCOD " _
'            & "         LEFT JOIN (SELECT a.NMovNroRef, ISNULL(SUM(ISNULL(me1.nMovMEImporte,c.nMovImporte)),0) as nMovImporte " _
'            & "                    FROM MovRef a JOIN MovCta c ON c.NMovNro = a.NMovNro " _
'            & "                                JOIN Mov   m ON m.NMovNro = c.NMovNro LEFT JOIN MovMe me1 ON me1.NMovNro = c.NMovNro and me1.NMovItem = c.NMovItem " _
'            & "                    WHERE left(m.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' and m.NMovEstado = 10 and (not m.NMovFlag  IN (1,2,3,5) or (m.NMovFlag = 2 and EXISTS (SELECT NMovNro FROM MovRef WHERE NMovNroRef = m.NMovNro and LEFT(cMovNro,8) > '" & Format(pdFecha, gsFormatoMovFecha) & "'))) " _
'            & "                        and c.cCtaContCod IN ('" & psCtaCod & "') GROUP BY  a.NMovNroRef ) as Ref ON Ref.NMovNroRef = m.NMovNro " _
'            & " WHERE   LEFT(M.CMOVNRO,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' AND MAR.cTpoArendir in (" & lsTipoArendir & ") " _
'            & "         AND MC.CCTACONTCOD IN ('" & psCtaCod & "') AND Substring(M.cOpeCod,3,1)='" & pnMoneda & "' AND M.nMovEstado = " & gMovEstContabMovContable & " And M.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "') " _
'            & "         AND ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) + ISNULL(Ref.nMovImporte,0) <> 0 "
'
'    sSql = sSql & " UNION ALL "
'    sSql = sSql & " SELECT  Convert(datetime, LEFT(M.cMovNro,8), 103) dFecTran, M.NMOVNRO as NMovNroAtenc, M.cMovNro, M.cMovDesc, ISNULL(CONVERT(CHAR(10),MD.dDocFecha,103),SUBSTRING(M.CMOVNRO,7,2) + '/' + SUBSTRING(M.CMOVNRO,5,2)+ '/' + SUBSTRING(M.CMOVNRO,1,4)) AS dDocFecha , atend.cCtaContCod, dbo.GetCtaContDesc(atend.cCtaContCod,2,2) cCtaContDesc, ISNULL(D.cDocAbrev,'EFEC') AS cDocAbrev, ISNULL(MD.cDocNro,'Efectivo') as cDocNro, " _
'            & "         CONVERT(CHAR(50),P.CPERSNOMBRE) AS cPersona, ISNULL(ATEND.NATENDIDO,0) AS NMOVIMPORTE,  " _
'            & "         P.CPERSCOD, ISNULL(MD.nDocTpo,''), ISNULL(ATEND.NATENDIDO,0) + ISNULL(Ref.nMovImporte,0) AS nSaldo, 2 Rendir" _
'            & " FROM    MOV M JOIN MOVARENDIR MAR ON M.NMOVNRO = MAR.NMOVNRO " _
'            & "         JOIN PERSONA            P       ON P.CPERSCOD = MAR.CPERSCOD " _
'            & "         JOIN MOVDOC         MD  ON MD.nMOVNRO = M.nMOVNRO AND nDocTpo IN (" & TpoDocRecViaticosARendirCuenta & ") " _
'            & "         JOIN DOCUMENTO      D   ON D.nDocTpo = MD.nDocTpo " _
'            & "         JOIN (  SELECT  mc.cCtaContCod, SUM(ISNULL(ME.nMovMEImporte,MC.NMOVIMPORTE)) AS NATENDIDO, MR.NMOVNROREF " _
'            & "                 FROM    MOVCTA MC JOIN MOV M1 ON M1.NMOVNRO = MC.NMOVNRO LEFT JOIN MovMe ME ON me.nMovNRo = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'            & "                         JOIN MOVREF MR ON MR.NMOVNRO  = M1.NMOVNRO And MC.cCtaContCod Not Like (Select nConsSisValor from constsistema where nConsSisCod = 157) " _
'            & "                 WHERE   LEFT(m1.CMOVNRO,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' AND MC.NMOVIMPORTE>0 AND M1.NMOVESTADO = " & gMovEstContabMovContable & " AND M1.NMOVFLAG =" & gMovFlagVigente & "  " _
'            & "                         AND SUBSTRING(M1.COPECOD,1,5) IN ('" & Mid(gCGArendirViatAtencMN, 1, 5) & "','" & Mid(gCGArendirViatAtencME, 1, 5) & "') " _
'            & "                 GROUP  BY mc.cCtaContCod, MR.NMOVNROREF ) AS ATEND  ON ATEND.NMOVNROREF = M.NMOVNRO " _
'            & "        LEFT JOIN (SELECT a.NMovNroRef, ISNULL(SUM(ISNULL(me1.nMovMEImporte,c.nMovImporte)),0) as nMovImporte " _
'            & "                  FROM MovRef a JOIN MovCta c ON c.NMovNro = a.NMovNro " _
'            & "                                JOIN Mov   m ON m.NMovNro = c.NMovNro LEFT JOIN MovMe me1 ON me1.NMovNro = c.NMovNro and me1.NMovItem = c.NMovItem " _
'            & "                  WHERE left(m.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' AND NOT SUBSTRING(M.COPECOD,1,5) IN ('40123','40223') " _
'            & "                        and m.NMovEstado = 10 and (not m.NMovFlag  IN (1,2,3,5) or (m.NMovFlag = 2 and EXISTS (SELECT NMovNro FROM MovRef WHERE NMovNroRef = m.NMovNro and LEFT(cMovNro,8) > '" & Format(pdFecha, gsFormatoMovFecha) & "'))) " _
'            & "                        and c.cCtaContCod IN ('" & psCtaCod & "') GROUP BY  a.NMovNroRef ) as Ref ON Ref.NMovNroRef = m.NMovNro " _
'
'    sSql = sSql + " WHERE   MAR.cTpoArendir ='" & gArendirTipoViaticos & "' AND SUBSTRING(M.COPECOD,3,1)='" & pnMoneda & "' " _
'            & "         AND M.nMovFlag =" & gMovFlagVigente & " " _
'            & "         AND ISNULL(ATEND.NATENDIDO,0) + ISNULL(Ref.nMovImporte,0)<> 0 " _
'            & " ORDER BY Rendir, M.cMovNro "
sSQL = "exec stp_sel_DevolverARendirPendientesTotal '" & pnMoneda & "', '" & psCtaCod & "', '" & Format(pdFecha, gsFormatoMovFecha) & "', " & IIf(pbCajaChica, 1, 0) & ""
'***Modificado por ELRO el 20121114******************************************
Dim oCon As DConecta
Set oCon = New DConecta
oCon.AbreConexion
Set ARendirPendientesTotal = oCon.CargaRecordSet(sSQL)
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function ARendirPendientes(ByVal pnTipoArendir As ArendirTipo, ByVal psDocTpo As String, _
                                ByVal pnMoneda As Moneda, ByVal lsAreaArendir As String, ByVal lsAgeArendir As String) As ADODB.Recordset
Dim oConect As DConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConect = New DConecta
Set rs = New ADODB.Recordset

If oConect.AbreConexion = False Then Exit Function
Select Case pnTipoArendir
    Case gArendirTipoCajaGeneral, gArendirTipoAgencias
        sql = " Select  MD.cDocNro,  " _
            & "         ISNULL(A.cAreaDescripcion,'') cAreaDescripcion , ISNULL(AG.cAgeDescripcion,'') AS cAgeDescripcion, " _
            & "         P.cPersNombre,  convert(varchar(10),md.dDocFecha,103) as Fecha, " _
            & "         MC.nMovMonto,M.cMovDesc, ISNULL(A.cAreaCod,'') AS cAreaCod , P.cPersCod, M.nMovNro, MD.nDocTpo, " _
            & "         ISNULL(AG.cAgeCod,'') AS cAgeCod " _
            & " From    Mov m          Join MovARendir MR  ON MR.nMovNro = M.nMovNro " _
            & "         JOIN MovDoc MD  ON MD.nMovNro = MR.nMovNro JOIN MOVCONT MC ON MC.nMOVNRO = M.nMOVNRO  " _
            & "         JOIN " & vsBaseComunes & "Documento D    ON D.nDocTpo = MD.nDocTpo " _
            & "         JOIN " & vsBasePesonas & "Persona P      ON P.cPersCod = MR.cPersCod " _
            & "         JOIN RRHH RH ON RH.cPersCod = P.cPersCod " _
            & "         LEFT JOIN " & vsBaseComunes & "Areas A   ON A.cAreaCod = RH.cAreaCodActual   " _
            & "         LEFT JOIN " & vsBaseComunes & "Agencias AG   ON AG.cAgeCod = RH.cAgenciaActual " _
            & "         JOIN " & vsBaseComunes & "Constante C    ON MR.cTpoARendir = C.nConsValor " _
            & "         AND C.nCONSCOD = '" & gArendirTipo & "' " _
            & " WHERE   MR.cTpoArendir = " & pnTipoArendir & " " _
            & "         And SUBSTRING(M.COPECOD,3,1)='" & pnMoneda & "' " _
            & "         AND MR.cAreaCod ='" & lsAreaArendir & "' and MR.cAgeCod ='" & lsAgeArendir & "' " _
            & "         AND M.nMovFlag IN (" & gMovFlagVigente & ") AND MD.nDocTpo ='" & psDocTpo & "'  " _
            & "         AND NOT EXISTS (    Select  MRef.nMovNroRef " _
            & "                             FROM    MOVRef MRef JOIN Mov M1 on M1.nMovNro = MRef.nMovNro " _
            & "                             WHERE   MRef.nMovNroRef = M.nMovNro and M1.nMovFlag Not in ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "')) " _
            & " ORDER BY P.cPersNombre "

    Case gArendirTipoViaticos
        'm.NMOVNRO , m.cMovNro, MAR.NMOVSALDO, IsNull(ATEND.NATENDIDO, 0), SOL.NSOLICITADO
        sql = " SELECT  D.CDOCABREV + '-' + MD.CDOCNRO AS DocNro, " _
            & "         ISNULL(A.CAREADESCRIPCION,'') AS CAREADESCRIPCION, ISNULL(AG.CAGEDESCRIPCION,'') CAGEDESCRIPCION, " _
            & "         P.CPERSNOMBRE, convert(varchar(10),md.dDocFecha,103) AS FECHA, " _
            & "         (ISNULL(SOL.NSOLICITADO,0) - ISNULL(ATEND.NATENDIDO,0) ) AS MONTOPEND , " _
            & "         M.cMovDesc, MAR.CAREACOD, P.CPERSCOD, M.NMOVNRO AS nVIATICOMOVNRO, MD.nDocTpo, MAR.CAGECOD  " _
            & " FROM    MOV M JOIN MOVARENDIR MAR ON M.NMOVNRO = MAR.NMOVNRO " _
            & "         JOIN PERSONA            P       ON P.CPERSCOD = MAR.CPERSCOD " _
            & "         JOIN RRHH               RH      ON RH.CPERSCOD = P.CPERSCOD " _
            & "         LEFT JOIN AREAS         A       ON A.CAREACOD = RH.cAreaCodActual " _
            & "         LEFT JOIN AGENCIAS  AG  ON AG.CAGECOD = RH.cAgenciaActual " _
            & "         JOIN MOVDOC         MD  ON MD.nMOVNRO = MAR.nMOVNRO AND nDocTpo IN (" & TpoDocRecViaticosARendirCuenta & ") " _
            & "         JOIN DOCUMENTO      D   ON D.nDocTpo = MD.nDocTpo " _
            & "         LEFT JOIN ( SELECT  SUM( Case " & pnMoneda & " When 1 Then MC.NMOVIMPORTE Else ME.nMovMEImporte End) AS NATENDIDO, MR.NMOVNROREF " _
            & "                     FROM    MOVCTA MC LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO And  ME.NMOVITEM = MC.NMOVITEM JOIN MOV M1 ON M1.NMOVNRO = MC.NMOVNRO " _
            & "                             JOIN MOVREF MR ON MR.NMOVNRO  = M1.NMOVNRO And MC.cCtaContCod Not Like (Select nConsSisValor from constsistema where nConsSisCod = 157) " _
            & "                     WHERE   MC.NMOVIMPORTE>0 AND M1.NMOVFLAG = " & gMovFlagVigente & "  " _
            & "                             AND SUBSTRING(M1.COPECOD,1,5) IN ('" & Mid(gCGArendirViatAtencMN, 1, 5) & "','" & Mid(gCGArendirViatAtencME, 1, 5) & "') " _
            & "                     GROUP  BY MR.NMOVNROREF ) AS ATEND " _
            & "                     ON ATEND.NMOVNROREF = M.NMOVNRO "
        sql = sql & "   LEFT JOIN ( SELECT  SUM(NMOVIMPORTE) AS NSOLICITADO, MV.NVIATICOMOVNRO " _
            & "                     FROM    MOVOBJVIATICOS MC JOIN MOV M1 ON M1.NMOVNRO = MC.NMOVNRO " _
            & "                             JOIN MOVVIATICOS MV ON MV.NMOVNRO = M1.NMOVNRO " _
            & "                     Where M1.NMOVFLAG = " & gMovFlagVigente & "  " _
            & "                     GROUP BY MV.NVIATICOMOVNRO ) AS SOL " _
            & "                     ON SOL.NVIATICOMOVNRO = M.NMOVNRO " _
            & " WHERE   MAR.cTpoArendir ='" & gArendirTipoViaticos & "' AND SUBSTRING(M.COPECOD,3,1)='" & pnMoneda & "' " _
            & "         AND M.nMovFlag =" & gMovFlagVigente & " " _
            & "         AND MAR.cAreaCod ='" & lsAreaArendir & "' and MAR.cAgeCod ='" & lsAgeArendir & "' " _
            & "         AND (SOL.NSOLICITADO- ISNULL(ATEND.NATENDIDO,0) )>0 "
    
    Case Else
End Select
Set rs = oConect.CargaRecordSet(sql)
Set ARendirPendientes = rs
oConect.CierraConexion: Set oConect = Nothing

End Function
Public Function GetDatosArendir(ByVal lsMovNroArendirAtencion As String, ByVal psCtaContArendir As String, ByVal psCtaContPendiente As String, ByVal pnMoneda As Moneda, Optional ByVal pnTipoArendir As ArendirTipo = gArendirTipoCajaGeneral) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set rs = New ADODB.Recordset
Set oConect = New DConecta

If oConect.AbreConexion = False Then Exit Function

Select Case pnTipoArendir
    Case gArendirTipoCajaGeneral, gArendirTipoAgencias
        sql = " SELECT  M.CMOVNRO AS CMOVSOL , MAT.CMOVNRO AS  CMOVATENC, " _
            & "         MAR.CPERSCOD , P.CPERSNOMBRE, ISNULL(A.CAREACOD,'') cAreaCod, ISNULL(A.CAREADESCRIPCION,'') CAREADESCRIPCION, " _
            & "         ISNULL(AG.CAGECOD,'') AS CAGECOD , ISNULL(AG.CAGEDESCRIPCION,'') CAGEDESCRIPCION,  ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS ImporteArendir, " _
            & "         MAR.CAREACOD as cAreaCodArendir, A1.CAREADESCRIPCION as cAreaARendir ,  ISNULL(MAR.CAGECOD,'') AS cAgeCodArendir , ISNULL(AG1.CAGEDESCRIPCION,'') cAgeArendir ,  " _
            & "         MAR.NMOVSALDO , MD.nDocTpo , D.CDOCABREV , MD.CDOCNRO, MD.dDocFecha as dDocFechaSol , " _
            & "         MD1.nDocTpo AS nDocTpoATENC, D1.CDOCABREV AS CDOCABREVATENC, MD1.CDOCNRO AS CDOCNROATENC , D1.cDocDesc  AS cDocDescAtenc, MD1.dDocFecha as dFechaDocAtenc, " _
            & "         SUST.cMovSust , SUST.cMovDesc  AS cDescSust, SUST.cDocAbrev as cDocNroAbrevSust, " _
            & "         SUST.cDocNro AS CDOCNROSUST , SUST.nDocTpo AS nDocTpoSUST, SUST.dDocFecha AS dDocFechaSust, " _
            & "         SUST.cPersNombre AS cPersCodSust, SUST.nDocImporte " _
            & "  FROM   MOV M JOIN MOVREF MR ON M.nMOVNRO = MR.nMOVNROREF " _
            & "         JOIN MOVARENDIR MAR ON MAR.nMOVNRO = M.nMOVNRO " _
            & "         JOIN " & vsBasePesonas & "PERSONA P ON P.CPERSCOD = MAR.CPERSCOD " _
            & "         LEFT JOIN RRHH RH ON RH.CPERSCOD = P.CPERSCOD " _
            & "         LEFT JOIN " & vsBaseComunes & "AREAS A ON A.CAREACOD = RH.cAreaCodActual   " _
            & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG ON AG.CAGECOD = RH.cAgenciaActual " _
            & "         LEFT JOIN " & vsBaseComunes & "AREAS A1 ON A1.CAREACOD = MAR.CAREACOD " _
            & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG1 ON AG1.CAGECOD = MAR.CAGECOD " _
            & "         JOIN MOVDOC MD ON MD.nMOVNRO = M.nMOVNRO " _
            & "         JOIN " & vsBaseComunes & "DOCUMENTO D ON D.nDocTpo = MD.nDocTpo " _
            & "         JOIN MOV AS MAT ON MAT.nMOVNRO = MR.nMOVNRO " _
            & "         JOIN MOVCTA MC ON MC.nMOVNRO  = MAT.nMOVNRO " _
            & "         LEFT JOIN MOVME ME ON ME.nMOVNRO =MC.nMOVNRO AND ME.nMOVITEM=MC.nMOVITEM " _
            & "         LEFT JOIN MOVDOC  MD1 ON MD1.nMOVNRO = MAT.nMOVNRO AND MD1.nDocTpo NOT IN ('" & TpoDocVoucherEgreso & "') " _
            & "         LEFT JOIN " & vsBaseComunes & "DOCUMENTO D1 ON D1.nDocTpo =MD1.nDocTpo " _
            & "         Join "
        sql = sql + "  (SELECT  M.cMovNro AS cMovSust, M.cMovDesc, D.cDocAbrev, MD.cDocNro, MD.nDocTpo, p.cPersNombre,MD.dDocFecha, " _
            & "                      ABS(SUM(IsNull(ME.NMOVMEIMPORTE, Mc.NMOVIMPORTE))) As nDocImporte , MRSUST.nMOVNROREF  AS nMovAtenc " _
            & "              FROM    MOVREF  MRSUST " _
            & "                      JOIN MOV M          ON MRSUST.nMOVNRO = M.nMOVNRO " _
            & "               JOIN MOVCTA MC      ON MC.nMOVNRO = M.nMOVNRO " _
            & "                      LEFT JOIN MOVME ME  ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM =MC.nMOVITEM " _
            & "                      JOIN MOVDOC MD      ON MD.nMOVNRO = M.nMOVNRO " _
            & "                      JOIN " & vsBaseComunes & "DOCUMENTO D    ON D.nDocTpo = MD.nDocTpo " _
            & "                      JOIN MovGasto MG    ON MG.nMOVNRO = M.nMOVNRO " _
            & "                      JOIN " & vsBasePesonas & "PERSONA P      ON P.CPERSCOD = MG.CPERSCOD " _
            & "              WHERE   cCtaContCod IN ('" & psCtaContArendir & "','" & psCtaContPendiente & "') " _
            & "                      And  M.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "') AND MRSUST.nMOVNROREF='" & lsMovNroArendirAtencion & "'  " _
            & "              GROUP BY M.cMovNro, M.cMovDesc, D.cDocAbrev, MD.cDocNro, MD.nDocTpo, p.cPersNombre,MD.dDocFecha, MRSUST.nMOVNROREF ) as SUST " _
            & "                     ON SUST.nMovAtenc= MR.nMovNro " _
            & " WHERE SUBSTRING(M.COPECOD,3,1)='" & pnMoneda & "' and  MC.CCTACONTCOD IN ('" & psCtaContArendir & "')  AND MR.nMOVNRO =" & lsMovNroArendirAtencion & " "
                   
      Case gArendirTipoViaticos
      
        sql = "SELECT   M.CMOVNRO AS CMOVSOL, M.CMOVNRO AS  CMOVATENC, MAR.CPERSCOD , P.CPERSNOMBRE, " _
            & "         isnull(A.CAREACOD,'') CAREACOD, isnull(A.CAREADESCRIPCION,'') CAREADESCRIPCION, ISNULL(AG.CAGECOD,'') AS CAGECOD , ISNULL(AG.CAGEDESCRIPCION,'') CAGEDESCRIPCION, " _
            & "         MAR.CAREACOD as cAreaCodArendir, A1.CAREADESCRIPCION as cAreaARendir ,  ISNULL(MAR.CAGECOD,'') AS cAgeCodArendir , ISNULL(AG1.CAGEDESCRIPCION,'') cAgeArendir ,  " _
            & "         REND.ImporteArendir, MAR.NMOVSALDO ,  MD.nDocTpo , D.CDOCABREV , (D.CDOCABREV + '-'  + MD.CDOCNRO) as CDOCNRO , MD.dDocFecha as dDocFechaSol ,  " _
            & "         ISNULL(Rend.nDocTpo,0) AS nDocTpoATENC,  ISNULL(Rend.CDOCABREV,'') AS CDOCABREVATENC, ISNULL(Rend.CDOCNRO,'NINGUNO') AS CDOCNROATENC , " _
            & "         Rend.CDOCDesc AS cDocDescAtenc, Rend.dDocFecha as dFechaDocAtenc, " _
            & "         SUST.nMovSust , SUST.cMovDesc  AS cDescSust, SUST.cDocAbrev as cDocNroAbrevSust, " _
            & "         SUST.cDocNro AS CDOCNROSUST , SUST.nDocTpo AS nDocTpoSUST, SUST.dDocFecha AS dDocFechaSust,  " _
            & "         SUST.cPersNombre AS cPersCodSust, SUST.nDocImporte " _
            & " FROM    MOV M " _
            & "         JOIN MOVARENDIR MAR                         ON MAR.nMOVNRO = M.nMOVNRO " _
            & "         JOIN " & vsBasePesonas & "PERSONA P         ON P.CPERSCOD = MAR.CPERSCOD " _
            & "         LEFT JOIN RRHH   RH                         ON RH.CPERSCOD = P.CPERSCOD " _
            & "         LEFT JOIN " & vsBaseComunes & "AREAS A      ON A.CAREACOD = RH.cAreaCodActual   " _
            & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG  ON AG.CAGECOD = RH.cAgenciaActual " _
            & "         JOIN " & vsBaseComunes & "AREAS A1          ON A1.CAREACOD = MAR.CAREACOD " _
            & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS     AG1 ON AG1.CAGECOD = MAR.CAGECOD " _
            & "         JOIN MOVDOC MD                              ON MD.nMOVNRO = M.nMOVNRO AND MD.nDocTpo IN ('" & TpoDocRecViaticosARendirCuenta & "') " _
            & "         JOIN DOCUMENTO D                            ON D.nDocTpo = MD.nDocTpo " _
            & "         Join (  SELECT  MC.cCtaContCod, MIN(md.cDocNro) cDocNro, MIN(md.nDocTpo) nDocTpo, MIN(d.cDocAbrev) cDocAbrev, MIN(d.cDocDesc) cDocDesc, MIN(md.dDocFecha) dDocFecha, " _
            & "                         SUM (IsNull(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)) as ImporteArendir   , " _
            & "                         MRSUST.nMOVNROREF AS nMovSol "
sql = sql + "                   FROM    MOVREF  MRSUST " _
            & "                         JOIN MOV M          ON MRSUST.nMOVNRO = M.nMOVNRO " _
            & "                         JOIN MOVCTA MC      ON MC.nMOVNRO = M.nMOVNRO " _
            & "                         LEFT JOIN MOVME ME  ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM =MC.nMOVITEM " _
            & "                         LEFT JOIN MovDoc md      ON md.nMovNro = m.nMovNro and md.nDocTpo IN (" & TpoDocOrdenPago & "," & TpoDocNotaAbono & "," & TpoDocCheque & "," & TpoDocCarta & ") LEFT JOIN Documento d ON d.nDocTpo = md.nDocTpo " _
            & "                 WHERE   M.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagModificado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "')  AND SUBSTRING(M.COPECOD,1,5) IN ('" & Mid(gCGArendirViatAtencMN, 1, 5) & "','" & Mid(gCGArendirViatAtencME, 1, 5) & "') " _
            & "                         AND MRSUST.nMOVNROREF ='" & lsMovNroArendirAtencion & "' AND MC.cCtaContCod  IN ('" & psCtaContArendir & "','" & psCtaContPendiente & "') " _
            & "                         " _
            & "                 GROUP BY MC.cCtacontCod, MRSUST.nMOVNROREF ) as REND " _
            & "                ON M.nMovNro = REND.nMovSol "
sql = sql + "         Join   (SELECT    M.nMovNro AS nMovSust, M.cMovDesc, D.cDocAbrev, MD.cDocNro, " _
            & "                         MD.nDocTpo, p.cPersNombre,MD.dDocFecha, " _
            & "                         abs(SUM(IsNull(ME.NMOVMEIMPORTE, MC.NMOVIMPORTE))) As nDocImporte , " _
            & "                         MRSUST.nMOVNROREF  AS nMovNroSol " _
            & "                 FROM    MOVREF  MRSUST " _
            & "                         JOIN MOV M          ON MRSUST.nMOVNRO = M.nMOVNRO " _
            & "                         JOIN MOVCTA MC      ON MC.nMOVNRO = M.nMOVNRO " _
            & "                         LEFT JOIN MOVME ME  ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM =MC.nMOVITEM " _
            & "                         JOIN MOVDOC MD      ON MD.nMOVNRO = M.nMOVNRO " _
            & "                         JOIN DOCUMENTO D    ON D.nDocTpo = MD.nDocTpo " _
            & "                         JOIN MovGasto MG    ON MG.nMOVNRO = M.nMOVNRO " _
            & "                         JOIN PERSONA P      ON P.CPERSCOD = MG.CPERSCOD " _
            & "                 WHERE   cCtaContCod IN ('" & psCtaContArendir & "','" & psCtaContPendiente & "')  AND SUBSTRING(M.COPECOD,1,5) NOT IN ('" & Mid(gCGArendirViatAtencMN, 1, 5) & "','" & Mid(gCGArendirViatAtencME, 1, 5) & "') " _
            & "                         And  M.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "') AND MRSUST.nMOVNROREF='" & lsMovNroArendirAtencion & "' " _
            & "                 GROUP BY M.nMovNro, M.cMovDesc, D.cDocAbrev, MD.cDocNro, MD.nDocTpo, " _
            & "                         p.cPersNombre,MD.dDocFecha, MRSUST.nMOVNROREF ) as SUST " _
            & "         ON SUST.nMovNroSol = M.nMovNro " _
            & " WHERE   SUBSTRING(M.COPECOD,3,1)='" & pnMoneda & "' and M.nMOVNRO ='" & lsMovNroArendirAtencion & "' "
    
    Case Else
        
End Select
Set rs = oConect.CargaRecordSet(sql)
Set GetDatosArendir = rs
oConect.CierraConexion
Set oConect = Nothing
End Function
Public Function GrabaAtencionArendir(ByVal pnTipoArendir As ArendirTipo, _
            ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, ByVal psCtaDebe As String, ByVal psCtaHaber As String, _
            ByVal psPersCod As String, ByVal pnImporte As Currency, ByVal psTpoIF As String, ByVal psPersCodIf As String, ByVal psCtaIf As String, _
            ByVal rsBilletaje As ADODB.Recordset, ByVal psDocTpo As String, ByVal psDocNro As String, ByVal psFechaDoc As String, _
            ByVal psDocTpoVoucher As String, ByVal psDocVoucher As String, _
            ByVal pnMovRef As Long, ByVal psCuentaAho As String, Optional gbBitCentral As Boolean = False, Optional pbGrabaOpeNegocio As Boolean = True, Optional lsCtaITFD As String, Optional lsCtaITFH As String, Optional pnTasaITF As Double) As Integer
            
Dim lnMovItem As Long
Dim lnMovOrden As Integer
Dim lnMovNro As Long
Dim oDMov As DMov
'Dim lnMontoITF As Currency
Dim lnMontoITF As Double '*** PEAC 20110331

'Graba Abono Cta
Dim oCon As New DConecta
Dim oDis As New NRHProcesosCierre

Dim lbTrans As Boolean
Dim lsMovNro As String
Dim lnSaldoArendir As Currency
Dim lbGraboNegocio As Boolean
Set oDMov = New DMov
Dim lnConfirmarAprobacion As Integer '***Agregado por ELRO el 20120928, según INC1209250017 y INC1209260002

On Error GoTo ErrorGrabaAtencionArendir

'***Agregado por ELRO el 20120928, según INC1209250017 y INC1209260002
'***Modificado por ELRO el 20121217, para que permita ampliar un viatico
'If psOpeCod = gCGArendirViatAprobMN Or psOpeCod = gCGArendirViatAprobME Then
'    lnConfirmarAprobacion = oDMov.verificarAprobacionARendirViatico(pnMovRef)
'    If lnConfirmarAprobacion >= 1 Then
'      Exit Function
'    End If
'Else
'***Fin Modificado por ELRO el 20121217, para que permita ampliar un viatico
If psOpeCod = gCGArendirCtaAprobMN Or psOpeCod = gCGArendirCtaAprobME Then
    lnConfirmarAprobacion = oDMov.verificarAprobacionARendirCuenta(pnMovRef)
    If lnConfirmarAprobacion >= 1 Then
      Exit Function
    End If
End If
'***Fin Agregado por ELRO el 20121001*********************************

oDMov.BeginTrans
lbTrans = True
lbGraboNegocio = False

'VAPI RS179-2014, Si es una operacion de aprobacion de A rendir o A rendir viaticos se registra como un movimiento no contable (13)
'oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente   '<--- COMENTADO POR VAPI, SE REEMPLAZÓ POR EL CODIGO DE ABAJO

If psOpeCod = "401125" Or psOpeCod = "402125" Or psOpeCod = "401225" Or psOpeCod = "402225" Then
    oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
Else
    oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
End If
'FIN VAPI

lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"

'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1: lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, pnImporte
'guardamos el billetaje si este posee

If Not rsBilletaje Is Nothing Then
    oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsBilletaje, psCtaHaber, "H"
Else
    'guardamos la cuenta del haber
    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnImporte * -1
    If psPersCodIf <> "" Then
        'guardamos la Cta de la IF si es que utiliza esta
        lnMovOrden = lnMovOrden + 1
        oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCodIf, psTpoIF, psCtaIf

        'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
'        oDMov.InsertaMovCtaIF lnMovNro, psPersCodIf, psTpoIf, psCtaIf, CGCtaIFConCapital, pnImporte * -1
'        oDMov.ActualizaSaldoCtaIF CDate(GetFechaMov(psMovNro, True)), _
'                                    psPersCodIf, psTpoIf, psCtaIf
    End If
End If
If psDocNro <> "" Then
   oDMov.InsertaMovDoc lnMovNro, psDocTpo, psDocNro, Format(CDate(psFechaDoc), gsFormatoFecha)
End If
If psDocVoucher <> "" Then
   oDMov.InsertaMovDoc lnMovNro, psDocTpoVoucher, psDocVoucher, Format(CDate(psFechaDoc), gsFormatoFecha)
End If
If pnTipoArendir = gArendirTipoViaticos Then
   GrabarReferenciaViat oDMov, pnMovRef, lnMovNro
End If
oDMov.InsertaMovGasto lnMovNro, psPersCod, "0"

lnSaldoArendir = GetSaldoARendir(pnMovRef)
lnSaldoArendir = lnSaldoArendir + pnImporte
oDMov.ActualizaMovArendir pnMovRef, pnTipoArendir, , , , lnSaldoArendir
oDMov.InsertaMovRef lnMovNro, pnMovRef


If pnTasaITF <> 0 Then
   lnMontoITF = fgTruncar(pnImporte * pnTasaITF, 2)
   If lsCtaITFD <> "" And lsCtaITFH <> "" And pnTasaITF > 0 Then
       lnMovItem = lnMovItem + 1
       oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFD, lnMontoITF
       lnMovItem = lnMovItem + 1
       oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), lsCtaITFH, lnMontoITF * -1
   End If
End If


If Mid(psOpeCod, 3, 1) = "2" Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If

oDMov.ActualizaSaldoMovimiento psMovNro, "+"
'***Modificado por ELRO el 20120410, según OYP-RFC005-2012
'If psDocTpo = TpoDocNotaAbono Or psDocTpo = TpoDocNotaCargo Then
'***Fin Modificado por ELRO***********************************
If psDocTpo = CStr(TpoDocNotaAbono) Or psDocTpo = CStr(TpoDocNotaCargo) Then
    oDMov.InsertaRegDocCuenta psDocTpo, psDocNro, psCuentaAho
    If pbGrabaOpeNegocio Then
        Dim oCapta As NCapMovimientos
        Dim lnSaldo As Double
        Set oCapta = New NCapMovimientos
        If gbBitCentral Then
           lbGraboNegocio = True
           lnSaldo = oCapta.CapAbonoCuentaAho(psCuentaAho, pnImporte, gCGPagProvAbonoCent, psMovNro, psMovDesc, TpoDocNotaAbono, psDocNro, , , , , , , , , , True, False, , , oDMov.GetConexion, , True, CDbl(lnMontoITF))
        Else
            If oCon.AbreConexion(Left(psCuentaAho, 2)) Then
                oCon.BeginTrans
                lbGraboNegocio = True
                lnSaldo = oDis.Abono(psCuentaAho, pnImporte, gCGPagProvAbonoDist, gCGPagProvAbonoDist, "112" & Left(psCuentaAho, 2), Right(psMovNro, 4), psDocNro, "PAGO : " & psMovDesc, oCon, GetFechaMov(psMovNro, True))
            End If
            Set oDis = Nothing
        End If
    End If
End If


oDMov.CommitTrans
If pbGrabaOpeNegocio And lbGraboNegocio And Not gbBitCentral Then
    oCon.CommitTrans
    oCon.CierraConexion
    Set oCon = Nothing
End If
lbTrans = False
GrabaAtencionArendir = 0
Set oDMov = Nothing
Exit Function
ErrorGrabaAtencionArendir:
    lsMsgErr = Err.Description
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
        If lbGraboNegocio Then
            oCon.RollbackTrans
        End If
    End If
    If lbGraboNegocio Then
        lsMsgErr = lsMsgErr & Chr(10) & Chr(10) & "EL SISTEMA HIZO EL ABONO EN LA CUENTA DEL CLIENTE. POR FAVOR EXTORNAR ESTA OPERACION"
    End If
    Err.Raise vbObjectError + 100, "GrabaAtencionArendir", TextErr(lsMsgErr)
End Function

Public Function GrabaAtencionArendirNew(ByVal pnTipoArendir As ArendirTipo, ByVal pdFecha As Date, ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psDocNro As String, ByVal psMovDesc As String, _
                                        ByVal psPersCod As String, ByVal pnImporte As Double, ByVal pnMovRef As Long, ByVal psCuentaAho As String, Optional gbBitCentral As Boolean = False) As Integer
            
Dim lnMovItem As Long
Dim lnMovOrden As Integer
Dim lnMovNro As Long
Dim oDMov As DMov

Dim oCon As New DConecta

Dim lbTrans As Boolean
Dim lnSaldoArendir As Currency
Set oDMov = New DMov

Dim pnMoneda As String
Dim rsMovGenAs As New ADODB.Recordset
Dim rsOpeCta As New ADODB.Recordset
Dim rsExisMovCta As New ADODB.Recordset
Dim lnMovItemME As Long
Dim psDocTpo As String
Dim pnMonto As Currency

On Error GoTo ErrorGrabaAtencionArendirNew

oCon.AbreConexion
oCon.BeginTrans

pnMoneda = Mid(psOpeCod, 3, 1)
lnMovItem = 0: lnMovOrden = 0
lnMovItemME = 0
pnMonto = 0
'Inserta en Mov, MovCap, MovCapDet
If psOpeCod = "401125" Or psOpeCod = "402125" Or psOpeCod = "401225" Or psOpeCod = "402225" Then
    oDMov.InsertaMovAprobacionER_Viaticos pdFecha, psMovNro, psOpeCod, gITFCobroCargo, Mid(psOpeCod, 3, 1), psCuentaAho, pnImporte
End If
'****

lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"

lnSaldoArendir = GetSaldoARendir(pnMovRef)
lnSaldoArendir = lnSaldoArendir + pnImporte
oDMov.ActualizaMovArendir pnMovRef, pnTipoArendir, , , , lnSaldoArendir
oDMov.InsertaMovRef lnMovNro, pnMovRef

Set rsMovGenAs = ObtieneMovGeneraAsiento(lnMovNro, psCuentaAho)
If Not (rsMovGenAs.BOF And rsMovGenAs.EOF) Then
    Do While Not rsMovGenAs.EOF
        'Se obtiene Cta Contable Real
        Set rsOpeCta = GetCuentaContableReal(rsMovGenAs!cCtaCod, rsMovGenAs!nMovNro, rsMovGenAs!cOpeCod, rsMovGenAs!nConceptoCod, rsMovGenAs!nMonto, pdFecha)
        'guardamos la cuenta en MovCta
        Do While Not rsOpeCta.EOF
            If rsOpeCta!cTipo = "0" Then
                Set rsExisMovCta = ExistMovCta(rsMovGenAs!nMovNro, rsOpeCta!cCtaCont, "1") 'MovCta
                If Not rsExisMovCta.BOF And Not rsExisMovCta.EOF Then
                    lnMovItem = rsExisMovCta!nMovItem
                    pnMonto = rsExisMovCta!nMovImporte
                Else
                    lnMovItem = lnMovItem + 1
                End If
                If rsOpeCta!cOpeCtaDH = "D" Then
                    oDMov.InsertaMovCtaNew_Update rsOpeCta!nMovNro, lnMovItem, rsOpeCta!cCtaCont, pnMonto + rsOpeCta!nDebe
                Else
                    oDMov.InsertaMovCtaNew_Update rsOpeCta!nMovNro, lnMovItem, rsOpeCta!cCtaCont, pnMonto + (rsOpeCta!nHaber * -1)
                End If
            Else
                lnMovItemME = lnMovItemME + 1
                If rsOpeCta!cOpeCtaDH = "D" Then
                    oDMov.InsertaMovMeNew rsOpeCta!nMovNro, lnMovItemME, rsOpeCta!cCtaCont, rsOpeCta!nDebe
                Else
                    oDMov.InsertaMovMeNew rsOpeCta!nMovNro, lnMovItemME, rsOpeCta!cCtaCont, rsOpeCta!nHaber * -1
                End If
             End If
            pnMonto = 0
            rsOpeCta.MoveNext
        Loop
        rsMovGenAs.MoveNext
    Loop
End If

psDocTpo = CStr(TpoDocNotaAbono)
oDMov.InsertaRegDocCuenta psDocTpo, psDocNro, psCuentaAho

'If psDocNro <> "" Then
   'oDMov.InsertaMovDoc lnMovNro, psDocTpo, psDocNro, Format(CDate(pdFecha), gsFormatoFecha)
'End If

If pnTipoArendir = gArendirTipoViaticos Then
   GrabarReferenciaViat oDMov, pnMovRef, lnMovNro
End If
oDMov.InsertaMovGasto lnMovNro, psPersCod, "0"

oDMov.ActualizaSaldoMovimiento psMovNro, "+"

oCon.CommitTrans
oCon.CierraConexion
Set rsMovGenAs = Nothing
Set rsOpeCta = Nothing
Set rsExisMovCta = Nothing
GrabaAtencionArendirNew = 0
Exit Function
ErrorGrabaAtencionArendirNew:
    Set oCon = Nothing
    Err.Raise vbObjectError + 100, "GrabaAtencionArendir", TextErr(lsMsgErr)
End Function '**************NAGL ERS075-2017 20171127

Public Sub GrabarReferenciaViat(ByRef DMov As DMov, ByVal pnMovNroViat As Long, _
                                ByVal pnMovNroAtenc As Long)
Dim rs As ADODB.Recordset
Dim sql As String
'Dim oCon As DConecta
Set rs = New ADODB.Recordset
'Set oCon = New DConecta

sql = " SELECT  nViaticoMovNro, MV.nMovNro " _
    & " FROM    MovViaticos MV JOIN MOV M ON M.nMOVNRO=MV.nMOVNRO " _
    & "         LEFT JOIN MOVREF MR ON MR.nMOVNROREF= M.nMOVNRO " _
    & " WHERE   M.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagExtornado & "','" & gMovFlagModificado & "') " _
    & "         AND MV.nViaticoMovNro = '" & pnMovNroViat & "' AND MR.nMOVNROREF IS NULL"
    

Set rs = DMov.CargaRecordSet(sql)
Do While Not rs.EOF
    DMov.InsertaMovRef pnMovNroAtenc, rs!nMovNro
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

End Sub
Public Function GrabaAtencionArendirCH(ByVal pnTipoArendir As ArendirTipo, psFormatoFecha As String, _
                                        ByVal pnMovNroArendirSol As Long, _
                                        ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                        ByVal pnImporte As Currency, _
                                        ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnNroProcCH As Integer, _
                                        ByVal psCtaContDebe As String, ByVal psCtaContHaber As String, _
                                        ByVal psDocNro As String, ByVal pdDocFecha As Date, _
                                        Optional ByRef pnMovNroAte As Long = 0) As Integer
                                        '***Parametro pnMovNroAte agregado por ELRO el 20120614, según OYP-RFC047-2012
                                        
                                        
Dim oDMov As DMov
Dim lnItem As Integer
Dim lnOrden As Integer
Dim lsSubCta As String
Dim oCon As NContFunciones
Dim lnSaldoNuevo As Currency
Dim lnSaldo As Currency
Dim oCh As NCajaChica
Dim lnMovNro As Long
Dim lbTrans As Boolean
On Error GoTo GrabaAtencionArendirCHErr

Set oCon = New NContFunciones
Set oDMov = New DMov
Set oCh = New NCajaChica

oDMov.BeginTrans
lbTrans = True

oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
lnItem = lnItem + 1
oDMov.InsertaMovCta lnMovNro, lnItem, psCtaContDebe, pnImporte
lnItem = lnItem + 1
lsSubCta = oCon.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaContHaber, psAreaCh + psAgeCh, False)
'***Modificado por ELRO el 20120221, según SATI INC1301300007
'oDMov.InsertaMovCta lnMovNro, lnItem, psCtaContHaber + lsSubCta, pnImporte * -1
oDMov.InsertaMovCta lnMovNro, lnItem, psCtaContHaber, pnImporte * -1
'***Fin Modificado por ELRO el 20120221**********************
oDMov.InsertaMovObj lnMovNro, lnItem, 1, ObjCMACAgenciaArea
oDMov.InsertaMovObjAgenciaArea lnMovNro, lnItem, 1, psAgeCh, psAreaCh

oDMov.InsertaMovDoc lnMovNro, TpoDocRecArendirCuenta, psDocNro, Format(pdDocFecha, gsFormatoFecha)

oDMov.InsertaMovRef lnMovNro, pnMovNroArendirSol
oDMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, pnNroProcCH, gCHTipoProcArendir
oDMov.ActualizaMovArendir pnMovNroArendirSol, gArendirTipoCajaChica, , , , pnImporte

lnSaldo = oCh.GetDatosCajaChica(psAreaCh, psAgeCh, SaldoActual)
lnSaldoNuevo = lnSaldo - pnImporte
oDMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProcCH, , , , , lnSaldoNuevo
'***Agregado por ELRO el 20120614, según OYP-RFC047-2012
pnMovNroAte = lnMovNro
'***Agregado por ELRO***********************************
GrabaAtencionArendirCH = 0
oDMov.CommitTrans
lbTrans = False
Set oCon = Nothing
Set oDMov = Nothing
Set oCh = Nothing
Exit Function
GrabaAtencionArendirCHErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaAtencionArendirCH", "Error Graba Solicitud Arendir de Caja Chica"
                                        
End Function
Public Function GrabaSolicitudArendir(ByVal pnTipoArendir As ArendirTipo, _
                                        ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psAgeCod As String, ByVal psAreaCod As String, _
                                        ByVal psMovDesc As String, ByVal psPersCod As String, _
                                        ByVal pnImporte As Currency, ByVal psDocNro As String, ByVal psDocTpo As String, _
                                        ByVal psFechaDoc As String, ByVal psAreaCajaCh As String, _
                                        ByVal psAgeCajaCh As String, ByVal pnNroPro As Integer, _
                                        Optional ByRef pnMovNroSol As Long = 0) As Integer
                                        '***Agregado por ELRO el parametro pnMovNroSol el 20120615
Dim oDMov As DMov
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim sMsgErr As String

On Error GoTo GrabaSolicitudArendirErr

Set oDMov = New DMov

oDMov.BeginTrans
lbTrans = True
'Grabamos Mov
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
If pnTipoArendir = gArendirTipoCajaChica Then
    oDMov.InsertaMovARendir lnMovNro, pnTipoArendir, psAreaCajaCh, psAgeCajaCh, psPersCod, 0, ""
Else
    oDMov.InsertaMovARendir lnMovNro, pnTipoArendir, psAreaCod, psAgeCod, psPersCod, 0, ""
End If
oDMov.InsertaMovDoc lnMovNro, psDocTpo, psDocNro, psFechaDoc
If pnTipoArendir = gArendirTipoCajaChica Then
    oDMov.InsertaMovCajaChica lnMovNro, psAreaCajaCh, psAgeCajaCh, pnNroPro, gCHTipoProcArendir
End If
If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If
oDMov.CommitTrans
lbTrans = False
'***Agregado por ELRO el 20120615, según OYP-RFC047-2012
pnMovNroSol = lnMovNro
'***Fin Agregado por ELRO*******************************
GrabaSolicitudArendir = 0
Set oDMov = Nothing
Exit Function
GrabaSolicitudArendirErr:
    sMsgErr = Err.Description
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    Err.Raise vbObjectError + 100, "GrabaAtencionArendir", sMsgErr
End Function
'TORE - Se cargaron dos campos mas para compesar los valores de entregas a rendir (campos con datos no validos)
Public Function GetAtencionPendArendir(ByVal pbTodos As Boolean, ByVal psAgeCod As String, _
                                       ByVal psAreaCod As String, ByVal pnTipoArendir As ArendirTipo, _
                                       ByVal psCtaArendir As String, ByVal pnMoneda As Moneda, _
                                       ByVal psAreaArendir As String, ByVal psAgeArendir As String, Optional pbRendidos As Boolean = False) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsFiltroAge  As String
Dim lsFiltroArea  As String

Dim gcITF As String
Dim oCon As NConstSistemas
Set oCon = New NConstSistemas

Set oConect = New DConecta
Set rs = New Recordset

gcITF = oCon.LeeConstSistema(157)

If oConect.AbreConexion = False Then Exit Function
If psAgeCod <> "" Then
    lsFiltroAge = " and RH.cAgenciaActual='" & psAgeCod & "' "
End If
If psAreaCod <> "" Then
    lsFiltroArea = " and RH.cAreaCodActual='" & psAreaCod & "' "
End If



Select Case pnTipoArendir
    Case gArendirTipoCajaGeneral, gArendirTipoAgencias
        sql = "SELECT   ISNULL(D1.cDocAbrev,'EFEC') AS cAbreDocAtenc, ISNULL(MD.cDocNro,'Efectivo') as cNroDocAtenc, ISNULL(CONVERT(CHAR(10),MD.dDocFecha,103),SUBSTRING(M.CMOVNRO,7,2) + '/' + SUBSTRING(M.CMOVNRO,5,2)+ '/' + SUBSTRING(M.CMOVNRO,1,4)) AS dDocAtencion , " _
            & "         MDAR.cDocNro, CONVERT(CHAR(10),MDAR.dDocFecha,103) AS dDocSolicitud,CONVERT(CHAR(50),P.CPERSNOMBRE) AS cPersNombre, ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS NIMPORTE,  " _
            & "         M.cMovDesc , P.cPersCod, M.NMOVNRO as NMovNroAtenc, ISNULL(MD.nDocTpo,'') as cTpoDocAtenc, MAR.nMovSaldo ,  " _
            & "         ISNULL(A.cAreaDescripcion,'') cAreaDescripcion,ISNULL(A.cAreaCod,'') cAreaCod, ISNULL(D1.cDocDesc,'Atención Efectivo') as cDocDesc, MAR.NMOVNRO as NMovNroSol,   " _
            & "         cAgeCod = (SELECT ISNULL(cAgeCod,'') FROM AreaAgencia WHERE cAreaCod = aa.cAreaCod and cAgeCod = ag.cAgeCod), D2.cDocAbrev, MDAR.nDocTpo, cRHCargoCategoria = 1, cTpoArendir = 1 " _
            & "  FROM   MOV M " _
            & "         JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO " _
            & "         LEFT JOIN MOVME ME ON ME.NMOVNRO =MC.NMOVNRO AND MC.NMOVITEM = ME.NMOVITEM " _
            & "         LEFT JOIN MOVDOC MD ON MD.NMOVNRO = M.NMOVNRO  AND MD.nDocTpo NOT IN ('" & Format(TpoDocVoucherEgreso, "00") & "') " _
            & "         LEFT JOIN " & vsBaseComunes & "DOCUMENTO D1 ON D1.nDocTpo = MD.nDocTpo " _
            & "         JOIN MOVREF MR ON MR.NMOVNRO = M.NMOVNRO " _
            & "         JOIN MOVARENDIR MAR ON MAR.NMOVNRO=MR.NMOVNROREF " _
            & "         JOIN " & vsBasePesonas & "PERSONA P ON P.CPERSCOD  = MAR.CPERSCOD LEFT JOIN RRHH RH ON RH.CPERSCOD= P.CPERSCOD " _
            & "         LEFT JOIN " & vsBaseComunes & "AREAS A ON A.CAREACOD = RH.cAreaCodActual LEFT JOIN (SELECT Distinct cAreaCod FROM AreaAgencia ) AA ON AA.cAreaCod = A.cAreaCod " _
            & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG ON AG.CAGECOD = RH.cAgenciaActual " _
            & "         JOIN MOVDOC MDAR ON MDAR.NMOVNRO = MAR.NMOVNRO " _
            & "         JOIN " & vsBaseComunes & "DOCUMENTO D2 ON D2.nDocTpo = MDAR.nDocTpo " _
            & " WHERE   MAR.cTpoArendir ='" & pnTipoArendir & "' " & lsFiltroAge & lsFiltroArea _
            & "         AND MC.CCTACONTCOD IN ('" & psCtaArendir & "') AND Substring(M.cOpeCod,3,1)='" & pnMoneda & "' AND M.nMovEstado = " & gMovEstContabMovContable & " And M.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "') " _
            & IIf(pbRendidos, "", " AND NOT EXISTS (SELECT  M1.NMOVNRO FROM MOV AS M1 JOIN MOVREF MR1 ON M1.NMOVNRO=MR1.NMOVNRO         " _
            & "                         WHERE   MR1.NMOVNROREF=M.NMOVNRO AND SUBSTRING(M1.COPECOD,1,5) IN ('" & Mid(gCGArendirCtaRendMN, 1, 5) & "','" & Mid(gCGArendirCtaRendME, 1, 5) & "','" & Mid(gCGArendirCtaRendDevMN, 1, 5) & "','" & Mid(gCGArendirCtaRendDevME, 1, 5) & "','" & Mid(gCGArendirCtaRendDevMNAge, 1, 5) & "','" & Mid(gCGArendirCtaRendDevMEAge, 1, 5) & "','" & Mid(gCGArendirCtaRendMNAge, 1, 5) & "','" & Mid(gCGArendirCtaRendMEAge, 1, 5) & "') " _
            & "                                 AND M1.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "') ) ") _
            & "         AND MAR.cAreacod ='" & psAreaArendir & "' and MAR.cAgeCod='" & psAgeArendir & "' " _
            & " ORDER BY MDAR.cDocNro "

    Case gArendirTipoViaticos
        
        sql = " SELECT  D.CDOCABREV, MD.CDOCNRO AS DocNro, CONVERT(CHAR(12),MD.dDocFecha,103) AS dDocFecha , MD.CDOCNRO AS DocNro, CONVERT(CHAR(12),MD.dDocFecha,103) AS dDocFecha , " _
            & "         P.CPERSNOMBRE, ISNULL(ATEND.NATENDIDO,0) AS MONTOATENDIDO , " _
            & "         M.cMovDesc,  P.CPERSCOD, M.NMOVNRO, MD.nDocTpo, MAR.NMOVSALDO , " _
            & "         ISNULL(A.CAREADESCRIPCION,'') AS CAREADESCRIPCION, A.CAREACOD,  D.cDocDesc, " _
            & "         M.NMOVNRO, ISNULL(AG.CAGEDESCRIPCION,'') CAGEDESCRIPCION, cAgeCod = (SELECT ISNULL(cAgeCod,'') FROM AreaAgencia WHERE cAreaCod = aa.cAreaCod and cAgeCod = ag.cAgeCod)  , " _
            & "         P.CPERSNOMBRE, CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHA, D.CDOCABREV , MD.nDocTpo, cRHCargoCategoria = 1, cTpoArendir = 1 " _
            & " FROM    MOV M JOIN MOVARENDIR MAR ON M.NMOVNRO = MAR.NMOVNRO " _
            & "         JOIN PERSONA            P       ON P.CPERSCOD = MAR.CPERSCOD " _
            & "         LEFT JOIN RRHH          RH      ON RH.CPERSCOD = P.CPERSCOD " _
            & "         LEFT JOIN AREAS         A       ON A.CAREACOD = RH.cAreaCodActual LEFT JOIN (SELECT Distinct cAreaCod FROM AreaAgencia ) AA ON AA.cAreaCod = A.cAreaCod  " _
            & "         LEFT JOIN AGENCIAS  AG  ON AG.CAGECOD = RH.cAgenciaActual " _
            & "         JOIN MOVDOC         MD  ON MD.nMOVNRO = MAR.nMOVNRO AND nDocTpo IN (" & TpoDocRecViaticosARendirCuenta & ") " _
            & "         JOIN DOCUMENTO      D   ON D.nDocTpo = MD.nDocTpo " _
            & "         JOIN (  SELECT  SUM(ISNULL(ME.nMovMEImporte,MC.NMOVIMPORTE)) AS NATENDIDO, MR.NMOVNROREF " _
            & "                 FROM    MOVCTA MC JOIN MOV M1 ON M1.NMOVNRO = MC.NMOVNRO LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
            & "                         JOIN MOVREF MR ON MR.NMOVNRO  = M1.NMOVNRO " _
            & "                 WHERE   MC.NMOVIMPORTE>0 AND M1.NMOVFLAG =" & gMovFlagVigente & "  " _
            & "                         AND SUBSTRING(M1.COPECOD,1,5) IN ('" & Mid(gCGArendirViatAtencMN, 1, 5) & "','" & Mid(gCGArendirViatAtencME, 1, 5) & "') and MC.cCtacontcod not like ('" & gcITF & "') " _
            & "                 GROUP  BY MR.NMOVNROREF ) AS ATEND " _
            & "         ON ATEND.NMOVNROREF = M.NMOVNRO "
    sql = sql + " WHERE   MAR.cTpoArendir ='" & pnTipoArendir & "' AND SUBSTRING(M.COPECOD,3,1)='" & pnMoneda & "' " _
            & "         AND M.nMovFlag =" & gMovFlagVigente & " " & lsFiltroAge & lsFiltroArea _
            & "         AND MAR.cAreacod ='" & psAreaArendir & "' and MAR.cAgeCod='" & psAgeArendir & "' " _
            & IIf(pbRendidos, "", " AND NOT EXISTS (SELECT  M2.NMOVNRO " _
            & "                         FROM    MOV M2 " _
            & "                         JOIN    MOVREF MR ON MR.NMOVNRO = M2.NMOVNRO " _
            & "                         Where   MR.NMOVNROREF = MAR.NMOVNRO " _
            & "                                 AND M2.nMovEstado = " & gMovEstContabMovContable & " AND M2.nMovFlag = " & gMovFlagVigente & " " _
            & "                                 AND SUBSTRING(M2.COPECOD,1,5) IN ('" & Mid(gCGArendirViatRendMN, 1, 5) & "','" & Left(gCGArendirViatRendDevMN, 5) & "','" & Left(gCGArendirViatRendDevME, 5) & "','" & Mid(gCGArendirViatRendME, 1, 5) & "')) ") _
            & " ORDER BY MD.CDOCNRO "

End Select

Set rs = oConect.CargaRecordSet(sql)
Set GetAtencionPendArendir = rs
oConect.CierraConexion
Set oConect = Nothing
End Function
Public Function GetAtencionSinSustentacion(ByVal psAgeCod As String, ByVal psAreaCod As String, _
                                            ByVal pnTipoArendir As ArendirTipo, ByVal psCtaArendir As String, _
                                            ByVal psAreaArendir As String, ByVal psAgeArendir As String, _
                                            ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal pnMoneda As Moneda, _
                                            Optional ByVal pbTodo As Boolean = True, _
                                            Optional ByVal psOpeCod As String = "") As ADODB.Recordset
                                            '***Parametros pbTodo y psOpeCod agregado por ELRO el 20120410, según OYP-RFC005-2012
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsFiltroAge  As String
Dim lsFiltroArea  As String
Set oConect = New DConecta
Set rs = New Recordset

    If psAgeCod <> "" Then
        lsFiltroAge = " and RH.cAgenciaActual='" & psAgeCod & "' "
    End If
    If psAreaCod <> "" Then
        lsFiltroArea = " and RH.cAreaCodActual='" & psAreaCod & "' "
    End If

If oConect.AbreConexion = False Then Exit Function
Select Case pnTipoArendir
    Case gArendirTipoCajaGeneral, gArendirTipoAgencias
        If psOpeCod = CStr(gCGArendirCtaExtAprobMN) Or psOpeCod = CStr(gCGArendirCtaExtAprobME) Then
            sql = "exec stp_sel_ObtenerAtencionSinSustentacion_3  " & IIf(pbTodo, 1, 0) & ", '" & psAreaCod & "', '" & psAgeCod & "' , '" & psCtaArendir & "', '" & Format(pdDesde, "yyyymmdd") & "', '" & Format(pdHasta, "yyyymmdd") & "', '" & pnMoneda & "', " & gArendirTipoCajaGeneral & ""
        Else
        sql = "SELECT   ISNULL(D1.cDocAbrev,'EFEC') AS cAbreDocAtenc, ISNULL(MD.cDocNro,'Efectivo') as cNroDocAtenc, ISNULL(CONVERT(CHAR(10),MD.dDocFecha,103),SUBSTRING(M.CMOVNRO,7,2) + '/' + SUBSTRING(M.CMOVNRO,5,2)+ '/' + SUBSTRING(M.CMOVNRO,1,4)) AS dDocAtencion , " _
            & "         MDAR.cDocNro, CONVERT(CHAR(10),MDAR.dDocFecha,103) AS dDocSolicitud,CONVERT(CHAR(50),P.CPERSNOMBRE) AS cPersNombre, ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS NIMPORTE,  " _
            & "         M.cMovDesc , P.cPersCod, M.nMOVNRO as cMovNroAtenc, ISNULL(MD.nDocTpo,'') as cTpoDocAtenc, MAR.nMovSaldo ,  " _
            & "         ISNULL(A.cAreaDescripcion,'') cAreaDescripcion, ISNULL(A.cAreaCod,'') cAreaCod, ISNULL(D1.cDocDesc,'Atención Efectivo') as cDocDesc, MAR.nMOVNRO as cMovNroSol,   " _
            & "         ISNULL(AG.cAgeDescripcion,'') cAgeDescripcion , ISNULL(AG.cAgeCod,'') cAgeCod, D2.cDocAbrev, MDAR.nDocTpo " _
            & "  FROM   MOV M " _
            & "         JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO " _
            & "         LEFT JOIN MOVME ME ON ME.nMOVNRO =MC.nMOVNRO AND MC.nMOVITEM = ME.nMOVITEM " _
            & "         LEFT JOIN MOVDOC MD ON MD.nMOVNRO = M.nMOVNRO  AND MD.nDocTpo NOT IN ('" & Format(TpoDocVoucherEgreso, "00") & "') " _
            & "         LEFT JOIN " & vsBaseComunes & "DOCUMENTO D1 ON D1.nDocTpo = MD.nDocTpo " _
            & "         JOIN MOVREF MR ON MR.nMOVNRO = M.nMOVNRO " _
            & "         JOIN MOVARENDIR MAR ON MAR.nMOVNRO=MR.nMOVNROREF " _
            & "         JOIN " & vsBasePesonas & "PERSONA P ON P.CPERSCOD  = MAR.CPERSCOD " _
            & "         JOIN RRHH               RH          ON RH.CPERSCOD = MAR.CPERSCOD " _
            & "         JOIN " & vsBaseComunes & "AREAS A   ON A.CAREACOD = RH.cAreaCodActual " _
            & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG ON AG.CAGECOD = RH.cAgenciaActual " _
            & "         JOIN MOVDOC MDAR ON MDAR.nMOVNRO = MAR.nMOVNRO " _
            & "         JOIN " & vsBaseComunes & "DOCUMENTO D2 ON D2.nDocTpo = MDAR.nDocTpo "
        
        sql = sql + " WHERE MAR.cTpoArendir ='" & pnTipoArendir & "' " & lsFiltroAge & lsFiltroArea _
            & "             AND MC.CCTACONTCOD IN ('" & psCtaArendir & "') And SUBSTRING(M.COPECOD,3,1)='" & pnMoneda & "' " _
            & "             AND M.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "') " _
            & "             AND MAR.CAREACOD ='" & psAreaArendir & "' AND MAR.CAGECOD ='" & psAgeArendir & "' " _
            & "             AND NOT EXISTS (    SELECT  M1.CMOVNRO FROM    MOV M1 JOIN MOVREF MR1 ON M1.nMOVNRO=MR1.nMOVNRO " _
            & "                                 WHERE   MR1.nMOVNROREF = M.nMovNro " _
            & "                                         AND M1.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "') ) " _
            & "             AND  SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " _
            & " ORDER BY M.CMOVNRO, MDAR.cDocNro "
         End If
            
    Case gArendirTipoViaticos
'***Modificado por ELRO el 20120410, según OYP-RFC005-2012
'        sql = " SELECT  ISNULL(D1.cDocAbrev,'EFEC') AS cAbreDocAtenc, " _
'            & "         ISNULL(MD.cDocNro,'Efectivo') as cNroDocAtenc, " _
'            & "         ISNULL(CONVERT(CHAR(10),MD.dDocFecha,103), " _
'            & "         SUBSTRING(M.CMOVNRO,7,2) + '/' + SUBSTRING(M.CMOVNRO,5,2)+ '/' + SUBSTRING(M.CMOVNRO,1,4)) AS dDocAtencion , " _
'            & "         D2.cDocAbrev + '-' + MDAR.cDocNro AS cDocNro, CONVERT(CHAR(10),MDAR.dDocFecha,103) AS dDocSolicitud,CONVERT(CHAR(50),P.CPERSNOMBRE) AS cPersNombre, " _
'            & "         ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS NIMPORTE, " _
'            & "         M.cMovDesc , P.cPersCod, M.NMOVNRO as cMovNroAtenc, ISNULL(MD.nDocTpo,'') as cTpoDocAtenc, " _
'            & "         MAR.nMovSaldo ,           ISNULL(A.cAreaDescripcion,'') cAreaDescripcion,ISNULL(A.cAreaCod,'') cAreaCod, " _
'            & "         ISNULL(D1.cDocDesc,'Atención Efectivo') as cDocDesc, MAR.NMOVNRO as NMovNroSol, " _
'            & "         ISNULL(AG.cAgeDescripcion,'') cAgeDescripcion , ISNULL(AG.cAgeCod,'') cAgeCod, D2.cDocAbrev, MDAR.nDocTpo " _
'            & " FROM    MOV M " _
'            & "         JOIN MOVCTA MC                              ON MC.NMOVNRO = M.NMOVNRO " _
'            & "         LEFT JOIN MOVME ME                          ON ME.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM = ME.NMOVITEM " _
'            & "         LEFT JOIN MOVDOC MD                         ON MD.NMOVNRO = M.NMOVNRO  AND MD.nDocTpo NOT IN ('" & Format(TpoDocVoucherEgreso, "00") & "') " _
'            & "         LEFT JOIN DOCUMENTO D1                      ON D1.nDocTpo = MD.nDocTpo " _
'            & "         JOIN MOVREF MR                              ON MR.NMOVNRO = M.NMOVNRO " _
'            & "         JOIN MOVARENDIR MAR                         ON MAR.NMOVNRO= MR.NMOVNROREF " _
'            & "         JOIN " & vsBasePesonas & "PERSONA P         ON P.CPERSCOD = MAR.CPERSCOD " _
'            & "         JOIN RRHH               RH                  ON RH.CPERSCOD= MAR.CPERSCOD " _
'            & "         JOIN " & vsBaseComunes & "AREAS A           ON A.CAREACOD = RH.cAreaCodActual " _
'            & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG  ON AG.CAGECOD = RH.cAgenciaActual " _
'            & "         JOIN MOVDOC MDAR                            ON MDAR.NMOVNRO = MAR.NMOVNRO AND MDAR.nDocTpo IN ('" & Format(TpoDocRecViaticosARendirCuenta, "00") & "') " _
'            & "         JOIN DOCUMENTO D2                           ON D2.nDocTpo = MDAR.nDocTpo "
'sql = sql + " WHERE   MAR.cTpoArendir ='" & pnTipoArendir & "' And SUBSTRING(M.COPECOD,1,5) in ('" & Mid(gCGArendirViatAtencMN, 1, 5) & "','" & Mid(gCGArendirViatAtencME, 1, 5) & "') " & lsFiltroAge & lsFiltroArea _
'            & "         AND MC.CCTACONTCOD IN ('" & psCtaArendir & "') And SUBSTRING(M.COPECOD,3,1)='" & pnMoneda & "' " _
'            & "         AND M.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "') " _
'            & "         AND MAR.CAREACOD ='" & psAreaArendir & "' AND MAR.CAGECOD ='" & psAgeArendir & "' " _
'            & "         AND NOT EXISTS ( SELECT M1.CMOVNRO " _
'            & "                          FROM   MOV M1 JOIN MOVREF MR1 ON M1.NMOVNRO=MR1.NMOVNRO " _
'            & "                          Where  MR1.NMOVNROREF = MAR.NMovNro " _
'            & "                                 AND M1.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "') " _
'            & "                                 And SUBSTRING(M1.COPECOD,1,5) IN ('" & Left(gCGArendirViatSustMN, 5) & "','" & Left(gCGArendirViatSustME, 5) & "','" & Left(gCGArendirViatRendMN, 5) & "','" & Left(gCGArendirViatRendME, 5) & "') ) " _
'            & "         AND  SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " _
'            & " ORDER BY M.NMOVNRO, MDAR.cDocNro "
If psOpeCod = CStr(gCGArendirViatExtAprobMN) Or psOpeCod = CStr(gCGArendirViatExtAprobME) Then
    sql = "exec stp_sel_ObtenerAtencionSinSustentacion_2  " & IIf(pbTodo, 1, 0) & ", '" & psAreaCod & "', '" & psAgeCod & "' , '" & psCtaArendir & "', '" & Format(pdDesde, "yyyymmdd") & "', '" & Format(pdHasta, "yyyymmdd") & "', '" & pnMoneda & "', " & gArendirTipoViaticos & ""
Else
    sql = "exec stp_sel_ObtenerAtencionSinSustentacion  " & IIf(pbTodo, 1, 0) & ", '" & psAreaCod & "', '" & psAgeCod & "', '" & psCtaArendir & "', '" & psAreaArendir & "', '" & psAgeArendir & "', '" & Format(pdDesde, "yyyymmdd") & "', '" & Format(pdHasta, "yyyymmdd") & "', '" & pnMoneda & "' "
End If
'***Fin Modificado por ELRO***********************************
End Select
Set rs = oConect.CargaRecordSet(sql)
Set GetAtencionSinSustentacion = rs
oConect.CierraConexion
Set oConect = Nothing
End Function
Public Function GetDocSustentariosArendir(ByVal psMovNro As String, ByVal psCtaContArendir As String, ByVal psCtaContPendiente As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Set oConect = New DConecta
Set rs = New ADODB.Recordset

If oConect.AbreConexion = False Then Exit Function
'***Modificado por ELRO el 20120618, según OYP-RFC047-2012
'sql = " SELECT  M.cMovNro, M.cMovDesc, D.cDocAbrev, MD.cDocNro, MD.nDocTpo, p.cPersNombre,MD.dDocFecha, " _
'    & "         SUM(IsNull(ME.NMOVMEIMPORTE, MC.NMOVIMPORTE) * -1 ) As nDocImporte " _
'    & " FROM    MOVREF  MRSUST " _
'    & "         JOIN MOV M          ON MRSUST.nMOVNRO = M.nMOVNRO " _
'    & "         JOIN MOVCTA MC      ON MC.nMOVNRO = M.nMOVNRO " _
'    & "         LEFT JOIN MOVME ME  ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM =MC.nMOVITEM " _
'    & "         JOIN MOVDOC MD      ON MD.nMOVNRO = M.nMOVNRO and not md.nDocTpo in (" & TpoDocOrdenPago & "," & TpoDocCheque & "," & TpoDocNotaAbono & "," & TpoDocVoucherEgreso & ", " & TpoDoc.TpoDocCarta & " ) " _
'    & "         JOIN " & vsBaseComunes & "DOCUMENTO D    ON D.nDocTpo = MD.nDocTpo " _
'    & "         JOIN MovGasto MG  ON MG.nMOVNRO = M.nMOVNRO " _
'    & "         JOIN " & vsBasePesonas & "PERSONA P      ON P.CPERSCOD = MG.CPERSCOD " _
'    & " WHERE   MRSUST.nMOVNROREF =" & psMovNro & " AND cCtaContCod IN ('" & psCtaContArendir & "','" & psCtaContPendiente & "') " _
'    & "         And  M.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "') " _
'    & "         And M.nMovNro not in (select nMovNro from ColocRecupGastos Where  nMovNro = M.nMovNro and nColocRecGastoEstado in (2) ) " _
'    & " GROUP BY M.cMovNro, M.cMovDesc, D.cDocAbrev, MD.cDocNro, MD.nDocTpo, p.cPersNombre,MD.dDocFecha"
sql = "exec stp_sel_obtenerDocSustentariosArendir " & CLng(psMovNro) & ", '" & psCtaContArendir & "', '" & psCtaContPendiente & "'"
'***Fin Modificado por ELRO*******************************
Set rs = oConect.CargaRecordSet(sql)
Set GetDocSustentariosArendir = rs
oConect.CierraConexion: Set oConect = Nothing
End Function

'FUNCION QUE DEVUELVE EL SALDO DE UN ARENDIR
Public Function GetSaldoARendir(ByVal pnMovNro As Long) As Currency
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Set oConect = New DConecta

Set rs = New ADODB.Recordset
'cMovNro                   cTpoArendir cAgeAreaCodSol cPersCod      nMovSaldo
sql = " Select  nMovNro, nMovSaldo " _
    & " From    MovArendir Where nMovNro='" & pnMovNro & "' "
    
If oConect.AbreConexion = False Then Exit Function
GetSaldoARendir = 0
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetSaldoARendir = rs!nMovSaldo
End If
RSClose rs
oConect.CierraConexion: Set oConect = Nothing
End Function
Public Function EliminaMovDocSust(ByVal psMovNroSust As String, ByVal psMovNroSol As String, ByVal pnTipoArendir As ArendirTipo, ByVal pnImporte As Currency, ByVal pnSaldoArendir As Currency, ByVal psMovNroActualiza As String) As Integer
    On Error GoTo EliminaMovDocSustErr
    Dim oMov As DMov
    Dim sql As String
    Dim lnSaldoArendir As Currency
    Set oMov = New DMov
    EliminaMovDocSust = 1
    lnSaldoArendir = pnSaldoArendir
    lnSaldoArendir = lnSaldoArendir + pnImporte
    
    'oMov.EliminaMov psMovNroSust, psMovNroActualiza
    oMov.ActualizaEstadoMov oMov.GetnMovNro(psMovNroSust), gMovFlagEliminado
    oMov.InsertaMovModifica psMovNroActualiza, psMovNroSust

    oMov.ActualizaMovArendir psMovNroSol, pnTipoArendir, , , , lnSaldoArendir
    EliminaMovDocSust = 0
    Exit Function
EliminaMovDocSustErr:
    Call RaiseError(MyUnhandledError, "NARendir:EliminaMovDocSust Method")
End Function
'MADM 20110507
Public Sub EliminaGastoRecupera(ByVal psMovNroSust As String)
    Dim oMov As DMov
    Dim lsSql As String
    Dim oConect As DConecta
    Dim RnMovNro As Long
    Set oMov = New DMov
    Set oConect = New DConecta
   
    If oConect.AbreConexion = False Then Exit Sub
    RnMovNro = 0
    RnMovNro = oMov.GetnMovNro(psMovNroSust)
    lsSql = "stp_del_RecupGastoMovFinan " & RnMovNro
    oConect.Ejecutar lsSql
    
    oConect.CierraConexion: Set oConect = Nothing
End Sub
'END MADM

Public Function GrabaRechazoSolARendir(ByVal psMovNroRechazo As String, ByVal psOpeCod As String, ByVal psMovNroAten As Long, _
                                        ByVal psGlosa As String) As Integer
    On Error GoTo GrabaRechazoSolARendirErr
    Dim oMov As DMov
    Dim sql As String
    Dim lnMovNro As Long
    Dim lbTrans As Boolean
    Dim lnSaldoArendir As Currency
    Set oMov = New DMov
    GrabaRechazoSolARendir = 1
    
    oMov.BeginTrans
    lbTrans = True
    oMov.InsertaMov psMovNroRechazo, psOpeCod, psGlosa, gMovEstContabNoContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(psMovNroRechazo)
    'oMov.InsertaMovCont lnMovNro, 0, "0", "0"
    oMov.ActualizaMov psMovNroAten, , gMovEstContabRechazado, gMovFlagEliminado
    oMov.InsertaMovRef lnMovNro, psMovNroAten
    
    GrabaRechazoSolARendir = 0
    oMov.CommitTrans
    lbTrans = False
    Exit Function
GrabaRechazoSolARendirErr:
    If lbTrans Then oMov.RollbackTrans
    Call RaiseError(MyUnhandledError, "NARendir:GrabaRechazoSolARendir Method")
End Function


Private Sub GetArendirRendido(ByVal pnMovNro As Long, ByVal psOpeCod As String, ByVal psCtaArendir As String, ByVal psCtaArendirMas As String, ByRef pnRendArendir As Currency, ByRef pnRendArendirMas As Currency)
On Error GoTo GetArendirRendidoErr
Dim sSQL As String
Dim prs  As Recordset
Dim oConect As DConecta
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Sub
 sSQL = "SELECT f.cCtaContCod, SUM(ISNULL(Mme.nMovMeImporte, f.nMovImporte))*-1 as nImporte " _
      & "FROM  MovRef a JOIN Mov m ON m.nMovNro = a.nMovNro " _
      & "               JOIN MovCta f ON f.nMovNro = m.nMovNro " _
      & "          LEFT JOIN MovMe Mme ON Mme.nMovNro = f.nMovNro and Mme.nMovItem = f.nMovItem " _
      & "               JOIN MovDoc b ON b.nMovNro = a.nMovNro JOIN OpeDoc od ON od.nDocTpo = b.nDocTpo " _
      & "               JOIN Documento e ON e.nDocTpo = b.nDocTpo " _
      & "WHERE a.nMovNroRef = '" & pnMovNro & "' and m.nMovEstado IN (" & gMovEstContabMovContable & "," & gMovEstContabNoContable & "," & gMovEstContabSustPendRendir & ") and m.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') and " _
      & "      f.cCtaContCod IN ('" & psCtaArendir & "','" & psCtaArendirMas & "') and od.cOpeCod IN ('" & Format(Val(Mid(psOpeCod, 1, 4)) - 1, "0000") & "00', '" & psOpeCod & "') and od.cOpeDocMetodo = " & OpeDocMetDigitado _
      & "GROUP BY f.cCtaContCod "
      
   Set prs = oConect.CargaRecordSet(sSQL)
   pnRendArendir = 0: pnRendArendirMas = 0
   If Not prs.EOF Then
      Do While Not prs.EOF
         If prs!cCtaContCod = psCtaArendir Then
            pnRendArendir = pnRendArendir + prs!nImporte
         End If
         If prs!cCtaContCod = psCtaArendirMas Then
            pnRendArendirMas = pnRendArendirMas + prs!nImporte
         End If
         prs.MoveNext
      Loop
   Else
      pnRendArendir = 0
      pnRendArendirMas = 0
   End If
   RSClose prs
   Set oConect = Nothing
Exit Sub
GetArendirRendidoErr:
   Call Err.Raise(Err.Number, "NArendir: GetArendirRendido Method", Err.Description)
End Sub

Public Sub GrabaSustentacionArendir(ByVal pbNewProv As Boolean, ByVal psPersCod As String, ByVal psProvEstado As String, ByVal psUltiActProv As String, _
       ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, ByVal pnMonto As Currency, ByVal pnDestino As Integer, _
       ByVal psCtaProvision As String, ByVal psCtaPendiente As String, ByVal pnTipoArendir As ArendirTipo, pnTotArendir As Currency, _
       ByVal psDocNro As String, ByVal pnDocTpo As TpoDoc, ByVal pdDocFecha As Date, ByVal pnTotal As Currency, ByVal pnMovNroAtencion As Long, _
       ByVal pnMovNroSol As Long, ByVal psAgeCod As String, ByVal psAreaCod As String, ByVal rsCtasCont As ADODB.Recordset, ByVal rsObjetos As ADODB.Recordset, _
       ByVal rsImp As ADODB.Recordset, Optional ByVal psAreaCh As String = "", Optional ByVal psAgeCh As String = "", _
       Optional ByVal pnNroProc As Integer = 0, Optional pnTCVariable As Currency = 0, _
       Optional ByVal lsCtaITFD As String, Optional ByVal lsCtaITFH As String, Optional ByVal nTasaITF As Double = 0, _
       Optional pnTCVPonderado As Currency = 0, Optional pbDocDolares As Boolean = 0, Optional psCompraVentaIng As String, Optional psCompraVetaEgr As String, Optional ByRef psCtasInexis As String = "", _
       Optional ByVal psMatrizViaticos As Variant = "", Optional ByVal pnNumAgesViaticos As Integer = 0, Optional ByRef pnMovNro As Long = 0, _
       Optional ByVal psLugarMovilidad As String = "", Optional ByVal pnRetencSistPensProvMontoBase As Currency = 0#, Optional ByVal pnRetencSistPensProvTotal As Currency = 0#, _
       Optional ByVal pnTpoProvSistPens As Integer = 0, Optional ByVal poMatDatosRetencSistPensProv As Variant = "")      'MADM 20110506 pnMovNro
       '***Parametro psLugarMovilidad agregado por ELRO el 20130111, según OYP-RFC126-2012
       'pnRetencSistPensProvMontoBase,pnRetencSistPensProvTotal, poMatDatosRetencSistPensProv EJVG20140729
    Dim sql As String
    Dim oMov As DMov
    Dim lsSubCta As String
    Dim lnItem As Integer
    Dim lnOrdenObj As Integer
    Dim lsAgeCod As String
    Dim lsAreaCod As String
    Dim lnImporte As Currency
    Dim lnImportePend As Currency
    Dim lbTrans As Boolean
    Dim lnMovNro As Long
    Dim i As Integer
    Dim lnMontoNoGrabado As Currency
    Dim lcCtaCont As String
    Dim RsBuscaCta As ADODB.Recordset
    
   On Error GoTo GrabaSustentacionArendirErr

   Dim nRendArendir     As Currency
   Dim nRendArendirMas  As Currency
   Dim nRendArendirImp     As Currency
   Dim nRendArendirImpMas  As Currency
   Dim nFaltaArendir    As Currency
   Dim lnItemDistri As Integer
    Dim lcCodObj As String
    Dim rsAgesDistrib As ADODB.Recordset
    Dim lsSql As String
    Dim lnMonto As Currency
    Dim lnRegImporte As Currency
    Dim rsBuscaCuenta As ADODB.Recordset
    Dim lsCtasInexis As String
    Dim lnItemPrin As Integer
    Dim lnImpoPrin As Currency
    Dim lnMontoCtaOrd As Currency, sTexto As String, rsCtasOrden As ADODB.Recordset
    
    lsCtasInexis = ""
    
   If pbDocDolares Then
      pnTotal = Format(pnTotal * pnTCVariable, "#.00")
      pnMonto = Format(pnMonto * pnTCVPonderado, "#.00")
   End If

   GetArendirRendido pnMovNroAtencion, psOpeCod, psCtaProvision, psCtaPendiente, nRendArendir, nRendArendirMas
   nFaltaArendir = 0
   nRendArendirImpMas = 0
   If pnTotArendir - nRendArendir > 0 Then
      nFaltaArendir = pnTotArendir - nRendArendir
   End If
   If nFaltaArendir <= 0 Then
      nRendArendirImp = 0
      nRendArendirImpMas = pnTotal
   Else
      If nFaltaArendir - nRendArendirMas > pnTotal Then
         nRendArendirImp = pnTotal + nRendArendirMas
         nRendArendirImpMas = nRendArendirMas * -1
      Else
         nRendArendirImp = nFaltaArendir
         nRendArendirImpMas = pnTotal - nFaltaArendir
      End If
   End If
    
    Set oMov = New DMov
    
    oMov.BeginTrans
    
    lbTrans = True
'    If pbNewProv Then
'        oMov.InsertProv psPersCod, psProvEstado, psUltiActProv
'    End If
    oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabSustPendRendir, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(psMovNro)
    oMov.InsertaMovCont lnMovNro, pnMonto, "0", "0"
    oMov.InsertaMovGasto lnMovNro, psPersCod, ""
    'Grabamos en MovCta, MovObj y MovCant
     
     pnMovNro = lnMovNro 'MADM 20110506
    
    lnItem = 0
    lnItemDistri = 0
    
    If Not rsCtasCont Is Nothing Then
        Do While Not rsCtasCont.EOF
            lsSubCta = ""
            lnItem = IIf(lnItemDistri > 0, lnItemDistri + 1, IIf(lnItem > 0, lnItem + 1, IIf(lnItem = 0, lnItem + 1, lnItem)))
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
                If pbDocDolares Then
                   lnImporte = Format(lnImporte * pnTCVPonderado, "#.00")
                End If
                
                For i = 6 To rsCtasCont.Fields.Count - 1
                    rsImp.Move i - 6, adBookmarkFirst
                    If rsImp.Fields(10) = "0" Then
                        If pbDocDolares Then
                            lnImporte = lnImporte + Format(rsCtasCont.Fields(i).value * pnTCVPonderado, "#.00")
                        Else
                            lnImporte = lnImporte + rsCtasCont.Fields(i).value
                        End If
                    End If
                    If rsCtasCont.Fields(i).value > 0 Then
                        If pbDocDolares Then
                            oMov.InsertaMovOtrosItem lnMovNro, lnItem, rsImp!CtaCont, Format(rsCtasCont.Fields(i).value * pnTCVPonderado, "#.00"), 3
                        Else
                            oMov.InsertaMovOtrosItem lnMovNro, lnItem, rsImp!CtaCont, rsCtasCont.Fields(i).value, 3
                        End If
                    End If
                Next

                '*** PEAC 20101019
                    'Call Distribuye(pbDocDolares, psOpeCod, psAreaCod, IIf(psAgeCod = "", "01", psAgeCod), lnMovNro, lnItem, pnDocTpo, lnImporte, rsCtasCont, rsObjetos, lnItemDistri, lsCtasInexis)
                    Call Distribuye(pbDocDolares, psOpeCod, psAreaCod, lsSubCta, lnMovNro, lnItem, pnDocTpo, lnImporte, rsCtasCont, rsObjetos, lnItemDistri, lsCtasInexis, psMatrizViaticos, pnNumAgesViaticos)

'                If pbDocDolares Then
'                    oMov.InsertaMovCta lnMovNro, lnItem, Left(rsCtasCont!Código, 2) & "2" & Trim(Mid(rsCtasCont!Código, 4, 50)) & lsSubCta, IIf(rsCtasCont!DH = "D", lnImporte, lnImporte * -1) * IIf(pnDocTpo = 7, -1, 1)
'                Else
'                    oMov.InsertaMovCta lnMovNro, lnItem, rsCtasCont!Código + lsSubCta, IIf(rsCtasCont!DH = "D", lnImporte, lnImporte * -1) * IIf(pnDocTpo = 7, -1, 1)
'                End If
                '*** FIN PEAC
                
            Else
                lnImporte = rsCtasCont!Monto
                
                If pbDocDolares Then
                   lnImporte = Format(lnImporte * pnTCVPonderado, "#.00")
                End If
                
                For i = 6 To rsCtasCont.Fields.Count - 1
                    rsImp.Move i - 6, adBookmarkFirst
                    If rsCtasCont.Fields(i) > 0 Then
                        If rsCtasCont.Fields(i).Name = "AJUSTE" Then
                            If pbDocDolares Then
                                lnImporte = lnImporte + Format(rsCtasCont.Fields(i) * pnTCVPonderado, "#.00")
                                oMov.InsertaMovOtrosItem lnMovNro, lnItem, rsImp!CtaCont, Format(rsCtasCont.Fields(i) * pnTCVPonderado, "#.00"), 1
                            Else
                                lnImporte = lnImporte + rsCtasCont.Fields(i)
                                oMov.InsertaMovOtrosItem lnMovNro, lnItem, rsImp!CtaCont, rsCtasCont.Fields(i), 1
                            End If
                            'oMov.InsertaMovCta lnMovNro, lnItem, rsCtasCont!Código + lsSubCta, IIf(rsCtasCont!DH = "D", lnImporte, lnImporte * -1)
                        Else
                            'oMov.InsertaMovCta lnMovNro, lnItem, rsImp!CtaCont, IIf(rsImp!cDocImpDH = "D", lnImporte, lnImporte * -1)
                            If pbDocDolares Then
                                oMov.InsertaMovOtrosItem lnMovNro, lnItem, rsImp!CtaCont, Format(rsCtasCont.Fields(i) * pnTCVPonderado, "#.00"), 1
                            Else
                                oMov.InsertaMovOtrosItem lnMovNro, lnItem, rsImp!CtaCont, rsCtasCont.Fields(i), 1
                            End If
                        End If
                    End If
                Next
                
                '*** PEAC 20101019
                '*** PEAC 20101209
                'Call Distribuye(pbDocDolares, psOpeCod, psAreaCod, IIf(psAgeCod = "", "01", psAgeCod), lnMovNro, lnItem, pnDocTpo, lnImporte, rsCtasCont, rsObjetos, lnItemDistri, lsCtasInexis)
                Call Distribuye(pbDocDolares, psOpeCod, psAreaCod, lsSubCta, lnMovNro, lnItem, pnDocTpo, lnImporte, rsCtasCont, rsObjetos, lnItemDistri, lsCtasInexis, psMatrizViaticos, pnNumAgesViaticos)
                
'                If pbDocDolares Then
'                     oMov.InsertaMovCta lnMovNro, lnItem, Left(rsCtasCont!Código, 2) & "2" & Trim(Mid(rsCtasCont!Código, 4, 50)) + lsSubCta, IIf(rsCtasCont!DH = "D", lnImporte, lnImporte * -1) * IIf(pnDocTpo = 7, -1, 1)
'                Else
'                     oMov.InsertaMovCta lnMovNro, lnItem, rsCtasCont!Código + lsSubCta, IIf(rsCtasCont!DH = "D", lnImporte, lnImporte * -1) * IIf(pnDocTpo = 7, -1, 1)
'                End If
                '*** FIN PEAC
                
            End If
            
            'lnImporte = rsCtasCont!Monto
            'For I = 6 To rsCtasCont.Fields.Count - 1
            '    If (pnDestino = 1 Or pnDestino = 2 Or rsCtasCont.Fields(I).Name = "AJUSTE") Then
            '        lnImporte = lnImporte + rsCtasCont.Fields(I).value
            '    End If
            'Next
            'oMov.InsertaMovCta lnMovNro, lnItem, rsCtasCont!Código + lsSubCta, IIf(rsCtasCont!DH = "D", lnImporte, lnImporte * -1)

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

                                'INSERTAMOS DENTRO DE LA TABLA MOVCTAIF PARA EL CONTROL OPERATIVO DE LAS CUENTAS DE BANCOS
'                                oMov.InsertaMovCtaIF lnMovNro, Mid(rsObjetos!Código, 4, 13), Mid(rsObjetos!Código, 1, 2), Mid(rsObjetos!Código, 18, 10), CGCtaIFConCapital, IIf(rsCtasCont!DH = "D", lnImporte, lnImporte * -1)
'                                oMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
'                                    Mid(rsObjetos!Código, 4, 13), Mid(rsObjetos!Código, 1, 2), Mid(rsObjetos!Código, 18, 10)

                            Case Else
                                lnOrdenObj = lnOrdenObj + 1
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código
                        End Select
                    End If
                    rsObjetos.MoveNext
                Loop
            End If
'            For I = 6 To rsCtasCont.Fields.Count - 1
'                rsImp.MoveFirst
'                If rsCtasCont.Fields(I).value <> 0 Then
'                    rsImp.Find "Impuesto='" & rsCtasCont.Fields(I).Name & "'"
'                    If Not rsImp.EOF Then
'                        oMov.InsertaMovOtrosItem lnMovNro, lnItem, rsImp!CTACONT, rsCtasCont.Fields(I).value, IIf(pnDestino = -1, "", pnDestino + 1)
'                    End If
'                End If
'            Next
            rsCtasCont.MoveNext
        Loop
    End If
'    If Not rsImp Is Nothing Then
'        rsImp.MoveFirst
'        Do While Not rsImp.EOF
'            If rsImp!Ok = "1" And (pnDestino = -1 Or pnDestino = 0 Or pnDestino = 1) Then
'                lnImporte = rsImp!Monto * IIf(rsImp!cDocImpDH = "D", 1, -1)
'                If lnImporte <> 0 And rsImp!Impuesto <> "AJUSTE" Then
'                    lnItem = lnItem + 1
'                    oMov.InsertaMovCta lnMovNro, lnItem, rsImp!CTACONT, lnImporte
'                End If
'            End If
'            rsImp.MoveNext
'        Loop
'    End If
If Not rsImp Is Nothing Then
    rsImp.MoveFirst
    Do While Not rsImp.EOF
        If rsImp!Ok = "1" And (pnDestino = -1 Or pnDestino = 0) Then
            If pbDocDolares Then
                lnImporte = Format(rsImp!Monto * pnTCVPonderado, "#.00") * IIf(rsImp!cDocImpDH = "D", 1, -1)
            Else
                lnImporte = rsImp!Monto * IIf(rsImp!cDocImpDH = "D", 1, -1)
            End If
            
            If lnImporte <> 0 And rsImp!Impuesto <> "AJUSTE" Then
                'lnItem = lnItem + 1
                lnItem = IIf(lnItemDistri > 0, lnItemDistri, lnItem) + 1
                lnItemDistri = lnItem
                oMov.InsertaMovCta lnMovNro, lnItem, rsImp!CtaCont, lnImporte * IIf(pnDocTpo = 7, -1, 1)
            End If
        ElseIf rsImp!Ok = "1" And (pnDestino = 1) Then
            If pbDocDolares Then
                lnImporte = Format(rsImp!Monto * pnTCVPonderado, "#.00") * IIf(rsImp!cDocImpDH = "D", 1, -1)
            Else
                lnImporte = rsImp!Monto * IIf(rsImp!cDocImpDH = "D", 1, -1)
            End If
            
            lnMontoNoGrabado = 0
            rsCtasCont.MoveFirst
            While Not rsCtasCont.EOF
                If rsCtasCont.Fields(4) = 1 Then
                    If pbDocDolares Then
                        lnMontoNoGrabado = lnMontoNoGrabado + (Format(rsCtasCont.Fields(rsImp.Bookmark + 5) * pnTCVPonderado, "#.00"))
                    Else
                        lnMontoNoGrabado = lnMontoNoGrabado + rsCtasCont.Fields(rsImp.Bookmark + 5)
                    End If
                End If
                rsCtasCont.MoveNext
            Wend
            
            If rsImp.Fields(10) = 0 Then
                lnImporte = lnImporte - lnMontoNoGrabado
            End If
            
            If lnImporte <> 0 And rsImp!Impuesto <> "AJUSTE" Then
                'lnItem = lnItem + 1
                lnItem = IIf(lnItemDistri > 0, lnItemDistri, lnItem) + 1
                lnItemDistri = lnItem
                oMov.InsertaMovCta lnMovNro, lnItem, rsImp!CtaCont, lnImporte * IIf(pnDocTpo = 7, -1, 1)
            End If
        ElseIf rsImp!Ok = "1" And rsImp.Fields(10) = 1 Then
            If pbDocDolares Then
                lnImporte = Format(rsImp!Monto * pnTCVPonderado, "#.00") * IIf(rsImp!cDocImpDH = "D", 1, -1)
            Else
                lnImporte = rsImp!Monto * IIf(rsImp!cDocImpDH = "D", 1, -1)
            End If
            
            If lnImporte <> 0 And rsImp!Impuesto <> "AJUSTE" Then
                'lnItem = lnItem + 1
                lnItem = IIf(lnItemDistri > 0, lnItemDistri, lnItem) + 1
                lnItemDistri = lnItem
                oMov.InsertaMovCta lnMovNro, lnItem, rsImp!CtaCont, lnImporte * IIf(pnDocTpo = 7, -1, 1)
            End If
        End If
        rsImp.MoveNext
    Loop
End If
    
    'Cuenta de Provisión / Caja Chica / A rendir Cuenta
    If nRendArendirImp <> 0 Then
    
        '*** PEAC 20101019
        '***Modificado por ELRO el 20120921, según OYP-RFC111-2012
        'If (psOpeCod = "401140" Or psOpeCod = "402140" Or psOpeCod = "401240" Or psOpeCod = "402240") And Right(lsSubCta, 2) = "01" And lnItemDistri > 0 Then
        If (psOpeCod = "401140" Or psOpeCod = "402140" Or psOpeCod = "401141" Or psOpeCod = "402141" Or psOpeCod = "401240" Or psOpeCod = "402240" Or psOpeCod = "401241" Or psOpeCod = "402241" Or psOpeCod = "401344") And Right(lsSubCta, 2) = "01" And lnItemDistri > 0 Then
        '***Fin Modificado por ELRO el 20120921*******************
            lnItem = lnItemDistri + 1
        Else
            lnItem = lnItem + 1
        End If
        '    lnItem = lnItem + 1
        '*** FIN PEAC
        
        'EJVG20140729 ***
        If pnRetencSistPensProvMontoBase > 0 Then 'Entra al calculo acumulativo
            oMov.InsertaMovProvRetencSistPens lnMovNro, pnTpoProvSistPens, pnRetencSistPensProvMontoBase, pnRetencSistPensProvTotal
        End If
        If IsArray(poMatDatosRetencSistPensProv) Then
            If poMatDatosRetencSistPensProv(1) > 0 Then
                oMov.InsertaMovCta lnMovNro, lnItem, poMatDatosRetencSistPensProv(2), poMatDatosRetencSistPensProv(1) * -1
                oMov.InsertaMovProvRetencSistPensDet lnMovNro, lnItem, ConceptoRetencionSistemaPensionProveedor.Aporte, poMatDatosRetencSistPensProv(1), poMatDatosRetencSistPensProv(1)
                lnItem = lnItem + 1
            End If
            If poMatDatosRetencSistPensProv(3) > 0 Then
                oMov.InsertaMovCta lnMovNro, lnItem, poMatDatosRetencSistPensProv(4), poMatDatosRetencSistPensProv(3) * -1
                oMov.InsertaMovProvRetencSistPensDet lnMovNro, lnItem, ConceptoRetencionSistemaPensionProveedor.SeguroAFP, poMatDatosRetencSistPensProv(3), poMatDatosRetencSistPensProv(3)
                lnItem = lnItem + 1
            End If
            If poMatDatosRetencSistPensProv(5) > 0 Then
                oMov.InsertaMovCta lnMovNro, lnItem, poMatDatosRetencSistPensProv(6), poMatDatosRetencSistPensProv(5) * -1
                oMov.InsertaMovProvRetencSistPensDet lnMovNro, lnItem, ConceptoRetencionSistemaPensionProveedor.ComsionAFP, poMatDatosRetencSistPensProv(5), poMatDatosRetencSistPensProv(5)
                lnItem = lnItem + 1
            End If
        End If
        'END EJVG *******
       oMov.InsertaMovCta lnMovNro, lnItem, psCtaProvision, nRendArendirImp * -1 * IIf(pnDocTpo = 7, -1, 1)
    End If
    If nRendArendirImpMas <> 0 Then
       lnItem = lnItem + 1
       lnItemDistri = lnItem
       '*** PEAC 20110104 - BUSCA SI ESTA EN ULTIMO NIVEL, SI NO LO ASIGNA A LA PRINCIPAL, a fin de no generar error en la clase
       Set RsBuscaCta = oMov.CargaRecordSet("Select cCtaContCod From CtaCont where cCtaContCod Like '" & psCtaPendiente & "%' And cCtaContCod <> '" & psCtaPendiente & "' And nCtaEstado = 1 ")
       If Not (RsBuscaCta.EOF And RsBuscaCta.BOF) Then
            psCtaPendiente = psCtaPendiente + Right(lsSubCta, 2)
       End If
       '*** FIN PEAC
       
       oMov.InsertaMovCta lnMovNro, lnItem, psCtaPendiente, nRendArendirImpMas * -1 * IIf(pnDocTpo = 7, -1, 1)
    End If
    ''Grabamos Documento en MovDoc
    '***Modificado por ELRO el 20130111, según OYP-RFC126-2012
    'oMov.InsertaMovDoc lnMovNro, pnDocTpo, psDocNro, Format(pdDocFecha, gsFormatoFecha)
    If pnDocTpo = 97 Then
        oMov.InsertaMovDoc lnMovNro, pnDocTpo, psDocNro, Format(pdDocFecha, gsFormatoFecha), psLugarMovilidad
    Else
        oMov.InsertaMovDoc lnMovNro, pnDocTpo, psDocNro, Format(pdDocFecha, gsFormatoFecha)
    End If
    '***Fin Modificado por ELRO el 20130111*******************
    'Grabamos la Referencia al Movimiento que se regulariza
    oMov.InsertaMovRef lnMovNro, pnMovNroAtencion
    If pnTipoArendir = gArendirTipoCajaChica Then
       oMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, pnNroProc, gCHTipoProcArendir
    End If
    
    'Actualizamos el saldo del aRendir
    oMov.ActualizaMovArendir pnMovNroSol, pnTipoArendir, , , , pnTotArendir - (nRendArendir + nRendArendirMas + (nRendArendirImp * IIf(pnDocTpo = 7, -1, 1)) + nRendArendirImpMas)
    oMov.InsertaMovArendirSust pnMovNroSol, lnMovNro

'--Calculo del ITF
'Dim lnMontoITF As Currency
Dim lnMontoITF As Double '*** PEAC 20110331


'If nTasaITF <> 0 And Not pnDocTpo = TpoDocRecEgreso Then
   'lnMontoITF = Round((nRendArendir + nRendArendirMas + nRendArendirImp + nRendArendirImpMas) * nTasaITF, 2)

'BLOQUEO POR SOLICITUD CUSCO
'If nTasaITF <> 0 Then
''   lnMontoITF = Round((nRendArendirImp) * nTasaITF, 2)
'    lnMontoITF = Round((nRendArendirImp + nRendArendirImpMas) * nTasaITF, 2)
'
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
    
    If Mid(psOpeCod, 3, 1) = gMonedaNacional And pbDocDolares Then
        oMov.InsertaMovAjusteCV_Soles_Doc_MonedaExt lnMovNro, Format(lnItem + 1, "#0"), psCompraVentaIng, psCompraVetaEgr
    End If
    'Sustentación de Arendir no pasa a Contabilidad
'    oMov.ActualizaSaldoMovimiento psMovNro, "+"

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
        lnItem = lnItem + 1
        oMov.InsertaMovCta lnMovNro, lnItem, rsCtasOrden!cCtaDebe, lnMontoCtaOrd
        lnItem = lnItem + 1
        oMov.InsertaMovCta lnMovNro, lnItem, rsCtasOrden!cCtaHaber, lnMontoCtaOrd * -1
    End If
End If

''*** FIN PEAC



    oMov.CommitTrans

    If lsCtasInexis <> "" Then
        psCtasInexis = lsCtasInexis
    End If
    
    lbTrans = False
    Set oMov = Nothing
Exit Sub
GrabaSustentacionArendirErr:
    If lbTrans Then oMov.RollbackTrans
    With Err ' Pass the error to the client
        .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
    End With

End Sub
Public Function GetAsientosSustARendir(ByVal lsMovNroArendirAtencion As String, ByVal psCtaContArendir As String, Optional pnTipoArendir As ArendirTipo = gArendirTipoCajaGeneral) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set rs = New ADODB.Recordset
Set oConect = New DConecta

Select Case pnTipoArendir
    Case gArendirTipoCajaGeneral, gArendirTipoAgencias
        sql = "SELECT   M.CMOVNRO AS CMOVSOL , MAT.CMOVNRO AS CMOVATENC," _
        & "             MAR.CPERSCOD , P.CPERSNOMBRE, ISNULL(A.CAREACOD,'') CAREACOD, ISNULL(A.CAREADESCRIPCION,'') CAREADESCRIPCION, " _
        & "             ISNULL(AG.CAGECOD,'') AS CAGECOD  , ISNULL(AG.CAGEDESCRIPCION,'') CAGEDESCRIPCION , ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS ImporteArendir,  " _
        & "             MAR.CAREACOD as cAreaCodArendir, A1.CAREADESCRIPCION as cAreaARendir , " _
        & "             ISNULL(MAR.CAGECOD,'') AS cAgeCodArendir , ISNULL(AG1.CAGEDESCRIPCION,'') cAgeArendir , " _
        & "             MAR.NMOVSALDO, MD.nDocTpo , D.CDOCABREV , MD.CDOCNRO, MD.dDocFecha as dDocFechaSol , " _
        & "             ISNULL(MD1.nDocTpo,'') AS nDocTpoATENC, ISNULL(D1.CDOCABREV,'') AS CDOCABREVATENC, ISNULL(MD1.CDOCNRO,'') AS CDOCNROATENC, ISNULL(D1.cDocDesc,'')  AS cDocDescAtenc, MD1.dDocFecha as dFechaDocAtenc, " _
        & "             SUST.cCtaContCod as CtaContSust, IsNull(ABS(SUST.Debe),0) AS Debe , IsNull(ABS(SUST.HABER),0) HABER " _
        & "     FROM    MOV M JOIN MOVREF MR ON M.nMOVNRO = MR.nMOVNROREF " _
        & "             JOIN MOVARENDIR MAR ON MAR.nMOVNRO = M.nMOVNRO " _
        & "             JOIN " & vsBasePesonas & "PERSONA P ON P.CPERSCOD = MAR.CPERSCOD " _
        & "             LEFT JOIN RRHH RH ON P.CPERSCOD = RH.CPERSCOD " _
        & "             LEFT JOIN " & vsBaseComunes & "AREAS A ON A.CAREACOD = RH.CAREACODActual " _
        & "             LEFT JOIN " & vsBaseComunes & "AGENCIAS AG ON AG.CAGECOD = RH.cAgenciaActual " _
        & "             LEFT JOIN " & vsBaseComunes & "AREAS A1 ON A1.CAREACOD = MAR.CAREACOD " _
        & "             LEFT JOIN " & vsBaseComunes & "AGENCIAS AG1 ON AG1.CAGECOD = MAR.CAGECOD " _
        & "             JOIN MOVDOC MD ON MD.nMOVNRO = M.nMOVNRO " _
        & "             JOIN " & vsBaseComunes & "DOCUMENTO D ON D.nDocTpo = MD.nDocTpo " _
        & "             JOIN MOV AS MAT ON MAT.nMOVNRO = MR.nMOVNRO " _
        & "             JOIN MOVCTA MC ON MC.nMOVNRO  = MAT.nMOVNRO " _
        & "             LEFT JOIN MOVME ME ON ME.nMOVNRO =MC.nMOVNRO AND ME.nMOVITEM=MC.nMOVITEM " _
        & "             LEFT JOIN MOVDOC  MD1 ON MD1.nMOVNRO = MAT.nMOVNRO AND MD1.nDocTpo NOT IN ('" & TpoDocVoucherEgreso & "') " _
        & "             LEFT JOIN " & vsBaseComunes & "DOCUMENTO D1 ON D1.nDocTpo =MD1.nDocTpo " _
        & "             Join "
sql = sql + "             (SELECT MC.cCtaContCod, " _
        & "                     SUM ( CASE  WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)>0 THEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE) " _
        & "                                 WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)<0 THEN 0 END ) AS DEBE, " _
        & "                     SUM ( CASE  WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)<0 THEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE) " _
        & "                                 WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)>0 THEN 0 END ) AS HABER, " _
        & "                     MRSUST.nMOVNROREF AS nMovAtenc " _
        & "              FROM MOVREF  MRSUST " _
        & "                     JOIN MOV M          ON MRSUST.nMOVNRO = M.nMOVNRO   " _
        & "                     JOIN MOVCTA MC      ON MC.nMOVNRO = M.nMOVNRO " _
        & "                     LEFT JOIN MOVME ME  ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM =MC.nMOVITEM " _
        & "                     WHERE   M.nMovFlag = " & gMovFlagVigente & " and not LEFT(m.cOpecod,5) IN ('" & Left(gCGArendirCtaRendMN, 5) & "', '" & Left(gCGArendirCtaRendME, 5) & "') " _
        & "                             AND MRSUST.nMOVNROREF =" & lsMovNroArendirAtencion & " " _
        & "               GROUP BY MC.cCtacontCod , MRSUST.nMOVNROREF ) as SUST " _
        & "                 ON SUST.nMovAtenc= MR.nMovNro " _
        & " WHERE MC.CCTACONTCOD in ('" & psCtaContArendir & "') AND MR.nMOVNRO ='" & lsMovNroArendirAtencion & "' " _
        & " ORDER BY SUST.cCtaContCod "
        
    Case gArendirTipoViaticos
        sql = " SELECT  M.CMOVNRO AS CMOVSOL , M.CMOVNRO AS CMOVATENC, MAR.CPERSCOD, P.CPERSNOMBRE, " _
            & "         ISNULL(A.CAREACOD,'') CAREACOD, ISNULL(A.CAREADESCRIPCION,'') CAREADESCRIPCION, ISNULL(AG.CAGECOD,'') AS CAGECOD  , ISNULL(AG.CAGEDESCRIPCION,'') CAGEDESCRIPCION , " _
            & "         MAR.CAREACOD as cAreaCodArendir, A1.CAREADESCRIPCION as cAreaARendir , ISNULL(MAR.CAGECOD,'') AS cAgeCodArendir , ISNULL(AG1.CAGEDESCRIPCION,'') cAgeArendir , " _
            & "         ABS(REND.Debe) AS ImporteArendir,  MAR.NMOVSALDO, MD.nDocTpo, " _
            & "         (D.CDOCABREV + '-' + MD.CDOCNRO) AS CDOCNRO, MD.dDocFecha as dDocFechaSol , " _
            & "         MD.dDocFecha as dFechaDocAtenc , MD.CDOCNRO as cDocNroAtenc , D.CDOCABREV AS cDocDescAtenc, " _
            & "         SUST.cCtaContCod as CtaContSust, IsNull(ABS(SUST.Debe),0) AS Debe , IsNull(ABS(SUST.HABER),0) HABER " _
            & " FROM    MOVARENDIR MAR " _
            & "         JOIN MOV M                                  ON M.nMOVNRO = MAR.nMOVNRO " _
            & "         JOIN " & vsBasePesonas & "PERSONA P         ON P.CPERSCOD = MAR.CPERSCOD " _
            & "         LEFT JOIN RRHH RH                           ON RH.CPERSCOD = P.CPERSCOD " _
            & "         LEFT JOIN " & vsBaseComunes & "AREAS A      ON A.CAREACOD = RH.cAreaCodActual   " _
            & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG  ON AG.CAGECOD = RH.cAgenciaActual " _
            & "         LEFT JOIN " & vsBaseComunes & "AREAS A1          ON A1.CAREACOD = MAR.CAREACOD " _
            & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG1 ON AG1.CAGECOD = MAR.CAGECOD " _
            & "         JOIN MOVDOC MD                              ON MD.NMOVNRO = M.NMOVNRO  AND MD.nDocTpo ='" & TpoDocRecViaticosARendirCuenta & "' " _
            & "         JOIN DOCUMENTO D                            ON D.nDocTpo = MD.nDocTpo "
    
    sql = sql + "      Join (SELECT    MC.cCtaContCod, " _
            & "                         SUM ( CASE " _
            & "                                 WHEN IsNull(ME.NMOVMEIMPORTE, MC.NMOVIMPORTE)>0 THEN IsNull(ME.NMOVMEIMPORTE, Mc.NMOVIMPORTE) " _
            & "                                 WHEN IsNull(ME.NMOVMEIMPORTE, MC.NMOVIMPORTE)<0 THEN 0 END ) AS DEBE, " _
            & "                         SUM ( CASE " _
            & "                         WHEN IsNull(Me.NMOVMEIMPORTE, Mc.NMOVIMPORTE)<0 THEN IsNull(Me.NMOVMEIMPORTE, Mc.NMOVIMPORTE) " _
            & "                         WHEN IsNull(Me.NMOVMEIMPORTE, Mc.NMOVIMPORTE)>0 THEN 0 END ) AS HABER, " _
            & "                         MRSUST.NMOVNROREF AS NMovSol " _
            & "                 FROM    MOVREF  MRSUST " _
            & "                         JOIN MOV M          ON MRSUST.NMOVNRO = M.NMOVNRO " _
            & "                         JOIN MOVCTA MC      ON MC.NMOVNRO = M.NMOVNRO " _
            & "                         LEFT JOIN MOVME ME  ON ME.NMOVNRO = MC.NMOVNRO AND ME.NMOVITEM =MC.NMOVITEM " _
            & "                 WHERE   M.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "')  " _
            & "                         AND SUBSTRING(M.COPECOD,1,5) IN ('" & Mid(gCGArendirViatAtencMN, 1, 5) & "','" & Mid(gCGArendirViatAtencME, 1, 5) & "') " _
            & "                         AND MRSUST.NMOVNROREF ='" & lsMovNroArendirAtencion & "' AND MC.cCtaContCod  IN ('" & psCtaContArendir & "') " _
            & "                 GROUP BY MC.cCtacontCod , MRSUST.NMOVNROREF  ) as REND " _
            & "          ON M.NMovNro = REND.NMovSol "
            
    sql = sql + "          Join  (SELECT  MC.cCtaContCod, " _
            & "                         SUM ( CASE " _
            & "                                 WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)>0 THEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE) " _
            & "                                 WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)<0 THEN 0 END ) AS DEBE, " _
            & "                         SUM ( CASE " _
            & "                                 WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)<0 THEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE) " _
            & "                                 WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)>0 THEN 0 END ) AS HABER, " _
            & "                         MRSUST.NMOVNROREF AS NMovSol " _
            & "                 FROM    MOVREF  MRSUST " _
            & "                         JOIN MOV M          ON MRSUST.NMOVNRO = M.NMOVNRO " _
            & "                         JOIN MOVCTA MC      ON MC.NMOVNRO = M.NMOVNRO " _
            & "                         LEFT JOIN MOVME ME  ON ME.NMOVNRO = MC.NMOVNRO AND ME.NMOVITEM =MC.NMOVITEM " _
            & "                 WHERE   M.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "') " _
            & "                         AND SUBSTRING(M.COPECOD,1,5) NOT IN ('" & Mid(gCGArendirViatAtencMN, 1, 5) & "','" & Mid(gCGArendirViatAtencME, 1, 5) & "')  and not LEFT(m.cOpecod,5) IN ('" & Left(gCGArendirViatRendMN, 5) & "', '" & Left(gCGArendirViatRendME, 5) & "')  " _
            & "                         AND MRSUST.NMOVNROREF ='" & lsMovNroArendirAtencion & "' GROUP BY MC.cCtacontCod , MRSUST.NMOVNROREF  ) as SUST " _
            & "         ON SUST.NMovSol= M.NMovNro " _
            & " WHERE   M.NMOVNRO ='" & lsMovNroArendirAtencion & "' " _
            & " ORDER BY SUST.cCtaContCod "
            
End Select
If oConect.AbreConexion = False Then Exit Function

Set rs = oConect.CargaRecordSet(sql)
Set GetAsientosSustARendir = rs
oConect.CierraConexion
Set oConect = Nothing
End Function
Public Function GetDetDocCtaCont(ByVal lsMovNroArendirAtencion As String, ByVal psCtaContArendir As String, ByVal psCtaContPendiente As String, Optional pnTipoArendir As ArendirTipo = gArendirTipoCajaGeneral) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set rs = New ADODB.Recordset
Set oConect = New DConecta

If oConect.AbreConexion = False Then Exit Function

Select Case pnTipoArendir
    Case gArendirTipoCajaGeneral, gArendirTipoAgencias
        sql = "     Select  M.CMOVNRO AS CMOVNROSOL, MR.nMOVNRO AS nMOVATENC , " _
            & "             SUST.cCtaContCod as CtaContSust, IsNull(ABS(SUST.Debe),0) Debe , IsNull(ABS(SUST.HABER),0) HABER , " _
            & "             SUST.nDocTpo , SUST.cDocDesc, SUST.cDocAbrev, SUST.CDOCNRO, SUST.DDOCFECHA, SUST.cMovDesc " _
            & "     FROM    MOV M JOIN MOVREF MR ON M.nMOVNRO = MR.nMOVNROREF " _
            & "             Join " _
            & "             (SELECT MC.cCtaContCod, " _
            & "                     CASE    WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)>0 THEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE) " _
            & "                             WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)<0 THEN 0 END  AS DEBE, " _
            & "                     CASE    WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)<0 THEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE) " _
            & "                             WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)>0 THEN 0 END AS HABER, " _
            & "                     MRSUST.nMovNroRef as nMovAtenc, MD.nDocTpo, D.cDocDesc, D.cDocAbrev, MD.CDOCNRO, MD.DDOCFECHA , M.cMovDesc " _
            & "              FROM   MOVREF  MRSUST " _
            & "                     JOIN MOV M              ON MRSUST.nMOVNRO = M.nMOVNRO " _
            & "                     JOIN MOVCTA MC          ON MC.nMOVNRO = M.nMOVNRO " _
            & "                     JOIN MOVDOC MD      ON MD.nMOVNRO = M.nMOVNRO " _
            & "                     JOIN DOCUMENTO D    ON D.nDocTpo = MD.nDocTpo " _
            & "                     LEFT JOIN MOVME ME  ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM =MC.nMOVITEM " _
            & "              WHERE  M.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "')  " _
            & "                     AND CCTACONTCOD NOT IN('" & psCtaContArendir & "','" & psCtaContPendiente & "') AND MRSUST.nMovNroRef=" & lsMovNroArendirAtencion & " ) as SUST " _
            & "                     ON SUST.nMovAtenc= MR.nMovNro " _
            & "     WHERE   MR.nMOVNRO =" & lsMovNroArendirAtencion & " ORDER BY CtaContSust "
    
    Case gArendirTipoViaticos
    
        sql = " SELECT  CtaContSust, NMovNroSol, NMovNroAtenc , nDocTpo, cDocDesc, cDocAbrev, CDOCNRO, DDOCFECHA, cMovDesc," _
            & "         IsNull(SUM(C.DEBE),0) AS DEBE , IsNull(SUM(C.HABER),0) AS HABER " _
            & " FROM ( SELECT   MC.cCtaContCod as CtaContSust , " _
            & "         CASE    WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)>0 THEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE) " _
            & "                 WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)<0 THEN 0 END  AS DEBE, " _
            & "         CASE    WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)<0 THEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE) " _
            & "                 WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)>0 THEN 0 END AS HABER, " _
            & "         MRSUST.nMovNroRef as nMovNroSol, MRSUST.nMovNroRef as nMovNroAtenc, " _
            & "         MD.nDocTpo, D.cDocDesc, D.cDocAbrev, MD.CDOCNRO, MD.DDOCFECHA , m.cMovDesc " _
            & " FROM    MOVREF  MRSUST " _
            & "         JOIN MOV M          ON MRSUST.nMOVNRO = M.nMOVNRO " _
            & "         JOIN MOVCTA MC          ON MC.nMOVNRO = M.nMOVNRO " _
            & "         JOIN MOVDOC MD          ON MD.nMOVNRO = M.nMOVNRO " _
            & "         JOIN DOCUMENTO D        ON D.nDocTpo = MD.nDocTpo " _
            & "         LEFT JOIN MOVME ME      ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM =MC.nMOVITEM " _
            & " WHERE   M.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "')  " _
            & "         AND substring(copecod,1,5) NOT IN ('" & Mid(gCGArendirViatAtencMN, 1, 5) & "','" & Mid(gCGArendirViatAtencME, 1, 5) & "') " _
            & "         AND CCTACONTCOD NOT IN('" & psCtaContArendir & "','" & psCtaContPendiente & "') " _
            & "         AND MRSUST.nMovNroRef=" & lsMovNroArendirAtencion & " ) AS C " _
            & " GROUP BY  CtaContSust, nMovNroSol, nMovNroAtenc , nDocTpo, cDocDesc, cDocAbrev, CDOCNRO, DDOCFECHA, cMovDesc " _
            & " ORDER BY C.CtaContSust "
    
    Case Else
    
End Select
Set rs = oConect.CargaRecordSet(sql)
Set GetDetDocCtaCont = rs
oConect.CierraConexion
Set oConect = Nothing
End Function
Public Function GrabaRendicionExacta(ByVal pnTipoArendir As ArendirTipo, psFormatoFecha As String, _
                                     ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                     ByVal pnMovNroAtenc As Long, ByVal pnMovNroSol As Long, _
                                     Optional ByVal psAreaCh As String = "", Optional ByVal psAgeCh As String = "", _
                                     Optional ByVal pnNroProc As Integer = 0, Optional ByVal pnSaldo As Currency) As Integer
Dim oDMov As DMov
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim lsMsgErr As String
Set oDMov = New DMov

On Error GoTo GrabaRendicionExactaErr
oDMov.BeginTrans
lbTrans = True
If pnTipoArendir = gArendirTipoCajaChica Then
    oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
Else
    oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
End If
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, 0, "0", "0"
oDMov.InsertaMovRef lnMovNro, pnMovNroAtenc
ActualizaMovSust pnMovNroSol, psMovNro, oDMov, , pnTipoArendir, pnNroProc, lnMovNro
oDMov.ActualizaMovArendir pnMovNroSol, pnTipoArendir, , , , 0
If pnTipoArendir = gArendirTipoCajaChica Then
    oDMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, pnNroProc, gCHTipoProcArendir
End If
oDMov.CommitTrans
lbTrans = False
GrabaRendicionExacta = 0
Set oDMov = Nothing
Exit Function
GrabaRendicionExactaErr:
   lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaRendicionExacta", lsMsgErr
End Function
Public Function GetOpeRendicion(ByVal psOpeCod As String, ByVal psDocTpo As TpoDoc, ByVal psCtaContArendir As String, ByVal psCtaContPend As String, Optional ByVal pbCtaCont As Boolean = False, Optional psDH As String = "") As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsFiltroCta As String
Dim oConect As DConecta
Dim lsFiltroDH As String

Set oConect = New DConecta
 
If oConect.AbreConexion = False Then Exit Function
If pbCtaCont = False Then
    lsFiltroCta = " AND OC.CCTACONTCOD IS NULL "
Else
    lsFiltroCta = " AND OC.CCTACONTCOD IS NOT NULL AND OC.cCtaContCod not in ('" & psCtaContArendir & "') "
End If
 
If psDH <> "" Then
    lsFiltroDH = " and cOpeCtaDH ='" & psDH & "'"
End If
sql = " Select  O.COPECOD, OD.nDocTpo " _
    & " From    " & vsBaseComunes & "OpeTpo O " _
    & "         LEFT JOIN " & vsBaseComunes & "OpeCta OC  ON O.cOpeCod= OC.cOpeCod " _
    & "         Left Join " & vsBaseComunes & "OpeDoc OD on OD.cOpeCod = O.cOpeCod " _
    & " Where   Substring(O.cOpeCod,1,LEN('" & psOpeCod & "')) ='" & psOpeCod & "' " _
    & "         and OD.nDocTpo " & IIf(psDocTpo = -1, " IS NULL ", " ='" & psDocTpo & "'") _
    & "         " & lsFiltroCta & lsFiltroDH _
    & " GROUP BY O.COPECOD, OD.nDocTpo "
 
Set rs = oConect.CargaRecordSet(sql)
GetOpeRendicion = ""
If Not rs.EOF And Not rs.BOF Then
    GetOpeRendicion = rs!cOpeCod
End If
rs.Close
Set rs = Nothing
oConect.CierraConexion
Set oConect = Nothing

'If oConect.AbreConexion = False Then Exit Function
'If psDocTpo = -1 Then
'    lsFiltroCta = " AND OC.CCTACONTCOD IS NULL "
'Else
'    lsFiltroCta = " AND OC.CCTACONTCOD IS NOT NULL AND OC.cCtaContCod not in ('" & psCtaContArendir & "') " ','" & psCtaContPend & "')"
'End If
'
'If psDH <> "" Then
'    lsFiltroDH = " and cOpeCtaDH ='" & psDH & "'"
'End If
'sql = " Select  O.COPECOD, OD.nDocTpo " _
'    & " From    " & vsBaseComunes & "OpeTpo O " _
'    & "         LEFT JOIN " & vsBaseComunes & "OpeCta OC  ON O.cOpeCod= OC.cOpeCod " _
'    & "         Left Join " & vsBaseComunes & "OpeDoc OD on OD.cOpeCod = O.cOpeCod " _
'    & " Where   Substring(O.cOpeCod,1,LEN('" & psOpeCod & "')) ='" & psOpeCod & "' " _
'    & "         and OD.nDocTpo " & IIf(psDocTpo = -1, " IS NULL ", " ='" & psDocTpo & "'") _
'    & "         " & lsFiltroCta & lsFiltroDH _
'    & " GROUP BY O.COPECOD, OD.nDocTpo "
'
'Set rs = oConect.CargaRecordSet(sql)
'GetOpeRendicion = ""
'If Not rs.EOF And Not rs.BOF Then
'    GetOpeRendicion = rs!COPECOD
'End If
'rs.Close
'Set rs = Nothing
'oConect.CierraConexion
'Set oConect = Nothing
End Function
Public Function GrabaRendicionEfectivo(ByVal pnTipoArendir As ArendirTipo, ByVal psFormatoFecha As String, _
                                        ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                        ByVal psCtaPendiente As String, ByVal psCtaEfectivo As String, _
                                        ByVal pnImporte As Currency, ByVal rsBilletaje As ADODB.Recordset, _
                                        ByVal psMovAtencion As String, ByVal pnMovNroSol As Long, _
                                        ByVal pbIngreso As Boolean, _
                                        ByVal psCtaDiferencia As String, ByVal pnMontoDiferencia As Currency, _
                                        Optional rsPend As ADODB.Recordset) As Integer
            
Dim lnMovItem As Long
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
On Error GoTo GrabaRendicionEfectivoErr

Set oDMov = New DMov

'GrabaAtencionArendir = 1
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"

'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1: lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaPendiente, pnImporte * IIf(pbIngreso = True, -1, 1)

'Guardamos el billetaje si este posee
oDMov.GrabaMovimientoEfectivo lnMovNro, lnMovItem, rsBilletaje, psCtaEfectivo, IIf(pbIngreso, "D", "H")

If pnMontoDiferencia <> 0 Then
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDiferencia, pnMontoDiferencia * IIf(pbIngreso = True, 1, -1)
End If
If pnTipoArendir <> -1 Then
    oDMov.InsertaMovRef lnMovNro, psMovAtencion
    oDMov.ActualizaMovArendir pnMovNroSol, pnTipoArendir, , , , 0
    ActualizaMovSust pnMovNroSol, psMovNro, oDMov, , pnTipoArendir, , lnMovNro
    'ActualizaMovSust pnMovNroSol, psMovNro, oDMov, , pnTipoArendir
Else
    If Not IsNull(rsPend) Then
        rsPend.MoveFirst
        Do While Not rsPend.EOF
           If rsPend!Ok = 1 Then
                'oDMov.InsertaMovRef lnMovNro, rsPend!nMovNro
                oDMov.ActualizaMovPendientesRend rsPend!nMovNro, psCtaPendiente, rsPend!Rendicion
            End If
            rsPend.MoveNext
        Loop
    End If
End If
If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If

oDMov.ActualizaSaldoMovimiento psMovNro, "+"

GrabaRendicionEfectivo = 0

oDMov.CommitTrans

lbTrans = False

Set oDMov = Nothing
Exit Function
GrabaRendicionEfectivoErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Call RaiseError(MyUnhandledError, "NARendir:GrabaRendicionEfectivoErr Method")
End Function

Public Function GrabaRendicionVentanilla(ByVal pnTipoArendir As ArendirTipo, ByVal psFormatoFecha As String, _
                                         ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                         ByVal psCtaPendiente As String, ByVal psCtaPendVentanilla As String, _
                                         ByVal pnImporte As Currency, ByVal rsPago As ADODB.Recordset, _
                                         ByVal psMovAtencion As String, ByVal pnMovNroSol As Long, _
                                         ByVal pbIngreso As Boolean, _
                                         ByVal psCtaDiferencia As String, ByVal pnMontoDiferencia As Currency, _
                                         Optional psAgeCodRef As String, Optional rsPend As ADODB.Recordset, Optional pbBitCentral As Boolean = False) As Integer
Dim lnMovItem As Long
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lnImporteVenta As Currency
Dim oNeg As New NNegOpePendientes

Set oDMov = New DMov
On Error GoTo GrabaRendicionVentanillaErr

'GrabaAtencionArendir = 1
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"

lnImporteVenta = 0
If Not rsPago Is Nothing Then
    If Not rsPago.State = adStateClosed Then
        Do While Not rsPago.EOF
            If rsPago.Fields(1).value = "1" Then   'Check Seleccionado
               lnImporteVenta = lnImporteVenta + rsPago.Fields(5).value
               If pbBitCentral Then
                  oNeg.ActualizaNegocioReferencia rsPago.Fields(10), Right(rsPago.Fields(7), 2), psMovNro
               End If
               oDMov.InsertaMovRef lnMovNro, rsPago.Fields(10), Right(rsPago.Fields(7), 2)
            End If
            rsPago.MoveNext
        Loop
        Set oNeg = Nothing
    End If
End If
If lnImporteVenta = 0 Then
    oDMov.RollbackTrans
    lbTrans = False
    Err.Raise "50001", "NARendir:GrabaRendicionVentanilla", "No se Seleccionó Ingreso hecho por Ventanilla"
End If

'Guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1: lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaPendVentanilla, lnImporteVenta * IIf(pbIngreso = True, 1, -1)
lnMovItem = lnMovItem + 1
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaPendiente, pnImporte * IIf(pbIngreso = True, -1, 1)

If pnMontoDiferencia <> 0 Then
    lnMovItem = lnMovItem + 1
    oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDiferencia, pnMontoDiferencia * IIf(pbIngreso = True, 1, -1)
End If
If rsPend Is Nothing Then  'A rendir Cuenta
    oDMov.InsertaMovRef lnMovNro, IIf(psMovAtencion = "", "0", psMovAtencion)
    oDMov.ActualizaMovArendir pnMovNroSol, pnTipoArendir, , , , 0
    ActualizaMovSust pnMovNroSol, psMovNro, oDMov, , pnTipoArendir, , lnMovNro
Else
    If Not rsPend.State = adStateClosed Then
        Do While Not rsPend.EOF
           If rsPend!Ok = 1 Then
                oDMov.ActualizaMovPendientesRend rsPend!nMovNro, psCtaPendiente, rsPend!Rendicion
                If psAgeCodRef <> "" Then
                    If Not pbBitCentral Then
                        oNeg.ActualizaNegocioReferencia rsPend!nMovNro, psAgeCodRef, psMovNro
                    End If
                    oDMov.InsertaMovRef lnMovNro, rsPend!nMovNro, psAgeCodRef
                Else
                    oDMov.InsertaMovRef lnMovNro, rsPend!nMovNro
                End If
            End If
            rsPend.MoveNext
        Loop
    End If
End If
GrabaRendicionVentanilla = 0
If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If
oDMov.ActualizaSaldoMovimiento psMovNro, "+"

oDMov.CommitTrans

lbTrans = False

Set oDMov = Nothing
Exit Function
GrabaRendicionVentanillaErr:

    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Call RaiseError(MyUnhandledError, "NARendir:GrabaRendicionVentanilla Method")
End Function

Private Function ActualizaMovSust(ByVal pnMovNroSol As Long, ByVal psMovRendicion As String, ByRef oMov As DMov, Optional ByVal pbExtorno As Boolean = False, Optional pnArendirTipo As ArendirTipo = gArendirTipoCajaGeneral, Optional pnProcNro = 0, Optional pnMovRendicion As Long = 0)
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsMovNroNew As String
Dim lsMovNroActualizaSaldo As String

lsMovNroActualizaSaldo = ""
Set rs = New ADODB.Recordset
sql = " Select  M.cMovNro, MA.nMovNro, MA.nMovNroSust, MA.nMovNroRend, M.cOpeCod " _
    & " From    Mov M JOIN MovArendirSust MA ON MA.nMOVNROSUST = M.nMOVNRO " _
    & " Where   MA.nMovNro = " & pnMovNroSol & " AND M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagExtornado & "','" & gMovFlagModificado & "') "

Set rs = oMov.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    Do While Not rs.EOF
        If Not pbExtorno Then
              lsMovNroNew = NuevoMovNro(psMovRendicion, rs!cMovNro, oMov)
              'Actualizamos los cMovNro de las Sustentaciones con la misma fecha de la
              'Rendición
              oMov.ActualizaMovArendirSust pnMovRendicion, rs!nMovNroSust
              If pnArendirTipo = gArendirTipoCajaChica Then
                 oMov.ActualizaMov rs!nMovNroSust, , gMovEstContabNoContable, , lsMovNroNew
                 If pnProcNro > 0 Then
                    oMov.ActualizaMovCajaChicaNroProc oMov.GetnMovNro(lsMovNroNew), pnProcNro
                 End If
              Else
                 oMov.ActualizaMov rs!nMovNroSust, , gMovEstContabMovContable, , lsMovNroNew
                 If Mid(rs!cOpeCod, 3, 1) = gMonedaExtranjera And Left(rs!cMovNro, 6) <> Left(lsMovNroNew, 6) Then
                    oMov.RevierteGeneraMovME rs!nMovNroSust
                    oMov.GeneraMovME rs!nMovNroSust, lsMovNroNew
                 End If
                 lsMovNroActualizaSaldo = lsMovNroActualizaSaldo & "'" & lsMovNroNew & "',"
              End If
        Else 'Extorno
              If pnArendirTipo = gArendirTipoCajaChica Then
                 oMov.ActualizaMov rs!nMovNroSust, , gMovEstContabSustPendRendir
              Else
                 oMov.ActualizaMov rs!nMovNroSust, , gMovEstContabNoContable
                 lsMovNroActualizaSaldo = lsMovNroActualizaSaldo & "'" & rs!cMovNro & "',"
              End If
        End If
        rs.MoveNext
    Loop
End If
RSClose rs
If lsMovNroActualizaSaldo <> "" Then
    lsMovNroActualizaSaldo = Left(lsMovNroActualizaSaldo, Len(lsMovNroActualizaSaldo) - 1)
    If Not pbExtorno Then
       oMov.ActualizaSaldoMovimiento lsMovNroActualizaSaldo, "+"
    Else
       oMov.ActualizaSaldoMovimiento lsMovNroActualizaSaldo, "-"
    End If
End If
End Function

Private Function NuevoMovNro(ByVal psMovNroRend As String, ByVal psMovNroSust As String, ByRef oMov As DMov) As String
Dim ldFechaArendir As Date
Dim lsAge As String
Dim lsCmac As String
Dim lsUser As String

ldFechaArendir = CDate(Mid(psMovNroRend, 7, 2) & "/" & Mid(psMovNroRend, 5, 2) & "/" & Mid(psMovNroRend, 1, 4))
lsCmac = Mid(psMovNroSust, 15, 3)
lsAge = Mid(psMovNroSust, 18, 2)
lsUser = Right(psMovNroSust, 4)
NuevoMovNro = oMov.GeneraMovNro(ldFechaArendir, lsAge, lsUser)

End Function
Public Function GrabaOtrosIngEgresos(ByVal pnTipoArendir As ArendirTipo, psFormatoFecha As String, _
                           ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                           ByVal pnMovNroAtenc As Long, ByVal psMovNroSol As String, _
                           ByVal psCtaPendiente As String, ByVal psCtaOtros As String, _
                           ByVal pnImporte As Currency, ByVal pbIngreso As Boolean) As Integer
Dim oDMov As DMov
Dim lnMovItem As Integer
Dim lnObjOrden As Integer
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lsMsgErr As String
Set oDMov = New DMov

On Error GoTo GrabaOtrosIngEgresosErr
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, 0, "0", "0"
lnMovItem = 0
lnMovItem = lnMovItem + 1
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaPendiente, pnImporte * IIf(pbIngreso = True, -1, 1)
oDMov.InsertaMovRef lnMovNro, pnMovNroAtenc
ActualizaMovSust psMovNroSol, psMovNro, oDMov, , pnTipoArendir, , lnMovNro
oDMov.CommitTrans
lbTrans = False
GrabaOtrosIngEgresos = 0
Set oDMov = Nothing
Exit Function
GrabaOtrosIngEgresosErr:
    lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaOtrosIngEgresos", Err.Description
End Function
Public Function GrabaRendicionIngresoCheque(ByVal pnTipoArendir As ArendirTipo, ByVal psFormatoFecha As String, _
                                            ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                            ByVal psCtaPendiente As String, ByVal psCtaContCheque As String, _
                                            ByVal psObjetoCodProd As String, ByVal psObjetoArea As String, ByVal psAgeCod As String, _
                                            ByVal pnImporte As Currency, ByVal pnMovAtencion As Long, ByVal pnMovNroSol As Long, _
                                            ByVal psNumCheque As String, ByVal psPersCodIf As String, ByVal psTipoIF As String, ByVal pnPlaza As ChequePlaza, _
                                            ByVal psCtaCheque As String, ByVal pdFechaReg As Date, ByVal pdFechaVal As Date, ByVal pnMonedaCheque As Moneda, _
                                            Optional ByVal pnEstadoCheque As ChequeEstado = gChqEstEnValorizacion, _
                                            Optional ByVal pnEstadoCG As CGEstadosChq = gCGEstadosChqRecibido, _
                                            Optional ByVal pnConfirmaCaja As CGEstadoConfCheque = ChqCGSinConfirmacion, _
                                            Optional ByVal psAreaCodChq As String, Optional ByVal psAgeCodChq As String) As Integer
                
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim lsSubCtaDebe As String
Dim oContFunct As NContFunciones


Set oDMov = New DMov
Set oContFunct = New NContFunciones

On Error GoTo GrabaRendicionChequeErr

'GrabaAtencionArendir = 1
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaContCheque, pnImporte
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, psObjetoCodProd
lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, psAgeCod, psObjetoArea

'guardamos la cuenta de ContraResta
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaPendiente, pnImporte * -1
If psNumCheque <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocCheque, psNumCheque, Format(pdFechaReg, gsFormatoFecha)
End If
'grabacion dentro de la tabla de cheques
oDMov.InsertaCheque TpoDocCheque, psNumCheque, psPersCodIf, psTipoIF, pnPlaza, psCtaCheque, pnImporte, pdFechaReg, pdFechaVal, pnEstadoCG, pnConfirmaCaja, pnMonedaCheque, , psAreaCodChq, psAgeCodChq
oDMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIf, psTipoIF, pdFechaReg, gChqEstEnValorizacion, psMovNro, psCtaCheque

oDMov.InsertaMovRef lnMovNro, pnMovAtencion
ActualizaMovSust pnMovNroSol, psMovNro, oDMov, , pnTipoArendir, , lnMovNro
oDMov.ActualizaMovArendir pnMovNroSol, pnTipoArendir, , , , 0
If Mid(psOpeCod, 3, 1) = "2" Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If

oDMov.CommitTrans
lbTrans = False
GrabaRendicionIngresoCheque = 0

Set oDMov = Nothing
Exit Function
GrabaRendicionChequeErr:
  lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, Err.Source, lsMsgErr
End Function
Public Function GrabaRendicionGiroDocumento(ByVal pnTipoArendir As ArendirTipo, _
                                            ByVal psMovNro As String, ByVal pnMovNroSol As Long, ByVal psMovNroAtenc As String, _
                                            ByVal psOpeCod As String, ByVal psMovDesc As String, ByVal psCtaDebe As String, _
                                            ByVal psCtaHaber As String, ByVal psPersCod As String, ByVal pnImporte As Currency, _
                                            ByVal psDocTpo As TpoDoc, ByVal psDocNro As String, _
                                            ByVal psFechaDoc As String, ByVal psDocNroVoucher As String, _
                                            ByVal psPersCodIf As String, ByVal psTipoIF As String, ByVal psCodCtaIf As String, _
                                            Optional ByVal psAreaCh As String = "", Optional ByVal psAgeCh As String = "", _
                                            Optional pnNroProc As Integer = 0, Optional pnMotivo As Long = -1, Optional psCtaCod As String = "", Optional gbBitCentral As Boolean = False, Optional pbGrabaOpeNegocio As Boolean = False, Optional psPersNombre As String = "") As Integer
                        
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim oDocRec As NDocRec
Dim lsMsg   As String
Set oDMov = New DMov
Set oDocRec = New NDocRec
On Error GoTo ErrorGrabaRendicionGiroDocumento

'GrabaAtencionArendir = 1
oDMov.BeginTrans
lbTrans = True
If pnTipoArendir = gArendirTipoCajaChica Then
    oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
Else
    oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
End If
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1: lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaDebe, Abs(pnImporte)
If pnTipoArendir = gArendirTipoCajaChica And psAreaCh <> "" Then
    oDMov.InsertaMovObj lnMovNro, lnMovItem, "1", ObjCMACAgenciaArea
    oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, "1", psAgeCh, psAreaCh
End If
'guardamos la cuenta del haber
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oDMov.InsertaMovCta lnMovNro, Format(lnMovItem, "#0"), psCtaHaber, Abs(pnImporte) * -1
If psPersCodIf <> "" And psCodCtaIf <> "" Then
    lnMovOrden = lnMovOrden + 1
    oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjEntidadesFinancieras
    oDMov.InsertaMovObjIF lnMovNro, Format(lnMovItem, "#0"), Format(lnMovOrden, "#0"), psPersCodIf, psTipoIF, psCodCtaIf
    
End If
Select Case Val(psDocTpo)
    Case TpoDocNotaAbono, TpoDocNotaCargo
        oDMov.InsertaNotaAbonoCargo psDocTpo, psDocNro, gNCNARegistrado, pnMotivo, Abs(pnImporte)
        oDMov.InsertaNotaAbonoCargoEst psDocTpo, psDocNro, gNCNARegistrado, psMovNro
        oDMov.InsertaRegDocCuenta psDocTpo, psDocNro, psCtaCod
    
        If pbGrabaOpeNegocio Then
            Dim oCapta As NCapMovimientos
            Dim lnSaldo As Double
            Dim lsMovNroNegocio  As String
            Set oCapta = New NCapMovimientos
            If gbBitCentral Then
                lsMovNroNegocio = oDMov.GeneraMovNro(, , , psMovNro)
                If psDocTpo = TpoDocNotaAbono Then
                    lnSaldo = oCapta.CapAbonoCuentaAho(psCtaCod, pnImporte, gCGArendirAtencAbonoCent, lsMovNroNegocio, psMovDesc, TpoDocNotaAbono, psDocNro, , , , , , , , , , True, False, , , oDMov.GetConexion)
                Else
                    lnSaldo = oCapta.CapCargoCuentaAho(psCtaCod, pnImporte, gCGArendirAtencCargoCent, lsMovNroNegocio, psMovDesc, psDocTpo, psDocNro, , , , , , , , , , True, "", , oDMov.GetConexion, False)
                End If
            Else
                Dim oCon As New DConecta
                Dim oDis As New NRHProcesosCierre
                
                If oCon.AbreConexion Then 'Remota(Left(psCtaCod, 2))
                    oCon.BeginTrans
                    If psDocTpo = TpoDocNotaAbono Then
                        lnSaldo = oDis.Abono(psCtaCod, pnImporte, gCGArendirAtencAbonoDist, gCGArendirAtencAbonoDist, "112" & Left(psCtaCod, 2), Right(psMovNro, 4), psDocNro, "PAGO : " & psMovDesc, oCon, GetFechaMov(psMovNro, True))
                    Else
                        lnSaldo = oDis.Cargo(psCtaCod, pnImporte, gCGArendirAtencCargoDist, "112" & Left(psCtaCod, 2), Right(psMovNro, 4), psDocNro, "CARGO : " & psMovDesc, oCon, GetFechaMov(psMovNro, True))
                    End If
                    oCon.CommitTrans
                End If
                oCon.CierraConexion
                Set oCon = Nothing
                Set oDis = Nothing
            End If
        End If
    
    Case TpoDocOrdenPago
        
End Select
If psDocNro <> "" Then
    oDMov.InsertaMovDoc lnMovNro, psDocTpo, psDocNro, psFechaDoc
End If
If psDocNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocNroVoucher, psFechaDoc
End If
If psPersCod <> "" Then
    oDMov.InsertaMovGasto lnMovNro, psPersCod, ""
End If
If psPersNombre <> "" Then
    oDMov.InsertaMovOtrosItem lnMovNro, 1, "PersChq", Abs(pnImporte), psPersNombre
End If
oDMov.InsertaMovRef lnMovNro, psMovNroAtenc
ActualizaMovSust pnMovNroSol, psMovNro, oDMov, , pnTipoArendir, pnNroProc, lnMovNro
oDMov.ActualizaMovArendir pnMovNroSol, pnTipoArendir, , , , 0
If pnTipoArendir = gArendirTipoCajaChica Then
    Dim lnSaldoCH As Currency
    Dim oCh As NCajaChica
    Set oCh = New NCajaChica
    lnSaldoCH = oCh.GetDatosCajaChica(psAreaCh, psAgeCh, SaldoActual)
    oDMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProc, , , , , lnSaldoCH + pnImporte
    Set oCh = Nothing
    oDMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, pnNroProc, gCHTipoProcArendir
End If
If Mid(psOpeCod, 3, 1) = "2" Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If
If Not pnTipoArendir = gArendirTipoCajaChica Then
    oDMov.ActualizaSaldoMovimiento psMovNro, "+"
End If
GrabaRendicionGiroDocumento = 0

oDMov.CommitTrans
lbTrans = False
Set oDMov = Nothing
Exit Function
ErrorGrabaRendicionGiroDocumento:
    lsMsg = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaRendicionGiroDocumento", lsMsg
End Function
Public Function EliminaArendir(ByVal psMovNroExt As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                               ByVal psMovNroAnt As String, ByVal psMovNroSol As String, ByVal pnTipoArendir As ArendirTipo, _
                               ByVal pnImporte As Currency, ByVal pnSaldo As Currency, ByVal pnArendirFase As ARendirFases, _
                               Optional ByVal pbEjecBatch As Boolean = True, _
                               Optional psAreaCh As String = "", Optional psAgeCh As String = "", _
                               Optional pnNroProc As Integer = 0) As Integer
                                       
    On Error GoTo EliminaArendirErr
    Dim oMov As DMov
    Dim sql As String
    Dim lnSaldoArendir As Currency
    Set oMov = New DMov
    EliminaArendir = 1
    lnSaldoArendir = pnSaldo - pnImporte
    oMov.InsertaMov psMovNroExt, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
    oMov.ActualizaMov psMovNroAnt, , gMovEstContabRechazado, gMovFlagEliminado
    oMov.ActualizaMovArendir psMovNroSol, pnTipoArendir, , , , lnSaldoArendir
    oMov.InsertaMovRef psMovNroExt, psMovNroAnt
    If pnArendirFase = ArendirExtornoRendicion Then
        ActualizaMovSust psMovNroSol, psMovNroAnt, oMov, True, pnTipoArendir
        oMov.ActualizaArendirSust psMovNroSol
    End If
    If pnTipoArendir = gArendirTipoCajaChica Then
        Dim lnSaldoCH As Currency
        Dim oCh As NCajaChica
        Set oCh = New NCajaChica
        lnSaldoCH = oCh.GetDatosCajaChica(psAreaCh, psAgeCh, SaldoActual)
        oMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProc, , , , , lnSaldoCH + pnImporte
        Set oCh = Nothing
    End If
    EliminaArendir = 0
    Exit Function
EliminaArendirErr:
    Call RaiseError(MyUnhandledError, "NARendir:EliminaArendir Method")
End Function
Public Function ExtornaArendir(ByVal psMovNroExt As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                               ByVal pnMovNroAnt As Long, ByVal pnMovNroSol As Long, _
                               ByVal pnTipoArendir As ArendirTipo, _
                               ByVal pnImporte As Currency, ByVal pnSaldo As Currency, _
                               ByVal pnArendirFase As ARendirFases, _
                               Optional psAreaCh As String = "", Optional psAgeCh As String = "", _
                               Optional pnNroProc As Integer = 0, _
                               Optional ByRef pnMovNroExtSolCH As Long = 0) As Integer
    
    On Error GoTo ExtornaArendirErr
    Dim oMov As DMov
    Dim lnMovNro As Long
    Dim lbTrans As Boolean
    Dim sql As String
    Dim lnSaldoArendir As Currency
    
    Set oMov = New DMov
    ExtornaArendir = 1
    oMov.BeginTrans
    lbTrans = True
    'oMov.InsertaMov psMovNroExt, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagDeExtorno
    'lnMovNro = oMov.GetnMovNro(psMovNroExt)
    oMov.ExtornaMovimiento psMovNroExt, pnMovNroAnt, psOpeCod, psMovDesc, True
    
    'oMov.ActualizaMov pnMovNroAnt, , , gMovFlagExtornado
    lnSaldoArendir = pnSaldo - pnImporte
    oMov.ActualizaMovArendir pnMovNroSol, pnTipoArendir, , , , lnSaldoArendir
    'oMov.InsertaMovRef lnMovNro, pnMovNroAnt
    If pnArendirFase = ArendirExtornoRendicion Then
        ActualizaMovSust pnMovNroSol, pnMovNroAnt, oMov, True, pnTipoArendir
        oMov.ActualizaArendirSust pnMovNroSol
    End If
    If pnTipoArendir = gArendirTipoCajaChica Then
        Dim lnSaldoCH As Currency
        Dim oCh As NCajaChica
        Set oCh = New NCajaChica
        lnSaldoCH = oCh.GetDatosCajaChica(psAreaCh, psAgeCh, SaldoActual)
        oMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProc, , , , , lnSaldoCH + pnImporte
        Set oCh = Nothing
    End If
    ExtornaArendir = 0
    '***Agregado por ELRO el 20120625, según OYP-RFC047-2012
    pnMovNroExtSolCH = pnMovNroSol
    '***Fin Agregado por ELRO*******************************
    oMov.CommitTrans
    lbTrans = False
    Exit Function
ExtornaArendirErr:
    If lbTrans Then
        lbTrans = False
        oMov.RollbackTrans
    End If
    Call RaiseError(MyUnhandledError, "NARendir:ExtornaArendir Method")
End Function

'***Modificado por ELRO el 20120427, según OYP-RFC005-2012
'Public Function GetRendicionesArendir(ByVal psAgeCod As String, ByVal psAreaCod As String, _
'                                        ByVal psAreaArendir As String, ByVal psAgeArendir As String, _
'                                        ByVal pnTipoArendir As ArendirTipo, _
'                                        ByVal psCtaArendir As String, ByVal psCtaPendiente As String, _
'                                        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal pnMoneda As Moneda) As ADODB.Recordset
Public Function GetRendicionesArendir(ByVal psAgeCod As String, _
                                      ByVal psAreaCod As String, _
                                      ByVal psAreaArendir As String, _
                                      ByVal psAgeArendir As String, _
                                      ByVal pnTipoArendir As ArendirTipo, _
                                      ByVal psCtaArendir As String, _
                                      ByVal psCtaPendiente As String, _
                                      ByVal pdDesde As Date, _
                                      ByVal pdHasta As Date, _
                                      ByVal pnMoneda As Moneda, _
                                      Optional ByVal pbTodo As Boolean = False) As ADODB.Recordset

Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsFiltroAge  As String
Dim lsFiltroArea  As String
Set oConect = New DConecta
Set rs = New Recordset
If psAgeCod <> "" Then
    lsFiltroAge = " and RH.cAgenciaActual='" & psAgeCod & "'  "
End If
If psAreaCod <> "" Then
    lsFiltroArea = " and RH.cAreaCodActual='" & psAreaCod & "' "
End If

If oConect.AbreConexion = False Then Exit Function
Select Case pnTipoArendir
    Case gArendirTipoCajaGeneral, gArendirTipoAgencias
    
'***Modificado por ELRO el 20120510, según OYP-RFC016-2012
'    sql = "SELECT  Rend.cOpeDesc, Rend.cDocAbrevRend, Rend.cDocNroRen, ISNULL(CONVERT(CHAR(10),Rend.dDocFechaRend,103), SUBSTRING(Rend.cMovRendicion,7,2) + '/' + SUBSTRING(Rend.cMovRendicion,5,2) + '/' + SUBSTRING(Rend.cMovRendicion,1,4) ) AS dDocFechaRend, MDAR.cDocNro as cDocNroSol ,   " _
'            & "         CONVERT(CHAR(10),MDAR.dDocFecha,103) AS dDocSolicitud, CONVERT(CHAR(50),P.CPERSNOMBRE) AS cPersNombre, ISNULL(Abs(Rend.MontoRendicion),0) AS MontoRend, ISNULL(Rend.MontoRendicion,0) AS MontoRendReal, " _
'            & "         ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS NIMPORTE, MAR.nMovSaldo, ISNULL(A.cAreaDescripcion,'') cAreaDescripcion , " _
'            & "         P.cPersCod, Rend.nMovRendicion, M.nMOVNRO as cMovNroAtenc, MAR.nMOVNRO as cMovNroSol, Rend.CMOVDESC, Rend.nDocTpoRend,  " _
'            & "         MDAR.nDocTpo as nDocTpoSol,ISNULL(A.cAreaCod,'') cAreaCod ,ISNULL(AG.cAgeCod,'') cAgeCod,  " _
'            & "         ISNULL(AG.cAgeDescripcion,'') cAgeDescripcion , Rend.cOpeCod , D2.cDocAbrev  as cDocAbrevSol , Rend.cCtaCod   " _
'            & " FROM    MOV M " _
'            & "         JOIN MOVCTA MC          ON MC.nMOVNRO = M.nMOVNRO " _
'            & "         LEFT JOIN MOVME ME      ON ME.nMOVNRO =MC.nMOVNRO AND MC.nMOVITEM = ME.nMOVITEM " _
'            & "         JOIN MOVREF MR          ON MR.nMOVNRO = M.nMOVNRO " _
'            & "         JOIN MOVARENDIR MAR     ON MAR.nMOVNRO=MR.nMOVNROREF " _
'            & "         JOIN PERSONA P          ON P.CPERSCOD  = MAR.CPERSCOD " _
'            & "         JOIN RRHH RH            ON RH.CPERSCOD  = P.CPERSCOD " _
'            & "         JOIN AREAS A            ON A.CAREACOD = RH.cAreaCodActual "
'
'    sql = sql + "       LEFT JOIN AGENCIAS AG   ON AG.CAGECOD = RH.cAgenciaActual " _
'            & "         JOIN MOVDOC MDAR        ON MDAR.nMOVNRO = MAR.nMOVNRO " _
'            & "         JOIN DOCUMENTO D2       ON D2.nDocTpo = MDAR.nDocTpo " _
'            & "         JOIN (  SELECT  O.cOpeCod, O.cOpeDesc, M1.cMOVNRO AS cMovRendicion, M1.nMOVNRO as nMovRendicion  , M1.CMOVDESC , MR1.nMOVNROREF as nMovAtenc , " _
'            & "                         ISNULL(D.CDOCABREV,'') AS cDocAbrevRend, ISNULL(MD.nDocTpo,'') as nDocTpoRend, MD.dDocFecha  AS dDocFechaRend, " _
'            & "                         ISNULL(MD.cDocNro,'') as cDocNroRen, SUM(ISNULL(Me.NMOVMEIMPORTE, MC.NMOVIMPORTE)) As MontoRendicion ,  ISNULL(MCP.cCtaCod,'') AS cCtaCod  " _
'            & "                 FROM    MOV M1  JOIN MOVREF MR1 ON M1.nMOVNRO=MR1.nMOVNRO " _
'            & "                         JOIN OPETPO O           ON O.COPECOD = M1.COPECOD " _
'            & "                         LEFT JOIN MOVCTA MC     ON MC.nMOVNRO = M1.nMOVNRO AND MC.CCTACONTCOD IN ('" & psCtaArendir & "','" & psCtaPendiente & "')  " _
'            & "                         LEFT JOIN MOVME ME      ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVNRO = MC.nMOVNRO " _
'            & "                         LEFT JOIN MOVDOC MD     ON MD.nMOVNRO = M1.nMOVNRO  AND MD.nDocTpo NOT IN ('" & TpoDocVoucherEgreso & "') " _
'            & "                         LEFT JOIN DOCUMENTO D   ON D.nDocTpo = MD.nDocTpo " _
'            & "                         LEFT JOIN MOVCAP MCP     ON MCP.NMOVNRO = M1.NMOVNRO "
'
'        sql = sql + "           WHERE   SUBSTRING(M1.COPECOD,3,1)='" & pnMoneda & "' AND M1.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
'            & "                         And SUBSTRING(M1.COPECOD,1,5) IN ('" & Mid(gCGArendirCtaRendMN, 1, 5) & "','" & Mid(gCGArendirCtaRendME, 1, 5) & "','" & Mid(gCGArendirCtaRendDevMN, 1, 5) & "','" & Mid(gCGArendirCtaRendDevME, 1, 5) & "','" & Mid(gCGArendirCtaRendDevMNAge, 1, 5) & "','" & Mid(gCGArendirCtaRendDevMEAge, 1, 5) & "','" & Mid(gCGArendirCtaRendMNAge, 1, 5) & "','" & Mid(gCGArendirCtaRendMEAge, 1, 5) & "') " _
'            & "                         AND SUBSTRING(M1.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " _
'            & "                 GROUP BY  O.cOpeCod,O.cOpeDesc, M1.cMOVNRO,M1.CMOVDESC , MR1.nMOVNROREF, D.CDOCABREV,MD.nDocTpo, MD.cDocNro, MD.dDocFecha , M1.nMOVNRO, MCP.cCtaCod) AS Rend " _
'            & "         ON Rend.nMovAtenc= M.nMOVNRO " _
'            & " WHERE   MAR.cTpoArendir ='" & pnTipoArendir & "'  " & lsFiltroAge & lsFiltroArea _
'            & "         AND MC.CCTACONTCOD IN ('" & psCtaArendir & "') AND SUBSTRING(M.COPECOD,3,1)='" & pnMoneda & "'  " _
'            & "         AND MAR.CAREACOD ='" & psAreaArendir & "' AND MAR.CAGECOD ='" & psAgeArendir & "' " _
'            & "         And M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
'            & " ORDER BY Rend.cMovRendicion, MDAR.cDocNro "
    If pnTipoArendir = gArendirTipoAgencias Then
         sql = "SELECT  Rend.cOpeDesc, Rend.cDocAbrevRend, Rend.cDocNroRen, ISNULL(CONVERT(CHAR(10),Rend.dDocFechaRend,103), SUBSTRING(Rend.cMovRendicion,7,2) + '/' + SUBSTRING(Rend.cMovRendicion,5,2) + '/' + SUBSTRING(Rend.cMovRendicion,1,4) ) AS dDocFechaRend, MDAR.cDocNro as cDocNroSol ,   " _
            & "         CONVERT(CHAR(10),MDAR.dDocFecha,103) AS dDocSolicitud, CONVERT(CHAR(50),P.CPERSNOMBRE) AS cPersNombre, ISNULL(Abs(Rend.MontoRendicion),0) AS MontoRend, ISNULL(Rend.MontoRendicion,0) AS MontoRendReal, " _
            & "         ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS NIMPORTE, MAR.nMovSaldo, ISNULL(A.cAreaDescripcion,'') cAreaDescripcion , " _
            & "         P.cPersCod, Rend.nMovRendicion, M.nMOVNRO as cMovNroAtenc, MAR.nMOVNRO as cMovNroSol, Rend.CMOVDESC, Rend.nDocTpoRend,  " _
            & "         MDAR.nDocTpo as nDocTpoSol,ISNULL(A.cAreaCod,'') cAreaCod ,ISNULL(AG.cAgeCod,'') cAgeCod,  " _
            & "         ISNULL(AG.cAgeDescripcion,'') cAgeDescripcion , Rend.cOpeCod , D2.cDocAbrev  as cDocAbrevSol , Rend.cCtaCod   " _
            & " FROM    MOV M " _
            & "         JOIN MOVCTA MC          ON MC.nMOVNRO = M.nMOVNRO " _
            & "         LEFT JOIN MOVME ME      ON ME.nMOVNRO =MC.nMOVNRO AND MC.nMOVITEM = ME.nMOVITEM " _
            & "         JOIN MOVREF MR          ON MR.nMOVNRO = M.nMOVNRO " _
            & "         JOIN MOVARENDIR MAR     ON MAR.nMOVNRO=MR.nMOVNROREF " _
            & "         JOIN PERSONA P          ON P.CPERSCOD  = MAR.CPERSCOD " _
            & "         JOIN RRHH RH            ON RH.CPERSCOD  = P.CPERSCOD " _
            & "         JOIN AREAS A            ON A.CAREACOD = RH.cAreaCodActual "

    sql = sql + "       LEFT JOIN AGENCIAS AG   ON AG.CAGECOD = RH.cAgenciaActual " _
            & "         JOIN MOVDOC MDAR        ON MDAR.nMOVNRO = MAR.nMOVNRO " _
            & "         JOIN DOCUMENTO D2       ON D2.nDocTpo = MDAR.nDocTpo " _
            & "         JOIN (  SELECT  O.cOpeCod, O.cOpeDesc, M1.cMOVNRO AS cMovRendicion, M1.nMOVNRO as nMovRendicion  , M1.CMOVDESC , MR1.nMOVNROREF as nMovAtenc , " _
            & "                         ISNULL(D.CDOCABREV,'') AS cDocAbrevRend, ISNULL(MD.nDocTpo,'') as nDocTpoRend, MD.dDocFecha  AS dDocFechaRend, " _
            & "                         ISNULL(MD.cDocNro,'') as cDocNroRen, SUM(ISNULL(Me.NMOVMEIMPORTE, MC.NMOVIMPORTE)) As MontoRendicion ,  ISNULL(MCP.cCtaCod,'') AS cCtaCod  " _
            & "                 FROM    MOV M1  JOIN MOVREF MR1 ON M1.nMOVNRO=MR1.nMOVNRO " _
            & "                         JOIN OPETPO O           ON O.COPECOD = M1.COPECOD " _
            & "                         LEFT JOIN MOVCTA MC     ON MC.nMOVNRO = M1.nMOVNRO AND MC.CCTACONTCOD IN ('" & psCtaArendir & "','" & psCtaPendiente & "')  " _
            & "                         LEFT JOIN MOVME ME      ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVNRO = MC.nMOVNRO " _
            & "                         LEFT JOIN MOVDOC MD     ON MD.nMOVNRO = M1.nMOVNRO  AND MD.nDocTpo NOT IN ('" & TpoDocVoucherEgreso & "') " _
            & "                         LEFT JOIN DOCUMENTO D   ON D.nDocTpo = MD.nDocTpo " _
            & "                         LEFT JOIN MOVCAP MCP     ON MCP.NMOVNRO = M1.NMOVNRO "

        sql = sql + "           WHERE   SUBSTRING(M1.COPECOD,3,1)='" & pnMoneda & "' AND M1.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
            & "                         And SUBSTRING(M1.COPECOD,1,5) IN ('" & Mid(gCGArendirCtaRendMN, 1, 5) & "','" & Mid(gCGArendirCtaRendME, 1, 5) & "','" & Mid(gCGArendirCtaRendDevMN, 1, 5) & "','" & Mid(gCGArendirCtaRendDevME, 1, 5) & "','" & Mid(gCGArendirCtaRendDevMNAge, 1, 5) & "','" & Mid(gCGArendirCtaRendDevMEAge, 1, 5) & "','" & Mid(gCGArendirCtaRendMNAge, 1, 5) & "','" & Mid(gCGArendirCtaRendMEAge, 1, 5) & "') " _
            & "                         AND SUBSTRING(M1.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " _
            & "                 GROUP BY  O.cOpeCod,O.cOpeDesc, M1.cMOVNRO,M1.CMOVDESC , MR1.nMOVNROREF, D.CDOCABREV,MD.nDocTpo, MD.cDocNro, MD.dDocFecha , M1.nMOVNRO, MCP.cCtaCod) AS Rend " _
            & "         ON Rend.nMovAtenc= M.nMOVNRO " _
            & " WHERE   MAR.cTpoArendir ='" & pnTipoArendir & "'  " & lsFiltroAge & lsFiltroArea _
            & "         AND MC.CCTACONTCOD IN ('" & psCtaArendir & "') AND SUBSTRING(M.COPECOD,3,1)='" & pnMoneda & "'  " _
            & "         AND MAR.CAREACOD ='" & psAreaArendir & "' AND MAR.CAGECOD ='" & psAgeArendir & "' " _
            & "         And M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
            & " ORDER BY Rend.cMovRendicion, MDAR.cDocNro "
    Else
        sql = "exec stp_sel_ObtenerRendicionesArendirCuentas '" & psAgeCod & "', '" & psAreaCod & "', '" & psAreaArendir & "', '" & psAgeArendir & "', " & pnTipoArendir & ", '" & psCtaArendir & "', '" & psCtaPendiente & "', '" & Format(pdDesde, "yyyyMMdd") & "', '" & Format(pdHasta, "yyyyMMdd") & "', '" & CStr(pnMoneda) & "', " & IIf(pbTodo, 1, 0) & ""
    End If
'***Fin Modificado por ELRO*******************************
    Case gArendirTipoViaticos
'***Modificado por ELRO el 20120427, según OYP-RFC005-2012
'        sql = "SELECT   Rend.cOpeDesc, Rend.cDocAbrevRend, Rend.cDocNroRen, ISNULL(CONVERT(CHAR(10),Rend.dDocFechaRend,103), " _
'            & "         SUBSTRING(Rend.cMovRendicion,7,2) + '/' + SUBSTRING(Rend.cMovRendicion,5,2) + '/' + SUBSTRING(Rend.cMovRendicion,1,4) ) AS dDocFechaRend, " _
'            & "         MDAR.cDocNro as cDocNroSol, CONVERT(CHAR(10),MDAR.dDocFecha,103) AS dDocSolicitud, " _
'            & "         CONVERT(CHAR(50),P.CPERSNOMBRE) AS cPersNombre, ISNULL(Abs(Rend.MontoRendicion),0) AS MontoRend, " _
'            & "         ISNULL(Rend.MontoRendicion,0) AS MontoRendReal,  0 AS NIMPORTE, " _
'            & "         MAR.nMovSaldo, ISNULL(A.cAreaDescripcion,'') cAreaDescripcion , " _
'            & "         P.cPersCod, Rend.NMovRendicion, M.NMOVNRO as cMovNroAtenc, MAR.NMOVNRO as NMovNroSol, " _
'            & "         Rend.CMOVDESC, Rend.nDocTpoRend,MDAR.nDocTpo as nDocTpoSol, " _
'            & "         ISNULL(A.cAreaCod,'') cAreaCod ,ISNULL(AG.cAgeCod,'') cAgeCod, " _
'            & "         ISNULL(AG.cAgeDescripcion,'') cAgeDescripcion , Rend.cOpeCod , D2.cDocAbrev  as cDocAbrevSol , Rend.cCtaCod " _
'            & " FROM    MOV M  " _
'            & "         JOIN MOVARENDIR MAR     ON M.NMOVNRO= MAR.NMOVNRO   " _
'            & "         JOIN PERSONA P          ON P.CPERSCOD  = MAR.CPERSCOD " _
'            & "         JOIN RRHH RH            ON RH.CPERSCOD  = P.CPERSCOD " _
'            & "         JOIN AREAS A            ON A.CAREACOD = RH.cAreaCodActual " _
'            & "         LEFT JOIN AGENCIAS AG   ON AG.CAGECOD = RH.cAgenciaActual " _
'            & "         JOIN MOVDOC MDAR        ON MDAR.NMOVNRO = M.NMOVNRO  AND MDAR.nDocTpo IN (" & TpoDocRecViaticosARendirCuenta & ") " _
'            & "         JOIN DOCUMENTO D2       ON D2.nDocTpo = MDAR.nDocTpo "
'
'        sql = sql + "         JOIN (  SELECT  O.cOpeCod, O.cOpeDesc, M1.NMOVNRO AS NMovRendicion, M1.CMOVNRO AS CMovRendicion, " _
'            & "                         M1.CMOVDESC , MR1.NMOVNROREF as NMovNroSol , " _
'            & "                         ISNULL(D.CDOCABREV,'') AS cDocAbrevRend, ISNULL(MD.nDocTpo,'') as nDocTpoRend, " _
'            & "                         MD.dDocFecha  AS dDocFechaRend, ISNULL(MD.cDocNro,'') as cDocNroRen, " _
'            & "                         SUM(ISNULL(Me.NMOVMEIMPORTE, ISNULL(MC.NMOVIMPORTE,0))) As MontoRendicion , ISNULL(MCP.cCtaCod,'') AS cCtaCod " _
'            & "                 FROM    MOV M1  JOIN MOVREF MR1 ON M1.NMOVNRO=MR1.NMOVNRO " _
'            & "                         JOIN OPETPO O           ON O.COPECOD = M1.COPECOD " _
'            & "                         LEFT JOIN MOVCTA MC     ON MC.NMOVNRO = M1.NMOVNRO AND MC.CCTACONTCOD IN ('" & psCtaArendir & "','" & psCtaPendiente & "') " _
'            & "                         LEFT JOIN MOVME ME      ON ME.NMOVNRO = MC.NMOVNRO AND ME.NMOVITEM = MC.NMOVITEM " _
'            & "                         LEFT JOIN MOVDOC MD     ON MD.NMOVNRO = M1.NMOVNRO  AND MD.nDocTpo NOT IN (" & TpoDocVoucherEgreso & ") " _
'            & "                         LEFT JOIN DOCUMENTO D   ON D.nDocTpo = MD.nDocTpo " _
'            & "                         LEFT JOIN MOVCAP MCP    ON MCP.NMOVNRO = M1.NMOVNRO " _
'            & "                 WHERE   SUBSTRING(M1.COPECOD,3,1)='" & pnMoneda & "' " _
'            & "                         AND M1.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "')  " _
'            & "                         And SUBSTRING(M1.COPECOD,1,5) IN ('" & Mid(gCGArendirViatRendMN, 1, 5) & "','" & Mid(gCGArendirViatRendME, 1, 5) & "','" & Mid(gCGArendirViatRendDevMN, 1, 5) & "','" & Mid(gCGArendirViatRendDevME, 1, 5) & "') " _
'            & "                         AND SUBSTRING(M1.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " _
'            & "                 GROUP BY  O.cOpeCod,O.cOpeDesc,M1.NMOVNRO, M1.CMOVNRO ,M1.CMOVDESC , MR1.NMOVNROREF, " _
'            & "                         D.CDOCABREV,MD.nDocTpo, MD.cDocNro, MD.dDocFecha, MCP.cCtaCod) AS Rend " _
'            & "          ON Rend.NMovNroSol= MAR.NMOVNRO " _
'            & "  WHERE   MAR.cTpoArendir ='" & pnTipoArendir & "' " & lsFiltroAge & lsFiltroArea _
'            & "          AND Substring(M.COPECOD,3,1)='" & pnMoneda & "' " _
'            & "          AND MAR.CAREACOD ='" & psAreaArendir & "' AND MAR.CAGECOD ='" & psAgeArendir & "' " _
'            & "          And M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "')  " _
'            & "  ORDER BY Rend.nMovRendicion, MDAR.cDocNro "
sql = "exec stp_sel_ObtenerRendicionesArendir '" & psAgeCod & "', '" & psAreaCod & "', '" & psCtaArendir & "', '" & psCtaPendiente & "', '" & Format(pdDesde, "yyyyMMdd") & "', '" & Format(pdHasta, "yyyyMMdd") & "', '" & CStr(pnMoneda) & "', '" & pnTipoArendir & "', " & IIf(pbTodo, 1, 0) & ""
'***Fin Modificado por ELRO*******************************
End Select
Set rs = oConect.CargaRecordSet(sql)
Set GetRendicionesArendir = rs
oConect.CierraConexion
Set oConect = Nothing
End Function
Public Function GetConceptosViaticos(ByVal psCategCod As ViaticosCateg, ByVal psDestinoCod As ViaticosDestino, _
                                    ByVal psTranspIdaCod As ViaticosTransporte, ByVal psTranspVueltaCod As ViaticosTransporte, _
                                    ByVal pnNumDias As Integer) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConec As DConecta
Set rs = New Recordset
Set oConec = New DConecta

If oConec.AbreConexion = False Then Exit Function

sql = "   SELECT    O.cObjetoCod, CONVERT(CHAR(40),cObjetoDesc) AS cObjetoDesc, 1 as Nivel ,  " _
    & "         ISNULL(NOTRANSP.ImporteViaticos,0) + CASE WHEN ISNULL(IDA.ImporteViaticos,0) > ISNULL(VUELTA.ImporteViaticos,0) THEN ISNULL(IDA.ImporteViaticos,0) ELSE ISNULL(VUELTA.ImporteViaticos,0) END  AS ImporteConcepto " _
    & " FROM    OBJETO O " _
    & "         Left Join " _
    & "         (Select  O.cObjetoCod, " _
    & "                  max(CASE " _
    & "                         WHEN cViaticoAfectoA = '" & gViaticoAfectoADias & "' THEN nViaticoImporte*" & pnNumDias & "" _
    & "                 WHEN cViaticoAfectoA IN ('" & gViaticoAfectoANinguno & "','" & gViaticoAfectoATransporte & "')  THEN nViaticoImporte " _
    & "                 END ) As ImporteViaticos, max(nViaticoImporte) nViaticoImporte " _
    & "          From    Viaticos V JOIN " & vsBaseComunes & "Objeto O on O.cObjetoCod = V.cObjetoCod  " _
    & "          Where   O.cObjetoCod Like '" & ObjConceptosARendir & "__' And cCategCod ='" & psCategCod & "' and cDestinoCod ='" & psDestinoCod & "' " _
    & "                  and cViaticoAfectoA NOT IN ('" & gViaticoAfectoATransporte & "') and cTranspCod IN ('0','" & psTranspIdaCod & "','" & psTranspVueltaCod & "') " _
    & "         GROUP BY  O.cObjetoCod  ) as NOTRANSP " _
    & "         ON O.COBJETOCOD = NOTRANSP.COBJETOCOD " _
    & "         Left Join " _
    & "         (Select O.cObjetoCod, " _
    & "                 CASE " _
    & "                     WHEN cViaticoAfectoA = '" & gViaticoAfectoADias & "' THEN nViaticoImporte*" & pnNumDias & " " _
    & "                     WHEN cViaticoAfectoA IN ('" & gViaticoAfectoANinguno & "','" & gViaticoAfectoATransporte & "')   THEN nViaticoImporte " _
    & "                 END As ImporteViaticos, nViaticoImporte " _
    & "          From    Viaticos V JOIN " & vsBaseComunes & "Objeto O on O.cObjetoCod = V.cObjetoCod " _
    & "          Where   O.cObjetoCod Like '" & ObjConceptosARendir & "__'  and cTranspCod='" & psTranspIdaCod & "' " _
    & "                  and cViaticoAfectoA IN ('" & gViaticoAfectoATransporte & "') And cCategCod ='" & psCategCod & "'  and cDestinoCod ='" & psDestinoCod & "' ) AS IDA " _
    & "          ON O.COBJETOCOD  = IDA.COBJETOCOD " _
    & "          Left Join "
sql = sql + "  (Select  O.cObjetoCod, " _
    & "                     CASE " _
    & "                             WHEN cViaticoAfectoA = '" & gViaticoAfectoADias & "'  THEN nViaticoImporte*" & pnNumDias & " " _
    & "                             WHEN cViaticoAfectoA IN ('" & gViaticoAfectoANinguno & "','" & gViaticoAfectoATransporte & "')   THEN nViaticoImporte " _
    & "                     END As ImporteViaticos, nViaticoImporte " _
    & "             From    Viaticos V JOIN " & vsBaseComunes & "Objeto O on O.cObjetoCod = V.cObjetoCod " _
    & "             Where   O.cObjetoCod Like '" & ObjConceptosARendir & "__'  and cTranspCod='" & psTranspVueltaCod & "' " _
    & "                     and cViaticoAfectoA IN ('" & gViaticoAfectoATransporte & "') And cCategCod ='" & psCategCod & "' and cDestinoCod='" & psDestinoCod & "'  ) AS VUELTA " _
    & "             ON  O.COBJETOCOD  = VUELTA.COBJETOCOD " _
    & " WHERE   O.COBJETOCOD LIKE '" & ObjConceptosARendir & "__'"
 
Set rs = oConec.CargaRecordSet(sql)
Set GetConceptosViaticos = rs
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GrabaSolicitudViaticos(ByVal psMovViatico As String, ByVal pdFecsis As Date, ByVal pnImporte As Currency, _
                                 ByVal psOpeCod As String, _
                                 ByVal psMovDesc As String, _
                                 ByVal psNroDocViatico As String, ByVal pdFechaViatico As Date, _
                                 ByVal psMovNroDocRef As String, ByVal pdFechaDocRef As Date, _
                                 ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal psPersCod As String, _
                                 ByVal psCategCod As ViaticosCateg, _
                                 ByVal rsDetViat As ADODB.Recordset, ByVal rsAux As ADODB.Recordset, _
                                 Optional pbAmpliacion As Boolean = False) As Integer
Dim oMov As DMov
Dim oCont As NContFunciones
Dim psDestinoCod As ViaticosDestino
Dim psTranspIda As ViaticosTransporte
Dim psTranspVuelta As ViaticosTransporte
Dim pdPartida As Date
Dim pdRetorno As Date 'PASI TI-ERS060 2014
Dim pnMovViaticosDias As String
Dim psDestinoDesc As String
Dim lsMovNro As String
Dim lsGlosa As String
Dim lsMovNroViat As String
Dim lnItem As Long
Dim lnImporte As Currency
Dim i As Integer

Dim lnMovNro As Long
Dim lnMovNroViat As Long
'***Agregado por ELRO el 20120404, según OYP-RFC005-2012
Dim lnMovNroViat2 As Long
'***Fin Agregado por ELRO***********************************
Dim lbTrans As Boolean

On Error GoTo GrabaSolicitudViaticosErr

GrabaSolicitudViaticos = 1
Set oMov = New DMov
Set oCont = New NContFunciones

oMov.BeginTrans
lbTrans = True
lsMovNroViat = psMovViatico
If pbAmpliacion = False Then
    oMov.InsertaMov lsMovNroViat, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
    lnMovNroViat = oMov.GetnMovNro(lsMovNroViat)
    oMov.InsertaMovCont lnMovNroViat, pnImporte, "0", "0"
    '***Agregado por ELRO el 20120926, según TIC1209270003
    psNroDocViatico = oMov.DevolverCorrelativoMovDoc(62)
    psNroDocViatico = Format(psNroDocViatico, String(8, "0"))
    '***Fin Agregado por ELRO el 20120926
    If psNroDocViatico <> "" Then
       'Grabamos en MovDoc el Recibo de Egresos de viaticos
       oMov.InsertaMovDoc lnMovNroViat, TpoDocRecViaticosARendirCuenta, psNroDocViatico, Format(pdFechaViatico, gsFormatoFecha)
    End If
    If psMovNroDocRef <> "" Then
        oMov.InsertaMovDoc lnMovNroViat, TpoDocCarta, psMovNroDocRef, Format(pdFechaDocRef, gsFormatoFecha)
    End If
    oMov.InsertaMovARendir lnMovNroViat, gArendirTipoViaticos, psAreaCod, psAgeCod, psPersCod, 0, psCategCod
Else
    '***Modificado por ELRO el 20120328, según OYP-RFC005-2012
    'lnMovNroViat = oMov.GetnMovNro(lsMovNroViat)
    If psOpeCod = gCGArendirViatSolcEditMN Or psOpeCod = gCGArendirViatSolcEditME Then
        lnMovNroViat = oMov.GetnMovNro(lsMovNroViat)
        Call oMov.actualizarMovCont(lnMovNroViat, pnImporte)
        Call oMov.actualizarMovDoc(lnMovNroViat, psMovNroDocRef)
        Call oMov.actualizarMovARendir(lnMovNroViat, psCategCod)
    Else
        lnMovNroViat = oMov.GetnMovNro(lsMovNroViat)
    End If
    '***Fin Modificado por ELRO***********************************

    
End If
Do While Not rsDetViat.EOF
    If rsDetViat!Importe > 0 And Trim(rsDetViat!cMovNroDet) = "" Then
        lsMovNro = oMov.GeneraMovNro(pdFecsis, Mid(lsMovNroViat, 18, 2), Right(lsMovNroViat, 4))
        lsGlosa = rsDetViat!Motivo
        oMov.InsertaMov lsMovNro, psOpeCod, lsGlosa, gMovEstContabNoContable, gMovFlagVigente
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        
        oMov.InsertaMovCont lnMovNro, rsDetViat!Importe, "0", "0"
        '***Modificado por ELRO el 20120925, según SATI INC1209250017 y INC1209260002
        'psDestinoCod = Val(Right(Trim(rsDetViat!Destino), 2))
        psDestinoCod = Val(Right(Trim(rsDetViat!Destino), 4))
        '***Fin Modificado por ELRO el 20120925**************************************
        psTranspIda = Val(Right(Trim(rsDetViat!Ida), 2))
        psTranspVuelta = Val(Right(Trim(rsDetViat!Vuelta), 2))
        pdPartida = rsDetViat!Partida
        pnMovViaticosDias = rsDetViat!Dias
        psDestinoDesc = rsDetViat!Lugar
        pdRetorno = rsDetViat!Retorno 'PASI TI-ERS060 2014
        
        oMov.InsertaMovViaticos lnMovNroViat, lnMovNro, psDestinoCod, psTranspIda, psTranspVuelta, pdPartida, pnMovViaticosDias, psDestinoDesc, pdRetorno 'pdRetorno Agregado por PASI TI-ERS060 2014
        
        lnItem = 0
        For i = 1 To rsAux.Fields.Count - 1
            lnImporte = rsAux.Fields(i).value
            If lnImporte > 0 Then
                lnItem = lnItem + 1
                oMov.InsertaMovObj lnMovNro, lnItem, 1, ObjConceptosARendir
                oMov.InsertaMovObjViaticos lnMovNro, lnItem, rsAux.Fields(i).Name, lnImporte
            End If
        Next
    '***Agregado por ELRO el 20120328, según OYP-RFC005-2012
    ElseIf rsDetViat!Importe > 0 And Trim(rsDetViat!cMovNroDet) <> "" And (psOpeCod = CStr(gCGArendirViatSolcEditMN) Or psOpeCod = CStr(gCGArendirViatSolcEditME)) Then
        lnMovNroViat2 = oMov.GetnMovNro(rsDetViat!cMovNroDet)
        Call oMov.actualizarMovCont(lnMovNroViat2, rsDetViat!Importe)
        Call oMov.actualizarMovViaticos(lnMovNroViat2, Trim(Right(rsDetViat!Destino, 4)), Trim(Right(rsDetViat!Ida, 4)), Trim(Right(rsDetViat!Vuelta, 4)), Format(rsDetViat!Partida, "yyyyMMdd"), rsDetViat!Dias, rsDetViat!Lugar)
        Call oMov.actualizarGlosaViaticos(lnMovNroViat2, rsDetViat!Motivo)
        For i = 1 To rsAux.Fields.Count - 1
            Call oMov.actualizarMovObjViaticos(lnMovNroViat2, rsAux.Fields(i).Name, rsAux.Fields(i).value)
        Next i
    '***Fin Agregado por ELRO***********************************
    End If
    rsDetViat.MoveNext
    rsAux.MoveNext
Loop
oMov.CommitTrans
lbTrans = False
GrabaSolicitudViaticos = 0
Set oMov = Nothing
Exit Function
GrabaSolicitudViaticosErr:
    lsMsgErr = Err.Description
    If lbTrans Then
        lbTrans = False
        oMov.RollbackTrans
    End If
    Err.Raise Err.Number, Err.Source, lsMsgErr
End Function
Public Function GetDatosViaticos(ByVal psNroDocViatico As String)
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function

sql = " SELECT  M.CMOVNRO, MAR.nMOVNRO, MD.CDOCNRO AS DOCVIAT, MD1.CDOCNRO AS NRODOCREF , MD1.DDOCFECHA , " _
    & "         P.CPERSCOD, P.CPERSNOMBRE, " _
    & "         A.CAREACOD, ISNULL(A.CAREADESCRIPCION,'') AS CAREADESCRIPCION , AG.CAGECOD, ISNULL(AG.cAgeDescripcion,'') cAgeDescripcion,  M.cOpeCod ,  " _
    & "         MAR.CAREACOD as cCodAreaArendir, ISNULL(A1.CAREADESCRIPCION,'') AS cAreaArendir, MAR.CAGECOD as cAgeCodArendir, ISNULL(AG1.cAgeDescripcion,'') cAgeArendir,  " _
    & "         ISNULL(C.cRHCargoCod,'') AS cRHCargoCod, ISNULL(C.cRHCargoDescripcion,'') AS cRHCargoDescripcion,  " _
    & "         ISNULL((SELECT cPersIDnro FROM PERSID WHERE cPersIDTpo ='" & gPersIdDNI & "' AND cPersCod = p.cPersCod ),'') AS DNI, " _
    & "         ISNULL((SELECT cPersIDnro FROM PERSID WHERE cPersIDTpo ='" & gPersIdRUC & "' AND cPersCod = p.cPersCod ),'') AS RUC  " _
    & " FROM    MOVARENDIR MAR " _
    & "         LEFT JOIN " & vsBaseComunes & "AREAS A1       ON A1.cAreaCod = MAR.cAreaCod  " _
    & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG1   ON AG1.cAgeCod = MAR.cAgeCod " _
    & "         JOIN MOV    M           ON M.nMOVNRO =  MAR.nMOVNRO " _
    & "         JOIN MOVDOC MD          ON MD.nMOVNRO = MAR.nMOVNRO AND MD.nDocTpo =" & TpoDocRecViaticosARendirCuenta & " " _
    & "         JOIN MOVDOC MD1         ON MD1.nMOVNRO = MAR.nMOVNRO  AND MD1.nDocTpo<>" & TpoDocRecViaticosARendirCuenta & " " _
    & "         JOIN " & vsBasePesonas & "PERSONA P          ON P.CPERSCOD = MAR.CPERSCOD " _
    & "         JOIN RRHH                       RH           ON RH.CPERSCOD = P.CPERSCOD " _
    & "         LEFT JOIN " & vsBaseComunes & "AREAS A       ON A.cAreaCod = RH.cAreaCodActual  " _
    & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG   ON AG.cAgeCod = RH.cAgenciaActual " _
    & "         LEFT JOIN ( SELECT  dRHCargoFecha as FechaCargo, C.cPersCod, " _
    & "                             C.cRHCargoCod, CT.cRHCargoDescripcion " _
    & "                     FROM    RHCargos C  JOIN RHCargosTabla CT  ON CT.cRHCargoCod = C.cRHCargoCod " _
    & "                     WHERE   dRHCargoFecha = (SELECT MaX(dRHCargoFecha) FROM RHCargos WHERE cPersCod = c.cPersCod)  ) AS C  ON C.CPERSCOD = P.CPERSCOD "
sql = sql + " WHERE   MAR.cTpoArendir ='" & gArendirTipoViaticos & "' " _
    & "         AND M.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
    & "         AND MD.CDOCNRO ='" & psNroDocViatico & "' " _
    & "         AND NOT EXISTS (SELECT  M1.nMOVNRO " _
    & "                         FROM    MOV M1 " _
    & "                                 JOIN MOVREF MR ON MR.nMOVNRO = M1.nMOVNRO " _
    & "                         WHERE   MR.nMOVNROREF=M.nMOVNRO  AND SUBSTRING(M1.COPECOD,1,5) IN ('" & Mid(gCGArendirViatSustMN, 1, 5) & "','" & Mid(gCGArendirViatSustME, 1, 5) & "') " _
    & "                                 AND M1.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') )"


Set rs = oCon.CargaRecordSet(sql)
Set GetDatosViaticos = rs
oCon.CierraConexion
Set oCon = Nothing

End Function
Public Function GetDetalleViaticos(ByVal psMovnroViaticos As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function
'***Modificado por ELRO el 20120328, según OYP-RFC005-2012
'sql = " SELECT  M.CMOVNRO, MV.nMOVNRO, M.cMovDesc AS Motivo , MV.cDestinoDesc as Lugar,  " _
'    & "         dPartida, nMovViaticosDias, " _
'    & "         C.nConsValor as CodCat,  C.cConsDescripcion  as Categoria , " _
'    & "         C1.nConsValor as CodDestino,  C1.cConsDescripcion  as Destino , " _
'    & "         C2.nConsValor as CodTrasnIda,  C2.cConsDescripcion  as TransIda,  " _
'    & "         C3.nConsValor as CodTransVuelta,  C3.cConsDescripcion  as TrasnVuelta , MC.nMovMonto , M.cOpeCod, ISNULL(Ref.nMovNro,0) nMovNroAtencion " _
'    & " FROM    MOVVIATICOS MV  JOIN MOVARENDIR MAR ON MAR.nMOVNRO = MV.nVIATICOMOVNRO " _
'    & "         JOIN MOV    M   ON M.nMOVNRO = MV.nMovNro " _
'    & "         LEFT JOIN MOVCONT    MC  ON MC.nMOVNRO =M.nMOVNRO " _
'    & "         JOIN " & vsBaseComunes & "CONSTANTE C    ON C.nConsValor  = MAR.cCategCod " _
'    & "         JOIN " & vsBaseComunes & "CONSTANTE C1   ON C1.nConsValor  = MV.cDestinoCod " _
'    & "         JOIN " & vsBaseComunes & "CONSTANTE C2   ON C2.nConsValor  = MV.cTranspIda " _
'    & "         JOIN " & vsBaseComunes & "CONSTANTE C3   ON C3.nConsValor  = MV.cTranspVuelta " _
'    & "         LEFT JOIN (SELECT m1.nMovNro, mr1.nMovNroRef FROM Mov m1 JOIN MovRef mr1 ON mr1.nMovNro = m1.nMovNro " _
'    & "                    WHERE m1.nMovEstado = " & gMovEstContabMovContable & " and not nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
'    & "                   ) Ref ON Ref.nMovNroRef = M.nMovNro " _
'    & " WHERE   M.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
'    & "         AND C.nCONSCOD like '" & gViaticosCateg & "'  AND C1.nCONSCOD like '" & gViaticosDestino & "' " _
'    & "         AND C2.nCONSCOD like '" & gViaticosTransporte & "' AND C3.nCONSCOD like '" & gViaticosTransporte & "' " _
'    & "         AND MV.nViaticoMovNro = '" & psMovnroViaticos & "' "
sql = "exec stp_sel_ObtenerDetalleViaticos " & CLng(psMovnroViaticos)
'***Fin Modificado por ELRO***********************************
Set rs = oCon.CargaRecordSet(sql)
Set GetDetalleViaticos = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetCostosViaticos(ByVal psMovNroDetViat As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function
'***Modificado por ELRO el 20120328, según OYP-RFC005-2012
'sql = " SELECT  M1.CMOVNRO, MC.cObjetoCod , ISNULL(MC.nMovImporte,0 ) as Importe " _
'    & " FROM    MOVVIATICOS     MV " _
'    & "         JOIN MOV    M   ON M.nMOVNRO = MV.nViaticoMovNro " _
'    & "         JOIN MOV        M1  ON M1.nMOVNRO = MV.NMOVNRO " _
'    & "         JOIN MOVOBJVIATICOS     MC  ON MC.nMOVNRO = M1.NMOVNRO " _
'    & "         " _
'    & "          " _
'    & " WHERE   M1.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') AND  " _
'    & "         M1.CMOVNRO = '" & psMovNroDetViat & "'  "
sql = "exec stp_sel_ObtenerCostosViaticos  '" & psMovNroDetViat & "' "
'***Fin Modificado por ELRO***********************************
Set rs = oCon.CargaRecordSet(sql)
Set GetCostosViaticos = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetDatosTotalesViaticos(ByVal psMovnroViatico As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta

sql = " SELECT  M.CMOVNRO as MovViatico, MD.CDOCNRO AS DOCVIAT, MD1.CDOCNRO AS NRODOCREF , MD1.DDOCFECHA , M.cOpeCod ,  " _
    & "         P.CPERSCOD, P.CPERSNOMBRE, " _
    & "         ISNULL((SELECT cPersIDnro FROM PERSID WHERE cPersIDTpo ='" & gPersIdDNI & "' AND cPersCod = p.cPersCod ),'') AS DNI, " _
    & "         ISNULL((SELECT cPersIDnro FROM PERSID WHERE cPersIDTpo ='" & gPersIdRUC & "' AND cPersCod = p.cPersCod ),'') AS RUC,  " _
    & "         A.CAREACOD, ISNULL(A.CAREADESCRIPCION,'') AS CAREADESCRIPCION , AG.CAGECOD, ISNULL(AG.cAgeDescripcion,'') cAgeDescripcion, " _
    & "         A1.CAREACOD as cCodAreaArendir, ISNULL(A1.CAREADESCRIPCION,'') AS cAreaARendir , IsNull(MAR.CAGECOD,'') as cCodAgeArendir, ISNULL(AG1.cAgeDescripcion,'') cAgeArendir ,  " _
    & "         CT.nConsValor as CodCat,  CT.cConsDescripcion  as Categoria,   " _
    & "         ISNULL(C.cRHCargoCod,'') AS cRHCargoCod, ISNULL(C.cRHCargoDescripcion,'') AS cRHCargoDescripcion, " _
    & "         DV.CMOVNRO, DV.NMOVNRO , DV.Motivo, DV.Lugar , DV.dPartida, DV.nMovViaticosDias, " _
    & "         DV.CodDestino, DV.Destino , DV.CodTrasnIda , DV.TransIda, DV.CodTransVuelta, " _
    & "         DV.TrasnVuelta , DV.nMovMonto " _
    & " FROM    MOVARENDIR MAR  " _
    & "         JOIN MOV    M           ON M.nMOVNRO =  MAR.nMOVNRO " _
    & "         JOIN MOVDOC MD          ON MD.nMOVNRO = MAR.nMOVNRO AND MD.nDocTpo =" & TpoDocRecViaticosARendirCuenta & " " _
    & "         LEFT JOIN MOVDOC MD1         ON MD1.nMOVNRO = MAR.nMOVNRO  AND MD1.nDocTpo<>" & TpoDocRecViaticosARendirCuenta & " " _
    & "         JOIN " & vsBasePesonas & "PERSONA P          ON P.CPERSCOD = MAR.CPERSCOD JOIN CONSTANTE CT  ON CT.nConsValor  = MAR.cCategCod " _
    & "         JOIN RRHH RH            ON RH.CPERSCOD = P.CPERSCOD  " _
    & "         LEFT JOIN " & vsBaseComunes & "AREAS A       ON A.cAreaCod = RH.cAreaCodActual  " _
    & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG   ON AG.cAgeCod = RH.cAgenciaActual " _
    & "         LEFT JOIN " & vsBaseComunes & "AREAS A1      ON A1.cAreaCod = MAR.cAreaCod " _
    & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG1  ON AG1.cAgeCod = MAR.cAgeCod "

sql = sql + "   LEFT JOIN ( SELECT      dRHCargoFecha as FechaCargo,C.cPersCod, C.cRHCargoCod, CT.cRHCargoDescripcion " _
    & "                     FROM        RHCargos C " _
    & "                                 JOIN RHCargosTabla CT  ON CT.cRHCargoCod = C.cRHCargoCod " _
    & "                     WHERE dRHCargoFecha = (SELECT MaX(dRHCargoFecha) FROM RHCargos WHERE cPersCod = c.cPersCod) ) AS C " _
    & "                     ON C.CPERSCOD = P.CPERSCOD " _
    & "         JOIN    (SELECT     MV.nViaticoMovNro , M.CMOVNRO,   MV.nMOVNRO , M.cMovDesc AS Motivo , MV.cDestinoDesc as Lugar, " _
    & "                             dPartida, nMovViaticosDias, " _
    & "                             C1.nConsValor as CodDestino,  C1.cConsDescripcion  as Destino , " _
    & "                             C2.nConsValor as CodTrasnIda,  C2.cConsDescripcion  as TransIda, " _
    & "                             C3.nConsValor as CodTransVuelta,  C3.cConsDescripcion  as TrasnVuelta , ISNULL(MC.nMovMonto ,0)  AS nMovMonto " _
    & "                  FROM       MOVVIATICOS  MV " _
    & "                             JOIN MOV            M       ON M.nMOVNRO = MV.NMovNro " _
    & "                             LEFT JOIN MOVCONT   MC      ON MC.nMOVNRO =M.nMOVNRO " _
    & "                             JOIN " & vsBaseComunes & "CONSTANTE C1           ON C1.nConsValor  = MV.cDestinoCod " _
    & "                             JOIN " & vsBaseComunes & "CONSTANTE C2           ON C2.nConsValor  = MV.cTranspIda " _
    & "                             JOIN " & vsBaseComunes & "CONSTANTE C3           ON C3.nConsValor  = MV.cTranspVuelta "
    
sql = sql + "           WHERE      M.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagExtornado & "','" & gMovFlagModificado & "') " _
    & "                            AND C1.nCONSCOD like '" & gViaticosDestino & "' " _
    & "                            AND C2.nCONSCOD like '" & gViaticosTransporte & "' AND C3.nCONSCOD like '" & gViaticosTransporte & "' )  AS DV " _
    & "         ON DV.nViaticoMovNro = MAR.nMOVNRO " _
    & " WHERE   MAR.cTpoArendir ='" & gArendirTipoViaticos & "' AND M.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagExtornado & "','" & gMovFlagModificado & "') " _
    & "         AND M.CMOVNRO ='" & psMovnroViatico & "' AND CT.nCONSCOD like '" & gViaticosCateg & "' " _
    & "         AND NOT EXISTS (SELECT  M1.nMOVNRO " _
    & "                         FROM    MOV M1 " _
    & "                                 JOIN MOVREF MR ON MR.nMOVNRO = M1.nMOVNRO " _
    & "                        WHERE   MR.nMOVNROREF=M.nMOVNRO  AND SUBSTRING(M.COPECOD,1,5) IN ('" & Mid(gCGArendirViatSustMN, 1, 5) & "','" & Mid(gCGArendirViatSustME, 1, 5) & "') " _
    & "                                AND M1.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagExtornado & "', '" & gMovFlagModificado & "')) "

If oCon.AbreConexion = False Then Exit Function
Set rs = oCon.CargaRecordSet(sql)
Set GetDatosTotalesViaticos = rs

oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetDetalleCostosViaticos(ByVal psMovNroDetViat As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Set rs = New ADODB.Recordset

sql = "SELECT   O.cObjetoCod as Concepto, CONVERT(CHAR(40),O.cObjetoDesc) AS Descripcion , ISNULL(MO.Importe,0) AS IMPORTE " _
    & " FROM    OBJETO O " _
    & "         Left Join " _
    & "         (SELECT  M1.CMOVNRO, MO.cObjetoCod , ISNULL(MO.nMovImporte,0 ) as Importe " _
    & "          FROM    MOVVIATICOS     MV " _
    & "                  JOIN MOV            M   ON M.NMOVNRO = MV.NViaticoMovNro " _
    & "                  JOIN MOV            M1  ON M1.NMOVNRO = MV.NMOVNRO " _
    & "                  JOIN MOVOBJVIATICOS MO  ON MO.NMOVNRO = M1.NMOVNRO  " _
    & "                   " _
    & "                   " _
    & "         WHERE   M1.CMOVNRO = '" & psMovNroDetViat & "' AND M1.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagExtornado & "','" & gMovFlagModificado & "')) AS MO " _
    & "         ON MO.COBJETOCOD= O.COBJETOCOD " _
    & " WHERE O.COBJETOCOD LIKE '" & ObjConceptosARendir & "__'"

If oCon.AbreConexion = False Then Exit Function
Set rs = oCon.CargaRecordSet(sql)
Set GetDetalleCostosViaticos = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function EmiteCajasChicasNuevas(Optional pbOrganigrama As Boolean = False) As ADODB.Recordset
On Error GoTo ErrorEmiteCajasChicasNuevas
Dim rs As ADODB.Recordset
Dim sSQL As String
Set rs = New ADODB.Recordset

Dim oConect As DConecta
Set oConect = New DConecta

If oConect.AbreConexion = False Then Exit Function


    If pbOrganigrama Then
        sSQL = " SELECT  AA.CAREACOD +  AA.CAGECOD AS CODIGO, A.CAREADESCRIPCION + ' ' + ISNULL(AG.CAGEDESCRIPCION,'') AS AREADESCRIPCION , LEN(A.cAreaEstruc) AS NIVEL " _
            & "  FROM    " & vsBaseComunes & "AREAAGENCIA AA " _
            & "          LEFT JOIN ( SELECT     CH.CAREACOD , CH.CAGECOD, ISNULL(AG.CAGEDESCRIPCION ,A.CAREADESCRIPCION ) as Descripcion " _
            & "                      FROM       CAJACHICA CH JOIN AREAAGENCIA AA ON AA.CAREACOD = CH.CAREACOD " _
            & "                                 JOIN " & vsBaseComunes & "AREAS A ON A.CAREACOD = CH.CAREACOD " _
            & "                                 LEFT JOIN " & vsBaseComunes & "AGENCIAS AG ON AG.CAGECOD = CH.CAGECOD " _
            & "                      GROUP BY CH.CAREACOD, CH.CAGECOD , A.CAREADESCRIPCION, AG.CAGEDESCRIPCION ) AS CH " _
            & "         ON AA.CAREACOD = CH.CAREACOD AND AA.CAGECOD = CH.CAGECOD " _
            & "         JOIN " & vsBaseComunes & "AREAS A ON A.CAREACOD=AA.CAREACOD " _
            & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG ON AG.CAGECOD = AA.CAGECOD " _
            & "     Where CH.CAREACOD Is Null ORDER BY AA.CAREACOD +  AA.CAGECOD  "
    Else
        sSQL = " SELECT Codigo , Descripcion AREADESCRIPCION, Nivel " _
            & " From  (   Select    AA.cAreaCod as  Codigo , A.nAgrupacion EST," _
            & "                     A.cAreaDescripcion as Descripcion , " _
            & "                     Len(AA.cAreaCod) As Nivel, AA.nArendir " _
            & "            From     AreaAgencia AA " _
            & "                     JOIN " & vsBaseComunes & "Areas A on A.cAreaCod = AA.cAreaCod " _
            & "            Union " _
            & "            Select   AA.cAreaCod  + AA.cAgecod as  Codigo , ARE.nAgrupacion EST," _
            & "                     AG.cAgeDescripcion AS Descripcion, 10 As Nivel, AA.nArendir " _
            & "            From     " & vsBaseComunes & "AreaAgencia AA " _
            & "                     JOIN " & vsBaseComunes & "Agencias AG on AA.cAgecod = Ag.cAgecod  JOIN " & vsBaseComunes & "Areas ARE on ARE.cAreaCod = AA.cAreacod" _
            & "              Union" _
            & "              Select  right('0000' + rtrim(ltrim(str(CO.nConsValor))),4) Codigo , rtrim(ltrim(str(CO.nConsValor))) EST," _
            & "                      CO.cConsDescripcion AS Descripcion, 1 As Nivel, 0 nArendir " _
            & "              From     Constante CO" _
            & "         Where nConsCod Like '1043' And nConsCod <> nConsValor" _
            & " ) AS AgeArea " _
            & "          LEFT JOIN ( SELECT     CH.CAREACOD , CH.CAGECOD  " _
            & "                      FROM       CAJACHICA CH JOIN AREAAGENCIA AA ON AA.CAREACOD = CH.CAREACOD " _
            & "                                 JOIN " & vsBaseComunes & "AREAS A ON A.CAREACOD = CH.CAREACOD " _
            & "                                 LEFT JOIN " & vsBaseComunes & "AGENCIAS AG ON AG.CAGECOD = CH.CAGECOD " _
            & "                      GROUP BY CH.CAREACOD, CH.CAGECOD , A.CAREADESCRIPCION, AG.CAGEDESCRIPCION ) AS CH " _
            & "         ON AgeArea.Codigo = CH.CAREACOD + CH.CAGECOD " _
            & "     Where CH.CAREACOD Is Null " _
            & " GROUP BY Codigo, Descripcion, Nivel, EST Order by EST , Codigo, Nivel "
    End If


Set rs = oConect.CargaRecordSet(sSQL)
Set EmiteCajasChicasNuevas = rs

oConect.CierraConexion
Set oConect = Nothing
Exit Function
ErrorEmiteCajasChicasNuevas:
    Err.Raise vbObjectError + 100, "EmiteCajasChicasNuevas", "Error emitir Cajas Chicas Nuevas"
End Function
Public Function CapAbonoCuentaAhoMov(ByVal psFormatoFecha As String, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", Optional sCodIF As String = "", _
            Optional dFechaValor As Date, Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
            Optional ByVal pbMovContable As Boolean = False, Optional ByVal pnTipoArendir As ArendirTipo = gArendirTipoCajaChica, _
            Optional ByVal psCtaDebe As String = "", Optional ByVal psCtaHaber As String = "", _
            Optional ByVal pnMovNroAtenc As Long, Optional ByVal pnMovNroSol As Long, _
            Optional ByVal pdFecsis As Date, Optional pbImprimeBoleta As Boolean = False, _
            Optional psNomAge As String = "", Optional ByVal psCodUser As String = "") As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nSaldoInac As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, nExtracto As Long
Dim sMsgOpe As String
Dim nMovNro As Long
Dim lnMovItem As Integer
Dim lnMovOrden  As Integer

bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, True, False)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing
Dim oMov As DMov
Set oMov = New DMov
'Inicia la transaccion
Set clsCap = New DCapMovimientos

On Error GoTo ErrGraba
Dim oCap As NCapMovimientos
Set oCap = New NCapMovimientos

oMov.BeginTrans
bTrans = True
If sGlosa = "" Then sGlosa = "Abono Cuenta = " & sCuenta
If bActivaCta Then
    If bInactiva Then
        nSaldoInac = nSaldoCnt
    End If
End If
'Calcula intereses
nDiasTranscurridos = DateDiff("d", dUltMov, pdFecsis) - 1
nIntGanado = oCap.GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, pdFecsis)
nSaldoCnt = nSaldoCnt + nMonto
oMov.InsertaMov sMovNro, nOperacion, sGlosa
nMovNro = oMov.GetnMovNro(sMovNro)
If sNroDoc = "" Then ' Si la operacion es en efectivo
    oMov.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
    nSaldoDisp = nSaldoDisp + nMonto
    oMov.ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
    sMsgOpe = "Depósito Efectivo"
Else
    If nTipoDoc = TpoDocCheque Then 'Si el abono es con cheque
        nSaldoDisp = nSaldoDisp + 0
        nIntGanado = 0
        oMov.ActualizaAbonoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
        oMov.AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMovNro
        sMsgOpe = "Depósito Cheque"
    ElseIf nTipoDoc = TpoDocNotaAbono Then 'Si el abono es con nota de abono
        oMov.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
        nSaldoDisp = nSaldoDisp + nMonto
        oMov.ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
        sMsgOpe = "Depósito Nota Abono"
    End If
End If
oMov.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
If bInactiva Then
    oMov.AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
    oMov.ActualizaEstadoCuenta sCuenta, gCapEstActiva, pdFecsis, sMovNro
End If
oMov.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
oMov.UltimaActualizacionCuenta sCuenta, sMovNro

'graba el movimiento realizado por caja general en el arendir cuenta
oMov.InsertaMovCont nMovNro, nMonto, "0", "0"
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1: lnMovOrden = lnMovOrden + 1
oMov.InsertaMovCta nMovNro, Format(lnMovItem, "#0"), psCtaDebe, Abs(nMonto)
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oMov.InsertaMovCta nMovNro, Format(lnMovItem, "#0"), psCtaHaber, Abs(nMonto) * -1
oMov.InsertaMovRef nMovNro, pnMovNroAtenc
If nTipoDoc <> TpoDocCheque Then
    oMov.InsertaMovDoc nMovNro, nTipoDoc, sNroDoc, Format(pdFecsis, gsFormatoFecha)
End If
If nTipoDoc = TpoDocNotaAbono Or nTipoDoc = TpoDocNotaCargo Then
    oMov.ActualizaNotaAbonoCargo nTipoDoc, sNroDoc, gNCNAEnMovimiento, , Abs(nMonto)
End If
ActualizaMovSust pnMovNroSol, sMovNro, oMov, , pnTipoArendir, , nMovNro
oMov.ActualizaMovArendir pnMovNroSol, pnTipoArendir, , , , 0

oMov.CommitTrans
'clsCap.dbCmact.CommitTrans
CapAbonoCuentaAhoMov = nSaldoCnt
If pbImprimeBoleta Then
    If sNroDoc = "" Then
        EmiteBoleta "RETIRO AHORROS", sMsgOpe, sCuenta, nMonto, nOperacion, nSaldoDisp, nSaldoCnt, nIntGanado, nExtracto, , , , , , pdFecsis, psNomAge, psCodUser
    Else
        EmiteBoleta "RETIRO AHORROS", sMsgOpe, sCuenta, nMonto, nOperacion, nSaldoDisp, nSaldoCnt, nIntGanado, nExtracto, True, nTipoDoc, sNroDoc, , , pdFecsis, psNomAge
    End If
End If
Exit Function
ErrGraba:
    If bTrans Then
        oMov.RollbackTrans ' clsCap.dbCmact.RollbackTrans
        bTrans = False
    End If
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapAbonoCuentaAhoMov = 0
End Function
Public Function CapCargoCuentaAhoMov(ByVal psFormatoFecha As String, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
            Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional bOPExiste As Boolean = False, _
            Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
            Optional bExtorno As Boolean = False, _
            Optional ByVal pbMovContable As Boolean = False, Optional ByVal pnTipoArendir As ArendirTipo = gArendirTipoCajaChica, _
            Optional ByVal psCtaDebe As String = "", Optional ByVal psCtaHaber As String = "", _
            Optional ByVal pnMovNroAtenc As Long, Optional ByVal pnMovNroSol As Long, _
            Optional ByVal pdFecsis As Date, Optional pbImprimeBoleta As Boolean = False, _
            Optional psNomAge As String = "", Optional ByVal psCodUser As String = "") As Double

Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nSaldoInac As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, i As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String
Dim oMov As DMov
Dim lnMovItem As Integer
Dim lnMovOrden  As Integer
Dim oCap As NCapMovimientos
Set oCap = New NCapMovimientos

Set oMov = New DMov
bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
Randomize
For i = 0 To Rnd(2000) * 1000
Next i
If Not oCap.ValidaSaldoCuenta(sCuenta, nMonto) Then  'Valida el saldo de la cuenta nuevamente
    clsCap.dbCmact.RollbackTrans
    bTrans = False
    Set clsCap = Nothing
    Err.Raise 1000, "CapCargoCuentaAho", "Cuenta NO Posee Saldo Suficiente"
    CapCargoCuentaAhoMov = 0
    Exit Function
End If
'Inicia la transaccion
oMov.BeginTrans
bTrans = True
If bExtorno Then
    sGlosa = sGlosa & " Extorno Abono Cuenta = " & sCuenta
Else
    sGlosa = sGlosa & " Cargo Cuenta = " & sCuenta
End If
If bActivaCta Then
    If bInactiva Then
        nSaldoInac = nSaldoCnt
    End If
End If
'Calcula los intereses
nDiasTranscurridos = DateDiff("d", dUltMov, pdFecsis) - 1
nIntGanado = oCap.GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, pdFecsis)
nSaldoCnt = nSaldoCnt - nMonto
oMov.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
nSaldoDisp = nSaldoDisp - nMonto
oMov.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
If sNroDoc <> "" Then
    If nTipoDoc = TpoDocOrdenPago Then 'Si es con orden de pago
        If bOPExiste Then
            oMov.AgregaOrdenPagoEstado sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstCobrada
        Else
            oMov.AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMonto, gCapOPEstCobrada
        End If
        sMsgOpe = "Retiro OP "
    ElseIf nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = "Retiro NC"
    End If
Else
    sMsgOpe = "Retiro Efectivo"
End If
oMov.InsertaMov sMovNro, nOperacion, sGlosa
nMovNro = oMov.GetnMovNro(sMovNro)
oMov.AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
If bInactiva Then
    oMov.AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
    oMov.ActualizaEstadoCuenta sCuenta, gCapEstActiva, pdFecsis, sMovNro
End If
oMov.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
oMov.UltimaActualizacionCuenta sCuenta, sMovNro
'graba el movimiento realizado por caja general en el arendir cuenta
oMov.InsertaMovCont nMovNro, nMonto, "0", "0"
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1: lnMovOrden = lnMovOrden + 1

oMov.InsertaMovCta nMovNro, Format(lnMovItem, "#0"), psCtaDebe, Abs(nMonto)
lnMovItem = lnMovItem + 1: lnMovOrden = 0
oMov.InsertaMovCta nMovNro, Format(lnMovItem, "#0"), psCtaHaber, Abs(nMonto) * -1

'VER INSERTA DOCUMETNO PUESTO QUE NO LO REALIZA EN PARTE DE CAPTACIONES
If nTipoDoc = TpoDocOrdenPago Then 'Si es con orden de pago
    If bOPExiste Then
        oMov.InsertaMovDoc nMovNro, nTipoDoc, sNroDoc, Format(pdFecsis, gsFormatoFecha)
    End If
Else
    oMov.InsertaMovDoc nMovNro, nTipoDoc, sNroDoc, Format(pdFecsis, gsFormatoFecha)
End If
If nTipoDoc = TpoDocNotaAbono Or nTipoDoc = TpoDocNotaCargo Then
    oMov.ActualizaNotaAbonoCargo nTipoDoc, sNroDoc, gNCNAEnMovimiento, , Abs(nMonto)
End If
oMov.InsertaMovRef nMovNro, pnMovNroAtenc
ActualizaMovSust pnMovNroSol, sMovNro, oMov, , pnTipoArendir, , nMovNro
oMov.ActualizaMovArendir pnMovNroSol, pnTipoArendir, , , , 0

oMov.CommitTrans
Set oMov = Nothing
CapCargoCuentaAhoMov = nSaldoCnt
If pbImprimeBoleta Then
    If sNroDoc = "" Then
        EmiteBoleta "RETIRO AHORROS", sMsgOpe, sCuenta, nMonto, nOperacion, nSaldoDisp, nSaldoCnt, nIntGanado, nExtracto, , , , , , pdFecsis, psNomAge, psCodUser
    Else
        EmiteBoleta "RETIRO AHORROS", sMsgOpe, sCuenta, nMonto, nOperacion, nSaldoDisp, nSaldoCnt, nIntGanado, nExtracto, True, nTipoDoc, sNroDoc, , , pdFecsis, psNomAge
    End If
End If
Exit Function
ErrGraba:
    If bTrans Then
        oMov.RollbackTrans
        Set oMov = Nothing
        bTrans = False
    End If
    Err.Raise Err.Number, "", Err.Description
    CapCargoCuentaAhoMov = 0
End Function

Public Function CapExtornoAbonoAhoMov(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", _
            Optional pbMovContable As Boolean = False, Optional ByVal pnMovNroAnt As Long, Optional ByVal pnMovNroSol As Long, _
            Optional ByVal pnTipoArendir As ArendirTipo, Optional ByVal pnSaldoArendir As Currency, _
            Optional ByVal pnArendirFase As ARendirFases, Optional pdFecsis As Date, Optional pbImprimeBoleta As Boolean = False, _
            Optional psNomAge As String = "", Optional ByVal psCodUser As String = "") As Double

Dim nOpeExt As CaptacOperacion
Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, i As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String
Dim oMov As DMov
Set oMov = New DMov
bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
Set clsMant = Nothing

nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing
nOpeExt = nOperacion
Dim oCap As NCapMovimientos
Set oCap = New NCapMovimientos

Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
Randomize
For i = 0 To Rnd(2000) * 1000
Next i
If Not oCap.ValidaSaldoCuenta(sCuenta, nMonto) Then
    Err.Raise 1000, "CapExtornoAbonoAho", "Cuenta NO Posee Saldo Suficiente"
    CapExtornoAbonoAhoMov = 0
    Exit Function
End If
'Inicia la transaccion
oMov.BeginTrans
bTrans = True
sGlosa = sGlosa & "Extorno Abono Cuenta = " & sCuenta
nDiasTranscurridos = DateDiff("d", dUltMov, pdFecsis) - 1
nIntGanado = oCap.GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, pdFecsis)
nSaldoCnt = nSaldoCnt - nMonto
oMov.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp

If sNroDoc <> "" Then
    If nTipoDoc = TpoDocCheque Then
        nSaldoDisp = nSaldoDisp - 0
        oMov.ActualizaCargoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
        sMsgOpe = "Ext. Depósito Chq."
    ElseIf nTipoDoc = TpoDocNotaCargo Then
        nSaldoDisp = nSaldoDisp - nMonto
        sMsgOpe = "Ext. Depósito NC"
    End If
Else
    nSaldoDisp = nSaldoDisp - nMonto
    oMov.ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
    sMsgOpe = "Ext. Depósito Efectivo"
End If
oMov.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
'oMov.AgregaMov sMovNro, nOpeExt, sGlosa
oMov.InsertaMov sMovNro, nOpeExt, sGlosa
nMovNro = clsCap.GetnMovNro(sMovNro)
oMov.AgregaMovCap nMovNro, nOpeExt, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
oMov.AgregaMovCapDet nMovNro, nOpeExt, sCuenta, gConcCapital, nMonto
oMov.UltimaActualizacionCuenta sCuenta, sMovNro
If pbMovContable Then
    Dim lnSaldoArendir As Currency
    oMov.ExtornaMovimiento sMovNro, pnMovNroAnt
    lnSaldoArendir = pnSaldoArendir - nMonto
    oMov.ActualizaMovArendir pnMovNroSol, pnTipoArendir, , , , lnSaldoArendir
    ActualizaMovSust pnMovNroSol, pnMovNroAnt, oMov, True, pnTipoArendir
    oMov.ActualizaArendirSust pnMovNroSol
End If
oMov.CommitTrans
bTrans = False
CapExtornoAbonoAhoMov = nSaldoCnt
If pbImprimeBoleta Then
    If sNroDoc = "" Then
        EmiteBoleta "EXTORNO AHORROS", sMsgOpe, sCuenta, nMonto, nOperacion, nSaldoDisp, nSaldoCnt, nIntGanado, nExtracto, , , , , , pdFecsis, psNomAge, psCodUser
    Else
        EmiteBoleta "EXTORNO AHORROS", sMsgOpe, sCuenta, nMonto, nOperacion, nSaldoDisp, nSaldoCnt, nIntGanado, nExtracto, True, nTipoDoc, sNroDoc, , , pdFecsis, psNomAge, psCodUser
    End If
End If
Exit Function
ErrGraba:
    If bTrans Then
        oMov.RollbackTrans
        bTrans = False
    End If
    Set clsCap = Nothing
    Err.Raise Err.Number, "", Err.Description
    CapExtornoAbonoAhoMov = 0
End Function
Public Function CapExtornoCargoAhoMov(ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", _
            Optional pbMovContable As Boolean = False, Optional ByVal pnMovNroAnt As Long, Optional ByVal pnMovNroSol As Long, _
            Optional ByVal pnTipoArendir As ArendirTipo, Optional ByVal pnSaldoArendir As Currency, _
            Optional ByVal pnArendirFase As ARendirFases, Optional pdFecsis As Date, Optional pbImprimeBoleta As Boolean = False, _
            Optional psNomAge As String = "", Optional ByVal psCodUser As String = "") As Double

Dim nOpeExt As CaptacOperacion
Dim nMovNro As Long
Dim clsMant As DCapMantenimiento
Dim clsCap As DCapMovimientos
Dim rsCta As Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, nExtracto As Long
Dim sMsgOpe As String
Dim oMov As DMov
Set oMov = New DMov

nOpeExt = nOperacion
bTrans = False
'Obtiene los datos para el calculo
Set clsMant = New DCapMantenimiento
Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
Set clsMant = Nothing
nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, True, False)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing
Dim oCap As NCapMovimientos
Set oCap = New NCapMovimientos

'Inicia la transaccion
Set clsCap = New DCapMovimientos
On Error GoTo ErrGraba
oMov.BeginTrans
bTrans = True
sGlosa = sGlosa + " Extorno Cargo Cuenta = " & sCuenta
nDiasTranscurridos = DateDiff("d", dUltMov, pdFecsis) - 1
nIntGanado = oCap.GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, pdFecsis)
nSaldoCnt = nSaldoCnt + nMonto
nSaldoDisp = nSaldoDisp + nMonto
oMov.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
oMov.ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
If sNroDoc = "" Then
    sMsgOpe = "Ext. Retiro Efectivo"
Else
    If nTipoDoc = TpoDocOrdenPago Then
        If nOperacion = gAhoRetOPCanje Then
            sMsgOpe = "Ext. Retiro OP Canje"
        Else
            sMsgOpe = "Ext. Retiro OP"
        End If
        oMov.AgregaOrdenPagoEstado sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstExtornada
    ElseIf nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = " Ext. Retiro NC"
    End If
End If
oMov.ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
oMov.InsertaMov sMovNro, nOpeExt, sGlosa, , gMovFlagDeExtorno
nMovNro = oMov.GetnMovNro(sMovNro)
oMov.AgregaMovCap nMovNro, nOpeExt, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
oMov.AgregaMovCapDet nMovNro, nOpeExt, sCuenta, gConcCapital, nMonto
oMov.UltimaActualizacionCuenta sCuenta, sMovNro

    Dim lnSaldoArendir As Currency
oMov.ExtornaMovimiento sMovNro, pnMovNroAnt, , , False
lnSaldoArendir = pnSaldoArendir + nMonto
oMov.ActualizaMovArendir pnMovNroSol, pnTipoArendir, , , , lnSaldoArendir
ActualizaMovSust pnMovNroSol, pnMovNroAnt, oMov, True, pnTipoArendir
oMov.ActualizaArendirSust pnMovNroSol

oMov.CommitTrans
Set oMov = Nothing
CapExtornoCargoAhoMov = nSaldoCnt
If pbImprimeBoleta Then
    If sNroDoc = "" Then
        EmiteBoleta "EXTORNO AHORROS", sMsgOpe, sCuenta, nMonto, nOpeExt, nSaldoDisp, nSaldoCnt, nIntGanado, nExtracto, , , , , , pdFecsis, psNomAge, psCodUser
    Else
        EmiteBoleta "EXTORNO AHORROS", sMsgOpe, sCuenta, nMonto, nOpeExt, nSaldoDisp, nSaldoCnt, nIntGanado, nExtracto, True, nTipoDoc, sNroDoc, , , pdFecsis, psNomAge, psCodUser
    End If
End If
Exit Function
ErrGraba:
    If bTrans Then
        oMov.RollbackTrans
        bTrans = False
    End If
    Set oMov = Nothing
    Err.Raise Err.Number, "CapExtornoCargoAho", Err.Description
    CapExtornoCargoAhoMov = 0
End Function

Public Function GetAtencionPendArendirViaticos(ByVal psUsuario As String, ByVal psAgeCod As String, _
                                       ByVal psAreaCod As String, ByVal psCtaArendir As String, _
                                       ByVal psAreaArendir As String, ByVal psFechaIni As String, _
                                       ByVal psFechaFin As String) As ADODB.Recordset
'GetSustentacionArendirCuenta
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsFiltroAge  As String
Dim lsFiltroArea  As String
Dim lsFiltroUser  As String

Dim gcITF As String
Dim oCon As NConstSistemas
Set oCon = New NConstSistemas

Set oConect = New DConecta
Set rs = New Recordset

psFechaFin = Mid(psFechaFin, 7, 4) + Mid(psFechaFin, 4, 2) + Mid(psFechaFin, 1, 2)
psFechaIni = Mid(psFechaIni, 7, 4) + Mid(psFechaIni, 4, 2) + Mid(psFechaIni, 1, 2)

gcITF = oCon.LeeConstSistema(157)

If oConect.AbreConexion = False Then Exit Function
If psUsuario = "" Then
    lsFiltroUser = ""
Else
    lsFiltroUser = " and RH.cUser In ('" & psUsuario & "') "
End If
If psAgeCod = "" Then
    lsFiltroAge = ""
Else
    lsFiltroAge = " and RH.cAgenciaActual In ('" & psAgeCod & "') "
End If
If psAreaCod = "" Then
    lsFiltroArea = ""
Else
    lsFiltroArea = " and RH.cAreaCodActual In ('" & psAreaCod & "') "
End If


'sql = " Select nMovNro, cPersNombre, dDocFecha, MontoAtendido,cAreaDescripcion, cAgeDescripcion,cDestinoDesc, cAreaCodActual," _
'    & "     cAgeCod, Fecha, dPartida, dLlegada, cMovDesc, CONVERT(CHAR(10),CONVERT(DATETIME,SUBSTRING(FechaRend,1,8)),103) FechaRend" _
'    & " From(SELECT M.NMOVNRO,P.CPERSNOMBRE, CONVERT(CHAR(10),MD.dDocFecha,103) AS dDocFecha , " _
'    & "             ISNULL(ATEND.NATENDIDO,0) AS MONTOATENDIDO, " _
'    & "             ISNULL(A.CAREADESCRIPCION,'') AS CAREADESCRIPCION," _
'    & "             ISNULL(AG.CAGEDESCRIPCION,'') CAGEDESCRIPCION," _
'    & "             MV.cDestinoDesc,RH.cAreaCodActual," _
'    & "             cAgeCod = (SELECT ISNULL(cAgeCod,'') FROM AreaAgencia WHERE cAreaCod = aa.cAreaCod and cAgeCod = ag.cAgeCod)," _
'    & "             CONVERT(CHAR(10),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHA," _
'    & "             MV.dPartida, MV.dPartida + nMovViaticosDias as dLlegada, m.cMovDesc," _
'    & "             (SELECT  M2.cMOVNRO FROM MOV M2 JOIN MOVREF MR ON MR.NMOVNRO = M2.NMOVNRO Where MR.NMOVNROREF = MAR.NMOVNRO AND M2.nMovEstado = 10 AND M2.nMovFlag = 0 AND SUBSTRING(M2.COPECOD,1,5) IN ('" & Mid(gCGArendirViatRendMN, 1, 5) & "','" & Left(gCGArendirViatRendDevMN, 5) & "','" & Left(gCGArendirViatRendDevME, 5) & "','" & Mid(gCGArendirViatRendME, 1, 5) & "')) FechaRend" _
'    & "      FROM MOV M JOIN MOVARENDIR MAR ON M.NMOVNRO = MAR.NMOVNRO " _
'    & "           JOIN PERSONA    P   ON P.CPERSCOD = MAR.CPERSCOD " _
'    & "           LEFT JOIN RRHH  RH  ON RH.CPERSCOD = P.CPERSCOD " _
'    & "           LEFT JOIN AREAS A   ON A.CAREACOD = RH.cAreaCodActual LEFT JOIN (SELECT Distinct cAreaCod FROM AreaAgencia ) AA ON AA.cAreaCod = A.cAreaCod  " _
'    & "           LEFT JOIN AGENCIAS  AG  ON AG.CAGECOD = RH.cAgenciaActual " _
'    & "           JOIN MOVDOC     MD  ON MD.nMOVNRO = MAR.nMOVNRO AND nDocTpo IN (62) " _
'    & "           JOIN DOCUMENTO  D   ON D.nDocTpo = MD.nDocTpo " _
'    & "           JOIN (  SELECT  SUM(ISNULL(ME.nMovMEImporte,MC.NMOVIMPORTE)) AS NATENDIDO, MR.NMOVNROREF " _
'    & "                 FROM    MOVCTA MC JOIN MOV M1 ON M1.NMOVNRO = MC.NMOVNRO LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'    & "                         JOIN MOVREF MR ON MR.NMOVNRO  = M1.NMOVNRO " _
'    & "                 WHERE   MC.NMOVIMPORTE>0 AND M1.NMOVFLAG =" & gMovFlagVigente & "  " _
'    & "                         AND SUBSTRING(M1.COPECOD,1,5) IN ('" & Mid(gCGArendirViatAtencMN, 1, 5) & "','" & Mid(gCGArendirViatAtencME, 1, 5) & "') and MC.cCtacontcod not like ('" & gcITF & "') " _
'    & "                 GROUP  BY MR.NMOVNROREF ) AS ATEND ON ATEND.NMOVNROREF = M.NMOVNRO " _
'    & "           JOIN MOVviaticos MV On M.nMovnro = MV.nViaticoMovNro"
'sql = sql & " WHERE MAR.cTpoArendir ='2' AND SUBSTRING(M.COPECOD,3,1)= '1' " _
'    & "         AND M.nMovFlag =" & gMovFlagVigente & " " & lsFiltroAge & lsFiltroArea & lsFiltroUser & "" _
'    & "         AND EXISTS (SELECT  M2.NMOVNRO " _
'    & "                     FROM MOV M2 " _
'    & "                         JOIN MOVREF MR ON MR.NMOVNRO = M2.NMOVNRO " _
'    & "                     Where MR.NMOVNROREF = MAR.NMOVNRO " _
'    & "                         AND M2.nMovEstado = " & gMovEstContabMovContable & " AND M2.nMovFlag = " & gMovFlagVigente & " " _
'    & "                         AND SUBSTRING(M2.COPECOD,1,5) IN ('" & Mid(gCGArendirViatRendMN, 1, 5) & "','" & Left(gCGArendirViatRendDevMN, 5) & "','" & Left(gCGArendirViatRendDevME, 5) & "','" & Mid(gCGArendirViatRendME, 1, 5) & "'))) As ZZ" _
'    & " WHERE Substring(FechaRend,1,8) Between '" & psFechaIni & "' And '" & psFechaFin & "'" _
'    & " Order By Substring(FechaRend,1,8)"
    
    '***PASI 20140315-********************************************************
    sql = " Select nMovNro, cPersNombre, dDocFecha, MontoAtendido,cAreaDescripcion, cAgeDescripcion,cDestinoDesc, cAreaCodActual," _
    & "     cAgeCod, Fecha, dPartida, dLlegada, cMovDesc, CONVERT(CHAR(10),CONVERT(DATETIME,SUBSTRING(FechaRend,1,8)),103) FechaRend" _
    & " From(SELECT M.NMOVNRO,P.CPERSNOMBRE, CONVERT(CHAR(10),MD.dDocFecha,103) AS dDocFecha , " _
    & "             ISNULL(ATEND.NATENDIDO,0) AS MONTOATENDIDO, " _
    & "             ISNULL(A.CAREADESCRIPCION,'') AS CAREADESCRIPCION," _
    & "             ISNULL(AG.CAGEDESCRIPCION,'') CAGEDESCRIPCION," _
    & "             MV.cDestinoDesc,RH.cAreaCodActual," _
    & "             cAgeCod = (SELECT ISNULL(cAgeCod,'') FROM AreaAgencia WHERE cAreaCod = aa.cAreaCod and cAgeCod = ag.cAgeCod)," _
    & "             CONVERT(CHAR(10),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHA," _
    & "             MV.dPartida, MV.dPartida + nMovViaticosDias as dLlegada, m.cMovDesc," _
    & "             (SELECT  M2.cMOVNRO FROM MOV M2 JOIN MOVREF MR ON MR.NMOVNRO = M2.NMOVNRO Where MR.NMOVNROREF = MAR.NMOVNRO AND M2.nMovEstado = 10 AND M2.nMovFlag = 0 AND SUBSTRING(M2.COPECOD,1,5) IN ('" & Mid(gCGArendirViatRendMN, 1, 5) & "','" & Left(gCGArendirViatRendDevMN, 5) & "','" & Left(gCGArendirViatRendDevME, 5) & "','" & Mid(gCGArendirViatRendME, 1, 5) & "')) FechaRend" _
    & "      FROM MOV M JOIN MOVARENDIR MAR ON M.NMOVNRO = MAR.NMOVNRO " _
    & "           JOIN PERSONA    P   ON P.CPERSCOD = MAR.CPERSCOD " _
    & "           LEFT JOIN RRHH  RH  ON RH.CPERSCOD = P.CPERSCOD " _
    & "           LEFT JOIN AREAS A   ON A.CAREACOD = RH.cAreaCodActual LEFT JOIN (SELECT Distinct cAreaCod FROM AreaAgencia ) AA ON AA.cAreaCod = A.cAreaCod  " _
    & "           LEFT JOIN AGENCIAS  AG  ON AG.CAGECOD = RH.cAgenciaActual " _
    & "           JOIN MOVDOC     MD  ON MD.nMOVNRO = MAR.nMOVNRO AND nDocTpo IN (62) " _
    & "           JOIN DOCUMENTO  D   ON D.nDocTpo = MD.nDocTpo " _
    & "           JOIN (  SELECT  SUM(ISNULL(ME.nMovMEImporte,MC.NMOVIMPORTE)) AS NATENDIDO, MR.NMOVNROREF " _
    & "                 FROM    MOVCTA MC JOIN MOV M1 ON M1.NMOVNRO = MC.NMOVNRO LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
    & "                         JOIN MOVREF MR ON MR.NMOVNRO  = M1.NMOVNRO " _
    & "                 WHERE   MC.NMOVIMPORTE>0 AND M1.NMOVFLAG =" & gMovFlagVigente & "  " _
    & "                         AND SUBSTRING(M1.COPECOD,1,5) IN ('" & Mid(gCGArendirViatAtencMN, 1, 5) & "','" & Mid(gCGArendirViatAtencME, 1, 5) & "','40122','40222') and MC.cCtacontcod not like ('" & gcITF & "') " _
    & "                 GROUP  BY MR.NMOVNROREF ) AS ATEND ON ATEND.NMOVNROREF = M.NMOVNRO " _
    & "           JOIN MOVviaticos MV On M.nMovnro = MV.nViaticoMovNro"
sql = sql & " WHERE MAR.cTpoArendir ='2' AND SUBSTRING(M.COPECOD,3,1)= '1' " _
    & "         AND M.nMovFlag =" & gMovFlagVigente & " " & lsFiltroAge & lsFiltroArea & lsFiltroUser & "" _
    & "         AND EXISTS (SELECT  M2.NMOVNRO " _
    & "                     FROM MOV M2 " _
    & "                         JOIN MOVREF MR ON MR.NMOVNRO = M2.NMOVNRO " _
    & "                     Where MR.NMOVNROREF = MAR.NMOVNRO " _
    & "                         AND M2.nMovEstado = " & gMovEstContabMovContable & " AND M2.nMovFlag = " & gMovFlagVigente & " " _
    & "                         AND SUBSTRING(M2.COPECOD,1,5) IN ('" & Mid(gCGArendirViatRendMN, 1, 5) & "','" & Left(gCGArendirViatRendDevMN, 5) & "','" & Left(gCGArendirViatRendDevME, 5) & "','" & Mid(gCGArendirViatRendME, 1, 5) & "'))) As ZZ" _
    & " WHERE Substring(FechaRend,1,8) Between '" & psFechaIni & "' And '" & psFechaFin & "'" _
    & " Order By Substring(FechaRend,1,8)"
    'PASI FIN*****************************************************************************
Set rs = oConect.CargaRecordSet(sql)
Set GetAtencionPendArendirViaticos = rs
oConect.CierraConexion
Set oConect = Nothing
End Function

Public Function GetDocSustentariosArendirViaticos(ByVal psMovNro As String, ByVal psCtaContArendir As String, ByVal psCtaContPendiente As String) As ADODB.Recordset
'GetDocSustentariosArendirCuenta
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Set oConect = New DConecta
Set rs = New ADODB.Recordset

If oConect.AbreConexion = False Then Exit Function


sql = " SELECT  MD.dDocFecha, D.cDocAbrev, MD.cDocNro, p.cPersNombre," _
    & "         M.cMovDesc, SUM(IsNull(ME.NMOVMEIMPORTE, MC.NMOVIMPORTE) * -1 ) As nDocImporte " _
    & " FROM    MOVREF  MRSUST " _
    & "         JOIN MOV M          ON MRSUST.nMOVNRO = M.nMOVNRO " _
    & "         JOIN MOVCTA MC      ON MC.nMOVNRO = M.nMOVNRO " _
    & "         LEFT JOIN MOVME ME  ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM =MC.nMOVITEM " _
    & "         JOIN MOVDOC MD      ON MD.nMOVNRO = M.nMOVNRO and not md.nDocTpo in (" & TpoDocOrdenPago & "," & TpoDocCheque & "," & TpoDocNotaAbono & "," & TpoDocVoucherEgreso & ", " & TpoDoc.TpoDocCarta & " ) " _
    & "         JOIN " & vsBaseComunes & "DOCUMENTO D    ON D.nDocTpo = MD.nDocTpo " _
    & "         JOIN MovGasto MG  ON MG.nMOVNRO = M.nMOVNRO " _
    & "         JOIN " & vsBasePesonas & "PERSONA P      ON P.CPERSCOD = MG.CPERSCOD " _
    & " WHERE   MRSUST.nMOVNROREF =" & psMovNro & " AND cCtaContCod IN ('" & psCtaContArendir & "','" & psCtaContPendiente & "') " _
    & "         And  M.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "') " _
    & " GROUP BY M.cMovNro, M.cMovDesc, D.cDocAbrev, MD.cDocNro, MD.nDocTpo, p.cPersNombre,MD.dDocFecha"

Set rs = oConect.CargaRecordSet(sql)
Set GetDocSustentariosArendirViaticos = rs
oConect.CierraConexion: Set oConect = Nothing
End Function

'*** PEAC 20101019
Public Sub Distribuye(ByRef pbDocDolares As Boolean, ByVal psOpeCod As String, _
    ByVal psCodArea As String, ByVal psSubCta As String, ByVal pnMovNro As Long, ByVal pnItem As Integer, ByVal pnDocTpo As TpoDoc, ByVal pnImporte As Currency, _
    ByVal prsCtasCont As ADODB.Recordset, ByVal prsObjetos As ADODB.Recordset, Optional ByRef pnItemDistri As Integer = 0, Optional ByRef psCtasInexis As String = "", _
    Optional ByVal psMatrizViaticos As Variant = "", Optional ByVal pnNumAgesViaticos As Integer = 0)

    Dim oMovi As DMov
    Set oMovi = New DMov
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim sql As String
    Dim rsAgesDistrib As ADODB.Recordset
    Dim lnMonto As Currency, lnRegImporte As Currency, lcCtaContCod As String
    Dim lcCodObj As String
    Dim lnItemDistri As Integer
    Dim lnItemPrin As Integer, lnImpoPrin As Currency
    Dim lsCtasInexis As String
    Dim lsSql As String
    Dim lcCtaCont As String
    
    Dim i As Integer, j As Integer
    
    Dim rsBuscaCuenta As ADODB.Recordset
    Set rsBuscaCuenta = New ADODB.Recordset
        
    lnItemDistri = 0
    lnMonto = 0
    lnRegImporte = 0
        
    If pbDocDolares Then
        lcCtaContCod = Left(prsCtasCont!Código, 2) & "2" & Trim(Mid(prsCtasCont!Código, 4, 50)) + psSubCta
    Else
        lcCtaContCod = prsCtasCont!Código + psSubCta
    End If
    
    
    '' 401140 a rendir cta - 401240 a rendir viaticos
    'If (psOpeCod = "401140" Or psOpeCod = "401240") And Right(psSubCta, 2) = "01" And psCodArea <> "026" Then
    '***Modificado por ELRO el 20120921, según OYP-RFC111-2012
    'If (psOpeCod = "401140" Or psOpeCod = "402140" Or psOpeCod = "401240" Or psOpeCod = "402240") And Right(lcCtaContCod, 2) = "01" Then
    If (psOpeCod = "401140" Or psOpeCod = "402140" Or psOpeCod = "401141" Or psOpeCod = "402141" Or psOpeCod = "401240" Or psOpeCod = "402240" Or psOpeCod = "401241" Or psOpeCod = "402241" Or psOpeCod = "401344") And Right(lcCtaContCod, 2) = "01" Then
    '***Fin Modificado por ELRO el 20120921*******************
        
        If psCodArea = "026" Or psCodArea = "041" Or psCodArea = "067" Then '*** PEAC 20110304
            pnItem = IIf(pnItem = 0, pnItem + 1, pnItem)
            oMovi.InsertaMovCta pnMovNro, pnItem, lcCtaContCod, IIf(prsCtasCont!DH = "D", pnImporte, pnImporte * -1) * IIf(pnDocTpo = 7, -1, 1)
        Else
            If Not prsObjetos Is Nothing Then
                prsObjetos.MoveFirst
                Do While Not prsObjetos.EOF
                    If prsObjetos!CtaCont = prsCtasCont!Código And prsObjetos!ItemCtaCont = prsCtasCont!ItemCtaCont Then
                        lcCodObj = prsObjetos!Código
                    End If
                    prsObjetos.MoveNext
                Loop
            End If
            
            '*** PEAC 20110219 - SOLO DEBE DISTRIBUIR LA "45"
            If lcCodObj <> "02601" And Left(lcCtaContCod, 2) = "45" Then
                
                If pnNumAgesViaticos > 0 Then
                
                    lnMonto = 0
                    lnRegImporte = 0
                    lnItemDistri = pnItem
                
                    For i = 1 To pnNumAgesViaticos
                                
                        lcCtaCont = Left(lcCtaContCod, Len(lcCtaContCod) - 2) + psMatrizViaticos(1, i)
                        
                        Set rsBuscaCuenta = oMovi.CargaRecordSet("Select top 1 isnull(max(cCtaContCod),'X') cCtaContCod From CtaCont where cCtaContCod Like '" & lcCtaCont & "%' And cCtaContCod <> '" & lcCtaCont & "' And nCtaEstado = 1")
                        If rsBuscaCuenta!cCtaContCod <> "X" Then
                            lcCtaContCod = Mid(Trim(rsBuscaCuenta!cCtaContCod), 1, Len(Trim(rsBuscaCuenta!cCtaContCod)) - 2)
                            lcCtaCont = lcCtaContCod + psMatrizViaticos(1, i)
                        End If
                        
                        Set rsBuscaCuenta = oMovi.CargaRecordSet("select * From CtaCont where cCtaContCod like '" & lcCtaCont & "' ")
                        If Not (rsBuscaCuenta.EOF And rsBuscaCuenta.BOF) Then
                            'lnItemDistri = IIf(lnItemDistri = 1, lnItemDistri, lnItemDistri + 1)
                            
                            '***Modificado por ELRO el 20120921, según OYP-RFC111-2012
                            'If psOpeCod = "401140" Or psOpeCod = "402140" Then
                            If psOpeCod = "401140" Or psOpeCod = "402140" Or psOpeCod = "401141" Or psOpeCod = "402141" Or psOpeCod = "401344" Then
                            '***Fin Modificado por ELRO el 20120921*******************
                                lnRegImporte = Round((IIf(prsCtasCont!DH = "D", psMatrizViaticos(3, i), psMatrizViaticos(3, i) * -1) * IIf(pnDocTpo = 7, -1, 1)), 2)
                            Else
                                lnRegImporte = Round((IIf(prsCtasCont!DH = "D", pnImporte, pnImporte * -1) * IIf(pnDocTpo = 7, -1, 1)) / pnNumAgesViaticos, 2)
                            End If
                            
                            If i = 1 Then
                                lnItemPrin = lnItemDistri
                                lnImpoPrin = lnRegImporte
                            End If
                            oMovi.InsertaMovCta pnMovNro, lnItemDistri, lcCtaCont, lnRegImporte
                            lnMonto = lnMonto + lnRegImporte
                            lnItemDistri = lnItemDistri + 1
                        Else
                            lsCtasInexis = lsCtasInexis + lcCtaCont + "-"
                        End If
                    Next i
                
                Else
                
                    Set rsAgesDistrib = New ADODB.Recordset
                    '**************************************************
                    'Modificado por PASI20131111 segun TI-ERS107-2013
                    'lssql = " exec stp_sel_AgenciaPorcentajeGastos "
                      lsSql = "EXEC stp_sel_RecuperaAgencias '01'"
                    '**************************************************
                    Set rsAgesDistrib = oMovi.CargaRecordSet(lsSql)
                    lnMonto = 0
                    lnRegImporte = 0
                    'lnItemDistri = pnItemDistri
                    lnItemDistri = IIf(pnItem = 0, pnItem + 1, pnItem)
                    
                    Do While Not rsAgesDistrib.EOF
                        'lcCtaCont = lcCtaContCod + rsAgesDistrib!cAgeCod
                        lcCtaCont = Left(lcCtaContCod, Len(lcCtaContCod) - 2) + rsAgesDistrib!cAgeCod
                                        
                        Set rsBuscaCuenta = oMovi.CargaRecordSet("Select top 1 isnull(max(cCtaContCod),'X') cCtaContCod From CtaCont where cCtaContCod Like '" & lcCtaCont & "%' And cCtaContCod <> '" & lcCtaCont & "' And nCtaEstado = 1")
                        If rsBuscaCuenta!cCtaContCod <> "X" Then
                            lcCtaContCod = Mid(Trim(rsBuscaCuenta!cCtaContCod), 1, Len(Trim(rsBuscaCuenta!cCtaContCod)) - 2)
                            lcCtaCont = lcCtaContCod + rsAgesDistrib!cAgeCod
                        End If
                        
                        Set rsBuscaCuenta = oMovi.CargaRecordSet("select * From CtaCont where cCtaContCod like '" & lcCtaCont & "' ")
                        If Not (rsBuscaCuenta.EOF And rsBuscaCuenta.BOF) Then
                            'lnItemDistri = lnItemDistri + 1
                            'lnRegImporte = Round((IIf(prsCtasCont!DH = "D", pnImporte, pnImporte * -1) * IIf(pnDocTpo = 7, -1, 1)) * rsAgesDistrib!nAgePorcentaje / 100, 2) '***Comentado por PASI20131111 segun TI-ERS107-2013
                            lnRegImporte = pnImporte 'PASI20131111 segun TI-ERS107-2013
                            
                            If rsAgesDistrib!cAgeCod = "01" Then
                                lnItemPrin = lnItemDistri
                                lnImpoPrin = lnRegImporte
                            End If
                            
                            oMovi.InsertaMovCta pnMovNro, lnItemDistri, lcCtaCont, lnRegImporte
                            lnMonto = lnMonto + lnRegImporte
                            lnItemDistri = lnItemDistri + 1
                        Else
                            lsCtasInexis = lsCtasInexis + lcCtaCont + "-"
                        End If
                        
                        rsAgesDistrib.MoveNext
                    Loop
                End If
                
                If lnMonto > pnImporte Then
                    If pnNumAgesViaticos > 0 Then '*** PEAC 20110303
                        oMovi.ActualizaMovCta pnMovNro, lnItemPrin, , lnImpoPrin - (lnMonto - pnImporte) '' la dif manda al primer registro
                    Else
                        'lcCtaCont = lcCtaContCod + "01"
                        lcCtaCont = Left(lcCtaContCod, Len(lcCtaContCod) - 2) + "01"
                        Set rsBuscaCuenta = oMovi.CargaRecordSet("select * From CtaCont where cCtaContCod like '" & lcCtaCont & "' ")
                        If Not (rsBuscaCuenta.EOF And rsBuscaCuenta.BOF) Then
                            oMovi.ActualizaMovCta pnMovNro, lnItemPrin, , lnImpoPrin - (lnMonto - pnImporte) '' la dif manda a la ag princ
                        Else
                            oMovi.ActualizaMovCta pnMovNro, lnItemDistri, , lnRegImporte - (lnMonto - pnImporte) '' si noexiste la age prin manda al ultimo age registrado
                        End If
                    End If
                ElseIf pnImporte > lnMonto Then
                    If pnNumAgesViaticos > 0 Then '*** PEAC 20110303
                        oMovi.ActualizaMovCta pnMovNro, lnItemPrin, , lnImpoPrin - (lnMonto - pnImporte) '' la dif manda al primer registro
                    Else
                        'lcCtaCont = lcCtaContCod + "01"
                        lcCtaCont = Left(lcCtaContCod, Len(lcCtaContCod) - 2) + "01"
                        Set rsBuscaCuenta = oMovi.CargaRecordSet("select * From CtaCont where cCtaContCod like '" & lcCtaCont & "' ")
                        If Not (rsBuscaCuenta.EOF And rsBuscaCuenta.BOF) Then
                            oMovi.ActualizaMovCta pnMovNro, lnItemPrin, , lnImpoPrin + (pnImporte - lnMonto) '' la dif manda a la ag princ
                        Else
                            oMovi.ActualizaMovCta pnMovNro, lnItemDistri, , lnRegImporte + (pnImporte - lnMonto)
                        End If
                    End If
                End If
                
                RSClose rsAgesDistrib
    
                Set rsAgesDistrib = Nothing
                
                pnItemDistri = lnItemDistri
                psCtasInexis = lsCtasInexis
                
            Else
                'lnItemDistri = IIf(lnItemDistri = 0, pnItem + 1, lnItemDistri)
                'pnItem = pnItem + 1
                oMovi.InsertaMovCta pnMovNro, pnItem, prsCtasCont!Código + psSubCta, IIf(prsCtasCont!DH = "D", pnImporte, pnImporte * -1) * IIf(pnDocTpo = 7, -1, 1)
            End If
        End If
    Else
         'oMovi.InsertaMovCta pnMovNro, pnItem, lcCtaContCod + psSubCta, IIf(prsCtasCont!DH = "D", pnImporte, pnImporte * -1) * IIf(pnDocTpo = 7, -1, 1)
         oMovi.InsertaMovCta pnMovNro, pnItem, lcCtaContCod, IIf(prsCtasCont!DH = "D", pnImporte, pnImporte * -1) * IIf(pnDocTpo = 7, -1, 1)
    End If
    
    '**** PEAC 20110128
    If lnItemDistri = 0 Then
        pnItemDistri = pnItem
    End If
    '*** FIN PEAC
    
End Sub

'*** PEAC 20101108
Public Function GetSustentacionArendirCuenta(ByVal psFechaIni As String, _
                                       ByVal psFechaFin As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsFiltroAge  As String
Dim lsFiltroArea  As String
Dim lsFiltroUser  As String

Dim gcITF As String
Dim oCon As NConstSistemas
Set oCon = New NConstSistemas

Set oConect = New DConecta
Set rs = New Recordset

psFechaFin = Mid(psFechaFin, 7, 4) + Mid(psFechaFin, 4, 2) + Mid(psFechaFin, 1, 2)
psFechaIni = Mid(psFechaIni, 7, 4) + Mid(psFechaIni, 4, 2) + Mid(psFechaIni, 1, 2)

gcITF = oCon.LeeConstSistema(157)

If oConect.AbreConexion = False Then Exit Function

sql = " exec stp_sel_GetSustentacionArendirCuenta '" & psFechaIni & "','" & psFechaFin & "'"
    
Set rs = oConect.CargaRecordSet(sql)
Set GetSustentacionArendirCuenta = rs
oConect.CierraConexion
Set oConect = Nothing
End Function

'*** PEAC 20101108
Public Function GetDocSustentariosArendirCuenta(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Set oConect = New DConecta
Set rs = New ADODB.Recordset

If oConect.AbreConexion = False Then Exit Function

sql = " exec stp_sel_GetDocSustentariosArendirCuenta " & pnMovNro

Set rs = oConect.CargaRecordSet(sql)
Set GetDocSustentariosArendirCuenta = rs
oConect.CierraConexion: Set oConect = Nothing
End Function

'*** PEAC 20110106
Public Function BuscaPendienteRendicion(ByVal psPersCod As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta

Set rs = New ADODB.Recordset
Set oConect = New DConecta

If oConect.AbreConexion = False Then Exit Function

sql = " exec stp_sel_BuscaRendicionPendiente '" & psPersCod & "'"

Set rs = oConect.CargaRecordSet(sql)
Set BuscaPendienteRendicion = rs
oConect.CierraConexion
Set oConect = Nothing
End Function

'***Agregado por ELRO el 20120321, según OYP-RFC005-2012 y OYP-RFC016-2012

Public Function actualizarSolicitudArendir(ByVal pnMovNro As Long, ByVal psMovDesc As String, ByVal pnMovMonto As Currency) As Boolean
Dim oDMov As DMov
Dim lbTrans As Boolean
Dim sMsgErr As String

On Error GoTo ErractualizarSolicitudArendir

Set oDMov = New DMov

oDMov.BeginTrans
lbTrans = True
Call oDMov.actualizarConceptoSolicitudARendirCta(pnMovNro, psMovDesc)
Call oDMov.actualizarMovCont(pnMovNro, pnMovMonto)
oDMov.CommitTrans
lbTrans = False
actualizarSolicitudArendir = True
Set oDMov = Nothing
Exit Function
ErractualizarSolicitudArendir:
    sMsgErr = Err.Description
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    Err.Raise vbObjectError + 100, "actualizarSolicitudArendir", sMsgErr
End Function

Public Function obtenerSolicitudARendirCta(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsSolicitud As ADODB.Recordset
Set rsSolicitud = New ADODB.Recordset

On Error GoTo ErrobtenerSolicitudARendirCta

    If oDConecta.AbreConexion = False Then Exit Function
    sql = " exec stp_sel_ObtenerSolicitudARendirCta " & pnMovNro & ""
    oDConecta.Ejecutar sql
    Set rsSolicitud = oDConecta.CargaRecordSet(sql)
    Set obtenerSolicitudARendirCta = rsSolicitud
    oDConecta.CierraConexion
    sql = ""
    Set oDConecta = Nothing
    Set rsSolicitud = Nothing
    Exit Function

ErrobtenerSolicitudARendirCta:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:obtenerSolicitudARendirCta", Err.Description
End Function

Public Function devolverARendirViaticosPendientes(ByVal pbTodo As Integer, ByVal psMoneda As String, ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal pnTpoArendir As Integer) As ADODB.Recordset
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsPendientes As ADODB.Recordset
Set rsPendientes = New ADODB.Recordset

On Error GoTo ErrdevolverARendirViaticosPendientes

    If oDConecta.AbreConexion = False Then Exit Function
    
    sql = " exec stp_sel_DevolverARendirViaticosPendientes " & pbTodo & ", '" & psMoneda & "', '" & psAreaCod & "', '" & psAgeCod & "', " & pnTpoArendir & ""
    
    Set rsPendientes = oDConecta.CargaRecordSet(sql)
    Set devolverARendirViaticosPendientes = rsPendientes
    
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsPendientes = Nothing
    sql = ""
    Exit Function
    
ErrdevolverARendirViaticosPendientes:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:devolverARendirViaticosPendientes", Err.Description
End Function
Public Function devolverARendirCtaPendientes(ByVal pbTodo As Integer, ByVal psMoneda As String, ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal pnTpoArendir As Integer) As ADODB.Recordset
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsPendientes As ADODB.Recordset
Set rsPendientes = New ADODB.Recordset

On Error GoTo ErrdevolverARendirCtaPendientes

    If oDConecta.AbreConexion = False Then Exit Function
    
    sql = " exec stp_sel_DevolverARendirCuentaPendientes " & pbTodo & ", '" & psMoneda & "', '" & psAreaCod & "', '" & psAgeCod & "', " & pnTpoArendir & ""
    
    Set rsPendientes = oDConecta.CargaRecordSet(sql)
    Set devolverARendirCtaPendientes = rsPendientes
    
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsPendientes = Nothing
    sql = ""
    Exit Function
    
ErrdevolverARendirCtaPendientes:
    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:devolverARendirCtaPendientes", Err.Description
End Function

Public Sub eliminarSolicitudARendirViatico(ByVal pnViaticoMovNro As Long)

Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta

On Error GoTo ErreliminarSolicitudARendirViatico

    If oDConecta.AbreConexion = False Then Exit Sub
    sql = " exec stp_sel_EliminarSolicitudARendirViaticos " & pnViaticoMovNro
    oDConecta.Ejecutar sql
    oDConecta.CierraConexion
    sql = ""
    Set oDConecta = Nothing
    Exit Sub

ErreliminarSolicitudARendirViatico:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:eliminarSolicitudARendirViatico", Err.Description

End Sub

Public Sub eliminarSolicitudARendirCuenta(ByVal pnCuentaMovNro As Long)

Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta

On Error GoTo ErreliminarSolicitudARendirCuenta

    If oDConecta.AbreConexion = False Then Exit Sub
    sql = " exec stp_sel_EliminarSolicitudARendirCuenta " & pnCuentaMovNro
    oDConecta.Ejecutar sql
    oDConecta.CierraConexion
    sql = ""
    Set oDConecta = Nothing
    Exit Sub

ErreliminarSolicitudARendirCuenta:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:eliminarSolicitudARendirCuenta", Err.Description

End Sub

Public Sub obtenerSaldoARendirViaticos(ByVal psPersCod As String, _
                                       ByRef nSaldoPendienteMN As Currency, _
                                       ByRef nSaldoPendienteME As Currency)
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsSaldo As ADODB.Recordset
Set rsSaldo = New ADODB.Recordset

On Error GoTo ErrobtenerSaldoARendirViaticos

    If oDConecta.AbreConexion = False Then Exit Sub
    
    sql = " exec stp_sel_ObtenerSaldoARendirViaticos  '" & psPersCod & "'"
    
    Set rsSaldo = oDConecta.CargaRecordSet(sql)
    
    If Not rsSaldo.BOF And Not rsSaldo.EOF Then
        Do While Not rsSaldo.EOF
            If rsSaldo!cMoneda = "1" Then
               nSaldoPendienteMN = rsSaldo!nSaldo
            Else
               nSaldoPendienteME = rsSaldo!nSaldo
            End If
            rsSaldo.MoveNext
        Loop
    Else
        nSaldoPendienteMN = 0#
        nSaldoPendienteME = 0#
    End If
    
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsSaldo = Nothing
    sql = ""
    Exit Sub
ErrobtenerSaldoARendirViaticos:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:obtenerSaldoARendirViaticos", Err.Description
End Sub

Public Sub obtenerSaldoARendirCuentas(ByVal psPersCod As String, _
                                      ByRef nSaldoPendienteMN As Currency, _
                                      ByRef nSaldoPendienteME As Currency)
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsSaldo As ADODB.Recordset
Set rsSaldo = New ADODB.Recordset

On Error GoTo ErrobtenerSaldoARendirCuentas

    If oDConecta.AbreConexion = False Then Exit Sub
    
    sql = " exec stp_sel_ObtenerSaldoARendirCuentas  '" & psPersCod & "'"
    
    Set rsSaldo = oDConecta.CargaRecordSet(sql)
    
    If Not rsSaldo.BOF And Not rsSaldo.EOF Then
        Do While Not rsSaldo.EOF
            If rsSaldo!cMoneda = "1" Then
               nSaldoPendienteMN = rsSaldo!nSaldo
            Else
               nSaldoPendienteME = rsSaldo!nSaldo
            End If
            rsSaldo.MoveNext
        Loop
    Else
        nSaldoPendienteMN = 0#
        nSaldoPendienteME = 0#
    End If
    
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsSaldo = Nothing
    sql = ""
    Exit Sub
ErrobtenerSaldoARendirCuentas:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:obtenerSaldoARendirCuentas", Err.Description
End Sub

Public Function obtenerARendirViaticosParaSustentar(ByVal psMoneda As String, ByVal psPersCod As String) As ADODB.Recordset
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsSustentar As ADODB.Recordset
Set rsSustentar = New ADODB.Recordset

On Error GoTo ErrobtenerARendirViaticosParaSustentar

    If oDConecta.AbreConexion = False Then Exit Function
    
    sql = " exec stp_sel_ObtenerARendirViaticosParaSustentar '" & psMoneda & "', '" & psPersCod & "'"
    
    Set rsSustentar = oDConecta.CargaRecordSet(sql)
    Set obtenerARendirViaticosParaSustentar = rsSustentar
    
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsSustentar = Nothing
    sql = ""
    Exit Function
ErrobtenerARendirViaticosParaSustentar:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:obtenerARendirViaticosParaSustentar", Err.Description
End Function

Public Function obtenerARendirCuentasParaSustentar(ByVal psMoneda As String, ByVal psPersCod As String) As ADODB.Recordset
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsSustentar As ADODB.Recordset
Set rsSustentar = New ADODB.Recordset

On Error GoTo ErrobtenerARendirCuentasParaSustentar

    If oDConecta.AbreConexion = False Then Exit Function
    
    sql = " exec stp_sel_ObtenerARendirCuentasParaSustentar '" & psMoneda & "', '" & psPersCod & "'"
    
    Set rsSustentar = oDConecta.CargaRecordSet(sql)
    Set obtenerARendirCuentasParaSustentar = rsSustentar
    
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsSustentar = Nothing
    sql = ""
    Exit Function
ErrobtenerARendirCuentasParaSustentar:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:obtenerARendirCuentasParaSustentar", Err.Description
End Function


Public Function obtenerFechaLlegadaColaborador(ByVal pnViaticoMovNro As Long, ByVal pnTpoArendir As Integer) As Date
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsFecha As ADODB.Recordset
Set rsFecha = New ADODB.Recordset

On Error GoTo ErrobtenerFechaLlegadaColaborador

    If oDConecta.AbreConexion = False Then Exit Function
    
    sql = "exec stp_sel_ObtenerFechaLlegadaColaborador " & pnViaticoMovNro & ", " & pnTpoArendir & ""
    
    Set rsFecha = oDConecta.CargaRecordSet(sql)
    
    If Not rsFecha.BOF And Not rsFecha.EOF Then
      obtenerFechaLlegadaColaborador = rsFecha!dFechaLlegada
    Else
      obtenerFechaLlegadaColaborador = "01/01/1900"
    End If
        
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsFecha = Nothing
    sql = ""
    Exit Function
ErrobtenerFechaLlegadaColaborador:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:obtenerFechaLlegadaColaborador", Err.Description
End Function
'TORE - Automatización de Prorroga
Public Function obtenerDiasTrancurridos(ByVal pdFechaLLegada As String, ByVal psAgeCod As String, ByVal pnMovNro As Long, ByVal pdFechaSis As String, ByVal psUsuario As String, ByVal pnTipoArendir As Integer) As ADODB.Recordset
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsDias As ADODB.Recordset
Set rsDias = New ADODB.Recordset

On Error GoTo ErrorobtenerDiasTrancurridos

    If oDConecta.AbreConexion = False Then Exit Function
    
    sql = "exec stp_sel_DiasProrroga '" & pdFechaLLegada & "', '" & psAgeCod & "'," & pnMovNro & ",'" & pdFechaSis & "','" & psUsuario & "', " & pnTipoArendir & ""
    
    Set rsDias = oDConecta.CargaRecordSet(sql)
    Set obtenerDiasTrancurridos = rsDias

    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsDias = Nothing
    sql = ""
    Exit Function
ErrorobtenerDiasTrancurridos:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:ErrorobtenerDiasTrancurridos", Err.Description
End Function
'END
Public Function obtenerLimiteSustentarCategoria(ByVal psRHCargoCategoria As String, ByVal pnTipoArendir As Integer) As Integer
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsLimite As ADODB.Recordset
Set rsLimite = New ADODB.Recordset

On Error GoTo ErrobtenerLimiteSustentarCategoria

    If oDConecta.AbreConexion = False Then Exit Function
    
    sql = " exec stp_sel_ObtenerLimiteSustentarCategoria '" & psRHCargoCategoria & "', " & pnTipoArendir & ""
    
    Set rsLimite = oDConecta.CargaRecordSet(sql)
    
    If Not rsLimite.BOF And Not rsLimite.EOF Then
      obtenerLimiteSustentarCategoria = rsLimite!nNroDia
    Else
      obtenerLimiteSustentarCategoria = 0
    End If
        
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsLimite = Nothing
    sql = ""
    Exit Function
    
ErrobtenerLimiteSustentarCategoria:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:obtenerLimiteSustentarCategoria", Err.Description
End Function

Public Function obtenerPlanillaRendicionSustentar(ByVal pnMovNro As Long, _
                                                  ByVal psCtaContArendir As String, _
                                                  ByVal psMoneda As String) As ADODB.Recordset
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsSustentar As ADODB.Recordset
Set rsSustentar = New ADODB.Recordset

On Error GoTo ErrobtenerPlanillaRendicionSustentar

    If oDConecta.AbreConexion = False Then Exit Function
    
    sql = " exec stp_sel_ObtenerPlanillaRendicionSustentar " & pnMovNro & ", '" & psCtaContArendir & "', '" & psMoneda & "'"
    
    Set rsSustentar = oDConecta.CargaRecordSet(sql)
    Set obtenerPlanillaRendicionSustentar = rsSustentar
    
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsSustentar = Nothing
    sql = ""
    Exit Function
ErrobtenerPlanillaRendicionSustentar:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:obtenerPlanillaRendicionSustentar", Err.Description
End Function


Public Function obtenerPlanillaRendicionSustentar_2(ByVal pnMovNro As Long, _
                                                    ByVal psCtaContArendir As String, _
                                                    ByVal psMoneda As String) As ADODB.Recordset
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsSustentar As ADODB.Recordset
Set rsSustentar = New ADODB.Recordset

On Error GoTo ErrobtenerPlanillaRendicionSustentar_2

    If oDConecta.AbreConexion = False Then Exit Function
    
    sql = " exec stp_sel_ObtenerPlanillaRendicionSustentar_2 " & pnMovNro & ", '" & psCtaContArendir & "', '" & psMoneda & "'"
    
    Set rsSustentar = oDConecta.CargaRecordSet(sql)
    Set obtenerPlanillaRendicionSustentar_2 = rsSustentar
    
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsSustentar = Nothing
    sql = ""
    Exit Function
ErrobtenerPlanillaRendicionSustentar_2:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:obtenerPlanillaRendicionSustentar_2", Err.Description
End Function

Public Function obtenerPlanillaAsientoContablesSustentar(ByVal pnMovNro As Long, _
                                                         ByVal psCtaContArendir As String, _
                                                         ByVal psMoneda As String) As ADODB.Recordset
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsSustentar As ADODB.Recordset
Set rsSustentar = New ADODB.Recordset

On Error GoTo ErrobtenerPlanillaAsientoContablesSustentar

    If oDConecta.AbreConexion = False Then Exit Function
    
    sql = " exec stp_sel_ObtenerPlanillaAsientoContablesSustentar " & pnMovNro & ", '" & psCtaContArendir & "', '" & psMoneda & "'"
    
    Set rsSustentar = oDConecta.CargaRecordSet(sql)
    Set obtenerPlanillaAsientoContablesSustentar = rsSustentar
    
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsSustentar = Nothing
    sql = ""
    Exit Function
    
ErrobtenerPlanillaAsientoContablesSustentar:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:obtenerPlanillaAsientoContablesSustentar", Err.Description
End Function

Public Function obtenerPlanillaAsientoContablesSustentar_2(ByVal pnMovNro As Long, _
                                                           ByVal psCtaContArendir As String, _
                                                           ByVal psMoneda As String) As ADODB.Recordset
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsSustentar As ADODB.Recordset
Set rsSustentar = New ADODB.Recordset

On Error GoTo ErrobtenerPlanillaAsientoContablesSustentar_2

    If oDConecta.AbreConexion = False Then Exit Function
    
    sql = " exec stp_sel_ObtenerPlanillaAsientoContablesSustentar_2 " & pnMovNro & ", '" & psCtaContArendir & "', '" & psMoneda & "'"
    
    Set rsSustentar = oDConecta.CargaRecordSet(sql)
    Set obtenerPlanillaAsientoContablesSustentar_2 = rsSustentar
    
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsSustentar = Nothing
    sql = ""
    Exit Function
    
ErrobtenerPlanillaAsientoContablesSustentar_2:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:obtenerPlanillaAsientoContablesSustentar_2", Err.Description
End Function

Public Function obtenerARendirViaticosParaRendir(ByVal psMoneda As String, ByVal psAreaCod As String, ByVal pbTodo As Boolean) As ADODB.Recordset
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsSustentar As ADODB.Recordset
Set rsSustentar = New ADODB.Recordset

On Error GoTo ErrobtenerARendirViaticosParaRendir

    If oDConecta.AbreConexion = False Then Exit Function
    
    sql = " exec stp_sel_ObtenerARendirViaticosParaRendir '" & psMoneda & "', '" & psAreaCod & "', " & IIf(pbTodo, 1, 0) & ""
    
    Set rsSustentar = oDConecta.CargaRecordSet(sql)
    Set obtenerARendirViaticosParaRendir = rsSustentar
    
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsSustentar = Nothing
    sql = ""
    Exit Function
ErrobtenerARendirViaticosParaRendir:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:obtenerARendirViaticosParaRendir", Err.Description
End Function

Public Function obtenerARendirCuentasParaRendir(ByVal psMoneda As String, ByVal psAreaCod As String, ByVal pbTodo As Boolean) As ADODB.Recordset
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsSustentar As ADODB.Recordset
Set rsSustentar = New ADODB.Recordset

On Error GoTo ErrobtenerARendirCuentasParaRendir

    If oDConecta.AbreConexion = False Then Exit Function
    
    sql = " exec stp_sel_ObtenerARendirCuentasParaRendir '" & psMoneda & "', '" & psAreaCod & "', " & IIf(pbTodo, 1, 0) & ""
    
    Set rsSustentar = oDConecta.CargaRecordSet(sql)
    Set obtenerARendirCuentasParaRendir = rsSustentar
    
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsSustentar = Nothing
    sql = ""
    Exit Function
ErrobtenerARendirCuentasParaRendir:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:obtenerARendirCuentasParaRendir", Err.Description
End Function


Public Sub registrarProrrogaSustentar(ByVal pnDias As Integer, ByVal pnMovNro As Long, ByVal psMovNro As String, ByVal pnNroProrroga As Integer) '****pnNroProrroga Agregado por PASI20140102 TI-ERS107-2013
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta

On Error GoTo ErrregistrarProrrogaSustentar

    If oDConecta.AbreConexion = False Then Exit Sub
    
    '****Modificado por PASI20140102 TI-ERS107-2013
    'sql = " exec stp_ins_RegistrarProrrogaSustentar " & pnDias & ", " & pnMovNro & ", '" & psMovNro & "'"
        sql = " exec stp_ins_RegistrarProrrogaSustentar " & pnDias & ", " & pnMovNro & ", '" & psMovNro & "'," & pnNroProrroga & ""
    '***FIN PASI
    oDConecta.Ejecutar (sql)
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    sql = ""
    Exit Sub
ErrregistrarProrrogaSustentar:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:registrarProrrogaSustentar", Err.Description
End Sub

'TORE - Automatizacion de Prorroga
Public Sub RegProrrogaSustentarAut(ByVal pnMovNro As Long, ByVal psMovNro As String, ByVal psGlosa As String)
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta

On Error GoTo ErrregistrarProrrogaSustentar

    If oDConecta.AbreConexion = False Then Exit Sub
    sql = " exec stp_ins_ProrrogaSustentarAut " & pnMovNro & ", '" & psMovNro & "','" & psGlosa & "'"
    oDConecta.Ejecutar (sql)
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    sql = ""
    Exit Sub
ErrregistrarProrrogaSustentar:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:registrarProrrogaSustentar", Err.Description
End Sub
'END TORE


Public Function obtenerExistenciaProrrogaSustentar(ByVal pnMovNro As Long) As Integer
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsSustentar As ADODB.Recordset
Set rsSustentar = New ADODB.Recordset

On Error GoTo ErrobtenerExistenciaProrrogaSustentar

    If oDConecta.AbreConexion = False Then Exit Function
    
    sql = " exec stp_sel_ObtenerExistenciaProrrogaSustentar " & pnMovNro & ""
    
    Set rsSustentar = oDConecta.CargaRecordSet(sql)
    If Not rsSustentar.BOF And Not rsSustentar.EOF Then
         '***Modificado por PASI20140102 TI-ERS107-2013
        'obtenerExistenciaProrrogaSustentar = rsSustentar.Fields(0)
        obtenerExistenciaProrrogaSustentar = rsSustentar!nNroProrroga
        'Fin PASI
    Else
        obtenerExistenciaProrrogaSustentar = 0
    End If
    
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsSustentar = Nothing
    sql = ""
    Exit Function
ErrobtenerExistenciaProrrogaSustentar:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:obtenerExistenciaProrrogaSustentar", Err.Description
End Function

Public Function devolverProrrogaSustentar(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsSustentar As ADODB.Recordset
Set rsSustentar = New ADODB.Recordset

On Error GoTo ErrdevolverProrrogaSustentar

    If oDConecta.AbreConexion = False Then Exit Function
    
    sql = " exec stp_sel_ObtenerExistenciaProrrogaSustentar_2 " & pnMovNro & ""
    
    Set rsSustentar = oDConecta.CargaRecordSet(sql)
    Set devolverProrrogaSustentar = rsSustentar
    
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsSustentar = Nothing
    sql = ""
    Exit Function
ErrdevolverProrrogaSustentar:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:devolverProrrogaSustentar", Err.Description
End Function

Public Sub registrarRendicionDsctoPlanillaViaticos(ByVal psMovNro As String, _
                                                        ByVal psOpeCod As String, _
                                                        ByVal psMovDesc As String, _
                                                        ByVal pnMovNroSolicitud As Long, _
                                                        ByVal pnMontoDescontar As Currency, _
                                                        ByVal psCtaContArendir As String, _
                                                        ByVal psCtaContPendiente As String, _
                                                        ByVal pnArendirTipo As ArendirTipo, _
                                                        ByRef pnMovNroCon As Long)
Dim oDMov As DMov
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim lsMsgErr As String
Set oDMov = New DMov

On Error GoTo ErrregistrarRendicionPlanilla
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnMontoDescontar, 0, "0"
oDMov.InsertaMovRef lnMovNro, pnMovNroSolicitud
oDMov.InsertaMovCta lnMovNro, 1, psCtaContPendiente, pnMontoDescontar
oDMov.InsertaMovCta lnMovNro, 2, psCtaContArendir, pnMontoDescontar * -1
oDMov.ActualizaMovArendir pnMovNroSolicitud, pnArendirTipo, , , , 0
oDMov.ActualizaSaldoMovimiento psMovNro, "+"
oDMov.CommitTrans
lbTrans = False
pnMovNroCon = lnMovNro
Set oDMov = Nothing
Exit Sub
ErrregistrarRendicionPlanilla:
   pnMovNroCon = 0
   lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "NARendir:registrarRendicionPlanillaViaticos", lsMsgErr
End Sub

Public Sub registrarRendicionDsctoPlanillaARendirCuenta(ByVal psMovNro As String, _
                                                        ByVal psOpeCod As String, _
                                                        ByVal psMovDesc As String, _
                                                        ByVal pnMovNroSolicitud As Long, _
                                                        ByVal pnMovNroAtencion As Long, _
                                                        ByVal pnMontoDescontar As Currency, _
                                                        ByVal psCtaContArendir As String, _
                                                        ByVal psCtaContPendiente As String, _
                                                        ByVal pnArendirTipo As ArendirTipo, _
                                                        ByRef pnMovNroCon As Long)
Dim oDMov As DMov
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim lsMsgErr As String
Set oDMov = New DMov

On Error GoTo ErrregistrarRendicionDsctoPlanillaARendirCuenta
oDMov.BeginTrans
lbTrans = True
oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)
oDMov.InsertaMovCont lnMovNro, pnMontoDescontar, 0, "0"
oDMov.InsertaMovRef lnMovNro, pnMovNroAtencion
oDMov.InsertaMovCta lnMovNro, 1, psCtaContPendiente, pnMontoDescontar
oDMov.InsertaMovCta lnMovNro, 2, psCtaContArendir, pnMontoDescontar * -1
oDMov.ActualizaMovArendir pnMovNroSolicitud, pnArendirTipo, , , , 0
oDMov.ActualizaSaldoMovimiento psMovNro, "+"
oDMov.CommitTrans
lbTrans = False
pnMovNroCon = lnMovNro
Set oDMov = Nothing
Exit Sub
ErrregistrarRendicionDsctoPlanillaARendirCuenta:
   pnMovNroCon = 0
   lsMsgErr = Err.Description
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "NARendir:registrarRendicionDsctoPlanillaARendirCuenta", lsMsgErr
End Sub

'***Fin Agregado por ELRO**************************************************
'***Agregado por ELRO el 20120529, según OYP-RFC047-2012
Public Function devolverCodigoJefeContabilidad() As String
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsCodigo As ADODB.Recordset
Set rsCodigo = New ADODB.Recordset

On Error GoTo ErrdevolverCodigoJefeContabilidad

    If oDConecta.AbreConexion = False Then Exit Function

    sql = " exec stp_sel_ObtenerCodigoJefeContabilidad"

    Set rsCodigo = oDConecta.CargaRecordSet(sql)

    If Not rsCodigo.BOF And Not rsCodigo.EOF Then
        devolverCodigoJefeContabilidad = rsCodigo.Fields(0)
    Else
        devolverCodigoJefeContabilidad = ""
    End If


    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Set rsCodigo = Nothing
    sql = ""
    Exit Function

ErrdevolverCodigoJefeContabilidad:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:devolverCodigoJefeContabilidad", Err.Description
End Function

Public Function devolverDetalleDocSustentariosArendir(ByVal psMovNro As String, ByVal psCtaContArendir As String) As ADODB.Recordset
Dim sql As String
Dim rsDocSus As ADODB.Recordset
Set rsDocSus = New ADODB.Recordset
Dim oDConecta As DConecta
Set oDConecta = New DConecta

On Error GoTo devolverDetalleDocSustentariosArendirErr

If oDConecta.AbreConexion = False Then Exit Function
sql = "exec stp_sel_obtenerDetalleDocSustentariosArendir " & CLng(psMovNro) & ", '" & psCtaContArendir & "'"
Set rsDocSus = oDConecta.CargaRecordSet(sql)
Set devolverDetalleDocSustentariosArendir = rsDocSus
oDConecta.CierraConexion: Set oDConecta = Nothing
Exit Function

devolverDetalleDocSustentariosArendirErr:
Set oDConecta = Nothing
Err.Raise Err.Number, "NARendir:devolverDetalleDocSustentariosArendirErr", Err.Description
End Function
'***Fin Agregado por ELRO*******************************

'***Agregado por ELRO el 20130115, según OYP-RF126-2012
Public Function devolverAgencias() As ADODB.Recordset
On Error GoTo devolverAgenciasErr
    Dim sSQL As String
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta

    sSQL = "exec stp_sel_RFC1262012_DevolverAgencias"

    If oDConecta.AbreConexion() Then
        Set devolverAgencias = oDConecta.CargaRecordSet(sSQL)
        oDConecta.CierraConexion
    End If
    Set oDConecta = Nothing
    Exit Function
devolverAgenciasErr:
Call RaiseError(MyUnhandledError, "NARendir:devolverAgencias Method")
End Function

Public Function devolverUsuarios() As ADODB.Recordset
On Error GoTo devolverUsuariosErr
    Dim sSQL As String
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta

    sSQL = "exec stp_sel_RFC1262012_DevolverUsuarios"

    If oDConecta.AbreConexion() Then
        Set devolverUsuarios = oDConecta.CargaRecordSet(sSQL)
        oDConecta.CierraConexion
    End If
    Set oDConecta = Nothing
    Exit Function
devolverUsuariosErr:
Call RaiseError(MyUnhandledError, "NARendir:devolverUsuarios Method")
End Function

Public Function devolverGastoMovilidadResumen(ByVal psFechaIni As String, _
                                              ByVal psFechaFin As String, _
                                              ByVal psUser As String, _
                                              ByVal psAgencia As String) As ADODB.Recordset
On Error GoTo devolverGastoMovilidadResumenErr
    Dim sSQL As String
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta

    sSQL = "exec stp_sel_RFC1262012_DevolverGastosMovilidadResumen '" & psFechaIni & "', '" & psFechaFin & "', '" & psUser & "', '" & psAgencia & "'"

    If oDConecta.AbreConexion() Then
        Set devolverGastoMovilidadResumen = oDConecta.CargaRecordSet(sSQL)
        oDConecta.CierraConexion
    End If
    Set oDConecta = Nothing
    Exit Function
devolverGastoMovilidadResumenErr:
Call RaiseError(MyUnhandledError, "NARendir:devolverGastoMovilidadResumen Method")
End Function

Public Function devolverGastoMovilidadDetalle(ByVal psFechaIni As String, _
                                              ByVal psFechaFin As String, _
                                              ByVal psUser As String, _
                                              ByVal psAgencia As String) As ADODB.Recordset
On Error GoTo devolverGastoMovilidadDetalleErr
    Dim sSQL As String
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta

    sSQL = "exec stp_sel_RFC1262012_DevolverGastosMovilidadDetalle '" & psFechaIni & "', '" & psFechaFin & "', '" & psUser & "', '" & psAgencia & "'"

    If oDConecta.AbreConexion() Then
        Set devolverGastoMovilidadDetalle = oDConecta.CargaRecordSet(sSQL)
        oDConecta.CierraConexion
    End If
    Set oDConecta = Nothing
    Exit Function
devolverGastoMovilidadDetalleErr:
Call RaiseError(MyUnhandledError, "NARendir:devolverGastoMovilidadDetalle Method")
End Function

Public Function devolverNombreAgenciasLineal(ByVal psAgencia As String) As String
On Error GoTo devolverNombreAgenciasLinealErr
    Dim sSQL As String
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta
    Dim rsAgencias As ADODB.Recordset
    Set rsAgencias = New ADODB.Recordset

    sSQL = "exec stp_sel_RS240009_DevolverAgencias '" & psAgencia & "'"

    If oDConecta.AbreConexion() Then
        Set rsAgencias = oDConecta.CargaRecordSet(sSQL)
        devolverNombreAgenciasLineal = rsAgencias!cAgencias
        oDConecta.CierraConexion
    End If
    Set oDConecta = Nothing
    Exit Function
devolverNombreAgenciasLinealErr:
Call RaiseError(MyUnhandledError, "NARendir:devolverNombreAgenciasLineal Method")
End Function

'***Fin Agregado por ELRO el 20130115******************

'*********Agregado por PASI20131118 segun TI-ERS107-2013
Public Function ObtenerARendirCuentasParaRendirxPersona(ByVal psPersCod As String, ByVal psMoneda As String) As ADODB.Recordset
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsRendicion As ADODB.Recordset
Set rsRendicion = New ADODB.Recordset
 
On Error GoTo ErroObtenerARendirCuentasParaRendirxPersona
    If oDConecta.AbreConexion = False Then Exit Function
    sql = "exec stp_sel_ObtenerARendirCuentasParaRendirxPersona '" & psMoneda & "','','true','" & psPersCod & "'"
    Set rsRendicion = oDConecta.CargaRecordSet(sql)
    Set ObtenerARendirCuentasParaRendirxPersona = rsRendicion
    Exit Function
ErroObtenerARendirCuentasParaRendirxPersona:
Set oDConecta = Nothing
Err.Raise Err.Number, "NARendir:ObtenerARendirCuentasParaRendirxPersona", Err.Description
End Function
Public Function ObtenerARendirViaticosParaRendirxPersona(ByVal psPersCod As String, ByVal psMoneda As String) As ADODB.Recordset
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta
Dim rsRendicion As ADODB.Recordset
Set rsRendicion = New ADODB.Recordset

On Error GoTo ErroObtenerARendirViaticosParaRendirxPersona
    If oDConecta.AbreConexion = False Then Exit Function
    sql = "exec stp_sel_ObtenerARendirViaticosParaRendirxPersona '" & psMoneda & "','','true','" & psPersCod & "'"
    Set rsRendicion = oDConecta.CargaRecordSet(sql)
    Set ObtenerARendirViaticosParaRendirxPersona = rsRendicion
    Exit Function
ErroObtenerARendirViaticosParaRendirxPersona:
Set oDConecta = Nothing
Err.Raise Err.Number, "NARendir:ObtenerARendirViaticosParaRendirxPersona", Err.Description
End Function
'*********Fin Agregado por PASI
'******************Agregado por PASI20140102 TI-ERS107-2013
Public Sub ActualizarProrrogaSustentar(ByVal pnDias As Integer, ByVal pnMovNro As Long, ByVal psMovNro As String, ByVal pnNroProrroga As Integer) '****pnNroProrroga Agregado por PASI20140102 TI-ERS107-2013
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta

On Error GoTo ErrregistrarProrrogaSustentar

    If oDConecta.AbreConexion = False Then Exit Sub
        sql = "exec stp_upd_ActualizarExistenciaProrrogaSustentar " & pnDias & ", " & pnMovNro & ", '" & psMovNro & "'," & pnNroProrroga & ""
    oDConecta.Ejecutar (sql)
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    sql = ""
    Exit Sub
ErrregistrarProrrogaSustentar:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:registrarProrrogaSustentar", Err.Description
End Sub
'******************Fin PASI

'TORE
Public Sub ActProrrogaSustentarAut(ByVal pnDias As Integer, ByVal pnMovNro As Long, ByVal psMovNro As String, ByVal pnNroProrroga As Integer, ByVal psGlosa As String)
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta

On Error GoTo ErrregistrarProrrogaSustentar

    If oDConecta.AbreConexion = False Then Exit Sub
        sql = "exec stp_upd_ActualizarExistenciaProrrogaSustentarAut " & pnDias & ", " & pnMovNro & ", '" & psMovNro & "'," & pnNroProrroga & ",'" & psGlosa & "'"
    oDConecta.Ejecutar (sql)
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    sql = ""
    Exit Sub
ErrregistrarProrrogaSustentar:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:registrarProrrogaSustentar", Err.Description
End Sub

Public Sub RegLogProrrogaAut(ByVal pnDias As Integer, ByVal pnMovNro As Long, ByVal psMovNro As String, ByVal pnNroProrroga As Integer, ByVal psGlosa As String)
Dim sql As String
Dim oDConecta As DConecta
Set oDConecta = New DConecta

On Error GoTo ErrregistrarProrrogaSustentar

    If oDConecta.AbreConexion = False Then Exit Sub
        sql = "exec stp_upd_ActualizarExistenciaProrrogaSustentarAut " & pnDias & ", " & pnMovNro & ", '" & psMovNro & "'," & pnNroProrroga & ",'" & psGlosa & "'"
    oDConecta.Ejecutar (sql)
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    sql = ""
    Exit Sub
ErrregistrarProrrogaSustentar:

    Set oDConecta = Nothing
    Err.Raise Err.Number, "NARendir:registrarProrrogaSustentar", Err.Description
End Sub

'END TORE


'VAPI SEGUN ERS 179-2014
Public Function ObtenerSolicitudViaticosXPersona(ByVal psPersCod As String) As ADODB.Recordset
    Dim sql As String
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta
    Dim rsSolicitud As ADODB.Recordset
    Set rsSolicitud = New ADODB.Recordset
     
    On Error GoTo ErroObtenerSolicitudViaticosXPersona
    
        If oDConecta.AbreConexion = False Then Exit Function
        sql = "exec stp_ERS1792014_sel_SolicitudViaticosXPersona '" & psPersCod & "'"
        Set rsSolicitud = oDConecta.CargaRecordSet(sql)
        Set ObtenerSolicitudViaticosXPersona = rsSolicitud
        
        oDConecta.CierraConexion
        Set oDConecta = Nothing
        
        Exit Function
        
ErroObtenerSolicitudViaticosXPersona:
        Set oDConecta = Nothing
        Err.Raise Err.Number, "NARendir:ErroObtenerSolicitudViaticosXPersona", Err.Description
End Function
'***FIN VAPI********************************************************

Public Function ExistProcesoPost(ByVal pnMovNro As Long) As Boolean
Dim sql As String
Dim oConec As New DConecta
Dim rs As New ADODB.Recordset
oConec.AbreConexion
sql = "Exec stp_sel_ExistMov_ProcesoPost " & pnMovNro & ""
Set rs = oConec.CargaRecordSet(sql)
If Not rs.BOF And Not rs.EOF Then
   ExistProcesoPost = True
Else
   ExistProcesoPost = False
End If
oConec.CierraConexion
Exit Function
End Function 'NAGL ERS 075-2017 20171116

Public Function ObtenerDatosGeneralComisionado(ByVal psPersCod As String) As ADODB.Recordset
Dim sql As String
Dim oDConecta As New DConecta
Dim RsDatos As New ADODB.Recordset
On Error GoTo ErroObtenerDatosGeneralComisionado
    oDConecta.AbreConexion
    sql = "Exec stp_sel_DatosComisionadoGeneral '" & psPersCod & "'"
    Set RsDatos = oDConecta.CargaRecordSet(sql)
    Set ObtenerDatosGeneralComisionado = RsDatos
    oDConecta.CierraConexion
    Exit Function
ErroObtenerDatosGeneralComisionado:
        Set oDConecta = Nothing
        Err.Raise Err.Number, "NARendir:ErroObtenerDatosGeneralComisionado", Err.Description
End Function 'NAGL ERS 075-2017 20171116

Public Function ObtenerCtaAhorroComis(ByVal psPersCod As String, ByVal psMoneda As String) As ADODB.Recordset
Dim sql As String
Dim oDConecta As New DConecta
Dim RsCtaAh As New ADODB.Recordset
On Error GoTo ErroObtenerCtaAhorroComis
    oDConecta.AbreConexion
    sql = "Exec stp_sel_ObtieneCtaAhorrosComisionado '" & psPersCod & "','" & psMoneda & "'"
    Set RsCtaAh = oDConecta.CargaRecordSet(sql)
    Set ObtenerCtaAhorroComis = RsCtaAh
    oDConecta.CierraConexion
    Exit Function
ErroObtenerCtaAhorroComis:
        Set oDConecta = Nothing
        Err.Raise Err.Number, "NARendir:ErroObtenerCtasAhorroComi", Err.Description
End Function 'NAGL ERS 075-2017 20171116

Public Function ObtieneMovGeneraAsiento(ByVal pnMovNro As Long, ByVal psCtaCod As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
Dim oConecta As New DConecta
Dim RsDatos  As New ADODB.Recordset
On Error GoTo ErrObtieneMovGeneraAsiento
oConecta.AbreConexion
sql = "Exec ObtieneMovCaptacGenerarAsiento " & pnMovNro & ",'" & psCtaCod & "'"
Set RsDatos = oConecta.CargaRecordSet(sql)
Set ObtieneMovGeneraAsiento = RsDatos
oConecta.CierraConexion
Exit Function
ErrObtieneMovGeneraAsiento:
        Set oConecta = Nothing
        Err.Raise Err.Number, "NARendir:ErroObtenerMovparaGenerarAsiento", Err.Description
End Function 'NAGL ERS075-2017 20171127
Public Function GetCuentaContableReal(ByVal psCtaCod As String, ByVal pnMovNro As Long, ByVal psOpeCod As String, ByVal pnConcepto As Integer, ByVal pnMonto As Double, ByVal pdFecha As Date) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
Dim oConecta As New DConecta
Dim rsCtaCont  As New ADODB.Recordset
On Error GoTo ErrGetCuentaContableReal
oConecta.AbreConexion
sql = "Exec stp_ERS0752017_sel_CargaOpeCuenta '" & psCtaCod & "', " & pnMovNro & ", '" & psOpeCod & "', " & pnConcepto & ", " & pnMonto & ",'" & Format(pdFecha, "YYYYMMDD") & "' "
Set rsCtaCont = oConecta.CargaRecordSet(sql)
Set GetCuentaContableReal = rsCtaCont
oConecta.CierraConexion
Exit Function
ErrGetCuentaContableReal:
        Set oConecta = Nothing
        Err.Raise Err.Number, "NARendir:ErrGetCuentaContableReal", Err.Description
End Function 'NAGL ERS075-2017 20171127

Public Function ExistMovCta(ByVal pnMovNro As Long, ByVal psCtaCod As String, ByVal psTipo As String) As ADODB.Recordset
    Dim oConec As New DConecta
    Dim psSql As String
    On Error GoTo ErroExistMovCta
    psSql = "Exec stp_sel_ExistMovCtaMovME " & pnMovNro & ", '" & psCtaCod & "', '" & psTipo & "'"
    oConec.AbreConexion
    Set ExistMovCta = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    Exit Function
ErroExistMovCta:
    Call RaiseError(MyUnhandledError, "Carga MovCta - MovMe")
End Function 'NAGL ERS075-2017 20171127

