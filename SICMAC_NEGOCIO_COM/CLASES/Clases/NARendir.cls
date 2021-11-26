VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NARendir"
Attribute VB_GlobalNameSpace = True
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

sSQL = " SELECT  CH.CAREACOD + CH.CAGECOD AS Codigo, ISNULL(AG.CAGEDESCRIPCION ,A.CAREADESCRIPCION ) as Descripcion  , 1 AS NIVEL " _
    & "  FROM    CAJACHICA CH JOIN " & vsBaseComunes & "AREAAGENCIA AA ON AA.CAREACOD = CH.CAREACOD " _
    & "          JOIN " & vsBaseComunes & "AREAS A ON A.CAREACOD = CH.CAREACOD " _
    & "          LEFT JOIN " & vsBaseComunes & "AGENCIAS AG ON AG.CAGECOD = CH.CAGECOD " _
    & "  GROUP BY CH.CAREACOD, CH.CAGECOD , A.CAREADESCRIPCION, AG.CAGEDESCRIPCION "


Set rs = oConect.CargaRecordSet(sSQL)
Set EmiteCajasChicas = rs

oConect.CierraConexion
Set oConect = Nothing
Exit Function
ErrorEmiteCajasChicas:
    Err.Raise vbObjectError + 100, "EmiteCajasChicas", "Error emitir Cajas Chicas"
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
    Case gArendirTipoCajaGeneral
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
            & "         LEFT JOIN ( SELECT  SUM(MC.NMOVIMPORTE) AS NATENDIDO, MR.NMOVNROREF " _
            & "                     FROM    MOVCTA MC JOIN MOV M1 ON M1.NMOVNRO = MC.NMOVNRO " _
            & "                             JOIN MOVREF MR ON MR.NMOVNRO  = M1.NMOVNRO " _
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
    Case gArendirTipoCajaGeneral
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
            & "                      SUM(ABS(IsNull(ME.NMOVMEIMPORTE, Mc.NMOVIMPORTE))) As nDocImporte , MRSUST.nMOVNROREF  AS nMovAtenc " _
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
            & "         A.CAREACOD, A.CAREADESCRIPCION, ISNULL(AG.CAGECOD,'') AS CAGECOD , ISNULL(AG.CAGEDESCRIPCION,'') CAGEDESCRIPCION, " _
            & "         MAR.CAREACOD as cAreaCodArendir, A1.CAREADESCRIPCION as cAreaARendir ,  ISNULL(MAR.CAGECOD,'') AS cAgeCodArendir , ISNULL(AG1.CAGEDESCRIPCION,'') cAgeArendir ,  " _
            & "         REND.ImporteArendir, MAR.NMOVSALDO ,  MD.nDocTpo , D.CDOCABREV , (D.CDOCABREV + '-'  + MD.CDOCNRO) as CDOCNRO , MD.dDocFecha as dDocFechaSol ,  " _
            & "         Rend.nDocTpo AS nDocTpoATENC,  Rend.CDOCABREV AS CDOCABREVATENC, Rend.CDOCNRO AS CDOCNROATENC , " _
            & "         Rend.CDOCDesc AS cDocDescAtenc, Rend.dDocFecha as dFechaDocAtenc, " _
            & "         SUST.nMovSust , SUST.cMovDesc  AS cDescSust, SUST.cDocAbrev as cDocNroAbrevSust, " _
            & "         SUST.cDocNro AS CDOCNROSUST , SUST.nDocTpo AS nDocTpoSUST, SUST.dDocFecha AS dDocFechaSust,  " _
            & "         SUST.cPersNombre AS cPersCodSust, SUST.nDocImporte " _
            & " FROM    MOV M " _
            & "         JOIN MOVARENDIR MAR                         ON MAR.nMOVNRO = M.nMOVNRO " _
            & "         JOIN " & vsBasePesonas & "PERSONA P         ON P.CPERSCOD = MAR.CPERSCOD " _
            & "         JOIN RRHH   RH                              ON RH.CPERSCOD = P.CPERSCOD " _
            & "         JOIN " & vsBaseComunes & "AREAS A           ON A.CAREACOD = RH.cAreaCodActual   " _
            & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG  ON AG.CAGECOD = RH.cAgenciaActual " _
            & "         JOIN " & vsBaseComunes & "AREAS A1          ON A1.CAREACOD = MAR.CAREACOD " _
            & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS     AG1 ON AG1.CAGECOD = MAR.CAGECOD " _
            & "         JOIN MOVDOC MD                              ON MD.nMOVNRO = M.nMOVNRO AND MD.nDocTpo IN ('" & TpoDocRecViaticosARendirCuenta & "') " _
            & "         JOIN DOCUMENTO D                            ON D.nDocTpo = MD.nDocTpo " _
            & "         Join (  SELECT  MC.cCtaContCod, MIN(md.cDocNro) cDocNro, MIN(md.nDocTpo) nDocTpo, MIN(d.cDocAbrev) cDocAbrev, MIN(d.cDocDesc) cDocDesc, MIN(md.dDocFecha) dDocFecha, " _
            & "                         SUM (IsNull(MC.NMOVIMPORTE,ME.NMOVMEIMPORTE)) as ImporteArendir   , " _
            & "                         MRSUST.nMOVNROREF AS nMovSol "
sql = sql + "                   FROM    MOVREF  MRSUST " _
            & "                         JOIN MOV M          ON MRSUST.nMOVNRO = M.nMOVNRO " _
            & "                         JOIN MOVCTA MC      ON MC.nMOVNRO = M.nMOVNRO " _
            & "                         LEFT JOIN MOVME ME  ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM =MC.nMOVITEM " _
            & "                         JOIN MovDoc md      ON md.nMovNro = m.nMovNro and md.nDocTpo = " & TpoDocOrdenPago & " JOIN Documento d ON d.nDocTpo = md.nDocTpo " _
            & "                 WHERE   M.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagModificado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "')  AND SUBSTRING(M.COPECOD,1,5) IN ('" & Mid(gCGArendirViatAtencMN, 1, 5) & "','" & Mid(gCGArendirViatAtencME, 1, 5) & "') " _
            & "                         AND MRSUST.nMOVNROREF ='" & lsMovNroArendirAtencion & "' AND MC.cCtaContCod  IN ('" & psCtaContArendir & "','" & psCtaContPendiente & "') " _
            & "                         " _
            & "                 GROUP BY MC.cCtacontCod, MRSUST.nMOVNROREF ) as REND " _
            & "                ON M.nMovNro = REND.nMovSol "
sql = sql + "         Join   (SELECT    M.nMovNro AS nMovSust, M.cMovDesc, D.cDocAbrev, MD.cDocNro, " _
            & "                         MD.nDocTpo, p.cPersNombre,MD.dDocFecha, " _
            & "                         SUM(ABS(IsNull(MC.NMOVIMPORTE, Me.NMOVMEIMPORTE))) As nDocImporte , " _
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
            ByVal pnImporte As Currency, ByVal psTpoIf As String, ByVal psPersCodIf As String, ByVal psCtaIf As String, _
            ByVal rsBilletaje As ADODB.Recordset, ByVal psDocTpo As String, ByVal psDocNro As String, ByVal psFechaDoc As String, _
            ByVal psDocTpoVoucher As String, ByVal psDocVoucher As String, _
            ByVal pnMovRef As Long) As Integer
            
Dim lnMovItem As Long
Dim lnMovOrden As Integer
Dim lnMovNro As Long
Dim oDMov As DMov
Dim lbTrans As Boolean
Dim lsMovNro As String
Dim lnSaldoArendir As Currency

Set oDMov = New DMov
On Error GoTo ErrorGrabaAtencionArendir
oDMov.BeginTrans
lbTrans = True

oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
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
        oDMov.InsertaMovObjIF lnMovNro, lnMovItem, lnMovOrden, psPersCodIf, psTpoIf, psCtaIf
        
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

lnSaldoArendir = GetSaldoARendir(pnMovRef)
lnSaldoArendir = lnSaldoArendir + pnImporte
oDMov.ActualizaMovArendir pnMovRef, pnTipoArendir, , , , lnSaldoArendir
oDMov.InsertaMovRef lnMovNro, pnMovRef

If Mid(psOpeCod, 3, 1) = "2" Then
    oDMov.GeneraMovME lnMovNro, psMovNro
End If

oDMov.ActualizaSaldoMovimiento psMovNro, "+"

oDMov.CommitTrans
lbTrans = False
GrabaAtencionArendir = 0
Set oDMov = Nothing
Exit Function
ErrorGrabaAtencionArendir:
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
    Err.Raise vbObjectError + 100, "GrabaAtencionArendir", "Error Graba Atencion Arendir "
End Function
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
                                        ByVal psDocNro As String, ByVal pdDocFecha As Date) As Integer
                                        
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
oDMov.InsertaMovCta lnMovNro, lnItem, psCtaContHaber + lsSubCta, pnImporte * -1
oDMov.InsertaMovObj lnMovNro, lnItem, 1, ObjCMACAgenciaArea
oDMov.InsertaMovObjAgenciaArea lnMovNro, lnItem, 1, psAgeCh, psAreaCh

oDMov.InsertaMovDoc lnMovNro, TpoDocRecArendirCuenta, psDocNro, Format(pdDocFecha, gsFormatoFecha)

oDMov.InsertaMovRef lnMovNro, pnMovNroArendirSol
oDMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, pnNroProcCH, gCHTipoProcArendir
oDMov.ActualizaMovArendir pnMovNroArendirSol, gArendirTipoCajaChica, , , , pnImporte

lnSaldo = oCh.GetDatosCajaChica(psAreaCh, psAgeCh, SaldoActual)
lnSaldoNuevo = lnSaldo - pnImporte
oDMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProcCH, , , , , lnSaldoNuevo
    
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
                                        ByVal psAgeCajaCh As String, ByVal pnNroPro As Integer) As Integer
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
Public Function GetAtencionPendArendir(ByVal pbTodos As Boolean, ByVal psAgeCod As String, _
                                       ByVal psAreaCod As String, ByVal pnTipoArendir As ArendirTipo, _
                                       ByVal psCtaArendir As String, ByVal pnMoneda As Moneda, _
                                       ByVal psAreaArendir As String, ByVal psAgeArendir As String, Optional pbRendidos As Boolean = False) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As DConecta
Dim lsFiltroAge  As String
Dim lsFiltroArea  As String
Set oConect = New DConecta
Set rs = New Recordset

If oConect.AbreConexion = False Then Exit Function
If psAgeCod <> "" Then
    lsFiltroAge = " and RH.cAgenciaActual='" & psAgeCod & "' "
End If
If psAreaCod <> "" Then
    lsFiltroArea = " and RH.cAreaCodActual='" & psAreaCod & "' "
End If

Select Case pnTipoArendir
    Case gArendirTipoCajaGeneral
        sql = "SELECT   ISNULL(D1.cDocAbrev,'EFEC') AS cAbreDocAtenc, ISNULL(MD.cDocNro,'Efectivo') as cNroDocAtenc, ISNULL(CONVERT(CHAR(10),MD.dDocFecha,103),SUBSTRING(M.CMOVNRO,7,2) + '/' + SUBSTRING(M.CMOVNRO,5,2)+ '/' + SUBSTRING(M.CMOVNRO,1,4)) AS dDocAtencion , " _
            & "         MDAR.cDocNro, CONVERT(CHAR(10),MDAR.dDocFecha,103) AS dDocSolicitud,CONVERT(CHAR(50),P.CPERSNOMBRE) AS cPersNombre, ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS NIMPORTE,  " _
            & "         M.cMovDesc , P.cPersCod, M.NMOVNRO as NMovNroAtenc, ISNULL(MD.nDocTpo,'') as cTpoDocAtenc, MAR.nMovSaldo ,  " _
            & "         ISNULL(A.cAreaDescripcion,'') cAreaDescripcion,ISNULL(A.cAreaCod,'') cAreaCod, ISNULL(D1.cDocDesc,'Atención Efectivo') as cDocDesc, MAR.NMOVNRO as NMovNroSol,   " _
            & "         cAgeCod = (SELECT ISNULL(cAgeCod,'') FROM AreaAgencia WHERE cAreaCod = aa.cAreaCod and cAgeCod = ag.cAgeCod), D2.cDocAbrev, MDAR.nDocTpo " _
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
            & "                         WHERE   MR1.NMOVNROREF=M.NMOVNRO AND SUBSTRING(M1.COPECOD,1,5) IN ('" & Mid(gCGArendirCtaRendMN, 1, 5) & "','" & Mid(gCGArendirCtaRendME, 1, 5) & "','" & Mid(gCGArendirCtaRendDevMN, 1, 5) & "','" & Mid(gCGArendirCtaRendDevME, 1, 5) & "') " _
            & "                                 AND M1.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "') ) ") _
            & "         AND MAR.cAreacod ='" & psAreaArendir & "' and MAR.cAgeCod='" & psAgeArendir & "' " _
            & " ORDER BY MDAR.cDocNro "

    Case gArendirTipoViaticos
        
        sql = " SELECT  D.CDOCABREV, MD.CDOCNRO AS DocNro, CONVERT(CHAR(12),MD.dDocFecha,103) AS dDocFecha , MD.CDOCNRO AS DocNro, CONVERT(CHAR(12),MD.dDocFecha,103) AS dDocFecha , " _
            & "         P.CPERSNOMBRE, ISNULL(ATEND.NATENDIDO,0) AS MONTOATENDIDO , " _
            & "         M.cMovDesc,  P.CPERSCOD, M.NMOVNRO, MD.nDocTpo, MAR.NMOVSALDO , " _
            & "         ISNULL(A.CAREADESCRIPCION,'') AS CAREADESCRIPCION, A.CAREACOD,  D.cDocDesc, " _
            & "         M.NMOVNRO, ISNULL(AG.CAGEDESCRIPCION,'') CAGEDESCRIPCION, cAgeCod = (SELECT ISNULL(cAgeCod,'') FROM AreaAgencia WHERE cAreaCod = aa.cAreaCod and cAgeCod = ag.cAgeCod)  , " _
            & "         P.CPERSNOMBRE, CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHA, D.CDOCABREV , MD.nDocTpo " _
            & " FROM    MOV M JOIN MOVARENDIR MAR ON M.NMOVNRO = MAR.NMOVNRO " _
            & "         JOIN PERSONA            P       ON P.CPERSCOD = MAR.CPERSCOD " _
            & "         JOIN RRHH               RH      ON RH.CPERSCOD = P.CPERSCOD " _
            & "         LEFT JOIN AREAS         A       ON A.CAREACOD = RH.cAreaCodActual JOIN (SELECT Distinct cAreaCod FROM AreaAgencia ) AA ON AA.cAreaCod = A.cAreaCod  " _
            & "         LEFT JOIN AGENCIAS  AG  ON AG.CAGECOD = RH.cAgenciaActual " _
            & "         JOIN MOVDOC         MD  ON MD.nMOVNRO = MAR.nMOVNRO AND nDocTpo IN (" & TpoDocRecViaticosARendirCuenta & ") " _
            & "         JOIN DOCUMENTO      D   ON D.nDocTpo = MD.nDocTpo " _
            & "         JOIN (  SELECT  SUM(MC.NMOVIMPORTE) AS NATENDIDO, MR.NMOVNROREF " _
            & "                 FROM    MOVCTA MC JOIN MOV M1 ON M1.NMOVNRO = MC.NMOVNRO " _
            & "                         JOIN MOVREF MR ON MR.NMOVNRO  = M1.NMOVNRO " _
            & "                 WHERE   MC.NMOVIMPORTE>0 AND M1.NMOVFLAG =" & gMovFlagVigente & "  " _
            & "                         AND SUBSTRING(M1.COPECOD,1,5) IN ('" & Mid(gCGArendirViatAtencMN, 1, 5) & "','" & Mid(gCGArendirViatAtencME, 1, 5) & "') " _
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
                                            ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal pnMoneda As Moneda) As ADODB.Recordset
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
    Case gArendirTipoCajaGeneral
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
            
    Case gArendirTipoViaticos
        
        sql = " SELECT  ISNULL(D1.cDocAbrev,'EFEC') AS cAbreDocAtenc, " _
            & "         ISNULL(MD.cDocNro,'Efectivo') as cNroDocAtenc, " _
            & "         ISNULL(CONVERT(CHAR(10),MD.dDocFecha,103), " _
            & "         SUBSTRING(M.CMOVNRO,7,2) + '/' + SUBSTRING(M.CMOVNRO,5,2)+ '/' + SUBSTRING(M.CMOVNRO,1,4)) AS dDocAtencion , " _
            & "         D2.cDocAbrev + '-' + MDAR.cDocNro AS cDocNro, CONVERT(CHAR(10),MDAR.dDocFecha,103) AS dDocSolicitud,CONVERT(CHAR(50),P.CPERSNOMBRE) AS cPersNombre, " _
            & "         ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS NIMPORTE, " _
            & "         M.cMovDesc , P.cPersCod, M.NMOVNRO as cMovNroAtenc, ISNULL(MD.nDocTpo,'') as cTpoDocAtenc, " _
            & "         MAR.nMovSaldo ,           ISNULL(A.cAreaDescripcion,'') cAreaDescripcion,ISNULL(A.cAreaCod,'') cAreaCod, " _
            & "         ISNULL(D1.cDocDesc,'Atención Efectivo') as cDocDesc, MAR.NMOVNRO as NMovNroSol, " _
            & "         ISNULL(AG.cAgeDescripcion,'') cAgeDescripcion , ISNULL(AG.cAgeCod,'') cAgeCod, D2.cDocAbrev, MDAR.nDocTpo " _
            & " FROM    MOV M " _
            & "         JOIN MOVCTA MC                              ON MC.NMOVNRO = M.NMOVNRO " _
            & "         LEFT JOIN MOVME ME                          ON ME.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM = ME.NMOVITEM " _
            & "         LEFT JOIN MOVDOC MD                         ON MD.NMOVNRO = M.NMOVNRO  AND MD.nDocTpo NOT IN ('" & Format(TpoDocVoucherEgreso, "00") & "') " _
            & "         LEFT JOIN DOCUMENTO D1                      ON D1.nDocTpo = MD.nDocTpo " _
            & "         JOIN MOVREF MR                              ON MR.NMOVNRO = M.NMOVNRO " _
            & "         JOIN MOVARENDIR MAR                         ON MAR.NMOVNRO= MR.NMOVNROREF " _
            & "         JOIN " & vsBasePesonas & "PERSONA P         ON P.CPERSCOD = MAR.CPERSCOD " _
            & "         JOIN RRHH               RH                  ON RH.CPERSCOD= MAR.CPERSCOD " _
            & "         JOIN " & vsBaseComunes & "AREAS A           ON A.CAREACOD = RH.cAreaCodActual " _
            & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG  ON AG.CAGECOD = RH.cAgenciaActual " _
            & "         JOIN MOVDOC MDAR                            ON MDAR.NMOVNRO = MAR.NMOVNRO AND MDAR.nDocTpo IN ('" & Format(TpoDocRecViaticosARendirCuenta, "00") & "') " _
            & "         JOIN DOCUMENTO D2                           ON D2.nDocTpo = MDAR.nDocTpo "
sql = sql + " WHERE   MAR.cTpoArendir ='" & pnTipoArendir & "' And SUBSTRING(M.COPECOD,1,5) in ('" & Mid(gCGArendirViatAtencMN, 1, 5) & "','" & Mid(gCGArendirViatAtencME, 1, 5) & "') " & lsFiltroAge & lsFiltroArea _
            & "         AND MC.CCTACONTCOD IN ('" & psCtaArendir & "') And SUBSTRING(M.COPECOD,3,1)='" & pnMoneda & "' " _
            & "         AND M.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "') " _
            & "         AND MAR.CAREACOD ='" & psAreaArendir & "' AND MAR.CAGECOD ='" & psAgeArendir & "' " _
            & "         AND NOT EXISTS ( SELECT M1.CMOVNRO " _
            & "                          FROM   MOV M1 JOIN MOVREF MR1 ON M1.NMOVNRO=MR1.NMOVNRO " _
            & "                          Where  MR1.NMOVNROREF = MAR.NMovNro " _
            & "                                 AND M1.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "') " _
            & "                                 And SUBSTRING(M1.COPECOD,1,5) IN ('" & Left(gCGArendirViatSustMN, 5) & "','" & Left(gCGArendirViatSustME, 5) & "','" & Left(gCGArendirViatRendMN, 5) & "','" & Left(gCGArendirViatRendME, 5) & "') ) " _
            & "         AND  SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " _
            & " ORDER BY M.NMOVNRO, MDAR.cDocNro "

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

sql = " SELECT  M.cMovNro, M.cMovDesc, D.cDocAbrev, MD.cDocNro, MD.nDocTpo, p.cPersNombre,MD.dDocFecha, " _
    & "         SUM(ABS(IsNull(ME.NMOVMEIMPORTE, MC.NMOVIMPORTE))) As nDocImporte " _
    & " FROM    MOVREF  MRSUST " _
    & "         JOIN MOV M          ON MRSUST.nMOVNRO = M.nMOVNRO " _
    & "         JOIN MOVCTA MC      ON MC.nMOVNRO = M.nMOVNRO " _
    & "         LEFT JOIN MOVME ME  ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM =MC.nMOVITEM " _
    & "         JOIN MOVDOC MD      ON MD.nMOVNRO = M.nMOVNRO " _
    & "         JOIN " & vsBaseComunes & "DOCUMENTO D    ON D.nDocTpo = MD.nDocTpo " _
    & "         JOIN MovGasto MG  ON MG.nMOVNRO = M.nMOVNRO " _
    & "         JOIN " & vsBasePesonas & "PERSONA P      ON P.CPERSCOD = MG.CPERSCOD " _
    & " WHERE   MRSUST.nMOVNROREF =" & psMovNro & " AND cCtaContCod IN ('" & psCtaContArendir & "','" & psCtaContPendiente & "') " _
    & "         And  M.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagEliminado & "','" & gMovFlagModificado & "') " _
    & " GROUP BY M.cMovNro, M.cMovDesc, D.cDocAbrev, MD.cDocNro, MD.nDocTpo, p.cPersNombre,MD.dDocFecha"

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
    oMov.EliminaMov psMovNroSust, psMovNroActualiza
    oMov.ActualizaMovArendir psMovNroSol, pnTipoArendir, , , , lnSaldoArendir
    EliminaMovDocSust = 0
    Exit Function
EliminaMovDocSustErr:
    Call RaiseError(MyUnhandledError, "NARendir:EliminaMovDocSust Method")
End Function
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
      & "WHERE a.nMovNroRef = '" & pnMovNro & "' and m.nMovEstado IN (" & gMovEstContabMovContable & "," & gMovEstContabNoContable & ") and m.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') and " _
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


Public Sub GrabaSustentacionArendir(ByVal psFormatoFecha As String, ByVal pbNewProv As Boolean, ByVal psPersCod As String, ByVal psProvEstado As String, ByVal psUltiActProv As String, _
       ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
       ByVal pnMonto As Currency, ByVal pnDestino As Integer, _
       ByVal psCtaProvision As String, ByVal psCtaPendiente As String, _
       ByVal pnTipoArendir As ArendirTipo, pnTotArendir As Currency, _
       ByVal psDocNro As String, ByVal pnDocTpo As TpoDoc, ByVal pdDocFecha As Date, _
       ByVal pnTotal As Currency, ByVal pnMovNroAtencion As Long, _
       ByVal pnMovNroSol As Long, ByVal psAgeCod As String, ByVal psAreaCod As String, _
       ByVal rsCtasCont As ADODB.Recordset, ByVal rsObjetos As ADODB.Recordset, _
       ByVal rsImp As ADODB.Recordset, _
       Optional ByVal psAreaCh As String = "", Optional ByVal psAgeCh As String = "", _
       Optional ByVal pnNroProc As Integer = 0)
    
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
    Dim I As Integer
    
   On Error GoTo GrabaSustentacionArendirErr

   Dim nRendArendir     As Currency
   Dim nRendArendirMas  As Currency
   Dim nRendArendirImp     As Currency
   Dim nRendArendirImpMas  As Currency
   Dim nFaltaArendir    As Currency
   
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
    If pbNewProv Then
        oMov.InsertProv psPersCod, psProvEstado, psUltiActProv
    End If
    oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabSustPendRendir, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(psMovNro)
    oMov.InsertaMovCont lnMovNro, pnMonto, "0", "0"
    oMov.InsertaMovGasto lnMovNro, psPersCod, ""
    'Grabamos en MovCta, MovObj y MovCant
    lnItem = 0
    If Not rsCtasCont Is Nothing Then
        Do While Not rsCtasCont.EOF
            lsSubCta = ""
            lnItem = lnItem + 1
            'generamos la subcta recorriendo el recordset de los objetos
            If Not rsObjetos Is Nothing Then
                rsObjetos.MoveFirst
                Do While Not rsObjetos.EOF
                    If rsObjetos!CTACONT = rsCtasCont!Código And rsObjetos!ItemCtaCont = rsCtasCont!ItemCtaCont Then
                        lsSubCta = lsSubCta + rsObjetos![SubCta]
                    End If
                    rsObjetos.MoveNext
                Loop
            End If
            lnImporte = rsCtasCont!Monto
            For I = 6 To rsCtasCont.Fields.Count - 1
                If (pnDestino = 1 Or pnDestino = 2 Or rsCtasCont.Fields(I).Name = "AJUSTE") Then
                    lnImporte = lnImporte + rsCtasCont.Fields(I).value
                End If
            Next
            oMov.InsertaMovCta lnMovNro, lnItem, rsCtasCont!Código + lsSubCta, IIf(rsCtasCont!DH = "D", lnImporte, lnImporte * -1)
            lnOrdenObj = 0
            If Not rsObjetos Is Nothing Then
                rsObjetos.MoveFirst
                Do While Not rsObjetos.EOF
                    If rsObjetos!CTACONT = rsCtasCont!Código And rsObjetos!ItemCtaCont = rsCtasCont!ItemCtaCont Then
                        Select Case rsObjetos!Objpadre
                            Case ObjCMACAgencias
                                lnOrdenObj = lnOrdenObj + 1
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!Objpadre
                                oMov.InsertaMovObjAgenciaArea lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código, ""
                            Case ObjCMACArea
                                lnOrdenObj = lnOrdenObj + 1
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!Objpadre
                                oMov.InsertaMovObjAgenciaArea lnMovNro, lnItem, lnOrdenObj, "", rsObjetos!Código
                            Case ObjCMACAgenciaArea
                                lsAreaCod = Mid(rsObjetos!Código, 1, 3)
                                lsAgeCod = Mid(rsObjetos!Código, 4, 2)
                                lnOrdenObj = lnOrdenObj + 1
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!Objpadre
                                oMov.InsertaMovObjAgenciaArea lnMovNro, lnItem, lnOrdenObj, lsAgeCod, lsAreaCod
                            Case ObjDescomEfectivo
                                lnOrdenObj = lnOrdenObj + 1
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!Objpadre
                                oMov.InsertaMovObjEfectivo lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código
                            Case ObjEntidadesFinancieras
                                lnOrdenObj = lnOrdenObj + 1
                                oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!Objpadre
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
            For I = 6 To rsCtasCont.Fields.Count - 1
                rsImp.MoveFirst
                If rsCtasCont.Fields(I).value <> 0 Then
                    rsImp.Find "Impuesto='" & rsCtasCont.Fields(I).Name & "'"
                    If Not rsImp.EOF Then
                        oMov.InsertaMovOtrosItem lnMovNro, lnItem, rsImp!CTACONT, rsCtasCont.Fields(I).value, IIf(pnDestino = -1, "", pnDestino + 1)
                    End If
                End If
            Next
            rsCtasCont.MoveNext
        Loop
    End If
    If Not rsImp Is Nothing Then
        rsImp.MoveFirst
        Do While Not rsImp.EOF
            If rsImp!Ok = "1" And (pnDestino = -1 Or pnDestino = 0 Or pnDestino = 1) Then
                lnImporte = rsImp!Monto * IIf(rsImp!cDocImpDH = "D", 1, -1)
                If lnImporte <> 0 And rsImp!Impuesto <> "AJUSTE" Then
                    lnItem = lnItem + 1
                    oMov.InsertaMovCta lnMovNro, lnItem, rsImp!CTACONT, lnImporte
                End If
            End If
            rsImp.MoveNext
        Loop
    End If
    'Cuenta de Provisión / Caja Chica / A rendir Cuenta
    If nRendArendirImp <> 0 Then
       lnItem = lnItem + 1
       oMov.InsertaMovCta lnMovNro, lnItem, psCtaProvision, nRendArendirImp * -1
    End If
    If nRendArendirImpMas <> 0 Then
       lnItem = lnItem + 1
       oMov.InsertaMovCta lnMovNro, lnItem, psCtaPendiente, nRendArendirImpMas * -1
    End If
    ''Grabamos Documento en MovDoc
    oMov.InsertaMovDoc lnMovNro, pnDocTpo, psDocNro, Format(pdDocFecha, gsFormatoFecha)
    'Grabamos la Referencia al Movimiento que se regulariza
    oMov.InsertaMovRef lnMovNro, pnMovNroAtencion
    If pnTipoArendir = gArendirTipoCajaChica Then
       oMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, pnNroProc, gCHTipoProcArendir
    End If
    
    'Actualizamos el saldo del aRendir
    oMov.ActualizaMovArendir pnMovNroSol, pnTipoArendir, , , , pnTotArendir - (nRendArendir + nRendArendirMas + nRendArendirImp + nRendArendirImpMas)
    oMov.InsertaMovArendirSust pnMovNroSol, lnMovNro

    If Mid(psOpeCod, 3, 1) = gMonedaExtranjera Then
        oMov.GeneraMovME lnMovNro, psMovNro
    End If
    
    'Sustentación de Arendir no pasa a Contabilidad
'    oMov.ActualizaSaldoMovimiento psMovNro, "+"
    
    oMov.CommitTrans
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
    Case gArendirTipoCajaGeneral
        sql = "SELECT   M.CMOVNRO AS CMOVSOL , MAT.CMOVNRO AS CMOVATENC," _
        & "             MAR.CPERSCOD , P.CPERSNOMBRE, ISNULL(A.CAREACOD,'') CAREACOD, ISNULL(A.CAREADESCRIPCION,'') CAREADESCRIPCION, " _
        & "             ISNULL(AG.CAGECOD,'') AS CAGECOD  , ISNULL(AG.CAGEDESCRIPCION,'') CAGEDESCRIPCION , ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS ImporteArendir,  " _
        & "             MAR.CAREACOD as cAreaCodArendir, A1.CAREADESCRIPCION as cAreaARendir , " _
        & "             ISNULL(MAR.CAGECOD,'') AS cAgeCodArendir , ISNULL(AG1.CAGEDESCRIPCION,'') cAgeArendir , " _
        & "             MAR.NMOVSALDO, MD.nDocTpo , D.CDOCABREV , MD.CDOCNRO, MD.dDocFecha as dDocFechaSol , " _
        & "             ISNULL(MD1.nDocTpo,'') AS nDocTpoATENC, ISNULL(D1.CDOCABREV,'') AS CDOCABREVATENC, ISNULL(MD1.CDOCNRO,'') AS CDOCNROATENC, ISNULL(D1.cDocDesc,'')  AS cDocDescAtenc, MD1.dDocFecha as dFechaDocAtenc, " _
        & "             SUST.cCtaContCod as CtaContSust, ABS(SUST.Debe) AS Debe , ABS(SUST.HABER) HABER " _
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
            & "         A.CAREACOD, A.CAREADESCRIPCION, ISNULL(AG.CAGECOD,'') AS CAGECOD  , ISNULL(AG.CAGEDESCRIPCION,'') CAGEDESCRIPCION , " _
            & "         MAR.CAREACOD as cAreaCodArendir, A1.CAREADESCRIPCION as cAreaARendir , ISNULL(MAR.CAGECOD,'') AS cAgeCodArendir , ISNULL(AG1.CAGEDESCRIPCION,'') cAgeArendir , " _
            & "         ABS(REND.Debe) AS ImporteArendir,  MAR.NMOVSALDO, MD.nDocTpo, " _
            & "         (D.CDOCABREV + '-' + MD.CDOCNRO) AS CDOCNRO, MD.dDocFecha as dDocFechaSol , " _
            & "         MD.dDocFecha as dFechaDocAtenc , MD.CDOCNRO as cDocNroAtenc , D.CDOCABREV AS cDocDescAtenc, " _
            & "         SUST.cCtaContCod as CtaContSust, ABS(SUST.Debe) AS Debe , ABS(SUST.HABER) HABER " _
            & " FROM    MOVARENDIR MAR " _
            & "         JOIN MOV M                                  ON M.nMOVNRO = MAR.nMOVNRO " _
            & "         JOIN " & vsBasePesonas & "PERSONA P         ON P.CPERSCOD = MAR.CPERSCOD " _
            & "         JOIN RRHH RH                                ON RH.CPERSCOD = P.CPERSCOD " _
            & "         JOIN " & vsBaseComunes & "AREAS A           ON A.CAREACOD = RH.cAreaCodActual   " _
            & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG  ON AG.CAGECOD = RH.cAgenciaActual " _
            & "         JOIN " & vsBaseComunes & "AREAS A1          ON A1.CAREACOD = MAR.CAREACOD " _
            & "         LEFT JOIN " & vsBaseComunes & "AGENCIAS AG1 ON AG1.CAGECOD = MAR.CAGECOD " _
            & "         JOIN MOVDOC MD                              ON MD.NMOVNRO = M.NMOVNRO  AND MD.nDocTpo ='" & TpoDocRecViaticosARendirCuenta & "' " _
            & "         JOIN DOCUMENTO D                            ON D.nDocTpo = MD.nDocTpo "
    
    sql = sql + "      Join (SELECT    MC.cCtaContCod, " _
            & "                         SUM ( CASE " _
            & "                                 WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)>0 THEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE) " _
            & "                                 WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)<0 THEN 0 END ) AS DEBE, " _
            & "                         SUM ( CASE " _
            & "                         WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)<0 THEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE) " _
            & "                         WHEN IsNull(MC.NMOVIMPORTE, ME.NMOVMEIMPORTE)>0 THEN 0 END ) AS HABER, " _
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
    Case gArendirTipoCajaGeneral
    
        sql = "     Select  M.CMOVNRO AS CMOVNROSOL, MR.nMOVNRO AS nMOVATENC , " _
            & "             SUST.cCtaContCod as CtaContSust, ABS(SUST.Debe) Debe , ABS(SUST.HABER) HABER , " _
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
            & "         SUM (C.DEBE) AS DEBE , SUM(C.HABER) AS HABER " _
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
            & "         AND substring(copecod,1,5) NOT IN ('" & Mid(gCGArendirViatAtencMN, 1, 5) & "','" & Mid(gCGArendirViatAtencMN, 1, 5) & "') " _
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
ActualizaMovSust pnMovNroSol, psMovNro, oDMov, , pnTipoArendir, pnNroProc
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
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaRendicionExacta", "Error Grabar Rendicion a Arendir Exacta "
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
    lsFiltroCta = " AND OC.CCTACONTCOD IS NOT NULL AND OC.cCtaContCod not in ('" & psCtaContArendir & "') " ','" & psCtaContPend & "')"
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
    GetOpeRendicion = rs!COPECOD
End If
rs.Close
Set rs = Nothing
oConect.CierraConexion
Set oConect = Nothing
End Function
Public Function GrabaRendicionEfectivo(ByVal pnTipoArendir As ArendirTipo, ByVal psFormatoFecha As String, _
                                        ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                        ByVal psCtaPendiente As String, ByVal psCtaEfectivo As String, _
                                        ByVal pnImporte As Currency, ByVal rsBilletaje As ADODB.Recordset, _
                                        ByVal psMovAtencion As String, ByVal pnMovNroSol As Long, _
                                        ByVal pbIngreso As Boolean, _
                                        ByVal psCtaDiferencia As String, ByVal pnMontoDiferencia As Currency) As Integer
            
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

oDMov.InsertaMovRef lnMovNro, psMovAtencion
oDMov.ActualizaMovArendir pnMovNroSol, pnTipoArendir, , , , 0
ActualizaMovSust pnMovNroSol, psMovNro, oDMov, , pnTipoArendir
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
                                         ByVal psCtaDiferencia As String, ByVal pnMontoDiferencia As Currency) As Integer
Dim lnMovItem As Long
Dim lnMovOrden As Integer
Dim oDMov As DMov
Dim lsMovNro As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Dim lnImporteVenta As Currency

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
        Dim oNeg As New NNegOpePendientes
        Do While Not rsPago.EOF
            If rsPago.Fields(1).value = "1" Then   'Check Seleccionado
               lnImporteVenta = lnImporteVenta + rsPago.Fields(5).value
               If rsPago.Fields(4) = rsPago.Fields(5) Then
                  'Actualizamos referencia en Negocio
                  oNeg.ActualizaNegocioReferencia rsPago.Fields(10), rsPago.Fields(7), psMovNro
               Else
                  oDMov.InsertaMovPendientesRend lnMovNro, rsPago!cCtaContCod, rsPago.Fields(4) - rsPago.Fields(5)
               End If
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

oDMov.InsertaMovRef lnMovNro, psMovAtencion
oDMov.ActualizaMovArendir pnMovNroSol, pnTipoArendir, , , , 0
ActualizaMovSust pnMovNroSol, psMovNro, oDMov, , pnTipoArendir
GrabaRendicionVentanilla = 0
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


Private Function ActualizaMovSust(ByVal pnMovNroSol As Long, ByVal psMovRendicion As String, ByRef oMov As DMov, Optional ByVal pbExtorno As Boolean = False, Optional pnArendirTipo As ArendirTipo = gArendirTipoCajaGeneral, Optional pnProcNro = 0)
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsMovNroNew As String

Set rs = New ADODB.Recordset
sql = " Select  M.cMovNro, MA.nMovNro, MA.nMovNroSust, MA.nMovNroRend    " _
    & " From    Mov M JOIN MovArendirSust MA ON " & IIf(pbExtorno, "MA.nMovNroRend", "MA.nMOVNROSUST") & "= M.nMOVNRO " _
    & " Where   MA.nMovNro = " & pnMovNroSol & " AND M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagExtornado & "','" & gMovFlagModificado & "') "

Set rs = oMov.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    Do While Not rs.EOF
        If pbExtorno = False Then
           If Left(rs!cMovNro, 8) <> Left(psMovRendicion, 8) Then
              lsMovNroNew = NuevoMovNro(psMovRendicion, oMov)
              'Actualizamos los cMovNro de las Sustentaciones con la misma fecha de la
              'Rendición
              If pnArendirTipo = gArendirTipoCajaChica Then
                oMov.ActualizaMovimiento lsMovNroNew, rs!nMovNroSust, gMovEstContabNoContable, , , False
                If pnProcNro > 0 Then
                    oMov.ActualizaMovCajaChicaNroProc oMov.GetnMovNro(lsMovNroNew), pnProcNro
                End If
              Else
                oMov.ActualizaMovimiento lsMovNroNew, rs!nMovNroSust, gMovEstContabMovContable, , , False
              End If
           Else
              lsMovNroNew = NuevoMovNro(psMovRendicion, oMov)
              
              If pnArendirTipo = gArendirTipoCajaChica Then
                oMov.ActualizaMovDatos rs!nMovNroSust, lsMovNroNew, gMovEstContabNoContable
              Else
                oMov.ActualizaMovDatos rs!nMovNroSust, lsMovNroNew, gMovEstContabMovContable
              End If
           End If

        Else
            oMov.ActualizaMovimiento oMov.GetcMovNro(rs!nMovNroSust), rs!nMovNroRend, gMovEstContabSustPendRendir, , , True
        End If
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing
End Function
Private Function NuevoMovNro(ByVal psMovNroRend As String, ByRef oMov As DMov) As String
Dim ldFechaArendir As Date
Dim lsAge As String
Dim lsCmac As String
Dim lsUser As String

ldFechaArendir = CDate(Mid(psMovNroRend, 7, 2) & "/" & Mid(psMovNroRend, 5, 2) & "/" & Mid(psMovNroRend, 1, 4))
lsCmac = Mid(psMovNroRend, 15, 3)
lsAge = Mid(psMovNroRend, 18, 2)
lsUser = Right(psMovNroRend, 4)
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
ActualizaMovSust psMovNroSol, psMovNro, oDMov, , pnTipoArendir
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
oDMov.InsertaChequeEstado TpoDocCheque, psNumCheque, psPersCodIf, psTipoIF, pdFechaReg, gChqEstEnValorizacion, psMovNro

oDMov.InsertaMovRef lnMovNro, pnMovAtencion
ActualizaMovSust pnMovNroSol, psMovNro, oDMov, , pnTipoArendir
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
                                            ByVal psCtaHaber As String, ByVal pnImporte As Currency, _
                                            ByVal psDocTpo As TpoDoc, ByVal psDocNro As String, _
                                            ByVal psFechaDoc As String, ByVal psDocNroVoucher As String, _
                                            ByVal psPersCodIf As String, ByVal psTipoIF As String, ByVal psCodCtaIf As String, _
                                            Optional ByVal psAreaCh As String = "", Optional ByVal psAgeCh As String = "", _
                                            Optional pnNroProc As Integer = 0, Optional pnMotivo As Long = -1, Optional psCtaCod As String = "") As Integer
            
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
    
    'INSERTAMOS DENTRO DE  LOS MOVIMIENTOS DE CUENTAS DE INSTITUCIONES FINANCIERAS
'    oDMov.InsertaMovCtaIF lnMovNro, psPersCodIf, psTipoIF, psCodCtaIf, CGCtaIFConCapital, pnImporte * -1
'    oDMov.ActualizaSaldoCtaIF CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Mid(psMovNro, 1, 4)), _
'                                    psPersCodIf, psTipoIF, psCodCtaIf
End If
Select Case Val(psDocTpo)
    Case TpoDocNotaAbono, TpoDocNotaCargo
        oDMov.InsertaNotaAbonoCargo psDocTpo, psDocNro, gNCNARegistrado, pnMotivo, Abs(pnImporte)
        oDMov.InsertaNotaAbonoCargoEst psDocTpo, psDocNro, gNCNARegistrado, psMovNro
        oDMov.InsertaRegDocCuenta psDocTpo, psDocNro, psCtaCod
    Case TpoDocOrdenPago
        
End Select
If psDocNro <> "" Then
    oDMov.InsertaMovDoc lnMovNro, psDocTpo, psDocNro, psFechaDoc
End If
If psDocNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psDocNroVoucher, psFechaDoc
End If
oDMov.InsertaMovRef lnMovNro, psMovNroAtenc
ActualizaMovSust pnMovNroSol, psMovNro, oDMov, , pnTipoArendir
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
                               Optional pnNroProc As Integer = 0) As Integer
    
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
    oMov.ExtornaMovimiento psMovNroExt, pnMovNroAnt, psOpeCod, psMovDesc
    
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
Public Function GetRendicionesArendir(ByVal psAgeCod As String, ByVal psAreaCod As String, _
                                        ByVal psAreaArendir As String, ByVal psAgeArendir As String, _
                                        ByVal pnTipoArendir As ArendirTipo, _
                                        ByVal psCtaArendir As String, ByVal psCtaPendiente As String, _
                                        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal pnMoneda As Moneda) As ADODB.Recordset
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
    Case gArendirTipoCajaGeneral
            
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
            & "                         And SUBSTRING(M1.COPECOD,1,5) IN ('" & Mid(gCGArendirCtaRendMN, 1, 5) & "','" & Mid(gCGArendirCtaRendME, 1, 5) & "','" & Mid(gCGArendirCtaRendDevMN, 1, 5) & "','" & Mid(gCGArendirCtaRendDevME, 1, 5) & "') " _
            & "                         AND SUBSTRING(M1.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " _
            & "                 GROUP BY  O.cOpeCod,O.cOpeDesc, M1.cMOVNRO,M1.CMOVDESC , MR1.nMOVNROREF, D.CDOCABREV,MD.nDocTpo, MD.cDocNro, MD.dDocFecha , M1.nMOVNRO, MCP.cCtaCod) AS Rend " _
            & "         ON Rend.nMovAtenc= M.nMOVNRO " _
            & " WHERE   MAR.cTpoArendir ='" & pnTipoArendir & "'  " & lsFiltroAge & lsFiltroArea _
            & "         AND MC.CCTACONTCOD IN ('" & psCtaArendir & "') AND SUBSTRING(M.COPECOD,3,1)='" & pnMoneda & "'  " _
            & "         AND MAR.CAREACOD ='" & psAreaArendir & "' AND MAR.CAGECOD ='" & psAgeArendir & "' " _
            & "         And M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
            & " ORDER BY Rend.cMovRendicion, MDAR.cDocNro "
   
    Case gArendirTipoViaticos
        
        sql = "SELECT   Rend.cOpeDesc, Rend.cDocAbrevRend, Rend.cDocNroRen, ISNULL(CONVERT(CHAR(10),Rend.dDocFechaRend,103), " _
            & "         SUBSTRING(Rend.cMovRendicion,7,2) + '/' + SUBSTRING(Rend.cMovRendicion,5,2) + '/' + SUBSTRING(Rend.cMovRendicion,1,4) ) AS dDocFechaRend, " _
            & "         MDAR.cDocNro as cDocNroSol, CONVERT(CHAR(10),MDAR.dDocFecha,103) AS dDocSolicitud, " _
            & "         CONVERT(CHAR(50),P.CPERSNOMBRE) AS cPersNombre, ISNULL(Abs(Rend.MontoRendicion),0) AS MontoRend, " _
            & "         ISNULL(Rend.MontoRendicion,0) AS MontoRendReal,  0 AS NIMPORTE, " _
            & "         MAR.nMovSaldo, ISNULL(A.cAreaDescripcion,'') cAreaDescripcion , " _
            & "         P.cPersCod, Rend.NMovRendicion, M.NMOVNRO as cMovNroAtenc, MAR.NMOVNRO as NMovNroSol, " _
            & "         Rend.CMOVDESC, Rend.nDocTpoRend,MDAR.nDocTpo as nDocTpoSol, " _
            & "         ISNULL(A.cAreaCod,'') cAreaCod ,ISNULL(AG.cAgeCod,'') cAgeCod, " _
            & "         ISNULL(AG.cAgeDescripcion,'') cAgeDescripcion , Rend.cOpeCod , D2.cDocAbrev  as cDocAbrevSol , Rend.cCtaCod " _
            & " FROM    MOV M  " _
            & "         JOIN MOVARENDIR MAR     ON M.NMOVNRO= MAR.NMOVNRO   " _
            & "         JOIN PERSONA P          ON P.CPERSCOD  = MAR.CPERSCOD " _
            & "         JOIN RRHH RH            ON RH.CPERSCOD  = P.CPERSCOD " _
            & "         JOIN AREAS A            ON A.CAREACOD = RH.cAreaCodActual " _
            & "         LEFT JOIN AGENCIAS AG   ON AG.CAGECOD = RH.cAgenciaActual " _
            & "         JOIN MOVDOC MDAR        ON MDAR.NMOVNRO = M.NMOVNRO  AND MDAR.nDocTpo IN (" & TpoDocRecViaticosARendirCuenta & ") " _
            & "         JOIN DOCUMENTO D2       ON D2.nDocTpo = MDAR.nDocTpo "
        
        sql = sql + "         JOIN (  SELECT  O.cOpeCod, O.cOpeDesc, M1.NMOVNRO AS NMovRendicion, M1.CMOVNRO AS CMovRendicion, " _
            & "                         M1.CMOVDESC , MR1.NMOVNROREF as NMovNroSol , " _
            & "                         ISNULL(D.CDOCABREV,'') AS cDocAbrevRend, ISNULL(MD.nDocTpo,'') as nDocTpoRend, " _
            & "                         MD.dDocFecha  AS dDocFechaRend, ISNULL(MD.cDocNro,'') as cDocNroRen, " _
            & "                         SUM(ISNULL(Me.NMOVMEIMPORTE, MC.NMOVIMPORTE)) As MontoRendicion , ISNULL(MCP.cCtaCod,'') AS cCtaCod " _
            & "                 FROM    MOV M1  JOIN MOVREF MR1 ON M1.NMOVNRO=MR1.NMOVNRO " _
            & "                         JOIN OPETPO O           ON O.COPECOD = M1.COPECOD " _
            & "                         LEFT JOIN MOVCTA MC     ON MC.NMOVNRO = M1.NMOVNRO AND MC.CCTACONTCOD IN ('" & psCtaArendir & "','" & psCtaPendiente & "') " _
            & "                         LEFT JOIN MOVME ME      ON ME.NMOVNRO = MC.NMOVNRO AND ME.NMOVITEM = MC.NMOVITEM " _
            & "                         LEFT JOIN MOVDOC MD     ON MD.NMOVNRO = M1.NMOVNRO  AND MD.nDocTpo NOT IN (" & TpoDocVoucherEgreso & ") " _
            & "                         LEFT JOIN DOCUMENTO D   ON D.nDocTpo = MD.nDocTpo " _
            & "                         LEFT JOIN MOVCAP MCP    ON MCP.NMOVNRO = M1.NMOVNRO " _
            & "                 WHERE   SUBSTRING(M1.COPECOD,3,1)='" & pnMoneda & "' " _
            & "                         AND M1.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "')  " _
            & "                         And SUBSTRING(M1.COPECOD,1,5) IN ('" & Mid(gCGArendirViatRendMN, 1, 5) & "','" & Mid(gCGArendirViatRendME, 1, 5) & "') AND MC.NMOVIMPORTE IS NOT  NULL " _
            & "                         AND SUBSTRING(M1.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " _
            & "                 GROUP BY  O.cOpeCod,O.cOpeDesc,M1.NMOVNRO, M1.CMOVNRO ,M1.CMOVDESC , MR1.NMOVNROREF, " _
            & "                         D.CDOCABREV,MD.nDocTpo, MD.cDocNro, MD.dDocFecha, MCP.cCtaCod) AS Rend " _
            & "          ON Rend.NMovNroSol= MAR.NMOVNRO " _
            & "  WHERE   MAR.cTpoArendir ='" & pnTipoArendir & "' " & lsFiltroAge & lsFiltroArea _
            & "          AND Substring(M.COPECOD,3,1)='" & pnMoneda & "' " _
            & "          AND MAR.CAREACOD ='" & psAreaArendir & "' AND MAR.CAGECOD ='" & psAgeArendir & "' " _
            & "          And M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "')  " _
            & "  ORDER BY Rend.nMovRendicion, MDAR.cDocNro "
        
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
    & "         ISNULL(NOTRANSP.ImporteViaticos,0) + ISNULL(IDA.ImporteViaticos,0) + ISNULL(VUELTA.ImporteViaticos,0) AS ImporteConcepto " _
    & " FROM    OBJETO O " _
    & "         Left Join " _
    & "         (Select  O.cObjetoCod, " _
    & "                  CASE " _
    & "                         WHEN cViaticoAfectoA = '" & gViaticoAfectoADias & "' THEN nViaticoImporte*" & pnNumDias & "" _
    & "                 WHEN cViaticoAfectoA IN ('" & gViaticoAfectoANinguno & "','" & gViaticoAfectoATransporte & "')  THEN nViaticoImporte " _
    & "                 END As ImporteViaticos, nViaticoImporte " _
    & "          From    Viaticos V JOIN " & vsBaseComunes & "Objeto O on O.cObjetoCod = V.cObjetoCod  " _
    & "          Where   O.cObjetoCod Like '" & ObjConceptosARendir & "__' And cCategCod ='" & psCategCod & "' and cDestinoCod ='" & psDestinoCod & "' " _
    & "                  and cViaticoAfectoA NOT IN ('" & gViaticoAfectoATransporte & "') " _
    & "         GROUP BY  O.cObjetoCod , cViaticoAfectoA , nViaticoImporte ) as NOTRANSP " _
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
Dim pnMovViaticosDias As String
Dim psDestinoDesc As String
Dim lsMovNro As String
Dim lsGlosa As String
Dim lsMovNroViat As String
Dim lnItem As Long
Dim lnImporte As Currency
Dim I As Integer

Dim lnMovNro As Long
Dim lnMovNroViat As Long
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
    If psNroDocViatico <> "" Then
       'Grabamos en MovDoc el Recibo de Egresos de viaticos
       oMov.InsertaMovDoc lnMovNroViat, TpoDocRecViaticosARendirCuenta, psNroDocViatico, Format(pdFechaViatico, gsFormatoFecha)
    End If
    If psMovNroDocRef <> "" Then
        oMov.InsertaMovDoc lnMovNroViat, TpoDocCarta, psMovNroDocRef, Format(pdFechaDocRef, gsFormatoFecha)
    End If
    oMov.InsertaMovARendir lnMovNroViat, gArendirTipoViaticos, psAreaCod, psAgeCod, psPersCod, 0, psCategCod
Else
    lnMovNroViat = oMov.GetnMovNro(lsMovNroViat)
End If
Do While Not rsDetViat.EOF
    If rsDetViat!Importe > 0 And Trim(rsDetViat!cMovNroDet) = "" Then
        lsMovNro = oMov.GeneraMovNro(pdFecsis, Mid(lsMovNroViat, 18, 2), Right(lsMovNroViat, 4))
        lsGlosa = rsDetViat!Motivo
        oMov.InsertaMov lsMovNro, psOpeCod, lsGlosa, gMovEstContabNoContable, gMovFlagVigente
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        
        oMov.InsertaMovCont lnMovNro, rsDetViat!Importe, "0", "0"
        psDestinoCod = Val(Right(Trim(rsDetViat!Destino), 2))
        psTranspIda = Val(Right(Trim(rsDetViat!Ida), 2))
        psTranspVuelta = Val(Right(Trim(rsDetViat!Vuelta), 2))
        pdPartida = rsDetViat!Partida
        pnMovViaticosDias = rsDetViat!Dias
        psDestinoDesc = rsDetViat!Lugar
        
        oMov.InsertaMovViaticos lnMovNroViat, lnMovNro, psDestinoCod, psTranspIda, psTranspVuelta, pdPartida, pnMovViaticosDias, psDestinoDesc
        
        lnItem = 0
        For I = 1 To rsAux.Fields.Count - 1
            lnImporte = rsAux.Fields(I).value
            If lnImporte > 0 Then
                lnItem = lnItem + 1
                oMov.InsertaMovObj lnMovNro, lnItem, 1, ObjConceptosARendir
                oMov.InsertaMovObjViaticos lnMovNro, lnItem, rsAux.Fields(I).Name, lnImporte
            End If
        Next
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

sql = " SELECT  M.CMOVNRO, MV.nMOVNRO, M.cMovDesc AS Motivo , MV.cDestinoDesc as Lugar,  " _
    & "         dPartida, nMovViaticosDias, " _
    & "         C.nConsValor as CodCat,  C.cConsDescripcion  as Categoria , " _
    & "         C1.nConsValor as CodDestino,  C1.cConsDescripcion  as Destino , " _
    & "         C2.nConsValor as CodTrasnIda,  C2.cConsDescripcion  as TransIda,  " _
    & "         C3.nConsValor as CodTransVuelta,  C3.cConsDescripcion  as TrasnVuelta , MC.nMovMonto , M.cOpeCod, ISNULL(Ref.nMovNro,0) nMovNroAtencion " _
    & " FROM    MOVVIATICOS MV  JOIN MOVARENDIR MAR ON MAR.nMOVNRO = MV.nVIATICOMOVNRO " _
    & "         JOIN MOV    M   ON M.nMOVNRO = MV.nMovNro " _
    & "         LEFT JOIN MOVCONT    MC  ON MC.nMOVNRO =M.nMOVNRO " _
    & "         JOIN " & vsBaseComunes & "CONSTANTE C    ON C.nConsValor  = MAR.cCategCod " _
    & "         JOIN " & vsBaseComunes & "CONSTANTE C1   ON C1.nConsValor  = MV.cDestinoCod " _
    & "         JOIN " & vsBaseComunes & "CONSTANTE C2   ON C2.nConsValor  = MV.cTranspIda " _
    & "         JOIN " & vsBaseComunes & "CONSTANTE C3   ON C3.nConsValor  = MV.cTranspVuelta " _
    & "         LEFT JOIN (SELECT m1.nMovNro, mr1.nMovNroRef FROM Mov m1 JOIN MovRef mr1 ON mr1.nMovNro = m1.nMovNro " _
    & "                    WHERE m1.nMovEstado = " & gMovEstContabMovContable & " and not nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") " _
    & "                   ) Ref ON Ref.nMovNroRef = M.nMovNro " _
    & " WHERE   M.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') " _
    & "         AND C.nCONSCOD like '" & gViaticosCateg & "'  AND C1.nCONSCOD like '" & gViaticosDestino & "' " _
    & "         AND C2.nCONSCOD like '" & gViaticosTransporte & "' AND C3.nCONSCOD like '" & gViaticosTransporte & "' " _
    & "         AND MV.nViaticoMovNro = '" & psMovnroViaticos & "' "

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

sql = " SELECT  M1.CMOVNRO, MC.cObjetoCod , ISNULL(MC.nMovImporte,0 ) as Importe " _
    & " FROM    MOVVIATICOS     MV " _
    & "         JOIN MOV    M   ON M.nMOVNRO = MV.nViaticoMovNro " _
    & "         JOIN MOV        M1  ON M1.nMOVNRO = MV.NMOVNRO " _
    & "         JOIN MOVOBJVIATICOS     MC  ON MC.nMOVNRO = M1.NMOVNRO " _
    & "         " _
    & "          " _
    & " WHERE   M1.nMovFlag NOT IN('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') AND  " _
    & "         M1.CMOVNRO = '" & psMovNroDetViat & "'  "

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
Public Function EmiteCajasChicasNuevas() As ADODB.Recordset
On Error GoTo ErrorEmiteCajasChicasNuevas
Dim rs As ADODB.Recordset
Dim sSQL As String
Set rs = New ADODB.Recordset

Dim oConect As DConecta
Set oConect = New DConecta

If oConect.AbreConexion = False Then Exit Function

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
bInactiva = IIf(rsCta("bInactiva") = 1, True, False)
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
ActualizaMovSust pnMovNroSol, sMovNro, oMov, , pnTipoArendir
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
Dim nNumExtracto As Long, nDiasTranscurridos As Long, I As Long
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
For I = 0 To Rnd(2000) * 1000
Next I
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
ActualizaMovSust pnMovNroSol, sMovNro, oMov, , pnTipoArendir
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
Dim nNumExtracto As Long, nDiasTranscurridos As Long, I As Long
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
For I = 0 To Rnd(2000) * 1000
Next I
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
bInactiva = IIf(rsCta("bInactiva") = 1, True, False)
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





