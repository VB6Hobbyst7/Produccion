VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCajaChica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim lsBaseCom As String
Dim lsBasePers As String

Dim oError As COMConecta.COMErrorHandling

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
End Enum
Private Sub Class_Initialize()
Dim oIni As COMConecta.DCOMClasIni
Set oIni = New COMConecta.DCOMClasIni
Set oError = New COMConecta.COMErrorHandling
lsBaseCom = oIni.BaseComunes
lsBasePers = oIni.BasePersonas
Set oIni = Nothing
End Sub
Public Function GetDatosCajaChica(ByVal psAreaCh As String, ByVal psAgeCh As String, Optional ByVal DatosCajaChica As DatosCajaChica = SaldoActual) As Currency
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function

sql = "SELECT  cAreaCod, cAgeCod, nProcNro, nMontoAsig, nTopeEgresos,nMontoDesem,nSaldoAnt,nSaldo " _
    & " FROM   CAJACHICA  C " _
    & " WHERE  C.nProcNro = (SELECT MAX(nProcNro) " _
    & "                      FROM CAJACHICA C1 " _
    & "                      WHERE C1.cAreaCod = C.cAreaCod and C1.cAgeCod = C.cAgeCod ) " _
    & "        And cAreaCod='" & psAreaCh & "' and cAgeCod= '" & psAgeCh & "'"
    
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
'Dim oCon As COMConecta.DCOMConecta
'Set oCon = New COMConecta.DCOMConecta
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
'Dim oCon As COMConecta.DCOMConecta
'Set oCon = New COMConecta.DCOMConecta
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
Dim oGen As COMDConstSistema.DCOMGeneral
Dim lnTasaCajaCh As Currency
Dim lnSaldo As String
Set oGen = New COMDConstSistema.DCOMGeneral

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
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

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
    & "                         WHERE   MR.nMOVNROREF=M.nMOVNRO AND M1.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "') ) " _
    & "         AND M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "') " _
    & "         AND MCH.cAreaCod='" & psAreaCh & "' AND MCH.cAgeCod='" & psAgeCh & "' AND nProcNro=" & pnProcCH

Set rs = oCon.CargaRecordSet(sql)
Set GetSolicitudesArendirCH = rs

oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetAtencionesArendir(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnProcCH As Integer, ByVal psCtaArendir As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function

sql = " SELECT  MD.CDOCNRO, CONVERT(CHAR(12), MD.dDocFecha ,103  ) AS Fecha ,P.CPERSNOMBRE, ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS IMPORTE,  " _
    & "         nMovSaldo, M1.CMOVDESC AS DESCRIPCION , ISNULL(A.CAREADESCRIPCION,'')  AS CAREADESC,  " _
    & "         M.nMOVNRO AS CMOVNROSOL, M1.nMOVNRO AS CMOVNROATENC , MAR.CAREACOD, " _
    & "         ISNULL(AG.CAGEDESCRIPCION,'') AS CAGEDESC, MAR.CAGECOD , MAR.CPERSCOD " _
    & "  FROM   MOV  M " _
    & "         JOIN MOVCONT MCT ON MCT.nMOVNRO = M.nMOVNRO " _
    & "         JOIN MOVARENDIR MAR ON MAR.nMOVNRO = M.nMOVNRO " _
    & "         JOIN PERSONA P     ON P.CPERSCOD = MAR.CPERSCOD " _
    & "         LEFT JOIN AREAS A ON A.CAREACOD = MAR.CAREACOD " _
    & "         LEFT JOIN AGENCIAS AG ON AG.CAGECOD=MAR.CAGECOD " _
    & "         JOIN MOVCAJACHICA MCH ON MCH.nMOVNRO = MAR.nMOVNRO " _
    & "         JOIN MOVDOC MD ON MD.nMOVNRO = M.nMOVNRO " _
    & "         JOIN DOCUMENTO  D ON D.nDocTpo = MD.nDocTpo " _
    & "         JOIN MOVREF MR ON MR.nMOVNROREF= M.nMOVNRO " _
    & "         JOIN MOV M1 ON M1.nMOVNRO = MR.nMOVNRO " _
    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M1.nMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "  WHERE  MAR.cTpoArendir='" & gArendirTipoCajaChica & "' AND cCtaContCod in ('" & psCtaArendir & "') " _
    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "         AND M1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "         AND MCH.cAreaCod='" & psAreaCh & "' AND MCH.cAgeCod='" & psAgeCh & "' AND nProcNro=" & pnProcCH & " " _
    & "         AND NOT EXISTS ( SELECT M2.CMOVNRO FROM MOV M2 JOIN MOVREF MR2 ON M2.nMOVNRO = MR2.nMOVNRO " _
    & "                          WHERE  MR2.nMOVNROREF =M1.nMOVNRO AND SUBSTRING(M2.COPECOD,1,5) IN ('" & Mid(gCHArendirCtaSustMN, 1, 5) & "','" & Mid(gCHArendirCtaSustME, 1, 5) & "')  " _
    & "                                 AND M2.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ")) "

Set rs = oCon.CargaRecordSet(sql)
Set GetAtencionesArendir = rs

oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetCHSustSinRendicion(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnProcCH As Integer, ByVal psCtaArendir As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function

sql = " SELECT  MD.CDOCNRO, CONVERT(CHAR(12),MD.DDOCFECHA,103 ) AS Fecha ,P.CPERSNOMBRE, ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS IMPORTE,  " _
    & "         nMovSaldo, M1.CMOVDESC AS DESCRIPCION , ISNULL(A.CAREADESCRIPCION,'')  AS CAREADESC,  " _
    & "         M.nMOVNRO AS CMOVNROSOL, M1.nMOVNRO AS CMOVNROATENC , MAR.CAREACOD, " _
    & "         ISNULL(AG.CAGEDESCRIPCION,'') AS CAGEDESC, MAR.CAGECOD, MAR.CPERSCOD, MCH.nProcNro " _
    & "  FROM   MOV  M " _
    & "         JOIN MOVCONT MCT ON MCT.nMOVNRO = M.nMOVNRO " _
    & "         JOIN MOVARENDIR MAR ON MAR.nMOVNRO = M.nMOVNRO " _
    & "         JOIN PERSONA P     ON P.CPERSCOD = MAR.CPERSCOD " _
    & "         LEFT JOIN AREAS A ON A.CAREACOD = MAR.CAREACOD " _
    & "         LEFT JOIN AGENCIAS AG ON AG.CAGECOD=MAR.CAGECOD " _
    & "         JOIN MOVCAJACHICA MCH ON MCH.nMOVNRO = MAR.nMOVNRO " _
    & "         JOIN MOVDOC MD ON MD.nMOVNRO = M.nMOVNRO " _
    & "         JOIN DOCUMENTO  D ON D.nDocTpo = MD.nDocTpo " _
    & "         JOIN MOVREF MR ON MR.nMOVNROREF= M.nMOVNRO " _
    & "         JOIN MOV M1 ON M1.nMOVNRO = MR.nMOVNRO " _
    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M1.nMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "  WHERE  MAR.cTpoArendir='" & gArendirTipoCajaChica & "' AND cCtaContCod in ('" & psCtaArendir & "')  AND nProcNro=" & pnProcCH & " " _
    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "         AND M1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "         AND MCH.cAreaCod='" & psAreaCh & "' AND MCH.cAgeCod='" & psAgeCh & "'  " _
    & "         AND NOT EXISTS ( SELECT M2.CMOVNRO FROM MOV M2 JOIN MOVREF MR2 ON M2.nMOVNRO = MR2.nMOVNRO " _
    & "                          WHERE  MR2.nMOVNROREF =M1.nMOVNRO AND SUBSTRING(M2.COPECOD,1,5) IN ('" & Mid(gCHArendirCtaRendMN, 1, 5) & "','" & Mid(gCHArendirCtaRendME, 1, 5) & "')  " _
    & "                                 AND M2.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ")) "

'
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
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltro As String
Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function
Select Case pnRend
    Case Exacta
        lsFiltro = " and Rend.MontoRendido=0"
    Case ConIngreso
        lsFiltro = " and Rend.MontoRendido<0"
    Case ConEgreso
        lsFiltro = " and Rend.MontoRendido>0"
End Select

sql = " SELECT  MD.CDOCNRO, CONVERT(CHAR(12),MD.DDOCFECHA,103 ) AS Fecha, P.CPERSNOMBRE, " _
    & "         Rend.MontoRendido  as MontoRealRend, abs(Rend.MontoRendido)  as MontoRend, " _
    & "         Rend.CMOVDESC AS DESCRIPCION , ISNULL(A.CAREADESCRIPCION,'')  AS CAREADESC, " _
    & "         M.nMOVNRO AS nMOVNROSOL, M1.nMOVNRO AS nMOVNROATENC , MAR.CAREACOD, " _
    & "         ISNULL(AG.CAGEDESCRIPCION,'') AS CAGEDESC, MAR.CAGECOD , MAR.CPERSCOD, Rend.nMOVNROREND " _
    & " FROM    MOV  M " _
    & "         JOIN MOVCONT MCT    ON MCT.nMOVNRO = M.nMOVNRO " _
    & "         JOIN MOVARENDIR MAR     ON MAR.nMOVNRO = M.nMOVNRO " _
    & "         JOIN PERSONA P      ON P.CPERSCOD = MAR.CPERSCOD " _
    & "         LEFT JOIN AREAS A   ON A.CAREACOD = MAR.CAREACOD " _
    & "         LEFT JOIN AGENCIAS AG   ON AG.CAGECOD=MAR.CAGECOD " _
    & "         JOIN MOVCAJACHICA MCH   ON MCH.nMOVNRO = MAR.nMOVNRO " _
    & "         JOIN MOVDOC MD      ON MD.nMOVNRO = M.nMOVNRO " _
    & "         JOIN DOCUMENTO  D   ON D.nDocTpo = MD.nDocTpo " _
    & "         JOIN MOVREF MR      ON MR.nMOVNROREF= M.nMOVNRO " _
    & "         JOIN MOV M1         ON M1.nMOVNRO = MR.nMOVNRO " _
    & "         JOIN MOVCTA MC      ON MC.nMOVNRO = M1.nMOVNRO " _
    & "         LEFT JOIN MOVME ME  ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "         JOIN (  SELECT  M2.nMOVNRO  AS nMOVNROREND, MR2.nMOVNROREF as nMovNroAtenc, " _
    & "                         IsNull(SUM(ISNULL(ME1.NMOVMEIMPORTE, MC2.NMOVIMPORTE)),0) AS MontoRendido, M2.cMovDesc " _
    & "                 FROM    MOV M2 " _
    & "                         JOIN MOVREF MR2 ON M2.nMOVNRO = MR2.nMOVNRO " _
    & "                         LEFT JOIN MOVCTA MC2 ON MC2.nMOVNRO = M2.nMOVNRO  AND MC2.cCtaContCod IN ('" & psCtaArendir & "','" & psCtaPendiente & "') " _
    & "                         LEFT JOIN MOVME ME1 ON ME1.nMOVNRO = MC2.nMOVNRO AND ME1.nMOVITEM = MC2.nMOVITEM " _
    & "                 WHERE   SUBSTRING(M2.COPECOD,1,5) IN ('" & Mid(gCHArendirCtaRendMN, 1, 5) & "','" & Mid(gCHArendirCtaRendME, 1, 5) & "') "
sql = sql + "                    AND M2.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "                 GROUP BY M2.nMOVNRO , MR2.nMOVNROREF, M2.cMovDesc ) as REND " _
    & "         ON REND.nMovNroAtenc = M1.nMOVNRO " _
    & " WHERE   MAR.cTpoArendir='" & gArendirTipoCajaChica & "' AND MC.cCtaContCod in ('" & psCtaArendir & "') " _
    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "         AND M1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "         AND MCH.cAreaCod='" & psAreaCh & "' AND MCH.cAgeCod='" & psAgeCh & "'  AND nProcNro=" & pnProcCH & "" & lsFiltro
    
Set rs = oCon.CargaRecordSet(sql)
Set GetCHRendiciones = rs

oCon.CierraConexion
Set oCon = Nothing
End Function
Public Sub GrabaSolEgresoDirecto(ByVal psFormatoFecha As String, ByVal pbNewProv As Boolean, _
                                 ByVal psPersCod As String, ByVal psProvEstado As String, ByVal psMovUltActProv As String, _
                                 ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, ByVal pnDestino As Integer, _
                                 ByVal psCtaProvision As String, _
                                 ByVal psDocNro As String, ByVal pnDocTpo As TpoDoc, ByVal pdDocFecha As Date, _
                                 ByVal pnTotal As Currency, _
                                 ByVal psAgeCh As String, ByVal psAreaCh As String, ByVal pnNroProc As Integer, _
                                 ByVal rsCtasCont As ADODB.Recordset, ByVal rsObjetos As ADODB.Recordset, _
                                 ByVal rsImp As ADODB.Recordset)
Dim sql As String
Dim oMov As COMDMov.DCOMMov
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
Dim i As Integer

On Error GoTo GrabaSolEgresoDirectoErr
Set oMov = New COMDMov.DCOMMov
'oMov.inicio psFormatoFecha
oMov.BeginTrans
lbTrans = True

If pbNewProv Then
    oMov.InsertProv psPersCod, psProvEstado, psMovUltActProv
End If
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovCont lnMovNro, pnTotal, "0", "0"
oMov.InsertaMovGasto lnMovNro, psPersCod, IIf(pnDestino = -1, "", pnDestino)
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
                    lsSubCta = lsSubCta + rsObjetos![Sub Cuenta]
                End If
                rsObjetos.MoveNext
            Loop
        End If
        lnImporte = rsCtasCont!Monto
        For i = 6 To rsCtasCont.Fields.Count - 1
            If (pnDestino = 1 Or pnDestino = 2 Or rsCtasCont.Fields(i).Name = "AJUSTE") Then
                lnImporte = lnImporte + rsCtasCont.Fields(i).value
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
                        Case ObjPersona
                            lnOrdenObj = lnOrdenObj + 1
                            oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!Objpadre
                            oMov.InsertaMovObjPers lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código
                        Case ObjDescomEfectivo
                            lnOrdenObj = lnOrdenObj + 1
                            oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!Objpadre
                            oMov.InsertaMovObjEfectivo lnMovNro, lnItem, lnOrdenObj, rsObjetos!Código
                        Case ObjEntidadesFinancieras
                            lnOrdenObj = lnOrdenObj + 1
                            oMov.InsertaMovObj lnMovNro, lnItem, lnOrdenObj, rsObjetos!Objpadre
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
        For i = 6 To rsCtasCont.Fields.Count - 1
            rsImp.MoveFirst
            If rsCtasCont.Fields(i).value <> 0 Then
                rsImp.Find "Impuesto='" & rsCtasCont.Fields(i).Name & "'"
                If Not rsImp.EOF Then
                    oMov.InsertaMovOtrosItem lnMovNro, lnItem, rsImp!CTACONT, rsCtasCont.Fields(i).value, ""
                End If
            End If
        Next
        rsCtasCont.MoveNext
    Loop
End If
If Not rsImp Is Nothing Then
    rsImp.MoveFirst
    Do While Not rsImp.EOF
        If rsImp!Ok = "1" And (pnDestino = 0 Or pnDestino = 1) Then
            lnImporte = rsImp!Monto * IIf(rsImp!cDocImpDH = "D", 1, -1)
            If lnImporte <> 0 And rsImp!Impuesto <> "AJUSTE" Then
                lnItem = lnItem + 1
                oMov.InsertaMovCta lnMovNro, lnItem, rsImp!CTACONT, lnImporte
            End If
        End If
        rsImp.MoveNext
    Loop
End If
''Cuenta de Provisión / Caja Chica / A rendir Cuenta
lnItem = lnItem + 1
Dim oCont As COMNContabilidad.NCOMContFunciones
Set oCont = New COMNContabilidad.NCOMContFunciones

psCtaProvision = psCtaProvision + oCont.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaProvision, psAreaCh + psAgeCh, False)
Set oCont = Nothing
oMov.InsertaMovCta lnMovNro, lnItem, psCtaProvision, pnTotal * -1
oMov.InsertaMovObj lnMovNro, lnItem, 1, ObjCMACAgenciaArea
oMov.InsertaMovObjAgenciaArea lnMovNro, lnItem, 1, psAgeCh, psAreaCh
''Grabamos Documento en MovDoc
oMov.InsertaMovDoc lnMovNro, pnDocTpo, psDocNro, pdDocFecha
'Grabamos la Referencia al Movimiento que se regulariza
oMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, pnNroProc, gCHTipoProcEgresoDirecto

If Mid(psOpeCod, 3, 1) = COMDConstantes.gMonedaExtranjera Then
'    GeneraMovME psMovNro
End If
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
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function
 
sql = "Select   0 as OK , D.CDOCABREV , MD.CDOCNRO AS nDocTpo,  " _
    & "         CONVERT(VARCHAR(12),MD.dDocFecha,103) AS dDocFecha , P.CPERSNOMBRE, ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)) AS IMPORTE , " _
    & "         M.CMOVDESC , M.nMovNro, MD.nDocTpo " _
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
    & " Where   MCH.cProcTpo ='" & gCHTipoProcEgresoDirecto & "' AND M.nMovEstado = " & gMovEstContabNoContable & " " _
    & "         AND MC.CCTACONTCOD LIKE '" & psCtaFondFijo & "%' " _
    & "         AND NOT EXISTS (SELECT  M1.CMOVNRO " _
    & "                         FROM    MOV M1 JOIN MOVREF MR ON MR.nMOVNROREF = M1.nMOVNRO " _
    & "                         WHERE   MR.nMOVNRO=M.nMOVNRO AND M1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ")) " _
    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "         AND MCH.cAreaCod='" & psAreaCh & "' AND MCH.cAgeCod='" & psAgeCh & "'  AND MCH.nProcNro=" & pnProcCH & "" _
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
    Dim oMov As COMDMov.DCOMMov
    Dim sql As String
    Dim lnSaldoArendir As Currency
    Dim lbTrans As Boolean
    Dim lnMovNro As Long
    Set oMov = New COMDMov.DCOMMov
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
    Call oError.RaiseError(oError.MyUnhandledError, "NARendir:GrabaRechazoSolEgrDir Method")
End Function
Public Function GrabaAtenEgresoDirec(ByVal psFormatoFecha As String, _
                                     ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                     ByVal pnMovNroSol As Long, _
                                     ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                     ByVal pnNroProc As Integer, ByVal pnImporte As Currency) As Integer
Dim oMov As COMDMov.DCOMMov
Dim lnSaldo As Currency
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New COMDMov.DCOMMov

On Error GoTo GrabaAtenEgresoDirecErr
'oMov.inicio psFormatoFecha
oMov.BeginTrans
lbTrans = True
oMov.ActualizaMovimiento psMovNro, pnMovNroSol, gMovEstContabNoContable, , True, False, psOpeCod
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovRef lnMovNro, pnMovNroSol
lnSaldo = GetDatosCajaChica(psAreaCh, psAgeCh)
lnSaldo = lnSaldo - pnImporte
oMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProc, , , , , lnSaldo
oMov.CommitTrans
lbTrans = False
GrabaAtenEgresoDirec = 0
Set oMov = Nothing
Exit Function
GrabaAtenEgresoDirecErr:
    If lbTrans Then
        lbTrans = False
        oMov.RollbackTrans
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
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

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
    & "         JOIN PERSONA  P ON P.CPERSCOD = MG.CPERSCOD  " _
    & "         JOIN MOVREF MR ON MR.nMOVNRO = M.nMOVNRO  JOIN MOV M1 ON M1.NMOVNRO =MR.NMOVNROREF " _
    & " Where   MCH.cProcTpo ='" & gCHTipoProcEgresoDirecto & "' " _
    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
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
Dim oMov As COMDMov.DCOMMov
Dim lnSaldo As Currency
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New COMDMov.DCOMMov

On Error GoTo GrabaExtornoAtenEgresoDirecErr

GrabaExtornoAtenEgresoDirec = 1

'oMov.inicio psFormatoFecha
oMov.BeginTrans
lbTrans = True

oMov.InsertaMov psMovNroExt, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagDeExtorno
lnMovNro = oMov.GetnMovNro(psMovNroExt)
oMov.InsertaMovCont lnMovNro, pnImporte, 0, 0
'se actualiza la atencion con la solicitud de nuevo
'los movimientos de la atencion pasan a ser de la solicitud
oMov.ActualizaMovimiento psMovNroSol, pnMovNroAtenc, gMovEstContabNoContable, , False, True, , , False
oMov.InsertaMovRef lnMovNro, pnMovNroAtenc

lnSaldo = GetDatosCajaChica(psAreaCh, psAgeCh)
lnSaldo = lnSaldo + pnImporte
oMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProc, , , , , lnSaldo
oMov.CommitTrans
lbTrans = False
GrabaExtornoAtenEgresoDirec = 0
Set oMov = Nothing
Exit Function
GrabaExtornoAtenEgresoDirecErr:
    If lbTrans Then
        lbTrans = False
        oMov.RollbackTrans
    End If
    
    Err.Raise vbObjectError + 100, "GrabaExtornoAtenEgresoDirec", "Error Grabar Extorno de Atencion de Egreso Directo"
End Function
Public Function GetRSDatosCajaChica(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnProcCH As Integer, Optional ByRef oMov As COMDMov.DCOMMov) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function
 
 
sql = "SELECT   CH.cAreaCod, CH.cAgeCod, CH.nProcNro, nMontoAsig , nSaldo,  " _
    & "         CH.cPersCod , P.cPersNombre, CONVERT(VARCHAR(12),CONVERT(DATETIME, SUBSTRING(M.cMovNro,1,8)),103) AS dDesembolso, " _
    & "         nTopeEgresos,nMontoDesem,nSaldoAnt,nSaldo " _
    & " FROM    CAJACHICA CH " _
    & "         JOIN " & lsBaseCom & "PERSONA P ON P.CPERSCOD = CH.CPERSCOD " _
    & "         LEFT JOIN MOVCAJACHICA MCH ON MCH.cAreaCod = CH.cAreaCod AND MCH.cAgeCod= CH.cAgeCod " _
    & "                                 AND MCH.nProcNro = CH.nProcNro " _
    & "         JOIN MOV M ON M.NMOVNRO =MCH.NMOVNRO  " _
    & " WHERE   CH.cAreaCod = '" & psAreaCh & "' AND CH.cAgeCod = '" & psAgeCh & "' AND CH.nProcNro  = " & pnProcCH & ""

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
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function


sql = "SELECT   MD.CDOCNRO, P.CPERSNOMBRE, CONVERT(CHAR(12), MD.dDocFecha ,103  ) AS Fecha ,          " _
    & "         ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS IMPORTE, nMovSaldo, M1.CMOVDESC AS DESCRIPCION, M1.CMOVNRO AS CMOVNROATENC, " _
    & "         ISNULL(REND.CMOVNROREND,'') AS CMOVNROREND , ISNULL(Rend.MontoRendido,0), MCH.nProcNro " _
    & " FROM    MOV  M JOIN MOVCONT MCT ON MCT.nMOVNRO = M.nMOVNRO " _
    & "         JOIN MOVARENDIR MAR ON MAR.nMOVNRO = M.nMOVNRO " _
    & "         JOIN PERSONA P     ON P.CPERSCOD = MAR.CPERSCOD " _
    & "         JOIN MOVCAJACHICA MCH ON MCH.nMOVNRO = MAR.nMOVNRO " _
    & "         JOIN MOVDOC MD ON MD.nMOVNRO = M.nMOVNRO " _
    & "         JOIN DOCUMENTO  D ON D.nDocTpo = MD.nDocTpo " _
    & "         JOIN MOVREF MR ON MR.nMOVNROREF= M.nMOVNRO " _
    & "         JOIN MOV M1 ON M1.nMOVNRO = MR.nMOVNRO " _
    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M1.nMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "         LEFT JOIN ( SELECT  M2.cMOVNRO AS CMOVNROREND, MR2.nMOVNROREF as nMovNroAtenc, " _
    & "                             IsNull(Sum(IsNull(ME1.NMOVMEIMPORTE, IsNull(MC2.NMOVIMPORTE, 0))), 0) As MontoRendido " _
    & "                     FROM    MOV M2 JOIN MOVREF MR2 ON M2.nMOVNRO = MR2.nMOVNRO " _
    & "                             LEFT JOIN MOVCTA MC2 ON MC2.nMOVNRO = M2.nMOVNRO  AND MC2.cCtaContCod IN ('" & psCtaArendir & "','" & psCtaPendiente & "') " _
    & "                             LEFT JOIN MOVME ME1 ON ME1.nMOVNRO = MC2.nMOVNRO AND ME1.nMOVITEM = MC2.nMOVITEM " _
    & "                     WHERE   SUBSTRING(M2.COPECOD,1,5) IN ('" & Mid(gCHArendirCtaRendMN, 1, 5) & "','" & Mid(gCHArendirCtaRendME, 1, 5) & "') " _
    & "                             AND M2.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "') " _
    & "                     GROUP BY    M2.cMOVNRO , MR2.nMOVNROREF ) as REND " _
    & "         ON REND.nMovNroAtenc = M1.nMOVNRO " _
    & " WHERE   MAR.cTpoArendir='" & gArendirTipoCajaChica & "'  AND cCtaContCod IN ('" & psCtaArendir & "')  " _
    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "         AND M1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") "
sql = sql + "      AND MCH.cAreaCod='" & psAreaCh & "' AND MCH.cAgeCod='" & psAgeCh & "' AND MCH.nProcNro=" & pnProcCH & " " _
    & " Union " _
    & " SELECT  MD.CDOCNRO, P.CPERSNOMBRE, CONVERT(CHAR(12), MD.dDocFecha ,103  ) AS Fecha , " _
    & "         ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) AS IMPORTE, nMovSaldo, M1.CMOVDESC AS DESCRIPCION, M1.CMOVNRO AS CMOVNROATENC, " _
    & "         SUST.CMOVNROSOL , IsNull(SUST.MontoRendido, 0), MCH.nProcNro " _
    & " FROM    MOV  M   JOIN MOVCONT MCT ON MCT.nMOVNRO = M.nMOVNRO " _
    & "         JOIN MOVARENDIR MAR ON MAR.nMOVNRO = M.nMOVNRO " _
    & "         JOIN PERSONA P     ON P.CPERSCOD = MAR.CPERSCOD " _
    & "         JOIN MOVCAJACHICA MCH ON MCH.nMOVNRO = MAR.nMOVNRO " _
    & "         JOIN MOVDOC MD ON MD.nMOVNRO = M.nMOVNRO " _
    & "         JOIN DOCUMENTO  D ON D.nDocTpo = MD.nDocTpo " _
    & "         JOIN MOVREF MR ON MR.nMOVNROREF= M.nMOVNRO " _
    & "         JOIN MOV M1 ON M1.nMOVNRO = MR.nMOVNRO " _
    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M1.nMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & "         JOIN ( SELECT           '' AS CMOVNROSOL, MR2.nMOVNROREF as nMovNroAtenc, " _
    & "                                 IsNull(Sum(IsNull(ME1.NMOVMEIMPORTE, IsNull(MC2.NMOVIMPORTE, 0))), 0) As MontoRendido " _
    & "                     FROM        MOV M2 JOIN MOVREF MR2 ON M2.nMOVNRO = MR2.nMOVNRO " _
    & "                                 JOIN MOVCAJACHICA MCH2 ON MCH2.nMOVNRO = M2.nMOVNRO " _
    & "                                 JOIN MOVCTA MC2 ON MC2.nMOVNRO = M2.nMOVNRO  AND MC2.cCtaContCod IN ('" & psCtaArendir & "','" & psCtaPendiente & "') " _
    & "                                 LEFT JOIN MOVME ME1 ON ME1.nMOVNRO = MC2.nMOVNRO AND ME1.nMOVITEM = MC2.nMOVITEM " _
    & "                     WHERE       M2.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "                                 AND M2.COPECOD IN ('" & gCHArendirCtaSustMN & "','" & gCHArendirCtaSustMN & "') " _
    & "                                 AND MCH2.cAreaCod='" & psAreaCh & "' AND MCH2.cAgeCod='" & psAgeCh & "' AND MCH2.nProcNro=" & pnProcCH & "" _
    & "                                 AND NOT EXISTS (SELECT  M3.nMOVNRO "
sql = sql + "                                           FROM    MOV M3 JOIN MOVREF MR3 ON MR3.nMOVNRO=M3.nMOVNRO" _
    & "                                                 WHERE   MR2.nMOVNROREF = MR3.nMOVNROREF AND Substring(M3.COPECOD,1,5) IN ('" & Mid(gCHArendirCtaRendMN, 1, 5) & "','" & Mid(gCHArendirCtaRendIngME, 1, 5) & "')" _
    & "                                                         AND M3.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ")) " _
    & "                     GROUP BY    MR2.nMOVNROREF " _
    & "               Union " _
    & "                     SELECT      M2.CMOVNRO AS CMOVNROREND, MR2.nMOVNROREF as nMovNroAtenc, " _
    & "                                 IsNull(Sum(IsNull(ME1.NMOVMEIMPORTE, IsNull(MC2.NMOVIMPORTE, 0))), 0) As MontoRendido " _
    & "                     FROM        MOV M2 JOIN MOVREF MR2 ON M2.nMOVNRO = MR2.nMOVNRO " _
    & "                                 JOIN MOVCAJACHICA MCH2 ON MCH2.nMOVNRO = M2.nMOVNRO " _
    & "                                 JOIN MOVCTA MC2 ON MC2.nMOVNRO = M2.nMOVNRO  AND MC2.cCtaContCod IN ('" & psCtaArendir & "','" & psCtaPendiente & "') " _
    & "                                 LEFT JOIN MOVME ME1 ON ME1.nMOVNRO = MC2.nMOVNRO AND ME1.nMOVITEM = MC2.nMOVITEM " _
    & "                     WHERE       M2.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "                                 AND Substring(M2.COPECOD,1,5) IN ('" & Mid(gCHArendirCtaRendMN, 1, 5) & "','" & Mid(gCHArendirCtaRendIngME, 1, 5) & "')" _
    & "                                 AND MCH2.cAreaCod='" & psAreaCh & "' AND MCH2.cAgeCod='" & psAgeCh & "' AND MCH2.nProcNro=" & pnProcCH & "" _
    & "                     GROUP BY    M2.CMOVNRO, MR2.nMOVNROREF ) as SUST " _
    & "         ON SUST.nMovNroAtenc = M1.nMOVNRO " _
    & " WHERE   MAR.cTpoArendir='" & gArendirTipoCajaChica & "' AND cCtaContCod IN ('" & psCtaArendir & "')  " _
    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "         AND M1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "         AND MCH.cAreaCod='" & psAreaCh & "'  AND MCH.cAgeCod='" & psAgeCh & "' AND MCH.nProcNro<>" & pnProcCH & ""

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
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function

sql = "SELECT   D.CDOCABREV, MD.CDOCNRO, CONVERT(VARCHAR(12),MD.DDOCFECHA,103) AS FECHADOC," _
    & "         P.CPERSNOMBRE, ABS(SUM(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE))) AS Importe, " _
    & "         M.cMovNro , M1.CMOVNRO AS CMOVNROREF , M.cMovDesc AS Glosa" _
    & " FROM    MOV M " _
    & "         JOIN MOVREF MR ON MR.nMOVNRO = M.nMOVNRO  JOIN MOV M1 ON M1.NMOVNRO = MR.NMOVNROREF " _
    & "         JOIN MOVGASTO  MG ON MG.nMOVNRO = M.nMOVNRO " _
    & "         JOIN PERSONA P ON P.CPERSCOD = MG.CPERSCOD " _
    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.nMOVNRO=MC.nMOVNRO AND MC.nMOVITEM = ME.nMOVITEM " _
    & "         JOIN MOVDOC MD ON MD.nMOVNRO =M.nMOVNRO " _
    & "         JOIN DOCUMENTO D ON D.nDocTpo = MD.nDocTpo " _
    & " WHERE   M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "         AND M1.CMOVNRO='" & psMovNroAtenc & "' " _
    & "         AND MC.CCTACONTCOD IN('" & psCtaArendir & "','" & psCtaPendiente & "') " _
    & " GROUP BY D.CDOCABREV, MD.CDOCNRO, MD.DDOCFECHA , P.CPERSNOMBRE , M.CMOVNRO, M1.CMOVNRO , M.cMovDesc "

Set rs = oCon.CargaRecordSet(sql)
Set GetCHSustAtenArendir = rs

oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GetDatosRendArendirCH(ByVal psMovNroAtenc As String, ByVal psCtaFondoFijo As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function

sql = " Select  CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) AS FechaRend ,  " _
    & "         M.CMOVNRO, M.cMovDesc, ISNULL(ME.NMOVMEIMPORTE,MC.nMovImporte) AS nMovImporte " _
    & " FROM    MOV M JOIN MOVREF MR ON MR.nMOVNRO = M.nMOVNRO JOIN MOV M1 ON M1.NMOVNRO=MR.NMOVNROREF " _
    & "         JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM = MC.nMOVITEM " _
    & " WHERE   M1.CMOVNRO ='" & psMovNroAtenc & "' AND MC.CCTACONTCOD LIKE '" & psCtaFondoFijo & "%'"

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
Dim oCon As New COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
If oCon.AbreConexion = False Then Exit Function

sql = "SELECT  MC.CCTACONTCOD, " _
    & "         SUM(CASE WHEN MC.NMOVIMPORTE>0 THEN ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE) " _
    & "             ELSE 0 END ) AS DEBE, " _
    & "         SUM(CASE WHEN MC.NMOVIMPORTE<0 THEN ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)) " _
    & "         ELSE 0 END)  AS HABER " _
    & " FROM    MOV  M " _
    & "         JOIN MOVCAJACHICA MCH ON M.NMOVNRO = MCH.NMOVNRO " _
    & "         JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO AND ME.NMOVITEM = MC.NMOVITEM " _
    & "         JOIN CTACONT CC ON CC.CCTACONTCOD = MC.CCTACONTCOD " _
    & " WHERE   SUBSTRING(M.COPECOD,1,5) NOT IN ('" & Mid(gCHDesembOrdenPagoMN, 1, 5) & "','" & Mid(gCHDesembOrdenPagoME, 1, 5) & "') AND  M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "         AND MCH.CAREACOD='" & psAreaCh & "' AND MCH.CAGECOD='" & psAgeCh & "' AND nProcNro = " & pnProcCH & " " _
    & " GROUP BY MC.CCTACONTCOD "

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
Dim oCon As New COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
If oCon.AbreConexion = False Then Exit Function

sql = "SELECT   MC.CCTACONTCOD, ISNULL(D.CDOCABREV,'') AS CDOCABREV, ISNULL(MD.CDOCNRO,'') AS CDOCNRO, " _
    & "         ISNULL(Convert(Varchar(10),MD.dDocFecha,103),'') as dDocFecha, M.CMOVDESC, " _
    & "         CASE WHEN MC.NMOVIMPORTE>0 THEN ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE ) " _
    & "             ELSE 0 END  AS DEBE, " _
    & "         CASE WHEN MC.NMOVIMPORTE<0 THEN ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE )) " _
    & "             ELSE 0 END  AS HABER " _
    & " FROM    MOV  M " _
    & "         JOIN MOVCAJACHICA MCH ON M.NMOVNRO = MCH.NMOVNRO " _
    & "         JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO AND ME.NMOVITEM = MC.NMOVITEM " _
    & "         JOIN CTACONT CC ON CC.CCTACONTCOD = MC.CCTACONTCOD " _
    & "         LEFT JOIN MOVDOC MD ON MD.NMOVNRO = M.NMOVNRO " _
    & "         LEFT JOIN DOCUMENTO D ON D.nDocTpo = MD.nDocTpo " _
    & " WHERE   SUBSTRING(M.COPECOD,1,5) NOT IN ('" & Mid(gCHDesembOrdenPagoMN, 1, 5) & "','" & Mid(gCHDesembOrdenPagoME, 1, 5) & "') AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "         AND NOT MC.cCtaContCod LIKE '" & psCtaFondoFijo & "%' " _
    & "         AND MCH.CAREACOD='" & psAreaCh & "' AND MCH.CAGECOD='" & psAgeCh & "' AND nProcNro = " & pnProcCH & " " _
    & " ORDER BY MC.cCtaContCod, MD.nDocTpo, MD.cDocNro"

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
                            ByVal pdFecSis As Date, ByVal psCodUser As String, ByVal psAgencia As String) As Integer


 On Error GoTo GrabaRendicionChErr
 Dim rs As ADODB.Recordset
 Dim lnNuevoProcCH As Integer
 Dim lbTrans As Boolean
 Dim lnMovNro As Long
    Dim oMov As COMDMov.DCOMMov
    Dim sql As String
    Dim lnSaldoArendir As Currency
    Set oMov = New COMDMov.DCOMMov
    GrabaRendicionCh = 1
    oMov.BeginTrans
    lbTrans = True
    oMov.InsertaMov psMovNroRend, psOpeCod, psGlosa
    lnMovNro = oMov.GetnMovNro(psMovNroRend)
    oMov.InsertaMovCont lnMovNro, pnMontoRendido, "0", "0"
    oMov.InsertaMovCajaChica lnMovNro, psAreaCod, psAgeCod, pnNroProc, gCHTipoProcRendicion
    'ver***************************************
    ActualizaMovCajaChica oMov, lnMovNro, psAreaCod, psAgeCod, pnNroProc, pdFecSis, psCodUser, psAgencia
    Set rs = New ADODB.Recordset
    Set rs = GetRSDatosCajaChica(psAreaCod, psAgeCod, pnNroProc, oMov)
    If Not rs.EOF And Not rs.BOF Then
        lnNuevoProcCH = GetNuevaCajaChica(psAreaCod, psAgeCod)
        oMov.InsertaCajaChica psAreaCod, psAgeCod, lnNuevoProcCH, rs!cPersCod, rs!nMontoAsig, _
                    rs!nTopeEgresos, rs!nMontoDesem, pnSaldo, pnSaldo
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
    Call oError.RaiseError(oError.MyUnhandledError, "NARendir:GrabaRendicionCh Method")
End Function
Public Function GrabaAutorizacionCh(ByVal psMovNroAut As String, ByVal psMovNroRend As String, _
                                    ByVal psFormatoFecha As String, _
                                    ByVal psOpeCod As String, ByVal psGlosa As String, _
                                    ByVal pnMontoAutorizado As String, ByVal pnSaldo As Currency, _
                                    ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal pnNroProc As Integer) As Integer


 On Error GoTo GrabaAutorizacionChErr
 Dim rs As ADODB.Recordset
 Dim lnNuevoProcCH As Integer
 Dim lbTrans As Boolean
 Dim lnMovNro As Long
    Dim oMov As COMDMov.DCOMMov
    Dim sql As String
    
    Dim lnSaldoArendir As Currency
    Set oMov = New COMDMov.DCOMMov
    
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
    GrabaAutorizacionCh = 0
    Exit Function
GrabaAutorizacionChErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Call oError.RaiseError(oError.MyUnhandledError, "NARendir:GrabaAutorizacionCh Method")
End Function
Private Sub ActualizaMovCajaChica(ByRef oMov As COMDMov.DCOMMov, ByVal pnMovNroRend As Long, ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnNroProc As Integer, _
                                  ByVal pdFecSis As Date, ByVal psCodUser As String, ByVal psAgencia As String)
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsMovNroNew As String
Dim lnMovNronew As Long
Set rs = New ADODB.Recordset

sql = "SELECT   M.NMOVNRO " _
    & " FROM    MOV M     " _
    & "         JOIN MOVCAJACHICA MCH ON MCH.NMOVNRO=M.NMOVNRO JOIN MOVCTA MC ON MC.NMOVNRO=M.NMOVNRO   " _
    & "  WHERE  M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "') " _
    & "         AND MCH.CAREACOD='" & psAreaCh & "' AND MC.NMOVIMPORTE>0 " _
    & "         AND CAGECOD='" & psAgeCh & "' AND nProcNro =" & pnNroProc & " " _
    & "         AND (SUBSTRING(M.COPECOD,1,5)>= '" & Mid(gCHArendirCtaAtencMN, 1, 5) & "' OR " _
    & "              SUBSTRING(M.COPECOD,1,5)>= '" & Mid(gCHArendirCtaAtencME, 1, 5) & "')  GROUP BY M.NMOVNRO"

Set rs = oMov.CargaRecordSet(sql)
Do While Not rs.EOF
    lsMovNroNew = oMov.GeneraMovNro(pdFecSis, psAgencia, psCodUser)
    'Debug.Print lsMovNroNew, rs!cMovNro
    oMov.ActualizaMovimiento lsMovNroNew, rs!nMovNro, gMovEstContabMovContable, True, True, True
    lnMovNronew = oMov.GetnMovNro(lsMovNroNew)
    oMov.InsertaMovCajaChicaRend pnMovNroRend, rs!nMovNro, lnMovNronew
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
End Sub
Public Function GetNuevaCajaChica(ByVal psAreaCh As String, ByVal psAgeCh As String) As Integer
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

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
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function
 
 
sql = "SELECT   MCH.nProcNro " _
    & "FROM     MOV M JOIN MOVCAJACHICA MCH ON M.nMOVNRO=MCH.nMOVNRO " _
    & " WHERE   MCH.cProcTpo IN ('" & gCHTipoProcRendicion & "')         " _
    & "         And MCH.cAreaCod='" & psAreaCh & "' and MCH.cAgeCod= '" & psAgeCh & "' " _
    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & " group by  MCH.CAREACOD,MCH.CAGECOD, nProcNro "
    
Set rs = oCon.CargaRecordSet(sql)
Set GetCHRendidas = rs
oCon.CierraConexion
Set oCon = Nothing

End Function
Public Function GetCHRendSinAutorizacion(ByVal psAreaCh As String, ByVal psAgeCh As String) As ADODB.Recordset

Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function
 
 
sql = "SELECT   MCH.nProcNro " _
    & "FROM     MOV M JOIN MOVCAJACHICA MCH ON M.NMOVNRO=MCH.NMOVNRO " _
    & " WHERE   MCH.cProcTpo IN ('" & gCHTipoProcRendicion & "')         " _
    & "         And MCH.cAreaCod='" & psAreaCh & "' and MCH.cAgeCod= '" & psAgeCh & "' " _
    & "         AND M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "') " _
    & "         AND NOT EXISTS ( SELECT M1.CMOVNRO " _
    & "                          FROM   MOV M1 JOIN MOVREF MR ON MR.NMOVNRO = M1.NMOVNRO" _
    & "                          WHERE  MR.NMOVNROREF=M.NMOVNRO AND M1.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "')) " _
    & " Group BY MCH.CAREACOD,MCH.CAGECOD, nProcNro "
    
Set rs = oCon.CargaRecordSet(sql)
Set GetCHRendSinAutorizacion = rs
oCon.CierraConexion
Set oCon = Nothing

End Function
Public Function GetMovCHProceso(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnNroProc As Integer, ByVal pnTipoProc As CHTipoProc) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function

sql = "SELECT  MCH.CAREACOD,MCH.CAGECOD, nProcNro, M.cMovNro, MCH.NMovNro " _
    & " FROM   MOV M JOIN MOVCAJACHICA MCH ON M.NMOVNRO=MCH.NMOVNRO  " _
    & " WHERE  MCH.cProcTpo IN ('" & pnTipoProc & "') " _
    & "        And MCH.cAreaCod='" & psAreaCh & "' and MCH.cAgeCod= '" & psAgeCh & "' " _
    & "        And MCH.nProcNro=" & pnNroProc & " " _
    & "        AND M.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "') "
GetMovCHProceso = ""
Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetMovCHProceso = rs!cMovnro
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing
End Function



Public Function GetCHAutorizadas(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnNroProc As Integer) As ADODB.Recordset

Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

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
    & "                          WHERE   M1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
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
                                    ByVal pbEfectivo As Boolean) As ADODB.Recordset

Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltroDoc As String

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset
If pbEfectivo Then
    lsFiltroDoc = " AND MD.CDOCNRO Is Null"
Else
    lsFiltroDoc = " AND NOT MD.CDOCNRO Is Null"
End If

If oCon.AbreConexion = False Then Exit Function
 
 sql = "SELECT  0 as Ok, ISNULL(D.CDOCABREV,'EFEC') AS DocAbrev, ISNULL(MD.CDOCNRO,'EFECTIVO') as NroDoc , " _
    & "         CONVERT(CHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) AS FECHADESEMB, " _
    & "         M.CMOVDESC, SUM(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)) AS MONTOAUTORIZADO, " _
    & "         SUM(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)) AS MONTOAUTORIZADO, CH.nSaldo,   " _
    & "         M.NMOVNRO AS NMovNroDesemb " _
    & " FROM    MOVCAJACHICA MCH " _
    & "         JOIN CAJACHICA CH ON CH.cAreaCod= MCH.cAreaCod AND CH.cAgeCod = MCH.cAgeCod  AND CH.nProcNro= MCH.nProcNro " _
    & "         JOIN MOV M ON M.nMOVNRO = MCH.nMOVNRO " _
    & "         JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO " _
    & "         LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO AND MC.NMOVITEM = ME.NMOVITEM " _
    & "         LEFT JOIN MOVDOC MD ON MD.NMOVNRO=M.NMOVNRO AND MD.nDocTpo NOT IN (" & TpoDocVoucherEgreso & ")  " _
    & "         LEFT JOIN DOCUMENTO D ON D.nDocTpo = MD.nDocTpo " _
    & " WHERE   MCH.CAREACOD='" & psAreaCh & "' AND MCH.CAGECOD= '" & psAgeCh & "' AND MCH.nProcNro=" & pnNroProc & " " _
    & "         AND MCH.cProcTpo IN ('" & gCHTipoProcDesembolso & "')  AND MC.CCTACONTCOD LIKE '" & psCtaFondoFijo & "%' " & lsFiltroDoc _
    & "         AND M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "         AND NOT EXISTS (SELECT  M1.NMOVNRO " _
    & "                         FROM    MOV M1 JOIN MOVCAJACHICA MCH1 ON MCH1.NMOVNRO =M1.NMOVNRO  " _
    & "                         WHERE   MCH1.cAreaCod= MCH.cAreaCod AND MCH1.cAgeCod = MCH.cAgeCod  AND MCH1.nProcNro= MCH.nProcNro " _
    & "                                 AND M1.NMOVFLAG =" & gMovFlagVigente & " AND  MCH1.cProcTpo IN ('" & gCHTipoProcEgresoDirecto & "','" & gCHTipoProcArendir & "')) " _
    & " GROUP BY D.CDOCABREV, MD.CDOCNRO, M.NMOVNRO , M.CMOVNRO  , CH.nSaldo, CH.nMontoDesem, M.CMOVDESC "
 
Set rs = oCon.CargaRecordSet(sql)
Set GetCHDesembolsadas = rs
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Function GrabaDesembolsoCH(ByVal psMovNro As String, ByVal pnMovNroAut As Long, _
                                        ByVal psFormatoFecha As String, ByVal rsBilletaje As ADODB.Recordset, _
                                        ByVal psOpeCod As String, ByVal psMovDesc As String, ByVal pnImporte As Currency, _
                                        ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnNroProc As Integer, _
                                        ByVal psTpoDoc As String, ByVal psNroOrden As String, ByVal psFechaOrden As String, _
                                        ByVal psNroVoucher As String, ByVal rsAut As ADODB.Recordset, _
                                        ByVal psCtaDiferencia As String, ByVal pnMontoDiferencia As Currency, _
                                        ByVal psCtaDebe As String, ByVal psCtaHaber As String, ByVal pnSaldoCh As Currency) As Integer
            
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oMov As COMDMov.DCOMMov
Dim oContF As COMNContabilidad.NCOMContFunciones
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New COMDMov.DCOMMov
On Error GoTo GrabaDesembolsoCHErr

Set oContF = New COMNContabilidad.NCOMContFunciones

GrabaDesembolsoCH = 1
'oMov.inicio psFormatoFecha
oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)

oMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe
lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1:
lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaDebe, psAreaCh + psAgeCh, False)
oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe + lsSubCta, pnImporte
lnMovOrden = lnMovOrden + 1
oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
oMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, psAgeCh, psAreaCh
'guardamos el billetaje si este posee
If Not rsBilletaje Is Nothing Then
    If rsBilletaje.State = adStateOpen Then
        If Not rsBilletaje.EOF And Not rsBilletaje.BOF Then
            Do While Not rsBilletaje.EOF
                If rsBilletaje!Monto > 0 Then
                    lnMovItem = lnMovItem + 1: lnMovOrden = 0
                    oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, rsBilletaje!Monto * -1
                    lnMovOrden = lnMovOrden + 1
                    oMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjDescomEfectivo
                    oMov.InsertaMovObjEfectivo lnMovNro, lnMovItem, lnMovOrden, rsBilletaje!cEfectivoCod
                End If
                rsBilletaje.MoveNext
            Loop
        End If
        rsBilletaje.Close: Set rsBilletaje = Nothing
    Else
        lnMovItem = lnMovItem + 1
        oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnImporte * -1
    End If
Else
    lnMovItem = lnMovItem + 1
    oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, pnImporte * -1
End If
If Not rsAut Is Nothing Then
    Do While Not rsAut.EOF
        If rsAut!Ok = "1" Then
            oMov.InsertaMovRef lnMovNro, pnMovNroAut
            If rsAut!nProcNro <> pnNroProc Then
                oMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, rsAut!nProcNro, gCHTipoProcCancelada
            End If
        End If
        rsAut.MoveNext
    Loop
End If
oMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, pnNroProc, gCHTipoProcDesembolso
If psTpoDoc <> "" And psNroOrden <> "" Then
    oMov.InsertaMovDoc lnMovNro, psTpoDoc, psNroOrden, psFechaOrden
End If
If psNroVoucher <> "" Then
    oMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, psFechaOrden
End If
pnSaldoCh = pnSaldoCh + pnImporte
oMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProc, , , pnImporte, , pnSaldoCh
GrabaDesembolsoCH = 0
oMov.CommitTrans
lbTrans = False
Set oMov = Nothing
Set oContF = Nothing
Exit Function
GrabaDesembolsoCHErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaDesembolsoCH", Err.Description & " Error Graba Rendicion a rendir Efectivo "
End Function
Public Function GrabaExtornoDesembolso(ByVal pdFecSis As Date, ByVal pdFechaDesemb As Date, ByVal psFormatoFecha As String, _
                                     ByVal pnMovNroDesemb As Long, ByVal psMovNroExt As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                     ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                     ByVal pnNroProc As Integer, ByVal pnImporte As Currency) As Integer
Dim oMov As COMDMov.DCOMMov
Dim lnSaldo As Currency
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New COMDMov.DCOMMov

On Error GoTo GrabaExtornoDesembolsoErr
GrabaExtornoDesembolso = 1

'oMov.inicio psFormatoFecha
oMov.BeginTrans
lbTrans = True

If pdFecSis <> pdFechaDesemb Then
    oMov.ExtornaMovimiento psMovNroExt, pnMovNroDesemb
Else
    oMov.InsertaMov psMovNroExt, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagDeExtorno
    lnMovNro = oMov.GetnMovNro(psMovNroExt)
    oMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
    oMov.ActualizaMov pnMovNroDesemb, , , gMovFlagEliminado
    oMov.InsertaMovRef pnMovNroDesemb, lnMovNro
End If
lnSaldo = GetDatosCajaChica(psAreaCh, psAgeCh)
lnSaldo = lnSaldo - pnImporte
oMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProc, , , 0, , lnSaldo
GrabaExtornoDesembolso = 0
oMov.CommitTrans
lbTrans = False
Set oMov = Nothing
Exit Function
    If lbTrans Then
        lbTrans = False
        oMov.RollbackTrans
    End If
GrabaExtornoDesembolsoErr:
    Err.Raise vbObjectError + 100, "GrabaExtornoDesembolso", "Error Grabar Extorno de Desembolso de Caja chica"
End Function
Public Function GrabaHabNuevaCH(ByVal psFormatoFecha As String, _
                                ByVal psMovNroHab As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                ByVal pnNroProc As Integer, ByVal pnImporte As Currency, _
                                ByVal psPersCod As String, ByVal pnMontoAsig As Currency, ByVal pnMontoTope As Currency) As Integer
Dim oMov As COMDMov.DCOMMov
Dim lnSaldo As Currency
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oMov = New COMDMov.DCOMMov

On Error GoTo GrabaHabNuevaCHErr
GrabaHabNuevaCH = 1

'oMov.inicio psFormatoFecha
oMov.BeginTrans
oMov.InsertaMov psMovNroHab, psOpeCod, psMovDesc, gMovEstContabNoContable
lnMovNro = oMov.GetnMovNro(psMovNroHab)
oMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
oMov.InsertaCajaChica psAreaCh, psAgeCh, pnNroProc, psPersCod, pnMontoAsig, pnMontoTope, 0, 0, 0
oMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, pnNroProc, gCHTipoProcHabilitacion
oMov.CommitTrans
lbTrans = False
GrabaHabNuevaCH = 0
Set oMov = Nothing
Exit Function
GrabaHabNuevaCHErr:
    If lbTrans Then
        lbTrans = False
        oMov.RollbackTrans
    End If
    Err.Raise vbObjectError + 100, "GrabaHabNuevaCH", "Error Grabar nueva Caja chica"
End Function
Public Function ActualizaCH(ByVal psFormatoFecha As String, _
                           ByVal psAreaCh As String, ByVal psAgeCh As String, _
                           ByVal pnNroProc As Integer, _
                           ByVal psPersCod As String, ByVal pnMontoAsig As Currency, ByVal pnMontoTope As Currency) As Integer
Dim oMov As COMDMov.DCOMMov
Dim lnSaldo As Currency
Dim lbTrans As Boolean
Set oMov = New COMDMov.DCOMMov

On Error GoTo ActualizaCHErr
ActualizaCH = 1

'oMov.inicio psFormatoFecha
oMov.BeginTrans
lbTrans = True
oMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProc, psPersCod, pnMontoAsig, pnMontoTope
oMov.CommitTrans
lbTrans = False
ActualizaCH = 0

Set oMov = Nothing
Exit Function
ActualizaCHErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "ActualizaCH", "Error Actualizacion de Caja chica"
End Function



Public Function GetMontoNoDesembolsado(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnNroProc As Integer) As Currency

Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion = False Then Exit Function
 
sql = " SELECT  ISNULL(Sum(MC.NMOVMONTO),0) as Total " _
    & " FROM    CAJACHICA CH " _
    & "         JOIN MOVCAJACHICA MCH ON MCH.CAREACOD = CH.CAREACOD AND MCH.CAGECOD=CH.CAGECOD AND CH.nProcNro = MCH.nProcNro " _
    & "         JOIN MOV M ON M.NMOVNRO = MCH.NMOVNRO " _
    & "         JOIN MOVCONT MC ON MC.NMOVNRO = M.NMOVNRO " _
    & "         JOIN PERSONA P ON P.CPERSCOD = CH.CPERSCOD " _
    & " WHERE   CH.CAREACOD='" & psAreaCh & "' AND CH.CAGECOD= '" & psAgeCh & "' " _
    & "         AND MCH.cProcTpo IN ('" & gCHTipoProcHabilitacion & "') AND " _
    & "         M.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "         AND NOT EXISTS  (SELECT  M1.CMOVNRO " _
    & "                          FROM    MOV M1 JOIN MOVREF MR ON MR.NMOVNRO = M1.NMOVNRO " _
    & "                          WHERE   M1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & ") " _
    & "                                  AND MR.NMOVNROREF=M.NMOVNRO ) "

'AND CH.nProcNro=" & pnNroProc & "
Set rs = oCon.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetMontoNoDesembolsado = rs!Total
End If
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function GetCHArqueoRendidos(ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                    ByVal pnNroProc As Integer) As ADODB.Recordset

Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltroDoc As String

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset


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
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltroDoc As String

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset


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
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltroDoc As String

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset


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
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltroDoc As String

Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset


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

Public Function GrabaArqueoCajaChica(ByVal psFormatoFecha As String, _
                                    ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                    ByVal psCtaDebe As String, ByVal psCtaHaber As String, _
                                    ByVal rsBillIngreso As ADODB.Recordset, ByVal rsMonIng As ADODB.Recordset, _
                                    ByVal rsTotales As ADODB.Recordset, ByVal rsOP As ADODB.Recordset, ByVal rsCh As ADODB.Recordset, _
                                    ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                    ByVal psAgeCod As String, ByVal psCodPers As String, _
                                    ByVal psHoraIni As String, ByVal psHoraFin As String, ByVal pnTotalOP As Currency) As Integer
            
Dim oMov As COMDMov.DCOMMov
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


Set oMov = New COMDMov.DCOMMov
On Error GoTo GrabaArqueoCajaChicaErr

lnMovItem = 0: lnMovOrden = 0
GrabaArqueoCajaChica = 1
'oMov.inicio psFormatoFecha
oMov.BeginTrans
lbTrans = True
oMov.InsertaMov psMovNro, psOpeCod, psMovDesc
lnMovNro = oMov.GetnMovNro(psMovNro)

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
                    Case 5
                        lnTotalFondo = rsTotales!Importe
                    Case 6
                        lnFondoAsig = rsTotales!Importe
                    Case 7
                        lnDiferencia = rsTotales!Importe
                        pnTotal = lnDiferencia
                End Select
                rsTotales.MoveNext
            Loop
        End If
        rsTotales.Close: Set rsTotales = Nothing
        oMov.InsertaMovArqueo lnMovNro, psAgeCod, psCodPers, psHoraIni, psHoraFin, lnTotalEfec, _
                        pnTotalOP, lnTotalDoc, lnDocContab, lnARendir, lnDiferencia
    End If
End If
'guardamos el asiento de diferencia
If pnTotal <> 0 Then
    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, Abs(pnTotal)
    lnMovItem = lnMovItem + 1: lnMovOrden = 0
    oMov.InsertaMovCta lnMovNro, lnMovItem, psCtaHaber, Abs(pnTotal) * -1
End If
'guardamos el efectivo de arqueo
If Not rsBillIngreso Is Nothing Then
    If rsBillIngreso.State = adStateOpen Then
        If Not rsBillIngreso.EOF And Not rsBillIngreso.BOF Then
            Do While Not rsBillIngreso.EOF
                If rsBillIngreso!Monto > 0 Then
                    oMov.InsertaMovArqueoEfectivo lnMovNro, rsBillIngreso!cEfectivoCod, rsBillIngreso!Monto
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
                    oMov.InsertaMovArqueoEfectivo lnMovNro, rsMonIng!cEfectivoCod, rsMonIng!Monto
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
                oMov.InsertaMovArqueoOP lnMovNro, rsOP!Orden, rsOP!Importe
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
                oMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, rsCh(0), gCHTipoProcArqueo
                rsCh.MoveNext
            Loop
        End If
        rsCh.Close: Set rsCh = Nothing
    End If
End If

oMov.CommitTrans
lbTrans = False
GrabaArqueoCajaChica = 0
Set oMov = Nothing
Exit Function
GrabaArqueoCajaChicaErr:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise vbObjectError + 100, "GrabaArqueoCajaChica", Err.Description & " Error Devolucion de Efectivo de Cajero a Boveda"
End Function

Private Sub Class_Terminate()
Set oError = Nothing
End Sub
