VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCajaCtaIF"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A80C4B70138"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String


Public Function GeneraCodCtaIF(psCtaIfCod As String, psMoneda As String, psTpoCta As String) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oConect As COMConecta.DCOMConecta
    
    Set rs = New ADODB.Recordset
    Set oConect = New COMConecta.DCOMConecta
    GeneraCodCtaIF = ""
    If oConect.AbreConexion = False Then Exit Function
    
    sql = "Select   Max(cCtaIfCod) as MaxCtaIF " _
        & " From    CTAIF " _
        & " Where   cPersCod='" & psCtaIfCod & "' " _
        & "         and Substring(cCtaIfCod,1,2)='" & psTpoCta & "' and Substring(cCtaIfCod,3,1)='" & psMoneda & "' "

    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        If IsNull(rs!MaxCtaIF) Then
            GeneraCodCtaIF = psTpoCta + psMoneda + Format(1, String(7 - Len(psTpoCta + psMoneda), "0"))
        Else
            GeneraCodCtaIF = Format(IIf(IsNull(rs!MaxCtaIF), 0, rs!MaxCtaIF) + 1, String(Len(rs!MaxCtaIF), "0"))
        End If
    End If
    rs.Close
    Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
End Function

Public Function NombreIF(ByVal psPersCodIF As String) As String
Dim oCtasIF As DCajaCtasIF
Set oCtasIF = New DCajaCtasIF
NombreIF = oCtasIF.NombreIF(psPersCodIF)
Set oCtasIF = Nothing
End Function


Public Function CargaCtasIF(psMoneda As Moneda, Optional psFiltroCtaIF As String, Optional pnMuestraIF As MuestraIF = MuestraCuentas, Optional psCanje As String = "", Optional pbFiltraActivos As Boolean = True) As ADODB.Recordset
Dim oCtasIF As DCajaCtasIF
Set oCtasIF = New DCajaCtasIF
Set CargaCtasIF = oCtasIF.CargaCtasIF(psMoneda, psFiltroCtaIF, pnMuestraIF, psCanje, , pbFiltraActivos)
Set oCtasIF = Nothing
End Function

Public Function EmiteTipoCuentaIF(psCodCtaIf As String) As String
    On Error GoTo EmiteTipoCuentaIFErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oConect As COMConecta.DCOMConecta
    
    EmiteTipoCuentaIF = ""
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    
    Set rs = New ADODB.Recordset
    sql = "Select * from " & vsServerCom & "Constante where nCONSCOD ='" & gCGTipoCtaIF & "'  and nConsValor ='" & Mid(psCodCtaIf, 2, 1) & "'"
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        EmiteTipoCuentaIF = Trim(rs!cConsDescripcion)
    End If
    rs.Close: Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing

    Exit Function
EmiteTipoCuentaIFErr:
    MsgBox Err.Description, vbCritical
End Function

Public Function ValidaBanco(psCtaIf As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set oConec = New COMConecta.DCOMConecta
    ValidaBanco = ""
    If oConec.AbreConexion = False Then Exit Function
    sSQL = "select cPersCod+'.'+cCtaIfCod as CtaIfCod, cCtaIFDesc from CTAIF where cPersCod='" & psCtaIf & "' "
    Set rs = oConec.CargaRecordSet(sSQL)
    If Not (rs.EOF And rs.BOF) Then
        ValidaBanco = Trim(rs!cCtaIFDesc)
    End If
    RSClose rs
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GetCtaIFDesc(psPersCod As String, psMoneda As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set oConec = New COMConecta.DCOMConecta
    GetCtaIFDesc = ""
    If oConec.AbreConexion = False Then Exit Function
    sSQL = "SELECT cPersCod+'.'+cCtaIfCod as CtaIfCod, cCtaIFDesc " _
         & "FROM CtaIf WHERE cPersCod = '" & psPersCod & "' and cCtaIfCod LIKE '__" & psMoneda & "____' "
    Set rs = oConec.CargaRecordSet(sSQL)
    If Not (rs.EOF And rs.BOF) Then
        GetCtaIFDesc = Trim(rs!cCtaIFDesc)
    End If
    RSClose rs
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GetIFDesc(psPersCod As String, Optional lbAnexaTpoCuenta As Boolean = False, Optional psCtaIfCod As String = "") As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set oConec = New COMConecta.DCOMConecta
    GetIFDesc = ""
    If oConec.AbreConexion = False Then Exit Function
    If Not lbAnexaTpoCuenta Then
        sSQL = "SELECT cPersCod CtaIfCod, cPersNombre cCtaIFDesc " _
             & "FROM Persona WHERE cPersCod = '" & psPersCod & "' "
    Else
        sSQL = "SELECT p.cPersCod+'.'+cCtaIfCod as CtaIfCod, RTRIM(cPersNombre) + ' - ' + ci.cCtaIFDesc cCtaIFDesc " _
             & "FROM   Persona p JOIN CtaIf ci ON ci.cPersCod = p.cPersCod " _
             & "WHERE  p.cPersCod = '" & psPersCod & "' and cCtaIfCod LIKE '" & Left(psCtaIfCod, 5) & "' "
    End If
    Set rs = oConec.CargaRecordSet(sSQL)
    If Not (rs.EOF And rs.BOF) Then
        GetIFDesc = Trim(rs!cCtaIFDesc)
    End If
    RSClose rs
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Private Sub Class_Initialize()
    Dim oImp As COMDConstSistema.DCOMImpresoras
    Set oImp = New COMDConstSistema.DCOMImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
    'Dim oIni  As COMConecta.DCOMClasIni
    'Set oIni = New COMConecta.DCOMClasIni
    '
    'vsServerCom = oIni.BaseComunes
    'vsServerPers = oIni.BasePersonas
End Sub

Public Function GetInstFinancieras(Optional ByVal psFiltroTipoCtaIF As String = "") As Recordset
Dim oConec As COMConecta.DCOMConecta
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsIFFiltro As String
Dim lsFiltroCta As String
Dim Pos As String
Dim lsCadAux As String
Dim lsFiltroTipoIF As String
lsIFFiltro = ""
If psFiltroTipoCtaIF <> "" Then
    lsIFFiltro = " WHERE I.cIfTpo LIKE '" & psFiltroTipoCtaIF & "' "
End If

Set oConec = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = "SELECT  cPersCod  AS cCodigo, cPersNombre as Descripcion, Nivel " _
    & " From " _
    & "         (   SELECT  I.cIFTpo cPersCod, c.cConsDescripcion cPersNombre, 1 as Nivel " _
    & "             FROM    INSTITUCIONFINANC I JOIN Constante c ON c.nConsValor = convert(int,I.cIFTpo) and c.nconscod like '" & gCGTipoIF & "' " _
    & "         " & lsIFFiltro & "" _
    & "             GROUP BY I.cIFTpo, c.cConsDescripcion " _
    & "             UNION ALL " _
    & "             SELECT  I.cIFTpo + '.' + P.cPersCod cPersCod, CONVERT(CHAR(30),P.cPersNombre ) AS cPersNombre , 2 AS Nivel " _
    & "             FROM    INSTITUCIONFINANC I  " _
    & "                     JOIN PERSONA P ON P.cPersCod = I.cPersCod  " & lsIFFiltro _
    & "         ) AS INSTFIN  " _
    & "     ORDER BY cPersCod "
Set rs = oConec.CargaRecordSet(sql)
Set GetInstFinancieras = rs
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetCtasInstFinancieras(Optional ByVal psFiltroTipoCtaIF As String = "", Optional ByVal psCtacont As String = "", Optional pnNivel As MuestraIF = MuestraCuentas, Optional ByVal psFiltroTipoIF As String) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsIFFiltro As String
Dim lsFiltroCta As String
Dim Pos As String
Dim lsCadAux As String
Dim lsFiltroTipoIF As String
lsFiltroCta = ""
If psCtacont <> "" Then
    lsFiltroCta = " WHERE CIF.cCtaContCod LIKE '" & psCtacont & "'  "
End If
lsIFFiltro = ""
If psFiltroTipoCtaIF <> "" Then
    If psCtacont = "" Then
        lsIFFiltro = " WHERE "
    Else
        lsIFFiltro = "  And   "
    End If
    lsIFFiltro = lsIFFiltro + " CI.cCtaIfCod LIKE '" & psFiltroTipoCtaIF & "%' "
End If
lsFiltroTipoIF = ""
If psFiltroTipoIF <> "" Then
    If lsFiltroCta = "" And lsIFFiltro = "" Then
        lsFiltroTipoIF = " WHERE "
    Else
        lsFiltroTipoIF = " AND "
    End If
    lsFiltroTipoIF = lsFiltroTipoIF + " I.cIFTpo LIKE '" & psFiltroTipoIF & "'"
End If

Set oConec = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = "SELECT  cPersCod  AS cCodigo, cPersNombre as Descripcion, Nivel " _
    & " From " _
    & "         (   SELECT  P.cPersCod, CONVERT(CHAR(30),P.cPersNombre ) AS cPersNombre , 1 AS Nivel " _
    & "             FROM    INSTITUCIONFINANC I JOIN CTAIF CI ON CI.cPersCod = I.cPersCod " _
    & "                     JOIN CTAIFFILTRO CIF ON ( CIF.cPersCod =CI.cPersCod AND SUBSTRING(CI.cCtaIFCod,1,5) = SUBSTRING(CIF.cCtaIFCod,1,5)) " _
    & "                     JOIN PERSONA P ON P.cPersCod = I.cPersCod  " _
    & "             " & lsFiltroCta & lsIFFiltro & lsFiltroTipoIF _
    & "             GROUP  BY P.cPersCod, P.cPersNombre " _
    & "             Union " _
    & "             SELECT  CI.cIFTpo+'.'+CI.CPERSCOD + '.' + CI.cCtaIFCod, CONVERT(CHAR(40),CI.cCtaIFDesc) AS CTAIFDESC, " _
    & "                     LEN(CI.cCtaIFCod) As Nivel " _
    & "             FROM    INSTITUCIONFINANC I JOIN CTAIF CI ON CI.cPersCod = I.cPersCod " _
    & "                     JOIN CTAIFFILTRO CIF ON ( CIF.cPersCod =CI.cPersCod AND SUBSTRING(CI.cCtaIFCod,1,5) = SUBSTRING(CIF.cCtaIFCod,1,5)) " _
    & "                     JOIN PERSONA P ON P.cPersCod = I.cPersCod " _
    & "             " & lsFiltroCta & lsIFFiltro & lsFiltroTipoIF _
    & "         ) AS INSTFIN  WHERE INSTFIN.Nivel =" & pnNivel & " " _
    & "     ORDER BY cPersCod "

Set rs = oConec.CargaRecordSet(sql)
Set GetCtasInstFinancieras = rs
oConec.CierraConexion
Set oConec = Nothing
End Function
Public Function GetNewCuentaIF(ByVal psPersCod As String, ByVal pnTipoIF As CGTipoIF, _
                              ByVal pnTipoCtaIf As CGTipoCtaIF, ByVal pnMoneda As Moneda, ByVal psSubCtaIF As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set rs = New ADODB.Recordset
    Set oConec = New COMConecta.DCOMConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSQL = "Select  REPLACE(STR(MAX(Right(cCtaifCod,4)) +1,4,0),' ','0' ) as cCtaCodNew " _
        & " From    CtaIF CIF " _
        & " WHERE   CIF.cPersCod ='" & psPersCod & "' and CIF.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         and Substring(cCtaIfCod,2,1)='" & pnTipoCtaIf & "' and Substring(cCtaIfCod,3,1)='" & pnMoneda & "' " _
        & "         AND LEN(cCtaIFCod)>5 " _
        & " "

    GetNewCuentaIF = ""
    Set rs = oConec.CargaRecordSet(sSQL)
    If Not (rs.EOF And rs.BOF) Then
        GetNewCuentaIF = Mid(psSubCtaIF, 1, 3) & Trim(IIf(IsNull(rs!cCtaCodNew), Format(1, "0000"), rs!cCtaCodNew))
    Else
        GetNewCuentaIF = Mid(psSubCtaIF, 1, 3) & Format(1, "0000")
    End If
    rs.Close
    Set rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function GetVerificaSubCuentaIF(ByVal psPersCod As String, ByVal pnTipoIF As CGTipoIF, _
                                        ByVal pnTipoCtaIf As CGTipoCtaIF, ByVal pnMoneda As Moneda) As Boolean
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSQL = "Select  cCtaIfCod " _
        & " From    CtaIF CIF " _
        & " WHERE   CIF.cPersCod ='" & psPersCod & "' and CIF.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         and Substring(cCtaIfCod,2,1)='" & pnTipoCtaIf & "' and Substring(cCtaIfCod,3,1)='" & pnMoneda & "' " _
        & "         AND LEN(cCtaIFCod)<=5 " _
    
    GetVerificaSubCuentaIF = False
    Set rs = oConec.CargaRecordSet(sSQL)
    If Not (rs.EOF And rs.BOF) Then
        GetVerificaSubCuentaIF = True
    End If
    rs.Close
    Set rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function GetSubCuentaIF(ByVal psPersCod As String, ByVal pnTipoIF As CGTipoIF, _
                                        ByVal pnTipoCtaIf As CGTipoCtaIF, ByVal pnMoneda As Moneda) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSQL = "Select  cCtaIfCod " _
        & " From    CtaIF CIF " _
        & " WHERE   CIF.cPersCod ='" & psPersCod & "' and CIF.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         and Substring(cCtaIfCod,2,1)='" & pnTipoCtaIf & "' and Substring(cCtaIfCod,3,1)='" & pnMoneda & "' " _
        & "         AND LEN(cCtaIFCod)<5 " _
    
    GetSubCuentaIF = ""
    Set rs = oConec.CargaRecordSet(sSQL)
    If Not (rs.EOF And rs.BOF) Then
        GetSubCuentaIF = Trim(rs!cCtaIFCod)
    Else
        GetSubCuentaIF = Format(pnTipoCtaIf, "00") & pnMoneda
    End If
    rs.Close
    Set rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function GeneraSubCuentaIF(ByVal psPersCod As String, ByVal pnTipoIF As CGTipoIF, _
                                  ByVal psCtacont As String, ByVal psCtaIfCod As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSQL = "Select  cCtaIfCod, C.cCtaContCod, cCtaIFSubCta,C.cCtaContCod + C.cCtaIFSubCta" _
        & " From    CtaIFFiltro  C   " _
        & "         JOIN CTACONT CC ON CC.CCTACONTCOD  = C.cCtaContCod+ C.cCtaIFSubCta " _
        & " Where   C.cCtaContCod Like '" & psCtacont & "' " _
        & "         And cCtaIfCod Like '" & Mid(psCtaIfCod, 1, 5) & "%' And cPersCod ='" & psPersCod & "' AND cIFTpo ='" & Format(pnTipoIF, "00") & "' "
    
    GeneraSubCuentaIF = ""
    Set rs = oConec.CargaRecordSet(sSQL)
    If Not (rs.EOF And rs.BOF) Then
        GeneraSubCuentaIF = Trim(rs!cCtaIFSubCta)
    End If
    rs.Close
    Set rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GetSaldoCtaIfCalendario(ByVal psPersCod As String, ByVal pnTipoIF As CGTipoIF, _
                             ByVal psCtaIfCod As String, ByVal pdFecha As Date) As Currency
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSQL = "SELECT ISNULL( round(nMontoPrestado/ISNULL(CASE WHEN SubString(cia.cCtaIFCod,3,1) = 1 and cMonedaPago = 2 THEN iv.nIndiceVac ELSE 1 END,1),2) - CapPagado,nMontoPrestado)  nSaldo " _
         & "FROM ctaifadeudados cia " _
         & "  LEFT JOIN (select ci.cPersCod, ci.cIFTpo, ci.cCtaIFCod, dCtaIFAper, sum(cic.nCapital) CapPagado " _
         & "            from ctaif ci join ctaifcalendario cic on cic.cperscod = ci.cperscod and cic.ciftpo = ci.ciftpo and cic.cctaifcod = ci.cctaifcod " _
         & "       where cEstado = '" & gTpoEstCuotaAdeudCanc & "' and LEFT(cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' " _
         & "       group by ci.cPersCod, ci.cIFTpo, ci.cCtaIFCod, dCtaIFAper " _
         & "           ) ci on cia.cperscod = ci.cperscod and cia.ciftpo = ci.ciftpo and cia.cctaifcod = ci.cctaifcod " _
         & " LEFT JOIN indicevac iv on iv.dIndiceVac = ci.dCtaIFAper " _
         & "where cia.cPersCod='" & psPersCod & "' AND cia.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
         & "     AND cia.cCtaIFCod='" & psCtaIfCod & "' "

    GetSaldoCtaIfCalendario = 0
    Set rs = oConec.CargaRecordSet(sSQL)
    If Not (rs.EOF And rs.BOF) Then
        GetSaldoCtaIfCalendario = Trim(rs!nSaldo)
    End If
    rs.Close
    Set rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GetSaldoCtaIf(ByVal psPersCod As String, ByVal pnTipoIF As CGTipoIF, _
                             ByVal psCtaIfCod As String, ByVal pdFecha As Date) As Currency
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSQL = "Select  MS.cPersCod, MS.cIFTpo, MS.cCtaIFCod, nSaldo" _
        & " From    CtaIFSaldo MS " _
        & " Where   MS.dCtaIFSaldo = (  Select  Max(MS1.dCtaIFSaldo) " _
        & "                             From    CtaIFSaldo MS1 " _
        & "                             Where   MS.cCtaContCod = MS1.cCtaContCod and MS.cPersCod = MS1.cPersCod AND " _
        & "                                     MS.cIFTpo = MS1.cIFTpo And MS.cCtaIFCod = MS1.cCtaIFCod " _
        & "                                     AND MS1.dCtaIFSaldo <='" & Format(pdFecha, gsFormatoFecha) & "') " _
        & "         AND MS.cPersCod='" & psPersCod & "' AND MS.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         AND MS.cCtaIFCod='" & psCtaIfCod & "' "
        
    GetSaldoCtaIf = 0
    Set rs = oConec.CargaRecordSet(sSQL)
    If Not (rs.EOF And rs.BOF) Then
        GetSaldoCtaIf = Trim(rs!nSaldo)
    End If
    rs.Close
    Set rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function GetCuentasIFCanc(ByVal psOpeCod As String, ByVal pdFecha As Date, Optional ByVal pnOpeObjOrden As Integer = 1) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = "SELECT   CI.cCtaIFDesc, cPersNombre AS cInstFinanc , " _
    & "         ISNULL(CASE WHEN SubString(CI.cCtaIFCod,3,1) = 1 THEN Saldo.nSaldo ELSE Saldo.nSaldoME END,0) AS nCapital, 0 as ValorInt, CI.nInteres, " _
    & "         CI.dCtaIFAper, 0 Dias , 0 as IntCalculado, CI.cPersCod, CI.cIFTpo, CI.cCtaIFCod, CI.nCtaIFPlazo " _
    & " FROM    CTAIF CI " _
    & "         JOIN INSTITUCIONFINANC I ON I.cPersCod = CI.cPersCod AND I.cIFTpo = CI.cIFTpo " _
    & "         JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD " _
    & "         LEFT JOIN (     SELECT  CS.dCtaIFSaldo, CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod, CS.nSaldo, CS.nSaldoME " _
    & "                         FROM    CTAIFSALDO CS " _
    & "                         WHERE   CS.dCtaIFSaldo = (      SELECT  MAX(dCtaIFSaldo) " _
    & "                                                         FROM    CTAIFSALDO CS1 " _
    & "                                                         Where CS1.cPersCod = CS.cPersCod And CS1.cIFTpo = CS.cIFTpo And CS1.cCtaIFCod = CS.cCtaIFCod " _
    & "                                                         AND CS1.dCtaIFSaldo<='" & Format(pdFecha, gsFormatoFecha) & "')) AS SALDO " _
    & "         ON SALDO.cPersCod=CI.cPersCod and SALDO.cIFTpo = CI.cIFTpo and SALDO.cCtaIFCod= CI.cCtaIFCod, OPEOBJ O " _
    & " WHERE   LEN(CI.cCtaIFCod)=(Select Max(Len(cCtaIFCod)) From CtaIf) AND cCtaIFEstado ='1' " _
    & "         AND o.cOpeCod = '" & psOpeCod & "' and cOpeObjOrden='" & gEstadoCtaIFActiva & "' " _
    & "         And CI.cIFTpo+CI.cCtaIFCod Like cOpeObjFiltro And Saldo.nSaldo>0 "

Set rs = oConec.CargaRecordSet(sql)
Set GetCuentasIFCanc = rs
oConec.CierraConexion
Set oConec = Nothing
End Function
Public Function GetInteresCtaIf(ByVal psPersCod As String, ByVal pnTipoIF As CGTipoIF, _
                             ByVal psCtaIfCod As String) As Currency
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSQL = "Select  MS.cPersCod, MS.cIFTpo, MS.cCtaIFCod, nInteres " _
        & " From    CtaIF MS " _
        & " WHERE   MS.cPersCod='" & psPersCod & "' AND MS.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         AND MS.cCtaIFCod='" & psCtaIfCod & "'"
        
    GetInteresCtaIf = 0
    Set rs = oConec.CargaRecordSet(sSQL)
    If Not (rs.EOF And rs.BOF) Then
        GetInteresCtaIf = Trim(rs!nInteres)
    End If
    rs.Close
    Set rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function GetRsInteresCtasIF(ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, ByVal psCtaIfCod As String) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = "Select convert(char(10),dCtaIFIntRegistro,103) as FechaInt, nCtaIFIntValor , nCtaIFIntPeriodo, cCtaIFCod ,cPersCod, cIFTpo,  1 as  Registrado " _
    & " from CtaifInteres " _
    & " Where cPersCod='" & psPersCod & "' and cIFTpo ='" & Format(pnIFTpo, "00") & "' and cCtaIFCod='" & psCtaIfCod & "'"

Set rs = oConec.CargaRecordSet(sql)
Set GetRsInteresCtasIF = rs
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetDatosCtaIf(ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, ByVal psCtaIfCod As String) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = " SELECT  C.dCtaIFAper, C.dCtaIfCap, C.dCtaIfInt, C.nCtaIFPlazo, C.nInteres, C.cCtaIFEstado,C.nMontoEuros MontoEurosCtaIF,CC.cConsDescripcion AS cEstadoCons, cTpoC.cConsDescripcion cTpoCtaDesc, Isnull(SALDO.nSaldo,0) as Saldo, dCtaIFVenc, cia.* " _
    & " FROM    CTAIF C  LEFT JOIN CtaIFAdeudados cia ON cia.cPersCod = c.cPersCod and cia.cIFTpo = c.cIFTpo and cia.cCtaIFCod = c.cCtaIFCod " _
    & "         JOIN CONSTANTE CC ON    CC.NCONSVALOR =  C.cCtaIFEstado LEFT JOIN Constante cTpoC ON cTpoC.nConsValor = cia.nTpoCuota AND cTpoC.nCONSCOD = " & gCGAdeudCalTpoCuota & " " _
    & "         LEFT JOIN ( SELECT  CS.dCtaIFSaldo, CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod, CASE WHEN SubString(cs.cCtaIFCod,3,1) = '1' THEN CS.nSaldo ELSE cs.nSaldoME END nSaldo" _
    & "                     FROM    CTAIFSALDO CS " _
    & "                     WHERE   CS.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
    & "                                                 FROM    CTAIFSALDO CS1 " _
    & "                                                 Where   CS1.cPersCod = CS.cPersCod And CS1.cIFTpo = CS.cIFTpo And CS1.cCtaIFCod = CS.cCtaIFCod )) AS SALDO " _
    & "                     ON SALDO.cPersCod=C.cPersCod and SALDO.cIFTpo = C.cIFTpo and SALDO.cCtaIFCod= C.cCtaIFCod " _
    & " WHERE CC.nCONSCOD = " & gCGEstadoCtaIF & " AND C.cPersCod='" & psPersCod & "' and C.cIFTpo ='" & Format(pnIFTpo, "00") & "' and C.cCtaIFCod = '" & psCtaIfCod & "' ORDER BY nSaldo DESC "

Set rs = oConec.CargaRecordSet(sql)
Set GetDatosCtaIf = rs
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetCtaIfInteres(ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, ByVal psCtaIfCod As String) As Currency
Dim oConec As COMConecta.DCOMConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = " SELECT  nCtaIFIntValor " _
    & " FROM    CTAIFINTERES  C  " _
    & " WHERE   dCtaIFIntRegistro = (SELECT Max(dCtaIFIntRegistro) FROM CtaIFInteres cia WHERE cia.cPersCod = c.cPersCod and cia.cIFTpo = c.cIFTpo and cia.cCtaIFCod = c.cCtaIFCod )" _
    & "  and C.cPersCod='" & psPersCod & "' and C.cIFTpo ='" & Format(pnIFTpo, "00") & "' and C.cCtaIFCod = '" & psCtaIfCod & "'"

Set rs = oConec.CargaRecordSet(sql)
If Not rs.EOF Then
    GetCtaIfInteres = rs!nCtaIFIntValor
Else
    GetCtaIfInteres = 0
End If
RSClose rs
oConec.CierraConexion
Set oConec = Nothing
End Function


Public Function GetFlujosCtaIF(ByVal psOpeCod As String, _
                            ByVal psCtaIfCodDesde As String, ByVal psCtaIfCodHasta As String, _
                            ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset

Dim oConec As COMConecta.DCOMConecta
Dim sql As String
Dim rs As ADODB.Recordset


Set oConec = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

 

sql = " SELECT CI.cPersCod, CI.cIFTpo, CI.cCtaIFCod, P.CPERSNOMBRE, CI1.cCtaIFDesc, CI.cCtaIFDesc," _
    & "        RTRIM(P.CPERSNOMBRE) + ' ' + CI1.cCtaIFDesc + ' ' + CI.cCtaIFDesc AS cNomCtaIF, " _
    & "        ISNULL(SALDO.nSaldo,0) AS nSaldoIni, ISNULL(nSaldoME,0) nSaldoIniME, " _
    & "        CI.nCtaIFPlazo, CI.nInteres , " _
    & "        M.CMOVNRO, M.CMOVDESC, IsNull(D.cDocAbrev,'')  + ' ' + IsNull(MD.cDocNro,'') as Documento , " _
    & "        MO.nMovImporte,  ISNULL(ME.nMovMEImporte,0) nMovMEImporte, " _
    & "        Case WHEN MO.nMovImporte > 0 THEN ABS(MO.nMovImporte) ELSE 0 END AS DEBE , Case WHEN ISNULL(ME.nMovMEImporte,0) > 0 THEN ABS(ME.nMovMEImporte) ELSE 0 END AS DEBEME, " _
    & "        Case WHEN MO.nMovImporte < 0 THEN ABS(MO.nMovImporte) ELSE 0 END AS HABER, Case WHEN ISNULL(ME.nMovMEImporte,0) < 0 THEN ABS(ME.nMovMEImporte) ELSE 0 END AS HABERME " _
    & " FROM   MOV M " _
    & "        JOIN MovCta mo ON mo.nMovNro = m.nMovNro LEFT JOIN MovME me ON me.nMovNro = mo.nMovNro and me.nMovItem = mo.nMovItem " _
    & "        JOIN MOVOBJIF MC ON MC.nMovNro = mo.nMovNro and mc.nMovItem = mo.nMovItem " _
    & "        LEFT JOIN MOVDOC MD ON MD.NMOVNRO = M.NMOVNRO AND MD.nDocTpo <>" & TpoDocVoucherEgreso & "  " _
    & "        JOIN CTAIF CI ON MC.cPersCod=CI.cPersCod and MC.cIFTpo = MC.cIFTpo and MC.cCtaIFCod= CI.cCtaIFCod " _
    & "        JOIN CTAIF CI1 ON CI1.cPersCod = CI.cPersCod AND CI1.cIFTpo = CI.cIFTpo  AND SUBSTRING(CI.cCtaIFCod,1,LEN(CI1.cCtaIFCod)) = CI1.cCtaIFCod and Len(CI1.cCtaIFCod)=3 " _
    & "        JOIN INSTITUCIONFINANC I ON I.cPersCod = CI.cPersCod AND I.cIFTpo = CI.cIFTpo " _
    & "        JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD " _
    & "         LEFT JOIN ( SELECT  CS.cCtaContCod, CS.dCtaIFSaldo, CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod, CS.nSaldo, CS.nSaldoME " _
    & "                     FROM    CTAIFSALDO CS " _
    & "                     WHERE   CS.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
    & "                                                 FROM    CTAIFSALDO CS1 " _
    & "                                                 Where   cs1.cCtaContCod = cs.cCtaContCod and CS1.cPersCod = CS.cPersCod And CS1.cIFTpo = CS.cIFTpo " _
    & "                                                         And CS1.cCtaIFCod = CS.cCtaIFCod " _
    & "                                                         AND CS1.dCtaIFSaldo<='" & Format(DateAdd("d", -1, pdDesde), "mm/dd/yyyy") & "' )) AS SALDO" _
    & "         ON SALDO.cPersCod=CI.cPersCod and SALDO.cIFTpo = CI.cIFTpo and SALDO.cCtaIFCod= CI.cCtaIFCod and Saldo.cCtaContCod = mo.cCtaContCod " _
    & "         LEFT JOIN DOCUMENTO D ON D.nDocTpo = MD.nDocTpo, OPEOBJ O "
sql = sql + "WHERE   LEN(CI.cCtaIFCod)= (Select max(len(cCtaIFCod)) From CtaIf) AND CI.cCtaIFEstado ='1' " _
    & "         AND o.cOpeCod = '" & psOpeCod & "' and cOpeObjOrden='0' " _
    & "         AND CI.cIFTpo + CI.cCtaIFCod like cOpeObjFiltro " _
    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "         AND CI.cPersCod + '.' + CI.cCtaIFCod BETWEEN '" & psCtaIfCodDesde & "' AND '" & psCtaIfCodHasta & "' " _
    & "         and ( M.NMOVFLAG IN (" & gMovFlagVigente & "," & gMovFlagDeExtorno & ") or (m.NMovFlag = " & gMovFlagExtornado & " and EXISTS (SELECT mr1.NMovNro FROM MovRef mr1 JOIN Mov m1 ON m1.nMovNro = mr1.nMovNro WHERE mr1.NMovNroRef = m.NMovNro and LEFT(m1.cMovNro,6) > LEFT(m.cMovNro,6) ))) " _
    & "         and mo.cCtaContCod LIKE '1[13]_[23]%'  " _
    & " Order by CI.cIFTpo+ CI.cPersCod + CI.cCtaIFCod , M.CMOVNRO"


Set rs = oConec.CargaRecordSet(sql)
Set GetFlujosCtaIF = rs
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetRepSaldosCtaIf(ByVal psOpeCod As String, ByVal pdFecha As Date, Optional ByVal psOpeObjOrden As String = "'0'", Optional psCtaIFEstado As String = "", Optional psMoneda As String = "") As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = " SELECT  Z.cCtaIfCod , Z.cCtaIFDesc, Z.cBancoDesc, Z.cCtaIFEstado, " _
    & "         CASE    WHEN SUM(z.nImporte) > 0 THEN str(SUM(z.nImporte), 14, 2) " _
    & "                 WHEN SUM(z.nImporte) < 0 THEN '0' END as Debe, " _
    & "         CASE    WHEN SUM(z.nImporte) < 0 THEN str(SUM(z.nImporte) * -1, 14, 2) " _
    & "                 WHEN SUM(z.nImporte) > 0 THEN '0' END as Haber " _
    & " FROM (  " _
    & "         SELECT  CtaIf.cCtaIfCod, CtaIF.cCtaIFDesc, CtaIF.cBancoDesc, CtaIf.cCtaIFEstado, " & IIf(psMoneda = "1", " SUM(CI.NSALDO) ", " SUM(CI.NSALDOME) ") & " AS nImporte " _
    & "         FROM    CTAIFSALDO CI " _
    & "                 JOIN " _
    & "                 (   SELECT  CI.cIFTpo + CI.cPersCod + CI.cCtaIFCod AS cCtaIfCod, CI.cCtaIFDesc, P.cPersNombre cBancoDesc, CI.cCtaIFEstado " _
    & "                     FROM    CTAIF CI JOIN INSTITUCIONFINANC I ON  I.CPERSCOD = CI.CPERSCOD  AND CI.CIFTPO = I.CIFTPO " _
    & "                             JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD, OPEOBJ O " _
    & "                     WHERE   O.cOpeCod = '" & psOpeCod & "' and O.cOpeObjOrden IN (" & psOpeObjOrden & ") and CI.cCtaIFEstado IN ('" & psCtaIFEstado & "') " _
    & "                             AND CI.cIFTpo + CI.cCtaIFCod like O.cOpeObjFiltro " _
    & "                 ) as CtaIf ON  SUBSTRING(CI.cIFTpo + CI.cPersCod + CI.cCtaIFCod , 1, LEN(CtaIf.cCtaIfCod)) = CtaIf.cCtaIfCod , OPEOBJ O JOIN OpeCta oc ON oc.cOpeCod = o.cOpeCod " _
    & "         WHERE   O.cOpeCod = '" & psOpeCod & "' and ci.cCtaContCod LIKE LEFT(oc.cCtaContCod,2)+'" & psMoneda & "'+SubString(oc.cCtaContCod,4,22) + '%' and oc.cOpeCtaOrden IN (" & psOpeObjOrden & ") and O.cOpeObjOrden IN (" & psOpeObjOrden & ") " _
    & "                 AND CI.cIFTpo + CI.cCtaIFCod Like O.cOpeObjFiltro " & IIf(psMoneda = "", "", " and CI.cCtaIFCod LIKE '__" & psMoneda & "%' ")
sql = sql + "           AND CI.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
    & "                                         FROM    CTAIFSALDO CS1 " _
    & "                                         Where   cs1.cCtaContCod = ci.cCtaContCod and CS1.cPersCod = CI.cPersCod And CS1.cIFTpo = CI.cIFTpo And CS1.cCtaIFCod = CI.cCtaIFCod " _
    & "                                                 AND CS1.dCtaIFSaldo<='" & Format(pdFecha, gsFormatoFecha) & "') " _
    & "         GROUP   BY CtaIf.cCtaIfCod , CtaIF.cCtaIFDesc, CtaIF.cBancoDesc, CtaIF.cCtaIFEstado ) AS Z " _
    & " WHERE nImporte <> 0 " _
    & " GROUP  BY Z.cCtaIfCod , Z.cCtaIFDesc, Z.cBancoDesc, Z.cCtaIFEstado " _
    & " ORDER BY Z.cBancoDesc, Z.cCtaIfCod"
    
    
    '& " ORDER BY  Z.cCtaIFEstado, Z.cBancoDesc, Z.cCtaIfCod"

Set rs = oConec.CargaRecordSet(sql)
Set GetRepSaldosCtaIf = rs
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetCtaIFCont(ByVal pdFecha As Date, ByVal pcCtaContCod As String) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = " Select cCtaContCod + cCtaIFSubCta cCtaContCod,(case when cCtaContCod LIKE '11[12]7%'  then " _
      & " dbo.GetSaldoCtaIF('" & Format(pdFecha, gsFormatoFecha) & "',cCtaContCod+cCtaIFSubCta,ci.cPersCod,ci.cIfTpo,ci.cCtaIfCod,substring(ci.cCtaIfCod,3,1)) " _
      & " else 0 end) Restringido, " _
      & " (case when cctacontcod LIKE '11[12]3%'  then " _
      & " dbo.GetSaldoCtaIF('" & Format(pdFecha, gsFormatoFecha) & "',cCtaContCod+cCtaIFSubCta,ci.cPersCod,ci.cIfTpo,ci.cCtaIfCod,substring(ci.cCtaIfCod,3,1))" _
      & " else 0 end) capital, " _
      & " dbo.GetSaldoCtaIF('" & Format(pdFecha, gsFormatoFecha) & "',cCtaContCod+cCtaIFSubCta,ci.cPersCod,ci.cIfTpo,ci.cCtaIfCod,substring(ci.cCtaIfCod,3,1))Total" _
      & " From CtaIf ci Inner Join CtaIfFiltro cif on cif.cPersCod=ci.cPersCod and cif.cIfTpo=ci.cIfTpo" _
      & " and  cif.cCtaIfCod=ci.cCtaIfCod and cctacontcod like '11[12][37]%'" _
      & " Inner Join Persona p on ci.cPersCod=p.cPersCod " _
      & " where ci.ciftpo='01' and ci.cCtaIfCod like '03%' and cctaifestado=1 " _
      & " and    dbo.GetSaldoCtaIF('" & Format(pdFecha, gsFormatoFecha) & "',cCtaContCod+cCtaIFSubCta,ci.cperscod,ci.ciftpo,ci.cctaifcod, " _
      & " substring(ci.cCtaIfCod,3,1))>0  AND ci.cPersCod+'.'+ ci.cIfTpo+'.'+ci.cCtaIfCod='" & pcCtaContCod & "'"

Set rs = oConec.CargaRecordSet(sql)
Set GetCtaIFCont = rs
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetCtaIFContArb(ByVal pdFecha As Date) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = " Select DISTINCT ci.cperscod+'.'+ci.ciftpo+'.'+ci.cCtaIFcod as cCtaCont,cPersNombre +' . '+ ci.cctaifdesc,1 " _
      & " From CtaIf ci Inner Join CtaIfFiltro cif on cif.cPersCod=ci.cPersCod and cif.cIfTpo=ci.cIfTpo" _
      & " and  cif.cCtaIfCod=ci.cCtaIfCod and cctacontcod like '11[12][37]%' " _
      & " Inner Join Persona p on ci.cPersCod=p.cPersCod " _
      & " where ci.ciftpo='01' and ci.cCtaIfCod like '03%' and cctaifestado=1 " _
      & " and   dbo.GetSaldoCtaIF('" & Format(pdFecha, gsFormatoFecha) & "',cCtaContCod+cCtaIFSubCta,ci.cperscod,ci.ciftpo,ci.cctaifcod, " _
      & " substring(ci.cCtaIfCod,3,1))>0 "

Set rs = oConec.CargaRecordSet(sql)
Set GetCtaIFContArb = rs
oConec.CierraConexion
Set oConec = Nothing
End Function


Public Function fbEliminarGarantia(ByVal psCodPers As String, ByVal psIFTpo As String, ByVal psCtaIfCod As String, ByVal psCodPersPF As String, ByVal psIFTpoPF As String, ByVal psCtaIfCodPF As String, psMovNro) As Boolean
    Dim sSQL As String
    Dim oCon As New COMConecta.DCOMConecta
  
    On Error GoTo mError
    
    sSQL = " Update cgCtaIFAdeudadosGarantia set cMovNroFin='" & psMovNro & "' where cPersCod='" & psCodPers & "' and cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & psCtaIfCod & "' and cPersCodPF='" & psCodPersPF & "'  and  cIFTpoPF='" & psIFTpoPF & "' and  cCtaIFCodPF='" & psCtaIfCodPF & "'"
    oCon.AbreConexion
    oCon.Ejecutar sSQL
    oCon.CierraConexion
    fbEliminarGarantia = 1
    Exit Function
mError:
    fbEliminarGarantia = 0
    Err.Raise vbObjectError + 100, "EliminarGarantiaRestringido", Err.Description
End Function

Public Function GrabaAdeudoVinculado(ByVal psPersCodV As String, ByVal pnIFTpoV As String, ByVal psCtaIfCodV As String, _
                           ByVal pnTipoCambioEuro As Currency, ByVal RsDatos As ADODB.Recordset) As Integer

Dim nbEuros As Integer
Dim MontoEuros As Currency
Dim sSQL As String
Dim oCon As New COMConecta.DCOMConecta
  
    On Error GoTo mError
    oCon.AbreConexion
    
    While Not RsDatos.EOF
         If RsDatos!Valor <> 1 Then

        sSQL = " Update CtaIFAdeudados set " & _
               " cPersCodVin='" & psPersCodV & "' ," & _
               " cIFTpoVin='" & pnIFTpoV & "', " & _
               " cCtaIFCodVin='" & psCtaIfCodV & "' " & _
               " where cPersCod='" & Mid(RsDatos(0), 4, 13) & "' and cIFTpo='" & Mid(RsDatos(0), 1, 2) & "' " & _
               " and cCtaIFCod='" & Mid(RsDatos(0), 18, 10) & "' "
       
        oCon.Ejecutar sSQL
        End If
        RsDatos.MoveNext
      
    Wend
    GrabaAdeudoVinculado = 1
    
    oCon.CierraConexion
    Exit Function
mError:
    GrabaAdeudoVinculado = 0
    Err.Raise vbObjectError + 100, "ActualizaDatosAdeudadoVinculados", Err.Description
End Function

Public Function EliminarAdeudoVinculado(ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIfCod As String) As Boolean

Dim sSQL As String
    Dim oCon As New COMConecta.DCOMConecta
  
    On Error GoTo mError
    
    sSQL = " Update CtaIFAdeudados set " & _
           " cPersCodVin='' ,  cIFTpoVin='' , cCtaIFCodVin='' " & _
           " where cPersCod='" & psPersCod & "' and cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & psCtaIfCod & "' "
           
    oCon.AbreConexion
    
    oCon.Ejecutar sSQL
    oCon.CierraConexion
    EliminarAdeudoVinculado = True
    Exit Function
mError:
    EliminarAdeudoVinculado = False
    Err.Raise vbObjectError + 100, "EliminarGarantiaRestringido", Err.Description
End Function

Public Function GetCarteraRecupCapital(ByVal psFecha As String)
    Dim sSQL As String
    Dim psFecha1 As String
    Dim oCon As New COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    psFecha1 = Mid(psFecha, 4, 2) & Mid(psFecha, 7, 4)
    oCon.AbreConexion
    sSQL = " Exec Stp_Sel_RecuperacionCapitalCartera '" & psFecha1 & "'"
    Set GetCarteraRecupCapital = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
 End Function
