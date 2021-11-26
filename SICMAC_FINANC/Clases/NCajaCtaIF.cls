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

Public Function GetCtaIfFiltro(ByVal psCtaContCod As String, psCtaIFCod As String, psPersCod As String, Optional lbMuestraCta As Boolean = True) As String
Dim oCtaIf As DCajaCtasIF
Set oCtaIf = New DCajaCtasIF
GetCtaIfFiltro = oCtaIf.GetCtaIfFiltro(psCtaContCod, psCtaIFCod, psPersCod, lbMuestraCta)
Set oCtaIf = Nothing
End Function


Public Function GeneraCodCtaIF(psCtaIFCod As String, psMoneda As String, psTpoCta As String) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oConect As DConecta
    
    Set rs = New ADODB.Recordset
    Set oConect = New DConecta
    GeneraCodCtaIF = ""
    If oConect.AbreConexion = False Then Exit Function
    
    sql = "Select   Max(cCtaIfCod) as MaxCtaIF " _
        & " From    CTAIF " _
        & " Where   cPersCod='" & psCtaIFCod & "' " _
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
Public Function NombreIF(ByVal psPersCodIf As String) As String
Dim oCtasIF As DCajaCtasIF
Set oCtasIF = New DCajaCtasIF
NombreIF = oCtasIF.NombreIF(psPersCodIf)
Set oCtasIF = Nothing
End Function
Public Function SubCuentaIF(ByVal psPersCodIf As String) As String
Dim oCtasIF As DCajaCtasIF
Set oCtasIF = New DCajaCtasIF
SubCuentaIF = oCtasIF.SubCuentaIF(psPersCodIf)
Set oCtasIF = Nothing
End Function
'Public Function CargaCtasIF(psMoneda As Moneda, Optional psFiltroCtaIF As String, Optional pnMuestraIF As MuestraIF = MuestraCuentas, Optional psCanje As String = "", Optional pbFiltraActivos As Boolean = True, Optional ByVal bMostrarSoloVigentes As Boolean = False) As ADODB.Recordset
Public Function CargaCtasIF(psMoneda As Moneda, Optional psFiltroCtaIF As String, Optional pnMuestraIF As MuestraIF = MuestraCuentas, Optional psCanje As String = "", Optional pbFiltraActivos As Boolean = True, Optional ByVal bMostrarSoloVigentes As Boolean = False, Optional ByVal psListaPersCod As String = "") As ADODB.Recordset
Dim oCtasIF As DCajaCtasIF
Set oCtasIF = New DCajaCtasIF
Set CargaCtasIF = oCtasIF.CargaCtasIF(psMoneda, psFiltroCtaIF, pnMuestraIF, psCanje, , pbFiltraActivos, bMostrarSoloVigentes, psListaPersCod)
Set oCtasIF = Nothing
End Function
Public Function EmiteTipoCuentaIF(psCodCtaIf As String) As String
    On Error GoTo EmiteTipoCuentaIFErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oConect As DConecta
    
    EmiteTipoCuentaIF = ""
    Set oConect = New DConecta
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
    Call RaiseError(MyUnhandledError, "NCajaCtaIF:EmiteTipoCuentaIF Method")
End Function
Public Function EmiteCuentaContBco(psCodCtaIf As String) As String
    On Error GoTo EmiteCuentaContBcoErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oConect As DConecta
    
    EmiteCuentaContBco = ""
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then Exit Function
    
    Set rs = New ADODB.Recordset
    sql = "Select cCtaContCod, cCtaIFSubCta from " & vsServerCom & "CtaIfFiltro " _
        & "where cIFTpo + '.' + cPerscod + '.' + cCtaIfcod = '" & psCodCtaIf & "' and cCtaContCod Like '11_3%'"
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        EmiteCuentaContBco = Trim(rs!cCtaContCod + rs!cCtaIFSubCta)
    End If
    rs.Close: Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing

    Exit Function
EmiteCuentaContBcoErr:
    Call RaiseError(MyUnhandledError, "NCajaCtaIF:EmiteCuentaContBco Method")
End Function

Public Function ValidaBanco(psCtaIf As String) As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set oConec = New DConecta
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
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set oConec = New DConecta
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

Public Function GetIFDesc(psPersCod As String, Optional lbAnexaTpoCuenta As Boolean = False, Optional psCtaIFCod As String = "") As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set oConec = New DConecta
    GetIFDesc = ""
    If oConec.AbreConexion = False Then Exit Function
    If Not lbAnexaTpoCuenta Then
        sSQL = "SELECT cPersCod CtaIfCod, cPersNombre cCtaIFDesc " _
             & "FROM Persona WHERE cPersCod = '" & psPersCod & "' "
    Else
        sSQL = "SELECT p.cPersCod+'.'+cCtaIfCod as CtaIfCod, RTRIM(cPersNombre) + ' - ' + ci.cCtaIFDesc cCtaIFDesc " _
             & "FROM   Persona p JOIN CtaIf ci ON ci.cPersCod = p.cPersCod " _
             & "WHERE  p.cPersCod = '" & psPersCod & "' and cCtaIfCod LIKE '" & Left(psCtaIFCod, 5) & "' "
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
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni  As ClasIni
Set oIni = New ClasIni

vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
End Sub

Public Function GetInstFinancieras(Optional ByVal psFiltroTipoCtaIF As String = "") As Recordset
Dim oConec As DConecta
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

Set oConec = New DConecta
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

Public Function GetCtasInstFinancieras(Optional ByVal psFiltroTipoCtaIF As String = "", Optional ByVal psCtaCont As String = "", Optional pnNivel As MuestraIF = MuestraCuentas, Optional ByVal psFiltroTipoIF As String) As ADODB.Recordset
Dim oConec As DConecta
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsIFFiltro As String
Dim lsFiltroCta As String
Dim Pos As String
Dim lsCadAux As String
Dim lsFiltroTipoIF As String
lsFiltroCta = ""
If psCtaCont <> "" Then
    lsFiltroCta = " WHERE CIF.cCtaContCod LIKE '" & psCtaCont & "'  "
End If
lsIFFiltro = ""
If psFiltroTipoCtaIF <> "" Then
    If psCtaCont = "" Then
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

Set oConec = New DConecta
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
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set rs = New ADODB.Recordset
    Set oConec = New DConecta
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
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New DConecta
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
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New DConecta
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
                                  ByVal psCtaCont As String, ByVal psCtaIFCod As String) As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New DConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSQL = "Select  cCtaIfCod, C.cCtaContCod, cCtaIFSubCta,C.cCtaContCod + C.cCtaIFSubCta" _
        & " From    CtaIFFiltro  C   " _
        & "         JOIN CTACONT CC ON CC.CCTACONTCOD  = C.cCtaContCod+ C.cCtaIFSubCta " _
        & " Where   C.cCtaContCod Like '" & psCtaCont & "' " _
        & "         And cCtaIfCod Like '" & Mid(psCtaIFCod, 1, 5) & "%' And cPersCod ='" & psPersCod & "' AND cIFTpo ='" & Format(pnTipoIF, "00") & "' "
    
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
                             ByVal psCtaIFCod As String, ByVal pdFecha As Date) As Currency
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New DConecta
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
         & "     AND cia.cCtaIFCod='" & psCtaIFCod & "' "

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
                             ByVal psCtaIFCod As String, ByVal pdFecha As Date, _
                             Optional ByVal psCtaContCod As String = "", Optional ByVal pnMoneda As Moneda = gMonedaNacional) As Currency
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New DConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSQL = "Select  MS.cPersCod, MS.cIFTpo, MS.cCtaIFCod, nSaldo=ISNULL(nSaldo,0.00), nSaldoME=ISNULL(nSaldoME,0.00)" _
        & " From    CtaIFSaldo MS " _
        & " Where   MS.dCtaIFSaldo = (  Select  Max(MS1.dCtaIFSaldo) " _
        & "                             From    CtaIFSaldo MS1 " _
        & "                             Where   MS.cCtaContCod = MS1.cCtaContCod and MS.cPersCod = MS1.cPersCod AND " _
        & "                                     MS.cIFTpo = MS1.cIFTpo And MS.cCtaIFCod = MS1.cCtaIFCod " _
        & "                                     AND MS1.dCtaIFSaldo <='" & Format(pdFecha, gsFormatoFecha) & "') " _
        & "         AND MS.cPersCod='" & psPersCod & "' AND MS.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         AND MS.cCtaIFCod='" & psCtaIFCod & "' " & IIf(psCtaContCod <> "", " AND MS.cCtaContCod='" & psCtaContCod & "' ", "")
        
    GetSaldoCtaIf = 0
    Set rs = oConec.CargaRecordSet(sSQL)
    If Not (rs.EOF And rs.BOF) Then
        'GetSaldoCtaIf = Trim(rs!nSaldo)
        GetSaldoCtaIf = IIf(pnMoneda = gMonedaNacional, rs!nSaldo, rs!nSaldoME)
    End If
    rs.Close
    Set rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GetInteresDevengados(ByVal psOpeCod As String, ByVal pdFecha As Date, Optional ByVal pnOpeObjOrden As Integer = 1, Optional psCodCmac As String = "112", Optional lbCapitaliza As Boolean = False) As ADODB.Recordset
Dim oConec As DConecta
Dim oOpe   As DOperacion
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsCtaCapital As String

Set oOpe = New DOperacion
Set oConec = New DConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

Set rs = oOpe.CargaOpeCtaUltimoNivel(psOpeCod, "D", "2")
'lsCtaCapital = RSMuestraLista(rs) 'Comentado by NAGL 20190612
lsCtaCapital = RSMuestraListaCadenaNew(rs) 'Agregado by NAGL 20190612
Set oOpe = Nothing
If lsCtaCapital = "" Then
    Err.Raise 50001, "NCajaCtaIF:GetIntereses", "No se definión Cuenta Contable en Operación"
End If

'Se modifico la formula del calcilo para el interes de acuerdo a formula de caja, se quito nInteres para que no sume con la capital
'EJVG20120430 Se puso que muestren las ctas con saldo mayor a cero
'CASE WHEN Periodo >0 Then  (POWER((1 + IntPorc*1.00000000),(Dias*1.00000000/Periodo*1.00000000))-1)*(nCapital) ELSE 0 END as IntCalculado,"_'***Modificado ELRO el 20130613, según SATI INC1306040019
'sql = " Select   cCtaIFDesc, cInstFinanc,nCapital ," _
'    & "         IntValor, nInteres, Convert(CHAR(12),UltFecha,103) UltFecha, Dias, " _
'    & "         CASE WHEN Periodo >0 Then  (POWER((1 + IntPorc*1.00000000),(Dias*1.00000000/Periodo*1.00000000))-1)*(nCapital+nInteres) ELSE 0 END as IntCalculado, " _
'    & "         cPersCod , cIFTpo, cCtaIFCod, nCtaIFPlazo, dctaifvenc " _
'    & " From  ( " _
'    & "         SELECT  CI.cPersCod, CI.cIFTpo, CI.cCtaIFCod, CI.cCtaIFDesc, cPersNombre AS cInstFinanc , " _
'    & "                 CI.dCtaIfCap, CI.nInteres, CI.nCtaIFPlazo, ISNULL(CASE WHEN Saldo.cCtaIFCod LIKE '__2%' THEN Saldo.nSaldoME ELSE Saldo.nSaldo END,0) AS nCapital, " _
'    & "                 IsNull(INTERES.nCtaIFIntPeriodo,0) as Periodo, " _
'    & "                 IsNull(INTERES.nCtaIFIntValor,0) as IntValor, " _
'    & "                 IsNull(INTERES.nCtaIFIntValor,0)/100 as IntPorc, " _
'    & "                 CASE WHEN CI.dCtaIfInt >= CI.dCtaIfCap THEN  CI.dCtaIfInt " _
'    & "                      ELSE  CI.dCtaIfCap END  as UltFecha , " _
'    & "                 DateDiff(day,  CASE WHEN CI.dCtaIfInt >= CI.dCtaIfCap THEN CI.dCtaIfInt " _
'    & "                                 ELSE  CI.dCtaIfCap END, '" & Format(pdFecha, gsFormatoFecha) & "') " & IIf(psCodCmac = "102", "+ CASE WHEN CI.nInteres <> 0 THEN " & IIf(lbCapitaliza, "-1", "0") & " ELSE " & IIf(lbCapitaliza, "0", "1") & " END", "") & " AS Dias, CI.dctaifvenc " _
'    & "          FROM   CTAIF CI " _
'    & "                 JOIN INSTITUCIONFINANC I ON I.cPersCod = CI.cPersCod AND I.cIFTpo = CI.cIFTpo " _
'    & "                 JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD " _
'    & "                 LEFT JOIN ( SELECT  CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod, SUM(CS.nSaldo) nSaldo, SUM(CS.nSaldoME) nSaldoME FROM CTAIFSALDO CS " _
'    & "                             WHERE   CS.cCtaContCod in ( " & lsCtaCapital & ") and CS.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
'    & "                                                         FROM    CTAIFSALDO CS1 " _
'    & "                                                         Where   cs1.cCtaContCod = cs.cCtaContCod and CS1.cPersCod = CS.cPersCod And CS1.cIFTpo = CS.cIFTpo And CS1.cCtaIFCod = CS.cCtaIFCod " _
'    & "                                                                 AND CS1.dCtaIFSaldo<='" & Format(pdFecha, gsFormatoFecha) & "') GROUP BY CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod ) AS SALDO " _
'    & "                 ON SALDO.cPersCod=CI.cPersCod and SALDO.cIFTpo = CI.cIFTpo and SALDO.cCtaIFCod= CI.cCtaIFCod " _
'    & "                 LEFT JOIN ( SELECT  I.cPersCod, I.cIFTpo, I.cCtaIFCod, I.dCtaIFIntRegistro, I.nCtaIFIntPeriodo, I.nCtaIFIntValor "
'sql = sql + "                       FROM    CtaIFInteres I " _
'    & "                             WHERE   I.dCtaIFIntRegistro = ( SELECT  MAX(dCtaIFIntRegistro) " _
'    & "                                                             FROM    CtaIFInteres I1 " _
'    & "                                                             Where   I1.cPersCod = I.cPersCod And I1.cIFTpo = I.cIFTpo And I1.cCtaIFCod = I.cCtaIFCod " _
'    & "                                                                     AND I1.dCtaIFIntRegistro<='" & Format(pdFecha, gsFormatoFecha) & "') ) AS INTERES " _
'    & "                 ON INTERES.cPersCod=CI.cPersCod and INTERES.cIFTpo = CI.cIFTpo and INTERES.cCtaIFCod= CI.cCtaIFCod, OPEOBJ O "
'    'Modificado PASI20161209***********************
'    '& "         WHERE   ISNULL(CASE WHEN Saldo.cCtaIFCod LIKE '__2%' THEN Saldo.nSaldoME ELSE Saldo.nSaldo END,0) > 0  And LEN(CI.cCtaIFCod)=(Select Max(Len(cCtaIFCod)) From CtaIf) AND cCtaIFEstado ='" & gEstadoCtaIFActiva & "' " _
'    '& "                 AND o.cOpeCod = '" & psOpeCod & "' and cOpeObjOrden='" & pnOpeObjOrden & "' and " _
'    '& "                 CI.cIFTpo+CI.cCtaIFCod Like cOpeObjFiltro  " _
'    '& "         ) AS DATOS " _
'    '& " Order By cInstFinanc "
'  sql = sql + "         WHERE   ISNULL(CASE WHEN Saldo.cCtaIFCod LIKE '__2%' THEN Saldo.nSaldoME ELSE Saldo.nSaldo END,0) > 0  And LEN(CI.cCtaIFCod) >= 7 AND cCtaIFEstado ='" & gEstadoCtaIFActiva & "' " _
'    & "                 AND o.cOpeCod = '" & psOpeCod & "' and cOpeObjOrden='" & pnOpeObjOrden & "' and " _
'    & "                 CI.cIFTpo+CI.cCtaIFCod Like cOpeObjFiltro  " _
'    & "         ) AS DATOS " _
'    & " Order By cInstFinanc "
'    'PASI END **************************************
sql = "Exec stp_sel_GetIntereseDevengados '" & psOpeCod & "', '" & lsCtaCapital & "','" & Format(pdFecha, "yyyymmdd") & "', " & pnOpeObjOrden & ", '" & psCodCmac & "', " & IIf(lbCapitaliza, 1, 0) & ""
Set rs = oConec.CargaRecordSet(sql)
Set GetInteresDevengados = rs
oConec.CierraConexion
Set oConec = Nothing
End Function
Public Function GetCuentasIFCanc(ByVal psOpeCod As String, ByVal pdFecha As Date, Optional ByVal pnOpeObjOrden As Integer = 1) As ADODB.Recordset
Dim oConec As DConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New DConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

'sql = "SELECT   CI.cCtaIFDesc, cPersNombre AS cInstFinanc , " _
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
    
sql = "stp_sel_GetCuentasIFCanc '" & psOpeCod & "','" & Format(pdFecha, gsFormatoFecha) & "'" 'PASI20161209

Set rs = oConec.CargaRecordSet(sql)
Set GetCuentasIFCanc = rs
oConec.CierraConexion
Set oConec = Nothing
End Function
Public Function GetInteresCtaIf(ByVal psPersCod As String, ByVal pnTipoIF As CGTipoIF, _
                             ByVal psCtaIFCod As String) As Currency
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New DConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSQL = "Select  MS.cPersCod, MS.cIFTpo, MS.cCtaIFCod, nInteres " _
        & " From    CtaIF MS " _
        & " WHERE   MS.cPersCod='" & psPersCod & "' AND MS.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         AND MS.cCtaIFCod='" & psCtaIFCod & "'"
        
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
Public Function GetRsInteresCtasIF(ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, ByVal psCtaIFCod As String) As ADODB.Recordset
Dim oConec As DConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New DConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = "Select convert(char(10),dCtaIFIntRegistro,103) as FechaInt, nCtaIFIntValor , nCtaIFIntPeriodo, cCtaIFCod ,cPersCod, cIFTpo,  1 as  Registrado " _
    & " from CtaifInteres " _
    & " Where cPersCod='" & psPersCod & "' and cIFTpo ='" & Format(pnIFTpo, "00") & "' and cCtaIFCod='" & psCtaIFCod & "'"

Set rs = oConec.CargaRecordSet(sql)
Set GetRsInteresCtasIF = rs
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetDatosCtaIf(ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, ByVal psCtaIFCod As String) As ADODB.Recordset
Dim oConec As DConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New DConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = " SELECT  C.dCtaIFAper, C.dCtaIfCap, C.dCtaIfInt, C.nCtaIFPlazo, C.nInteres, C.cCtaIFEstado,C.nMontoEuros MontoEurosCtaIF,CC.cConsDescripcion AS cEstadoCons, cTpoC.cConsDescripcion cTpoCtaDesc, Isnull(SALDO.nSaldo,0) as Saldo, dCtaIFVenc,bContable=isnull(C.bContable,0),nMovNroOriginal=C.nMovNroOriginal, cia.* " _
    & " FROM    CTAIF C  LEFT JOIN CtaIFAdeudados cia ON cia.cPersCod = c.cPersCod and cia.cIFTpo = c.cIFTpo and cia.cCtaIFCod = c.cCtaIFCod " _
    & "         JOIN CONSTANTE CC ON    CC.NCONSVALOR =  C.cCtaIFEstado LEFT JOIN Constante cTpoC ON cTpoC.nConsValor = cia.nTpoCuota AND cTpoC.nCONSCOD = " & gCGAdeudCalTpoCuota & " " _
    & "         LEFT JOIN ( SELECT  CS.dCtaIFSaldo, CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod, CASE WHEN SubString(cs.cCtaIFCod,3,1) = '1' THEN CS.nSaldo ELSE cs.nSaldoME END nSaldo" _
    & "                     FROM    CTAIFSALDO CS " _
    & "                     WHERE   CS.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
    & "                                                 FROM    CTAIFSALDO CS1 " _
    & "                                                 Where   CS1.cPersCod = CS.cPersCod And CS1.cIFTpo = CS.cIFTpo And CS1.cCtaIFCod = CS.cCtaIFCod )) AS SALDO " _
    & "                     ON SALDO.cPersCod=C.cPersCod and SALDO.cIFTpo = C.cIFTpo and SALDO.cCtaIFCod= C.cCtaIFCod " _
    & " WHERE CC.nCONSCOD = " & gCGEstadoCtaIF & " AND C.cPersCod='" & psPersCod & "' and C.cIFTpo ='" & Format(pnIFTpo, "00") & "' and C.cCtaIFCod = '" & psCtaIFCod & "' ORDER BY nSaldo DESC "

Set rs = oConec.CargaRecordSet(sql)
Set GetDatosCtaIf = rs
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetDatosAcumRestante(ByVal pdFecha As Date, ByVal psPersCod As String, ByVal psCtaIFCod As String) As ADODB.Recordset
Dim oConec As New DConecta
Dim sql As String
Dim rs As New ADODB.Recordset
On Error GoTo ErroGetDatosAcumRestante
    sql = "Exec sp_sel_MontoAedudadosRestante '" & Format(pdFecha, "yyyymmdd") & "','" & psPersCod & "', '" & psCtaIFCod & "'"
    oConec.AbreConexion
    Set GetDatosAcumRestante = oConec.CargaRecordSet(sql)
    oConec.CierraConexion
Exit Function
ErroGetDatosAcumRestante:
Call RaiseError(MyUnhandledError, "Lista Datos Adudados Restantes")
End Function 'NAGL 20170519

Public Function GetCtaIfInteres(ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, ByVal psCtaIFCod As String) As Currency
Dim oConec As DConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New DConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = " SELECT  nCtaIFIntValor " _
    & " FROM    CTAIFINTERES  C  " _
    & " WHERE   dCtaIFIntRegistro = (SELECT Max(dCtaIFIntRegistro) FROM CtaIFInteres cia WHERE cia.cPersCod = c.cPersCod and cia.cIFTpo = c.cIFTpo and cia.cCtaIFCod = c.cCtaIFCod )" _
    & "  and C.cPersCod='" & psPersCod & "' and C.cIFTpo ='" & Format(pnIFTpo, "00") & "' and C.cCtaIFCod = '" & psCtaIFCod & "'"

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

Public Function ActualizaCtas(ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, _
                                ByVal psCtaIFCod As String, ByVal psCtaIFDesc As String, _
                                ByVal pdCtaIFAper As Date, ByVal pdCtaIFCap As Date, _
                                ByVal pdCtaIFInt As Date, ByVal pnCtaIFPlazo As Integer, _
                                ByVal pnInteres As Currency, ByVal pnCtaIFEstado As CGEstadoCtaIF, _
                                ByVal psMovUltAct As String, ByVal rsInt As ADODB.Recordset, _
                                ByVal pnMontoPrestado As Currency, pnNroCuotas As Integer, _
                                ByVal pnGracia As Currency, pnComisionIni As Currency, pnComisionMonto As Currency, _
                                ByVal psInterno As String, pnCuotaPagok As Integer, pdFecVencimiento As Date, pdFecUltPago As Date, _
                                ByVal pnTpoCuota As CGAdeudCalTpoCuota, ByVal pnTramo As Currency, ByVal rsAdeud As ADODB.Recordset, _
                                ByVal pnMonedaPago As Moneda, ByVal pnComisionCuota As Currency, psLinCredCod As String, pnFechaFija As Integer, nMontoPrestadoVac As Currency, ByVal pnPorcentAfect As Currency, _
                                Optional ByVal pnMalPg As Integer, Optional pnMontoEuros As Currency, Optional pnMontoPresConce As Currency = 0, Optional pnContable As Integer = 1, Optional psMovNro As String = "", Optional pnMovNroOriginal As Long = 0) As Integer
Dim oMov As DMov
Dim lbTrans As Boolean
Dim lnMovNro As String
Set oMov = New DMov
On Error GoTo ErrorActualizaCtas
oMov.BeginTrans
lbTrans = True
oMov.ActualizaCtaIF psPersCod, pnIFTpo, psCtaIFCod, Trim(psCtaIFDesc), _
                    Format(pdCtaIFAper, gsFormatoFecha), Format(pdCtaIFCap, gsFormatoFecha), Format(pdCtaIFInt, gsFormatoFecha), pnCtaIFPlazo, Format(pdFecVencimiento, gsFormatoFecha), pnCtaIFEstado, pnInteres, psMovUltAct, , pnMontoEuros, 1
If Not rsInt Is Nothing Then
    Do While Not rsInt.EOF
        If Val(rsInt!lbnuevo) = 2 Then
            'oMov.EliminaCuentaIFInteres
            oMov.ActualizaCuentaIFInteres psPersCod, pnIFTpo, psCtaIFCod, rsInt!Registro, rsInt![Período], rsInt!Interes, psMovUltAct
        End If
        If Val(rsInt!lbnuevo) <> 1 And Val(rsInt!lbnuevo) <> 2 Then
            oMov.InsertaCuentaIFInteres psPersCod, pnIFTpo, psCtaIFCod, rsInt!Registro, rsInt![Período], rsInt!Interes, psMovUltAct
        End If
        rsInt.MoveNext
    Loop
End If

If Left(psCtaIFCod, 2) = gTpoCtaIFCtaAdeud Then
    'ALPA20130719********
   'oMov.InsertaDatosAdeudados rsAdeud, psPersCod, pnIFTpo, psCtaIFCod, pnMontoPrestado, pnNroCuotas, pnGracia, pnComisionIni, pnComisionMonto, psInterno, pnCuotaPagok, Format(pdFecVencimiento, gsFormatoFecha), CInt(pnMonedaPago), Format(pdFecUltPago, "YYYY/MM/DD"), pnTpoCuota, pnTramo, pnComisionCuota, psLinCredCod, pnFechaFija, psMovUltAct, pnMalPg, nMontoPrestadoVac, pnPorcentAfect, pnMontoEuros
   oMov.InsertaDatosAdeudados rsAdeud, psPersCod, pnIFTpo, psCtaIFCod, pnMontoPrestado, pnNroCuotas, pnGracia, pnComisionIni, pnComisionMonto, psInterno, pnCuotaPagok, Format(pdFecVencimiento, gsFormatoFecha), CInt(pnMonedaPago), Format(pdFecUltPago, "YYYY/MM/DD"), pnTpoCuota, pnTramo, pnComisionCuota, psLinCredCod, pnFechaFija, psMovUltAct, pnMalPg, nMontoPrestadoVac, pnPorcentAfect, pnMontoEuros, pnMontoPresConce
   '*********************
End If
If pnContable = 0 Then
    Dim lsMovNro As String
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    Set oRs = oMov.ObtenerMov(pnMovNroOriginal)
    If Not (oRs.BOF Or oRs.EOF) Then
        lsMovNro = psMovNro
        oMov.InsertaMov lsMovNro, oRs!cOpeCod, oRs!cMovdesc, gMovEstContabMovContable
        lnMovNro = oMov.GetnMovNro(psMovNro)
        oMov.ObtenerMovContMIVIVIENDA lnMovNro, pnMovNroOriginal
        oMov.ObtenerMovCtaMIVIVIENDA lnMovNro, pnMovNroOriginal
        oMov.ObtenerMovObjMIVIVIENDA lnMovNro, pnMovNroOriginal
        oMov.ObtenerMovObjIFMIVIVIENDA lnMovNro, pnMovNroOriginal
        oMov.ObtenerMovObjEfectivoMIVIVIENDA lnMovNro, pnMovNroOriginal
        oMov.ObtenerMovObjAreaAgenciaMIVIVIENDA lnMovNro, pnMovNroOriginal
        oMov.ObtenerMovRefMIVIVIENDA lnMovNro, pnMovNroOriginal
        oMov.ObtenerMovDocMIVIVIENDA lnMovNro, pnMovNroOriginal
        If Mid(oRs!cOpeCod, 3, 1) = gMonedaExtranjera Then
            oMov.GeneraMovME lnMovNro, lsMovNro
        End If
        oMov.ActualizaSaldoMovimiento lsMovNro, "+"
    End If

End If
oMov.CommitTrans
lbTrans = False
Exit Function
ErrorActualizaCtas:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, "Actualiza Ctas", Err.Description
End Function

'Public Function GetFlujosCtaIF(ByVal psOpeCod As String, _
'                            ByVal psCtaIfCodDesde As String, ByVal psCtaIfCodHasta As String, _
'                            ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
'Dim oConec As DConecta
'Dim sql As String
'Dim rs As ADODB.Recordset
'
'Set oConec = New DConecta
'Set rs = New ADODB.Recordset
'If oConec.AbreConexion() = False Then Exit Function
'
'sql = " SELECT CI.cPersCod, CI.cIFTpo, CI.cCtaIFCod, P.CPERSNOMBRE, CI1.cCtaIFDesc, CI.cCtaIFDesc," _
'    & "        RTRIM(P.CPERSNOMBRE) + ' ' + CI1.cCtaIFDesc + ' ' + CI.cCtaIFDesc AS cNomCtaIF, " _
'    & "        ISNULL(SALDO.nSaldo,0) AS nSaldoIni, ISNULL(nSaldoME,0) nSaldoIniME, " _
'    & "        CI.nCtaIFPlazo, CI.nInteres , " _
'    & "        M.CMOVNRO, M.CMOVDESC, IsNull(D.cDocAbrev,'')  + ' ' + IsNull(MD.cDocNro,'') as Documento , " _
'    & "        MO.nMovImporte,  ISNULL(ME.nMovMEImporte,0) nMovMEImporte, " _
'    & "        Case WHEN MO.nMovImporte > 0 THEN ABS(MO.nMovImporte) ELSE 0 END AS DEBE , Case WHEN ISNULL(ME.nMovMEImporte,0) > 0 THEN ABS(ME.nMovMEImporte) ELSE 0 END AS DEBEME, " _
'    & "        Case WHEN MO.nMovImporte < 0 THEN ABS(MO.nMovImporte) ELSE 0 END AS HABER, Case WHEN ISNULL(ME.nMovMEImporte,0) < 0 THEN ABS(ME.nMovMEImporte) ELSE 0 END AS HABERME " _
'    & " FROM   MOV M " _
'    & "        JOIN MovCta mo ON mo.nMovNro = m.nMovNro LEFT JOIN MovME me ON me.nMovNro = mo.nMovNro and me.nMovItem = mo.nMovItem " _
'    & "        JOIN MOVOBJIF MC ON MC.nMovNro = mo.nMovNro and mc.nMovItem = mo.nMovItem " _
'    & "        LEFT JOIN MOVDOC MD ON MD.NMOVNRO = M.NMOVNRO AND MD.nDocTpo <>" & TpoDocVoucherEgreso & "  " _
'    & "        JOIN CTAIF CI ON MC.cPersCod=CI.cPersCod and MC.cIFTpo = MC.cIFTpo and MC.cCtaIFCod= CI.cCtaIFCod " _
'    & "        JOIN CTAIF CI1 ON CI1.cPersCod = CI.cPersCod AND CI1.cIFTpo = CI.cIFTpo  AND SUBSTRING(CI.cCtaIFCod,1,LEN(CI1.cCtaIFCod)) = CI1.cCtaIFCod and Len(CI1.cCtaIFCod)=3 " _
'    & "        JOIN INSTITUCIONFINANC I ON I.cPersCod = CI.cPersCod AND I.cIFTpo = CI.cIFTpo " _
'    & "        JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD " _
'    & "         LEFT JOIN ( SELECT  CS.cCtaContCod, CS.dCtaIFSaldo, CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod, CS.nSaldo, CS.nSaldoME " _
'    & "                     FROM    CTAIFSALDO CS " _
'    & "                     WHERE   CS.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
'    & "                                                 FROM    CTAIFSALDO CS1 " _
'    & "                                                 Where   cs1.cCtaContCod = cs.cCtaContCod and CS1.cPersCod = CS.cPersCod And CS1.cIFTpo = CS.cIFTpo " _
'    & "                                                         And CS1.cCtaIFCod = CS.cCtaIFCod " _
'    & "                                                         AND CS1.dCtaIFSaldo<='" & Format(DateAdd("d", -1, pdDesde), "mm/dd/yyyy") & "' )) AS SALDO" _
'    & "         ON SALDO.cPersCod=CI.cPersCod and SALDO.cIFTpo = CI.cIFTpo and SALDO.cCtaIFCod= CI.cCtaIFCod and Saldo.cCtaContCod = mo.cCtaContCod " _
'    & "         LEFT JOIN DOCUMENTO D ON D.nDocTpo = MD.nDocTpo, OPEOBJ O "
'sql = sql + "WHERE   LEN(CI.cCtaIFCod)= (Select max(len(cCtaIFCod)) From CtaIf) AND CI.cCtaIFEstado ='1' " _
'    & "         AND o.cOpeCod = '" & psOpeCod & "' and cOpeObjOrden='0' " _
'    & "         AND CI.cIFTpo + CI.cCtaIFCod like cOpeObjFiltro " _
'    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " _
'    & "         AND CI.cPersCod + '.' + CI.cCtaIFCod BETWEEN '" & psCtaIfCodDesde & "' AND '" & psCtaIfCodHasta & "' " _
'    & "         and ( M.NMOVFLAG IN (" & gMovFlagVigente & "," & gMovFlagDeExtorno & ") or (m.NMovFlag = " & gMovFlagExtornado & " and EXISTS (SELECT mr1.NMovNro FROM MovRef mr1 JOIN Mov m1 ON m1.nMovNro = mr1.nMovNro WHERE mr1.NMovNroRef = m.NMovNro and LEFT(m1.cMovNro,6) > LEFT(m.cMovNro,6) ))) " _
'    & "         and mo.cCtaContCod LIKE '11_[23]%'  " _
'    & " Order by CI.cIFTpo+ CI.cPersCod + CI.cCtaIFCod , M.CMOVNRO"
'
'Set rs = oConec.CargaRecordSet(sql)
'Set GetFlujosCtaIF = rs
'oConec.CierraConexion
'Set oConec = Nothing
'End Function

Public Function GetFlujosCtaIF(ByVal psOpeCod As String, _
                            ByVal psCtaIfCodDesde As String, ByVal psCtaIfCodHasta As String, _
                            ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset

Dim oConec As DConecta
Dim sql As String
Dim rs As ADODB.Recordset


Set oConec = New DConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

 '*** PEAC 20110906
'sql = " SELECT CI.cPersCod, CI.cIFTpo, CI.cCtaIFCod, P.CPERSNOMBRE, CI1.cCtaIFDesc, CI.cCtaIFDesc," _
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
'sql = sql + "WHERE   LEN(CI.cCtaIFCod)= (Select max(len(cCtaIFCod)) From CtaIf) AND CI.cCtaIFEstado ='1' " _
    & "         AND o.cOpeCod = '" & psOpeCod & "' and cOpeObjOrden='0' " _
    & "         AND CI.cIFTpo + CI.cCtaIFCod like cOpeObjFiltro " _
    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "         AND CI.cPersCod + '.' + CI.cCtaIFCod BETWEEN '" & psCtaIfCodDesde & "' AND '" & psCtaIfCodHasta & "' " _
    & "         and ( M.NMOVFLAG IN (" & gMovFlagVigente & "," & gMovFlagDeExtorno & ") or (m.NMovFlag = " & gMovFlagExtornado & " and EXISTS (SELECT mr1.NMovNro FROM MovRef mr1 JOIN Mov m1 ON m1.nMovNro = mr1.nMovNro WHERE mr1.NMovNroRef = m.NMovNro and mr1.NMovNro = " & gMovFlagVigente & " and LEFT(m1.cMovNro,6) > LEFT(m.cMovNro,6) ))) " _
    & "         and mo.cCtaContCod LIKE '1[13]_[23]%'  " _
    & " Order by CI.cIFTpo+ CI.cPersCod + CI.cCtaIFCod , M.CMOVNRO"

'    sql = "exec stp_sel_GetFlujosCtaIF " & TpoDocVoucherEgreso & ",'" & Format(DateAdd("d", -1, pdDesde), "yyyymmdd") & "','" & Trim(psOpeCod) & "','" & Format(pdDesde, "yyyymmdd") & "','" & Format(pdHasta, "yyyymmdd") & "','" & Trim(psCtaIfCodDesde) & "','" & Trim(psCtaIfCodHasta) & "'"
sql = "EXEC stp_sel_GetFlujosCtaIF_NEW '" & psOpeCod & "','" & psCtaIfCodDesde & "','" & psCtaIfCodHasta & "','" & Format(pdDesde, "yyyymmdd") & "','" & Format(pdHasta, "yyyymmdd") & "'" 'EJVG20151009 x Inversiones

Set rs = oConec.CargaRecordSet(sql)
Set GetFlujosCtaIF = rs
oConec.CierraConexion
Set oConec = Nothing
End Function



'Public Function GetRepSaldosCtaIf(ByVal psOpeCod As String, ByVal pdFecha As Date, Optional ByVal psOpeObjOrden As String = "'0'", Optional psCtaIFEstado As String = "", Optional psMoneda As String = "") As ADODB.Recordset
'Dim oConec As DConecta
'Dim sql As String
'Dim rs As ADODB.Recordset
'
'Set oConec = New DConecta
'Set rs = New ADODB.Recordset
'If oConec.AbreConexion() = False Then Exit Function
'
'sql = " SELECT  Z.cCtaIfCod , Z.cCtaIFDesc, Z.cBancoDesc, Z.cCtaIFEstado, " _
'    & "         CASE    WHEN SUM(z.nImporte) > 0 THEN str(SUM(z.nImporte), 14, 2) " _
'    & "                 WHEN SUM(z.nImporte) < 0 THEN '0' END as Debe, " _
'    & "         CASE    WHEN SUM(z.nImporte) < 0 THEN str(SUM(z.nImporte) * -1, 14, 2) " _
'    & "                 WHEN SUM(z.nImporte) > 0 THEN '0' END as Haber " _
'    & " FROM (  " _
'    & "         SELECT  CtaIf.cCtaIfCod, CtaIF.cCtaIFDesc, CtaIF.cBancoDesc, CtaIf.cCtaIFEstado, " & IIf(psMoneda = "1", " SUM(CI.NSALDO) ", " SUM(CI.NSALDOME) ") & " AS nImporte " _
'    & "         FROM    CTAIFSALDO CI " _
'    & "                 JOIN " _
'    & "                 (   SELECT  CI.cIFTpo + CI.cPersCod + CI.cCtaIFCod AS cCtaIfCod, CI.cCtaIFDesc, P.cPersNombre cBancoDesc, CI.cCtaIFEstado " _
'    & "                     FROM    CTAIF CI JOIN INSTITUCIONFINANC I ON  I.CPERSCOD = CI.CPERSCOD  AND CI.CIFTPO = I.CIFTPO " _
'    & "                             JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD, OPEOBJ O " _
'    & "                     WHERE   O.cOpeCod = '" & psOpeCod & "' and O.cOpeObjOrden IN (" & psOpeObjOrden & ") and CI.cCtaIFEstado IN ('" & psCtaIFEstado & "') " _
'    & "                             AND CI.cIFTpo + CI.cCtaIFCod like O.cOpeObjFiltro " _
'    & "                 ) as CtaIf ON  SUBSTRING(CI.cIFTpo + CI.cPersCod + CI.cCtaIFCod , 1, LEN(CtaIf.cCtaIfCod)) = CtaIf.cCtaIfCod , OPEOBJ O JOIN OpeCta oc ON oc.cOpeCod = o.cOpeCod " _
'    & "         WHERE   O.cOpeCod = '" & psOpeCod & "' and ci.cCtaContCod LIKE LEFT(oc.cCtaContCod,2)+'" & psMoneda & "'+SubString(oc.cCtaContCod,4,22) + '%' and oc.cOpeCtaOrden IN (" & psOpeObjOrden & ") and O.cOpeObjOrden IN (" & psOpeObjOrden & ") " _
'    & "                 AND CI.cIFTpo + CI.cCtaIFCod Like O.cOpeObjFiltro " & IIf(psMoneda = "", "", " and CI.cCtaIFCod LIKE '__" & psMoneda & "%' ")
'sql = sql + "           AND CI.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
'    & "                                         FROM    CTAIFSALDO CS1 " _
'    & "                                         Where   cs1.cCtaContCod = ci.cCtaContCod and CS1.cPersCod = CI.cPersCod And CS1.cIFTpo = CI.cIFTpo And CS1.cCtaIFCod = CI.cCtaIFCod " _
'    & "                                                 AND CS1.dCtaIFSaldo<='" & Format(pdFecha, gsFormatoFecha) & "') " _
'    & "         GROUP   BY CtaIf.cCtaIfCod , CtaIF.cCtaIFDesc, CtaIF.cBancoDesc, CtaIF.cCtaIFEstado ) AS Z " _
'    & " WHERE nImporte <> 0 " _
'    & " GROUP  BY Z.cCtaIfCod , Z.cCtaIFDesc, Z.cBancoDesc, Z.cCtaIFEstado " _
'    & " ORDER BY Z.cBancoDesc, Z.cCtaIfCod"
'
'
'    '& " ORDER BY  Z.cCtaIFEstado, Z.cBancoDesc, Z.cCtaIfCod"
'
'Set rs = oConec.CargaRecordSet(sql)
'Set GetRepSaldosCtaIf = rs
'oConec.CierraConexion
'Set oConec = Nothing
'
'End Function

Public Function GetRepSaldosCtaIf(ByVal psOpeCod As String, ByVal pdFecha As Date, Optional ByVal psOpeObjOrden As String = "'0'", Optional psCtaIFEstado As String = "", Optional psMoneda As String = "") As ADODB.Recordset
Dim oConec As DConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New DConecta
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
Dim oConec As DConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New DConecta
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
Dim oConec As DConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New DConecta
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

Public Function Restringir(ByVal psMovNro As String, ByVal psOpeCod As String, _
                            ByVal pnCtaCod As String, ByVal pnImporte As Currency, _
                            ByVal psPersCod As String, ByVal psIFTpo As String, _
                            ByVal psCtaIFCod As String, ByVal pnCtaCodRe As String, _
                            ByVal psCodPersA As String, ByVal psIFTpoA As String, _
                            ByVal psCodCtaIFA As String) As Boolean

Dim oConec As DConecta
Dim oMov As DMov
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim sql As String
Dim pSql As String
Dim rs As ADODB.Recordset


On Error GoTo ErrorRestringirCta
Set oMov = New DMov
Set rs = New ADODB.Recordset
Set oConec = New DConecta

oMov.BeginTrans
lbTrans = True
'Insertamos Movimientos
oMov.InsertaMov psMovNro, psOpeCod, "Restringido", gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"

'Insertamos  en el "Debe"
oMov.InsertaMovCta lnMovNro, 0, pnCtaCodRe, pnImporte
oMov.InsertaMovObj lnMovNro, 0, "1", "01"
oMov.InsertaMovObjIF lnMovNro, 0, "1", psPersCod, psIFTpo, psCtaIFCod

'guardamos la cuenta en el haber
oMov.InsertaMovCta lnMovNro, 1, pnCtaCod, pnImporte * -1
oMov.InsertaMovObj lnMovNro, 1, "1", "01"
oMov.InsertaMovObjIF lnMovNro, 1, "1", psPersCod, psIFTpo, psCtaIFCod

If Mid(psCtaIFCod, 3, 1) = "2" Then
    oMov.GeneraMovME lnMovNro, psMovNro
End If


'Registramos la Garantia
    oMov.fbACGRegistrarGarantia lnMovNro, psCodPersA, psIFTpoA, psCodCtaIFA, psPersCod, psIFTpo, psCtaIFCod, psMovNro
    
    oMov.fbACGRegistrarCtaIFFiltro psPersCod, psIFTpo, psCtaIFCod, Left(pnCtaCodRe, 6), Right(pnCtaCodRe, Len(pnCtaCodRe) - 6)
    
    oMov.ActualizaSaldoMovimiento psMovNro, "+"
 
oMov.CommitTrans
Restringir = True
lbTrans = False
Exit Function
ErrorRestringirCta:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
        Restringir = False
    End If
    Err.Raise Err.Number, "Restringir Ctas", Err.Description
End Function

Public Function Habilitar(ByVal psMovNro As String, ByVal psOpeCod As String, _
                          ByVal pnCtaCod As String, ByVal pnImporte As Currency, _
                          ByVal psPersCod As String, ByVal psIFTpo As String, _
                          ByVal psCtaIFCod As String, ByVal pnCtaCodRe As String, _
                          ByVal psCodPersA As String, ByVal psIFTpoA As String, _
                          ByVal psCodCtaIFA As String) As Boolean
                          

Dim oConec As DConecta
Dim oMov As DMov
Dim lnMovNro As Long
Dim lbTrans As Boolean
Dim sql As String
Dim pSql As String
Dim rs As ADODB.Recordset


On Error GoTo ErrorRestringirCta
Set oMov = New DMov
Set rs = New ADODB.Recordset
Set oConec = New DConecta

oMov.BeginTrans
lbTrans = True
'Insertamos Movimientos
oMov.InsertaMov psMovNro, psOpeCod, "habilitar", gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oMov.GetnMovNro(psMovNro)
oMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"

'Insertamos  en el "Debe"
oMov.InsertaMovCta lnMovNro, 0, pnCtaCod, pnImporte
'guardamos la Cta de la IF si es que utiliza esta
oMov.InsertaMovObj lnMovNro, 0, "1", "01"
oMov.InsertaMovObjIF lnMovNro, 0, "1", psPersCod, psIFTpo, psCtaIFCod

'guardamos la cuenta en el haber
oMov.InsertaMovCta lnMovNro, 1, pnCtaCodRe, pnImporte * -1
oMov.InsertaMovObj lnMovNro, 1, "1", "01"
oMov.InsertaMovObjIF lnMovNro, 1, "1", psPersCod, psIFTpo, psCtaIFCod

If Mid(psCtaIFCod, 3, 1) = "2" Then
    oMov.GeneraMovME lnMovNro, psMovNro
End If


''Registramos la Garantia
' '   oMov.fbACGRegistrarGarantia lnMovNro, psCodPersA, psIFTpoA, psCodCtaIFA, psPersCod, psIFTpo, psCtaIfCod, psMovNro
    
'  '  oMov.fbACGRegistrarCtaIFFiltro psPersCod, psIFTpo, psCtaIfCod, Left(pnCtaCodRe, 6), Right(pnCtaCodRe, Len(pnCtaCodRe) - 6)
    
    oMov.ActualizaSaldoMovimiento psMovNro, "+"
 
oMov.CommitTrans
Habilitar = True
lbTrans = False
Exit Function
ErrorRestringirCta:
    If lbTrans Then
        oMov.RollbackTrans
        lbTrans = False
        Habilitar = False
    End If
    Err.Raise Err.Number, "Restringir Ctas", Err.Description


End Function

Public Function foACGBuscarGarantia(ByVal psCodPers As String, ByVal psIFTpo As String, ByVal psCtaIFCod As String, ByVal psFecha As String) As ADODB.Recordset
    Dim sSQL As String
    Dim rs As New ADODB.Recordset
    Dim oCon As New DConecta
  
    On Error GoTo foACGBuscarGarantiaErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = " Select Codigo,Descripcion,Max(CtaK)CtaK,Max(CtaR)CtaR,sum(SK)SK,sum(SR)SR From ("
    sSQL = sSQL & " Select ca.cPersCodPF+'.'+ca.cIfTpoPF+'.'+ca.cCtaIFCodPF Codigo,rTrim(p.cpersnombre) +' - '+ci.cCtaIFDesc Descripcion,"
    sSQL = sSQL & " case when substring(cif.cCtaContCod,4,1)='3' then cif.cCtaContCod+cif.cCtaIFSubCta Else '' end CtaK,"
    sSQL = sSQL & " case when substring(cif.cCtaContCod,4,1)='7' then cif.cCtaContCod+cif.cCtaIFSubCta Else '' end CtaR,"
    sSQL = sSQL & " case when substring(cif.cCtaContCod,4,1)='3' then dbo.GetSaldoCtaIF('" & Format(psFecha, "yyyy/mm/dd") & "',cCtaContCod+cCtaIFSubCta,ca.cPersCodPF,ca.cIfTpoPF,ca.cCtaIFCodPF,substring(ca.cCtaIFCodPF,3,1)) Else 0 end SK,"
    sSQL = sSQL & " case when substring(cif.cCtaContCod,4,1)='7' then dbo.GetSaldoCtaIF('" & Format(psFecha, "yyyy/mm/dd") & "',cCtaContCod+cCtaIFSubCta,ca.cPersCodPF,ca.cIfTpoPF,ca.cCtaIFCodPF,substring(ca.cCtaIFCodPF,3,1)) Else 0 end SR"
    sSQL = sSQL & " from cgCtaIFAdeudadosGarantia ca"
    sSQL = sSQL & " Inner Join CtaIF ci on ca.cPersCodPF=ci.cPersCod and ca.cIFTpoPF=ci.cIFTpo and ca.cCtaIFCodPF=ci.cCtaIFCod"
    sSQL = sSQL & " Inner Join ctaIfFiltro cif on  ca.cPersCodPF=cif.cPersCod and ca.cIFTpoPF=cif.cIFTpo and ca.cCtaIFCodPF=cif.cCtaIFCod and cif.cctacontcod like '11[12][37]%'"
    sSQL = sSQL & " Inner Join Persona p on ca.cPersCodPF=p.cPersCod"
    sSQL = sSQL & " where ca.cPersCod='" & psCodPers & "' and ca.cIfTpo='" & psIFTpo & "' and ca.cCtaIfCod='" & psCtaIFCod & "' and cMovNroFin is NULL "
    sSQL = sSQL & " and nmovnro=(select max(nmovnro)from cgCtaIFAdeudadosGarantia where cPersCod='" & psCodPers & "' and cIfTpo='" & psIFTpo & "' and cCtaIfCod='" & psCtaIFCod & "')"
    sSQL = sSQL & " )Z Group by Codigo,Descripcion"

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set foACGBuscarGarantia = rs
    Exit Function
foACGBuscarGarantiaErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias ")
End Function
Public Function fbEliminarGarantia(ByVal psCodPers As String, ByVal psIFTpo As String, ByVal psCtaIFCod As String, ByVal psCodPersPF As String, ByVal psIFTpoPF As String, ByVal psCtaIfCodPF As String, psMovNro) As Boolean
    Dim sSQL As String
    Dim oCon As New DConecta
  
    On Error GoTo mError
    
    sSQL = " Update cgCtaIFAdeudadosGarantia set cMovNroFin='" & psMovNro & "' where cPersCod='" & psCodPers & "' and cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & psCtaIFCod & "' and cPersCodPF='" & psCodPersPF & "'  and  cIFTpoPF='" & psIFTpoPF & "' and  cCtaIFCodPF='" & psCtaIfCodPF & "'"
    oCon.AbreConexion
    oCon.Ejecutar sSQL
    oCon.CierraConexion
    fbEliminarGarantia = 1
    Exit Function
mError:
    fbEliminarGarantia = 0
    Err.Raise vbObjectError + 100, "EliminarGarantiaRestringido", Err.Description
End Function




Public Function GetCtaIFMontos(ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIFCod As String) As ADODB.Recordset
Dim sSQL As String
    Dim rs As New ADODB.Recordset
    Dim oCon As New DConecta
  
    On Error GoTo GetCtaIFMontosErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = " Select nMontoPrestado,nSaldoCap" & _
           " From CtaIFAdeudados " & _
           " where cPersCod ='" & psPersCod & "' " & _
           " and cIFTpo ='" & psIFTpo & "' " & _
           " and cCtaIFCod='" & psCtaIFCod & "' "
    

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCtaIFMontos = rs
    Exit Function
GetCtaIFMontosErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function

Public Function GrabaAdeudoVinculado(ByVal psPersCodV As String, ByVal pnIFTpoV As String, ByVal psCtaIfCodV As String, _
                           ByVal pnTipoCambioEuro As Currency, ByVal RsDatos As ADODB.Recordset) As Integer

Dim nbEuros As Integer
Dim MontoEuros As Currency
Dim sSQL As String
Dim oCon As New DConecta
  
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

Public Function GetCtaIFVinculados(ByVal psPersCodV As String, ByVal psIFTpoV As String, ByVal psCtaIfCodV As String) As ADODB.Recordset
Dim sSQL As String
    Dim rs As New ADODB.Recordset
    Dim oCon As New DConecta
  
    On Error GoTo GetCtaIFMontosErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = " Select c.cIFTpo+'.'+c.cPersCod+'.'+c.cCtaIFCod Codigo," & _
           " ci.cCtaIFDesc,p.cPersNombre,c.nMontoPrestado,c.nSaldoCap, " & _
           " IsNull(Convert(int,c.bEuros),0) bEuros,Isnull(c.nMontoEuros,0)nMontoEuros " & _
           " From CtaIFAdeudados c JOIN CtaIF ci ON ci.cPersCod = c.cPersCod AND ci.cIFTpo= c.cIFTpo AND ci.cCtaIFCod=c.cCtaIFCod" & _
           " JOIN INSTITUCIONFINANC i ON i.cPersCod = ci.cPersCod" & _
           " JOIN PERSONA p ON p.cPersCod = ci.cPersCod" & _
           " where c.cPersCodVin ='" & psPersCodV & "' " & _
           " and c.cIFTpoVin ='" & psIFTpoV & "' " & _
           " and c.cCtaIFCodVin='" & psCtaIfCodV & "' "
    

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCtaIFVinculados = rs
    Exit Function
GetCtaIFMontosErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function

Public Function EliminarAdeudoVinculado(ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIFCod As String) As Boolean

Dim sSQL As String
    Dim oCon As New DConecta
  
    On Error GoTo mError
    
    sSQL = " Update CtaIFAdeudados set " & _
           " cPersCodVin='' ,  cIFTpoVin='' , cCtaIFCodVin='' " & _
           " where cPersCod='" & psPersCod & "' and cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & psCtaIFCod & "' "
           
    oCon.AbreConexion
    
    oCon.Ejecutar sSQL
    oCon.CierraConexion
    EliminarAdeudoVinculado = True
    Exit Function
mError:
    EliminarAdeudoVinculado = False
    Err.Raise vbObjectError + 100, "EliminarGarantiaRestringido", Err.Description
End Function

Public Function GetReporteAdeudadoVin(ByVal plnMoneda As Integer) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetReporteAdeudadoVinErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = " Select p.cPersNombre,c.nMontoPrestado, " & _
            "    Monto MontoRestringido, X.nMontoEuros MontoEuros, "
    sSQL = sSQL & " ISNULL(( Select datediff(day,( Select Max(dVencimiento) " & _
                "     FROM CtaIFCalendario  where cEstado = 1 " & _
                "     and cPErsCod = ci.cPersCod and cIFTpo =ci.cIFTpo " & _
                "     and cCtaIFCod =ci.cCtaIFCod), " & _
                "         (Select Min(dVencimiento)" & _
                "          FROM CtaIFCalendario where cEstado = 0 " & _
                "          and cPErsCod =ci.cPersCod and cIFTpo =ci.cIFTpo " & _
                "          and cCtaIFCod =ci.cCtaIFCod))),0) Plazo," & _
                "      ci.nInteres Interes,c.nSaldoCap SaldoCapital," & _
                "     p1.cPersNombre cPersNombreVin , " & _
                "     isnull(( Select datediff(day,( Select Max(dVencimiento) " & _
                "               FROM CtaIFCalendario  where cEstado = 1 " & _
                "               and cPErsCod = ci1.cPersCod and cIFTpo =ci1.cIFTpo " & _
                "               and cCtaIFCod =ci1.cCtaIFCod), " & _
                "               (Select Min(dVencimiento) " & _
                "                FROM CtaIFCalendario where cEstado = 0 " & _
                "                and cPErsCod =ci1.cPersCod and cIFTpo =ci1.cIFTpo " & _
                "                   and cCtaIFCod =ci1.cCtaIFCod))),0) Plazo, " & _
                " ci1.nInteres Interes"
    sSQL = sSQL & " FROM CtaIFAdeudados c  " & _
            " JOIN CtaIF ci ON ci.cPersCod =c.cPersCod and ci.cIFTpo =c.cIFTpo and ci.cCtaIFCod=c.cCtaIFCod" & _
            " JOIN Persona p ON p.cPersCod=c.cPersCod" & _
            " JOIN CtaIF ci1 ON ci1.cPersCod =c.cPersCodVin and ci1.cIFTpo =c.cIFTpoVin and ci1.cCtaIFCod=c.cCtaIFCodVin" & _
            " JOIN Persona p1 ON p1.cPersCod=c.cPersCodVin"
    sSQL = sSQL & " LEFT JOIN (select cia.cperscod,cia.ciftpo,cia.cctaifcod ,sum(nMovImporte)Monto ,sum(nmontoeuros)nMontoEuros " & _
            "                   from cgctaifadeudadosgarantia cia" & _
            "                   Inner join mov m on m.nMovNro=cia.nMovNro" & _
            "                   Inner join movCta mc on mc.nMovNro=m.nMovNro Inner join CtaIF ci on ci.cperscod=ciA.cperscodPF and ci.ciftpo=ciA.ciftpoPF and ci.cctaifcod=ciA.cctaifcodPF " & _
            "                   where " & _
            "                   m.nmovflag = 0 and cctacontcod like '11" & plnMoneda & "7%' and cmovnrofin is null" & _
            "                   group by cia.cperscod,cia.ciftpo,cia.cctaifcod  )" & _
            "                   X ON X.cperscod=ci1.cperscod and  case when len(X.ciftpo)=1 then '0'+x.ciftpo else x.ciftpo end=ci1.ciftpo and X.cctaifcod=ci1.cctaifcod " & _
            " where  substring(c.cCtaIFCod,3,1) = '" & plnMoneda & "'" & _
            " order by cPersNombreVin "

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetReporteAdeudadoVin = rs
    Exit Function
GetReporteAdeudadoVinErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function


Public Function CGAjusteAdeudadosEuros(ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psCtaIFCod As String, ByVal pnFactor As Double, ByVal psMovNro As String, pdFecha As Date, psAdeudoDesc As String, psOpeCod As String) As String
    Dim sSQL As String
    Dim rs As New ADODB.Recordset
    Dim oMov As New DMov
    Set oMov = New DMov
    On Error GoTo GetCtaIFMontosErr
    Set rs = New ADODB.Recordset

    Dim lsCuentaCorto As String
    Dim lsCuentaLargo As String
    Dim lsCtaInteres As String

    Dim lsPersCod As String
    Dim lsIFTpo As String
    Dim cTpoCuota As String
    Dim lnMovNro As String
    Dim item As Long
    Dim lsCtaIFCod As String
    Dim nNroCuota As Long
    
    Dim lnCapitalReal As Currency
    Dim lnInteresProvReal As Currency
    Dim lnInteresReal As Currency
    Dim lnComisionReal As Currency
    Dim lnSKAntReal As Currency
    Dim lnSKNewReal As Currency
    Dim lnSKLPAntReal As Currency
    Dim lnSKLPNewReal As Currency
    Dim dFecUltActSaldos As Date
    Dim ldFechaPago As Date
    Dim dFecUltPago As Date
    Dim cEstadoCuota As String
    'No se modificaran fechas de calculo de saldos
    
    lsPersCod = psPersCod
    lsIFTpo = psIFTpo
    lsCtaIFCod = psCtaIFCod
    'Solo cuotas tipo 2
    cTpoCuota = "2"
    
    lnCapitalReal = 0
    lnInteresProvReal = 0
    lnInteresReal = 0
    lnComisionReal = 0
    lnSKAntReal = 0
    lnSKNewReal = 0
    lnSKLPAntReal = 0
    lnSKLPNewReal = 0
    
    Dim oAdeud As New NCajaAdeudados
    Dim sCadena  As String
    sCadena = " AND CIF.cPersCod = '" & lsPersCod & "' AND CIF.cIFTpo = '" & lsIFTpo & "'  And CIF.cCtaIFCod = '" & lsCtaIFCod & "'"
    
    Set rs = oAdeud.GetDatosAdeudadoPago(pdFecha, "2", sCadena)

    dFecUltActSaldos = rs!dcuotaultmodsaldos
    dFecUltPago = rs!dCuotaUltPago

    nNroCuota = rs!nNroCuota
    lnSKAntReal = rs!nSaldoCap
    lnSKNewReal = Format(rs!nSaldoCap * pnFactor, "#.00")
    lnSKLPAntReal = rs!nSaldoCapLP
    lnSKLPNewReal = Format(rs!nSaldoCapLP * pnFactor, "#.00")
    lnCapitalReal = lnSKNewReal - lnSKAntReal
    
    'CP LP
    lsCuentaCorto = oAdeud.GetOpeCta(gOpeCGAdeudaPagoCuotaME, "D", "5", lsPersCod, lsIFTpo)
    lsCuentaLargo = "26" & Mid(lsCuentaCorto, 3, 50)
    lsCtaInteres = oAdeud.GetOpeCta(gOpeCGAdeudaPagoCuotaME, "D", "2", lsPersCod, lsIFTpo)
        
    oMov.BeginTrans
    
    oMov.InsertaMov psMovNro, psOpeCod, "ACTUALIZACION DE SALDO EN DOLARES DE ADEUDOS " & psAdeudoDesc
    lnMovNro = oMov.GetnMovNro(psMovNro)
    
    'actualiza saldos contables
    item = 0
    'Corto Plazo
    If (lnCapitalReal - (lnSKLPNewReal - lnSKLPAntReal)) <> 0 Then
            item = item + 1
        oMov.InsertaMovCta lnMovNro, item, lsCuentaCorto, (lnCapitalReal - (lnSKLPNewReal - lnSKLPAntReal)) * -1
        oMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
        oMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
    End If
    'Largo Plazo
    If (lnSKLPNewReal - lnSKLPAntReal) <> 0 Then
        item = item + 1
        oMov.InsertaMovCta lnMovNro, item, lsCuentaLargo, (lnSKLPNewReal - lnSKLPAntReal) * -1
        oMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
        oMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
    End If
    'Contra Cuenta
    item = item + 1
    oMov.InsertaMovCta lnMovNro, item, lsCtaInteres, lnCapitalReal
    oMov.InsertaMovObj lnMovNro, item, 1, ObjEntidadesFinancieras
    oMov.InsertaMovObjIF lnMovNro, item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
    
    'Graba Saldo Estadistico
    oMov.ActualizaDatosFAdeudado lsPersCod, lsIFTpo, lsCtaIFCod, False, pdFecha, False, pdFecha, True, lnSKNewReal, True, lnSKLPNewReal, , pdFecha, , False
        
    'Graba estadistica
    oMov.GrabaCtaIFEstadistica lsPersCod, lsIFTpo, lsCtaIFCod, cTpoCuota, nNroCuota, lnMovNro, 5, _
                               lnCapitalReal, lnInteresProvReal, lnInteresReal, lnComisionReal, 0, 0, _
                               lnSKAntReal, lnSKNewReal, _
                               lnSKLPAntReal, lnSKLPNewReal, 0, _
                               lnCapitalReal, lnInteresProvReal, lnInteresReal, lnComisionReal, 0, 0, _
                               lnSKAntReal, lnSKNewReal, lnSKLPAntReal, lnSKLPNewReal, _
                               dFecUltActSaldos, dFecUltActSaldos, dFecUltPago, dFecUltPago, "0"
    
    'Copia el BackUp de calendario
    sSQL = " Insert CtaIFcalendarioHist(cPerscod, cIFTpo, cCtaIFCod, cTpoCuota, nNroCuota, dVencimiento, cEstado, nDiasPago, nCapital, nInteres, nInteresPagado, nComision, cMovNro, bvigente, nInteresGracia, nTotalCuota, cUltimaActualizacion, nMovNro)" _
         & " Select cPersCod, cIFTpo, cCtaIFCod, cTpoCuota, nNroCuota, dVencimiento, cEstado, nDiasPago, nCapital, nInteres, nInteresPagado, nComision, cMovNro, bvigente, nInteresGracia, nTotalCuota, cUltimaActualizacion, " & lnMovNro & " From CtaIFcalendario " _
         & " Where cPerscod = '" & lsPersCod & "' And cIFTpo = '" & lsIFTpo & "'  And cCtaIFCod = '" & lsCtaIFCod & "' And bVigente = 1"
    oMov.Ejecutar sSQL
    
    'Actualiza Calendario
    sSQL = " Update CtaIFcalendario " _
         & " Set nCapital = Convert(Decimal(20,2),nCapital * " & pnFactor & ") , nInteres = Convert(Decimal(20,2),nInteres * " & pnFactor & "), nComision = Convert(Decimal(20,2),nComision  * " & pnFactor & "), nInteresGracia = Convert(Decimal(20,2), IsNull(nInteresGracia,0) * " & pnFactor & "), nTotalCuota = Convert(Decimal(20,2),nCapital * " & pnFactor & ") + Convert(Decimal(20,2),nInteres * " & pnFactor & ") + Convert(Decimal(20,2),nComision  * " & pnFactor & ") + Convert(Decimal(20,2), IsNull(nInteresGracia,0) * " & pnFactor & "), cUltimaActualizacion = '" & psMovNro & "' From CtaIFcalendario " _
         & " Where cPerscod = '" & lsPersCod & "' And cIFTpo = '" & lsIFTpo & "'  And cCtaIFCod = '" & lsCtaIFCod & "' And bVigente = 1 And cEstado In (0,2)"
    oMov.Ejecutar sSQL
    
    oMov.GeneraMovME lnMovNro, psMovNro
    
    oMov.CommitTrans
    
    Exit Function
GetCtaIFMontosErr:
    oMov.RollbackTrans
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function


Public Function GetCreditosVigentes(ByVal psFecha As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

'    sSQL = " SELECT SUM( " & _
'        " CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' " & _
'        "    AND nDiasAtraso < 8 THEN " & _
'        " C.nSaldoCap " & _
'        " Else " & _
'        " (   CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc)<8 THEN " & _
'        "    ( CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) <0 " & _
'        "        AND SUBSTRING(C.cCtaCod,6,1)IN('3','4') " & _
'        "        AND C.nDiasAtraso BETWEEN 31 AND 90 THEN 0 " & _
'        "        Else P.nCapital -P.nCapPag " & _
'        "        End  )  ELSE 0 END ) END ) " & _
'        " as Venc07, " & _
'        " SUM( " & _
'        " CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso  BETWEEN 8 AND 15 THEN " & _
'        " C.nSaldoCap " & _
'        " Else " & _
'        " (   CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 8 AND 15 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END " & _
'        "    ) " & _
'        "     END ) " & _
'        " as Venc08_15,"
'
'
'    sSQL = sSQL & " SUM(CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' " & _
'        "   AND nDiasAtraso BETWEEN 16 AND 30 THEN " & _
'        " C.nSaldoCap " & _
'        " Else " & _
'        " ( " & _
'        "    CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 16 AND 30 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END " & _
'        "    ) " & _
'        "     END ) " & _
'        " as Venc016_30, " & _
'        " SUM( " & _
'        " CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305'AND  nDiasAtraso BETWEEN 31 AND 60 " & _
'            " THEN 0 "
'
'
'
'    sSQL = sSQL & " ELSE " & _
'        " ( CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 31 AND 60 THEN " & _
'        "          P.nCapital-P.nCapPag " & _
'        "    ELSE 0 END " & _
'        "     ) END ) as Venc031_60, " & _
'        " SUM( CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 61 AND 90 " & _
'        "    THEN 0 " & _
'        " Else " & _
'        " (   CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 61 AND 90 THEN " & _
'        "        P.nCapital-P.nCapPag " & _
'        "     ELSE 0 END ) END ) as Venc061_90, " & _
'        " SUM( CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 91 AND 180 " & _
'        "    THEN 0 " & _
'        " Else " & _
'        " (       CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 91 AND 180 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END  )  END )  as Venc091_180, "
'
'
'
'    sSQL = sSQL & " SUM( " & _
'        " CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 181 AND 360 " & _
'        "    THEN 0 " & _
'        " Else " & _
'        " (CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 181 AND 360 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END  ) END ) " & _
'        " as Venc0181_360, " & _
'        " SUM( CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 361 AND 720 " & _
'        "    THEN   0 " & _
'        " Else " & _
'        " ( CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 361 AND 720 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END )   END ) as Venc02_A, " & _
'        " SUM( CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 721 AND 1080 " & _
'        "    THEN  0 " & _
'        " Else " & _
'        " ( CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 721 AND 1080 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END ) END )as Venc03_A,"
'
'
'     sSQL = sSQL & "SUM(CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 1081 AND 1440 " & _
'        "    THEN 0 " & _
'        " Else " & _
'        " ( CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 1081 AND 1440 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END  ) END ) as Venc04_A, " & _
'        " SUM( CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 1441 AND 1800 " & _
'        "    THEN  0 " & _
'        " Else " & _
'        " (       CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 1441 AND 1800 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END " & _
'        "    )  END ) as Venc05_A," & _
'        " SUM(CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 1801 AND 3600 " & _
'        "    THEN 0 " & _
'        " Else ( CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 1801 AND 3600 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END )  END )as Venc06_10A, "
'
'
'     sSQL = sSQL & " SUM(CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 3601 AND 7200 " & _
'        "    THEN 0 " & _
'        "Else " & _
'        " (       CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 3601 AND 7200 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END ) END )  as Venc11_20A, " & _
'        " SUM(CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso >=7201 " & _
'        "     THEN  0 " & _
'        " Else " & _
'        " ( CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) >=7201 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END ) END ) as Venc21_Amas ,SUBSTRING(C.cCtaCod,9,1)"
'
'     sSQL = sSQL & "FROM DBConsolidada..CreditoConsol C LEFT JOIN DBConsolidada..PlanDesPagConsol P ON C.cCtaCod = P.cCtaCod " & _
'        " AND nTipo= 1 AND nEstado=0 " & _
'        " WHERE nPrdEstado IN (2020,2021,2022,2101,2104,2106,2107) " & _
'        " AND (   (SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso<=30) " & _
'        " OR  " & _
'        " (   SUBSTRING(C.cCtaCod,6,3)<>'305' " & _
'        "    AND " & _
'        "    (   (SUBSTRING(C.cCtaCod,6,1)='1' AND nDiasAtraso <= 15 ) " & _
'        "        OR (SUBSTRING(C.cCtaCod,6,1)='2' AND nDiasAtraso <= 30 ) " & _
'        "        OR (SUBSTRING(C.cCtaCod,6,1)IN('3','4')AND nDiasAtraso <= 90 ) " & _
'        "    )  ) ) " & _
'        " GROUP BY SUBSTRING(C.cCtaCod,9,1) " & _
'        " ORDER BY SUBSTRING(C.cCtaCod,9,1) "
    sSQL = "exec B2_stp_sel_creditosVigentes '" & psFecha & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCreditosVigentes = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function


'Public Function GetCreditosRefinanciados(ByVal psFecha As String) As ADODB.Recordset
'Dim sSQL As String
'Dim rs As New ADODB.Recordset
'Dim oCon As New DConecta
'
'On Error GoTo GetCreditosVigentesErr
'    oCon.AbreConexion
'    Set rs = New ADODB.Recordset
'
'    sSQL = " SELECT SUM(CASE WHEN DATEDIFF(d,'" & psFecha & "',C.dFecVenc)<8 THEN " & _
'           " (CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1' AND nDiasAtraso <= 15 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)='2' AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN('3','4')AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN ('3','4') AND nDiasAtraso BETWEEN 31 AND 90 " & _
'        "    THEN nSaldoCap-ISNULL(nCapVencido,0) END " & _
'        " ) " & _
'        " ELSE 0 END) as Venc07, " & _
'        " SUM(CASE WHEN DATEDIFF(d,'" & psFecha & "',C.dFecVenc) BETWEEN 8 AND 15 THEN " & _
'        " (CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1' AND nDiasAtraso <= 15 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)='2' AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN('3','4')AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN ('3','4') AND nDiasAtraso BETWEEN 31 AND 90 " & _
'        "    THEN nSaldoCap-ISNULL(nCapVencido,0) END " & _
'        " ) " & _
'        " ELSE 0 END) as Venc08_15, " & _
'        " SUM(CASE WHEN DATEDIFF(d,'" & psFecha & "',C.dFecVenc) BETWEEN 16 AND 30 THEN " & _
'        "    (CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1' AND nDiasAtraso <= 15 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)='2' AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN('3','4')AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN ('3','4') AND nDiasAtraso BETWEEN 31 AND 90 " & _
'        "    THEN nSaldoCap-ISNULL(nCapVencido,0) END " & _
'        " ) " & _
'        " ELSE 0 END) as Venc016_30,"
'    sSQL = sSQL & "SUM(CASE WHEN DATEDIFF(d,'" & psFecha & "',C.dFecVenc) BETWEEN 31 AND 60 THEN " & _
'            " (CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1' AND nDiasAtraso <= 15 THEN nSaldoCap " & _
'            " WHEN SUBSTRING(C.cCtaCod,6,1)='2' AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'            " WHEN SUBSTRING(C.cCtaCod,6,1)IN('3','4')AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'            " WHEN SUBSTRING(C.cCtaCod,6,1)IN ('3','4') AND nDiasAtraso BETWEEN 31 AND 90 " & _
'            "THEN nSaldoCap-ISNULL(nCapVencido,0) END " & _
'        " ) " & _
'        " ELSE 0 END) as Venc031_60, " & _
'        " SUM(CASE WHEN DATEDIFF(d,'" & psFecha & "',C.dFecVenc) BETWEEN 61 AND 90 THEN " & _
'        " (CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1' AND nDiasAtraso <= 15 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)='2' AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN('3','4')AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN ('3','4') AND nDiasAtraso BETWEEN 31 AND 90 " & _
'        "    THEN nSaldoCap-ISNULL(nCapVencido,0) END " & _
'        " ) " & _
'        " ELSE 0 END) as Venc061_90, " & _
'        " SUM(CASE WHEN DATEDIFF(d,'" & psFecha & "',C.dFecVenc) BETWEEN 91 AND 180 THEN " & _
'        " (CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1' AND nDiasAtraso <= 15 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)='2' AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN('3','4')AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN ('3','4') AND nDiasAtraso BETWEEN 31 AND 90 " & _
'        "    THEN nSaldoCap-ISNULL(nCapVencido,0) END " & _
'        " ) "
'
'    sSQL = sSQL & "ELSE 0 END) as Venc091_180," & _
'        " SUM(CASE WHEN DATEDIFF(d,'" & psFecha & "',C.dFecVenc) BETWEEN 181 AND 360 THEN " & _
'        " (CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1' AND nDiasAtraso <= 15 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)='2' AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN('3','4')AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN ('3','4') AND nDiasAtraso BETWEEN 31 AND 90 " & _
'        "    THEN nSaldoCap-ISNULL(nCapVencido,0) END " & _
'        " ) " & _
'        " ELSE 0 END) as Venc0181_360, "
'    sSQL = sSQL & " SUM(CASE WHEN DATEDIFF(d,'" & psFecha & "',C.dFecVenc) BETWEEN 361 AND 720 THEN " & _
'        " (CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1' AND nDiasAtraso <= 15 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)='2' AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN('3','4')AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN ('3','4') AND nDiasAtraso BETWEEN 31 AND 90 " & _
'        "    THEN nSaldoCap-ISNULL(nCapVencido,0) END " & _
'        " ) " & _
'        " ELSE 0 END) as Venc02_A, " & _
'        " SUM(CASE WHEN DATEDIFF(d,'" & psFecha & "',C.dFecVenc) BETWEEN 721 AND 1080 THEN " & _
'        " (CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1' AND nDiasAtraso <= 15 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)='2' AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN('3','4')AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN ('3','4') AND nDiasAtraso BETWEEN 31 AND 90 " & _
'        "    THEN nSaldoCap-ISNULL(nCapVencido,0) END " & _
'        ") " & _
'        " ELSE 0 END) as Venc03_A, " & _
'        " SUM(CASE WHEN DATEDIFF(d,'" & psFecha & "',C.dFecVenc) BETWEEN 1081 AND 1440 THEN " & _
'        " (CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1' AND nDiasAtraso <= 15 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)='2' AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN('3','4')AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN ('3','4') AND nDiasAtraso BETWEEN 31 AND 90 " & _
'        "    THEN nSaldoCap-ISNULL(nCapVencido,0) END " & _
'        " ) " & _
'        " ELSE 0 END) as Venc04_A,"
'    sSQL = sSQL & " SUM(CASE WHEN DATEDIFF(d,'" & psFecha & "',C.dFecVenc) BETWEEN 1441 AND 1800 THEN " & _
'        "(CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1' AND nDiasAtraso <= 15 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)='2' AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN('3','4')AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN ('3','4') AND nDiasAtraso BETWEEN 31 AND 90 " & _
'        "    THEN nSaldoCap-ISNULL(nCapVencido,0) END " & _
'        " ) " & _
'        " ELSE 0 END) as Venc05_A, " & _
'        " SUM(CASE WHEN DATEDIFF(d,'" & psFecha & "',C.dFecVenc) BETWEEN 1801 AND 3600 THEN " & _
'        " (CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1' AND nDiasAtraso <= 15 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)='2' AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN('3','4')AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN ('3','4') AND nDiasAtraso BETWEEN 31 AND 90 " & _
'        "    THEN nSaldoCap-ISNULL(nCapVencido,0) END " & _
'        " ) " & _
'        " ELSE 0 END) as Venc06_10A, " & _
'        " SUM(CASE WHEN DATEDIFF(d,'" & psFecha & "',C.dFecVenc) BETWEEN 3601 AND 7200 THEN " & _
'        " (CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1' AND nDiasAtraso <= 15 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)='2' AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN('3','4')AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "    WHEN SUBSTRING(C.cCtaCod,6,1)IN ('3','4') AND nDiasAtraso BETWEEN 31 AND 90 " & _
'        "    THEN nSaldoCap-ISNULL(nCapVencido,0) END " & _
'        " ) " & _
'        " ELSE 0 END) as Venc11_20A,"
'     sSQL = sSQL & " SUM(CASE WHEN DATEDIFF(d,'" & psFecha & "',C.dFecVenc) >=7201 THEN " & _
'        "    (CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1' AND nDiasAtraso <= 15 THEN nSaldoCap " & _
'        "        WHEN SUBSTRING(C.cCtaCod,6,1)='2' AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "        WHEN SUBSTRING(C.cCtaCod,6,1)IN('3','4')AND nDiasAtraso <= 30 THEN nSaldoCap " & _
'        "        WHEN SUBSTRING(C.cCtaCod,6,1)IN ('3','4') AND nDiasAtraso BETWEEN 31 AND 90 " & _
'        "        THEN nSaldoCap-ISNULL(nCapVencido,0) END " & _
'        "    ) " & _
'        "    ELSE 0 END) as Venc21_Amas ,SUBSTRING(cCtaCod,9,1) " & _
'        " FROM DBConsolidada..CreditoConsol C " & _
'        " WHERE nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) AND cRefinan='R' " & _
'        " GROUP BY SUBSTRING(cCtaCod,9,1)"
'
'    oCon.AbreConexion
'    Set rs = oCon.CargaRecordSet(sSQL)
'    oCon.CierraConexion
'    Set GetCreditosRefinanciados = rs
'    Exit Function
'GetCreditosVigentesErr:
'    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
'End Function

''MIOL 20120430, SEGUN RQ12049 ***************************************************************
'Public Function GetCreditosVigentesN(ByVal psFecha As String) As ADODB.Recordset
'Dim sSQL As String
'Dim rs As New ADODB.Recordset
'Dim oCon As New DConecta
'
'On Error GoTo GetCreditosVigentesErr
'    oCon.AbreConexion
'    Set rs = New ADODB.Recordset
'
'    sSQL = "exec B2_stp_sel_creditosVigentesN '" & psFecha & "'"
'    oCon.AbreConexion
'    Set rs = oCon.CargaRecordSet(sSQL)
'    oCon.CierraConexion
'    Set GetCreditosVigentes = rs
'    Exit Function
'GetCreditosVigentesErr:
'    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
'End Function
''END MIOL ***********************************************************************************

Public Function GetCreditosRefinanciados(ByVal psFecha As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    
'    sSQL = " SELECT SUM( " & _
'        " CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' " & _
'        "    AND nDiasAtraso < 8 THEN " & _
'        " C.nSaldoCap " & _
'        " Else " & _
'        " (   CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc)<8 THEN " & _
'        "    ( CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) <0 " & _
'        "        AND SUBSTRING(C.cCtaCod,6,1)IN('3','4') " & _
'        "        AND C.nDiasAtraso BETWEEN 31 AND 90 THEN 0 " & _
'        "        Else P.nCapital -P.nCapPag " & _
'        "        End  )  ELSE 0 END ) END ) " & _
'        " as Venc07, " & _
'        " SUM( " & _
'        " CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso  BETWEEN 8 AND 15 THEN " & _
'        " C.nSaldoCap " & _
'        " Else " & _
'        " (   CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 8 AND 15 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END " & _
'        "    ) " & _
'        "     END ) " & _
'        " as Venc08_15,"
'
'    sSQL = sSQL & " SUM(CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' " & _
'        "   AND nDiasAtraso BETWEEN 16 AND 30 THEN " & _
'        " C.nSaldoCap " & _
'        " Else " & _
'        " ( " & _
'        "    CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 16 AND 30 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END " & _
'        "    ) " & _
'        "     END ) " & _
'        " as Venc016_30, " & _
'        " SUM( " & _
'        " CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305'AND  nDiasAtraso BETWEEN 31 AND 60 " & _
'            " THEN 0 "
'
'    sSQL = sSQL & " ELSE " & _
'        " ( CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 31 AND 60 THEN " & _
'        "          P.nCapital-P.nCapPag " & _
'        "    ELSE 0 END " & _
'        "     ) END ) as Venc031_60, " & _
'        " SUM( CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 61 AND 90 " & _
'        "    THEN 0 " & _
'        " Else " & _
'        " (   CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 61 AND 90 THEN " & _
'        "        P.nCapital-P.nCapPag " & _
'        "     ELSE 0 END ) END ) as Venc061_90, " & _
'        " SUM( CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 91 AND 180 " & _
'        "    THEN 0 " & _
'        " Else " & _
'        " (       CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 91 AND 180 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END  )  END )  as Venc091_180, "
'
'    sSQL = sSQL & " SUM( " & _
'        " CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 181 AND 360 " & _
'        "    THEN 0 " & _
'        " Else " & _
'        " (CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 181 AND 360 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END  ) END ) " & _
'        " as Venc0181_360, " & _
'        " SUM( CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 361 AND 720 " & _
'        "    THEN   0 " & _
'        " Else " & _
'        " ( CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 361 AND 720 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END )   END ) as Venc02_A, " & _
'        " SUM( CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 721 AND 1080 " & _
'        "    THEN  0 " & _
'        " Else " & _
'        " ( CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 721 AND 1080 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END ) END )as Venc03_A,"
'
'     sSQL = sSQL & "SUM(CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 1081 AND 1440 " & _
'        "    THEN 0 " & _
'        " Else " & _
'        " ( CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 1081 AND 1440 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END  ) END ) as Venc04_A, " & _
'        " SUM( CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 1441 AND 1800 " & _
'        "    THEN  0 " & _
'        " Else " & _
'        " (       CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 1441 AND 1800 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END " & _
'        "    )  END ) as Venc05_A," & _
'        " SUM(CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 1801 AND 3600 " & _
'        "    THEN 0 " & _
'        " Else ( CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 1801 AND 3600 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END )  END )as Venc06_10A, "
'
'     sSQL = sSQL & " SUM(CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso BETWEEN 3601 AND 7200 " & _
'        "    THEN 0 " & _
'        "Else " & _
'        " (       CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) BETWEEN 3601 AND 7200 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END ) END )  as Venc11_20A, " & _
'        " SUM(CASE WHEN SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso >=7201 " & _
'        "     THEN  0 " & _
'        " Else " & _
'        " ( CASE WHEN DATEDIFF(d,' " & psFecha & " ',P.dFecVenc) >=7201 THEN " & _
'        "    P.nCapital -P.nCapPag " & _
'        "    ELSE 0 END ) END ) as Venc21_Amas ,SUBSTRING(C.cCtaCod,9,1)"
'
'     sSQL = sSQL & "FROM DBConsolidada..CreditoConsol C LEFT JOIN DBConsolidada..PlanDesPagConsol P ON C.cCtaCod = P.cCtaCod " & _
'        " AND nTipo= 1 AND nEstado=0 " & _
'        " WHERE nPrdEstado IN (2030,2031,2032) " & _
'        " AND (   (SUBSTRING(C.cCtaCod,6,3)='305' AND nDiasAtraso<=30) " & _
'        " OR  " & _
'        " (   SUBSTRING(C.cCtaCod,6,3)<>'305' " & _
'        "    AND " & _
'        "    (   (SUBSTRING(C.cCtaCod,6,1)='1' AND nDiasAtraso <= 15 ) " & _
'        "        OR (SUBSTRING(C.cCtaCod,6,1)='2' AND nDiasAtraso <= 30 ) " & _
'        "        OR (SUBSTRING(C.cCtaCod,6,1)IN('3','4')AND nDiasAtraso <= 90 ) " & _
'        "    )  ) ) " & _
'        " GROUP BY SUBSTRING(C.cCtaCod,9,1) " & _
'        " ORDER BY SUBSTRING(C.cCtaCod,9,1) "
   
    sSQL = "exec B2_stp_sel_creditosRefinanciados '" & psFecha & "'"
    
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCreditosRefinanciados = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function


Public Function GetCarteraCreditos(ByVal psFecha As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset
    
    '*****BRGO BASILEA II
    sSQL = "SELECT cAgeDescripcion, " & _
    "    cEstado = CASE WHEN cEstado IN ('4') THEN 'REFINANCIADO' " & _
    "                WHEN cEstado ='1'THEN 'VIGENTE' " & _
    "                WHEN cEstado='5'THEN 'VENCIDO' " & _
    "               WHEN cEstado='6'THEN 'JUDICIAL' " & _
    "            END, " & _
    "    cMoneda= CASE WHEN cMoneda='1'THEN /*'SOLES'*/ (select cConsDescripcion from Constante where nConsValor = 102 and nConsCod = 10087) ELSE'DOLARES'END, " & _
    "    nSaldo = Sum(nSaldo) " & _
    " FROM DBConsolidada..CredSaldosCont C " & _
    " INNER JOIN DBCmacMaynas..Agencias A ON A.cAgeCod = C.cAgeCod " & _
    " WHERE cConcepto='C' AND DATEDIFF(d,dFecha,'" & psFecha & "')=0 " & _
    " GROUP BY A.cAgeCod,cAgeDescripcion, " & _
    "        CASE WHEN cEstado IN ('4') THEN 'REFINANCIADO' " & _
    "                WHEN cEstado ='1'THEN 'VIGENTE' " & _
    "                WHEN cEstado='5'THEN 'VENCIDO' " & _
    "                WHEN cEstado='6'THEN 'JUDICIAL' " & _
    "            END, cMoneda " & _
    " ORDER BY A.cAgeCod "
    '******************
    
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCarteraCreditos = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function


Public Function GetInteresCreditos(ByVal psFecha As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    'by capi 09122008
'    sSQL = " SELECT cAgeDescripcion,cConsDescripcion," & _
'            " cEstado = CASE WHEN cRefinanc = 'R' and cEstado IN ('1','3') THEN 'REFINANCIADO' " & _
'            "        WHEN cEstado ='1'THEN 'VIGENTE' " & _
'            "        WHEN cEstado='5'THEN 'VENCIDO' " & _
'            "        WHEN cEstado='6'THEN 'JUDICIAL' " & _
'            "    END, " & _
'            " cMoneda= CASE WHEN cMoneda='1'THEN'SOLES'ELSE'DOLARES'END, " & _
'            " nSaldo = Sum(nSaldo) " & _
'            " FROM DBConsolidada..CredSaldosCont C " & _
'            " INNER JOIN DBCmacMaynas..Constante P ON C.cProdCod = P.nConsValor AND nConsCod= 1001 " & _
'            " INNER JOIN DBCmacMaynas..Agencias A ON A.cAgeCod = C.cAgeCod " & _
'            " WHERE cConcepto='I' AND DATEDIFF(d,dFecha,'" & psFecha & "')=0 " & _
'            " GROUP BY A.cAgeCod,cAgeDescripcion,cConsDescripcion, " & _
'            "   CASE WHEN cRefinanc = 'R' and cEstado IN ('1','3') THEN 'REFINANCIADO' " & _
'            "        WHEN cEstado ='1'THEN 'VIGENTE' " & _
'            "        WHEN cEstado='5'THEN 'VENCIDO' " & _
'            "        WHEN cEstado='6'THEN 'JUDICIAL' " & _
'            "    END, cMoneda " & _
'            " ORDER BY A.cAgeCod"
    
     sSQL = " SELECT cAgeDescripcion,cConsDescripcion," & _
            " cEstado = CASE WHEN cRefinanc = 'R' and cEstado IN ('4') THEN 'REFINANCIADO' " & _
            "        WHEN cEstado ='1' THEN 'VIGENTE' " & _
            "        WHEN cEstado='5'THEN 'VENCIDO' " & _
            "        WHEN cEstado='6'THEN 'JUDICIAL' " & _
            "        WHEN cEstado='7'THEN 'ALINEACION' " & _
            "    END, " & _
            " cMoneda= CASE WHEN cMoneda='1'THEN'SOLES'ELSE'DOLARES'END, " & _
            " nSaldo = Sum(nSaldo), CASE WHEN cEstado ='1' and cCalGen in ('0','1','2') and cRefinanc = 'N' Then 'ID' else 'IS' End TI" & _
            " FROM DBConsolidada..CredSaldosCont C " & _
            " INNER JOIN DBCmacMaynas..Constante P ON C.cProdCod = P.nConsValor AND nConsCod= 3034 " & _
            " INNER JOIN DBCmacMaynas..Agencias A ON A.cAgeCod = C.cAgeCod " & _
            " WHERE cConcepto='I' AND DATEDIFF(d,dFecha,'" & psFecha & "')=0 AND RIGHT(cProdCod,2) != '81' " & _
            " GROUP BY A.cAgeCod,cAgeDescripcion,cConsDescripcion, " & _
            "   CASE WHEN cRefinanc = 'R' and cEstado IN ('4') THEN 'REFINANCIADO' " & _
            "        WHEN cEstado ='1'THEN 'VIGENTE' " & _
            "        WHEN cEstado='5'THEN 'VENCIDO' " & _
            "        WHEN cEstado='6'THEN 'JUDICIAL' " & _
            "        WHEN cEstado='7'THEN 'ALINEACION' " & _
            "    END, cMoneda,CASE WHEN cEstado ='1' and cCalGen in ('0','1','2') and cRefinanc = 'N' Then 'ID' else 'IS' End " & _
            " ORDER BY A.cAgeCod"
  'End by

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetInteresCreditos = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function



Public Function GetCreditosCastigados(ByVal psFecha As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

'    sSQL = " SELECT cAgeDescripcion Descripcion," & _
'            " CASE WHEN SUBSTRING(cCtaCod,9,1)='1'THEN 'SOLES'ELSE'DOLARES'END Moneda,SUM(nSaldoCap)Saldo,COUNT(*)Cantidad " & _
'            " FROM DBConsolidada..CreditoConsol INNER JOIN DBCmacMaynas..Agencias A ON A.cAgeCod = SUBSTRING(cCtaCod,4,2) " & _
'            " WHERE nPrdEstado IN(2202,2206) " & _
'            " GROUP BY cAgeCod,cAgeDescripcion,SUBSTRING(cCtaCod,9,1) " & _
'            " ORDER BY cAgeCod"

'----- By BRGO 20100611 BASILEA II
   sSQL = " SELECT cAgeDescripcion Descripcion," & _
            " CASE WHEN SUBSTRING(CC.cCtaCod,9,1)='1'THEN 'SOLES'ELSE'DOLARES'END Moneda, " & _
            " SUM(CC.nSaldoCap)Saldo,COUNT(*)Cantidad " & _
            " FROM DBConsolidada..CreditoConsol CC " & _
            "    INNER JOIN DBConsolidada..CreditoConsolTotal C ON C.cCtaCod = CC.cCtaCod" & _
            "    INNER JOIN DBCmacMaynas..Agencias A ON A.cAgeCod = C.cAgeCodAct " & _
            " WHERE CC.nPrdEstado IN(2202,2206) " & _
            " GROUP BY A.cAgeCod,A.cAgeDescripcion,SUBSTRING(CC.cCtaCod,9,1)" & _
            " ORDER BY A.cAgeCod "
    '---- End BRGO
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCreditosCastigados = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function

'Reportes Planeamiento
'JEOM
Public Function GetCarteraCreditosPla(ByVal psFecha As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

'    sSQL = " SELECT cAgeDescripcion Agencia," & _
'           "CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1'THEN 'COMERCIAL' " & _
'           " WHEN SUBSTRING(C.cCtaCod,6,1)='2'THEN 'MICROEMPRESA' " & _
'           " WHEN SUBSTRING(C.cCtaCod,6,3)='305'THEN 'PRENDARIO' " & _
'           " WHEN SUBSTRING(C.cCtaCod,6,1)='3'THEN 'CONSUMO' " & _
'           " WHEN SUBSTRING(C.cCtaCod,6,1)='4'THEN 'HIPOTECARIO' " & _
'           " END Tipo, " & _
'           " CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1'THEN 'SOLES' ELSE 'DOLARES'END Moneda, " & _
'           " COUNT(*)Cantidad, " & _
'           " SUM(nSaldoCap)Saldo " & _
'           " FROM DBConsolidada..CreditoConsol C " & _
'           " INNER JOIN Agencias A ON SUBSTRING(C.cCtaCod,6,3) = A.cAgeCod " & _
'           " WHERE nPrdEstado IN(2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107) " & _
'           " GROUP BY cAgeCod,cAgeDescripcion, "
'    sSQL = sSQL & "CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1'THEN 'COMERCIAL' " & _
'           " WHEN SUBSTRING(C.cCtaCod,6,1)='2'THEN 'MICROEMPRESA' " & _
'           " WHEN SUBSTRING(C.cCtaCod,6,3)='305'THEN 'PRENDARIO' " & _
'           " WHEN SUBSTRING(C.cCtaCod,6,1)='3'THEN 'CONSUMO' " & _
'           " WHEN SUBSTRING(C.cCtaCod,6,1)='4'THEN 'HIPOTECARIO' " & _
'           " END,SUBSTRING(C.cCtaCod,9,1) " & _
'           " ORDER BY cAgeCod,SUBSTRING(C.cCtaCod,9,1), " & _
'           " CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1'THEN 'COMERCIAL' " & _
'           " WHEN SUBSTRING(C.cCtaCod,6,1)='2'THEN 'MICROEMPRESA' " & _
'           " WHEN SUBSTRING(C.cCtaCod,6,3)='305'THEN 'PRENDARIO' " & _
'           " WHEN SUBSTRING(C.cCtaCod,6,1)='3'THEN 'CONSUMO' " & _
'           " WHEN SUBSTRING(C.cCtaCod,6,1)='4'THEN 'HIPOTECARIO'" & _
'           " END"

'******* BRGO BASILEA II
    sSQL = " SELECT cAgeDescripcion Agencia," & _
           " CASE WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='1'THEN 'CORPORATIVO' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='2' THEN 'GRANDES EMPRESAS' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='3' THEN 'MEDIANAS EMPRESAS' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='4' THEN 'PEQUEÑAS EMPRESAS' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='5' THEN 'MICROEMPRESAS' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='6' THEN 'CONSUMO REVOLVENTE' " & _
           " WHEN CCT.cTpoCredCod = '755'THEN 'PRENDARIO' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='7' THEN 'CONSUMO NO REVOLVENTE' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='8' THEN 'HIPOTECARIOS' " & _
           " END Tipo, " & _
           " CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1'THEN /*'SOLES'*/ (select cConsDescripcion from Constante where nConsValor = 102 and nConsCod = 10087) ELSE 'DOLARES'END Moneda, " & _
           " COUNT(*)Cantidad, " & _
           " SUM(C.nSaldoCap)Saldo " & _
           " FROM DBConsolidada..CreditoConsol C " & _
           " INNER JOIN DBConsolidada..CreditoConsolTotal CCT ON C.cCtaCod = CCT.cCtaCod " & _
           " INNER JOIN Agencias A ON CCT.cAgeCodAct = A.cAgeCod " & _
           " WHERE C.nPrdEstado IN(2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107) " & _
           " GROUP BY cAgeCod,cAgeDescripcion, "
    sSQL = sSQL & " CASE WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='1'THEN 'CORPORATIVO' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='2' THEN 'GRANDES EMPRESAS' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='3' THEN 'MEDIANAS EMPRESAS' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='4' THEN 'PEQUEÑAS EMPRESAS' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='5' THEN 'MICROEMPRESAS' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='6' THEN 'CONSUMO REVOLVENTE' " & _
           " WHEN CCT.cTpoCredCod = '755'THEN 'PRENDARIO' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='7' THEN 'CONSUMO NO REVOLVENTE' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='8' THEN 'HIPOTECARIOS' " & _
           " END,SUBSTRING(C.cCtaCod,9,1) " & _
           " ORDER BY cAgeCod,SUBSTRING(C.cCtaCod,9,1), " & _
           " CASE WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='1'THEN 'CORPORATIVO' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='2' THEN 'GRANDES EMPRESAS' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='3' THEN 'MEDIANAS EMPRESAS' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='4' THEN 'PEQUEÑAS EMPRESAS' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='5' THEN 'MICROEMPRESAS' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='6' THEN 'CONSUMO REVOLVENTE' " & _
           " WHEN CCT.cTpoCredCod = '755' THEN 'PRENDARIO' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='7' THEN 'CONSUMO NO REVOLVENTE' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='8' THEN 'HIPOTECARIOS' " & _
           " END"
'***********************************

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCarteraCreditosPla = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function


Public Function GetCredDesembolsosPla(ByVal psFecha As String, psTipoCambio As Currency) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
Dim lsAnio As String
Dim lsMes As String

  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset
    
    lsAnio = Mid(psFecha, 1, 4)
    lsMes = Mid(psFecha, 5, 2)

'    sSQL = " SELECT cAgeDescripcion Agencia," & _
'            " CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1'THEN 'COMERCIAL' " & _
'            " WHEN SUBSTRING(C.cCtaCod,6,1)='2'THEN 'MICROEMPRESA' " & _
'            " WHEN SUBSTRING(C.cCtaCod,6,3)='305'THEN 'PRENDARIO' " & _
'            " WHEN SUBSTRING(C.cCtaCod,6,1)='3'THEN 'CONSUMO' " & _
'            " WHEN SUBSTRING(C.cCtaCod,6,1)='4'THEN 'HIPOTECARIO' " & _
'            " END Tipo, " & _
'            " Descripcion=ISNULL(cConsDescripcion,'PRENDARIO') , " & _
'            " COUNT(*)Cantidad, " & _
'            " SUM(CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1'THEN nMontoDesemb ELSE nMontoDesemb * " & psTipoCambio & " END )Monto " & _
'            " FROM DBConsolidada..CreditoConsol C " & _
'            " INNER JOIN Agencias A ON SUBSTRING(C.cCtaCod,4,2)=A.cAgeCod " & _
'            " LEFT JOIN Constante ON nConsValor = nCondCre AND nConsCod= 3015 " & _
'            " Where Year(dFecVig) = '" & lsAnio & "' And Month(dFecVig) = '" & lsMes & "' " & _
'            " GROUP BY cAgeCod,cAgeDescripcion, " & _
'            " CASE WHEN SUBSTRING(C.cCtaCod,6,1)='1'THEN 'COMERCIAL' " & _
'            " WHEN SUBSTRING(C.cCtaCod,6,1)='2'THEN 'MICROEMPRESA' " & _
'            " WHEN SUBSTRING(C.cCtaCod,6,3)='305'THEN 'PRENDARIO' " & _
'            " WHEN SUBSTRING(C.cCtaCod,6,1)='3'THEN 'CONSUMO' " & _
'            " WHEN SUBSTRING(C.cCtaCod,6,1)='4'THEN 'HIPOTECARIO' " & _
'            " END,nCondCre,cConsDescripcion " & _
'            " ORDER BY cAgeCod,nCondCre"

'******* BRGO BASILEA II
    sSQL = " SELECT cAgeDescripcion Agencia," & _
           " CASE WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='1'THEN 'CORPORATIVO' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='2' THEN 'GRANDES EMPRESAS' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='3' THEN 'MEDIANAS EMPRESAS' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='4' THEN 'PEQUEÑAS EMPRESAS' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='5' THEN 'MICROEMPRESAS' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='6' THEN 'CONSUMO REVOLVENTE' " & _
           " WHEN CCT.cTpoCredCod = '755'THEN 'PRENDARIO' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='7' THEN 'CONSUMO NO REVOLVENTE' " & _
           " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='8' THEN 'HIPOTECARIOS' " & _
           " END Tipo, " & _
           " Descripcion=ISNULL(cConsDescripcion,'PRENDARIO') , " & _
           " COUNT(*)Cantidad, " & _
           " SUM(CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1'THEN C.nMontoDesemb ELSE C.nMontoDesemb * " & psTipoCambio & " END )Monto "

    sSQL = sSQL & " FROM DBConsolidada..CreditoConsol C " & _
            " INNER JOIN DBConsolidada..CreditoConsolTotal CCT ON C.cCtaCod = CCT.cCtaCod " & _
            " INNER JOIN Agencias A ON CCT.cAgeCodAct = A.cAgeCod " & _
            " LEFT JOIN Constante ON nConsValor = C.nCondCre AND nConsCod= 3015 " & _
            " Where Year(C.dFecVig) = '" & lsAnio & "' And Month(C.dFecVig) = '" & lsMes & "' " & _
            " GROUP BY cAgeCod, cAgeDescripcion, " & _
            " CASE WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='1'THEN 'CORPORATIVO' " & _
            " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='2' THEN 'GRANDES EMPRESAS' " & _
            " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='3' THEN 'MEDIANAS EMPRESAS' " & _
            " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='4' THEN 'PEQUEÑAS EMPRESAS' " & _
            " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='5' THEN 'MICROEMPRESAS' " & _
            " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='6' THEN 'CONSUMO REVOLVENTE' " & _
            " WHEN CCT.cTpoCredCod = '755' THEN 'PRENDARIO' " & _
            " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='7' THEN 'CONSUMO NO REVOLVENTE' " & _
            " WHEN SUBSTRING(CCT.cTpoCredCod,1,1)='8' THEN 'HIPOTECARIOS' " & _
            " END, C.nCondCre, cConsDescripcion " & _
            " ORDER BY cAgeCod, C.nCondCre"
'*********************************


    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCredDesembolsosPla = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function


Public Function GetCarteraVencidaPla(ByVal psFecha As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset
    'EJVG20120809 ***
''********* BRGO BASILEA II
'    sSQL = " SELECT cAgeDescripcion Descripcion," & _
'           " CASE WHEN SUBSTRING(C.cProdCod,1,1)='1'THEN 'CORPORATIVO' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='2' THEN 'GRANDES EMPRESAS' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='3' THEN 'MEDIANAS EMPRESAS' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='4' THEN 'PEQUEÑAS EMPRESAS' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='5' THEN 'MICROEMPRESAS' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='6' THEN 'CONSUMO REVOLVENTE' " & _
'           " WHEN C.cProdCod = '755' THEN 'PRENDARIO' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='7' THEN 'CONSUMO NO REVOLVENTE' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='8' THEN 'HIPOTECARIOS' " & _
'           " END Tipo, " & _
'           " Moneda= CASE WHEN cMoneda='1' THEN 'SOLES' ELSE 'DOLARES' END, " & _
'           " Saldo = Sum(nSaldo) "
'    sSQL = sSQL & " FROM DBConsolidada..CredSaldosCont C " & _
'           " INNER JOIN Constante P ON C.cProdCod = P.nConsValor AND nConsCod= 3034 " & _
'           " INNER JOIN Agencias A ON A.cAgeCod = C.cAgeCod " & _
'           " WHERE cConcepto='C' AND DATEDIFF(d,dFecha,'" & psFecha & "')=0 " & _
'           " AND cEstado ='5' " & _
'           " GROUP BY A.cAgeCod,cAgeDescripcion, " & _
'           " CASE WHEN SUBSTRING(C.cProdCod,1,1)='1'THEN 'CORPORATIVO' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='2' THEN 'GRANDES EMPRESAS' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='3' THEN 'MEDIANAS EMPRESAS' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='4' THEN 'PEQUEÑAS EMPRESAS' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='5' THEN 'MICROEMPRESAS' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='6' THEN 'CONSUMO REVOLVENTE' " & _
'           " WHEN C.cProdCod = '755' THEN 'PRENDARIO' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='7' THEN 'CONSUMO NO REVOLVENTE' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='8' THEN 'HIPOTECARIOS' " & _
'           " End " & _
'           " , cMoneda " & _
'           " ORDER BY A.cAgeCod, cMoneda ASC"
''***********************
    sSQL = "Exec stp_sel_GetCarteraVencidaPla '" & psFecha & "'"
    'END EJVG *******
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCarteraVencidaPla = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function



Public Function GetCarteraRefinanciadaPla(ByVal psFecha As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

'    sSQL = " SELECT cAgeDescripcion," & _
'        " CASE WHEN SUBSTRING(C.cProdCod,1,1)='1'THEN 'COMERCIAL' " & _
'        "    WHEN SUBSTRING(C.cProdCod,1,1)='2'THEN 'MICROEMPRESA' " & _
'        "    WHEN (C.cProdCod)='305'THEN 'PRENDARIO' " & _
'        "    WHEN SUBSTRING(C.cProdCod,1,1)='3'THEN 'CONSUMO' " & _
'        "    WHEN SUBSTRING(C.cProdCod,1,1)='4'THEN 'HIPOTECARIO' " & _
'        " END, " & _
'        " cMoneda= CASE WHEN cMoneda='1'THEN'SOLES'ELSE'DOLARES'END, " & _
'        " nSaldo = Sum(nSaldo) " & _
'        " FROM DBConsolidada..CredSaldosCont C " & _
'        " INNER JOIN DBCmacMaynas..Constante P ON C.cProdCod = P.nConsValor AND nConsCod= 1001 " & _
'        " INNER JOIN DBCmacMaynas..Agencias A ON A.cAgeCod = C.cAgeCod " & _
'        " WHERE cConcepto='C' AND DATEDIFF(d,dFecha,'" & psFecha & "')=0 " & _
'        " AND cRefinanc = 'R' and cEstado IN('1','3')" & _
'        " GROUP BY A.cAgeCod,cAgeDescripcion, " & _
'        " CASE WHEN SUBSTRING(C.cProdCod,1,1)='1'THEN 'COMERCIAL' " & _
'        " WHEN SUBSTRING(C.cProdCod,1,1)='2'THEN 'MICROEMPRESA' " & _
'        " WHEN (C.cProdCod)='305'THEN 'PRENDARIO' " & _
'        " WHEN SUBSTRING(C.cProdCod,1,1)='3'THEN 'CONSUMO' " & _
'        " WHEN SUBSTRING(C.cProdCod,1,1)='4'THEN 'HIPOTECARIO' " & _
'        " End " & _
'        " , cMoneda " & _
'        " ORDER BY A.cAgeCod, cMoneda DESC"

'EJVG20120813 ***
'    sSQL = " SELECT cAgeDescripcion," & _
'           " CASE WHEN SUBSTRING(C.cProdCod,1,1)='1'THEN 'CORPORATIVO' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='2' THEN 'GRANDES EMPRESAS' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='3' THEN 'MEDIANAS EMPRESAS' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='4' THEN 'PEQUEÑAS EMPRESAS' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='5' THEN 'MICROEMPRESAS' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='6' THEN 'CONSUMO REVOLVENTE' " & _
'           " WHEN C.cProdCod = '755' THEN 'PRENDARIO' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='7' THEN 'CONSUMO NO REVOLVENTE' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='8' THEN 'HIPOTECARIOS' " & _
'           " END, " & _
'           " cMoneda= CASE WHEN cMoneda='1'THEN'SOLES'ELSE'DOLARES'END, " & _
'           " nSaldo = Sum(nSaldo) "
'    sSQL = sSQL & " FROM DBConsolidada..CredSaldosCont C " & _
'        " INNER JOIN DBCmacMaynas..Constante P ON C.cProdCod = P.nConsValor AND nConsCod= 3034 " & _
'        " INNER JOIN DBCmacMaynas..Agencias A ON A.cAgeCod = C.cAgeCod " & _
'        " WHERE cConcepto='C' AND DATEDIFF(d,dFecha,'" & psFecha & "')=0 " & _
'        " AND cEstado IN ('4')" & _
'        " GROUP BY A.cAgeCod,cAgeDescripcion, " & _
'        " CASE WHEN SUBSTRING(C.cProdCod,1,1)='1'THEN 'CORPORATIVO' " & _
'        " WHEN SUBSTRING(C.cProdCod,1,1)='2' THEN 'GRANDES EMPRESAS' " & _
'        " WHEN SUBSTRING(C.cProdCod,1,1)='3' THEN 'MEDIANAS EMPRESAS' " & _
'        " WHEN SUBSTRING(C.cProdCod,1,1)='4' THEN 'PEQUEÑAS EMPRESAS' " & _
'        " WHEN SUBSTRING(C.cProdCod,1,1)='5' THEN 'MICROEMPRESAS' " & _
'        " WHEN SUBSTRING(C.cProdCod,1,1)='6' THEN 'CONSUMO REVOLVENTE' " & _
'        " WHEN C.cProdCod = '755' THEN 'PRENDARIO' " & _
'        " WHEN SUBSTRING(C.cProdCod,1,1)='7' THEN 'CONSUMO NO REVOLVENTE' " & _
'        " WHEN SUBSTRING(C.cProdCod,1,1)='8' THEN 'HIPOTECARIOS' " & _
'        " End " & _
'        " , cMoneda " & _
'        " ORDER BY A.cAgeCod, cMoneda DESC"
    sSQL = "Exec stp_sel_GetCarteraRefinanciadaPla '" & psFecha & "'"
'END EJVG *******
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCarteraRefinanciadaPla = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function
'By Capi 18122007 para planeamiento
Public Function GetCarteraRecupCapital(ByVal psFecha As String)
    Dim sSQL As String
    Dim psFecha1 As String
    Dim oCon As New DConecta
    Set oCon = New DConecta
    psFecha1 = Mid(psFecha, 4, 2) & Mid(psFecha, 7, 4)
    oCon.AbreConexion
    sSQL = " Exec Stp_Sel_RecuperacionCapitalCartera '" & psFecha1 & "'"
    Set GetCarteraRecupCapital = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
 End Function
 'By Capi 09122008
 Public Sub ActualizaIntDevConAlineamiento(ByVal psFecha As String)
    Dim sSQL As String
    Dim psFecha1 As String
    Dim oCon As New DConecta
    Set oCon = New DConecta
    psFecha1 = Mid(psFecha, 4, 2) & Mid(psFecha, 7, 4)
    oCon.AbreConexion
    sSQL = " Exec stp_upd_Consolida_ReclasificaIntDevConAlineamiento '" & psFecha & "'"
    oCon.Ejecutar (sSQL)
    
    sSQL = " Exec stp_upd_Consolida_ReclasificaIntDevConAlineamientoAF '" & psFecha & "'"
    oCon.Ejecutar (sSQL)
    
    oCon.CierraConexion
    Set oCon = Nothing
 End Sub



'GetNumCliCtasEnAhorrosCreditos

Public Function GetCarteraJudicialesPla(ByVal psFecha As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset
    'EJVG20120817 ***
'    '******  BRGO - BASILEA II
'    sSQL = "SELECT cAgeDescripcion Agencia," & _
'           " Producto= CASE WHEN SUBSTRING(C.cProdCod,1,1)='1'THEN 'CORPORATIVO' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='2' THEN 'GRANDES EMPRESAS' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='3' THEN 'MEDIANAS EMPRESAS' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='4' THEN 'PEQUEÑAS EMPRESAS' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='5' THEN 'MICROEMPRESAS' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='6' THEN 'CONSUMO REVOLVENTE' " & _
'           " WHEN C.cProdCod = '755' THEN 'PRENDARIO' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='7' THEN 'CONSUMO NO REVOLVENTE' " & _
'           " WHEN SUBSTRING(C.cProdCod,1,1)='8' THEN 'HIPOTECARIOS' " & _
'           " END, " & _
'           " Moneda= CASE WHEN cMoneda='1'THEN'SOLES'ELSE'DOLARES'END, " & _
'           " COUNT(*) Cantidad, " & _
'           " Saldo = Sum(nSaldo) "
'    sSQL = sSQL & " FROM DBConsolidada..CredSaldosCont C " & _
'            " INNER JOIN Constante P ON C.cProdCod = P.nConsValor AND nConsCod= 3034 " & _
'            " INNER JOIN Agencias A ON A.cAgeCod = C.cAgeCod " & _
'            " WHERE cConcepto='C' AND DATEDIFF(d,dFecha,'" & psFecha & "')=0 AND cEstado='6' " & _
'            " GROUP BY A.cAgeCod,cAgeDescripcion, " & _
'            " CASE WHEN SUBSTRING(C.cProdCod,1,1)='1'THEN 'CORPORATIVO' " & _
'            " WHEN SUBSTRING(C.cProdCod,1,1)='2' THEN 'GRANDES EMPRESAS' " & _
'            " WHEN SUBSTRING(C.cProdCod,1,1)='3' THEN 'MEDIANAS EMPRESAS' " & _
'            " WHEN SUBSTRING(C.cProdCod,1,1)='4' THEN 'PEQUEÑAS EMPRESAS' " & _
'            " WHEN SUBSTRING(C.cProdCod,1,1)='5' THEN 'MICROEMPRESAS' " & _
'            " WHEN SUBSTRING(C.cProdCod,1,1)='6' THEN 'CONSUMO REVOLVENTE' " & _
'            " WHEN C.cProdCod = '755' THEN 'PRENDARIO' " & _
'            " WHEN SUBSTRING(C.cProdCod,1,1)='7' THEN 'CONSUMO NO REVOLVENTE' " & _
'            " WHEN SUBSTRING(C.cProdCod,1,1)='8' THEN 'HIPOTECARIOS' " & _
'            " END,cMoneda " & _
'            " ORDER BY A.cAgeCod,cMoneda DESC"
'    '**********************
    sSQL = "Exec stp_sel_GetCarteraJudicialPla '" & psFecha & "'"
    'END EJVG ********
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCarteraJudicialesPla = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function


Public Function GetInteresesDiferidos(ByVal psFecha As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset
    'By Capi 14112007 porque no generaba correctamente el calculo solo se agrego a 30 30.0
    sSQL = "Select substring(cctacod,4,2)Agencia," & _
           " Sum(Round( (POWER (1+nTasaInt /100 ,(datediff(day,dFecUltPago,'" & psFecha & "')/30.0 )) -1) * nSaldoCap,2)) " & _
           " from DBConsolidada..CreditoConsol " & _
           " where nPrdEstado in ('2104','2101','2106','2107') " & _
           " group by substring(cctacod,4,2) " & _
           " order by substring(cctacod,4,2)"
    '" where nPrdEstado in ('2104','2101','2106','20107') " & _ by capi 14112007 porque el estado 20107 no existe
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetInteresesDiferidos = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function



Public Function GetPignoraticiosVigentes() As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = " Select substring(cCtaCod,4,2)Agencia,nPlazoApr,sum(nsaldocap)Saldo" & _
           " From dbconsolidada..CreditoConsol " & _
           " where substring(cctacod,6,3)='305' and nprdestado in ('2101','2104','2106','2107') " & _
           " and ndiasatraso <=30 " & _
           " group by  substring(cCtaCod,4,2),nPlazoApr  order by  substring(cCtaCod,4,2),nPlazoApr "

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetPignoraticiosVigentes = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function

Public Function GetPignoraticiosVencidos() As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = " Select substring(cCtaCod,4,2)Agencia,nPlazoApr,sum(nsaldocap)Saldo " & _
            " From dbconsolidada..CreditoConsol " & _
            " where substring(cctacod,6,3)='305' and nprdestado in ('2101','2104','2106','2107') " & _
            " and ndiasatraso > 30 " & _
            " group by  substring(cCtaCod,4,2),nPlazoApr " & _
            " order by  substring(cCtaCod,4,2),nPlazoApr"

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetPignoraticiosVencidos = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function



Public Function GetCreditosCondonados(ByVal psFecha As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

'JUEZ 20130429 *********************************************************************
'    sSQL = "Select p.cCtaCod, Case When nPrdEstado = 2203 Then 'Judicial' Else 'Castigado' End Tipo, cl.cDescripcion,  nSaldoIntComp, nSaldoIntMor,(nSaldoIntComp+nSaldoIntMor) Total, " & _
'        " Case substring(p.cCtaCod,9,1) When '1' Then 'Soles' Else 'Dolares' End  Moneda " & _
'        " From Producto P " & _
'        " Join ColocRecup CR On P.cCtaCod = CR.cCtaCod " & _
'        " Join Colocaciones C On C.cCtaCod = CR.cCtaCod " & _
'        " Inner Join ColocLineaCredito CL On CL.cLineaCred = left(C.cLineaCred,4) " & _
'        " Where nPRDEstado In (2203,2204) And nsaldoIntComp > 0 " & _
'        " And cUltimaActualizacion like '" & psFecha & "%' "
sSQL = "stp_sel_GetCreditosCondonados '" & psFecha & "'"
'END JUEZ **************************************************************************
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCreditosCondonados = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function


Public Function GetProvisionCreditos() As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "Select substring(cCtaCod,4,2)Agencia,Case cCalNor when '0' then 'Generica' Else 'Especifico' End," & _
            " Moneda = Case substring(cCtaCod,9,1) When '1' Then 'Soles' Else 'Dolares' End, " & _
            " Sum(nProvision) Provision " & _
            " From ColocCalifProv " & _
            " where nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205) " & _
            " group by substring(cCtaCod,4,2),cCalNor,substring(cCtaCod,9,1) " & _
            " order by substring(cCtaCod,4,2),cCalNor,substring(cCtaCod,9,1) "

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetProvisionCreditos = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function


Public Function GetIntCreditosRefinanciados(ByVal psFechaIni As String, ByVal psFechaFin As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

'**** By BRGO 20100611 BASILEA II
    sSQL = " EXEC stp_sel_GetIntCreditosRefinanciados '" & psFechaIni & "', '" & psFechaFin & "'"    'JIPR20210803
'***********************************************************

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetIntCreditosRefinanciados = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function


Public Function GetProvisionCartasFianzas() As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = " Select Agencia =Substring(c.cCtaCod,4,2),c.cCtaCod, " & _
            " Case substring(c.cCtaCod,6,3) When '121' Then 'COMERCIAL' When '221' Then 'MES' End Producto, " & _
            " cl.cDescripcion,Moneda = Case substring(cCtaCod,9,1) When '1' Then 'Soles' Else 'Dolares' End, " & _
            " SUM( nProvision ) nSaldo " & _
            " FROM ColocCalifProv c " & _
            " join  (Select distinct substring(E.cLineaCred,1,4) Lin, cCtacont " & _
            " from DBConsolidada..ColocLineaCreditoEquiv E " & _
            " ) E    on substring(c.cLineaCred,1,4)  = lin " & _
            " Inner Join ColocLineaCredito CL On left(C.cLineaCred,4) = CL.cLineaCred " & _
            " WHERE nprdEstado in ('2020', '2021', '2022', '2030', '2031', '2032','2101', '2104', '2106', '2107', '2201','2205','2092') " & _
            " And substring(cctacod,7,2) = '21' " & _
            " Group By  Substring(cCtaCod,4,2),cCtaCod,cl.cDescripcion,substring(cCtaCod,9,1), " & _
            " Case substring(c.cCtaCod,6,3) When '121' Then 'COMERCIAL' When '221' Then 'MES' End "


    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetProvisionCartasFianzas = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function



Public Function GetPlazoFijoRiesgos(ByVal psFecha As String, ByVal psMoneda As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = " Select 'MEN_007',SUM(nSaldCnt) as nSaldo " & _
        " from CapSaldosDiarios CS " & _
        " Inner Join Captaciones C ON CS.cctacod=C.cctacod " & _
        " Where datediff(day,CS.dfecha, '" & psFecha & "')=0   and substring(CS.cctacod,9,1)='" & psMoneda & "' " & _
        " and substring(CS.cctacod,6,3)=233 and " & _
        " DATEDIFF(d,dApertura,'" & psFecha & "')BETWEEN 0 AND 7 " & _
        " group by substring(CS.cctacod,6,3) " & _
        " Union " & _
        " Select '008_015',SUM(nSaldCnt) as nSaldo" & _
        " from CapSaldosDiarios CS " & _
        " Inner Join Captaciones C ON CS.cctacod=C.cctacod " & _
        " Where datediff(day,CS.dfecha, '" & psFecha & "')=0   and substring(CS.cctacod,9,1)='" & psMoneda & "' " & _
        " and substring(CS.cctacod,6,3)=233 and " & _
        " DATEDIFF(d,dApertura,'" & psFecha & "')BETWEEN 8 AND 15 " & _
        " group by substring(CS.cctacod,6,3)"
    sSQL = sSQL & "union " & _
        " Select '016_030',SUM(nSaldCnt) as nSaldo " & _
        " from CapSaldosDiarios CS " & _
        " Inner Join Captaciones C ON CS.cctacod=C.cctacod " & _
        " Where datediff(day,CS.dfecha, '" & psFecha & "')=0   and substring(CS.cctacod,9,1)='" & psMoneda & "' " & _
        " and substring(CS.cctacod,6,3)=233 and " & _
        " DATEDIFF(d,dApertura,'" & psFecha & "')BETWEEN 16 AND 30 " & _
        " group by substring(CS.cctacod,6,3) " & _
        " Union " & _
        " Select '031_060',SUM(nSaldCnt) as nSaldo " & _
        " from CapSaldosDiarios CS " & _
        " Inner Join Captaciones C ON CS.cctacod=C.cctacod " & _
        " Where datediff(day,CS.dfecha, '" & psFecha & "')=0   and substring(CS.cctacod,9,1)='" & psMoneda & "' " & _
        " and substring(CS.cctacod,6,3)=233 and " & _
        " DATEDIFF(d,dApertura,'" & psFecha & "')BETWEEN 31 AND 60 " & _
        " group by substring(CS.cctacod,6,3)"
     sSQL = sSQL & "union " & _
        " Select '061_090',SUM(nSaldCnt) as nSaldo " & _
        " from CapSaldosDiarios CS " & _
        "Inner Join Captaciones C ON CS.cctacod=C.cctacod " & _
        " Where datediff(day,CS.dfecha, '" & psFecha & "')=0   and substring(CS.cctacod,9,1)='" & psMoneda & "' " & _
        " and substring(CS.cctacod,6,3)=233 and " & _
        " DATEDIFF(d,dApertura,'" & psFecha & "')BETWEEN 61 AND 90 " & _
        " group by substring(CS.cctacod,6,3) " & _
        " Union " & _
        " Select '091_180',SUM(nSaldCnt) as nSaldo " & _
        " from CapSaldosDiarios CS " & _
        " Inner Join Captaciones C ON CS.cctacod=C.cctacod " & _
        " Where datediff(day,CS.dfecha, '" & psFecha & "')=0   and substring(CS.cctacod,9,1)='" & psMoneda & "' " & _
        " and substring(CS.cctacod,6,3)=233 and " & _
        " DATEDIFF(d,dApertura,'" & psFecha & "')BETWEEN 91 AND 180 " & _
        " group by substring(CS.cctacod,6,3) "
     sSQL = sSQL & "union " & _
        " Select '181_360',SUM(nSaldCnt) as nSaldo " & _
        " from CapSaldosDiarios CS " & _
        " Inner Join Captaciones C ON CS.cctacod=C.cctacod " & _
        " Where datediff(day,CS.dfecha, '" & psFecha & "')=0   and substring(CS.cctacod,9,1)='" & psMoneda & "' " & _
        " and substring(CS.cctacod,6,3)=233 and " & _
        " DATEDIFF(d,dApertura,'" & psFecha & "')BETWEEN 181 AND 360 " & _
        " group by substring(CS.cctacod,6,3) " & _
        " Union " & _
        " Select '02_AÑOS',SUM(nSaldCnt) as nSaldo " & _
        " from CapSaldosDiarios CS " & _
        " Inner Join Captaciones C ON CS.cctacod=C.cctacod " & _
        " Where datediff(day,CS.dfecha, '" & psFecha & "')=0   and substring(CS.cctacod,9,1)='" & psMoneda & "' " & _
        " and substring(CS.cctacod,6,3)=233 and " & _
        " DATEDIFF(d,dApertura,'" & psFecha & "')BETWEEN 361 AND 720 " & _
        " group by substring(CS.cctacod,6,3) "
     sSQL = sSQL & "union " & _
        " Select '03_AÑOS',SUM(nSaldCnt) as nSaldo " & _
        " from CapSaldosDiarios CS " & _
        " Inner Join Captaciones C ON CS.cctacod=C.cctacod " & _
        " Where datediff(day,CS.dfecha, '" & psFecha & "')=0   and substring(CS.cctacod,9,1)='" & psMoneda & "' " & _
        " and substring(CS.cctacod,6,3)=233 and " & _
        " DATEDIFF(d,dApertura,'" & psFecha & "')BETWEEN 721 AND 1080 " & _
        " group by substring(CS.cctacod,6,3) " & _
        " Union " & _
        " Select '04_AÑOS',SUM(nSaldCnt) as nSaldo " & _
        " from CapSaldosDiarios CS " & _
        " Inner Join Captaciones C ON CS.cctacod=C.cctacod " & _
        " Where datediff(day,CS.dfecha, '" & psFecha & "')=0   and substring(CS.cctacod,9,1)='" & psMoneda & "' " & _
        " and substring(CS.cctacod,6,3)=233 and " & _
        " DATEDIFF(d,dApertura,'" & psFecha & "')BETWEEN 1081 AND 1440 " & _
        " group by substring(CS.cctacod,6,3)"
     sSQL = sSQL & "union " & _
        " Select '05_AÑOS',SUM(nSaldCnt) as nSaldo " & _
        " from CapSaldosDiarios CS " & _
        " Inner Join Captaciones C ON CS.cctacod=C.cctacod " & _
        " Where datediff(day,CS.dfecha, '" & psFecha & "')=0   and substring(CS.cctacod,9,1)='" & psMoneda & "' " & _
        " and substring(CS.cctacod,6,3)=233 and " & _
        " DATEDIFF(d,dApertura,'" & psFecha & "')BETWEEN 1441 AND 1800 " & _
        " group by substring(CS.cctacod,6,3) " & _
        " Union " & _
        " Select '06_10AÑOS', SUM(nSaldCnt) as nSaldo " & _
        " from CapSaldosDiarios CS " & _
        " Inner Join Captaciones C ON CS.cctacod=C.cctacod " & _
        " Where datediff(day,CS.dfecha, '" & psFecha & "')=0   and substring(CS.cctacod,9,1)='" & psMoneda & "' " & _
        " and substring(CS.cctacod,6,3)=233 and " & _
        " DATEDIFF(d,dApertura,'" & psFecha & "')BETWEEN 1801 AND 3600 " & _
        " group by substring(CS.cctacod,6,3)"
      sSQL = sSQL & "union " & _
        " Select '11_20AÑOS',SUM(nSaldCnt) as nSaldo" & _
        " from CapSaldosDiarios CS " & _
        " Inner Join Captaciones C ON CS.cctacod=C.cctacod " & _
        " Where datediff(day,CS.dfecha, '" & psFecha & "')=0   and substring(CS.cctacod,9,1)='" & psMoneda & "' " & _
        " and substring(CS.cctacod,6,3)=233 and " & _
        " DATEDIFF(d,dApertura,'" & psFecha & "')BETWEEN 3601 AND 7200 " & _
        " group by substring(CS.cctacod,6,3) " & _
        " Union " & _
        " Select '21_MAS',SUM(nSaldCnt) as nSaldo " & _
        " from CapSaldosDiarios CS " & _
        " Inner Join Captaciones C ON CS.cctacod=C.cctacod " & _
        " Where datediff(day,CS.dfecha, '" & psFecha & "')=0   and substring(CS.cctacod,9,1)='" & psMoneda & "' " & _
        " and substring(CS.cctacod,6,3)=233 and " & _
        " DATEDIFF(d,dApertura,'" & psFecha & "')>=7201 " & _
        " group by substring(CS.cctacod,6,3) "


    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetPlazoFijoRiesgos = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function




Public Function GetITFQuincena(ByVal psFechaIni As String, ByVal psFechaFin As String, Optional nTipoReporte As Integer = 1) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetITFQuincenaErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset
'ALPA 20081010***********************************************************
''
''    sSQL = " Select Oficina, " & _
''           " IsNull(Sum(Case When aaa.cOpecod In ('990101','990102','990109','990301','990302') Then nMontoDetSol End),0) MontoSolAho, " & _
''           " IsNull(Sum(Case When aaa.cOpecod In ('990101','990102','990109','990301','990302') Then nMontoDetDol End),0) As MonDolAho, " & _
''           " IsNull(Sum(Case When aaa.cOpecod In ('990103','990105','990106','990107','990108','990303') Then nMontoDetSol End),0) As MonSolCre, " & _
''           " IsNull(Sum(Case When aaa.cOpecod In ('990103','990105','990106','990107','990108','990303') Then nMontoDetDol End),0) As MonDolCre, " & _
''           " IsNull(Sum(Case When aaa.cOpecod In ('990104','990304') Then nMontoDetSol End),0) As MonSolPre,  " & _
''           " IsNull(Sum(Case When aaa.cOpecod In ('1') Then nMontoDetSol End),0) As MonSolCaj,  " & _
''           " IsNull(Sum(Case When aaa.cOpecod In ('2') Then nMontoDetDol End),0) As MonDolCaj  " & _
''           " From ( SELECT  MC.CCTACOD, (SELECT max(cperscod) FROM ProductoPersona pp where pp.cctacod = MC.cctacod And nPrdPersRelac = 20) CodPers , " & _
''           "    Substring(MC.CCTACOD,4,2) oficina,MC.COPECOD,O.cOpeDesc, " & _
''           "    SUM(Case Substring(MC.CCTACOD,9,1) when '2' then  round(MD.nMonto,2) else 0.00 end) as nMontoDetDol , " & _
''           "    SUM(Case Substring(MC.CCTACOD,9,1) when '1' then  Round(MD.nMonto,2) else 0.00 end) as nMontoDetSol " & _
''           " FROM MOV M " & _
''           "    JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO " & _
''           "    JOIN MOVCOLDET MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.COPECOD = MC.COPECOD AND MD.CCTACOD = MC.CCTACOD " & _
''           "    JOIN OPETPO O ON O.COPECOD = MC.COPECOD " & _
''           "    JOIN PRODUCTOCONCEPTO PC ON PC.nPrdConceptoCod = MD.nPrdConceptoCod " & _
''           " WHERE MC.COPECOD LIKE '99%' AND LEFT(M.CMOVNRO,8) BETWEEN '" & psFechaIni & "' AND '" & psFechaFin & "' AND M.NMOVFLAG = 0 " & _
''           " GROUP BY M.cMovNro,MC.COPECOD,O.cOpeDesc,MC.CCTACOD " & _
''           " Union All "
''     sSQL = sSQL & " SELECT  MC.CCTACOD, (SELECT max(cperscod) FROM ProductoPersona pp where pp.cctacod = MC.cctacod And nPrdPersRelac = 10) CodPers, " & _
''                   " Substring(MC.CCTACOD,4,2) oficina, MC.COPECOD,O.cOpeDesc, " & _
''                   " SUM(Case Substring(MC.CCTACOD,9,1) when '2' then  round(MD.nMonto,2) else 0.00 end) as nMontDetDol , " & _
''                   " SUM(Case Substring(MC.CCTACOD,9,1) when '1' then  Round(MD.nMonto,2) else 0.00 end) as nMontDetSol " & _
''                   " FROM    MOV M  " & _
''                   " JOIN MOVCAP MC ON MC.NMOVNRO = M.NMOVNRO " & _
''                   " JOIN MOVCAPDET MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.COPECOD = MC.COPECOD AND MD.CCTACOD = MC.CCTACOD " & _
''                   " JOIN OPETPO O ON O.COPECOD = MC.COPECOD " & _
''                   " JOIN PRODUCTOCONCEPTO PC ON PC.nPrdConceptoCod = MD.nConceptoCod " & _
''                   " WHERE MC.COPECOD LIKE '99%' AND LEFT(M.CMOVNRO,8) BETWEEN '" & psFechaIni & "' AND '" & psFechaFin & "' AND M.NMOVFLAG = 0 " & _
''                   " GROUP BY M.cMovNro,MC.COPECOD,O.cOpeDesc,MC.CCTACOD " & _
''                   " Union All " & _
''                   " SELECT  '1089800000272' cCtaCod, '1089800000272' CodPers, " & _
''                   " Substring(M.cMovNro,18,2) oficina,'2' as COPECOD, 'DOLARES' As cOpeDesc, " & _
''                   " SUM(round(ABS(ME.NMOVMEIMPORTE),2))  as nMontDetDol, 0.00 as nMontDetSol " & _
''                   " FROM    MOV M  " & _
''                   " JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO " & _
''                   " LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO and ME.nMovItem = MC.nMovItem " & _
''                   " JOIN OPETPO O ON O.COPECOD = M.COPECOD " & _
''                   " WHERE LEFT(M.CMOVNRO,8) BETWEEN '" & psFechaIni & "' AND '" & psFechaFin & "' " & _
''                   " AND M.NMOVFLAG = 0 AND M.COPECOD Not In ('701105') AND MC.cCtaContCod LIKE '21240207%'  " & _
''                   " and MC.nMovImporte < 0 " & _
''                   " Group By M.cMovNro " & _
''                   " Union All "
''     sSQL = sSQL & " SELECT  '1089800000272' cCtaCod, '1089800000272' CodPers, " & _
''                   " Substring(M.cMovNro,18,2) oficina,'1' as COPECOD, 'SOLES' As cOpeDesc, " & _
''                   " 0.00 As nMontDetDol,SUM(round(ABS(MC.NMOVIMPORTE),2))  as nMontDetSol  " & _
''                   " FROM    MOV M " & _
''                   " JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO " & _
''                   " LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO and ME.nMovItem = MC.nMovItem " & _
''                   " JOIN OPETPO O ON O.COPECOD = M.COPECOD " & _
''                   " WHERE LEFT(M.CMOVNRO,8) BETWEEN '" & psFechaIni & "' AND '" & psFechaFin & "' " & _
''                   " AND M.NMOVFLAG = 0 AND M.COPECOD Not In ('701105') AND MC.cCtaContCod LIKE '21140207%'  " & _
''                   " and MC.nMovImporte < 0 Group By M.cMovNro) As aaa " & _
''                   " left join persona pe on aaa.codpers =  pe.cperscod " & _
''                   " left join persid bb on aaa.codpers =  bb.cperscod and bb.cPersIDTpo = 1 " & _
''                   " left join persid cc on aaa.codpers =  cc.cperscod and cc.cPersIDTpo = 2 " & _
''                   " left join persid dd on aaa.codpers =  dd.cperscod and dd.cPersIDTpo = 4 " & _
''                   " left join persid ee on aaa.codpers =  ee.cperscod and ee.cPersIDTpo = 11 " & _
''                   " left join persid ff on aaa.codpers =  ff.cperscod and ff.cPersIDTpo = 10 " & _
''                   " Group by  oficina Order By oficina "
    sSQL = "exec stp_sel_ReporteGetITFQuincena '" & psFechaIni & "','" & psFechaFin & "', " & nTipoReporte
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetITFQuincena = rs
    Exit Function
GetITFQuincenaErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function


Public Function GetDetalleGarantias(ByVal psFecha As String, ByVal psTipoCambio As Currency, ByVal pnMoneda As Integer) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

'    sSQL = " SELECT substring(cg.cCtaCod,6,3) Prod,cc.cCtaCod, CASE WHEN cFF is NULL THEN '01' ELSE cFF END FF, substring(cg.cCtaCod,4,2) Agencia " & _
'            " ,ROUND(nMonto*ROUND( CASE WHEN SUBSTRING(CC.CCTACOD,9,1)='1' " & _
'            " THEN CC.nSaldoCap ELSE ROUND(CC.nSaldoCap* " & psTipoCambio & " ,2) END / (CASE WHEN ISNULL(D.SaldoCap,0)=0 THEN 1 ELSE ISNULL(D.SaldoCap,0) END), 2)*1,2) Monto " & _
'            " ,CASE WHEN cg.nMoneda ='1' THEN 'SOLES' ELSE 'DOLARES' END Moneda, " & _
'            " replace(replace(replace(replace(replace(replace(c.cConsDescripcion,' ','_'),'.',''),'/','_'),')',''),'(',''),':','') cConsDescripcion , " & _
'            " convert(int,cTipoGarant) cTipoGarant " & _
'            " FROM DBConsolidada..CredGarantiasConsol cg " & _
'            " Inner Join DBConsolidada..ProductoPersonaConsol PP On CG.cCtaCod = PP.cCtaCod " & _
'            " Inner Join DBConsolidada..CreditoConsol CC On CC.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 20 " & _
'            " Inner Join (select  GC.cPersCod, Sum(  Case Substring(B.cctacod,9,1) When '1' then B.nSaldoCap else B.nSaldoCap * " & psTipoCambio & " end) SaldoCap, Count(*) Num " & _
'            " from    DBConsolidada..ProductoPersonaConsol GC " & _
'            " INNER jOIN DBConsolidada..creditoconsol B On GC.cCtaCod = B.cCtaCod and GC.nPrdPersRelac = 20 " & _
'            " group by GC.cPersCod " & _
'            " ) As D On PP.cPersCod = D.cPersCod" & _
'            " inner join Constante c ON c.nConsValor = convert(int,cg.cTipoGarant) " & _
'            " WHERE datediff(d,CG.dFecha,'" & psFecha & "') = 0 and c.nConsCod = 1027 and cg.nMoneda= " & pnMoneda & " " & _
'            " order by Agencia,prod,cTipoGarant,Moneda"
    'ALPA 20090812**********************************************************************************************
    
    '*** PEAC 20110305 - SE MANDO A UN STP
    
'    sSQL = " SELECT substring(cc.cCtaCod,6,3) Prod,cc.cCtaCod, CASE WHEN cFF is NULL THEN '01' ELSE cFF END FF, substring(cg.cCtaCod,4,2) Agencia, "
'    sSQL = sSQL & "sum(round(CG.nMonto,2)) Monto"
'    ''sSQL = sSQL & "round(sum(CG.nMonto),2) Monto"
'    sSQL = sSQL & " ,CASE WHEN cg.nMoneda ='1' THEN 'SOLES' ELSE 'DOLARES' END Moneda, "
'    sSQL = sSQL & " replace(replace(replace(replace(replace(replace(c.cConsDescripcion,' ','_'),'.',''),'/','_'),')',''),'(',''),':','') cConsDescripcion , "
'    sSQL = sSQL & " convert(int,cTipoGarant) cTipoGarant "
'    sSQL = sSQL & " FROM DBConsolidada..CredGarantiasConsol cg "
'    sSQL = sSQL & " Inner Join DBConsolidada..ProductoPersonaConsol PP On CG.cCtaCod = PP.cCtaCod "
'    sSQL = sSQL & " Inner Join DBConsolidada..CreditoConsol CC On CC.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 20 "
'    sSQL = sSQL & " inner join Constante c ON c.nConsValor = convert(int,cg.cTipoGarant) "
'    sSQL = sSQL & " WHERE datediff(d,CG.dFecha,'" & psFecha & "') = 0 and c.nConsCod = 1027 and cg.nMoneda= " & pnMoneda & " "
'    'ALPA 20091103************************
'    'sSQL = sSQL & " and convert(int,CG.cTipoGarant) in (1,2,3,4,6,7,9,11,26,27,28,29,30) "
'    'ALPA 20100304****************************************************************************
'    'sSQL = sSQL & " and convert(int,CG.cTipoGarant) in (1,2,3,4,6,7,9,11,12,26,27,28,29,30) "
'    sSQL = sSQL & " and convert(int,CG.cTipoGarant) in (1,2,3,4,6,7,9,11,12,26,27,28,29,30,32) "
'    '*****************************************************************************************
'    '********************************************
'    sSQL = sSQL & " group by substring(cc.cCtaCod,6,3), cc.cCtaCod,case when cFF is null then '01' else cFF end,"
'    sSQL = sSQL & " substring(CG.cCtaCod,4,2),cg.nMoneda "
'    sSQL = sSQL & " ,c.cConsDescripcion "
'    sSQL = sSQL & " ,cTipoGarant "
'    sSQL = sSQL & " order by Agencia,prod,cTipoGarant,Moneda"
'    '************************************************************************************************************
    
    sSQL = " exec stp_sel_ReporteContDetalleGarantias '" & Format(psFecha, "yyyymmdd") & "'," & pnMoneda
    
    '*** FIN PEAC
        
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetDetalleGarantias = rs
    Exit Function
GetCreditosVigentesErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function

'*** PEAC 20080915
Public Function GetNumCliCtasEnAhorrosPorAgencias() As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta

On Error GoTo GetNumCliCtasEnAhorrosCreditosErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_ReporteNumCliCtasEnAhorros "


    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetNumCliCtasEnAhorrosPorAgencias = rs
    Exit Function
GetNumCliCtasEnAhorrosCreditosErr:
    Call RaiseError(MyUnhandledError, "clsSicmact-GetNumCliCtasEnAhorrosPorAgencias")
End Function

'*** PEAC 20080915
Public Function GetNumCliCtasEnAhorrosConsolidado() As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta

On Error GoTo GetNumCliCtasCarteraEnCreditosErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_ReporteNumCliCtasEnAhorrosConsol "

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetNumCliCtasEnAhorrosConsolidado = rs
    Exit Function
GetNumCliCtasCarteraEnCreditosErr:
    Call RaiseError(MyUnhandledError, "clsSicmact-GetNumCliCtasEnAhorrosConsolidado")
End Function


'*** PEAC 20080915
Public Function GetCarteraNumCliYCtasCredito(ByVal psTipoCambio As Currency) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta

On Error GoTo GetCarteraNumCliYCtasCreditoErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_ObtieneResumenCartera " & psTipoCambio

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCarteraNumCliYCtasCredito = rs
    Exit Function
GetCarteraNumCliYCtasCreditoErr:
    Call RaiseError(MyUnhandledError, "clsSicmact-GetCarteraNumCliYCtasCredito")
End Function

'*** PEAC 20080915
Public Function GetUniversoClientesCartera() As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta

On Error GoTo GetUniversoClientesCarteraErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_ObtieneUniversoClientesCartera "

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetUniversoClientesCartera = rs
    Exit Function
GetUniversoClientesCarteraErr:
    Call RaiseError(MyUnhandledError, "clsSicmact-GetUniversoClientesCartera")
End Function

'*** PEAC 20080915
Public Function GetObtienePeriodoConsol() As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta

On Error GoTo GetObtienePeriodoConsolErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_ObtienePeriodoConsol "

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetObtienePeriodoConsol = rs
    Exit Function
GetObtienePeriodoConsolErr:
    Call RaiseError(MyUnhandledError, "clsSicmact-GetObtienePeriodoConsol")
End Function

Public Function GetDaotCostos(ByVal psAnio As String, ByVal pnNroUIT As Integer, ByVal pnTipCamb As Double) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetDAOTCostosErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_GeneraGetCostosDAOT '" & psAnio & "'," & pnNroUIT & "," & pnTipCamb
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetDaotCostos = rs
    Exit Function
GetDAOTCostosErr:
    Call RaiseError(MyUnhandledError, "Genera Archivo DAOT ")
End Function

Public Function GetDaotIngresos(ByVal psAnio As String, ByVal pnNroUIT As Integer, ByVal pnTipCamb As Double) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetDAOTIngresosErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_GeneraGetIngresosDAOT '" & psAnio & "'," & pnNroUIT & "," & pnTipCamb
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetDaotIngresos = rs
    Exit Function
GetDAOTIngresosErr:
    Call RaiseError(MyUnhandledError, "Genera Archivo Ingresos DAOT ")
End Function

'ALPA 20090309*****************************************************
Public Function GetInteresesCreditosCastigados(ByVal psFecha As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetICreditosCastigadosErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

'    sSQL = " SELECT cAgeDescripcion Descripcion," & _
'            " CASE WHEN SUBSTRING(cCtaCod,9,1)='1'THEN 'SOLES'ELSE'DOLARES'END Moneda,SUM(nIntDev)Saldo,COUNT(*)Cantidad " & _
'            " FROM DBConsolidada..CreditoConsol INNER JOIN DBCmacMaynas..Agencias A ON A.cAgeCod = SUBSTRING(cCtaCod,4,2) " & _
'            " WHERE nPrdEstado IN(2202,2206) " & _
'            " GROUP BY cAgeCod,cAgeDescripcion,SUBSTRING(cCtaCod,9,1) " & _
'            " ORDER BY SUBSTRING(cCtaCod,9,1),cAgeCod"

    sSQL = " SELECT cAgeDescripcion Descripcion,"
    sSQL = sSQL & " CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1'"
    sSQL = sSQL & "        THEN 'SOLES'ELSE'DOLARES'END Moneda,"
    sSQL = sSQL & "    SUM(nIntDev)Saldo,COUNT(*)Cantidad "
    sSQL = sSQL & " FROM DBConsolidada..CreditoConsol CC "
    sSQL = sSQL & "    INNER JOIN DBCmacMaynas..Colocaciones C ON C.cCtaCod = CC.cCtaCod"
    sSQL = sSQL & "    INNER JOIN DBCmacMaynas..Agencias A ON A.cAgeCod = C.cAgeCodAct"
    sSQL = sSQL & " WHERE nPrdEstado IN(2202,2206)"
    sSQL = sSQL & " GROUP BY A.cAgeCod,cAgeDescripcion,SUBSTRING(C.cCtaCod,9,1) "
    sSQL = sSQL & " ORDER BY SUBSTRING(C.cCtaCod,9,1), A.cAgeCod"

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetInteresesCreditosCastigados = rs
    Exit Function
GetICreditosCastigadosErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function
'******************************************************************
'JUEZ 20130116 ****************************************************************
Public Function GetGastosCreditosCastigados(ByVal psFecha As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetGastosCreditosCastigadosErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec sp_sel_ReporteGastosCreditosCastigadosCont"

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetGastosCreditosCastigados = rs
    Exit Function
GetGastosCreditosCastigadosErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Garantias Montos ")
End Function
'END JUEZ *********************************************************************
'JACA 20110503****************************************************************************
Public Function GetPlazoFijosDeOtros(ByVal psFecha As String, ByVal pnTpoCam As Currency) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetPlazoFijosDeOtrosErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_PlazoFijo_DeOtros '" & psFecha & "'," & pnTpoCam & ""
    
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetPlazoFijosDeOtros = rs
    Exit Function
GetPlazoFijosDeOtrosErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Buscar Plazo Fijo de Otras ")
End Function

Public Function GetPlazoFijoAnexo7(ByVal psFecha As String, ByVal pnTipoCambio As Currency) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetPlazoFijoAnexo7Err
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_PlazoFijo_Anexo7 '" & psFecha & "'," & pnTipoCambio & ""

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetPlazoFijoAnexo7 = rs
    Exit Function
GetPlazoFijoAnexo7Err:
    Call RaiseError(MyUnhandledError, "Obtener Plazo Fijo Anexo 7")
End Function
Public Function GetCaleAdeudadosXTramos(ByVal psFecha As String, ByVal pnMoneda As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCaleAdeudadosXTramosErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_CaleAdeudadosXTramos '" & psFecha & "','" & pnMoneda & "'"

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCaleAdeudadosXTramos = rs
    Exit Function
GetCaleAdeudadosXTramosErr:
    Call RaiseError(MyUnhandledError, "Obtener Calendario de Adeudados por Tramos")
End Function
'JACA END*********************************************************************************
'BRGO ************************************************************************************
Public Function GetCaleAdeudadosCapIntXTramos(ByVal psFecha As String, ByVal pnMoneda As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCaleAdeudadosXTramosErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_CaleAdeudadosCapIntXTramos '" & psFecha & "','" & pnMoneda & "'"

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCaleAdeudadosCapIntXTramos = rs
    Exit Function
GetCaleAdeudadosXTramosErr:
    Call RaiseError(MyUnhandledError, "Obtener Calendario de Adeudados por Tramos")
End Function
Public Function GetCapPrincipalesClientes(ByVal pnLimite As Double, ByVal pnTipoCambio As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCapPrincClientesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_ObtenerCapPrincipalesClientes " & pnLimite & ",'" & pnTipoCambio & "'"

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCapPrincipalesClientes = rs
    Exit Function
GetCapPrincClientesErr:
    Call RaiseError(MyUnhandledError, "Obtener Calendario de Adeudados por Tramos")
End Function
Public Function GetCapTotalDepositos(ByVal psFecha As String, ByVal pnTipoCambio As String) As Double
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
Dim nTotalDep As Double
On Error GoTo GetCapTotalDepositosErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "SELECT TotalDep = SUM(CASE WHEN nMoneda=1 THEN 1 ELSE " & pnTipoCambio & " END * nSaldo)" & _
            " FROM CapEstadSaldo " & _
            " WHERE DATEDIFF(DAY,dEstad,'" & psFecha & "')=0"

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    nTotalDep = rs!TotalDep
    oCon.CierraConexion
    GetCapTotalDepositos = nTotalDep
    Exit Function
GetCapTotalDepositosErr:
    Call RaiseError(MyUnhandledError, "Obtener Calendario de Adeudados por Tramos")
End Function
Public Function GetCapDepositosEntidadesPublicas(ByVal pnTipoCambio As String) As Double
    Dim sSQL As String
    Dim rs As New ADODB.Recordset
    Dim oCon As New DConecta
    On Error GoTo GetCapDepositosEntidPublicasErr
    Dim nTotalDep As Double
        oCon.AbreConexion
        Set rs = New ADODB.Recordset
    
        sSQL = "exec stp_sel_ObtenerDepositosEntidadesPublicas " & pnTipoCambio
    
        oCon.AbreConexion
        Set rs = oCon.CargaRecordSet(sSQL)
        nTotalDep = rs!nSaldo
        oCon.CierraConexion
        GetCapDepositosEntidadesPublicas = nTotalDep
        Exit Function
GetCapDepositosEntidPublicasErr:
        Call RaiseError(MyUnhandledError, "Obtener Calendario de Adeudados por Tramos")
End Function
'END BRGO *****************************************************************************
'EJVG20121002 ***
Public Function obtenerConsolidadoCtasxConcentracionFondos(ByVal pdFecha As Date, ByVal pnTpoMoneda As Integer, ByVal psListaTpoIfi As String, ByVal psListaCodPersIFiExcluidas As String) As ADODB.Recordset
    Dim lsSql As String
    Dim oCon As New DConecta
On Error GoTo ErrObtenerConsolidadoCtasxConcentracionFondos
    oCon.AbreConexion
    lsSql = "Exec stp_sel_ListaConsolCtaRptConcentracion '" & Format(pdFecha, "yyyymmdd") & "','" & pnTpoMoneda & "','" & psListaTpoIfi & "','" & psListaCodPersIFiExcluidas & "'"
    oCon.AbreConexion
    Set obtenerConsolidadoCtasxConcentracionFondos = oCon.CargaRecordSet(lsSql)
    oCon.CierraConexion
    Exit Function
ErrObtenerConsolidadoCtasxConcentracionFondos:
    Call RaiseError(MyUnhandledError, "Recupera Ctas Concentración")
End Function
Public Function obtenerDPFyOvernyInversion(ByVal pdFecha As Date, ByVal pnTpoMoneda As Integer, ByVal psListaTpoIfi As String, ByVal psListaCodPersIFiExcluidas As String) As ADODB.Recordset
    Dim lsSql As String
    Dim oCon As New DConecta
On Error GoTo ErrobtenerPFyOvernyInvers
    oCon.AbreConexion
    lsSql = "Exec stp_sel_ListaPFyOvernyInversion '" & Format(pdFecha, "yyyymmdd") & "','" & pnTpoMoneda & "','" & psListaTpoIfi & "','" & psListaCodPersIFiExcluidas & "'"
    oCon.AbreConexion
    Set obtenerDPFyOvernyInversion = oCon.CargaRecordSet(lsSql)
    oCon.CierraConexion
    Exit Function
ErrobtenerPFyOvernyInvers:
    Call RaiseError(MyUnhandledError, "Recupera PF y Overnight Concentración")
End Function

Public Function obtenerBCRInversionDPF(ByVal pdFecha As Date) As ADODB.Recordset
    Dim lsSql As String
    Dim oCon As New DConecta
On Error GoTo ErrobtenerBCRInversion
    oCon.AbreConexion
    lsSql = "Exec sp_ERS0792016_sel_ListarBCRPxDPF '" & Format(pdFecha, "yyyymmdd") & "'"
    oCon.AbreConexion
    Set obtenerBCRInversionDPF = oCon.CargaRecordSet(lsSql)
    oCon.CierraConexion
    Exit Function
ErrobtenerBCRInversion:
    Call RaiseError(MyUnhandledError, "Recupera PF y Overnight en BCRP")
End Function '*****NAGL ERS 079-2016 20170407

Public Function obtenerTotalAhorroEnCMACxFecha(ByVal pdFecha As Date, ByVal pnTpoMoneda As Integer, psCodPersIFI As String) As ADODB.Recordset
    Dim lsSql As String
    Dim oCon As New DConecta
    On Error GoTo ErrobtenerTotalAhorroEnCMACxFecha
    oCon.AbreConexion
    lsSql = "select  ISNULL(SUM(CASE WHEN SUBSTRING(CSD.cCtaCod,6,3)='232' THEN CSD.nSaldDisp ELSE 0 END),0) nSaldoCtaCte," & _
           "        ISNULL(SUM(CASE WHEN SUBSTRING(CSD.cCtaCod,6,3)='233' THEN CSD.nSaldDisp ELSE 0 END),0) nSaldoDPF " & _
           "From CapSaldosDiarios CSD " & _
           "Inner Join ProductoPersona PP on PP.cCtaCod=CSD.cCtaCod And PP.nPrdPersRelac=10 " & _
           "Where PP.cPersCod='" & psCodPersIFI & "'" & _
           "    And Datediff(d,csd.dFecha,'" & Format(pdFecha, "yyyymmdd") & "')=0" & _
           "    And SUBSTRING(CSD.cCtaCod,9,1)='" & pnTpoMoneda & "'"
    oCon.AbreConexion
    Set obtenerTotalAhorroEnCMACxFecha = oCon.CargaRecordSet(lsSql)
    oCon.CierraConexion
    Exit Function
ErrobtenerTotalAhorroEnCMACxFecha:
    Call RaiseError(MyUnhandledError, "Recupera Total Ahorros en CMAC")
End Function
'END EJVG *******
'MIOL 20121130, SEGUN RQ12349 ********************************************
Public Function GetCreditosVigentesSBSanx7(ByVal psFecha As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosVigentesSBSanx7Err
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec B2_stp_sel_creditosVigentesANX7 '" & psFecha & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCreditosVigentesSBSanx7 = rs
    Exit Function
GetCreditosVigentesSBSanx7Err:
    Call RaiseError(MyUnhandledError, "Creditos Vigentes")
End Function

Public Function GetCtaPlazoFijoSBSanx7(ByVal psFecha As String, ByVal pcMoneda As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCtaPlazoFijoSBSanx7Err
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_PasivoCuentasaPlazoANX7 '" & psFecha & "','" & pcMoneda & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCtaPlazoFijoSBSanx7 = rs
    Exit Function
GetCtaPlazoFijoSBSanx7Err:
    Call RaiseError(MyUnhandledError, "Interes Creditos Vigentes")
End Function

Public Function GetObligxCtaAhorrosSBSanx7(ByVal psFecha As String, ByVal pcMoneda As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetObligxCtaAhorrosSBSanx7Err
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_ObligacionesxCtaAhorrosANX7 '" & psFecha & "','" & pcMoneda & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetObligxCtaAhorrosSBSanx7 = rs
    Exit Function
GetObligxCtaAhorrosSBSanx7Err:
    Call RaiseError(MyUnhandledError, "Obligaciones x Cuentas de Ahorros")
End Function

Public Function GetDepositosIFIsSBSanx7(ByVal psFecha As String, ByVal pcMoneda As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetDepositosIFIsSBSanx7Err
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_DepositosxIFIsANX7 '" & psFecha & "','" & pcMoneda & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetDepositosIFIsSBSanx7 = rs
    Exit Function
GetDepositosIFIsSBSanx7Err:
    Call RaiseError(MyUnhandledError, "Depositos IFIs")
End Function

Public Function GetDepositosPFIFIsSBSanx7(ByVal psFecha As String, ByVal pcMoneda As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetDepositosPFIFIsSBSanx7Err
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_DepositosPFxIFIsANX7 '" & psFecha & "','" & pcMoneda & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetDepositosPFIFIsSBSanx7 = rs
    Exit Function
GetDepositosPFIFIsSBSanx7Err:
    Call RaiseError(MyUnhandledError, "Depositos IFIs")
End Function

Public Function GetCreditosCuentasaPFSBSanx7(ByVal psFecha As String, ByVal pcMoneda As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCreditosCuentasaPFSBSanx7Err
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_PasivoCuentasaPlazoANX7 '" & psFecha & "','" & pcMoneda & "','',''"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCreditosCuentasaPFSBSanx7 = rs
    Exit Function
GetCreditosCuentasaPFSBSanx7Err:
    Call RaiseError(MyUnhandledError, "Cuentas a Plazo Fijo")
End Function

Public Function GetCTSSBSanx7(ByVal psFecha As String, ByVal pcMoneda As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetCTSSBSanx7Err
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_PasivoCTSANX7 '" & psFecha & "','" & pcMoneda & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetCTSSBSanx7 = rs
    Exit Function
GetCTSSBSanx7Err:
    Call RaiseError(MyUnhandledError, "CTS")
End Function
'END MIOL ****************************************************************
'MIOL 20130212 RFC138-2012 ***********************************************
Public Function GetSaldoMEPosCambiaria(ByVal pdFecha As String, ByVal pcCtaCont As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetSaldoMEPosCambiariaErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_SaldosMEPosCambiaria '" & pdFecha & "','" & pcCtaCont & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetSaldoMEPosCambiaria = rs
    Exit Function
GetSaldoMEPosCambiariaErr:
    Call RaiseError(MyUnhandledError, "CTS")
End Function
'END MIOL *****************************************************************

'NAGL ERS 002-2017 20170803*********************************************************************
Public Function GetSaldoMEPosCambiariaNewYBalanc(ByVal pnTipoProc As String, ByVal pdFecha As Date, ByVal pcCtaCont As String, Optional lsMovNro As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetSaldoMEPosCambiariaNewYBalancErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset
    If pnTipoProc = "PC" Then
        sSQL = "exec stp_sel_SaldosMEPosCambiariaNew '" & Format(pdFecha, "yyyymmdd") & "','" & pcCtaCont & "', '" & lsMovNro & "'"
    Else
        sSQL = "exec stp_sel_SaldosMEPosCambiariaBalance '" & Format(pdFecha, "yyyymmdd") & "', '" & pcCtaCont & "'"
    End If
        oCon.AbreConexion
        Set rs = oCon.CargaRecordSet(sSQL)
        oCon.CierraConexion
        Set GetSaldoMEPosCambiariaNewYBalanc = rs
        Exit Function
GetSaldoMEPosCambiariaNewYBalancErr:
    Call RaiseError(MyUnhandledError, "PosCamNewBal")
End Function
Public Sub setSaldoMEPosCambiariaFormulas(ByVal pdFecha As Date, ByVal pcCtaCont As String, ByVal pnSaldoActual As Currency, ByVal pnSaldoAnterior As Currency, ByVal lsMovNro As String)
Dim sSQL As String
Dim oCon As New DConecta
    oCon.AbreConexion
    sSQL = " Exec stp_ins_SaldosMEPosCambiariaFormulas '" & Format(pdFecha, "yyyymmdd") & "','" & pcCtaCont & "'," & pnSaldoActual & "," & pnSaldoAnterior & ", '" & lsMovNro & "' "
    oCon.Ejecutar (sSQL)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'END NAGL *****************************************************************

'ALPA 20131220*************************************************************
Public Function GetObligxCtaAhorrosSBSanx6A(ByVal psFecha As String, ByVal psMoneda As String, ByVal psProducto As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetObligxCtaAhorrosSBSanx6AErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_ObligacionesxCtaAhorrosANX6A '" & psFecha & "','" & psMoneda & "', '" & psProducto & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetObligxCtaAhorrosSBSanx6A = rs
    Exit Function
GetObligxCtaAhorrosSBSanx6AErr:
    Call RaiseError(MyUnhandledError, "Obligaciones x Cuentas de Ahorros")
End Function

Public Function GetDepositosInmovilizadosSBSanx16A(ByVal psFecha As String, ByVal psMoneda As String, ByVal pnTipoCambio As Currency) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetDepositosInmovilizadosSBSanx16AErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset
    
    If (psMoneda = 1) Then
        sSQL = "exec sp_ERS0792016_sel_DepositosInmovilMNanx16A  '" & psFecha & "'"
        oCon.AbreConexion
        Set rs = oCon.CargaRecordSet(sSQL)
    Else
        sSQL = "exec sp_ERS0792016_sel_DepositosInmovilMEanx16A '" & psFecha & "'," & pnTipoCambio & ""
        oCon.AbreConexion
        Set rs = oCon.CargaRecordSet(sSQL)
    End If
    
    oCon.CierraConexion
    Set GetDepositosInmovilizadosSBSanx16A = rs
    Exit Function
GetDepositosInmovilizadosSBSanx16AErr:
    Call RaiseError(MyUnhandledError, "Depósitos Inmovilizados")
End Function '**NAGL ERS079-2016 20170407***'

Public Function GetObligacionesVistaSBSanx16A(ByVal psFecha As String, ByVal psMoneda As String, ByVal pnTipoCambio As Currency) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetObligacionesVistaSBSanx16AErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset
    
    If (psMoneda = 1) Then
        sSQL = "exec sp_ERS0792016_sel_ObligacionesVistaMNAnx16A  '" & psFecha & "'"
        oCon.AbreConexion
        Set rs = oCon.CargaRecordSet(sSQL)
    Else
        sSQL = "exec sp_ERS0792016_sel_ObligacionesVistaMEanx16A '" & psFecha & "'," & pnTipoCambio & ""
        oCon.AbreConexion
        Set rs = oCon.CargaRecordSet(sSQL)
    End If
    oCon.CierraConexion
    Set GetObligacionesVistaSBSanx16A = rs
    Exit Function
GetObligacionesVistaSBSanx16AErr:
    Call RaiseError(MyUnhandledError, "Obligaciones a la Vista")
End Function '**NAGL ERS079-2016 20170407***'

Public Function GetPlazoFijoRangosPersoneriaRangoAnexo6(ByVal psFecha As String, ByVal psFiltro As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo GetPlazoFijoRangosPersoneriaRangoAnexo6Err
    oCon.AbreConexion
    Set rs = New ADODB.Recordset
    sSQL = "exec stp_sel_PlazoFijoRangosPersoneriaRangoAnexo6 '" & Format(psFecha, "YYYY/MM/DD") & "','" & psFiltro & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetPlazoFijoRangosPersoneriaRangoAnexo6 = rs
    Exit Function
GetPlazoFijoRangosPersoneriaRangoAnexo6Err:
    Call RaiseError(MyUnhandledError, "Plazo Fijo y Rangos de Personeria del Anexo 6")
End Function
'**************************************************************************
'EJVG20131230 ***
Public Function ObtieneEfectivoDiaxRptConcentracionGastos(ByVal pdFecha As Date, ByVal pnMoneda As Moneda) As Currency
    Dim lsSql As String
    Dim oCon As New DConecta
    Dim rs As New ADODB.Recordset
On Error GoTo ErrobtenerPFyOvernyInvers
    If pnMoneda = gMonedaNacional Then
        lsSql = "Exec stp_sel_EfectivoDiarioMNxRptConcentraGastos '" & Format(pdFecha, "yyyymmdd") & "'"
    Else
        lsSql = "Exec stp_sel_EfectivoDiarioMExRptConcentraGastos '" & Format(pdFecha, "yyyymmdd") & "'"
    End If
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(lsSql)
    If Not rs.EOF Then
        ObtieneEfectivoDiaxRptConcentracionGastos = rs!nMonto
    End If
    oCon.CierraConexion
    Set rs = Nothing
    Set oCon = Nothing
    Exit Function
ErrobtenerPFyOvernyInvers:
    Call RaiseError(MyUnhandledError, "Recupera Efectivo Diario x Reporte Concentración Gastos")
End Function
'END EJVG *******
'PASI20160127 **
Public Function CargaActivosDisponiblesANX07(ByVal psFecha As String, ByVal psMoneda As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim oCon As New DConecta

On Error GoTo ERRORCargaActivosDisponiblesANX07
    oCon.AbreConexion
    Set rs = New ADODB.Recordset
    
    sSQL = "stp_sel_ObtieneDisponiblexAnx07 '" & psFecha & "','" & psMoneda & "'"
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set CargaActivosDisponiblesANX07 = rs
    Exit Function
ERRORCargaActivosDisponiblesANX07:
 Call RaiseError(MyUnhandledError, "Creditos Disponibles")
End Function
Public Function CargaActivosInversionesDispxVentANX07(ByVal psFecha As String, ByVal psMoneda As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim oCon As New DConecta

On Error GoTo ERRORCargaActivosInversionesDispxVentANX07
    oCon.AbreConexion
    Set rs = New ADODB.Recordset
    
    sSQL = "stp_sel_InversionesDispxVenta '" & psFecha & "','" & psMoneda & "'"
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set CargaActivosInversionesDispxVentANX07 = rs
    Exit Function
ERRORCargaActivosInversionesDispxVentANX07:
 Call RaiseError(MyUnhandledError, "Inversiones Disponibles para la Venta")
End Function
Public Function CargaActivosCtasxCobrarSensiblesANX07(ByVal psFecha As String, ByVal psMoneda As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim oCon As New DConecta

On Error GoTo ERRORCargaActivosCtasxCobrarSensiblesANX07
    oCon.AbreConexion
    Set rs = New ADODB.Recordset
    
    sSQL = "stp_sel_CuentasxCobrarSensibles '" & psFecha & "','" & psMoneda & "'"
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set CargaActivosCtasxCobrarSensiblesANX07 = rs
    Exit Function
ERRORCargaActivosCtasxCobrarSensiblesANX07:
 Call RaiseError(MyUnhandledError, "Cuentas Sensibles x Cobrar")
End Function
'END PASI**


'*** PEAC 20201127
Public Function GetReporte35(ByVal psMes As String, ByVal psAnio As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta

On Error GoTo GetReporte35Err
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_Reporte35 '" & psAnio & "','" & Right("00" & psMes, 2) & "' "

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetReporte35 = rs
    Exit Function
GetReporte35Err:
    Call RaiseError(MyUnhandledError, "GetReporte35")
End Function

'*** PEAC 20201127
Public Function GetExisteRCD_Reporte35(ByVal psMes As Integer, ByVal psAnio As Integer) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta

On Error GoTo GetExisteRCD_Reporte35Err
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_VerificaExisteRCD '" & psAnio & "','" & Right("00" & psMes, 2) & "' "

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetExisteRCD_Reporte35 = rs
    Exit Function
GetExisteRCD_Reporte35Err:
    Call RaiseError(MyUnhandledError, "GetExisteRCD_Reporte35")
End Function

'*** PEAC 20210503
Public Function GetAnexoRepo35(ByVal psMes As Integer, ByVal psAnio As Integer, ByVal pnTC As Double) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta

On Error GoTo GetAnexoRepo35Err
    oCon.AbreConexion
    Set rs = New ADODB.Recordset

    sSQL = "exec stp_sel_DetalleIntDevPerAnexo '" & psAnio & "','" & Right("00" & psMes, 2) & "', " & pnTC

    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set GetAnexoRepo35 = rs
    Exit Function
GetAnexoRepo35Err:
    Call RaiseError(MyUnhandledError, "GetAnexoRepo35")
End Function

