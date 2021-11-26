VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCajaGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0

Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String

Dim vsFormatoFecha As String
Dim oError As COMConecta.COMErrorHandling

Public Function GetCtaIfFiltro(ByVal psCtaContCod As String, psCtaIfCod As String, psPersCod As String, Optional lbMuestraCta As Boolean = True) As String
Dim oCtaIf As COMDCajaGeneral.DCOMCajaCtasIF
Set oCtaIf = New COMDCajaGeneral.DCOMCajaCtasIF
GetCtaIfFiltro = oCtaIf.GetCtaIfFiltro(psCtaContCod, psCtaIfCod, psPersCod, lbMuestraCta)
Set oCtaIf = Nothing
End Function

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
Dim oCtasIF As COMDCajaGeneral.DCOMCajaCtasIF
Set oCtasIF = New COMDCajaGeneral.DCOMCajaCtasIF
NombreIF = oCtasIF.NombreIF(psPersCodIF)
Set oCtasIF = Nothing
End Function

Public Function SubCuentaIF(ByVal psPersCodIF As String) As String
Dim oCtasIF As COMDCajaGeneral.DCOMCajaCtasIF
Set oCtasIF = New COMDCajaGeneral.DCOMCajaCtasIF
SubCuentaIF = oCtasIF.SubCuentaIF(psPersCodIF)
Set oCtasIF = Nothing

End Function

Public Function CargaCtasIF(psMoneda As COMDConstantes.Moneda, Optional psFiltroCtaIF As String, Optional pnMuestraIF As COMDConstantes.MuestraIF = COMDConstantes.MuestraCuentas, Optional psCanje As String = "") As ADODB.Recordset

Dim oCtasIF As COMDCajaGeneral.DCOMCajaCtasIF
Set oCtasIF = New COMDCajaGeneral.DCOMCajaCtasIF
Set CargaCtasIF = oCtasIF.CargaCtasIF(psMoneda, psFiltroCtaIF, pnMuestraIF, psCanje)
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
    Call oError.RaiseError(oError.MyUnhandledError, "NCOMCajaCtaIF:EmiteTipoCuentaIF Method")
End Function
Private Function DescTipoCtaIF(psTpoCtaIF As String) As String

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
    rs.Close
     
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
    rs.Close
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Private Sub Class_Initialize()
    Dim oImp As COMDConstSistema.DCOMImpresoras
    Set oImp = New COMDConstSistema.DCOMImpresoras
    Dim oImpresora As New COMFunciones.FCOMVarImpresion
    
    Set oError = New COMConecta.COMErrorHandling
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni  As COMConecta.DCOMClasIni
Set oIni = New COMConecta.DCOMClasIni

vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
End Sub

Public Function GetInstFinancieras(Optional ByVal psFiltroTipoCtaIF As String = "") As Recordset
Dim oConec As COMConecta.DCOMConecta
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsIFFiltro As String
Dim lsFiltroCta As String
Dim pos As String
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

Public Function GetCtasInstFinancieras(Optional ByVal psFiltroTipoCtaIF As String = "", Optional ByVal psCtaCont As String = "", Optional pnNivel As COMDConstantes.MuestraIF = COMDConstantes.MuestraCuentas, Optional ByVal psFiltroTipoIF As String) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsIFFiltro As String
Dim lsFiltroCta As String
Dim pos As String
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
    & "             SELECT  CI.CPERSCOD + '.' + CI.cCtaIFCod, CONVERT(CHAR(40),CI.cCtaIFDesc) AS CTAIFDESC, " _
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

Public Function GetVerificaSubCuentaIF(ByVal psPersCod As String, ByVal pnTipoIF As COMDConstantes.CGTipoIF, _
                                        ByVal pnTipoCtaIf As COMDConstantes.CGTipoCtaIF, ByVal pnMoneda As COMDConstantes.Moneda) As Boolean
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

Public Function GetSubCuentaIF(ByVal psPersCod As String, ByVal pnTipoIF As COMDConstantes.CGTipoIF, _
                                        ByVal pnTipoCtaIf As COMDConstantes.CGTipoCtaIF, ByVal pnMoneda As COMDConstantes.Moneda) As String
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
        GetSubCuentaIF = Trim(rs!cCtaIfCod)
    Else
        GetSubCuentaIF = Format(pnTipoCtaIf, "00") & pnMoneda & "00"
    End If
    rs.Close
    Set rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GeneraSubCuentaIF(ByVal psPersCod As String, ByVal pnTipoIF As COMDConstantes.CGTipoIF, _
                                  ByVal psCtaCont As String, ByVal psCtaIfCod As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSQL = "Select  cCtaIfCod, C.cCtaContCod, cCtaIFSubCta,C.cCtaContCod + C.cCtaIFSubCta" _
        & " From    CtaIFFiltro  C   " _
        & "         JOIN CTACONT CC ON CC.CCTACONTCOD  = C.cCtaContCod+ C.cCtaIFSubCta " _
        & " Where   C.cCtaContCod Like '" & psCtaCont & "' " _
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

Public Function GetSaldoCtaIf(ByVal psPersCod As String, ByVal pnTipoIF As COMDConstantes.CGTipoIF, _
                             ByVal psCtaIfCod As String, ByVal pdFecha As Date) As Currency
    
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Set rs = New ADODB.Recordset
    Dim oFormat As New COMFunciones.FCOMVarPublicas
    
    Set oConec = New COMConecta.DCOMConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSQL = "Select  MS.cPersCod, MS.cIFTpo, MS.cCtaIFCod, nSaldo" _
        & " From    CtaIFSaldo MS " _
        & " Where   MS.dCtaIFSaldo = (  Select  Max(MS1.dCtaIFSaldo) " _
        & "                             From    CtaIFSaldo MS1 " _
        & "                             Where   MS.cPersCod = MS1.cPersCod AND " _
        & "                                     MS.cIFTpo = MS1.cIFTpo And MS.cCtaIFCod = MS1.cCtaIFCod " _
        & "                                     AND MS1.dCtaIFSaldo <='" & Format(pdFecha, oFormat.gsFormatoFecha) & "') " _
        & "         AND MS.cPersCod='" & psPersCod & "' AND MS.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         AND MS.cCtaIFCod='" & psCtaIfCod & "'"
        
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

Public Function GetInteresDevengados(ByVal psOpeCod As String, ByVal pdFecha As Date, Optional ByVal pnOpeObjOrden As Integer = 1) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim oOpe   As NCOMCajaCtaIF
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsCtaCapital As String

Set oOpe = New NCOMCajaCtaIF
Set oConec = New COMConecta.DCOMConecta
Dim oFormat As New COMFunciones.FCOMVarPublicas

Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

lsCtaCapital = oOpe.EmiteOpeCta(psOpeCod, "H", 2)
Set oOpe = Nothing

sql = " Select   cCtaIFDesc, cInstFinanc,nCapital ," _
    & "         IntValor, nInteres, Convert(CHAR(12),UltFecha,103) UltFecha, Dias, " _
    & "         CASE WHEN Periodo >0 Then  (POWER((1 + IntPorc*1.0),(Dias*1.0/Periodo*1.0))-1)*(nCapital+nInteres) ELSE 0 END as IntCalculado, " _
    & "         cPersCod , cIFTpo, cCtaIFCod " _
    & " From  ( " _
    & "         SELECT  CI.cPersCod, CI.cIFTpo, CI.cCtaIFCod, CI.cCtaIFDesc, cPersNombre AS cInstFinanc , " _
    & "                 CI.dCtaIfCap, CI.nInteres , ISNULL(Saldo.nSaldo,0) AS nCapital, " _
    & "                 IsNull(INTERES.nCtaIFIntPeriodo,0) as Periodo, " _
    & "                 IsNull(INTERES.nCtaIFIntValor,0) as IntValor, " _
    & "                 IsNull(INTERES.nCtaIFIntValor,0)/100 as IntPorc, " _
    & "                 CASE WHEN CI.dCtaIfInt >= CI.dCtaIfCap THEN  CI.dCtaIfInt " _
    & "                      ELSE  CI.dCtaIfCap END  as UltFecha , " _
    & "                 DateDiff(day,  CASE WHEN CI.dCtaIfInt >= CI.dCtaIfCap THEN CI.dCtaIfInt " _
    & "                                 ELSE  CI.dCtaIfCap END, '" & Format(pdFecha, oFormat.gsFormatoFecha) & "') AS Dias " _
    & "          FROM   CTAIF CI " _
    & "                 JOIN INSTITUCIONFINANC I ON I.cPersCod = CI.cPersCod AND I.cIFTpo = CI.cIFTpo " _
    & "                 JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD " _
    & "                 LEFT JOIN ( SELECT  CS.dCtaIFSaldo, CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod, CS.nSaldo " _
    & "                             FROM    CTAIFSALDO CS WHERE cs.cCtaContCod LIKE '" & lsCtaCapital & "%'" _
    & "                                 and CS.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
    & "                                                         FROM    CTAIFSALDO CS1 " _
    & "                                                         Where   cs1.cCtaContCod = cs.cCtaContCod and CS1.cPersCod = CS.cPersCod And CS1.cIFTpo = CS.cIFTpo And CS1.cCtaIFCod = CS.cCtaIFCod " _
    & "                                                                 AND CS1.dCtaIFSaldo<='" & Format(pdFecha, oFormat.gsFormatoFecha) & "')) AS SALDO " _
    & "                 ON SALDO.cPersCod=CI.cPersCod and SALDO.cIFTpo = CI.cIFTpo and SALDO.cCtaIFCod= CI.cCtaIFCod " _
    & "                 LEFT JOIN ( SELECT  I.cPersCod, I.cIFTpo, I.cCtaIFCod, I.dCtaIFIntRegistro, I.nCtaIFIntPeriodo, I.nCtaIFIntValor "
sql = sql + "                       FROM    CtaIFInteres I " _
    & "                             WHERE   I.dCtaIFIntRegistro = ( SELECT  MAX(dCtaIFIntRegistro) " _
    & "                                                             FROM    CtaIFInteres I1 " _
    & "                                                             Where   I1.cPersCod = I.cPersCod And I1.cIFTpo = I.cIFTpo And I1.cCtaIFCod = I.cCtaIFCod " _
    & "                                                                     AND I1.dCtaIFIntRegistro<='" & Format(pdFecha, oFormat.gsFormatoFecha) & "') ) AS INTERES " _
    & "                 ON INTERES.cPersCod=CI.cPersCod and INTERES.cIFTpo = CI.cIFTpo and INTERES.cCtaIFCod= CI.cCtaIFCod, OPEOBJ O " _
    & "         WHERE   LEN(CI.cCtaIFCod)=(Select Max(Len(cCtaIFCod)) From CtaIf) AND cCtaIFEstado ='" & gEstadoCtaIFActiva & "' " _
    & "                 AND o.cOpeCod = '" & psOpeCod & "' and cOpeObjOrden='" & pnOpeObjOrden & "' and " _
    & "                 CI.cIFTpo+CI.cCtaIFCod Like cOpeObjFiltro And Saldo.nSaldo>0  " _
    & "         ) AS DATOS " _
    & " Order By cInstFinanc "

Set rs = oConec.CargaRecordSet(sql)
Set GetInteresDevengados = rs
oConec.CierraConexion
Set oConec = Nothing
End Function
Public Function GetCuentasIFCanc(ByVal psOpeCod As String, ByVal pdFecha As Date, Optional ByVal pnOpeObjOrden As Integer = 1) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New COMConecta.DCOMConecta
Dim oFormat As New COMFunciones.FCOMVarPublicas

Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = "SELECT   CI.cCtaIFDesc, cPersNombre AS cInstFinanc , " _
    & "         ISNULL(Saldo.nSaldo,0) AS nCapital, 0 as ValorInt, CI.nInteres, " _
    & "         CI.dCtaIFAper, 0 Dias , 0 as IntCalculado, CI.cPersCod, CI.cIFTpo, CI.cCtaIFCod " _
    & " FROM    CTAIF CI " _
    & "         JOIN INSTITUCIONFINANC I ON I.cPersCod = CI.cPersCod AND I.cIFTpo = CI.cIFTpo " _
    & "         JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD " _
    & "         LEFT JOIN (     SELECT  CS.dCtaIFSaldo, CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod, CS.nSaldo " _
    & "                         FROM    CTAIFSALDO CS " _
    & "                         WHERE   CS.dCtaIFSaldo = (      SELECT  MAX(dCtaIFSaldo) " _
    & "                                                         FROM    CTAIFSALDO CS1 " _
    & "                                                         Where CS1.cPersCod = CS.cPersCod And CS1.cIFTpo = CS.cIFTpo And CS1.cCtaIFCod = CS.cCtaIFCod " _
    & "                                                         AND CS1.dCtaIFSaldo<='" & Format(pdFecha, oFormat.gsFormatoFecha) & "')) AS SALDO " _
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

sql = " SELECT  C.dCtaIFAper, C.dCtaIfCap, C.dCtaIfInt, C.nCtaIFPlazo, C.nInteres, C.cCtaIFEstado, CC.cConsDescripcion AS cEstadoCons, Isnull(SALDO.nSaldo,0) as Capital" _
    & " FROM    CTAIF C " _
    & "         JOIN CONSTANTE CC ON    CC.NCONSVALOR =  C.cCtaIFEstado  " _
    & "         LEFT JOIN ( SELECT  CS.dCtaIFSaldo, CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod, CS.nSaldo " _
    & "                     FROM    CTAIFSALDO CS " _
    & "                     WHERE   CS.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
    & "                                                 FROM    CTAIFSALDO CS1 " _
    & "                                                 Where   CS1.cPersCod = CS.cPersCod And CS1.cIFTpo = CS.cIFTpo And CS1.cCtaIFCod = CS.cCtaIFCod )) AS SALDO " _
    & "                     ON SALDO.cPersCod=C.cPersCod and SALDO.cIFTpo = C.cIFTpo and SALDO.cCtaIFCod= C.cCtaIFCod " _
    & " Where CC.NCONSCOD = " & gCGEstadoCtaIF & " AND C.cPersCod='" & psPersCod & "' and C.cIFTpo ='" & Format(pnIFTpo, "00") & "' and C.cCtaIFCod='" & psCtaIfCod & "'"

Set rs = oConec.CargaRecordSet(sql)
Set GetDatosCtaIf = rs
oConec.CierraConexion
Set oConec = Nothing
End Function
Public Function ActualizaCtas(ByVal psFormatoFecha As String, ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, _
                                ByVal psCtaIfCod As String, ByVal psCtaIFDesc As String, _
                                ByVal pdCtaIFAper As Date, ByVal pdCtaIFCap As Date, _
                                ByVal pdCtaIFInt As Date, ByVal pnCtaIFPlazo As Integer, _
                                ByVal pnInteres As Currency, ByVal pnCtaIFEstado As CGEstadoCtaIF, _
                                ByVal psMovUltAct As String, ByVal rsInt As ADODB.Recordset) As Integer
Dim oMov As COMDMov.DCOMMov
Dim lbTrans As Boolean
Dim lnMovNro As String
Set oMov = New COMDMov.DCOMMov
On Error GoTo ErrorActualizaCtas
Dim oFormat As New COMFunciones.FCOMVarPublicas

oMov.BeginTrans
lbTrans = True
oMov.ActualizaCtaIF psPersCod, pnIFTpo, psCtaIfCod, Trim(psCtaIFDesc), _
                    Format(pdCtaIFAper, oFormat.gsFormatoFecha), Format(pdCtaIFCap, oFormat.gsFormatoFecha), Format(pdCtaIFInt, oFormat.gsFormatoFecha), pnCtaIFPlazo, pnCtaIFEstado, pnInteres, psMovUltAct
If Not rsInt Is Nothing Then
    Do While Not rsInt.EOF
        If Val(rsInt!lbNuevo) <> 1 Then
            oMov.InsertaCuentaIFInteres psPersCod, pnIFTpo, psCtaIfCod, rsInt!Registro, rsInt![Período], rsInt!Interes, psMovUltAct
        End If
        rsInt.MoveNext
    Loop
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

Public Function GetFlujosCtaIF(ByVal psOpeCod As String, _
                            ByVal psCtaIfCodDesde As String, ByVal psCtaIfCodHasta As String, _
                            ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = " SELECT  CI.cPersCod, CI.cIFTpo, CI.cCtaIFCod, P.CPERSNOMBRE, CI1.cCtaIFDesc, CI.cCtaIFDesc," _
    & "         RTRIM(P.CPERSNOMBRE) + ' ' + CI1.cCtaIFDesc + ' ' + CI.cCtaIFDesc AS cNomCtaIF, " _
    & "         ISNULL(SALDO.nSaldo,0) AS nSaldoIni, " _
    & "         CI.nCtaIFPlazo, CI.nInteres , " _
    & "         M.CMOVNRO, M.CMOVDESC, IsNull(D.cDocAbrev,'')  + ' ' + IsNull(MD.cDocNro,'') as Documento , " _
    & "         MO.nMovImporte ,  " _
    & "         Case WHEN MO.nMovImporte > 0 THEN ABS(MO.nMovImporte) ELSE 0 END AS DEBE , " _
    & "         Case WHEN MO.nMovImporte < 0 THEN ABS(MO.nMovImporte) ELSE 0 END AS HABER " _
    & " FROM    MOV M " _
    & "        JOIN MovCta mo ON mo.nMovNro = m.nMovNro " _
    & "        JOIN MOVOBJIF MC ON MC.nMovNro = mo.nMovNro and mc.nMovItem = mo.nMovItem " _
    & "        LEFT JOIN MOVDOC MD ON MD.NMOVNRO = M.NMOVNRO AND MD.nDocTpo <>" & TpoDocVoucherEgreso & "  " _
    & "        JOIN CTAIF CI ON MC.cPersCod=CI.cPersCod and MC.cIFTpo = MC.cIFTpo and MC.cCtaIFCod= CI.cCtaIFCod " _
    & "        JOIN CTAIF CI1 ON CI1.cPersCod = CI.cPersCod AND CI1.cIFTpo = CI.cIFTpo  AND SUBSTRING(CI.cCtaIFCod,1,LEN(CI1.cCtaIFCod)) = CI1.cCtaIFCod and Len(CI1.cCtaIFCod)=5 " _
    & "        JOIN INSTITUCIONFINANC I ON I.cPersCod = CI.cPersCod AND I.cIFTpo = CI.cIFTpo " _
    & "        JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD " _
    & "         LEFT JOIN ( SELECT  CS.dCtaIFSaldo, CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod, CS.nSaldo " _
    & "                     FROM    CTAIFSALDO CS " _
    & "                     WHERE   CS.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
    & "                                                 FROM    CTAIFSALDO CS1 " _
    & "                                                 Where   CS1.cPersCod = CS.cPersCod And CS1.cIFTpo = CS.cIFTpo " _
    & "                                                         And CS1.cCtaIFCod = CS.cCtaIFCod " _
    & "                                                         AND CS1.dCtaIFSaldo<='" & Format(DateAdd("d", -1, pdDesde), "mm/dd/yyyy") & "' )) AS SALDO" _
    & "         ON SALDO.cPersCod=CI.cPersCod and SALDO.cIFTpo = CI.cIFTpo and SALDO.cCtaIFCod= CI.cCtaIFCod " _
    & "         LEFT JOIN DOCUMENTO D ON D.nDocTpo = MD.nDocTpo, OPEOBJ O "
sql = sql + "WHERE   LEN(CI.cCtaIFCod)= (Select max(len(cCtaIFCod)) From CtaIf) AND CI.cCtaIFEstado ='1' " _
    & "         AND o.cOpeCod = '" & psOpeCod & "' and cOpeObjOrden='0' " _
    & "         AND CI.cIFTpo + CI.cCtaIFCod like cOpeObjFiltro " _
    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "         AND CI.cPersCod + '.' + CI.cCtaIFCod BETWEEN '" & psCtaIfCodDesde & "' AND '" & psCtaIfCodHasta & "' " _
    & "         and M.NMOVFLAG IN (" & gMovFlagVigente & "," & gMovFlagDeExtorno & ") " _
    & "         and mo.cCtaContCod LIKE '11_[23]%' " _
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
Dim oFormat As New COMFunciones.FCOMVarPublicas

Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = " SELECT  Z.cCtaIfCod , Z.cCtaIFDesc, Z.cBancoDesc, Z.cCtaIFEstado, " _
    & "         CASE    WHEN SUM(z.nImporte) > 0 THEN str(SUM(z.nImporte), 14, 2) " _
    & "                 WHEN SUM(z.nImporte) < 0 THEN '0' END as Debe, " _
    & "         CASE    WHEN SUM(z.nImporte) < 0 THEN str(SUM(z.nImporte) * -1, 14, 2) " _
    & "                 WHEN SUM(z.nImporte) > 0 THEN '0' END as Haber " _
    & " FROM (  " _
    & "         SELECT  CtaIf.cCtaIfCod, CtaIF.cCtaIFDesc, CtaIF.cBancoDesc, CtaIf.cCtaIFEstado, SUM(CI.NSALDO) AS nImporte " _
    & "         FROM    CTAIFSALDO CI " _
    & "                 JOIN " _
    & "                 (   SELECT  CI.cIFTpo + CI.cPersCod + CI.cCtaIFCod AS cCtaIfCod, CI.cCtaIFDesc, P.cPersNombre cBancoDesc, CI.cCtaIFEstado " _
    & "                     FROM    CTAIF CI JOIN INSTITUCIONFINANC I ON  I.CPERSCOD = CI.CPERSCOD  AND CI.CIFTPO = I.CIFTPO " _
    & "                             JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD, OPEOBJ O " _
    & "                     WHERE   O.cOpeCod = '" & psOpeCod & "' and O.cOpeObjOrden IN (" & psOpeObjOrden & ") and CI.cCtaIFEstado IN ('" & psCtaIFEstado & "') " _
    & "                             AND CI.cIFTpo + CI.cCtaIFCod like O.cOpeObjFiltro " _
    & "                 ) as CtaIf ON  SUBSTRING(CI.cIFTpo + CI.cPersCod + CI.cCtaIFCod , 1, LEN(CtaIf.cCtaIfCod)) = CtaIf.cCtaIfCod , OPEOBJ O JOIN OpeCta oc ON oc.cOpeCod = o.cOpeCod " _
    & "         WHERE   O.cOpeCod = '" & psOpeCod & "' and ci.cCtaContCod LIKE oc.cCtaContCod + '%' and oc.cOpeCtaOrden IN (" & psOpeObjOrden & ") and O.cOpeObjOrden IN (" & psOpeObjOrden & ") " _
    & "                 AND CI.cIFTpo + CI.cCtaIFCod Like O.cOpeObjFiltro " & IIf(psMoneda = "", "", " and CI.cCtaIFCod LIKE '__" & psMoneda & "%' ")
sql = sql + "           AND CI.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
    & "                                         FROM    CTAIFSALDO CS1 " _
    & "                                         Where   CS1.cPersCod = CI.cPersCod And CS1.cIFTpo = CI.cIFTpo And CS1.cCtaIFCod = CI.cCtaIFCod " _
    & "                                                 AND CS1.dCtaIFSaldo<='" & Format(pdFecha, oFormat.gsFormatoFecha) & "') " _
    & "         GROUP   BY CtaIf.cCtaIfCod , CtaIF.cCtaIFDesc, CtaIF.cBancoDesc, CtaIF.cCtaIFEstado ) AS Z " _
    & " WHERE nImporte <> 0 " _
    & " GROUP  BY Z.cCtaIfCod , Z.cCtaIFDesc, Z.cBancoDesc, Z.cCtaIFEstado " _
    & " ORDER BY  Z.cCtaIFEstado, Z.cBancoDesc, Z.cCtaIfCod"

Set rs = oConec.CargaRecordSet(sql)
Set GetRepSaldosCtaIf = rs
oConec.CierraConexion
Set oConec = Nothing

End Function


'*********** Antes se encontraba en la Clase DOperacion
'***************** Comentar temporalmente hasta contar con las modificaciones de la DLL Constante

Public Function GetOpeObj(ByVal psOpeCod As String, ByVal psOpeObjOrden As String, _
                        Optional psCtaContCod As String = "", _
                        Optional ByVal psFiltro As String, Optional ByVal psFiltroAdd As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta
''Dim oRHAreas As DActualizaDatosArea
Dim oContFunct As COMNContabilidad.NCOMContFunciones
Dim oCtaIf As NCOMCajaCtaIF
Dim oEfect As COMDCajaGeneral.DCOMEfectivo
Dim lsFiltro As String
Dim lsFiltroIf As String
Dim lsFiltroCta As String

Set oCtaIf = New NCOMCajaCtaIF
Set oContFunct = New COMNContabilidad.NCOMContFunciones
Set oConect = New COMConecta.DCOMConecta
''Set oRHAreas = New DActualizaDatosArea
Set oEfect = New COMDCajaGeneral.DCOMEfectivo

Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
If oConect.AbreConexion = False Then
    Exit Function
End If
sql = " SELECT cOpeCod, cOpeObjOrden, cObjetoCod, cOpeObjFiltro, nOpeObjNiv " _
    & " FROM  OPEOBJ " _
    & " WHERE COPECOD ='" & psOpeCod & "' and cOpeObjOrden='" & psOpeObjOrden & "'"
    
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    If psFiltro = "" Then
        lsFiltro = Trim(rs!cOpeObjFiltro) & psFiltroAdd
    Else
        lsFiltro = psFiltro & psFiltroAdd
    End If
    Select Case Val(rs!cObjetoCod)
            Case ObjCMACAgencias
                ''Set rs1 = oRHAreas.GetAgencias(lsFiltro)
            Case ObjCMACAgenciaArea
                ''Set rs1 = oRHAreas.GetAgenciasAreas(lsFiltro)
            Case ObjCMACArea
                ''Set rs1 = oRHAreas.GetAreas(lsFiltro)
            Case ObjEntidadesFinancieras
                lsFiltroIf = lsFiltro
                lsFiltroCta = lsFiltro
                If rs!nOpeObjNiv = 1 Then
                    Set rs1 = oCtaIf.GetInstFinancieras(lsFiltro)
                Else
                    Set rs1 = oCtaIf.CargaCtasIF(Mid(psOpeCod, 3, 1), lsFiltro, rs!nOpeObjNiv)
                End If
            Case ObjDescomEfectivo
                Set rs1 = oEfect.GetBilletajes(lsFiltro)
            Case ObjPersona
                Set rs1 = Nothing
            Case Else
                'Set rs1 = oContFunct.GetObjetos(Val(IIf(lsFiltro = "", rs!cObjetoCod, rs!cOpeObjFiltro)), IIf(rs!cOpeObjFiltro = "", true, false))
                Set rs1 = oContFunct.GetObjetos(Val(IIf(lsFiltro = "", rs!cObjetoCod, rs!cOpeObjFiltro)), False)
        End Select
End If
rs.Close
Set rs = Nothing
oConect.CierraConexion: Set oConect = Nothing

Set oCtaIf = Nothing
Set oContFunct = Nothing
Set oConect = Nothing
''Set oRHAreas = Nothing
Set oEfect = Nothing

Set GetOpeObj = rs1

End Function

Public Function GrabaExtornoMov(ByVal pdFecSis As Date, ByVal pdFechaMov As Date, ByVal psMovNroExt As String, _
                                  ByVal pnMovNroAnt As Long, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                  ByVal pnImporte As Currency)

On Error GoTo ErrorGrabaExtornoMov
Dim oMov As COMDMov.DCOMMov
Dim lbTrans As Boolean
Set oMov = New COMDMov.DCOMMov
oMov.BeginTrans
lbTrans = True
    oMov.ExtornaMovimiento psMovNroExt, pnMovNroAnt, psOpeCod, psMovDesc
oMov.CommitTrans
lbTrans = False
Set oMov = Nothing
Exit Function
ErrorGrabaExtornoMov:
    If lbTrans Then oMov.RollbackTrans
    Err.Raise Err.Number, "GrabaExtornoMov", Err.Description
End Function

