VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCajaCtaIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Dim vsConexion As String
Dim vsServerCom As String
Dim vsServerPers As String

Dim vsFormatoFecha As String


Public Function GetCtaIfFiltro(ByVal psCtaContCod As String, psCtaIfCod As String, psPersCod As String, Optional lbMuestraCta As Boolean = True) As String
Dim oCtaIf As DCajaCtasIF
Set oCtaIf = New DCajaCtasIF
GetCtaIfFiltro = oCtaIf.GetCtaIfFiltro(psCtaContCod, psCtaIfCod, psPersCod, lbMuestraCta)
Set oCtaIf = Nothing
End Function
Public Function GeneraCodCtaIF(psCtaIfCod As String, psMoneda As String, psTpoCta As String) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oconect As DConecta
    
    Set rs = New ADODB.Recordset
    Set oconect = New DConecta
    GeneraCodCtaIF = ""
    If oconect.AbreConexion = False Then Exit Function
    
    sql = "Select   Max(cCtaIfCod) as MaxCtaIF " _
        & " From    CTAIF " _
        & " Where   cPersCod='" & psCtaIfCod & "' " _
        & "         and Substring(cCtaIfCod,1,2)='" & psTpoCta & "' and Substring(cCtaIfCod,3,1)='" & psMoneda & "' "

    Set rs = oconect.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        If IsNull(rs!MaxCtaIF) Then
            GeneraCodCtaIF = psTpoCta + psMoneda + Format(1, String(7 - Len(psTpoCta + psMoneda), "0"))
        Else
            GeneraCodCtaIF = Format(IIf(IsNull(rs!MaxCtaIF), 0, rs!MaxCtaIF) + 1, String(Len(rs!MaxCtaIF), "0"))
        End If
    End If
    rs.Close
    Set rs = Nothing
    oconect.CierraConexion
    Set oconect = Nothing
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
Public Function CargaCtasIF(psMoneda As Moneda, Optional psFiltroCtaIF As String, Optional pnMuestraIF As MuestraIF = MuestraCuentas, Optional psCanje As String = "") As ADODB.Recordset
Dim oCtasIF As DCajaCtasIF
Set oCtasIF = New DCajaCtasIF
Set CargaCtasIF = oCtasIF.CargaCtasIF(psMoneda, psFiltroCtaIF, pnMuestraIF, psCanje)
Set oCtasIF = Nothing
End Function
Public Function EmiteTipoCuentaIF(psCodCtaIf As String) As String
    On Error GoTo EmiteTipoCuentaIFErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oconect As DConecta
    
    EmiteTipoCuentaIF = ""
    Set oconect = New DConecta
    If oconect.AbreConexion = False Then Exit Function
    
    Set rs = New ADODB.Recordset
    sql = "Select * from " & vsServerCom & "Constante where nCONSCOD ='" & gCGTipoCtaIF & "'  and nConsValor ='" & Mid(psCodCtaIf, 2, 1) & "'"
    Set rs = oconect.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        EmiteTipoCuentaIF = Trim(rs!cConsDescripcion)
    End If
    rs.Close: Set rs = Nothing
    oconect.CierraConexion
    Set oconect = Nothing

    Exit Function
EmiteTipoCuentaIFErr:
    Call RaiseError(MyUnhandledError, "NCajaCtaIF:EmiteTipoCuentaIF Method")
End Function
Private Function DescTipoCtaIF(psTpoCtaIF As String) As String

End Function
Public Function ValidaBanco(psCtaIf As String) As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Set oConec = New DConecta
    ValidaBanco = ""
    If oConec.AbreConexion = False Then Exit Function
    sSql = "select cPersCod+'.'+cCtaIfCod as CtaIfCod, cCtaIFDesc from CTAIF where cPersCod='" & psCtaIf & "' "
    Set rs = oConec.CargaRecordSet(sSql)
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
    Dim sSql As String
    Set oConec = New DConecta
    GetCtaIFDesc = ""
    If oConec.AbreConexion = False Then Exit Function
    sSql = "SELECT cPersCod+'.'+cCtaIfCod as CtaIfCod, cCtaIFDesc " _
         & "FROM CtaIf WHERE cPersCod = '" & psPersCod & "' and cCtaIfCod LIKE '__" & psMoneda & "____' "
    Set rs = oConec.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        GetCtaIFDesc = Trim(rs!cCtaIFDesc)
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

vsConexion = oIni.CadenaConexion
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
End Sub
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
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New DConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSql = "Select  REPLACE(STR(MAX(Right(cCtaifCod,4)) +1,4,0),' ','0' ) as cCtaCodNew " _
        & " From    CtaIF CIF " _
        & " WHERE   CIF.cPersCod ='" & psPersCod & "' and CIF.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         and Substring(cCtaIfCod,2,1)='" & pnTipoCtaIf & "' and Substring(cCtaIfCod,3,1)='" & pnMoneda & "' " _
        & "         AND LEN(cCtaIFCod)>5 " _
        & " "

    GetNewCuentaIF = ""
    Set rs = oConec.CargaRecordSet(sSql)
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
    Dim sSql As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New DConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSql = "Select  cCtaIfCod " _
        & " From    CtaIF CIF " _
        & " WHERE   CIF.cPersCod ='" & psPersCod & "' and CIF.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         and Substring(cCtaIfCod,2,1)='" & pnTipoCtaIf & "' and Substring(cCtaIfCod,3,1)='" & pnMoneda & "' " _
        & "         AND LEN(cCtaIFCod)<=5 " _
    
    GetVerificaSubCuentaIF = False
    Set rs = oConec.CargaRecordSet(sSql)
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
    Dim sSql As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New DConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSql = "Select  cCtaIfCod " _
        & " From    CtaIF CIF " _
        & " WHERE   CIF.cPersCod ='" & psPersCod & "' and CIF.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         and Substring(cCtaIfCod,2,1)='" & pnTipoCtaIf & "' and Substring(cCtaIfCod,3,1)='" & pnMoneda & "' " _
        & "         AND LEN(cCtaIFCod)<5 " _
    
    GetSubCuentaIF = ""
    Set rs = oConec.CargaRecordSet(sSql)
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
Public Function GeneraSubCuentaIF(ByVal psPersCod As String, ByVal pnTipoIF As CGTipoIF, _
                                  ByVal psCtaCont As String, ByVal psCtaIfCod As String) As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New DConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSql = "Select  cCtaIfCod, C.cCtaContCod, cCtaIFSubCta,C.cCtaContCod + C.cCtaIFSubCta" _
        & " From    CtaIFFiltro  C   " _
        & "         JOIN CTACONT CC ON CC.CCTACONTCOD  = C.cCtaContCod+ C.cCtaIFSubCta " _
        & " Where   C.cCtaContCod Like '" & psCtaCont & "' " _
        & "         And cCtaIfCod Like '" & Mid(psCtaIfCod, 1, 5) & "%' And cPersCod ='" & psPersCod & "' AND cIFTpo ='" & Format(pnTipoIF, "00") & "' "
    
    GeneraSubCuentaIF = ""
    Set rs = oConec.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        GeneraSubCuentaIF = Trim(rs!cCtaIFSubCta)
    End If
    rs.Close
    Set rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function GetSaldoCtaIf(ByVal psPersCod As String, ByVal pnTipoIF As CGTipoIF, _
                             ByVal psCtaIfCod As String, ByVal pdFecha As Date) As Currency
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New DConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSql = "Select  MS.cPersCod, MS.cIFTpo, MS.cCtaIFCod, nSaldo" _
        & " From    CtaIFSaldo MS " _
        & " Where   MS.dCtaIFSaldo = (  Select  Max(MS1.dCtaIFSaldo) " _
        & "                             From    CtaIFSaldo MS1 " _
        & "                             Where   MS.cPersCod = MS1.cPersCod AND " _
        & "                                     MS.cIFTpo = MS1.cIFTpo And MS.cCtaIFCod = MS1.cCtaIFCod " _
        & "                                     AND MS1.dCtaIFSaldo <='" & Format(pdFecha, gsFormatoFecha) & "') " _
        & "         AND MS.cPersCod='" & psPersCod & "' AND MS.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         AND MS.cCtaIFCod='" & psCtaIfCod & "'"
        
    GetSaldoCtaIf = 0
    Set rs = oConec.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        GetSaldoCtaIf = Trim(rs!nSaldo)
    End If
    rs.Close
    Set rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GetInteresDevengados(ByVal psOpeCod As String, ByVal pdFecha As Date, Optional ByVal pnOpeObjOrden As Integer = 1) As ADODB.Recordset
Dim oConec As DConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New DConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

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
    & "                                 ELSE  CI.dCtaIfCap END, '" & Format(pdFecha, gsFormatoFecha) & "') AS Dias " _
    & "          FROM   CTAIF CI " _
    & "                 JOIN INSTITUCIONFINANC I ON I.cPersCod = CI.cPersCod AND I.cIFTpo = CI.cIFTpo " _
    & "                 JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD " _
    & "                 LEFT JOIN ( SELECT  CS.dCtaIFSaldo, CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod, CS.nSaldo " _
    & "                             FROM    CTAIFSALDO CS " _
    & "                             WHERE   CS.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
    & "                                                         FROM    CTAIFSALDO CS1 " _
    & "                                                         Where   CS1.cPersCod = CS.cPersCod And CS1.cIFTpo = CS.cIFTpo And CS1.cCtaIFCod = CS.cCtaIFCod " _
    & "                                                                 AND CS1.dCtaIFSaldo<='" & Format(pdFecha, gsFormatoFecha) & "')) AS SALDO " _
    & "                 ON SALDO.cPersCod=CI.cPersCod and SALDO.cIFTpo = CI.cIFTpo and SALDO.cCtaIFCod= CI.cCtaIFCod " _
    & "                 LEFT JOIN ( SELECT  I.cPersCod, I.cIFTpo, I.cCtaIFCod, I.dCtaIFIntRegistro, I.nCtaIFIntPeriodo, I.nCtaIFIntValor "
sql = sql + "                       FROM    CtaIFInteres I " _
    & "                             WHERE   I.dCtaIFIntRegistro = ( SELECT  MAX(dCtaIFIntRegistro) " _
    & "                                                             FROM    CtaIFInteres I1 " _
    & "                                                             Where   I1.cPersCod = I.cPersCod And I1.cIFTpo = I.cIFTpo And I1.cCtaIFCod = I.cCtaIFCod " _
    & "                                                                     AND I1.dCtaIFIntRegistro<='" & Format(pdFecha, gsFormatoFecha) & "') ) AS INTERES " _
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
Dim oConec As DConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New DConecta
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
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    Set rs = New ADODB.Recordset
    
    Set oConec = New DConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSql = "Select  MS.cPersCod, MS.cIFTpo, MS.cCtaIFCod, nInteres " _
        & " From    CtaIF MS " _
        & " WHERE   MS.cPersCod='" & psPersCod & "' AND MS.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         AND MS.cCtaIFCod='" & psCtaIfCod & "'"
        
    GetInteresCtaIf = 0
    Set rs = oConec.CargaRecordSet(sSql)
    If Not (rs.EOF And rs.BOF) Then
        GetInteresCtaIf = Trim(rs!nInteres)
    End If
    rs.Close
    Set rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function GetRsInteresCtasIF(ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, ByVal psCtaIfCod As String) As ADODB.Recordset
Dim oConec As DConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New DConecta
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
Dim oConec As DConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New DConecta
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
Dim oMov As DMov
Dim lbTrans As Boolean
Dim lnMovNro As String
Set oMov = New DMov
On Error GoTo ErrorActualizaCtas
oMov.BeginTrans
lbTrans = True
oMov.ActualizaCtaIF psPersCod, pnIFTpo, psCtaIfCod, Trim(psCtaIFDesc), _
                    pdCtaIFAper, pdCtaIFCap, pdCtaIFInt, pnCtaIFPlazo, pnCtaIFEstado, pnInteres, psMovUltAct
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
        oMov.RollBackTrans
        lbTrans = False
    End If
    Err.Raise Err.Number, "Actualiza Ctas", Err.Description
End Function

Public Function GetFlujosCtaIF(ByVal psOpeCod As String, _
                            ByVal psCtaIfCodDesde As String, ByVal psCtaIfCodHasta As String, _
                            ByVal pdDesde As Date, ByVal pdHasta As Date) As ADODB.Recordset
Dim oConec As DConecta
Dim sql As String
Dim rs As ADODB.Recordset

Set oConec = New DConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = " SELECT  CI.cPersCod, CI.cIFTpo, CI.cCtaIFCod, P.CPERSNOMBRE, CI1.cCtaIFDesc, CI.cCtaIFDesc," _
    & "         RTRIM(P.CPERSNOMBRE) + ' ' + CI1.cCtaIFDesc + ' ' + CI.cCtaIFDesc AS cNomCtaIF, " _
    & "         ISNULL(SALDO.nSaldo,0) AS nSaldoIni, " _
    & "         CI.nCtaIFPlazo, CI.nInteres , " _
    & "         M.CMOVNRO, M.CMOVDESC, IsNull(D.cDocAbrev,'')  + ' ' + IsNull(MD.cDocNro,'') as Documento , " _
    & "         MC.nConceptoCod, MC.nDiasTrans,  MC.nMovImporte ,  " _
    & "         Case WHEN MC.nMovImporte > 0 THEN ABS(MC.nMovImporte) ELSE 0 END AS DEBE , " _
    & "         Case WHEN MC.nMovImporte < 0 THEN ABS(MC.nMovImporte) ELSE 0 END AS HABER " _
    & " FROM    CTAIF CI " _
    & "         JOIN CTAIF CI1 ON CI1.cPersCod = CI.cPersCod AND CI1.cIFTpo = CI.cIFTpo  AND SUBSTRING(CI.cCtaIFCod,1,LEN(CI1.cCtaIFCod)) = CI1.cCtaIFCod and Len(CI1.cCtaIFCod)=5 " _
    & "         JOIN INSTITUCIONFINANC I ON I.cPersCod = CI.cPersCod AND I.cIFTpo = CI.cIFTpo " _
    & "         JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD " _
    & "         LEFT JOIN ( SELECT  CS.dCtaIFSaldo, CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod, CS.nSaldo " _
    & "                     FROM    CTAIFSALDO CS " _
    & "                     WHERE   CS.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
    & "                                                 FROM    CTAIFSALDO CS1 " _
    & "                                                 Where   CS1.cPersCod = CS.cPersCod And CS1.cIFTpo = CS.cIFTpo " _
    & "                                                         And CS1.cCtaIFCod = CS.cCtaIFCod " _
    & "                                                         AND CS1.dCtaIFSaldo<='" & Format(DateAdd("d", -1, pdDesde), "mm/dd/yyyy") & "' )) AS SALDO" _
    & "         ON SALDO.cPersCod=CI.cPersCod and SALDO.cIFTpo = CI.cIFTpo and SALDO.cCtaIFCod= CI.cCtaIFCod " _
    & "         JOIN MOVCTAIF MC ON MC.cPersCod=CI.cPersCod and MC.cIFTpo = MC.cIFTpo and MC.cCtaIFCod= CI.cCtaIFCod " _
    & "         JOIN MOV M ON M.NMOVNRO = MC.NMOVNRO " _
    & "         LEFT JOIN MOVDOC MD ON MD.NMOVNRO = M.NMOVNRO AND MD.nDocTpo <>" & TpoDocVoucherEgreso & "  " _
    & "         LEFT JOIN DOCUMENTO D ON D.nDocTpo = MD.nDocTpo, OPEOBJ O "
sql = sql + "WHERE   LEN(CI.cCtaIFCod)= (Select max(len(cCtaIFCod)) From CtaIf) AND CI.cCtaIFEstado ='1' " _
    & "         AND o.cOpeCod = '" & psOpeCod & "' and cOpeObjOrden='0' " _
    & "         AND CI.cIFTpo + CI.cCtaIFCod like cOpeObjFiltro  AND MC.nConceptoCod = '1' " _
    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "         AND CI.cPersCod + '.' + CI.cCtaIFCod BETWEEN '" & psCtaIfCodDesde & "' AND '" & psCtaIfCodHasta & "' " _
    & "         and M.NMOVFLAG IN (" & gMovFlagVigente & "," & gMovFlagDeExtorno & ") " _
    & " Order by CI.cIFTpo+ CI.cPersCod + CI.cCtaIFCod , M.CMOVNRO"

Set rs = oConec.CargaRecordSet(sql)
Set GetFlujosCtaIF = rs
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetRepSaldosCtaIf(ByVal psOpeCod As String, ByVal pdFecha As Date, Optional ByVal psOpeObjOrden As String = "'0'", Optional psCtaIFEstado As String = "") As ADODB.Recordset
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
    & "         SELECT  CtaIf.cCtaIfCod, CtaIF.cCtaIFDesc, CtaIF.cBancoDesc, CtaIf.cCtaIFEstado, SUM(CI.NSALDO) AS nImporte " _
    & "         FROM    CTAIFSALDO CI " _
    & "                 JOIN " _
    & "                 (   SELECT  CI.cIFTpo + CI.cPersCod + CI.cCtaIFCod AS cCtaIfCod, CI.cCtaIFDesc, P.cPersNombre cBancoDesc, CI.cCtaIFEstado " _
    & "                     FROM    CTAIF CI JOIN INSTITUCIONFINANC I ON  I.CPERSCOD = CI.CPERSCOD  AND CI.CIFTPO = I.CIFTPO " _
    & "                             JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD, OPEOBJ O " _
    & "                     WHERE   O.cOpeCod = '" & psOpeCod & "' and O.cOpeObjOrden IN (" & psOpeObjOrden & ") and CI.cCtaIFEstado IN ('" & psCtaIFEstado & "') " _
    & "                             AND CI.cIFTpo + CI.cCtaIFCod like O.cOpeObjFiltro " _
    & "                 ) as CtaIf " _
    & "         ON  SUBSTRING(CI.cIFTpo + CI.cPersCod + CI.cCtaIFCod , 1, LEN(CtaIf.cCtaIfCod)) = CtaIf.cCtaIfCod , OPEOBJ O " _
    & "         WHERE   O.cOpeCod = '" & psOpeCod & "' and O.cOpeObjOrden IN (" & psOpeObjOrden & ") " _
    & "                 AND CI.cIFTpo + CI.cCtaIFCod Like O.cOpeObjFiltro "
sql = sql + "           AND CI.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
    & "                                         FROM    CTAIFSALDO CS1 " _
    & "                                         Where   CS1.cPersCod = CI.cPersCod And CS1.cIFTpo = CI.cIFTpo And CS1.cCtaIFCod = CI.cCtaIFCod " _
    & "                                                 AND CS1.dCtaIFSaldo<='" & Format(pdFecha, gsFormatoFecha) & "') " _
    & "         GROUP   BY CtaIf.cCtaIfCod , CtaIF.cCtaIFDesc, CtaIF.cBancoDesc, CtaIF.cCtaIFEstado ) AS Z " _
    & "                 WHERE nImporte <> 0 " _
    & " GROUP  BY Z.cCtaIfCod , Z.cCtaIFDesc, Z.cBancoDesc, Z.cCtaIFEstado " _
    & " ORDER BY  Z.cCtaIFEstado, Z.cCtaIfCod"

Set rs = oConec.CargaRecordSet(sql)
Set GetRepSaldosCtaIf = rs
oConec.CierraConexion
Set oConec = Nothing

End Function

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

