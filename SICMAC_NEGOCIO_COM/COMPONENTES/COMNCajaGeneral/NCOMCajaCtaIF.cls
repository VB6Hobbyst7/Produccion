VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCajaCtaIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0

Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String

Dim vsFormatoFecha As String
Dim oerror As COMConecta.COMErrorHandling

Public Function GetCtaIfFiltro(ByVal psCtaContCod As String, psCtaIfCod As String, psPersCod As String, Optional lbMuestraCta As Boolean = True) As String
Dim oCtaIf As COMDCajaGeneral.DCOMCajaCtasIF
Set oCtaIf = New COMDCajaGeneral.DCOMCajaCtasIF
GetCtaIfFiltro = oCtaIf.GetCtaIfFiltro(psCtaContCod, psCtaIfCod, psPersCod, lbMuestraCta)
Set oCtaIf = Nothing
End Function

Public Function GeneraCodCtaIF(psCtaIfCod As String, psMoneda As String, psTpoCta As String) As String
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim oConect As COMConecta.DCOMConecta
    
    Set Rs = New ADODB.Recordset
    Set oConect = New COMConecta.DCOMConecta
    GeneraCodCtaIF = ""
    If oConect.AbreConexion = False Then Exit Function
    
    Sql = "Select   Max(cCtaIfCod) as MaxCtaIF " _
        & " From    CTAIF " _
        & " Where   cPersCod='" & psCtaIfCod & "' " _
        & "         and Substring(cCtaIfCod,1,2)='" & psTpoCta & "' and Substring(cCtaIfCod,3,1)='" & psMoneda & "' "

    Set Rs = oConect.CargaRecordSet(Sql)
    If Not Rs.EOF And Not Rs.BOF Then
        If IsNull(Rs!MaxCtaIF) Then
            GeneraCodCtaIF = psTpoCta + psMoneda + Format(1, String(7 - Len(psTpoCta + psMoneda), "0"))
        Else
            GeneraCodCtaIF = Format(IIf(IsNull(Rs!MaxCtaIF), 0, Rs!MaxCtaIF) + 1, String(Len(Rs!MaxCtaIF), "0"))
        End If
    End If
    Rs.Close
    Set Rs = Nothing
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

Public Function CargaCtasIF(psMoneda As Moneda, Optional psFiltroCtaIF As String, Optional pnMuestraIF As MuestraIF = MuestraCuentas, Optional psCanje As String = "") As ADODB.Recordset
Dim oCtasIF As COMDCajaGeneral.DCOMCajaCtasIF
Set oCtasIF = New COMDCajaGeneral.DCOMCajaCtasIF
Set CargaCtasIF = oCtasIF.CargaCtasIF(psMoneda, psFiltroCtaIF, pnMuestraIF, psCanje)
Set oCtasIF = Nothing
End Function

Public Function EmiteTipoCuentaIF(psCodCtaIf As String) As String
    On Error GoTo EmiteTipoCuentaIFErr
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim oConect As COMConecta.DCOMConecta
    
    EmiteTipoCuentaIF = ""
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion = False Then Exit Function
    
    Set Rs = New ADODB.Recordset
    Sql = "Select * from " & vsServerCom & "Constante where nCONSCOD ='" & gCGTipoCtaIF & "'  and nConsValor ='" & Mid(psCodCtaIf, 2, 1) & "'"
    Set Rs = oConect.CargaRecordSet(Sql)
    If Not Rs.EOF And Not Rs.BOF Then
        EmiteTipoCuentaIF = Trim(Rs!cConsDescripcion)
    End If
    Rs.Close: Set Rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing

    Exit Function
EmiteTipoCuentaIFErr:
    Call oerror.RaiseError(oerror.MyUnhandledError, "NCajaCtaIF:EmiteTipoCuentaIF Method")
End Function
Private Function DescTipoCtaIF(psTpoCtaIF As String) As String

End Function
Public Function ValidaBanco(psCtaIf As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim Rs As New ADODB.Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    ValidaBanco = ""
    If oConec.AbreConexion = False Then Exit Function
    sSql = "select cPersCod+'.'+cCtaIfCod as CtaIfCod, cCtaIFDesc from CTAIF where cPersCod='" & psCtaIf & "' "
    Set Rs = oConec.CargaRecordSet(sSql)
    If Not (Rs.EOF And Rs.BOF) Then
        ValidaBanco = Trim(Rs!cCtaIFDesc)
    End If
      Rs.Close
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GetCtaIFDesc(psPersCod As String, psMoneda As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim Rs As ADODB.Recordset
    Dim sSql As String
    Set oConec = New COMConecta.DCOMConecta
    GetCtaIFDesc = ""
    If oConec.AbreConexion = False Then Exit Function
    sSql = "SELECT cPersCod+'.'+cCtaIfCod as CtaIfCod, cCtaIFDesc " _
         & "FROM CtaIf WHERE cPersCod = '" & psPersCod & "' and cCtaIfCod LIKE '__" & psMoneda & "____' "
    Set Rs = oConec.CargaRecordSet(sSql)
    If Not (Rs.EOF And Rs.BOF) Then
        GetCtaIFDesc = Trim(Rs!cCtaIFDesc)
    End If
    Rs.Close
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Private Sub Class_Initialize()
    Dim oImp As COMDConstSistema.DCOMImpresoras
    Set oImp = New COMDConstSistema.DCOMImpresoras
    Dim oImpresora As New COMFunciones.FCOMVarImpresion
    
    Set oerror = New COMConecta.COMErrorHandling
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni  As COMConecta.DCOMClasIni
Set oIni = New COMConecta.DCOMClasIni

vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
End Sub

Public Function GetInstFinancieras(Optional ByVal psFiltroTipoCtaIF As String = "") As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim Sql As String
Dim Rs As ADODB.Recordset
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
Set Rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

Sql = "SELECT  cPersCod  AS cCodigo, cPersNombre as Descripcion, Nivel " _
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
Set Rs = oConec.CargaRecordSet(Sql)
Set GetInstFinancieras = Rs
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetCtasInstFinancieras(Optional ByVal psFiltroTipoCtaIF As String = "", Optional ByVal psCtaCont As String = "", Optional pnNivel As MuestraIF = MuestraCuentas, Optional ByVal psFiltroTipoIF As String) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim Sql As String
Dim Rs As ADODB.Recordset
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
Set Rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

Sql = "SELECT  cPersCod  AS cCodigo, cPersNombre as Descripcion, Nivel " _
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

Set Rs = oConec.CargaRecordSet(Sql)
Set GetCtasInstFinancieras = Rs
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetNewCuentaIF(ByVal psPersCod As String, ByVal pnTipoIF As COMDConstantes.CGTipoIF, _
                              ByVal pnTipoCtaIf As COMDConstantes.CGTipoCtaIF, ByVal pnMoneda As COMDConstantes.Moneda, ByVal psSubCtaIF As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim Rs As ADODB.Recordset
    Dim sSql As String
    Set Rs = New ADODB.Recordset
    Set oConec = New COMConecta.DCOMConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSql = "Select  REPLACE(STR(MAX(Right(cCtaifCod,4)) +1,4,0),' ','0' ) as cCtaCodNew " _
        & " From    CtaIF CIF " _
        & " WHERE   CIF.cPersCod ='" & psPersCod & "' and CIF.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         and Substring(cCtaIfCod,2,1)='" & pnTipoCtaIf & "' and Substring(cCtaIfCod,3,1)='" & pnMoneda & "' " _
        & "         AND LEN(cCtaIFCod)>5 " _
        & " "

    GetNewCuentaIF = ""
    Set Rs = oConec.CargaRecordSet(sSql)
    If Not (Rs.EOF And Rs.BOF) Then
        GetNewCuentaIF = Mid(psSubCtaIF, 1, 3) & Trim(IIf(IsNull(Rs!cCtaCodNew), Format(1, "0000"), Rs!cCtaCodNew))
    Else
        GetNewCuentaIF = Mid(psSubCtaIF, 1, 3) & Format(1, "0000")
    End If
    Rs.Close
    Set Rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GetVerificaSubCuentaIF(ByVal psPersCod As String, ByVal pnTipoIF As COMDConstantes.CGTipoIF, _
                                        ByVal pnTipoCtaIf As COMDConstantes.CGTipoCtaIF, ByVal pnMoneda As COMDConstantes.Moneda) As Boolean
    Dim oConec As COMConecta.DCOMConecta
    Dim Rs As ADODB.Recordset
    Dim sSql As String
    Set Rs = New ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSql = "Select  cCtaIfCod " _
        & " From    CtaIF CIF " _
        & " WHERE   CIF.cPersCod ='" & psPersCod & "' and CIF.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         and Substring(cCtaIfCod,2,1)='" & pnTipoCtaIf & "' and Substring(cCtaIfCod,3,1)='" & pnMoneda & "' " _
        & "         AND LEN(cCtaIFCod)<=5 " _
    
    GetVerificaSubCuentaIF = False
    Set Rs = oConec.CargaRecordSet(sSql)
    If Not (Rs.EOF And Rs.BOF) Then
        GetVerificaSubCuentaIF = True
    End If
    Rs.Close
    Set Rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function GetSubCuentaIF(ByVal psPersCod As String, ByVal pnTipoIF As COMDConstantes.CGTipoIF, _
                                        ByVal pnTipoCtaIf As COMDConstantes.CGTipoCtaIF, ByVal pnMoneda As COMDConstantes.Moneda) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim Rs As ADODB.Recordset
    Dim sSql As String
    Set Rs = New ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSql = "Select  cCtaIfCod " _
        & " From    CtaIF CIF " _
        & " WHERE   CIF.cPersCod ='" & psPersCod & "' and CIF.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         and Substring(cCtaIfCod,2,1)='" & pnTipoCtaIf & "' and Substring(cCtaIfCod,3,1)='" & pnMoneda & "' " _
        & "         AND LEN(cCtaIFCod)<5 " _
    
    GetSubCuentaIF = ""
    Set Rs = oConec.CargaRecordSet(sSql)
    If Not (Rs.EOF And Rs.BOF) Then
        GetSubCuentaIF = Trim(Rs!cCtaIfCod)
    Else
        GetSubCuentaIF = Format(pnTipoCtaIf, "00") & pnMoneda & "00"
    End If
    Rs.Close
    Set Rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GeneraSubCuentaIF(ByVal psPersCod As String, ByVal pnTipoIF As COMDConstantes.CGTipoIF, _
                                  ByVal psCtaCont As String, ByVal psCtaIfCod As String) As String
    Dim oConec As COMConecta.DCOMConecta
    Dim Rs As ADODB.Recordset
    Dim sSql As String
    Set Rs = New ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSql = "Select  cCtaIfCod, C.cCtaContCod, cCtaIFSubCta,C.cCtaContCod + C.cCtaIFSubCta" _
        & " From    CtaIFFiltro  C   " _
        & "         JOIN CTACONT CC ON CC.CCTACONTCOD  = C.cCtaContCod+ C.cCtaIFSubCta " _
        & " Where   C.cCtaContCod Like '" & psCtaCont & "' " _
        & "         And cCtaIfCod Like '" & Mid(psCtaIfCod, 1, 5) & "%' And cPersCod ='" & psPersCod & "' AND cIFTpo ='" & Format(pnTipoIF, "00") & "' "
    
    GeneraSubCuentaIF = ""
    Set Rs = oConec.CargaRecordSet(sSql)
    If Not (Rs.EOF And Rs.BOF) Then
        GeneraSubCuentaIF = Trim(Rs!cCtaIFSubCta)
    End If
    Rs.Close
    Set Rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GetSaldoCtaIf(ByVal psPersCod As String, ByVal pnTipoIF As COMDConstantes.CGTipoIF, _
                             ByVal psCtaIfCod As String, ByVal pdFecha As Date) As Currency
    
    Dim oConec As COMConecta.DCOMConecta
    Dim Rs As ADODB.Recordset
    Dim sSql As String
    Set Rs = New ADODB.Recordset
    Dim oFun As COMFunciones.FCOMVarPublicas
    
    Set oConec = New COMConecta.DCOMConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSql = "Select  MS.cPersCod, MS.cIFTpo, MS.cCtaIFCod, nSaldo" _
        & " From    CtaIFSaldo MS " _
        & " Where   MS.dCtaIFSaldo = (  Select  Max(MS1.dCtaIFSaldo) " _
        & "                             From    CtaIFSaldo MS1 " _
        & "                             Where   MS.cPersCod = MS1.cPersCod AND " _
        & "                                     MS.cIFTpo = MS1.cIFTpo And MS.cCtaIFCod = MS1.cCtaIFCod " _
        & "                                     AND MS1.dCtaIFSaldo <='" & Format(pdFecha, oFun.gsFormatoFecha) & "') " _
        & "         AND MS.cPersCod='" & psPersCod & "' AND MS.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         AND MS.cCtaIFCod='" & psCtaIfCod & "'"
        
    GetSaldoCtaIf = 0
    Set Rs = oConec.CargaRecordSet(sSql)
    If Not (Rs.EOF And Rs.BOF) Then
        GetSaldoCtaIf = Trim(Rs!nSaldo)
    End If
    Rs.Close
    Set Rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function EmiteOpeCta(ByVal psOpeCod As String, ByVal psDH As String, Optional psOpeCtaOrden As String = "0", Optional ByVal psObjetoCod As String = "", Optional pnFiltro As COMDConstantes.TpoObjetos, Optional lbMuestraSubCta As Boolean = True) As String
    On Error GoTo EmiteOpeCodErr
    Dim sSql As String
    Dim Rs As ADODB.Recordset
    Dim oConect As COMConecta.DCOMConecta
    Dim lsListaCtasCont As String
    Dim lsSubCta As String
    Dim oCtaIf As NCOMCajaGeneral
    Dim oCtaCont As COMDContabilidad.DCOMCtaCont
    Dim oContFunc As COMNContabilidad.NCOMContFunciones
    Dim oOpe As New COMDConstSistema.DCOMOperacion
    
    Set oCtaCont = New COMDContabilidad.DCOMCtaCont
    Set oCtaIf = New NCOMCajaGeneral
    Set oContFunc = New COMNContabilidad.NCOMContFunciones
    
    Set oConect = New COMConecta.DCOMConecta
    Set Rs = New ADODB.Recordset
    Dim lnNumCtas As Integer
    Dim lsNuevaCta As String
    Dim lsCuenta As String
    Dim lsAux As String
    Dim lbExisteFiltro As Boolean
    Dim lbExisteCuenta As Boolean
    If oConect.AbreConexion = False Then Exit Function
    
    sSql = " SELECT  OC.cCtaContCod , OC.cOpeCtaOrden, OC.cOpeCtaDH, CC.cCtaContDesc " _
        & " FROM     OPECTA OC  JOIN CTACONT CC ON CC.CCTACONTCOD = OC.CCTACONTCOD   " _
        & " WHERE    OC.cOpeCod ='" & psOpeCod & "'  AND OC.cOpeCtaDH = '" & psDH & "' AND OC.cOpeCtaOrden = '" & psOpeCtaOrden & "'"
    lnNumCtas = 0
    Set Rs = oConect.CargaRecordSet(sSql)
    lnNumCtas = Rs.RecordCount
    lbExisteCuenta = False
    Do While Not Rs.EOF
        lsListaCtasCont = Rs!cCtaContCod
        If lsListaCtasCont <> "" Then
            'lsListaCtasCont = Mid(lsListaCtasCont, 1, Len(lsListaCtasCont) - 3)
            lsNuevaCta = lsListaCtasCont
            lsSubCta = oOpe.EmiteSubCtaOpeObj(psOpeCod, lsListaCtasCont)
            lsAux = ""
            If lsSubCta = "" Then
                If psObjetoCod <> "" Or lsSubCta <> "" Then
                    lsSubCta = ""
                    lsAux = oContFunc.GetFiltroObjetos(pnFiltro, lsListaCtasCont, psObjetoCod, lbMuestraSubCta, lbExisteFiltro)
                    If lbExisteFiltro Then
                        lbExisteCuenta = True
                    End If
                Else
                    lsAux = IIf(lbMuestraSubCta, lsListaCtasCont, "")
                End If
            Else
                lsAux = IIf(lbMuestraSubCta, lsListaCtasCont, "") & lsSubCta
            End If
            If lsNuevaCta = lsAux And Not lbExisteFiltro And lnNumCtas > 1 And lsSubCta = "" Then
                lsAux = ""
            Else
                lsCuenta = lsAux
            End If
        End If
        lbExisteFiltro = False
        Rs.MoveNext
    Loop
    Rs.Close: Set Rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    EmiteOpeCta = lsCuenta
    Set oCtaIf = Nothing
    Set oCtaCont = Nothing
    Exit Function
EmiteOpeCodErr:
    Call oerror.RaiseError(oerror.MyUnhandledError, "COMDMov.DCOMMov:EmiteOpeCod Method")
End Function

Public Function GetInteresDevengados(ByVal psOpeCod As String, ByVal pdFecha As Date, Optional ByVal pnOpeObjOrden As Integer = 1) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
'Dim oOpe   As DOperacion
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim lsCtaCapital As String

Dim oFun As COMFunciones.FCOMVarPublicas
'Set oOpe = New DOperacion
Set oConec = New COMConecta.DCOMConecta
Set Rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

lsCtaCapital = EmiteOpeCta(psOpeCod, "H", 2)
'Set oOpe = Nothing

Sql = " Select   cCtaIFDesc, cInstFinanc,nCapital ," _
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
    & "                                 ELSE  CI.dCtaIfCap END, '" & Format(pdFecha, oFun.gsFormatoFecha) & "') AS Dias " _
    & "          FROM   CTAIF CI " _
    & "                 JOIN INSTITUCIONFINANC I ON I.cPersCod = CI.cPersCod AND I.cIFTpo = CI.cIFTpo " _
    & "                 JOIN PERSONA P ON P.CPERSCOD = I.CPERSCOD " _
    & "                 LEFT JOIN ( SELECT  CS.dCtaIFSaldo, CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod, CS.nSaldo " _
    & "                             FROM    CTAIFSALDO CS WHERE cs.cCtaContCod LIKE '" & lsCtaCapital & "%'" _
    & "                                 and CS.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
    & "                                                         FROM    CTAIFSALDO CS1 " _
    & "                                                         Where   cs1.cCtaContCod = cs.cCtaContCod and CS1.cPersCod = CS.cPersCod And CS1.cIFTpo = CS.cIFTpo And CS1.cCtaIFCod = CS.cCtaIFCod " _
    & "                                                                 AND CS1.dCtaIFSaldo<='" & Format(pdFecha, oFun.gsFormatoFecha) & "')) AS SALDO " _
    & "                 ON SALDO.cPersCod=CI.cPersCod and SALDO.cIFTpo = CI.cIFTpo and SALDO.cCtaIFCod= CI.cCtaIFCod " _
    & "                 LEFT JOIN ( SELECT  I.cPersCod, I.cIFTpo, I.cCtaIFCod, I.dCtaIFIntRegistro, I.nCtaIFIntPeriodo, I.nCtaIFIntValor "
Sql = Sql + "                       FROM    CtaIFInteres I " _
    & "                             WHERE   I.dCtaIFIntRegistro = ( SELECT  MAX(dCtaIFIntRegistro) " _
    & "                                                             FROM    CtaIFInteres I1 " _
    & "                                                             Where   I1.cPersCod = I.cPersCod And I1.cIFTpo = I.cIFTpo And I1.cCtaIFCod = I.cCtaIFCod " _
    & "                                                                     AND I1.dCtaIFIntRegistro<='" & Format(pdFecha, oFun.gsFormatoFecha) & "') ) AS INTERES " _
    & "                 ON INTERES.cPersCod=CI.cPersCod and INTERES.cIFTpo = CI.cIFTpo and INTERES.cCtaIFCod= CI.cCtaIFCod, OPEOBJ O " _
    & "         WHERE   LEN(CI.cCtaIFCod)=(Select Max(Len(cCtaIFCod)) From CtaIf) AND cCtaIFEstado ='" & gEstadoCtaIFActiva & "' " _
    & "                 AND o.cOpeCod = '" & psOpeCod & "' and cOpeObjOrden='" & pnOpeObjOrden & "' and " _
    & "                 CI.cIFTpo+CI.cCtaIFCod Like cOpeObjFiltro And Saldo.nSaldo>0  " _
    & "         ) AS DATOS " _
    & " Order By cInstFinanc "

Set Rs = oConec.CargaRecordSet(Sql)
Set GetInteresDevengados = Rs
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetCuentasIFCanc(ByVal psOpeCod As String, ByVal pdFecha As Date, Optional ByVal pnOpeObjOrden As Integer = 1) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim Sql As String
Dim Rs As ADODB.Recordset

Dim oFun As COMFunciones.FCOMVarPublicas
Set oConec = New COMConecta.DCOMConecta
Set Rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

Sql = "SELECT   CI.cCtaIFDesc, cPersNombre AS cInstFinanc , " _
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
    & "                                                         AND CS1.dCtaIFSaldo<='" & Format(pdFecha, oFun.gsFormatoFecha) & "')) AS SALDO " _
    & "         ON SALDO.cPersCod=CI.cPersCod and SALDO.cIFTpo = CI.cIFTpo and SALDO.cCtaIFCod= CI.cCtaIFCod, OPEOBJ O " _
    & " WHERE   LEN(CI.cCtaIFCod)=(Select Max(Len(cCtaIFCod)) From CtaIf) AND cCtaIFEstado ='1' " _
    & "         AND o.cOpeCod = '" & psOpeCod & "' and cOpeObjOrden='" & gEstadoCtaIFActiva & "' " _
    & "         And CI.cIFTpo+CI.cCtaIFCod Like cOpeObjFiltro And Saldo.nSaldo>0 "

Set Rs = oConec.CargaRecordSet(Sql)
Set GetCuentasIFCanc = Rs
oConec.CierraConexion
Set oConec = Nothing
End Function
Public Function GetInteresCtaIf(ByVal psPersCod As String, ByVal pnTipoIF As CGTipoIF, _
                             ByVal psCtaIfCod As String) As Currency
    Dim oConec As COMConecta.DCOMConecta
    Dim Rs As ADODB.Recordset
    Dim sSql As String
    Set Rs = New ADODB.Recordset
    
    Set oConec = New COMConecta.DCOMConecta
    If oConec.AbreConexion = False Then Exit Function
    
    sSql = "Select  MS.cPersCod, MS.cIFTpo, MS.cCtaIFCod, nInteres " _
        & " From    CtaIF MS " _
        & " WHERE   MS.cPersCod='" & psPersCod & "' AND MS.cIFTpo='" & Format(pnTipoIF, "00") & "' " _
        & "         AND MS.cCtaIFCod='" & psCtaIfCod & "'"
        
    GetInteresCtaIf = 0
    Set Rs = oConec.CargaRecordSet(sSql)
    If Not (Rs.EOF And Rs.BOF) Then
        GetInteresCtaIf = Trim(Rs!nInteres)
    End If
    Rs.Close
    Set Rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function GetRsInteresCtasIF(ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, ByVal psCtaIfCod As String) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim Sql As String
Dim Rs As ADODB.Recordset

Set oConec = New COMConecta.DCOMConecta
Set Rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

Sql = "Select convert(char(10),dCtaIFIntRegistro,103) as FechaInt, nCtaIFIntValor , nCtaIFIntPeriodo, cCtaIFCod ,cPersCod, cIFTpo,  1 as  Registrado " _
    & " from CtaifInteres " _
    & " Where cPersCod='" & psPersCod & "' and cIFTpo ='" & Format(pnIFTpo, "00") & "' and cCtaIFCod='" & psCtaIfCod & "'"

Set Rs = oConec.CargaRecordSet(Sql)
Set GetRsInteresCtasIF = Rs
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetDatosCtaIf(ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, ByVal psCtaIfCod As String) As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim Sql As String
Dim Rs As ADODB.Recordset

Set oConec = New COMConecta.DCOMConecta
Set Rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

Sql = " SELECT  C.dCtaIFAper, C.dCtaIfCap, C.dCtaIfInt, C.nCtaIFPlazo, C.nInteres, C.cCtaIFEstado, CC.cConsDescripcion AS cEstadoCons, Isnull(SALDO.nSaldo,0) as Capital" _
    & " FROM    CTAIF C " _
    & "         JOIN CONSTANTE CC ON    CC.NCONSVALOR =  C.cCtaIFEstado  " _
    & "         LEFT JOIN ( SELECT  CS.dCtaIFSaldo, CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod, CS.nSaldo " _
    & "                     FROM    CTAIFSALDO CS " _
    & "                     WHERE   CS.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
    & "                                                 FROM    CTAIFSALDO CS1 " _
    & "                                                 Where   CS1.cPersCod = CS.cPersCod And CS1.cIFTpo = CS.cIFTpo And CS1.cCtaIFCod = CS.cCtaIFCod )) AS SALDO " _
    & "                     ON SALDO.cPersCod=C.cPersCod and SALDO.cIFTpo = C.cIFTpo and SALDO.cCtaIFCod= C.cCtaIFCod " _
    & " Where CC.NCONSCOD = " & gCGEstadoCtaIF & " AND C.cPersCod='" & psPersCod & "' and C.cIFTpo ='" & Format(pnIFTpo, "00") & "' and C.cCtaIFCod='" & psCtaIfCod & "'"

Set Rs = oConec.CargaRecordSet(Sql)
Set GetDatosCtaIf = Rs
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
Dim oFun As COMFunciones.FCOMVarPublicas
On Error GoTo ErrorActualizaCtas
oMov.BeginTrans
lbTrans = True
oMov.ActualizaCtaIF psPersCod, pnIFTpo, psCtaIfCod, Trim(psCtaIFDesc), _
                    Format(pdCtaIFAper, oFun.gsFormatoFecha), Format(pdCtaIFCap, oFun.gsFormatoFecha), Format(pdCtaIFInt, oFun.gsFormatoFecha), pnCtaIFPlazo, pnCtaIFEstado, pnInteres, psMovUltAct
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
Dim Sql As String
Dim Rs As ADODB.Recordset

Set oConec = New COMConecta.DCOMConecta
Set Rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

Sql = " SELECT  CI.cPersCod, CI.cIFTpo, CI.cCtaIFCod, P.CPERSNOMBRE, CI1.cCtaIFDesc, CI.cCtaIFDesc," _
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
Sql = Sql + "WHERE   LEN(CI.cCtaIFCod)= (Select max(len(cCtaIFCod)) From CtaIf) AND CI.cCtaIFEstado ='1' " _
    & "         AND o.cOpeCod = '" & psOpeCod & "' and cOpeObjOrden='0' " _
    & "         AND CI.cIFTpo + CI.cCtaIFCod like cOpeObjFiltro " _
    & "         AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & Format(pdDesde, "yyyymmdd") & "' AND '" & Format(pdHasta, "yyyymmdd") & "' " _
    & "         AND CI.cPersCod + '.' + CI.cCtaIFCod BETWEEN '" & psCtaIfCodDesde & "' AND '" & psCtaIfCodHasta & "' " _
    & "         and M.NMOVFLAG IN (" & gMovFlagVigente & "," & gMovFlagDeExtorno & ") " _
    & "         and mo.cCtaContCod LIKE '11_[23]%' " _
    & " Order by CI.cIFTpo+ CI.cPersCod + CI.cCtaIFCod , M.CMOVNRO"

Set Rs = oConec.CargaRecordSet(Sql)
Set GetFlujosCtaIF = Rs
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function GetRepSaldosCtaIf(ByVal psOpeCod As String, ByVal pdFecha As Date, Optional ByVal psOpeObjOrden As String = "'0'", Optional psCtaIFEstado As String = "", Optional psMoneda As String = "") As ADODB.Recordset
Dim oConec As COMConecta.DCOMConecta
Dim Sql As String
Dim Rs As ADODB.Recordset

Dim oFun As COMFunciones.FCOMVarPublicas

Set oConec = New COMConecta.DCOMConecta
Set Rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

Sql = " SELECT  Z.cCtaIfCod , Z.cCtaIFDesc, Z.cBancoDesc, Z.cCtaIFEstado, " _
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
Sql = Sql + "           AND CI.dCtaIFSaldo = (  SELECT  MAX(dCtaIFSaldo) " _
    & "                                         FROM    CTAIFSALDO CS1 " _
    & "                                         Where   CS1.cPersCod = CI.cPersCod And CS1.cIFTpo = CI.cIFTpo And CS1.cCtaIFCod = CI.cCtaIFCod " _
    & "                                                 AND CS1.dCtaIFSaldo<='" & Format(pdFecha, oFun.gsFormatoFecha) & "') " _
    & "         GROUP   BY CtaIf.cCtaIfCod , CtaIF.cCtaIFDesc, CtaIF.cBancoDesc, CtaIF.cCtaIFEstado ) AS Z " _
    & " WHERE nImporte <> 0 " _
    & " GROUP  BY Z.cCtaIfCod , Z.cCtaIFDesc, Z.cBancoDesc, Z.cCtaIFEstado " _
    & " ORDER BY  Z.cCtaIFEstado, Z.cBancoDesc, Z.cCtaIfCod"

Set Rs = oConec.CargaRecordSet(Sql)
Set GetRepSaldosCtaIf = Rs
oConec.CierraConexion
Set oConec = Nothing

End Function

Public Function GrabaDesembolsoVaucherCH(ByVal psMovNro As String, _
                                        ByVal psPersCod As String, ByVal psFormatoFecha As String, _
                                        ByVal psOpeCod As String, ByVal psMovDesc As String, ByVal pnImporte As Currency, _
                                        ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnNroProc As Integer, _
                                        ByVal psFechaDoc As String, _
                                        ByVal psNroVoucher As String, ByVal pnMovNroRef As Long, _
                                        ByVal psCtaDebe As String, ByVal pnSaldoCh As Currency) As Integer
            
Dim lnMovItem As Integer
Dim lnMovOrden As Integer
Dim oDMov As COMDMov.DCOMMov
Dim oContF As COMNContabilidad.NCOMContFunciones
Dim lsMovNro As String
Dim lsSubCta As String
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New COMDMov.DCOMMov
On Error GoTo GrabaDesembolsoCHErr

Set oContF = New COMNContabilidad.NCOMContFunciones

'GrabaDesembolsoCH = 1
oDMov.BeginTrans
lbTrans = True

oDMov.InsertaMov psMovNro, psOpeCod, psMovDesc, gMovEstContabMovContable, gMovFlagVigente
lnMovNro = oDMov.GetnMovNro(psMovNro)

oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
'guardamos la cuenta en el debe

lnMovItem = 0: lnMovOrden = 0
lnMovItem = lnMovItem + 1:

'lsSubCta = oContF.GetFiltroObjetos(ObjCMACAgenciaArea, psCtaDebe, psAreaCh + psAgeCh, False)
'oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe + lsSubCta, pnImporte
oDMov.InsertaMovCta lnMovNro, lnMovItem, psCtaDebe, pnImporte

lnMovOrden = lnMovOrden + 1
oDMov.InsertaMovObj lnMovNro, lnMovItem, lnMovOrden, ObjCMACAgenciaArea
oDMov.InsertaMovObjAgenciaArea lnMovNro, lnMovItem, lnMovOrden, psAgeCh, psAreaCh


oDMov.InsertaMovCajaChica lnMovNro, psAreaCh, psAgeCh, pnNroProc, gCHTipoProcDesembolso

If psNroVoucher <> "" Then
    oDMov.InsertaMovDoc lnMovNro, TpoDocVoucherEgreso, psNroVoucher, Format(psFechaDoc, psFormatoFecha)
End If

pnSaldoCh = pnSaldoCh - pnImporte

oDMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProc, , , pnImporte, , pnSaldoCh
'GrabaDesembolsoCH = 0

oDMov.InsertaMovGasto lnMovNro, psPersCod, "0"

oDMov.InsertaMovRef lnMovNro, pnMovNroRef

GrabaDesembolsoVaucherCH = 1
oDMov.CommitTrans
lbTrans = False
Set oDMov = Nothing
Set oContF = Nothing
Exit Function
GrabaDesembolsoCHErr:
    If lbTrans Then
        oDMov.RollbackTrans
        lbTrans = False
    End If
    GrabaDesembolsoVaucherCH = 0
    Err.Raise vbObjectError + 100, "GrabaDesembolsoCH", Err.Description & " Error Graba Rendicion a rendir Efectivo "
End Function


Public Function GrabaExtornoDesembolsoVaucher(ByVal pdFecSis As Date, ByVal pdFechaDesemb As Date, ByVal psFormatoFecha As String, _
                                     ByVal pnMovNroDesemb As Long, ByVal psMovNroExt As String, ByVal psOpeCod As String, ByVal psMovDesc As String, _
                                     ByVal psAreaCh As String, ByVal psAgeCh As String, _
                                     ByVal pnNroProc As Integer, ByVal pnImporte As Currency) As Integer

Dim oDMov As COMDMov.DCOMMov
Dim lnSaldo As Currency
Dim lbTrans As Boolean
Dim lnMovNro As Long
Set oDMov = New COMDMov.DCOMMov
Dim oCaja As COMDCajaGeneral.DCOMDocumento

On Error GoTo GrabaExtornoDesembolsoErr

'GrabaExtornoDesembolso = 1

oDMov.BeginTrans
lbTrans = True

    oDMov.InsertaMov psMovNroExt, psOpeCod, psMovDesc, gMovEstContabNoContable, gMovFlagExtornado
    lnMovNro = oDMov.GetnMovNro(psMovNroExt)
    oDMov.InsertaMovCont lnMovNro, pnImporte, "0", "0"
    oDMov.ActualizaMov pnMovNroDesemb, , , gMovFlagEliminado
    oDMov.InsertaMovRef lnMovNro, pnMovNroDesemb

Set oCaja = New COMDCajaGeneral.DCOMDocumento
    lnSaldo = oCaja.GetDatosCajaChica(psAreaCh, psAgeCh, gSaldoActual)
Set oCaja = Nothing

lnSaldo = lnSaldo + pnImporte
oDMov.ActualizaCajaChica psAreaCh, psAgeCh, pnNroProc, , , 0, , lnSaldo

GrabaExtornoDesembolsoVaucher = 1
oDMov.CommitTrans
lbTrans = False

Set oDMov = Nothing
Exit Function
    If lbTrans Then
        lbTrans = False
        oDMov.RollbackTrans
    End If
GrabaExtornoDesembolsoErr:
    GrabaExtornoDesembolsoVaucher = 0
    Err.Raise vbObjectError + 100, "GrabaExtornoDesembolso", "Error Grabar Extorno de Desembolso de Caja chica"
End Function

