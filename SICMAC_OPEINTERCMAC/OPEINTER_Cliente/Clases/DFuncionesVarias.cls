VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dFuncionesNeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum Moneda
    gMonedaNacional = 1
    gMonedaExtranjera = 2
End Enum

Dim dbCmact As DConecta
Public oImp As NImpresion

Dim sSQL As String
Public Enum TpoDoc
    TpoDocCheque = 47
    TpoDocOrdenPago = 48
    TpoDocNotaCargo = 57
    TpoDocNotaAbono = 58
End Enum

Public Function GetCapParametro(ByVal nParametro As CaptacParametro) As Double
    Dim rsVar As ADODB.Recordset
    Set dbCmact = New DConecta
    dbCmact.AbreConexion
    
    sSQL = "SELECT nParValor FROM Parametro WHERE nParCod = " & nParametro & " " _
        & "And nParProd = " & gPrdParamCaptac
    Set rsVar = New ADODB.Recordset
    rsVar.CursorLocation = adUseClient
    
    'rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsVar = dbCmact.CargaRecordSet(sSQL)
    
    Set rsVar.ActiveConnection = Nothing
    If rsVar.EOF And rsVar.BOF Then
        GetCapParametro = 0
    Else
        GetCapParametro = rsVar("nParValor")
    End If
    rsVar.Close
    dbCmact.CierraConexion
    Set rsVar = Nothing
End Function

'ALPA 20081003*******************************************************************************
'Se agrego el parametro nValPondREU para el REU
'********************************************************************************************
Public Function EmiteTipoCambio(ByVal dFecha As Date, ByVal nTpoTipoCambio As TipoCambio) As Double
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Set dbCmact = New DConecta

    EmiteTipoCambio = 0

    If dbCmact.AbreConexion = False Then Exit Function
    rs.CursorLocation = adUseClient
    sql = "Select isnull(nValPondREU,0) nValPondREU,nValFijo, nValFijoDia, nValVent, nValComp, nValVentEsp, nValCompEsp, nValPond,nValPondVenta From TipoCambio " _
        & " WHERE dFecCamb = (   Select Max(dFecCamb)" _
        & "                      From TipoCambio " _
        & "                      Where datediff(day,dFecCamb,'" & Format$(dFecha, "mm/dd/yyyy") & "')=0)"
  
    Set rs = dbCmact.CargaRecordSet(sql)
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF And Not rs.BOF Then
        Select Case nTpoTipoCambio
            Case TCFijoMes
                EmiteTipoCambio = rs("nValFijo")
            Case TCFijoDia
                EmiteTipoCambio = rs("nValFijoDia")
            Case TCVenta
                EmiteTipoCambio = rs("nValVent")
            Case TCCompra
                EmiteTipoCambio = rs("nValComp")
            Case TCVentaEsp
                EmiteTipoCambio = rs("nValVentEsp")
            Case TCCompraEsp
                EmiteTipoCambio = rs("nValCompEsp")
            Case TCPonderado
                EmiteTipoCambio = rs("nValPond")
            Case TCPondVenta
                EmiteTipoCambio = IIf(IsNull(rs("nValPondVenta")), 0, rs("nValPondVenta"))
            Case TCPondREU
                EmiteTipoCambio = rs("nValPondREU")
        End Select
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Function GeneraMovNro(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String
    On Error GoTo GeneraMovNroErr
    Dim rs As ADODB.Recordset
    Dim sql As String
    Set rs = New ADODB.Recordset
    Set dbCmact = New DConecta

    If dbCmact.AbreConexion = False Then Exit Function
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & dbCmact.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & Right(psCodAge, 2) & "','" & psUser & "'"
    Else
       sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    Set rs = dbCmact.Ejecutar(sql)
    If Not rs.EOF Then
        GeneraMovNro = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
    dbCmact.CierraConexion
    'Set dbCmact = Nothing
    Exit Function
GeneraMovNroErr:
    Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function

Public Function GetCuentaAbonoIF(ByVal sPersCod As String, ByVal nMoneda As Moneda) As String
Dim rsCta As ADODB.Recordset
Dim sSQL As String

Set dbCmact = New DConecta

If dbCmact.AbreConexion = False Then Exit Function

sSQL = "Select C.cCtaCod FROM ProductoPersona PP INNER JOIN CuentaIFEspecial C ON " _
    & " PP.cCtaCod = C.cCtaCod Where PP.cPersCod = '" & sPersCod & "' And C.bAbono = 1 " _
    & " And C.cCtaCod LIKE '________" & Trim(nMoneda) & "%' AND SUBSTRING(C.cCtaCod,6,3)='232' "

Set rsCta = New ADODB.Recordset
'rsCta.CursorLocation = adUseClient
'rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsCta = dbCmact.CargaRecordSet(sSQL)

Set rsCta.ActiveConnection = Nothing
If Not (rsCta.EOF And rsCta.BOF) Then
    GetCuentaAbonoIF = rsCta("cCtaCod")
Else
    GetCuentaAbonoIF = ""
End If
dbCmact.CierraConexion
Set rsCta = Nothing
End Function

Public Function GetnMovNro(ByVal sMovNro As String) As Long
Dim rsMov As ADODB.Recordset
Dim sSQL As String

Set dbCmact = New DConecta

sSQL = "Select nMovNro From Mov Where cMovNro = '" & sMovNro & "'"
Set rsMov = New ADODB.Recordset
If dbCmact.AbreConexion = False Then Exit Function
rsMov.CursorLocation = adUseClient
'rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMov = dbCmact.CargaRecordSet(sSQL)

Set rsMov.ActiveConnection = Nothing
If rsMov.EOF And rsMov.BOF Then
    GetnMovNro = 0
Else
    GetnMovNro = rsMov("nMovNro")
End If
dbCmact.CierraConexion
rsMov.Close
Set rsMov = Nothing
End Function

Public Function IncrementaMovNro(ByVal cMovnro As String) As String
Dim cTexto1 As String
Dim cTexto2 As String
Dim cTexto3 As String
Dim cNumero As Integer
Dim nNumero As Integer
Dim sNumero As String
cTexto1 = Mid(cMovnro, 1, 19)
cTexto2 = Mid(cMovnro, 20, 2)
cTexto3 = Mid(cMovnro, 22, 4)
nNumero = Val(cTexto2) + 1
sNumero = Right("00" & Trim(Str(nNumero)), 2)
IncrementaMovNro = cTexto1 + sNumero + cTexto3
End Function

Public Function GetDatosCuentaAho(ByVal sCuenta As String) As ADODB.Recordset
Dim rsCta As ADODB.Recordset
Set dbCmact = New DConecta
If dbCmact.AbreConexion = False Then Exit Function

sSQL = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
    & "(C.nSaldoDisp-Isnull(A.nRetencion,0)) as nSaldoDisp,A.nRetencion, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, AG.cAgeDescripcion, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.bOrdPag, A.dUltContacto, T.cConsDescripcion cEstado, " _
    & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.bInactiva, A.bInmovilizada, A.nSobregiro, " _
    & "C.nFirmasMin, C.cAlias, nTpoPrograma=isnull(c.nTpoPrograma,0), ISNULL(C.nBloqueoParcial,0) nBloqueoParcial,nPlazo,nMontoAbono,dRenoPanderito FROM Producto P " _
    & "INNER JOIN Captaciones C INNER JOIN CaptacAhorros A ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod " _
    & "WHERE P.cCtaCod = '" & sCuenta & "' AND T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    & "AND T2.nConsCod = " & gCaptacTipoTasa
'*****************************************************************************************
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
'rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta = dbCmact.CargaRecordSet(sSQL)

Set GetDatosCuentaAho = rsCta
dbCmact.CierraConexion
Set rsCta.ActiveConnection = Nothing
Set rsCta = Nothing

End Function

Public Function GetProductoPersona(ByVal sCuenta As String) As ADODB.Recordset
Dim sSQL As String
Dim rsCta As ADODB.Recordset
Set dbCmact = New DConecta
If dbCmact.AbreConexion = False Then Exit Function

sSQL = " Select PP.cPersCod, P.cPersNombre Nombre, (UPPER(RTRIM(T.cConsDescripcion)) " _
    & " + SPACE(75 - len((UPPER(RTRIM(T.cConsDescripcion))))) + CONVERT(Varchar(2),PP.nPrdPersRelac)) cRelacion, P.nPersPersoneria, Obligatorio=case when pp.cobligatorio='S' then 'SI' when (pp.cobligatorio='N' or pp.cobligatorio is null) then 'NO' else 'OPCIONAL' end   " _
    & " FROM " & sDBPersona & "Persona P INNER JOIN ProductoPersona PP ON " _
    & " P.cPersCod = PP.cPersCod INNER JOIN " & sDBComunes & "Constante T " _
    & " ON PP.nPrdPersRelac = T.nConsValor WHERE PP.cCtaCod = '" & sCuenta & "' " _
    & " AND T.nConsCod = " & gCaptacRelacPersona & "" _
    & " Order by PP.nPrdPersRelac "
    
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient

'rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta = dbCmact.CargaRecordSet(sSQL)

Set rsCta.ActiveConnection = Nothing
Set GetProductoPersona = rsCta
Set rsCta = Nothing
End Function

Public Function ValidaSaldoCuenta(ByVal sCuenta As String, ByVal nMonto As Double, Optional nOperacion As CaptacOperacion, Optional pnComiRetOtraAge As Double, Optional pnComix31Ope As Double) As Boolean
Dim nMoneda As Moneda
Dim nProducto As Producto
'Dim clsMant As NCOMCaptaGenerales
'Dim clsPar As NCOMCaptaDefinicion
Dim rsCta As New ADODB.Recordset
Dim nSaldo As Double, nMontoMinimo As Double
Dim bOrdPag As Boolean
Dim nPersoneria As PersPersoneria
Dim nEstado As CaptacEstado
Dim nBloqueoParcial As Double
'Set clsMant = New NCOMCaptaGenerales
Set rsCta = GetDatosCuenta(sCuenta)
nProducto = CLng(Mid(sCuenta, 6, 3))
bOrdPag = False
If nProducto = gCapAhorros Then
    nSaldo = rsCta("nSaldoDisp")
    bOrdPag = rsCta("bOrdPag")
ElseIf nProducto = gCapCTS Then
    If nOperacion = gCTSDepChq Or nOperacion = gCTSDepEfec Or nOperacion = gCTSDepPlanRRHH Or nOperacion = gCTSDepTransf Or nOperacion = "220303" Then
        nSaldo = rsCta("nSaldo")
    Else
        nSaldo = rsCta("nSaldRetiro")
    End If
End If

nPersoneria = rsCta("nPersoneria")
nEstado = rsCta("nPrdEstado")
nBloqueoParcial = rsCta("nBloqueoParcial")

rsCta.Close
Set rsCta = Nothing
'Set clsMant = Nothing
nMoneda = CLng(Mid(sCuenta, 9, 1))

If nEstado <> gCapEstActiva Then
    ValidaSaldoCuenta = False
Else
    If CLng(Mid(sCuenta, 6, 3)) = gCapAhorros Then
        'Set clsPar = New NCOMCaptaDefinicion
        nMontoMinimo = GetSaldoMinimoPersoneria(nProducto, nMoneda, nPersoneria, bOrdPag)
        'Set clsPar = Nothing
    Else
        nMontoMinimo = 0
    End If
    ' Agrego montos de Comision AVMM -- 03-06-2006
    If nSaldo - nBloqueoParcial - nMontoMinimo - nMonto - pnComiRetOtraAge - pnComix31Ope >= 0 Then
        ValidaSaldoCuenta = True
    Else
        ValidaSaldoCuenta = False
    End If
End If
End Function

Public Function GetSaldoMinimoPersoneria(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal nPersoneria As PersPersoneria, ByVal bOrdPag As Boolean, _
        Optional ByVal nTpoPrograma As Integer = 0) As Double
    Dim rsPar As ADODB.Recordset
    Dim sOrdPag As String
    sOrdPag = IIf(bOrdPag, "1", "0")
    Set rsPar = New ADODB.Recordset
    Set dbCmact = New DConecta
    If dbCmact.AbreConexion = False Then Exit Function
    
    rsPar.CursorLocation = adUseClient
    sSQL = "Select nSaldoMin From CapPersParam Where nPersoneria = " & nPersoneria & " And " _
        & "nProducto = " & nProducto & " And nMoneda = " & nMoneda & " And cOrdPag = '" & sOrdPag & "' " _
        & "And nTpoPrograma = " & nTpoPrograma
    Set rsPar = dbCmact.CargaRecordSet(sSQL)
    'rsPar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsPar.ActiveConnection = Nothing
    If rsPar.EOF And rsPar.BOF Then
        GetSaldoMinimoPersoneria = 0
    Else
        GetSaldoMinimoPersoneria = rsPar("nSaldoMin")
    End If
    rsPar.Close
    Set rsPar = Nothing
End Function

Public Function GetInteres(ByVal nCapital As Double, ByVal nTasa As Double, _
            ByVal nPlazo As Long, Optional nTipoInteres As TipoCalculoInteres = TpoCalcIntSimple) As Double

Dim nIntEfe As Double
If nTipoInteres = TpoCalcIntSimple Then
    GetInteres = Round((nTasa / 36000) * nPlazo * nCapital, 2)
ElseIf nTipoInteres = TpoCalcIntCompuesto Then
    GetInteres = Round(((nTasa / 36000 + 1) ^ nPlazo - 1) * nCapital, 2)
ElseIf nTipoInteres = TpoCalcIntAdelantado Then
    nIntEfe = ((nTasa / 36000 + 1) ^ nPlazo - 1)
    nIntEfe = nIntEfe / (1 + nIntEfe)
    GetInteres = Round(nIntEfe * nCapital, 2)
End If
End Function

Public Function GetNombreTitulares(ByVal sCuenta As String, Optional pbNombreCompleto As Boolean = False, Optional pnTitulares As Integer = 0, Optional pnMargenIzquierdo As Integer = 0) As String
Dim sSQL As String
Dim rsTit As ADODB.Recordset
Dim sCadena As String
Dim sCadAux As String
Dim sCadAux1 As String
Dim sCadAux2 As String

Dim nPosApe As Long
Dim nPosNom  As Long
Dim sCadAux3 As String
Dim nContador As Long

Dim lstOpe As String * 30

Dim lnNumTit As Integer
Dim oFuncion As New COMFunciones.FCOMCadenas
Set dbCmact = New DConecta
If dbCmact.AbreConexion = False Then Exit Function

sSQL = "SELECT P.cPersNombre FROM " & sDBPersona & "Persona P INNER JOIN ProductoPersona PP " _
    & "INNER JOIN Producto D ON PP.cCtaCod = D.cCtaCod ON P.cPersCod = PP.cPersCod " _
    & "WHERE D.cCtaCod = '" & sCuenta & "' AND PP.nPrdPersRelac IN (" & gCapRelPersTitular & "," & gGiroRelPersRemitente & ")"
Set rsTit = New ADODB.Recordset
rsTit.CursorLocation = adUseClient
Set rsTit = dbCmact.CargaRecordSet(sSQL)
'rsTit.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsTit.ActiveConnection = Nothing

If pbNombreCompleto Then
    If Not (rsTit.EOF And rsTit.BOF) Then
        lnNumTit = 0
        sCadena = ""
        While Not rsTit.EOF And lnNumTit < pnTitulares
            If sCadena = "" Then
                sCadena = sCadena & oFuncion.PstaNombre(rsTit("cPersNombre"), False) & Chr(10)
            Else
                sCadena = sCadena & Space(pnMargenIzquierdo) & oFuncion.PstaNombre(rsTit("cPersNombre"), False) & Chr(10)
            End If
            lnNumTit = lnNumTit + 1
            rsTit.MoveNext
        Wend
    End If
    GetNombreTitulares = sCadena
    
    rsTit.Close
    Set rsTit = Nothing
    Exit Function
End If

If rsTit.EOF And rsTit.BOF Then
    GetNombreTitulares = ""
Else
    nContador = 1
    lstOpe = oFuncion.PstaNombre(rsTit("cPersNombre"), False)
    sCadAux = Trim(lstOpe)
    sCadAux3 = rsTit("cPersNombre")
    rsTit.MoveNext
    If Not rsTit.EOF Then sCadAux = ""
    Do While Not rsTit.EOF And nContador <> 4
        If sCadAux <> "" Then
            sCadAux3 = rsTit("cPersNombre")
        End If
        nPosApe = InStr(1, sCadAux3, "/", vbTextCompare)
        nPosNom = InStr(1, sCadAux3, ",", vbTextCompare)
        sCadAux1 = Left(Left(sCadAux3, nPosApe), 18)
        sCadAux2 = Left(Mid(sCadAux3, nPosNom + 1), 18)
        nPosNom = InStr(1, Trim(sCadAux2), " ", vbTextCompare)
        If nPosNom <> 0 Then
            sCadAux2 = Mid(sCadAux2, 1, nPosNom)
        End If
        sCadAux2 = Trim(sCadAux2)
        If sCadAux <> "" Then
            If nContador Mod 2 = 0 Then
                sCadAux = sCadAux & "*" & Left(sCadAux1 & sCadAux2, 20) & Space(20 - Len(Left(sCadAux1 & sCadAux2, 20)))
            Else
                sCadAux = sCadAux & "," & Left(sCadAux1 & sCadAux2, 20) & Space(20 - Len(Left(sCadAux1 & sCadAux2, 20)))
            End If
            nContador = nContador + 1
            rsTit.MoveNext
        Else
            sCadAux = Left(sCadAux1 & sCadAux2, 20) & Space(23 - Len(Left(sCadAux1 & sCadAux2, 20)))
        End If
    Loop
    GetNombreTitulares = sCadAux
End If
'rsTit.Close
Set oFuncion = Nothing
Set rsTit = Nothing
End Function

Public Function PstaNombre(psNombre As String, Optional pbNombApell As Boolean = False) As String
Dim Total As Long
Dim Pos As Long
Dim CadAux As String
Dim lsApellido As String
Dim lsNombre As String
Dim lsMaterno As String
Dim lsConyugue As String
Dim CadAux2 As String
Dim posAux As Integer
Dim lbVda As Boolean
lbVda = False
Total = Len(Trim(psNombre))
Pos = InStr(psNombre, "/")
If Pos <> 0 Then
    lsApellido = Left(psNombre, Pos - 1)
    CadAux = Mid(psNombre, Pos + 1, Total)
    Pos = InStr(CadAux, "\")
    If Pos <> 0 Then
        lsMaterno = Left(CadAux, Pos - 1)
        CadAux = Mid(CadAux, Pos + 1, Total)
        Pos = InStr(CadAux, ",")
        If Pos > 0 Then
            CadAux2 = Left(CadAux, Pos - 1)
            posAux = InStr(CadAux, "VDA")
            If posAux = 0 Then
                lsConyugue = CadAux2
            Else
                lbVda = True
                lsConyugue = CadAux2
            End If
        Else
            lsMaterno = CadAux
        End If
    Else
        CadAux = Mid(CadAux, Pos + 1, Total)
        Pos = InStr(CadAux, ",")
        If Pos <> 0 Then
            lsMaterno = Left(CadAux, Pos - 1)
            lsConyugue = ""
        Else
            lsMaterno = CadAux
        End If
    End If
    lsNombre = Mid(CadAux, Pos + 1, Total)
    If pbNombApell = True Then
        If Len(Trim(lsConyugue)) > 0 Then
            PstaNombre = Trim(lsNombre) & " " & Trim(lsApellido) & " " & Trim(lsMaterno) & IIf(lbVda = False, " DE ", " ") & Trim(lsConyugue)
        Else
            PstaNombre = Trim(lsNombre) & " " & Trim(lsApellido) & " " & Trim(lsMaterno)
        End If
    Else
        If Len(Trim(lsConyugue)) > 0 Then
            PstaNombre = Trim(lsApellido) & " " & Trim(lsMaterno) & IIf(lbVda = False, " DE ", " ") & Trim(lsConyugue) & " " & Trim(lsNombre)
        Else
            PstaNombre = Trim(lsApellido) & " " & Trim(lsMaterno) & " " & Trim(lsNombre)
        End If
    End If
Else
    PstaNombre = Trim(psNombre)
End If
End Function

'Prepara una cadena especial (cadena con caracteres con tilde y/o otros)
' para que se imprima en el modo FREEFILE.
Public Function ImpreCarEsp(ByVal vCadena As String) As String
    vCadena = Replace(vCadena, "á", Chr(160), , , vbTextCompare)
    vCadena = Replace(vCadena, "é", Chr(130), , , vbTextCompare)
    vCadena = Replace(vCadena, "í", Chr(161), , , vbTextCompare)
    vCadena = Replace(vCadena, "ó", Chr(162), , , vbTextCompare)
    vCadena = Replace(vCadena, "ú", Chr(163), , , vbTextCompare)
    vCadena = Replace(vCadena, "ñ", Chr(164), , , vbTextCompare)
    vCadena = Replace(vCadena, "Ñ", Chr(165), , , vbTextCompare)
    vCadena = Replace(vCadena, "°", Chr(248), , , vbTextCompare)
    vCadena = Replace(vCadena, "¦", Chr(179), , , vbTextCompare)
    ImpreCarEsp = vCadena
End Function

Public Function GetCaptaOperacion(psOpeCod As String) As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set dbCmact = New DConecta
    Set rs = New ADODB.Recordset
    
    If dbCmact.AbreConexion = False Then Exit Function
    
    sql = "Select IsNull(cOpeDesc,'') cOpeDesc From OpeTpo where copecod = '" & Trim(psOpeCod) & "'"
    'rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rs = dbCmact.CargaRecordSet(sql)
    If rs.EOF And rs.BOF Then
        GetCaptaOperacion = ""
    Else
        GetCaptaOperacion = rs!cOpeDesc
    End If
    dbCmact.CierraConexion
End Function

Public Function LeeConstSistema(ByVal psConstSistCod As ConstSistemas) As String
    Dim rsVar As Recordset
    Dim sSQL As String
    Dim oCon  As DConecta
    Set oCon = New DConecta
    
    If oCon.AbreConexion = False Then Exit Function
    sSQL = "Select nConsSisDesc, nConsSisValor From ConstSistema where nConsSisCod =" & psConstSistCod & ""
    Set rsVar = New Recordset
    Set rsVar = oCon.CargaRecordSet(sSQL)
    LeeConstSistema = ""
    If Not rsVar.EOF And Not rsVar.BOF Then
        LeeConstSistema = rsVar("nConsSisValor")
    End If
    rsVar.Close
    Set rsVar = Nothing
    Set oCon = Nothing
End Function

'******** Captaciones **********'
'********* Negocios **********'
Public Sub CapOpeAhoCMACPIT(ByVal sMovNro As String, ByVal nMoneda As Moneda, ByVal nOperacion As CaptacOperacion, _
            ByVal sDescOperacion As String, ByVal nMonto As Double, ByVal sCuenta As String, ByVal nSaldo As Double, ByVal sCodCmac As String, _
            ByVal sNomCmac As String, Optional sNomAge As String = "", Optional nComision As Double = 0, Optional pnMovNro As Long = 0, _
            Optional pnMovNroCom As Long = 0, Optional psGlosa As String)
            
    'Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    'Dim clsCap As DFunciones.dFuncionesVarias
    Dim sMsgOpe As String, sCodUser As String, sCodage As String
    Dim dFecSis As Date
    Dim bTrans As Boolean
    Dim nMovNro As Long
    'Dim oCont As COMNContabilidad.NCOMContFunciones
    Dim nTempo As Currency
'
'    Dim sMovNroComITF As String
'    Dim nMovNroComITF As String
    
    Dim nBan As Boolean
    'Dim oVarImpre As New COMFunciones.FCOMVarImpresion
    
    dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
    sCodUser = Right(sMovNro, 4)
    sCodage = Mid(sMovNro, 18, 2)
    bTrans = False
    
    'Inicia la transaccion
    'Set clsCap = New DFunciones.dFuncionesVarias
    On Error GoTo ErrGraba
    'clsCap.dbCmact.BeginTrans
    'clsCap.bTransaccion = True
    
    bTrans = True
    If psGlosa = "" Then
        sGlosa = sDescOperacion & ", " & sNomCmac & ", Cuenta = " & sCuenta
    Else
        sGlosa = sGlosa & Chr$(13) & sDescOperacion & ", " & sNomCmac & ", Cuenta = " & sCuenta
    End If
    
    'clsCap.AgregaMov sMovNro, nOperacion, sGlosa
    Call AgregaMov(sMovNro, nOperacion, sGlosa)
    nMovNro = GetnMovNro(sMovNro)
    
    'ALPA 20081010*************
    pnMovNro = nMovNro
'    AgregaMovCMAC nMovNro, sCodCmac, nMoneda, nOperacion, nMonto, sDocumento, sCuenta, sCliente
'    If Len(Trim(sDocumento)) > 0 And nTpoDoc = TpoDocOrdenPago And nOperacion = gCMACOTAhoRetOP Then
'        AgregaMovDoc nMovNro, nTpoDoc, sDocumento, dFecSis
'    End If
    'xxx eliminado una transacc clsCap.dbCmact.CommitTrans
    'Fin transaccion
    
    'bTrans = False
    Dim bSaldo As Boolean

    If nComision > 0 Then
        Dim sMovNroCom As String
        Dim nMovNroCom As Long
        'Set oCont = New COMNContabilidad.NCOMContFunciones
        
        'xxx Agregado
        'sMovNroCom = oCont.IncrementaMovNro(sMovNro)
        sMovNroCom = IncrementaMovNro(sMovNro)
            
        'xxx eliminado sMovNroCom = oCont.GeneraMovNro(dFecSis, sCodAge, sCodUser)
        Set oCont = Nothing
        
        'Inicio Transaccion
        'xxx eliminado clsCap.dbCmact.BeginTrans
        
        AgregaMov sMovNroCom, gOtrOpeComisionCMACLlam, sGlosa
        nMovNroCom = GetnMovNro(sMovNroCom)
        AgregaMovOpeVarias nMovNroCom, nComision, sCuenta, sCodCmac, nMonedaComision
        AgregaMovRef nMovNro, nMovNroCom
        
'        If pnITFValorCom > 0 Then
'            sMovNroComITF = Left(nMovNroCom, 2) & "2" & Right(nMovNroCom, 4)
'            AgregaMov sMovNroComITF, gITFCobroEfectivo, sGlosa
'            nMovNroComITF = clsCap.GetnMovNro(sMovNroComITF)
'            AgregaMovOpeVarias nMovNroComITF, pnITFValorCom, sCuenta, sCodCmac, nMonedaComision
'            AgregaMovRef nMovNroCom, nMovNroComITF
'        End If
        pnMovNroCom = nMovNroCom
    End If

'    If sCuenta <> "" Then
'        Dim sMovNroReg As String
'        Dim nMovNroReg As Long
'        'Set oCont = New COMNContabilidad.NCOMContFunciones
'
'        'xxx agregado
'        If Len(Trim(sMovNroCom)) > 0 Then
'            sMovNroReg = IncrementaMovNro(sMovNroCom)
'        Else
'            sMovNroReg = IncrementaMovNro(sMovNro)
'        End If
        
        'Set oCont = Nothing
'        Select Case nOperacion
'            Case gCMACOTAhoDepEfec
'                'ALPA 20081009***************************************************
'                nTempo = CapAbonoCuentaAho(psCuenta, nMonto, gAhoDepRegCMACLlam, sMovNroReg, sGlosa, , , , , , , , , sNomAge, , , True, , , , , , , , , , , , , , , , False)
'                If nTempo = 0 Then
'                  ' If clsCap.bTransaccion = False Then clsCap.dbCmact.RollbackTrans
'                    'clsCap.dbCmact.RollbackTrans
'                    'Set clsCap = Nothing
'                    Exit Sub
'                End If
'            Case gCMACOTAhoRetEfec, gCMACOTAhoRetOP
'                'ALPA 20081009***************************************************
'                'By capi 23022009 se agrego paramerto lsMensajeValidacion
'                'nTempo = CapCargoCuentaAho(sCuentaAho, nMonto, gAhoRetRegCMACLlam, sMovNroReg, sGlosa, , , , , , , , , sNomAge, sLpt, psPersLavDinero, True, , , , clsCap.dbCmact, , , , , , , , psPersBenLavDinero, , , , , , False, psTitPersLavDinero, psOrdPersLavDinero, psReaPersLavDinero, psBenPersLavDinero, psVisPersLavDinero)
'                'nTempo = CapCargoCuentaAho(sCuentaAho, nMonto, gAhoRetRegCMACLlam, sMovNroReg, sGlosa, , , , , , , , , sNomAge, sLpt, psPersLavDinero, True, , , , clsCap.dbCmact, , , , , , , , psPersBenLavDinero, , , , , , False, psTitPersLavDinero, psOrdPersLavDinero, psReaPersLavDinero, psBenPersLavDinero, psVisPersLavDinero)
'                nTempo = CapCargoCuentaAho(psCuenta, nMonto, 200330, sMovNroReg, "", "", "", "")
'                If nTempo = 0 Then
'                    'If clsCap.bTransaccion = False Then clsCap.dbCmact.RollbackTrans
'                    'clsCap.dbCmact.RollbackTrans Comentado Por MAVM Porq sale error
'                    Set clsCap = Nothing
'                    Exit Sub
'                End If
'        End Select
'        nMovNroReg = GetnMovNro(sMovNroReg)
'        AgregaMovRef nMovNro, nMovNroReg
'    End If

    'clsCap.dbCmact.CommitTrans
    'clsCap.bTransaccion = False
    
    'Imprime Documentos
    'MsgBox "Se imprimirán las boletas de regularización", vbInformation, "Aviso"
'    Dim vVez As Integer
'
'    vVez = 1
'    nBan = True
    
'    If nOperacion = gCMACOTAhoDepEfec Or nOperacion = gCMACOTAhoDepChq Then
'        bSaldo = False
'    Else
'        bSaldo = True
'    End If
    
'    psImpBoleta = psImpBoleta & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'    If sDocumento = "NULL" Then
'        'psImpBoleta = psImpBoleta & ImprimeBoletaCMAC("OPERACION CMAC LLAMADA", sDescOperacion, Trim(nMonto), sCliente, sCuenta, "", nSaldo, nMoneda, "", nExtracto, nSaldo, True, True, , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, True, sCodAge, bSaldo) & oImp.gPrnSaltoLinea
'    Else
'        'psImpBoleta = psImpBoleta & ImprimeBoletaCMAC("OPERACION CMAC LLAMADA", sDescOperacion & " N° " & sDocumento, Trim(nMonto), sCliente, sCuenta, "", nSaldo, nMoneda, "", nExtracto, nSaldo, True, True, , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, True, sCodAge, bSaldo) & oImp.gPrnSaltoLinea
'    End If
'
'    psImpBoleta = psImpBoleta & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'    psImpBoleta = psImpBoleta & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'    psImpBoleta = psImpBoleta & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'    'psImpBoleta = psImpBoleta & ImprimeBoletaCMAC("OPERACION CMAC LLAMADA", "Comision Ope Llam", Trim(nComision), sCliente, sCuenta, "", 0, nMonedaComision, "", 0, 0, True, True, , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, True, sCodAge, False) & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'    psImpBoleta = psImpBoleta & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
'    'Impresion de Boleta de Regulariazacion
    'psImpBoleta = psImpBoleta & ImprimeBoletaCMACRegula("OPERACION CMAC LLAMADA ", sDescOperacion, Trim(nMonto), sCliente, sCuenta, "", 0, nMoneda, "", 0, 0, True, True, , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, True, sCodAge, bSaldo, psCuentaPers) & oImp.gPrnSaltoLinea
   
    
'Set clsCap = Nothing

Exit Sub
ErrGraba:
    'If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
'    MsgBox Err.Description, vbExclamation, "Error"
    Err.Raise Err.Number, "Error", Err.Description
    
End Sub

Public Function CapCargoCuentaAho(ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As Double, ByVal sMovNro As String, ByVal sGlosa As String, _
            Optional nTipoDoc As String, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional bOPExiste As Boolean = False, _
            Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
            Optional sCodCmac As String = "", Optional sNomCmac As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "", Optional bTransaccion As Boolean = False, _
            Optional psCtaBanco As String = "", Optional pnMoneda As Integer, _
            Optional ByVal psCodCmac As String, Optional pCon As ADODB.Connection = Nothing, _
            Optional ByVal psCodAge As String, Optional ByVal pbProcesaAutorizacion As Boolean = False, _
            Optional pbITFAplica As Boolean = True, Optional pnITFValor As Currency = 0, Optional pbITFAsumido As Boolean = False, _
            Optional psItfOperacion As String = "", Optional ByVal sOPedesc As String = "", Optional ByVal sBenPersLavDinero As String = "", _
            Optional ByRef psMensajeValidacion As String = "", Optional ByRef psImpBoleta As String, Optional ByRef psITFImpBoleta As String, _
            Optional ByVal pnComiRetOtraAge As Double = 0, Optional pnComiRetxMaxOpe As Double = 0, Optional pbImpTMU As Boolean = False, _
            Optional psTitPersLavDinero As String = "", Optional psOrdPersLavDinero As String = "", Optional psReaPersLavDinero As String = "", _
            Optional psBenPersLavDinero As String = "", Optional psVisPersLavDinero As String = "", _
            Optional pnMovNro As Long = 0) As Currency
            'ALPA 20081009***************************************************
Dim lnOperacion As CaptacOperacion
Dim lnMonto As Double

'Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
'Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim rsCta As ADODB.Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nSaldoInac As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean, bInmovilizada As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, i As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
'Dim oCon As COMDConstSistema.NCOMConstSistema
'Set oCon = New COMDConstSistema.NCOMConstSistema
Dim lnMontoCV As Currency
Dim lnSaldoCont As Double
Dim lnSaldoDisp As Double
Dim bImpreSaldos As Boolean
Dim bMaxOpeRet As Boolean
Dim nSumaComiRet As Double

Dim lsCodCliente As String

'Dim oGen As New COMDConstSistema.DCOMGeneral

'Para el manejo del Tipo de Cambio
Dim lnTipCambio As Currency
Dim lnTipCambioV As Currency
Dim lnTipCambioC As Currency

Dim oImpre As New COMFunciones.FCOMImpresion
Dim nComisionesRetiro As Double

Dim pnTipoITF As Integer

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False

'Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento

Dim lsrt As New ADODB.Recordset
'Obtiene los datos para el calculo
'Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set rsCta = GetDatosCuentaAho(sCuenta)
Set lsrt = GetProductoPersona(sCuenta)
Set clsMant = Nothing

lsCodCliente = lsrt("cPersCod")

nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
bInmovilizada = IIf(rsCta("bInmovilizada") = 0, False, True)

nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion
'Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento

On Error GoTo ErrGraba

If Not pCon Is Nothing Then
    clsCap.SetConexion pCon
    clsCap.bTransaccion = True
    bTransaccion = True
End If

'If Not bTransaccion Then clsCap.dbCmact.BeginTrans
bTrans = True

Randomize

For i = 0 To Rnd(2000) * 1000
Next i

If Not ValidaSaldoCuenta(sCuenta, nMonto, , pnComiRetOtraAge, pnComiRetxMaxOpe) Then 'Valida el saldo de la cuenta nuevamente
    clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    'MsgBox "Cuenta NO Posee Saldo Suficiente", vbInformation, "Aviso"
    psMensajeValidacion = "Cuenta NO Posee Saldo Suficiente o Estado es diferente a Activa"
    CapCargoCuentaAho = 0
    Exit Function
End If

If bActivaCta Then
    If bInactiva Then
        nSaldoInac = nSaldoCnt
    End If
End If

If Not ValidaSaldoCuenta(sCuenta, nMonto, , pnComiRetOtraAge, pnComiRetxMaxOpe) Then 'Valida el saldo de la cuenta nuevamente
    clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    psMensajeValidacion = "Cuenta NO Posee Saldo Suficiente"
    CapCargoCuentaAho = 0
    Exit Function
End If

'----------------------------------------

'Calcula los intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
If nDiasTranscurridos < 0 Then
   nDiasTranscurridos = 0
End If
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)

dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt - nMonto

AgregaMov sMovNro, nOperacion, sGlosa
nMovNro = GetnMovNro(sMovNro)
pnMovNro = nMovNro
ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
nSaldoDisp = nSaldoDisp - nMonto

lnSaldoDisp = nSaldoDisp
lnSaldoCont = nSaldoCnt

If pbITFAplica And pnITFValor > 0 Then
    If pbITFAsumido Then
       ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, bNumExtracto
       AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
       AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFAsumido, pnITFValor
    Else
    
        If psItfOperacion = gITFCobroEfectivo Then 'cuando se cobra en efectivo
            pnTipoITF = 1
            ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, bNumExtracto
            AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
            AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFCliente, pnITFValor
        Else
            pnTipoITF = 2
            ActualizaCargoCaptacion sCuenta, nMonto + pnITFValor, nMonto + pnITFValor, nIntGanado, dUltMov, sMovNro, bNumExtracto
            AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp - pnITFValor, nSaldoCnt - pnITFValor
            AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFCliente, pnITFValor
        End If
    End If
Else
    ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, bNumExtracto
End If

' *** COBRO DE COMISION X RETIROS  ***
nComisionesRetiro = pnComiRetOtraAge + pnComiRetxMaxOpe

If nComisionesRetiro <> 0 Then
     ActualizaCargoCaptacion sCuenta, nComisionesRetiro, nComisionesRetiro, nIntGanado, dUltMov, sMovNro, bNumExtracto
     
     If pnComiRetOtraAge <> 0 Then
        If pnComiRetxMaxOpe <> 0 Then
            pnComiRetOtraAge = pnComiRetOtraAge + pnComiRetxMaxOpe
            bMaxOpeRet = True
        End If
        AgregaMovCap nMovNro, gAhoRetComision, sCuenta, pnComiRetOtraAge, nSaldoDisp - pnITFValor - pnComiRetOtraAge, nSaldoCnt - pnITFValor - pnComiRetOtraAge
        AgregaMovCapDet nMovNro, gAhoRetComision, sCuenta, gConcOpeRetOtraAge, pnComiRetOtraAge
     End If
     If bMaxOpeRet = False Then
        If pnComiRetxMaxOpe <> 0 Then
           AgregaMovCap nMovNro, gAhoRetComision, sCuenta, pnComiRetxMaxOpe, nSaldoDisp - pnITFValor - pnComiRetxMaxOpe, nSaldoCnt - pnITFValor - pnComiRetxMaxOpe
           AgregaMovCapDet nMovNro, gAhoRetComision, sCuenta, gConcOpeRetxMaxOpe, pnComiRetxMaxOpe
        End If
     End If
End If

If sNroDoc <> "" Then
    If nTipoDoc = TpoDocOrdenPago Then 'Si es con orden de pago
        
        Dim nEstadoOP As CaptacOrdPagoEstado
        If nOperacion <> gAhoRetComOrdPagDev Then
            nEstadoOP = gCapOPEstCobrada
        Else
            nEstadoOP = gCapOPEstRechazada
        End If
        If nOperacion = gAhoRetOPCanje Then
            sMsgOpe = "Retiro OP Canje"
        ElseIf nOperacion = gAhoRetOP Then
            sMsgOpe = "Retiro OP"
        ElseIf nOperacion = gAhoOPCertificacion Then
            sMsgOpe = "Certif. OP"
            nEstadoOP = gCapOPEstCertifiCada
        End If
        
        AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, Right(sCodIF, 13), sMovNro, nMovNro, nMonto, nEstadoOP
    End If
Else
    If nOperacion = gAhoDctoEmiOP Then
        sMsgOpe = "Retiro Emisión OP"
    ElseIf nOperacion = gAhoRetTransf Then
        sMsgOpe = "Retiro Transferencia"
    ElseIf nOperacion = gAhoRetEfec Then
        sMsgOpe = "Retiro Efectivo"
    ElseIf nOperacion = gAhoRetRetencionJudicial Then
        sMsgOpe = "Retiro Retención Jud."
    ElseIf nOperacion = gAhoRetDuplicadoTarj Then
        sMsgOpe = "Retiro Dup.Tarj."
    ElseIf nOperacion = gAhoRetComServHidrandina Then
        sMsgOpe = "Retiro Serv Hidrandina"
    ElseIf nOperacion = gAhoRetComServSEDALIB Then
        sMsgOpe = "Retiro Serv Sedalib"
    ElseIf nOperacion = gAhoRetComServEDELNOR Then
        sMsgOpe = "Retiro Serv Edelnor"
    ElseIf nOperacion = gAhoRetEmiChq Then
        sMsgOpe = "Retiro Emisión Cheque"
    ElseIf nOperacion = gAhoRetEmiChqCanjeOP Then
        sMsgOpe = "Ret Emi.Chq Canje OP"
    ElseIf nOperacion = gAhoDctoEmiExt Then
        sMsgOpe = "Ret Emision Ext."
    ElseIf nOperacion = gAhoRetOtrosConceptos Then
        sMsgOpe = "Retiro otros Conceptos"
    ElseIf nOperacion >= 200331 And nOperacion <= 200357 Then
        sMsgOpe = Trim(sOPedesc)
        
    End If
    
End If
AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
ActualizadUltContactoAho sCuenta, dFecSis

If bInmovilizada Then
    If oCon.LeeConstSistema(gConstSistBitCtrCtasInmovil) = 1 Then
        AgregaMovCap nMovNro, gAhoEstInmovAct, sCuenta, nSaldoInac, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro, gAhoEstInmovAct, sCuenta, gConcCapital, nSaldoInac
        ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
        ActualizaInmovilizadaAct sCuenta
        ActualizaInactivaAct sCuenta
    End If
End If

If bInactiva And Not bInmovilizada Then
    AgregaMovCap nMovNro, gAhoEstInacAct, sCuenta, nSaldoInac, nSaldoDisp, nSaldoCnt
    AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
    ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
    ActualizaInactivaAct sCuenta
End If

AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
If sCodCmac <> "" Then
    AgregaMovCMAC nMovNro, sCodCmac, CLng(Mid(sCuenta, 9, 1)), nOperacion, nMonto
End If
If nOperacion = gAhoRetTransf Or nOperacion = gAhoRetEmiChq Then
    AgregaMovTranferBanco nMovNro, sCodIF, psCtaBanco, pnMoneda

    If pnMoneda <> Mid(sCuenta, 9, 1) Then
        'GetTipCambio dFecSis
        lnTipCambio = oGen.GetTipCambio(dFecSis, TCFijoMes)
        lnTipCambioC = oGen.GetTipCambio(dFecSis, TCCompra)
        lnTipCambioV = oGen.GetTipCambio(dFecSis, TCVenta)
        
        If Mid(sCuenta, 9, 1) = gMonedaNacional Then
            lnMontoCV = Round((nMonto / lnTipCambioV) * lnTipCambio, 2)
            clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeVenta, lnMontoCV
            clsCap.AgregaMovTipoCambio nMovNro, lnTipCambioV
            If lnTipCambioV > lnTipCambio Then
                clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - nMonto)
            Else
                clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - nMonto)
            End If
        Else
            lnMontoCV = Round(nMonto * lnTipCambio, 2)
            clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeVenta, lnMontoCV
            clsCap.AgregaMovTipoCambio nMovNro, lnTipCambioC
            If lnTipCambioC < lnTipCambio Then
                clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - Round(nMonto * lnTipCambioC, 2))
            Else
                clsCap.AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - Round(nMonto * lnTipCambioC, 2))
            End If
        End If
    End If
        
End If

UltimaActualizacionCuenta sCuenta, sMovNro

'If Not bTransaccion Then clsCap.dbCmact.CommitTrans
bTrans = False

CapCargoCuentaAho = nSaldoCnt

Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String
sTipDep = IIf(Mid(sCuenta, 9, 1) = "1", "SOLES", "DOLARES")
sCodOpe = Trim(nOperacion)
sModDep = sMsgOpe
sTipApe = "RETIRO AHORROS"
'Dim clsTit As NCOMCaptaGenerales
'Set clsTit = New NCOMCaptaGenerales

sNomTit = oImpre.ImpreCarEsp(GetNombreTitulares(sCuenta))
Set clsTit = Nothing

If pbITFAplica And pnITFValor > 0 And pbITFAsumido = False Then
    
        If (Val(psItfOperacion) = gITFCobroCargo Or Val(psItfOperacion) = gITFCobroCMACCargo) Then
            nSaldoDisp = nSaldoDisp - pnITFValor
            nSaldoCnt = nSaldoCnt - pnITFValor
            
        End If
End If

If nComisionesRetiro <> 0 Then
    nSaldoDisp = nSaldoDisp - nComisionesRetiro
    nSaldoCnt = nSaldoCnt - nComisionesRetiro
End If

If sMsgOpe = "" Then sMsgOpe = GetCaptaOperacion(Trim(Str(nOperacion)))

'Do
    If sNroDoc <> "" Then
        Select Case nTipoDoc
            Case TpoDocOrdenPago
                If sCodCmac <> "" Then
                    sTipApe = "RETIRO AHORROS CMACMAYNAS"
                    psImpBoleta = psImpBoleta & ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, False, False, , , , True, , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, , psCodCmac, , , bImpreSaldos, CStr(nOperacion), , , , , , , , False, pnTipoITF, CDbl(pnITFValor), True, pnComiRetOtraAge, pnComiRetxMaxOpe, pbImpTMU)
                Else
                    psImpBoleta = psImpBoleta & ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, False, False, , , , True, , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCmac, , , bImpreSaldos, CStr(nOperacion), , , , , , , , False, pnTipoITF, CDbl(pnITFValor), True, pnComiRetOtraAge, pnComiRetxMaxOpe, pbImpTMU)
                End If
        End Select
    Else

        If sCodCmac <> "" Then
            sTipApe = "RETIRO AHORROS CMACMAYNAS"
            psImpBoleta = psImpBoleta & ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, False, False, , , , True, , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, , psCodCmac, , , bImpreSaldos, CStr(nOperacion), , , , , , , , True, pnTipoITF, CDbl(pnITFValor), True, pnComiRetOtraAge, pnComiRetxMaxOpe, pbImpTMU)
        Else
            psImpBoleta = psImpBoleta & ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sMsgOpe), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, False, False, , , , True, , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCmac, , , bImpreSaldos, CStr(nOperacion), , , , , , , , True, pnTipoITF, CDbl(pnITFValor), True, pnComiRetOtraAge, pnComiRetxMaxOpe, pbImpTMU)
        End If
    End If

Exit Function
ErrGraba:
    If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "OCURRIO UN ERROR, VERIFIQUE ESTA OPERACION", "Error"
    'MsgBox "OCURRIO UN ERROR, VERIFIQUE ESTA OPERACION ERROR:" & Err.Description, vbOKOnly + vbInformation, "AVISO"
    CapCargoCuentaAho = 0
End Function

'*********** Capataciones ****************'
Public Function ImprimeBoleta(ByVal psTit As String, ByVal psText As String, ByVal psnMovNro As String, ByVal psnMovNroCom As String, ByVal psCtaCod As String, _
                              ByVal pdFecSis As Date, ByVal sNomAge As String, ByVal psNomCMACDest As String, ByVal psNumTarj As String, ByVal psCodAutDest As String, _
                              ByVal psUser As String, ByVal psOpeCod As String, ByVal psDNI As String, ByVal pnMoneda As Integer, Optional pbImpTMU As Boolean = False) As String

    Dim sFecha As String
    Dim sHora As String
    Dim sSep As Integer
    Dim sIni As Integer
    Dim sMax As Integer
    Dim saux As Integer
    Dim sCad As String
    Dim lsNomAge As String
    Dim lsMoneda As String

    '** Para la Impresion del Codigo Antiguo en las Boletas
    Dim lscad As String

    Dim lsSalLin As String
    'Dim loConstSistema As DFunciones.dFuncionesVarias
    'Set loConstSistema = New DFunciones.dFuncionesVarias
    
    'Set loConstSistema = Nothing
    Set oImp = New NImpresion
    lsSalLin = Chr$(10)
ETIQ:

    On Error GoTo ERROR

    lnTope = 0 '6 'Tope de lineas en Boleta

    lsNroExt = Str(pnNumExt)
    
    If pnMoneda = 1 Then
        lsMoneda = "SOLES"
    Else
        lsMoneda = "DOLARES"
    End If

    lscad = lscad & Chr$(27) & Chr$(64)
    lscad = lscad & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
    lscad = lscad & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
    lscad = lscad & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
    lscad = lscad & Chr$(27) & Chr$(107) & Chr$(0)     'Tipo de Letra Sans Serif
    lscad = lscad & Chr$(27) & Chr$(18) ' cancela condensada
    lscad = lscad & Chr$(27) & Chr$(72) ' desactiva negrita

    sSep = 15
    sIni = 0
    sMax = 33
    saux = 5

    sFecha = Format$(pdFecSis, "dd/mm/yyyy")
    sHora = Format$(Time, "hh:mm:ss")

    lsNomAge = psNomAge
    '' Declaracion para los mensajes
    'Dim clsGen As COMDConstSistema.DCOMGeneral
    'Set clsGen = New COMDConstSistema.DCOMGeneral

    Dim oImpre As New COMFunciones.FCOMImpresion

    '' liquidacion Trujillo
    If pbImpTMU = False Then
        lscad = lscad & Chr$(27) & Chr$(71)  'Activa Negrita
        lscad = lscad & lsSalLin
        lscad = lscad & Space(sIni) & "CMAC-MAYNAS" & Space(16 + sSep + saux) & "CMAC-MAYNAS" & lsSalLin 'oImp.gPrnSaltoLinea
        lscad = lscad & lsNomAge & Space(5) & psnMovNro & Space(saux + sMax + sSep - Len(lsNomAge) - Len(psnMovNro)) & lsNomAge & Space(5) & psnMovNro & lsSalLin & lsSalLin
        lscad = lscad & sFecha & Space(2) & sHora & Space(saux + sMax + sSep - (Len(sFecha) + 2) - Len(sHora)) & sFecha & Space(2) & sHora & lsSalLin & lsSalLin
        lscad = lscad & Space(15) & psTit & Space(saux + sMax + sSep - Len(psTit)) & Space(15) & psTit & lsSalLin & lsSalLin
        lscad = lscad & psNomCMACDest & Space(10) & psCodAutDest & Space(saux + sMax + sSep - Len(psNomCMACDest) - Len(psCodAutDest)) & psNomCMACDest & Space(10) & psCodAutDest & lsSalLin
        
        Select Case psOpeCod
            Case "260503"  'Retiro
                lscad = lscad & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & Space(5) & lsMoneda & Space(16 + sSep + saux) & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & Space(5) & lsMoneda & lsSalLin
                lscad = lscad & "Nº CTA :" & psCtaCod & Space(16 + sSep + saux) & "Nº CTA :" & psCtaCod & lsSalLin
                lscad = lscad & psText & Space(16 + sSep + saux) & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "260501" 'Deposito
                lscad = lscad & Right(psNumTarj, 6) & Space(5) & lsMoneda & Space(16 + sSep + saux) & Right(psNumTarj, 6) & Space(5) & lsMoneda & lsSalLin
                lscad = lscad & "Nº CTA :" & psCtaCod & Space(16 + sSep + saux) & "Nº CTA :" & psCtaCod & lsSalLin
                lscad = lscad & psText & Space(16 + sSep + saux) & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "107001" 'Pago de Creditos
                lscad = lscad & "Nº CTA :" & psCtaCod & Space(16 + sSep + saux) & "Nº CTA :" & psCtaCod & lsSalLin
                lscad = lscad & psText & Space(16 + sSep + saux) & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "260505" 'Consulta de Cuentas de Ahorro
                lscad = lscad & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & Space(16 + sSep + saux) & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & lsSalLin & lsSalLin
                lscad = lscad & psText & Space(16 + sSep + saux) & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "260506" 'Consulta de Movimientos de Cuentas de Ahorro
                lscad = lscad & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & Space(5) & lsMoneda & Space(16 + sSep + saux) & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & Space(5) & lsMoneda & lsSalLin
                lscad = lscad & "Nº CTA :" & psCtaCod & Space(16 + sSep + saux) & "Nº CTA :" & psCtaCod & lsSalLin
                lscad = lscad & psText & Space(16 + sSep + saux) & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "107004" 'Consulta de Cuentas de Credito
                lscad = lscad & psDNI & " / " & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & Space(16 + sSep + saux) & psDNI & " / " & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & lsSalLin
                lscad = lscad & psText & Space(16 + sSep + saux) & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
        End Select
        
        lscad = lscad & "COMISION TRANSACCION" & Space(5) & "S/. 1.00" & Space(5) & psnMovNroCom & Space(16 + sSep + saux) & "COMISION TRANSACCION" & Space(5) & "S/. 1.00" & Space(5) & psnMovNroCom & lsSalLin & lsSalLin & lsSalLin
        lscad = lscad & psUser & " / " & sFecha & " : " & sHora & Space(16 + sSep + saux) & psUser & " / " & sFecha & " : " & sHora & lsSalLin
        lscad = lscad & "CARACTERISTICAS CMAC ORIGEN" & Space(16 + sSep + saux) & "CARACTERISTICAS CMAC ORIGEN"
            
        ImprimeBoleta = lscad & lsSalLin
    Else
        lscad = lscad & Chr$(27) & Chr$(71)  'Activa Negrita
        lscad = lscad & lsSalLin
        lscad = lscad & Space(sIni) & "CMAC-MAYNAS" & lsSalLin 'oImp.gPrnSaltoLinea
        lscad = lscad & lsNomAge & Space(10) & psnMovNro & lsSalLin & lsSalLin
        lscad = lscad & sFecha & Space(2) & sHora & lsSalLin & lsSalLin
        lscad = lscad & Space(15) & psTit & lsSalLin & lsSalLin
        lscad = lscad & psNomCMACDest & Space(10) & psCodAutDest & lsSalLin
        
        Select Case psOpeCod
            Case "260503"  'Retiro
                lscad = lscad & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & Space(5) & lsMoneda & lsSalLin
                lscad = "Nº CTA :" & psCtaCod & lsSalLin
                lscad = lscad & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "260501" 'Deposito
                lscad = lscad & Right(psNumTarj, 6) & Space(5) & lsMoneda & lsSalLin
                lscad = "Nº CTA :" & psCtaCod & lsSalLin
                lscad = lscad & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "107001" 'Pago de Creditos
                lscad = "Nº CTA :" & psCtaCod & lsSalLin
                lscad = lscad & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "260505" 'Consulta de Cuentas de Ahorro
                lscad = lscad & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & lsSalLin & lsSalLin
                lscad = lscad & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "260506" 'Consulta de Movimientos de Cuentas de Ahorro
                lscad = lscad & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & Space(5) & lsMoneda & lsSalLin
                lscad = "Nº CTA :" & psCtaCod & lsSalLin
                lscad = lscad & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
            Case "107004" 'Consulta de Cuentas de Credito
                lscad = lscad & psDNI & " / " & Left(psNumTarj, 4) & " **** " & Right(psNumTarj, 4) & lsSalLin
                lscad = lscad & psText & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin
        End Select
                
        lscad = lscad & "COMISION TRANSACCION" & Space(5) & "S/. 1.00" & Space(5) & psNroRecIngXCom & lsSalLin & lsSalLin & lsSalLin
        lscad = lscad & psUser & " / " & sFecha & " : " & sHora & lsSalLin
        lscad = lscad & "CARACTERISTICAS CMAC ORIGEN"

        lscad = lscad & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & lsSalLin & Chr(27) & "m" & Chr(10)
        ImprimeBoleta = lscad
    End If
Exit Function
ERROR:
    Err.Raise Err.Number, "Comprueba la conexion de su impresora... ", Err.Description
'    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
'        GoTo ETIQ
'    End If
End Function

Public Function CapAbonoCuentaAho(ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, _
        Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", Optional sCodIF As String = "", _
        Optional dFechaValor As Date, Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
        Optional sCodCmac As String = "", Optional sNomCmac As String = "", Optional sNomAge As String = "", _
        Optional sLpt As String = "", Optional sPersLavDinero As String = "", Optional bTransaccion As Boolean = False, _
        Optional bImprimeBoleta As Boolean = True, Optional pnMovNroTransfer As Long = -1, _
        Optional pnMonedaTrans As Moneda = gMonedaNacional, Optional ByVal psCodCmac As String, Optional pCon As ADODB.Connection = Nothing, _
        Optional pbITFAplica As Boolean = True, Optional pnITFValor As Double = 0, Optional pbITFAsumido As Boolean = False, Optional psItfOperacion As String = "", Optional sOPedesc As String = "", Optional ByRef pnExtracto As Long, _
        Optional sBenPersLavDinero As String, Optional ByRef psMensajeValidacion As String = "", Optional ByRef psImpBoleta As String, Optional ByRef psITFImprimeBoleta As String, Optional pbImpTMU As Boolean, _
        Optional psTitPersLavDinero As String = "", Optional psOrdPersLavDinero As String = "", Optional psReaPersLavDinero As String = "", _
        Optional psBenPersLavDinero As String = "", Optional psVisPersLavDinero As String = "", Optional pnMovNro As Long = 0) As Double
        'By Capi 12022008 se adiciono Optional sOrdPersLavDinero As String, Optional sReaPersLavDinero As String, Optional VisPersLavDinero As String
        
        
    'Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    'Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim rsCta As New ADODB.Recordset
    Dim nEstado As CaptacEstado
    Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
    Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
    Dim nSaldoInac As Double
    Dim dUltMov As Date
    Dim bInactiva As Boolean, bTrans As Boolean, bInamovilizada As Boolean
    Dim nNumExtracto As Long, nDiasTranscurridos As Long, nExtracto As Long
    Dim sMsgOpe As String, sCodUser As String, sAgencia As String
    Dim nMovNro As Long
    Dim dFecSis As Date
    Dim oCon As COMDConstSistema.NCOMConstSistema
    Set oCon = New COMDConstSistema.NCOMConstSistema
    Dim oGen As New COMDConstSistema.DCOMGeneral
    Dim oImpre As New COMFunciones.FCOMImpresion

    Dim lnSaldoDisp As Double
    Dim lnSaldoCont As Double
    Dim bImpreSaldos As Boolean

    Dim lnMontoCV As Currency

    'Para el manejo del Tipo de Cambio
    Dim lnTipCambio As Currency
    Dim lnTipCambioV As Currency
    Dim lnTipCambioC As Currency
    
    dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
    sCodUser = Right(sMovNro, 4)
    sAgencia = Mid(sMovNro, 18, 2)

    'Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento

    Dim MontoItf As Double
    nMovNro = GetnMovNro(sMovNro)

    bTrans = False
    'Obtiene los datos para el calculo
    'Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsCta = GetDatosCuentaAho(sCuenta)
    Set clsMant = Nothing
    nEstado = rsCta("nPrdEstado")
    nSaldoDisp = rsCta("nSaldoDisp")
    nSaldoCnt = rsCta("nSaldo")
    dUltMov = rsCta("dUltCierre")
    nTasa = rsCta("nTasaInteres")
    nIntAcum = rsCta("nIntAcum")
    bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
    bInamovilizada = IIf(rsCta("bInmovilizada") = 0, False, True)
    nExtracto = rsCta("nTransacc") + 1
    pnExtracto = nExtracto
    
    rsCta.Close
    Set rsCta = Nothing
    
    'Inicia la transaccion
    'Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
    
    On Error GoTo ErrGraba
    
    If Not pCon Is Nothing Then
        clsCap.SetConexion pCon
        clsCap.bTransaccion = True
        bTransaccion = True
    End If
    
    'If Not bTransaccion Then clsCap.dbCmact.BeginTrans
    bTrans = True


    If Not ValidaEstadoCuenta(sCuenta, False) Then   'Valida el estado de la cuenta nuevamente
        clsCap.dbCmact.RollbackTrans
        Set clsCap = Nothing
        psMensajeValidacion = "Cuenta NO Tiene un estado valido para la Operacion, consulte con el Asistente de Agencia."
        CapAbonoCuentaAho = 0
        Exit Function
    End If
    
    If sGlosa = "" Then sGlosa = "Abono Cuenta = " & sCuenta
    If bActivaCta Then
        If bInactiva Then
            nSaldoInac = nSaldoCnt
        End If
    End If

    'Calcula intereses
    nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
    If nDiasTranscurridos < 0 Then
        nDiasTranscurridos = 0
    End If
    nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
    dUltMov = DateAdd("d", -1, dFecSis)
    nSaldoCnt = nSaldoCnt + nMonto
    
    AgregaMov sMovNro, nOperacion, sGlosa
    
    nMovNro = GetnMovNro(sMovNro)
    
    If sNroDoc = "" Then ' Si la operacion es en efectivo
        ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
        nSaldoDisp = nSaldoDisp + nMonto
        ActualizadUltContactoAho sCuenta, dFecSis
    Else
        If nTipoDoc = TpoDocCheque And nOperacion <> "200252" Then 'Si el abono es con cheque
            nSaldoDisp = nSaldoDisp + 0
            nIntGanado = 0
            clsCap.ActualizadUltContactoAho sCuenta, dFecSis
        ElseIf nTipoDoc = TpoDocCheque And nOperacion = "200252" Then
            clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
            nSaldoDisp = nSaldoDisp + nMonto
            clsCap.ActualizadUltContactoAho sCuenta, dFecSis
        ElseIf nTipoDoc = TpoDocNotaAbono Then 'Si el abono es con nota de abono
            clsCap.ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
            nSaldoDisp = nSaldoDisp + nMonto
        End If
    End If

    '******** para Boleta ITF
    
    lnSaldoDisp = nSaldoDisp
    lnSaldoCont = nSaldoCnt
    '-----
    AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
    
    If pbITFAplica And pnITFValor > 0 Then
        If pbITFAsumido Then
            clsCap.AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
            clsCap.AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFAsumido, pnITFValor
        Else
            If psItfOperacion = gITFCobroEfectivo Then
                clsCap.AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
                clsCap.AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFCliente, pnITFValor
            Else
                clsCap.AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp - pnITFValor, nSaldoCnt - pnITFValor
                clsCap.AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFCliente, pnITFValor
            End If
        End If
    End If

    If sNroDoc = "" Then ' Si la operacion es en efectivo
        If (pbITFAplica And pnITFValor > 0) And (Not pbITFAsumido) And (Val(psItfOperacion) = gITFCobroCargo Or Val(psItfOperacion) = gITFCobroCMACCargo) Then
            ActualizaAbonoCaptacion sCuenta, nMonto - pnITFValor, nMonto - pnITFValor, nIntGanado, dUltMov, sMovNro, bNumExtracto
        Else
            ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, bNumExtracto
        End If
        
        If nOperacion = gAhoDepEfec Then
            sMsgOpe = "Depósito Efectivo"
        ElseIf nOperacion = gAhoDepTransf Then
            sMsgOpe = "Depósito Transferencia"
        ElseIf nOperacion = gAhoDepDevFondoGar Then
            sMsgOpe = "Depósito Dev Fondo Gar"
        ElseIf nOperacion = gAhoDepPagServEdelnor Then
            sMsgOpe = "Depósito Serv Edelnor"
        ElseIf nOperacion = gAhoDepPagServHidrandina Then
            sMsgOpe = "Depósito Serv Hidrandina"
        ElseIf nOperacion = gAhoDepPagServSedalib Then
            sMsgOpe = "Depósito Serv Sedalib"
        ElseIf nOperacion = gAhoDepSobCaja Then
            sMsgOpe = "Depósito Sobrante"
        ElseIf nOperacion = gAhoDepAboOtrosConceptos Then
            sMsgOpe = "Depósito Otros Conceptos"
        ElseIf nOperacion = gAhoDepDevCredPersonales Then
            sMsgOpe = "Depósito Dev Cred Pers"
        ElseIf nOperacion = "200242" Then
            sMsgOpe = "Abono por otros Servicios"
        ElseIf nOperacion = "200209" Then
            sMsgOpe = "Otros Conceptos RRHH"
        ElseIf nOperacion >= 200247 And nOperacion <= 200251 Then
            sMsgOpe = Trim(sOPedesc)
        
        End If
    Else
        If nTipoDoc = TpoDocCheque Then  'Si el abono es con cheque
        
           If nOperacion <> "200252" Then
              If (pbITFAplica And pnITFValor > 0) And (Not pbITFAsumido) And (Val(psItfOperacion) = gITFCobroCargo Or Val(psItfOperacion) = gITFCobroCMACCargo) Then
                   ActualizaAbonoCaptacion sCuenta, nMonto - pnITFValor, (pnITFValor * -1), nIntGanado, dUltMov, sMovNro, True
              Else
                   ActualizaAbonoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
              End If
           Else
              If (pbITFAplica And pnITFValor > 0) And (Not pbITFAsumido) And (Val(psItfOperacion) = gITFCobroCargo Or Val(psItfOperacion) = gITFCobroCMACCargo) Then
                    ActualizaAbonoCaptacion sCuenta, nMonto - pnITFValor, nMonto - pnITFValor, nIntGanado, dUltMov, sMovNro, bNumExtracto
              Else
                    ActualizaAbonoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
              End If
           End If
             
           If nOperacion <> "200252" Then
                 Dim rstemp1 As ADODB.Recordset
                 
                 Set rstemp1 = clsCap.GetDatosCheque(sNroDoc, sCodIF, "01")
                
                 AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMovNro, nMonto, , rstemp1!cIFCta
           End If
            
            sMsgOpe = "Dep.Chq"
        End If
    End If

    If bInamovilizada Then
        If LeeConstSistema(gConstSistBitCtrCtasInmovil) = 1 Then
            AgregaMovCap nMovNro, gAhoEstInmovAct, sCuenta, nSaldoInac, nSaldoDisp, nSaldoCnt
            AgregaMovCapDet nMovNro, gAhoEstInmovAct, sCuenta, gConcCapital, nSaldoInac
            ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
            ActualizaInmovilizadaAct sCuenta
            ActualizaInactivaAct sCuenta
        End If
    End If
    
    If bInactiva And Not bInamovilizada Then
        AgregaMovCap nMovNro, gAhoEstInacAct, sCuenta, nSaldoInac, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
        ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
        ActualizaInactivaAct sCuenta
    End If

    If pnMovNroTransfer <> -1 Then
        AgregaMovRef nMovNro, pnMovNroTransfer
        If pnMonedaTrans <> Mid(sCuenta, 9, 1) Then
            lnTipCambio = EmiteTipoCambio(dFecSis, TCFijoMes)
            lnTipCambioC = EmiteTipoCambio(dFecSis, TCCompra)
            lnTipCambioV = EmiteTipoCambio(dFecSis, TCVenta)
            
            If Mid(sCuenta, 9, 1) = gMonedaNacional Then
                ActualizaMovPendientesRend pnMovNroTransfer, Round(nMonto / lnTipCambioC, 2)
                lnMontoCV = Round((nMonto / lnTipCambioC) * lnTipCambio, 2)
                AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeCompra, lnMontoCV
                AgregaMovTipoCambio nMovNro, lnTipCambioC
                If lnTipCambioC < lnTipCambio Then
                    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - nMonto)
                Else
                    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - nMonto)
                End If
            Else
                ActualizaMovPendientesRend pnMovNroTransfer, Round(nMonto * lnTipCambioV, 2)
                lnMontoCV = Round(nMonto * lnTipCambio, 2)
                AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeVenta, lnMontoCV
                AgregaMovTipoCambio nMovNro, lnTipCambioV
                If lnTipCambioV > lnTipCambio Then
                    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - Round(nMonto * lnTipCambioV, 2))
                Else
                    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - Round(nMonto * lnTipCambioV, 2))
                End If
            End If
        Else
            ActualizaMovPendientesRend pnMovNroTransfer, nMonto
        End If
    End If

    If sCodCmac <> "" Then
        AgregaMovCMAC nMovNro, sCodCmac, CLng(Mid(sCuenta, 9, 1)), nOperacion, nMonto
    End If
    
    UltimaActualizacionCuenta sCuenta, sMovNro
    

    CapAbonoCuentaAho = nSaldoCnt

    pnMovNro = nMovNro
    '******************************************************************
    If (Not bImprimeBoleta) Or nOperacion = gAhoDepEntConv Then
            Exit Function
    End If

    '****************************************
    '*************************************************'
    'If Not bTransaccion Then clsCap.dbCmact.CommitTrans
    bTrans = False
    
    
    Dim sTipDep As String, sCodOpe As String
    Dim sModDep As String, sTipApe As String
    Dim sNomTit As String
    sTipDep = IIf(Mid(sCuenta, 9, 1) = Moneda.gMonedaNacional, "SOLES", "DOLARES")
    sCodOpe = Trim(nOperacion)
    'sModDep = sMsgOpe
    'sTipApe = "DEPOSITO AHORROS"
    

    'Dim clsTit As NCOMCaptaGenerales
    'Set clsTit = New NCOMCaptaGenerales
    'sNomTit = oImpre.ImpreCarEsp(GetNombreTitulares(sCuenta))
    'Set clsMant = Nothing
    
    If pbITFAplica And pnITFValor > 0 And pbITFAsumido = False Then
        
            If (Val(psItfOperacion) = gITFCobroCargo Or Val(psItfOperacion) = gITFCobroCMACCargo) Then
                nSaldoDisp = nSaldoDisp - pnITFValor
                nSaldoCnt = nSaldoCnt - pnITFValor
                
            End If
    End If
    
    Dim nTipoPag As Integer


    If pnITFValor > 0 Then
        If psItfOperacion = gITFCobroEfectivo Then
            nTipoPag = 1
        Else
            nTipoPag = 2
        End If
    End If
    'Do
'    If sNroDoc <> "" Then
'        Select Case nTipoDoc
'            Case TpoDocCheque
'                If sCodCmac <> "" Then
'                    psImpBoleta = psImpBoleta & ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntGanado, "Fecha Valor", nExtracto, nSaldoCnt, True, , , , , , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, , psCodCmac, sAgencia, , bImpreSaldos, CStr(nOperacion), , , , , , , , False, nTipoPag, CDbl(pnITFValor), True, , , pbImpTMU)
'                Else
'                    psImpBoleta = psImpBoleta & ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntGanado, "Fecha Valor", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCmac, sAgencia, , bImpreSaldos, CStr(nOperacion), , , , , , , , False, nTipoPag, CDbl(pnITFValor), True, , , pbImpTMU)
'                End If
'            Case TpoDocNotaAbono
'                psImpBoleta = psImpBoleta & ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCmac, sAgencia, , bImpreSaldos, CStr(nOperacion), , , , , , , , False, nTipoPag, CDbl(pnITFValor), True, , , pbImpTMU)
'        End Select
'    Else
'        If sCodCmac <> "" Then
'            sTipApe = "DEPOSITO CMAC AHORROS"
'            psImpBoleta = psImpBoleta & ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, , psCodCmac, sAgencia, , bImpreSaldos, CStr(nOperacion), , , , , , , , False, nTipoPag, CDbl(pnITFValor), True, , , pbImpTMU)
'        Else
'            psImpBoleta = psImpBoleta & ImprimeBoleta(sTipApe, oImpre.ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLpt, , psCodCmac, sAgencia, , bImpreSaldos, CStr(nOperacion), , , , , , , , False, nTipoPag, CDbl(pnITFValor), True, , , pbImpTMU)
'        End If
'    End If

Exit Function
ErrGraba:
    'If bTrans Then clsCap.dbCmact.RollbackTrans
    'MsgBox "OCURRIO UN ERROR, VERIFIQUE ESTA OPERACION ERROR:" & Err.Description, vbOKOnly + vbInformation, "AVISO"
    Err.Raise Err.Number, "OCURRIO UN ERROR, VERIFIQUE ESTA OPERACION ", Err.Description
    Set clsCap = Nothing
    'Err.Raise Err.Number, "", Err.Description
    CapAbonoCuentaAho = 0
End Function
Public Sub AgregaMov(ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
        ByVal sGlosa As String, Optional nMovEstado As MovEstado = gMovEstContabMovContable, _
        Optional nMovFlag As MovFlag = gMovFlagVigente)
Set dbCmact = New DConecta
dbCmact.AbreConexion

sSQL = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sGlosa & "'," & nMovEstado & "," & nMovFlag & ")"
    
dbCmact.Ejecutar (sSQL)
dbCmact.CierraConexion
End Sub
Public Sub AgregaMovCMAC(ByVal nMovNro As Long, ByVal sPersCod As String, _
        ByVal nMoneda As Moneda, ByVal nOperacion As CaptacOperacion, _
        ByVal nMonto As Double, Optional sDocumento As String = "NULL", _
        Optional sCuenta As String = "NULL", Optional sCliente As String = "")
Dim sIFTipo As String
Dim sCta As String, sDoc As String
Set dbCmact = New DConecta
dbCmact.AbreConexion
'sCta = "NULL"
'sDoc = "NULL"
'If sCuenta = "NULL" Then sCuenta = ""
'If sDocumento = "NULL" Then sDocumento = ""
sIFTipo = Format$(gTpoIFCmac, "00")
If sCuenta <> "NULL" Then sCta = sCuenta
If sDocumento <> "NULL" Then sDoc = sDocumento
sSQL = "INSERT MovCMAC (nMovNro,cPersCod,cIFTpo,nMoneda,cCtaIF,cDocumento,cOpeCod,nMonto, cCliente) " _
    & "VALUES (" & nMovNro & ",'" & sPersCod & "','" & sIFTipo & "'," & nMoneda & ",'" & sCta & "','" & sDoc & "','" & Trim(nOperacion) & "'," & Trim(nMonto) & ", '" & sCliente & "')"
'dbCmact.Execute sSql
dbCmact.Ejecutar (sSQL)
dbCmact.CierraConexion
End Sub
Public Sub AgregaMovDoc(ByVal nMovNro As Long, ByVal nTipoDoc As TpoDoc, _
        ByVal sDocumento As String, ByVal dFecDocumento As Date)
Dim sFecha As String
Set dbCmact = New DConecta
dbCmact.AbreConexion

sFecha = Format$(dFecDocumento, "mm/dd/yyyy")
sSQL = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
    & "VALUES (" & nMovNro & "," & nTipoDoc & ",'" & sDocumento & "','" & sFecha & "')"
dbCmact.Ejecutar (sSQL)
dbCmact.CierraConexion
End Sub
Public Sub AgregaMovOpeVarias(ByVal pnMovNro As Long, ByVal pnMovImporte As Double, _
        ByVal psNroDoc As String, ByVal psReferencia As String, ByVal pnMoneda As Moneda, _
        Optional nOperacion As CaptacOperacion)
    Dim sql As String
    Set dbCmact = New DConecta
    dbCmact.AbreConexion
    
    sql = " Insert MovOpeVarias (nMovNro,nMovImporte,cNroDoc,cReferencia,nMoneda,cOpeCod)" _
        & " Values(" & pnMovNro & "," & pnMovImporte & ",'" & psNroDoc & "','" & psReferencia & "'," & pnMoneda & ",'" & Trim(nOperacion) & "')"
    
    dbCmact.Ejecutar (sql)
    dbCmact.CierraConexion
End Sub
Public Sub AgregaMovRef(ByVal nMovNro As Long, ByVal nMovRef As Long)
Set dbCmact = New DConecta
dbCmact.AbreConexion
sSQL = "INSERT MovRef (nMovNro,nMovNroRef) " _
    & "VALUES (" & nMovNro & "," & nMovRef & ")"

dbCmact.Ejecutar (sSQL)
dbCmact.CierraConexion
End Sub

Private Sub Class_Terminate()
Set dbCmact = Nothing
End Sub
Public Function GetDatosCuenta(ByVal sCuenta As String) As ADODB.Recordset
Dim rsCta As ADODB.Recordset
Dim nProd As Producto
'Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
nProd = CInt(Mid(sCuenta, 6, 3))
'Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Select Case nProd
    Case gCapAhorros
        Set rsCta = GetDatosCuentaAho(sCuenta)
        
    Case gCapPlazoFijo
        Set rsCta = GetDatosCuentaPF(sCuenta)
        
    Case gCapCTS
        Set rsCta = GetDatosCuentaCTS(sCuenta)
        
End Select
Set GetDatosCuenta = rsCta
Set rsCta = Nothing
Set clsMant = Nothing
End Function
Public Function GetDatosCuentaCTS(ByVal sCuenta As String) As ADODB.Recordset
Dim sSQL As String
Dim rsCta As ADODB.Recordset
Set dbCmact = New DConecta
dbCmact.AbreConexion

sSQL = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado,RetAdiCTS=Isnull(A.nRetAdiCTS,0), " _
    & "(C.nSaldoDisp-Isnull(A.nRetencion,0)) as nSaldoDisp,A.nRetencion, C.nPersoneria, C.nFirmas,c.nfirmasmin ,C.nIntAcum, dUltCierre=( case when dbo.ultimodiames(dUltCierre)=1 then dUltCierre else dateadd(dd,1, dUltCierre) end )  , C.dApertura, AG.cAgeDescripcion, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, (A.nSaldRetiro-Isnull(A.nRetencion,0)) as nSaldRetiro, A.nIntSaldo, PE.cPersNombre cInstitucion, " _
    & "T.cConsDescripcion cEstado, T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.cCodInst,C.cAlias,ISNULL(C.nBloqueoParcial,0) nBloqueoParcial, IsNull(PEJUR.cPersJurSigla,'') EmpSiglas " _
    & "FROM Producto P INNER JOIN Captaciones C INNER JOIN CaptacCTS A INNER JOIN " & sDBPersona & "" _
    & "Persona PE ON A.cCodInst = PE.cPersCod ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod " _
    & "Left JOIN " & sDBPersona & "PersonaJur PEJUR ON PE.cPersCod = PEJUR.cPersCod " _
    & "WHERE P.cCtaCod = '" & sCuenta & "' AND T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    & "AND T2.nConsCod = " & gCaptacTipoTasa
'******************************************************************************************
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient

dbCmact.Ejecutar (sSQL)
dbCmact.CierraConexion

Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaCTS = rsCta
Set rsCta = Nothing
End Function

Public Function GetDatosCuentaPF(ByVal sCuenta As String) As ADODB.Recordset
    Dim sSQL As String
    Dim rsCta As ADODB.Recordset
    Set dbCmact = New DConecta
    dbCmact.AbreConexion
    'By Capi 07082008 se adiciono columna bTasaPactada
    sSQL = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
        & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, AG.cAgeDescripcion, " _
        & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.nPlazo, A.nIntPag, A.dRenovacion, A.nApertura, A.nIntAdelRev, " _
        & "A.nFormaRetiro, A.dAuxiliar, A.nDuplicado, T.cConsDescripcion cEstado, A.bBusCredPend, " _
        & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, T3.cConsDescripcion cRetiro, " _
        & "A.dUltCierreAnt, A.dAuxiliarAnt, ISNULL(CI.cCtaCodAbono,'') cCtaCodAbono, a.nformaretiro,c.calias,c.nfirmasmin,nIntDevChq,C.cUltimaActualizacion,A.bTasaPactada FROM Producto P " _
        & "INNER JOIN Captaciones C INNER JOIN CaptacPlazoFijo A LEFT JOIN CaptacCtaAboIntPF CI " _
        & "ON A.cCtaCod = CI.cCtaCod ON C.cCtaCod = A.cCtaCod ON " _
        & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
        & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
        & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Constante T3 ON A.nFormaRetiro = T3.nConsValor " _
        & "INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod WHERE P.cCtaCod = '" & sCuenta & "' AND " _
        & "T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
        & "AND T2.nConsCod = " & gCaptacTipoTasa & " AND T3.nConsCod = " & gCaptacPFFormaRetiro
        
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    
    Set rsCta = dbCmact.CargaRecordSet(sSQL)
    dbCmact.CierraConexion
    
    Set rsCta.ActiveConnection = Nothing
    Set GetDatosCuentaPF = rsCta
    Set rsCta = Nothing
End Function

Public Sub ActualizaSaldoAnteriorAho(ByVal sCuenta As String, ByVal nSaldoAnt As Double)
    Set dbCmact = New DConecta
    dbCmact.AbreConexion
    
    sSQL = "Update CaptacAhorros Set nSaldoAnterior = " & nSaldoAnt & "  WHERE cCtaCod = '" & sCuenta & "'"
    dbCmact.Ejecutar (sSQL)
    dbCmact.CierraConexion
End Sub

Public Sub ActualizaCargoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)
    Set dbCmact = New DConecta
    dbCmact.AbreConexion
    
    If bActExtracto Then
        sSQL = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
            & "WHERE cCtaCod = '" & sCuenta & "'"
    Else
        sSQL = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & " " _
            & "WHERE cCtaCod = '" & sCuenta & "'"
    End If
    dbCmact.Ejecutar (sSQL)

    sSQL = "Update Captaciones Set nSaldoDisp = nSaldoDisp - " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
         & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
         & " WHERE cCtaCod = '" & sCuenta & "'"
    dbCmact.Ejecutar (sSQL)
    dbCmact.CierraConexion
End Sub
Public Sub AgregaMovCap(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double)
Set dbCmact = New DConecta
dbCmact.AbreConexion
    
'Dim OCon As COMDConstSistema.FCOMITF
'Set OCon = New COMDConstSistema.FCOMITF
'    If Mid(nOperacion, 1, 2) = "99" Then
'        nMonto = OCon.fgTruncar(nMonto, 2)
'    End If
'Set OCon = Nothing
sSQL = "INSERT MovCap (nMovNro,cOpeCod,cCtaCod,nMonto,nSaldoDisponible,nSaldoContable) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nMonto & "," & nSaldoDisp & "," & nSaldoCnt & ")"
dbCmact.Ejecutar (sSQL)
dbCmact.CierraConexion
End Sub
Public Sub AgregaMovCapDet(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nConcepto As CaptacConcepto, ByVal nMonto As Double)
Set dbCmact = New DConecta
dbCmact.AbreConexion

'Dim OCon As COMDConstSistema.FCOMITF
'Set OCon = New COMDConstSistema.FCOMITF
'    If Mid(nOperacion, 1, 2) = "99" Then
'        nMonto = OCon.fgTruncar(nMonto, 2)
'    End If
'Set OCon = Nothing
sSQL = "INSERT MovCapDet (nMovNro,cOpeCod,cCtaCod,nConceptoCod,nMonto) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nConcepto & "," & nMonto & ")"
dbCmact.Ejecutar (sSQL)
dbCmact.CierraConexion
End Sub
Public Sub AgregaCuentaDocumento(ByVal sCuenta As String, ByVal nTpoDoc As TpoDoc, _
        ByVal sNroDoc As String, ByVal sCodIF As String, ByVal sMovNro As String, _
        ByVal nMovNro As Long, Optional nMonto As Double = 0, _
        Optional nEstadoOP As CaptacOrdPagoEstado = gCapOPEstCobrada, _
        Optional psCtaIf As String = "")
Dim sFecha As String
Set dbCmact = New DConecta
dbCmact.AbreConexion
sFecha = Mid(sMovNro, 1, 4) & "/" & Mid(sMovNro, 5, 2) & "/" & Mid(sMovNro, 7, 2) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)


Select Case nTpoDoc
    Case TpoDocCheque
        sSQL = "INSERT DocRecCapta (nTpoDoc,cNroDoc,cPersCod,cIFTpo,nMovNro,cCtaCod,nMonto,cIFCta) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "'," & nMovNro & ",'" & sCuenta & "'," & nMonto & ", '" & psCtaIf & "')"
        dbCmact.Ejecutar (sSQL)
        sSQL = "Update DocRec Set nEstado = " & gChqEstEnValorizacion & " Where cNroDoc = '" & sNroDoc & "' And " _
            & "cIFTPo = '" & Format$(gTpoIFBanco, "00") & "' And cPersCod = '" & sCodIF & "' and CIFCTA='" & psCtaIf & "'"
        dbCmact.Ejecutar (sSQL)
        sSQL = "INSERT DocRecEst (nTpoDoc,cNroDoc,cPersCod,cIFTpo,cMovNro,nEstado,CIFCTA ) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "','" & sMovNro & "'," & gChqEstEnValorizacion & ",'" & psCtaIf & "')"
        dbCmact.Ejecutar (sSQL)
    Case TpoDocOrdenPago
        sSQL = "INSERT DocRecOP (nTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers,nEstado,cMovNro) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCuenta & "'," & nMonto & ",'" & sCodIF & "'," & nEstadoOP & ",'" & sMovNro & "')"
        dbCmact.Ejecutar (sSQL)
End Select
sSQL = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
    & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
dbCmact.Ejecutar (sSQL)
End Sub
Public Sub ActualizadUltContactoAho(ByVal sCuenta As String, ByVal dFecha As Date)
    Set dbCmact = New DConecta
    dbCmact.AbreConexion
    sSQL = "Update CaptacAhorros Set dUltContacto = '" & Format(dFecha, gsFormatoFecha) & "'  WHERE cCtaCod = '" & sCuenta & "'"
    dbCmact.Ejecutar (sSQL)
    dbCmact.CierraConexion
End Sub
Public Sub ActualizaEstadoCuenta(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, _
        ByVal dFecha As Date, ByVal sMovNro As String)
Set dbCmact = New DConecta
dbCmact.AbreConexion
'Actualiza el ultimo estado a la tabla de producto
sSQL = "Update Producto Set nPrdEstado = " & nEstado & ", " _
    & "dPrdEstado = '" & Format$(dFecha, "mm/dd/yyyy") & "' WHERE " _
    & "cCtaCod = '" & sCuenta & "'"
dbCmact.Ejecutar (sSQL)
'Agrega un registro mas de estado a la historia de estados de la cuenta
sSQL = "INSERT CaptacEstado (cCtaCod,nPrdEstado,cMovNro) " _
    & "VALUES ('" & sCuenta & "'," & nEstado & ",'" & sMovNro & "')"
dbCmact.Ejecutar (sSQL)
dbCmact.CierraConexion
End Sub
Public Sub ActualizaInmovilizadaAct(ByVal sCuenta As String)
Set dbCmact = New DConecta
    dbCmact.AbreConexion
    sSQL = "Update CaptacAhorros Set bInactiva = 0,  bInmovilizada = 0 WHERE cCtaCod = '" & sCuenta & "'"
    dbCmact.Ejecutar (sSQL)
    dbCmact.CierraConexion
End Sub
Public Sub ActualizaInactivaAct(ByVal sCuenta As String)
Set dbCmact = New DConecta
    dbCmact.AbreConexion
    sSQL = "Update CaptacAhorros Set bInactiva = 0 WHERE cCtaCod = '" & sCuenta & "'"
    dbCmact.Ejecutar (sSQL)
    dbCmact.CierraConexion
End Sub
Public Sub AgregaMovTranferBanco(ByVal pnMovNro As Long, ByVal psPersCodIF As String, psCtaBco As String, ByVal pnMoneda As Moneda)
Set dbCmact = New DConecta
    dbCmact.AbreConexion
sSQL = " Insert MovTransferBco (nMovNro, cPersCodIF, cCtaBco, nMoneda)" _
     & " Values (" & pnMovNro & ",'" & psPersCodIF & "','" & psCtaBco & "'," & pnMoneda & ")"
    dbCmact.Ejecutar (sSQL)
    dbCmact.CierraConexion
End Sub
Public Sub AgregaMovTipoCambio(ByVal nMovNro As Long, ByVal pnTipoCambio As Currency)
Set dbCmact = New DConecta
dbCmact.AbreConexion
sSQL = " INSERT MovTpoCambio (nMovNro, nMovTpoCambio) " _
     & " VALUES (" & nMovNro & "," & pnTipoCambio & ")"
dbCmact.Ejecutar (sSQL)
dbCmact.CierraConexion
End Sub
Public Sub UltimaActualizacionCuenta(ByVal sCuenta As String, ByVal sMovNro As String)
Set dbCmact = New DConecta
dbCmact.AbreConexion
'Actualiza la fecha de ultima actualizacion de la cuenta de captaciones
sSQL = "Update Captaciones Set cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Ejecutar (sSQL)
dbCmact.CierraConexion
End Sub
Public Property Get gsFormatoFecha() As Variant
    gsFormatoFecha = "mm/dd/yyyy"
End Property
Public Function ValidaEstadoCuenta(ByVal sCuenta As String, Optional pbParaRetiro As Boolean = True) As Boolean
'Dim clsMant As NCOMCaptaGenerales
'Set clsMant = New NCOMCaptaGenerales
Dim rsCta As ADODB.Recordset
Set rsCta = New ADODB.Recordset
Dim nEstado As CaptacEstado

Set rsCta = GetDatosCuenta(sCuenta)

nEstado = rsCta("nPrdEstado")

rsCta.Close
Set rsCta = Nothing
'Set clsMant = Nothing

If pbParaRetiro Then
    If nEstado = gCapEstActiva Then
        ValidaEstadoCuenta = True
    Else
        ValidaEstadoCuenta = False
    End If
Else
    If nEstado = gCapEstActiva Or nEstado = gCapEstBloqRetiro Or nEstado = gCapEstBloqTotal Then
        ValidaEstadoCuenta = True
    Else
        ValidaEstadoCuenta = False
    End If
End If
End Function
Public Sub ActualizaAbonoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)
Set dbCmact = New DConecta
dbCmact.AbreConexion

If bActExtracto Then
    sSQL = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSQL = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
dbCmact.Ejecutar (sSQL)
sSQL = "Update Captaciones Set nSaldoDisp = nSaldoDisp + " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, "mm/dd/yyyy") & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
dbCmact.Ejecutar (sSQL)
dbCmact.CierraConexion
End Sub
Public Sub ActualizaMovPendientesRend(ByVal nMovNro As Long, ByVal pnMonto As Currency)
Set dbCmact = New DConecta
    dbCmact.AbreConexion
    sSQL = " UPDATE MovPendientesRend " _
         & " Set nSaldo = nSaldo - " & pnMonto & "" _
         & " Where nMovNro = " & nMovNro & ""
    dbCmact.Ejecutar (sSQL)
End Sub
Public Function GetTarifaParametro(ByVal nOperacion As CaptacOperacion, _
            ByVal nMoneda As Moneda, ByVal nParametro As CaptacParametro) As Recordset
    Dim rsPar As ADODB.Recordset
    Set dbCmact = New DConecta
    dbCmact.AbreConexion
    
    Set rsPar = New ADODB.Recordset
    rsPar.CursorLocation = adUseClient
    sSQL = "Select P.nParValor From Parametro P JOIN Tarifas T ON P.nParProd = T.nParProd And " _
        & "P.nParCod = T.nParCod Where T.cOpeCod = '" & nOperacion & "' And " _
        & "T.nMoneda = " & nMoneda & " And T.nParCod = " & nParametro
    
    Set rsPar = dbCmact.CargaRecordSet(sSQL)
    'rsPar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    Set rsPar.ActiveConnection = Nothing
    Set GetTarifaParametro = rsPar
    Set rsPar = Nothing
    dbCmact.CierraConexion

End Function
Public Sub nOpeCMACLlamadaCredPignoraticio(ByVal pnOpeCod As Long, _
        ByVal psFechaHora As String, ByVal psMovNro As String, _
        ByVal psGlosa As String, ByVal psPersCodCMAC As String, _
        ByVal psIFTipo As String, ByVal pnMoneda As Moneda, ByVal psCtaIf As String, _
        ByVal psDocumento As String, ByVal pnMontoTransac As Double, _
        Optional pbEjecBatch As Boolean = False, Optional nComision As Double = 0, _
        Optional nMonedaComision As Moneda = gMonedaNacional, Optional sCuentaAho As String = "", _
        Optional sNomAge As String = "", Optional sLpt As String = "", Optional sCliente As String = "", Optional psPersLavDinero As String = "", Optional ByRef psmensaje As String, _
        Optional ByRef psBoleta As String, Optional ByRef psBoletaITF As String, Optional ByVal pbImpTMU As Boolean = False, _
        Optional psTitPersLavDinero As String = "", Optional psOrdPersLavDinero As String = "", Optional psReaPersLavDinero As String = "", _
        Optional psBenPersLavDinero As String = "", Optional psVisPersLavDinero As String = "", _
        Optional pnMovNro As Long = 0)

        'By Capi 18022008 se agrego 5 parametros opcionales al final
        '** dInsertMov
    '** dInsertMovCMAC
    '************************************
    
    Dim lsSql As String
    'Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
    Dim lnMovNro As Long
    Dim lnMovNroRef As Long
    Dim lsOpeCod As String, sCodUser As String, sCodage As String
    'Dim oCont As New COMNContabilidad.NCOMContFunciones
    'Dim oCon As New DConecta
    Dim dFecSis As Date
    'Set dbCmact = New DConecta
    'dbCmact.AbreConexion
    
    dFecSis = CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4))
    sCodUser = Right(psMovNro, 4)
    sCodage = Mid(psMovNro, 18, 2)
    
    'Codigo de Operacion
    lsOpeCod = pnOpeCod
    
    'On Error GoTo ErrorModPig
    'loRegPig.dBeginTrans
    mbTrans = True

    '** Inserta Mov
    Call AgregaMov(psMovNro, lsOpeCod, psGlosa, gMovEstContabMovContable, gMovFlagVigente)

    ' Obtiene nMovNro
    lnMovNro = GetnMovNro(psMovNro)
    'ALPA 20081010****************************
    pnMovNro = lnMovNro
    '** Inserta MovCMAC
'    Call AgregaMovCMAC(lnMovNro, psPersCodCMAC, pnMoneda, pnOpeCod, _
'                         pnMontoTransac, psDocumento, psCtaIf, sCliente)

    If nComision > 0 Then
        Dim sMovNroCom As String
        Dim nMovNroCom As Long
        'Set oCont = New COMNContabilidad.NCOMContFunciones

        sMovNroCom = IncrementaMovNro(psMovNro)

        'xxx sMovNroCom = oCont.GeneraMovNro(dFecSis, sCodAge, sCodUser)

        Set oCont = Nothing
        'loRegPig.dBeginTrans
        Call AgregaMov(sMovNroCom, gOtrOpeComisionCMACLlam, psGlosa, gMovEstContabMovContable, gMovFlagVigente)
        nMovNroCom = GetnMovNro(sMovNroCom)
        Call AgregaMovOpeVarias(nMovNroCom, nComision, psCtaIf, psPersCodCMAC, nMonedaComision)
        Call AgregaMovRef(lnMovNro, nMovNroCom)
        'loRegPig.dCommitTrans
    End If

    If sCuentaAho <> "" Then
        'Dim oCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
        Dim sMovNroReg As String
        'Set oCont = New COMNContabilidad.NCOMContFunciones
        'sMovNroReg = oCont.GeneraMovNro(dFecSis, sCodAge, sCodUser)

        'xxx agregado
        If Len(Trim(sMovNroCom)) > 0 Then
            sMovNroReg = IncrementaMovNro(sMovNroCom)
        Else
            sMovNroReg = IncrementaMovNro(psMovNro)
        End If

        'Set oCont = Nothing

        'Set oCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento

        CapAbonoCuentaAho sCuentaAho, pnMontoTransac, gAhoDepRegCMACLlam, sMovNroReg, psGlosa, , , , , , , , , sNomAge, sLpt, psPersLavDinero, True, , , , , , , , , , , , , psmensaje, psBoleta, psBoletaITF, pbImpTMU

        'Set oCapMov = Nothing

        'Pepe
        lnMovNroRef = GetnMovNro(sMovNroReg)
        Call AgregaMovRef(lnMovNro, lnMovNroRef)
        'Fin Pepe

    End If

    'loRegPig.dCommitTrans

    Set loRegPig = Nothing

    'Set oBase = Nothing
    'oCon.CierraConexion
Exit Sub
ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nCancelacionCredPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub
Public Sub GetOperaciones(rsOpeLLama As ADODB.Recordset, ByRef rsCMACs As ADODB.Recordset)
    Set dbCmact = New DConecta
    dbCmact.AbreConexion
  
    sSQL = "SELECT O.cOpeCod, O.cOpeDesc, O.cOpeVisible, O.nOpeNiv FROM OpeTpo O  " _
         & "Where O.cOpeCod in ('107000','107001','107004','260500','260501','260503','260504','260505','260506') And cOpeVisible = 1 Order by O.cOpeCod, O.nOpeNiv "
    
    Set rsOpeLLama = New ADODB.Recordset
    Set rsOpeLLama = dbCmact.CargaRecordSet(sSQL)
    
'    sSQL = "Select I.cPersCod, P.cPersNombre, Nivel = 1, C.cConsDescripcion + space(50) + I.cIFTpo as sTipo, C.cConsDescripcion, I.cSubCtaContCod " _
'        & "From InstitucionFinanc I JOIN Persona P on I.cPersCod = P.cPersCod " _
'        & "Inner JOIN Constante C ON convert(Int,I.cIFTpo) = C.nConsValor " _
'        & "WHERE C.nConsCod = '" & COMDConstantes.gCGTipoIF & "' And nConsValor <> '" & COMDConstantes.gCGTipoIF & "' " _
'        & "AND Convert(smallint,I.cIFTpo) = 3 Order By P.cPersNombre "

    sSQL = "Select CO.cPerscod, CO.cPersDesc as cPersNombre, Nivel = 1, CO.cCodAutOIT" _
         & " from CMACsOperInterCajas CO"
         
    Set rsCMACs = New ADODB.Recordset
    Set rsCMACs = dbCmact.CargaRecordSet(sSQL)
    
    dbCmact.CierraConexion
    Set dbCmact = Nothing

End Sub
Public Function GetVarSistema() As ADODB.Recordset
    Dim rsPar As ADODB.Recordset
    Set dbCmact = New DConecta
    dbCmact.AbreConexion
    
    Set rsPar = New ADODB.Recordset
    sSQL = " Select nConsSisCod,nConsSisValor,nConsSisDesc From ConstSistema " _
            & " Where nConsSisCod in (" & gConstSistFechaSistema & "," & gConstSistNombreAbrevCMAC & "," _
            & gConstSistRutaBackup & "," & gConstSistCodCMAC & "," & gConstSistMargenSupCartas & "," & gConstSistMagenIzqCartas & "," _
            & gConstSistMargenDerCartas & "," & gConstSistNroLineasPagina & "," & gConstSistNroLineasOrdenPago & "," _
            & gConstSistCtaConversionMEDol & "," & gConstSistCtaConversiónMESoles & "," & gConstSistTipoConverión & "," _
            & gConstSistNombreModulo & "," & gConstSistFechaInicioDia & "," & gConstSistDominio & "," & gConstSistPDC _
            & ",40,50," & gConstPersCodCMACT & "," & gConstSistAgenciaEspecial & "," & gConstSistCierreMesNegocio & ", " _
            & gConstSistRutaIcono & ",84,85," & gConstSistVerificaRegistroEfectivo & ",301,302,303,304,305) " _
            & "ORDER BY nConsSisCod"

    Set rsPar = dbCmact.CargaRecordSet(sSQL)
    
    Set rsPar.ActiveConnection = Nothing
    Set GetVarSistema = rsPar
    Set rsPar = Nothing
    dbCmact.CierraConexion
End Function
Public Function DevuelveHoraServer() As String
    Dim rsH As New ADODB.Recordset
    Set dbCmact = New DConecta

    If dbCmact.AbreConexion = False Then Exit Function
    sSQL = "Select Convert(varchar(10),getdate(),108) as sHora"
    Set rsH = dbCmact.CargaRecordSet(sSQL)
    If Not rsH.EOF Then
        GetHoraServer = rsH!sHora
    Else
        GetHoraServer = Format(Time, "hh:mm:ss")
    End If
    rsH.Close

    dbCmact.CierraConexion
    Set dbCmact = Nothing
End Function
Public Function GetMensajeBoletas(ByVal sCuenta As String, Optional sAgencia As String) As String
Dim sSQL As String
Dim rsMsg As New ADODB.Recordset
Set dbCmact = New DConecta

If dbCmact.AbreConexion = False Then Exit Function
rsMsg.CursorLocation = adUseClient
sSQL = "Select cMensaje from Mensajes where cCodPro = '" & Mid(sCuenta, 6, 3) & "' And bMenEst = 1"
If sAgencia <> "" Then
    sSQL = sSQL & " And cCodAge = '" & sAgencia & "'"
End If

'rsMsg.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsMsg = dbCmact.CargaRecordSet(sSQL)

Set rsMsg.ActiveConnection = Nothing
If rsMsg.EOF And rsMsg.BOF Then
    GetMensajeBoletas = ""
Else
    GetMensajeBoletas = Trim(rsMsg("cMensaje"))
End If
rsMsg.Close
Set rsMsg = Nothing
End Function
Public Function DevuelveCodAutorizaCMAC(ByVal psPersCod As String, ByRef psPersDesc As String) As String
    Dim rsA As New ADODB.Recordset
    Set dbCmact = New DConecta

    If dbCmact.AbreConexion = False Then Exit Function
    sSQL = "Select cPersDesc, cCodAutOIT From CMACsOperInterCajas Where cPerscod = '" & psPersCod & "'"
    Set rsA = dbCmact.CargaRecordSet(sSQL)
    If Not rsA.EOF Then
        DevuelveCodAutorizaCMAC = rsA!cCodAutOIT
        psPersDesc = rsA!cPersDesc
    Else
        DevuelveCodAutorizaCMAC = ""
        psPersDesc = ""
    End If
    rsA.Close

    dbCmact.CierraConexion
    Set dbCmact = Nothing
End Function
Public Sub nOpeCMACConsultaCTAAhoCreMov(ByVal psMovNro As String, ByRef pnMovNro As Long, ByVal pnMoneda As Integer, ByVal pnComision As Double, ByVal psOpeCod As String, _
                                 ByRef pnMovNroCom As Long, ByVal psOpeDescripcion As String, ByVal psNomAge As String, ByVal sLpt As String, ByVal lsBoleta As String, _
                                 Optional ByVal psCuenta As String, Optional ByVal psPersCodCMAC As String, Optional ByVal psNombreCMAC As String, _
                                 Optional psCliente As String, Optional ByVal psGlosa As String)
                                 
Dim sMsgOpe As String, sCodUser As String, sCodage As String
Dim dFecSis As Date
Dim bTrans As Boolean
Dim nMovNro As Long
Dim nTempo As Currency

Dim nBan As Boolean

    
    dFecSis = CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4))
    sCodUser = Right(sMovNro, 4)
    sCodage = Mid(sMovNro, 18, 2)
    bTrans = False
    

    On Error GoTo ErrGraba
    'clsCap.dbCmact.BeginTrans
    'clsCap.bTransaccion = True
    
    bTrans = True
    If sGlosa = "" Then
        sGlosa = sDescOperacion & ", " & sNomCmac & ", Cuenta = " & sCuenta
    Else
        sGlosa = sGlosa & Chr$(13) & sDescOperacion & ", " & sNomCmac & ", Cuenta = " & sCuenta
    End If
    
    'clsCap.AgregaMov sMovNro, nOperacion, sGlosa
    Call AgregaMov(psMovNro, Val(psOpeCod), sGlosa)
    nMovNro = GetnMovNro(sMovNro)
    pnMovNro = nMovNro
    'bTrans = False

    If pnComision > 0 Then
        Dim sMovNroCom As String
        Dim nMovNroCom As Long
        'Set oCont = New COMNContabilidad.NCOMContFunciones
        
        'xxx Agregado
        'sMovNroCom = oCont.IncrementaMovNro(sMovNro)
        sMovNroCom = IncrementaMovNro(sMovNro)
            
        'xxx eliminado sMovNroCom = oCont.GeneraMovNro(dFecSis, sCodAge, sCodUser)
        Set oCont = Nothing
        
        'Inicio Transaccion
        'xxx eliminado clsCap.dbCmact.BeginTrans
        
        AgregaMov sMovNroCom, gOtrOpeComisionCMACLlam, sGlosa
        nMovNroCom = GetnMovNro(sMovNroCom)
        AgregaMovOpeVarias nMovNroCom, nComision, sCuenta, sCodCmac, nMonedaComision
        AgregaMovRef nMovNro, nMovNroCom
        pnMovNroCom = nMovNroCom
'        If pnITFValorCom > 0 Then
'            sMovNroComITF = Left(nMovNroCom, 2) & "2" & Right(nMovNroCom, 4)
'            AgregaMov sMovNroComITF, gITFCobroEfectivo, sGlosa
'            nMovNroComITF = clsCap.GetnMovNro(sMovNroComITF)
'            AgregaMovOpeVarias nMovNroComITF, pnITFValorCom, sCuenta, sCodCmac, nMonedaComision
'            AgregaMovRef nMovNroCom, nMovNroComITF
'        End If
    
    End If

'    If sCuentaAho <> "" And nOperacion <> gCMACOTAhoDepChq Then
'        Dim sMovNroReg As String
'        Dim nMovNroReg As Long
'        'Set oCont = New COMNContabilidad.NCOMContFunciones
'
'        'xxx agregado
'        If Len(Trim(sMovNroCom)) > 0 Then
'            sMovNroReg = IncrementaMovNro(sMovNroCom)
'        Else
'            sMovNroReg = IncrementaMovNro(sMovNro)
'        End If
'
'        'Set oCont = Nothing
'        Select Case nOperacion
'            Case gCMACOTAhoDepEfec
'                'ALPA 20081009***************************************************
'                nTempo = CapAbonoCuentaAho(sCuentaAho, nMonto, gAhoDepRegCMACLlam, sMovNroReg, sGlosa, , , , , , , , , sNomAge, sLpt, psPersLavDinero, True, , , , , , , , , , , , psPersBenLavDinero, , psImpBoleta, , False, psTitPersLavDinero, psOrdPersLavDinero, psReaPersLavDinero, psBenPersLavDinero, psVisPersLavDinero)
'                If nTempo = 0 Then
'                  ' If clsCap.bTransaccion = False Then clsCap.dbCmact.RollbackTrans
'                    'clsCap.dbCmact.RollbackTrans
'                    'Set clsCap = Nothing
'                    Exit Sub
'                End If
'            Case gCMACOTAhoRetEfec, gCMACOTAhoRetOP
'
'                nTempo = CapCargoCuentaAho(sCuentaAho, nMonto, 200330, sMovNroReg, "", "", "", "", , , , , , , sLpt, , , , , , , , , , , , , , , , psImpBoleta)
'                If nTempo = 0 Then
'                    'If clsCap.bTransaccion = False Then clsCap.dbCmact.RollbackTrans
'                    'clsCap.dbCmact.RollbackTrans Comentado Por MAVM Porq sale error
'                    Set clsCap = Nothing
'                    Exit Sub
'                End If
'        End Select
'        nMovNroReg = GetnMovNro(sMovNroReg)
'        AgregaMovRef nMovNro, nMovNroReg
'    End If
    
    'clsCap.dbCmact.CommitTrans
    'clsCap.bTransaccion = False
    
    'Imprime Documentos
    'MsgBox "Se imprimirán las boletas de regularización", vbInformation, "Aviso"
    

    'Impresion de Boleta de Regulariazacion
    'psImpBoleta = psImpBoleta & ImprimeBoletaCMACRegula("OPERACION CMAC LLAMADA ", sDescOperacion, Trim(nMonto), sCliente, sCuenta, "", 0, nMoneda, "", 0, 0, True, True, , , sNomCmac, , dFecSis, sNomAge, sCodUser, sLpt, True, sCodAge, bSaldo, psCuentaPers) & oImp.gPrnSaltoLinea
   
    
'Set clsCap = Nothing

Exit Sub
ErrGraba:
    'If bTrans Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
'    MsgBox Err.Description, vbExclamation, "Error"
    Err.Raise Err.Number, "Error", Err.Description
End Sub
