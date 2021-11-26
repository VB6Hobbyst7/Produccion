VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCaptaDefinicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public dbCmact As Connection
Dim dbCmactHist As Connection
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String
Dim sSql As String

Public Function ActualizaParametros(ByVal nVariable As Long, ByVal sDesc As String, ByVal nValor As Double) As Boolean
    Dim rsPar As New ADODB.Recordset
    Dim sDescH As String, sVar As String
    Dim nVal As Double
    On Error GoTo ErrGetParam
    sSql = "Update Parametro Set cParDescripcion = '" & sDesc & "', nParValor = " & nValor & " WHERE " _
        & "nParCod = " & nVariable & " And nParProd = " & gPrdParamCaptac
    dbCmact.Execute sSql
    ActualizaParametros = True
    Exit Function
ErrGetParam:
    Err.Raise Err.Number, "NCapDefinicion:ActualizaParametros", Err.Description
    ActualizaParametros = False
End Function

'JUEZ 20141115 Se agregó parametro pnCategoria
Public Function GetParametros(Optional ByVal pnCategoria As Integer = 0) As ADODB.Recordset
    Dim rsParam As ADODB.Recordset

    On Error GoTo ErrGetParam
    Set rsParam = New ADODB.Recordset
    rsParam.CursorLocation = adUseClient
    'JUEZ 20141115 Convertido a SP ************************
    'sSql = "Select nParCod Variable, cParDescripcion Descripcion, nParValor Valor, Tag = '' " _
    '    & "FROM Parametro WHERE nParProd = " & gPrdParamCaptac
    sSql = "stp_sel_CapParametros " & pnCategoria
    'END JUEZ *********************************************
    rsParam.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rsParam.ActiveConnection = Nothing
    Set GetParametros = rsParam
    Set rsParam = Nothing
    Exit Function
ErrGetParam:
    Err.Raise Err.Number, "NCapDefinicion:GetParametros", Err.Description
End Function
Public Function GetParametrosPIT() As ADODB.Recordset
    Dim rsParamPIT As ADODB.Recordset

    On Error GoTo ErrGetParam
    Set rsParamPIT = New ADODB.Recordset
    rsParamPIT.CursorLocation = adUseClient
    sSql = "exec stp_sel_RecuperaParamPIT "
    rsParamPIT.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rsParamPIT.ActiveConnection = Nothing
    Set GetParametrosPIT = rsParamPIT
    Set rsParamPIT = Nothing
    Exit Function
ErrGetParam:
    Err.Raise Err.Number, "NCapDefinicion:GetParametros", Err.Description
End Function
Public Function GetTarifario(ByVal sProd As Producto, ByVal sMoneda As Moneda, _
        ByVal nTipo As CaptacTipoTasa, ByVal sAgencia As String, _
        Optional ByVal nTpoPrograma As Integer = 0, Optional ByVal nTipoTarifario As Integer = 0) As ADODB.Recordset
        'JIPR20190109 ERS077-2018 , Optional ByVal nTipoTarifario As Integer = 0
    Dim rsParam As ADODB.Recordset

    Set rsParam = New ADODB.Recordset
    rsParam.CursorLocation = adUseClient
    
    'By Capi 15012008 se adiciono nuevos campos
    '& "nTasaValor, nTasaCod From CaptacTasas Where nTasaProd = " & sProd & " And nTasaMon = " & sMoneda & " And "
    'JUEZ 2013020 *********************************************************************
        'sSql = "Select nValorIni, nValorFin, nPlazoIni, nPlazoFin, (Case When cOrdPag='1' Then 'Si' Else 'No' End), " _
        '& "nTasaValor, nTasaCod,IsNull(cLogFechaCambio,''),cActiva From CaptacTasas Where nTasaProd = " & sProd & " And nTasaMon = " & sMoneda & " And " _
        '& "nTasaTpo = " & nTipo & " And cCodAge = '" & sAgencia & "' And nTpoPrograma = " & nTpoPrograma & " " _
        '& "Order by cActiva,nPlazoIni,nPlazoFin,nTasaCod,cLogFechaCambio"
     sSql = " exec stp_sel_GetTarifario '" & sProd & "','" & sMoneda & "','" & sAgencia & "','" & nTpoPrograma & "','" & nTipoTarifario & "'" 'JIPR20190109 ERS077-2018 ,'" & nTipoTarifario & "'"
    'END JUEZ *************************************************************************
    rsParam.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsParam.ActiveConnection = Nothing

    Set GetTarifario = rsParam
    Set rsParam = Nothing
    Exit Function
ErrGetParam:
    Err.Raise Err.Number, "NCapDefinicion", Err.Description
End Function

Public Sub ActualizaTasa(ByVal nCodigo As Integer, ByVal nMonIni As Double, _
        ByVal nMonFin As Double, ByVal nPlazoIni As Long, ByVal nPlazoFin As Long, _
        ByVal bOrdPag As Boolean, ByVal nValor As Double, Optional ByVal pcCambio As String = "", Optional ByVal bActiva As Boolean = True, _
        Optional ByVal nProd As Integer = 0, Optional ByVal nTpoPrograma As Integer = 0, Optional ByVal nMoneda As Integer = 0, Optional ByVal sAgencia As String = "", _
        Optional ByVal nTipoTarifario As Integer = 0)
        'JIPR20190109 ERS077-2018 , Optional ByVal nTipoTarifario As Integer = 0
        
        'By Capi 15012008 se agrego parametros cCambio y bActiva
        'JUEZ 20140220 nProd, nTpoPrograma, sAgencia
    Dim sActiva As String
    Dim sOrdPag As String
    On Error GoTo ErrGetParam

    sOrdPag = IIf(bOrdPag, "1", "0")
    'By Capi 15012008
    sActiva = IIf(bActiva, "1", "0")
    
    
    
'       ssql = "Update CaptacTasas Set nValorIni = " & nMonIni & ", nValorFin = " & nMonFin & ", nTasaValor = " & nValor & ", " _
'            & "nPlazoIni = " & nPlazoIni & ", nPlazoFin = " & nPlazoFin & ", " _
'            & "cOrdPag = '" & sOrdPag & "',cActiva = '" & sActiva & "'" & " Where nTasaCod = " & nCodigo
'
        'By Capi 20012008
        'JUEZ 20140220 *****************************************
        'sSql = "Update CaptacTasas Set cActiva = '" & sActiva & "'" & " Where nTasaCod = " & nCodigo
        sSql = "EXEC stp_sel_ActualizaTasa " & nProd & "," & nTpoPrograma & "," & nMoneda & ",'" & sAgencia & "','" & nTipoTarifario & "'" 'JIPR20190109 ERS077-2018  ,'" & nTipoTarifario & "'"
        
        'END JUEZ **********************************************

    dbCmact.Execute sSql
    Exit Sub
ErrGetParam:
    Err.Raise Err.Number, "NCapDefinicion:ActualizaTasa", Err.Description
End Sub

Public Sub NuevaTasa(ByVal nProd As Producto, ByVal nMon As Moneda, _
        ByVal nTipoTasa As CaptacTipoTasa, ByVal nMonIni As Double, nMonFin As Double, _
        ByVal nPlazoIni As Long, ByVal nPlazoFin As Long, ByVal bOrdPag As Boolean, _
        ByVal sAgencia As String, ByVal nValor As Double, Optional ByVal nTpoPrograma As Integer = 0, _
        Optional ByVal pcCambio As String = "", Optional ByVal bActiva As Boolean = True, _
        Optional ByVal pnTasaCodAnt As Double, Optional ByVal nTipoTarifario As Integer = 0) 'JIPR20190109 ERS077-2018 , Optional ByVal nTipoTarifario As Integer = 0
        'By capi 15012008 se agrego 2 parametros opcionales

    Dim sOrdPag As String
    Dim sActiva As String
    On Error GoTo ErrTasa
    sOrdPag = IIf(bOrdPag, "1", "0")
    sActiva = IIf(bActiva, "1", "0")
    'JUEZ 20140220 ***************************************************************
    'sSql = "Insert CaptacTasas (nTasaProd,nTasaMon,nTasaTpo,nValorIni,nValorFin,nPlazoIni,nPlazoFin,cOrdPag,cCodAge,nTasaValor,nTpoPrograma,clogFechaCambio,cActiva,nTasaCodAnt) " _
    '        & "Values (" & nProd & "," & nMon & "," & nTipoTasa & "," & nMonIni & "," & nMonFin & "," & nPlazoIni & "," & nPlazoFin & ",'" & sOrdPag & "','" & sAgencia & "'," & nValor & "," & nTpoPrograma & ",'" & pcCambio & "','" & sActiva & "'," & pnTasaCodAnt & ")"
    sSql = "EXEC stp_sel_NuevaTasa " & nProd & "," & nMon & "," & nTipoTasa & "," & nMonIni & "," & nMonFin & "," & nPlazoIni & "," & nPlazoFin & "," & sOrdPag & ",'" & sAgencia & "'," & nValor & "," & nTpoPrograma & ",'" & pcCambio & "','" & nTipoTarifario & "'" 'JIPR20190109 ERS077-2018  ,'" & nTipoTarifario & "'"
    'END JUEZ ********************************************************************
    
    dbCmact.Execute sSql
    Exit Sub
ErrTasa:
    Err.Raise Err.Number, "NCapDefinicion:NuevaTasa", Err.Description
End Sub


Public Sub ActualizaTasaCaptacion(ByVal nProd As Producto, ByVal nMon As Moneda, _
                    ByVal sTipoAnt As String, ByVal sTipoNuevo As String, _
                    ByVal nValorAnt As Double, nValorNuevo As Double, _
                    ByVal dFecha As Date)

    Dim cmdTasa As Command
    Dim prmTasa As Parameter

    Set cmdTasa = New Command
    cmdTasa.CommandText = "spCapActualizaTasas"
    cmdTasa.Name = "spActualizaTasa"
    cmdTasa.CommandType = adCmdStoredProc

    Set prmTasa = cmdTasa.CreateParameter("Producto", adVarChar, adParamInput, 3, Trim(Str(nProd)))
    cmdTasa.Parameters.Append prmTasa
    Set prmTasa = cmdTasa.CreateParameter("Moneda", adVarChar, adParamInput, 1, Trim(Str(nMon)))
    cmdTasa.Parameters.Append prmTasa
    Set prmTasa = cmdTasa.CreateParameter("TipoAnt", adVarChar, adParamInput, 3, sTipoAnt)
    cmdTasa.Parameters.Append prmTasa
    Set prmTasa = cmdTasa.CreateParameter("TipoNuevo", adVarChar, adParamInput, 3, sTipoNuevo)
    cmdTasa.Parameters.Append prmTasa
    Set prmTasa = cmdTasa.CreateParameter("ValorAnt", adCurrency, adParamInput, nValorAnt)
    cmdTasa.Parameters.Append prmTasa
    Set prmTasa = cmdTasa.CreateParameter("ValorNuevo", adCurrency, adParamInput, nValorNuevo)
    cmdTasa.Parameters.Append prmTasa
    Set prmTasa = cmdTasa.CreateParameter("Fecha", adDate, adParamInput, dFecha)
    cmdTasa.Parameters.Append prmTasa
    cmdTasa.Parameters.Refresh
    cmdTasa.ActiveConnection = dbCmact
    dbCmact.spActualizaTasa

    Set prmTasa = Nothing
    Set cmdTasa = Nothing
End Sub
Public Function ValidaPlazoMin_SubProductoAHO(ByVal pnPlazo As Integer, ByVal pnPersoneria As Integer, _
ByVal pnMoneda As Moneda, ByVal pnTpoPrograma As Integer, ByRef nPlazoMin As Integer, Optional bOrdPag As Boolean = False) As String

    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sOrdPag As String
    sOrdPag = IIf(bOrdPag, "1", "0")

    sql = "Select nPlazoMin from CapPersParam where nProducto = 232 and cOrdPag = '" & sOrdPag & "' and "
    sql = sql & " nMoneda = " & pnMoneda & " and nTpoPrograma = " & pnTpoPrograma & " and nPersoneria=" & pnPersoneria
    rs.CursorLocation = adUseClient
    rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    nPlazoMin = 0
    If Not (rs.EOF And rs.BOF) Then
        If IsNull(rs!nPlazoMin) Then
            ValidaPlazoMin_SubProductoAHO = "No se a Definido el Plazo Minimo para este Tipo de Personeria"
        Else
            If rs!nPlazoMin <= pnPlazo Then
                ValidaPlazoMin_SubProductoAHO = ""
            Else
                ValidaPlazoMin_SubProductoAHO = "Plazo Incorrecto el Mìmimo es: " & rs!nPlazoMin
                nPlazoMin = rs!nPlazoMin
            End If
        End If
    Else
        ValidaPlazoMin_SubProductoAHO = "No se a Definido el Plazo Minimo para este Tipo de Personeria"
    End If
    Set rs = Nothing
End Function
Public Function ValidaMontoMinDep_SubProductoAHO(ByVal pnMontoMinDep As Currency, ByVal pnPersoneria As Integer, _
ByVal pnMoneda As Moneda, ByVal pnTpoPrograma As Integer, ByRef nMontoMinDep As Currency, Optional bOrdPag As Boolean = False) As String

    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sOrdPag As String
    sOrdPag = IIf(bOrdPag, "1", "0")

    sql = "Select nMontoMinDep from CapPersParam where nProducto = 232 and cOrdPag = '" & sOrdPag & "' and "
    sql = sql & " nMoneda = " & pnMoneda & " and nTpoPrograma = " & pnTpoPrograma & " and nPersoneria=" & pnPersoneria
    rs.CursorLocation = adUseClient
    rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    nMontoMinDep = 0
    If Not (rs.EOF And rs.BOF) Then
        If IsNull(rs!nMontoMinDep) Then
            ValidaMontoMinDep_SubProductoAHO = "No se a Definido el Monto Minimo de Deposito para este Tipo de Personeria"
        Else
            If rs!nMontoMinDep <= pnMontoMinDep Then
                ValidaMontoMinDep_SubProductoAHO = ""
            Else
                nMontoMinDep = rs!nMontoMinDep
                ValidaMontoMinDep_SubProductoAHO = "Monto Minimo de Deposito Incorrecto el Mìmimo es: " & rs!nMontoMinDep
            End If
        End If
    Else
        ValidaMontoMinDep_SubProductoAHO = "No se a Definido el Monto Minimo de Deposito para este Tipo de Personeria"
    End If
    Set rs = Nothing
End Function

Public Function GetITF_AsumidoAgencia(ByVal psCodAge As String) As Boolean
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    sql = "select bITFApert from Agencias  where cAgeCod  = '" & psCodAge & "'"
    rs.CursorLocation = adUseClient
    rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    GetITF_AsumidoAgencia = False

    If Not (rs.EOF Or rs.BOF) Then
        If rs("bITFApert") = "1" Then
            GetITF_AsumidoAgencia = True
        End If
    End If
    rs.Close
    Set rs = Nothing

End Function
Public Function GetCapTasaInteres(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
    ByVal nTipoTasa As CaptacTipoTasa, Optional nPlazo As Long = 0, Optional nValor As Double = 0, _
    Optional sCodAge As String = "", Optional bOrdPag As Boolean = False, Optional nSubProd As Integer = 0) As Double

    Dim rsTasa As ADODB.Recordset
    Dim cOrdPag As String

    'OJO 03-11-2005
    'If sCodage =  "" Then sCodage = gsCodAge
'COMENTO APRI
'    If nProducto = gCapAhorros Or nProducto = gCapCTS Then
'        cOrdPag = IIf(bOrdPag, "1", "0")
'        sSql = "SELECT nTasaValor FROM CaptacTasas WHERE nTasaProd = " & nProducto & " AND " _
'            & "nTasaMon = " & nMoneda & " AND nTasaTpo = " & nTipoTasa & " And (" & nValor & " >= nValorIni " _
'            & "And " & nValor & " < nValorFin) And cOrdPag = '" & cOrdPag & "' And " _
'            & "nTpoPrograma = " & nSubProd & " And cCodAge = '" & sCodAge & "' And cActiva='1'"
'            'By Capi 20012008 se agrego para que solo tome tasas activas
'    Else
'        sSql = "SELECT nTasaValor FROM CaptacTasas WHERE nTasaProd = " & nProducto & " AND " _
'            & "nTasaMon = " & nMoneda & " AND nTasaTpo = " & nTipoTasa & " And (" & nPlazo & " " _
'            & ">= nPlazoIni AND " & nPlazo & " <= nPlazoFin) And (" & nValor & " >= nValorIni " _
'            & "And " & nValor & " < = nValorFin) And nTpoPrograma = " & nSubProd & " " _
'            & "And cCodAge = '" & sCodAge & "' And cActiva='1'"
'            'By Capi 20012008 se agrego para que solo tome tasas activas
'    End If
    sSql = "EXEC stp_sel_GetCapTasaInteres " & nProducto & "," & nSubProd & "," & nMoneda & "," & nTipoTasa & "," _
    & IIf(bOrdPag, 1, 0) & ",'" & sCodAge & "'," & nPlazo & "," & nValor 'APRI20190109 ERS077-2018
    
    Set rsTasa = New ADODB.Recordset
    rsTasa.CursorLocation = adUseClient
    rsTasa.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsTasa.ActiveConnection = Nothing
    If Not (rsTasa.EOF Or rsTasa.BOF) Then
        GetCapTasaInteres = rsTasa("nTasaValor")
    Else
        GetCapTasaInteres = 0
    End If
    rsTasa.Close
    Set rsTasa = Nothing
End Function

Public Function GetCapParametro(ByVal nParametro As CaptacParametro) As Double
    Dim rsVar As ADODB.Recordset
    sSql = "SELECT nParValor FROM Parametro WHERE nParCod = " & nParametro & " " _
        & "And nParProd = " & gPrdParamCaptac
    Set rsVar = New ADODB.Recordset
    rsVar.CursorLocation = adUseClient
    rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsVar.ActiveConnection = Nothing
    If rsVar.EOF And rsVar.BOF Then
        GetCapParametro = 0
    Else
        GetCapParametro = rsVar("nParValor")
    End If
    rsVar.Close
    Set rsVar = Nothing
End Function

Public Function GetCapParametroDesc(ByVal nParametro As CaptacParametro) As String
    Dim rsVar As ADODB.Recordset
    sSql = "SELECT cParDescripcion FROM Parametro WHERE nParCod = " & nParametro
    Set rsVar = New ADODB.Recordset
    rsVar.CursorLocation = adUseClient
    rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsVar.ActiveConnection = Nothing
    If rsVar.EOF And rsVar.BOF Then
        GetCapParametroDesc = ""
    Else
        GetCapParametroDesc = rsVar("cParDescripcion")
    End If
    rsVar.Close
    Set rsVar = Nothing
End Function

Public Function GetCapMovTipo() As ADODB.Recordset
    Dim rsVar As ADODB.Recordset
    sSql = "Select O.cOpeCod, O.cOpeDesc, Prod = CASE WHEN M.Prod IS NULL THEN " _
        & "CASE WHEN O.cOpeCod LIKE '2[03]%' Or O.cOpeCod LIKE '2601%' Or O.cOpeCod LIKE '1001%' THEN 'AHORROS' + Space(75) + '" & gCapAhorros & "' " _
        & "     WHEN O.cOpeCod LIKE '2[14]%' Or O.cOpeCod LIKE '2602%' THEN 'PLAZO FIJO' + Space(75) + '" & gCapPlazoFijo & "' " _
        & "     WHEN O.cOpeCod LIKE '2[25]%' Or O.cOpeCod LIKE '2603%' THEN 'CTS' + Space(75) + '" & gCapCTS & "' END " _
        & "ELSE M.Prod END, M.Tipo, Estad = CONVERT(Int,ISNULL(M.Estad,0)) From " _
        & "(Select M.cOpeCod, UPPER(K1.cConsDescripcion) + Space(75) + Convert(Varchar(3),nProducto) Prod, " _
        & "UPPER(K.cConsDescripcion) + Space(75) + Convert(Varchar(2),nCapMovTpo) Tipo, bEstadistica Estad From CapMovTipo M INNER JOIN Constante K " _
        & "ON M.nCapMovTpo = K.nConsValor INNER JOIN Constante K1 ON M.nProducto = K1.nConsValor Where K.nConsCod = " & gCaptacMovTipo & "  " _
        & "And K1.nConsCod = " & gProducto & ") M " _
        & "RIGHT JOIN OpeTpo O ON M.cOpeCod = O.cOpeCod Where (O.cOpeCod LIKE '2[0-7]%' OR O.cOpeCod LIKE '1001%') " _
        & "Order by O.cOpeCod"
    Set rsVar = New ADODB.Recordset
    rsVar.CursorLocation = adUseClient
    rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsVar.ActiveConnection = Nothing
    Set GetCapMovTipo = rsVar
    Set rsVar = Nothing
End Function

Private Sub InsertaCapMovTipo(ByVal sOpeCod As String, ByVal nProducto As Producto, _
        ByVal nCapMovTpo As CaptacMovTipo, ByVal bEstadistica As Boolean)
    sSql = "Insert CapMovTipo (cOpeCod,nProducto,nCapMovTpo,bEstadistica) " _
        & "Values ('" & sOpeCod & "'," & nProducto & "," & nCapMovTpo & ", " & IIf(bEstadistica, 1, 0) & ")"
    dbCmact.Execute sSql
End Sub

Private Sub EilminaCapMovTipo(ByVal sOpeCod As String)
    sSql = "Delete CapMovTipo Where cOpeCod = '" & sOpeCod & "'"
    dbCmact.Execute sSql
End Sub

Public Sub ActualizaCaptacionOperacion(ByVal rsOpe As Recordset)
    Dim sMod As String, sOpeCod As String
    Dim nProducto As Producto
    Dim nCapMovTipo As CaptacMovTipo
    Dim bEstad As Boolean
    dbCmact.BeginTrans
    rsOpe.MoveFirst
    Do While Not rsOpe.EOF
        sMod = rsOpe("Mod")
        If sMod = "M" Then
            sOpeCod = rsOpe("Código")
            nProducto = CLng(Trim(Right(rsOpe("Producto"), 4)))
            nCapMovTipo = CLng(Trim(Right(rsOpe("Tipo"), 4)))
            bEstad = CLng(Trim(Right(rsOpe("Estad"), 4)))
            EilminaCapMovTipo sOpeCod
            InsertaCapMovTipo sOpeCod, nProducto, nCapMovTipo, bEstad
        End If
        rsOpe.MoveNext
    Loop
    dbCmact.CommitTrans
End Sub

Public Function GetnMovNro(ByVal sMovNro As String) As Long
    Dim rsMov As ADODB.Recordset
    sSql = "Select nMovNro From Mov Where cMovNro = '" & sMovNro & "'"
    Set rsMov = New ADODB.Recordset
    rsMov.CursorLocation = adUseClient
    rsMov.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsMov.ActiveConnection = Nothing
    If rsMov.EOF And rsMov.BOF Then
        GetnMovNro = 0
    Else
        GetnMovNro = rsMov("nMovNro")
    End If
    rsMov.Close
    Set rsMov = Nothing
End Function

Private Sub ActualizaPFBusqCredPend(ByVal sCuenta As String, ByVal bValor As Boolean)
    sSql = "Update CaptacPlazoFijo Set bBusCredPend = " & IIf(bValor, 1, 0) & " " _
        & "Where cCtaCod = '" & sCuenta & "'"
    dbCmact.Execute sSql
End Sub

Public Sub AgregaMov(ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
        ByVal sGlosa As String, Optional nMovEstado As MovEstado = gMovEstContabMovContable, _
        Optional nMovFlag As MovFlag = gMovFlagVigente)
    sSql = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
        & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sGlosa & "'," & nMovEstado & "," & nMovFlag & ")"
    dbCmact.Execute sSql
End Sub

Public Sub AgregaMovCap(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double)
    sSql = "INSERT MovCap (nMovNro,cOpeCod,cCtaCod,nMonto,nSaldoDisponible,nSaldoContable) " _
        & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nMonto & "," & nSaldoDisp & "," & nSaldoCnt & ")"
    dbCmact.Execute sSql
End Sub

Public Sub ActualizaCredPendPagoPF(ByVal sCuenta As String, ByVal bValor As Boolean, _
            ByVal sMovNro As String, ByVal sGlosa As String)
    Dim bTrans As Boolean
    Dim nMovNro As Long
    On Error GoTo ErrGraba
    dbCmact.BeginTrans
    bTrans = True
    AgregaMov sMovNro, gPFBusqCredOend, sGlosa, gMovEstContabNoContable, gMovFlagVigente
    nMovNro = GetnMovNro(sMovNro)
    AgregaMovCap nMovNro, gPFBusqCredOend, sCuenta, 0, 0, 0
    ActualizaPFBusqCredPend sCuenta, bValor
    dbCmact.CommitTrans
    Exit Sub
ErrGraba:
    If bTrans Then dbCmact.RollbackTrans
    Err.Raise Err.Number, "NCapDefinicion:ActualizaCredPendPagoPF", Err.Description
End Sub

Public Sub ActualizaDocRecParam(ByVal nTipoCheque As ChequeTipo, ByVal nPlaza As ChequePlaza, _
        ByVal nMoneda As Moneda, ByVal nDias As Double)
    sSql = "Update DocRecParametro Set nDiasVal = " & nDias & " " _
        & "Where nPlaza = " & nPlaza & " And nTpoCheque = " & nTipoCheque & " And nMoneda = " & nMoneda
    dbCmact.Execute sSql
End Sub

Public Sub ActualizaDocRecParametro(ByVal rsParam As Recordset)
    Dim nTipoCheque As ChequeTipo
    Dim nPlaza As ChequePlaza
    Dim nMoneda As Moneda
    Dim nDias As Double
    Dim bTrans As Boolean

    On Error GoTo ErrGraba
    dbCmact.BeginTrans
    bTrans = True
    Do While Not rsParam.EOF
        If rsParam("Mod") = "M" Then
            nTipoCheque = CLng(Right(rsParam("Tipo"), 1))
            nPlaza = CLng(Right(rsParam("Plaza"), 1))
            nMoneda = CLng(Right(rsParam("Moneda"), 1))
            nDias = CDbl(rsParam(3))
            ActualizaDocRecParam nTipoCheque, nPlaza, nMoneda, nDias
        End If
        rsParam.MoveNext
    Loop
    dbCmact.CommitTrans
    bTrans = False
    Exit Sub
ErrGraba:
    If bTrans Then dbCmact.RollbackTrans
    Err.Raise Err.Number, "NCapDefinicion:ActualizaDocRecParametro", Err.Description
End Sub

Public Function GetDocRecParametro() As ADODB.Recordset
    Dim rsDoc As ADODB.Recordset
    sSql = "Select UPPER(K.cConsDescripcion) + Space(50) + Convert(Varchar(1),C.nMoneda) Moneda, " _
        & "UPPER(K1.cConsDescripcion) + Space(50) + Convert(Varchar(1),C.nTpoCheque) TipoCheque, " _
        & "UPPER(K2.cConsDescripcion) + Space(50) + Convert(Varchar(1),C.nPlaza) Plaza, C.nDiasVal " _
        & "From DocRecParametro C JOIN Constante K ON C.nMoneda = K.nConsValor " _
        & "JOIN Constante K1 ON C.nTpoCheque = K1.nConsValor JOIN " _
        & "Constante K2 ON C.nPlaza = K2.nConsValor Where K.nConsCod = " & gMoneda & " " _
        & "And K1.nConsCod = " & gChequeTipo & " And K2.nConsCod = " & gChequePlaza & " Order by C.nMoneda, C.nTpoCheque, Plaza"

    Set rsDoc = New ADODB.Recordset
    rsDoc.CursorLocation = adUseClient
    rsDoc.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsDoc.ActiveConnection = Nothing
    Set GetDocRecParametro = rsDoc
    Set rsDoc = Nothing
End Function

Public Function GetDiasMinValorizacion(ByVal nMoneda As Moneda, _
    ByVal nPlaza As ChequePlaza, ByVal nTipoCheque As ChequeTipo) As Integer
    Dim rsChq As ADODB.Recordset

    sSql = "Select nDiasVal From DocRecParametro Where nTpoCheque = " & nTipoCheque & " " _
        & "And nPlaza = " & nPlaza & " And nMoneda = " & nMoneda

    Set rsChq = New ADODB.Recordset
    rsChq.CursorLocation = adUseClient
    rsChq.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsChq.ActiveConnection = Nothing
    If rsChq.EOF And rsChq.BOF Then
        GetDiasMinValorizacion = 0
    Else
        GetDiasMinValorizacion = rsChq("nDiasVal")
    End If
    rsChq.Close
    Set rsChq = Nothing
End Function

Public Function GetCapPersoneriaParam(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal bOrdPag As Boolean, Optional ByVal nTpoPrograma As Integer = 0) As ADODB.Recordset
    Dim rsPar As ADODB.Recordset
    Dim sOrdPag As String
    sOrdPag = IIf(bOrdPag, "1", "0")
    Set rsPar = New ADODB.Recordset
    rsPar.CursorLocation = adUseClient
    sSql = "Select K.cConsDescripcion + Space(100) + Convert(Varchar(1),K.nConsValor), ISNULL(P.nMontoMinApert,0), ISNULL(P.nSaldoMin,0) " _
        & "From Constante K LEFT JOIN (Select nPersoneria, nMontoMinApert, nSaldoMin From CapPersParam Where nProducto = " & nProducto & " " _
        & "And nMoneda = " & nMoneda & " And cOrdPag = '" & sOrdPag & "' And nTpoPrograma = " & nTpoPrograma & ") P ON K.nConsValor = P.nPersoneria Where " _
        & "K.nConsCod = " & gPersPersoneria & " and K.nConsValor <> " & gPersPersoneria & "Order by K.nConsValor"
    rsPar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsPar.ActiveConnection = Nothing
    Set GetCapPersoneriaParam = rsPar
    Set rsPar = Nothing
End Function

Public Function GetSaldoMinimoPersoneria(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal nPersoneria As PersPersoneria, ByVal bOrdPag As Boolean, _
        Optional ByVal nTpoPrograma As Integer = 0, Optional ByVal psCtaCod As String = "") As Double
'APRI20190109 ERS077-2018 add psCtaCod
    Dim rsPar As ADODB.Recordset
    Dim sOrdPag As String
    sOrdPag = IIf(bOrdPag, "1", "0")
    Set rsPar = New ADODB.Recordset
    rsPar.CursorLocation = adUseClient
'    sSql = "Select nSaldoMin From CapPersParam Where nPersoneria = " & nPersoneria & " And " _
'        & "nProducto = " & nProducto & " And nMoneda = " & nMoneda & " And cOrdPag = '" & sOrdPag & "' " _
'        & "And nTpoPrograma = " & nTpoPrograma
    sSql = "EXEC stp_sel_GetSaldoMinimoPersoneria " & nPersoneria & "," & nProducto & "," & nMoneda & ",'" & sOrdPag & "'," & nTpoPrograma & ",'" & psCtaCod & "'" 'APRI20190109 ERS077-2018
    rsPar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsPar.ActiveConnection = Nothing
    If rsPar.EOF And rsPar.BOF Then
        GetSaldoMinimoPersoneria = 0
    Else
        GetSaldoMinimoPersoneria = rsPar("nSaldoMin")
    End If
    rsPar.Close
    Set rsPar = Nothing
End Function

Public Function GetMontoMinimoAperturaPersoneria(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal nPersoneria As PersPersoneria, ByVal bOrdPag As Boolean, _
        Optional ByVal nTpoPrograma As Integer = 0) As Double
    Dim rsPar As ADODB.Recordset
    Dim sOrdPag As String
    sOrdPag = IIf(bOrdPag, "1", "0")

    Set rsPar = New ADODB.Recordset
    rsPar.CursorLocation = adUseClient
    sSql = "Select nMontoMinApert From CapPersParam Where nPersoneria = " & nPersoneria & " And " _
        & "nProducto = " & nProducto & " And nMoneda = " & nMoneda & " And cOrdPag = '" & sOrdPag & "' " _
        & "And nTpoPrograma = " & nTpoPrograma
    rsPar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsPar.ActiveConnection = Nothing
    If rsPar.EOF And rsPar.BOF Then
        GetMontoMinimoAperturaPersoneria = 0
    Else
        GetMontoMinimoAperturaPersoneria = rsPar("nMontoMinApert")
    End If
    rsPar.Close
    Set rsPar = Nothing
End Function

Public Function GetMontoMinimoDepPersoneria(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal nPersoneria As PersPersoneria, ByVal bOrdPag As Boolean, _
        Optional ByVal nTpoPrograma As Integer = 0) As Double
    Dim rsPar As ADODB.Recordset
    Dim sOrdPag As String
    sOrdPag = IIf(bOrdPag, "1", "0")

    Set rsPar = New ADODB.Recordset
    rsPar.CursorLocation = adUseClient
    sSql = "Select nMontoMinDep From CapPersParam Where nPersoneria = " & nPersoneria & " And " _
        & "nProducto = " & nProducto & " And nMoneda = " & nMoneda & " And cOrdPag = '" & sOrdPag & "' " _
        & "And nTpoPrograma = " & nTpoPrograma
    rsPar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsPar.ActiveConnection = Nothing
    If rsPar.EOF And rsPar.BOF Then
        GetMontoMinimoDepPersoneria = 0
    Else
        GetMontoMinimoDepPersoneria = IIf(IsNull(rsPar("nMontoMinDep")), 0, rsPar("nMontoMinDep"))
    End If
    rsPar.Close
    Set rsPar = Nothing
End Function

Public Function GetMontoMinimoRetPersoneria(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal nPersoneria As PersPersoneria, ByVal bOrdPag As Boolean, Optional psCtaCod As String = "") As Double
'APRI20190109 ERS077-2018 Add psCtaCod
    Dim rsPar As ADODB.Recordset
    Dim sOrdPag As String
    sOrdPag = IIf(bOrdPag, "1", "0")

    Set rsPar = New ADODB.Recordset
    rsPar.CursorLocation = adUseClient
'    sSql = "Select nMontoMinRet From CapPersParam Where nPersoneria = " & nPersoneria & " And " _
'        & "nProducto = " & nProducto & " And nMoneda = " & nMoneda & " And cOrdPag = '" & sOrdPag & "'"
    sSql = "exec stp_sel_GetMontoMinimoRetPersoneria '" & psCtaCod & "'," & nProducto & "," & nPersoneria & "," & nMoneda & "," & sOrdPag 'APRI20190109 ERS077-2018
    rsPar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsPar.ActiveConnection = Nothing
    If rsPar.EOF And rsPar.BOF Then
        GetMontoMinimoRetPersoneria = 0
    Else
        GetMontoMinimoRetPersoneria = IIf(IsNull(rsPar("nMontoMinRet")), 0, rsPar("nMontoMinRet"))
    End If
    rsPar.Close
    Set rsPar = Nothing
End Function


Public Function GetMontoMinimoRetOPPersoneria(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal nPersoneria As PersPersoneria, ByVal bOrdPag As Boolean, Optional psCtaCod As String = "") As Double
'APRI20190109 ERS077-2018 Add psCtaCod
    Dim rsPar As ADODB.Recordset
    Dim sOrdPag As String
    sOrdPag = IIf(bOrdPag, "1", "0")

    Set rsPar = New ADODB.Recordset
    rsPar.CursorLocation = adUseClient
'    sSql = "Select nMontoMinRetOP From CapPersParam Where nPersoneria = " & nPersoneria & " And " _
'        & "nProducto = " & nProducto & " And nMoneda = " & nMoneda & " And cOrdPag = '" & sOrdPag & "'"
    sSql = "Exec stp_sel_GetMontoMinimoRetPersoneria '" & psCtaCod & "'," & nProducto & "," & nPersoneria & "," & nMoneda & "," & sOrdPag 'APRI20190109 ERS077-2018
    rsPar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsPar.ActiveConnection = Nothing
    If rsPar.EOF And rsPar.BOF Then
        GetMontoMinimoRetOPPersoneria = 0
    Else
        GetMontoMinimoRetOPPersoneria = IIf(IsNull(rsPar("nMontoMinRet")), 0, rsPar("nMontoMinRet"))
    End If
    rsPar.Close
    Set rsPar = Nothing
End Function

Public Sub ActualizaCapPersoneriaParam(ByVal rsPar As ADODB.Recordset, ByVal nProducto As Producto, _
        ByVal nMoneda As Moneda, ByVal bOrdPag As Boolean, ByVal sMovNro As String)
    Dim bTrans As Boolean
    Dim nMonto As Double, nSaldo As Double
    Dim nPers As PersPersoneria
    Dim sOrdPag As String
    sOrdPag = IIf(bOrdPag, "1", "0")

    On Error GoTo ErrGraba
    dbCmact.BeginTrans
    bTrans = True
    sSql = "Delete CapPersParam Where nProducto = " & nProducto & " And nMoneda = " & nMoneda & " And cOrdPag = '" & sOrdPag & "'"
    dbCmact.Execute sSql
    Do While Not rsPar.EOF
        nPers = CLng(Right(rsPar("Personería"), 1))
        nMonto = CDbl(rsPar("Mon Min Apert"))
        nSaldo = CDbl(rsPar("Saldo Minimo"))
        sSql = "Insert CapPersParam (nProducto,nMoneda,nPersoneria,cOrdPag,nMontoMinApert,nSaldoMin,cMovNro) " _
            & "Values (" & nProducto & "," & nMoneda & "," & nPers & ",'" & sOrdPag & "'," & nMonto & "," & nSaldo & ",'" & sMovNro & "')"
        dbCmact.Execute sSql
        rsPar.MoveNext
    Loop
    dbCmact.CommitTrans
    bTrans = False
    Exit Sub
ErrGraba:
    If bTrans Then dbCmact.RollbackTrans
    'Err.Description , vbCritical, "Error"
End Sub

Public Function GetTarifaParametro(ByVal nOperacion As CaptacOperacion, _
            ByVal nMoneda As Moneda, ByVal nParametro As CaptacParametro, Optional psCtaCod As String = "") As Recordset
'APRI20190109 ERS077-2018 Add psCtaCod
    Dim rsPar As ADODB.Recordset

    Set rsPar = New ADODB.Recordset
    rsPar.CursorLocation = adUseClient
'    sSql = "Select P.nParValor From Parametro P JOIN Tarifas T ON P.nParProd = T.nParProd And " _
'        & "P.nParCod = T.nParCod Where T.cOpeCod = '" & nOperacion & "' And " _
'        & "T.nMoneda = " & nMoneda & " And T.nParCod = " & nParametro
    sSql = "Exec stp_sel_GetTarifaParametro '" & nOperacion & "'," & nMoneda & "," & nParametro & ",'" & psCtaCod & "'" 'APRI20190109 ERS077-2018
    rsPar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

    Set rsPar.ActiveConnection = Nothing
    Set GetTarifaParametro = rsPar
    Set rsPar = Nothing

End Function


Public Function EsOperacionEfectivo(ByVal sOperacion As String) As Boolean
    Dim rsEfec As ADODB.Recordset

    'sSql = "Select GO.cOpeCod From GrupoOpe G JOIN GruposOpe GO ON G.cGrupoCod = GO.cGrupoCod " _
    '    & "WHERE GO.cOpeCod = '" & sOperacion & "' And G.nEfectivo = 1"
    
    sSql = " SELECT GO.cOpeCod From GrupoOpe G JOIN GruposOpe GO ON G.cGrupoCod = GO.cGrupoCod " _
            & " WHERE GO.cOpeCod = '" & sOperacion & "' AND G.nEfectivo = 1 " _
            & " And Not cOpeCod like '9010%' and cOpeCod NOT IN (401401, 402401) " _
            & " And cOpeCod not in (Select cOpeCod from OpeDoc Where nDocTpo = 47) "

    Set rsEfec = New ADODB.Recordset
    rsEfec.CursorLocation = adUseClient
    rsEfec.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsEfec.ActiveConnection = Nothing

    If rsEfec.EOF And rsEfec.BOF Then
        EsOperacionEfectivo = False
    Else
        EsOperacionEfectivo = True
    End If

    rsEfec.Close
    Set rsEfec = Nothing
End Function

Public Function ObtenerFeriado(ByVal pdFechaI As Date, ByVal pdFechaF As Date) As Integer

    Dim rs As ADODB.Recordset
    sSql = " select count(dferiado) NroFeriado from feriado where dferiado >= '" & Format(pdFechaI, "YYYYMMDD") & "' and  dferiado <= ' " & Format(pdFechaF, "YYYYMMDD") & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing

    If rs.EOF And rs.BOF Then
        ObtenerFeriado = 0
    Else
        ObtenerFeriado = rs!NroFeriado
    End If

    Set rs = Nothing

End Function


Public Function ObtenerDiasValoriza(ByVal psDoc As String) As Integer
    Dim rs As ADODB.Recordset
    
    sSql = "select dFecReg = CONVERT(Varchar(12),CONVERT(Datetime,SUBSTRING(E.cMovNro,1,8)),103)," & _
        "       dValor = CASE WHEN D.nEstado = 2 THEN Convert(Varchar(12),D.dValorizacion,103) ELSE Convert(Varchar(12),D.dValorizaRef,103) End " & _
        " from DocRec D " & _
        " INNER JOIN  DocRecEst E ON D.nTpoDoc = E.nTpoDoc And D.cNroDoc = E.cNroDoc And D.cPersCod = E.cPersCod And D.cIFTpo = E.cIFTpo" & _
        " WHERE D.cNroDoc = '" & psDoc & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing

    If rs.EOF And rs.BOF Then
        ObtenerDiasValoriza = 0
    Else
        ObtenerDiasValoriza = CDate(rs!dValor) - CDate(rs!dFecReg)
    End If

    Set rs = Nothing

End Function

Private Sub Class_Initialize()
    Dim sConn As String
    Dim ClsIni As COMConecta.DCOMClasIni
    Set ClsIni = New COMConecta.DCOMClasIni
    sConn = ClsIni.CadenaConexion
    sDBComunes = ClsIni.BaseComunes
    sDBPersona = ClsIni.BasePersonas
    sDBImagenes = ClsIni.BaseImagenes
    Set dbCmact = New Connection
    dbCmact.Open sConn
    Set dbCmactHist = New Connection
    'sConn = ClsIni.CadenaConexion(, "HISTORICA")
    dbCmactHist.Open sConn
    Set ClsIni = Nothing
End Sub

Private Sub Class_Terminate()
    dbCmact.Close
    Set dbCmact = Nothing
    dbCmactHist.Close
    Set dbCmactHist = Nothing
End Sub
'Add By GITU 11-08-2010
Public Function GetCapTasaInteresPF(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
    ByVal nTipoTasa As CaptacTipoTasa, Optional nPlazo As Long = 0, Optional nValor As Double = 0, _
    Optional sCodAge As String = "", Optional bOrdPag As Boolean = False, Optional nSubProd As Integer = 0) As Double

    Dim rsTasaPF As ADODB.Recordset
    Dim cOrdPag As String
    
    sSql = "Exec stp_sel_RecuperaTasasPF " & nProducto & "," & nMoneda & "," & nTipoTasa & "," & nPlazo & "," _
            & nValor & "," & nSubProd & ",'" & sCodAge & "'"

    Set rsTasaPF = New ADODB.Recordset
    rsTasaPF.CursorLocation = adUseClient
    rsTasaPF.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsTasaPF.ActiveConnection = Nothing
    If Not (rsTasaPF.EOF Or rsTasaPF.BOF) Then
        GetCapTasaInteresPF = rsTasaPF("nTasaValor")
    Else
        GetCapTasaInteresPF = 0
    End If
    rsTasaPF.Close
    Set rsTasaPF = Nothing
End Function


Public Function GetTarifarioPF(ByVal sProd As Producto, ByVal sMoneda As Moneda, _
        ByVal nTipo As CaptacTipoTasa, ByVal sAgencia As String, _
        Optional ByVal nTpoPrograma As Integer = 0) As ADODB.Recordset
    Dim rsParam As ADODB.Recordset

    Set rsParam = New ADODB.Recordset
    rsParam.CursorLocation = adUseClient
    
    'By Capi 15012008 se adiciono nuevos campos
    '& "nTasaValor, nTasaCod From CaptacTasas Where nTasaProd = " & sProd & " And nTasaMon = " & sMoneda & " And "
    '    sSql = "Select nValorIni, nValorFin, nPlazoIni, nPlazoFin, (Case When cOrdPag='1' Then 'Si' Else 'No' End), " _
    '    & "nTasaValor, nTasaCod,IsNull(cLogFechaCambio,''),cActiva From CaptacTasasPF Where nTasaProd = " & sProd & " And nTasaMon = " & sMoneda & " And " _
    '    & "nTasaTpo = " & nTipo & " And cCodAge = '" & sAgencia & "' And nTpoPrograma = " & nTpoPrograma & " " _
    '    & "Order by cActiva,nPlazoIni,nPlazoFin,nTasaCod,cLogFechaCambio"
    sSql = " exec stp_sel_GetTarifarioPF '" & sProd & "','" & sMoneda & "','" & sAgencia & "','" & nTpoPrograma & "'"
    rsParam.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsParam.ActiveConnection = Nothing

    Set GetTarifarioPF = rsParam
    Set rsParam = Nothing
    Exit Function
ErrGetParam:
    Err.Raise Err.Number, "NCapDefinicion", Err.Description
End Function

Public Sub ActualizaTasaPF(ByVal nCodigo As Integer, ByVal nMonIni As Double, _
        ByVal nMonFin As Double, ByVal nPlazoIni As Long, ByVal nPlazoFin As Long, _
        ByVal bOrdPag As Boolean, ByVal nValor As Double, Optional ByVal pcCambio As String = "", Optional ByVal bActiva As Boolean = True, _
        Optional ByVal nProd As Integer = 0, Optional ByVal nTpoPrograma As Integer = 0, Optional ByVal nMoneda As Integer = 0, Optional ByVal sAgencia As String = "")
        'By Capi 15012008 se agrego parametros cCambio y bActiva
    Dim sActiva As String
    Dim sOrdPag As String
    On Error GoTo ErrGetParam

    sOrdPag = IIf(bOrdPag, "1", "0")
    'By Capi 15012008
    sActiva = IIf(bActiva, "1", "0")
    
    'By Capi 20012008
    'JUEZ 20140220 *****************************************
    'sSql = "Update CaptacTasasPF Set cActiva = '" & sActiva & "'" & " Where nTasaCod = " & nCodigo
    sSql = "EXEC stp_sel_ActualizaTasaPF " & nProd & "," & nTpoPrograma & "," & nMoneda & ",'" & sAgencia & "'"
    'END JUEZ **********************************************
    
    dbCmact.Execute sSql
    Exit Sub
ErrGetParam:
    Err.Raise Err.Number, "NCapDefinicion:ActualizaTasaPF", Err.Description
End Sub

Public Sub NuevaTasaPF(ByVal nProd As Producto, ByVal nMon As Moneda, _
        ByVal nTipoTasa As CaptacTipoTasa, ByVal nMonIni As Double, nMonFin As Double, _
        ByVal nPlazoIni As Long, ByVal nPlazoFin As Long, ByVal bOrdPag As Boolean, _
        ByVal sAgencia As String, ByVal nValor As Double, Optional ByVal nTpoPrograma As Integer = 0, _
        Optional ByVal pcCambio As String = "", Optional ByVal bActiva As Boolean = True, _
        Optional ByVal pnTasaCodAnt As Double)
        'By capi 15012008 se agrego 2 parametros opcionales

    Dim sOrdPag As String
    Dim sActiva As String
    On Error GoTo ErrTasa
    sOrdPag = IIf(bOrdPag, "1", "0")
    sActiva = IIf(bActiva, "1", "0")
    'JUEZ 20140220 ***************************************************************
    'sSql = "Insert CaptacTasasPF (nTasaProd,nTasaMon,nTasaTpo,nValorIni,nValorFin,nPlazoIni,nPlazoFin,cOrdPag,cCodAge,nTasaValor,nTpoPrograma,clogFechaCambio,cActiva,nTasaCodAnt) " _
    '        & "Values (" & nProd & "," & nMon & "," & nTipoTasa & "," & nMonIni & "," & nMonFin & "," & nPlazoIni & "," & nPlazoFin & ",'" & sOrdPag & "','" & sAgencia & "'," & nValor & "," & nTpoPrograma & ",'" & pcCambio & "','" & sActiva & "'," & pnTasaCodAnt & ")"
    sSql = "EXEC stp_sel_NuevaTasaPF " & nProd & "," & nMon & "," & nTipoTasa & "," & nMonIni & "," & nMonFin & "," & nPlazoIni & "," & nPlazoFin & "," & sOrdPag & ",'" & sAgencia & "'," & nValor & "," & nTpoPrograma & ",'" & pcCambio & "'"
    'END JUEZ ********************************************************************
    
    dbCmact.Execute sSql
    Exit Sub
ErrTasa:
    Err.Raise Err.Number, "NCapDefinicion:NuevaTasa", Err.Description
End Sub

Public Function ActualizaParametrosPIT(ByVal nVariable As Long, ByVal sDesc As String, ByVal nValor As Double) As Boolean
    Dim rsPar As New ADODB.Recordset
    Dim sDescH As String, sVar As String
    Dim nVal As Double
    On Error GoTo ErrGetParam
    sSql = "Update DBTarjeta..PITParametro Set cDescripcion = '" & sDesc & "', nValor = " & nValor & " WHERE " _
        & "nParametroId = " & nVariable
    dbCmact.Execute sSql
    ActualizaParametrosPIT = True
    Exit Function
ErrGetParam:
    Err.Raise Err.Number, "NCapDefinicion:ActualizaParametros", Err.Description
    ActualizaParametrosPIT = False
End Function

'JUEZ 20130828 ********************************************************************
'Public Function GetParametrosComision(Optional ByVal psParCod As String = "", Optional ByVal pbConsultaDifAho As Boolean = False) As ADODB.Recordset
Public Function GetParametrosComision(Optional ByVal psParCod As String = "", Optional ByVal pbConsultaDifAho As Boolean = False, Optional ByVal psParProd As String = "") As ADODB.Recordset 'JUEZ 20150930
    Dim rsParam As ADODB.Recordset
    'sSql = "exec stp_sel_GetParametrosComision '" & psParCod & "','" & IIf(pbConsultaDifAho, "1", "0,1") & "'"
    sSql = "exec stp_sel_GetParametrosComision '" & psParCod & "','" & IIf(pbConsultaDifAho, "1", "0,1") & "','" & psParProd & "'" 'JUEZ 20150930
    Set rsParam = New ADODB.Recordset
    rsParam.CursorLocation = adUseClient
    rsParam.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rsParam.ActiveConnection = Nothing
    Set GetParametrosComision = rsParam
    Set rsParam = Nothing
    Exit Function
End Function
'Public Function ObtenerCodigoParametroComision() As String
Public Function ObtenerCodigoParametroComision(ByVal psParProd As String) As String 'JUEZ 20150930
    Dim rs As ADODB.Recordset
    sSql = "exec stp_sel_ObtenerCodigoParametroComision"
    sSql = "exec stp_sel_ObtenerCodigoParametroComision '" & psParProd & "'" 'JUEZ 20150930
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    ObtenerCodigoParametroComision = rs!cCodigo
    Set rs = Nothing
    Exit Function
End Function
Public Sub ActualizaParametrosComision(ByVal pnOperacion As Integer, ByVal psParCod As String, ByVal psParDesc As String, _
                                       ByVal pnMonto As Double, ByVal pnMoneda As Integer, ByVal pnTipo As Integer, ByVal psParProd As String)
                                       'JUEZ 20150930 Se agregó psParProd
    Dim rs As ADODB.Recordset
    If pnOperacion = 1 Then
        'sSql = "exec stp_ins_RegistrarParametrosComision '" & psParCod & "','" & psParDesc & "'," & pnMonto & "," & pnMoneda & "," & pnTipo
        sSql = "exec stp_ins_RegistrarParametrosComision '" & psParCod & "','" & psParDesc & "'," & pnMonto & "," & pnMoneda & "," & pnTipo & ",'" & psParProd & "'" 'JUEZ 20150930
    ElseIf pnOperacion = 2 Then
        sSql = "exec stp_upd_ActualizarParametrosComision '" & psParCod & "','" & psParDesc & "'," & pnMonto & "," & pnMoneda & "," & pnTipo
    ElseIf pnOperacion = 3 Then
        sSql = "exec stp_upd_EliminarParametrosComision '" & psParCod & "'"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    Set rs = Nothing
    Exit Sub
End Sub
'END JUEZ *************************************************************************

' *** RIRO20131206 ERS137 ***
Public Function getMinComisionTransferencia() As Double
    Dim rs As ADODB.Recordset
    Dim nMinComisionTranf As Double
    
    On Error GoTo Error
    
    sSql = "exec stp_sel_ComisionMinimoTransferencia"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            nMinComisionTranf = CDbl(rs!nConsSisValor)
        Else
            nMinComisionTranf = 0
        End If
    Else
        nMinComisionTranf = 0
    End If
    
    getMinComisionTransferencia = nMinComisionTranf
    
    Set rs = Nothing
    Exit Function
    
Error:
        Err.Raise Err.Number, "NCapDefinicion:getMinComisionTransferencia", Err.Description
        
End Function

Public Function ActualizaMinComisionTransferencia(ByVal pnComision As Double) As Boolean
           
    On Error GoTo Error
    sSql = "stp_upd_ComisionMinimoTransferencia " & pnComision
    dbCmact.Execute sSql
    
    ActualizaMinComisionTransferencia = True
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:ActualizaMinComisionTransferencia", Err.Description
    ActualizaMinComisionTransferencia = False
    
End Function

Public Function RegistraComisionTransferencia(ByVal cIdBanco As String, _
                                              ByVal nTipoOpe As Integer, _
                                              ByVal nPlaza As Integer, _
                                              ByVal nMonedaOpe As Integer, _
                                              ByVal nDesde As Double, _
                                              ByVal nHasta As Double, _
                                              ByVal nTipoCalc As Integer, _
                                              ByVal nMonedaComi As Integer, _
                                              ByVal nMontoComi As Double, _
                                              ByVal nMontoMinimo As Integer, _
                                              ByVal nEstado As Integer) As Boolean
                                                           
    On Error GoTo Error
    
    sSql = "stp_ins_ComiTransferenciaBanc '" & cIdBanco & "', " & nTipoOpe & ", "
    sSql = sSql & nPlaza & ", " & nMonedaOpe & ", " & nDesde & ", " & nHasta & ", "
    sSql = sSql & nTipoCalc & ", " & nMonedaComi & ", "
    sSql = sSql & nMontoComi & ", " & nMontoMinimo & ", " & nEstado
    
    dbCmact.Execute sSql
    
    RegistraComisionTransferencia = True
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:RegistraComisionTransferencia", Err.Description
    RegistraComisionTransferencia = False
    
End Function

Public Function ActualizaComisionTransferencia(ByVal nIDComision As Double, _
                                               ByVal cIdBanco As String, _
                                               ByVal nTipoOpe As Integer, _
                                               ByVal nPlaza As Integer, _
                                               ByVal nMonedaOpe As Integer, _
                                               ByVal nDesde As Double, _
                                               ByVal nHasta As Double, _
                                               ByVal nTipoCalc As Integer, _
                                               ByVal nMonedaComi As Integer, _
                                               ByVal nMontoComi As Double, _
                                               ByVal nMontoMinimo As Integer, _
                                               ByVal nEstado As Integer) As Boolean
                                                

           
    On Error GoTo Error
    sSql = "stp_upd_ComiTransferenciaBanc " & nIDComision & ", '" & cIdBanco & "', " & nTipoOpe & ", "
    sSql = sSql & nPlaza & ", " & nMonedaOpe & ", " & nDesde & ", " & nHasta & ", "
    sSql = sSql & nTipoCalc & ", " & nMonedaComi & ", "
    sSql = sSql & nMontoComi & ", " & nMontoMinimo & ", " & nEstado
    dbCmact.Execute sSql
    
    ActualizaComisionTransferencia = True
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:ActualizaComisionTransferencia", Err.Description
    ActualizaComisionTransferencia = False
    
End Function

Public Function GetComisionTransferencia() As ADODB.Recordset
    
    Dim rsParam As ADODB.Recordset
    Set rsParam = New ADODB.Recordset
    
    On Error GoTo Error
    
    rsParam.CursorLocation = adUseClient
    sSql = "stp_sel_ComisionesTransferencia"
    rsParam.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsParam.ActiveConnection = Nothing

    Set GetComisionTransferencia = rsParam
    Set rsParam = Nothing
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:GetComisionTransferencia", Err.Description
End Function

Public Function BuscarComisionTransferencia(ByVal psIdBanco As String, _
                                            ByVal pnPlaza As Integer, _
                                            ByVal pnMoneda As Integer, _
                                            ByVal pnTipoOpe As Integer, _
                                            ByVal pnMonto As Double) As Recordset

    
    On Error GoTo Error
    
    Dim rs As ADODB.Recordset
    Dim nComisionTranf As Double
    sSql = "exec stp_sel_BuscarComisionTransferencia '" & psIdBanco _
                                                        & "', " _
                                                        & pnPlaza _
                                                        & ", " _
                                                        & IIf(pnMoneda = 1, 301, 302) _
                                                        & ", " _
                                                        & pnTipoOpe _
                                                        & ", " _
                                                        & pnMonto
    
    
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            Set BuscarComisionTransferencia = rs
        Else
            Set BuscarComisionTransferencia = Nothing
        End If
    Else
        Set BuscarComisionTransferencia = Nothing
    End If
            
    Exit Function
    
Error:
    Err.Raise Err.Number, "NCapDefinicion:GetComisionTransferencia", Err.Description
    
End Function

Public Function getCalculaComision(ByVal idBanco As String, _
                            ByVal nPlaza As Integer, _
                            ByVal nMoneda As Integer, _
                            ByVal nTipo As Integer, _
                            ByVal nMonto As Double, _
                            ByVal sFecha As String) As Double

Dim nComision, nTC As Double
Dim nMontoMinimo As Double
Dim rsComision As Recordset
Dim oDefinicion As COMNCaptaGenerales.NCOMCaptaDefinicion

On Error GoTo Error

Set oDefinicion = New COMNCaptaGenerales.NCOMCaptaDefinicion
Set rsComision = New ADODB.Recordset

Dim clsTC As COMDConstSistema.NCOMTipoCambio
Set clsTC = New COMDConstSistema.NCOMTipoCambio

Set rsComision = oDefinicion.BuscarComisionTransferencia(idBanco, nPlaza, nMoneda, nTipo, nMonto)
nMontoMinimo = oDefinicion.getMinComisionTransferencia

If Not rsComision Is Nothing Then
    If Not rsComision.BOF And Not rsComision.EOF Then
        nComision = CDbl(rsComision!nMontoPagar)
        If nMoneda = gMonedaNacional Then
            nTC = clsTC.EmiteTipoCambio(sFecha, TCVenta)
            If rsComision!moneda_comi = "501" Or rsComision!moneda_comi = "503" Then ' Soles
            ElseIf rsComision!moneda_comi = "502" Then ' Dolares
                nComision = nComision * nTC
            End If
            If nTipo <> 101 Then
                If nComision < nMontoMinimo Then
                    nComision = nMontoMinimo
                End If
            End If
        ElseIf nMoneda = gMonedaExtranjera Then
            nTC = clsTC.EmiteTipoCambio(sFecha, TCVenta)
            If rsComision!moneda_comi = "501" Then ' Soles
                nComision = nComision / nTC
            ElseIf rsComision!moneda_comi = "502" Or rsComision!moneda_comi = "503" Then ' Dolares
            End If
            If nTipo <> 101 Then
                If nComision * nTC < nMontoMinimo Then
                    nComision = nMontoMinimo / nTC
                End If
            End If
        End If
    Else
        If nMoneda = gMonedaNacional Then
            nComision = nMontoMinimo
        ElseIf nMoneda = gMonedaExtranjera Then
            nTC = clsTC.EmiteTipoCambio(sFecha, TCVenta)
            nComision = nMontoMinimo / nTC
        End If
        If nTipo = 101 Then
            nComision = 0
        End If
        
    End If
Else
        If nMoneda = gMonedaNacional Or nMoneda = 0 Then
            nComision = nMontoMinimo
        ElseIf nMoneda = gMonedaExtranjera Then
            nTC = clsTC.EmiteTipoCambio(sFecha, TCVenta)
            nComision = nMontoMinimo / nTC
        End If
        If nTipo = 101 Then
            nComision = 0
        End If
End If
getCalculaComision = nComision
Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:getCalculaComision", Err.Description
    getCalculaComision = 0
End Function

' *** FIN RIRO ***
'RIRO20140112 ERS162 *****************

Public Function getSiniestros(ByVal pnTipoSeguro As Double) As Recordset
    On Error GoTo Error
    Dim rs As ADODB.Recordset
    Dim nComisionTranf As Double
    sSql = "exec stp_sel_Listar_Siniestros " & pnTipoSeguro
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            Set getSiniestros = rs
        Else
            Set getSiniestros = Nothing
        End If
    Else
        Set getSiniestros = Nothing
    End If
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:getTipoSeguro", Err.Description
End Function
Public Function getDocumentos(ByVal pnTipoSeguro As TipoSeguro) As Recordset
    
    On Error GoTo Error
    Dim rs As ADODB.Recordset
    Dim nComisionTranf As Double
    sSql = "exec stp_sel_Listar_SegDocumentos " & pnTipoSeguro
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            Set getDocumentos = rs
        Else
            Set getDocumentos = Nothing
        End If
    Else
        Set getDocumentos = Nothing
    End If
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:getTipoSeguro", Err.Description
End Function
Public Function getTipoSeguro() As Recordset
    
    On Error GoTo Error
    Dim rs As ADODB.Recordset
    Dim nComisionTranf As Double
    sSql = "exec stp_sel_Listar_Tipo_Seguro"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            Set getTipoSeguro = rs
        Else
            Set getTipoSeguro = Nothing
        End If
    Else
        Set getTipoSeguro = Nothing
    End If
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:getTipoSeguro", Err.Description
End Function
Public Function getTipoSeguroDestinatario(Optional ByVal nTipo As TipoSeguro = 0) As Recordset
    
    On Error GoTo Error
    Dim rs As ADODB.Recordset
    Dim nComisionTranf As Double
    sSql = "exec stp_sel_Listar_Tipo_Seguro_Destinatario " & nTipo
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            Set getTipoSeguroDestinatario = rs
        Else
            Set getTipoSeguroDestinatario = Nothing
        End If
    Else
        Set getTipoSeguroDestinatario = Nothing
    End If
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:getTipoSeguroDestinatario", Err.Description
End Function
Public Function GrabarParametros(ByVal pArrDocumento As Variant, _
                                 ByVal pnOpeDocumento As Integer, _
                                 ByVal pArrTpoSiniestro As Variant, _
                                 ByVal pnOpeSiniestro As Integer, _
                                 ByVal psDestinatario As String, _
                                 ByVal psEntidad As String, _
                                 ByVal pnTipoSeguro As Double, _
                                 ByVal psTipoSeguro As String, _
                                 Optional pnDestinatarioOp As Integer = 0) As Boolean

    On Error GoTo Error
    Dim sArrTmp() As String
    Dim bTrans As Boolean
    Dim nI As Integer
        
    dbCmact.BeginTrans
    bTrans = True
    'Operaciones con documentos: 1: Nuevo, 2: Edicion, 0: Ninguno ****
    If pnOpeDocumento = 1 Then
        For nI = 1 To UBound(pArrDocumento)
            sArrTmp = Split(pArrDocumento(nI), "|")
            sSql = "stp_ins_SegDocumento " & sArrTmp(1) & ", '" & sArrTmp(0) & "'"
            dbCmact.Execute sSql
        Next
    ElseIf pnOpeDocumento = 2 Then
        For nI = 1 To UBound(pArrDocumento)
            sArrTmp = Split(pArrDocumento(nI), "|")
            sSql = "stp_upd_SegParametro " & sArrTmp(1) & ", '" & sArrTmp(0) & "', '',''"
            dbCmact.Execute sSql
        Next
    End If
    'Operaciones con Siniestros
    If pnOpeSiniestro = 1 Then
        For nI = 1 To UBound(pArrTpoSiniestro)
            sArrTmp = Split(pArrTpoSiniestro(nI), "|")
            sSql = "stp_ins_SegSiniestro " & sArrTmp(1) & ", '" & sArrTmp(0) & "'"
            dbCmact.Execute sSql
        Next
    ElseIf pnOpeSiniestro = 2 Then
        For nI = 1 To UBound(pArrTpoSiniestro)
            sArrTmp = Split(pArrTpoSiniestro(nI), "|")
            sSql = "stp_upd_SegParametro " & sArrTmp(1) & ", '" & sArrTmp(0) & "', '',''"
            dbCmact.Execute sSql
        Next
    End If
    'Operaciones con los destinatarios
    If pnDestinatarioOp > 0 Then
        sSql = "stp_upd_SegParametro " & pnTipoSeguro & ", '" & psTipoSeguro & "', '" & psDestinatario & "','" & psEntidad & "'"
        dbCmact.Execute sSql
    End If
    dbCmact.CommitTrans
    bTrans = False
    
    GrabarParametros = True
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:RegistrarDocumento", Err.Description
    GrabarParametros = False
    
End Function
'Public Function RegistrarSiniestro(ByVal pTipoSeguro As Double, ByVal pSiniestro As String) As Boolean
'
'    On Error GoTo Error
'
'    sSql = "STP_INS_SINIESTRO " & pTipoSeguro & ", '" & pSiniestro & "'"
'    dbCmact.Execute sSql
'    RegistrarSiniestro = True
'    Exit Function
'Error:
'    Err.Raise Err.Number, "NCapDefinicion:RegistrarDocumento", Err.Description
'    RegistrarSiniestro = False
'
'End Function
Public Function ActualizarParametro(ByVal pIndice As Double, ByVal pParametro As String, _
                                    Optional ByVal sDestinatario As String = "", _
                                    Optional ByVal sEntidad As String = "") As Boolean
           
    On Error GoTo Error
    
    sSql = "STP_UPD_PARAMETRO " & pIndice & ", '" & pParametro & "', '" & sDestinatario & "','" & sEntidad & "'"
    
    dbCmact.Execute sSql
    ActualizarParametro = True
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:RegistrarDocumento", Err.Description
    ActualizarParametro = False
    
End Function
Public Function QuitarParametro(ByVal pIndice As Double) As Boolean
           
    On Error GoTo Error
    
    sSql = "stp_upd_Quitar_SegParametro " & pIndice
    dbCmact.Execute sSql
    QuitarParametro = True
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:RegistrarDocumento", Err.Description
    QuitarParametro = False
End Function
Public Function getCreditoCobertura(Optional ByVal sPersCod As String = "", _
                                    Optional ByVal nTipoSeguro As Long = 0, _
                                    Optional ByVal nSolicitudcobertura As Long = 0, _
                                    Optional ByVal dFechaSiniestro As Date = "1900-01-01", _
                                    Optional ByVal dFechaCancelacion As Date = "1900-01-01", _
                                    Optional ByVal nEstadoSolicitud As Integer = 1) As ADODB.Recordset
           
    On Error GoTo Error
    
    Dim rs As ADODB.Recordset
    Dim nComisionTranf As Double
    
    If nSolicitudcobertura = 0 Then
        sSql = "exec stp_sel_CreditosCliente '" & Format(dFechaSiniestro, "yyyymmdd") & "', '" & sPersCod & "'"
    Else
        sSql = "exec stp_sel_CreditoCobertura " & nSolicitudcobertura & ", " & nEstadoSolicitud & ", '" & Format(dFechaCancelacion, "yyyymmdd") & "'"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            Set getCreditoCobertura = rs
        Else
            Set getCreditoCobertura = Nothing
        End If
    Else
        Set getCreditoCobertura = Nothing
    End If

    Set rs = Nothing
    
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:getCreditoCobertura", Err.Description
End Function
Public Function getDatosCarta() As ADODB.Recordset
           
    On Error GoTo Error
    
    Dim rs As ADODB.Recordset
    Dim nComisionTranf As Double
    sSql = "exec STP_SEL_DATOSCARTA "
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            Set getDatosCarta = rs
        Else
            Set getDatosCarta = Nothing
        End If
    Else
        Set getDatosCarta = Nothing
    End If
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:getDatosCarta", Err.Description
    
End Function
Public Function RegistrarSolicitudCobertura(ByVal nIdSeguro As Double, ByVal nIdSiniestro As Double, _
                                            ByVal sPersCod As String, ByVal dFechaSiniestro As Date, _
                                            ByVal dFechaSolicitud As Date, ByVal nEstado As Integer, _
                                            ByVal nCarta As Integer, ByVal nAnio As Integer, _
                                            ByVal lDocAdjuntos As Variant, ByVal lCredCliente As Variant) As Long
    On Error GoTo Error

    Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim rs As ADODB.Recordset
    Dim nIdSolicitud As Long, nDocAdjunto As Long, nCredCliente As Long
    Dim i As Integer, J As Integer
    Dim lsTmp() As String

    sSql = "STP_INS_SOLICITUDCOBERTURA " & nIdSeguro & ", " & nIdSiniestro & ", '" & sPersCod & "', '" & Format(dFechaSiniestro, "yyyymmdd") & "', '" & _
                                           Format(dFechaSolicitud, "yyyymmdd") & "', " & nEstado & ", " & nCarta & ", " & nAnio

    Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
    
    clsCap.SetConexion dbCmact
    clsCap.bTransaccion = True
    clsCap.dbCmact.BeginTrans
    
    'Registrando la Solicitud de Cobertura
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            nIdSolicitud = rs!nIdSolicitudCobertura
        Else
            nIdSolicitud = 0
        End If
    Else
        nIdSolicitud = 0
    End If
    
    'Registrando la Documentos Adjuntos
    For i = 1 To UBound(lDocAdjuntos)
        lsTmp = Split(lDocAdjuntos(i), "|")
        nDocAdjunto = RegistrarDocsAdjuntos(nIdSolicitud, lsTmp(1), lsTmp(2))
        If nDocAdjunto = 0 Then
            RegistrarSolicitudCobertura = 0
            clsCap.dbCmact.RollbackTrans
            Set clsCap = Nothing
            Exit Function
        End If
    Next
    
    'Registrando la Creditos Cliente
    For i = 1 To UBound(lCredCliente)
        lsTmp = Split(lCredCliente(i), "|")
        nCredCliente = RegistrarCredCliente(nIdSolicitud, lsTmp(1), lsTmp(2), lsTmp(3), lsTmp(4))
        If nCredCliente = 0 Then
            RegistrarSolicitudCobertura = 0
            clsCap.dbCmact.RollbackTrans
            Set clsCap = Nothing
            Exit Function
        End If
    Next
    
    clsCap.dbCmact.CommitTrans
    Set clsCap = Nothing
    RegistrarSolicitudCobertura = nIdSolicitud
    
    Exit Function
    
Error:

    RegistrarSolicitudCobertura = 0
    clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
End Function
Public Function RegistrarDocsAdjuntos(ByVal nIdSolicitudCobertura As Long, ByVal nIdDocAdj As Long, ByVal nTipoDocumento As Long, Optional ByVal pnTpoCober As Integer = 1, Optional ByVal pnTpoSeguro As Long = 0) As Long
           
    On Error GoTo Error

    Dim rs As ADODB.Recordset
    Dim nDocAdjunto As Long
    
    sSql = "STP_INS_DocAdjuntosCobertura " & nIdSolicitudCobertura & ", " & nIdDocAdj & ", " & nTipoDocumento & ", " & 1 & "," & pnTpoCober & "," & pnTpoSeguro
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            nDocAdjunto = rs!idDocumentosAdjuntos
        Else
            nDocAdjunto = 0
        End If
    Else
        nDocAdjunto = 0
    End If
    
    RegistrarDocsAdjuntos = nDocAdjunto
    
    Exit Function
Error:
    RegistrarDocsAdjuntos = 0
End Function
Public Function RegistrarCredCliente(ByVal nSolicitudcobertura As Long, ByVal sCtaCod As String, ByVal nCapital As Double, ByVal nInteres As Double, ByVal nEstado As Integer) As Long
    On Error GoTo Error

    Dim nCredCliente As Long
    Dim rs As ADODB.Recordset
    
    sSql = "STP_INS_CredClienteCobertura " & nSolicitudcobertura & ", '" & sCtaCod & "', " & nCapital & ", " & nInteres & ", " & 1

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            nCredCliente = rs!idCredCliente
        Else
            nCredCliente = 0
        End If
    Else
        nCredCliente = 0
    End If
    
    RegistrarCredCliente = nCredCliente
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:RegistrarDocumento", Err.Description
    RegistrarCredCliente = nCredCliente
End Function
Public Function getSolicitudCobertura(Optional ByVal nCarta As Long = 0, Optional ByVal nAnio As Integer = 0, _
                                      Optional ByVal nIdSeguro As Double = 0, Optional ByVal nEstado As Byte = 1) As ADODB.Recordset
    ' Valores: 1:Activo     0: Anulado
    On Error GoTo Error
    
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_Buscar_Solicitud_Cobertura " & nCarta & ", " & nAnio & ", " & nIdSeguro & ", " & nEstado
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            Set getSolicitudCobertura = rs
        Else
            Set getSolicitudCobertura = Nothing
        End If
    Else
        Set getSolicitudCobertura = Nothing
    End If
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:getSolicitudCobertura", Err.Description
End Function
Public Function ActualizarSolicitudCobertura(ByVal nIdSolicitudCobertura As Long, ByVal nEstado As Integer) As Boolean
           
    On Error GoTo Error
    sSql = "STP_UPD_SOLICITUD_COBERTURA " & nIdSolicitudCobertura & ", " & nEstado
    dbCmact.Execute sSql
    ActualizarSolicitudCobertura = True
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:ActualizarSolicitudCobertura", Err.Description
    ActualizarSolicitudCobertura = False
End Function
Public Function AceptarSolicitud(ByVal nIdSolicitudCobertura As Long, ByVal sCtaSoles As String, ByVal sCtaDolares As String, ByVal dFechaSoles As Date, ByVal dFechaDolares As Date, _
                                  ByVal nMontoSoles As Long, ByVal nMontoDolares As Long, ByVal nMovPSoles As Long, ByVal nMovPDolares As Long, _
                                  ByVal sMovOPSoles As String, ByVal sMovOPDolares As String, ByVal lCreditos As Variant) As Boolean

On Error GoTo Error

    Dim nMovOPSoles As Long, nMovOPDolares As Long
    Dim sLista() As String
    Dim i As Integer
    Dim clsMant As New COMDCaptaGenerales.DCOMCaptaMovimiento
    'corregir el registro de los pendientes
    clsMant.dbCmact.BeginTrans
    If nMontoSoles > 0 Then
         clsMant.AgregaMov sMovOPSoles, "100232", "", gMovEstContabNoContable, gMovFlagVigente
         nMovOPSoles = clsMant.GetnMovNro(sMovOPSoles)
         clsMant.AgregaMovRef nMovOPSoles, nMovPSoles
    End If
    If nMontoDolares > 0 Then
         clsMant.AgregaMov sMovOPSoles, "100232", "", gMovEstContabNoContable, gMovFlagVigente
         nMovOPDolares = clsMant.GetnMovNro(sMovOPDolares)
         clsMant.AgregaMovRef nMovOPDolares, nMovPDolares
    End If
    sSql = "stp_upd_Aceptar_Solicitad_Cobertura " & _
                                              nIdSolicitudCobertura & ", '" & _
                                              sCtaSoles & "', '" & _
                                              sCtaDolares & "', '" & _
                                              dFechaSoles & "', '" & _
                                              dFechaDolares & "', " & _
                                              nMontoSoles & ", " & _
                                              nMontoDolares & ", " & _
                                              nMovPSoles & ", " & _
                                              nMovPDolares & ", " & _
                                              nMovOPSoles & ", " & _
                                              nMovOPDolares
    clsMant.dbCmact.Execute sSql
    
    For i = 1 To UBound(lCreditos)
        sLista = Split(lCreditos(i), "|")
        sSql = "stp_upd_Actualizar_Credito " & sLista(0) & ", " & Format(sLista(1), "#0.00") & ", 2"
        clsMant.dbCmact.Execute sSql
    Next
    
    clsMant.dbCmact.CommitTrans
    AceptarSolicitud = True
    Set clsMant = Nothing
    
Exit Function
Error:
    AceptarSolicitud = False
    clsMant.dbCmact.RollbackTrans
End Function
Public Function getParametroSolicitud(ByVal nCodParametro As Double) As String

    On Error GoTo Error
    
    Dim rs As ADODB.Recordset
    Dim sParametro As String
    
    sSql = "exec STP_SEL_PARAMETROSOLICITUD " & nCodParametro
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            sParametro = rs!cDescripcion
        Else
            sParametro = ""
        End If
    Else
        sParametro = ""
    End If
    getParametroSolicitud = IIf(IsNumeric(sParametro), sParametro, 0)
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:getParametroSolicitud", Err.Description

End Function

'END RIRO ****************************
'JUEZ 20140220 *****************************************************************
Public Function GetTarifarioHistorico(ByVal sProd As Producto, ByVal sMoneda As Moneda, ByVal sAgencias As String, ByVal nTpoPrograma As Integer, Optional ByVal nTipoTarifario As Integer) As ADODB.Recordset 'JIPR20190109 ERS077-2018 ,Optional ByVal nTipoTarifario As Integer
    Dim rsParam As ADODB.Recordset

    Set rsParam = New ADODB.Recordset
    rsParam.CursorLocation = adUseClient
    
    sSql = " exec stp_sel_GetTarifarioHistorico '" & sProd & "','" & sMoneda & "','" & sAgencias & "','" & nTpoPrograma & "','" & nTipoTarifario & "'" 'JIPR20190109 ERS077-2018 ,'" & nTipoTarifario & "'"
    rsParam.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsParam.ActiveConnection = Nothing

    Set GetTarifarioHistorico = rsParam
    Set rsParam = Nothing
    Exit Function
ErrGetParam:
    Err.Raise Err.Number, "NCapDefinicion", Err.Description
End Function
Public Function GetTarifarioHistoricoPF(ByVal sProd As Producto, ByVal sMoneda As Moneda, ByVal sAgencias As String, ByVal nTpoPrograma As Integer) As ADODB.Recordset
    Dim rsParam As ADODB.Recordset

    Set rsParam = New ADODB.Recordset
    rsParam.CursorLocation = adUseClient
    
    sSql = " exec stp_sel_GetTarifarioHistoricoPF '" & sProd & "','" & sMoneda & "','" & sAgencias & "','" & nTpoPrograma & "'"
    rsParam.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsParam.ActiveConnection = Nothing

    Set GetTarifarioHistoricoPF = rsParam
    Set rsParam = Nothing
    Exit Function
ErrGetParam:
    Err.Raise Err.Number, "NCapDefinicion", Err.Description
End Function
'END JUEZ **********************************************************************
'JUEZ 20140908 **************************************************************************************
Public Function GetCapParametroNew(ByVal pnProd As Producto, ByVal pnTpoPrograma As Integer, Optional ByVal pcCtaCod As String = "") As ADODB.Recordset
'APRI20190109 ERS077-2018 ADD pcCtaCod
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_CapParametroNew " & pnProd & "," & pnTpoPrograma & ",'" & pcCtaCod & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    Set GetCapParametroNew = rs
    Set rs = Nothing
End Function
Public Function ActualizaCapParametroNew(ByVal pnProd As Producto, ByVal pnTpoPrograma As Integer, ByVal psMovNro As String, pMatParametros() As String) As Boolean
On Error GoTo Error
    If pnProd = gCapAhorros Then
        sSql = "stp_ins_CapAhoParametro " & pnTpoPrograma & "," & pMatParametros(0) & "," & pMatParametros(1) & "," & pMatParametros(2) & "," & pMatParametros(3) & _
                                        "," & pMatParametros(4) & "," & pMatParametros(5) & "," & pMatParametros(6) & "," & pMatParametros(7) & "," & pMatParametros(8) & _
                                        "," & pMatParametros(9) & "," & pMatParametros(10) & "," & pMatParametros(11) & "," & pMatParametros(12) & "," & pMatParametros(13) & _
                                        "," & pMatParametros(14) & ",'" & psMovNro & "'"
    ElseIf pnProd = gCapPlazoFijo Then
        sSql = "stp_ins_CapPFParametro " & pnTpoPrograma & "," & pMatParametros(0) & "," & pMatParametros(1) & "," & pMatParametros(2) & "," & pMatParametros(3) & _
                                        "," & pMatParametros(4) & "," & pMatParametros(5) & "," & pMatParametros(6) & "," & pMatParametros(7) & "," & pMatParametros(8) & _
                                        "," & pMatParametros(9) & "," & pMatParametros(10) & "," & pMatParametros(11) & "," & pMatParametros(12) & "," & pMatParametros(13) & _
                                        "," & pMatParametros(14) & ",'" & psMovNro & "'"
    ElseIf pnProd = gCapCTS Then
        sSql = "stp_ins_CapCTSParametro " & pnTpoPrograma & "," & pMatParametros(0) & "," & pMatParametros(1) & "," & pMatParametros(2) & "," & pMatParametros(3) & _
                                        "," & pMatParametros(4) & "," & pMatParametros(5) & ",'" & psMovNro & "'"
    End If
    dbCmact.Execute sSql
    
    ActualizaCapParametroNew = True
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:ActualizaCapParametroNew", Err.Description
    ActualizaCapParametroNew = False
End Function
Public Function ValidaCantOperaciones(ByVal psOpeCod As CaptacOperacion, ByVal pnProd As Producto, ByVal pnOperacion As CaptacMovTipo) As Integer
On Error GoTo ErrorValidaCantOperaciones
    Dim oGen As COMDConstSistema.DCOMGeneral
    Dim sOpeValida As String
    Dim sOpeCod() As String
    Dim nConsSisCod As Integer
    Dim i As Integer
    
    If pnProd = gCapAhorros And pnOperacion = gCapMovDeposito Then
        nConsSisCod = 351
    ElseIf pnProd = gCapAhorros And pnOperacion = gCapMovRetiro Then
        nConsSisCod = 352
    ElseIf pnProd = gCapCTS And pnOperacion = gCapMovDeposito Then
        nConsSisCod = 353
    'APRI20190109 ERS077-2018
    ElseIf pnProd = gCapCTS And pnOperacion = gCapMovRetiro Then
        nConsSisCod = 356
    'END APRI
    ElseIf pnProd = gCapPlazoFijo And pnOperacion = gCapMovDeposito Then
        nConsSisCod = 354
    End If
    
    Set oGen = New COMDConstSistema.DCOMGeneral
    sOpeValida = oGen.LeeConstSistema(nConsSisCod)
    
    sOpeCod = Split(sOpeValida, ",")
    ValidaCantOperaciones = False
    
    For i = 0 To UBound(sOpeCod)
        If psOpeCod = sOpeCod(i) Then
            ValidaCantOperaciones = True
            Exit Function
        End If
    Next i
    
    Exit Function
ErrorValidaCantOperaciones:
    MsgBox Err.Description, vbCritical, "Error"
End Function
'END JUEZ *******************************************************************************************

'RIRO20160406 **********************

Public Function ObtenerGruposComision(ByVal nTipo As Integer) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    'nTipo=1 --> Tasas
    'nTipo=2 --> Comisiones
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_TarifarioTasaComisionGrupos " & nTipo
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    Set ObtenerGruposComision = rs
    Set rs = Nothing
End Function
Public Function AgregaComisionVersion(poComision As tComision) As Boolean

    On Error GoTo Error

    Dim Cmd As ADODB.Command
    Dim Prm As ADODB.Parameter
    Set Cmd = New ADODB.Command
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nProducto", adInteger, adParamInput, , poComision.Producto)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSubProducto", adInteger, adParamInput, , poComision.SubProducto)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nPersoneria", adInteger, adParamInput, , poComision.Personeria)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVersion", adInteger, adParamInput, , poComision.Version)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFechaRegistro", adDate, adParamInput, , poComision.FechaRegistro)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cGlosa", adVarChar, adParamInput, 100, poComision.Glosa)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nEstado", adInteger, adParamInput, , poComision.Estado)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFechaVigencia", adDate, adParamInput, , poComision.FechaVigencia)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cGrupo", adVarChar, adParamInput, 1, poComision.Grupo)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cMovNro", adVarChar, adParamInput, 25, poComision.MovNro)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaDepositoSolesTipo", adInteger, adParamInput, , poComision.CtaDepositoSolesTipo)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaDepositoSolesComision", adDouble, adParamInput, , poComision.CtaDepositoSolesComision)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaDepositoSolesMin", adDouble, adParamInput, , poComision.CtaDepositoSolesMin)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaDepositoSolesMax", adDouble, adParamInput, , poComision.CtaDepositoSolesMax)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaDepositoDolaresTipo", adInteger, adParamInput, , poComision.CtaDepositoDolaresTipo)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaDepositoDolaresComision", adDouble, adParamInput, , poComision.CtaDepositoDolaresComision)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaDepositoDolaresMin", adDouble, adParamInput, , poComision.CtaDepositoDolaresMin)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaDepositoDolaresMax", adDouble, adParamInput, , poComision.CtaDepositoDolaresMax)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaRetiroSolesTipo", adInteger, adParamInput, , poComision.CtaRetiroSolesTipo)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaRetiroSolesComision", adDouble, adParamInput, , poComision.CtaRetiroSolesComision)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaRetiroSolesMin", adDouble, adParamInput, , poComision.CtaRetiroSolesMin)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaRetiroSolesMax", adDouble, adParamInput, , poComision.CtaRetiroSolesMax)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaRetiroDolaresTipo", adInteger, adParamInput, , poComision.CtaRetiroDolaresTipo)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaRetiroDolaresComision", adDouble, adParamInput, , poComision.CtaRetiroDolaresComision)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaRetiroDolaresMin", adDouble, adParamInput, , poComision.CtaRetiroDolaresMin)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaRetiroDolaresMax", adDouble, adParamInput, , poComision.CtaRetiroDolaresMax)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaTransfSolesTipo", adInteger, adParamInput, , poComision.CtaTransfSolesTipo)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaTransfSolesComision", adDouble, adParamInput, , poComision.CtaTransfSolesComision)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaTransfSolesMin", adDouble, adParamInput, , poComision.CtaTransfSolesMin)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaTransfSolesMax", adDouble, adParamInput, , poComision.CtaTransfSolesMax)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaTransfDolaresTipo", adInteger, adParamInput, , poComision.CtaTransfDolaresTipo)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaTransfDolaresComision", adDouble, adParamInput, , poComision.CtaTransfDolaresComision)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaTransfDolaresMin", adDouble, adParamInput, , poComision.CtaTransfDolaresMin)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaTransfDolaresMax", adDouble, adParamInput, , poComision.CtaTransfDolaresMax)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajOperacionesLibresATM", adInteger, adParamInput, , poComision.CajOperacionesLibresATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajOperacionesLibresMON", adInteger, adParamInput, , poComision.CajOperacionesLibresMON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperacionesATM", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperacionesATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroDolaresExcesoOperacionesATM", adDouble, adParamInput, , poComision.CajRetiroDolaresExcesoOperacionesATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajConsultaSaldosSolesExcesoOperacionesATM", adDouble, adParamInput, , poComision.CajConsultaSaldosSolesExcesoOperacionesATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajConsultaSaldosDolaresExcesoOperacionesATM", adDouble, adParamInput, , poComision.CajConsultaSaldosDolaresExcesoOperacionesATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajConsultaMovSolesExcesoOperacionesATM", adDouble, adParamInput, , poComision.CajConsultaMovSolesExcesoOperacionesATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajConsultaMovDolaresExcesoOperacionesATM", adDouble, adParamInput, , poComision.CajConsultaMovDolaresExcesoOperacionesATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajCambioClaveSolesExcesoOperacionesATM", adDouble, adParamInput, , poComision.CajCambioClaveSolesExcesoOperacionesATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajCambioClaveDolaresExcesoOperacionesATM", adDouble, adParamInput, , poComision.CajCambioClaveDolaresExcesoOperacionesATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones05MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones05MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones10MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones10MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones15MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones15MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones20MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones20MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones25MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones25MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones30MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones30MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones35MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones35MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones40MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones40MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones45MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones45MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesOTRO", adDouble, adParamInput, , poComision.CajRetiroSolesOTRO)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroDolaresOTRO", adDouble, adParamInput, , poComision.CajRetiroDolaresOTRO)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajConsultaSaldosSolesOTRO", adDouble, adParamInput, , poComision.CajConsultaSaldosSolesOTRO)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajConsultaSaldosDolaresOTRO", adDouble, adParamInput, , poComision.CajConsultaSaldosDolaresOTRO)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesATMextranjero", adDouble, adParamInput, , poComision.CajRetiroSolesATMextranjero)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroDolaresATMextranjero", adDouble, adParamInput, , poComision.CajRetiroDolaresATMextranjero)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajConsultaSaldosSolesATMextranjero", adDouble, adParamInput, , poComision.CajConsultaSaldosSolesATMextranjero)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajConsultaSaldosDolaresATMextranjero", adDouble, adParamInput, , poComision.CajConsultaSaldosDolaresATMextranjero)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajCambioClaveSolesATMextranjero", adDouble, adParamInput, , poComision.CajCambioClaveSolesATMextranjero)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajCambioClaveDolaresATMextranjero", adDouble, adParamInput, , poComision.CajCambioClaveDolaresATMextranjero)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajCompraSolesNacionalVISA", adDouble, adParamInput, , poComision.CajCompraSolesNacionalVISA)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajCompraDolaresNacionalVISA", adDouble, adParamInput, , poComision.CajCompraDolaresNacionalVISA)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajCompraSolesInternacionalVISA", adDouble, adParamInput, , poComision.CajCompraSolesInternacionalVISA)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajCompraDolaresInternacionalVISA", adDouble, adParamInput, , poComision.CajCompraDolaresInternacionalVISA)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntOperacionesLibresRetiro", adInteger, adParamInput, , poComision.VntOperacionesLibresRetiro)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntOperacionesLibresDeposito", adInteger, adParamInput, , poComision.VntOperacionesLibresDeposito)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntOperacionesLibresTransf", adInteger, adParamInput, , poComision.VntOperacionesLibresTransf)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntOperacionesLibresOrdenesPago", adInteger, adParamInput, , poComision.VntOperacionesLibresOrdenesPago)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntOperacionesLibresInterplazaRetiro", adInteger, adParamInput, , poComision.VntOperacionesLibresInterplazaRetiro)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntOperacionesLibresInterplazaDeposito", adInteger, adParamInput, , poComision.VntOperacionesLibresInterplazaDeposito)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntOperacionesLibresInterplazaTransf", adInteger, adParamInput, , poComision.VntOperacionesLibresInterplazaTransf)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntExcesoOperacionesOrdenesPagoSoles", adDouble, adParamInput, , poComision.VntExcesoOperacionesOrdenesPagoSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntExcesoOperacionesOrdenesPagoDolares", adDouble, adParamInput, , poComision.VntExcesoOperacionesOrdenesPagoDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntRetiroSolesExcesoOperaciones", adDouble, adParamInput, , poComision.VntRetiroSolesExcesoOperaciones)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntRetiroDolaresExcesoOperaciones", adDouble, adParamInput, , poComision.VntRetiroDolaresExcesoOperaciones)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntDepositoSolesExcesoOperaciones", adDouble, adParamInput, , poComision.VntDepositoSolesExcesoOperaciones)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntDepositoDolaresExcesoOperaciones", adDouble, adParamInput, , poComision.VntDepositoDolaresExcesoOperaciones)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntTranfSolesExcesoOperaciones", adDouble, adParamInput, , poComision.VntTranfSolesExcesoOperaciones)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntTranfDolaresExcesoOperaciones", adDouble, adParamInput, , poComision.VntTranfDolaresExcesoOperaciones)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntExtractoCtaAhorrosSoles", adDouble, adParamInput, , poComision.VntExtractoCtaAhorrosSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntExtractoCtaAhorrosDolares", adDouble, adParamInput, , poComision.VntExtractoCtaAhorrosDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntExtractoCtaAhorrosValor", adInteger, adParamInput, , poComision.VntExtractoCtaAhorrosValor)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntConsultaSaldosVentanillaSoles", adDouble, adParamInput, , poComision.VntConsultaSaldosVentanillaSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntConsultaSaldosVentanillaDolares", adDouble, adParamInput, , poComision.VntConsultaSaldosVentanillaDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntConsultaSaldosVentanillaValor", adInteger, adParamInput, , poComision.VntConsultaSaldosVentanillaValor)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntRetiroSinTarjetaDebitoSoles", adDouble, adParamInput, , poComision.VntRetiroSinTarjetaDebitoSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntRetiroSinTarjetaDebitoDolares", adDouble, adParamInput, , poComision.VntRetiroSinTarjetaDebitoDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntRetiroSinTarjetaDebitoValor", adInteger, adParamInput, , poComision.VntRetiroSinTarjetaDebitoValor)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvMantenimientoCtaSoles", adDouble, adParamInput, , poComision.SrvMantenimientoCtaSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvMantenimientoCtaDolares", adDouble, adParamInput, , poComision.SrvMantenimientoCtaDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvMantenimientoCtaValor", adInteger, adParamInput, , poComision.SrvMantenimientoCtaValor)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvDebitoServicioRecaudoSoles", adDouble, adParamInput, , poComision.SrvDebitoServicioRecaudoSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvDebitoServicioRecaudoDolares", adDouble, adParamInput, , poComision.SrvDebitoServicioRecaudoDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvDebitoServicioRecaudoValor", adInteger, adParamInput, , poComision.SrvDebitoServicioRecaudoValor)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvDebitoPagoCreditoSoles", adDouble, adParamInput, , poComision.SrvDebitoPagoCreditoSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvDebitoPagoCreditoDolares", adDouble, adParamInput, , poComision.SrvDebitoPagoCreditoDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvDebitoPagoCreditoValor", adInteger, adParamInput, , poComision.SrvDebitoPagoCreditoValor)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvEnvioEstadoCtaFisicoSoles", adDouble, adParamInput, , poComision.SrvEnvioEstadoCtaFisicoSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvEnvioEstadoCtaFisicoDolares", adDouble, adParamInput, , poComision.SrvEnvioEstadoCtaFisicoDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvEnvioEstadoCtaFisicoValor", adInteger, adParamInput, , poComision.SrvEnvioEstadoCtaFisicoValor)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvReposicionTarjetaDebitoSoles", adDouble, adParamInput, , poComision.SrvReposicionTarjetaDebitoSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvReposicionTarjetaDebitoDolares", adDouble, adParamInput, , poComision.SrvReposicionTarjetaDebitoDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvReposicionTarjetaDebitoValor", adInteger, adParamInput, , poComision.SrvReposicionTarjetaDebitoValor)
    Cmd.Parameters.Append Prm
 
    Cmd.Parameters.Refresh
    Cmd.ActiveConnection = dbCmact
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "stp_ins_TarifarioComisionAdd"
    Cmd.Execute

    Set Cmd = Nothing
    Set Prm = Nothing
    
    AgregaComisionVersion = True
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:AgregaComisionVersion", Err.Description
    AgregaComisionVersion = False
        
    Set Cmd = Nothing
    Set Prm = Nothing
        
End Function
Public Function ActualizaComisionVersion(poComision As tComision) As Boolean
                                                           
    On Error GoTo Error
      
    Dim Cmd As ADODB.Command
    Dim Prm As ADODB.Parameter
    Set Cmd = New ADODB.Command
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nIdComision", adInteger, adParamInput, , poComision.IdComision)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nProducto", adInteger, adParamInput, , poComision.Producto)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSubProducto", adInteger, adParamInput, , poComision.SubProducto)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nPersoneria", adInteger, adParamInput, , poComision.Personeria)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVersion", adInteger, adParamInput, , poComision.Version)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFechaRegistro", adDate, adParamInput, , poComision.FechaRegistro)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cGlosa", adVarChar, adParamInput, 100, poComision.Glosa)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nEstado", adInteger, adParamInput, , poComision.Estado)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFechaVigencia", adDate, adParamInput, , poComision.FechaVigencia)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cGrupo", adVarChar, adParamInput, 1, poComision.Grupo)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cMovNro", adVarChar, adParamInput, 25, poComision.MovNro)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaDepositoSolesTipo", adInteger, adParamInput, , poComision.CtaDepositoSolesTipo)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaDepositoSolesComision", adDouble, adParamInput, , poComision.CtaDepositoSolesComision)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaDepositoSolesMin", adDouble, adParamInput, , poComision.CtaDepositoSolesMin)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaDepositoSolesMax", adDouble, adParamInput, , poComision.CtaDepositoSolesMax)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaDepositoDolaresTipo", adInteger, adParamInput, , poComision.CtaDepositoDolaresTipo)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaDepositoDolaresComision", adDouble, adParamInput, , poComision.CtaDepositoDolaresComision)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaDepositoDolaresMin", adDouble, adParamInput, , poComision.CtaDepositoDolaresMin)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaDepositoDolaresMax", adDouble, adParamInput, , poComision.CtaDepositoDolaresMax)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaRetiroSolesTipo", adInteger, adParamInput, , poComision.CtaRetiroSolesTipo)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaRetiroSolesComision", adDouble, adParamInput, , poComision.CtaRetiroSolesComision)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaRetiroSolesMin", adDouble, adParamInput, , poComision.CtaRetiroSolesMin)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaRetiroSolesMax", adDouble, adParamInput, , poComision.CtaRetiroSolesMax)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaRetiroDolaresTipo", adInteger, adParamInput, , poComision.CtaRetiroDolaresTipo)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaRetiroDolaresComision", adDouble, adParamInput, , poComision.CtaRetiroDolaresComision)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaRetiroDolaresMin", adDouble, adParamInput, , poComision.CtaRetiroDolaresMin)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaRetiroDolaresMax", adDouble, adParamInput, , poComision.CtaRetiroDolaresMax)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaTransfSolesTipo", adInteger, adParamInput, , poComision.CtaTransfSolesTipo)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaTransfSolesComision", adDouble, adParamInput, , poComision.CtaTransfSolesComision)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaTransfSolesMin", adDouble, adParamInput, , poComision.CtaTransfSolesMin)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaTransfSolesMax", adDouble, adParamInput, , poComision.CtaTransfSolesMax)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaTransfDolaresTipo", adInteger, adParamInput, , poComision.CtaTransfDolaresTipo)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaTransfDolaresComision", adDouble, adParamInput, , poComision.CtaTransfDolaresComision)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaTransfDolaresMin", adDouble, adParamInput, , poComision.CtaTransfDolaresMin)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCtaTransfDolaresMax", adDouble, adParamInput, , poComision.CtaTransfDolaresMax)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajOperacionesLibresATM", adInteger, adParamInput, , poComision.CajOperacionesLibresATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajOperacionesLibresMON", adInteger, adParamInput, , poComision.CajOperacionesLibresMON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperacionesATM", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperacionesATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroDolaresExcesoOperacionesATM", adDouble, adParamInput, , poComision.CajRetiroDolaresExcesoOperacionesATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajConsultaSaldosSolesExcesoOperacionesATM", adDouble, adParamInput, , poComision.CajConsultaSaldosSolesExcesoOperacionesATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajConsultaSaldosDolaresExcesoOperacionesATM", adDouble, adParamInput, , poComision.CajConsultaSaldosDolaresExcesoOperacionesATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajConsultaMovSolesExcesoOperacionesATM", adDouble, adParamInput, , poComision.CajConsultaMovSolesExcesoOperacionesATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajConsultaMovDolaresExcesoOperacionesATM", adDouble, adParamInput, , poComision.CajConsultaMovDolaresExcesoOperacionesATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajCambioClaveSolesExcesoOperacionesATM", adDouble, adParamInput, , poComision.CajCambioClaveSolesExcesoOperacionesATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajCambioClaveDolaresExcesoOperacionesATM", adDouble, adParamInput, , poComision.CajCambioClaveDolaresExcesoOperacionesATM)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones05MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones05MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones10MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones10MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones15MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones15MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones20MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones20MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones25MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones25MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones30MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones30MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones35MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones35MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones40MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones40MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesExcesoOperaciones45MON", adDouble, adParamInput, , poComision.CajRetiroSolesExcesoOperaciones45MON)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesOTRO", adDouble, adParamInput, , poComision.CajRetiroSolesOTRO)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroDolaresOTRO", adDouble, adParamInput, , poComision.CajRetiroDolaresOTRO)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajConsultaSaldosSolesOTRO", adDouble, adParamInput, , poComision.CajConsultaSaldosSolesOTRO)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajConsultaSaldosDolaresOTRO", adDouble, adParamInput, , poComision.CajConsultaSaldosDolaresOTRO)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroSolesATMextranjero", adDouble, adParamInput, , poComision.CajRetiroSolesATMextranjero)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajRetiroDolaresATMextranjero", adDouble, adParamInput, , poComision.CajRetiroDolaresATMextranjero)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajConsultaSaldosSolesATMextranjero", adDouble, adParamInput, , poComision.CajConsultaSaldosSolesATMextranjero)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajConsultaSaldosDolaresATMextranjero", adDouble, adParamInput, , poComision.CajConsultaSaldosDolaresATMextranjero)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajCambioClaveSolesATMextranjero", adDouble, adParamInput, , poComision.CajCambioClaveSolesATMextranjero)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajCambioClaveDolaresATMextranjero", adDouble, adParamInput, , poComision.CajCambioClaveDolaresATMextranjero)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajCompraSolesNacionalVISA", adDouble, adParamInput, , poComision.CajCompraSolesNacionalVISA)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajCompraDolaresNacionalVISA", adDouble, adParamInput, , poComision.CajCompraDolaresNacionalVISA)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajCompraSolesInternacionalVISA", adDouble, adParamInput, , poComision.CajCompraSolesInternacionalVISA)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCajCompraDolaresInternacionalVISA", adDouble, adParamInput, , poComision.CajCompraDolaresInternacionalVISA)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntOperacionesLibresRetiro", adInteger, adParamInput, , poComision.VntOperacionesLibresRetiro)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntOperacionesLibresDeposito", adInteger, adParamInput, , poComision.VntOperacionesLibresDeposito)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntOperacionesLibresTransf", adInteger, adParamInput, , poComision.VntOperacionesLibresTransf)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntOperacionesLibresOrdenesPago", adInteger, adParamInput, , poComision.VntOperacionesLibresOrdenesPago)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntOperacionesLibresInterplazaRetiro", adInteger, adParamInput, , poComision.VntOperacionesLibresInterplazaRetiro)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntOperacionesLibresInterplazaDeposito", adInteger, adParamInput, , poComision.VntOperacionesLibresInterplazaDeposito)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntOperacionesLibresInterplazaTransf", adInteger, adParamInput, , poComision.VntOperacionesLibresInterplazaTransf)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntExcesoOperacionesOrdenesPagoSoles", adDouble, adParamInput, , poComision.VntExcesoOperacionesOrdenesPagoSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntExcesoOperacionesOrdenesPagoDolares", adDouble, adParamInput, , poComision.VntExcesoOperacionesOrdenesPagoDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntRetiroSolesExcesoOperaciones", adDouble, adParamInput, , poComision.VntRetiroSolesExcesoOperaciones)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntRetiroDolaresExcesoOperaciones", adDouble, adParamInput, , poComision.VntRetiroDolaresExcesoOperaciones)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntDepositoSolesExcesoOperaciones", adDouble, adParamInput, , poComision.VntDepositoSolesExcesoOperaciones)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntDepositoDolaresExcesoOperaciones", adDouble, adParamInput, , poComision.VntDepositoDolaresExcesoOperaciones)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntTranfSolesExcesoOperaciones", adDouble, adParamInput, , poComision.VntTranfSolesExcesoOperaciones)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntTranfDolaresExcesoOperaciones", adDouble, adParamInput, , poComision.VntTranfDolaresExcesoOperaciones)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntExtractoCtaAhorrosSoles", adDouble, adParamInput, , poComision.VntExtractoCtaAhorrosSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntExtractoCtaAhorrosDolares", adDouble, adParamInput, , poComision.VntExtractoCtaAhorrosDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntExtractoCtaAhorrosValor", adInteger, adParamInput, , poComision.VntExtractoCtaAhorrosValor)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntConsultaSaldosVentanillaSoles", adDouble, adParamInput, , poComision.VntConsultaSaldosVentanillaSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntConsultaSaldosVentanillaDolares", adDouble, adParamInput, , poComision.VntConsultaSaldosVentanillaDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntConsultaSaldosVentanillaValor", adInteger, adParamInput, , poComision.VntConsultaSaldosVentanillaValor)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntRetiroSinTarjetaDebitoSoles", adDouble, adParamInput, , poComision.VntRetiroSinTarjetaDebitoSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntRetiroSinTarjetaDebitoDolares", adDouble, adParamInput, , poComision.VntRetiroSinTarjetaDebitoDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nVntRetiroSinTarjetaDebitoValor", adInteger, adParamInput, , poComision.VntRetiroSinTarjetaDebitoValor)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvMantenimientoCtaSoles", adDouble, adParamInput, , poComision.SrvMantenimientoCtaSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvMantenimientoCtaDolares", adDouble, adParamInput, , poComision.SrvMantenimientoCtaDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvMantenimientoCtaValor", adInteger, adParamInput, , poComision.SrvMantenimientoCtaValor)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvDebitoServicioRecaudoSoles", adDouble, adParamInput, , poComision.SrvDebitoServicioRecaudoSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvDebitoServicioRecaudoDolares", adDouble, adParamInput, , poComision.SrvDebitoServicioRecaudoDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvDebitoServicioRecaudoValor", adInteger, adParamInput, , poComision.SrvDebitoServicioRecaudoValor)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvDebitoPagoCreditoSoles", adDouble, adParamInput, , poComision.SrvDebitoPagoCreditoSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvDebitoPagoCreditoDolares", adDouble, adParamInput, , poComision.SrvDebitoPagoCreditoDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvDebitoPagoCreditoValor", adInteger, adParamInput, , poComision.SrvDebitoPagoCreditoValor)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvEnvioEstadoCtaFisicoSoles", adDouble, adParamInput, , poComision.SrvEnvioEstadoCtaFisicoSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvEnvioEstadoCtaFisicoDolares", adDouble, adParamInput, , poComision.SrvEnvioEstadoCtaFisicoDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvEnvioEstadoCtaFisicoValor", adInteger, adParamInput, , poComision.SrvEnvioEstadoCtaFisicoValor)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvReposicionTarjetaDebitoSoles", adDouble, adParamInput, , poComision.SrvReposicionTarjetaDebitoSoles)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvReposicionTarjetaDebitoDolares", adDouble, adParamInput, , poComision.SrvReposicionTarjetaDebitoDolares)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSrvReposicionTarjetaDebitoValor", adInteger, adParamInput, , poComision.SrvReposicionTarjetaDebitoValor)
    Cmd.Parameters.Append Prm
 
    Cmd.Parameters.Refresh
    Cmd.ActiveConnection = dbCmact
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "stp_upd_TarifarioComisionUp"
    Cmd.Execute

    Set Cmd = Nothing
    Set Prm = Nothing
    
    'ALPA20160427*******************
    'ActualizaComision = True
    ActualizaComisionVersion = True
    '******************************
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:ActualizaComision", Err.Description
    'ALPA20160427*******************
    'ActualizaComision = False
    ActualizaComisionVersion = False
    '******************************
        
    Set Cmd = Nothing
    Set Prm = Nothing
    
End Function
Public Function ObtenerUltimaVersionComision(ByVal sGrupo As String, _
                                             ByVal nProducto As Integer, _
                                             ByVal nSubProducto As Integer, _
                                             ByVal nPersoneria As Integer) As Integer

    On Error GoTo Error
    
    Dim rs As ADODB.Recordset
    Dim nVersion As Integer
    
    sSql = "exec stp_sel_TarifarioComisionUltimaVersion '" & sGrupo & "', " & nProducto & ", " & nSubProducto & ", " & nPersoneria
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            nVersion = rs!nUltimaVersion
        Else
            nVersion = -1
        End If
    Else
        nVersion = -1
    End If
    ObtenerUltimaVersionComision = nVersion
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:ObtenerUltimaVersionComision", Err.Description

End Function
Public Function ObtenerVersionesComision(ByVal sGrupo As String, _
                                         ByVal nProducto As Integer, _
                                         ByVal nSubProducto As Integer, _
                                         ByVal nPersoneria As Integer) As ADODB.Recordset

    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_TarifarioComisionExaminar '" & sGrupo & "'," & nProducto & ", " & nSubProducto & ", " & nPersoneria

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    Set ObtenerVersionesComision = rs
    Set rs = Nothing
End Function
Public Function ObtenerComision(ByVal pIdComision As Integer, ByRef poComision As tComision) As Boolean

    On Error GoTo Error
    
    Dim rs As ADODB.Recordset
    Dim bRespuesta As Boolean
    bRespuesta = False
    sSql = "exec stp_sel_comision " & pIdComision
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            poComision.IdComision = rs!nIDComision
            poComision.Producto = rs!nProducto
            poComision.SubProducto = rs!nSubProducto
            poComision.Personeria = rs!nPersoneria
            poComision.Version = rs!nVersion
            poComision.FechaRegistro = rs!dFechaRegistro
            poComision.Glosa = rs!cGlosa
            poComision.Estado = rs!nEstado
            poComision.FechaVigencia = rs!dFechaVigencia
            poComision.Grupo = rs!cGrupo
            poComision.MovNro = rs!cMovNro
            poComision.CtaDepositoSolesTipo = rs!nCtaDepositoSolesTipo
            poComision.CtaDepositoSolesComision = rs!nCtaDepositoSolesComision
            poComision.CtaDepositoSolesMin = rs!nCtaDepositoSolesMin
            poComision.CtaDepositoSolesMax = rs!nCtaDepositoSolesMax
            poComision.CtaDepositoDolaresTipo = rs!nCtaDepositoDolaresTipo
            poComision.CtaDepositoDolaresComision = rs!nCtaDepositoDolaresComision
            poComision.CtaDepositoDolaresMin = rs!nCtaDepositoDolaresMin
            poComision.CtaDepositoDolaresMax = rs!nCtaDepositoDolaresMax
            poComision.CtaRetiroSolesTipo = rs!nCtaRetiroSolesTipo
            poComision.CtaRetiroSolesComision = rs!nCtaRetiroSolesComision
            poComision.CtaRetiroSolesMin = rs!nCtaRetiroSolesMin
            poComision.CtaRetiroSolesMax = rs!nCtaRetiroSolesMax
            poComision.CtaRetiroDolaresTipo = rs!nCtaRetiroDolaresTipo
            poComision.CtaRetiroDolaresComision = rs!nCtaRetiroDolaresComision
            poComision.CtaRetiroDolaresMin = rs!nCtaRetiroDolaresMin
            poComision.CtaRetiroDolaresMax = rs!nCtaRetiroDolaresMax
            poComision.CtaTransfSolesTipo = rs!nCtaTransfSolesTipo
            poComision.CtaTransfSolesComision = rs!nCtaTransfSolesComision
            poComision.CtaTransfSolesMin = rs!nCtaTransfSolesMin
            poComision.CtaTransfSolesMax = rs!nCtaTransfSolesMax
            poComision.CtaTransfDolaresTipo = rs!nCtaTransfDolaresTipo
            poComision.CtaTransfDolaresComision = rs!nCtaTransfDolaresComision
            poComision.CtaTransfDolaresMin = rs!nCtaTransfDolaresMin
            poComision.CtaTransfDolaresMax = rs!nCtaTransfDolaresMax
            poComision.CajOperacionesLibresATM = rs!nCajOperacionesLibresATM
            poComision.CajOperacionesLibresMON = rs!nCajOperacionesLibresMON
            poComision.CajRetiroSolesExcesoOperacionesATM = rs!nCajRetiroSolesExcesoOperacionesATM
            poComision.CajRetiroDolaresExcesoOperacionesATM = rs!nCajRetiroDolaresExcesoOperacionesATM
            poComision.CajConsultaSaldosSolesExcesoOperacionesATM = rs!nCajConsultaSaldosSolesExcesoOperacionesATM
            poComision.CajConsultaSaldosDolaresExcesoOperacionesATM = rs!nCajConsultaSaldosDolaresExcesoOperacionesATM
            poComision.CajConsultaMovSolesExcesoOperacionesATM = rs!nCajConsultaMovSolesExcesoOperacionesATM
            poComision.CajConsultaMovDolaresExcesoOperacionesATM = rs!nCajConsultaMovDolaresExcesoOperacionesATM
            poComision.CajCambioClaveSolesExcesoOperacionesATM = rs!nCajCambioClaveSolesExcesoOperacionesATM
            poComision.CajCambioClaveDolaresExcesoOperacionesATM = rs!nCajCambioClaveDolaresExcesoOperacionesATM
            poComision.CajRetiroSolesExcesoOperaciones05MON = rs!nCajRetiroSolesExcesoOperaciones05MON
            poComision.CajRetiroSolesExcesoOperaciones10MON = rs!nCajRetiroSolesExcesoOperaciones10MON
            poComision.CajRetiroSolesExcesoOperaciones15MON = rs!nCajRetiroSolesExcesoOperaciones15MON
            poComision.CajRetiroSolesExcesoOperaciones20MON = rs!nCajRetiroSolesExcesoOperaciones20MON
            poComision.CajRetiroSolesExcesoOperaciones25MON = rs!nCajRetiroSolesExcesoOperaciones25MON
            poComision.CajRetiroSolesExcesoOperaciones30MON = rs!nCajRetiroSolesExcesoOperaciones30MON
            poComision.CajRetiroSolesExcesoOperaciones35MON = rs!nCajRetiroSolesExcesoOperaciones35MON
            poComision.CajRetiroSolesExcesoOperaciones40MON = rs!nCajRetiroSolesExcesoOperaciones40MON
            poComision.CajRetiroSolesExcesoOperaciones45MON = rs!nCajRetiroSolesExcesoOperaciones45MON
            poComision.CajRetiroSolesOTRO = rs!nCajRetiroSolesOTRO
            poComision.CajRetiroDolaresOTRO = rs!nCajRetiroDolaresOTRO
            poComision.CajConsultaSaldosSolesOTRO = rs!nCajConsultaSaldosSolesOTRO
            poComision.CajConsultaSaldosDolaresOTRO = rs!nCajConsultaSaldosDolaresOTRO
            poComision.CajRetiroSolesATMextranjero = rs!nCajRetiroSolesATMextranjero
            poComision.CajRetiroDolaresATMextranjero = rs!nCajRetiroDolaresATMextranjero
            poComision.CajConsultaSaldosSolesATMextranjero = rs!nCajConsultaSaldosSolesATMextranjero
            poComision.CajConsultaSaldosDolaresATMextranjero = rs!nCajConsultaSaldosDolaresATMextranjero
            poComision.CajCambioClaveSolesATMextranjero = rs!nCajCambioClaveSolesATMextranjero
            poComision.CajCambioClaveDolaresATMextranjero = rs!nCajCambioClaveDolaresATMextranjero
            poComision.CajCompraSolesNacionalVISA = rs!nCajCompraSolesNacionalVISA
            poComision.CajCompraDolaresNacionalVISA = rs!nCajCompraDolaresNacionalVISA
            poComision.CajCompraSolesInternacionalVISA = rs!nCajCompraSolesInternacionalVISA
            poComision.CajCompraDolaresInternacionalVISA = rs!nCajCompraDolaresInternacionalVISA
            poComision.VntOperacionesLibresRetiro = rs!nVntOperacionesLibresRetiro
            poComision.VntOperacionesLibresDeposito = rs!nVntOperacionesLibresDeposito
            poComision.VntOperacionesLibresTransf = rs!nVntOperacionesLibresTransf
            poComision.VntOperacionesLibresOrdenesPago = rs!nVntOperacionesLibresOrdenesPago
            poComision.VntOperacionesLibresInterplazaRetiro = rs!nVntOperacionesLibresInterplazaRetiro
            poComision.VntOperacionesLibresInterplazaDeposito = rs!nVntOperacionesLibresInterplazaDeposito
            poComision.VntOperacionesLibresInterplazaTransf = rs!nVntOperacionesLibresInterplazaTransf
            poComision.VntExcesoOperacionesOrdenesPagoSoles = rs!nVntExcesoOperacionesOrdenesPagoSoles
            poComision.VntExcesoOperacionesOrdenesPagoDolares = rs!nVntExcesoOperacionesOrdenesPagoDolares
            poComision.VntRetiroSolesExcesoOperaciones = rs!nVntRetiroSolesExcesoOperaciones
            poComision.VntRetiroDolaresExcesoOperaciones = rs!nVntRetiroDolaresExcesoOperaciones
            poComision.VntDepositoSolesExcesoOperaciones = rs!nVntDepositoSolesExcesoOperaciones
            poComision.VntDepositoDolaresExcesoOperaciones = rs!nVntDepositoDolaresExcesoOperaciones
            poComision.VntTranfSolesExcesoOperaciones = rs!nVntTranfSolesExcesoOperaciones
            poComision.VntTranfDolaresExcesoOperaciones = rs!nVntTranfDolaresExcesoOperaciones
            poComision.VntExtractoCtaAhorrosSoles = rs!nVntExtractoCtaAhorrosSoles
            poComision.VntExtractoCtaAhorrosDolares = rs!nVntExtractoCtaAhorrosDolares
            poComision.VntExtractoCtaAhorrosValor = rs!nVntExtractoCtaAhorrosValor
            poComision.VntConsultaSaldosVentanillaSoles = rs!nVntConsultaSaldosVentanillaSoles
            poComision.VntConsultaSaldosVentanillaDolares = rs!nVntConsultaSaldosVentanillaDolares
            poComision.VntConsultaSaldosVentanillaValor = rs!nVntConsultaSaldosVentanillaValor
            poComision.VntRetiroSinTarjetaDebitoSoles = rs!nVntRetiroSinTarjetaDebitoSoles
            poComision.VntRetiroSinTarjetaDebitoDolares = rs!nVntRetiroSinTarjetaDebitoDolares
            poComision.VntRetiroSinTarjetaDebitoValor = rs!nVntRetiroSinTarjetaDebitoValor
            poComision.SrvMantenimientoCtaSoles = rs!nSrvMantenimientoCtaSoles
            poComision.SrvMantenimientoCtaDolares = rs!nSrvMantenimientoCtaDolares
            poComision.SrvMantenimientoCtaValor = rs!nSrvMantenimientoCtaValor
            poComision.SrvDebitoServicioRecaudoSoles = rs!nSrvDebitoServicioRecaudoSoles
            poComision.SrvDebitoServicioRecaudoDolares = rs!nSrvDebitoServicioRecaudoDolares
            poComision.SrvDebitoServicioRecaudoValor = rs!nSrvDebitoServicioRecaudoValor
            poComision.SrvDebitoPagoCreditoSoles = rs!nSrvDebitoPagoCreditoSoles
            poComision.SrvDebitoPagoCreditoDolares = rs!nSrvDebitoPagoCreditoDolares
            poComision.SrvDebitoPagoCreditoValor = rs!nSrvDebitoPagoCreditoValor
            poComision.SrvEnvioEstadoCtaFisicoSoles = rs!nSrvEnvioEstadoCtaFisicoSoles
            poComision.SrvEnvioEstadoCtaFisicoDolares = rs!nSrvEnvioEstadoCtaFisicoDolares
            poComision.SrvEnvioEstadoCtaFisicoValor = rs!nSrvEnvioEstadoCtaFisicoValor
            poComision.SrvReposicionTarjetaDebitoSoles = rs!nSrvReposicionTarjetaDebitoSoles
            poComision.SrvReposicionTarjetaDebitoDolares = rs!nSrvReposicionTarjetaDebitoDolares
            poComision.SrvReposicionTarjetaDebitoValor = rs!nSrvReposicionTarjetaDebitoValor
            bRespuesta = True
        End If
    End If
    ObtenerComision = bRespuesta
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:ObtenerComision", Err.Description
End Function
Public Function ObtenerVersionesXtipo(ByVal pnTipo As Integer) As ADODB.Recordset

    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_TarifarioListaVersionesXtipo '" & pnTipo & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    Set ObtenerVersionesXtipo = rs
    Set rs = Nothing
End Function
'*-- Create By GITU 21-07-2016
Public Function ObtenerAgenciaGrupoTasaComisionXAgencia(ByVal psCodAge As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    sSql = "exec stp_sel_TarifarioObtenerGrupoTasaComisionXAgencia '" + psCodAge + "'"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    Set ObtenerAgenciaGrupoTasaComisionXAgencia = rs
    Set rs = Nothing
End Function
Public Function ObtenerAgenciaGrupoTasaComision() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    sSql = "exec stp_sel_TarifarioObtenerGrupoTasaComision "
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    Set ObtenerAgenciaGrupoTasaComision = rs
    Set rs = Nothing
End Function
Public Function ActualizaTarifarioGrupo(ByVal prsTarifGrupo As ADODB.Recordset, ByVal psUser As String, ByVal psAgeCod As String) As Boolean
On Error GoTo Error
    Dim sSql As String
    If Not prsTarifGrupo Is Nothing Then
        If Not prsTarifGrupo.EOF And Not prsTarifGrupo.BOF Then
            If prsTarifGrupo.RecordCount > 0 Then
                dbCmact.BeginTrans
                Do While Not prsTarifGrupo.EOF And Not prsTarifGrupo.BOF
                    sSql = "stp_ins_TarifarioGrupoAgenciaTasaComision " & _
                            prsTarifGrupo!nIdGrupo & ", '" & _
                            psUser & "', '" & _
                            prsTarifGrupo!cAgeCod & "', '" & _
                            prsTarifGrupo!Tasa & "', '" & _
                            prsTarifGrupo!Comision & "', '" & _
                            psUser & "'"
                    dbCmact.Execute sSql
                    prsTarifGrupo.MoveNext
                Loop
                dbCmact.CommitTrans
            End If
        End If
    End If
    ActualizaTarifarioGrupo = True
    Exit Function
Error:
    dbCmact.RollbackTrans
    ActualizaTarifarioGrupo = False
End Function
Public Function ObtenerUltimaVersionTasa(ByVal sGrupo As String, _
                                         ByVal nProducto As Integer, _
                                         ByVal nSubProducto As Integer, _
                                         ByVal nPersoneria As Integer) As Integer

    On Error GoTo Error
    
    Dim rs As ADODB.Recordset
    Dim nVersion As Integer
    
    sSql = "exec stp_sel_TarifarioTasaUltimaVersion '" & sGrupo & "', " & nProducto & ", " & nSubProducto & ", " & nPersoneria
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            nVersion = rs!nUltimaVersion
        Else
            nVersion = -1
        End If
    Else
        nVersion = -1
    End If
    ObtenerUltimaVersionTasa = nVersion
    Exit Function
Error:
    Err.Raise Err.Number, "NCapDefinicion:ObtenerUltimaVersionTasa", Err.Description

End Function
Public Function AgregaTasaVersion(poTasa As tTasa, _
                                  ByVal prsTasaSoles As ADODB.Recordset, _
                                  ByVal prsTasaDolares As ADODB.Recordset) As Boolean
On Error GoTo Error

    Dim sSql As String
    Dim nIdTasa As Integer
    Dim RSTMP As New ADODB.Recordset
    Dim dbCon As New COMConecta.DCOMConecta
            
    dbCmact.BeginTrans
        
    sSql = "stp_ins_TarifarioTasaCabAdd " & _
            poTasa.Producto & ", " & _
            poTasa.SubProducto & ", " & _
            poTasa.Personeria & ", '" & _
            poTasa.Grupo & "', '" & _
            poTasa.Version & "', '" & _
            poTasa.MovNro & "', '" & _
            Format(poTasa.FechaRegistro, "yyyyMMdd") & "', '" & _
            poTasa.Glosa & "', " & _
            poTasa.Estado & ", " & _
            poTasa.AplicaInmediato

    Set RSTMP = dbCmact.Execute(sSql)

    'insertando la cabecera
    nIdTasa = -1
    If Not RSTMP Is Nothing Then
        If Not RSTMP.EOF Then
            nIdTasa = RSTMP!TarifarioTasaCab
        End If
    End If
    If nIdTasa = -1 Then
        dbCmact.RollbackTrans
        AgregaTasaVersion = False
        Exit Function
    End If
    'insertando las tasas en Soles
    If Not prsTasaSoles Is Nothing Then
        If Not prsTasaSoles.EOF And Not prsTasaSoles.BOF Then
            If prsTasaSoles.RecordCount > 0 Then
                Do While Not prsTasaSoles.EOF And Not prsTasaSoles.BOF
                    sSql = "stp_ins_TarifarioTasaDetAdd " & _
                            nIdTasa & ", " & _
                            prsTasaSoles("Monto Ini") & ", " & _
                            prsTasaSoles("Monto Fin") & ", " & _
                            prsTasaSoles("Plazo Ini") & ", " & _
                            prsTasaSoles("Plazo Fin") & ", " & _
                            prsTasaSoles("Tasa Interes") & ", " & _
                            prsTasaSoles("Tasa Ord. Pago") & ", " & _
                            1 & ", " & _
                            1
                    dbCmact.Execute (sSql)
                    prsTasaSoles.MoveNext
                Loop
            End If
        End If
    End If
    'insertando las tasas en Dolares
    If Not prsTasaDolares Is Nothing Then
        If Not prsTasaDolares.EOF And Not prsTasaDolares.BOF Then
            If prsTasaDolares.RecordCount > 0 Then
                Do While Not prsTasaDolares.EOF And Not prsTasaDolares.BOF
                    sSql = "stp_ins_TarifarioTasaDetAdd " & _
                            nIdTasa & ", " & _
                            prsTasaDolares("Monto Ini") & ", " & _
                            prsTasaDolares("Monto Fin") & ", " & _
                            prsTasaDolares("Plazo Ini") & ", " & _
                            prsTasaDolares("Plazo Fin") & ", " & _
                            prsTasaDolares("Tasa Interes") & ", " & _
                            prsTasaDolares("Tasa Ord. Pago") & ", " & _
                            2 & ", " & _
                            1
                    dbCmact.Execute (sSql)
                    prsTasaDolares.MoveNext
                Loop
            End If
        End If
    End If
    dbCmact.CommitTrans
    
    AgregaTasaVersion = True
    Exit Function
Error:
    dbCmact.RollbackTrans
    AgregaTasaVersion = False
End Function
Public Function ObtenerVersionesTasa(ByVal sGrupo As String, _
                                     ByVal nProducto As Integer, _
                                     ByVal nSubProducto As Integer, _
                                     ByVal nPersoneria As Integer) As ADODB.Recordset

    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_TarifarioTasaExaminar '" & sGrupo & "'," & nProducto & ", " & nSubProducto & ", " & nPersoneria

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    Set ObtenerVersionesTasa = rs
    Set rs = Nothing
End Function
Public Function ObtenerTasaVersion(ByRef oTasa As tTasa) As Boolean

    On Error GoTo Error
    
    Dim rs As ADODB.Recordset
    Dim bRespuesta As Boolean
    bRespuesta = False
    sSql = "exec stp_sel_TarifarioTasa " & oTasa.IdTasa
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            If rs.RecordCount > 0 Then
                With oTasa
                    'cabecera de la version
                    .Producto = rs!nProducto
                    .SubProducto = rs!nSubProducto
                    .Personeria = rs!nPersoneria
                    .Grupo = rs!cGrupo
                    .Version = rs!nVersion
                    .FechaRegistro = rs!dFechaRegistro
                    .AplicaInmediato = rs!nAplicaInmediato
                    .Glosa = rs!cGlosa
                    .MovNro = rs!cMovNro
                    'Detalle
                    Set .rsTasa = rs
                End With
                bRespuesta = True
            End If
        End If
    End If
    ObtenerTasaVersion = bRespuesta
    Exit Function
Error:
    ObtenerTasaVersion = False
End Function
Public Function ActualizaTasaVersion(poTasa As tTasa, _
                                     ByVal prsSol As ADODB.Recordset, _
                                     ByVal prsDol As ADODB.Recordset, _
                                     ByVal plstElim As Variant) As Boolean
On Error GoTo Error

    Dim sSql As String
    Dim nIdTasa As Integer
    Dim RSTMP As New ADODB.Recordset
    Dim dbCon As New COMConecta.DCOMConecta
    Dim R As Variant

    dbCmact.BeginTrans

    'actualizando la cabecera de tasa versionada
    sSql = "stp_upd_TarifarioTasaCabUpd " & _
            poTasa.IdTasa & ", '" & _
            poTasa.MovNro & "', '" & _
            poTasa.Glosa & "', " & _
            poTasa.AplicaInmediato

    dbCmact.Execute (sSql)
    'actualizando el detalle en soles de las tasas versionadas
    If Not prsSol Is Nothing Then
        If Not prsSol.EOF And Not prsSol.BOF Then
            If prsSol.RecordCount > 0 Then
                Do While Not prsSol.EOF And Not prsSol.BOF
                    If prsSol!nTipo = 1 Then 'Nuevo
                        sSql = "stp_ins_TarifarioTasaDetAdd " & _
                                prsSol!nIdTarifarioTasaCab & ", " & _
                                prsSol!nMontoIni & ", " & _
                                prsSol!nMontoFin & ", " & _
                                prsSol!nPlazoIni & ", " & _
                                prsSol!nPlazoFin & ", " & _
                                prsSol!ntasainteres & ", " & _
                                prsSol!nTasaOrdenPago & ", " & _
                                prsSol!nMoneda & ", " & _
                                prsSol!nEstado
                    ElseIf prsSol!nTipo = 2 Then ' Edición
                        sSql = "stp_upd_TarifarioTasaDetUpd " & _
                                prsSol!nIdTarifarioTasaDet & ", " & _
                                prsSol!nMontoIni & ", " & _
                                prsSol!nMontoFin & ", " & _
                                prsSol!nPlazoIni & ", " & _
                                prsSol!nPlazoFin & ", " & _
                                prsSol!ntasainteres & ", " & _
                                prsSol!nTasaOrdenPago & ", " & _
                                prsSol!nMoneda & ", " & _
                                prsSol!nEstado
                    End If
                    dbCmact.Execute (sSql)
                    prsSol.MoveNext
                Loop
            End If
        End If
    End If
    'actualizando el detalle en dolares de las tasas versionadas
    If Not prsDol Is Nothing Then
        If Not prsDol.EOF And Not prsDol.BOF Then
            If prsDol.RecordCount > 0 Then
                Do While Not prsDol.EOF And Not prsDol.BOF
                    If prsDol!nTipo = 1 Then 'Nuevo
                        sSql = "stp_ins_TarifarioTasaDetAdd " & _
                                prsDol!nIdTarifarioTasaCab & ", " & _
                                prsDol!nMontoIni & ", " & _
                                prsDol!nMontoFin & ", " & _
                                prsDol!nPlazoIni & ", " & _
                                prsDol!nPlazoFin & ", " & _
                                prsDol!ntasainteres & ", " & _
                                prsDol!nTasaOrdenPago & ", " & _
                                prsDol!nMoneda & ", " & _
                                prsDol!nEstado
                    ElseIf prsDol!nTipo = 2 Then ' Edición
                        sSql = "stp_upd_TarifarioTasaDetUpd " & _
                                prsDol!nIdTarifarioTasaDet & ", " & _
                                prsDol!nMontoIni & ", " & _
                                prsDol!nMontoFin & ", " & _
                                prsDol!nPlazoIni & ", " & _
                                prsDol!nPlazoFin & ", " & _
                                prsDol!ntasainteres & ", " & _
                                prsDol!nTasaOrdenPago & ", " & _
                                prsDol!nMoneda & ", " & _
                                prsDol!nEstado
                    End If
                    dbCmact.Execute (sSql)
                    prsDol.MoveNext
                Loop
            End If
        End If
    End If
    'actualizando las tasas eliminadas
    If UBound(plstElim) > 0 Then
        For Each R In plstElim
            sSql = "stp_upd_TarifarioTasaDetEliminar " & R
            dbCmact.Execute (sSql)
        Next
    End If
    dbCmact.CommitTrans
    
    ActualizaTasaVersion = True
    Exit Function
Error:
    dbCmact.RollbackTrans
    ActualizaTasaVersion = False
End Function
Public Function ObtieneProgramacionXfechaTipo(ByVal pnTipo As Integer, _
                                              ByVal pdFechaVigencia As Date, _
                                              ByRef pRs As ADODB.Recordset) As Boolean
On Error GoTo Error

    Dim sSql As String
    Dim bRespuesta As Boolean
    
    sSql = "STP_SEL_TARIFARIOPROGRAMACIONXFECHA " & pnTipo & ", '" & Format(pdFechaVigencia, "yyyymmdd") & "'"
    Set pRs = dbCmact.Execute(sSql)
    
    ObtieneProgramacionXfechaTipo = False
    If Not pRs Is Nothing Then
        If Not pRs.EOF And Not pRs.BOF Then
                ObtieneProgramacionXfechaTipo = True
        End If
    End If
    
    Exit Function
Error:
    dbCmact.RollbackTrans
    ObtieneProgramacionXfechaTipo = False
End Function
Public Function RegistraProgramacion(ByVal pdFecha As Date, _
                                     ByVal psMovNro As String, _
                                     ByVal pnEstado As Integer, _
                                     ByVal plstTasCom As Variant, _
                                     ByVal pnTipo As Integer) As Boolean
On Error GoTo Error

    Dim sSql As String
    Dim bRes As Boolean
    Dim nIdProgramacion As Integer
    Dim rs As ADODB.Recordset
    Dim l As Variant
    Dim i As Integer
    dbCmact.BeginTrans
    
    sSql = "STP_INS_TARIFARIOPROGRAMACIONADD '" & Format(pdFecha, "yyyymmdd") & "', '" & psMovNro & "', " & pnEstado
    Set rs = dbCmact.Execute(sSql)

    If Not rs Is Nothing Then
        If Not rs.EOF And Not rs.BOF Then
            nIdProgramacion = rs!nIdProgramacion
            bRes = True
        End If
    End If
    If Not bRes Then
        dbCmact.RollbackTrans
        Set rs = Nothing
    Else
        For i = 1 To UBound(plstTasCom)
            sSql = "STP_INS_TARIFARIOPROGRAMACIONDETADD " & nIdProgramacion & ", " & plstTasCom(i) & ", " & pnTipo
            dbCmact.Execute sSql
        Next i
    End If
    
    dbCmact.CommitTrans
    RegistraProgramacion = True
    Exit Function
Error:
    dbCmact.RollbackTrans
    RegistraProgramacion = False
End Function

'END RIRO **************************
'JUEZ 20160415 *************************************************************************
Public Function GetCaptacCampanas(Optional ByVal pnCampanaCod As Long = 0, Optional ByVal pnProd As Producto = 0, _
                                  Optional ByVal pnSubProd As Integer = -1, Optional ByVal pnMoneda As Moneda = 0) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_CaptacCampanas " & pnCampanaCod & "," & pnProd & "," & pnSubProd & "," & pnMoneda

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    Set GetCaptacCampanas = rs
    Set rs = Nothing
End Function
Public Function InsertaCaptacCampanas(ByVal psCampanaDesc As String, ByVal pnMoneda As Moneda, ByVal pnProd As Producto, ByVal pnSubProd As Integer, _
                                      ByVal pnMontoMin As Double, ByVal pnPeriodoVigMeses As Integer, pdFechaIni As Date, pdFechaFin As Date, _
                                      ByVal pbPersJur As Boolean, ByVal psMovNro As String, ByVal pMatAgencias As Variant, ByVal pbDineroNuevo As Boolean) As Integer
'APRI20210621 ERS031-2021 Add pbDineroNuevo
Dim rs As ADODB.Recordset
Dim nCampanaCod As Integer
Dim i As Integer

On Error GoTo Error
    InsertaCaptacCampanas = 0
    dbCmact.BeginTrans
    
    sSql = "stp_ins_CaptacCampanas '" & psCampanaDesc & "'," & pnMoneda & "," & pnProd & "," & pnSubProd & "," & pnMontoMin & "," & _
                                        pnPeriodoVigMeses & ",'" & Format(pdFechaIni, "yyyyMMdd") & "','" & Format(pdFechaFin, "yyyyMMdd") & "'," & _
                                        IIf(pbPersJur, 1, 0) & ",'" & psMovNro & "'," & IIf(pbDineroNuevo, 1, 0)
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    nCampanaCod = CInt(rs("nCampanaCod"))
    Set rs = Nothing
    
    If nCampanaCod <> 0 Then
        For i = 0 To UBound(pMatAgencias)
            InsertaCaptacCampanasAge nCampanaCod, pMatAgencias(i)
        Next i
    Else
        dbCmact.RollbackTrans
        Exit Function
    End If
    
    dbCmact.CommitTrans
    InsertaCaptacCampanas = nCampanaCod
    Exit Function
Error:
    dbCmact.RollbackTrans
    InsertaCaptacCampanas = 0
    Err.Raise Err.Number, "NCapDefinicion:InsertaCaptacCampanas", Err.Description
End Function
Public Function ActualizaCaptacCampanas(ByVal pnCampanaCod As Long, ByVal pcCampanaDesc As String, ByVal pnMoneda As Integer, ByVal pnProd As Integer, _
                                        ByVal pnSubProd As Integer, ByVal pnMontoMin As Double, ByVal pnPeriodoVigMeses As Integer, ByVal pdFechaIni As Date, _
                                        ByVal pdFechaFin As Date, ByVal pbPersJur As Boolean, ByVal psMovNro As String, ByVal pMatAgencias As Variant, ByVal pbDineroNuevo As Boolean) As Boolean
'APRI20210621 ERS031-2021 Add pbDineroNuevo
Dim rs As ADODB.Recordset
Dim nNuevaCampanaCod As Integer
Dim i As Integer

On Error GoTo Error
    ActualizaCaptacCampanas = False
    dbCmact.BeginTrans
    
    sSql = "stp_upd_CaptacCampanas " & pnCampanaCod & ",'" & pcCampanaDesc & "'," & pnMoneda & "," & pnProd & "," & pnSubProd & "," & pnMontoMin & "," & _
                                       pnPeriodoVigMeses & ",'" & Format(pdFechaIni, "yyyyMMdd") & "','" & Format(pdFechaFin, "yyyyMMdd") & "'," & _
                                       IIf(pbPersJur, 1, 0) & ",'" & psMovNro & "'," & IIf(pbDineroNuevo, 1, 0)
    dbCmact.Execute sSql
    
    EliminaCaptacCampanasAge pnCampanaCod
    For i = 0 To UBound(pMatAgencias)
        InsertaCaptacCampanasAge pnCampanaCod, pMatAgencias(i)
    Next i
    
    dbCmact.CommitTrans
    ActualizaCaptacCampanas = True
    Exit Function
Error:
    dbCmact.RollbackTrans
    ActualizaCaptacCampanas = False
    Err.Raise Err.Number, "NCapDefinicion:ActualizaCaptacCampanas", Err.Description
End Function
Public Sub EliminaCaptacCampanas(ByVal pnCampanaCod As Long)

On Error GoTo Error
    sSql = "stp_del_CaptacCampanas " & pnCampanaCod
    dbCmact.Execute sSql
    Exit Sub
Error:
    Err.Raise Err.Number, "NCapDefinicion:EliminaCaptacCampanas", Err.Description
End Sub
Public Function GetCaptacCampanasAge(Optional ByVal pnCampanaCod As Long = 0) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_CaptacCampanasAge " & pnCampanaCod

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    Set GetCaptacCampanasAge = rs
    Set rs = Nothing
End Function
Public Sub InsertaCaptacCampanasAge(ByVal pnCampanaCod As Long, ByVal cAgeCod As String)

On Error GoTo Error
    sSql = "stp_ins_CaptacCampanasAge " & pnCampanaCod & ",'" & cAgeCod & "'"
    dbCmact.Execute sSql
    Exit Sub
Error:
    dbCmact.RollbackTrans
    Err.Raise Err.Number, "NCapDefinicion:InsertaCaptacCampanasAge", Err.Description
End Sub
Public Sub EliminaCaptacCampanasAge(ByVal pnCampanaCod As Long)

On Error GoTo Error
    sSql = "stp_del_CaptacCampanasAge " & pnCampanaCod
    dbCmact.Execute sSql
    Exit Sub
Error:
    dbCmact.RollbackTrans
    Err.Raise Err.Number, "NCapDefinicion:EliminaCaptacCampanasAge", Err.Description
End Sub
Public Function GetCaptacTasasCampanas(ByVal pnCampanaCod As Long) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    sSql = "exec stp_sel_CaptacTasasCampanas " & pnCampanaCod

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    rs.ActiveConnection = Nothing
    Set GetCaptacTasasCampanas = rs
    Set rs = Nothing
End Function
Public Sub NuevaTasaCampana(ByVal pnCampanaCod As Long, ByVal pbOrdPag As Boolean, ByVal pnPlazoIni As Long, ByVal pnPlazoFin As Long, _
                            ByVal pnMontoIni As Double, ByVal pnMontoFin As Double, ByVal pnValorTasa As Double, ByVal psMovNro As String)
On Error GoTo Error
    sSql = "stp_ins_CaptacTasasCampanas " & pnCampanaCod & "," & IIf(pbOrdPag, 1, 0) & "," & pnPlazoIni & "," & pnPlazoFin & "," & _
                                            pnMontoIni & "," & pnMontoFin & "," & pnValorTasa & ",'" & psMovNro & "'"
    dbCmact.Execute sSql
    Exit Sub
Error:
    Err.Raise Err.Number, "NCapDefinicion:NuevaTasaCampana", Err.Description
End Sub
Public Sub ActualizaTasaCampana(ByVal pnCampanaCod As Long)
On Error GoTo Error
    sSql = "stp_upd_CaptacTasasCampanas " & pnCampanaCod
    dbCmact.Execute sSql
    Exit Sub
Error:
    Err.Raise Err.Number, "NCapDefinicion:ActualizaTasaCampana", Err.Description
End Sub
Public Function GetCapTasaInteresCamp(ByVal pnProducto As Producto, ByVal pnSubProducto As Integer, ByVal pnMoneda As Integer, ByVal pnPlazo As Long, _
                                 ByVal pnMonto As Double, ByVal psCodAge As String, ByVal pdFecSis As Date, ByVal pbPersJur As Boolean, _
                                 ByVal pbOrdPag As Boolean, ByRef pnCampanaCod As Long, Optional ByRef psCampanaDesc As String = "", Optional ByRef pnDineroNuevo As Integer = 0) As Double 'JUEZ 20160420
'APRI20210621 ERS031-2021 ADD pnDineroNuevo

    Dim rsTasa As ADODB.Recordset
    Dim cOrdPag As String

    sSql = "EXEC stp_sel_GetCapTasaInteresCamp " & pnProducto & "," & pnSubProducto & "," & pnMoneda & "," & pnPlazo & "," & pnMonto & ",'" & psCodAge & "','" & _
                                                   Format(pdFecSis, "yyyyMMdd") & "'," & IIf(pbPersJur, 1, 0) & "," & IIf(pbOrdPag, 1, 0)

    Set rsTasa = New ADODB.Recordset
    rsTasa.CursorLocation = adUseClient
    rsTasa.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsTasa.ActiveConnection = Nothing
    If Not (rsTasa.EOF Or rsTasa.BOF) Then
        GetCapTasaInteresCamp = rsTasa("nTasaValor")
        pnCampanaCod = CLng(rsTasa("nCampanaCod"))
        psCampanaDesc = rsTasa("cCampanaDesc")
        pnDineroNuevo = rsTasa("bDineroNuevo")
    Else
        GetCapTasaInteresCamp = 0
        pnCampanaCod = 0
        psCampanaDesc = ""
        pnDineroNuevo = 0
    End If
    rsTasa.Close
    Set rsTasa = Nothing
End Function
'END JUEZ ******************************************************************************
'APRI20190109 ERS077-2018
Public Function GetCapCostoMantenimiento(ByVal nProducto As Integer, ByVal nSubProd As Integer, ByVal nMoneda As Moneda, ByVal nMonto As Double) As Currency

    Dim rsMant As ADODB.Recordset

    sSql = "EXEC stp_sel_GetCapCostoMantenimiento " & nProducto & "," & nSubProd & "," & nMoneda & "," & nMonto
    
    Set rsMant = New ADODB.Recordset
    rsMant.CursorLocation = adUseClient
    rsMant.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsMant.ActiveConnection = Nothing
    If Not (rsMant.EOF Or rsMant.BOF) Then
        GetCapCostoMantenimiento = rsMant("nValor")
    Else
        GetCapCostoMantenimiento = 0
    End If
    rsMant.Close
    Set rsMant = Nothing
End Function
'END APRI



