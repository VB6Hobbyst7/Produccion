VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nCapDefinicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public dbCmact As Connection
Dim dbCmactHist As Connection
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String
Dim sSQL As String

Public Function ActualizaParametros(ByVal nVariable As Long, ByVal sDesc As String, ByVal nValor As Double) As Boolean
Dim rsPar As Recordset
Dim sDescH As String, sVar As String
Dim nVal As Double
On Error GoTo ErrGetParam
sSQL = "Update Parametro Set cParDescripcion = '" & sDesc & "', nParValor = " & nValor & " WHERE " _
    & "nParCod = " & nVariable & " And nParProd = " & gPrdParamCaptac
dbCmact.Execute sSQL
ActualizaParametros = True
Exit Function
ErrGetParam:
    Err.Raise Err.Number, "NCapDefinicion:ActualizaParametros", Err.Description
    ActualizaParametros = False
End Function

Public Function GetParametros() As Recordset
Dim rsParam As Recordset

On Error GoTo ErrGetParam
Set rsParam = New Recordset
rsParam.CursorLocation = adUseClient
sSQL = "Select nParCod Variable, cParDescripcion Descripcion, nParValor Valor, Tag = '' " _
    & "FROM Parametro WHERE nParProd = " & gPrdParamCaptac
rsParam.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
rsParam.ActiveConnection = Nothing
Set GetParametros = rsParam
Set rsParam = Nothing
Exit Function
ErrGetParam:
    MsgBox "Error Open DataBase.", vbExclamation, "Aviso"
End Function

Public Function GetTarifario(ByVal sProd As Producto, ByVal sMoneda As Moneda, _
        ByVal nTipo As CaptacTipoTasa, ByVal sAgencia As String) As Recordset
Dim rsParam As Recordset

Set rsParam = New Recordset
rsParam.CursorLocation = adUseClient
sSQL = "Select nValorIni, nValorFin, nPlazoIni, nPlazoFin, cOrdPag, " _
    & "nTasaValor, nTasaCod From CaptacTasas Where nTasaProd = " & sProd & " And nTasaMon = " & sMoneda & " And " _
    & "nTasaTpo = " & nTipo & " And cCodAge = '" & sAgencia & "' Order by nTasaCod"
rsParam.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsParam.ActiveConnection = Nothing

Set GetTarifario = rsParam
Set rsParam = Nothing
Exit Function
ErrGetParam:
    Err.Raise Err.Number, "NCapDefinicion", Err.Description
End Function

Public Sub ActualizaTasa(ByVal nCodigo As Integer, ByVal nMonIni As Double, _
        ByVal nMonFin As Double, ByVal nPlazoIni As Long, ByVal nPlazoFin As Long, _
        ByVal bOrdPag As Boolean, ByVal nValor As Double)

Dim sOrdPag As String
On Error GoTo ErrGetParam

sOrdPag = IIf(bOrdPag, "1", "0")
sSQL = "Update CaptacTasas Set nValorIni = " & nMonIni & ", nValorFin = " & nMonFin & ", nTasaValor = " & nValor & ", " _
    & "nPlazoIni = " & nPlazoIni & ", nPlazoFin = " & nPlazoFin & ", " _
    & "cOrdPag = '" & sOrdPag & "' Where nTasaCod = " & nCodigo

dbCmact.Execute sSQL
Exit Sub
ErrGetParam:
    Err.Raise Err.Number, "NCapDefinicion:ActualizaTasa", Err.Description
End Sub

Public Sub NuevaTasa(ByVal nProd As Producto, ByVal nMon As Moneda, _
        ByVal nTipoTasa As CaptacTipoTasa, ByVal nMonIni As Double, nMonFin As Double, _
        ByVal nPlazoIni As Long, ByVal nPlazoFin As Long, ByVal bOrdPag As Boolean, _
        ByVal sAgencia As String, ByVal nValor As Double)

Dim sOrdPag As String
On Error GoTo ErrTasa
sOrdPag = IIf(bOrdPag, "1", "0")
sSQL = "Insert CaptacTasas (nTasaProd,nTasaMon,nTasaTpo,nValorIni,nValorFin,nPlazoIni,nPlazoFin,cOrdPag,cCodAge,nTasaValor) " _
    & "Values (" & nProd & "," & nMon & "," & nTipoTasa & "," & nMonIni & "," & nMonFin & "," & nPlazoIni & "," & nPlazoFin & ",'" & sOrdPag & "','" & sAgencia & "'," & nValor & ")"

dbCmact.Execute sSQL
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

Public Function GetCapTasaInteres(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
    ByVal nTipoTasa As CaptacTipoTasa, Optional nPlazo As Long = 0, Optional nValor As Double = 0, _
    Optional sCodage As String = "", Optional bOrdPag As Boolean = False) As Double

Dim rsTasa As Recordset
Dim cOrdPag As String

If sCodage = "" Then sCodage = gsCodAge


If nProducto = gCapAhorros Or nProducto = gCapCTS Then
    cOrdPag = IIf(bOrdPag, "1", "0")
    sSQL = "SELECT nTasaValor FROM CaptacTasas WHERE nTasaProd = " & nProducto & " AND " _
        & "nTasaMon = " & nMoneda & " AND nTasaTpo = " & nTipoTasa & " And (" & nValor & " >= nValorIni " _
        & "And " & nValor & " < nValorFin) And cOrdPag = '" & cOrdPag & "' And " _
        & "cCodAge = '" & sCodage & "'"
Else
    If nPlazo < 31 Then nPlazo = 31

    sSQL = "SELECT nTasaValor FROM CaptacTasas WHERE nTasaProd = " & nProducto & " AND " _
        & "nTasaMon = " & nMoneda & " AND nTasaTpo = " & nTipoTasa & " And (" & nPlazo & " " _
        & ">= nPlazoIni AND " & nPlazo & " <= nPlazoFin) And (" & nValor & " >= nValorIni " _
        & "And " & nValor & " < nValorFin) And cCodAge = '" & sCodage & "'"
End If
Set rsTasa = New Recordset
rsTasa.CursorLocation = adUseClient
rsTasa.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
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
Dim rsVar As Recordset
sSQL = "SELECT nParValor FROM Parametro WHERE nParCod = " & nParametro & " " _
    & "And nParProd = " & gPrdParamCaptac
Set rsVar = New Recordset
rsVar.CursorLocation = adUseClient
rsVar.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
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
Dim rsVar As Recordset
sSQL = "SELECT cParDescripcion FROM Parametro WHERE nParCod = " & nParametro
Set rsVar = New Recordset
rsVar.CursorLocation = adUseClient
rsVar.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsVar.ActiveConnection = Nothing
If rsVar.EOF And rsVar.BOF Then
    GetCapParametroDesc = ""
Else
    GetCapParametroDesc = rsVar("cParDescripcion")
End If
rsVar.Close
Set rsVar = Nothing
End Function

Public Function GetCapMovTipo() As Recordset
Dim rsVar As Recordset
sSQL = "Select O.cOpeCod, O.cOpeDesc, Prod = CASE WHEN M.Prod IS NULL THEN " _
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
Set rsVar = New Recordset
rsVar.CursorLocation = adUseClient
rsVar.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsVar.ActiveConnection = Nothing
Set GetCapMovTipo = rsVar
Set rsVar = Nothing
End Function

Private Sub InsertaCapMovTipo(ByVal sOpeCod As String, ByVal nProducto As Producto, _
        ByVal nCapMovTpo As CaptacMovTipo, ByVal bEstadistica As Boolean)
sSQL = "Insert CapMovTipo (cOpeCod,nProducto,nCapMovTpo,bEstadistica) " _
    & "Values ('" & sOpeCod & "'," & nProducto & "," & nCapMovTpo & ", " & IIf(bEstadistica, 1, 0) & ")"
dbCmact.Execute sSQL
End Sub

Private Sub EilminaCapMovTipo(ByVal sOpeCod As String)
sSQL = "Delete CapMovTipo Where cOpeCod = '" & sOpeCod & "'"
dbCmact.Execute sSQL
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
Dim rsMov As Recordset
sSQL = "Select nMovNro From Mov Where cMovNro = '" & sMovNro & "'"
Set rsMov = New Recordset
rsMov.CursorLocation = adUseClient
rsMov.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
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
sSQL = "Update CaptacPlazoFijo Set bBusCredPend = " & IIf(bValor, 1, 0) & " " _
    & "Where cCtaCod = '" & sCuenta & "'"
dbCmact.Execute sSQL
End Sub

Public Sub AgregaMov(ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
        ByVal sGlosa As String, Optional nMovEstado As MovEstado = gMovEstContabMovContable, _
        Optional nMovFlag As MovFlag = gMovFlagVigente)
sSQL = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sGlosa & "'," & nMovEstado & "," & nMovFlag & ")"
dbCmact.Execute sSQL
End Sub

Public Sub AgregaMovCap(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double)
sSQL = "INSERT MovCap (nMovNro,cOpeCod,cCtaCod,nMonto,nSaldoDisponible,nSaldoContable) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nMonto & "," & nSaldoDisp & "," & nSaldoCnt & ")"
dbCmact.Execute sSQL
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
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Public Sub ActualizaDocRecParam(ByVal nTipoCheque As ChequeTipo, ByVal nPlaza As ChequePlaza, _
        ByVal nMoneda As Moneda, ByVal nDias As Double)
sSQL = "Update DocRecParametro Set nDiasVal = " & nDias & " " _
    & "Where nPlaza = " & nPlaza & " And nTpoCheque = " & nTipoCheque & " And nMoneda = " & nMoneda
dbCmact.Execute sSQL
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
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Public Function GetDocRecParametro() As Recordset
Dim rsDoc As Recordset
sSQL = "Select UPPER(K.cConsDescripcion) + Space(50) + Convert(Varchar(1),C.nMoneda) Moneda, " _
    & "UPPER(K1.cConsDescripcion) + Space(50) + Convert(Varchar(1),C.nTpoCheque) TipoCheque, " _
    & "UPPER(K2.cConsDescripcion) + Space(50) + Convert(Varchar(1),C.nPlaza) Plaza, C.nDiasVal " _
    & "From DocRecParametro C JOIN Constante K ON C.nMoneda = K.nConsValor " _
    & "JOIN Constante K1 ON C.nTpoCheque = K1.nConsValor JOIN " _
    & "Constante K2 ON C.nPlaza = K2.nConsValor Where K.nConsCod = " & gMoneda & " " _
    & "And K1.nConsCod = " & gChequeTipo & " And K2.nConsCod = " & gChequePlaza & " Order by C.nMoneda, C.nTpoCheque, C.Plaza"

Set rsDoc = New Recordset
rsDoc.CursorLocation = adUseClient
rsDoc.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsDoc.ActiveConnection = Nothing
Set GetDocRecParametro = rsDoc
Set rsDoc = Nothing
End Function

Public Function GetDiasMinValorizacion(ByVal nMoneda As Moneda, _
    ByVal nPlaza As ChequePlaza, ByVal nTipoCheque As ChequeTipo) As Integer
Dim rsChq As Recordset

sSQL = "Select nDiasVal From DocRecParametro Where nTpoCheque = " & nTipoCheque & " " _
    & "And nPlaza = " & nPlaza & " And nMoneda = " & nMoneda

Set rsChq = New Recordset
rsChq.CursorLocation = adUseClient
rsChq.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
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
        ByVal bOrdPag As Boolean) As Recordset
Dim rsPar As Recordset
Dim sOrdPag As String
sOrdPag = IIf(bOrdPag, "1", "0")
Set rsPar = New Recordset
rsPar.CursorLocation = adUseClient
sSQL = "Select K.cConsDescripcion + Space(100) + Convert(Varchar(1),K.nConsValor), ISNULL(P.nMontoMinApert,0), ISNULL(P.nSaldoMin,0) " _
    & "From Constante K LEFT JOIN (Select nPersoneria, nMontoMinApert, nSaldoMin From CapPersParam Where nProducto = " & nProducto & " " _
    & "And nMoneda = " & nMoneda & " And cOrdPag = '" & sOrdPag & "') P ON K.nConsValor = P.nPersoneria Where " _
    & "K.nConsCod = " & gPersPersoneria & " Order by K.nConsValor"
rsPar.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsPar.ActiveConnection = Nothing
Set GetCapPersoneriaParam = rsPar
Set rsPar = Nothing
End Function

Public Function GetSaldoMinimoPersoneria(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal nPersoneria As PersPersoneria, ByVal bOrdPag As Boolean) As Double
Dim rsPar As Recordset
Dim sOrdPag As String
sOrdPag = IIf(bOrdPag, "1", "0")
Set rsPar = New Recordset
rsPar.CursorLocation = adUseClient
sSQL = "Select nSaldoMin From CapPersParam Where nPersoneria = " & nPersoneria & " And " _
    & "nProducto = " & nProducto & " And nMoneda = " & nMoneda & " And cOrdPag = '" & sOrdPag & "'"
rsPar.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
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
        ByVal nPersoneria As PersPersoneria, ByVal bOrdPag As Boolean) As Double
Dim rsPar As Recordset
Dim sOrdPag As String
sOrdPag = IIf(bOrdPag, "1", "0")

Set rsPar = New Recordset
rsPar.CursorLocation = adUseClient
sSQL = "Select nMontoMinApert From CapPersParam Where nPersoneria = " & nPersoneria & " And " _
    & "nProducto = " & nProducto & " And nMoneda = " & nMoneda & " And cOrdPag = '" & sOrdPag & "'"
rsPar.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
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
        ByVal nPersoneria As PersPersoneria, ByVal bOrdPag As Boolean) As Double
Dim rsPar As Recordset
Dim sOrdPag As String
sOrdPag = IIf(bOrdPag, "1", "0")

Set rsPar = New Recordset
rsPar.CursorLocation = adUseClient
sSQL = "Select nMontoMinDep From CapPersParam Where nPersoneria = " & nPersoneria & " And " _
    & "nProducto = " & nProducto & " And nMoneda = " & nMoneda & " And cOrdPag = '" & sOrdPag & "'"
rsPar.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsPar.ActiveConnection = Nothing
If rsPar.EOF And rsPar.BOF Then
    GetMontoMinimoDepPersoneria = 0
Else
    GetMontoMinimoDepPersoneria = rsPar("nMontoMinDep")
End If
rsPar.Close
Set rsPar = Nothing
End Function

Public Function GetMontoMinimoRetPersoneria(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal nPersoneria As PersPersoneria, ByVal bOrdPag As Boolean) As Double
Dim rsPar As Recordset
Dim sOrdPag As String
sOrdPag = IIf(bOrdPag, "1", "0")

Set rsPar = New Recordset
rsPar.CursorLocation = adUseClient
sSQL = "Select nMontoMinRet From CapPersParam Where nPersoneria = " & nPersoneria & " And " _
    & "nProducto = " & nProducto & " And nMoneda = " & nMoneda & " And cOrdPag = '" & sOrdPag & "'"
rsPar.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsPar.ActiveConnection = Nothing
If rsPar.EOF And rsPar.BOF Then
    GetMontoMinimoRetPersoneria = 0
Else
    GetMontoMinimoRetPersoneria = rsPar("nMontoMinRet")
End If
rsPar.Close
Set rsPar = Nothing
End Function

Public Function GetMontoMinimoRetOPPersoneria(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal nPersoneria As PersPersoneria, ByVal bOrdPag As Boolean) As Double
Dim rsPar As Recordset
Dim sOrdPag As String
sOrdPag = IIf(bOrdPag, "1", "0")

Set rsPar = New Recordset
rsPar.CursorLocation = adUseClient
sSQL = "Select nMontoMinRetOP From CapPersParam Where nPersoneria = " & nPersoneria & " And " _
    & "nProducto = " & nProducto & " And nMoneda = " & nMoneda & " And cOrdPag = '" & sOrdPag & "'"
rsPar.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsPar.ActiveConnection = Nothing
If rsPar.EOF And rsPar.BOF Then
    GetMontoMinimoRetOPPersoneria = 0
Else
    GetMontoMinimoRetOPPersoneria = rsPar("nMontoMinRetOP")
End If
rsPar.Close
Set rsPar = Nothing
End Function

Public Sub ActualizaCapPersoneriaParam(ByVal rsPar As Recordset, ByVal nProducto As Producto, _
        ByVal nMoneda As Moneda, ByVal bOrdPag As Boolean, ByVal sMovNro As String)
Dim bTrans As Boolean
Dim nMonto As Double, nSaldo As Double
Dim nPers As PersPersoneria
Dim sOrdPag As String
sOrdPag = IIf(bOrdPag, "1", "0")

On Error GoTo ErrGraba
dbCmact.BeginTrans
bTrans = True
sSQL = "Delete CapPersParam Where nProducto = " & nProducto & " And nMoneda = " & nMoneda & " And cOrdPag = '" & sOrdPag & "'"
dbCmact.Execute sSQL
Do While Not rsPar.EOF
    nPers = CLng(Right(rsPar("Personería"), 1))
    nMonto = CDbl(rsPar("Mon Min Apert"))
    nSaldo = CDbl(rsPar("Saldo Minimo"))
    sSQL = "Insert CapPersParam (nProducto,nMoneda,nPersoneria,cOrdPag,nMontoMinApert,nSaldoMin,cMovNro) " _
        & "Values (" & nProducto & "," & nMoneda & "," & nPers & ",'" & sOrdPag & "'," & nMonto & "," & nSaldo & ",'" & sMovNro & "')"
    dbCmact.Execute sSQL
    rsPar.MoveNext
Loop
dbCmact.CommitTrans
bTrans = False
Exit Sub
ErrGraba:
    If bTrans Then dbCmact.RollbackTrans
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Public Function GetTarifaParametro(ByVal nOperacion As CaptacOperacion, _
            ByVal nMoneda As Moneda, ByVal nParametro As CaptacParametro) As Recordset
Dim rsPar As Recordset

Set rsPar = New Recordset
rsPar.CursorLocation = adUseClient
sSQL = "Select P.nParValor From Parametro P JOIN Tarifas T ON P.nParProd = T.nParProd And " _
    & "P.nParCod = T.nParCod Where T.cOpeCod = " & nOperacion & " And " _
    & "T.nMoneda = " & nMoneda & " And T.nParCod = " & nParametro
rsPar.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set rsPar.ActiveConnection = Nothing
Set GetTarifaParametro = rsPar
Set rsPar = Nothing

End Function

Public Function EsOperacionEfectivo(ByVal sOperacion As String) As Boolean
Dim rsEfec As Recordset

'sSql = "Select GO.cOpeCod From GrupoOpe G JOIN GruposOpe GO ON G.cGrupoCod = GO.cGrupoCod " _
'    & "WHERE GO.cOpeCod = '" & sOperacion & "' And G.nEfectivo = 1"
    
sSQL = " SELECT GO.cOpeCod From GrupoOpe G JOIN GruposOpe GO ON G.cGrupoCod = GO.cGrupoCod " _
        & " WHERE GO.cOpeCod = '" & sOperacion & "' AND G.nEfectivo = 1 " _
        & " And Not cOpeCod like '9010%' and cOpeCod not in (401401, 402401) " _
        & " And cOpeCod not in (Select cOpeCod from OpeDoc Where nDocTpo = 47) "

Set rsEfec = New Recordset
rsEfec.CursorLocation = adUseClient
rsEfec.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsEfec.ActiveConnection = Nothing

If rsEfec.EOF And rsEfec.BOF Then
    EsOperacionEfectivo = False
Else
    EsOperacionEfectivo = True
End If

rsEfec.Close
Set rsEfec = Nothing

End Function

Private Sub Class_Initialize()
Dim sConn As String
Dim ClsIni As ClsIni.ClasIni
Set ClsIni = New ClsIni.ClasIni
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



