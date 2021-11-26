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

Public Function ActualizaParametros(ByVal nVariable As Long, ByVal sDesc As String, ByVal nValor As Double) As Boolean
Dim sSql As String
Dim rsPar As Recordset
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

Public Function GetParametros() As Recordset
Dim sSql As String
Dim rsParam As Recordset

On Error GoTo ErrGetParam
Set rsParam = New Recordset
rsParam.CursorLocation = adUseClient
sSql = "Select nParCod Variable, cParDescripcion Descripcion, nParValor Valor, Tag = '' " _
    & "FROM Parametro WHERE nParProd = " & gPrdParamCaptac
rsParam.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
rsParam.ActiveConnection = Nothing
Set GetParametros = rsParam
Set rsParam = Nothing
Exit Function
ErrGetParam:
    MsgBox "Error Open DataBase.", vbExclamation, "Aviso"
End Function

Public Function GetTarifario(ByVal sProd As Producto, ByVal sMoneda As Moneda, _
        ByVal nTipo As CaptacTipoTasa, ByVal sAgencia As String) As Recordset
Dim sSql As String
Dim rsParam As Recordset

Set rsParam = New Recordset
rsParam.CursorLocation = adUseClient
sSql = "Select nValorIni, nValorFin, nPlazoIni, nPlazoFin, cOrdPag, " _
    & "nTasaValor, nTasaCod From CaptacTasas Where nTasaProd = " & sProd & " And nTasaMon = " & sMoneda & " And " _
    & "nTasaTpo = " & nTipo & " And cCodAge = '" & sAgencia & "' Order by nTasaCod"
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
        ByVal bOrdPag As Boolean, ByVal nValor As Double)

Dim sSql As String, sOrdPag As String
On Error GoTo ErrGetParam

sOrdPag = IIf(bOrdPag, "1", "0")
sSql = "Update CaptacTasas Set nValorIni = " & nMonIni & ", nValorFin = " & nMonFin & ", nTasaValor = " & nValor & ", " _
    & "nPlazoIni = " & nPlazoIni & ", nPlazoFin = " & nPlazoFin & ", " _
    & "cOrdPag = '" & sOrdPag & "' Where nTasaCod = " & nCodigo

dbCmact.Execute sSql
Exit Sub
ErrGetParam:
    Err.Raise Err.Number, "NCapDefinicion:ActualizaTasa", Err.Description
End Sub

Public Sub NuevaTasa(ByVal nProd As Producto, ByVal nMon As Moneda, _
        ByVal nTipoTasa As CaptacTipoTasa, ByVal nMonIni As Double, nMonFin As Double, _
        ByVal nPlazoIni As Long, ByVal nPlazoFin As Long, ByVal bOrdPag As Boolean, _
        ByVal sAgencia As String, ByVal nValor As Double)

Dim sSql As String, sOrdPag As String
On Error GoTo ErrTasa
sOrdPag = IIf(bOrdPag, "1", "0")
sSql = "Insert CaptacTasas (nTasaProd,nTasaMon,nTasaTpo,nValorIni,nValorFin,nPlazoIni,nPlazoFin,cOrdPag,cCodAge,nTasaValor) " _
    & "Values (" & nProd & "," & nMon & "," & nTipoTasa & "," & nMonIni & "," & nMonFin & "," & nPlazoIni & "," & nPlazoFin & ",'" & sOrdPag & "','" & sAgencia & "'," & nValor & ")"

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

Public Function GetCapTasaInteres(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
    ByVal nTipoTasa As CaptacTipoTasa, Optional nPlazo As Long = 0, Optional nValor As Double = 0, _
    Optional sCodAge As String = "", Optional bOrdPag As Boolean = False) As Double

Dim rsTasa As Recordset
Dim sSql As String
Dim cOrdPag As String
cOrdPag = IIf(bOrdPag, "1", "0")

If sCodAge = "" Then sCodAge = gsCodAge
sSql = "SELECT nTasaValor FROM CaptacTasas WHERE nTasaProd = " & nProducto & " AND " _
    & "nTasaMon = " & nMoneda & " AND nTasaTpo = " & nTipoTasa & " And " & nPlazo & " " _
    & "BETWEEN nPlazoIni AND nPlazoFin And " & nValor & " BETWEEN nValorIni And nValorFin " _
    & "And cOrdPag = '" & cOrdPag & "' And cCodAge = '" & sCodAge & "'"

Set rsTasa = New Recordset
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
Dim rsVar As Recordset
Dim sSql As String
sSql = "SELECT nParValor FROM Parametro WHERE nParCod = " & nParametro & " " _
    & "And nParProd = " & gPrdParamCaptac
Set rsVar = New Recordset
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
Dim rsVar As Recordset
Dim sSql As String
sSql = "SELECT cParDescripcion FROM Parametro WHERE nParCod = " & nParametro
Set rsVar = New Recordset
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

Public Function GetCapMovTipo() As Recordset
Dim rsVar As Recordset
Dim sSql As String
sSql = "Select O.cOpeCod, O.cOpeDesc, Prod = CASE WHEN M.Prod IS NULL THEN " _
    & "CASE WHEN O.cOpeCod LIKE '2[03]%' Or O.cOpeCod LIKE '2601%' THEN 'AHORROS' + Space(75) + '" & gCapAhorros & "' " _
    & "     WHEN O.cOpeCod LIKE '2[14]%' Or O.cOpeCod LIKE '2602%' THEN 'PLAZO FIJO' + Space(75) + '" & gCapPlazoFijo & "' " _
    & "     WHEN O.cOpeCod LIKE '2[25]%' Or O.cOpeCod LIKE '2603%' THEN 'CTS' + Space(75) + '" & gCapCTS & "' END " _
    & "ELSE M.Prod END, M.Tipo, Estad = CONVERT(Int,ISNULL(M.Estad,0)) From " _
    & "(Select M.cOpeCod, UPPER(K1.cConsDescripcion) + Space(75) + Convert(Varchar(3),nProducto) Prod, " _
    & "UPPER(K.cConsDescripcion) + Space(75) + Convert(Varchar(2),nCapMovTpo) Tipo, bEstadistica Estad From CapMovTipo M INNER JOIN Constante K " _
    & "ON M.nCapMovTpo = K.nConsValor INNER JOIN Constante K1 ON M.nProducto = K1.nConsValor Where K.nConsCod = " & gCaptacMovTipo & "  " _
    & "And K1.nConsCod = " & gProducto & ") M " _
    & "RIGHT JOIN OpeTpo O ON M.cOpeCod = O.cOpeCod Where O.cOpeCod LIKE '2[0-7]%' " _
    & "Order by O.cOpeCod"
Set rsVar = New Recordset
rsVar.CursorLocation = adUseClient
rsVar.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsVar.ActiveConnection = Nothing
Set GetCapMovTipo = rsVar
Set rsVar = Nothing
End Function

Private Sub InsertaCapMovTipo(ByVal sOpeCod As String, ByVal nProducto As Producto, _
        ByVal nCapMovTpo As CaptacMovTipo, ByVal bEstadistica As Boolean)
Dim sSql As String
sSql = "Insert CapMovTipo (cOpeCod,nProducto,nCapMovTpo,bEstadistica) " _
    & "Values ('" & sOpeCod & "'," & nProducto & "," & nCapMovTpo & ", " & IIf(bEstadistica, 1, 0) & ")"
dbCmact.Execute sSql
End Sub

Private Sub EilminaCapMovTipo(ByVal sOpeCod As String)
Dim sSql As String
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







