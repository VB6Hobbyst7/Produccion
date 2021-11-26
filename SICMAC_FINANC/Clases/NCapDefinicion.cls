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

Public dbCmact As DConecta
Dim dbCmactHist As Connection
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String

Public Function ActualizaParametros(ByVal sVariable As String, ByVal sDesc As String, ByVal nValor As Double) As Boolean
Dim sSql As String
Dim rsPar As Recordset
Dim sDescH As String, sVar As String
Dim nVal As Double

On Error GoTo ErrGetParam

sSql = "Select cParamVar, cParamDesc, nParamValor FROM CaptacParametro WHERE " _
    & "cParamVar = '" & sVariable & "'"
Set rsPar = New Recordset
rsPar.CursorLocation = adUseClient
rsPar.Open sSql, dbCmactHist, adOpenStatic, adLockReadOnly, adCmdText
Set rsPar.ActiveConnection = Nothing

If Not (rsPar.EOF And rsPar.BOF) Then
    sVar = sVariable
    sDescH = rsPar("cParamDesc")
    nVal = rsPar("cParamValor")
    rsPar.Close
    Set rsPar = Nothing
    sSql = "Insert CaptacParametro (cParamVar,dUltAct,cCodUsu,cParamDesc,nParamValor) " _
        & "VALUES ('" & sVariable & "','" & sDescH & "'," & nVal & ")"
    dbCmactHist.Execute sSql
End If

sSql = "Update CaptacParametro Set cParamDesc = '" & sDesc & "', nParamValor = " & nValor & " WHERE " _
    & "cParamVar = '" & sVariable & "'"
dbCmact.Ejecutar sSql
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
sSql = "Select cParamVar Variable, cParamDesc Descripcion, nParamValor Valor, Tag = '' from CaptacParametro"
Set rsParam = dbCmact.CargaRecordSet(sSql)
Set GetParametros = rsParam
Set rsParam = Nothing
Exit Function
ErrGetParam:
    MsgBox "Error Open DataBase.", vbExclamation, "Aviso"
End Function

Public Function GetTarifario(ByVal sProd As Producto, ByVal sMoneda As Moneda, ByVal sTipo As String) As Recordset
Dim sSql As String
Dim rsParam As Recordset

Set rsParam = New Recordset
rsParam.CursorLocation = adUseClient
sSql = "Select nTasaCod, nValorIni, nValorFin, nTasaValor From CaptacTasas " _
    & "Where cTasaProd = '" & sProd & "' And cTasaMon = '" & sMoneda & "' And cTasaTpo = '" & sTipo & "' Order by nTasaCod"
rsParam.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsParam.ActiveConnection = Nothing

Set GetTarifario = rsParam
Set rsParam = Nothing
Exit Function
ErrGetParam:
    Err.Raise Err.Number, "NCapDefinicion", Err.Description
End Function

Public Sub ActualizaTasa(ByVal nProd As Producto, ByVal nMon As Moneda, _
                            ByVal sTipo As String, ByVal nCodigo As Integer, _
                            ByVal nMonIni As Currency, nMonFin As Currency, _
                            ByVal nValor As Currency)

Dim sSql As String
On Error GoTo ErrGetParam

sSql = "Update CaptacTasas Set nValorIni = " & nMonIni & ", nValorFin = " & nMonFin & ", nTasaValor = " & nValor & " " _
    & "Where cTasaProd = '" & nProd & "' And cTasaMon = '" & nMon & "' And cTasaTpo = '" & sTipo & "' And nTasaCod = " & nCodigo

dbCmact.Ejecutar sSql
Exit Sub
ErrGetParam:
    Err.Raise Err.Number, "NCapDefinicion:ActualizaTasa", Err.Description
End Sub

Public Sub NuevaTasa(ByVal nProd As Producto, ByVal nMon As Moneda, _
                    ByVal sTipo As String, ByVal nCodigo As Integer, _
                    ByVal nMonIni As Double, nMonFin As Double, _
                    ByVal nValor As Double)

Dim sSql As String
On Error GoTo ErrTasa
sSql = "Insert CaptacTasas (cTasaProd,cTasaMon,cTasaTpo,nTasaCod,nValorIni,nValorFin,nTasaValor) " _
    & "Values ('" & nProd & "','" & nMon & "','" & sTipo & "'," & nCodigo & "," & nMonIni & "," & nMonFin & "," & nValor & ")"

dbCmact.Ejecutar sSql
Exit Sub
ErrTasa:
    Err.Raise Err.Number, "NCapDefinicion:NuevaTasa", Err.Description
End Sub

'
'Public Sub ActualizaTasaCaptacion(ByVal nProd As Producto, ByVal nMon As Moneda, _
'                    ByVal sTipoAnt As String, ByVal sTipoNuevo As String, _
'                    ByVal nValorAnt As Double, nValorNuevo As Double, _
'                    ByVal dFecha As Date)
'
'Dim cmdTasa As Command
'Dim prmTasa As Parameter
'
'Set cmdTasa = New Command
'cmdTasa.CommandText = "spCapActualizaTasas"
'cmdTasa.Name = "spActualizaTasa"
'cmdTasa.CommandType = adCmdStoredProc
'
'Set prmTasa = cmdTasa.CreateParameter("Producto", adVarChar, adParamInput, 3, Trim(Str(nProd)))
'cmdTasa.Parameters.Append prmTasa
'Set prmTasa = cmdTasa.CreateParameter("Moneda", adVarChar, adParamInput, 1, Trim(Str(nMon)))
'cmdTasa.Parameters.Append prmTasa
'Set prmTasa = cmdTasa.CreateParameter("TipoAnt", adVarChar, adParamInput, 3, sTipoAnt)
'cmdTasa.Parameters.Append prmTasa
'Set prmTasa = cmdTasa.CreateParameter("TipoNuevo", adVarChar, adParamInput, 3, sTipoNuevo)
'cmdTasa.Parameters.Append prmTasa
'Set prmTasa = cmdTasa.CreateParameter("ValorAnt", adCurrency, adParamInput, nValorAnt)
'cmdTasa.Parameters.Append prmTasa
'Set prmTasa = cmdTasa.CreateParameter("ValorNuevo", adCurrency, adParamInput, nValorNuevo)
'cmdTasa.Parameters.Append prmTasa
'Set prmTasa = cmdTasa.CreateParameter("Fecha", adDate, adParamInput, dFecha)
'cmdTasa.Parameters.Append prmTasa
'cmdTasa.Parameters.Refresh
'cmdTasa.ActiveConnection = dbCmact
'dbCmact.spActualizaTasa
'
'Set prmTasa = Nothing
'Set cmdTasa = Nothing
'End Sub

Public Function GetInstitucion() As Recordset
Dim sSql As String
Dim rsInst As Recordset
sSql = "Select FROM "
Set rsInst = dbCmact.CargaRecordSet(sSql)

Set GetInstitucion = rsInst
Set rsInst = Nothing
End Function

Public Function GetCapTasa(ByVal nProducto As Producto, Optional nMoneda As Moneda = gMonedaNacional, _
        Optional nTipoTasa As CaptacTipoTasa = gCapTasaNormal, Optional nPlazo As Long = 0) As Double
Dim sSql As String
Dim rsTasa As Recordset
sSql = "SELECT nTasaValor FROM CaptacTasas WHERE cTasaProd = " & nProducto & " And " _
    & "cTasaTpo = " & nTipoTasa & " And cTasaMon = " & nMoneda
If nProducto = gCapPlazoFijo Then
    sSql = sSql & " And " & nPlazo & " BETWEEN nValorIni And nValorFin"
End If
Set rsTasa = dbCmact.CargaRecordSet(sSql)
If rsTasa.EOF And rsTasa.BOF Then
    GetCapTasa = 0
Else
    GetCapTasa = rsTasa("nTasaValor")
End If
rsTasa.Close
Set rsTasa = Nothing
End Function

Public Function GetCapParametro(ByVal nParametro As CaptacParametro) As Double
Dim rsVar As Recordset
Dim sSql As String
sSql = "SELECT nParValor FROM Parametro where nParCod = " & nParametro
Set rsVar = dbCmact.CargaRecordSet(sSql)
If rsVar.EOF And rsVar.BOF Then
    GetCapParametro = 0
Else
    GetCapParametro = rsVar("nParValor")
End If
RSClose rsVar
End Function

Public Function GetCapParametroDesc(ByVal nParametro As CaptacParametro) As String
Dim rsVar As Recordset
Dim sSql As String
sSql = "SELECT cParamDesc FROM CaptacParametro WHERE cParamVar = " & nParametro
Set rsVar = dbCmact.CargaRecordSet(sSql)
If rsVar.EOF And rsVar.BOF Then
    GetCapParametroDesc = ""
Else
    GetCapParametroDesc = rsVar("cParamDesc")
End If
rsVar.Close
Set rsVar = Nothing
End Function
 
Public Function CuentasCapVigentes(ByVal dFecha As Date, ByVal nMoneda As Moneda, ByVal nProducto As Producto, ByRef nMonto As Currency, _
        Optional bConsol As Boolean = False, Optional sAgencia As String = "") As Long

Dim rsVar As Recordset
Dim sSql As String, sInicio As String, sFin As String
Dim sProducto As String

Dim rsAux As ADODB.Recordset
Set rsAux = New ADODB.Recordset

sInicio = Format$(dFecha, "mm/dd/yyyy")
sFin = Format$(DateAdd("d", 1, dFecha), "mm/dd/yyyy")

If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
    sProducto = " And nProducto IN (" & gCapPlazoFijo & "," & gCapCTS & ") "
Else
    sProducto = " And nProducto IN (" & nProducto & ") "
End If

Dim lbNoConsideraBlqAhoFijas As Boolean
lbNoConsideraBlqAhoFijas = True

'sSQL = " SELECT  isnull(Sum(nSaldCnt),0) Monto, count(*) Cantidad FROM capsaldosdiarios where datediff(day,dfecha,'" & Format(dFecha, gsFormatoFecha) & "') = 0 and ((nInmovilizada = 1) or nTpoBloqueo in (3,15) or nCuentaCredMV = '1' ) and substring(cCtaCod,6,4) = '232" & Trim(Str(nMoneda)) & "'"
'sSQL = " SELECT  isnull(Sum(nSaldCnt),0) Monto, count(*) Cantidad FROM capsaldosdiarios where dfecha >= '" & Format(dFecha, gsFormatoFecha) & "' And dfecha < '" & Format(DateAdd("d", 1, dFecha), gsFormatoFecha) & "  and ((nInmovilizada = 1) or nTpoBloqueo in (3,15) ) and substring(cCtaCod,6,4) = '232" & Trim(Str(nMoneda)) & "'"
If Not lbNoConsideraBlqAhoFijas Then
    sSql = " SELECT  isnull(Sum(nSaldCnt),0) Monto, count(*) Cantidad FROM capsaldosdiarios where dfecha >= '" & Format(dFecha, gsFormatoFecha) & "' And dfecha < '" & Format(DateAdd("d", 1, dFecha), gsFormatoFecha) & "  and ((nInmovilizada = 1) or nTpoBloqueo in (3,15) ) and substring(cCtaCod,6,4) = '232" & Trim(Str(nMoneda)) & "'"
Else
    sSql = " SELECT  0 Monto, 0 Cantidad "
End If
Set rsAux = dbCmact.CargaRecordSet(sSql)


If bConsol Then
    If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
        sSql = "Select Numero = ISNULL(Sum(nNumCtas),0)  + " & rsAux!Cantidad & "  , Monto = ISNULL(Sum(nSaldo),0)  + " & rsAux!Monto & "  From CapEstadSaldo Where " _
            & "nMoneda = " & nMoneda & sProducto & "And " _
            & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
    Else
        sSql = "Select Numero = ISNULL(Sum(nNumCtas),0)  - " & rsAux!Cantidad & "  , Monto = ISNULL(Sum(nSaldo),0)  - " & rsAux!Monto & "  From CapEstadSaldo Where " _
            & "nMoneda = " & nMoneda & sProducto & "And " _
            & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
    End If
Else
    sSql = "Select Numero = ISNULL(Sum(nNumCtas),0), Monto = ISNULL(Sum(nSaldo),0) From CapEstadSaldo Where " _
        & "nMoneda = " & nMoneda & sProducto & " And cCodAge = '" & sAgencia & "' And " _
        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
End If

Set rsVar = dbCmact.CargaRecordSet(sSql)
Set rsVar.ActiveConnection = Nothing

If rsVar.EOF And rsVar.BOF Then
    MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
    CuentasCapVigentes = 0
    nMonto = 0
Else
    CuentasCapVigentes = rsVar("Numero")
    nMonto = rsVar("Monto")
End If
rsVar.Close
Set rsVar = Nothing
End Function

Public Function CuentasCapApertura(ByVal dFecIni As Date, ByVal dFecFin As Date, _
        ByVal nMoneda As Moneda, ByVal nProducto As Producto, ByRef nMonto As Currency, _
        Optional bConsol As Boolean = False, Optional sAgencia As String = "", Optional pnTC As Double = 1) As Long

Dim rsVar As Recordset
Dim sSql As String, sInicio As String, sFin As String
Dim sProducto As String
sInicio = Format$(dFecIni, "mm/dd/yyyy")
sFin = Format$(DateAdd("d", 1, dFecFin), "mm/dd/yyyy")

If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
    sProducto = " And nProducto IN (" & gCapPlazoFijo & "," & gCapCTS & ") "
Else
    sProducto = " And nProducto IN (" & nProducto & ") "
End If

If bConsol Then
        sSql = "Select Numero = ISNULL(Sum(nNumAper),0), Monto = ISNULL(Sum(nMonAper),0) "
        sSql = sSql & " From CapEstadMovimiento "
        sSql = sSql & " Where nMoneda = " & nMoneda & sProducto
        sSql = sSql & " And dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
Else
    sSql = "Select Numero = ISNULL(Sum(nNumAper),0), Monto = ISNULL(Sum(nMonAper),0) From CapEstadMovimiento Where " _
        & "nMoneda = " & nMoneda & sProducto & " And cCodAge = '" & sAgencia & "' And " _
        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
End If


'If bConsol Then
'    sSQL = "Select "
'    sSQL = sSQL & " Numero = ISNULL(COUNT(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 1 and not c.copecod like '99%'  THEN C.nMovNro END),0),"
'    sSQL = sSQL & " Monto = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 1 and not c.copecod like '99%'  THEN Abs(CD.nMonto) END),0) "
'    sSQL = sSQL & " From  (Select nMovNro, cOpeCod From Mov Where LEFT(cMovNro,8) >='" & Format(dFecIni, "YYYYMMdd") & "' And LEFT(cMovNro,8) <='" & Format(dFecFin, "YYYYMMdd") & "' And nMovFlag = 0) M "
'    sSQL = sSQL & " INNER JOIN MovCap C ON M.nMovNro = C.nMovNro "
'    sSQL = sSQL & " INNER JOIN MovCapDet CD ON CD.nMovNro = C.nMovNro And Cd.cCtaCod = C.cCtaCod and CD.cOpeCod = C.cOpeCod "
'    sSQL = sSQL & " INNER JOIN CapMovTipo  EO on EO.cOpeCod = C.cOpeCod "
'    If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
'        sSQL = sSQL & " Where C.cCtaCod LIKE '108__23[34]" & nMoneda & "%' And EO.bEstadistica = 1 "
'    Else
'        sSQL = sSQL & " Where C.cCtaCod LIKE '108__232" & nMoneda & "%' And EO.bEstadistica = 1 "
'    End If
'Else
'    sSQL = "Select "
'    sSQL = sSQL & " Numero = ISNULL(COUNT(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 1 and not c.copecod like '99%'  THEN C.nMovNro END),0),"
'    sSQL = sSQL & " Monto = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 1 and not c.copecod like '99%'  THEN Abs(CD.nMonto) END),0) "
'    sSQL = sSQL & " From  (Select nMovNro, cOpeCod From Mov Where LEFT(cMovNro,8) >='" & Format(dFecIni, "YYYYMMdd") & "' And LEFT(cMovNro,8) <='" & Format(dFecFin, "YYYYMMdd") & "' And nMovFlag = 0) M "
'    sSQL = sSQL & " INNER JOIN MovCap C ON M.nMovNro = C.nMovNro "
'    sSQL = sSQL & " INNER JOIN MovCapDet CD ON CD.nMovNro = C.nMovNro And Cd.cCtaCod = C.cCtaCod and CD.cOpeCod = C.cOpeCod "
'    sSQL = sSQL & " INNER JOIN CapMovTipo  EO on EO.cOpeCod = C.cOpeCod "
'    If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
'        sSQL = sSQL & " Where C.cCtaCod LIKE '108" & sAgencia & "23[34]" & nMoneda & "%' And EO.bEstadistica = 1 "
'    Else
'        sSQL = sSQL & " Where C.cCtaCod LIKE '108" & sAgencia & "232" & nMoneda & "%' And EO.bEstadistica = 1 "
'    End If
'End If



'If bConsol Then
''        sSQL = "Select ISNULL(Count(m.cMovNro),0) as Numero, isnull(Sum(nMonto),0) as Monto "
''        sSQL = sSQL & " From Mov M "
''        sSQL = sSQL & "     left Join MovCap MC "
''        sSQL = sSQL & "         On M.nMovNro=MC.nMovNro "
''        If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
''            sSQL = sSQL & " Where M.cOpeCod like '2[12]01%' "
''            sSQL = sSQL & "     And left(M.cMovNro,8) >='" & Format(dFecIni, "YYYYMMdd") & "' And Left(M.cMovNro,8)<='" & Format(dFecFin, "YYYYMMdd") & "' "
''            sSQL = sSQL & "     And (MC.cCtaCod like '_____" & gCapPlazoFijo & nMoneda & "%' or MC.cCtaCod like '_____" & gCapCTS & nMoneda & "%') "
''        Else
''            sSQL = sSQL & " Where M.cOpeCod like '2001%' "
''            sSQL = sSQL & "     And left(M.cMovNro,8) >='" & Format(dFecIni, "YYYYMMdd") & "' And Left(M.cMovNro,8)<='" & Format(dFecFin, "YYYYMMdd") & "' "
''            sSQL = sSQL & "     And (MC.cCtaCod like '_____" & gCapAhorros & nMoneda & "%') "
''        End If
''        sSQL = sSQL & "     And M.nMovFlag=0 and not MC.copecod like '99%' "
'
'
'        sSQL = " Select "
'        sSQL = sSQL & "    Numero = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 1 and not c.copecod like '99%' THEN 1 ELSE 0 END),0),"
'        sSQL = sSQL & "        Monto = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 1 and not c.copecod like '99%' THEN Abs(CD.nMonto) END),0)"
'        sSQL = sSQL & "    From    MOV M"
'        sSQL = sSQL & "        INNER JOIN MovCap C ON C.nMovNro  = M.nMovNro"
'        sSQL = sSQL & "        INNER JOIN MovCapDet CD ON CD.nMovNro = C.nMovNro And Cd.cCtaCod=C.cCtaCod  AND CD.cOpeCod=C.cOpeCod"
'        sSQL = sSQL & "        INNER JOIN CapMovTipo EO ON EO.cOpeCod = C.cOpeCod"
'        If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
'            sSQL = sSQL & "    Where   (C.cCtaCod LIKE '_____" & gCapPlazoFijo & nMoneda & "%' OR  C.cCtaCod LIKE '_____" & gCapCTS & nMoneda & "%' ) And EO.bEstadistica = 1  and M.nMovFlag = 0"
'        Else
'            sSQL = sSQL & "    Where   C.cCtaCod LIKE '_____" & gCapAhorros & nMoneda & "%' And EO.bEstadistica = 1  and M.nMovFlag = 0"
'        End If
'        sSQL = sSQL & "        AND LEFT(M.CMOVNRO,8) BETWEEN convert(varchar(8),'" & Format(dFecIni, "YYYYMMdd") & "',112) AND convert(varchar(8),'" & Format(dFecFin, "YYYYMMdd") & "',112)"
'
'Else
'        sSQL = "Select ISNULL(Count(m.cMovNro),0) as Numero, ISNULL(Sum(nMonto),0) as Monto "
'        sSQL = sSQL & " From Mov M "
'        sSQL = sSQL & "     left Join MovCap MC "
'        sSQL = sSQL & "         On M.nMovNro=MC.nMovNro "
'        If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
'            sSQL = sSQL & " Where M.cOpeCod like '2[12]01%' "
'            sSQL = sSQL & "     And left(M.cMovNro,8) >='" & Format(sInicio, "YYYYMMdd") & "' And Left(M.cMovNro,8)<='" & Format(dFecIni, "YYYYMMdd") & "' "
'            sSQL = sSQL & "     And (MC.cCtaCod like '___" & sAgencia & gCapPlazoFijo & nMoneda & "%' or MC.cCtaCod like '___" & sAgencia & gCapCTS & nMoneda & "%') "
'        Else
'            sSQL = sSQL & " Where M.cOpeCod like '2001%' "
'            sSQL = sSQL & "     And left(M.cMovNro,8) >='" & Format(sInicio, "YYYYMMdd") & "' And Left(M.cMovNro,8)<='" & Format(dFecIni, "YYYYMMdd") & "' "
'            sSQL = sSQL & "     And (MC.cCtaCod like '___" & sAgencia & gCapAhorros & nMoneda & "%') "
'        End If
'        sSQL = sSQL & "     And M.nMovFlag=0 and not MC.copecod like '99%' "
'End If



Set rsVar = dbCmact.CargaRecordSet(sSql)
Set rsVar.ActiveConnection = Nothing

If rsVar.EOF And rsVar.BOF Then
    MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
    CuentasCapApertura = 0
    nMonto = 0
Else
    CuentasCapApertura = rsVar("Numero")
    nMonto = Round(rsVar("Monto") * IIf(nMoneda = gMonedaExtranjera, pnTC, 1), 2)
End If
rsVar.Close
Set rsVar = Nothing
End Function

Public Function CuentasCapAbono(ByVal dFecIni As Date, ByVal dFecFin As Date, _
        ByVal nMoneda As Moneda, ByVal nProducto As Producto, ByRef nMonto As Currency, _
        Optional bConsol As Boolean = False, Optional sAgencia As String = "", Optional pnTC As Double = 1) As Long

Dim rsVar As Recordset
Dim sSql As String, sInicio As String, sFin As String
Dim sProducto As String
sInicio = Format$(dFecIni, "mm/dd/yyyy")
sFin = Format$(DateAdd("d", 1, dFecFin), "mm/dd/yyyy")

If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
    sProducto = " And nProducto IN (" & gCapPlazoFijo & "," & gCapCTS & ") "
Else
    sProducto = " And nProducto IN (" & nProducto & ") "
End If

If bConsol Then
    sSql = "Select Numero = ISNULL(Sum(nNumDep),0), Monto = ISNULL(Sum(nMonDep),0) From CapEstadMovimiento Where " _
        & "nMoneda = " & nMoneda & sProducto & "And " _
        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
Else
    sSql = "Select Numero = ISNULL(Sum(nNumDep),0), Monto = ISNULL(Sum(nMonDep),0) From CapEstadMovimiento Where " _
        & "nMoneda = " & nMoneda & sProducto & " And cCodAge = '" & sAgencia & "' And " _
        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
End If

'If bConsol Then
''    sSQL = "Select "
''    sSQL = sSQL & " Numero = ISNULL(COUNT(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 7 and not c.copecod like '99%'  THEN C.nMovNro END),0), "
''    sSQL = sSQL & " Monto = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 7 and not c.copecod like '99%' THEN Abs(CD.nMonto) END),0) "
''    sSQL = sSQL & " From  (Select nMovNro, cOpeCod From Mov Where LEFT(cMovNro,8) >='" & Format(dFecIni, "YYYYMMdd") & "' And LEFT(cMovNro,8) <='" & Format(dFecFin, "YYYYMMdd") & "' And nMovFlag = 0) M "
''    sSQL = sSQL & " INNER JOIN MovCap C ON M.nMovNro = C.nMovNro "
''    sSQL = sSQL & " INNER JOIN MovCapDet CD ON CD.nMovNro = C.nMovNro And Cd.cCtaCod = C.cCtaCod and CD.cOpeCod = C.cOpeCod "
''    sSQL = sSQL & " INNER JOIN CapMovTipo  EO on EO.cOpeCod = C.cOpeCod "
''    If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
''        sSQL = sSQL & " Where C.cCtaCod LIKE '108__23[34]" & nMoneda & "%' And EO.bEstadistica = 1 "
''    Else
''        sSQL = sSQL & " Where C.cCtaCod LIKE '108__232" & nMoneda & "%' And EO.bEstadistica = 1 "
''    End If
'
'    sSQL = "SELECT "
'    sSQL = sSQL & " Numero = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 7 and not c.copecod like '99%' THEN 1 ELSE 0 END),0), "
'    sSQL = sSQL & " Monto = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 7 and not c.copecod like '99%' THEN Abs(CD.nMonto) END),0) "
'    sSQL = sSQL & "     From    MOV M"
'    sSQL = sSQL & "         INNER JOIN MovCap C ON C.nMovNro  = M.nMovNro"
'    sSQL = sSQL & "         INNER JOIN MovCapDet CD ON CD.nMovNro = C.nMovNro And Cd.cCtaCod=C.cCtaCod  AND CD.cOpeCod=C.cOpeCod"
'    sSQL = sSQL & "         INNER JOIN CapMovTipo EO ON EO.cOpeCod = C.cOpeCod"
'    If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
'        sSQL = sSQL & "     Where   (C.cCtaCod LIKE '_____" & gCapPlazoFijo & nMoneda & "%' OR C.cCtaCod LIKE '_____" & gCapCTS & nMoneda & "%') And EO.bEstadistica = 1  and M.nMovFlag = 0"
'    Else
'        sSQL = sSQL & "     Where   C.cCtaCod LIKE '_____232" & nMoneda & "%'  And EO.bEstadistica = 1  and M.nMovFlag = 0"
'    End If
'    sSQL = sSQL & "         AND LEFT(M.CMOVNRO,8) BETWEEN convert(varchar(8),'" & Format(dFecIni, "YYYYMMdd") & "',112) AND convert(varchar(8),'" & Format(dFecFin, "YYYYMMdd") & "',112)"
'
'
'Else
'    sSQL = "Select "
'    sSQL = sSQL & " Numero = ISNULL(COUNT(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 7 and not c.copecod like '99%'  THEN C.nMovNro END),0), "
'    sSQL = sSQL & " Monto = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 7 and not c.copecod like '99%' THEN Abs(CD.nMonto) END),0) "
'    sSQL = sSQL & " From  (Select nMovNro, cOpeCod From Mov Where LEFT(cMovNro,8) >='" & Format(dFecIni, "YYYYMMdd") & "' And LEFT(cMovNro,8) <='" & Format(dFecFin, "YYYYMMdd") & "' And nMovFlag = 0) M "
'    sSQL = sSQL & " INNER JOIN MovCap C ON M.nMovNro = C.nMovNro "
'    sSQL = sSQL & " INNER JOIN MovCapDet CD ON CD.nMovNro = C.nMovNro And Cd.cCtaCod = C.cCtaCod and CD.cOpeCod = C.cOpeCod "
'    sSQL = sSQL & " INNER JOIN CapMovTipo  EO on EO.cOpeCod = C.cOpeCod "
'    If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
'        sSQL = sSQL & " Where C.cCtaCod LIKE '108" & sAgencia & "23[34]" & nMoneda & "%' And EO.bEstadistica = 1 "
'    Else
'        sSQL = sSQL & " Where C.cCtaCod LIKE '108" & sAgencia & "232" & nMoneda & "%' And EO.bEstadistica = 1 "
'    End If
'End If

Set rsVar = dbCmact.CargaRecordSet(sSql)
Set rsVar.ActiveConnection = Nothing

If rsVar.EOF And rsVar.BOF Then
    MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
    CuentasCapAbono = 0
    nMonto = 0
Else
    CuentasCapAbono = rsVar("Numero")
    nMonto = Round(rsVar("Monto") * pnTC, 2)
End If
rsVar.Close
Set rsVar = Nothing
End Function

Public Function CuentasCapIntCapitalizado(ByVal dFecIni As Date, ByVal dFecFin As Date, _
        ByVal nMoneda As Moneda, ByVal nProducto As Producto, _
        Optional bConsol As Boolean = False, Optional sAgencia As String = "", Optional pnTC As Double = 1) As Double

Dim rsVar As Recordset
Dim sSql As String, sInicio As String, sFin As String
Dim sProducto As String
sInicio = Format$(dFecIni, "mm/dd/yyyy")
sFin = Format$(DateAdd("d", 1, dFecFin), "mm/dd/yyyy")

If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
    sProducto = " And nProducto IN (" & gCapPlazoFijo & "," & gCapCTS & ") "
Else
    sProducto = " And nProducto IN (" & nProducto & ") "
End If

If bConsol Then
    sSql = "Select Monto = ISNULL(Sum(nIntCap),0) From CapEstadMovimiento "
    sSql = sSql & " Where nMoneda = " & nMoneda & sProducto
    sSql = sSql & " And dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
Else
    sSql = "Select Monto = ISNULL(Sum(nIntCap),0) From CapEstadMovimiento "
    sSql = sSql & " Where nMoneda = " & nMoneda & sProducto
    sSql = sSql & " And cCodAge = '" & sAgencia & "' And dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
End If

'If bConsol Then
'
''    sSQL = "Select Monto=ISNULL(SUM(NMONTO),0) from movcapdet mc "
''    sSQL = sSQL & " join mov m on m.nmovnro=mc.nmovnro "
''    sSQL = sSQL & " JOIN (  SELECT DISTINCT CCTACOD "
''    sSQL = sSQL & "         FROM CAPSALDOSDIARIOS "
''    sSQL = sSQL & "         Where CONVERT(varchar(8),DFECHA,112) >='" & Format(dFecIni, "YYYYMMdd") & "' And CONVERT(varchar(8),DFECHA,112) <='" & Format(dFecFin, "YYYYMMdd") & "' "
''    If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
''        sSQL = sSQL & "             AND CCTACOD LIKE '108__23[34]" & nMoneda & "%') C ON C.CCTACOD=MC.CCTACOD "
''    Else
''        sSQL = sSQL & "             AND CCTACOD LIKE '108__" & nProducto & nMoneda & "%') C ON C.CCTACOD=MC.CCTACOD "
''    End If
''    sSQL = sSQL & " Where m.nmovflag=0 and Left(m.cmovnro,8)>='" & Format(dFecIni, "YYYYMMdd") & "' And Left(m.cMovNro,8) <='" & Format(dFecFin, "YYYYMMdd") & "' "
''    If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
''        sSQL = sSQL & " and mc.copecod IN ('210401', '210501', '220501') AND MC.NCONCEPTOCOD=1 "
''    Else
''        sSQL = sSQL & " and mc.copecod IN ('200500', '200501', '200502', '200503') AND MC.NCONCEPTOCOD=1 "
''    End If
'
'        sSQL = " Select"
'        sSQL = sSQL & "        Monto = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 8 and not c.copecod like '99%' THEN Abs(CD.nMonto) END),0)"
'        sSQL = sSQL & "        From    MOV M"
'        sSQL = sSQL & "            INNER JOIN MovCap C ON C.nMovNro  = M.nMovNro"
'        sSQL = sSQL & "            INNER JOIN MovCapDet CD ON CD.nMovNro = C.nMovNro And Cd.cCtaCod=C.cCtaCod  AND CD.cOpeCod=C.cOpeCod"
'        sSQL = sSQL & "            INNER JOIN CapMovTipo EO ON EO.cOpeCod = C.cOpeCod"
'        If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
'            sSQL = sSQL & "    Where   (C.cCtaCod LIKE '_____" & gCapPlazoFijo & nMoneda & "%' OR  C.cCtaCod LIKE '_____" & gCapCTS & nMoneda & "%' ) And EO.bEstadistica = 1  and M.nMovFlag = 0"
'        Else
'            sSQL = sSQL & "    Where   C.cCtaCod LIKE '_____" & gCapAhorros & nMoneda & "%' And EO.bEstadistica = 1  and M.nMovFlag = 0"
'        End If
'        sSQL = sSQL & "        AND LEFT(M.CMOVNRO,8) BETWEEN convert(varchar(8),'" & Format(dFecIni, "YYYYMMdd") & "',112) AND convert(varchar(8),'" & Format(dFecFin, "YYYYMMdd") & "',112)"
'
'Else
'    sSQL = "Select Monto=ISNULL(SUM(NMONTO),0) from movcapdet mc "
'    sSQL = sSQL & " join mov m on m.nmovnro=mc.nmovnro "
'    sSQL = sSQL & " JOIN (  SELECT DISTINCT CCTACOD "
'    sSQL = sSQL & "         FROM CAPSALDOSDIARIOS "
'    sSQL = sSQL & "         Where CONVERT(varchar(8),DFECHA,112) >='" & Format(dFecIni, "YYYYMMdd") & "' And CONVERT(varchar(8),DFECHA,112) <='" & Format(dFecFin, "YYYYMMdd") & "' "
'    If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
'        sSQL = sSQL & "             AND CCTACOD LIKE '108__23[34]" & nMoneda & "%') C ON C.CCTACOD=MC.CCTACOD "
'    Else
'        sSQL = sSQL & "             AND CCTACOD LIKE '108__" & nProducto & nMoneda & "%') C ON C.CCTACOD=MC.CCTACOD "
'    End If
'    sSQL = sSQL & " Where m.nmovflag=0 and Left(m.cmovnro,8)>='" & Format(dFecIni, "YYYYMMdd") & "' And Left(m.cMovNro,8) <='" & Format(dFecFin, "YYYYMMdd") & "' "
'    If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
'        sSQL = sSQL & " and mc.copecod IN ('210401', '210501', '220501') AND MC.NCONCEPTOCOD=1 "
'    Else
'        sSQL = sSQL & " and mc.copecod IN ('200500', '200501', '200502', '200503') AND MC.NCONCEPTOCOD=1 "
'    End If
'End If

Set rsVar = dbCmact.CargaRecordSet(sSql)
Set rsVar.ActiveConnection = Nothing

If rsVar.EOF And rsVar.BOF Then
    MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
    CuentasCapIntCapitalizado = 0
Else
    CuentasCapIntCapitalizado = Round(rsVar("Monto") * pnTC, 2)
End If
rsVar.Close
Set rsVar = Nothing
End Function

Public Function CuentasCapCancelada(ByVal dFecIni As Date, ByVal dFecFin As Date, _
        ByVal nMoneda As Moneda, ByVal nProducto As Producto, ByRef nMonto As Currency, _
        Optional bConsol As Boolean = False, Optional sAgencia As String = "", Optional pnTC As Double = 1) As Long

Dim rsVar As Recordset
Dim sSql As String, sInicio As String, sFin As String
Dim sProducto As String
sInicio = Format$(dFecIni, "mm/dd/yyyy")
sFin = Format$(DateAdd("d", 1, dFecFin), "mm/dd/yyyy")

If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
    sProducto = " And nProducto IN (" & gCapPlazoFijo & "," & gCapCTS & ") "
Else
    sProducto = " And nProducto IN (" & nProducto & ") "
End If

If bConsol Then
    sSql = "Select Numero = ISNULL(Sum(nNumCanc+nNumCancInac),0), Monto = ISNULL(Sum(nMonCanc+nMonCancInac),0) From CapEstadMovimiento Where " _
        & "nMoneda = " & nMoneda & sProducto & "And " _
        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
Else
    sSql = "Select Numero = ISNULL(Sum(nNumCanc+nNumCancInac),0), Monto = ISNULL(Sum(nMonCanc+nMonCancInac),0) From CapEstadMovimiento Where " _
        & "nMoneda = " & nMoneda & sProducto & " And cCodAge = '" & sAgencia & "' And " _
        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
End If

'If bConsol Then
''    sSQL = "Select "
''    sSQL = sSQL & " nNumCancAct = ISNULL(COUNT(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 2 and not c.copecod like '99%'  THEN C.nMovNro END),0),"
''    sSQL = sSQL & " nMonCancAct = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 2 and not c.copecod like '99%'  THEN Abs(CD.nMonto) END)*-1,0),"
''    sSQL = sSQL & " nNumCancInact = ISNULL(COUNT(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 3 and not c.copecod like '99%'  THEN C.nMovNro END),0),"
''    sSQL = sSQL & " nMonCancInact = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 3 and not c.copecod like '99%'  THEN Abs(CD.nMonto) END)*-1,0)"
''    sSQL = sSQL & " From  (Select nMovNro, cOpeCod From Mov Where LEFT(cMovNro,8) >='" & Format(dFecIni, "YYYYMMdd") & "' And LEFT(cMovNro,8) <='" & Format(dFecFin, "YYYYMMdd") & "' And nMovFlag = 0) M "
''    sSQL = sSQL & " INNER JOIN MovCap C ON M.nMovNro = C.nMovNro "
''    sSQL = sSQL & " INNER JOIN MovCapDet CD ON CD.nMovNro = C.nMovNro And Cd.cCtaCod = C.cCtaCod and CD.cOpeCod = C.cOpeCod "
''    sSQL = sSQL & " INNER JOIN CapMovTipo  EO on EO.cOpeCod = C.cOpeCod "
''    If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
''        sSQL = sSQL & " Where C.cCtaCod LIKE '108__23[34]" & nMoneda & "%' And EO.bEstadistica = 1 "
''    Else
''        sSQL = sSQL & " Where C.cCtaCod LIKE '108__232" & nMoneda & "%' And EO.bEstadistica = 1 "
''    End If
'
'
'        sSQL = " Select"
'        sSQL = sSQL & "    nNumCancAct = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 2 and not c.copecod like '99%' THEN 1 ELSE 0 END),0),"
'        sSQL = sSQL & "        nNumCancInact = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 3 and not c.copecod like '99%' THEN 1 ELSE 0 END),0),"
'        sSQL = sSQL & "    nMonCancAct = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 2 and not c.copecod like '99%' THEN Abs(CD.nMonto) END)*-1,0),"
'        sSQL = sSQL & "        nMonCancInact = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 3 and not c.copecod like '99%' THEN Abs(CD.nMonto) END)*-1,0)"
'        sSQL = sSQL & "From    MOV M"
'        sSQL = sSQL & "        INNER JOIN MovCap C ON C.nMovNro  = M.nMovNro"
'        sSQL = sSQL & "        INNER JOIN MovCapDet CD ON CD.nMovNro = C.nMovNro And Cd.cCtaCod=C.cCtaCod  AND CD.cOpeCod=C.cOpeCod"
'        sSQL = sSQL & "        INNER JOIN CapMovTipo EO ON EO.cOpeCod = C.cOpeCod"
'        If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
'            sSQL = sSQL & "     Where   (C.cCtaCod LIKE '_____" & gCapPlazoFijo & nMoneda & "%' OR C.cCtaCod LIKE '_____" & gCapCTS & nMoneda & "%') And EO.bEstadistica = 1  and M.nMovFlag = 0"
'        Else
'            sSQL = sSQL & "     Where   C.cCtaCod LIKE '_____232" & nMoneda & "%'  And EO.bEstadistica = 1  and M.nMovFlag = 0"
'        End If
'        sSQL = sSQL & "         AND LEFT(M.CMOVNRO,8) BETWEEN convert(varchar(8),'" & Format(dFecIni, "YYYYMMdd") & "',112) AND convert(varchar(8),'" & Format(dFecFin, "YYYYMMdd") & "',112)"
'
'Else
'    sSQL = "Select "
'    sSQL = sSQL & " nNumCancAct = ISNULL(COUNT(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 2 and not c.copecod like '99%'  THEN C.nMovNro END),0),"
'    sSQL = sSQL & " nMonCancAct = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 2 and not c.copecod like '99%'  THEN Abs(CD.nMonto) END)*-1,0),"
'    sSQL = sSQL & " nNumCancInact = ISNULL(COUNT(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 3 and not c.copecod like '99%'  THEN C.nMovNro END),0),"
'    sSQL = sSQL & " nMonCancInact = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 3 and not c.copecod like '99%'  THEN Abs(CD.nMonto) END)*-1,0)"
'    sSQL = sSQL & " From  (Select nMovNro, cOpeCod From Mov Where LEFT(cMovNro,8) >='" & Format(dFecIni, "YYYYMMdd") & "' And LEFT(cMovNro,8) <='" & Format(dFecFin, "YYYYMMdd") & "' And nMovFlag = 0) M "
'    sSQL = sSQL & " INNER JOIN MovCap C ON M.nMovNro = C.nMovNro "
'    sSQL = sSQL & " INNER JOIN MovCapDet CD ON CD.nMovNro = C.nMovNro And Cd.cCtaCod = C.cCtaCod and CD.cOpeCod = C.cOpeCod "
'    sSQL = sSQL & " INNER JOIN CapMovTipo  EO on EO.cOpeCod = C.cOpeCod "
'    If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
'        sSQL = sSQL & " Where C.cCtaCod LIKE '108" & sAgencia & "23[34]" & nMoneda & "%' And EO.bEstadistica = 1 "
'    Else
'        sSQL = sSQL & " Where C.cCtaCod LIKE '108" & sAgencia & "232" & nMoneda & "%' And EO.bEstadistica = 1 "
'    End If
'End If



Set rsVar = dbCmact.CargaRecordSet(sSql)
Set rsVar.ActiveConnection = Nothing

If rsVar.EOF And rsVar.BOF Then
    MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
    CuentasCapCancelada = 0
    nMonto = 0
Else
    CuentasCapCancelada = rsVar("Numero") '+ rsVar("nNumCancInact")
    'nMonto = Round((rsVar("nMonCancAct") + rsVar("nMonCancInact")) * pnTC, 2)
    nMonto = Round((rsVar("Monto")) * pnTC, 2)
End If
rsVar.Close
Set rsVar = Nothing
End Function

Public Function CuentasCapRetiro(ByVal dFecIni As Date, ByVal dFecFin As Date, _
        ByVal nMoneda As Moneda, ByVal nProducto As Producto, ByRef nMonto As Currency, _
        Optional bConsol As Boolean = False, Optional sAgencia As String = "", Optional pnTC As Double = 1) As Long

Dim rsVar As Recordset
Dim sSql As String, sInicio As String, sFin As String
Dim sProducto As String
sInicio = Format$(dFecIni, "mm/dd/yyyy")
sFin = Format$(DateAdd("d", 1, dFecFin), "mm/dd/yyyy")

If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
    sProducto = " And nProducto IN (" & gCapPlazoFijo & "," & gCapCTS & ") "
Else
    sProducto = " And nProducto IN (" & nProducto & ") "
End If

If bConsol Then
    sSql = "Select Numero = ISNULL(Sum(nNumRet),0), Monto = ISNULL(Sum(nRetInac+nMonRet),0) From CapEstadMovimiento Where " _
        & "nMoneda = " & nMoneda & sProducto & "And " _
        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
Else
    sSql = "Select Numero = ISNULL(Sum(nNumRet),0), Monto = ISNULL(Sum(nRetInac+nMonRet),0) From CapEstadMovimiento Where " _
        & "nMoneda = " & nMoneda & sProducto & " And cCodAge = '" & sAgencia & "' And " _
        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
End If

If bConsol Then
'    sSQL = "Select "
'    sSQL = sSQL & " nMonRetInact = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 4 and not c.copecod like '99%'  THEN Abs(CD.nMonto) END)*-1,0), "
'    sSQL = sSQL & " Numero = ISNULL(COUNT(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 6 and not c.copecod like '99%'  THEN C.nMovNro END),0), "
'    sSQL = sSQL & " nMonRet = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 6 and not c.copecod like '99%'  THEN Abs(CD.nMonto) END)*-1,0) "
'    sSQL = sSQL & " From  (Select nMovNro, cOpeCod From Mov Where LEFT(cMovNro,8) >='" & Format(dFecIni, "YYYYMMdd") & "' And LEFT(cMovNro,8) <='" & Format(dFecFin, "YYYYMMdd") & "' And nMovFlag = 0) M "
'    sSQL = sSQL & " INNER JOIN MovCap C ON M.nMovNro = C.nMovNro "
'    sSQL = sSQL & " INNER JOIN MovCapDet CD ON CD.nMovNro = C.nMovNro And Cd.cCtaCod = C.cCtaCod and CD.cOpeCod = C.cOpeCod "
'    sSQL = sSQL & " INNER JOIN CapMovTipo  EO on EO.cOpeCod = C.cOpeCod "
'    If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
'        sSQL = sSQL & " Where C.cCtaCod LIKE '108__23[34]" & nMoneda & "%' And EO.bEstadistica = 1 "
'    Else
'        sSQL = sSQL & " Where C.cCtaCod LIKE '108__232" & nMoneda & "%' And EO.bEstadistica = 1 "
'    End If

        sSql = " Select "
        sSql = sSql & "        nMonRetInact = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 4 and not c.copecod like '99%' THEN Abs(CD.nMonto) END)*-1,0),"
        sSql = sSql & "        Numero = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 6 and not c.copecod like '99%' THEN 1 ELSE 0 END),0), "
        sSql = sSql & "        nMonRet = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 6 and not c.copecod like '99%' THEN Abs(CD.nMonto) END)*-1,0)  +  "
        sSql = sSql & "                  ISNULL(SUM(CASE WHEN CD.nConceptoCod = 20  and c.copecod in ('990101','990301')  And EO.nCapMovTpo = 6 THEN Abs(CD.nMonto) END)*-1,0) "
        If nProducto = gCapAhorros Then
            sSql = sSql & "        + ISNULL(SUM(CASE WHEN CD.nConceptoCod = 2 And EO.nCapMovTpo = 5 and not c.copecod like '99%' THEN Abs(CD.nMonto) END)*-1,0) "
        End If
        sSql = sSql & "        From    MOV M"
        sSql = sSql & "            INNER JOIN MovCap C ON C.nMovNro  = M.nMovNro"
        sSql = sSql & "            INNER JOIN MovCapDet CD ON CD.nMovNro = C.nMovNro And Cd.cCtaCod=C.cCtaCod  AND CD.cOpeCod=C.cOpeCod"
        sSql = sSql & "            INNER JOIN CapMovTipo EO ON EO.cOpeCod = C.cOpeCod"
        If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
            sSql = sSql & "     Where   (C.cCtaCod LIKE '_____" & gCapPlazoFijo & nMoneda & "%' OR C.cCtaCod LIKE '_____" & gCapCTS & nMoneda & "%') And EO.bEstadistica = 1  and M.nMovFlag = 0"
        Else
            sSql = sSql & "     Where   C.cCtaCod LIKE '_____232" & nMoneda & "%'  And EO.bEstadistica = 1  and M.nMovFlag = 0"
        End If
        sSql = sSql & "         AND LEFT(M.CMOVNRO,8) BETWEEN convert(varchar(8),'" & Format(dFecIni, "YYYYMMdd") & "',112) AND convert(varchar(8),'" & Format(dFecFin, "YYYYMMdd") & "',112)"

Else
    sSql = "Select "
    sSql = sSql & " nMonRetInact = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 4 and not c.copecod like '99%'  THEN Abs(CD.nMonto) END)*-1,0), "
    sSql = sSql & " Numero = ISNULL(COUNT(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 6 and not c.copecod like '99%'  THEN C.nMovNro END),0), "
    sSql = sSql & " nMonRet = ISNULL(SUM(CASE WHEN CD.nConceptoCod = 1 And EO.nCapMovTpo = 6 and not c.copecod like '99%'  THEN Abs(CD.nMonto) END)*-1,0) "
    sSql = sSql & " From  (Select nMovNro, cOpeCod From Mov Where LEFT(cMovNro,8) >='" & Format(dFecIni, "YYYYMMdd") & "' And LEFT(cMovNro,8) <='" & Format(dFecFin, "YYYYMMdd") & "' And nMovFlag = 0) M "
    sSql = sSql & " INNER JOIN MovCap C ON M.nMovNro = C.nMovNro "
    sSql = sSql & " INNER JOIN MovCapDet CD ON CD.nMovNro = C.nMovNro And Cd.cCtaCod = C.cCtaCod and CD.cOpeCod = C.cOpeCod "
    sSql = sSql & " INNER JOIN CapMovTipo  EO on EO.cOpeCod = C.cOpeCod "
    If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
        sSql = sSql & " Where C.cCtaCod LIKE '108" & sAgencia & "23[34]" & nMoneda & "%' And EO.bEstadistica = 1 "
    Else
        sSql = sSql & " Where C.cCtaCod LIKE '108" & sAgencia & "232" & nMoneda & "%' And EO.bEstadistica = 1 "
    End If
End If


Set rsVar = dbCmact.CargaRecordSet(sSql)
Set rsVar.ActiveConnection = Nothing

If rsVar.EOF And rsVar.BOF Then
    MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
    CuentasCapRetiro = 0
    nMonto = 0
Else
    CuentasCapRetiro = Abs(rsVar("Numero"))
    nMonto = Round((rsVar("nMonRetInact") + rsVar("nMonRet")) * pnTC, 2)
End If
rsVar.Close
Set rsVar = Nothing
End Function

Public Function GetSaldoEstadistica(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal dFecha As Date, Optional bInactiva As Boolean = False, _
        Optional bInteres As Boolean = False) As Recordset

Dim rsVar As Recordset
Dim sSql As String, sInicio As String, sFin As String
Dim sSaldo As String
sInicio = Format$(dFecha, "mm/dd/yyyy")
sFin = Format$(DateAdd("d", 1, dFecha), "mm/dd/yyyy")

sSaldo = " E.nSaldo -  E.nSaldoInactEmp"
If bInactiva Then sSaldo = " E.nSaldoInactEmp"
If bInteres Then sSaldo = " E.nIntDev"

'*** PEAC 20110423

 'Variable no definida gPersonaJurCFLFoncodes
'sSQL = "SELECT A.cAgeCod, A.cAgeDescripcion, " _
    & "nSaldPN = ISNULL(SUM(CASE WHEN nPersoneria = " & gPersonaNat & " THEN " & sSaldo & " END),0), " _
    & "nSaldPSFL = ISNULL(SUM(CASE WHEN nPersoneria = " & gPersonaJurSFL & " THEN " & sSaldo & " END),0), " _
    & "nSaldPCFL = ISNULL(SUM(CASE WHEN nPersoneria = " & gPersonaJurCFL & " THEN " & sSaldo & " END),0), " _
    & "nSaldCMAC = ISNULL(SUM(CASE WHEN nPersoneria = " & gPersonaJurCFLCMAC & " THEN " & sSaldo & " END),0), " _
    & "nSaldCRAC = ISNULL(SUM(CASE WHEN nPersoneria = " & gPersonaJurCFLCRAC & " THEN " & sSaldo & " END),0), " _
    & "nSaldFoncodes = ISNULL(SUM(CASE WHEN nPersoneria = " & gPersonaJurCFLFONCODES & " THEN " & sSaldo & " END),0) " _
    & "FROM Agencias A LEFT JOIN CapEstadSaldo E ON A.cAgeCod = E.cCodAge WHERE nMoneda = " & nMoneda & " " _
    & "And nProducto = " & nProducto & " And  dEstad  >= '" & sInicio & "' And  dEstad  < '" & sFin & "' " _
    & "GROUP BY A.cAgeCod, A.cAgeDescripcion"

'sSQL = "SELECT A.cAgeCod, A.cAgeDescripcion, " _
    & "nSaldPN = ISNULL(SUM(CASE WHEN nPersoneria = " & gPersonaNat & " THEN " & sSaldo & " END),0), " _
    & "nSaldPSFL = ISNULL(SUM(CASE WHEN nPersoneria = " & gPersonaJurSFL & " THEN " & sSaldo & " END),0), " _
    & "nSaldPCFL = ISNULL(SUM(CASE WHEN nPersoneria = " & gPersonaJurCFL & " THEN " & sSaldo & " END),0), " _
    & "nSaldCMAC = ISNULL(SUM(CASE WHEN nPersoneria = " & gPersonaJurCFLCMAC & " THEN " & sSaldo & " END),0), " _
    & "nSaldCRAC = ISNULL(SUM(CASE WHEN nPersoneria in (" & gPersonaJurCFLCRAC & ") THEN " & sSaldo & " END),0), " _
    & "nSaldBanco = ISNULL(SUM(CASE WHEN nPersoneria in (6) THEN " & sSaldo & " END),0), " _
    & "nSaldEdpym = ISNULL(SUM(CASE WHEN nPersoneria in (" & gPersonaJurCFLEdpyme & ") THEN " & sSaldo & " END),0) " _
    & "FROM Agencias A LEFT JOIN CapEstadSaldo E ON A.cAgeCod = E.cCodAge And nMoneda = " & nMoneda & " " _
    & "And nProducto = " & nProducto & " And  dEstad  >= '" & sInicio & "' And  dEstad  < '" & sFin & "' " _
    & "GROUP BY A.cAgeCod, A.cAgeDescripcion"


sSql = "exec stp_sel_GetSaldoEstadistica " & nProducto & ",'" & sInicio & "','" & sFin & "'," & nMoneda & ",'" & IIf(bInactiva, "1", "0") & "','" & IIf(bInteres, "1", "0") & "'"

'*** FIN PEAC

Set rsVar = dbCmact.CargaRecordSet(sSql)
Set rsVar.ActiveConnection = Nothing
Set GetSaldoEstadistica = rsVar
Set rsVar = Nothing
End Function

'Fin Agregado

'''''Modificado 09.02.2005
''''Public Function CuentasCapVigentes(ByVal dFecha As Date, ByVal nMoneda As Moneda, ByVal nProducto As Producto, ByRef nMonto As Currency, _
''''        Optional bConsol As Boolean = False, Optional sAgencia As String = "") As Long
''''
''''Dim rsVar As Recordset
''''Dim sSql As String, sInicio As String, sFin As String
''''Dim sProducto As String
''''Dim sProdCad As String
''''Dim sSigno As String
''''Dim nNumCta   As Currency, nMonCta As Currency
''''
''''sInicio = Format$(dFecha, "mm/dd/yyyy")
''''sFin = Format$(DateAdd("d", 1, dFecha), "mm/dd/yyyy")
''''
''''If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
''''    sProducto = " And nProducto IN (" & gCapPlazoFijo & "," & gCapCTS & ") "
''''    sProdCad = " IN ('" & gCapPlazoFijo & "', '" & gCapCTS & "') "
''''    sSigno = "+"
''''Else
''''    sSigno = "-"
''''    sProducto = " And nProducto IN (" & nProducto & ") "
''''    sProdCad = " IN ('" & nProducto & "') "
''''End If
''''
''''
''''sSql = " Select isnull(count(*),0) nNumCtas, isnull(sum(nSaldCnt),0) nSaldo From capsaldosdiarios" _
''''     & " where convert(varchar(10),dfecha,101) = '" & sInicio & "' " _
''''     & "   and substring(cCtaCod,6,3) " & sProdCad & " and substring(cCtaCod,9,1) = " & nMoneda & " " _
''''     & " "
''''
''''
''''Set rsVar = dbCmact.CargaRecordSet(sSql)
''''If Not rsVar.EOF Then
''''   nNumCta = rsVar!nNumCtas
''''   nMonCta = rsVar!nSaldo
''''End If
''''RSClose rsVar
''''
''''
''''If bConsol Then
''''    sSql = "Select Numero = ISNULL(Sum(nNumCtas),0) " & sSigno & nNumCta & ", Monto = ISNULL(Sum(nSaldo),0)" & sSigno & nMonCta & " From CapEstadSaldo Where " _
''''        & "nMoneda = " & nMoneda & sProducto & "And " _
''''        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
''''Else
''''    sSql = "Select Numero = ISNULL(Sum(nNumCtas),0) " & sSigno & nNumCta & ", Monto = ISNULL(Sum(nSaldo),0)" & sSigno & nMonCta & " From CapEstadSaldo Where " _
''''        & "nMoneda = " & nMoneda & sProducto & " And cCodAge = '" & sAgencia & "' And " _
''''        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
''''End If
''''
''''Set rsVar = dbCmact.CargaRecordSet(sSql)
''''Set rsVar.ActiveConnection = Nothing
''''
''''If rsVar.EOF And rsVar.BOF Then
''''    MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
''''    CuentasCapVigentes = 0
''''    nMonto = 0
''''Else
''''    CuentasCapVigentes = rsVar("Numero")
''''    nMonto = rsVar("Monto")
''''End If
''''rsVar.Close
''''Set rsVar = Nothing
''''End Function
''''
'''''''Public Function CuentasDepGarantia(ByVal dFecIni As Date, ByVal dFecFin As Date, _
'''''''        ByVal nMoneda As Moneda, ByVal nProducto As Producto, _
'''''''        Optional sAgencia As String = "", Optional pnTC As Double = 1) As Long
'''''''
'''''''Dim sSql As String, sInicio As String, sFin As String
'''''''Dim rsVar As ADODB.Recordset
'''''''sInicio = Format$(dFecIni - 1, "yyyymmdd")
'''''''sFin = Format$(dFecFin, "yyyymmdd")
'''''''sSql = "select isnull(sum(saldoact - saldoant),0) nSaldo from ( "
'''''''sSql = sSql & " select 0 SaldoAct, sum(nsaldcnt) SaldoAnt"
'''''''sSql = sSql & " from capsaldosdiarios where bgarantia = 1"
'''''''sSql = sSql & " and convert(Varchar(8),dfecha,112) = '" & sInicio & "'"
'''''''sSql = sSql & " and substring(cctacod,6,3) = '" & nProducto & "' and substring(cCtaCod,9,1) = " & nMoneda & IIf(sAgencia <> "", " and SubString(cCtaCod,4,2) = '" & sAgencia & "' ", "")
'''''''sSql = sSql & " Union All"
'''''''sSql = sSql & " select sum(nsaldcnt) SaldoAct, 0 SaldoAnt"
'''''''sSql = sSql & " from capsaldosdiarios where bgarantia = 1"
'''''''sSql = sSql & " and convert(Varchar(8),dfecha,112) = '" & sFin & "'"
'''''''sSql = sSql & " and substring(cctacod,6,3) = '" & nProducto & "' and substring(cCtaCod,9,1) = " & nMoneda & IIf(sAgencia <> "", " and SubString(cCtaCod,4,2) = '" & sAgencia & "' ", "")
'''''''sSql = sSql & " ) Sdo "
'''''''
'''''''Set rsVar = dbCmact.CargaRecordSet(sSql)
'''''''If Not rsVar.EOF Then
'''''''   CuentasDepGarantia = Round(rsVar!nSaldo * pnTC, 2)
'''''''End If
'''''''End Function
''''
''''Public Function CuentasCapApertura(ByVal dFecIni As Date, ByVal dFecFin As Date, _
''''        ByVal nMoneda As Moneda, ByVal nProducto As Producto, ByRef nMonto As Currency, _
''''        Optional bConsol As Boolean = False, Optional sAgencia As String = "", Optional pnTC As Double = 1) As Long
''''
''''Dim rsVar As Recordset
''''Dim sSql As String, sInicio As String, sFin As String
''''Dim sProducto As String
''''sInicio = Format$(dFecIni, "yyyymmdd")
''''sFin = Format$(DateAdd("d", 1, dFecFin), "yyyymmdd")
''''
''''If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
''''    sProducto = " And nProducto IN (" & gCapPlazoFijo & "," & gCapCTS & ") "
''''Else
''''    sProducto = " And nProducto IN (" & nProducto & ") "
''''End If
''''
''''If bConsol Then
''''    sSql = "Select Numero = ISNULL(Sum(nNumAper),0), Monto = ISNULL(Sum(nMonAper),0) From CapEstadMovimiento Where " _
''''        & "nMoneda = " & nMoneda & sProducto & "And " _
''''        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
''''Else
''''    sSql = "Select Numero = ISNULL(Sum(nNumAper),0), Monto = ISNULL(Sum(nMonAper),0) From CapEstadMovimiento Where " _
''''        & "nMoneda = " & nMoneda & sProducto & " And cCodAge = '" & sAgencia & "' And " _
''''        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
''''End If
''''
''''Set rsVar = dbCmact.CargaRecordSet(sSql)
''''Set rsVar.ActiveConnection = Nothing
''''
''''If rsVar.EOF And rsVar.BOF Then
''''    MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
''''    CuentasCapApertura = 0
''''    nMonto = 0
''''Else
''''    CuentasCapApertura = rsVar("Numero")
''''    nMonto = Round(rsVar("Monto") * IIf(nMoneda = gMonedaExtranjera, pnTC, 1), 2)
''''End If
''''rsVar.Close
''''Set rsVar = Nothing
''''End Function
''''
''''Public Function CuentasCapAbono(ByVal dFecIni As Date, ByVal dFecFin As Date, _
''''        ByVal nMoneda As Moneda, ByVal nProducto As Producto, ByRef nMonto As Currency, _
''''        Optional bConsol As Boolean = False, Optional sAgencia As String = "", Optional pnTC As Double = 1) As Long
''''
''''Dim rsVar As Recordset
''''Dim sSql As String, sInicio As String, sFin As String
''''Dim sProducto As String
''''sInicio = Format$(dFecIni, "mm/dd/yyyy")
''''sFin = Format$(DateAdd("d", 1, dFecFin), "mm/dd/yyyy")
''''
''''If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
''''    sProducto = " And nProducto IN (" & gCapPlazoFijo & "," & gCapCTS & ") "
''''Else
''''    sProducto = " And nProducto IN (" & nProducto & ") "
''''End If
''''
''''If bConsol Then
''''    sSql = "Select Numero = ISNULL(Sum(nNumDep),0), Monto = ISNULL(Sum(nMonDep),0) From CapEstadMovimiento Where " _
''''        & "nMoneda = " & nMoneda & sProducto & "And " _
''''        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
''''Else
''''    sSql = "Select Numero = ISNULL(Sum(nNumDep),0), Monto = ISNULL(Sum(nMonDep),0) From CapEstadMovimiento Where " _
''''        & "nMoneda = " & nMoneda & sProducto & " And cCodAge = '" & sAgencia & "' And " _
''''        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
''''End If
''''
''''Set rsVar = dbCmact.CargaRecordSet(sSql)
''''Set rsVar.ActiveConnection = Nothing
''''
''''If rsVar.EOF And rsVar.BOF Then
''''    MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
''''    CuentasCapAbono = 0
''''    nMonto = 0
''''Else
''''    CuentasCapAbono = rsVar("Numero")
''''    nMonto = Round(rsVar("Monto") * pnTC, 2)
''''End If
''''rsVar.Close
''''Set rsVar = Nothing
''''End Function
''''
''''Public Function CuentasCapIntCapitalizado(ByVal dFecIni As Date, ByVal dFecFin As Date, _
''''        ByVal nMoneda As Moneda, ByVal nProducto As Producto, _
''''        Optional bConsol As Boolean = False, Optional sAgencia As String = "", Optional pnTC As Double = 1) As Double
''''
''''Dim rsVar As Recordset
''''Dim sSql As String, sInicio As String, sFin As String
''''Dim sProducto As String
''''sInicio = Format$(dFecIni, "mm/dd/yyyy")
''''sFin = Format$(DateAdd("d", 1, dFecFin), "mm/dd/yyyy")
''''
''''If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
''''    sProducto = " And nProducto IN (" & gCapPlazoFijo & "," & gCapCTS & ") "
''''Else
''''    sProducto = " And nProducto IN (" & nProducto & ") "
''''End If
''''
''''If bConsol Then
''''    sSql = "Select Monto = ISNULL(Sum(nIntCap),0) From CapEstadMovimiento Where " _
''''        & "nMoneda = " & nMoneda & sProducto & "And " _
''''        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
''''Else
''''    sSql = "Select Monto = ISNULL(Sum(nIntCap),0) From CapEstadMovimiento Where " _
''''        & "nMoneda = " & nMoneda & sProducto & " And cCodAge = '" & sAgencia & "' And " _
''''        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
''''End If
''''
''''Set rsVar = dbCmact.CargaRecordSet(sSql)
''''Set rsVar.ActiveConnection = Nothing
''''
''''If rsVar.EOF And rsVar.BOF Then
''''    MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
''''    CuentasCapIntCapitalizado = 0
''''Else
''''    CuentasCapIntCapitalizado = Round(rsVar("Monto") * pnTC, 2)
''''End If
''''rsVar.Close
''''Set rsVar = Nothing
''''End Function
''''
''''Public Function CuentasCapCancelada(ByVal dFecIni As Date, ByVal dFecFin As Date, _
''''        ByVal nMoneda As Moneda, ByVal nProducto As Producto, ByRef nMonto As Currency, _
''''        Optional bConsol As Boolean = False, Optional sAgencia As String = "", Optional pnTC As Double = 1) As Long
''''
''''Dim rsVar As Recordset
''''Dim sSql As String, sInicio As String, sFin As String
''''Dim sProducto As String
''''sInicio = Format$(dFecIni, "mm/dd/yyyy")
''''sFin = Format$(DateAdd("d", 1, dFecFin), "mm/dd/yyyy")
''''
''''If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
''''    sProducto = " And nProducto IN (" & gCapPlazoFijo & "," & gCapCTS & ") "
''''Else
''''    sProducto = " And nProducto IN (" & nProducto & ") "
''''End If
''''
''''If bConsol Then
''''    sSql = "Select Numero = ISNULL(Sum(nNumCanc+nNumCancInac),0), Monto = ISNULL(Sum(nMonCanc+nMonCancInac),0) From CapEstadMovimiento Where " _
''''        & "nMoneda = " & nMoneda & sProducto & "And " _
''''        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
''''Else
''''    sSql = "Select Numero = ISNULL(Sum(nNumCanc+nNumCancInac),0), Monto = ISNULL(Sum(nMonCanc+nMonCancInac),0) From CapEstadMovimiento Where " _
''''        & "nMoneda = " & nMoneda & sProducto & " And cCodAge = '" & sAgencia & "' And " _
''''        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
''''End If
''''
''''Set rsVar = dbCmact.CargaRecordSet(sSql)
''''Set rsVar.ActiveConnection = Nothing
''''
''''If rsVar.EOF And rsVar.BOF Then
''''    MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
''''    CuentasCapCancelada = 0
''''    nMonto = 0
''''Else
''''    CuentasCapCancelada = rsVar("Numero")
''''    nMonto = Round(rsVar("Monto") * pnTC, 2)
''''End If
''''rsVar.Close
''''Set rsVar = Nothing
''''End Function
''''
''''Public Function CuentasCapRetiro(ByVal dFecIni As Date, ByVal dFecFin As Date, _
''''        ByVal nMoneda As Moneda, ByVal nProducto As Producto, ByRef nMonto As Currency, _
''''        Optional bConsol As Boolean = False, Optional sAgencia As String = "", Optional pnTC As Double = 1) As Long
''''
''''Dim rsVar As Recordset
''''Dim sSql As String, sInicio As String, sFin As String
''''Dim sProducto As String
''''sInicio = Format$(dFecIni, "mm/dd/yyyy")
''''sFin = Format$(DateAdd("d", 1, dFecFin), "mm/dd/yyyy")
''''
''''If nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
''''    sProducto = " And nProducto IN (" & gCapPlazoFijo & "," & gCapCTS & ") "
''''Else
''''    sProducto = " And nProducto IN (" & nProducto & ") "
''''End If
''''
''''If bConsol Then
''''    sSql = "Select Numero = ISNULL(Sum(nNumRet),0), Monto = ISNULL(Sum(nRetInac+nMonRet),0) From CapEstadMovimiento Where " _
''''        & "nMoneda = " & nMoneda & sProducto & "And " _
''''        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
''''Else
''''    sSql = "Select Numero = ISNULL(Sum(nNumRet),0), Monto = ISNULL(Sum(nRetInac+nMonRet),0) From CapEstadMovimiento Where " _
''''        & "nMoneda = " & nMoneda & sProducto & " And cCodAge = '" & sAgencia & "' And " _
''''        & "dEstad >= '" & sInicio & "' And dEstad < '" & sFin & "'"
''''End If
''''
''''Set rsVar = dbCmact.CargaRecordSet(sSql)
''''Set rsVar.ActiveConnection = Nothing
''''
''''If rsVar.EOF And rsVar.BOF Then
''''    MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
''''    CuentasCapRetiro = 0
''''    nMonto = 0
''''Else
''''    CuentasCapRetiro = rsVar("Numero")
''''    nMonto = Round(rsVar("Monto") * pnTC, 2)
''''End If
''''rsVar.Close
''''Set rsVar = Nothing
''''End Function
''''
''''Public Function GetSaldoEstadistica(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
''''        ByVal dFecha As Date, Optional bInactiva As Boolean = False, _
''''        Optional bInteres As Boolean = False) As Recordset
''''
''''Dim rsVar As Recordset
''''Dim sSql As String, sInicio As String, sFin As String
''''Dim sSaldo As String
''''sInicio = Format$(dFecha, "mm/dd/yyyy")
''''sFin = Format$(DateAdd("d", 1, dFecha), "mm/dd/yyyy")
''''
''''sSaldo = " E.nSaldo"
''''If bInactiva Then sSaldo = " E.nSaldoInactEmp"
''''If bInteres Then sSaldo = " E.nIntDev"
''''
''''sSql = "SELECT A.cAgeCod, A.cAgeDescripcion, " _
''''    & "nSaldPN = ISNULL(SUM(CASE WHEN nPersoneria = " & gPersonaNat & " THEN " & sSaldo & " END),0), " _
''''    & "nSaldPSFL = ISNULL(SUM(CASE WHEN nPersoneria = " & gPersonaJurSFL & " THEN " & sSaldo & " END),0), " _
''''    & "nSaldPCFL = ISNULL(SUM(CASE WHEN nPersoneria = " & gPersonaJurCFL & " THEN " & sSaldo & " END),0), " _
''''    & "nSaldCMAC = ISNULL(SUM(CASE WHEN nPersoneria = " & gPersonaJurCFLCMAC & " THEN " & sSaldo & " END),0), " _
''''    & "nSaldCRAC = ISNULL(SUM(CASE WHEN nPersoneria = " & gPersonaJurCFLCRAC & " THEN " & sSaldo & " END),0), " _
''''    & "nSaldFoncodes = ISNULL(SUM(CASE WHEN nPersoneria = " &  & " THEN " & sSaldo & " END),0) " _
''''    & "FROM Agencias A LEFT JOIN CapEstadSaldo E ON A.cAgeCod = E.cCodAge WHERE nMoneda = " & nMoneda & " " _
''''    & "And nProducto = " & nProducto & " And  dEstad  >= '" & sInicio & "' And  dEstad  < '" & sFin & "' " _
''''    & "GROUP BY A.cAgeCod, A.cAgeDescripcion"
''''
''''Set rsVar = dbCmact.CargaRecordSet(sSql)
''''Set rsVar.ActiveConnection = Nothing
''''Set GetSaldoEstadistica = rsVar
''''Set rsVar = Nothing
''''End Function

'Fin Agregado

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim ClsIni As DClassIni
Set ClsIni = New DClassIni
sDBComunes = ClsIni.BaseComunes
sDBPersona = ClsIni.BasePersonas
sDBImagenes = ClsIni.BaseImagenes
Set dbCmact = New DConecta
dbCmact.AbreConexion
Set ClsIni = Nothing
End Sub

Private Sub Class_Terminate()
dbCmact.CierraConexion
End Sub
'FRHU 20140107
Public Function GetMovTarjetaDebito(ByVal pdFecha As Date) As Recordset

Dim rsVar As Recordset
Dim sSql As String

sSql = "exec stp_sel_MovimientoTarjetaDebito '" & CStr(pdFecha) & "'"
Set rsVar = dbCmact.CargaRecordSet(sSql)
Set rsVar.ActiveConnection = Nothing
Set GetMovTarjetaDebito = rsVar
Set rsVar = Nothing
End Function
'FIN FRHU 20140107
'ALPA 20150427***********
Public Function ObtenerDatosAhorrosRestringidos(ByVal psMoneda As String, ByVal psFecha As String) As Recordset

Dim rsVar As Recordset
Dim sSql As String

sSql = "exec stp_sel_ObtenerDatosAhorrosRestringidos '" & psMoneda & "','" & psFecha & "'"
Set rsVar = dbCmact.CargaRecordSet(sSql)
Set rsVar.ActiveConnection = Nothing
Set ObtenerDatosAhorrosRestringidos = rsVar
Set rsVar = Nothing
End Function
'**********************/
