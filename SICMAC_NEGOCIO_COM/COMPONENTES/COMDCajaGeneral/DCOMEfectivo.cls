VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsConexion As String
Dim vsServerCom As String
Dim vsServerPers As String
Dim vsFormatoFecha As String

Public Sub inicio(ByVal psFormatoFecha As String)
vsFormatoFecha = psFormatoFecha
End Sub

Public Function EmiteBilletajes(ByVal psMonedas As COMDConstantes.Moneda, Optional ByVal psDenominacion As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta
Dim lsFiltroDenom As String
Set oConect = New COMConecta.DCOMConecta
If psDenominacion <> "" Then
    'By capi 23032009 se modifico para que no visualice anterior billetaje
    lsFiltroDenom = " and Substring(cEfectivoCod,2,1)='" & psDenominacion & "' "
    'lsFiltroDenom = " and nEfectivoValor=0 and Substring(cEfectivoCod,2,1)='" & psDenominacion & "' "
    '
End If
If oConect.AbreConexion = False Then Exit Function

sql = " Select     Case " _
    & "         WHEN Substring(cEfectivoCod,2,1)='B'  Then 'BILLETAJE' + Convert(char(10),nEfectivoValor) " _
    & "         WHEN Substring(cEfectivoCod,2,1)='M'  Then 'MONEDA   ' + Convert(char(10),nEfectivoValor) " _
    & "         END as Descripcion, 0 as Cantidad, 0  as Monto, cEfectivoCod, nEfectivoValor " _
    & " From Efectivo " _
    & " WHERE SUBSTRING(cEfectivoCod,1,1) ='" & psMonedas & "' " & lsFiltroDenom _
    & " ORDER BY  SUBSTRING(cEfectivoCod,2,1), nEfectivoValor DESC "
    
Set rs = oConect.CargaRecordSet(sql)
Set EmiteBilletajes = rs

oConect.CierraConexion: Set oConect = Nothing
End Function


Private Sub Class_Initialize()
Dim oIni As COMConecta.DCOMClasIni
Set oIni = New COMConecta.DCOMClasIni
vsConexion = oIni.CadenaConexion
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing
End Sub
Public Function GetBilletajes(ByVal psFiltro As String) As ADODB.Recordset
Dim sql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim lsFiltro As String

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function
If psFiltro <> "" Then
    lsFiltro = " WHERE cEfecTivoCod like '" & psFiltro & "%'"
End If

sql = "   SELECT  cEfectivoCod, " _
    & "         CASE " _
    & "         WHEN SUBSTRING(cEfectivoCod,2,1)='B' THEN 'BILLETES ' + Convert(Char(10),nEfectivoValor) " _
    & "         WHEN SUBSTRING(cEfectivoCod,2,1)='M' THEN 'MONEDAS  ' + Convert(Char(10),nEfectivoValor) " _
    & "         END AS Descripcion , 1  as Nivel " _
    & "  From " & vsServerCom & "EFECTIVO " & lsFiltro _
    & "  ORDER BY SUBSTRING(cEfectivoCod,2,1), nEfectivoValor DESC "

Set rs = oConect.CargaRecordSet(sql)
Set GetBilletajes = rs
oConect.CierraConexion: Set oConect = Nothing

End Function

Public Function GetBilletajesMov(ByVal psMovNro As String, Optional ByVal pnMoneda As COMDConstantes.Moneda, Optional ByVal psDenominacion As String = "")
Dim sql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim lsFiltroDenom As String

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

If psDenominacion <> "" Then
    lsFiltroDenom = " and Substring(E.cEfectivoCod,2,1)='" & psDenominacion & "' "
End If

sql = "  SELECT Case " _
    & "             WHEN Substring(E.cEfectivoCod,2,1)='B'  Then 'BILLETAJE' + Convert(char(10),E.nEfectivoValor) " _
    & "             WHEN Substring(E.cEfectivoCod,2,1)='M'  Then 'MONEDA   ' + Convert(char(10),E.nEfectivoValor) " _
    & "         END as Descripcion ,  ISNULL(MovEfect.Cantidad,0) as Cantidad , ISNULL(MovEfect.Importe,0) as Monto, " _
    & "         E.cEfectivoCod , E.nEfectivoValor " _
    & "  FROM   EFECTIVO E " _
    & "         Left Join " _
    & "         (SELECT ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE))/E.nEfectivoValor  AS CANTIDAD, " _
    & "                 ABS(ISNULL(ME.NMOVMEIMPORTE,MC.NMOVIMPORTE)) AS IMPORTE, MOE.cEfectivoCod, E.nEfectivoValor , M.cOpeCod " _
    & "          FROM   MOV M JOIN MOVCTA MC ON MC.nMovNro = M.nMovNro  " _
    & "                 LEFT JOIN MOVME ME ON ME.nMovNro = MC.nMovNro AND ME.nMovItem = MC.nMovItem " _
    & "                 JOIN MOVOBJ MO ON MO.nMovNro =MC.nMovNro AND MO.nMovItem = MC.nMovItem " _
    & "                 JOIN MOVOBJEFECTIVO MOE ON MOE.nMovNro = MO.nMovNro AND MO.nMovItem = MOE.nMovItem AND MO.nMovObjOrden= MOE.nMovObjOrden " _
    & "                 JOIN EFECTIVO E ON E.cEfectivoCod = MOE.cEfectivoCod " _
    & "          WHERE   M.CMOVNRO= '" & psMovNro & "') AS MovEfect " _
    & "     ON MovEfect.cEfectivoCod = E.cEfectivoCod " _
    & " WHERE E.cEfectivoCod LIKE '" & pnMoneda & "%' " & lsFiltroDenom & " Order BY E.nEfectivoValor Desc "

Set rs = oConect.CargaRecordSet(sql)
Set GetBilletajesMov = rs
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetEfectivoMovImp(ByVal psMovNro As String, Optional ByVal pbDebe As Boolean = True) As ADODB.Recordset
Dim sql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim lsFiltro As String

Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

sql = "SELECT  CMOVNRO, " _
    & "         CASE WHEN Tipo ='B' THEN nEfectivoValor  ELSE 0 END AS Billetes, " _
    & "         SUM(CASE WHEN Tipo = 'B' THEN ABS(Importe) ELSE 0 END ) AS MontoBilletes, " _
    & "         CASE WHEN Tipo ='M' THEN nEfectivoValor  ELSE 0 END AS Monedas, " _
    & "         SUM(CASE WHEN Tipo = 'M' THEN ABS(Importe) ELSE 0 END ) AS MontoMonedas , Moneda " _
    & " FROM (SELECT    M.CMOVNRO, E.nEfectivoValor,Substring(e.cefectivocod,2,1) AS  Tipo, Substring(e.cefectivocod,1,1) AS Moneda ,  " _
    & "                 ISNULL(Me.nMovMEImporte, MC.nMovImporte) As Importe " _
    & "        FROM     MOVOBJEFECTIVO MOE JOIN EFECTIVO E ON E.CEFECTIVOCOD = MOE.CEFECTIVOCOD  " _
    & "                 JOIN MOVCTA MC ON MC.NMOVNRO = MOE.NMOVNRO AND MC.NMOVITEM = MOE.NMOVITEM JOIN MOV M ON M.NMOVNRO =MC.NMOVNRO  " _
    & "                 LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO  AND ME.NMOVITEM = MC.NMOVITEM " _
    & "         WHERE  M.CMOVNRO ='" & psMovNro & "'  AND MC.NMOVIMPORTE" & IIf(pbDebe, ">", "<") & "0  ) AS MC " _
    & " GROUP BY CMOVNRO , nEfectivoValor , Tipo , Moneda " _
    & " ORDER BY nEfectivoValor "

Set rs = oConect.CargaRecordSet(sql)
Set GetEfectivoMovImp = rs
oConect.CierraConexion: Set oConect = Nothing
End Function

Public Function GetSaldoEfectivo(ByVal sAgencia As String, ByVal dFecha As Date, _
            ByVal nMoneda As COMDConstantes.Moneda, pdFechaDia As Date, pgsCodUser As String, Optional sUsuario As String = "", Optional bUltimo As Boolean = False) As Double

Dim rsSaldo As ADODB.Recordset
Dim ssql As String
Dim oConect As COMConecta.DCOMConecta
Dim sCondicion As String, sUltimo As String

Dim lsTablaDiaria As String

If dFecha = pdFechaDia Then
    lsTablaDiaria = " MovDiario "
Else
    lsTablaDiaria = " Mov "
End If

If bUltimo Then
    sUltimo = " And LEFT(M.cMovNro,8) IN (Select MAX(LEFT(M1.cMovNro,8)) From " & lsTablaDiaria & " M1 JOIN MovUserEfectivo E1 ON " _
        & "M1.nMovNro = E1.nMovNro Where M1.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & ") " _
        & "And LEFT(M1.cMovNro,8) <= '" & Format$(dFecha, "yyyymmdd") & "' AND Substring(M1.cMovNro,18,2) = '" & sAgencia & "' AND M1.cOpeCod in ('901007', '901016') ) "
        'MAVM 27102009 Se modifico la subconsulta detectado por JARO (M.cOpeCod)
    sCondicion = ""
Else
    sUltimo = ""
    sCondicion = " And M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "_________" & sAgencia & "%' "
End If

ssql = "Select ISNULL(SUM(E.nMonto),0) nSaldo From " & lsTablaDiaria & " M INNER JOIN MovUserEfectivo E " _
    & "ON M.nMovNro = E.nMovNro Where E.cEfectivoCod LIKE '" & nMoneda & "%' And SUBSTRING(M.cMovNro,18,2) = '" & sAgencia & "'" _
    & "AND M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & ") " & sCondicion & sUltimo

If sUsuario <> "" Then
    ssql = ssql & " And E.cUser = '" & sUsuario & "'"
End If

ssql = ssql & " AND M.cOpeCod in ('901007', '901016') "

Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

Set rsSaldo = oConect.CargaRecordSet(ssql)
GetSaldoEfectivo = rsSaldo("nSaldo")
rsSaldo.Close
Set rsSaldo = Nothing
oConect.CierraConexion
Set oConect = Nothing
End Function

'MADM 20110124
Public Function GetSaldoEfectivoPreCuadre(ByVal sAgencia As String, ByVal dFecha As Date, _
            ByVal nMoneda As COMDConstantes.Moneda, pdFechaDia As Date, pgsCodUser As String, Optional sUsuario As String = "", Optional bUltimo As Boolean = False) As Double

Dim rsSaldo As ADODB.Recordset
Dim ssql As String
Dim oConect As COMConecta.DCOMConecta
Dim sCondicion As String, sUltimo As String

Dim lsTablaDiaria As String

If dFecha = pdFechaDia Then
    lsTablaDiaria = " MovDiario "
Else
    lsTablaDiaria = " Mov "
End If

If nMoneda = gMonedaNacional Then

    If bUltimo Then
        sUltimo = " And LEFT(M.cMovNro,8) IN (Select MAX(LEFT(M1.cMovNro,8)) From " & lsTablaDiaria & " M1 JOIN MovPreCuadre E1 ON " _
            & "M1.nMovNro = E1.nMovNro Where M1.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & ") " _
            & "And LEFT(M1.cMovNro,8) <= '" & Format$(dFecha, "yyyymmdd") & "' AND Substring(M1.cMovNro,18,2) = '" & sAgencia & "' AND M1.cOpeCod in ('901040') ) "
        sCondicion = ""
    Else
        sUltimo = ""
        sCondicion = " And M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "_________" & sAgencia & "%' "
    End If
    
    ssql = "Select top 1 E.nMovNro, E.nMontoSol nSaldo From " & lsTablaDiaria & " M INNER JOIN MovPreCuadre E " _
        & "ON M.nMovNro = E.nMovNro Where SUBSTRING(M.cMovNro,18,2) = '" & sAgencia & "'" _
        & "AND M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & ") " & sCondicion & sUltimo
    
    If sUsuario <> "" Then
        ssql = ssql & " And right(M.cMovNro,4) = '" & sUsuario & "'"
    End If
    
    ssql = ssql & " AND M.cOpeCod in ('901040') Order by nMovNro desc "

Else

    If bUltimo Then
        sUltimo = " And LEFT(M.cMovNro,8) IN (Select MAX(LEFT(M1.cMovNro,8)) From " & lsTablaDiaria & " M1 JOIN MovPreCuadre E1 ON " _
            & "M1.nMovNro = E1.nMovNro Where M1.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & ") " _
            & "And LEFT(M1.cMovNro,8) <= '" & Format$(dFecha, "yyyymmdd") & "' AND Substring(M1.cMovNro,18,2) = '" & sAgencia & "' AND M1.cOpeCod in ('901040') ) "
        sCondicion = ""
    Else
        sUltimo = ""
        sCondicion = " And M.cMovNro LIKE '" & Format$(dFecha, "yyyymmdd") & "_________" & sAgencia & "%' "
    End If
    
    ssql = "Select top 1 E.nMovNro, E.nMontodol nSaldo from " & lsTablaDiaria & " M INNER JOIN MovPreCuadre E " _
        & "ON M.nMovNro = E.nMovNro Where SUBSTRING(M.cMovNro,18,2) = '" & sAgencia & "'" _
        & "AND M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & ") " & sCondicion & sUltimo
    
    If sUsuario <> "" Then
        ssql = ssql & " And right(M.cMovNro,4) = '" & sUsuario & "'"
    End If
    
    ssql = ssql & " AND M.cOpeCod in ('901040') Order by nMovNro desc "
    
End If
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Function

Set rsSaldo = oConect.CargaRecordSet(ssql)
GetSaldoEfectivoPreCuadre = rsSaldo("nSaldo")
rsSaldo.Close
Set rsSaldo = Nothing
oConect.CierraConexion
Set oConect = Nothing
End Function


