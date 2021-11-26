VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Defectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vsConexion As String
Dim vsServerCom As String
Dim vsServerPers As String
Dim vsFormatoFecha As String
Public Sub Inicio(ByVal psFormatoFecha As String)
vsFormatoFecha = psFormatoFecha
End Sub
Public Function EmiteBilletajes(ByVal psMonedas As Moneda, Optional ByVal psDenominacion As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oconect As DConecta
Dim lsFiltroDenom As String
Set oconect = New DConecta
If psDenominacion <> "" Then
    lsFiltroDenom = " and Substring(cEfectivoCod,2,1)='" & psDenominacion & "' "
End If
If oconect.AbreConexion = False Then Exit Function

sql = " Select     Case " _
    & "         WHEN Substring(cEfectivoCod,2,1)='B'  Then 'BILLETAJE' + Convert(char(10),nEfectivoValor) " _
    & "         WHEN Substring(cEfectivoCod,2,1)='M'  Then 'MONEDA   ' + Convert(char(10),nEfectivoValor) " _
    & "         END as Descripcion, 0 as Cantidad, 0  as Monto, cEfectivoCod, nEfectivoValor " _
    & " From Efectivo " _
    & " WHERE SUBSTRING(cEfectivoCod,1,1) ='" & psMonedas & "' " & lsFiltroDenom _
    & " ORDER BY  SUBSTRING(cEfectivoCod,2,1), nEfectivoValor DESC "
    
Set rs = oconect.CargaRecordSet(sql)
Set EmiteBilletajes = rs

oconect.CierraConexion: Set oconect = Nothing
End Function


Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

Dim oIni As ClasIni
Set oIni = New ClasIni
vsConexion = oIni.CadenaConexion
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas

End Sub
Public Function GetBilletajes(ByVal psFiltro As String) As ADODB.Recordset
Dim sql As String
Dim oconect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltro As String

Set rs = New ADODB.Recordset
Set oconect = New DConecta
If oconect.AbreConexion = False Then Exit Function
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

Set rs = oconect.CargaRecordSet(sql)
Set GetBilletajes = rs
oconect.CierraConexion: Set oconect = Nothing

End Function
Public Function GetBilletajesMov(ByVal psMovNro As String, Optional ByVal pnMoneda As Moneda, Optional ByVal psDenominacion As String = "")
Dim sql As String
Dim oconect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltroDenom As String

Set rs = New ADODB.Recordset
Set oconect = New DConecta
If oconect.AbreConexion = False Then Exit Function

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
    & "          FROM   MOV M JOIN MOVCTA MC ON MC.nMOVNRO = M.nMOVNRO  " _
    & "                 LEFT JOIN MOVME ME ON ME.nMOVNRO = MC.nMOVNRO AND ME.nMOVITEM=MC.nMOVITEM " _
    & "                 JOIN MOVOBJ MO      ON MO.nMOVNRO =MC.nMOVNRO AND MO.nMOVITEM=MC.nMOVITEM " _
    & "                 JOIN MOVOBJEFECTIVO MOE ON MOE.nMovNro =MO.nMovNro AND MO.nMOVITEM=MOE.nMOVITEM AND MO.nMovObjOrden= MOE.nMovObjOrden " _
    & "                 JOIN EFECTIVO E ON E.cEfectivoCod = MOE.cEfectivoCod " _
    & "          WHERE   M.CMOVNRO= '" & psMovNro & "') AS MovEfect " _
    & "     ON MovEfect.cEfectivoCod = E.cEfectivoCod " _
    & " WHERE E.cEfectivoCod LIKE '" & pnMoneda & "%' " & lsFiltroDenom & " Order BY E.nEfectivoValor Desc "

Set rs = oconect.CargaRecordSet(sql)
Set GetBilletajesMov = rs
oconect.CierraConexion: Set oconect = Nothing
End Function
Public Function GetEfectivoMovImp(ByVal psMovNro As String, Optional ByVal pbDebe As Boolean = True) As ADODB.Recordset
Dim sql As String
Dim oconect As DConecta
Dim rs As ADODB.Recordset
Dim lsFiltro As String

Set rs = New ADODB.Recordset
Set oconect = New DConecta
If oconect.AbreConexion = False Then Exit Function

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

Set rs = oconect.CargaRecordSet(sql)
Set GetEfectivoMovImp = rs
oconect.CierraConexion: Set oconect = Nothing

End Function
