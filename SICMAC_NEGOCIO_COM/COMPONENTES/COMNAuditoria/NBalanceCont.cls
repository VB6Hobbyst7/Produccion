VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NBalanceCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Dim dbConec As COMConecta.DCOMConecta
Dim psSql As String
Dim prs As ADODB.Recordset

Dim sCentralCom As String
Dim lsEmprLogo  As String

Public Enum ValidacionBalance
   gValidaCuadreAsiento = 1
   gValidaConvesionME = 2
   gValidaCuentasNoExistentes = 3
   gValidaCuentasNoExistentes2 = 4
   gValidaCuentasAnaliticas = 5
   gValidaCuentasDeOrden = 6
   gValidaSaldosContables = 7
   gValidaCuentasSinPadre = 8
   gValidaCuentasConsolidadas = 9
   gValidaAsientosDigitos = 10
End Enum

Private Sub Class_Initialize()
    Dim oImp As COMDConstSistema.DCOMImpresoras
    Set oImp = New COMDConstSistema.DCOMImpresoras
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    Set oImp = Nothing

    lsEmprLogo = "CMAC MAYNAS S.A."
    gbBitCentral = True
    gcConvTipo = "2"
    gsCodCMAC = "109"

    Set dbConec = New COMConecta.DCOMConecta

    If Not dbConec.AbreConexion() Then

       MsgBox Err.Description, vbCritical
    End If

End Sub

Private Sub Class_Terminate()
dbConec.CierraConexion
Set dbConec = Nothing
End Sub

Public Function getImporteBalanceMes(psCtaCod As String, pnTipoBala As Integer, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer) As Currency
On Error GoTo getImporteBalanceMesErr
getImporteBalanceMes = 0
psSql = "SELECT SubString(cCtaContCod,1,1) as cCtaContCod, " _
     & "       nSaldoFinImporte as nImporte " _
     & "From BalanceEstad WHERE cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
     & "     ' and cBalanceMes = '" & Format(pnMes, "00") & "' and " _
     & "cBalanceAnio = '" & Format(pnAnio, "0000") & "' and cCtaContCod = '" & psCtaCod & "' "
Set prs = dbConec.CargaRecordSet(psSql)
If Not prs.EOF Then
   getImporteBalanceMes = prs!nImporte
End If
Exit Function
getImporteBalanceMesErr:
   MsgBox Err.Description, vbCritical
End Function

Public Function GetCtasSaldoME(pdFecha As Date) As ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim sql As String

    sql = " Select B.cCtaContCod, dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, gsFormatoFecha) & "',B.cCtaContCod + '%',2) SaldoME FROM CtaContMEBalance A " _
        & " Inner Join CtaCont B On B.cCtaContCod Like A.cCtaContCod + '%'"

    oCon.AbreConexion
    Set GetCtasSaldoME = oCon.CargaRecordSet(sql)
End Function

Public Function GetTipCambioBalance(psMovFecha As String) As Currency
On Error GoTo getTipCambioBalanceErr
   psSql = "SELECT nMovOtroImporte FROM MovOtrosItem MO JOIN Mov M ON M.nMovNro = MO.nMovNro WHERE M.nMovEstado = '" & gMovEstContabMovContable & "' and not M.nMovFlag IN ('" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagExtornado & "','" & gMovFlagModificado & "') and " _
           & "M.cMovNro LIKE '" & psMovFecha & "%' and cmovotrovariable = 'TC2'"
   Set prs = dbConec.CargaRecordSet(psSql)
   If Not prs.EOF Then
      GetTipCambioBalance = prs!nMovOtroImporte
   Else
      If gbBitCentral Then
          psSql = "SELECT nValFijo FROM " & sCentralCom & "TipoCambio WHERE LEFT(convert(varchar(8),dFecCamb,112),6) LIKE '" & Left(psMovFecha, 6) & "%' "
      Else
          psSql = "SELECT nValFijo FROM DBcomunes..TipCambio WHERE LEFT(convert(varchar(8),dFecCamb,112),6) LIKE '" & Left(psMovFecha, 6) & "%' "
      End If
      Set prs = dbConec.CargaRecordSet(psSql)
      
      If Not prs.EOF Then
         GetTipCambioBalance = prs!nValFijo
      Else
         GetTipCambioBalance = 0
      End If
   End If
   prs.Close: Set prs = Nothing
Exit Function
getTipCambioBalanceErr:
   MsgBox Err.Description, vbCritical
End Function

Public Function ValidaBalance(lSoloUtilidad As Boolean, pdFechaIni As Date, pdFechaFin As Date, pnTipoBala As Integer, pnMoneda As Integer) As String
Dim nUtilidad As Currency
Dim nUtilidadMes As Currency
Dim nRei As Currency
Dim nDeduccion As Currency
Dim nDeduccion1 As Currency
Dim sValida    As String
Dim n5 As Currency, n4 As Currency
Dim n62 As Currency, n63 As Currency, n64 As Currency, n65 As Currency, n66 As Currency
nUtilidad = 0
nUtilidadMes = 0
If Month(pdFechaIni) > 1 Then
   nUtilidad = GetUtilidadAcumulada(Format(pnTipoBala, "#"), pnMoneda, Format(Month(pdFechaIni - 1), "00"), Format(Year(pdFechaIni - 1), "0000"))
End If

n5 = getImporteBalanceMes("5", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n62 = getImporteBalanceMes("62", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n63 = getImporteBalanceMes("63", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n64 = getImporteBalanceMes("64", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n65 = getImporteBalanceMes("65", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n66 = getImporteBalanceMes("66", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n4 = getImporteBalanceMes("4", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nUtilidadMes = n5 + n62 + n64 - (n4 + n63 + n65)

'69
nRei = getImporteBalanceMes("69", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
If gsCodCMAC = "102" Then
   nRei = nRei * -1
End If
nDeduccion = getImporteBalanceMes("67", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nDeduccion1 = getImporteBalanceMes("68", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nDeduccion = nDeduccion * -1
nDeduccion1 = nDeduccion1 * -1

If Not lSoloUtilidad Then
   Linea sValida, Cabecera("C A L C U L O   D E   L A   U T I L I D A D", 0, "", 80, , , lsEmprLogo)
   If pnMoneda = 0 Then
      Linea sValida, Centra("C O N S O L I D A D O", 80)
   End If
   Linea sValida, Centra("AL " & pdFechaFin, 80), 3
End If
Linea sValida, String(2, oImpresora.gPrnSaltoLinea) & Space(5) & Justifica("UTILIDAD ACUMULADA AL " & CDate(pdFechaIni) - 1, 40) & ": " & PrnVal(nUtilidad, 16, 2)
Linea sValida, Space(5) & Justifica("UTILIDAD DEL MES DE " & Mid(pdFechaFin, 4, 10), 40) & ": " & PrnVal(nUtilidadMes - nUtilidad, 16, 2)
Linea sValida, Space(45) & "--------------------"
Linea sValida, Space(5) & Justifica("UTILIDAD ACUMULADA AL " & pdFechaFin, 40) & ": " & PrnVal(nUtilidadMes, 16, 2)
Linea sValida, Space(5) & Justifica("R.E.I. " & Right(pdFechaFin, 4), 40) & ": " & PrnVal(nRei, 16, 2)
Linea sValida, Space(45) & "===================="
Linea sValida, Space(5) & Justifica("UTILIDAD ANTES DE IMPUESTO", 40) & ": " & PrnVal(nUtilidadMes + nRei, 16, 2)
Linea sValida, Space(5) & Justifica("PARTICIPACION LABORAL", 40) & ": " & PrnVal(nDeduccion, 16, 2)
Linea sValida, Space(5) & Justifica("IMPUESTO A LA RENTA", 40) & ": " & PrnVal(nDeduccion1, 16, 2)
Linea sValida, Space(45) & "--------------------"
Linea sValida, Space(5) & Justifica("UTILIDAD(PERDIDA) NETA AL " & pdFechaFin, 40) & ": " & PrnVal(nRei + nUtilidadMes + nDeduccion + nDeduccion1, 16, 2)
Linea sValida, Space(45) & "====================", 2

If Not lSoloUtilidad Then

   Dim nActivo As Currency
   Dim nPasivo As Currency
   Dim nPatri  As Currency

   'Eliminamos si Existe la Utilidad Acumulada del Mes
   Dim dBalance As New DbalanceCont
   dBalance.EliminaUtilidadAcumulada pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni), True
   dBalance.InsertaUtilidadAcumulada pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni), nUtilidadMes, True
   dBalance.EjecutaBatch

   nActivo = getImporteBalanceMes("1", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
   nPasivo = getImporteBalanceMes("2", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
   nPatri = getImporteBalanceMes("3", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))

   Linea sValida, BON & Centra(" CONSISTENCIA DE CUADRE DEL BALANCE ", 80) & BOFF
   Select Case pnMoneda
      Case 0: Linea sValida, Centra(" ( CONSOLIDADO ) ", 80), 2
      Case 1: Linea sValida, Centra(" ( MONEDA NACIONAL ) ", 80), 2
      Case 2: Linea sValida, Centra(" ( MONEDA EXTRANJERA ) ", 80), 2
   End Select
   Linea sValida, Justifica("ACTIVO", 40) & PrnVal(nActivo, 16, 2)
   Linea sValida, Justifica("PASIVO", 40) & Space(18) & PrnVal(nPasivo, 16, 2)
   Linea sValida, Justifica("PATRIMONIO", 40) & Space(18) & PrnVal(nPatri, 16, 2)
   Linea sValida, Justifica("UTILIDAD (PERDIDA) NETA", 40) & Space(18) & PrnVal(nRei + nUtilidadMes + nDeduccion + nDeduccion1, 16, 2)
   Linea sValida, Space(39) & String(36, "-")
   Linea sValida, Space(40) & PrnVal(nActivo, 16, 2) & "  " & PrnVal(nPasivo + nPatri + nRei + nUtilidadMes + nDeduccion + nDeduccion1, 16, 2)
   Linea sValida, Space(39) & String(36, "="), 2

   Linea sValida, Space(39) & String(18, "-")
   Linea sValida, Justifica("DIFERENCIA", 40) & PrnVal(nActivo - (nPasivo + nPatri + nRei + nUtilidadMes + nDeduccion + nDeduccion1), 16, 2)
   Linea sValida, Space(39) & String(18, "-")
End If
ValidaBalance = sValida
End Function

Public Function GetUtilidadAcumulada(psTipoBala As String, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer, Optional pbGetUltimo As Boolean = False, Optional pbMsg As Boolean = True) As Currency
On Error GoTo GetUtilidadAcumuladaErr
GetUtilidadAcumulada = 0
If pbGetUltimo Then
   psSql = "SELECT nUtilidadImporte FROM UtilidadAcumula U " _
     & "WHERE  cBalanceCate = '" & psTipoBala & "' " & IIf(pnMoneda = -1, "", " and cBalanceTipo = '" & pnMoneda & "'") _
     & " and cUtilidadAnio+cUtilidadMes = (SELECT MAX(cUtilidadAnio+cUtilidadMes) FROM UtilidadAcumula WHERE cBalanceCate = U.cBalanceCate and cBalanceTipo = U.cBalanceTipo and cUtilidadAnio+cUtilidadMes <= '" & Format(pnAnio, "0000") & Format(pnMes, "00") & "' )"
Else
   psSql = "SELECT nUtilidadImporte FROM UtilidadAcumula " _
     & "WHERE  cBalanceCate = '" & psTipoBala & "' and cBalanceTipo = '" & pnMoneda _
     & "' and cUtilidadMes = '" & Format(pnMes, "00") & "' and " _
     & "cUtilidadAnio = '" & Format(pnAnio, "0000") & "'"
End If
Set prs = dbConec.CargaRecordSet(psSql)
If prs.EOF Then
   If pbMsg Then
      RSClose prs
      MsgBox "No se encontraron datos de Utilidad Acumulada del Mes anterior"
   End If
   Exit Function
End If
GetUtilidadAcumulada = prs!nUtilidadImporte
prs.Close
Set prs = Nothing
Exit Function
GetUtilidadAcumuladaErr:
   MsgBox Err.Description, vbCritical
End Function

Public Function BalanceGeneradoHisto(pnTipoBala As Integer, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer) As Boolean
On Error GoTo BalanceGeneradoErr
                    'MsgBox "Antes de: BalanceGeneradoHisto = True"
BalanceGeneradoHisto = True
                    'MsgBox "Antes de: psSql"
psSql = "SELECT TOP 1 cCtaContCod as cCtaContCod " _
     & "FROM   BalanceEstad " _
     & "WHERE cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
     & "' and cBalanceMes = '" & Format(pnMes, "00") & "' " _
     & "  and cBalanceAnio = '" & Format(pnAnio, "0000") & "' "
                    'MsgBox "Antes de: Set prs"
Set prs = dbConec.CargaRecordSet(psSql)
                    'MsgBox "Antes de: If prs.EOF Then"
If prs.EOF Then
   BalanceGeneradoHisto = False
End If
                    'MsgBox "Antes de: prs.Close"
prs.Close: Set prs = Nothing

Exit Function

BalanceGeneradoErr:
   MsgBox Err.Description, vbCritical

End Function

Public Function LeeBalanceHisto(pnTipoBala As Integer, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer, Optional psCtaIni As String = "", Optional psCtaFin As String = "", Optional pnDigitos As Integer = 0, Optional pbSoloAnaliticas As Boolean = False) As Recordset
Dim lsCtaFiltro As String
On Error GoTo LeeBalanceHistoErr

If psCtaIni <> "" Then
    lsCtaFiltro = " and b.cCtaContCod >= '" & psCtaIni & "'"
End If
If psCtaFin <> "" Then
    lsCtaFiltro = lsCtaFiltro & " and b.cCtaContCod <= '" & psCtaFin & "'"
End If

psSql = "SELECT b.cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte " _
     & "FROM   BalanceEstad b join " & sCentralCom & "CtaCont c ON c.cCtaContCod = b.cCtaContCod " _
     & IIf(pbSoloAnaliticas, " join CtaCont c1 ON c1.cCtaContCod LIKE c.cCtaContCod + '%' ", "") _
     & "WHERE len(b.cCtaContCod) <= " & pnDigitos & " and cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
     & "' and cBalanceMes = '" & Format(pnMes, "00") & "' and " _
     & "cBalanceAnio = '" & Format(pnAnio, "0000") & "' " & lsCtaFiltro _
     & IIf(pbSoloAnaliticas, " GROUP BY b.cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte HAVING count(*) = 1", "") _
     & " ORDER BY b.cCtaContCod "
     
   Set LeeBalanceHisto = dbConec.CargaRecordSet(psSql)
Exit Function
LeeBalanceHistoErr:
   MsgBox Err.Description, vbCritical
End Function

Public Function TotalizaBalanceHisto(pnTipoBala As Integer, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer) As Recordset
On Error GoTo TotalizaBalanceHistoErr
psSql = "SELECT SUM(nDebe) as nDebe, SUM(nHaber) as nHaber FROM BalanceEstad " _
     & "WHERE cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda & "' " _
     & "  and cBalanceMes = '" & Format(pnMes, "00") & "' and " _
     & "      cBalanceAnio = '" & Format(pnAnio, "0000") & "' and LEN(cCtaContCod) = 1 "
Set TotalizaBalanceHisto = dbConec.CargaRecordSet(psSql)
Exit Function
TotalizaBalanceHistoErr:
   MsgBox Err.Description, vbCritical
End Function

Public Function GetBalanceSituacion(pnTipoBala As Integer, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer, psCaracter As String) As Recordset
On Error GoTo GetBalanceSituacionErr
psSql = "SELECT Substring(B.cCtaContCod,1,2) as cCtaContCod,nSaldoIniImporte as nInicial, nDebe as nDebe, nHaber as nHaber " _
     & "FROM   BalanceEstad B JOIN " & sCentralCom & "CtaContClase cls ON B.cCtaContcod LIKE RTRIM(cls.cCtaContCod) + '%' " _
     & "WHERE  cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda & "' " _
     & "   and cBalanceMes = '" & Format(pnMes, "00") & "' and " _
     & "       cBalanceAnio = '" & Format(pnAnio, "0000") & "' and cls.cCtaCaracter = '" & psCaracter & "' " _
     & "and Len(B.cCtaContCod) = 2 and (nSaldoIniImporte <> 0 or nDebe <> 0 or nHaber<> 0) " _
     & "ORDER  BY B.cCtaContCod "
Set GetBalanceSituacion = dbConec.CargaRecordSet(psSql)
Exit Function
GetBalanceSituacionErr:
  MsgBox Err.Description, vbCritical
End Function

