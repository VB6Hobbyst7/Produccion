Attribute VB_Name = "gFunGeneralTesoreria"
Global gsPersNombre As String

Public Sub EnviaPrevioTes(psImpre As String, psTitulo As String, ByVal pnLinPage As Integer, Optional plCondensado As Boolean = False)
Dim clsPrevioFinan As clsprevio
Set clsPrevioFinan = New clsprevio
clsPrevioFinan.Show psImpre, psTitulo, plCondensado, pnLinPage, gImpresora
Set clsPrevioFinan = Nothing
End Sub

Public Function AgregaUtilidad(lSoloUtilidad As Boolean, pdFechaIni As Date, pdFechaFin As Date, pnTipoBala As Integer, pnMoneda As Integer) As String
Dim nUtilidad As Currency
Dim nUtilidadMes As Currency
Dim nRei As Currency
Dim nDeduccion As Currency
Dim nDeduccion1 As Currency
Dim sValida As String
Dim n5 As Currency, n4 As Currency
Dim n62 As Currency, n63 As Currency, n64 As Currency, n65 As Currency, n66 As Currency
Dim glsArchivo As String
Dim ss As Integer

On Error GoTo ss
nUtilidad = 0
nUtilidadMes = 0
ss = 1
Dim oBal As COMNAuditoria.NBalanceCont
Set oBal = New COMNAuditoria.NBalanceCont

If Month(pdFechaIni) > 1 Then
   nUtilidad = oBal.GetUtilidadAcumulada(Format(pnTipoBala, "#"), pnMoneda, Format(Month(pdFechaIni - 1), "00"), Format(Year(pdFechaIni - 1), "0000"))
End If
ss = 2
n5 = oBal.getImporteBalanceMes("5", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n62 = oBal.getImporteBalanceMes("62", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n63 = oBal.getImporteBalanceMes("63", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n64 = oBal.getImporteBalanceMes("64", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n65 = oBal.getImporteBalanceMes("65", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n66 = oBal.getImporteBalanceMes("66", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n4 = oBal.getImporteBalanceMes("4", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nUtilidadMes = n5 + n62 + n64 - (n4 + n63 + n65)
ss = 3
'69
nRei = oBal.getImporteBalanceMes("69", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
If gsCodCMAC = "102" Then
   nRei = nRei * -1
End If
ss = 4
nDeduccion = oBal.getImporteBalanceMes("67", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nDeduccion1 = oBal.getImporteBalanceMes("68", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nDeduccion = nDeduccion * -1
nDeduccion1 = nDeduccion1 * -1
ss = 5
If Not lSoloUtilidad Then

   Dim nActivo As Currency
   Dim nPasivo As Currency
   Dim nPatri  As Currency

   Dim dBalance1 As COMNAuditoria.DbalanceCont
   Set dBalance1 = New COMNAuditoria.DbalanceCont
ss = 6
   dBalance1.EliminaUtilidadAcumulada pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni), True
   dBalance1.InsertaUtilidadAcumulada pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni), nUtilidadMes, True
   dBalance1.EjecutaBatch
ss = 7
   nActivo = oBal.getImporteBalanceMes("1", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
   nPasivo = oBal.getImporteBalanceMes("2", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
   nPatri = oBal.getImporteBalanceMes("3", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
ss = 8
   glsArchivo = glsArchivo & Space(20) & "CONSTANCIA DE CUADRE DE BALANCE" & oImpresora.gPrnSaltoLinea
   If pnMoneda = 0 Then
      glsArchivo = glsArchivo & Space(20) & "        (CONSOLIDADO)          " & oImpresora.gPrnSaltoLinea
    ElseIf pnMoneda = 2 Then
        glsArchivo = glsArchivo & Space(20) & "     (MONEDA EXTRANJERA)       " & oImpresora.gPrnSaltoLinea
    Else
        glsArchivo = glsArchivo & Space(20) & "      (MONEDA NACIONAL)        " & oImpresora.gPrnSaltoLinea
    End If
    glsArchivo = glsArchivo & Space(20) & String(35, "-") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea

    glsArchivo = glsArchivo & Space(5) & "ACTIVO" & Space(35) & PrnVal(nActivo, 16, 2) & oImpresora.gPrnSaltoLinea

    glsArchivo = glsArchivo & Space(5) & "PASIVO" & Space(55) & PrnVal(nPasivo, 16, 2) & oImpresora.gPrnSaltoLinea

    glsArchivo = glsArchivo & Space(5) & "PATRIMONIO" & Space(50) & PrnVal(nPatri, 16, 2) & oImpresora.gPrnSaltoLinea
ss = 9
    glsArchivo = glsArchivo & Space(5) & "UTILIDAD (PERDIDA) NETA" & Space(37) & PrnVal(nRei + nUtilidadMes + nDeduccion + nDeduccion1, 16, 2) & oImpresora.gPrnSaltoLinea

    glsArchivo = glsArchivo & Space(47) & String(38, "-") & oImpresora.gPrnSaltoLinea

    glsArchivo = glsArchivo & Space(45) & PrnVal(nActivo, 16, 2) & Space(5) & PrnVal(nPasivo + nPatri + nRei + nUtilidadMes + nDeduccion + nDeduccion1, 16, 2) & oImpresora.gPrnSaltoLinea

    glsArchivo = glsArchivo & Space(47) & String(38, "-") & oImpresora.gPrnSaltoLinea

    glsArchivo = glsArchivo & Space(5) & "DIFERENCIA" & Space(30) & PrnVal(nActivo - (nPasivo + nPatri + nRei + nUtilidadMes + nDeduccion + nDeduccion1), 16, 2)
    Set dBalance1 = Nothing
    Set oBal = Nothing
ss = 10
   AgregaUtilidad = glsArchivo
End If
Exit Function
ss:
MsgBox ss & " " & Err.Description & " " & Err.Number & " " & Err.Source
End Function
