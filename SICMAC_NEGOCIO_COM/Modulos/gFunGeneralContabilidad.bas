Attribute VB_Name = "gFunGeneralContabilidad"
Option Explicit

Public Function GetFechaHoraServer() As String
Dim oConect As DConecta
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Function
GetFechaHoraServer = oConect.GetFechaHoraServer()
oConect.CierraConexion
Set oConect = Nothing
End Function

Public Function ImprimeBalanceComprobacion(pdFechaIni As Date, pdFechaFin As Date, pdFecha As Date, pnTipoBala As Integer, _
                pnMoneda As Integer, pnLinPage As Integer, nTotDebe As Currency, nTotHaber As Currency, Optional psCtaIni As String = "", _
                Optional psCtaFin As String = "", Optional pnDigitos As Integer = 0, Optional pbSoloAnaliticas As Boolean = False, _
                Optional pnCierreAnio As Integer = 0, Optional pbExcel As Boolean = False, Optional xlHoja1 As Excel.Worksheet, Optional psTipo As Integer) As String

Dim prs As ADODB.Recordset
Set prs = New ADODB.Recordset
Dim n As Integer
Dim nLin As Integer, P As Integer
Dim lnFila As Integer
Dim nTotD As Currency, nTotH As Currency, nSaldo As Currency
Dim lOk As Boolean
Dim sCond As String
Dim sTexto As String
Dim lsImpre As String
Dim lnCont As Long
Dim sObj As String
Dim nImporte As Currency
Dim rsAge As New ADODB.Recordset
Dim sImporteME As String
Dim lnTotCta As Currency
Dim nTipCambio As Currency

Dim sTotales() As Currency

Dim nTotalCtaD As Currency
Dim nTotalCtaH As Currency
Dim nTotalCtaIni As Currency
Dim nTotalCtaFin As Currency

Dim nTotalAjuteAnt As Currency
Dim nTotalAjuteH As Currency
Dim nTotalAjuteD As Currency
Dim nTotalAjuteAct As Currency
Dim sCtaClase As String
Dim I, J, K As Integer
Dim nSaldoIniImporte1, nSaldoFinImporte1 As Double
Dim sCtaAux As String

Dim rsCtasME As ADODB.Recordset
Set rsCtasME = New ADODB.Recordset

Dim oBalanceGet As COMNAuditoria.NBalanceCont
Set oBalanceGet = New COMNAuditoria.NBalanceCont

If pnMoneda = 2 Then
    Set rsCtasME = oBalanceGet.GetCtasSaldoME(pdFechaFin)
End If

lsImpre = ""
lnCont = 0
lnTotCta = 0
sCtaClase = ""
nLin = 66
If pnMoneda = 2 Then
   nTipCambio = oBalanceGet.GetTipCambioBalance(Format(pdFechaFin, "yyyymmdd"))
End If
Set prs = oBalanceGet.LeeBalanceHisto(pnTipoBala, pnMoneda, Month(pdFechaIni) + pnCierreAnio, Year(pdFechaIni), psCtaIni, psCtaFin, pnDigitos, pbSoloAnaliticas)
If pbExcel Then
   ImprimeBalanceComprobacionCabExcel pdFechaIni, pdFechaFin, nTipCambio, P, pnTipoBala, pnMoneda, pdFecha, nLin, pnLinPage, xlHoja1
End If
lnFila = 7
ReDim sTotales(3, 8, 4)

sCtaClase = Val(Mid(prs!cCtaContCod, 1, 1))

Do While Not prs.EOF
   If Not pbExcel Then
      DoEvents
      sTexto = sTexto & ImprimeBalanceComprobacionCabecera(pdFechaIni, pdFechaFin, nTipCambio, P, pnTipoBala, pnMoneda, pdFecha, nLin, pnLinPage)
      If Len(prs!cCtaContCod) <= 4 Then
         LineaAuditoria sTexto, ""
         nLin = nLin + 1
         lnCont = lnCont + 1
           If lnCont Mod 600 = 0 Then
              lsImpre = lsImpre & sTexto
              sTexto = ""
           End If
      End If
   End If

   If sCtaClase <> Val(Mid$(prs!cCtaContCod, 1, 1)) Then
        If pbExcel And psTipo <> 0 Then
            xlHoja1.Cells(lnFila, 1) = "TOTAL MONEDA NACIONAL"
            xlHoja1.Cells(lnFila, 3) = sTotales(1, Val(sCtaClase), 2) ''prs!nSaldoIniImporte
            xlHoja1.Cells(lnFila, 4) = sTotales(1, Val(sCtaClase), 0) ''prs!nDebe
            xlHoja1.Cells(lnFila, 5) = sTotales(1, Val(sCtaClase), 1)  ''prs!nHaber
            xlHoja1.Cells(lnFila, 6) = sTotales(1, Val(sCtaClase), 3) ''prs!nSaldoFinImporte
            xlHoja1.Range("A" & lnFila & ":F" & lnFila).Font.Bold = True
            lnFila = lnFila + 1
            xlHoja1.Cells(lnFila, 1) = "TOTAL MONEDA EXTRANJERA"
            xlHoja1.Cells(lnFila, 3) = sTotales(2, Val(sCtaClase), 2)
            xlHoja1.Cells(lnFila, 4) = sTotales(2, Val(sCtaClase), 0)
            xlHoja1.Cells(lnFila, 5) = sTotales(2, Val(sCtaClase), 1)
            xlHoja1.Cells(lnFila, 6) = sTotales(2, Val(sCtaClase), 3)
            xlHoja1.Range("A" & lnFila & ":F" & lnFila).Font.Bold = True
            lnFila = lnFila + 1
            xlHoja1.Cells(lnFila, 1) = "TOTAL AJUSTE POR INFLACION"
            xlHoja1.Cells(lnFila, 3) = sTotales(0, Val(sCtaClase), 2)
            xlHoja1.Cells(lnFila, 4) = sTotales(0, Val(sCtaClase), 0)
            xlHoja1.Cells(lnFila, 5) = sTotales(0, Val(sCtaClase), 1)
            xlHoja1.Cells(lnFila, 6) = sTotales(0, Val(sCtaClase), 3)
            xlHoja1.Range("A" & lnFila & ":F" & lnFila).Font.Bold = True
            lnFila = lnFila + 1

            If sCtaClase = 3 Then
                nTotalCtaIni = (sTotales(1, 1, 2) + sTotales(2, 1, 2) + sTotales(0, 1, 2)) - (sTotales(1, 2, 2) + sTotales(2, 2, 2) + sTotales(0, 2, 2) + sTotales(1, 3, 2) + sTotales(2, 3, 2) + sTotales(0, 3, 2))
                nTotalCtaD = (sTotales(1, 1, 0) + sTotales(2, 1, 0) + sTotales(0, 1, 0)) + (sTotales(1, 2, 0) + sTotales(2, 2, 0) + sTotales(0, 2, 0) + sTotales(1, 3, 0) + sTotales(2, 3, 0) + sTotales(0, 3, 0))
                nTotalCtaH = (sTotales(1, 1, 1) + sTotales(2, 1, 1) + sTotales(0, 1, 1)) + (sTotales(1, 2, 1) + sTotales(2, 2, 1) + sTotales(0, 2, 1) + sTotales(1, 3, 1) + sTotales(2, 3, 1) + sTotales(0, 3, 1))

                xlHoja1.Cells(lnFila, 1) = "TOTAL CUENTAS DE BALANCE"
                xlHoja1.Cells(lnFila, 3) = nTotalCtaIni
                xlHoja1.Cells(lnFila, 4) = nTotalCtaD
                xlHoja1.Cells(lnFila, 5) = nTotalCtaH
                xlHoja1.Cells(lnFila, 6) = nTotalCtaIni + nTotalCtaD - nTotalCtaH
                xlHoja1.Range("A" & lnFila & ":F" & lnFila).Font.Bold = True
            ElseIf sCtaClase = 6 Then

                nTotalCtaD = 0: nTotalCtaH = 0: nTotalCtaIni = 0

                nTotalCtaIni = (sTotales(1, 4, 2) + sTotales(2, 4, 2) + sTotales(0, 4, 2)) - (sTotales(1, 5, 2) + sTotales(2, 5, 2) + sTotales(0, 5, 2) + sTotales(1, 6, 2) + sTotales(2, 6, 2) + sTotales(0, 6, 2))
                nTotalCtaD = (sTotales(1, 4, 0) + sTotales(2, 4, 0) + sTotales(0, 4, 0)) + (sTotales(1, 5, 0) + sTotales(2, 5, 0) + sTotales(0, 5, 0) + sTotales(1, 6, 0) + sTotales(2, 6, 0) + sTotales(0, 6, 0))
                nTotalCtaH = (sTotales(1, 4, 1) + sTotales(2, 4, 1) + sTotales(0, 4, 1)) + (sTotales(1, 5, 1) + sTotales(2, 5, 1) + sTotales(0, 5, 1) + sTotales(1, 6, 1) + sTotales(2, 6, 1) + sTotales(0, 6, 1))

                xlHoja1.Cells(lnFila, 1) = "TOTAL CUENTAS DE BALANCE"
                xlHoja1.Cells(lnFila, 3) = nTotalCtaIni
                xlHoja1.Cells(lnFila, 4) = nTotalCtaD
                xlHoja1.Cells(lnFila, 5) = nTotalCtaH
                xlHoja1.Cells(lnFila, 6) = nTotalCtaIni + nTotalCtaD - nTotalCtaH
                xlHoja1.Range("A" & lnFila & ":F" & lnFila).Font.Bold = True
            End If

            If sCtaClase = 3 Or sCtaClase = 6 Then
                xlHoja1.Range("A" & lnFila - 3 & ":F" & lnFila - 1).Borders(xlEdgeTop).LineStyle = xlContinuous
                xlHoja1.Range("A" & lnFila - 3 & ":F" & lnFila - 1).Borders(xlEdgeTop).Weight = xlThin
                xlHoja1.Range("A" & lnFila - 3 & ":F" & lnFila - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
                xlHoja1.Range("A" & lnFila - 3 & ":F" & lnFila - 1).Borders(xlEdgeBottom).Weight = xlThin
                xlHoja1.Range("A" & lnFila & ":F" & lnFila).Borders(xlEdgeBottom).LineStyle = xlContinuous
                xlHoja1.Range("A" & lnFila & ":F" & lnFila).Borders(xlEdgeBottom).Weight = xlMedium
                lnFila = lnFila + 2
            Else
                xlHoja1.Range("A" & lnFila - 3 & ":F" & lnFila - 1).Borders(xlEdgeTop).LineStyle = xlContinuous
                xlHoja1.Range("A" & lnFila - 3 & ":F" & lnFila - 1).Borders(xlEdgeTop).Weight = xlThin
                xlHoja1.Range("A" & lnFila - 3 & ":F" & lnFila - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
                xlHoja1.Range("A" & lnFila - 3 & ":F" & lnFila - 1).Borders(xlEdgeBottom).Weight = xlMedium
                lnFila = lnFila + 1
            End If

        ElseIf pbExcel = False And psTipo <> 0 Then

            LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & " --------------------------------------------------------------------------------------------------------------------------------------------"
            LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & oImpresora.gPrnBoldON & "   " & "TOTAL MONEDA NACIONAL     " & Space(44) & PrnVal(sTotales(1, Val(sCtaClase), 2), 16, 2) & " " _
                  & PrnVal(sTotales(1, Val(sCtaClase), 0), 16, 2) & " " & PrnVal(sTotales(1, Val(sCtaClase), 1), 16, 2) & " " _
                  & PrnVal(sTotales(1, Val(sCtaClase), 3), 16, 2) & oImpresora.gPrnBoldOFF & oImpresora.gPrnCondensadaOFF
            nLin = nLin + 1
            LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & oImpresora.gPrnBoldON & "   " & "TOTAL MONEDA EXTRANJERA   " & Space(44) & PrnVal(sTotales(2, Val(sCtaClase), 2), 16, 2) & " " _
                  & PrnVal(sTotales(2, Val(sCtaClase), 0), 16, 2) & " " & PrnVal(sTotales(2, Val(sCtaClase), 1), 16, 2) & " " _
                  & PrnVal(sTotales(2, Val(sCtaClase), 3), 16, 2) & oImpresora.gPrnBoldOFF & oImpresora.gPrnCondensadaOFF
            nLin = nLin + 1
            LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & oImpresora.gPrnBoldON & "   " & "TOTAL AJUSTE POR INFLACION" & Space(44) & PrnVal(sTotales(0, Val(sCtaClase), 2), 16, 2) & " " _
                  & PrnVal(sTotales(0, Val(sCtaClase), 0), 16, 2) & " " & PrnVal(sTotales(0, Val(sCtaClase), 1), 16, 2) & " " _
                  & PrnVal(sTotales(0, Val(sCtaClase), 3), 16, 2) & oImpresora.gPrnBoldOFF & oImpresora.gPrnCondensadaOFF
            nLin = nLin + 1
            If sCtaClase = 3 Then
                nTotalCtaIni = (sTotales(1, 1, 2) + sTotales(2, 1, 2) + sTotales(0, 1, 2)) - (sTotales(1, 2, 2) + sTotales(2, 2, 2) + sTotales(0, 2, 2) + sTotales(1, 3, 2) + sTotales(2, 3, 2) + sTotales(0, 3, 2))
                nTotalCtaD = (sTotales(1, 1, 0) + sTotales(2, 1, 0) + sTotales(0, 1, 0)) + (sTotales(1, 2, 0) + sTotales(2, 2, 0) + sTotales(0, 2, 0) + sTotales(1, 3, 0) + sTotales(2, 3, 0) + sTotales(0, 3, 0))
                nTotalCtaH = (sTotales(1, 1, 1) + sTotales(2, 1, 1) + sTotales(0, 1, 1)) + (sTotales(1, 2, 1) + sTotales(2, 2, 1) + sTotales(0, 2, 1) + sTotales(1, 3, 1) + sTotales(2, 3, 1) + sTotales(0, 3, 1))

                LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & oImpresora.gPrnBoldON & "   " & "TOTAL CUENTAS DE BALANCE  " & Space(44) & PrnVal(nTotalCtaIni, 16, 2) & " " _
                  & PrnVal(nTotalCtaD, 16, 2) & " " & PrnVal(nTotalCtaH, 16, 2) & " " _
                  & PrnVal(nTotalCtaIni + nTotalCtaD - nTotalCtaH, 16, 2) & oImpresora.gPrnBoldOFF & oImpresora.gPrnCondensadaOFF
                nLin = nLin + 1

            ElseIf sCtaClase = 6 Then

                nTotalCtaD = 0: nTotalCtaH = 0: nTotalCtaIni = 0
                nTotalCtaIni = (sTotales(1, 4, 2) + sTotales(2, 4, 2) + sTotales(0, 4, 2)) - (sTotales(1, 5, 2) + sTotales(2, 5, 2) + sTotales(0, 5, 2) + sTotales(1, 6, 2) + sTotales(2, 6, 2) + sTotales(0, 6, 2))
                nTotalCtaD = (sTotales(1, 4, 0) + sTotales(2, 4, 0) + sTotales(0, 4, 0)) + (sTotales(1, 5, 0) + sTotales(2, 5, 0) + sTotales(0, 5, 0) + sTotales(1, 6, 0) + sTotales(2, 6, 0) + sTotales(0, 6, 0))
                nTotalCtaH = (sTotales(1, 4, 1) + sTotales(2, 4, 1) + sTotales(0, 4, 1)) + (sTotales(1, 5, 1) + sTotales(2, 5, 1) + sTotales(0, 5, 1) + sTotales(1, 6, 1) + sTotales(2, 6, 1) + sTotales(0, 6, 1))

                LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & oImpresora.gPrnBoldON & "   " & "TOTAL CUENTAS DE BALANCE  " & Space(44) & PrnVal(nTotalCtaIni, 16, 2) & " " _
                  & PrnVal(nTotalCtaD, 16, 2) & " " & PrnVal(nTotalCtaH, 16, 2) & " " _
                  & PrnVal(nTotalCtaIni + nTotalCtaD - nTotalCtaH, 16, 2) & oImpresora.gPrnBoldOFF & oImpresora.gPrnCondensadaOFF
                nLin = nLin + 1
            End If
            LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & " --------------------------------------------------------------------------------------------------------------------------------------------"
            nLin = pnLinPage + 1
            sTexto = sTexto & ImprimeBalanceComprobacionCabecera(pdFechaIni, pdFechaFin, nTipCambio, P, pnTipoBala, pnMoneda, pdFecha, nLin, pnLinPage)

        End If

        sCtaClase = Val(Mid$(prs!cCtaContCod, 1, 1))
   End If

   If Len(prs!cCtaContCod) = 4 Then
       If Mid$(Trim(prs!cCtaContCod), 3, 1) = 1 And Val(Mid$(prs!cCtaContCod, 1, 1)) < 6 Then
            sTotales(1, Val(sCtaClase), 0) = Val(sTotales(1, Val(sCtaClase), 0)) + prs!nDebe
            sTotales(1, Val(sCtaClase), 1) = Val(sTotales(1, Val(sCtaClase), 1)) + prs!nHaber
            sTotales(1, Val(sCtaClase), 2) = Val(sTotales(1, Val(sCtaClase), 2)) + prs!nSaldoIniImporte
            sTotales(1, Val(sCtaClase), 3) = Val(sTotales(1, Val(sCtaClase), 3)) + prs!nSaldoFinImporte
       ElseIf Mid$(Trim(prs!cCtaContCod), 3, 1) = 1 And Val(Mid$(prs!cCtaContCod, 1, 1)) >= 6 Then
            sCtaAux = Mid$(prs!cCtaContCod, 1, 2)
            If (sCtaAux = 62 Or sCtaAux = 64 Or sCtaAux = 72 Or sCtaAux = 82 Or sCtaAux = 84 Or sCtaAux = 86) Then
                nSaldoIniImporte1 = -1 * prs!nSaldoIniImporte
                nSaldoFinImporte1 = -1 * prs!nSaldoFinImporte
            Else
                nSaldoIniImporte1 = prs!nSaldoIniImporte:  nSaldoFinImporte1 = prs!nSaldoFinImporte
            End If
            sTotales(1, Val(sCtaClase), 0) = Val(sTotales(1, Val(sCtaClase), 0)) + prs!nDebe
            sTotales(1, Val(sCtaClase), 1) = Val(sTotales(1, Val(sCtaClase), 1)) + prs!nHaber
            sTotales(1, Val(sCtaClase), 2) = Val(sTotales(1, Val(sCtaClase), 2)) - nSaldoIniImporte1
            sTotales(1, Val(sCtaClase), 3) = Val(sTotales(1, Val(sCtaClase), 3)) - nSaldoFinImporte1
       End If

        If Mid$(Trim(prs!cCtaContCod), 3, 1) = 2 And Val(Mid$(prs!cCtaContCod, 1, 1)) < 6 Then
             sTotales(2, Val(sCtaClase), 0) = Val(sTotales(2, Val(sCtaClase), 0)) + prs!nDebe
             sTotales(2, Val(sCtaClase), 1) = Val(sTotales(2, Val(sCtaClase), 1)) + prs!nHaber
             sTotales(2, Val(sCtaClase), 2) = Val(sTotales(2, Val(sCtaClase), 2)) + prs!nSaldoIniImporte
             sTotales(2, Val(sCtaClase), 3) = Val(sTotales(2, Val(sCtaClase), 3)) + prs!nSaldoFinImporte
        ElseIf Mid$(Trim(prs!cCtaContCod), 3, 1) = 2 And Val(Mid$(prs!cCtaContCod, 1, 1)) >= 6 Then
            sCtaAux = Mid$(prs!cCtaContCod, 1, 2)
            If sCtaAux = 62 Or sCtaAux = 64 Or sCtaAux = 72 Or sCtaAux = 82 Or sCtaAux = 84 Or sCtaAux = 86 Then
                nSaldoIniImporte1 = -1 * prs!nSaldoIniImporte
                nSaldoFinImporte1 = -1 * prs!nSaldoFinImporte
            Else
                nSaldoIniImporte1 = prs!nSaldoIniImporte
                nSaldoFinImporte1 = prs!nSaldoFinImporte
            End If
             sTotales(2, Val(sCtaClase), 0) = Val(sTotales(2, Val(sCtaClase), 0)) + prs!nDebe
             sTotales(2, Val(sCtaClase), 1) = Val(sTotales(2, Val(sCtaClase), 1)) + prs!nHaber
             sTotales(2, Val(sCtaClase), 2) = Val(sTotales(2, Val(sCtaClase), 2)) - nSaldoIniImporte1
             sTotales(2, Val(sCtaClase), 3) = Val(sTotales(2, Val(sCtaClase), 3)) - nSaldoFinImporte1

        End If

        If Mid$(Trim(prs!cCtaContCod), 3, 1) = 6 And Val(Mid$(prs!cCtaContCod, 1, 1)) < 6 Then
                 sTotales(0, Val(sCtaClase), 0) = Val(sTotales(0, Val(sCtaClase), 0)) + prs!nDebe
                 sTotales(0, Val(sCtaClase), 1) = Val(sTotales(0, Val(sCtaClase), 1)) + prs!nHaber
                 sTotales(0, Val(sCtaClase), 2) = Val(sTotales(0, Val(sCtaClase), 2)) + prs!nSaldoIniImporte
                 sTotales(0, Val(sCtaClase), 3) = Val(sTotales(0, Val(sCtaClase), 3)) + prs!nSaldoFinImporte
        ElseIf Mid$(Trim(prs!cCtaContCod), 3, 1) = 6 And Val(Mid$(prs!cCtaContCod, 1, 1)) >= 6 Then
            sCtaAux = Mid$(prs!cCtaContCod, 1, 2)
            If sCtaAux = 62 Or sCtaAux = 64 Or sCtaAux = 72 Or sCtaAux = 82 Or sCtaAux = 84 Or sCtaAux = 86 Then
                nSaldoIniImporte1 = -1 * prs!nSaldoIniImporte
                nSaldoFinImporte1 = -1 * prs!nSaldoFinImporte
            Else
                nSaldoIniImporte1 = prs!nSaldoIniImporte
                nSaldoFinImporte1 = prs!nSaldoFinImporte
            End If
            sTotales(0, Val(sCtaClase), 0) = Val(sTotales(0, Val(sCtaClase), 0)) + prs!nDebe
            sTotales(0, Val(sCtaClase), 1) = Val(sTotales(0, Val(sCtaClase), 1)) + prs!nHaber
            sTotales(0, Val(sCtaClase), 2) = Val(sTotales(0, Val(sCtaClase), 2)) - nSaldoIniImporte1
            sTotales(0, Val(sCtaClase), 3) = Val(sTotales(0, Val(sCtaClase), 3)) - nSaldoFinImporte1
        End If
   End If

   If pnMoneda = 2 Then
      rsCtasME.MoveFirst
      rsCtasME.Find "cCtaContCod Like '" & prs.Fields(0) & "'"
      If rsCtasME.EOF And rsCtasME.EOF Then
            sImporteME = PrnVal(Round(prs!nSaldoFinImporte / nTipCambio, 2), 15, 2)
      Else
            sImporteME = PrnVal(rsCtasME.Fields(1), 15, 2)
      End If
   End If
   Set rsCtasME = Nothing
   If pbExcel Then
      xlHoja1.Cells(lnFila, 1) = prs!cCtaContDesc
      xlHoja1.Cells(lnFila, 2) = prs!cCtaContCod
      xlHoja1.Cells(lnFila, 3) = prs!nSaldoIniImporte
      xlHoja1.Cells(lnFila, 4) = prs!nDebe
      xlHoja1.Cells(lnFila, 5) = prs!nHaber
      xlHoja1.Cells(lnFila, 6) = prs!nSaldoFinImporte
      xlHoja1.Cells(lnFila, 7) = IIf(pnMoneda = 2, sImporteME, "")
      lnFila = lnFila + 1
   Else
      If gsCodCMAC = "102" Then
         LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & "   " & Mid(prs!cCtaContDesc & Space(50), 1, 50) & " " _
                  & Mid(prs!cCtaContCod & Space(20), 1, 20) & " " & PrnVal(prs!nSaldoIniImporte, 16, 2) & " " _
                  & PrnVal(prs!nDebe, 16, 2) & " " _
                  & PrnVal(prs!nHaber, 16, 2) & " " _
                  & PrnVal(prs!nSaldoFinImporte, 16, 2) & " " _
                  & IIf(pnMoneda = 2, sImporteME, "") _
                  & oImpresora.gPrnCondensadaOFF
      Else
         LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & " " & Mid(prs!cCtaContCod & Space(20), 1, 20) & " " _
                  & Mid(prs!cCtaContDesc & Space(50), 1, 50) & " " _
                  & PrnVal(prs!nSaldoIniImporte, 16, 2) & " " _
                  & PrnVal(prs!nDebe, 16, 2) & " " _
                  & PrnVal(prs!nHaber, 16, 2) & " " _
                  & PrnVal(prs!nSaldoFinImporte, 16, 2) & " " _
                  & IIf(pnMoneda = 2, sImporteME, "") _
                  & oImpresora.gPrnCondensadaOFF
      End If
      nLin = nLin + 1
      lnCont = lnCont + 1
      If lnCont Mod 600 = 0 Then
         lsImpre = lsImpre & sTexto
         sTexto = ""
      End If
   End If
   prs.MoveNext
Loop
Set oBalanceGet = Nothing

If pbExcel And psTipo <> 0 Then
    xlHoja1.Cells(lnFila, 1) = "TOTAL MONEDA NACIONAL"
    xlHoja1.Cells(lnFila, 3) = sTotales(1, Val(sCtaClase), 2) ''prs!nSaldoIniImporte
    xlHoja1.Cells(lnFila, 4) = sTotales(1, Val(sCtaClase), 0) ''prs!nDebe
    xlHoja1.Cells(lnFila, 5) = sTotales(1, Val(sCtaClase), 1)  ''prs!nHaber
    xlHoja1.Cells(lnFila, 6) = sTotales(1, Val(sCtaClase), 3) ''prs!nSaldoFinImporte
    xlHoja1.Range("A" & lnFila & ":F" & lnFila).Font.Bold = True
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "TOTAL MONEDA EXTRANJERA"
    xlHoja1.Cells(lnFila, 3) = sTotales(2, Val(sCtaClase), 2)
    xlHoja1.Cells(lnFila, 4) = sTotales(2, Val(sCtaClase), 0)
    xlHoja1.Cells(lnFila, 5) = sTotales(2, Val(sCtaClase), 1)
    xlHoja1.Cells(lnFila, 6) = sTotales(2, Val(sCtaClase), 3)
    xlHoja1.Range("A" & lnFila & ":F" & lnFila).Font.Bold = True
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "TOTAL AJUSTE POR INFLACION"
    xlHoja1.Cells(lnFila, 3) = sTotales(0, Val(sCtaClase), 2)
    xlHoja1.Cells(lnFila, 4) = sTotales(0, Val(sCtaClase), 0)
    xlHoja1.Cells(lnFila, 5) = sTotales(0, Val(sCtaClase), 1)
    xlHoja1.Cells(lnFila, 6) = sTotales(0, Val(sCtaClase), 3)
    xlHoja1.Range("A" & lnFila & ":F" & lnFila).Font.Bold = True
    lnFila = lnFila + 1

    xlHoja1.Range("A" & lnFila - 3 & ":F" & lnFila - 1).Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range("A" & lnFila - 3 & ":F" & lnFila - 1).Borders(xlEdgeTop).Weight = xlThin
    xlHoja1.Range("A" & lnFila - 3 & ":F" & lnFila - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlHoja1.Range("A" & lnFila - 3 & ":F" & lnFila - 1).Borders(xlEdgeBottom).Weight = xlMedium

    lnFila = lnFila + 1

    xlHoja1.Cells(lnFila, 2) = "TOTALES "
    xlHoja1.Cells(lnFila, 4) = nTotDebe
    xlHoja1.Cells(lnFila, 5) = nTotHaber
    xlHoja1.Range("A" & lnFila & ":G" & lnFila).Font.Bold = True
    xlHoja1.Range("C7:G" & lnFila).NumberFormat = "#,##0.00"

    xlHoja1.Range("A" & lnFila & ":F" & lnFila).BorderAround xlContinuous, xlMedium

ElseIf pbExcel And psTipo = 0 Then
    xlHoja1.Cells(lnFila, 2) = "TOTALES "
    xlHoja1.Cells(lnFila, 4) = nTotDebe
    xlHoja1.Cells(lnFila, 5) = nTotHaber
    xlHoja1.Range("A" & lnFila & ":G" & lnFila).Font.Bold = True
    xlHoja1.Range("C7:G" & lnFila).NumberFormat = "#,##0.00"
    xlHoja1.Range("A" & lnFila & ":F" & lnFila).BorderAround xlContinuous, xlMedium

ElseIf pbExcel = False And psTipo <> 0 Then
    nLin = nLin + 1
    LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & " --------------------------------------------------------------------------------------------------------------------------------------------"
    LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & oImpresora.gPrnBoldON & "   " & "TOTAL MONEDA NACIONAL     " & Space(44) & PrnVal(sTotales(1, Val(sCtaClase), 2), 16, 2) & " " _
          & PrnVal(sTotales(1, Val(sCtaClase), 0), 16, 2) & " " & PrnVal(sTotales(1, Val(sCtaClase), 1), 16, 2) & " " _
          & PrnVal(sTotales(1, Val(sCtaClase), 3), 16, 2) & oImpresora.gPrnBoldOFF & oImpresora.gPrnCondensadaOFF
    nLin = nLin + 1
    LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & oImpresora.gPrnBoldON & "   " & "TOTAL MONEDA EXTRANJERA   " & Space(44) & PrnVal(sTotales(2, Val(sCtaClase), 2), 16, 2) & " " _
          & PrnVal(sTotales(2, Val(sCtaClase), 0), 16, 2) & " " & PrnVal(sTotales(2, Val(sCtaClase), 1), 16, 2) & " " _
          & PrnVal(sTotales(2, Val(sCtaClase), 3), 16, 2) & oImpresora.gPrnBoldOFF & oImpresora.gPrnCondensadaOFF
    nLin = nLin + 1
    LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & oImpresora.gPrnBoldON & "   " & "TOTAL AJUSTE POR INFLACION" & Space(44) & PrnVal(sTotales(0, Val(sCtaClase), 2), 16, 2) & " " _
          & PrnVal(sTotales(0, Val(sCtaClase), 0), 16, 2) & " " & PrnVal(sTotales(0, Val(sCtaClase), 1), 16, 2) & " " _
          & PrnVal(sTotales(0, Val(sCtaClase), 3), 16, 2) & oImpresora.gPrnBoldOFF & oImpresora.gPrnCondensadaOFF
    nLin = nLin + 1

   LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & " --------------------------------------------------------------------------------------------------------------------------------------------" & IIf(pnMoneda = 2, String(17, "-"), "")
   LineaAuditoria sTexto, oImpresora.gPrnBoldON & Space(88) & "TOTALES  " & PrnVal(nTotDebe, 14, 2) & " " & PrnVal(nTotHaber, 14, 2) & oImpresora.gPrnBoldOFF & oImpresora.gPrnCondensadaOFF, 2
   LineaAuditoria sTexto, " ============================================================================================================================================" & IIf(pnMoneda = 2, String(17, "="), "") & oImpresora.gPrnCondensadaOFF

ElseIf pbExcel = False And psTipo = 0 Then
   LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & " --------------------------------------------------------------------------------------------------------------------------------------------" & IIf(pnMoneda = 2, String(17, "-"), "")
   LineaAuditoria sTexto, oImpresora.gPrnBoldON & Space(88) & "TOTALES  " & PrnVal(nTotDebe, 14, 2) & " " & PrnVal(nTotHaber, 14, 2) & oImpresora.gPrnBoldOFF & oImpresora.gPrnCondensadaOFF, 2
   LineaAuditoria sTexto, " ============================================================================================================================================" & IIf(pnMoneda = 2, String(17, "="), "") & oImpresora.gPrnCondensadaOFF
End If
Set prs = Nothing
ImprimeBalanceComprobacion = lsImpre & sTexto
End Function

Private Function ImprimeBalanceComprobacionCabecera(pdFechaIni As Date, pdFechaFin As Date, pnTpoCambio As Currency, P As Integer, pnTipoBala As Integer, pnMoneda As Integer, pdFecha As Date, ByRef nLin As Integer, ByRef pnLinPage As Integer) As String
Dim sTexto As String
If nLin > pnLinPage - 6 Then
   LineaAuditoria sTexto, CabeceraCuscoAuditoria(" B A L A N C E   D E   C O M P R O B A C I O N   (" & IIf(pnTipoBala = 1, "HISTORICO", "AJUSTADO") & ") ", P, IIf(pnMoneda = 1, "S/.", IIf(pnMoneda = 2, "$", "")), 120, , Format(pdFecha, gsFormatoFechaView), "CMAC MAYNAS S.A."), 0
   If pnMoneda = 0 Then
      LineaAuditoria sTexto, Centra(" C O N S O L I D A D O ", 120)
   End If
   LineaAuditoria sTexto, Centra("( DEL " & pdFechaIni & " AL " & pdFechaFin & ")", 120) & oImpresora.gPrnBoldOFF
   If pnMoneda = 2 Then
      LineaAuditoria sTexto, " TIPO DE CAMBIO : " & Format(pnTpoCambio, "###,##0.00##")
   End If
   LineaAuditoria sTexto, oImpresora.gPrnCondensadaON, 0
   LineaAuditoria sTexto, " =============================================================================================================================================" & IIf(pnMoneda = 2, String(17, "="), "")
   If gsCodCMAC = "102" Then
      LineaAuditoria sTexto, "                                                            CUENTA               SALDO                    MOVIMIENTO                  SALDO   " & IIf(pnMoneda = 2, "      SALDO ME", "")
      LineaAuditoria sTexto, "    D E S C R I P C I O N                                   CONTABLE             INICIAL    ----------------------------------      ACUMULADO " & IIf(pnMoneda = 2, "     ACUMULADO", "")
   Else
      LineaAuditoria sTexto, "  CUENTA                                                                         SALDO                    MOVIMIENTO                  SALDO   " & IIf(pnMoneda = 2, "      SALDO ME", "")
      LineaAuditoria sTexto, "  CONTABLE         D E S C R I P C I O N                                         INICIAL    ----------------------------------      ACUMULADO " & IIf(pnMoneda = 2, "     ACUMULADO", "")
   End If
   LineaAuditoria sTexto, "                                                                                                    DEBE            HABER                     "
   LineaAuditoria sTexto, " ---------------------------------------------------------------------------------------------------------------------------------------------" & IIf(pnMoneda = 2, String(17, "-"), "") & oImpresora.gPrnCondensadaOFF
   nLin = 10
End If
ImprimeBalanceComprobacionCabecera = sTexto
End Function

Public Function ImprimeBalanceComprobacionCabExcel(pdFechaIni As Date, pdFechaFin As Date, pnTpoCambio As Currency, P As Integer, pnTipoBala As Integer, pnMoneda As Integer, pdFecha As Date, ByRef nLin As Integer, ByRef pnLinPage As Integer, Optional xlHoja1 As Excel.Worksheet) As String
xlHoja1.PageSetup.LeftMargin = 1.5
xlHoja1.PageSetup.RightMargin = 0
xlHoja1.PageSetup.BottomMargin = 1
xlHoja1.PageSetup.TopMargin = 1
xlHoja1.PageSetup.Zoom = 70
xlHoja1.Cells(1, 1) = "CMAC MAYNAS S.A."
xlHoja1.Cells(1, 6) = "Fecha : " & Format(pdFecha, gsFormatoFechaView)
xlHoja1.Cells(2, 1) = " B A L A N C E   D E   C O M P R O B A C I O N   (" & IIf(pnTipoBala = 1, "HISTORICO", "AJUSTADO") & ") "
If pnMoneda = 0 Then
   xlHoja1.Cells(2, 1) = xlHoja1.Cells(2, 1) & " C O N S O L I D A D O "
End If
xlHoja1.Cells(3, 1) = "( DEL " & pdFechaIni & " AL " & pdFechaFin & ")"
If pnMoneda = 2 Then
   xlHoja1.Cells(4, 1) = " TIPO DE CAMBIO : " & Format(pnTpoCambio, "###,##0.00##")
End If
xlHoja1.Cells(5, 2) = "CUENTA CONTABLE"
xlHoja1.Cells(5, 1) = "D E S C R I P C I O N"
xlHoja1.Cells(5, 3) = "SALDO INICIAL"
xlHoja1.Cells(5, 4) = "MOVIMIENTO"
xlHoja1.Cells(6, 4) = "DEBE"
xlHoja1.Cells(6, 5) = "HABER"
xlHoja1.Cells(5, 6) = "SALDO ACUMULADO"
If pnMoneda = 2 Then
   xlHoja1.Cells(5, 7) = "SALDO ME  ACUMULADO"
   xlHoja1.Range("G5:G6").MergeCells = True
End If
xlHoja1.Range("A1:G6").Font.Bold = True
xlHoja1.Range("A2:G2").MergeCells = True
xlHoja1.Range("A3:G3").MergeCells = True
xlHoja1.Range("A5:A6").MergeCells = True
xlHoja1.Range("B5:B6").MergeCells = True
xlHoja1.Range("C5:C6").MergeCells = True
xlHoja1.Range("D5:E5").MergeCells = True
xlHoja1.Range("F5:F6").MergeCells = True
xlHoja1.Range("A2:G3").HorizontalAlignment = xlCenter
xlHoja1.Range("A5:G6").HorizontalAlignment = xlCenter
xlHoja1.Range("A5:G6").VerticalAlignment = xlCenter
xlHoja1.Range("A5:G6").WrapText = True
xlHoja1.Range("A5:" & IIf(pnMoneda = "2", "G", "F") & "6").BorderAround xlContinuous, xlMedium
xlHoja1.Range("A5:" & IIf(pnMoneda = "2", "G", "F") & "6").Borders(xlInsideHorizontal).LineStyle = xlContinuous
xlHoja1.Range("A5:" & IIf(pnMoneda = "2", "G", "F") & "6").Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("A1:A1").ColumnWidth = 45
xlHoja1.Range("B1:B1").ColumnWidth = 20
xlHoja1.Range("C1:G1").ColumnWidth = 14
End Function

Public Function LineaAuditoria(psVarImpre As String, psTexto As String, Optional pnLineas As Integer = 1, Optional ByRef pnLinCnt As Integer = 0) As String
Dim K As Integer
psVarImpre = psVarImpre & psTexto
For K = 1 To pnLineas
   psVarImpre = psVarImpre & oImpresora.gPrnSaltoLinea
   pnLinCnt = pnLinCnt + 1
Next
End Function

Public Function CabeceraCuscoAuditoria(sTit As String, P As Integer, psSimbolo As String, pnColPage As Integer, Optional sCabe As String = "", Optional sFecha As String = "", Optional psEmprLogo As String = " CMAC - CUSCO ") As String
CabeceraCuscoAuditoria = ""
If sFecha = "" Then
   sFecha = Date
End If
   If P > 0 Then CabeceraCuscoAuditoria = oImpresora.gPrnSaltoPagina
   P = P + 1
   CabeceraCuscoAuditoria = CabeceraCuscoAuditoria + ImpreFormat(psEmprLogo, 80) & sFecha & "-" & Format(Time, "hh:mm:ss") & "            Pag. " & Format(P, "000") & oImpresora.gPrnSaltoLinea
   CabeceraCuscoAuditoria = CabeceraCuscoAuditoria + Space(72) & oImpresora.gPrnSaltoLinea
   CabeceraCuscoAuditoria = CabeceraCuscoAuditoria + oImpresora.gPrnBoldON & Centra(sTit, pnColPage) & oImpresora.gPrnBoldOFF & oImpresora.gPrnSaltoLinea
   If psSimbolo <> "" Then
      CabeceraCuscoAuditoria = CabeceraCuscoAuditoria + oImpresora.gPrnBoldON & Centra(" M O N E D A   " & IIf(psSimbolo = gcMN, "N A C I O N A L ", "E X T R A N J E R A "), pnColPage) & oImpresora.gPrnBoldOFF & oImpresora.gPrnSaltoLinea
   End If
   If sCabe <> "" Then
      CabeceraCuscoAuditoria = CabeceraCuscoAuditoria & sCabe
   End If
End Function

