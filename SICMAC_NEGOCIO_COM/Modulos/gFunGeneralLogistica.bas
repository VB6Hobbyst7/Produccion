Attribute VB_Name = "gFunGeneralLogistica"
Global gbBitTCPonderado As Boolean

Public Function GetTipCambioLog(pdFecha As Date, Optional LeedeAdmin As Boolean = True) As Boolean
    Dim oDGeneral As DGeneral
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    
    If LeedeAdmin Then
        oCon.AbreConexion 'Remota "07", , , "01"
        oCon.Ejecutar "Set dateformat mdy"
        
        sql = " Select top 1 nValVent,nValComp,nValFijo from dbcomunes..tipcambio where dFecCamb Between '" & Format(pdFecha, gcFormatoFecha) & "' And '" & Format(pdFecha, gcFormatoFecha) & " 23:59:59' order by dFecCamb desc"
        Set rs = oCon.CargaRecordSet(sql)
        
        If rs.EOF And rs.BOF Then
            gnTipCambio = 0
            gnTipCambioV = 0
            gnTipCambioC = 0
            
            gnTipCambioVE = 0
            gnTipCambioCE = 0
            gnTipCambioPonderado = 0
            
            MsgBox "Tipo de cambio no definido.", vbInformation, "Aviso"
        Else
            gnTipCambio = rs!nValFijo
            gnTipCambioV = rs!nValVent
            gnTipCambioC = rs!nValComp
        End If
        
        rs.Close
    Else
    
        Set oDGeneral = New DGeneral
        GetTipCambioLog = True
        gnTipCambio = 0
        gnTipCambioV = 0
        gnTipCambioC = 0
        
        gnTipCambioVE = 0
        gnTipCambioCE = 0
        gnTipCambioPonderado = 0
        
        gnTipCambio = oDGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
        gnTipCambioV = oDGeneral.EmiteTipoCambio(pdFecha, TCVenta)
        gnTipCambioC = oDGeneral.EmiteTipoCambio(pdFecha, TCCompra)
        
        gnTipCambioVE = oDGeneral.EmiteTipoCambio(pdFecha, TCVentaEsp)
        gnTipCambioCE = oDGeneral.EmiteTipoCambio(pdFecha, TCCompraEsp)
        gnTipCambioPonderado = oDGeneral.EmiteTipoCambio(pdFecha, TCPonderado)
        
        If gnTipCambio = 0 Then
            MsgBox "Tipo de Cambio aun no definido", vbInformation, "Aviso"
            GetTipCambioLog = False
        End If
    End If
End Function

Public Function ImprimeBalanceSituacion(pdFechaIni As Date, pdFechaFin As Date, pdFecha As Date, pnTipoBala As Integer, pnMoneda As Integer, nTotDebe As Currency, nTotHaber As Currency, Optional pnConCierreAnual As Integer = 0) As String
Dim prs1 As ADODB.Recordset
Set prs1 = New ADODB.Recordset
Dim P As Integer
Dim sTexto As String
Dim ntInicial As Currency, ntDebe As Currency, ntHaber As Currency
Dim lOrden As Boolean
Dim nFinal As Currency
Dim nFinalME As Currency
Dim nAcumME As Currency
Dim nTipCambio As Currency
lOrden = False
P = 0
sTexto = ""

Dim oBalGetTipo As COMNAuditoria.NBalanceCont
Set oBalGetTipo = New COMNAuditoria.NBalanceCont

If pnMoneda = 2 Then
   nTipCambio = oBalGetTipo.GetTipCambioBalance(Format(pdFechaFin, "yyyymmdd"))
End If

gFunGeneralContabilidad.LineaAuditoria sTexto, ImprimeBalanceSituacionCabecera(pdFechaIni, pdFechaFin, nTipCambio, P, pnTipoBala, pnMoneda, pdFecha), 4
gFunGeneralContabilidad.LineaAuditoria sTexto, "( I ) ACTIVO Y SALDOS DEUDORES "
gFunGeneralContabilidad.LineaAuditoria sTexto, "-------------------------------", 2
Set prs1 = oBalGetTipo.GetBalanceSituacion(pnTipoBala, pnMoneda, Month(pdFechaIni) + pnConCierreAnual, Year(pdFechaIni), "D")
nFinalME = 0
Do While Not prs1.EOF
      nFinal = prs1!nInicial + prs1!nDebe - prs1!nHaber
      If pnMoneda = 2 Then
         nFinalME = Round(nFinal / nTipCambio, 2)
      End If
      nAcumME = nAcumME + nFinalME
      Dim oGen As New DGeneral
      gFunGeneralContabilidad.LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & " " & prs1!cCtaContCod & Space(8) & Mid(oGen.CuentaNombre(prs1!cCtaContCod) & Space(60), 1, 60) & " " _
               & PrnVal(prs1!nInicial, 16, 2) & " " _
               & PrnVal(prs1!nDebe, 16, 2) & " " _
               & PrnVal(prs1!nHaber, 16, 2) & " " _
               & PrnVal(nFinal, 16, 2) & " " _
               & IIf(pnMoneda = 2, PrnVal(nFinalME, 16, 2), "") _
               & oImpresora.gPrnCondensadaOFF
      Set oGen = Nothing
   ntInicial = ntInicial + prs1!nInicial
   ntDebe = ntDebe + prs1!nDebe
   ntHaber = ntHaber + prs1!nHaber
   prs1.MoveNext
   If prs1.EOF Then Exit Do
   If Mid(prs1!cCtaContCod, 1, 1) = "8" And Not lOrden Then
      gFunGeneralContabilidad.LineaAuditoria sTexto, oImpresora.gPrnSaltoLinea & oImpresora.gPrnCondensadaON & Space(15) & "TOTAL ACTIVO Y SALDOS DEUDORES       " & Space(20) & PrnVal(ntInicial, 16, 2) & " " & PrnVal(ntDebe, 16, 2) & " " & PrnVal(ntHaber, 16, 2) & " " & PrnVal(ntInicial + ntDebe - ntHaber, 16, 2) & " " & IIf(pnMoneda = 2, PrnVal(nAcumME, 16, 2), "") & oImpresora.gPrnCondensadaOFF, 3
      nAcumME = 0
      gFunGeneralContabilidad.LineaAuditoria sTexto, "( II ) CUENTAS DE ORDEN DEUDORAS "
      gFunGeneralContabilidad.LineaAuditoria sTexto, "---------------------------------", 2
      ntInicial = 0: ntDebe = 0: ntHaber = 0
      lOrden = True
   End If
Loop
gFunGeneralContabilidad.LineaAuditoria sTexto, oImpresora.gPrnSaltoLinea & oImpresora.gPrnCondensadaON & Space(15) & "TOTAL CUENTAS DE ORDEN DEUDORAS      " & Space(20) & PrnVal(ntInicial, 16, 2) & " " & PrnVal(ntDebe, 16, 2) & " " & PrnVal(ntHaber, 16, 2) & " " & PrnVal(ntInicial + ntDebe - ntHaber, 16, 2) & " " & IIf(pnMoneda = 2, PrnVal(nAcumME, 16, 2), "") & oImpresora.gPrnCondensadaOFF

lOrden = False
ntInicial = 0: ntDebe = 0: ntHaber = 0
gFunGeneralContabilidad.LineaAuditoria sTexto, ImprimeBalanceSituacionCabecera(pdFechaIni, pdFechaFin, nTipCambio, P, pnTipoBala, pnMoneda, pdFecha), 4
gFunGeneralContabilidad.LineaAuditoria sTexto, "( III ) PASIVO Y SALDOS ACREEDORES"
gFunGeneralContabilidad.LineaAuditoria sTexto, "-----------------------------------", 2
Set prs1 = oBalGetTipo.GetBalanceSituacion(pnTipoBala, pnMoneda, Month(pdFechaIni) + pnConCierreAnual, Year(pdFechaFin), "A")
nFinalME = 0
Do While Not prs1.EOF
   nFinal = prs1!nInicial + prs1!nHaber - prs1!nDebe
   If pnMoneda = 2 Then
      nFinalME = Round(nFinal / nTipCambio, 2)
   End If
   nAcumME = nAcumME + nFinalME
   Set oGen = New DGeneral
   gFunGeneralContabilidad.LineaAuditoria sTexto, oImpresora.gPrnCondensadaON & " " & prs1!cCtaContCod & Space(8) & Mid(oGen.CuentaNombre(prs1!cCtaContCod) & Space(60), 1, 60) & " " _
         & PrnVal(prs1!nInicial, 16, 2) & " " _
         & PrnVal(prs1!nDebe, 16, 2) & " " _
         & PrnVal(prs1!nHaber, 16, 2) & " " _
         & PrnVal(nFinal, 16, 2) & " " _
         & IIf(pnMoneda = 2, PrnVal(nFinalME, 16, 2), "") _
         & oImpresora.gPrnCondensadaOFF
   Set oGen = Nothing
   ntInicial = ntInicial + prs1!nInicial
   ntDebe = ntDebe + prs1!nDebe
   ntHaber = ntHaber + prs1!nHaber
   prs1.MoveNext
   If prs1.EOF Then Exit Do
   If Mid(prs1!cCtaContCod, 1, 1) = "8" And Not lOrden Then
      gFunGeneralContabilidad.LineaAuditoria sTexto, oImpresora.gPrnSaltoLinea & oImpresora.gPrnCondensadaON & Space(15) & "TOTAL PASIVO Y SALDOS ACREEDORES     " & Space(20) & PrnVal(ntInicial, 16, 2) & " " & PrnVal(ntDebe, 16, 2) & " " & PrnVal(ntHaber, 16, 2) & " " & PrnVal(ntInicial - ntDebe + ntHaber, 16, 2) & " " & IIf(pnMoneda = 2, PrnVal(nAcumME, 16, 2), "") & oImpresora.gPrnCondensadaOFF, 3
      nAcumME = 0
      gFunGeneralContabilidad.LineaAuditoria sTexto, "( IV ) CUENTAS DE ORDEN ACREEDORAS"
      gFunGeneralContabilidad.LineaAuditoria sTexto, "-----------------------------------", 2
      ntInicial = 0: ntDebe = 0: ntHaber = 0
      lOrden = True
   End If
Loop
gFunGeneralContabilidad.LineaAuditoria sTexto, oImpresora.gPrnSaltoLinea & oImpresora.gPrnCondensadaON & Space(15) & "TOTAL CUENTAS DE ORDEN ACREEDORAS    " & Space(20) & PrnVal(ntInicial, 16, 2) & " " & PrnVal(ntDebe, 16, 2) & " " & PrnVal(ntHaber, 16, 2) & " " & PrnVal(ntInicial - ntDebe + ntHaber, 16, 2) & IIf(pnMoneda = 2, PrnVal(nAcumME, 16, 2), "") & oImpresora.gPrnCondensadaOFF, 2
gFunGeneralContabilidad.LineaAuditoria sTexto, oImpresora.gPrnBoldON & Space(15) & "TOTAL GENERAL" & Space(5) & "DEBE  : " & PrnVal(Val(Format(nTotDebe)), 16, 2)
gFunGeneralContabilidad.LineaAuditoria sTexto, Space(15) & "             " & Space(5) & "HABER : " & PrnVal(Val(Format(nTotHaber)), 16, 2) & oImpresora.gPrnBoldOFF

gFunGeneralContabilidad.LineaAuditoria sTexto, ImprimeBalanceSituacionCabecera(pdFechaIni, pdFechaFin, nTipCambio, P, pnTipoBala, pnMoneda, pdFecha), 4
sTexto = sTexto & oImpresora.gPrnSaltoLinea
sTexto = sTexto & oImpresora.gPrnSaltoLinea

If pnMoneda = 0 Then
      gFunGeneralContabilidad.LineaAuditoria sTexto, CentrarCadena("CALCULO DE LA UTILIDAD (HISTORICA)", 80)
      gFunGeneralContabilidad.LineaAuditoria sTexto, CentrarCadena("         (CONSOLIDADO)             ", 80)
      gFunGeneralContabilidad.LineaAuditoria sTexto, CentrarCadena("-----------------------------------", 80), 2
ElseIf pnMoneda = 2 Then
      gFunGeneralContabilidad.LineaAuditoria sTexto, CentrarCadena("CALCULO DE LA UTILIDAD (HISTORICA)", 80)
      gFunGeneralContabilidad.LineaAuditoria sTexto, CentrarCadena("         (MONEDA EXTRANJERA)             ", 80)
      gFunGeneralContabilidad.LineaAuditoria sTexto, CentrarCadena("-----------------------------------", 80), 2
Else
      gFunGeneralContabilidad.LineaAuditoria sTexto, CentrarCadena("CALCULO DE LA UTILIDAD (HISTORICA)", 80)
      gFunGeneralContabilidad.LineaAuditoria sTexto, CentrarCadena("         (MONEDA NACIONAL)             ", 80)
      gFunGeneralContabilidad.LineaAuditoria sTexto, CentrarCadena("-----------------------------------", 80), 2
End If

sTexto = sTexto & oImpresora.gPrnSaltoLinea
sTexto = sTexto & oImpresora.gPrnSaltoLinea

gFunGeneralContabilidad.LineaAuditoria sTexto, oBalGetTipo.ValidaBalance(True, pdFechaIni, pdFechaFin, pnTipoBala, pnMoneda)
'If prs1.State = adStateOpen Then prs1.Close:
Set prs1 = Nothing
Set oBalGetTipo = Nothing
ImprimeBalanceSituacion = sTexto

End Function

Private Function ImprimeBalanceSituacionCabecera(pdFechaIni As Date, pdFechaFin As Date, pnTpoCambio As Currency, P As Integer, pnTipoBala As Integer, pnMoneda As Integer, pdFecha As Date) As String
Dim sTexto As String
gFunGeneralContabilidad.LineaAuditoria sTexto, gFunGeneralContabilidad.CabeceraCuscoAuditoria(" B A L A N C E    D E    S I T U A C I O N   (" & IIf(pnTipoBala = 1, "HISTORICO", "AJUSTADO") & ") ", P, IIf(pnMoneda = 1, "S/.", IIf(pnMoneda = 2, "$", "")), 130, , Format(pdFecha, gsFormatoFecha), "CMAC MAYNAS S.A."), 0
If pnMoneda = 0 Then
   gFunGeneralContabilidad.LineaAuditoria sTexto, Centra(" C O N S O L I D A D O ", 130), 2
End If
gFunGeneralContabilidad.LineaAuditoria sTexto, Centra("( DEL " & pdFechaIni & " AL " & pdFechaFin & ")", 130) & oImpresora.gPrnBoldOFF
If pnMoneda = 2 Then
   gFunGeneralContabilidad.LineaAuditoria sTexto, " TIPO DE CAMBIO : " & Format(pnTpoCambio, "###,##0.00##")
End If
gFunGeneralContabilidad.LineaAuditoria sTexto, oImpresora.gPrnCondensadaON
gFunGeneralContabilidad.LineaAuditoria sTexto, " =============================================================================================================================================" & IIf(pnMoneda = 2, "=================", "")
gFunGeneralContabilidad.LineaAuditoria sTexto, "  CUENTA                                                                         SALDO                    MOVIMIENTO                  SALDO   "
gFunGeneralContabilidad.LineaAuditoria sTexto, "  CONTABLE         D E S C R I P C I O N                                         INICIAL    ----------------------------------      ACUMULADO " & IIf(pnMoneda = 2, "      ACUMULADO ", "")
gFunGeneralContabilidad.LineaAuditoria sTexto, "                                                                                                    DEBE            HABER             M.N.    " & IIf(pnMoneda = 2, "         M.E.   ", "")
gFunGeneralContabilidad.LineaAuditoria sTexto, " ---------------------------------------------------------------------------------------------------------------------------------------------" & IIf(pnMoneda = 2, "-----------------", "") & oImpresora.gPrnCondensadaOFF
ImprimeBalanceSituacionCabecera = sTexto
End Function


