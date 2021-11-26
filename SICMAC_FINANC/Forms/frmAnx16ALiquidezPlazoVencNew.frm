VERSION 5.00
Begin VB.Form frmAnx16ALiquidezPlazoVencNew 
   Caption         =   "Anexo 16: Liquidez por Plazos de Vecimiento"
   ClientHeight    =   975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5085
   Icon            =   "frmAnx16ALiquidezPlazoVencNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmAnx16ALiquidezPlazoVencNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'*** Nombre : frmAnx16ALiquidezPlazoVencNew
'*** Descripción : Formulario para la Generación del Anexo 16A - Cuadro de Liquidez por plazos de Vencimiento.
'*** Creación : NAGL - 20201211
'********************************************************************************

Public Sub ReporteAnexo16A(ByVal pdFecha As Date)
    Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja  As String
    Dim lsNombreAgencia As String
    Dim lsCodAgencia As String
    Dim lsMes As String
    Dim lnContador As Integer
    Dim lsArchivo As String
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim TituloProgress As String 'NAGL20170407
    Dim MensajeProgress As String 'NAGL20170407
    Dim oBarra As clsProgressBar 'NAGL20170407
    Dim nprogress As Integer 'NAGL20170407
    
    Dim oDbalanceCont As DbalanceCont
    Dim oDAnx As DAnexoRiesgos  'NAGL 20190518
    Dim nSaldoDiario1 As Currency
    Dim nSaldoDiario2 As Currency
    Dim pdFechaFinDeMes As Date
    Dim pdFechaFinDeMesMA As Date
    Dim dFechaAnte As Date
    Dim ldFechaPro As Date
    Dim nDia As Integer
    Dim oCambio As nTipoCambio
    Dim lnTipoCambioFC As Currency
    Dim lnTipoCambioProceso As Currency
    Dim nTipoCambioAn As Currency
    Dim loRs As ADODB.Recordset
    Dim oEst As New NEstadisticas 'NAGL
    Dim oPatrimonio As New DAnexoRiesgos 'NAGL 20170425
    
    Dim nTotalObligSugEncajMN As Currency
    'Dim nTotalTasaBaseEncajMN As Currency
    Dim nTotalObligSugEncajME As Currency
    'Dim nTotalTasaBaseEncajME  As Currency
    Dim lnTotalObligacionesAlDiaMN As Currency
    Dim lnTotalObligacionesAlDiaME  As Currency
    Dim lnTotalObligacionesAlDiaMN_MA As Currency
    Dim lnTotalObligacionesAlDiaME_MA  As Currency
    Dim lnTotalObligacionesAlDiaMN_DA As Currency
    Dim lnTotalObligacionesAlDiaME_DA As Currency
    Dim nTotalTasaBaseEncajMN_DA  As Currency
    Dim nTotalObligSugEncajMN_DA As Currency
    Dim nTotalTasaBaseEncajME_DA As Currency
    Dim nTotalObligSugEncajME_DA As Currency
    Dim oRs1 As ADODB.Recordset
    Dim ix As Integer
    Dim lnTotalObligacionesAlDiaPromedioMN As Currency
    Dim lnTotalObligacionesAlDiaPromedioME  As Currency
    Dim lnTotalObligacionesBase As Currency 'NAGL
    Dim lnTotalObligacionesxDiasDeclarar As Currency 'NAGL
    Dim lnTotalObligacionesAlDiaPromedioMN_MA As Currency
    Dim lnTotalObligacionesAlDiaPromedioME_MA  As Currency
    Dim nPromedTMMN As Double
    Dim nPromedTMME As Double
    Dim nTasaExigibleME As Currency 'NAGL
    Dim nTasaBaseMaginalMN As Currency
    Dim nTasaBaseMaginalME As Currency
    Dim nExigibleMaginalMN As Currency
    Dim nExigibleMaginalME As Currency
    Dim lnEncajeExigibleRGMN As Currency
    Dim lnEncajeExigibleRGME As Currency
    
    Dim lnOtrosDepMND1y2_C0 As Currency
    Dim lnOtrosDepMND4aM_C0 As Currency
    Dim lnOtrosDepMND1y2_C1 As Currency
    Dim lnOtrosDepMED1y2_C0 As Currency
    Dim lnOtrosDepMED4aM_C0 As Currency
    Dim lnOtrosDepMED1y2_C1 As Currency
    Dim lnSubastasMND1y2_C1 As Currency
    Dim lnSubastasMND1y2_C0 As Currency
    Dim lnSubastasMND4aM_C0 As Currency
    Dim lnSubastasMED1y2_C1 As Currency
    Dim lnSubastasMED1y2_C0 As Currency
    Dim lnSubastasMED4aM_C0 As Currency
    
    Dim lnSubastasMED3o3_C0 As Currency
    Dim lnSubastasMND3o3_C0 As Currency
    Dim lnSubastasMED3o3_C1 As Currency
    Dim lnSubastasMND3o3_C1 As Currency
    
    Dim lnOtrosDepMND3o3_C0 As Currency
    Dim lnOtrosDepMED3o3_C0 As Currency
    Dim lnOtrosDepMND3o3_C1 As Currency
    Dim lnOtrosDepMED3o3_C1 As Currency
    
    Dim lnSubastasMND4aM_C1 As Currency
    Dim lnSubastasMED4aM_C1 As Currency
    
    Dim nTotalAcredores20 As Currency
    Dim nTotalAcredores10 As Currency
    Dim nTotalAcredoresTo As Currency
    
    Dim nTotalDepositantes20 As Currency
    Dim nTotalDepositantes10 As Currency
    Dim nTotalDepositantesTo As Currency
    Dim nTotalAcredores201, nTotalAcredores101, nTotalAcredores202, nTotalAcredores102 As Currency
    Dim nTotalDepositantes201, nTotalDepositantes101, nTotalDepositantes202, nTotalDepositantes102 As Currency
    
On Error GoTo GeneraExcelErr
    'NAGL
    Set oBarra = New clsProgressBar
    oBarra.ShowForm frmAnx16ALiquidezPlazoVencNew
    oBarra.Max = 10
    nprogress = 0
    oBarra.Progress nprogress, "ANEXO 16A: Cuadro de Liquidez por plazos de Vencimiento", "GENERANDO EL ARCHIVO", "", vbBlue
    TituloProgress = "ANEXO 16A: Cuadro de Liquidez por plazos de Vencimiento"
    MensajeProgress = "GENERANDO EL ARCHIVO"
    'NAGL
    
    pdFechaFinDeMes = DateAdd("d", -Day(DateAdd("m", 1, pdFecha)), DateAdd("m", 1, pdFecha))
    pdFechaFinDeMesMA = DateAdd("d", -Day(pdFechaFinDeMes), pdFechaFinDeMes)
    nDia = Day(pdFecha)
    Set oDbalanceCont = New DbalanceCont
    Set oDAnx = New DAnexoRiesgos
    If nDia >= 15 Then
        dFechaAnte = DateAdd("d", -(nDia - 1), pdFecha)
    Else
        dFechaAnte = DateAdd("d", -(nDia - 1), DateAdd("m", -1, pdFecha))
    End If
    Set oCambio = New nTipoCambio
    
    If Month(pdFecha) = Month(DateAdd("d", 1, pdFecha)) Then
        lnTipoCambioFC = Format(oCambio.EmiteTipoCambio(pdFecha, TCFijoDia), "#,##0.0000")
    Else
        lnTipoCambioFC = Format(oCambio.EmiteTipoCambio(DateAdd("d", 1, pdFecha), TCFijoDia), "#,##0.0000")
    End If
    nTipoCambioAn = lnTipoCambioFC
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    
    lsArchivo = "Anexo16_SBS"
    lsArchivo1 = "\spooler\ANEXO_16A_" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If
    
    oBarra.Progress 1, TituloProgress, MensajeProgress, "", vbBlue

    '****************************OBLIGACIONES SUJETAS A ENCAJE*******************************
    lnTotalObligacionesAlDiaMN = 0 '
    lnTotalObligacionesAlDiaME = 0 '
    nTotalObligSugEncajMN = 0
    nTotalObligSugEncajME = 0
    lnTotalObligacionesAlDiaMN_DA = 0 '
    lnTotalObligacionesAlDiaME_DA = 0 '
    nTotalObligSugEncajMN_DA = 0
    nTotalObligSugEncajME_DA = 0

    ldFechaPro = DateAdd("d", -Day(pdFecha), pdFecha)

    For ix = 1 To Day(pdFecha)
         ldFechaPro = DateAdd("d", 1, ldFechaPro)
         If Month(ldFechaPro) = Month(DateAdd("d", 1, ldFechaPro)) Then
             lnTipoCambioProceso = Format(oCambio.EmiteTipoCambio(ldFechaPro, TCFijoDia), "#,##0.0000")
         Else
             lnTipoCambioProceso = Format(oCambio.EmiteTipoCambio(DateAdd("d", 1, ldFechaPro), TCFijoDia), "#,##0.0000")
         End If 'NAGL ERS 079-2016 20170407
         
         'SOLES
         nTotalObligSugEncajMN_DA = oDbalanceCont.SaldoCtas(1, "761201", ldFechaPro, pdFechaFinDeMes, lnTipoCambioProceso, lnTipoCambioProceso) 'Obligaciones Inmediatas
         nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "232") - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1, "232") 'Ahorros
         nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "233") + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "234")
         nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1, "233") - oDbalanceCont.SaldoCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1, "233") 'Depositos a plazo fijo
         'nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA - (oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "232") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "233") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "234"))
         lnTotalObligacionesAlDiaMN = lnTotalObligacionesAlDiaMN + nTotalObligSugEncajMN_DA '*************NAGL ERS079-2016 20170407

         nTotalObligSugEncajME_DA = oDbalanceCont.SaldoCtas(1, "762201", ldFechaPro, pdFechaFinDeMes, lnTipoCambioProceso, lnTipoCambioProceso) 'Obligaciones Inmediatas
         nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "232") - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2, "232") 'Ahorros
         nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "233") + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "234")
         nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2, "233") - oDbalanceCont.SaldoCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2, "233") 'Depositos a plazo fijo
         'nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA - (oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "232") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "233") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "234"))
         lnTotalObligacionesAlDiaME = lnTotalObligacionesAlDiaME + nTotalObligSugEncajME_DA '*************NAGL ERS079-2016 20170407
    Next ix
        
    lnTotalObligacionesBase = oDbalanceCont.ObtenerParamEncDiarioxCodigo("06") 'Tose Base Mes de Referencia ME
    nExigibleMaginalME = oDbalanceCont.ObtenerParamEncDiarioxCodigo("08")
    lnTotalObligacionesAlDiaPromedioME = Round(lnTotalObligacionesBase / oDbalanceCont.ObtenerParamEncDiarioxCodigo("10"), 2)
    nTasaExigibleME = Round(nExigibleMaginalME / oDbalanceCont.ObtenerParamEncDiarioxCodigo("10"), 2)
    lnTotalObligacionesxDiasDeclarar = lnTotalObligacionesAlDiaPromedioME * Day(pdFecha)
    nPromedTMME = Round((nTasaExigibleME / lnTotalObligacionesAlDiaPromedioME), 6)
    
    nTasaBaseMaginalME = (oDbalanceCont.ObtenerParamEncDiarioxCodigo("03") / 100)
    
    If lnTotalObligacionesAlDiaME > lnTotalObligacionesxDiasDeclarar Then
        lnEncajeExigibleRGME = (lnTotalObligacionesxDiasDeclarar * nPromedTMME) + (lnTotalObligacionesAlDiaME - lnTotalObligacionesxDiasDeclarar) * nTasaBaseMaginalME
    Else
        lnEncajeExigibleRGME = Round(lnTotalObligacionesAlDiaME * nPromedTMME, 2)
    End If
    
    lnEncajeExigibleRGMN = lnTotalObligacionesAlDiaMN * (oDbalanceCont.ObtenerParamEncDiarioxCodigo("32") / 100)
    '**********************************NAGL ERS 079-2016 20170407
    oBarra.Progress 2, TituloProgress, MensajeProgress, "", vbBlue

    lsNomHoja = "Anx16AMN"
    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    xlHoja1.Cells(3, 1) = "AL " & Format(pdFecha, "DD") & " DE " & UCase(Format(pdFecha, "MMMM")) & " DEL  " & Format(pdFecha, "YYYY")
    xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(3, 1)).Font.Bold = True
    
    '****************************OBLIGACIONES X CTA ***********************
     cargarPasivosObligxCtaAhorrosANX6 xlHoja1.Application, 1, pdFecha, lnTipoCambioFC
     cargarPasivosObligxCtaAhorrosANX6 xlHoja1.Application, 2, pdFecha, lnTipoCambioFC
     cargarPasivosObligxCtaCTSANX6 xlHoja1.Application, 1, pdFecha, lnTipoCambioFC
     cargarPasivosObligxCtaCTSANX6 xlHoja1.Application, 2, pdFecha, lnTipoCambioFC
     cargarDepositosInmovilizadosANX16 xlHoja1.Application, 1, pdFecha, lnTipoCambioFC '***NAGL ERS 079-2016 20170407
     cargarDepositosInmovilizadosANX16 xlHoja1.Application, 2, pdFecha, lnTipoCambioFC '***NAGL ERS 079-2016 20170407
     cargarObligacionesVistaANX16 xlHoja1.Application, 1, pdFecha, lnTipoCambioFC '***NAGL ERS 079-2016 20170407
     cargarObligacionesVistaANX16 xlHoja1.Application, 2, pdFecha, lnTipoCambioFC '***NAGL ERS 079-2016 20170407
     cargarPlazoFijoRangosPersoneriaRangoAnexo6 xlHoja1.Application, pdFecha, "1", Round(lnEncajeExigibleRGMN, 2), Round(lnEncajeExigibleRGME, 2), lnTipoCambioFC
     cargarFondeoAhPFCTSxProducto xlHoja1.Application, pdFecha, lnTipoCambioFC '***NAGL ERS006-2019 20190518
    '******************************************************************************
    oBarra.Progress 3, TituloProgress, MensajeProgress, "", vbBlue
    
    'Patrimonio Efectivo  NAGL 20170425
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oPatrimonio.CargaPatrimonioEfectivo(pdFecha)
    xlHoja1.Cells(1, 18) = Format(oRs1!dFechaPatrimonio, "dd/mm/yyyy")
    xlHoja1.Cells(2, 18) = Format(oRs1!nPatrimonioMN, "#,##0.00") 'NAGL 20170621
    xlHoja1.Cells(3, 18) = Format(lnTipoCambioFC, "#,##0.0000")
    xlHoja1.Cells(4, 18) = Format(oRs1!nPatrimonioME, "#,##0.00") 'NAGL 20170621
    
    'Disponible
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerOverNightTramosResidual("1", pdFecha, "2")
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(9, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If
    
    'Inversiones Disponibles para la venta
    Set oRs1 = oDbalanceCont.ObtenerInversionesVentaTramosResidual("1", pdFecha)
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(11, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If
    xlHoja1.Cells(11, 3) = CCur(xlHoja1.Cells(11, 3)) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("131402", pdFecha, "1", 0) 'CDBCRP
    '*************************NAGL ERS 079-2016 20170407
    
    Set oRs1 = New ADODB.Recordset
    'Inversiones a Vencimiento
    Set oRs1 = oDbalanceCont.ObtenerInversionesAVencimientoResidual(pdFecha, 1)
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(12, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If
    xlHoja1.Range(xlHoja1.Cells(9, 3), xlHoja1.Cells(12, 13)).NumberFormat = "#,##0.00;-#,##0.00"
    oBarra.Progress 4, TituloProgress, MensajeProgress, "", vbBlue
    
    'Creditos Vigentes - Tramos
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerCreditosTramosResidual("1", pdFecha)
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
           xlHoja1.Cells(oRs1!nPosFila, 3) = oRs1![1M]
           xlHoja1.Cells(oRs1!nPosFila, 4) = oRs1![2M]
           xlHoja1.Cells(oRs1!nPosFila, 5) = oRs1![3M]
           xlHoja1.Cells(oRs1!nPosFila, 6) = oRs1![4M]
           xlHoja1.Cells(oRs1!nPosFila, 7) = oRs1![5M]
           xlHoja1.Cells(oRs1!nPosFila, 8) = oRs1![6M]
           xlHoja1.Cells(oRs1!nPosFila, 9) = oRs1![7-9M]
           xlHoja1.Cells(oRs1!nPosFila, 10) = oRs1![10-12M]
           xlHoja1.Cells(oRs1!nPosFila, 11) = oRs1![1-2A]
           xlHoja1.Cells(oRs1!nPosFila, 12) = oRs1![2-5A]
           xlHoja1.Cells(oRs1!nPosFila, 13) = oRs1![m5A]
           oRs1.MoveNext
        Loop
    End If 'NAGL 202012 Según Acta N°094-2020
    
    'Creditos Refinanciados - Tramos
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerCreditosRTramosResidual("1", pdFecha)
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
           xlHoja1.Cells(oRs1!nPosFila, 3) = oRs1![1M]
           xlHoja1.Cells(oRs1!nPosFila, 4) = oRs1![2M]
           xlHoja1.Cells(oRs1!nPosFila, 5) = oRs1![3M]
           xlHoja1.Cells(oRs1!nPosFila, 6) = oRs1![4M]
           xlHoja1.Cells(oRs1!nPosFila, 7) = oRs1![5M]
           xlHoja1.Cells(oRs1!nPosFila, 8) = oRs1![6M]
           xlHoja1.Cells(oRs1!nPosFila, 9) = oRs1![7-9M]
           xlHoja1.Cells(oRs1!nPosFila, 10) = oRs1![10-12M]
           xlHoja1.Cells(oRs1!nPosFila, 11) = oRs1![1-2A]
           xlHoja1.Cells(oRs1!nPosFila, 12) = oRs1![2-5A]
           xlHoja1.Cells(oRs1!nPosFila, 13) = oRs1![m5A]
           oRs1.MoveNext
        Loop
    End If 'NAGL 202012 Según Acta N°094-2020
    
    'Intereses Devengados - Tramos
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerInteresDevengadoResidual(pdFecha, "1")
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
           xlHoja1.Cells(oRs1!nPosFila, 3) = oRs1!nSaldo
           oRs1.MoveNext
        Loop
    End If 'NAGL 202012 Según Acta N°094-2020
    
    'Cuentas por Cobrar - Operaciones de Reporte
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerCuentasxCobrarTramosResidual("1", pdFecha)
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(53, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If
    xlHoja1.Range(xlHoja1.Cells(53, 3), xlHoja1.Cells(53, 13)).NumberFormat = "#,##0.00;-#,##0.00"
    '**********NAGL ERS 079-2016 20170407
    
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerRestringidosxTramosxProducto("1", pdFecha, 0, "233")
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(60, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If
    
    Set oRs1 = oDbalanceCont.ObtenerDepositosSistemaFinancieroyOFIxTramosxOProducto(pdFecha, "1", "233")
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(62, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If
  
    Set oRs1 = oDbalanceCont.ObtenerMontoAdeudadosxPlazoyMonedaxCPxTramos(pdFecha, "1", "1")
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(63, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If
    'xlHoja1.Cells(63, 3) = CCur(xlHoja1.Cells(63, 3)) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("241802", pdFecha, "1", 0) 'Activado by NAGL 20190626 Primera Banda
    'Comentado by NAGL 202012
    
    Set oRs1 = oDbalanceCont.ObtenerMontoAdeudadosxPlazoyMonedaxCPxTramos(pdFecha, "1", "0")
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(64, 2 + oRs1!cRango) = oRs1!nSaldo
        oRs1.MoveNext
    Loop
    End If
    'xlHoja1.Cells(64, 3) = CCur(xlHoja1.Cells(64, 3)) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("241807", pdFecha, "1", 0) 'Activado by NAGL 20190626 Primera Banda
    'Comentado by NAGL 202012
    
    Set oRs1 = oDbalanceCont.ObtenerCtasxPagarOpeReactivaxTramos(pdFecha, "1")
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(68, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If 'NAGL 202012 Según ACTA N°094-2020
    xlHoja1.Range(xlHoja1.Cells(60, 3), xlHoja1.Cells(68, 13)).NumberFormat = "#,##0.00;-#,##0.00"
    
    'Inversiones a valor Razonable con cambios en resultados - supuesto
    nSaldoDiario1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1311", pdFecha, "1", 0)
    xlHoja1.Cells(75, 3) = nSaldoDiario1 '****NAGL ERS 079-2016 20170407
    
    nSaldoDiario1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2312", pdFecha, "1", 0)
    xlHoja1.Cells(97, 3) = nSaldoDiario1  'Depósitos Ifis/OFIs según Supuesto
    
    cargarDatosBalanceANX16 xlHoja1.Application, "1", pdFecha, lnTipoCambioFC '***NAGL ERS 079-2016 20170407
    oBarra.Progress 5, TituloProgress, MensajeProgress, "", vbBlue
    
    lsNomHoja = "Disponible"
    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    
    'InicioObligacionesVista
    Call PintaFondeoObligacionesVista(xlHoja1, pdFecha, lnTipoCambioFC, 1) 'NAGL ERS 079-2016 20170407
    xlHoja1.Cells(75, 2) = Format(lnTipoCambioFC, "#,##0.0000")
    oBarra.Progress 6, TituloProgress, MensajeProgress, "", vbBlue
    
    lsNomHoja = "Anx16AME"
    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    
    'Disponible
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerOverNightTramosResidual("2", pdFecha, "2")
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(9, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If
    
    'Inversiones Disponibles para la venta
    Set oRs1 = oDbalanceCont.ObtenerInversionesVentaTramosResidual("2", pdFecha)
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(11, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If
    xlHoja1.Cells(11, 3) = CCur(xlHoja1.Cells(11, 3)) + (oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("132402", pdFecha, "2", lnTipoCambioFC) / lnTipoCambioFC) 'CDBCRP
    '*************************NAGL ERS 079-2016 20170407
    
    Set oRs1 = New ADODB.Recordset
    'Inversiones a Vencimiento
    Set oRs1 = oDbalanceCont.ObtenerInversionesAVencimientoResidual(pdFecha, 2)
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(12, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If
    xlHoja1.Range(xlHoja1.Cells(9, 3), xlHoja1.Cells(12, 13)).NumberFormat = "#,##0.00;-#,##0.00"
   
    'Creditos Vigentes - Tramos
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerCreditosTramosResidual("2", pdFecha)
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
           xlHoja1.Cells(oRs1!nPosFila, 3) = oRs1![1M]
           xlHoja1.Cells(oRs1!nPosFila, 4) = oRs1![2M]
           xlHoja1.Cells(oRs1!nPosFila, 5) = oRs1![3M]
           xlHoja1.Cells(oRs1!nPosFila, 6) = oRs1![4M]
           xlHoja1.Cells(oRs1!nPosFila, 7) = oRs1![5M]
           xlHoja1.Cells(oRs1!nPosFila, 8) = oRs1![6M]
           xlHoja1.Cells(oRs1!nPosFila, 9) = oRs1![7-9M]
           xlHoja1.Cells(oRs1!nPosFila, 10) = oRs1![10-12M]
           xlHoja1.Cells(oRs1!nPosFila, 11) = oRs1![1-2A]
           xlHoja1.Cells(oRs1!nPosFila, 12) = oRs1![2-5A]
           xlHoja1.Cells(oRs1!nPosFila, 13) = oRs1![m5A]
           oRs1.MoveNext
        Loop
    End If 'NAGL 202012 Según Acta N°094-2020
    
    'Creditos Refinanciados - Tramos
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerCreditosRTramosResidual("2", pdFecha)
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
           xlHoja1.Cells(oRs1!nPosFila, 3) = oRs1![1M]
           xlHoja1.Cells(oRs1!nPosFila, 4) = oRs1![2M]
           xlHoja1.Cells(oRs1!nPosFila, 5) = oRs1![3M]
           xlHoja1.Cells(oRs1!nPosFila, 6) = oRs1![4M]
           xlHoja1.Cells(oRs1!nPosFila, 7) = oRs1![5M]
           xlHoja1.Cells(oRs1!nPosFila, 8) = oRs1![6M]
           xlHoja1.Cells(oRs1!nPosFila, 9) = oRs1![7-9M]
           xlHoja1.Cells(oRs1!nPosFila, 10) = oRs1![10-12M]
           xlHoja1.Cells(oRs1!nPosFila, 11) = oRs1![1-2A]
           xlHoja1.Cells(oRs1!nPosFila, 12) = oRs1![2-5A]
           xlHoja1.Cells(oRs1!nPosFila, 13) = oRs1![m5A]
           oRs1.MoveNext
        Loop
    End If 'NAGL 202012 Según Acta N°094-2020
    
    'Intereses Devengados - Tramos
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerInteresDevengadoResidual(pdFecha, "2")
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
           xlHoja1.Cells(oRs1!nPosFila, 3) = Round(oRs1!nSaldo / lnTipoCambioFC, 2)
           oRs1.MoveNext
        Loop
    End If 'NAGL 202012 Según Acta N°094-2020
    
    'Cuentas por Cobrar - Operaciones de Reporte
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerCuentasxCobrarTramosResidual("2", pdFecha)
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(53, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If
    xlHoja1.Range(xlHoja1.Cells(53, 3), xlHoja1.Cells(53, 13)).NumberFormat = "#,##0.00;-#,##0.00"
    '**********NAGL ERS 079-2016 20170407
    
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerRestringidosxTramosxProducto("2", pdFecha, 0, "233")
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(60, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If
    
    Set oRs1 = oDbalanceCont.ObtenerDepositosSistemaFinancieroyOFIxTramosxOProducto(pdFecha, "2", "233")
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(62, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If
  
    Set oRs1 = oDbalanceCont.ObtenerMontoAdeudadosxPlazoyMonedaxCPxTramos(pdFecha, "2", "1", lnTipoCambioFC) 'NAGL202012 Agregó lnTipoCambioFC
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(63, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If
    'xlHoja1.Cells(63, 3) = CCur(xlHoja1.Cells(63, 3)) + Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("242802", pdFecha, "2", lnTipoCambioFC) / lnTipoCambioFC, 2) 'Activado by NAGL 20190626 Primera Banda
    'Comentado by NAGL 202012
    
    Set oRs1 = oDbalanceCont.ObtenerMontoAdeudadosxPlazoyMonedaxCPxTramos(pdFecha, "2", "0", lnTipoCambioFC) 'NAGL202012 Agregó lnTipoCambioFC
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(64, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If
    'xlHoja1.Cells(64, 3) = CCur(xlHoja1.Cells(64, 3)) + Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("242807", pdFecha, "2", lnTipoCambioFC) / lnTipoCambioFC, 2) 'Activado by NAGL 20190626 Primera Banda
    'Comentado by NAGL 202012
    
    Set oRs1 = oDbalanceCont.ObtenerCtasxPagarOpeReactivaxTramos(pdFecha, "2")
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            xlHoja1.Cells(68, 2 + oRs1!cRango) = oRs1!nSaldo
            oRs1.MoveNext
        Loop
    End If 'NAGL 202012 Según ACTA N°094-2020
    xlHoja1.Range(xlHoja1.Cells(60, 3), xlHoja1.Cells(68, 13)).NumberFormat = "#,##0.00;-#,##0.00"
    
    'Inversiones a valor Razonable con cambios en resultados - supuesto
    nSaldoDiario1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1321", pdFecha, "2", lnTipoCambioFC) / lnTipoCambioFC
    xlHoja1.Cells(75, 3) = nSaldoDiario1 '****NAGL ERS 079-2016 20170407
    
    'Depósitos de empresas del sistema financiero y OFI
    nSaldoDiario1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2322", pdFecha, "2", lnTipoCambioFC)
    xlHoja1.Cells(97, 3) = Round(nSaldoDiario1 / lnTipoCambioFC, 2)  'Depósitos Ifis/OFIs según Supuesto
    
    cargarDatosBalanceANX16 xlHoja1.Application, "2", pdFecha, lnTipoCambioFC '***NAGL ERS 079-2016 20170407
    oBarra.Progress 7, TituloProgress, MensajeProgress, "", vbBlue
    
    lsNomHoja = "DisponibleDolares"
        For Each xlHoja1 In xlsLibro.Worksheets
           If xlHoja1.Name = lsNomHoja Then
                xlHoja1.Activate
             lbExisteHoja = True
            Exit For
           End If
        Next
        If lbExisteHoja = False Then
            Set xlHoja1 = xlsLibro.Worksheets
            xlHoja1.Name = lsNomHoja
        End If
    'InicioObligacionesVista
    Call PintaFondeoObligacionesVista(xlHoja1, pdFecha, lnTipoCambioFC, 2) '***NAGL ERS 079-2016 20170407
    
    lsNomHoja = "Anx16AInd"
    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    
    'InicioAhorro
    'stp_sel_Indicadores16Acredores
    nTotalAcredoresTo = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2102", pdFecha, "0", lnTipoCambioFC)
    nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("210303", pdFecha, "0", lnTipoCambioFC)
    nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("210305", pdFecha, "0", lnTipoCambioFC)
    nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2107", pdFecha, "0", lnTipoCambioFC)
    nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2302", pdFecha, "0", lnTipoCambioFC)
    nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2303", pdFecha, "0", lnTipoCambioFC)
    nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("24", pdFecha, "0", lnTipoCambioFC)
    nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("26", pdFecha, "0", lnTipoCambioFC)
    
    'Depositantes
    nTotalDepositantesTo = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2102", pdFecha, "0", lnTipoCambioFC)
    nTotalDepositantesTo = nTotalDepositantesTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("210303", pdFecha, "0", lnTipoCambioFC)
    nTotalDepositantesTo = nTotalDepositantesTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("210305", pdFecha, "0", lnTipoCambioFC)
    nTotalDepositantesTo = nTotalDepositantesTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2107", pdFecha, "0", lnTipoCambioFC)
    nTotalDepositantesTo = nTotalDepositantesTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2302", pdFecha, "0", lnTipoCambioFC)
    nTotalDepositantesTo = nTotalDepositantesTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2303", pdFecha, "0", lnTipoCambioFC)
    
    '**************************************ERS 079-2016 20170407
    'DEL REPORTE 1 - BCRP
    xlHoja1.Cells(4, 13) = Round(lnEncajeExigibleRGMN, 2)
    xlHoja1.Cells(5, 13) = Day(pdFecha)
    xlHoja1.Cells(4, 14) = Round(lnEncajeExigibleRGME, 2)
    
    'DE BALANCE CONSOLIDADO
    xlHoja1.Cells(9, 13) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("24", pdFecha, "0", lnTipoCambioFC)
    xlHoja1.Cells(10, 13) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1", pdFecha, "0", lnTipoCambioFC)
    
    'De Activos Liquidos Anexo 15C
    
    Dim CajasFFMN As Currency, CajasFFME As Currency
    Dim FondosBCRPMN As Currency, FondosBCRPME As Currency
    Dim FondosSFNMN As Currency, FondosSFNME As Currency
    Dim FondosNActNMN As Currency, FondosNActNME As Currency
    Dim ValorBCRPMN As Currency
    Dim ValorBCRPMN2 As Currency
    Dim ValorGCMN As Currency
    Dim SumaTotalCMN As Currency, SumaTotalCME As Currency
    
    FondosBCRPMN = 0
    FondosSFNMN = 0
    FondosNActNMN = 0
    ValorBCRPMN = 0
    ValorBCRPMN2 = 0
    CajasFFMN = 0
    FondosBCRPME = 0
    FondosSFNME = 0
    FondosNActNME = 0
    
    ldFechaPro = DateAdd("d", -Day(pdFecha), pdFecha)
    For ix = 1 To Day(pdFecha)
    ldFechaPro = DateAdd("d", 1, ldFechaPro)
    
    CajasFFMN = CajasFFMN + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "1", "A1", "300")
    FondosBCRPMN = FondosBCRPMN + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "1", "A1", "425")
    FondosSFNMN = FondosSFNMN + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "1", "A1", "450")
    FondosNActNMN = FondosNActNMN + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "1", "A1", "600")
    ValorBCRPMN = ValorBCRPMN + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "1", "A1", "725")
    ValorBCRPMN2 = ValorBCRPMN2 + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "1", "A1", "750")
    
    'ME
    CajasFFME = CajasFFME + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "2", "A1", "300")
    FondosBCRPME = FondosBCRPME + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "2", "A1", "425")
    FondosSFNME = FondosSFNME + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "2", "A1", "450")
    FondosNActNME = FondosNActNME + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "2", "A1", "600")
    Next ix
    
    SumaTotalCMN = CajasFFMN + FondosBCRPMN + FondosSFNMN + FondosNActNMN + ValorBCRPMN + ValorBCRPMN2
    SumaTotalCME = CajasFFME + FondosBCRPME + FondosSFNME + FondosNActNME
    
    xlHoja1.Cells(5, 6) = Round(SumaTotalCMN / Day(pdFecha), 2)
    xlHoja1.Cells(6, 6) = Round(SumaTotalCME / Day(pdFecha), 2)
    xlHoja1.Cells(7, 6) = Format(lnTipoCambioFC, "#,##0.0000")
    '**************************************ERS 079-2016 20170407
     
    Dim nSaldoFondeoTotal As Currency
    nSaldoDiario1 = oDbalanceCont.ObtenerAdeudadosExterior(pdFecha, "1", "0") + (oDbalanceCont.ObtenerAdeudadosExterior(pdFecha, "2", "0") / lnTipoCambioFC) '**NAGL
    xlHoja1.Cells(10, 3) = nSaldoDiario1 / oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2", pdFecha, "0", lnTipoCambioFC)
    'NAGL ERS 079-2016 20170407
    
    oBarra.Progress 8, TituloProgress, MensajeProgress, "", vbBlue
    '**********NAGL Agregó esta sección ERS006-2019******************'
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDAnx.ObtieneFondeoAnx16ANew(pdFecha, lnTipoCambioFC, "16AInd")
    If Not (oRs1.BOF Or oRs1.EOF) Then
        Do While Not oRs1.EOF
            If oRs1!nTipoCobertura = 1 Then
                    xlHoja1.Cells(11, 3) = Round(oRs1!nTotal / nTotalDepositantesTo, 2)
                    xlHoja1.Cells(33, 2) = Round(oRs1!nVenc30Dias, 2)
                    xlHoja1.Cells(33, 3) = Round(oRs1!nTotal, 2)
            ElseIf oRs1!nTipoCobertura = 0 Then
                    xlHoja1.Cells(34, 2) = Round(oRs1!nVenc30Dias, 2)
                    xlHoja1.Cells(34, 3) = Round(oRs1!nTotal, 2)
            ElseIf oRs1!nTipoCobertura = 3 Then
                    xlHoja1.Cells(29, 2) = Round(oRs1!nVenc30Dias, 2)
                    xlHoja1.Cells(29, 3) = Round(oRs1!nTotal, 2)
            End If
            oRs1.MoveNext
        Loop
    End If
    '***********************20190516*********************************'
    '****CUADRO DE VALIDACIÓN DE SALDOS
    xlHoja1.Cells(30, 3) = nTotalDepositantesTo
    xlHoja1.Cells(31, 3) = nTotalAcredoresTo
    xlHoja1.Cells(42, 3) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2", pdFecha, "0", lnTipoCambioFC)
    '**********NAGL ERS 079-2016 20170407
    
    cargarDetalleAcreedoresDepositantes xlHoja1.Application, 0, pdFecha, lnTipoCambioFC 'NAGL ERS 079-2016 20170407
    cargarDetalleAcreedoresDepositantes xlHoja1.Application, 1, pdFecha, lnTipoCambioFC 'NAGL ERS 079-2016 20170407
    oBarra.Progress 9, TituloProgress, MensajeProgress, "", vbBlue
    
    '**************************ERS 079-2016 20170407
    lsNomHoja = "Anx16BReg"
        For Each xlHoja1 In xlsLibro.Worksheets
           If xlHoja1.Name = lsNomHoja Then
                xlHoja1.Activate
             lbExisteHoja = True
            Exit For
           End If
        Next
        If lbExisteHoja = False Then
            Set xlHoja1 = xlsLibro.Worksheets
            xlHoja1.Name = lsNomHoja
        End If
        
    'Inversiones a valor razonable con cambios en resultados
    xlHoja1.Cells(76, 3) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("13110705", pdFecha, "1", 0)
    xlHoja1.Cells(76, 4) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("13210705", pdFecha, "2", lnTipoCambioFC) / lnTipoCambioFC
    
    'CDBCRP
    'xlHoja1.Cells(80, 3) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("131402", pdFecha, "1", 0)
    'xlHoja1.Cells(80, 4) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("132402", pdFecha, "2", lnTipoCambioFC) / lnTipoCambioFC
    'JIPR20211115  CAMBIA DE PÓSICIÓN DE 1304071901 A 13040201 CB BCRP - DSCTO 10%

    xlHoja1.Cells(79, 3) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("131402", pdFecha, "1", 0)
    xlHoja1.Cells(79, 4) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("132402", pdFecha, "2", lnTipoCambioFC) / lnTipoCambioFC
        
    'JIPR20211115  CAMBIA DE PÓSICIÓN DE 12M A MAS
    xlHoja1.Cells(80, 19) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("131405", pdFecha, "1", 0)
    xlHoja1.Cells(80, 20) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("132405", pdFecha, "2", lnTipoCambioFC) / lnTipoCambioFC
        
    Set oRs1 = oDbalanceCont.ObtenerInversionesVenta16B("1", pdFecha, "1314010101")
    Do While Not oRs1.EOF
        If (oRs1!cRango < 9) Then
            xlHoja1.Cells(79, 1 + oRs1!cRango * 2) = oRs1!nSaldo
        Else
            xlHoja1.Cells(79, 19) = CCur(xlHoja1.Cells(79, 19)) + oRs1!nSaldo
        End If
        oRs1.MoveNext
    Loop
                   
    Set oRs1 = oDbalanceCont.ObtenerInversionesVenta16B("1", pdFecha, "131407190[12]")
    Do While Not oRs1.EOF
        If (oRs1!cRango < 9) Then
            xlHoja1.Cells(81, 1 + oRs1!cRango * 2) = oRs1!nSaldo
        Else
            xlHoja1.Cells(81, 19) = CCur(xlHoja1.Cells(81, 19)) + oRs1!nSaldo
        End If
        oRs1.MoveNext
    Loop
         
    Set oRs1 = oDbalanceCont.ObtenerInversionesVenta16B("1", pdFecha, "13140507")
    Do While Not oRs1.EOF
        If (oRs1!cRango < 9) Then
            xlHoja1.Cells(82, 1 + oRs1!cRango * 2) = oRs1!nSaldo
        Else
            xlHoja1.Cells(82, 19) = CCur(xlHoja1.Cells(82, 19)) + oRs1!nSaldo
        End If
        oRs1.MoveNext
    Loop
             
    Set oRs1 = oDbalanceCont.ObtenerInversionesVenta16B("2", pdFecha, "1324010101")
    Do While Not oRs1.EOF
        If (oRs1!cRango < 9) Then
            xlHoja1.Cells(79, 2 + oRs1!cRango * 2) = oRs1!nSaldo
        Else
            xlHoja1.Cells(79, 20) = CCur(xlHoja1.Cells(79, 20)) + oRs1!nSaldo
        End If
        oRs1.MoveNext
    Loop

    Set oRs1 = oDbalanceCont.ObtenerInversionesVenta16B("2", pdFecha, "132407190[12]")
    Do While Not oRs1.EOF
        If (oRs1!cRango < 9) Then
            xlHoja1.Cells(81, 2 + oRs1!cRango * 2) = oRs1!nSaldo
        Else
            xlHoja1.Cells(81, 20) = CCur(xlHoja1.Cells(81, 20)) + oRs1!nSaldo
        End If
        oRs1.MoveNext
    Loop
         
    Set oRs1 = oDbalanceCont.ObtenerInversionesVenta16B("2", pdFecha, "13240507")
    Do While Not oRs1.EOF
        If (oRs1!cRango < 9) Then
            xlHoja1.Cells(82, 2 + oRs1!cRango * 2) = oRs1!nSaldo
        Else
            xlHoja1.Cells(82, 20) = CCur(xlHoja1.Cells(82, 20)) + oRs1!nSaldo
        End If
        oRs1.MoveNext
    Loop
    '********************************************NAGL ERS 079-2016 20170407

    oBarra.Progress 10, "ANEXO 16A: Cuadro de Liquidez por plazos de Vencimiento", "Generación Terminada", "", vbBlue
    oBarra.CloseForm frmAnx16ALiquidezPlazoVencNew
    Set oBarra = Nothing 'NAGL20170407

    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    
Exit Sub
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub

Private Sub cargarDetalleAcreedoresDepositantes(ByVal pobj_Excel As Excel.Application, ByVal pTipo As Integer, psFecha As Date, ByVal lnTipoCambioFC As Currency)
    Dim pcelda As Excel.Range
    Dim oDBalance As DbalanceCont '***NAGL
    Dim prs As ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Set prs = New ADODB.Recordset
    Set oDBalance = New DbalanceCont
    
    Set prs = oDBalance.ObtenerDetalleAcreedoresDepositantes(psFecha, lnTipoCambioFC, 20, pTipo)
    If Not prs.EOF Or prs.BOF Then
        If pTipo = 0 Then
            nFilas = 46
            Do While Not prs.EOF
                Set pcelda = pobj_Excel.Range("Anx16AInd!A" & nFilas)
                pcelda.value = IIf(pTipo = 0, prs(1), "")
                Set pcelda = pobj_Excel.Range("Anx16AInd!B" & nFilas)
                pcelda.value = IIf(pTipo = 0, prs(2), 0)
                Set pcelda = pobj_Excel.Range("Anx16AInd!C" & nFilas)
                pcelda.value = IIf(pTipo = 0, prs(3), 0)
                nFilas = nFilas + 1
                prs.MoveNext
            Loop
            ElseIf pTipo = 1 Then
            nFilas = 69
            Do While Not prs.EOF
                Set pcelda = pobj_Excel.Range("Anx16AInd!A" & nFilas)
                pcelda.value = IIf(pTipo = 1, prs(1), "")
                Set pcelda = pobj_Excel.Range("Anx16AInd!B" & nFilas)
                pcelda.value = IIf(pTipo = 1, prs(2), 0)
                Set pcelda = pobj_Excel.Range("Anx16AInd!C" & nFilas)
                pcelda.value = IIf(pTipo = 1, prs(3), 0)
                nFilas = nFilas + 1
                prs.MoveNext
            Loop
        End If
    End If
    Set prs = Nothing
End Sub '***NAGL ERS 079-2016 20170407*****'

Private Sub cargarPasivosObligxCtaAhorrosANX6(ByVal pobj_Excel As Excel.Application, ByVal cMoneda As String, psFecha As Date, ByVal lnTipoCambioFC As Currency)
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim objDAnexoRiesgos As DAnexoRiesgos
    Dim oDbalanceCont As New DbalanceCont '***NAGL
    Dim prs As ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Dim Cant As Integer
    Set objDAnexoRiesgos = New DAnexoRiesgos
    Cant = 0
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
        
        Set prs = oCtaIf.GetObligxCtaAhorrosSBSanx6A(Format(psFecha, "yyyymmdd"), cMoneda, "232")
        nFilas = 2
        If Not prs.EOF Or prs.BOF Then
            If cMoneda = 1 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("Ahorro2102!A" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("Ahorro2102!B" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(1), 0)
                    nFilas = nFilas + 1
                    Cant = Cant + 1
                    prs.MoveNext
                Loop
                
                If (Cant = 272) Then
                Set pcelda = pobj_Excel.Range("Ahorro2102!A274")
                pobj_Excel.Range("Ahorro2102!A274").value = psFecha
                Set pcelda = pobj_Excel.Range("Ahorro2102!B274")
                pobj_Excel.Range("Ahorro2102!B274").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2112", psFecha, "1", lnTipoCambioFC)
                End If '***NAGL ERS 079-2016 20170407***'
                
            ElseIf cMoneda = 2 Then
                Cant = 0
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("Ahorro2102!I" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("Ahorro2102!J" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(1), 0)
                    nFilas = nFilas + 1
                    prs.MoveNext
                    Cant = Cant + 1
                Loop
                If (Cant = 272) Then
                Set pcelda = pobj_Excel.Range("Ahorro2102!I274")
                pobj_Excel.Range("Ahorro2102!I274").value = psFecha
                Set pcelda = pobj_Excel.Range("Ahorro2102!J274")
                pobj_Excel.Range("Ahorro2102!J274").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2122", psFecha, "2", lnTipoCambioFC) / lnTipoCambioFC, 2)
                End If '***NAGL ERS 079-2016 20170407***'
            End If
        End If
        Set prs = Nothing
End Sub

Private Sub cargarPasivosObligxCtaCTSANX6(ByVal pobj_Excel As Excel.Application, ByVal cMoneda As String, psFecha As Date, ByVal lnTipoCambioFC As Currency)
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim objDAnexoRiesgos As DAnexoRiesgos
    Dim oDbalanceCont As New DbalanceCont '***NAGL
    Dim prs As ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Dim Cant As Integer
    Set objDAnexoRiesgos = New DAnexoRiesgos
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
    Cant = 0
        Set prs = oCtaIf.GetObligxCtaAhorrosSBSanx6A(Format(psFecha, "yyyymmdd"), cMoneda, "234")
        nFilas = 2
        If Not prs.EOF Or prs.BOF Then
            If cMoneda = 1 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("CTS210305!A" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("CTS210305!B" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(1), 0)
                    nFilas = nFilas + 1
                    Cant = Cant + 1
                    prs.MoveNext
                Loop
                
                If (Cant = 272) Then
                Set pcelda = pobj_Excel.Range("CTS210305!A274")
                pobj_Excel.Range("CTS210305!A274").value = psFecha
                Set pcelda = pobj_Excel.Range("CTS210305!B274")
                pobj_Excel.Range("CTS210305!B274").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("211305", psFecha, "1", lnTipoCambioFC)
                End If '***NAGL ERS 079-2016 20170407***'
                 
            End If
            If cMoneda = 2 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("CTS210305!I" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("CTS210305!J" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(1), 0)
                    nFilas = nFilas + 1
                    Cant = Cant + 1
                    prs.MoveNext
                Loop
                
                If (Cant = 272) Then
                Set pcelda = pobj_Excel.Range("CTS210305!I274")
                pobj_Excel.Range("CTS210305!I274").value = psFecha
                Set pcelda = pobj_Excel.Range("CTS210305!J274")
                pobj_Excel.Range("CTS210305!J274").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("212305", psFecha, "2", lnTipoCambioFC) / lnTipoCambioFC, 2)
                End If '***NAGL ERS 079-2016 20170407***'
            End If

        End If
        Set prs = Nothing
End Sub

Private Sub cargarDepositosInmovilizadosANX16(ByVal pobj_Excel As Excel.Application, ByVal cMoneda As String, psFecha As Date, ByVal lnTipoCambioFC As Currency)
    Dim pcelda As Excel.Range
    Dim oCtaIf As New NCajaCtaIF
    Dim oDbalanceCont As New DbalanceCont
    Dim prs As New ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Dim Cant As Integer
    
    Cant = 0
        Set prs = oCtaIf.GetDepositosInmovilizadosSBSanx16A(Format(psFecha, "yyyymmdd"), cMoneda, lnTipoCambioFC)
        nFilas = 2
        If Not prs.EOF Or prs.BOF Then
            If cMoneda = 1 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("DepInmov210701!A" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("DepInmov210701!B" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(1), 0)
                    nFilas = nFilas + 1
                    Cant = Cant + 1
                    prs.MoveNext
                Loop
                
                If (Cant = 272) Then
                Set pcelda = pobj_Excel.Range("DepInmov210701!A274")
                pobj_Excel.Range("DepInmov210701!A274").value = psFecha
                Set pcelda = pobj_Excel.Range("DepInmov210701!B274")
                pobj_Excel.Range("DepInmov210701!B274").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("211701", psFecha, "1", lnTipoCambioFC)
                End If
                 
            End If
            If cMoneda = 2 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("DepInmov210701!I" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("DepInmov210701!J" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(1), 0)
                    nFilas = nFilas + 1
                    Cant = Cant + 1
                    prs.MoveNext
                Loop
                
                If (Cant = 272) Then
                Set pcelda = pobj_Excel.Range("DepInmov210701!I274")
                pobj_Excel.Range("DepInmov210701!I274").value = psFecha
                Set pcelda = pobj_Excel.Range("DepInmov210701!J274")
                pobj_Excel.Range("DepInmov210701!J274").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("212701", psFecha, "2", lnTipoCambioFC) / lnTipoCambioFC, 2)
                End If
            End If

        End If
        Set prs = Nothing
End Sub '************NAGL ERS 079-2016 20170407*********'

Private Sub cargarObligacionesVistaANX16(ByVal pobj_Excel As Excel.Application, ByVal cMoneda As String, psFecha As Date, ByVal lnTipoCambioFC As Currency)
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim oDbalanceCont As New DbalanceCont
    Dim prs As ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Dim Cant As Integer
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
    Cant = 0
        Set prs = oCtaIf.GetObligacionesVistaSBSanx16A(Format(psFecha, "yyyymmdd"), cMoneda, lnTipoCambioFC)
        nFilas = 2
        If Not prs.EOF Or prs.BOF Then
            If cMoneda = 1 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("ObligVista2101!A" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("ObligVista2101!B" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(1), 0)
                    nFilas = nFilas + 1
                    Cant = Cant + 1
                    prs.MoveNext
                Loop
                
                If (Cant = 272) Then
                Set pcelda = pobj_Excel.Range("ObligVista2101!A274")
                pobj_Excel.Range("ObligVista2101!A274").value = psFecha
                Set pcelda = pobj_Excel.Range("ObligVista2101!B274")
                pobj_Excel.Range("ObligVista2101!B274").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2111", psFecha, "1", lnTipoCambioFC)
                End If
                 
            End If
            If cMoneda = 2 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("ObligVista2101!I" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("ObligVista2101!J" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(1), 0)
                    nFilas = nFilas + 1
                    Cant = Cant + 1
                    prs.MoveNext
                Loop
                
                If (Cant = 272) Then
                Set pcelda = pobj_Excel.Range("ObligVista2101!I274")
                pobj_Excel.Range("ObligVista2101!I274").value = psFecha
                Set pcelda = pobj_Excel.Range("ObligVista2101!J274")
                pobj_Excel.Range("ObligVista2101!J274").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2121", psFecha, "2", lnTipoCambioFC) / lnTipoCambioFC, 2)
                End If
            End If

        End If
        Set prs = Nothing
End Sub '************NAGL ERS 079-2016 20170407********'

Private Sub cargarFondeoAhPFCTSxProducto(ByVal pobj_Excel As Excel.Application, pdFecha As Date, nTipoCambio As Currency)
Dim pcelda As Excel.Range
Dim oDAnx As New DAnexoRiesgos
Dim rs As New ADODB.Recordset
Dim nColumnas As Integer

Set rs = oDAnx.ObtieneFondeoAnx16ANew(pdFecha, nTipoCambio)
nColumnas = 66
If Not (rs.BOF And rs.EOF) Then
    Do While Not rs.EOF
        If rs!cTipoProd = "233" Then
            If rs!cTipo = "FE" Then
                Set pcelda = pobj_Excel.Range("Anx16AMN!" & Chr(nColumnas + CInt(rs!cRango)) & 57)
                pobj_Excel.Range("Anx16AMN!" & Chr(nColumnas + CInt(rs!cRango)) & 57).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("Anx16AME!" & Chr(nColumnas + CInt(rs!cRango)) & 57)
                pobj_Excel.Range("Anx16AME!" & Chr(nColumnas + CInt(rs!cRango)) & 57).value = Format(rs!nSaldCntME, "#,##0.00")
                
            ElseIf rs!cTipo = "FME" Then
                Set pcelda = pobj_Excel.Range("Anx16AMN!" & Chr(nColumnas + CInt(rs!cRango)) & 58)
                pobj_Excel.Range("Anx16AMN!" & Chr(nColumnas + CInt(rs!cRango)) & 58).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("Anx16AME!" & Chr(nColumnas + CInt(rs!cRango)) & 58)
                pobj_Excel.Range("Anx16AME!" & Chr(nColumnas + CInt(rs!cRango)) & 58).value = Format(rs!nSaldCntME, "#,##0.00")
            
            ElseIf rs!cTipo = "FGA" Then
                Set pcelda = pobj_Excel.Range("Anx16AMN!" & Chr(nColumnas + CInt(rs!cRango)) & 59)
                pobj_Excel.Range("Anx16AMN!" & Chr(nColumnas + CInt(rs!cRango)) & 59).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("Anx16AME!" & Chr(nColumnas + CInt(rs!cRango)) & 59)
                pobj_Excel.Range("Anx16AME!" & Chr(nColumnas + CInt(rs!cRango)) & 59).value = Format(rs!nSaldCntME, "#,##0.00")
            End If
        ElseIf rs!cTipoProd = "232" Then
            If rs!cTipo = "FE" Then
                Set pcelda = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 36)
                pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 36).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 36)
                pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 36).value = Format(rs!nSaldCntME, "#,##0.00")
            
            ElseIf rs!cTipo = "FME" Then
               Set pcelda = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 37)
                pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 37).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 37)
                pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 37).value = Format(rs!nSaldCntME, "#,##0.00")
            
            ElseIf rs!cTipo = "FGA" Then
                Set pcelda = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 38)
                pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 38).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 38)
                pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 38).value = Format(rs!nSaldCntME, "#,##0.00")
            End If
        ElseIf rs!cTipoProd = "234" Then
            If rs!cTipo = "FE" Then
                Set pcelda = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 47)
                pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 47).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 47)
                pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 47).value = Format(rs!nSaldCntME, "#,##0.00")
            
            ElseIf rs!cTipo = "FME" Then
               Set pcelda = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 48)
                pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 48).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 48)
                pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 48).value = Format(rs!nSaldCntME, "#,##0.00")
            
            ElseIf rs!cTipo = "FGA" Then
                Set pcelda = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 49)
                pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 49).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 49)
                pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 49).value = Format(rs!nSaldCntME, "#,##0.00")
            End If
        End If
        rs.MoveNext
    Loop
End If
End Sub 'NAGL ERS006-2019 20190514

Private Sub PintaFondeoObligacionesVista(ByRef xlHoja1 As Excel.Worksheet, ByVal pdFecha As Date, ByVal pnTipoCambio As Currency, ByVal pnMoneda As Integer)
    Dim loRs As ADODB.Recordset
    Dim oDbalanceCont As DbalanceCont
    Set oDbalanceCont = New DbalanceCont
    Set loRs = New ADODB.Recordset
    Set loRs = oDbalanceCont.ObtenerFondeoObligVista(pdFecha, pnTipoCambio) 'Antes ObtenerFondeoGirosxPagar 20190518
    If Not (loRs.BOF Or loRs.EOF) Then
    Do While Not loRs.EOF
        If loRs!nTipoCobertura = 1 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                    xlHoja1.Cells(63, 3) = CCur(xlHoja1.Cells(63, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
            If loRs!nPersoneria = "3" Then
                    xlHoja1.Cells(63, 3) = CCur(xlHoja1.Cells(63, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                    xlHoja1.Cells(65, 3) = CCur(xlHoja1.Cells(65, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If

        End If
        If loRs!nTipoCobertura = 0 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                    xlHoja1.Cells(64, 3) = CCur(xlHoja1.Cells(64, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
            If loRs!nPersoneria = "3" Then
                    xlHoja1.Cells(64, 3) = CCur(xlHoja1.Cells(64, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                    xlHoja1.Cells(65, 3) = CCur(xlHoja1.Cells(65, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
        End If
        loRs.MoveNext
    Loop
    End If
End Sub

Private Sub cargarPlazoFijoRangosPersoneriaRangoAnexo6(ByVal pobj_Excel As Excel.Application, psFecha As Date, psFiltro As String, nPromedioEncajeMAMN As Currency, nPromedioEncajeMAME As Currency, nTipoCambio As Currency)
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim objDAnexoRiesgos As DAnexoRiesgos
    Dim prs As ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Dim nColumnas As Integer
    Dim oDbalanceCont As DbalanceCont
    Dim nRango As Integer 'ALPA20140211
    Set objDAnexoRiesgos = New DAnexoRiesgos
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
    Set oDbalanceCont = New DbalanceCont
    Dim ix As Integer
    Dim lnTipoCambioFCMA As Currency
    Dim lnToTalCajaFondosMN, lnToTalCajaFondosME As Currency
    Dim pdFecha As Date
    Dim oEst As New NEstadisticas 'NAGL
    Dim SaldoCajaME As Currency  'NAGL
    Dim oCambio As New nTipoCambio 'NAGL
    'Dim ldFechaAnt As String 'NAGL
    Dim lnTotalBCRPMN As Currency 'NAGL
    Dim lnTotalBCRPME As Currency 'NAGL

     'Inicio
    Dim pdFechaFinDeMes As Date
    Dim pdFechaFinDeMesMA As Date
    Dim nSaldoCajaDiarioMesAnteriorME As Currency
    Dim nSaldoCajaDiarioMesAnteriorMN As Currency
    Dim lnToTalOMN As Currency
    Dim lnToTalOME As Currency
    Dim ldFechaPro As Date
    Dim lnToTalTotalCajaFondosMN As Currency
    Dim lnToTalTotalCajaFondosME As Currency
    Dim lnTotalSaldoBCRPAnexoDiarioMN As Currency
    Dim lnTotalSaldoBCRPAnexoDiarioME As Currency
    
    nSaldoCajaDiarioMesAnteriorMN = 0
    pdFechaFinDeMes = DateAdd("d", -Day(DateAdd("m", 1, psFecha)), DateAdd("m", 1, psFecha))
    pdFechaFinDeMesMA = DateAdd("d", -Day(pdFechaFinDeMes), pdFechaFinDeMes)

    ldFechaPro = DateAdd("d", -Day(psFecha), psFecha)
    ldFechaPro = DateAdd("d", -Day(ldFechaPro), ldFechaPro)
    nSaldoCajaDiarioMesAnteriorME = 0
    
     lnTipoCambioFCMA = 0
    For ix = 1 To Day(pdFechaFinDeMesMA)
        ldFechaPro = DateAdd("d", 1, ldFechaPro)
        nSaldoCajaDiarioMesAnteriorMN = nSaldoCajaDiarioMesAnteriorMN + (oDbalanceCont.SaldoCtas(32, "761201", ldFechaPro, pdFechaFinDeMesMA, lnTipoCambioFCMA, lnTipoCambioFCMA) / Day(pdFechaFinDeMesMA))
    Next ix
    
    lnToTalCajaFondosMN = 0
    lnToTalCajaFondosME = 0
    lnTotalSaldoBCRPAnexoDiarioMN = 0
    lnTotalSaldoBCRPAnexoDiarioME = 0
    lnToTalTotalCajaFondosMN = 0
    lnToTalTotalCajaFondosME = 0
    lnTotalBCRPMN = 0 '***NAGL
    lnTotalBCRPME = 0 '****NAGL
    lnToTalOMN = 0
    lnToTalOME = 0
    
    ldFechaPro = DateAdd("d", -Day(psFecha), psFecha) '***NAGL
    'ldFechaAnt = ldFechaPro
    
   'CAJA  - DEPOSITOS EN EL BCRP
    For ix = 1 To Day(psFecha)
        ldFechaPro = DateAdd("d", 1, ldFechaPro)
        If Month(ldFechaPro) = Month(DateAdd("d", 1, ldFechaPro)) Then
               lnTipoCambioFCMA = Format(oCambio.EmiteTipoCambio(ldFechaPro, TCFijoDia), "#,##0.0000")
        Else
               lnTipoCambioFCMA = Format(oCambio.EmiteTipoCambio(DateAdd("d", 1, ldFechaPro), TCFijoDia), "#,##0.0000")
        End If
        
        'CAJA ANTERIOR MN
        'SaldoCajAnt = Round(oEst.GetCajaAnterior(ldFechaAnt, "761201", "32"), 2)
        lnToTalTotalCajaFondosMN = lnToTalTotalCajaFondosMN + Round(nSaldoCajaDiarioMesAnteriorMN, 2)
        
        'CAJA ME
        'SaldoCajaME = oDbalanceCont.SaldoCajasObligExoneradas(Format(ldFechaPro, "yyyymmdd"), 2)
        SaldoCajaME = oDbalanceCont.ObtenerCtaContSaldoDiario("1121", ldFechaPro) + oDbalanceCont.ObtenerCtaContSaldoDiario("112701", ldFechaPro)
        lnToTalTotalCajaFondosME = lnToTalTotalCajaFondosME + SaldoCajaME
       
        'DEPOSITOS BCRP
        lnTotalSaldoBCRPAnexoDiarioMN = oDbalanceCont.SaldoBCRPAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1)
        lnTotalBCRPMN = lnTotalBCRPMN + lnTotalSaldoBCRPAnexoDiarioMN
        
        lnTotalSaldoBCRPAnexoDiarioME = oDbalanceCont.SaldoBCRPAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2)
        lnTotalBCRPME = lnTotalBCRPME + lnTotalSaldoBCRPAnexoDiarioME
        
    Next ix
    
    lnToTalOMN = lnToTalTotalCajaFondosMN + lnTotalBCRPMN 'lnTotalSaldoBCRPAnexoDiarioMN
    lnToTalOME = lnToTalTotalCajaFondosME + lnTotalBCRPME 'lnTotalSaldoBCRPAnexoDiarioME
    'Fin

    'SOLES*********************************************************************
    Set pcelda = pobj_Excel.Range("Disponible!B17")
    pobj_Excel.Range("Disponible!B17").value = Round(nPromedioEncajeMAMN / Day(psFecha), 2) '***NAGL ERS 079-2016 20170407
    
    Set pcelda = pobj_Excel.Range("Disponible!B16") 'Superavit
    pobj_Excel.Range("Disponible!B16").value = Round((lnToTalOMN - nPromedioEncajeMAMN) / Day(psFecha), 2)
    pobj_Excel.Range("Disponible!B16:Disponible!B17").NumberFormat = "#,##0.00;-#,##0.00"
    
    Set pcelda = pobj_Excel.Range("Disponible!B20")
    'pobj_Excel.Range("Disponible!B20").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1117", psFecha, "1", 0) 'JIPR20211115 DEBE TOMAR SIN VENCIMIENTO
    pobj_Excel.Range("Disponible!B20").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("111701", psFecha, "1", 0)	
    Set pcelda = pobj_Excel.Range("Disponible!B21")
    pobj_Excel.Range("Disponible!B21").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioxLike("1113010_0[12]", psFecha, "1", 0)
    Set pcelda = pobj_Excel.Range("Disponible!B22")
    pobj_Excel.Range("Disponible!B22").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioxLike("111201", psFecha, "1", 0)
    
    Set pcelda = pobj_Excel.Range("Disponible!B23")
    pobj_Excel.Range("Disponible!B23").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1111", psFecha, "1", 0)
    
    Set pcelda = pobj_Excel.Range("Disponible!B24")
    pobj_Excel.Range("Disponible!B24").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1115", psFecha, "1", 0) 'NAGL 20170621
    
    Set pcelda = pobj_Excel.Range("Disponible!B25")
    pobj_Excel.Range("Disponible!B25").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1116", psFecha, "1", 0)
    
    Set pcelda = pobj_Excel.Range("Disponible!B26")
    pobj_Excel.Range("Disponible!B26").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1118", psFecha, "1", 0)
    
    pobj_Excel.Range("Disponible!B20:Disponible!B26").NumberFormat = "#,##0.00;-#,##0.00"
    'FIN SOLES
    
    'DOLARES*********************************************************************
    
    Set pcelda = pobj_Excel.Range("DisponibleDolares!B17") 'Encaje Exigible ME
    pobj_Excel.Range("DisponibleDolares!B17").value = Round(nPromedioEncajeMAME / Day(psFecha), 2) '***NAGL ERS 079-2016 20170407
    
    Set pcelda = pobj_Excel.Range("DisponibleDolares!B16") 'Superavit
    pobj_Excel.Range("DisponibleDolares!B16").value = Round((lnToTalOME - nPromedioEncajeMAME) / Day(psFecha), 2) 'NAGL ERS 079-2016 20170407
    pobj_Excel.Range("DisponibleDolares!B16:DisponibleDolares!B17").NumberFormat = "#,##0.00;-#,##0.00"
    
    Set pcelda = pobj_Excel.Range("DisponibleDolares!B20")
    'pobj_Excel.Range("DisponibleDolares!B20").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1127", psFecha, "2", nTipoCambio) / nTipoCambio, 2) 'JIPR20211115 DEBE TOMAR SIN VENCIMIENTO
    pobj_Excel.Range("DisponibleDolares!B20").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("112701", psFecha, "2", nTipoCambio) / nTipoCambio, 2)	
	Set pcelda = pobj_Excel.Range("DisponibleDolares!B21")
    pobj_Excel.Range("DisponibleDolares!B21").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioxLike("1123010_0[12]", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
    Set pcelda = pobj_Excel.Range("DisponibleDolares!B22")
    pobj_Excel.Range("DisponibleDolares!B22").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioxLike("112201", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
    
    Set pcelda = pobj_Excel.Range("DisponibleDolares!B23")
    pobj_Excel.Range("DisponibleDolares!B23").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1121", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
    
    Set pcelda = pobj_Excel.Range("DisponibleDolares!B24")
    pobj_Excel.Range("DisponibleDolares!B24").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1125", psFecha, "2", nTipoCambio) / nTipoCambio, 2) 'NAGL 20170621
    
    Set pcelda = pobj_Excel.Range("DisponibleDolares!B25")
    pobj_Excel.Range("DisponibleDolares!B25").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1126", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
    
    Set pcelda = pobj_Excel.Range("DisponibleDolares!B26")
    pobj_Excel.Range("DisponibleDolares!B26").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1128", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
    
    pobj_Excel.Range("DisponibleDolares!B20:DisponibleDolares!B26").NumberFormat = "#,##0.00;-#,##0.00"
    'FIN DOLARES
    
    Set prs = oCtaIf.GetPlazoFijoRangosPersoneriaRangoAnexo6(Format(psFecha, "yyyy/mm/dd"), psFiltro)
    nColumnas = 66
    If Not prs.EOF Or prs.BOF Then
        Do While Not prs.EOF
            If prs!cTipo = 1 Then
                Set pcelda = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(prs!cRango)) & 3)
                pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(prs!cRango)) & 3).value = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(prs!cRango)) & 3).value + prs!SALDO_MN
                
                Set pcelda = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(prs!cRango)) & 3)
                pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(prs!cRango)) & 3).value = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(prs!cRango)) & 3).value + prs!SALDO_ME
            End If
            If prs!cTipo = 2 Then
                Set pcelda = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(prs!cRango)) & 3)
                pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(prs!cRango)) & 3).value = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(prs!cRango)) & 3).value - prs!SALDO_MN
                pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(prs!cRango)) & 4).value = prs!SALDO_MN
                
                Set pcelda = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(prs!cRango)) & 3)
                pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(prs!cRango)) & 3).value = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(prs!cRango)) & 3).value - prs!SALDO_ME
                pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(prs!cRango)) & 4).value = prs!SALDO_ME
            End If
            prs.MoveNext
        Loop
    End If
    Set prs = Nothing
    
    '****NAGL 20190515 ERS006-2019 NUEVO FORMATO
    
    'Set pcelda = pobj_Excel.Range("Creditos!C51")
    'pobj_Excel.Range("Creditos!C51").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251703", psFecha, "1", 0)
    'Set pcelda = pobj_Excel.Range("Creditos!C52")
    'pobj_Excel.Range("Creditos!C52").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251704", psFecha, "1", 0)
    'Set pcelda = pobj_Excel.Range("Creditos!C53")
    'pobj_Excel.Range("Creditos!C53").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251705", psFecha, "1", 0)
    'Set pcelda = pobj_Excel.Range("Creditos!C55")
    'pobj_Excel.Range("Creditos!C55").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2312", psFecha, "1", 0)
    'Set pcelda = pobj_Excel.Range("Creditos!C56")
    'pobj_Excel.Range("Creditos!C56").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2313", psFecha, "1", 0)
    'Comentado by NAGL 202012*************************************
    
    'Set pcelda = pobj_Excel.Range("Creditos!E51")
    'pobj_Excel.Range("Creditos!E51").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252703", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
    'Set pcelda = pobj_Excel.Range("Creditos!E52")
    'pobj_Excel.Range("Creditos!E52").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252704", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
    'Set pcelda = pobj_Excel.Range("Creditos!E53")
    'pobj_Excel.Range("Creditos!E53").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252705", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
    'Set pcelda = pobj_Excel.Range("Creditos!E55")
    'pobj_Excel.Range("Creditos!E55").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2322", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
    'Set pcelda = pobj_Excel.Range("Creditos!E56")
    'pobj_Excel.Range("Creditos!E56").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2323", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
    'pobj_Excel.Range("Creditos!C51:Creditos!C56").NumberFormat = "#,##0.00;-#,##0.00"
    'pobj_Excel.Range("Creditos!E51:Creditos!E56").NumberFormat = "#,##0.00;-#,##0.00"
    'Comentado by NAGL 202012*************************************
    
    Set pcelda = pobj_Excel.Range("CreditosVIg_SinReactiva!C54")
    pobj_Excel.Range("CreditosVIg_SinReactiva!C54").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2116", psFecha, "1", 0)
    Set pcelda = pobj_Excel.Range("CreditosVIg_SinReactiva!E54")
    pobj_Excel.Range("CreditosVIg_SinReactiva!E54").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2126", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
    
    Set pcelda = pobj_Excel.Range("CreditosVIg_SinReactiva!A54")
    If (Month(psFecha) >= 1 And Month(psFecha) <= 4) Then
        nRango = (5 - CInt(Month(psFecha)))
    ElseIf Month(psFecha) = 11 Then
        nRango = 6
    ElseIf Month(psFecha) = 12 Then
        nRango = 5
    ElseIf (Month(psFecha) >= 5 And Month(psFecha) <= 10) Then
        nRango = (11 - CInt(Month(psFecha)))
    End If 'NAGL 20170519
    '1M  2M  3M  4M  5M  6M  7-9 M   10-12 M

    pobj_Excel.Range("CreditosVIg_SinReactiva!A54").value = nRango
    pdFechaFinDeMesMA = DateAdd("d", -Day(psFecha), psFecha)
    Set pcelda = pobj_Excel.Range("CreditosVIg_SinReactiva!I36")
    pobj_Excel.Range("CreditosVIg_SinReactiva!I36").value = pdFechaFinDeMesMA

    Set pcelda = pobj_Excel.Range("CreditosVIg_SinReactiva!I38")
    pobj_Excel.Range("CreditosVIg_SinReactiva!I38").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1404", pdFechaFinDeMesMA, "0", 0)
    Set pcelda = pobj_Excel.Range("CreditosVIg_SinReactiva!I39")
    pobj_Excel.Range("CreditosVIg_SinReactiva!I39").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1405", pdFechaFinDeMesMA, "0", 0)
    Set pcelda = pobj_Excel.Range("CreditosVIg_SinReactiva!I40")
    pobj_Excel.Range("CreditosVIg_SinReactiva!I40").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1406", pdFechaFinDeMesMA, "0", 0)
    
    'NAGL 20190516********
    Set pcelda = pobj_Excel.Range("CreditosVIg_SinReactiva!I42")
    pobj_Excel.Range("CreditosVIg_SinReactiva!I42").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("14", pdFechaFinDeMesMA, "0", 0)
    Set pcelda = pobj_Excel.Range("CreditosVIg_SinReactiva!I43")
    pobj_Excel.Range("CreditosVIg_SinReactiva!I43").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1408", pdFechaFinDeMesMA, "0", 0)
    Set pcelda = pobj_Excel.Range("CreditosVIg_SinReactiva!I44")
    pobj_Excel.Range("CreditosVIg_SinReactiva!I44").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1409", pdFechaFinDeMesMA, "0", 0)
    pobj_Excel.Range("CreditosVIg_SinReactiva!I38:CreditosVIg_SinReactiva!I44").NumberFormat = "#,##0.00;-#,##0.00"
    '***********************
End Sub

Private Sub cargarDatosBalanceANX16(ByVal pobj_Excel As Excel.Application, ByVal cMoneda As String, psFecha As Date, ByVal nTipoCambio As Currency)
    Dim pcelda As Excel.Range
    Dim oDbalanceCont As New DbalanceCont
    Dim prs As New ADODB.Recordset
    Dim nFilas As Integer
    Dim cColum As String
    Dim cCtaCnt As String
    Dim psNombHoja As String
    
    Set prs = oDbalanceCont.ObtieneCuentasBalanceANX16(cMoneda)
    If cMoneda = "1" Then
       psNombHoja = "Anx16AMN!"
       nTipoCambio = 1
    Else
       psNombHoja = "Anx16AME!"
    End If
    
    If Not (prs.EOF And prs.BOF) Then
      Do While Not prs.EOF
          cColum = prs!cColum
          nFilas = prs!nFila
          cCtaCnt = prs!cCuenta
          Set pcelda = pobj_Excel.Range(psNombHoja & cColum & nFilas)
          
          If oDbalanceCont.EvaluaCuentaLikeAnx16(cCtaCnt) = True Then
            pobj_Excel.Range(psNombHoja & cColum & nFilas).value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioxLike(cCtaCnt, psFecha, cMoneda, 0) / nTipoCambio, 2)
          Else
            pobj_Excel.Range(psNombHoja & cColum & nFilas).value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes(cCtaCnt, psFecha, cMoneda, 0) / nTipoCambio, 2)
          End If
        prs.MoveNext
      Loop
    End If
End Sub 'NAGL 202012 Según ACTA N°094-2020

