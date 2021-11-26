VERSION 5.00
Begin VB.Form frmAnx15AConsolBancos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte Consolidado de Bancos y Otras Entidades Financieras"
   ClientHeight    =   975
   ClientLeft      =   1620
   ClientTop       =   2235
   ClientWidth     =   4815
   Icon            =   "frmAnx15AConsolBancos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmAnx15AConsolBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lsArchivo As String
Dim lsNomArch As String
Dim lbExcel As Boolean

Dim lsOpeCuentas() As String

Dim nFil01 As Integer, nFil02 As Integer, nFil03 As Integer
Dim aBcoCon() As String
Dim aCmacs()  As String

Dim lsSaldosAhorrosCmacs() As String
Dim lsSaldosPFCmacs()  As String
Dim lsObjetosAnexo15() As String
Dim lnTipCambio        As Currency
Dim nFilBcoRestringido As Integer

Public Sub ImprimeConsolidaBancos(psOpeCod As String, psMoneda As String, pdFecha As Date)
On Error GoTo ImprimeConsolidaBancosErr
Dim oTC As New nTipoCambio
Dim oBarra As clsProgressBar
Set oBarra = New clsProgressBar
   lnTipCambio = oTC.EmiteTipoCambio(pdFecha, TCFijoMes)
   oBarra.ShowForm frmReportes
   oBarra.Max = 3
   oBarra.CaptionSyle = eCap_CaptionPercent
   oBarra.Progress 0, "CONSOLIDADO DE BANCOS", "", "Generando Reporte", vbBlue
   
    Set oTC = Nothing
   lsNomArch = "Anx15A_ConsolBancos_" & Format(pdFecha, "mmyyyy") & IIf(psMoneda = "1", "MN", "ME") & ".XLS"
   lsArchivo = App.path & "\SPOOLER\" & lsNomArch
   lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
   If lbExcel Then
      ExcelAddHoja Format(pdFecha, "dd-mm-yyyy"), xlLibro, xlHoja1
      oBarra.Progress 1, "CONSOLIDADO DE BANCOS", "", "Generando Reporte Bancos", vbBlue
      ConsolidaBancos psOpeCod, psMoneda, pdFecha
      oBarra.Progress 2, "CONSOLIDADO DE BANCOS", "", "Generando Reporte Cmacs", vbBlue
      ConsolidaCmacs psOpeCod, psMoneda, pdFecha
      oBarra.Progress 3, "CONSOLIDADO DE BANCOS", "", "Preparando Archivo Excel...", vbBlue
      GeneraReporte psOpeCod, psMoneda, pdFecha
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
      oBarra.Progress 4, "CONSOLIDADO DE BANCOS", "", "Reporte Generado", vbBlue
      If lsArchivo <> "" Then
         CargaArchivo lsArchivo, App.path & "\SPOOLER\"
      End If
   End If
   oBarra.CloseForm frmReportes
   Set oBarra = Nothing
Exit Sub
ImprimeConsolidaBancosErr:
  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
    If lbExcel = True Then
       ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
End Sub
Private Sub GeneraReporte(psOpeCod As String, psMoneda As String, pdFecha As Date)
    Dim lnFila As Integer, I As Integer, lnCol As Integer, j As Integer
    Dim Y1 As Currency, Y2 As Currency, N As Integer
    
    Dim lsSubTotalBancos() As String
    Dim lsTotSaldos As String
    Dim lsTotalSaldos As String
    Dim TotalBancosCmac As String
    Dim lbExisteHoja  As Boolean
    Dim lsRutaReferencia  As String
    Dim nTempFil As Integer
    
    TotalBancosCmac = ""
    ReDim lsSubTotalBancos(5)
    xlHoja1.PageSetup.Zoom = 80
    xlHoja1.Range("A1:R100").Font.Size = 8
    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(4, 8)).Font.Bold = True
    
    xlHoja1.Cells(1, 1) = gsNomCmac
    xlHoja1.Cells(1, 7) = "Fecha :" & Format(pdFecha, "dd mmmm yyyy")
    xlHoja1.Cells(2, 1) = "Area de Caja General"
    xlHoja1.Cells(3, 3) = "CONSOLIDADO DE BANCOS Y OTRAS ENTIDADES FINANCIERAS "
    xlHoja1.Cells(4, 4) = IIf(psMoneda = "1", "MONEDA NACIONAL", "MONEDA EXTRANJERA")
    
    lsRutaReferencia = "+^" & App.path & "\SPOOLER\[" & lsNomArch & "]" & Format(pdFecha, "dd-mm-yyyy") & "^!"

    lnFila = 6
    Y1 = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Borders.LineStyle = xlContinuous
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Borders.LineStyle = xlContinuous
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 7)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 7)).Borders.LineStyle = xlContinuous
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
    
    xlHoja1.Cells(lnFila, 2) = "CUENTAS CORRIENTES": xlHoja1.Cells(lnFila, 4) = "CUENTAS AHORROS": xlHoja1.Cells(lnFila, 6) = "CUENTAS A PLAZO"
    lnFila = lnFila + 1
    
    'definimos el ancho de las columnas del numero de cuentas
    xlHoja1.Range("A1").ColumnWidth = 25
    xlHoja1.Range("B1").ColumnWidth = 20
    xlHoja1.Range("D1").ColumnWidth = 20
    xlHoja1.Range("F1").ColumnWidth = 20
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
    
    xlHoja1.Cells(lnFila, 1) = "BANCOS":    xlHoja1.Cells(lnFila, 2) = "Nº CUENTA": xlHoja1.Cells(lnFila, 3) = "SALDOS " & IIf(psMoneda = "1", gcPEN_SIMBOLO, "$") 'marg ers044-2016
                                            xlHoja1.Cells(lnFila, 4) = "Nº CUENTA": xlHoja1.Cells(lnFila, 5) = "SALDOS " & IIf(psMoneda = "1", gcPEN_SIMBOLO, "$") 'marg ers044-2016
                                            xlHoja1.Cells(lnFila, 6) = "Nº CUENTA": xlHoja1.Cells(lnFila, 7) = "SALDOS " & IIf(psMoneda = "1", gcPEN_SIMBOLO, "$") 'marg ers044-2016
                                            xlHoja1.Cells(lnFila, 8) = "TOTAL": xlHoja1.Cells(lnFila, 9) = "ADEUDADOS"
    
    lnFila = lnFila + 1
    Y2 = lnFila
    ExcelCuadro xlHoja1, 1, CCur(Y1), 9, CCur(Y2)
    
   Dim oEst As New NEstadisticas
    For N = 1 To UBound(aBcoCon, 2)
        lsTotSaldos = ""
        If N = nFilBcoRestringido And N > 1 Then
            Y2 = lnFila
            ExcelCuadro xlHoja1, 1, CCur(Y1), 9, CCur(Y2)
            lnFila = lnFila + 1
            xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).Borders.LineStyle = xlContinuous
            xlHoja1.Cells(lnFila, 1) = "SUBTOTAL BANCOS"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).Formula = "=sum(" & Mid(lsSubTotalBancos(1), 1, Len(lsSubTotalBancos(1)) - 1) & ")"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Formula = "=sum(" & Mid(lsSubTotalBancos(2), 1, Len(lsSubTotalBancos(2)) - 1) & ")"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=sum(" & Mid(lsSubTotalBancos(3), 1, Len(lsSubTotalBancos(3)) - 1) & ")"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).Formula = "=sum(" & Mid(lsSubTotalBancos(4), 1, Len(lsSubTotalBancos(4)) - 1) & ")"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila, 9)).Formula = "=sum(" & Mid(lsSubTotalBancos(5), 1, Len(lsSubTotalBancos(5)) - 1) & ")"
            
            ReDim lsSubTotalBancos(5)
            TotalBancosCmac = TotalBancosCmac & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).Address(False, False)
            
            lnFila = lnFila + 1
            xlHoja1.Cells(lnFila, 1) = "BANCOS RESTRINGIDOS"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).Borders.LineStyle = xlContinuous
            
            Y1 = lnFila + 1
        End If
        
        lnFila = lnFila + 1
        If N = 1 Then
            Y1 = lnFila
        End If
        xlHoja1.Cells(lnFila, 1) = aBcoCon(1, N) 'nombre de banco
        xlHoja1.Cells(lnFila, 2) = "'" + aBcoCon(2, N) ' nº de cuenta corriente
        xlHoja1.Cells(lnFila, 3) = aBcoCon(3, N) ' Saldo de cuenta corriente
        'Ubicación de Cuentas de BCR
        If Trim(aBcoCon(2, N)) = Trim(lsObjetosAnexo15(0)) Then
            'oEst.EliminaEstadAnexos pdFecha, "FONDOSBCR", psMoneda
            'oEst.InsertaEstadAnexos pdFecha, "FONDOSBCR", psMoneda, lsRutaReferencia & xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).Address
        End If
    
        'Ubicacion de Cuentas del banco de la Nacion
        If Trim(aBcoCon(2, N)) = Trim(lsObjetosAnexo15(1)) Then
            oEst.EliminaEstadAnexos pdFecha, "BCONACION", psMoneda
            oEst.InsertaEstadAnexos pdFecha, "BCONACION", psMoneda, lsRutaReferencia & xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).Address
        End If
                
        lsSubTotalBancos(1) = lsSubTotalBancos(1) + xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).Address(False, False) & "+"
        lsTotSaldos = lsTotSaldos + xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).Address(False, False) & "+"

        xlHoja1.Cells(lnFila, 4) = aBcoCon(4, N) 'nº de cuenta de ahorros
        xlHoja1.Cells(lnFila, 5) = aBcoCon(5, N) 'saldo de cta de ahorros

        lsSubTotalBancos(2) = lsSubTotalBancos(2) + xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Address(False, False) & "+"
        lsTotSaldos = lsTotSaldos + xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Address(False, False) & "+"

        xlHoja1.Cells(lnFila, 6) = aBcoCon(6, N) 'Nº de cta plazo fijo
        xlHoja1.Cells(lnFila, 7) = aBcoCon(7, N) 'Saldo de cta de plazo fijo

        lsSubTotalBancos(3) = lsSubTotalBancos(3) + xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False) & "+"
        lsTotSaldos = lsTotSaldos + xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)

        xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).Formula = "=sum(" & Mid(lsTotSaldos, 1, Len(lsTotSaldos))
        lsSubTotalBancos(4) = lsSubTotalBancos(4) + xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).Address(False, False) & "+"
        xlHoja1.Cells(lnFila, 9) = aBcoCon(8, N) 'Adeudados
        lsSubTotalBancos(5) = lsSubTotalBancos(5) + xlHoja1.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila, 9)).Address(False, False) & "+"
        
    Next

   Y2 = lnFila
   ExcelCuadro xlHoja1, 1, Y1, 9, Y2

   lnFila = lnFila + 1
   xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
   xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).Borders.LineStyle = xlContinuous
   
    If nFilBcoRestringido = 0 Then
        xlHoja1.Cells(lnFila, 1) = "SUBTOTAL BANCOS"
        nTempFil = 0
    Else
        xlHoja1.Cells(lnFila, 1) = "SUBTOTAL BCOS. RESTRINGIDOS"
        nTempFil = lnFila
    End If
    
        
    If Len(lsSubTotalBancos(1)) - 1 >= 0 Then
        xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).Formula = "=sum(" & Mid(lsSubTotalBancos(1), 1, Len(lsSubTotalBancos(1)) - 1) & ")"
        xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Formula = "=sum(" & Mid(lsSubTotalBancos(2), 1, Len(lsSubTotalBancos(2)) - 1) & ")"
        xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=sum(" & Mid(lsSubTotalBancos(3), 1, Len(lsSubTotalBancos(3)) - 1) & ")"
        xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).Formula = "=sum(" & Mid(lsSubTotalBancos(4), 1, Len(lsSubTotalBancos(4)) - 1) & ")"
        xlHoja1.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila, 9)).Formula = "=sum(" & Mid(lsSubTotalBancos(5), 1, Len(lsSubTotalBancos(5)) - 1) & ")"
    End If
   
   oEst.EliminaEstadAnexos pdFecha, "RESTRBANCO", psMoneda
   
   If nTempFil > 0 Then
        'oEst.EliminaEstadAnexos pdFecha, "RESTRBANCO", psMoneda
        oEst.InsertaEstadAnexos pdFecha, "RESTRBANCO", psMoneda, lsRutaReferencia & xlHoja1.Range(xlHoja1.Cells(nTempFil, 8), xlHoja1.Cells(nTempFil, 8)).Address
   End If
   
   oEst.EliminaEstadAnexos pdFecha, "ADEUDADOBANCO", psMoneda
   oEst.InsertaEstadAnexos pdFecha, "ADEUDADOBANCO", psMoneda, lsRutaReferencia & xlHoja1.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila, 9)).Address
    
   TotalBancosCmac = TotalBancosCmac & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).Address(False, False)
   lnFila = lnFila + 1

   xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
   xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Borders.LineStyle = xlContinuous

   xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Merge True
   xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Borders.LineStyle = xlContinuous

   xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).HorizontalAlignment = xlHAlignCenter
   xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
   Y1 = lnFila
   xlHoja1.Cells(lnFila, 2) = "CUENTAS AHORROS": xlHoja1.Cells(lnFila, 4) = "CUENTAS A PLAZO"
   lnFila = lnFila + 1
   xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).HorizontalAlignment = xlHAlignCenter
   xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
   xlHoja1.Cells(lnFila, 1) = "CMACS": xlHoja1.Cells(lnFila, 2) = "Nº CUENTA": xlHoja1.Cells(lnFila, 3) = "SALDOS " & IIf(psMoneda = "1", gcPEN_SIMBOLO, "$") 'marg ers044-2016
                                        xlHoja1.Cells(lnFila, 4) = "Nº CUENTA": xlHoja1.Cells(lnFila, 5) = "SALDOS " & IIf(psMoneda = "1", gcPEN_SIMBOLO, "$") 'marg ers044-2016
                                        xlHoja1.Cells(lnFila, 6) = "INTERBANCARIOS": xlHoja1.Cells(lnFila, 7) = "TOTAL"
                                        xlHoja1.Cells(lnFila, 8) = "ADEUDADOS"


   lnFila = lnFila + 1
   Y2 = lnFila
   ExcelCuadro xlHoja1, 1, Y1, 8, Y2
   ReDim lsSubTotalBancos(5)
   For N = 1 To UBound(aCmacs, 2)
       If aCmacs(2, N) <> "SIN CUENTA" Then
           lsTotSaldos = ""
           lnFila = lnFila + 1
           If N = 1 Then
               Y1 = lnFila
           End If
           xlHoja1.Cells(lnFila, 1) = aCmacs(1, N) 'nombre de Cmac
           xlHoja1.Cells(lnFila, 2) = "'" & aCmacs(2, N) ' nº de cuenta Ahorros
           xlHoja1.Cells(lnFila, 3) = aCmacs(3, N) ' Saldo de cuenta Ahorros
           
           lsSubTotalBancos(1) = lsSubTotalBancos(1) + xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).Address(False, False) & "+"
           lsTotSaldos = lsTotSaldos + xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).Address(False, False) & "+"
   
           xlHoja1.Cells(lnFila, 4) = "'" & aCmacs(4, N) 'nº de cuenta de Plazo Fijo
           xlHoja1.Cells(lnFila, 5) = aCmacs(5, N) 'saldo de cta de Plazo fijo
   
           lsTotSaldos = lsTotSaldos + xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Address(False, False) & "+"
           lsSubTotalBancos(2) = lsSubTotalBancos(2) + xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Address(False, False) & "+"
           
           lsTotSaldos = lsTotSaldos + xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
           lsSubTotalBancos(3) = lsSubTotalBancos(3) + xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False) & "+"
               
           xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=sum(" & Mid(lsTotSaldos, 1, Len(lsTotSaldos))
           lsSubTotalBancos(4) = lsSubTotalBancos(4) + xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False) & "+"
           
           lsSubTotalBancos(5) = lsSubTotalBancos(5) + xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).Address(False, False) & "+"
       End If
   Next

    
   Y2 = lnFila
   ExcelCuadro xlHoja1, 1, Y1, 8, Y2
   lnFila = lnFila + 1
   xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
   xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 8)).Borders.LineStyle = xlContinuous
   xlHoja1.Cells(lnFila, 1) = "SUBTOTAL CMACS"
   If lsSubTotalBancos(1) <> "" Then
      xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).Formula = "=sum(" & Mid(lsSubTotalBancos(1), 1, Len(lsSubTotalBancos(1)) - 1) & ")"
      xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Formula = "=sum(" & Mid(lsSubTotalBancos(2), 1, Len(lsSubTotalBancos(2)) - 1) & ")"
      
      xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=sum(" & Mid(lsSubTotalBancos(3), 1, Len(lsSubTotalBancos(3)) - 1) & ")"
      xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=sum(" & Mid(lsSubTotalBancos(4), 1, Len(lsSubTotalBancos(4)) - 1) & ")"
      xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).Formula = "=sum(" & Mid(lsSubTotalBancos(5), 1, Len(lsSubTotalBancos(5)) - 1) & ")"
   End If
   TotalBancosCmac = TotalBancosCmac & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)
   
   Y1 = lnFila + 1
   lnFila = lnFila + 2
   MuestraAdeudadosCortoPlazo lnFila, Val(psMoneda), pdFecha
   oEst.EliminaEstadAnexos pdFecha, "ADEUDADO", psMoneda
   oEst.InsertaEstadAnexos pdFecha, "ADEUDADO", psMoneda, lsRutaReferencia & xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).Address
    
   lnFila = lnFila + 1
   xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
   xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 8)).Borders.LineStyle = xlContinuous
   xlHoja1.Cells(lnFila, 1) = "TOTAL BANCOS Y CMACS " & IIf(psMoneda = "1", gcPEN_SIMBOLO, "$") 'marg ers044-2016
   xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=sum(" & TotalBancosCmac & ")"
      
   lnFila = lnFila + 2
   xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
   xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 8)).Borders.LineStyle = xlContinuous
   xlHoja1.Cells(lnFila, 1) = "TOTAL FONDOS MUTUOS " & IIf(psMoneda = "1", gcPEN_SIMBOLO, "$") 'marg ers044-2016
   xlHoja1.Cells(lnFila, 7) = Format(oEst.GetFondosMutuos(psMoneda), "#,#0.00")
 
      
    oEst.EliminaEstadAnexos pdFecha, "BANCOS", psMoneda
    oEst.InsertaEstadAnexos pdFecha, "BANCOS", psMoneda, lsRutaReferencia & xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address
    
        
   Set oEst = Nothing
End Sub

Private Function ParametroOperacion(psOpeCod As String, psMoneda As String) As Boolean
Dim sql As String
Dim rs As New ADODB.Recordset
Dim I As Integer

'************ CUENTAS CONTABLES DE OPERACION *******************
ParametroOperacion = False
ReDim lsOpeCuentas(0)
I = 0
Dim oOpe As New DOperacion
Set rs = oOpe.CargaOpeCtaArbol(psOpeCod)
If Not RSVacio(rs) Then
   Do While Not rs.EOF
       ReDim Preserve lsOpeCuentas(I)
       lsOpeCuentas(I) = Trim(rs!cCtaContCod)
       I = I + 1
       rs.MoveNext
   Loop
Else
   MsgBox "No se han Definido Cuentas contables en la Operación", vbInformation, "Aviso"
End If

'OBJETOS PARA EL REPORTE ANEXO 15
Dim oIF As New NCajaCtaIF
ReDim Preserve lsObjetosAnexo15(2)
Set rs = oOpe.CargaOpeObj(psOpeCod, , "2")
If Not rs.EOF Then
   lsObjetosAnexo15(0) = oIF.GetCtaIFDesc(rs!cOpeObjFiltro, psMoneda)
Else
    MsgBox "No se han Definido Banco BCR en Operación", vbInformation, "Aviso"
End If

Set rs = oOpe.CargaOpeObj(psOpeCod, , "3")
If Not rs.EOF Then
   lsObjetosAnexo15(1) = oIF.GetCtaIFDesc(rs!cOpeObjFiltro, psMoneda)
Else
    MsgBox "No se han Definido Banco de la Nación en Operación", vbInformation, "Aviso"
End If
RSClose rs
Set oOpe = Nothing
Set oIF = Nothing
ParametroOperacion = True
End Function

Private Sub ConsolidaBancos(psOpeCod As String, psMoneda As String, pdFecha As Date)
Dim N As Integer, sTexto As String
Dim nLin As Integer, P As Integer
Dim sBanco   As String
Dim sSql As String
Dim rs As New ADODB.Recordset
Dim nPos      As Integer
Dim K As Integer

If Not ParametroOperacion(psOpeCod, psMoneda) Then
   Exit Sub
End If

ReDim aBcoCon(8, 0)
Dim oEst As New NCajaCtaIF
Dim oAdeud As New NCajaAdeudados

'    Set rs = oEst.GetRepSaldosCtaIf(psOpeCod, pdFecha, "0,4", gEstadoCtaIFActiva & "','" & gEstadoCtaIFRestringida, psMoneda)
'    If rs.EOF Then
'       RSClose rs
'       MsgBox "No existen datos registrados de Bancos ", vbInformation, "Aviso"
'       Exit Sub
'    End If

nFil01 = 0: nFil02 = 0: nFil03 = 0: nFilBcoRestringido = 0

For K = 1 To 2

    If K = 1 Then
        Set rs = oEst.GetRepSaldosCtaIf(psOpeCod, pdFecha, "0", gEstadoCtaIFActiva & "','" & gEstadoCtaIFRestringida, psMoneda)
    ElseIf K = 2 Then
        Set rs = oEst.GetRepSaldosCtaIf(psOpeCod, pdFecha, "4", gEstadoCtaIFActiva & "','" & gEstadoCtaIFRestringida, psMoneda)
        If rs.EOF Then
        Else
            nFilBcoRestringido = nFil01 + 1
        End If
    End If
    
    Do While Not rs.EOF
       nFil01 = Mayor(nFil01, nFil02, nFil03) + 1
       nFil02 = nFil01
       nFil03 = nFil01
       'If k = 2 Then
'            If rs!cCtaIFEstado = gEstadoCtaIFRestringida Then      'Empezamos con Bancos Restringidos
'                nFilBcoRestringido = nFil01
'            End If
       'End If
       ReDim Preserve aBcoCon(8, nFil01)
       sBanco = Trim(rs!cBancoDesc)
       aBcoCon(1, nFil01) = sBanco
        'Adeudados por IF
       aBcoCon(8, nFil01) = Format(oAdeud.GetAdeudadosSaldoCap(Mid(rs!cCtaIfCod, 3, 13), , , , pdFecha, Left(rs!cCtaIfCod, 2) & Format(gTpoCtaIFCtaAdeud, "00") & Mid(rs!cCtaIfCod, 18, 1)), gsFormatoNumeroView)
        'Fin Adeudados por IF
       Do While Trim(rs!cBancoDesc) = Trim(sBanco)
          ReDim Preserve aBcoCon(8, Mayor(nFil01, nFil02, nFil03))
          ' ************  Cuentas Corrientes ***********************
          If CInt(Mid(rs!cCtaIfCod, 16, 2)) = gTpoCtaIFCtaCte Then
            aBcoCon(2, nFil01) = Trim(rs!cCtaIFDesc)
            aBcoCon(3, nFil01) = Format(IIf(rs!Debe > 0, rs!Debe, rs!Haber * -1), "#,#0.00")
            nFil01 = nFil01 + 1
          End If
          ' **************** Ahorros ************************
          If CInt(Mid(rs!cCtaIfCod, 16, 2)) = gTpoCtaIFCtaAho Then
            aBcoCon(4, nFil02) = Trim(rs!cCtaIFDesc)
            aBcoCon(5, nFil02) = Format(IIf(rs!Debe > 0, rs!Debe, rs!Haber * -1), "#,#0.00")
            nFil02 = nFil02 + 1
          End If
          
          '****************  plazo fijo *******************
          If CInt(Mid(rs!cCtaIfCod, 16, 2)) = gTpoCtaIFCtaPF Then
            aBcoCon(6, nFil03) = Trim(rs!cCtaIFDesc)
            aBcoCon(7, nFil03) = Format(IIf(rs!Debe > 0, rs!Debe, rs!Haber * -1), "#,#0.00")
            nFil03 = nFil03 + 1
          End If
          rs.MoveNext
          If rs.EOF Then
             Exit Do
          End If
       Loop
       nFil01 = nFil01 - 1
       nFil02 = nFil02 - 1
       nFil03 = nFil03 - 1
    Loop
    RSClose rs
    Set oEst = Nothing
    RSClose rs
Next

End Sub
Private Function Mayor(N1 As Integer, n2 As Integer, n3 As Integer) As Integer
Dim N As Integer
If N1 > n2 Then
    If N1 > n3 Then
        N = N1
    Else
        N = n3
    End If
ElseIf n2 > n3 Then
   N = n2
Else
   N = n3
End If
Mayor = N
End Function
Private Sub ConsolidaCmacs(psOpeCod As String, psMoneda As String, pdFecha As Date)
Dim sSql As String
Dim rs As New ADODB.Recordset
Dim j As Integer
Dim Total, N As Long
Dim sCmac As String
ReDim aCmacs(6, 0)

Dim oEst As New NCajaCtaIF
Set rs = oEst.GetRepSaldosCtaIf(psOpeCod, pdFecha, "'1'", gEstadoCtaIFActiva & "','" & gEstadoCtaIFRestringida, psMoneda)
'If rs.EOF Then
'   MsgBox "No existen datos registrados de Cmacs", vbInformation, "Aviso"
'   Exit Sub
'End If
nFil01 = 0: nFil02 = 0: nFil03 = 0
Do While Not rs.EOF
   nFil01 = Mayor(nFil01, nFil02, nFil03) + 1
   nFil02 = nFil01
   nFil03 = nFil01
   ReDim Preserve aCmacs(6, nFil01)
   sCmac = Trim(rs!cBancoDesc)
   aCmacs(1, nFil01) = sCmac
   Do While Trim(rs!cBancoDesc) = sCmac
      ReDim Preserve aCmacs(6, Mayor(nFil01, nFil02, nFil03))
      ' ************  Cuentas aHorros ***********************
      If CInt(Mid(rs!cCtaIfCod, 16, 2)) = gTpoCtaIFCtaAho Then
        aCmacs(2, nFil01) = Trim(rs!cCtaIFDesc)
        aCmacs(3, nFil01) = Format(CCur(rs!Debe) - CCur(rs!Haber * -1), "#,#0.00")
        nFil01 = nFil01 + 1
      End If
      '****************  plazo fijo *******************
      If CInt(Mid(rs!cCtaIfCod, 16, 2)) = gTpoCtaIFCtaPF Then
        aCmacs(4, nFil02) = Trim(rs!cCtaIFDesc)
        aCmacs(5, nFil02) = Format(CCur(rs!Debe) - CCur(rs!Haber * -1), "#,#0.00")
        nFil02 = nFil02 + 1
      End If
      rs.MoveNext
      If rs.EOF Then
         Exit Do
      End If
   Loop
   nFil01 = nFil01 - 1
   nFil02 = nFil02 - 1
   nFil03 = nFil03 - 1
Loop
RSClose rs
Set oEst = Nothing

End Sub

Private Function MuestraAdeudadosCortoPlazo(ByRef pnFila As Integer, pnMoneda As Integer, pdFecha As Date) As Currency
Dim sSql As String
Dim rs   As ADODB.Recordset
Dim K As Integer
Dim lnTotal As Currency
Dim oCon As New DConecta
sSql = "SELECT LEFT(cs.cCtaContCod,4), c.cCtaContDesc, SUM(nCtaSaldoImporte) nSaldo, ISNULL(SUM(nCtaSaldoImporteME),0) nSaldoME " _
     & "FROM CtaSaldo cs JOIN CtaCont c ON c.cCtaContCod = LEFT(cs.cCtaContCod,4) " _
     & "WHERE (cs.cCtaContCod LIKE '24" & pnMoneda & "10[129]%' or cs.cCtaContCod LIKE '24" & pnMoneda & "[236]%' " _
     & "    or cs.cCtaContCod LIKE '24" & pnMoneda & "80[1236]%') and dCtaSaldoFecha = ( SELECT Max(dCtaSaldoFecha) FROM CtaSaldo " _
     & "       WHERE cs.cCtaContCod = cCtaContCod and dCtaSaldoFecha <= '" & Format(pdFecha, gsFormatoFecha) & "') " _
     & "GROUP BY LEFT(cs.cCtaContCod,4), cCtaContDesc "
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
If Not rs.EOF Then
   xlHoja1.Cells(pnFila, 1) = "ADEUDADOS CORTO PLAZO"
   xlHoja1.Cells(pnFila, 3) = "SALDO"
   xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 3)).Font.Bold = True
   ExcelCuadro xlHoja1, 1, CCur(pnFila), 3, CCur(pnFila)
End If
K = 1
lnTotal = 0
Do While Not rs.EOF
   xlHoja1.Cells(pnFila + K, 1) = rs!cCtaContDesc
   If pnMoneda = 1 Then
      xlHoja1.Cells(pnFila + K, 3) = rs!nSaldo
   Else
      xlHoja1.Cells(pnFila + K, 3) = rs!nSaldoME
   End If
   lnTotal = lnTotal + xlHoja1.Cells(pnFila + K, 3)
   K = K + 1
   rs.MoveNext
Loop
ExcelCuadro xlHoja1, 1, pnFila + 1, 3, pnFila + K - 1
ExcelCuadro xlHoja1, 1, pnFila + K, 3, pnFila + K
xlHoja1.Cells(pnFila + K, 1) = "TOTAL ADEUDADOS"
xlHoja1.Cells(pnFila + K, 3) = lnTotal
xlHoja1.Range(xlHoja1.Cells(pnFila + K, 3), xlHoja1.Cells(pnFila + K, 3)).NumberFormat = "#0.00"
pnFila = pnFila + K
RSClose rs

End Function

Private Sub Form_Load()
CentraForm Me
End Sub
