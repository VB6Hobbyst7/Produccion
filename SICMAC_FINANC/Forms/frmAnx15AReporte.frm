VERSION 5.00
Begin VB.Form frmAnx15AReporte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Tesorería y Liquidéz Anexo No 15A"
   ClientHeight    =   690
   ClientLeft      =   1350
   ClientTop       =   2340
   ClientWidth     =   4260
   Icon            =   "frmAnx15AReporte.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmAnx15AReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsArchivo As String
Dim lbExcel As Boolean
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lnObligInmMN As Currency
Dim lnObligInmME As Currency
Dim lnDifTC As Currency

Dim lnChqPlazoMN As Currency
Dim lnChqPlazoME As Currency
Dim lnChqAhorroME As Currency
Dim lnChqAhorroMN As Currency
Dim ldFecha  As Date
Dim oBarra As clsProgressBar
Dim oCon As DConecta

Public Sub ImprimeAnexo15A(psOpeCod As String, psMoneda As String, pdFecha As Date)
On Error GoTo GeneraEstadError
   ldFecha = pdFecha
   'GeneraEstadisticaDiaria psOpeCod, psMoneda, pdFecha
   lsArchivo = App.path & "\SPOOLER\" & "Anx15A_" & Format(pdFecha, "mmyyyy") & IIf(psMoneda = "1", "MN", "ME") & ".XLS"
   lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
   If lbExcel Then
      ExcelAddHoja Format(pdFecha, "dd-mmmm-yyyy"), xlLibro, xlHoja1
      CargaDatos psOpeCod, psMoneda, pdFecha
      GeneraEstadisticaDiaria psOpeCod, psMoneda, pdFecha
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
      If lsArchivo <> "" Then
         CargaArchivo lsArchivo, App.path & "\SPOOLER\"
      End If
   End If
Exit Sub
GeneraEstadError:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    If lbExcel = True Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
End Sub

Private Function Vinculo(lsDato As String, lsMoneda As String) As String
Dim oAnx As New NEstadisticas
Vinculo = Format(oAnx.GetImporteEstadAnexos(ldFecha, lsDato, lsMoneda), gsFormatoNumeroDato)
Set oAnx = Nothing
End Function
   
Private Sub GeneraEstadisticaDiaria(psOpeCod As String, psMoneda As String, pdFecha As Date)
    Dim lsTotalActivos() As String
    Dim lsTotalPasivos() As String
    Dim lbExisteHoja  As Boolean
    Dim lsTotalesActivos() As String
    Dim lsTotalesPasivos() As String
    Dim I  As Long
    Dim Y1 As Integer, Y2 As Integer
    Dim lnFila As Integer
    Dim lnFilaFondosCaja As Integer
   Set oBarra = New clsProgressBar
   oBarra.ShowForm frmReportes
   oBarra.Max = 100
   oBarra.Progress 0, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue
    
    xlHoja1.PageSetup.Zoom = 80
    For I = 2 To 30
        If I <> 6 Then
            xlHoja1.Range(xlHoja1.Cells(I, 3), xlHoja1.Cells(I, 5)).Merge True
        End If
    Next
    ReDim lsTotalActivos(2)
    ReDim lsTotalPasivos(2)
    
    xlHoja1.Range("A1:R100").Font.Size = 8
    
    xlHoja1.Range("A1").ColumnWidth = 7
    xlHoja1.Range("B1").ColumnWidth = 27
    xlHoja1.Range("C1").ColumnWidth = 36
    xlHoja1.Range("D1:E1").ColumnWidth = 10
    xlHoja1.Range("F1:G1").ColumnWidth = 15
    
       
    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(6, 8)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(10, 8)).HorizontalAlignment = xlCenter
    
    xlHoja1.Range("B1:B50").HorizontalAlignment = xlLeft
    
    lnFila = 1
    xlHoja1.Cells(lnFila, 2) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "ANEXO Nº 15A"
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "REPORTE DE TESORERIA Y POSICION DE LIQUIDEZ"
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "(EN NUEVOS SOLES)"
    lnFila = lnFila + 2
    xlHoja1.Cells(lnFila, 2) = "EMPRESA :" & gsNomCmac:
    xlHoja1.Cells(lnFila, 6) = "Fecha :" & Format(pdFecha, "dd mmmm yyyy")
    
    lnFila = lnFila + 2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila + 2, 8)).Font.Bold = True
    xlHoja1.Cells(lnFila, 2) = "I RATIOS DE LIQUIDEZ"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Borders.LineStyle = xlContinuous
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 6) = "MONEDA ": xlHoja1.Cells(9, 7) = "MONEDA "
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 6) = "NACIONAL": xlHoja1.Cells(10, 7) = "EXTRANJERA"
    
    ExcelCuadro xlHoja1, 2, lnFila - 2, 7, lnFila   'ExcelCuadro 2, 8, 7, 10
    
    lnFila = lnFila + 1
    Y1 = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
    xlHoja1.Cells(lnFila, 3) = "Activos Líquidos"
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "1101":
    xlHoja1.Cells(lnFila, 3) = "Caja":
    'lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=" & Vinculo("CAJA", "1")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=" & Vinculo("CAJA", "2")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lsTotalActivos(1) = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalActivos(2) = xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "1102+1103+1107.03"
    xlHoja1.Cells(lnFila, 3) = "Bancos y Otras Instituciones Financieras del País"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=" & Vinculo("BANCOS", "1") & "-" & Vinculo("RESTBANCO", "1")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=" & Vinculo("BANCOS", "2") & "-" & Vinculo("RESTBANCO", "2")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "1104.01"
    xlHoja1.Cells(lnFila, 3) = "Bancos del exterior de Primera Categoría"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "1201-2201"
    xlHoja1.Cells(lnFila, 3) = "Fondos InterBancarios netos de deudores"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "1302.01.01.01 + 1302.01.09.01+ 1302.02.01+ 1304.01.01.01 + 1304.01.02 + 1304.01.09.01 + 1304.02.01"
    xlHoja1.Cells(lnFila, 3) = "Titulos Representativos de Deuda del Gobierno Central y Títulos Emitidos por el Banco Central"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "1302.05.01 + 1302.05.03 + 1304.05.01 + 1304.05.03"
    xlHoja1.Cells(lnFila, 3) = "Certificados de Depósitos y Certificados Bancarios"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "1302.01.01.02+1304.01.01.02+1302.01.09.02+1304.01.09.02+(1302.05(p)-1302.05.01-1302.05.03)+1302.06(p)+(1304.05(p)-1304.05.01-1304.05.03)+1304.06(p) "
    xlHoja1.Cells(lnFila, 3) = "Titulos Representativos de Deuda Pública y Sistema Financiero del Exterior"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    
    lsTotalActivos(1) = lsTotalActivos(1) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalActivos(2) = lsTotalActivos(2) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)
    
'    Me.prgBarra.value = 25
    '*************** TOTALES ACTIVOS DE LIQUIDEZ *************************
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
    xlHoja1.Cells(lnFila, 3) = "TOTAL(A)"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=Sum(" & lsTotalActivos(1) & ")"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=Sum(" & lsTotalActivos(2) & ")"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    ReDim lsTotalesActivos(2)
    ReDim lsTotalesPasivos(2)
    
    lsTotalesActivos(1) = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalesActivos(2) = xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)
    
    lnFila = lnFila + 2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).Font.Bold = True
    xlHoja1.Cells(lnFila, 3) = "Pasivos de Corto Plazo"
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "2101+2104+2301+2105":
    xlHoja1.Cells(lnFila, 3) = "Obligaciones Inmediatas":
    'xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "0"
    'xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "0"
    xlHoja1.Cells(lnFila, 6) = lnObligInmMN
    xlHoja1.Cells(lnFila, 7) = lnObligInmME
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lsTotalPasivos(1) = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalPasivos(2) = xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "2201+1201"
    xlHoja1.Cells(lnFila, 3) = "Fondos Interbancarios netos acreedores"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "2102+2302"
    xlHoja1.Cells(lnFila, 3) = "Depósitos de Ahorros"
    
    AsignaSaldoAhorros lnFila, 1, pdFecha
    AsignaSaldoAhorros lnFila, 2, pdFecha

    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "2103-2103.05+2303"
    xlHoja1.Cells(lnFila, 3) = "Depósito a Plazo por Vencer dentro de 360 días"
    AsignaSaldoPlazoFijo lnFila, 1
    AsignaSaldoPlazoFijo lnFila, 2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "2400+2800"
    xlHoja1.Cells(lnFila, 3) = "Adeudados y Otras Obligaciones Financieras por Vencer dentro de 360 Días"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=" & Vinculo("ADEUDADO", "1")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=" & Vinculo("ADEUDADO", "2")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lsTotalPasivos(1) = lsTotalPasivos(1) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalPasivos(2) = lsTotalPasivos(2) + ":" + xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)
    
    '******************** TOTALES DE PASIVOS DE CORTO PLAZO ***************************
'    Me.prgBarra.value = 50
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 8)).Font.Bold = True
    xlHoja1.Cells(lnFila, 3) = "TOTAL(B)"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=Sum(" & lsTotalPasivos(1) & ")"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=Sum(" & lsTotalPasivos(2) & ")"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lsTotalesPasivos(1) = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsTotalesPasivos(2) = xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)
    
    Y2 = lnFila
    xlHoja1.Range(xlHoja1.Cells(Y1, 2), xlHoja1.Cells(Y2, 7)).Borders.LineStyle = xlContinuous
    
    lnFila = lnFila + 2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 5)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Cells(lnFila, 2) = "Ratios de Liquidez[(a)/(b)]*100"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=(" & lsTotalesActivos(1) & "/" & lsTotalesPasivos(1) & ")*100"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=(" & lsTotalesActivos(2) & "/" & lsTotalesPasivos(2) & ")*100"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    ExcelCuadro xlHoja1, 2, lnFila, 7, lnFila
    
    lnFila = lnFila + 2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 5)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Cells(lnFila, 2) = "Activos Liquidos Ajustados por Recursos Prestados(C) "
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    ExcelCuadro xlHoja1, 2, lnFila, 7, lnFila
    
    lnFila = lnFila + 2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 5)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Cells(lnFila, 2) = "Pasivos de corto Plazo Ajustados por Recursos Prestados (d) "
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    ExcelCuadro xlHoja1, 2, lnFila, 7, lnFila
    
    lnFila = lnFila + 2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 5)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Cells(lnFila, 2) = "Ratio de Liquidez Ajustados por Recusos Prestados [(C)*/(d)]*100"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    ExcelCuadro xlHoja1, 2, lnFila, 7, lnFila
    
   '**********************   ENCAJE **********************************************
    lnFila = lnFila + 2
    Y1 = lnFila
    xlHoja1.Cells(lnFila, 2) = "II. ENCAJE"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Cells.BorderAround
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 7)).Borders.LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Font.Bold = True
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 6) = "MONEDA ": xlHoja1.Cells(lnFila, 7) = "MONEDA "
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Font.Bold = True
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 6) = "NACIONAL": xlHoja1.Cells(lnFila, 7) = "EXTRANJERA"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).Font.Bold = True
    
    Y2 = lnFila
    ExcelCuadro xlHoja1, 2, Y1, 7, Y2
    
    lnFila = lnFila + 1
    Y1 = lnFila
    xlHoja1.Cells(lnFila, 2) = "1. Total de Obligaciones sujetas a Encaje consolidados a nivel nacional(TOSE)"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = " . Obligaciones a Plazo Mayor a 30 días"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=" & Vinculo("OBLIGPLAZO", "1")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=" & Vinculo("OBLIGPLAZO", "2")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = " . Ahorros"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=" & Vinculo("AHORROS", "1")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=" & Vinculo("AHORROS", "2")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    xlHoja1.Range(xlHoja1.Cells(Y1, 6), xlHoja1.Cells(Y1, 6)).Formula = "=SUM(F" & Y1 + 1 & ":F" & lnFila & ")"
    xlHoja1.Range(xlHoja1.Cells(Y1, 7), xlHoja1.Cells(Y1, 7)).Formula = "=SUM(G" & Y1 + 1 & ":G" & lnFila & ")"
    
    Y2 = lnFila
    ExcelCuadro xlHoja1, 2, Y1, 7, Y2
    
    
    lnFila = lnFila + 1
    Y1 = lnFila
    xlHoja1.Cells(lnFila, 2) = "2. Posición de Encaje"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
  
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "   2.1. Encaje Exigible"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=" & Vinculo("EXIGIBLE", "1")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=" & Vinculo("EXIGIBLE", "2")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lsPosEnResultados(1) = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsPosEnResultados(2) = xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)
    
    lnFila = lnFila + 1
    lnFilaFondosCaja = lnFila
    xlHoja1.Cells(lnFila, 2) = "   2.2. Fondos de Caja"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "      . Caja"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=" & Vinculo("FONDOCAJA", "1")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=" & Vinculo("FONDOCAJA", "2")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lsPosFondosCaja(1) = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsPosFondosCaja(2) = xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "      . Cuenta Corriente BCRP"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=" & Vinculo("FONDOSBCR", "1")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=" & Vinculo("FONDOSBCR", "2")
    lsPosFondosCaja(1) = lsPosFondosCaja(1) & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
    lsPosFondosCaja(2) = lsPosFondosCaja(2) & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address(False, False)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
'    Me.prgBarra.value = 75
    
    xlHoja1.Range(xlHoja1.Cells(lnFilaFondosCaja, 6), xlHoja1.Cells(lnFilaFondosCaja, 6)).Formula = "=Sum(" & lsPosFondosCaja(1) & ")"
    xlHoja1.Range(xlHoja1.Cells(lnFilaFondosCaja, 7), xlHoja1.Cells(lnFilaFondosCaja, 7)).Formula = "=Sum(" & lsPosFondosCaja(2) & ")"
    
    xlHoja1.Range(xlHoja1.Cells(lnFilaFondosCaja, 6), xlHoja1.Cells(lnFilaFondosCaja, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFilaFondosCaja, 7), xlHoja1.Cells(lnFilaFondosCaja, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lsPosEnResultados(1) = xlHoja1.Range(xlHoja1.Cells(lnFilaFondosCaja, 6), xlHoja1.Cells(lnFilaFondosCaja, 6)).Address(False, False) & "-" & lsPosEnResultados(1)
    lsPosEnResultados(2) = xlHoja1.Range(xlHoja1.Cells(lnFilaFondosCaja, 7), xlHoja1.Cells(lnFilaFondosCaja, 7)).Address(False, False) & "-" & lsPosEnResultados(2)
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "   2.3. Resultados del día(Fondos de Encaje - encaje Exigible)"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=" & lsPosEnResultados(1)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=" & lsPosEnResultados(2)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "   2.4. Posición de encaje acumulada del período"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = Vinculo("ENCAJE_ACUMULA", "1")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = Vinculo("ENCAJE_ACUMULA", "2")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    xlHoja1.Range(xlHoja1.Cells(Y1, 6), xlHoja1.Cells(Y1, 6)).Formula = "=SUM(F" & Y1 + 1 & ":F" & lnFila & ")"
    xlHoja1.Range(xlHoja1.Cells(Y1, 7), xlHoja1.Cells(Y1, 7)).Formula = "=SUM(G" & Y1 + 1 & ":G" & lnFila & ")"
    
    Y2 = lnFila
    ExcelCuadro xlHoja1, 2, Y1, 7, Y2
    
    lnFila = lnFila + 1
    Y1 = lnFila
    xlHoja1.Cells(lnFila, 2) = "3. Cheques deducidos del total de obligaciones sujetas a encaje"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "0"
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "   3.1. Deducidos de Obligaciones a la Vista y a Plazo hasta 30 días"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "0"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "   3.2. Deducidos de obligaciones a plazo mayor de 30 días"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=" & lnChqPlazoMN
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=" & lnChqPlazoME
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "   3.3. Deducidos de Ahorros"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=" & lnChqAhorroMN
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=" & lnChqAhorroME
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    xlHoja1.Range(xlHoja1.Cells(Y1, 6), xlHoja1.Cells(Y1, 6)).Formula = "=SUM(F" & Y1 + 1 & ":F" & lnFila & ")"
    xlHoja1.Range(xlHoja1.Cells(Y1, 7), xlHoja1.Cells(Y1, 7)).Formula = "=SUM(G" & Y1 + 1 & ":G" & lnFila & ")"
    
    Y2 = lnFila
    ExcelCuadro xlHoja1, 2, Y1, 7, Y2
    
    lnFila = lnFila + 1
    
    Y1 = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 10)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Merge True:
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 7)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 7)).Borders.LineStyle = xlContinuous
    xlHoja1.Cells(lnFila, 4) = "Tasas de Interés Promedio": xlHoja1.Cells(lnFila, 6) = "Saldo"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 10)).HorizontalAlignment = xlCenter
    xlHoja1.Cells(lnFila, 4) = "Moneda ":   xlHoja1.Cells(lnFila, 5) = "Moneda ":
    xlHoja1.Cells(lnFila, 6) = "Moneda ": xlHoja1.Cells(lnFila, 7) = "Moneda "
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 10)).HorizontalAlignment = xlCenter
    xlHoja1.Cells(lnFila, 4) = "Nacional":
    xlHoja1.Cells(lnFila, 5) = "Extranjera":
    xlHoja1.Cells(lnFila, 6) = "Nacional":
    xlHoja1.Cells(lnFila, 7) = "Extranjera"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    Y2 = lnFila
    
    ExcelCuadro xlHoja1, 2, Y1, 7, Y2
    
    lnFila = lnFila + 1
    Y1 = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    xlHoja1.Cells(lnFila, 2) = "4. Fondos Interbancarios(Saldos por Institución)"
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    xlHoja1.Cells(lnFila, 2) = "   4.1 Activos(cuenta 1201)"
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    xlHoja1.Cells(lnFila, 2) = "   4.2 Pasivos(cuenta 2201)"
    
    xlHoja1.Range(xlHoja1.Cells(Y1, 6), xlHoja1.Cells(Y1, 6)).Formula = "=SUM(F" & Y1 + 1 & ":F" & lnFila & ")"
    xlHoja1.Range(xlHoja1.Cells(Y1, 7), xlHoja1.Cells(Y1, 7)).Formula = "=SUM(G" & Y1 + 1 & ":G" & lnFila & ")"
    
    Y2 = lnFila
    ExcelCuadro xlHoja1, 2, Y1, 7, Y2
    
    lnFila = lnFila + 1
    Y1 = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    xlHoja1.Cells(lnFila, 2) = "5. Obligaciones con Banco de la nación"
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "0" '"=" & Vinculo("BCONACION", "1")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "0"  '"=" & Vinculo("BCONACION", "2")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lnFila = lnFila + 1
    
    Y2 = lnFila:
    ExcelCuadro xlHoja1, 2, Y1, 7, Y2
    Y1 = Y2
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    xlHoja1.Cells(lnFila, 2) = "6. Operaciones de reporte con CD BCRP"
    
    lnFila = lnFila + 1
    Y2 = lnFila
    ExcelCuadro xlHoja1, 2, Y1, 7, Y2
    Y1 = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 3)).Merge True
    xlHoja1.Cells(lnFila, 2) = "7. Créditos del BCr con fines de regulaciòn monetaria"
    Y2 = lnFila
    ExcelCuadro xlHoja1, 2, Y1, 7, Y2
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "8. Venta Temporal de Moneda Extranjera al BCR"
    Y2 = lnFila
    ExcelCuadro xlHoja1, 2, Y1, 3, Y2
        
    lnFila = lnFila + 2
    Y1 = lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Font.Bold = True
    xlHoja1.Cells(lnFila, 3) = "III Número de Dias de Redescuento "
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Merge True
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 5)).Font.Bold = True
    xlHoja1.Cells(lnFila, 3) = "LOS ULTIMOS 180 DIAS "
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Merge True
    ExcelCuadro xlHoja1, 3, Y1, 5, lnFila, False
    
    lnFila = lnFila + 2
    xlHoja1.Cells(lnFila, 3) = "VI.  POSICION DE CAMBIO "
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Cells(lnFila, 4) = "En Moneda Extranjera US$"
    ExcelCuadro xlHoja1, 3, lnFila, 5, lnFila
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "1. Balance "
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Cells(lnFila, 4) = lnDifTC
    ExcelCuadro xlHoja1, 3, lnFila, 5, lnFila
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "2. Global "
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 5)).Merge True
    xlHoja1.Cells(lnFila, 4) = lnDifTC
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    ExcelCuadro xlHoja1, 3, lnFila, 5, lnFila
        
    oBarra.Progress 100, "ANEXO 15A: Tesorería y Posición Diaria de Liquidez", "REPORTE TERMINADO", "", vbBlue
End Sub

Private Sub AsignaSaldoAhorros(pnFila As Integer, psMoneda As String, pdFecha As Date)
Dim oEst   As New NEstadisticas
Dim oSdo   As New NCtasaldo
Dim rs     As ADODB.Recordset
Dim nSaldo As Currency

Set rs = oEst.GetEstadisticaAhorro(gbBitCentral, pdFecha, pdFecha, psMoneda)
    If Not rs.EOF Then
      nSaldo = oSdo.GetCtaSaldo("21" & psMoneda & "202010199", Format(pdFecha, gsFormatoFecha))
      xlHoja1.Range(xlHoja1.Cells(pnFila, 5 + psMoneda), xlHoja1.Cells(pnFila, 5 + psMoneda)).Formula = "=" & rs!nSaldoAC & "+" & rs!nMonChqVal & "+" & rs!nChqCMAC & "+" & nSaldo
      If psMoneda = "1" Then
         lnChqAhorroMN = rs!nMonChqVal + rs!nChqCMAC
      Else
         lnChqAhorroME = rs!nMonChqVal + rs!nChqCMAC
      End If
    End If
    RSClose rs
Set oSdo = Nothing
Set oEst = Nothing
End Sub

Private Sub AsignaSaldoPlazoFijo(pnFila As Integer, pnMoneda As Integer)
    Dim sSql As String
    Dim rs   As ADODB.Recordset
   sSql = "SELECT SUM(nSaldoPF) nSaldoPF, SUM(nSaldCMAC) nSaldCMAC, SUM(nChqCmac) nChqCMAC, SUM(nMonChqVal) nMonChqVal, SUM(nSaldCRAC) nSaldCRAC " _
        & "FROM ( SELECT nSaldoPF, nSaldCMAC, 0 nChqCmac, nMonChqVal, nSaldCRAC " _
        & "       FROM EstadDiaPFConsol WHERE datediff(d,dEstadPF, '" & Format(ldFecha, gsFormatoFecha) & "') = 0 and cMoneda = '" & pnMoneda & "' " _
        & "       UNION " _
        & "       SELECT 0 nSaldoPF, 0 nSaldCMAC, 0 nChqCMAC, 0 nMonChqVal, 0 nSaldCRAC " _
        & "       FROM EstadDiaCTSConsol WHERE datediff(d,dEstadCTS, '" & Format(ldFecha, gsFormatoFecha) & "') = 0 and cMoneda = '" & pnMoneda & "' " _
        & "     ) a"
    Set oCon = New DConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
      xlHoja1.Range(xlHoja1.Cells(pnFila, 5 + pnMoneda), xlHoja1.Cells(pnFila, 5 + pnMoneda)).Formula = "=" & rs!nSaldoPF & "+" & rs!nMonChqVal & "+" & rs!nChqCMAC
      If pnMoneda = 1 Then
         lnChqPlazoMN = rs!nMonChqVal + rs!nChqCMAC
      Else
         lnChqPlazoME = rs!nMonChqVal + rs!nChqCMAC
      End If
    End If
    oCon.CierraConexion
    RSClose rs
End Sub

Sub CargaDatos(psOpeCod As String, psMoneda As String, pdFecha As Date)
lnDifTC = EmiteDifTipoCambio(pdFecha)
If lnDifTC = 0 Then
    MsgBox "No se ha realizado el Cálculo de Diferencia de Tipo de Cambio " + vbCrLf + _
            "Posiblemente no se ha realizado la consolidación diaria del Tipo Cambio " + _
            vbCrLf + "Por favor, consulte a las áreas involucradas", vbInformation, "¡Aviso!"
End If
lnObligInmMN = EmiteObligacionesInmediatas(psOpeCod, "1", pdFecha)
lnObligInmME = EmiteObligacionesInmediatas(psOpeCod, "2", pdFecha)
End Sub

Function EmiteDifTipoCambio(pdFecha As Date) As Currency
Dim sql      As String
Dim rs       As ADODB.Recordset
Dim nImporte As Currency
Dim oDol As New NCompraVenta
Set rs = oDol.GetImporteCompraVenta(gOpeCajeroMECompra, pdFecha, pdFecha, "")
If Not rs.EOF Then
   nImporte = rs!TotalDol
End If

Set rs = oDol.GetImporteCompraVenta(gOpeCajeroMEVenta, pdFecha, pdFecha, "")
If Not rs.EOF Then
   nImporte = nImporte - rs!TotalDol
End If
RSClose rs
Set oDol = Nothing
EmiteDifTipoCambio = nImporte
End Function

Function EmiteObligacionesInmediatas(psOpeCod As String, psMoneda As String, pdFecha As Date) As Currency
Dim oSdo    As New NCtasaldo
Dim lsOrden As String
If psMoneda = "1" Then
   lsOrden = "0"
Else
   lsOrden = "1"
End If
EmiteObligacionesInmediatas = oSdo.GetOpeCtaSaldo(psOpeCod, Format(pdFecha, gsFormatoFecha), IIf(psMoneda = "1", True, False), lsOrden)

End Function


Private Sub Form_Load()
Set oCon = New DConecta
oCon.AbreConexion
Set oBarra = New clsProgressBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
oCon.CierraConexion
Set oCon = Nothing
Set oBarra = Nothing
End Sub

