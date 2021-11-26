VERSION 5.00
Begin VB.Form frmAnx15AEfectivoCaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Descomposición en Efectivo - Caja General"
   ClientHeight    =   975
   ClientLeft      =   1425
   ClientTop       =   2700
   ClientWidth     =   4815
   Icon            =   "frmAnx15AEfectivoCaja.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmAnx15AEfectivoCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlAplicacion As Excel.Application
Dim xlLibro      As Excel.Workbook
Dim xlHoja1      As Excel.Worksheet
Dim lsArchivo    As String

Dim lsBilletesSoles()   As String
Dim lsMonedasSoles()    As String
Dim lsBilletesDolares() As String
Dim lsMonedaDolares()   As String
Dim lsTotBillSol()      As String
Dim lsTotBillDol()      As String
Dim lsTotMonSol()       As String
Dim lsTotMonDol()       As String

Dim lsTotalDol() As String
Dim lsTotalSol() As String

'************* matrices para el calculo de cobertura ***********************
Dim lsIndiceCob()    As String
Dim lsTotalSolesTC() As String
Dim lsNivCobEfect()  As String
Dim lsCobEfect()     As String

Dim lsCoberturado() As String
Dim lsCobertura() As String
Dim lsIndiceCobertura() As String

Dim lsTotalOro() As String
    
Dim lnTotalColumnas As Long
Dim lbExcel     As Boolean
Dim lnTipCambio As Currency
Dim oBarra As clsProgressBar

Public Sub ImprimeEfectivoCajaNew(psOpeCod As String, psMoneda As String, pdFecha As Date)
Dim lbHojaActiva As Boolean
Dim oTC As nTipoCambio
Set oTC = New nTipoCambio
Set oBarra = New clsProgressBar
    oBarra.Max = 100
    'ProgressShow oBarra, frmReportes, eCap_CaptionPercent, True
    oBarra.ShowForm frmReportes
    oBarra.Progress 10, "DESCOMPOSICION DE EFECTIVO", "Carga Inicial", , vbBlue
    lnTipCambio = oTC.EmiteTipoCambio(pdFecha, TCFijoMes)
    lsArchivo = App.path & "\SPOOLER\" & "Anx_DescompEfectivo_" & Format(pdFecha, "yyyymmdd") & IIf(psMoneda = "1", "MN", "ME") & "_" & Format(Time(), "HHMM") & ".XLS"
    lbHojaActiva = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    If lbHojaActiva Then
        ExcelAddHoja Format(pdFecha, "dd-mm-yyyy"), xlLibro, xlHoja1
        GeneraReporteNew psOpeCod, psMoneda, pdFecha, lnTipCambio
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
        If lsArchivo <> "" Then
            CargaArchivo lsArchivo, App.path & "\SPOOLER\"
        End If
    End If
End Sub 'NAGL 20180920 Según TIC1807210002

Private Sub GeneraReporteNew(psOpeCod As String, lsMoneda As String, pdFecha As Date, lnTipCambio As Currency)
Dim DEfec As New Defectivo
Dim rs As New ADODB.Recordset
Dim rsAge As New ADODB.Recordset
Dim nCantAge As Long
Dim liFilaIni As Long, liFila As Long, liColIni As Long, liCols As Long
Dim cTpoEfecAnt As String, cTpoEfec As String, cTpoAgeAnt As String, cTpoAge As String
Dim pscCtaContCod As String
Dim lsCadenaSum() As String
Dim oOpe As New DOperacion
Dim nTotal110701 As Currency
Dim dBalan As New DbalanceCont

'Set oBarra = New clsProgressBar
'oBarra.Max = 100
''ProgressShow oBarra, frmReportes, eCap_CaptionPercent, True
'oBarra.ShowForm frmReportes
'oBarra.Progress 10, "DESCOMPOSICION DE EFECTIVO", "Carga Inicial", , vbBlue
cTpoEfec = ""
cTpoEfecAnt = ""
ReDim lsCadenaSum(4)
lsCadenaSum(1) = ""
lsCadenaSum(2) = ""
lsCadenaSum(3) = ""
lsCadenaSum(4) = ""
nCantAge = 0
xlAplicacion.Range("A7:Z10000").Font.Size = 9
xlAplicacion.Range("A7:Z10000").Font.Name = "Arial"

xlAplicacion.Range("A1:M6").Font.Size = 11
xlAplicacion.Range("A1:M6").Font.Name = "Calibri"
 xlHoja1.Range("A1").ColumnWidth = 5
xlHoja1.Cells(1, 2) = gsNomCmac
xlHoja1.Range(xlHoja1.Cells(1, 2), xlHoja1.Cells(1, 3)).Merge True
xlHoja1.Range(xlHoja1.Cells(1, 2), xlHoja1.Cells(1, 3)).HorizontalAlignment = xlCenter
xlHoja1.Cells(1, 12) = "Fecha :" & Format(pdFecha, "dd mmmm yyyy")
xlHoja1.Range(xlHoja1.Cells(1, 12), xlHoja1.Cells(1, 13)).Merge True
xlHoja1.Range(xlHoja1.Cells(1, 12), xlHoja1.Cells(1, 13)).HorizontalAlignment = xlCenter
xlHoja1.Cells(3, 6) = "DESCOMPOSICION DE EFECTIVO " & IIf(lsMoneda = "1", "MONEDA NACIONAL", "MONEDA EXTRANJERA")
xlHoja1.Range(xlHoja1.Cells(3, 6), xlHoja1.Cells(3, 9)).Merge True
xlHoja1.Range(xlHoja1.Cells(3, 6), xlHoja1.Cells(3, 9)).HorizontalAlignment = xlCenter
xlHoja1.Cells(4, 6) = "Area de Caja General"
xlHoja1.Range(xlHoja1.Cells(4, 6), xlHoja1.Cells(4, 9)).Merge True
xlHoja1.Range(xlHoja1.Cells(4, 6), xlHoja1.Cells(4, 9)).HorizontalAlignment = xlCenter
xlHoja1.Range(xlHoja1.Cells(1, 2), xlHoja1.Cells(5, 13)).Font.Bold = True

'TIPO DE EFECTIVO BILLETAJE - MONEDA
oBarra.Progress 20, "DESCOMPOSICION DE EFECTIVO", "Cargando Tipo de Billetaje", , vbBlue
liFilaIni = 8
liFila = liFilaIni
Set rs = DEfec.GetCtaEfectivoTrans_Age(pdFecha, lsMoneda, "Age")
If Not (rs.EOF And rs.BOF) Then
    nCantAge = rs.RecordCount
End If
Set rs = DEfec.ObtieneTipoEfectivo(lsMoneda)
If Not (rs.EOF And rs.BOF) Then
    xlHoja1.Cells(liFila - 1, 2) = "Billetes"
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 2), xlHoja1.Cells(liFila - 1, 2)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 2), xlHoja1.Cells(liFila - 1, 2)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, 2), xlHoja1.Cells(liFila - 1, 2)).Interior.Color = 16764057
    ExcelCuadro xlHoja1, 2, liFila - 1, 2, CCur(liFila - 1)
    
    liCols = 2 + 1 + nCantAge + 1 'Billetaje - Caja - Agencias - Total
    xlHoja1.Cells(liFila - 1, liCols) = "TOTALES"
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, liCols), xlHoja1.Cells(liFila - 1, liCols)).Interior.Color = 16764057
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, liCols), xlHoja1.Cells(liFila - 1, liCols)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, liCols), xlHoja1.Cells(liFila - 1, liCols)).Font.Bold = True
    ExcelCuadro xlHoja1, liCols, liFila - 1, liCols, CCur(liFila - 1)
    
    Do While Not rs.EOF
       cTpoEfec = rs!cTipo
       If cTpoEfecAnt <> "" Then
            If (cTpoEfec = cTpoEfecAnt) Then
                 liFila = liFila + 1
            Else
               ExcelCuadro xlHoja1, 2, liFilaIni, 2, CCur(liFila + 1)
               liFila = liFila + 2
               xlHoja1.Cells(liFila, 2) = "TOTALES BILLETES"
               ExcelCuadro xlHoja1, 2, liFila, 2, CCur(liFila)
            
               liCols = 2 + 1 + nCantAge + 1
               ExcelCuadro xlHoja1, liCols, liFilaIni, liCols, CCur(liFila)
               xlHoja1.Cells(liFila, liCols).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(liFilaIni, liCols), xlHoja1.Cells(liFila - 2, liCols)).Address(False, False) & ")"
               xlHoja1.Range(xlHoja1.Cells(liFila, liCols), xlHoja1.Cells(liFila, liCols)).NumberFormat = "#,##0.00;-#,##0.00"
               xlHoja1.Range(xlHoja1.Cells(liFila, liCols), xlHoja1.Cells(liFila, liCols)).Font.Bold = True
               ExcelCuadro xlHoja1, liCols, liFila, liCols, CCur(liFila)
               ExcelCuadro xlHoja1, liCols, liFilaIni, liCols, CCur(liFila)
               
               xlHoja1.Cells(liFila + 2, 2) = "MONEDAS"
               xlHoja1.Range(xlHoja1.Cells(liFila + 2, 2), xlHoja1.Cells(liFila + 2, 2)).Interior.Color = 16764057
               xlHoja1.Range(xlHoja1.Cells(liFila + 2, 2), xlHoja1.Cells(liFila + 2, 2)).HorizontalAlignment = xlCenter
               ExcelCuadro xlHoja1, 2, liFila + 2, 2, CCur(liFila + 2)
               liFila = liFila + 3
               liFilaIni = liFila
            End If
       End If
       xlHoja1.Cells(liFila, 2) = rs!Descripcion
       liCols = 2 + 1 + nCantAge + 1 'Billetaje - Caja - Agencias - Total
       xlHoja1.Cells(liFila, liCols).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(liFila, 3), xlHoja1.Cells(liFila, liCols - 1)).Address(False, False) & ")"
       xlHoja1.Range(xlHoja1.Cells(liFila, liCols), xlHoja1.Cells(liFila, liCols)).NumberFormat = "#,##0.00;-#,##0.00"
       cTpoEfecAnt = cTpoEfec
       rs.MoveNext
    Loop
    xlHoja1.Cells(liFila + 2, 2) = "TOTALES MONEDAS"
    ExcelCuadro xlHoja1, 2, liFila + 2, 2, CCur(liFila + 2)
    ExcelCuadro xlHoja1, 2, liFilaIni, 2, CCur(liFila + 1)
    liFilaIni = 8
    xlHoja1.Range(xlHoja1.Cells(liFilaIni, 2), xlHoja1.Cells(liFila + 2, 2)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liFilaIni, 2), xlHoja1.Cells(liFilaIni, 2)).EntireColumn.AutoFit
End If

oBarra.Progress 40, "DESCOMPOSICION DE EFECTIVO", "Generando Billetes de Caja General y Agencias", , vbBlue
'CAJA GENERAL Y DESCOMP.EFECTIVO EN AGENCIAS
Set rs = Nothing
liColIni = 3
liFilaIni = 8
liCols = liColIni
liFila = liFilaIni
cTpoEfec = ""
cTpoEfecAnt = ""
cTpoAge = ""
cTpoAgeAnt = ""

'****Cuenta para Caja General
pscCtaContCod = oOpe.EmiteOpeCta(psOpeCod, "D", "0")
pscCtaContCod = Left(pscCtaContCod, 2) & lsMoneda & Mid(pscCtaContCod, 4, 22)
'****
Set rs = DEfec.EmiteBilletajeMonedaNew(pdFecha, lsMoneda, pscCtaContCod)
If Not (rs.EOF And rs.BOF) Then
    Do While Not rs.EOF
        cTpoEfec = rs!cTipo
        cTpoAge = rs!cAgeCod
        If cTpoAgeAnt <> "" Then
            If (cTpoAge <> cTpoAgeAnt) Then
                ExcelCuadro xlHoja1, liCols, liFilaIni, liCols, CCur(liFila + 1)
               liFila = liFila + 2
               xlHoja1.Cells(liFila, liCols).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(liFilaIni, liCols), xlHoja1.Cells(liFila - 2, liCols)).Address(False, False) & ")"
               xlHoja1.Range(xlHoja1.Cells(liFila, liCols), xlHoja1.Cells(liFila, liCols)).Font.Bold = True
               ExcelCuadro xlHoja1, liCols, liFila, liCols, CCur(liFila)
               xlHoja1.Range(xlHoja1.Cells(liFilaIni, liCols - 1), xlHoja1.Cells(liFila, liCols)).NumberFormat = "#,##0.00;-#,##0.00"
               
               '*****RESUMEN TOTAL*****
               lsCadenaSum(2) = xlHoja1.Range(xlHoja1.Cells(liFila, liCols), xlHoja1.Cells(liFila, liCols)).Address(False, False)
               lsCadenaSum(3) = xlHoja1.Range(xlHoja1.Cells(liFila + 3, liCols), xlHoja1.Cells(liFila + 3, liCols)).Address(False, False)
               xlHoja1.Cells(liFila + 5, liCols).Formula = "=" & "Sum" & "(" & lsCadenaSum(1) & "+" & lsCadenaSum(2) & "+" & lsCadenaSum(3) & ")"
               ExcelCuadro xlHoja1, liCols, liFila + 5, liCols, CCur(liFila + 5)
               '**********************
               liCols = liCols + 1
               liFilaIni = 8
               liFila = liFilaIni
               cTpoEfecAnt = ""
               xlHoja1.Range(xlHoja1.Cells(liFilaIni, liCols - 1), xlHoja1.Cells(liFilaIni, liCols - 1)).EntireColumn.AutoFit
            End If
        End If
        
        If cTpoEfecAnt <> "" Then
            If (cTpoEfec = cTpoEfecAnt) Then
               liFila = liFila + 1
            Else
               ExcelCuadro xlHoja1, liCols, liFilaIni, liCols, CCur(liFila + 1)
               liFila = liFila + 2
               xlHoja1.Cells(liFila, liCols).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(liFilaIni, liCols), xlHoja1.Cells(liFila - 2, liCols)).Address(False, False) & ")"
               xlHoja1.Range(xlHoja1.Cells(liFila, liCols), xlHoja1.Cells(liFila, liCols)).Font.Bold = True
               ExcelCuadro xlHoja1, liCols, liFila, liCols, CCur(liFila)
               xlHoja1.Range(xlHoja1.Cells(liFilaIni, liCols - 1), xlHoja1.Cells(liFila, liCols + 1)).NumberFormat = "#,##0.00;-#,##0.00"
               lsCadenaSum(1) = xlHoja1.Range(xlHoja1.Cells(liFila, liCols), xlHoja1.Cells(liFila, liCols)).Address(False, False)
               
               If liCols = 2 + 1 + nCantAge Then 'La Primera Celda del Total
                    lsCadenaSum(4) = xlHoja1.Range(xlHoja1.Cells(liFila, liCols + 1), xlHoja1.Cells(liFila, liCols + 1)).Address(False, False) 'Resumen Total
               End If
               
               xlHoja1.Cells(liFila + 2, liCols) = rs!cAgeDescripcion
               xlHoja1.Range(xlHoja1.Cells(liFila + 2, liCols), xlHoja1.Cells(liFila + 2, liCols)).Interior.Color = 16764057
               xlHoja1.Range(xlHoja1.Cells(liFila + 2, liCols), xlHoja1.Cells(liFila + 2, liCols)).HorizontalAlignment = xlCenter
               xlHoja1.Range(xlHoja1.Cells(liFila + 2, liCols), xlHoja1.Cells(liFila + 2, liCols)).Font.Bold = True
               ExcelCuadro xlHoja1, liCols, liFila + 2, liCols, CCur(liFila + 2)
               
               liFila = liFila + 3
               liFilaIni = liFila
            End If
        End If
        
        If xlHoja1.Cells(liFilaIni - 1, liCols) = "" Then
            xlHoja1.Cells(liFilaIni - 1, liCols) = rs!cAgeDescripcion
            xlHoja1.Range(xlHoja1.Cells(liFilaIni - 1, liCols), xlHoja1.Cells(liFilaIni - 1, liCols)).Interior.Color = 16764057
            xlHoja1.Range(xlHoja1.Cells(liFilaIni - 1, liCols), xlHoja1.Cells(liFilaIni - 1, liCols)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liFilaIni - 1, liCols), xlHoja1.Cells(liFilaIni - 1, liCols)).Font.Bold = True
            ExcelCuadro xlHoja1, liCols, liFilaIni - 1, liCols, CCur(liFilaIni - 1)
        End If
        xlHoja1.Cells(liFila, liCols) = Format(rs!nMonto, "#,###0.00")
        cTpoEfecAnt = cTpoEfec
        cTpoAgeAnt = cTpoAge
        rs.MoveNext
    Loop
    
    'EL ultimo Tramo de la última Agencia
    ExcelCuadro xlHoja1, liCols, liFilaIni, liCols, CCur(liFila + 1)
    liFila = liFila + 2
    xlHoja1.Cells(liFila, liCols).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(liFilaIni, liCols), xlHoja1.Cells(liFila - 2, liCols)).Address(False, False) & ")"
    ExcelCuadro xlHoja1, liCols, liFila, liCols, CCur(liFila)
    xlHoja1.Range(xlHoja1.Cells(liFila, liCols), xlHoja1.Cells(liFila, liCols)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liFila, liCols), xlHoja1.Cells(liFila, liCols)).EntireColumn.AutoFit
    xlHoja1.Range(xlHoja1.Cells(liFilaIni, liCols), xlHoja1.Cells(liFila, liCols)).NumberFormat = "#,##0.00;-#,##0.00"
    
    'PARA LOS TOTALES DESCOM
    liCols = liCols + 1
    xlHoja1.Cells(liFilaIni - 1, liCols) = "TOTALES"
    ExcelCuadro xlHoja1, liCols, liFilaIni - 1, liCols, CCur(liFilaIni - 1)
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 1, liCols), xlHoja1.Cells(liFilaIni - 1, liCols)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 1, liCols), xlHoja1.Cells(liFilaIni - 1, liCols)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 1, liCols), xlHoja1.Cells(liFilaIni - 1, liCols)).Interior.Color = 16764057
    
    ExcelCuadro xlHoja1, liCols, liFilaIni, liCols, CCur(liFila - 1)
    xlHoja1.Cells(liFila, liCols).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(liFilaIni, liCols), xlHoja1.Cells(liFila - 2, liCols)).Address(False, False) & ")"
    ExcelCuadro xlHoja1, liCols, liFila, liCols, CCur(liFila)
    xlHoja1.Range(xlHoja1.Cells(liFila, liCols), xlHoja1.Cells(liFila, liCols)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liFilaIni - 1, liCols), xlHoja1.Cells(liFilaIni - 1, liCols)).EntireColumn.AutoFit
    
    'PARTE DEL RESUMEN TOTAL PARA LA ULTIMA AGENCIA
    lsCadenaSum(2) = xlHoja1.Range(xlHoja1.Cells(liFila, liCols - 1), xlHoja1.Cells(liFila, liCols - 1)).Address(False, False)
    lsCadenaSum(3) = xlHoja1.Range(xlHoja1.Cells(liFila + 3, liCols - 1), xlHoja1.Cells(liFila + 3, liCols - 1)).Address(False, False)
    xlHoja1.Cells(liFila + 5, liCols - 1).Formula = "=" & "Sum" & "(" & lsCadenaSum(1) & "+" & lsCadenaSum(2) & "+" & lsCadenaSum(3) & ")"
    ExcelCuadro xlHoja1, liCols - 1, liFila + 5, liCols - 1, CCur(liFila + 5)
    '*******
    '***********Resumen Total Total
    xlHoja1.Range(xlHoja1.Cells(liFila, liCols), xlHoja1.Cells(liFila, liCols)).Font.Bold = True
    ExcelCuadro xlHoja1, liCols, liFila, liCols, CCur(liFila)
    xlHoja1.Range(xlHoja1.Cells(liFilaIni, liCols), xlHoja1.Cells(liFila, liCols)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(liFilaIni, liCols), xlHoja1.Cells(liFilaIni, liCols)).EntireColumn.AutoFit
    
    xlHoja1.Cells(liFila + 5, 2) = "RESUMEN TOTAL " & "(" & IIf(lsMoneda = "1", "MN", "ME") & ")"
    xlHoja1.Range(xlHoja1.Cells(liFila + 5, 2), xlHoja1.Cells(liFila + 5, 2)).Interior.Color = 16764057
    ExcelCuadro xlHoja1, 2, liFila + 5, 2, CCur(liFila + 5)
    xlHoja1.Range(xlHoja1.Cells(liFila + 5, 2), xlHoja1.Cells(liFila + 5, 2)).EntireColumn.AutoFit
    xlHoja1.Range(xlHoja1.Cells(liFila + 5, 2), xlHoja1.Cells(liFila + 5, nCantAge + 5)).Font.Bold = True 'Sombreo hasta el Final
    xlHoja1.Range(xlHoja1.Cells(liFila + 5, 3), xlHoja1.Cells(liFila + 5, nCantAge + 5)).NumberFormat = "#,##0.00;-#,##0.00"
    
    lsCadenaSum(2) = xlHoja1.Range(xlHoja1.Cells(liFila, liCols), xlHoja1.Cells(liFila, liCols)).Address(False, False)
    lsCadenaSum(3) = xlHoja1.Range(xlHoja1.Cells(liFila + 3, liCols), xlHoja1.Cells(liFila + 3, liCols)).Address(False, False)
    xlHoja1.Cells(liFila + 5, liCols).Formula = "=" & "Sum" & "(" & lsCadenaSum(4) & "+" & lsCadenaSum(2) & "+" & lsCadenaSum(3) & ")"
    ExcelCuadro xlHoja1, liCols, liFila + 5, liCols, CCur(liFila + 5)
    'El "4" referencia al Primer Total del Cuadro de Totales
    
    '******************** muestra de saldos del dia *******************************
    xlAplicacion.Range(xlHoja1.Cells(5, 2), xlHoja1.Cells(5, 3)).Font.Bold = True
    xlHoja1.Cells(5, 2) = "SALDO DEL DIA"
    xlHoja1.Range(xlHoja1.Cells(5, 3), xlHoja1.Cells(5, 3)).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(liFila + 5, liCols), xlHoja1.Cells(liFila + 5, liCols)).Address(False, False)
    xlHoja1.Range(xlHoja1.Cells(5, 3), xlHoja1.Cells(5, 3)).EntireColumn.AutoFit
    Set dBalan = New DbalanceCont
    Call dBalan.InsertaCtaContSaldoDiario("11" & lsMoneda & "1", pdFecha, psOpeCod, xlHoja1.Cells(5, 3))
    If lsMoneda = "1" Then
        nTotal110701 = dBalan.ObtenerCtaContSaldoBalanceDiario("11" & lsMoneda & "701", pdFecha, lsMoneda, 0)
    Else
        nTotal110701 = dBalan.ObtenerCtaContSaldoDiario("11" & lsMoneda & "701", pdFecha)
    End If
    Call dBalan.InsertaCtaContSaldoDiario("11" & lsMoneda & "701", pdFecha, psOpeCod, nTotal110701)
    '***********************
End If

oBarra.Progress 80, "DESCOMPOSICION DE EFECTIVO", "Generando Efectivo en Tránsito", , vbBlue
'EFECTIVO EN TRÁNSITO
Set rs = Nothing
liFila = liFila + 3
liFilaIni = liFila
liCols = 3
liColIni = liCols
Set rs = DEfec.GetCtaEfectivoTrans_Age(pdFecha, lsMoneda)
If Not (rs.EOF And rs.BOF) Then
    xlHoja1.Cells(liFila - 1, liCols - 1) = "Efectivo en Tránsito"
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, liCols - 1), xlHoja1.Cells(liFila - 1, liCols - 1)).Interior.Color = 16764057
    ExcelCuadro xlHoja1, liCols - 1, liFila - 1, liCols - 1, CCur(liFila - 1)
    xlHoja1.Cells(liFila, liCols - 1) = rs!TpoMoneda
    xlHoja1.Range(xlHoja1.Cells(liFila, liCols - 1), xlHoja1.Cells(liFila, liCols - 1)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(liFila, liCols - 1), xlHoja1.Cells(liFila, liCols - 1)).Font.Bold = True
    ExcelCuadro xlHoja1, liCols - 1, liFila, liCols - 1, CCur(liFila)
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, liCols - 1), xlHoja1.Cells(liFila - 1, liCols - 1)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, liCols - 1), xlHoja1.Cells(liFila - 1, liCols - 1)).Font.Bold = True
    Do While Not rs.EOF
        xlHoja1.Cells(liFila - 1, liCols) = rs!cAgeDescripcion
        xlHoja1.Range(xlHoja1.Cells(liFila - 1, liCols), xlHoja1.Cells(liFila - 1, liCols)).Interior.Color = 16764057
        ExcelCuadro xlHoja1, liCols, liFila - 1, liCols, CCur(liFila - 1)
        xlHoja1.Cells(liFila, liCols) = Format(rs!nSaldoEfectTrans, "#,###0.00")
        ExcelCuadro xlHoja1, liCols, liFila, liCols, CCur(liFila - 1)
        liCols = liCols + 1
        rs.MoveNext
    Loop
    xlHoja1.Cells(liFila - 1, liCols) = "TOTALES"
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, liCols), xlHoja1.Cells(liFila - 1, liCols)).Interior.Color = 16764057
    ExcelCuadro xlHoja1, liCols, liFila - 1, liCols, CCur(liFila - 1)
    xlHoja1.Cells(liFila, liCols).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(liFila, liColIni - 1), xlHoja1.Cells(liFila, liCols - 1)).Address(False, False) & ")"
    xlHoja1.Range(xlHoja1.Cells(liFila, liCols), xlHoja1.Cells(liFila, liCols)).Font.Bold = True
    ExcelCuadro xlHoja1, liCols, liFila, liCols, CCur(liFila)
    
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, liColIni), xlHoja1.Cells(liFila - 1, liCols)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(liFila - 1, liColIni), xlHoja1.Cells(liFila - 1, liCols)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liFila, liColIni - 1), xlHoja1.Cells(liFila, liCols)).NumberFormat = "#,##0.00;-#,##0.00"
End If

liFila = liFila + 6
liCols = 3
If (lsMoneda = "1") Then
    oBarra.Progress 90, "DESCOMPOSICION DE EFECTIVO", "Generando Indice de Cob. y Nivel de Caja", , vbBlue
    Call CargarIndiceCobNivelCaja(xlHoja1.Application, pdFecha, liFila, liCols, lnTipCambio, pscCtaContCod)     '***NAGL ERS 079-2017 20180123
End If
oBarra.Progress 100, "DESCOMPOSICION DE EFECTIVO", "Generando Efectivo en Tránsito", , vbBlue
oBarra.CloseForm frmReportes
Set oBarra = Nothing
End Sub 'NAGL 20180920 Según TIC1807210002

Private Sub CargarIndiceCobNivelCaja(ByVal xlAplicacion As Excel.Application, ByVal pdFecha As Date, ByVal liFila As Long, ByVal liCols As Long, ByVal lnTipCambio As Currency, ByVal pscCtaContCod As String)
Dim DEfec As New Defectivo
Dim rsNiv As New ADODB.Recordset
Set rsNiv = DEfec.ObtieneIndiceCobNivCaja(pdFecha, lnTipCambio, pscCtaContCod)
If Not (rsNiv.EOF And rsNiv.BOF) Then
    xlHoja1.Cells(liFila - 2, liCols - 1) = "INDICE DE COBERTURA"
    xlHoja1.Range(xlHoja1.Cells(liFila - 2, liCols - 1), xlHoja1.Cells(liFila - 1, liCols - 1)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(liFila - 2, liCols - 1), xlHoja1.Cells(liFila - 1, liCols - 1)).VerticalAlignment = xlCenter
    'xlHoja1.Range(xlHoja1.Cells(liFila - 2, liCols - 1), xlHoja1.Cells(liFila - 1, liCols - 1)).Interior.Color = 16764057
    ExcelCuadro xlHoja1, liCols - 1, liFila - 2, liCols - 1, CCur(liFila - 1)
    xlHoja1.Cells(liFila, liCols - 1) = "MN Y ME"
    xlHoja1.Range(xlHoja1.Cells(liFila, liCols - 1), xlHoja1.Cells(liFila, liCols - 1)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(liFila - 2, liCols - 1), xlHoja1.Cells(liFila - 1, liCols - 1)).Font.Bold = True
    ExcelCuadro xlHoja1, liCols - 1, liFila, liCols - 1, CCur(liFila)
    
    xlHoja1.Cells(liFila + 2, liCols - 1) = "NIVEL DE CAJA"
    xlHoja1.Range(xlHoja1.Cells(liFila + 2, liCols - 1), xlHoja1.Cells(liFila + 3, liCols - 1)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(liFila + 2, liCols - 1), xlHoja1.Cells(liFila + 3, liCols - 1)).VerticalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(liFila + 2, liCols - 1), xlHoja1.Cells(liFila + 3, liCols - 1)).HorizontalAlignment = xlCenter
    'xlHoja1.Range(xlHoja1.Cells(liFila + 2, liCols - 1), xlHoja1.Cells(liFila + 3, liCols - 1)).Interior.Color = 16764057
    ExcelCuadro xlHoja1, liCols - 1, liFila + 2, liCols - 1, CCur(liFila + 3)

    xlHoja1.Cells(liFila + 4, liCols - 1) = "TOTAL EN ME"
    xlHoja1.Cells(liFila + 5, liCols - 1) = "TOTAL EN MN"
    xlHoja1.Range(xlHoja1.Cells(liFila + 4, liCols - 1), xlHoja1.Cells(liFila + 5, liCols - 1)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(liFila + 2, liCols - 1), xlHoja1.Cells(liFila + 5, liCols - 1)).Font.Bold = True
    ExcelCuadro xlHoja1, liCols - 1, liFila + 4, liCols - 1, CCur(liFila + 5)

    Do While Not rsNiv.EOF
        '***Indice de Cobertura
        xlHoja1.Cells(liFila - 2, liCols) = rsNiv!cAgeDescripcion
        xlHoja1.Range(xlHoja1.Cells(liFila - 2, liCols), xlHoja1.Cells(liFila - 1, liCols)).MergeCells = True
        xlHoja1.Range(xlHoja1.Cells(liFila - 2, liCols), xlHoja1.Cells(liFila - 1, liCols)).VerticalAlignment = xlCenter
        'xlHoja1.Range(xlHoja1.Cells(liFila - 2, liCols), xlHoja1.Cells(liFila - 1, liCols)).Interior.Color = 16764057
        ExcelCuadro xlHoja1, liCols, liFila - 2, liCols, CCur(liFila - 1)
        xlHoja1.Range(xlHoja1.Cells(liFila - 2, liCols), xlHoja1.Cells(liFila - 1, liCols)).Font.Bold = True
        
        xlHoja1.Cells(liFila, liCols) = Format(rsNiv!lsIndiceCobertura, "#,###0.00")
        ExcelCuadro xlHoja1, liCols, liFila, liCols, CCur(liFila)
        '***End Indice de Cobertura
        
        '************Nivel de Caja*****************
        xlHoja1.Cells(liFila + 2, liCols) = rsNiv!cAgeDescripcion
        xlHoja1.Range(xlHoja1.Cells(liFila + 2, liCols), xlHoja1.Cells(liFila + 3, liCols)).MergeCells = True
        xlHoja1.Range(xlHoja1.Cells(liFila + 2, liCols), xlHoja1.Cells(liFila + 3, liCols)).VerticalAlignment = xlCenter
        'xlHoja1.Range(xlHoja1.Cells(liFila + 2, liCols), xlHoja1.Cells(liFila + 3, liCols)).Interior.Color = 16764057
        ExcelCuadro xlHoja1, liCols, liFila + 2, liCols, CCur(liFila + 3)
        xlHoja1.Range(xlHoja1.Cells(liFila + 2, liCols), xlHoja1.Cells(liFila + 3, liCols)).Font.Bold = True
        
        xlHoja1.Cells(liFila + 4, liCols) = Format(rsNiv!nMontoNivCajaMN, "#,###0.00")
        xlHoja1.Cells(liFila + 5, liCols) = Format(rsNiv!nMontoNivCajaME, "#,###0.00")
        ExcelCuadro xlHoja1, liCols, liFila + 4, liCols, CCur(liFila + 5)
        
        xlHoja1.Range(xlHoja1.Cells(liFila, liCols), xlHoja1.Cells(liFila, liCols)).NumberFormat = "#,##0.00;-#,##0.00"
        xlHoja1.Range(xlHoja1.Cells(liFila + 4, liCols), xlHoja1.Cells(liFila + 5, liCols)).NumberFormat = "#,##0.00;-#,##0.00"
        '******************************************
        liCols = liCols + 1
        rsNiv.MoveNext
    Loop
    End If
End Sub 'NAGL 20180920 Según TIC1807210002





'************METODOS ANTERIORES QUE INTERVENIAN EN LA DESCOMPOSICIÓN DE EFECTIVO
'
'Function BilletajeCajaGeneral(psOpeCod As String, psMoneda As String, psEfectivoCod As String, pdFecha As Date) As Currency
'Dim lsCtaCod As String
'Dim oOpe As New DOperacion
'lsCtaCod = oOpe.EmiteOpeCta(psOpeCod, "D", "0", psEfectivoCod)
'Set oOpe = Nothing
'lsCtaCod = Left(lsCtaCod, 2) & psMoneda & Mid(lsCtaCod, 4, 22)
'Dim oSdo As New NCtasaldo
'BilletajeCajaGeneral = oSdo.GetCtaEfectivoSaldo(lsCtaCod, Format(pdFecha, gsFormatoFecha), IIf(psMoneda = "1", True, False), psEfectivoCod)
'Set oSdo = Nothing
'End Function
'
'Private Sub ValoresBilletaje(psOpeCod As String, pdFecha As Date, lsMoneda As String)
'Dim sql As String
'Dim rs  As New ADODB.Recordset
'Dim rs1 As New ADODB.Recordset
'Dim i    As Integer
'Dim Col  As Integer
'Dim Fila As Integer
'
'Dim lnTotal As Long
'Dim N As Long
'
'Dim lsCodOPeDol As String
'Dim lsCodOpeSol As String
'
'If lsMoneda = "1" Then
'   lsCodOpeSol = psOpeCod
'Else
'   If Mid(psOpeCod, 3, 1) = "1" Then
'      lsCodOPeDol = Mid(psOpeCod, 1, 2) & "2" & Mid(psOpeCod, 4, Len(psOpeCod))
'      lsCodOpeSol = psOpeCod
'   Else
'      lsCodOPeDol = psOpeCod
'      lsCodOpeSol = Mid(psOpeCod, 1, 2) & "1" & Mid(psOpeCod, 4, Len(psOpeCod))
'   End If
'End If
'
'Dim oEfec As New Defectivo
'Dim oAge  As New DActualizaDatosArea
'Set oBarra = New clsProgressBar
'ProgressShow oBarra, frmReportes, eCap_CaptionPercent, True
'oBarra.Progress 0, "DESCOMPOSICION DE EFECTIVO", "Cargando Datos de Billetes", , vbBlue
'Set rs = oEfec.EmiteBilletajes(lsMoneda, "B")
'lnTotal = rs.RecordCount
'If lsMoneda = "1" Then
'    lsBilletesSoles(1, 0) = "BILLETES"
'    lsBilletesSoles(2, 0) = "CAJA GENERAL"
'Else
'    lsBilletesDolares(1, 0) = "BILLETES"
'    lsBilletesDolares(2, 0) = "CAJA GENERAL"
'End If
'Fila = 0
'oBarra.Max = rs.RecordCount
'Do While Not rs.EOF
'    oBarra.Progress rs.Bookmark, "DESCOMPOSICION DE EFECTIVO", "Generando Billetes de Caja General", , vbBlue
'    N = N + 1
'    Fila = Fila + 1
'    If lsMoneda = "1" Then
'        ReDim Preserve lsBilletesSoles(lnTotalColumnas, Fila)
'        lsBilletesSoles(1, Fila) = Trim(Replace(rs!Descripcion, "AGENCIA", ""))
'        lsBilletesSoles(2, Fila) = Format(BilletajeCajaGeneral(lsCodOpeSol, lsMoneda, rs!cEfectivoCod, pdFecha), gsFormatoNumeroView)
'    Else
'        ReDim Preserve lsBilletesDolares(lnTotalColumnas, Fila)
'        lsBilletesDolares(1, Fila) = Trim(Replace(rs!Descripcion, "AGENCIA", ""))
'        lsBilletesDolares(2, Fila) = Format(BilletajeCajaGeneral(lsCodOPeDol, lsMoneda, rs!cEfectivoCod, pdFecha), gsFormatoNumeroView)
'    End If
'    Col = 2
'
'    Set rs1 = oAge.GetAgencias(, False)
'    Do While Not rs1.EOF
'       oBarra.Progress rs.Bookmark, "DESCOMPOSICION DE EFECTIVO", "Generando Billetes de " & rs1!Descripcion, , vbBlue
'        Col = Col + 1
'        If lsMoneda = "1" Then
'            lsBilletesSoles(Col, 0) = Trim(Replace(rs1!Descripcion, "AGENCIA", ""))
'            'Centralizado
'            lsBilletesSoles(Col, Fila) = Format(oEfec.BilletajeCajaAgencias(rs!cEfectivoCod, lsMoneda, pdFecha, rs1!Codigo), gsFormatoNumeroView)
'
'            'Distribuido
'            'lsBilletesSoles(Col, Fila) = BilletajeCajaAgencias(rs!nEfectivoValor, lsMoneda, pdFecha, "B", gsCodCMAC & rs1!Codigo)
'        Else
'            lsBilletesDolares(Col, 0) = Trim(Replace(rs1!Descripcion, "AGENCIA", ""))
'            lsBilletesDolares(Col, Fila) = Format(oEfec.BilletajeCajaAgencias(rs!cEfectivoCod, lsMoneda, pdFecha, rs1!Codigo), gsFormatoNumeroView)
'            'lsBilletesDolares(Col, Fila) = BilletajeCajaAgencias(rs!nEfectivoValor, lsMoneda, pdFecha, "B", gsCodCMAC & rs1!Codigo)
'        End If
'        rs1.MoveNext
'    Loop
'    RSClose rs1
'    rs.MoveNext
'Loop
'RSClose rs
'ProgressClose oBarra, frmReportes, True
'ProgressShow oBarra, frmReportes, eCap_CaptionPercent, True
'oBarra.Progress 0, "DESCOMPOSICION DE EFECTIVO", "Cargando Datos de Monedas", , vbBlue
'Set rs = oEfec.EmiteBilletajes(lsMoneda, "M")
'Set oEfec = Nothing
'lnTotal = rs.RecordCount
'If lsMoneda = "1" Then
'   lsMonedasSoles(1, 0) = "MONEDAS"
'   lsMonedasSoles(2, 0) = "CAJA GENERAL"
'Else
'   lsMonedaDolares(1, 0) = "MONEDAS"
'   lsMonedaDolares(2, 0) = "CAJA GENERAL"
'End If
'Fila = 0
'N = 0
'oBarra.Max = rs.RecordCount
'Do While Not rs.EOF
'   oBarra.Progress rs.Bookmark, "DESCOMPOSICION DE EFECTIVO", "Generando Monedas de Caja General", , vbBlue
'   N = N + 1
'   Fila = Fila + 1
'   If lsMoneda = "1" Then
'      ReDim Preserve lsMonedasSoles(lnTotalColumnas, Fila)
'      lsMonedasSoles(1, Fila) = Trim(Replace(rs!Descripcion, "AGENCIA", ""))
'      lsMonedasSoles(2, Fila) = Format(BilletajeCajaGeneral(lsCodOpeSol, lsMoneda, rs!cEfectivoCod, pdFecha), gsFormatoNumeroView)
'   Else
'      ReDim Preserve lsMonedaDolares(lnTotalColumnas, Fila)
'      lsMonedaDolares(1, Fila) = Trim(Replace(rs!Descripcion, "AGENCIA", ""))
'      lsMonedaDolares(2, Fila) = Format(BilletajeCajaGeneral(lsCodOpeSol, lsMoneda, rs!cEfectivoCod, pdFecha), gsFormatoNumeroView)
'   End If
'   Col = 2
'
'   Set rs1 = oAge.GetAgencias(, False)
'   Do While Not rs1.EOF
'      oBarra.Progress rs.Bookmark, "DESCOMPOSICION DE EFECTIVO", "Generando Monedas de " & rs1!Descripcion, , vbBlue
'      Col = Col + 1
'      If lsMoneda = "1" Then
'         lsMonedasSoles(Col, 0) = Trim(Replace(rs1!Descripcion, "AGENCIA", ""))
'         lsMonedasSoles(Col, Fila) = Format(oEfec.BilletajeCajaAgencias(rs!cEfectivoCod, lsMoneda, pdFecha, rs1!Codigo), gsFormatoNumeroView)
'         'lsMonedasSoles(Col, Fila) = BilletajeCajaAgencias(rs!nEfectivoValor, lsMoneda, pdFecha, "M", gsCodCMAC & rs1!Codigo)
'      Else
'         lsMonedaDolares(Col, 0) = Trim(Replace(rs1!Descripcion, "AGENCIA", ""))
'         lsMonedaDolares(Col, Fila) = Format(oEfec.BilletajeCajaAgencias(rs!cEfectivoCod, lsMoneda, pdFecha, rs1!Codigo), gsFormatoNumeroView)
'         'lsMonedaDolares(Col, Fila) = BilletajeCajaAgencias(rs!nEfectivoValor, lsMoneda, pdFecha, "M", gsCodCMAC & rs1!Codigo)
'      End If
'      rs1.MoveNext
'   Loop
'   RSClose rs1
'   rs.MoveNext
'Loop
'ProgressClose oBarra, frmReportes, True
'oBarra.CloseForm Me
'
''ProgressClose oBarra, Me
'Set oBarra = Nothing
'Set oAge = Nothing
'Set oEfec = Nothing
'End Sub
'Private Sub TotalesMatrices()
'Dim i As Integer
'Dim j As Integer
'
''******************** Totales de billetes soles ***********************
'For j = 2 To UBound(lsBilletesSoles, 1) 'COLUMNAS
'    For i = 1 To UBound(lsBilletesSoles, 2) 'FILAS
'        lsTotBillSol(j - 1) = Format(CCur(IIf(lsTotBillSol(j - 1) = "", "0", lsTotBillSol(j - 1))) + CCur(IIf(lsBilletesSoles(j, i) = "", "0", lsBilletesSoles(j, i))), gsFormatoNumeroView)
'    Next i
'Next j
'' ********************** Totales de Monedas Soles ************************
'For j = 2 To UBound(lsMonedasSoles, 1) 'COLUMNA
'    For i = 1 To UBound(lsMonedasSoles, 2) 'FILAS
'        lsTotMonSol(j - 1) = Format(CCur(IIf(lsTotMonSol(j - 1) = "", "0", lsTotMonSol(j - 1))) + CCur(IIf(lsMonedasSoles(j, i) = "", "0", lsMonedasSoles(j, i))), gsFormatoNumeroView)
'    Next
'Next
''******************** Totales de billetes dolares ***********************
'For j = 2 To UBound(lsBilletesDolares, 1) 'COLUMNAS
'    For i = 1 To UBound(lsBilletesDolares, 2) 'FILAS
'        lsTotBillDol(j - 1) = Format(CCur(IIf(lsTotBillDol(j - 1) = "", "0", lsTotBillDol(j - 1))) + CCur(IIf(lsBilletesDolares(j, i) = "", "0", lsBilletesDolares(j, i))), gsFormatoNumeroView)
'    Next
'Next
''******************** Totales de monedas dolares***********************
'For j = 2 To UBound(lsMonedaDolares, 1)
'    For i = 1 To UBound(lsMonedaDolares, 2) 'FILAS
'        lsTotMonDol(j - 1) = Format(CCur(IIf(lsTotMonDol(j - 1) = "", "0", lsTotMonDol(j - 1))) + CCur(IIf(lsMonedaDolares(j, i) = "", "0", lsMonedaDolares(j, i))), gsFormatoNumeroView)
'    Next
'Next
'
''***************** TOTALES SOLES ****************************
'For i = 1 To UBound(lsTotBillSol)
'    lsTotalSol(i) = Format(CCur(IIf(lsTotalSol(i) = "", "0", lsTotalSol(i))) + CCur(IIf(lsTotBillSol(i) = "", "0", lsTotBillSol(i))), gsFormatoNumeroView)
'Next i
'For i = 1 To UBound(lsTotMonSol) - 1
'    lsTotalSol(i) = CCur(IIf(lsTotalSol(i) = "", "0", lsTotalSol(i))) + CCur(IIf(lsTotMonSol(i) = "", "0", lsTotMonSol(i)))
'    lsTotalSolesTC(i) = Format(Round(CCur(IIf(lsTotalSol(i) = "", "0", lsTotalSol(i))) / lnTipCambio, 2), gsFormatoNumeroView)
'Next i
'
''****************** TOTALES DOLARES ***************************
'For i = 1 To UBound(lsTotBillDol)
'    lsTotalDol(i) = Format(CCur(IIf(lsTotalDol(i) = "", "0", lsTotalDol(i))) + CCur(IIf(lsTotBillDol(i) = "", "0", lsTotBillDol(i))), gsFormatoNumeroView)
'Next i
'For i = 1 To UBound(lsTotMonDol)
'    lsTotalDol(i) = Format(CCur(IIf(lsTotalDol(i) = "", "0", lsTotalDol(i))) + CCur(IIf(lsTotMonDol(i) = "", "0", lsTotMonDol(i))), gsFormatoNumeroView)
'Next i
'End Sub
'
'Private Sub MatrizCoberturas()
'Dim sql As String
'Dim rs As New ADODB.Recordset
'Dim i As Integer
'Dim oAge As New DActualizaDatosArea
'Set rs = oAge.GetAreaAgenciasImporte()
'i = 1
'Do While Not rs.EOF
'    If rs!cAreaCod = "012" Then
'        lsCobertura(1) = rs!nCobertura
'    Else
'        i = i + 1
'        lsCobertura(i) = rs!nCobertura
'    End If
'    rs.MoveNext
'Loop
'RSClose rs
'End Sub
'
'Private Sub TotalOro(pdFecha As Date, lsMoneda As String)
'Dim sql As String
'Dim rs As New ADODB.Recordset
'Dim lnTotalOro As Currency
'Dim lnTotalOro01 As Currency
'Dim i As Integer
'
'Dim oVar As New NConstSistemas
'Dim nValorOro As Currency
'nValorOro = CCur(oVar.LeeConstSistema(gConstSistValorOroDolares))
'
'Dim oAge As New DActualizaDatosArea
'Set rs = oAge.GetAgencias(, False)
'Set oAge = Nothing
'i = 1
'Do While Not rs.EOF
'    i = i + 1
'    lnTotalOro = OroAgencia(pdFecha, lsMoneda, rs!Codigo)
'    lsTotalOro(i) = lnTotalOro * nValorOro
'    'Se veririca si la agencia es sede entonce el oro se
'    'asigna al item 1 que es de caja puesto que los posee caja general y
'    'no la agencia sede por sì misma
'    If CInt(rs!Codigo) = 7 Then
'        lsTotalOro(1) = CCur(lsTotalOro(i))
'        lsTotalOro(i) = "0"
'    End If
'    rs.MoveNext
'Loop
'RSClose rs
'End Sub
'
'Private Sub GeneraDesEfectivo(psOpeCod As String, psMoneda As String, pdFecha As Date)
'Dim i As Integer
'InicializaMatrices
'If psMoneda = "1" Then
'   ValoresBilletaje psOpeCod, pdFecha, psMoneda
'End If
'ValoresBilletaje psOpeCod, pdFecha, "2"
'TotalesMatrices
'MatrizCoberturas
'TotalOro pdFecha, psMoneda
'
'For i = 1 To UBound(lsCobertura)
'    If CCur(lsTotalOro(i)) > 0 Then
'        lsNivCobEfect(i) = Format(nVal(lsCobertura(i)) - CCur(lsTotalOro(i)), "#,#0.00")
'    Else
'        lsNivCobEfect(i) = "0"
'    End If
'Next i
'
'For i = 1 To UBound(lsCobertura)
'    lsCobEfect(i) = Format(CCur(lsTotalSolesTC(i)) + CCur(lsTotalDol(i)), "#,#0.00")
'Next
'
'For i = 1 To UBound(lsCobertura)
'    lsCoberturado(i) = Format(CCur(lsTotalOro(i)) + CCur(lsTotalDol(i)) + CCur(lsTotalSolesTC(i)), "#,#0.00")
'Next i
'For i = 1 To UBound(lsNivCobEfect)
'    If Val(lsNivCobEfect(i)) > 0 Then
'        lsIndiceCobertura(i) = Format(CCur(lsNivCobEfect(i)) / CCur(lsCobEfect(i)), "#,#0.00")
'    Else
'        If Val(lsCoberturado(i)) > 0 Then
'            lsIndiceCobertura(i) = Format(nVal(lsCobertura(i)) / CCur(lsCoberturado(i)), "#,#0.00")
'        Else
'            lsIndiceCobertura(i) = 0
'        End If
'    End If
'Next i
'
'End Sub
'Public Sub ImprimeEfectivoCaja(psOpeCod As String, psMoneda As String, pdFecha As Date)
'Dim lbHojaActiva As Boolean
'Dim oTC As nTipoCambio
'
''On Error GoTo ImprimeEfectivoCajaErr
'
'Set oTC = New nTipoCambio
'   lnTipCambio = oTC.EmiteTipoCambio(pdFecha, TCFijoMes)
'   GeneraDesEfectivo psOpeCod, psMoneda, pdFecha
'   lsArchivo = App.path & "\SPOOLER\" & "Anx15A_Efectivo_" & Format(pdFecha, "mmyyyy") & IIf(psMoneda = "1", "MN", "ME") & ".XLS"
'   lbHojaActiva = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
'   If lbHojaActiva Then
'      ExcelAddHoja Format(pdFecha, "dd-mm-yyyy"), xlLibro, xlHoja1
'      GeneraReporte psOpeCod, psMoneda, pdFecha
'      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
'      If lsArchivo <> "" Then
'          CargaArchivo lsArchivo, App.path & "\SPOOLER\"
'      End If
'   End If
'Exit Sub
'ImprimeEfectivoCajaErr:
'   MsgBox TextErr(Err.Description), vbInformation, "Aviso"
'   If lbHojaActiva Then
'      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
'   End If
'   lbHojaActiva = False
'End Sub
'Private Sub GeneraReporte(psOpeCod As String, lsMoneda As String, pdFecha As Date)
'   Dim lnFila As Integer, lnCol As Integer
'   Dim i      As Integer, j     As Integer, m As Integer
'   Dim lnFilaTotales As Integer
'   Dim lnColTotales  As Integer
'
'   Dim lsTotalesFilasBill() As String
'   Dim lsTotalesColBill()   As String
'
'   Dim lnSaldoDia As Currency
'   Dim lsTotalGeneral()    As String
'   Dim lsTotalesFilasMon() As String
'   Dim lsTotalesColMon()   As String
'
'   Dim Y1 As Currency, Y2 As Currency
'   Dim lbExisteHoja  As Boolean
'   Dim lsTotales()   As String
'   Dim nTotal111701 As Currency
'   Dim dBalan As DbalanceCont
'   xlHoja1.PageSetup.Zoom = 75
'   xlHoja1.PageSetup.Orientation = xlLandscape
'
'   xlHoja1.Range("A1:R100").Font.Size = 8
'   xlAplicacion.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(4, 50)).Font.Bold = True
'
'   xlHoja1.Range("A1").ColumnWidth = 5
'   xlHoja1.Range("B1:P1").ColumnWidth = 13
'
'   xlHoja1.Cells(1, 2) = gsNomCmac
'   xlHoja1.Cells(1, 12) = "Fecha :" & Format(pdFecha, "dd mmmm yyyy")
'   xlHoja1.Cells(3, 6) = "DESCOMPOSICION DE EFECTIVO " & IIf(lsMoneda = "1", "MONEDA NACIONAL", "MONEDA EXTRANJERA")
'   xlHoja1.Cells(4, 7) = "Area de Caja General"
'   lnFila = 8
'   If lsMoneda = "1" Then
'      '********************************* ****************************************************
'      '********************************* SOLES **********************************************
'      '**************************************************************************************
'
'      '******************************* BILLETES *********************************************
'      ReDim lsTotalesFilasBill(UBound(lsBilletesSoles, 2))
'      ReDim lsTotalesColBill(UBound(lsBilletesSoles, 1) + 1)
'      For i = 0 To UBound(lsBilletesSoles, 2) 'filas
'          lnFila = lnFila + 1
'          If i = 0 Then
'              lnFilaTotales = lnFila
'              Y1 = lnFila
'          End If
'          lnCol = 1
'          ReDim Preserve lsTotalesFilasBill(i)
'          For j = 1 To UBound(lsBilletesSoles, 1) 'columnas
'              lnCol = lnCol + 1
'               xlHoja1.Cells(lnFila, lnCol) = lsBilletesSoles(j, i)
'               If i > 0 Then
'                  If j > 1 Then
'                      lsTotalesFilasBill(i) = lsTotalesFilasBill(i) + xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'                      lsTotalesColBill(j) = lsTotalesColBill(j) + xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'                  Else
'                      lsTotalesColBill(j) = "TOTALES BILLETES"
'                  End If
'              Else
'                  xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Borders.LineStyle = xlContinuous
'                  xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'                  xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).HorizontalAlignment = xlCenter
'                  lsTotalesFilasBill(i) = "TOTALES"
'              End If
'          Next j
'      Next i
'
'      lnColTotales = lnCol + 1
'      '******************* Columna de Totales de billetes ******************************************
'
'      For i = 0 To UBound(lsTotalesFilasBill)
'          If i = 0 Then
'              xlAplicacion.Range(xlHoja1.Cells(lnFilaTotales, lnColTotales), xlHoja1.Cells(lnFilaTotales, lnColTotales)).Borders.LineStyle = xlContinuous
'              xlAplicacion.Range(xlHoja1.Cells(lnFilaTotales, lnColTotales), xlHoja1.Cells(lnFilaTotales, lnColTotales)).Font.Bold = True
'              xlAplicacion.Range(xlHoja1.Cells(lnFilaTotales, lnColTotales), xlHoja1.Cells(lnFilaTotales, lnColTotales)).HorizontalAlignment = xlCenter
'
'              xlHoja1.Cells(lnFilaTotales, lnColTotales) = lsTotalesFilasBill(i)
'          Else
'              xlHoja1.Range(xlHoja1.Cells(lnFilaTotales, lnColTotales), xlHoja1.Cells(lnFilaTotales, lnColTotales)).Formula = "=Sum(" & Mid(lsTotalesFilasBill(i), 1, Len(lsTotalesFilasBill(i)) - 1) & ")"
'              lsTotalesColBill(UBound(lsTotalesColBill)) = lsTotalesColBill(UBound(lsTotalesColBill)) + xlHoja1.Range(xlHoja1.Cells(lnFilaTotales, lnColTotales), xlHoja1.Cells(lnFilaTotales, lnColTotales)).Address(False, False) & "+"
'          End If
'          lnFilaTotales = lnFilaTotales + 1
'      Next
'
'      Y2 = lnFila + 1
'      ExcelCuadro xlHoja1, 2, Y1, CCur(lnColTotales), Y2
'
'      '****************** SUBTOTALES BILLETES  *****************************
'      lnFila = lnFila + 2
'      lnCol = 1
'      ReDim lsTotalGeneral(UBound(lsTotalesColBill))
'      For i = 1 To UBound(lsTotalesColBill)
'          lnCol = lnCol + 1
'          If i = 1 Then
'              xlHoja1.Cells(lnFila, lnCol) = lsTotalesColBill(i)
'              lsTotalGeneral(i) = "RESUMEN TOTAL"
'          Else
'              xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Formula = "=Sum(" & Mid(lsTotalesColBill(i), 1, Len(lsTotalesColBill(i)) - 1) & ")"
'              lsTotalGeneral(i) = xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'          End If
'          xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Borders.LineStyle = xlContinuous
'          xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'      Next
'
'      '*********************************** MONEDAS *******************************************
'      lnFila = lnFila + 1
'
'      ReDim lsTotalesFilasMon(UBound(lsMonedasSoles, 2))
'      ReDim lsTotalesColMon(UBound(lsMonedasSoles, 1) + 1)
'
'      For i = 0 To UBound(lsMonedasSoles, 2) 'filas
'          lnFila = lnFila + 1
'          If i = 0 Then
'              lnFilaTotales = lnFila
'              Y1 = lnFila
'          End If
'          lnCol = 1
'          For j = 1 To UBound(lsMonedasSoles, 1) 'columnas
'              lnCol = lnCol + 1
'               xlHoja1.Cells(lnFila, lnCol) = lsMonedasSoles(j, i)
'               If i > 0 Then
'                  If j > 1 Then
'                      lsTotalesFilasMon(i) = lsTotalesFilasMon(i) + xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'                      lsTotalesColMon(j) = lsTotalesColMon(j) + xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'                  Else
'                      lsTotalesColMon(j) = "TOTALES MONEDAS"
'                  End If
'              Else
'                  xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Borders.LineStyle = xlContinuous
'                  xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'                  xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).HorizontalAlignment = xlCenter
'                  lsTotalesFilasMon(i) = "TOTALES"
'              End If
'          Next j
'      Next i
'
'      '******************* Columna de Totales de MONEDAS ******************************************
'      lnColTotales = lnCol + 1
'
'      For i = 0 To UBound(lsTotalesFilasMon)
'          If i = 0 Then
'              xlAplicacion.Range(xlHoja1.Cells(lnFilaTotales, lnColTotales), xlHoja1.Cells(lnFilaTotales, lnColTotales)).Borders.LineStyle = xlContinuous
'              xlAplicacion.Range(xlHoja1.Cells(lnFilaTotales, lnColTotales), xlHoja1.Cells(lnFilaTotales, lnColTotales)).Font.Bold = True
'              xlAplicacion.Range(xlHoja1.Cells(lnFilaTotales, lnColTotales), xlHoja1.Cells(lnFilaTotales, lnColTotales)).HorizontalAlignment = xlCenter
'
'              xlHoja1.Cells(lnFilaTotales, lnColTotales) = lsTotalesFilasMon(i)
'          Else
'              xlHoja1.Range(xlHoja1.Cells(lnFilaTotales, lnColTotales), xlHoja1.Cells(lnFilaTotales, lnColTotales)).Formula = "=Sum(" & Mid(lsTotalesFilasMon(i), 1, Len(lsTotalesFilasMon(i)) - 1) & ")"
'              lsTotalesColMon(UBound(lsTotalesColMon)) = lsTotalesColMon(UBound(lsTotalesColMon)) + xlHoja1.Range(xlHoja1.Cells(lnFilaTotales, lnColTotales), xlHoja1.Cells(lnFilaTotales, lnColTotales)).Address(False, False) & "+"
'          End If
'          lnFilaTotales = lnFilaTotales + 1
'      Next
'
'      Y2 = lnFila + 1
'      ExcelCuadro xlHoja1, 2, Y1, CCur(lnColTotales), Y2
'
'      '******************     SUBTOTALES MONEDAS **********************
'      lnFila = lnFila + 2
'      lnCol = 1
'      For i = 1 To UBound(lsTotalesColMon)
'          lnCol = lnCol + 1
'          If i = 1 Then
'              xlHoja1.Cells(lnFila, lnCol) = lsTotalesColMon(i)
'              lsTotalGeneral(i) = "RESUMEN TOTAL " & IIf(Mid(psOpeCod, 3, 1) = "1", gcPEN_SIMBOLO, "$") 'marg ers044-2016
'          Else
'              xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Formula = "=Sum(" & Mid(lsTotalesColMon(i), 1, Len(lsTotalesColMon(i)) - 1) & ")"
'              lsTotalGeneral(i) = lsTotalGeneral(i) + xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False)
'          End If
'          xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'          xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Borders.LineStyle = xlContinuous
'      Next
'
'      '***********************  RESUMEN TOTAL ***************************************************************
'      lnFila = lnFila + 2
'      lnCol = 1
'      For i = 1 To UBound(lsTotalGeneral)
'          lnCol = lnCol + 1
'          If i = 1 Then
'              xlHoja1.Cells(lnFila, lnCol) = lsTotalGeneral(i)
'          Else
'              xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Formula = "=Sum(" & Mid(lsTotalGeneral(i), 1, Len(lsTotalGeneral(i))) & ")"
'          End If
'          xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Borders.LineStyle = xlContinuous
'          xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'      Next
'      '******************** muestra de saldos del dia *******************************
'      xlAplicacion.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 3)).Font.Bold = True
'      xlHoja1.Cells(6, 2) = "SALDO DEL DIA"
'      xlHoja1.Range(xlHoja1.Cells(6, 3), xlHoja1.Cells(6, 3)).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False)
'      Set dBalan = New DbalanceCont
'      Call dBalan.InsertaCtaContSaldoDiario("11" & lsMoneda & "1", pdFecha, psOpeCod, xlHoja1.Cells(lnFila, lnCol))
'      nTotal111701 = dBalan.ObtenerCtaContSaldoBalanceDiario("11" & lsMoneda & "701", pdFecha, lsMoneda, 0)
'      Call dBalan.InsertaCtaContSaldoDiario("11" & lsMoneda & "701", pdFecha, psOpeCod, nTotal111701)
'      '**************** ENCABEZADO DE INDICES DE COBERTURAS *************************
'      lnFila = lnFila + 2
'      lnCol = 2
'      Y1 = lnFila
'      xlHoja1.Cells(lnFila, lnCol) = "INDICE DE "
'      xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'      xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).HorizontalAlignment = xlCenter
'
'      xlHoja1.Cells(lnFila + 1, lnCol) = "COBERTURA"
'      xlAplicacion.Range(xlHoja1.Cells(lnFila + 1, lnCol), xlHoja1.Cells(lnFila + 1, lnCol)).Font.Bold = True
'      xlAplicacion.Range(xlHoja1.Cells(lnFila + 1, lnCol), xlHoja1.Cells(lnFila + 1, lnCol)).HorizontalAlignment = xlCenter
'
'      Y2 = lnFila + 1
'
'      For i = 2 To UBound(lsBilletesSoles, 1)
'          lnCol = lnCol + 1
'
'          xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'          xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).HorizontalAlignment = xlCenter
'
'          xlHoja1.Cells(lnFila, lnCol) = lsBilletesSoles(i, 0)
'      Next
'
'      ExcelCuadro xlHoja1, 2, Y1, CCur(lnCol), Y2
'
'      lnFila = lnFila + 2
'
'      Y1 = lnFila
'      lnCol = 2
'      xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'      xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).HorizontalAlignment = xlCenter
'      xlHoja1.Cells(lnFila, lnCol) = "MN Y ME"
'      For i = 1 To UBound(lsIndiceCobertura)
'          lnCol = lnCol + 1
'          xlHoja1.Cells(lnFila, lnCol) = lsIndiceCobertura(i)
'      Next
'      Y2 = lnFila
'      ExcelCuadro xlHoja1, 2, Y1, CCur(lnCol), Y2
'
'      '**************** ENCABEZADO DE INDICES DE NIVEL DE CAJA*************************
'      lnFila = lnFila + 2
'      lnCol = 2
'      Y1 = lnFila
'      xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'      xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).HorizontalAlignment = xlCenter
'
'      xlHoja1.Cells(lnFila, lnCol) = "NIVEL DE "
'
'      xlAplicacion.Range(xlHoja1.Cells(lnFila + 1, lnCol), xlHoja1.Cells(lnFila + 1, lnCol)).Font.Bold = True
'      xlAplicacion.Range(xlHoja1.Cells(lnFila + 1, lnCol), xlHoja1.Cells(lnFila + 1, lnCol)).HorizontalAlignment = xlCenter
'
'      xlHoja1.Cells(lnFila + 1, lnCol) = "CAJA"
'
'      For i = 2 To UBound(lsBilletesSoles, 1)
'          lnCol = lnCol + 1
'
'          xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'          xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).HorizontalAlignment = xlCenter
'
'          xlHoja1.Cells(lnFila, lnCol) = lsBilletesSoles(i, 0)
'      Next
'      Y2 = lnFila + 1
'      ExcelCuadro xlHoja1, 2, Y1, CCur(lnCol), Y2
'
'      lnFila = lnFila + 2
'      lnCol = 2
'      Y1 = lnFila
'      xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'      xlHoja1.Cells(lnFila, lnCol) = "TOTAL EN ME"
'      For i = 1 To UBound(lsCobEfect)
'          lnCol = lnCol + 1
'          xlHoja1.Cells(lnFila, lnCol) = lsCobEfect(i)
'      Next
'
'      lnFila = lnFila + 1
'      lnCol = 2
'      xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'      xlHoja1.Cells(lnFila, lnCol) = "TOTAL EN MN"
'      For i = 1 To UBound(lsCobEfect)
'          lnCol = lnCol + 1
'          xlHoja1.Cells(lnFila, lnCol) = Format(Round(CCur(lsCobEfect(i)) * lnTipCambio, 2), gsFormatoNumeroView)
'      Next
'      Y2 = lnFila
'      ExcelCuadro xlHoja1, 2, Y1, CCur(lnCol), Y2
'   Else
'      '********************************* ****************************************************
'      '********************************* DOLARES ********************************************
'      '**************************************************************************************
'
'
'      '********************* BILLETES *****************************
'      lnFila = 8
'      ReDim lsTotalesFilasBill(UBound(lsBilletesDolares, 2))
'      ReDim lsTotalesColBill(UBound(lsBilletesDolares, 1) + 1)
'
'      For i = 0 To UBound(lsBilletesDolares, 2) 'filas
'          lnFila = lnFila + 1
'          If i = 0 Then
'              lnFilaTotales = lnFila
'              Y1 = lnFila
'          End If
'          lnCol = 1
'          ReDim Preserve lsTotalesFilasBill(i)
'          For j = 1 To UBound(lsBilletesDolares, 1) 'columnas
'              lnCol = lnCol + 1
'               xlHoja1.Cells(lnFila, lnCol) = lsBilletesDolares(j, i)
'               If i > 0 Then
'                  If j > 1 Then
'                     lsTotalesFilasBill(i) = lsTotalesFilasBill(i) + xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'                     lsTotalesColBill(j) = lsTotalesColBill(j) + xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'                  Else
'                     lsTotalesColBill(j) = "TOTALES BILLETES"
'                  End If
'              Else
'                  xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Borders.LineStyle = xlContinuous
'                  xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'                  xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).HorizontalAlignment = xlCenter
'                  lsTotalesFilasBill(i) = "TOTALES"
'              End If
'          Next j
'      Next i
'
'      lnColTotales = lnCol + 1
'      '******************* Columna de Totales de billetes ******************************************
'
'      For i = 0 To UBound(lsTotalesFilasBill)
'          If i = 0 Then
'              xlAplicacion.Range(xlHoja1.Cells(lnFilaTotales, lnColTotales), xlHoja1.Cells(lnFilaTotales, lnColTotales)).Borders.LineStyle = xlContinuous
'              xlAplicacion.Range(xlHoja1.Cells(lnFilaTotales, lnColTotales), xlHoja1.Cells(lnFilaTotales, lnColTotales)).Font.Bold = True
'              xlAplicacion.Range(xlHoja1.Cells(lnFilaTotales, lnColTotales), xlHoja1.Cells(lnFilaTotales, lnColTotales)).HorizontalAlignment = xlCenter
'
'              xlHoja1.Cells(lnFilaTotales, lnColTotales) = lsTotalesFilasBill(i)
'          Else
'              xlHoja1.Range(xlHoja1.Cells(lnFilaTotales, lnColTotales), xlHoja1.Cells(lnFilaTotales, lnColTotales)).Formula = "=Sum(" & Mid(lsTotalesFilasBill(i), 1, Len(lsTotalesFilasBill(i)) - 1) & ")"
'              lsTotalesColBill(UBound(lsTotalesColBill)) = lsTotalesColBill(UBound(lsTotalesColBill)) + xlHoja1.Range(xlHoja1.Cells(lnFilaTotales, lnColTotales), xlHoja1.Cells(lnFilaTotales, lnColTotales)).Address(False, False) & "+"
'          End If
'          lnFilaTotales = lnFilaTotales + 1
'      Next
'
'      Y2 = lnFila + 1
'      ExcelCuadro xlHoja1, 2, Y1, CCur(lnColTotales), Y2
'
'      '****************** SUBTOTALES BILLETES  *****************************
'      lnFila = lnFila + 2
'      lnCol = 1
'      ReDim lsTotalGeneral(UBound(lsTotalesColBill))
'      For i = 1 To UBound(lsTotalesColBill)
'          lnCol = lnCol + 1
'          If i = 1 Then
'              xlHoja1.Cells(lnFila, lnCol) = lsTotalesColBill(i)
'              lsTotalGeneral(i) = "RESUMEN TOTAL"
'          Else
'              xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Formula = "=Sum(" & Mid(lsTotalesColBill(i), 1, Len(lsTotalesColBill(i)) - 1) & ")"
'              lsTotalGeneral(i) = xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False) & "+"
'          End If
'          xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Borders.LineStyle = xlContinuous
'          xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'      Next
'
'      '***********************  RESUMEN TOTAL ***************************************************************
'      lnFila = lnFila + 2
'      lnCol = 1
'      For i = 1 To UBound(lsTotalGeneral)
'          lnCol = lnCol + 1
'          If i = 1 Then
'              xlHoja1.Cells(lnFila, lnCol) = lsTotalGeneral(i)
'          Else
'              xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Formula = "=Sum(" & Mid(lsTotalGeneral(i), 1, Len(lsTotalGeneral(i)) - 1) & ")"
'          End If
'          xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Borders.LineStyle = xlContinuous
'          xlAplicacion.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Font.Bold = True
'      Next
'      '******************** muestra de saldos del dia *******************************
'      xlAplicacion.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 3)).Font.Bold = True
'      xlHoja1.Cells(6, 2) = "SALDO DEL DIA"
'      xlHoja1.Range(xlHoja1.Cells(6, 3), xlHoja1.Cells(6, 3)).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(lnFila, lnCol), xlHoja1.Cells(lnFila, lnCol)).Address(False, False)
'
'      Set dBalan = New DbalanceCont
'      Call dBalan.InsertaCtaContSaldoDiario("11" & lsMoneda & "1", pdFecha, psOpeCod, xlHoja1.Cells(lnFila, lnCol))
'      nTotal111701 = dBalan.ObtenerCtaContSaldoDiario("11" & lsMoneda & "701", pdFecha)
'      Call dBalan.InsertaCtaContSaldoDiario("11" & lsMoneda & "701", pdFecha, psOpeCod, nTotal111701)
'      '***************************************************************************
'   End If
'End Sub
'
'Private Function OroAgencia(pdFecha As Date, lsMoneda As String, lsCodAge As String) As Currency
'Dim sql As String
'Dim rs As New ADODB.Recordset
'Dim oEst As New NEstadisticas
'OroAgencia = 0
'Set rs = oEst.GetEstadisticaPrendario(gbBitCentral, pdFecha, lsMoneda, lsCodAge)
'Do While Not rs.EOF
'   OroAgencia = OroAgencia + rs!nOroVig + rs!nOroAdj
'   rs.MoveNext
'Loop
'RSClose rs
'Set oEst = Nothing
'End Function
'
'
'Private Sub InicializaMatrices()
'Dim rs As New ADODB.Recordset
'   Dim oAge As New DActualizaDatosArea
'   lnTotalColumnas = 0
'   Set rs = oAge.GetAgencias(, False)
'   lnTotalColumnas = rs.RecordCount + 2
'   RSClose rs
'
'    ReDim lsBilletesSoles(lnTotalColumnas, 0)
'    ReDim lsMonedasSoles(lnTotalColumnas, 0)
'    ReDim lsBilletesDolares(lnTotalColumnas, 0)
'    ReDim lsMonedaDolares(lnTotalColumnas, 0)
'
'    ReDim lsTotBillSol(lnTotalColumnas)
'    ReDim lsTotBillDol(lnTotalColumnas)
'    ReDim lsTotMonSol(lnTotalColumnas)
'    ReDim lsTotMonDol(lnTotalColumnas)
'    ReDim lsTotalDol(lnTotalColumnas)
'    ReDim lsTotalSol(lnTotalColumnas)
'    ReDim lsTotalSolesTC(lnTotalColumnas)
'
'    'MATRICES PARA EL CALCULO DE COBERTURA
'    ReDim lsIndiceCob(lnTotalColumnas - 1)
'
'    ReDim lsNivCobEfect(lnTotalColumnas - 1)
'    ReDim lsCobEfect(lnTotalColumnas - 1)
'    ReDim lsCoberturado(lnTotalColumnas - 1)
'    ReDim lsCobertura(lnTotalColumnas - 1)
'    ReDim lsTotalOro(lnTotalColumnas - 1)
'
'    ReDim lsIndiceCobertura(lnTotalColumnas - 1)
'
'End Sub
'
'Private Sub Form_Load()
'CentraForm Me
'End Sub
