VERSION 5.00
Begin VB.Form frmPresServDeuda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PRESUPUESTO - Servicio a la Deuda"
   ClientHeight    =   450
   ClientLeft      =   1080
   ClientTop       =   1320
   ClientWidth     =   4080
   Icon            =   "frmPresServDeuda.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmPresServDeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'lnConcepto de Pago
' 0 Capital
' 1 Intereses + Comisiones

Private Type ltServicioDeuda
    lnCol       As Integer
    lsPlaza     As String
    lnPlazo     As Integer
    lnConcepto  As Integer
    lnCapital   As Currency
    lnInteres   As Currency
End Type
Dim ltDatos() As ltServicioDeuda

Dim xlAplicacion As Excel.Application
Dim xlLibro      As Excel.Workbook
Dim xlHoja1      As Excel.Worksheet
Dim lsArchivo    As String
Dim lbExcel As Boolean
Dim oCon    As DConecta

Dim lnAnio    As Integer
Dim lnTCambio As Currency
Dim ldFechaAl As Date
Dim nCont As Integer

Public Sub ImprimeServicioDeuda(pnAnio As Integer, pdFechaAl As Date, pnTCambio As Currency)
Dim lsMsgErr As String
On Error GoTo ImprimeServicioDeudaErr
lbExcel = False
lnTCambio = pnTCambio
ldFechaAl = pdFechaAl
lnAnio = pnAnio
    GetSaldoPeriodoAnterior CDate("01/01/" & Year(pdFechaAl)) - 1
    GetDeudasPresentePeriodo CDate("01/01/" & Year(pdFechaAl)), pdFechaAl
    GetPagoDeudas CDate("01/01/" & Year(pdFechaAl)), pdFechaAl, False
    GetPagoDeudas CDate("01/01/" & Year(pdFechaAl)), pdFechaAl, True
    GeneraReporteExcel
If lsArchivo <> "" Then
    '*******Carga el Archivo Excel a Objeto Ole ******
    CargaArchivo lsArchivo, App.path & "\SPOOLER\"
End If
Exit Sub
ImprimeServicioDeudaErr:
    lsMsgErr = Err.Description
    If lbExcel = True Then
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
        lbExcel = False
    End If
    MsgBox TextErr(lsMsgErr), vbInformation, "Aviso"
End Sub

Private Sub GeneraReporteExcel()
Dim lsHoja As String
Dim lnFila As Integer
Dim Y1 As Currency
Dim Y2 As Currency
Dim I  As Integer
    If Month(ldFechaAl) <= 3 Then
        lsHoja = "I"
    ElseIf Month(ldFechaAl) <= 6 Then
        lsHoja = "II"
    ElseIf Month(ldFechaAl) <= 9 Then
        lsHoja = "III"
    Else
        lsHoja = "IV"
    End If
    lsHoja = lsHoja & " TRIMESTRE"
    lsArchivo = App.path & "\Spooler\" & "P_" & lnAnio & "_ServDeuda.XLS"
    
    ExcelBegin lsArchivo, xlAplicacion, xlLibro, True
    ExcelAddHoja lsHoja, xlLibro, xlHoja1, True
    
    xlHoja1.PageSetup.Zoom = 80
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlAplicacion.Range("A1:Z1000").Font.Size = 8
    
    xlHoja1.Range("A1").ColumnWidth = 25
    xlHoja1.Range("B1:Z1").ColumnWidth = 12
    xlHoja1.Range("I1").ColumnWidth = 17
    lnFila = 1
    xlHoja1.Cells(lnFila, 1) = "EJECUCION INSTITUCIONAL DEL"
    xlHoja1.Cells(lnFila + 1, 1) = "PRESUPUESTO PARA EL"
    xlHoja1.Cells(lnFila + 2, 1) = "AÑO FISCAL " & lnAnio
    lnFila = 4
    xlHoja1.Cells(lnFila, 2) = "DETALLE DE LAS DEUDAS DE LA ENTIDAD"
    xlHoja1.Cells(lnFila + 1, 2) = "AL " & lsHoja & " DEL AÑO " & lnAnio
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 8)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 2), xlHoja1.Cells(lnFila + 1, 8)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila + 1, 2)).Font.Size = 12
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 2), xlHoja1.Cells(lnFila + 1, 2)).HorizontalAlignment = xlCenter
    
    lnFila = lnFila + 2
    xlHoja1.Cells(lnFila, 9) = "ANEXO Nº E-2"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, 9)).Font.Size = 10
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "RAZON SOCIAL :" & gsNomCmac
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "(EN NUEVOS SOLES)"
    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
    
    lnFila = lnFila + 1
    Y1 = lnFila
    xlHoja1.Cells(lnFila, 1) = "CONCEPTO"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila + 1, 1)).MergeCells = True
    
    xlHoja1.Cells(lnFila, 2) = "CUENTAS POR PAGAR AL 31/12/" & (lnAnio - 1)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 4)).MergeCells = True
    
    xlHoja1.Cells(lnFila, 5) = "DEUDAS ORIGINADAS AL " & lsHoja & " " & lnAnio
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila + 1, 5)).MergeCells = True
    xlHoja1.Cells(lnFila, 6) = "PAGO DE DEUDAS AL " & lsHoja & " " & lnAnio
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 8)).MergeCells = True
    xlHoja1.Cells(lnFila, 9) = "CUENTAS POR PAGAR AL " & lsHoja & " " & lnAnio
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila + 1, 9)).MergeCells = True
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "VENCIDA"
    xlHoja1.Cells(lnFila, 3) = "POR VENCER"
    xlHoja1.Cells(lnFila, 4) = "TOTAL"
    xlHoja1.Cells(lnFila, 6) = "VENCIDA al 31/12/" & (lnAnio - 1)
    xlHoja1.Cells(lnFila, 7) = "DEL AÑO " & lnAnio
    xlHoja1.Cells(lnFila, 8) = "TOTAL"
        
    Y2 = lnFila
    ExcelCuadro xlHoja1, 1, Y1, 9, Y2
    
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 1), xlHoja1.Cells(lnFila, 9)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 1), xlHoja1.Cells(lnFila, 9)).RowHeight = 23
    With xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 1), xlHoja1.Cells(lnFila, 9))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
    End With
    
    lnFila = lnFila + 1
    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
    xlHoja1.Cells(lnFila, 1) = "I. CUENTAS POR PAGAR (A)"
    xlHoja1.Cells(lnFila + 1, 1) = "   SUNAT"
    xlHoja1.Cells(lnFila + 2, 1) = "   Aduanas"
    xlHoja1.Cells(lnFila + 3, 1) = "   ESSALUD (IPSS)"
    xlHoja1.Cells(lnFila + 4, 1) = "   Comp.Tiempo Serv. (CTS)"
    xlHoja1.Cells(lnFila + 5, 1) = "   Ex-COLFONAVI"
    xlHoja1.Cells(lnFila + 6, 1) = "   Proveedores"
    xlHoja1.Cells(lnFila + 7, 1) = "   Personal"
    xlHoja1.Cells(lnFila + 8, 1) = "   Otros 5/"
    lnFila = lnFila + 8
    ExcelCuadro xlHoja1, 1, Y2 + 1, 9, CCur(lnFila)
    Y2 = lnFila
    Dim lsTotPlaza  As String
    Dim lnFilaPlazo As Integer
    Dim lsTotal     As String
    lnFila = lnFila + 1
    lsTotal = "+B" & (lnFila + 1)
    Llenavalores lnFila, "1", 0, "2.1 Endeudamiento Externo", "Largo Plazo"
    
    lnFila = lnFila + 1
    Llenavalores lnFila, "1", 1, "", "Corto Plazo"
    lnFila = lnFila + 1
    lsTotal = lsTotal & "+B" & (lnFila + 1)
    Llenavalores lnFila, "0", 0, "2.2 Endeudamiento Interno", "Largo Plazo"
    lnFila = lnFila + 1
    Llenavalores lnFila, "0", 1, "", "Corto Plazo"
    
    'Totales
    lnFila = lnFila + 1
    lsTotPlaza = "+B" & lnFila & "+C" & lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 4)).Formula = "=" & lsTotPlaza
    lsTotPlaza = "+F" & lnFila & "+G" & lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).Formula = "=" & lsTotPlaza
    xlHoja1.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila, 9)).Formula = "=" & "D" & lnFila & "+E" & lnFila & "-H" & lnFila
    xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 4)).AutoFill xlHoja1.Range("D" & (Y2 + 2) & ":D" & lnFila), xlFillDefault
    xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).AutoFill xlHoja1.Range("H" & (Y2 + 2) & ":H" & lnFila), xlFillDefault
    xlHoja1.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila, 9)).AutoFill xlHoja1.Range("I" & (Y2 + 2) & ":I" & lnFila), xlFillDefault
    'End Totales
    
    ExcelCuadro xlHoja1, 1, Y2 + 1, 9, CCur(lnFila)
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "TOTAL (I+II)"
    ExcelCuadro xlHoja1, 1, CCur(lnFila), 9, CCur(lnFila)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).Formula = "=" & lsTotal
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).AutoFill xlHoja1.Range("B" & lnFila & ":I" & lnFila), xlFillDefault
    xlHoja1.Range(xlHoja1.Cells(11, 2), xlHoja1.Cells(lnFila, 9)).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
    
    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
    
End Sub

Private Sub Llenavalores(ByRef lnFila As Integer, lsUltPlaza As String, lnUltPlazo As Integer, lsTextPlaza As String, lsTextPlazo As String)
Dim lnCol As Integer
Dim Y2    As Integer
Dim lnFilaPlaza As Integer
If lsTextPlaza <> "" Then
    lnFila = lnFila + 1
    lnFilaPlaza = lnFila
    xlHoja1.Cells(lnFila, 1) = lsTextPlaza
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).Formula = "=B" & (lnFila + 1) & "+B" & (lnFila + 4)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).AutoFill xlHoja1.Range("B" & lnFila & ":I" & lnFila), xlFillDefault
End If
If lsTextPlazo <> "" Then
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "   " & lsTextPlazo
    xlHoja1.Cells(lnFila + 1, 1) = "      Amortización"
    xlHoja1.Cells(lnFila + 2, 1) = "      Interes y Comisiones"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).Formula = "=B" & (lnFila + 1) & "+B" & (lnFila + 2)
    xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).AutoFill xlHoja1.Range("B" & lnFila & ":I" & lnFila), xlFillDefault
    lnFila = lnFila + 1
End If
If UBound(ltDatos) > 0 Then
    For lnCol = 1 To UBound(ltDatos)
        If ltDatos(lnCol).lsPlaza = lsUltPlaza And ltDatos(lnCol).lnPlazo = lnUltPlazo Then
            xlHoja1.Cells(lnFila, ltDatos(lnCol).lnCol) = ltDatos(lnCol).lnCapital
            xlHoja1.Cells(lnFila + 1, ltDatos(lnCol).lnCol) = ltDatos(lnCol).lnInteres
        End If
    Next
End If
End Sub

Private Function GetSaldoPeriodoAnterior(pdFecha As Date) As Boolean
Dim sSql  As String
Dim rs    As ADODB.Recordset
Dim lnIndiceVac As Double

On Error GoTo ErrorGeneraDatos
    Dim oDAdeud As DCaja_Adeudados
    Set oDAdeud = New DCaja_Adeudados
    lnIndiceVac = oDAdeud.CargaIndiceVAC(CDate("01/" & Format(Month(gdFecSis), "00") & "/" & Format(Year(gdFecSis), "0000")) - 1)
    Set oDAdeud = Nothing

sSql = "SELECT nPlazo, cPlaza, SUM( CASE WHEN cConcepto = 1 THEN nSaldoCap END ) nCapital, " _
     & "       SUM( CASE WHEN cConcepto = 2 THEN nSaldoCap END ) nInteres " _
     & "FROM ( SELECT CASE WHEN nCtaIFPlazo * cia.nCtaIFCuotas > 365 THEN 0 ELSE 1 END nPlazo, cia.cPlaza, '1' cConcepto , " _
     & "       Round( cia.nSaldoCap * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '2' THEN " & lnTCambio & " ELSE 1 END ,2) + " _
     & "       ISNULL( (SELECT SUM(mc.nMovImporte) " _
     & "                FROM Mov m JOIN MovCta mc on mc.nMovNro = m.nMovNro " _
     & "                       JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem " _
     & "                       JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%' " _
     & "                WHERE m.nMovFlag = 0 and LEFT(m.cmovnro,8) > '" & Format(pdFecha, gsFormatoMovFecha) & "' and oc.copecod IN ('" & OpeCGAdeudRepGeneralMN & "', '" & OpeCGAdeudRepGeneralME & "') and not m.cOpecod in ('401804','402804') and oc.cOpeCtaOrden = '0' " _
     & "                  and mif.cPersCod = ci.cPersCod and mif.cCtaIFCod = ci.cCtaIFCod " _
     & "                and mif.cIFTpo = ci.cIFTpo " _
     & "             ), 0) nSaldoCap " _
     & "FROM CtaIF ci JOIN CtaIFAdeudados cia ON cia.cPersCod = ci.cPersCod and cia.cIFTpo = ci.cIFTpo and cia.cCtaIFCod = ci.cCtaIFCod " _
     & "     LEFT JOIN IndiceVac iv ON iv.dIndiceVac = ISNULL(cia.dCuotaUltPago, ci.dCtaIFAper) " _
     & "WHERE ci.cCtaIFEstado = '" & gEstadoCtaIFActiva & "' and ci.cCtaIFCod like '" & Format(gTpoCtaIFCtaAdeud, "00") & "%' " _
     & "  and ci.dCtaIFAper <= '" & Format(pdFecha, gsFormatoFecha) & "' "
sSql = sSql & "UNION ALL "
sSql = sSql & "SELECT CASE WHEN nCtaIFPlazo * cia.nCtaIFCuotas > 365 THEN 0 ELSE 1 END nPlazo, cia.cPlaza, '2' cConcepto, " _
     & "       Round( (cic.nInteres + cic.nComision) * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '2' THEN " & lnTCambio & " ELSE 1 END ,2) " _
     & "FROM CtaIF ci JOIN CtaIFAdeudados cia ON cia.cPersCod = ci.cPersCod and cia.cIFTpo = ci.cIFTpo and cia.cCtaIFCod = ci.cCtaIFCod " _
     & "              JOIN CtaIFCalendario cic ON cic.cPersCod = ci.cPersCod and cic.cIFTpo = ci.cIFTpo and cic.cCtaIFCod = ci.cCtaIFCod " _
     & "     LEFT JOIN IndiceVac iv ON iv.dIndiceVac = ISNULL(cia.dCuotaUltPago, ci.dCtaIFAper) " _
     & "WHERE ci.cCtaIFEstado = '" & gEstadoCtaIFActiva & "' and ci.cCtaIFCod like '" & Format(gTpoCtaIFCtaAdeud, "00") & "%' " _
     & "  and ci.dCtaIFAper <= '" & Format(pdFecha, gsFormatoFecha) & "' and cic.dVencimiento > '" & Format(pdFecha, gsFormatoFecha) & "' ) a " _
     & "GROUP BY nPlazo, cPlaza"

Set oCon = New DConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
Set oCon = Nothing
Do While Not rs.EOF
    nCont = nCont + 1
    ReDim Preserve ltDatos(nCont)
    ltDatos(nCont).lnCol = 3
    ltDatos(nCont).lnPlazo = rs!nPlazo
    ltDatos(nCont).lsPlaza = rs!cPlaza
    ltDatos(nCont).lnCapital = rs!nCapital
    ltDatos(nCont).lnInteres = rs!nInteres
    rs.MoveNext
Loop
Exit Function
ErrorGeneraDatos:
    MsgBox "Error N° [" & Err.Number & "] " & Err.Description, vbInformation, "aviso"
End Function

Private Function GetDeudasPresentePeriodo(pdFechaDel As Date, pdFechaAl As Date) As Boolean
Dim sSql  As String
Dim rs    As ADODB.Recordset
Dim lnIndiceVac As Double

On Error GoTo ErrorGeneraDatos
    Dim oDAdeud As DCaja_Adeudados
    Set oDAdeud = New DCaja_Adeudados
    lnIndiceVac = oDAdeud.CargaIndiceVAC(CDate("01/" & Format(Month(gdFecSis), "00") & "/" & Format(Year(gdFecSis), "0000")) - 1)
    Set oDAdeud = Nothing

sSql = "SELECT nPlazo, cPlaza, SUM(nCapital) nCapital, SUM(nInteres) nInteres " _
     & "FROM (SELECT CASE WHEN nCtaIFPlazo * cia.nCtaIFCuotas > 365 THEN 0 ELSE 1 END nPlazo, cia.cPlaza, " _
     & "        Round( (cic.nCapital) * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '2' THEN " & lnTCambio & " ELSE 1 END ,2) nCapital, " _
     & "        Round( (cic.nInteres + cic.nComision) * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '2' THEN " & lnTCambio & " ELSE 1 END ,2) nInteres " _
     & "      FROM CtaIF ci JOIN CtaIFAdeudados cia ON cia.cPersCod = ci.cPersCod and cia.cIFTpo = ci.cIFTpo and cia.cCtaIFCod = ci.cCtaIFCod " _
     & "                    JOIN CtaIFCalendario cic ON cic.cPersCod = ci.cPersCod and cic.cIFTpo = ci.cIFTpo and cic.cCtaIFCod = ci.cCtaIFCod " _
     & "               LEFT JOIN IndiceVac iv ON iv.dIndiceVac = ISNULL(cia.dCuotaUltPago, ci.dCtaIFAper) " _
     & "      WHERE ci.cCtaIFEstado = '" & gEstadoCtaIFActiva & "' and ci.cCtaIFCod like '" & Format(gTpoCtaIFCtaAdeud, "00") & "%' " _
     & "        and ci.dCtaIFAper Between '" & Format(pdFechaDel, gsFormatoFecha) & "' and '" & Format(pdFechaAl, gsFormatoFecha) & "' " _
     & ") a " _
     & "GROUP BY nPlazo, cPlaza "
     
Set oCon = New DConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
Set oCon = Nothing
Do While Not rs.EOF
    nCont = nCont + 1
    ReDim Preserve ltDatos(nCont)
    ltDatos(nCont).lnCol = 5
    ltDatos(nCont).lnPlazo = rs!nPlazo
    ltDatos(nCont).lsPlaza = rs!cPlaza
    ltDatos(nCont).lnCapital = rs!nCapital
    ltDatos(nCont).lnInteres = rs!nInteres
    rs.MoveNext
Loop
Exit Function
ErrorGeneraDatos:
    MsgBox "Error N° [" & Err.Number & "] " & Err.Description, vbInformation, "aviso"
End Function

Private Function GetPagoDeudas(pdFechaDel As Date, pdFechaAl As Date, pbPeriodoActual As Boolean) As Boolean
Dim sSql  As String
Dim rs    As ADODB.Recordset
Dim lnIndiceVac As Double
Dim lsFiltroPeriodo As String
On Error GoTo ErrorGeneraDatos
If pbPeriodoActual Then
     lsFiltroPeriodo = "  and ci.dCtaIFAper >= '" & Format(pdFechaDel, gsFormatoFecha) & "' "
Else
     lsFiltroPeriodo = "  and ci.dCtaIFAper < '" & Format(pdFechaDel, gsFormatoFecha) & "' "
End If

Dim oDAdeud As DCaja_Adeudados
Set oDAdeud = New DCaja_Adeudados
lnIndiceVac = oDAdeud.CargaIndiceVAC(CDate("01/" & Format(Month(gdFecSis), "00") & "/" & Format(Year(gdFecSis), "0000")) - 1)
Set oDAdeud = Nothing

sSql = "SELECT nPlazo, cPlaza, SUM(nCapital) nCapital, SUM(nInteres) nInteres " _
     & "FROM ( " _
     & "SELECT CASE WHEN nCtaIFPlazo * cia.nCtaIFCuotas > 365 THEN 0 ELSE 1 END nPlazo, cia.cPlaza, " _
     & "       ISNULL( (SELECT SUM(mc.nMovImporte) " _
     & "                FROM Mov m JOIN MovCta mc on mc.nMovNro = m.nMovNro " _
     & "                       JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem " _
     & "                       JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%' " _
     & "                WHERE m.nMovFlag = " & gMovFlagVigente & " and LEFT(m.cmovnro,8) BETWEEN '" & Format(pdFechaDel, gsFormatoMovFecha) & "' and '" & Format(pdFechaAl, gsFormatoMovFecha) & "' and oc.cOpeCod IN ('" & OpeCGAdeudRepGeneralMN & "', '" & OpeCGAdeudRepGeneralME & "') and oc.cOpeCtaOrden = '0' " _
     & "                  and mif.cPersCod = ci.cPersCod and mif.cCtaIFCod = ci.cCtaIFCod " _
     & "                  and mif.cIFTpo = ci.cIFTpo and m.cMovNro = cic.cMovNro " _
     & "       ), 0) as nCapital, " _
     & "       Round( cic.nInteresPagado * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '2' THEN " & lnTCambio & " ELSE 1 END ,2) nInteres " _
     & "FROM CtaIF ci JOIN CtaIFAdeudados cia ON cia.cPersCod = ci.cPersCod and cia.cIFTpo = ci.cIFTpo and cia.cCtaIFCod = ci.cCtaIFCod " _
     & "              JOIN CtaIFCalendario cic ON cic.cPersCod = ci.cPersCod and cic.cIFTpo = ci.cIFTpo and cic.cCtaIFCod = ci.cCtaIFCod and cic.cEstado = '" & gTpoEstCuotaAdeudCanc & "' " _
     & "     LEFT JOIN IndiceVac iv ON iv.dIndiceVac = ISNULL(cia.dCuotaUltPago, ci.dCtaIFAper) " _
     & "WHERE ci.cCtaIFEstado = '" & gEstadoCtaIFActiva & "' and ci.cCtaIFCod like '" & Format(gTpoCtaIFCtaAdeud, "00") & "%' " & lsFiltroPeriodo _
     & ") a " _
     & " GROUP BY nPlazo, cPlaza "
     
Set oCon = New DConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
Set oCon = Nothing
Do While Not rs.EOF
    nCont = nCont + 1
    ReDim Preserve ltDatos(nCont)
    ltDatos(nCont).lnCol = IIf(pbPeriodoActual, 7, 6)
    ltDatos(nCont).lnPlazo = rs!nPlazo
    ltDatos(nCont).lsPlaza = rs!cPlaza
    ltDatos(nCont).lnCapital = rs!nCapital
    ltDatos(nCont).lnInteres = rs!nInteres
    rs.MoveNext
Loop
Exit Function
ErrorGeneraDatos:
    MsgBox "Error N° [" & Err.Number & "] " & Err.Description, vbInformation, "aviso"
End Function


